/**
 * カラオケサロン リノア シフト表・タイムカード  バックエンド（Google Apps Script）
 *
 * - データはこのスクリプトに紐づくスプレッドシートに保存します（シートは自動で作られます）
 * - フロント（web/index.html）から POST された JSON を受け取り、LINE のアクセストークンで本人確認してから処理します
 * - シフト公開・締切前・退勤忘れ・前日（出勤の前の日のお昼12時）の LINE 通知を Messaging API で送ります
 * - 時給（夜10時からの割増は設定で。既定はなし）・ボトル歩合からお給料の見込みを出し、「payroll」シートに書きます（経営ボードが読む）
 * - 営業日は朝6時で切り替え（夜中1時の退勤も前の日の勤務として記録）
 *
 * 初回セットアップ（README の手順どおり）
 *   1. スクリプトのプロパティに LIFF_CHANNEL_ID / MESSAGING_TOKEN / LIFF_URL を入れる
 *   2. 関数 setup を1回実行（シート作成・権限の許可）
 *   3. 関数 setupTriggers を1回実行（毎朝9時の通知）
 *   4. デプロイ → 新しいデプロイ → ウェブアプリ（実行ユーザー: 自分 / アクセス: 全員）
 */

const VERSION = "2026-10-28.4";   // 画面側（index.html の SERVER_VERSION）と合わせる。再デプロイの確認用
const TZ = "Asia/Tokyo";
const PROP = PropertiesService.getScriptProperties();

const DEFAULT_SETTINGS = {
  name: "カラオケサロン リノア", eyebrow: "KARAOKE SALON  RINOA", pin: "1234", deadlineDay: 20,
  openStart: "11:00", openEnd: "18:00",
  required: [2, 2, 2, 2, 2, 3, 3],
  closed: [], closedHol: true,
  tcUnit: 30, tcMode: "nearest",
  notifyWish: true,
  phonePunch: false,     // スマホ（LINE）からの打刻を許可するか。標準は店舗端末（iPad）のみ
  kioskPin: true,        // 店舗端末での打刻に本人の暗証番号を求める
  remind: true,          // 出勤の前の日のお昼12時に本人へ知らせる
  notifyPublish: false,  // シフトを公開したとき全員へ知らせる（2026-10 やっさん：不要）
  notifyDeadline: false, // 希望の締切3日前・当日に未提出の人と管理者へ知らせる（2026-10 やっさん：不要）
  nightExtra: 0,         // 夜10時〜朝5時の割増（2026-10-28 やっさん：割増しなし＝0。古い nightPremium=0.25 の行は使わない）
  saleKinds: ["シャンパン", "ワイン"],
  saleTax: 0             // ボトルの値段は税込みで入れ、歩合も税込みに掛ける（0.1 にすると税抜き＝税込÷1.1 に掛ける）
};
const DAY_START_HOUR = 6;   // 朝6時までは前の日の営業として扱う

const SHEETS = {
  settings:  ["key", "value"],
  staff:     ["id", "name", "short", "emp", "order", "active", "userId", "note", "createdAt", "pin", "wage", "rate"],
  shifts:    ["month", "status", "publishedAt", "cells"],
  wishes:    ["month", "staffId", "days", "memo", "updatedAt"],
  timecards: ["month", "staffId", "days"],
  admins:    ["userId", "name", "addedAt"],
  devices:   ["id", "name", "key", "createdAt", "lastSeen"],
  log:       ["at", "userId", "action", "detail"],
  bottles:   ["id", "date", "staffId", "kind", "item", "price", "qty", "checked", "deleted", "by", "createdAt"],
  payroll:   ["month", "fixed", "variable", "total", "hours", "salesTotal", "basis", "updatedAt", "detail"],
  kstyle:    ["month", "data", "updatedAt"]   // 旧表 K-style（R8）の勤務時間の写し（hoursJob が1時間ごとに書く）
};
const JSON_COLS = { settings: ["value"], shifts: ["cells"], wishes: ["days"], timecards: ["days"], payroll: ["detail"], kstyle: ["data"] };

/* ============================================================
   HTTP 入口
   ============================================================ */
function doGet() {
  return json_({ ok: true, service: "rinoa-shift", version: VERSION, time: nowStr_() });
}

function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: "bad_request" }); }
  try {
    if (body.deviceKey) {
      const device = verifyDevice_(body.deviceKey);
      const res = handleDevice_(body.action, body.data || {}, device);
      return json_(Object.assign({ ok: true }, res));
    }
    const auth = verifyUser_(body.accessToken);
    const res = handle_(body.action, body.data || {}, auth);
    return json_(Object.assign({ ok: true }, res));
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ============================================================
   本人確認（LINE アクセストークンの検証 → userId）
   ============================================================ */
function verifyUser_(token) {
  if (!token) throw new Error("unauthorized");
  const cache = CacheService.getScriptCache();
  const key = "tok:" + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, token)).slice(0, 40);
  const hit = cache.get(key);
  if (hit) return JSON.parse(hit);

  const v = UrlFetchApp.fetch("https://api.line.me/oauth2/v2.1/verify?access_token=" + encodeURIComponent(token), { muteHttpExceptions: true });
  if (v.getResponseCode() !== 200) throw new Error("unauthorized");
  const vj = JSON.parse(v.getContentText());
  const expected = PROP.getProperty("LIFF_CHANNEL_ID");
  if (expected && String(vj.client_id) !== String(expected)) throw new Error("unauthorized");

  const p = UrlFetchApp.fetch("https://api.line.me/v2/profile", { headers: { Authorization: "Bearer " + token }, muteHttpExceptions: true });
  if (p.getResponseCode() !== 200) throw new Error("unauthorized");
  const pj = JSON.parse(p.getContentText());
  const auth = { userId: pj.userId, name: pj.displayName || "" };
  cache.put(key, JSON.stringify(auth), 600);
  return auth;
}

/* ============================================================
   処理の振り分け
   ============================================================ */
function handle_(action, data, auth) {
  const admins = readAll_("admins");
  const isAdmin = admins.some(a => a.userId === auth.userId);
  const staff = readAll_("staff");
  const me = staff.find(s => s.userId === auth.userId) || null;
  const settings = getSettings_();

  const needAdmin = () => { if (!isAdmin) throw new Error("forbidden"); };
  const needMe = () => { if (!me) throw new Error("not_registered"); };

  switch (action) {
    case "bootstrap": {
      if (isAdmin) ensureTriggers_();
      return {
        version: VERSION,
        profile: { userId: auth.userId, name: auth.name },
        settings: publicSettings_(settings),
        staff: staff.map(s => publicStaff_(s, isAdmin)),
        me: me ? publicStaff_(me, true) : null,
        bizToday: bizDate_(),
        isAdmin,
        devices: isAdmin ? readAll_("devices").map(d => ({ id: d.id, name: d.name, key: d.key, createdAt: d.createdAt, lastSeen: d.lastSeen })) : [],
        liffUrl: PROP.getProperty("LIFF_URL") || "",
        months: monthsFor_(data.months || [], me, isAdmin)
      };
    }
    case "getMonths":
      return { months: monthsFor_(data.months || [], me, isAdmin) };

    case "register": {
      if (me) return { me: publicStaff_(me, true), staff: staff.map(s => publicStaff_(s, isAdmin)) };
      const name = String(data.name || "").trim().slice(0, 20);
      if (!name) throw new Error("name_required");
      const s = { id: "s_" + Date.now().toString(36), name, short: "", emp: "", order: staff.length + 1, active: true, userId: auth.userId, note: "", createdAt: nowStr_() };
      upsert_("staff", ["id"], s);
      log_(auth, "register", name);
      const all = readAll_("staff");
      return { me: publicStaff_(s, true), staff: all.map(x => publicStaff_(x, isAdmin)) };
    }

    case "saveWish": {
      const sid = (isAdmin && data.staffId) ? data.staffId : (me && me.id);
      if (!sid) throw new Error("not_registered");
      const month = ym_(data.month);
      const shift = readAll_("shifts").find(r => r.month === month);
      if (!isAdmin && shift && shift.status && shift.status !== "collecting") throw new Error("closed");
      const w = { month, staffId: sid, days: data.days || {}, memo: String(data.memo || "").slice(0, 300), updatedAt: nowStr_() };
      upsert_("wishes", ["month", "staffId"], w);
      log_(auth, "saveWish", month + " " + sid);
      if (settings.notifyWish && !isAdmin) {
        const who = staff.find(s => s.id === sid);
        const xs = Object.keys(w.days).filter(d => w.days[d] && (w.days[d].t === "o" || w.days[d].t === "p")).length;
        pushTo_(admins.map(a => a.userId), `【${settings.name}】${who ? who.name : ""}さんが${+month.slice(5)}月の希望を提出しました（入れる日 ${xs}日）。`);
      }
      return { wish: w };
    }

    case "punch": {
      needMe();
      if ((data.type === "in" || data.type === "out") && !settings.phonePunch && !isAdmin) throw new Error("phone_punch_disabled");
      const r = punch_(me, data, "phone");
      log_(auth, "punch:" + data.type, r.timecard.month);
      return r;
    }

    case "setPin": {
      needMe();
      const pin = String(data.pin || "");
      if (!/^\d{4,6}$/.test(pin)) throw new Error("bad_pin_format");
      upsert_("staff", ["id"], Object.assign({}, me, { pin: pinHash_(me.id, pin) }));
      log_(auth, "setPin", me.id);
      return { me: publicStaff_(Object.assign({}, me, { pin: "set" }), true) };
    }

    case "resetPin": {
      needAdmin();
      const s = staff.find(x => x.id === data.staffId); if (!s) throw new Error("bad_staff");
      const pin = data.pin ? String(data.pin) : "";
      if (pin && !/^\d{4,6}$/.test(pin)) throw new Error("bad_pin_format");
      upsert_("staff", ["id"], Object.assign({}, s, { pin: pin ? pinHash_(s.id, pin) : "" }));
      log_(auth, "resetPin", s.id);
      return { staff: readAll_("staff").map(x => publicStaff_(x, true)) };
    }

    case "createDevice": {
      needAdmin();
      const seg = () => { const cs = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789"; let o = ""; for (let i = 0; i < 4; i++) o += cs[Math.floor(Math.random() * cs.length)]; return o; };
      const dev = { id: "d_" + Date.now().toString(36), name: String(data.name || "店舗iPad").slice(0, 20), key: "RINOA-" + seg() + "-" + seg(), createdAt: nowStr_(), lastSeen: "" };
      upsert_("devices", ["id"], dev);
      log_(auth, "createDevice", dev.name);
      return { devices: readAll_("devices") };
    }

    case "deleteDevice": {
      needAdmin();
      remove_("devices", r => r.id === data.id);
      log_(auth, "deleteDevice", data.id);
      return { devices: readAll_("devices") };
    }

    case "adminLogin": {
      if (String(data.pin) !== String(settings.pin)) throw new Error("bad_pin");
      if (!isAdmin) upsert_("admins", ["userId"], { userId: auth.userId, name: auth.name, addedAt: nowStr_() });
      log_(auth, "adminLogin", "");
      return { isAdmin: true };
    }

    case "saveSettings": {
      needAdmin();
      const next = Object.assign({}, settings, sanitizeSettings_(data.settings || {}));
      Object.keys(next).forEach(k => upsert_("settings", ["key"], { key: k, value: next[k] }));
      log_(auth, "saveSettings", "");
      return { settings: publicSettings_(next) };
    }

    case "upsertStaff": {
      needAdmin();
      const s = data.staff || {};
      if (!s.id || !String(s.name || "").trim()) throw new Error("bad_staff");
      const prev = staff.find(x => x.id === s.id) || {};
      const row = Object.assign({ createdAt: nowStr_() }, prev, {
        id: String(s.id), name: String(s.name).trim().slice(0, 20), short: String(s.short || "").slice(0, 4), emp: String(s.emp || ""),
        order: +s.order || 99, active: s.active !== false, note: String(s.note || "").slice(0, 100), userId: s.userId !== undefined ? s.userId : (prev.userId || "")
      });
      if (s.wage !== undefined) { const w = Math.round(+s.wage || 0); if (w < 0 || w > 100000) throw new Error("bad_wage"); row.wage = w; }
      if (s.rate !== undefined) { const r = +s.rate || 0; if (r < 0 || r > 1) throw new Error("bad_rate"); row.rate = r; }
      upsert_("staff", ["id"], row);
      log_(auth, "upsertStaff", row.name);
      return { staff: readAll_("staff").map(x => publicStaff_(x, true)) };
    }

    case "deleteStaff": {
      needAdmin();
      remove_("staff", r => r.id === data.id);
      log_(auth, "deleteStaff", data.id);
      return { staff: readAll_("staff").map(x => publicStaff_(x, true)) };
    }

    case "saveShift": {
      needAdmin();
      const month = ym_(data.month);
      const prev = readAll_("shifts").find(r => r.month === month) || {};
      const doc = { month, status: data.status || prev.status || "collecting", publishedAt: prev.publishedAt || "", cells: data.cells || {} };
      if (doc.status === "published" && prev.status !== "published") {
        doc.publishedAt = nowStr_();
      }
      if (doc.status === "published" && prev.status !== "published" && settings.notifyPublish === true) {
        const ids = staff.filter(s => s.active !== false && s.userId).map(s => s.userId);
        const url = PROP.getProperty("LIFF_URL") || "";
        pushTo_(ids, `【${settings.name}】${+month.slice(5)}月のシフトを公開しました。\n「シフト表」タブで確認してください。${url ? "\n" + url : ""}`);
      }
      upsert_("shifts", ["month"], doc);
      log_(auth, "saveShift", month + " " + doc.status);
      return { shift: doc };
    }

    case "saveTimecard": {
      needAdmin();
      const month = ym_(data.month), sid = String(data.staffId);
      const prev = readAll_("timecards").find(r => r.month === month && r.staffId === sid) || { month, staffId: sid, days: {} };
      const next = { month, staffId: sid, days: {}, edits: (prev.edits || []).slice() };
      const by = whoName_(auth, staff) + "（経営者）";
      const days = new Set(Object.keys(prev.days || {}).concat(Object.keys(data.days || {})));
      days.forEach(d => {
        const a = (prev.days || {})[d], b = (data.days || {})[d];
        if (!b) { if (a) addEdit_(next, d, by, "記録を削除（" + recText_(a) + "）"); return; }
        const rec = Object.assign({}, a || {}, { in: b.in, brk: +b.brk || 0 }); if (b.out) rec.out = b.out; else delete rec.out;
        if (!a) { rec.manual = true; addEdit_(next, d, by, "記録を追加（" + recText_(rec) + "）"); }
        else if (recText_(a) !== recText_(rec)) addEdit_(next, d, by, recText_(a) + " → " + recText_(rec));
        if (rec.req && (!a || recText_(a) !== recText_(rec))) delete rec.req;   // 依頼に対応したら依頼を消す
        if (a && recText_(a) !== recText_(rec)) rec.edited = true;
        if (!a) rec.edited = true;
        next.days[d] = rec;
      });
      upsert_("timecards", ["month", "staffId"], next);
      log_(auth, "saveTimecard", month + " " + sid);
      return { timecard: next };
    }

    case "requestFix": {
      needMe();
      const month = ym_(data.month), d = String(+data.day), text = String(data.text || "").trim().slice(0, 200);
      if (!text) throw new Error("text_required");
      const doc = readAll_("timecards").find(r => r.month === month && r.staffId === me.id);
      if (!doc || !doc.days || !doc.days[d]) throw new Error("no_record");
      doc.days[d].req = { text, at: nowStr_().slice(0, 16) };
      upsert_("timecards", ["month", "staffId"], doc);
      log_(auth, "requestFix", month + "/" + d);
      pushTo_(admins.map(a => a.userId), `【${settings.name}】${me.name}さんから打刻修正の依頼\n${+month.slice(5)}/${d}（現在 ${recText_(doc.days[d])}）\n「${text}」\n管理画面の勤怠 → 明細・修正 から直せます。`);
      return { timecard: doc };
    }

    /* ---------- ボトル売上（歩合のもと） ---------- */
    case "getBottles": {
      const month = ym_(data.month);
      const sid = (isAdmin && !data.staffId) ? null : ((isAdmin && data.staffId) ? data.staffId : (me && me.id));
      if (!isAdmin && !sid) throw new Error("not_registered");
      return { bottles: bottlesOf_(month).filter(b => !sid || b.staffId === sid) };
    }
    case "addBottle": {
      const sid = (isAdmin && data.staffId) ? String(data.staffId) : (me && me.id);
      if (!sid) throw new Error("not_registered");
      const date = String(data.date || "");
      if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error("bad_date");
      if (date > bizDate_()) throw new Error("future_date");
      if (!isAdmin && date < addDays_(bizDate_(), -40)) throw new Error("too_old");
      if ((settings.saleKinds || DEFAULT_SETTINGS.saleKinds).indexOf(data.kind) < 0) throw new Error("bad_kind");
      const price = Math.round(+data.price || 0), qty = Math.round(+data.qty || 1);
      if (!(price > 0 && price < 10000000)) throw new Error("bad_price");
      if (!(qty >= 1 && qty <= 50)) throw new Error("bad_qty");
      const b = { id: "b_" + Date.now().toString(36), date, staffId: sid, kind: data.kind, item: String(data.item || "").trim().slice(0, 40), price, qty, checked: false, deleted: false, by: auth.userId, createdAt: nowStr_() };
      upsert_("bottles", ["id"], b);
      log_(auth, "addBottle", date + " " + sid + " " + price + "x" + qty);
      return { bottle: b };
    }
    case "updateBottle": {
      const b = readAll_("bottles").find(x => x.id === data.id); if (!b) throw new Error("bad_bottle");
      const own = me && b.staffId === me.id;
      if (!isAdmin && (!own || isTrue_(b.checked))) throw new Error("bottle_locked");
      if (data.remove) b.deleted = true;
      if (data.checked !== undefined) { needAdmin(); b.checked = !!data.checked; }
      upsert_("bottles", ["id"], b);
      log_(auth, "updateBottle", b.id + (data.remove ? " remove" : "") + (data.checked !== undefined ? " checked=" + !!data.checked : ""));
      return { bottle: b };
    }

    /* ---------- 公式アカウントのトーク画面の下のメニュー ---------- */
    case "setupRichMenu": {
      needAdmin();
      const r = setupRichMenu_();
      log_(auth, "setupRichMenu", r.richMenuId || "");
      return r;
    }
    case "lineStatus": {
      needAdmin();
      const token = PROP.getProperty("MESSAGING_TOKEN");
      // ok は通信の成否に使うので、つながりの結果は connected で返す
      if (!token) return { connected: false, reason: "no_token" };
      const res = UrlFetchApp.fetch("https://api.line.me/v2/bot/info", { headers: { Authorization: "Bearer " + String(token).trim() }, muteHttpExceptions: true });
      if (res.getResponseCode() !== 200) return { connected: false, reason: "bad_token", code: res.getResponseCode() };
      const j = JSON.parse(res.getContentText());
      return { connected: true, name: j.displayName || "", basicId: j.basicId || "" };
    }

    /* ---------- お給料の見込み（経営ボードへも書く） ---------- */
    case "getPay": {
      needAdmin();
      return { pay: savePayroll_(ym_(data.month)) };
    }

    case "hoursBook": {   // 勤務時間表（見る用のスプレッドシート）を書き直してURLを返す
      needAdmin();
      return { url: hoursJob().getUrl() };
    }

    default:
      throw new Error("unknown_action");
  }
}

/* ============================================================
   お給料の見込み
   時間：打刻（丸め後・休憩を引く）があればそれ、なければ公開済みシフトの予定
   金額：時給×時間 ＋ 時給×割増×夜10時〜朝5時の時間 ＋ ボトル売上（税込）×歩合
   ============================================================ */
function bottlesOf_(month) {
  return readAll_("bottles").filter(b => String(b.date).slice(0, 7) === month && !isTrue_(b.deleted))
    .map(b => ({ id: b.id, date: b.date, staffId: b.staffId, kind: b.kind, item: b.item || "", price: +b.price || 0, qty: +b.qty || 1, checked: isTrue_(b.checked) }))
    .sort((a, b) => a.date < b.date ? 1 : a.date > b.date ? -1 : 0);
}
function hm2m_(hm) { const p = String(hm || "").split(":"); return (+p[0]) * 60 + (+p[1] || 0); }
function roundM_(m, kind, S) {
  const u = +S.tcUnit || 30, mode = S.tcMode || "nearest";
  if (mode === "floor") return Math.floor(m / u) * u;
  if (mode === "ceil") return Math.ceil(m / u) * u;
  if (mode === "favor") return kind === "in" ? Math.ceil(m / u) * u : Math.floor(m / u) * u;
  return Math.round(m / u) * u;
}
// 開始・終了（分。終了が前なら翌日）から、全体の分と夜10時〜朝5時の分
function spanMins_(a, b, brk, exact) {
  if (!exact && b <= a) b += 1440;
  const total = Math.max(0, b - a - (+brk || 0));
  let night = 0;
  [[22 * 60, 29 * 60], [-2 * 60, 5 * 60]].forEach(w => { night += Math.max(0, Math.min(b, w[1]) - Math.max(a, w[0])); });
  return { total, night: Math.min(night, total) };
}
function recMins_(rec, S) {
  if (!rec || !rec.in || !rec.out) return null;
  const a = roundM_(hm2m_(rec.in.slice(11, 16)), "in", S);
  let b = roundM_(hm2m_(rec.out.slice(11, 16)), "out", S);
  const dd = Math.round((new Date(rec.out.slice(0, 10) + "T00:00:00") - new Date(rec.in.slice(0, 10) + "T00:00:00")) / 86400000);
  b += dd * 1440;
  // 日付が同じで退勤の時刻が前のときだけ次の日とみなす（画面の calcRec と同じ）。
  // 丸めで出勤・退勤が同じ時刻になったら0時間（1:53→1:54 が24時間と数えられていた）
  if (dd === 0 && rec.out.slice(11, 16) < rec.in.slice(11, 16)) b += 1440;
  return spanMins_(a, Math.max(a, b), rec.brk, true);
}
function cellMins_(val, S) {
  const base = String(val || "").split("|")[0];
  if (!base || base === "OFF") return null;
  let s = S.openStart, e = S.openEnd;
  if (base.indexOf("C:") === 0) { const p = base.slice(2).split("-"); s = p[0]; e = p[1]; }
  return spanMins_(hm2m_(s), hm2m_(e), 0);
}
function computePay_(month) {
  const S = getSettings_();
  const prem = +S.nightExtra || 0;
  const tax = +S.saleTax || 0;
  const staff = readAll_("staff");
  const tcs = readAll_("timecards").filter(r => r.month === month);
  const sh = readAll_("shifts").find(r => r.month === month);
  const cells = sh && sh.status === "published" ? (sh.cells || {}) : {};
  const bottles = bottlesOf_(month);
  const ks = kstyleOf_(month);
  const y = +month.slice(0, 4), m = +month.slice(5, 7), n = new Date(y, m, 0).getDate();
  const rows = staff.map(s => {
    const doc = tcs.find(r => r.staffId === s.id) || { days: {} };
    const kd = kstyleDays_(ks, s.id);
    let days = 0, mins = 0, night = 0, planMins = 0;
    for (let d = 1; d <= n; d++) {
      let w = recMins_((doc.days || {})[d], S), plan = false;
      if (!w && kd[d]) w = { total: Math.round(kd[d] * 60), night: 0 };   // 打刻がない日は K-style（R8）の時間
      if (!w) { w = cellMins_(cells[d + "_" + s.id], S); plan = !!w; }
      if (!w || !w.total) continue;
      days++; mins += w.total; night += w.night; if (plan) planMins += w.total;
    }
    const wage = +s.wage || 0, rate = +s.rate || 0;
    const mine = bottles.filter(b => b.staffId === s.id);
    const salesTotal = mine.reduce((t, b) => t + b.price * b.qty, 0);
    const wagePay = Math.round(wage * mins / 60 + wage * prem * night / 60);
    const salesNet = Math.round(salesTotal / (1 + tax));
    const commission = Math.round(salesNet * rate);
    return { staffId: s.id, name: s.name, days, hours: Math.round(mins / 6) / 10, nightHours: Math.round(night / 6) / 10, planHours: Math.round(planMins / 6) / 10,
      wage, rate, wagePay, salesTotal, salesNet, bottles: mine.reduce((t, b) => t + b.qty, 0), commission, total: wagePay + commission };
  }).filter(r => r.days || r.salesTotal);
  const sum = k => rows.reduce((t, r) => t + r[k], 0);
  return { month, staff: rows, fixed: sum("wagePay"), variable: sum("commission"), total: sum("total"), hours: Math.round(sum("hours") * 10) / 10, salesTotal: sum("salesTotal"),
    salesNet: sum("salesNet"), saleTax: tax, noWage: rows.filter(r => r.hours && !r.wage).map(r => r.name), nightPremium: prem, basis: "シフト表からの見込み（打刻があれば打刻、なければK-style（R8）の時間、なければ公開済みシフトの予定）。freeeの給与とは別" };
}
// 経営ボード（rinoa-order）が読むシート。fixed＝時給の分（固定費）、variable＝歩合の分（変動費）
function savePayroll_(month) {
  const p = computePay_(month);
  upsert_("payroll", ["month"], { month, fixed: p.fixed, variable: p.variable, total: p.total, hours: p.hours, salesTotal: p.salesTotal, basis: p.basis, updatedAt: nowStr_(),
    detail: p.staff.map(r => ({ name: r.name, hours: r.hours, nightHours: r.nightHours, wagePay: r.wagePay, salesTotal: r.salesTotal, salesNet: r.salesNet, commission: r.commission, total: r.total })) });
  return p;
}
function payrollJob() {
  const now = new Date();
  [fmt_(now, "yyyy-MM"), fmt_(new Date(now.getFullYear(), now.getMonth() - 1, 1), "yyyy-MM")].forEach(m => { try { savePayroll_(m); } catch (e) { Logger.log("payroll " + m + " " + e); } });
}

/* ============================================================
   勤務時間表（見る用の別スプレッドシート「リノア 勤務時間表」）
   旧表 K-style（R8）と同じく、縦にスタッフ・横に日付で「何時間働いたか」を並べる。
   「昨日」シート＝きのう誰が何時から何時まで働いたか。1時間ごとに自動で書き直す（手で直しても上書きされる）
   ============================================================ */
const HOURS_WD = ["日", "月", "火", "水", "木", "金", "土"];
function hoursBook_() {
  const id = PROP.getProperty("HOURS_SS_ID");
  if (id) { try { return SpreadsheetApp.openById(id); } catch (e) { Logger.log("hoursBook " + e); } }
  const ss = SpreadsheetApp.create("リノア 勤務時間表");
  PROP.setProperty("HOURS_SS_ID", ss.getId());
  return ss;
}
function hoursSheet_(ss, name, index) {
  const sh = ss.getSheetByName(name) || ss.insertSheet(name, index);
  sh.clear(); sh.clearNotes(); sh.setFrozenRows(0); sh.setFrozenColumns(0);
  return sh;
}
// "2026-10-28T23:31:00" → "23:31"、日付が次の日なら "翌1:15"
function hoursClock_(stamp, baseYmd) {
  if (!stamp) return "";
  const t = String(stamp).slice(11, 16).replace(/^0/, "");
  return String(stamp).slice(0, 10) > baseYmd ? "翌" + t : t;
}
// スタッフ1人・1日ぶん：{ h:時間, plan:予定か, open:退勤未, rec }
function hoursDay_(doc, cells, s, ymd, S, kd) {
  const d = +ymd.slice(8), rec = ((doc && doc.days) || {})[d];
  if (rec && rec.in && !rec.out) return { open: true, rec };
  const w = recMins_(rec, S);
  if (w) return { h: w.total / 60, rec };
  if (kd && kd[d]) return { h: kd[d], ks: true };
  const p = cellMins_(cells[d + "_" + s.id], S);
  if (p && p.total) return { h: p.total / 60, plan: true };
  return null;
}
function hoursStaff_(month, tcs) {
  return readAll_("staff").filter(s => (s.active !== false && s.active !== "FALSE") || tcs.some(t => t.staffId === s.id && Object.keys(t.days || {}).length))
    .sort((a, b) => (+a.order || 99) - (+b.order || 99));
}
function writeHoursMonth_(ss, month) {
  const S = getSettings_();
  const y = +month.slice(0, 4), m = +month.slice(5, 7), n = new Date(y, m, 0).getDate();
  const tcs = readAll_("timecards").filter(r => r.month === month);
  const shRow = readAll_("shifts").find(r => r.month === month);
  const cells = shRow && shRow.status === "published" ? (shRow.cells || {}) : {};
  const pay = computePay_(month), payOf = id => pay.staff.find(r => r.staffId === id) || {};
  const ks = kstyleOf_(month);
  const staff = hoursStaff_(month, tcs).concat(Object.keys(ks).filter(l => !ks[l].staffId).map(l => ({ id: "", name: l + "（シフト表に未登録）", ksLabel: l })));
  const W = n + 8;   // A列＋日付n列＋右の7列
  const sh = hoursSheet_(ss, y + "年" + m + "月", 1);
  const vals = [], colors = [], notes = [];
  const blank = () => new Array(W).fill("");
  const r1 = blank(); r1[0] = "リノア 勤務時間表　" + y + "年" + m + "月"; r1[n + 1] = "更新：" + fmt_(new Date(), "M/d H:mm");
  const r2 = blank(); r2[0] = "黒い数字＝打刻した時間（" + (+S.tcUnit || 30) + "分刻み・休憩を引く）／灰色＝打刻がない日のシフトの予定／「出勤中」＝まだ退勤していない／青い数字＝旧表 K-style（R8）から写した時間。数字のマスに印（右上の三角）があれば、出勤・退勤の時刻が見られます";
  const r3 = ["日付"], r4 = ["曜日"];
  for (let d = 1; d <= n; d++) { r3.push(m + "/" + d); r4.push(HOURS_WD[new Date(y, m - 1, d).getDay()]); }
  r3.push("打刻・K-styleの時間", "予定も入れた時間", "出勤日数", "時給", "時給の分", "歩合", "お給料の見込み"); r4.push("", "", "", "", "", "", "");
  [r1, r2, r3, r4].forEach(r => { vals.push(r); colors.push(new Array(W).fill("#000000")); notes.push(new Array(W).fill("")); });
  const dayTot = new Array(n + 1).fill(0);
  staff.forEach(s => {
    const doc = tcs.find(t => t.staffId === s.id) || { days: {} };
    const kd = s.ksLabel ? ks[s.ksLabel].days : kstyleDays_(ks, s.id);
    const row = [s.name], col = ["#000000"], nt = [""];
    let real = 0;
    for (let d = 1; d <= n; d++) {
      const ymd = month + "-" + ("0" + d).slice(-2), x = hoursDay_(s.ksLabel ? null : doc, s.ksLabel ? {} : cells, s, ymd, S, kd);
      if (!x) { row.push(""); col.push("#000000"); nt.push(""); continue; }
      if (x.open) { row.push("出勤中"); col.push("#c0392b"); nt.push("出勤 " + hoursClock_(x.rec.in, ymd) + "（退勤がまだ）"); continue; }
      row.push(Math.round(x.h * 100) / 100); col.push(x.plan ? "#a0a0a0" : x.ks ? "#2a6fdb" : "#000000");
      nt.push(x.plan ? "打刻なし（シフトの予定）" : x.ks ? "K-style（R8）から写した時間" : "出勤 " + hoursClock_(x.rec.in, ymd) + " → 退勤 " + hoursClock_(x.rec.out, ymd) + (+x.rec.brk ? "（休憩" + x.rec.brk + "分）" : "") + (x.rec.edited ? "\n経営者が修正" : ""));
      if (!x.plan) { real += x.h; dayTot[d] += x.h; }
    }
    const p = s.ksLabel ? {} : payOf(s.id);
    if (s.ksLabel) row.push(Math.round(real * 100) / 100, Math.round(real * 100) / 100, Object.keys(kd).length, "", "", "", "");
    else row.push(Math.round(real * 100) / 100, p.hours || 0, p.days || 0, +s.wage || 0, p.wagePay || 0, p.commission || 0, p.total || 0);
    for (let i = 0; i < 7; i++) { col.push("#000000"); nt.push(""); }
    vals.push(row); colors.push(col); notes.push(nt);
  });
  const tot = ["合計"];
  for (let d = 1; d <= n; d++) tot.push(dayTot[d] ? Math.round(dayTot[d] * 100) / 100 : "");
  const sumCol = k => staff.reduce((t, s) => t + (+(payOf(s.id)[k]) || 0), 0);
  const extraH = staff.filter(s => s.ksLabel).reduce((t, s) => t + Object.values(ks[s.ksLabel].days).reduce((a, b) => a + b, 0), 0);
  tot.push(Math.round(dayTot.reduce((a, b) => a + b, 0) * 100) / 100, Math.round(((pay.hours || 0) + extraH) * 100) / 100, "", "", pay.fixed || 0, pay.variable || 0, pay.total || 0);
  vals.push(tot); colors.push(new Array(W).fill("#000000")); notes.push(new Array(W).fill(""));
  const rows = vals.length;
  [vals, colors, notes].forEach(a => a.forEach(r => { while (r.length < W) r.push(a === colors ? "#000000" : ""); r.length = W; }));   // 行の長さをそろえる（ずれると setValues が止まる）
  sh.getRange(1, 1, rows, W).setValues(vals).setFontColors(colors).setNotes(notes).setVerticalAlignment("middle");
  sh.getRange(1, 1).setFontSize(14).setFontWeight("bold");
  sh.getRange(2, 1).setFontColor("#666666").setFontSize(9);
  sh.getRange(3, 1, 2, W).setFontWeight("bold").setBackground("#f3f3f3").setHorizontalAlignment("center");
  for (let d = 1; d <= n; d++) {
    const w = new Date(y, m - 1, d).getDay();
    if (w === 0 || w === 6) sh.getRange(3, d + 1, 2, 1).setFontColor(w === 0 ? "#c0392b" : "#2a6fdb");
  }
  if (staff.length) {
    sh.getRange(5, 2, staff.length + 1, n + 3).setNumberFormat("0.0#").setHorizontalAlignment("center");
    sh.getRange(5, n + 4, staff.length + 1, 1).setNumberFormat("0").setHorizontalAlignment("center");
    sh.getRange(5, n + 5, staff.length + 1, 4).setNumberFormat("#,##0").setHorizontalAlignment("right");
    sh.getRange(5, 1, staff.length, 1).setFontWeight("bold");
  }
  sh.getRange(rows, 1, 1, W).setFontWeight("bold").setBackground("#fff7e0");
  sh.getRange(3, n + 2, rows - 2, 1).setBackground("#eef6ee");
  sh.getRange(3, W, rows - 2, 1).setBackground("#eef6ee").setFontWeight("bold");
  sh.getRange(3, 1, rows - 2, W).setBorder(true, true, true, true, true, true, "#cccccc", SpreadsheetApp.BorderStyle.SOLID);
  sh.setColumnWidth(1, 150); sh.setColumnWidths(2, n, 44); sh.setColumnWidths(n + 2, 7, 92);
  sh.setRowHeight(2, 20); sh.setFrozenRows(4); sh.setFrozenColumns(1);
  return sh;
}
// 「昨日」：きのうの営業日に、だれが何時から何時まで、何時間働いたか
function writeHoursYesterday_(ss) {
  const S = getSettings_();
  const ymd = addDays_(bizDate_(), -1), month = ymd.slice(0, 7), d = +ymd.slice(8);
  const p = ymd.split("-").map(Number), wd = HOURS_WD[new Date(p[0], p[1] - 1, p[2]).getDay()];
  const tcs = readAll_("timecards").filter(r => r.month === month);
  const shRow = readAll_("shifts").find(r => r.month === month);
  const cells = shRow && shRow.status === "published" ? (shRow.cells || {}) : {};
  const staff = hoursStaff_(month, tcs), ks = kstyleOf_(month);
  const sh = hoursSheet_(ss, "昨日", 0);
  const vals = [["きのう " + p[1] + "月" + p[2] + "日（" + wd + "）の勤務", "", "", "", "", "", "", ""],
    ["更新：" + fmt_(new Date(), "M/d H:mm") + "　（朝6時までは前の日として数えます）", "", "", "", "", "", "", ""],
    ["名前", "出勤", "退勤", "休憩（分）", "働いた時間", "時給", "時給の分", "メモ"]];
  let totH = 0, totPay = 0, any = 0;
  staff.forEach(s => {
    const doc = tcs.find(t => t.staffId === s.id) || { days: {} }, x = hoursDay_(doc, cells, s, ymd, S, kstyleDays_(ks, s.id));
    if (!x) return;
    any++;
    if (x.open) { vals.push([s.name, hoursClock_(x.rec.in, ymd), "", "", "", +s.wage || "", "", "退勤の打刻がありません"]); return; }
    if (x.plan) { vals.push([s.name, "", "", "", "", +s.wage || "", "", "打刻がありません（シフトでは " + Math.round(x.h * 10) / 10 + " 時間の予定）"]); return; }
    const pay = Math.round((+s.wage || 0) * x.h);
    totH += x.h; totPay += pay;
    if (x.ks) { vals.push([s.name, "", "", "", Math.round(x.h * 100) / 100, +s.wage || "", pay, "K-style（R8）から写した時間"]); return; }
    vals.push([s.name, hoursClock_(x.rec.in, ymd), hoursClock_(x.rec.out, ymd), +x.rec.brk || 0, Math.round(x.h * 100) / 100, +s.wage || "", pay, x.rec.edited ? "経営者が修正" : ""]);
  });
  if (!any) vals.push(["この日の勤務はありません", "", "", "", "", "", "", ""]);
  vals.push(["合計", "", "", "", Math.round(totH * 100) / 100, "", totPay, ""]);
  // 今月これまで（月の初めのころは先月分も見られるよう、月のシートへ案内）
  const pay = computePay_(month);
  vals.push(["", "", "", "", "", "", "", ""]);
  vals.push([p[1] + "月のこれまで（打刻・K-style（R8）＋これからの予定）", "", "", "", "", "", "", ""]);
  vals.push(["名前", "出勤日数", "時間", "", "時給の分", "歩合", "お給料の見込み", ""]);
  const sumStart = vals.length + 1;
  pay.staff.forEach(r => vals.push([r.name, r.days, r.hours, "", r.wagePay, r.commission, r.total, ""]));
  vals.push(["合計", "", pay.hours, "", pay.fixed, pay.variable, pay.total, "日ごとの表は「" + p[0] + "年" + p[1] + "月」シート"]);
  const rows = vals.length, lastY = 4 + Math.max(any, 1);
  sh.getRange(1, 1, rows, 8).setValues(vals).setVerticalAlignment("middle");
  sh.getRange(1, 1).setFontSize(16).setFontWeight("bold");
  sh.getRange(2, 1).setFontColor("#666666").setFontSize(9);
  [3, sumStart - 1].forEach(r => sh.getRange(r, 1, 1, 8).setFontWeight("bold").setBackground("#f3f3f3"));
  sh.getRange(sumStart - 2, 1).setFontSize(13).setFontWeight("bold");
  [lastY, rows].forEach(r => sh.getRange(r, 1, 1, 8).setFontWeight("bold").setBackground("#fff7e0"));
  sh.getRange(4, 2, lastY - 3, 2).setHorizontalAlignment("center");
  sh.getRange(4, 5, lastY - 3, 1).setNumberFormat("0.0#").setFontWeight("bold").setFontSize(12);
  sh.getRange(4, 6, lastY - 3, 2).setNumberFormat("#,##0");
  sh.getRange(sumStart, 3, rows - sumStart + 1, 1).setNumberFormat("0.0#");
  sh.getRange(sumStart, 5, rows - sumStart + 1, 3).setNumberFormat("#,##0");
  sh.getRange(3, 1, lastY - 2, 8).setBorder(true, true, true, true, true, true, "#cccccc", SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(sumStart - 1, 1, rows - sumStart + 2, 7).setBorder(true, true, true, true, true, true, "#cccccc", SpreadsheetApp.BorderStyle.SOLID);
  sh.setColumnWidth(1, 120); sh.setColumnWidths(2, 6, 90); sh.setColumnWidth(8, 300);
  sh.setFrozenRows(3);
  return sh;
}
/* ---------- 旧表 K-style（R8）の勤務時間を読む（読むだけ・旧表は変えない） ----------
   月シート（「10月」など）の「人件費合計」の行から「材料費合計」の行までのうち、名前の行（ドリンク・ボトル手当などを除く）の
   1日〜月末の数字＝働いた時間。昨日までの分を kstyle シートに写す。打刻がある日は打刻が優先 */
const KSTYLE_SS_ID = "12RQVnhF50hkIInYJQlL0NqPcpO6gR0Q9nKoppU0_BME";
const KSTYLE_SKIP = /ドリンク|ボトル|ゴルフ|タイムカード|手当|合計/;
// K-style の呼び名 → シフト表の名前に含まれる文字
const KSTYLE_ALIAS = { "真美": ["まみ", "真美"], "サキ": ["紗希", "サキ"], "弥生": ["弥生"], "侑冴": ["侑冴"], "みゆき": ["miyuki", "みゆき"], "佳代": ["佳代"], "えり": ["えり", "エリ"], "マオ": ["真央", "マオ"] };
function kstyleImport_() {
  const today = bizDate_(), months = [today.slice(0, 7)];
  if (+today.slice(8) <= 5) { const p = today.split("-").map(Number); months.push(fmt_(new Date(p[0], p[1] - 2, 1, 12), "yyyy-MM")); }
  const old = SpreadsheetApp.openById(KSTYLE_SS_ID), staff = readAll_("staff");
  months.forEach(month => {
    const sh = old.getSheetByName(+month.slice(5) + "月");
    if (!sh) return;
    const y = +month.slice(0, 4), m = +month.slice(5), n = new Date(y, m, 0).getDate();
    const last = month < today.slice(0, 7) ? n : +today.slice(8) - 1;   // 昨日まで
    const v = sh.getRange(1, 1, Math.min(sh.getLastRow(), 200), n + 1).getValues();
    const from = v.findIndex(r => String(r[0]).trim() === "人件費合計"), to = v.findIndex((r, i) => i > from && /材料費合計/.test(String(r[0])));
    if (from < 0) return;
    const data = {};
    v.slice(from + 1, to < 0 ? undefined : to).forEach(r => {
      const label = String(r[0]).trim();
      if (!label || KSTYLE_SKIP.test(label)) return;
      const days = {};
      for (let d = 1; d <= last; d++) { const h = +r[d]; if (h > 0 && h < 24) days[d] = h; }
      if (!Object.keys(days).length) return;
      const keys = KSTYLE_ALIAS[label] || [label];
      const st = staff.find(s => keys.some(k => String(s.name).indexOf(k) >= 0 || String(s.short || "").indexOf(k) >= 0));
      data[label] = { staffId: st ? st.id : "", days };
    });
    upsert_("kstyle", ["month"], { month, data, updatedAt: nowStr_() });
  });
}
function kstyleOf_(month) { const r = readAll_("kstyle").find(x => x.month === month); return (r && r.data) || {}; }
function kstyleDays_(ks, staffId) { const o = {}; Object.keys(ks).forEach(l => { if (staffId && ks[l].staffId === staffId) Object.assign(o, ks[l].days); }); return o; }
// 1時間ごと：昨日・今月（月の初め5日間は先月も）を書き直す
function hoursJob() {
  try { kstyleImport_(); } catch (e) { Logger.log("kstyle " + e); }
  const ss = hoursBook_();
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return ss;
  try {
    const today = bizDate_(), month = today.slice(0, 7);
    writeHoursYesterday_(ss);
    writeHoursMonth_(ss, month);
    if (+today.slice(8) <= 5) { const p = today.split("-").map(Number); writeHoursMonth_(ss, fmt_(new Date(p[0], p[1] - 2, 1, 12), "yyyy-MM")); }
    const def = ss.getSheetByName("シート1") || ss.getSheetByName("Sheet1");
    if (def && ss.getSheets().length > 1) ss.deleteSheet(def);
    // 月のシートは新しい順に並べる（「昨日」が先頭）
    ss.getSheets().filter(s => /^\d{4}年\d{1,2}月$/.test(s.getName()))
      .sort((a, b) => { const k = s => { const q = s.getName().match(/\d+/g); return +q[0] * 100 + +q[1]; }; return k(b) - k(a); })
      .forEach((s, i) => { ss.setActiveSheet(s); ss.moveActiveSheet(i + 2); });
    ss.setActiveSheet(ss.getSheetByName("昨日")); ss.moveActiveSheet(1);
  } finally { lock.releaseLock(); }
  return ss;
}

/* ============================================================
   リッチメニュー（トーク画面の下の3つのボタン）：タイムカード／シフト表／希望を出す
   画像は GitHub Pages の docs/richmenu.png。鍵は MESSAGING_TOKEN（スクリプトのプロパティ）を使う
   ============================================================ */
const LIFF_BASE = "https://liff.line.me/2012191949-slfaXSbA";
const RICHMENU_IMAGE = "https://sakai4528.github.io/rinoa-shift/docs/richmenu.png";
function setupRichMenu_() {
  const token = String(PROP.getProperty("MESSAGING_TOKEN") || "").trim();
  if (!token) throw new Error("no_token");
  const H = { Authorization: "Bearer " + token };
  const call = (method, url, payload, contentType) => {
    const o = { method, headers: H, muteHttpExceptions: true };
    if (payload !== undefined) { o.payload = payload; o.contentType = contentType || "application/json"; }
    const r = UrlFetchApp.fetch(url, o);
    if (r.getResponseCode() >= 300) throw new Error("line_api " + r.getResponseCode() + " " + r.getContentText().slice(0, 200));
    return r.getContentText() ? JSON.parse(r.getContentText()) : {};
  };
  const NAME = "リノア シフト表";
  (call("get", "https://api.line.me/v2/bot/richmenu/list").richmenus || []).forEach(m => { if (m.name === NAME) call("delete", "https://api.line.me/v2/bot/richmenu/" + m.richMenuId); });
  const W = 2500, cw = Math.floor(W / 3);
  const menu = { size: { width: W, height: 843 }, selected: true, name: NAME, chatBarText: "シフト表をひらく",
    areas: [["tc", 0], ["shift", 1], ["wish", 2]].map(a => ({ bounds: { x: a[1] * cw, y: 0, width: a[1] === 2 ? W - 2 * cw : cw, height: 843 }, action: { type: "uri", uri: LIFF_BASE + "?tab=" + a[0] } })) };
  const id = call("post", "https://api.line.me/v2/bot/richmenu", JSON.stringify(menu)).richMenuId;
  const img = UrlFetchApp.fetch(RICHMENU_IMAGE).getBlob();
  call("post", "https://api-data.line.me/v2/bot/richmenu/" + id + "/content", img.getBytes(), "image/png");
  call("post", "https://api.line.me/v2/bot/user/all/richmenu/" + id, "");
  return { richMenuId: id };
}

/* ============================================================
   前の日のお知らせ（毎日12時）：明日のシフト（公開済み）がある人へ1人ずつ
   ============================================================ */
function remindJob() {
  const settings = getSettings_();
  if (settings.remind === false) return;
  const t = new Date(Date.now() + 86400000);
  const month = fmt_(t, "yyyy-MM"), day = String(+fmt_(t, "dd")), label = (+fmt_(t, "M")) + "/" + day + "(" + "日月火水木金土"[t.getDay()] + ")";
  const sh = readAll_("shifts").find(r => r.month === month);
  if (!sh || sh.status !== "published") return;
  const staff = readAll_("staff").filter(s => s.active !== false && s.userId);
  let sent = 0;
  staff.forEach(s => {
    const base = String((sh.cells || {})[day + "_" + s.id] || "").split("|")[0];
    if (!base || base === "OFF") return;
    let st = settings.openStart, en = settings.openEnd;
    if (base.indexOf("C:") === 0) { const p = base.slice(2).split("-"); st = p[0]; en = p[1]; }
    const tags = String((sh.cells || {})[day + "_" + s.id] || "").split("|")[1] || "";
    pushTo_([s.userId], `${s.name}さん、お疲れさまです。\n明日 ${label} は ${dz_(st)}〜${hm2m_(en) <= hm2m_(st) ? "夜中" : ""}${dz_(en)} の出勤です。${tags.indexOf("d") >= 0 ? "（同伴出勤）" : ""}\nよろしくお願いします。\n\n${settings.name}`);
    sent++;
  });
  log_({ userId: "system" }, "remind", month + "/" + day + " " + sent + "人");
}
function dz_(hm) { return String(hm || "").replace(/^0/, ""); }

/* トリガー（毎朝9:05 締切・退勤忘れ・お給料の写し／毎日12時 前日のお知らせ／1時間ごと 勤務時間表）がなければ作る。管理者が開いたときに確認 */
function ensureTriggers_() {
  try {
    const cache = CacheService.getScriptCache();
    if (cache.get("triggers_ok2")) return;
    const have = ScriptApp.getProjectTriggers().map(t => t.getHandlerFunction());
    if (have.indexOf("dailyJob") < 0) ScriptApp.newTrigger("dailyJob").timeBased().everyDays(1).atHour(9).nearMinute(5).create();
    if (have.indexOf("remindJob") < 0) ScriptApp.newTrigger("remindJob").timeBased().everyDays(1).atHour(12).nearMinute(0).create();
    if (have.indexOf("hoursJob") < 0) ScriptApp.newTrigger("hoursJob").timeBased().everyHours(1).create();
    cache.put("triggers_ok2", "1", 21600);
  } catch (e) { Logger.log("ensureTriggers " + e); }
}

/* ============================================================
   店舗端末（iPad）：端末コードで認証。見えるのは打刻に必要な情報だけ
   ============================================================ */
function verifyDevice_(key) {
  const dev = readAll_("devices").find(d => d.key && String(d.key) === String(key).trim().toUpperCase());
  if (!dev) throw new Error("bad_device");
  return dev;
}
function todayStatus_(staff) {
  const now = new Date(), today = bizDate_(), month = today.slice(0, 7), day = String(+today.slice(8));
  const tcs = readAll_("timecards").filter(r => r.month === month);
  const shift = readAll_("shifts").find(r => r.month === month);
  const out = {};
  staff.forEach(s => {
    const t = tcs.find(r => r.staffId === s.id); const rec = t && t.days && t.days[day];
    const cell = shift && shift.status === "published" && shift.cells ? shift.cells[day + "_" + s.id] || "" : "";
    out[s.id] = { in: rec && rec.in ? rec.in.slice(11, 16) : "", out: rec && rec.out ? rec.out.slice(11, 16) : "", brk: rec ? (+rec.brk || 0) : 0, sched: cell };
  });
  return { date: today, status: out, serverTime: fmt_(now, "HH:mm") };
}
function handleDevice_(action, data, device) {
  const settings = getSettings_();
  const staff = readAll_("staff").filter(s => s.active !== false && s.active !== "FALSE");
  try { upsert_("devices", ["id"], Object.assign({}, device, { lastSeen: nowStr_() })); } catch (e) { /* 任意 */ }
  switch (action) {
    case "kioskBootstrap":
    case "kioskStatus": {
      const st = todayStatus_(staff);
      return {
        version: VERSION, device: { id: device.id, name: device.name },
        settings: { name: settings.name, eyebrow: settings.eyebrow, openStart: settings.openStart, openEnd: settings.openEnd, tcUnit: settings.tcUnit, tcMode: settings.tcMode, kioskPin: settings.kioskPin !== false },
        staff: staff.map(s => ({ id: s.id, name: s.name, short: s.short || "", order: +s.order || 99, hasPin: !!s.pin })).sort((a, b) => a.order - b.order),
        today: st
      };
    }
    case "kioskPunch": {
      const s = staff.find(x => x.id === data.staffId); if (!s) throw new Error("bad_staff");
      if (settings.kioskPin !== false) {
        if (!s.pin) throw new Error("pin_not_set");
        if (pinHash_(s.id, String(data.pin || "")) !== s.pin) throw new Error("bad_pin");
      }
      if (!["in", "out", "undoIn", "undoOut"].includes(data.type)) throw new Error("bad_type");
      const r = punch_(s, data, "kiosk:" + device.id);
      log_({ userId: "device:" + device.id }, "kioskPunch:" + data.type, s.name);
      const today = todayStatus_(staff);
      const day = String(+bizDate_().slice(8)); const rec = (r.timecard.days || {})[day];
      if (today.status[s.id]) { today.status[s.id].in = rec && rec.in ? rec.in.slice(11, 16) : ""; today.status[s.id].out = rec && rec.out ? rec.out.slice(11, 16) : ""; today.status[s.id].brk = rec ? (+rec.brk || 0) : 0; }
      return Object.assign(r, { today });
    }
    default: throw new Error("unknown_action");
  }
}
function pinHash_(sid, pin) {
  return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, sid + ":" + pin));
}
/** 打刻の本体（スマホ・店舗端末 共通） */
function punch_(staffRow, data, via) {
  const now = new Date();
  const today = bizDate_(now);   // 夜中1時の退勤も、前の日の勤務として記録する
  const month = today.slice(0, 7);
  const stamp = fmt_(now, "yyyy-MM-dd'T'HH:mm");
  let tcMonth = month, day = String(+today.slice(8));
  if (data.type === "fixOut") { tcMonth = ym_(data.month); day = String(+data.day); }
  const doc = readAll_("timecards").find(r => r.month === tcMonth && r.staffId === staffRow.id) || { month: tcMonth, staffId: staffRow.id, days: {} };
  doc.days = doc.days || {};
  if (data.type === "in") {
    if (doc.days[day] && doc.days[day].in) throw new Error("already_in");
    doc.days[day] = { in: stamp, brk: 0, via };
  } else if (data.type === "out") {
    const r = doc.days[day];
    if (!r || !r.in) throw new Error("not_in");
    r.out = stamp; r.brk = Math.max(0, +data.brk || 0); r.viaOut = via;
  } else if (data.type === "fixOut") {
    const r = doc.days[day];
    if (!r || !r.in) throw new Error("not_in");
    if (!/^\d{2}:\d{2}$/.test(String(data.hm || ""))) throw new Error("bad_time");
    r.out = r.in.slice(0, 10) + "T" + data.hm; r.manual = true; r.edited = true;
    addEdit_(doc, day, staffRow.name + "（本人）", "退勤を手入力 " + data.hm);
  } else if (data.type === "undoIn" || data.type === "undoOut") {
    const r = doc.days[day];
    if (!r || !r.in) throw new Error("not_in");
    const limit = 15 * 60000;
    if (data.type === "undoOut") {
      if (!r.out) throw new Error("not_out");
      if (now - new Date(r.out.replace("T", " ").replace(/-/g, "/")) > limit) throw new Error("undo_expired");
      addEdit_(doc, day, staffRow.name + "（本人・取り消し）", "退勤 " + r.out.slice(11) + " を取り消し");
      delete r.out; delete r.viaOut; r.edited = true;
    } else {
      if (r.out) throw new Error("already_out");
      if (now - new Date(r.in.replace("T", " ").replace(/-/g, "/")) > limit) throw new Error("undo_expired");
      addEdit_(doc, day, staffRow.name + "（本人・取り消し）", "出勤 " + r.in.slice(11) + " を取り消し");
      delete doc.days[day];
    }
  } else throw new Error("bad_type");
  upsert_("timecards", ["month", "staffId"], doc);
  return { timecard: doc, serverTime: stamp };
}

/* ---------- 修正履歴 ---------- */
function addEdit_(doc, day, by, what) {
  doc.edits = (doc.edits || []).concat([{ at: nowStr_().slice(0, 16), day: String(+day), by, what }]).slice(-300);
}
function recText_(r) { if (!r || !r.in) return "記録なし"; return r.in.slice(11, 16) + "〜" + (r.out ? r.out.slice(11, 16) : "未退勤") + (r.brk ? " 休憩" + r.brk + "分" : ""); }
function whoName_(auth, staff) { const s = (staff || []).find(x => x.userId === auth.userId); return s ? s.name : (auth.name || "管理者"); }

/* ---------- 月データ ---------- */
function monthsFor_(months, me, isAdmin) {
  const out = {};
  const shifts = readAll_("shifts"), wishes = readAll_("wishes"), tcs = readAll_("timecards");
  months.forEach(m => {
    const ym = ym_(m);
    const sh = shifts.find(r => r.month === ym) || null;
    const shift = sh ? { month: ym, status: sh.status || "collecting", publishedAt: sh.publishedAt || "", cells: (isAdmin || sh.status === "published") ? (sh.cells || {}) : {} } : null;
    // 希望：全員分を返す（誰が入れるかをスタッフ同士で見られるように）。連絡事項は本人と管理者のみ
    const w = {}; wishes.filter(r => r.month === ym).forEach(r => { const own = me && r.staffId === me.id; w[r.staffId] = (isAdmin || own) ? r : { month: r.month, staffId: r.staffId, days: r.days, updatedAt: r.updatedAt }; });
    const t = {}; tcs.filter(r => r.month === ym && (isAdmin || (me && r.staffId === me.id))).forEach(r => { t[r.staffId] = r; });
    out[ym] = { shift, wishes: w, timecards: t };
  });
  return out;
}

function publicSettings_(s) { const o = Object.assign({}, s); delete o.pin; return o; }
function publicStaff_(s, full) {
  const o = { id: s.id, name: s.name, short: s.short || "", emp: s.emp || "", order: +s.order || 99, active: s.active !== false && s.active !== "FALSE", linked: !!s.userId, hasPin: !!s.pin };
  o.hasRate = (+s.rate || 0) > 0;
  if (full) { o.note = s.note || ""; o.userId = s.userId || ""; o.wage = +s.wage || 0; o.rate = +s.rate || 0; }
  return o;
}
function sanitizeSettings_(s) {
  const o = {};
  if (s.name) o.name = String(s.name).slice(0, 20);
  if (s.eyebrow !== undefined) o.eyebrow = String(s.eyebrow).slice(0, 24);
  if (s.pin) o.pin = String(s.pin).slice(0, 8);
  if (s.deadlineDay) o.deadlineDay = Math.min(28, Math.max(1, +s.deadlineDay));
  if (/^\d{2}:\d{2}$/.test(s.openStart || "")) o.openStart = s.openStart;
  if (/^\d{2}:\d{2}$/.test(s.openEnd || "")) o.openEnd = s.openEnd;
  if (Array.isArray(s.required)) o.required = s.required.slice(0, 7).map(n => Math.max(0, +n || 0));
  if (Array.isArray(s.closed)) o.closed = s.closed.map(n => +n).filter(n => n >= 0 && n <= 6);
  if (s.closedHol !== undefined) o.closedHol = !!s.closedHol;
  if ([15, 30, 60].includes(+s.tcUnit)) o.tcUnit = +s.tcUnit;
  if (["nearest", "floor", "ceil", "favor"].includes(s.tcMode)) o.tcMode = s.tcMode;
  if (s.notifyWish !== undefined) o.notifyWish = !!s.notifyWish;
  if (s.phonePunch !== undefined) o.phonePunch = !!s.phonePunch;
  if (s.kioskPin !== undefined) o.kioskPin = !!s.kioskPin;
  if (s.remind !== undefined) o.remind = !!s.remind;
  if (s.notifyPublish !== undefined) o.notifyPublish = !!s.notifyPublish;
  if (s.notifyDeadline !== undefined) o.notifyDeadline = !!s.notifyDeadline;
  if (s.nightExtra !== undefined) o.nightExtra = Math.min(1, Math.max(0, +s.nightExtra || 0));
  return o;
}

/* ============================================================
   LINE 通知（Messaging API）
   ============================================================ */
function pushTo_(userIds, text) {
  const token = String(PROP.getProperty("MESSAGING_TOKEN") || "").trim();
  const ids = (userIds || []).filter(Boolean);
  if (!token || !ids.length) return;
  for (let i = 0; i < ids.length; i += 500) {
    try {
      UrlFetchApp.fetch("https://api.line.me/v2/bot/message/multicast", {
        method: "post", contentType: "application/json", muteHttpExceptions: true,
        headers: { Authorization: "Bearer " + token },
        payload: JSON.stringify({ to: ids.slice(i, i + 500), messages: [{ type: "text", text }] })
      });
    } catch (e) { /* 通知失敗は本処理を止めない */ }
  }
}

/** 毎朝9時：締切前の未提出者・締切日の提出状況・退勤忘れ を通知（setupTriggers で登録） */
function dailyJob() {
  const settings = getSettings_();
  const staff = readAll_("staff").filter(s => s.active !== false && s.active !== "FALSE");
  const admins = readAll_("admins");
  const now = new Date();
  const today = fmt_(now, "yyyy-MM-dd");
  const day = +today.slice(8);
  const nextYm = fmt_(new Date(now.getFullYear(), now.getMonth() + 1, 1), "yyyy-MM");
  const url = PROP.getProperty("LIFF_URL") || "";
  const dl = +settings.deadlineDay || 20;

  // 希望の締切 3日前と当日（設定でオンのときだけ）
  if (settings.notifyDeadline === true && (day === dl - 3 || day === dl)) {
    const wishes = readAll_("wishes").filter(w => w.month === nextYm);
    const missing = staff.filter(s => !wishes.some(w => w.staffId === s.id));
    pushTo_(missing.map(s => s.userId), `【${settings.name}】${+nextYm.slice(5)}月の希望休の締切は${day === dl ? "本日" : `${dl}日`}です。まだ提出がありません。「希望を出す」から送ってください。${url ? "\n" + url : ""}`);
    if (day === dl && admins.length) {
      pushTo_(admins.map(a => a.userId), `【${settings.name}】${+nextYm.slice(5)}月の希望：提出 ${staff.length - missing.length}／${staff.length}名${missing.length ? "\n未提出：" + missing.map(s => s.name).join("、") : ""}`);
    }
  }

  // 退勤忘れ（昨日）
  const y = new Date(now.getTime() - 86400000);
  const yYm = fmt_(y, "yyyy-MM"), yDay = String(+fmt_(y, "dd"));
  const tcs = readAll_("timecards").filter(r => r.month === yYm);
  const forgot = staff.filter(s => { const t = tcs.find(r => r.staffId === s.id); const rec = t && t.days && t.days[yDay]; return rec && rec.in && !rec.out; });
  forgot.forEach(s => pushTo_([s.userId], `【${settings.name}】昨日（${+fmt_(y, "M")}/${+yDay}）の退勤が押されていません。「タイムカード」を開いて退勤時刻を入れてください。${url ? "\n" + url : ""}`));

  // お給料の見込みを payroll シートへ（経営ボード用）
  payrollJob();
}

/* ============================================================
   セットアップ用（エディタから手で実行）
   ============================================================ */
function setup() {
  Object.keys(SHEETS).forEach(getSheet_);
  migrate_();
  const cur = getSettings_();
  Object.keys(DEFAULT_SETTINGS).forEach(k => { if (cur[k] === undefined) upsert_("settings", ["key"], { key: k, value: DEFAULT_SETTINGS[k] }); });
  UrlFetchApp.fetch("https://api.line.me/v2/profile", { muteHttpExceptions: true }); // 外部接続の権限を先に許可させる
  Logger.log("setup done. properties: " + JSON.stringify({ LIFF_CHANNEL_ID: !!PROP.getProperty("LIFF_CHANNEL_ID"), MESSAGING_TOKEN: !!PROP.getProperty("MESSAGING_TOKEN"), LIFF_URL: PROP.getProperty("LIFF_URL") }));
}
function setupTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => { if (["dailyJob", "remindJob", "hoursJob"].indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger("dailyJob").timeBased().everyDays(1).atHour(9).nearMinute(5).create();
  ScriptApp.newTrigger("remindJob").timeBased().everyDays(1).atHour(12).nearMinute(0).create();
  ScriptApp.newTrigger("hoursJob").timeBased().everyHours(1).create();
  Logger.log("triggers set (dailyJob 09:05 / remindJob 12:00 / hoursJob 毎時 JST)");
}
/** 通知のテスト：管理者全員に1通送る */
function testPush() {
  const admins = readAll_("admins");
  pushTo_(admins.map(a => a.userId), "【テスト】シフト表アプリからの通知です。届いていればOKです。");
  Logger.log("sent to " + admins.length + " admins");
}

/* ============================================================
   スプレッドシート入出力
   ============================================================ */
function ss_() {
  const id = PROP.getProperty("SPREADSHEET_ID");
  if (id) return SpreadsheetApp.openById(id);
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  const created = SpreadsheetApp.create("リノア シフト表データ");
  PROP.setProperty("SPREADSHEET_ID", created.getId());
  return created;
}
function getSheet_(name) {
  const ss = ss_();
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(SHEETS[name]); sh.setFrozenRows(1); sh.getRange(1, 1, sh.getMaxRows(), SHEETS[name].length).setNumberFormat("@"); return sh; }
  const want = SHEETS[name];
  const have = sh.getLastColumn() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] : [];
  if (want.some((h, i) => have[i] !== h)) sh.getRange(1, 1, 1, want.length).setValues([want]);   // 列が増えたとき見出しをそろえる
  return sh;
}
function readAll_(name) {
  const sh = getSheet_(name);
  const vals = sh.getDataRange().getValues();
  if (vals.length < 2) return [];
  const head = vals[0];
  const jc = JSON_COLS[name] || [];
  return vals.slice(1).filter(r => r.some(v => v !== "")).map(r => {
    const o = {};
    head.forEach((h, i) => {
      let v = cellStr_(h, r[i]);
      if (jc.includes(h)) { try { v = v === "" ? null : JSON.parse(v); } catch (e) { v = null; } }
      if (h === "active") v = (v === true || v === "TRUE" || v === "true" || v === "");
      if (h === "order" && v !== "") v = +v;
      o[h] = v;
    });
    return o;
  });
}
function upsert_(name, keys, obj) {
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const sh = getSheet_(name);
    const head = SHEETS[name];
    const jc = JSON_COLS[name] || [];
    const row = head.map(h => { let v = obj[h]; if (v === undefined || v === null) v = ""; if (jc.includes(h)) v = JSON.stringify(v); if (typeof v === "boolean") v = v ? "TRUE" : "FALSE"; return v; });
    const vals = sh.getDataRange().getValues();
    for (let i = 1; i < vals.length; i++) {
      if (keys.every(k => String(cellStr_(k, vals[i][head.indexOf(k)])) === String(obj[k]))) { const rg = sh.getRange(i + 1, 1, 1, head.length); rg.setNumberFormat("@"); rg.setValues([row]); SpreadsheetApp.flush(); return; }
    }
    const rg = sh.getRange(vals.length + 1, 1, 1, head.length); rg.setNumberFormat("@"); rg.setValues([row]);   // 文字列として保存（"2026-10" を日付に変換させない）
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); }
}
function remove_(name, pred) {
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const sh = getSheet_(name);
    const rows = readAll_(name);
    for (let i = rows.length - 1; i >= 0; i--) { if (pred(rows[i])) sh.deleteRow(i + 2); }
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); }
}
function getSettings_() {
  const o = Object.assign({}, DEFAULT_SETTINGS);
  readAll_("settings").forEach(r => { if (r.key && r.value !== null && r.value !== undefined && r.value !== "") o[r.key] = r.value; });
  return o;
}
function log_(auth, action, detail) {
  try { getSheet_("log").appendRow([nowStr_(), auth.userId, action, String(detail || "").slice(0, 200)]); } catch (e) { /* ログ失敗は無視 */ }
}

/* ---------- セル値の正規化：日付に自動変換された月・日時を文字列へ ---------- */
function cellStr_(h, v) {
  if (v instanceof Date) { if (isNaN(v.getTime())) return ""; return h === "month" ? fmt_(v, "yyyy-MM") : fmt_(v, "yyyy-MM-dd'T'HH:mm:ss"); }
  return v;
}
/** 既存データの修復：文字列書式にそろえ、日付化した月を戻し、同じキーの重複行をまとめる（setup から実行） */
function migrate_() {
  Object.keys(SHEETS).forEach(name => {
    const sh = getSheet_(name);
    sh.getRange(1, 1, Math.max(sh.getMaxRows(), 2), SHEETS[name].length).setNumberFormat("@");
    const keys = { settings: ["key"], staff: ["id"], shifts: ["month"], wishes: ["month", "staffId"], timecards: ["month", "staffId"], admins: ["userId"], devices: ["id"], bottles: ["id"], payroll: ["month"], kstyle: ["month"] }[name];
    if (!keys) return;
    const rows = readAll_(name);
    const merged = {}; const order = [];
    rows.forEach(r => {
      const k = keys.map(x => String(r[x])).join("|");
      if (!merged[k]) { merged[k] = r; order.push(k); return; }
      const a = merged[k];
      if (name === "timecards") { a.days = Object.assign({}, a.days || {}, r.days || {}); a.edits = (a.edits || []).concat(r.edits || []); }
      else merged[k] = r;   // それ以外は後の行を採用
    });
    if (order.length === rows.length && rows.every(r => keys.every(x => typeof r[x] === "string"))) { /* 重複なし・型も正常 */ }
    const head = SHEETS[name], jc = JSON_COLS[name] || [];
    const out = order.map(k => head.map(h => { let v = merged[k][h]; if (v === undefined || v === null) v = ""; if (jc.includes(h)) v = JSON.stringify(v); if (typeof v === "boolean") v = v ? "TRUE" : "FALSE"; return v; }));
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, Math.max(head.length, sh.getLastColumn())).clearContent();
    if (out.length) { const rg = sh.getRange(2, 1, out.length, head.length); rg.setNumberFormat("@"); rg.setValues(out); }
    SpreadsheetApp.flush();
  });
}

/* ---------- ユーティリティ ---------- */
function fmt_(d, p) { return Utilities.formatDate(d, TZ, p); }
function nowStr_() { return fmt_(new Date(), "yyyy-MM-dd'T'HH:mm:ss"); }
// 営業日（朝6時までは前の日）
function bizDate_(d) { return fmt_(new Date((d || new Date()).getTime() - DAY_START_HOUR * 3600000), "yyyy-MM-dd"); }
function addDays_(ymd, n) { const p = ymd.split("-").map(Number); return fmt_(new Date(p[0], p[1] - 1, p[2] + n, 12), "yyyy-MM-dd"); }
function isTrue_(v) { return v === true || v === "TRUE" || v === "true"; }
function ym_(s) { if (!/^\d{4}-\d{2}$/.test(String(s || ""))) throw new Error("bad_month"); return String(s); }
