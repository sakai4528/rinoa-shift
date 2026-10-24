/**
 * カラオケサロン リノア シフト表・タイムカード  バックエンド（Google Apps Script）
 *
 * - データはこのスクリプトに紐づくスプレッドシートに保存します（シートは自動で作られます）
 * - フロント（web/index.html）から POST された JSON を受け取り、LINE のアクセストークンで本人確認してから処理します
 * - シフト公開・締切前・退勤忘れの LINE 通知を Messaging API で送ります
 *
 * 初回セットアップ（README の手順どおり）
 *   1. スクリプトのプロパティに LIFF_CHANNEL_ID / MESSAGING_TOKEN / LIFF_URL を入れる
 *   2. 関数 setup を1回実行（シート作成・権限の許可）
 *   3. 関数 setupTriggers を1回実行（毎朝9時の通知）
 *   4. デプロイ → 新しいデプロイ → ウェブアプリ（実行ユーザー: 自分 / アクセス: 全員）
 */

const TZ = "Asia/Tokyo";
const PROP = PropertiesService.getScriptProperties();

const DEFAULT_SETTINGS = {
  name: "カラオケサロン リノア", eyebrow: "KARAOKE SALON  RINOA", pin: "1234", deadlineDay: 20,
  openStart: "11:00", openEnd: "18:00",
  required: [2, 2, 2, 2, 2, 3, 3],
  closed: [], closedHol: true,
  tcUnit: 30, tcMode: "nearest",
  notifyWish: true
};

const SHEETS = {
  settings:  ["key", "value"],
  staff:     ["id", "name", "short", "emp", "order", "active", "userId", "note", "createdAt"],
  shifts:    ["month", "status", "publishedAt", "cells"],
  wishes:    ["month", "staffId", "days", "memo", "updatedAt"],
  timecards: ["month", "staffId", "days"],
  admins:    ["userId", "name", "addedAt"],
  log:       ["at", "userId", "action", "detail"]
};
const JSON_COLS = { settings: ["value"], shifts: ["cells"], wishes: ["days"], timecards: ["days"] };

/* ============================================================
   HTTP 入口
   ============================================================ */
function doGet() {
  return json_({ ok: true, service: "rinoa-shift", time: nowStr_() });
}

function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: "bad_request" }); }
  try {
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
      return {
        profile: { userId: auth.userId, name: auth.name },
        settings: publicSettings_(settings),
        staff: staff.map(s => publicStaff_(s, isAdmin)),
        me: me ? publicStaff_(me, true) : null,
        isAdmin,
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
        const xs = Object.keys(w.days).filter(d => w.days[d] && w.days[d].t === "x").length;
        pushTo_(admins.map(a => a.userId), `【${settings.name}】${who ? who.name : ""}さんが${+month.slice(5)}月の希望を提出しました（希望休 ${xs}日）。`);
      }
      return { wish: w };
    }

    case "punch": {
      needMe();
      const now = new Date();
      const today = fmt_(now, "yyyy-MM-dd");
      const month = today.slice(0, 7);
      const stamp = fmt_(now, "yyyy-MM-dd'T'HH:mm");
      let tcMonth = month, day = String(+today.slice(8));
      if (data.type === "fixOut") { tcMonth = ym_(data.month); day = String(+data.day); }
      const doc = readAll_("timecards").find(r => r.month === tcMonth && r.staffId === me.id) || { month: tcMonth, staffId: me.id, days: {} };
      doc.days = doc.days || {};
      if (data.type === "in") {
        if (doc.days[day] && doc.days[day].in) throw new Error("already_in");
        doc.days[day] = { in: stamp, brk: 0 };
      } else if (data.type === "out") {
        const r = doc.days[day];
        if (!r || !r.in) throw new Error("not_in");
        r.out = stamp; r.brk = Math.max(0, +data.brk || 0);
      } else if (data.type === "fixOut") {
        const r = doc.days[day];
        if (!r || !r.in) throw new Error("not_in");
        if (!/^\d{2}:\d{2}$/.test(String(data.hm || ""))) throw new Error("bad_time");
        r.out = r.in.slice(0, 10) + "T" + data.hm; r.manual = true;
      } else throw new Error("bad_type");
      upsert_("timecards", ["month", "staffId"], doc);
      log_(auth, "punch:" + data.type, tcMonth + "/" + day);
      return { timecard: doc, serverTime: stamp };
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
      const doc = { month: ym_(data.month), staffId: String(data.staffId), days: data.days || {} };
      upsert_("timecards", ["month", "staffId"], doc);
      log_(auth, "saveTimecard", doc.month + " " + doc.staffId);
      return { timecard: doc };
    }

    default:
      throw new Error("unknown_action");
  }
}

/* ---------- 月データ ---------- */
function monthsFor_(months, me, isAdmin) {
  const out = {};
  const shifts = readAll_("shifts"), wishes = readAll_("wishes"), tcs = readAll_("timecards");
  months.forEach(m => {
    const ym = ym_(m);
    const sh = shifts.find(r => r.month === ym) || null;
    const shift = sh ? { month: ym, status: sh.status || "collecting", publishedAt: sh.publishedAt || "", cells: (isAdmin || sh.status === "published") ? (sh.cells || {}) : {} } : null;
    const w = {}; wishes.filter(r => r.month === ym && (isAdmin || (me && r.staffId === me.id))).forEach(r => { w[r.staffId] = r; });
    const t = {}; tcs.filter(r => r.month === ym && (isAdmin || (me && r.staffId === me.id))).forEach(r => { t[r.staffId] = r; });
    out[ym] = { shift, wishes: w, timecards: t };
  });
  return out;
}

function publicSettings_(s) { const o = Object.assign({}, s); delete o.pin; return o; }
function publicStaff_(s, full) {
  const o = { id: s.id, name: s.name, short: s.short || "", emp: s.emp || "", order: +s.order || 99, active: s.active !== false && s.active !== "FALSE", linked: !!s.userId };
  if (full) { o.note = s.note || ""; o.userId = s.userId || ""; }
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
  return o;
}

/* ============================================================
   LINE 通知（Messaging API）
   ============================================================ */
function pushTo_(userIds, text) {
  const token = PROP.getProperty("MESSAGING_TOKEN");
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

  // 希望の締切 3日前と当日
  if (day === dl - 3 || day === dl) {
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
}

/* ============================================================
   セットアップ用（エディタから手で実行）
   ============================================================ */
function setup() {
  Object.keys(SHEETS).forEach(getSheet_);
  const cur = getSettings_();
  Object.keys(DEFAULT_SETTINGS).forEach(k => { if (cur[k] === undefined) upsert_("settings", ["key"], { key: k, value: DEFAULT_SETTINGS[k] }); });
  UrlFetchApp.fetch("https://api.line.me/v2/profile", { muteHttpExceptions: true }); // 外部接続の権限を先に許可させる
  Logger.log("setup done. properties: " + JSON.stringify({ LIFF_CHANNEL_ID: !!PROP.getProperty("LIFF_CHANNEL_ID"), MESSAGING_TOKEN: !!PROP.getProperty("MESSAGING_TOKEN"), LIFF_URL: PROP.getProperty("LIFF_URL") }));
}
function setupTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === "dailyJob") ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger("dailyJob").timeBased().everyDays(1).atHour(9).nearMinute(5).create();
  Logger.log("daily trigger set (09:05 JST)");
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
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(SHEETS[name]); sh.setFrozenRows(1); }
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
      let v = r[i];
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
      if (keys.every(k => String(vals[i][head.indexOf(k)]) === String(obj[k]))) { sh.getRange(i + 1, 1, 1, head.length).setValues([row]); return; }
    }
    sh.appendRow(row);
  } finally { lock.releaseLock(); }
}
function remove_(name, pred) {
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const sh = getSheet_(name);
    const rows = readAll_(name);
    for (let i = rows.length - 1; i >= 0; i--) { if (pred(rows[i])) sh.deleteRow(i + 2); }
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

/* ---------- ユーティリティ ---------- */
function fmt_(d, p) { return Utilities.formatDate(d, TZ, p); }
function nowStr_() { return fmt_(new Date(), "yyyy-MM-dd'T'HH:mm:ss"); }
function ym_(s) { if (!/^\d{4}-\d{2}$/.test(String(s || ""))) throw new Error("bad_month"); return String(s); }
