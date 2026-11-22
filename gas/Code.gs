/**
 * @OnlyCurrentDoc
 * ↑ limits the permission to THIS spreadsheet only (not all of your Google Sheets).
 *
 * Animal Peekaboo + Piano Challenge — shared score sheet (Google Apps Script web app)
 *
 * v2 (2026-11-23): added the "game" column so each game has its own ranking.
 *   Requests without "game" are treated as 'peekaboo' (old Peekaboo pages keep working unchanged).
 *   Old rows with an empty game cell also count as 'peekaboo'.
 *   After pasting this version: Deploy → Manage deployments → ✏ Edit → Version: New version → Deploy
 *   (the /exec URL stays the same, so no device needs new settings).
 *
 * Setup (once):
 *   1. Create a new Google Sheet → Extensions → Apps Script → paste this file.
 *   2. Project Settings → Script properties → add  KEY = <any secret word, e.g. 8+ letters>.
 *      The game sends this key with every score. It is entered on each iPad's settings screen,
 *      never written in the game code (the GitHub repo is public).
 *   3. Run "setup" once (▶ button) and allow the permissions → creates the sheets.
 *   4. Deploy → New deployment → Web app
 *        Execute as: Me   /   Who has access: Anyone
 *      Copy the Web app URL (…/exec) and paste it into the game's settings.
 *   5. (Optional) Triggers → add "cleanup", time-driven, once a day → deletes rows older than KEEP_DAYS.
 *
 * Sheets:
 *   plays : 1 row = 1 play (this is the history)
 *   log   : operation log (who sent what, result, ms) — separate from plays on purpose
 */

const KEEP_DAYS = 365;
const LEVELS = ['beginner', 'intermediate', 'advanced'];
const GAMES = ['peekaboo', 'piano'];
const MAX_SCORE = 20000;   // anything above this is treated as a broken/fake request
// "game" is the LAST column on purpose: existing sheets keep their columns, the header cell is added on the fly
const PLAY_HEAD = ['id', 'ts_utc', 'ts_local', 'date', 'device', 'nickname', 'level', 'variant',
                   'score', 'rank_letter', 'cleared', 'max_combo', 'accuracy', 'input', 'user_agent', 'is_test', 'game'];
const game_ = g => (GAMES.indexOf(g) >= 0 ? g : 'peekaboo');
const LOG_HEAD = ['ts_utc', 'ts_local', 'action', 'device', 'target', 'result', 'detail', 'ms'];

function setup() {
  sheet_('plays', PLAY_HEAD);
  sheet_('log', LOG_HEAD);
}

/* ---------- GET: ping / ranking ---------- */
function doGet(e) {
  const t0 = Date.now(), p = (e && e.parameter) || {};
  try {
    // "games" tells new pages that this server understands the game column (v2)
    if (p.action === 'ping') return json_({ ok: true, time: new Date().toISOString(), version: 2, games: GAMES });
    if (p.action === 'ranking') {
      const level = LEVELS.indexOf(p.level) >= 0 ? p.level : 'beginner', game = game_(p.game);
      const r = ranking_(game, level, p.date || today_(), p.test === '1');
      return json_({ ok: true, game, level, date: p.date || today_(), total: r.length, top: r.slice(0, Number(p.limit) || 10) });
    }
    return json_({ ok: false, error: 'unknown action' });
  } catch (err) {
    log_('ranking', p.device, p.level, 'failed', String(err), Date.now() - t0);
    return json_({ ok: false, error: 'server error' });
  }
}

/* ---------- POST: add a play ----------
 * The game sends Content-Type text/plain (a "simple" request) so the browser does not send a
 * CORS preflight, which Apps Script cannot answer. The body is still JSON.
 */
function doPost(e) {
  const t0 = Date.now();
  let d = {};
  try { d = JSON.parse(e.postData.contents || '{}'); } catch (err) { return json_({ ok: false, error: 'bad json' }); }

  const key = PropertiesService.getScriptProperties().getProperty('KEY');
  if (!key || d.key !== key) {
    log_('add', d.device, '', 'failed', 'wrong key', Date.now() - t0);
    return json_({ ok: false, error: 'wrong key' });
  }
  const bad = validate_(d);
  if (bad) {
    log_('add', d.device, d.nickname, 'failed', bad, Date.now() - t0);
    return json_({ ok: false, error: bad });
  }

  const lock = LockService.getScriptLock();   // two iPads may send at the same moment
  lock.waitLock(10000);
  try {
    const sh = sheet_('plays', PLAY_HEAD);
    if (d.id && idExists_(sh, d.id)) {        // resent after a network drop → do not double count
      log_('add', d.device, d.nickname, 'duplicate', d.id, Date.now() - t0);
    } else {
      const now = new Date();
      sh.appendRow([d.id || Utilities.getUuid(), now.toISOString(), local_(now), today_(), clean_(d.device, 20),
        clean_(d.nickname, 10), d.level, clean_(d.variant || '', 10), Math.round(d.score), clean_(d.rankLetter || '', 2),
        Number(d.cleared) || 0, Number(d.maxCombo) || 0, Number(d.accuracy) || 0, clean_(d.input || '', 12),
        clean_(d.userAgent || '', 120), d.isTest ? true : false, game_(d.game)]);
      log_('add', d.device, d.nickname + ' / ' + game_(d.game) + ' / ' + d.level, 'success', 'score ' + d.score, Date.now() - t0);
    }
  } finally { lock.releaseLock(); }

  const r = ranking_(game_(d.game), d.level, today_(), !!d.isTest);
  const me = r.findIndex(x => x.nickname === clean_(d.nickname, 10));
  return json_({ ok: true, game: game_(d.game), rank: me + 1, total: r.length, best: me >= 0 ? r[me].score : d.score, top: r.slice(0, 10) });
}

/* ---------- ranking: best score per nickname, per game, per level, per day ---------- */
function ranking_(game, level, date, includeTest) {
  const sh = sheet_('plays', PLAY_HEAD);
  const rows = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, PLAY_HEAD.length).getValues() : [];
  const c = name => PLAY_HEAD.indexOf(name);
  const best = {};
  rows.forEach(r => {
    if (cell_(r[c('date')], 'yyyy-MM-dd') !== date || r[c('level')] !== level) return;
    if (game_(String(r[c('game')] || '')) !== game) return;   // empty (old rows) → peekaboo
    if (!includeTest && r[c('is_test')] === true) return;
    const n = String(r[c('nickname')]), s = Number(r[c('score')]);
    if (!best[n] || s > best[n].score) best[n] = { nickname: n, score: s, device: r[c('device')], time: cell_(r[c('ts_local')], 'yyyy-MM-dd HH:mm:ss').slice(11, 16) };
  });
  return Object.values(best).sort((a, b) => b.score - a.score);
}

/* ---------- retention ---------- */
function cleanup() {
  const sh = sheet_('plays', PLAY_HEAD), limit = Date.now() - KEEP_DAYS * 864e5;
  for (let i = sh.getLastRow(); i >= 2; i--) {
    if (new Date(sh.getRange(i, 2).getValue()).getTime() < limit) sh.deleteRow(i);
  }
  log_('cleanup', 'server', 'plays', 'success', 'older than ' + KEEP_DAYS + ' days', 0);
}

/* ---------- helpers ---------- */
function validate_(d) {
  const nick = clean_(d.nickname, 10);
  if (!nick) return 'nickname missing';
  if (LEVELS.indexOf(d.level) < 0) return 'unknown level';
  if (d.game && GAMES.indexOf(d.game) < 0) return 'unknown game';
  const s = Number(d.score);
  if (!isFinite(s) || s < 0 || s > MAX_SCORE) return 'score out of range';
  return '';
}
// strip characters that would turn a cell into a formula, trim to length
function clean_(v, max) { return String(v == null ? '' : v).replace(/^[=+\-@]+/, '').replace(/[\r\n\t]/g, ' ').trim().slice(0, max); }
function idExists_(sh, id) { return sh.getLastRow() > 1 && sh.getRange(2, 1, sh.getLastRow() - 1, 1).createTextFinder(id).matchEntireCell(true).findNext() !== null; }
function today_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function local_(d) { return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'); }
// Sheets may turn "2026-11-19" text into a Date cell → normalise both forms back to text
function cell_(v, fmt) { return v instanceof Date ? Utilities.formatDate(v, Session.getScriptTimeZone(), fmt) : String(v); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function sheet_(name, head) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(head); sh.setFrozenRows(1); sh.getRange('D:D').setNumberFormat('@'); }
  // a sheet made by an older version has fewer header cells → add the new ones (e.g. "game")
  else if (sh.getLastColumn() < head.length) sh.getRange(1, sh.getLastColumn() + 1, 1, head.length - sh.getLastColumn()).setValues([head.slice(sh.getLastColumn())]);
  return sh;
}
// logging must never break the real request
function log_(action, device, target, result, detail, ms) {
  try { const n = new Date(); sheet_('log', LOG_HEAD).appendRow([n.toISOString(), local_(n), action, clean_(device, 20), clean_(target, 40), result, clean_(detail, 120), ms]); } catch (e) {}
}
