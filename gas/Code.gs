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
 * v3 (2026-11-25): Instagram winners.
 *   - 2 new columns at the end of "plays": insta, insta_ok (filled later by action "insta").
 *   - POST action "winners" (key needed): today's #1 per game × level, with the Instagram ID
 *     ONLY of players who agreed. The public GET ranking never returns Instagram IDs.
 *   - POST action "posted" (key needed): staff mark a card as posted → "posts" sheet.
 *   Deploy the same way as v2 (Edit → New version), the URL stays the same.
 *
 * v4 (2026-11-25): played time + nickname check.
 *   - The game sends the time it was PLAYED ("ts"). ts_utc / ts_local / date now come from that time,
 *     so a score sent later (after the internet came back) still counts for the day it was played.
 *     An impossible time (more than 10 min in the future, or older than MAX_AGE_DAYS) is NOT used:
 *     the time the sheet received it is used instead, and the log says "bad device time".
 *   - New last column "received_utc" = when the sheet got the row.
 *   - GET action "nick" (game, nickname): is this nickname already used today in this game? (yes/no only)
 *   - POST action "unconsent" (key needed, staff page): a guest asked "do not post me" →
 *     clears insta / insta_ok on EVERY row with that Instagram ID, and replaces the ID in "posts" with "(removed)".
 *     The log keeps who / when / how many rows — never the ID itself.
 *   Deploy the same way (Edit → New version), the URL stays the same.
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
 *   posts : 1 row = 1 Instagram post made by staff (used for "NEW" and to avoid double posts)
 */

const KEEP_DAYS = 365;
const LEVELS = ['beginner', 'intermediate', 'advanced'];
const GAMES = ['peekaboo', 'piano'];
const MAX_SCORE = 20000;   // anything above this is treated as a broken/fake request
const MAX_AGE_DAYS = 7;    // a score waiting offline longer than this is dated by the time it arrives
const FUTURE_MS = 10 * 60 * 1000;   // device clocks can be a few minutes off; more than 10 min ahead = wrong clock
// "game" is the LAST column on purpose: existing sheets keep their columns, the header cell is added on the fly
const PLAY_HEAD = ['id', 'ts_utc', 'ts_local', 'date', 'device', 'nickname', 'level', 'variant',
                   'score', 'rank_letter', 'cleared', 'max_combo', 'accuracy', 'input', 'user_agent', 'is_test', 'game',
                   'insta', 'insta_ok', 'received_utc'];
const game_ = g => (GAMES.indexOf(g) >= 0 ? g : 'peekaboo');
const LOG_HEAD = ['ts_utc', 'ts_local', 'action', 'device', 'target', 'result', 'detail', 'ms'];
const POST_HEAD = ['ts_utc', 'ts_local', 'date', 'game', 'level', 'nickname', 'score', 'insta', 'play_id', 'staff'];

function setup() {
  sheet_('plays', PLAY_HEAD);
  sheet_('log', LOG_HEAD);
  sheet_('posts', POST_HEAD);
}

/* ---------- GET: ping / ranking ---------- */
function doGet(e) {
  const t0 = Date.now(), p = (e && e.parameter) || {};
  try {
    // "games" tells new pages that this server understands the game column (v2)
    if (p.action === 'ping') return json_({ ok: true, time: new Date().toISOString(), version: 4, games: GAMES });
    // nickname check: answers only yes/no (no scores, no devices)
    if (p.action === 'nick') {
      const n = clean_(p.nickname, 10).toLowerCase(), game = game_(p.game), date = today_(), c = name => PLAY_HEAD.indexOf(name);
      const exists = !!n && playRows_().some(r => cell_(r[c('date')], 'yyyy-MM-dd') === date && r[c('is_test')] !== true &&
        game_(String(r[c('game')] || '')) === game && String(r[c('nickname')]).toLowerCase() === n);
      return json_({ ok: true, exists });
    }
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
  // staff / follow-up actions (all need the key, checked above)
  try {
    if (d.action === 'insta') return setInsta_(d, t0);
    if (d.action === 'winners') { const w = winners_(); log_('winners', d.device, 'staff', 'success', '', Date.now() - t0); return json_(w); }
    if (d.action === 'posted') return posted_(d, t0);
    if (d.action === 'unconsent') return unconsent_(d, t0);
  } catch (err) {
    log_(d.action, d.device, '', 'failed', String(err), Date.now() - t0);
    return json_({ ok: false, error: 'server error' });
  }
  const bad = validate_(d);
  if (bad) {
    log_('add', d.device, d.nickname, 'failed', bad, Date.now() - t0);
    return json_({ ok: false, error: bad });
  }

  const played = playedAt_(d.ts);             // the time the game was played (checked)
  const lock = LockService.getScriptLock();   // two iPads may send at the same moment
  lock.waitLock(10000);
  try {
    const sh = sheet_('plays', PLAY_HEAD);
    if (d.id && idExists_(sh, d.id)) {        // resent after a network drop → do not double count
      log_('add', d.device, d.nickname, 'duplicate', d.id, Date.now() - t0);
    } else {
      const now = new Date(), at = played.at;
      if (!played.ok) log_('add', d.device, d.nickname, 'warning', 'bad device time: ' + clean_(d.ts, 40) + ' → used received time', 0);
      sh.appendRow([d.id || Utilities.getUuid(), at.toISOString(), local_(at), day_(at), clean_(d.device, 20),
        clean_(d.nickname, 10), d.level, clean_(d.variant || '', 10), Math.round(d.score), clean_(d.rankLetter || '', 2),
        Number(d.cleared) || 0, Number(d.maxCombo) || 0, Number(d.accuracy) || 0, clean_(d.input || '', 12),
        clean_(d.userAgent || '', 120), d.isTest ? true : false, game_(d.game), '', '', now.toISOString()]);
      log_('add', d.device, d.nickname + ' / ' + game_(d.game) + ' / ' + d.level, 'success', 'score ' + d.score, Date.now() - t0);
    }
  } finally { lock.releaseLock(); }

  // "You are #N" is for the day the game was played
  const r = ranking_(game_(d.game), d.level, day_(played.at), !!d.isTest);
  const me = r.findIndex(x => x.nickname === clean_(d.nickname, 10));
  return json_({ ok: true, game: game_(d.game), rank: me + 1, total: r.length, best: me >= 0 ? r[me].score : d.score, top: r.slice(0, 10) });
}

/* ---------- ranking: best score per nickname, per game, per level, per day ---------- */
// Public: only nickname / score / device / time. Never add Instagram fields here (anyone can call GET).
function ranking_(game, level, date, includeTest) {
  return bestRows_(playRows_(), game, level, date, includeTest)
    .map(b => ({ nickname: b.nickname, score: b.score, device: b.device, time: b.time }));
}
function playRows_() {
  const sh = sheet_('plays', PLAY_HEAD);
  return sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, PLAY_HEAD.length).getValues() : [];
}
// best row per nickname (internal: also keeps the row's id and Instagram consent)
function bestRows_(rows, game, level, date, includeTest) {
  const c = name => PLAY_HEAD.indexOf(name);
  const best = {};
  rows.forEach(r => {
    if (cell_(r[c('date')], 'yyyy-MM-dd') !== date || r[c('level')] !== level) return;
    if (game_(String(r[c('game')] || '')) !== game) return;   // empty (old rows) → peekaboo
    if (!includeTest && r[c('is_test')] === true) return;
    const n = String(r[c('nickname')]), s = Number(r[c('score')]);
    if (!best[n] || s > best[n].score) best[n] = { nickname: n, score: s, device: r[c('device')],
      time: cell_(r[c('ts_local')], 'yyyy-MM-dd HH:mm:ss').slice(11, 16), id: String(r[c('id')]),
      variant: String(r[c('variant')] || ''), accuracy: Number(r[c('accuracy')]) || 0,
      insta: r[c('insta_ok')] === true ? String(r[c('insta')] || '') : '' };
  });
  return Object.values(best).sort((a, b) => b.score - a.score);
}

/* ---------- Instagram (v3) ---------- */
// The player agreed on the result screen → store the ID on THAT play row. No agreement → store nothing.
function setInsta_(d, t0) {
  const insta = String(d.insta || '').replace(/^@+/, '').trim();
  if (!d.id) return json_({ ok: false, error: 'id missing' });
  if (d.agree !== true || d.adult !== true) return json_({ ok: false, error: 'no consent' });
  if (!/^[A-Za-z0-9._]{1,30}$/.test(insta)) return json_({ ok: false, error: 'bad instagram id' });
  const sh = sheet_('plays', PLAY_HEAD);
  const cell = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).createTextFinder(d.id).matchEntireCell(true).findNext() : null;
  if (!cell) { log_('insta', d.device, d.id, 'failed', 'play not found', Date.now() - t0); return json_({ ok: false, error: 'play not found' }); }
  sh.getRange(cell.getRow(), PLAY_HEAD.indexOf('insta') + 1, 1, 2).setValues([[insta, true]]);
  log_('insta', d.device, d.id, 'success', 'consent saved', Date.now() - t0);   // the ID itself is not logged
  return json_({ ok: true });
}
// Today's #1 per game × level. "pick" = best player who agreed (may differ from #1).
function winners_() {
  const rows = playRows_(), date = today_(), last = lastPosts_(date), out = [];
  GAMES.forEach(game => LEVELS.forEach(level => {
    const list = bestRows_(rows, game, level, date, false);
    const pick = list.find(b => b.insta) || null, lp = last[game + '|' + level] || null;
    out.push({ game, level, players: list.length,
      top: list[0] ? { nickname: list[0].nickname, score: list[0].score, agreed: !!list[0].insta } : null,
      pick: pick ? { nickname: pick.nickname, score: pick.score, insta: pick.insta, id: pick.id, variant: pick.variant,
                     accuracy: pick.accuracy, time: pick.time, rank: list.indexOf(pick) + 1 } : null,
      lastPosted: lp, isNew: !!pick && (!lp || lp.play_id !== pick.id) });
  }));
  return { ok: true, date, now: local_(new Date()).slice(11, 16), winners: out };
}
function lastPosts_(date) {
  const sh = sheet_('posts', POST_HEAD), m = {};
  if (sh.getLastRow() < 2) return m;
  sh.getRange(2, 1, sh.getLastRow() - 1, POST_HEAD.length).getValues().forEach(r => {
    if (cell_(r[2], 'yyyy-MM-dd') !== date) return;
    m[r[3] + '|' + r[4]] = { time: cell_(r[1], 'yyyy-MM-dd HH:mm:ss').slice(11, 16), nickname: String(r[5]), score: Number(r[6]), play_id: String(r[8]) };
  });
  return m;
}
function posted_(d, t0) {
  if (GAMES.indexOf(d.game) < 0 || LEVELS.indexOf(d.level) < 0 || !d.playId) return json_({ ok: false, error: 'bad request' });
  const n = new Date();
  sheet_('posts', POST_HEAD).appendRow([n.toISOString(), local_(n), today_(), d.game, d.level, clean_(d.nickname, 10),
    Number(d.score) || 0, clean_(d.insta, 30), clean_(d.playId, 60), clean_(d.staff || '', 20)]);
  log_('posted', d.device, d.game + ' / ' + d.level + ' / ' + clean_(d.nickname, 10), 'success', 'by ' + clean_(d.staff || '-', 20), Date.now() - t0);
  return json_({ ok: true });
}

// Guest withdrew consent: remove the Instagram ID everywhere (all days, all games), so the card never comes back.
function unconsent_(d, t0) {
  const sh = sheet_('plays', PLAY_HEAD), ci = PLAY_HEAD.indexOf('insta'), cid = PLAY_HEAD.indexOf('id');
  if (sh.getLastRow() < 2 || !d.playId) return json_({ ok: false, error: 'bad request' });
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  let rows = 0, posts = 0;
  try {
    const vals = sh.getRange(2, 1, sh.getLastRow() - 1, PLAY_HEAD.length).getValues();
    const hit = vals.find(r => String(r[cid]) === String(d.playId));
    if (!hit) { log_('unconsent', d.device, d.playId, 'failed', 'play not found', Date.now() - t0); return json_({ ok: false, error: 'play not found' }); }
    const insta = String(hit[ci] || '').toLowerCase();
    vals.forEach((r, i) => {
      const same = String(r[cid]) === String(d.playId) || (insta && String(r[ci] || '').toLowerCase() === insta);
      if (same && (r[ci] !== '' || r[ci + 1] === true)) { sh.getRange(i + 2, ci + 1, 1, 2).setValues([['', false]]); rows++; }
    });
    const ps = sheet_('posts', POST_HEAD), pi = POST_HEAD.indexOf('insta');
    if (insta && ps.getLastRow() > 1) {
      const pv = ps.getRange(2, pi + 1, ps.getLastRow() - 1, 1).getValues();
      pv.forEach((r, i) => { if (String(r[0]).toLowerCase() === insta) { ps.getRange(i + 2, pi + 1).setValue('(removed)'); posts++; } });
    }
  } finally { lock.releaseLock(); }
  log_('unconsent', d.device, clean_(d.game, 10) + ' / ' + clean_(d.level, 12) + ' / ' + clean_(d.nickname, 10), 'success',
       `rows ${rows}, posts ${posts}, by ${clean_(d.staff || '-', 20)}`, Date.now() - t0);   // the Instagram ID itself is not logged
  return json_({ ok: true, rows, posts });
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
function today_() { return day_(new Date()); }
function day_(d) { return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
// time sent by the game → used only when it is possible (not in the future, not too old)
function playedAt_(ts) {
  const now = new Date(), t = ts ? new Date(ts) : null;
  if (t && isFinite(t.getTime()) && t.getTime() <= now.getTime() + FUTURE_MS && t.getTime() >= now.getTime() - MAX_AGE_DAYS * 864e5) return { at: t, ok: true };
  return { at: now, ok: false };
}
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
