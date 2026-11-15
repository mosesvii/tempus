/**
 * JSON endpoint for the native Tempus app (Mac now, iPhone later).
 *
 * The web page is slow because every open builds ~2.4 MB of HTML and then makes one
 * google.script.run call per slice of data. The native app instead keeps its own copy of the data
 * and talks to this endpoint: one POST per action, JSON in and JSON out, running the same api_*
 * functions as the web page so every rule (rates, rounding, duplicates, assessments) stays here.
 *
 * Deploy it as its own web app deployment: Execute as "Me", Who has access "Anyone". The app has no
 * Google sign-in, so the key below is the lock: run setupNativeApiKey() once from the editor, copy
 * the key it logs into the app's Settings, and keep it private. Running it again makes a new key
 * (the old one stops working). doGet refuses to serve the web page to anonymous visitors, so that
 * deployment's URL never shows the full Tempus page to someone who isn't signed in.
 *
 * Request:  POST { "key": "...", "action": "bootstrap" | "<allowed api name>", "args": [ ... ] }
 * Response: { "ok": true, "result": ..., "serverTime": "<ISO>" } or { "ok": false, "error": "..." }
 */
var NATIVE_API_KEY_PROPERTY = 'tempus_native_api_key';
var NATIVE_API_VERSION = 1;

// Only these api_* functions can be called from the app. Add to it as the app grows.
var NATIVE_API_ALLOWED = {
  getSettings: true,
  getHourTypes: true,
  getContracts: true,
  getEntries: true,
  getPublicHolidays: true,
  getFeatureFlags: true,
  getAnnualSummary: true,
  addEntry: true,
  updateEntry: true,
  deleteEntry: true,
  saveDaySessions: true
};

function doPost(e) {
  var answer;
  try {
    var body = nativeApiParseBody_(e);
    if (!nativeApiKeyMatches_(body.key)) {
      answer = { ok: false, error: 'unauthorised' };
    } else {
      answer = { ok: true, result: nativeApiDispatch_(String(body.action || ''), body.args) };
    }
  } catch (err) {
    answer = { ok: false, error: String(err && err.message || err).slice(0, 500) };
  }
  answer.serverTime = new Date().toISOString();
  answer.apiVersion = NATIVE_API_VERSION;
  return ContentService.createTextOutput(JSON.stringify(answer)).setMimeType(ContentService.MimeType.JSON);
}

function nativeApiParseBody_(e) {
  var raw = e && e.postData && e.postData.contents ? e.postData.contents : '';
  if (!raw) throw new Error('Empty request.');
  var body = JSON.parse(raw);
  if (!body || typeof body !== 'object') throw new Error('Invalid request.');
  return body;
}

function nativeApiKeyMatches_(given) {
  var expected = PropertiesService.getScriptProperties().getProperty(NATIVE_API_KEY_PROPERTY);
  if (!expected || typeof given !== 'string' || given.length !== expected.length) return false;
  // Compare every character so the time taken doesn't hint at how much of the key was right.
  var diff = 0;
  for (var i = 0; i < expected.length; i++) diff |= expected.charCodeAt(i) ^ given.charCodeAt(i);
  return diff === 0;
}

function nativeApiDispatch_(action, args) {
  // Same gate as doGet: no reads or writes while the sheet still needs its upgrade.
  var migration = api_getMigrationStatus();
  if (migration && migration.required) throw new Error('Tempus needs its upgrade. Open Tempus in the browser once.');
  if (action === 'bootstrap') return nativeApiBootstrap_();
  if (!NATIVE_API_ALLOWED.hasOwnProperty(action)) throw new Error('Unknown action: ' + action);
  var scope = typeof globalThis !== 'undefined' ? globalThis : this;
  var fn = scope['api_' + action];
  if (typeof fn !== 'function') throw new Error('Missing function: api_' + action);
  return fn.apply(null, Array.isArray(args) ? args : []);
}

/** Everything the app needs to draw the calendar, agenda and annual view, in one round trip. */
function nativeApiBootstrap_() {
  var year = Number(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy'));
  var settings = api_getSettings();
  return {
    settings: nativeApiPickSettings_(settings),
    hourTypes: api_getHourTypes(),
    contracts: api_getContracts(),
    entries: api_getEntries({}),
    publicHolidays: api_getPublicHolidays({ year: year }),
    timeZone: Session.getScriptTimeZone()
  };
}

// Only the settings the app uses: the full set holds theme blobs and other things it doesn't need.
function nativeApiPickSettings_(settings) {
  var keep = ['round_to_nearest', 'default_time_view', 'public_holiday_state', 'assessment_time_hour_type_id',
    'week_start', 'week_starts_on', 'financial_year_start_month'];
  var out = {};
  keep.forEach(function(key) {
    if (settings && settings.hasOwnProperty(key)) out[key] = settings[key];
  });
  return out;
}

/** Run once from the Apps Script editor (select it, then Run). The new key appears in the execution log. */
function setupNativeApiKey() {
  var key = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperty(NATIVE_API_KEY_PROPERTY, key);
  Logger.log('Tempus app key (paste it into the app\'s Settings): ' + key);
  return key;
}

/**
 * True when nobody is signed in: the native deployment runs as the owner for anyone with its URL,
 * so doGet must not hand that visitor the web page. The normal deployment runs as the signed-in
 * visitor, who always has an email here.
 */
function nativeApiIsAnonymousVisitor_() {
  try {
    return !Session.getActiveUser().getEmail();
  } catch (e) {
    return true;
  }
}
