'use strict';
// backend/nativeApi.js: the native app's JSON endpoint (key check, allowlist, bootstrap, doGet guard).
const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const root = path.resolve(__dirname, '..');

function makeContext(options) {
  const props = { tempus_native_api_key: 'k'.repeat(64) };
  const calls = [];
  const context = {
    PropertiesService: { getScriptProperties: () => ({ getProperty: (k) => props[k] || null, setProperty: (k, v) => { props[k] = v; } }) },
    ContentService: { MimeType: { JSON: 'json' }, createTextOutput: (text) => ({ text, setMimeType() { return this; } }) },
    Session: { getScriptTimeZone: () => 'Australia/Sydney', getActiveUser: () => ({ getEmail: () => (options && options.email) || '' }) },
    Utilities: { formatDate: () => '2026', getUuid: () => '1234-5678' },
    Logger: { log() {} },
    api_getMigrationStatus: () => ({ required: !!(options && options.migrationRequired) }),
    api_getSettings: () => ({ round_to_nearest: 15, theme: 'dark', custom_themes: '[big]' }),
    api_getHourTypes: () => [{ id: 'work' }],
    api_getContracts: () => [{ id: 'c1' }],
    api_getEntries: (filters) => { calls.push(['getEntries', filters]); return [{ id: 'e1' }]; },
    api_getPublicHolidays: () => [],
    api_deleteEntry: (id) => { calls.push(['deleteEntry', id]); return { success: true }; },
    api_runUpgrade: () => { throw new Error('must not be reachable'); }
  };
  vm.createContext(context);
  vm.runInContext(fs.readFileSync(path.join(root, 'backend/nativeApi.js'), 'utf8'), context, { filename: 'backend/nativeApi.js' });
  return { context, calls, props };
}
function post(context, body) {
  return JSON.parse(context.doPost({ postData: { contents: JSON.stringify(body) } }).text);
}

exports.run = function(test) {
  test('native api rejects a missing or wrong key', () => {
    const { context, calls } = makeContext();
    assert.strictEqual(post(context, { action: 'getEntries' }).error, 'unauthorised');
    assert.strictEqual(post(context, { key: 'k'.repeat(63) + 'x', action: 'getEntries' }).error, 'unauthorised');
    assert.strictEqual(calls.length, 0);
  });

  test('native api runs allowed actions with their arguments and refuses the rest', () => {
    const { context, calls } = makeContext();
    const key = 'k'.repeat(64);
    const answer = post(context, { key, action: 'deleteEntry', args: ['e1'] });
    assert.strictEqual(answer.ok, true);
    assert.deepStrictEqual(calls, [['deleteEntry', 'e1']]);
    const refused = post(context, { key, action: 'runUpgrade' });
    assert.strictEqual(refused.ok, false);
    assert.ok(/Unknown action/.test(refused.error));
  });

  test('native api bootstrap returns the app data with only the settings it needs', () => {
    const { context } = makeContext();
    const answer = post(context, { key: 'k'.repeat(64), action: 'bootstrap' });
    assert.strictEqual(answer.ok, true);
    assert.deepStrictEqual(answer.result.settings, { round_to_nearest: 15 });
    assert.deepStrictEqual(answer.result.entries, [{ id: 'e1' }]);
    assert.strictEqual(answer.result.timeZone, 'Australia/Sydney');
  });

  test('native api refuses work while the sheet needs its upgrade', () => {
    const { context } = makeContext({ migrationRequired: true });
    const answer = post(context, { key: 'k'.repeat(64), action: 'bootstrap' });
    assert.strictEqual(answer.ok, false);
    assert.ok(/upgrade/.test(answer.error));
  });

  test('native api treats a visitor without an email as anonymous', () => {
    assert.strictEqual(makeContext().context.nativeApiIsAnonymousVisitor_(), true);
    assert.strictEqual(makeContext({ email: 'someone@example.com' }).context.nativeApiIsAnonymousVisitor_(), false);
  });
};
