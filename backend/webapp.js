/** Web entry + HTML includes */
function doGet(e) {
  // The native app's deployment runs as the owner for anyone with its URL (see backend/nativeApi.js).
  // Never serve the web page there: someone who isn't signed in gets a blank notice instead.
  if (nativeApiIsAnonymousVisitor_()) {
    return HtmlService.createHtmlOutput('<p style="font-family:system-ui;padding:24px">This address is for the Tempus app.</p>').setTitle('Tempus');
  }
  // Financial migrations are explicit: the user sees what will happen and a verified Drive backup
  // is created before any sheet is changed. The application cannot race a partial upgrade.
  var migrationStatus = api_getMigrationStatus();
  if (migrationStatus.required) return buildPreUpgradePage_();
  var view = e && e.parameter && e.parameter.view ? String(e.parameter.view).toLowerCase() : '';
  var templateName = view === 'mobile' ? 'views/mobile' : 'views/index';
  var tpl = HtmlService.createTemplateFromFile(templateName);
  // Read settings once for first-paint injection (the time view + the mobile-gate theme colours below).
  var bootSettings = getBootSettings();
  // Inject the user's saved time view (Agenda/Calendar) so the page paints directly on their preference
  // on any browser — no default-then-switch flash. Always defined so the template scriptlet can't throw.
  tpl.initialTimeView = pickInitialTimeView(bootSettings);
  var bootFlags = api_getFeatureFlags();
  var lilMode = !!(bootFlags.enable_lil_assessments_mode && bootFlags.enable_lil_assessments_mode.enabled);
  var requestedPage = e && e.parameter ? String(e.parameter.page || '') : '';
  tpl.initialPage = pickInitialPage_(lilMode, requestedPage);
  tpl.lilModeJson = JSON.stringify(lilMode);
  tpl.explicitPageJson = JSON.stringify(!!requestedPage);
  // The mobile redirect gate paints before the app boots, so it can't read the theme from CSS. Resolve
  // the active theme's core colours server-side here and hand them to the page (see views/index.html +
  // views/partials/mobile-detect) so the splash renders in the user's theme on the first frame.
  tpl.gateThemeJson = JSON.stringify(getGateThemeColours(bootSettings));
  // Hand the page its own /exec URL so the gate's "Open" button is ready on first paint (no spinner /
  // round-trip). The gate falls back to the async api_getWebAppUrl lookup if this is ever empty.
  tpl.webAppUrl = api_getWebAppUrl();
  // First-paint theming for the mobile view: the body's theme class (built-ins get their full CSS from
  // the head partial; custom resolves to ts-theme-custom) plus, for custom themes, a small inline block
  // so the dominant colours are right on frame one. The client fills in the complete token set after.
  tpl.bodyThemeClass = getActiveThemeBodyClass(bootSettings);
  tpl.customThemeFirstPaintCss = getCustomThemeFirstPaintCss(bootSettings);
  var html = tpl.evaluate()
    // Apps Script ignores <meta name="viewport"> placed inside the HTML (it only
    // affects the inner sandbox iframe). addMetaTag injects it into the outer
    // hosting page, which is what actually controls mobile scaling — without this
    // phones render the app at desktop width and zoom out (tiny, big margins).
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setFaviconUrl('https://raw.githubusercontent.com/brandonhinds/tempus/refs/heads/main/images/favicon.ico');
  return html;
}

function buildPreUpgradePage_() {
  var template = HtmlService.createTemplateFromFile('views/upgrade');
  template.webAppUrlJson = JSON.stringify(api_getWebAppUrl());
  return template.evaluate()
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setTitle('Upgrade Tempus');
}

function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

/**
 * Minimal page served while another execution is running the schema upgrade (see doGet). Styled on
 * the dark-brand fallback (GATE_THEME_FALLBACK) — settings can't be read yet, so no theme lookup.
 * Auto-retry is best-effort only: the Apps Script sandbox blocks programmatic top-frame navigation
 * outside a user gesture, so a real link is always rendered too (same approach as mobile-detect).
 */
function buildUpgradeSplashPage_() {
  var url = api_getWebAppUrl();
  var autoRetry = url
    ? '<script>setTimeout(function(){' +
      'try{window.top.location.href=' + JSON.stringify(url) + ';}' +
      'catch(e){try{window.location.reload();}catch(e2){}}' +
      '},5000);</script>'
    : '';
  var cta = url
    ? '<a target="_top" href="' + url + '" style="display:inline-block;margin-top:24px;background:#136f63;color:#ffffff;' +
      'text-decoration:none;font-weight:700;padding:13px 32px;border-radius:12px;">Continue</a>'
    : '<p style="margin-top:24px;color:#8b92a0;">Reload this page in a moment.</p>';
  var html =
    '<!DOCTYPE html><html><head><base target="_top"><title>Tempus</title></head>' +
    '<body style="margin:0;display:flex;flex-direction:column;align-items:center;justify-content:center;' +
    'min-height:100vh;background:#05060a;color:#e4e6eb;font-family:system-ui,-apple-system,sans-serif;text-align:center;padding:28px;box-sizing:border-box;">' +
    '<div style="width:40px;height:40px;border-radius:50%;border:4px solid #1f2634;border-top-color:#136f63;animation:ts-up-spin 1s linear infinite;"></div>' +
    '<style>@keyframes ts-up-spin{to{transform:rotate(360deg)}}</style>' +
    '<h1 style="margin:20px 0 8px;font-size:22px;font-weight:800;">Upgrading your spreadsheet…</h1>' +
    '<p style="margin:0;color:#8b92a0;">Tempus is applying a one-time update. This usually takes a few seconds.</p>' +
    cta +
    autoRetry +
    '</body></html>';
  return HtmlService.createHtmlOutput(html);
}

// One settings read for everything doGet injects on first paint (avoids hitting the sheet twice).
function getBootSettings() {
  try { return api_getSettings ? api_getSettings() : null; } catch (e) { return null; }
}

// The account-synced time-view preference, read server-side for first-paint injection (see doGet).
// Returns '' when the setting is unset (first-ever load, or an existing user who never saved it) so the
// client falls back to its localStorage hint and then to the Agenda default — this avoids flipping an
// existing Calendar user to Agenda just because their preference hadn't been persisted to the account yet.
function pickInitialTimeView(settings) {
  try {
    var v = settings && settings.default_time_view ? String(settings.default_time_view).toLowerCase() : '';
    return (v === 'calendar' || v === 'agenda') ? v : '';
  } catch (e) {
    return '';
  }
}

// ---- Mobile redirect-gate theme resolution ----------------------------------------------------
// The gate (views/partials/mobile-detect) renders before the app boots, so it can't read the theme
// from CSS or know which theme is active (there's no client-side cache). We resolve the active theme's
// core colours here and inject them. Custom themes store these raw in settings; built-in themes are read
// straight from the theme's CSS block (single source of truth — no duplicated colour list to maintain).
var GATE_THEME_FALLBACK = {
  bg: '#05060a', panel: '#0e1424', text: '#e4e6eb',
  muted: '#8b92a0', primary: '#136f63', primaryText: '#ffffff'
};

function getGateThemeColours(settings) {
  try {
    var theme = settings && settings.theme ? String(settings.theme) : 'dark';
    var base;
    if (theme === 'custom') {
      base = readLegacyCustomGateBase(settings);           // pre-migration single custom config
    } else if (theme.indexOf('custom:') === 0) {
      base = readCustomThemeGateBase(settings, theme.slice(7));
    } else {
      base = readBuiltinThemeGateBase(theme);
    }
    if (!base) return GATE_THEME_FALLBACK;
    var out = {
      bg: sanitizeGateHex(base.bg) || GATE_THEME_FALLBACK.bg,
      panel: sanitizeGateHex(base.panel) || GATE_THEME_FALLBACK.panel,
      text: sanitizeGateHex(base.text) || GATE_THEME_FALLBACK.text,
      muted: sanitizeGateHex(base.muted) || GATE_THEME_FALLBACK.muted,
      primary: sanitizeGateHex(base.primary) || GATE_THEME_FALLBACK.primary
    };
    out.primaryText = gatePrimaryText(out.primary);
    return out;
  } catch (e) {
    return GATE_THEME_FALLBACK;
  }
}

function readCustomThemeGateBase(settings, id) {
  try {
    var raw = settings ? settings.custom_themes : null;
    var arr = typeof raw === 'string' ? JSON.parse(raw) : raw;
    if (!Array.isArray(arr)) return null;
    for (var i = 0; i < arr.length; i++) {
      if (arr[i] && arr[i].id === id && arr[i].config) {
        var c = arr[i].config;
        return { bg: c.bg, panel: c.panel, text: c.text, muted: c.muted, primary: c.primary };
      }
    }
    return null;
  } catch (e) { return null; }
}

function readLegacyCustomGateBase(settings) {
  try {
    var raw = settings ? settings.custom_theme_config : null;
    var c = typeof raw === 'string' ? JSON.parse(raw) : raw;
    if (!c) return null;
    return { bg: c.bg, panel: c.panel, text: c.text, muted: c.muted, primary: c.primary };
  } catch (e) { return null; }
}

function readBuiltinThemeGateBase(key) {
  try {
    var safeKey = String(key || '').replace(/[^a-z0-9-]/gi, '');
    if (!safeKey) return null;
    // Read the theme's own CSS block from the head partial (the source of truth for built-in colours).
    var css = HtmlService.createHtmlOutputFromFile('views/partials/head').getContent();
    var re = new RegExp('\\.ts-theme-' + safeKey + '\\s*\\{([\\s\\S]*?)\\}');
    var m = re.exec(css);
    if (!m) return null;
    var block = m[1];
    return {
      bg: extractCssVar(block, '--bg'),
      panel: extractCssVar(block, '--panel'),
      text: extractCssVar(block, '--text'),
      muted: extractCssVar(block, '--muted'),
      primary: extractCssVar(block, '--primary')
    };
  } catch (e) { return null; }
}

function extractCssVar(block, name) {
  try {
    var re = new RegExp(name + '\\s*:\\s*([^;]+);');
    var m = re.exec(block);
    return m ? m[1].trim() : '';
  } catch (e) { return ''; }
}

// Only hex passes through to the page (values are injected raw for the gate), so a tampered config can't
// inject script; anything non-hex just drops to the fallback for that token.
function sanitizeGateHex(v) {
  if (!v) return '';
  v = String(v).trim();
  return /^#([0-9a-fA-F]{3}|[0-9a-fA-F]{6}|[0-9a-fA-F]{8})$/.test(v) ? v : '';
}

// Readable text colour for the primary button/icon tile (white on dark primaries, near-black on light).
function gatePrimaryText(hex) {
  var rgb = gateHexToRgb(hex);
  if (!rgb) return '#ffffff';
  var lin = function (c) { c /= 255; return c <= 0.03928 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4); };
  var lum = 0.2126 * lin(rgb.r) + 0.7152 * lin(rgb.g) + 0.0722 * lin(rgb.b);
  return lum > 0.5 ? '#1a1a1a' : '#ffffff';
}

function gateHexToRgb(hex) {
  try {
    var h = String(hex || '').replace('#', '');
    if (h.length === 3) h = h[0] + h[0] + h[1] + h[1] + h[2] + h[2];
    if (h.length < 6) return null;
    return { r: parseInt(h.substr(0, 2), 16), g: parseInt(h.substr(2, 2), 16), b: parseInt(h.substr(4, 2), 16) };
  } catch (e) { return null; }
}

function gateRgbToHex(r, g, b) {
  var to = function (n) { n = Math.max(0, Math.min(255, Math.round(n))).toString(16); return n.length === 1 ? '0' + n : n; };
  return '#' + to(r) + to(g) + to(b);
}

function gateMix(hexA, hexB, t) {
  var a = gateHexToRgb(hexA), b = gateHexToRgb(hexB);
  if (!a || !b) return hexA;
  var m = function (x, y) { return x * (1 - t) + y * t; };
  return gateRgbToHex(m(a.r, b.r), m(a.g, b.g), m(a.b, b.b));
}

function gateIsLight(hex) {
  var rgb = gateHexToRgb(hex);
  if (!rgb) return false;
  var lin = function (c) { c /= 255; return c <= 0.03928 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4); };
  return (0.2126 * lin(rgb.r) + 0.7152 * lin(rgb.g) + 0.0722 * lin(rgb.b)) > 0.6;
}

// Body theme class for first paint: a built-in maps to its own class (full CSS lives in the head partial),
// a resolvable custom theme to ts-theme-custom, anything unknown to a safe dark fallback.
function getActiveThemeBodyClass(settings) {
  try {
    var theme = settings && settings.theme ? String(settings.theme) : 'dark';
    if (theme === 'custom') return readLegacyCustomGateBase(settings) ? 'ts-theme-custom' : 'ts-theme-dark';
    if (theme.indexOf('custom:') === 0) return readCustomThemeGateBase(settings, theme.slice(7)) ? 'ts-theme-custom' : 'ts-theme-dark';
    var safeKey = theme.replace(/[^a-z0-9-]/gi, '');
    return (safeKey && readBuiltinThemeGateBase(safeKey)) ? 'ts-theme-' + safeKey : 'ts-theme-dark';
  } catch (e) { return 'ts-theme-dark'; }
}

// For a custom theme, a small <style> overriding .ts-theme-custom with the dominant tokens so the page
// paints in-theme on frame one; the client injects the complete token set a moment later. '' for built-ins.
function getCustomThemeFirstPaintCss(settings) {
  try {
    var theme = settings && settings.theme ? String(settings.theme) : '';
    var raw = theme === 'custom' ? readLegacyCustomGateBase(settings)
      : (theme.indexOf('custom:') === 0 ? readCustomThemeGateBase(settings, theme.slice(7)) : null);
    if (!raw) return '';
    var bg = sanitizeGateHex(raw.bg) || GATE_THEME_FALLBACK.bg;
    var panel = sanitizeGateHex(raw.panel) || GATE_THEME_FALLBACK.panel;
    var text = sanitizeGateHex(raw.text) || GATE_THEME_FALLBACK.text;
    var muted = sanitizeGateHex(raw.muted) || GATE_THEME_FALLBACK.muted;
    var primary = sanitizeGateHex(raw.primary) || GATE_THEME_FALLBACK.primary;
    var border = gateMix(panel, bg, 0.35);
    var surfaceSubtle = gateMix(bg, gateIsLight(bg) ? '#000000' : '#ffffff', 0.04);
    var css = '.ts-theme-custom{'
      + '--bg:' + bg + ';--panel:' + panel + ';--text:' + text + ';--muted:' + muted + ';'
      + '--primary:' + primary + ';--border:' + border + ';--border-color:' + border + ';'
      + '--surface-subtle:' + surfaceSubtle + ';--text-secondary:' + muted + ';}';
    return '<style id="custom-theme-firstpaint">' + css + '</style>';
  } catch (e) { return ''; }
}

function api_getWebAppUrl() {
  var url = '';
  try {
    url = ScriptApp.getService().getUrl() || '';
  } catch (e) {
    Logger.log('Error getting service URL: ' + e.toString());
  }
  return url;
}

/**
 * Adds a custom menu to the spreadsheet when it opens.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Tempus')
    .addItem('Get Web App URL', 'getTempusWebAppUrl')
    .addToUi();
}

/**
 * Gets the Tempus web app URL and displays it, or shows setup instructions.
 */
function getTempusWebAppUrl() {
  var ui = SpreadsheetApp.getUi();
  var url = null;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var targetSheetName = 'INTRO';
  var targetCell = 'B12';
  var instructionsCell = 'M2';
  var linkLabel = 'Open Tempus';
  var settings = {};
  var storedUrl = '';

  try {
    url = ScriptApp.getService().getUrl();
  } catch (e) {
    Logger.log('Error getting service URL: ' + e.toString());
  }

  try {
    settings = api_getSettings ? api_getSettings() : {};
    storedUrl = settings && settings.tempus_url ? String(settings.tempus_url) : '';
  } catch (e) {
    Logger.log('Error loading settings: ' + e.toString());
  }

  var targetSheet = ss.getSheetByName(targetSheetName);
  var clearIntroInstructions = function() {
    if (!targetSheet) return;
    targetSheet.getRange(instructionsCell).setValue('');
  };

  if (storedUrl) {
    clearIntroInstructions();
    Logger.log('getTempusWebAppUrl: opening stored Tempus URL: ' + storedUrl);
    var html = HtmlService.createHtmlOutput(
      '<html><body style="font-family:Arial, sans-serif; padding:16px;">' +
      '<p style="margin:0 0 12px;">Opening Tempus in a new tab.<br><br>If your browser blocks the pop-up, click the button below.</p>' +
      '<button id="open-tempus" style="padding:6px 10px; font-size:12px;">Open Tempus</button>' +
      '<script>' +
      'var opened = window.open("' + storedUrl + '", "_blank", "noopener");' +
      'document.getElementById("open-tempus").addEventListener("click", function(){' +
      '  window.open("' + storedUrl + '", "_blank", "noopener");' +
      '});' +
      '</script>' +
      '</body></html>'
    ).setWidth(360).setHeight(150);
    ui.showModelessDialog(html, 'Opening Tempus');
    return storedUrl;
  }

  var message = '';

  // Check if we have a test deployment URL (contains /dev)
  if (url && url.indexOf('/dev') !== -1) {
    if (!targetSheet) {
      message = 'Your Tempus Web App URL:\n\n' +
                url + '\n\n' +
                'Note: The "' + targetSheetName + '" sheet was not found, so the link could not be written.';
    } else {
      var hyperlinkFormula = '=HYPERLINK("' + url + '","' + linkLabel + '")';
      targetSheet.getRange(targetCell).setFormula(hyperlinkFormula);
      clearIntroInstructions();
      message = 'Your Tempus Web App URL:\n\n' +
                url + '\n\n' +
                'Saved as a clickable link in ' + targetSheetName + '!' + targetCell + '.';
    }
  } else {
    var instructions = 'Unable to retrieve the test deployment.\n\n' +
                       'To create a test deployment, or get an existing URL that the script cannot find:\n\n' +
                       '1. Click "Extensions" → "Apps Script" in the toolbar above\n' +
                       '2. In Apps Script Editor, click "Deploy" → "Test deployments"\n' +
                       '3. Click the "Copy" button under the URL\n\n' +
                       'Bookmark this URL for easy access.\n\n' +
                       'The test deployment URL automatically updates when you push updates to Tempus, so this is a one-time process.\n\n' +
                       'To share access with others:\n' +
                       '1. Share your Google Sheet with them (Editor or Viewer)\n' +
                       '2. Give them the URL\n';
    if (targetSheet) {
      targetSheet.getRange(instructionsCell).setValue(instructions);
      ss.setActiveSheet(targetSheet);
      message = 'To get the URL for the Tempus frontend follow the instructions that have been posted in the INTRO sheet.';
    } else {
      message = instructions + '\n' +
                'Note: The "' + targetSheetName + '" sheet was not found, so the instructions could not be written.';
    }
  }

  ui.alert('Tempus Web App URL', message, ui.ButtonSet.OK);

  return url;
}

/**
 * ⚠️ DANGER: Deletes all sheets except 'INTRO' to reset dev environment.
 *
 * This function is ONLY for development/testing purposes.
 * It will permanently delete all data sheets (timesheet_entries, contracts, etc.)
 * while preserving only the INTRO sheet.
 *
 * MANUAL EXECUTION ONLY - Do NOT create buttons or UI triggers for this function.
 * Run from Apps Script Editor: Select function → Run
 *
 * Use case: Resetting dev environment to factory settings for testing.
 */
function DANGER_resetToFactorySettings() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  // First confirmation
  var response1 = ui.alert(
    '⚠️ DANGER: Reset to Factory Settings',
    'This will PERMANENTLY DELETE all sheets except INTRO.\n\n' +
    'ALL DATA WILL BE LOST:\n' +
    '- timesheet_entries\n' +
    '- contracts\n' +
    '- user_settings\n' +
    '- feature_flags\n' +
    '- projects\n' +
    '- And any other sheets\n\n' +
    'Are you absolutely sure you want to continue?',
    ui.ButtonSet.YES_NO
  );

  if (response1 !== ui.Button.YES) {
    Logger.log('Reset cancelled by user at first confirmation');
    return { success: false, message: 'Reset cancelled' };
  }

  // Second confirmation
  var response2 = ui.alert(
    '⚠️ FINAL WARNING',
    'This action CANNOT BE UNDONE.\n\n' +
    'All your time entries, contracts, and settings will be permanently deleted.\n\n' +
    'Click YES to proceed with deletion.',
    ui.ButtonSet.YES_NO
  );

  if (response2 !== ui.Button.YES) {
    Logger.log('Reset cancelled by user at second confirmation');
    return { success: false, message: 'Reset cancelled' };
  }

  // Proceed with deletion
  var sheets = ss.getSheets();
  var deletedSheets = [];
  var errors = [];

  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var sheetName = sheet.getName();

    if (sheetName !== 'INTRO') {
      try {
        ss.deleteSheet(sheet);
        deletedSheets.push(sheetName);
        Logger.log('Deleted sheet: ' + sheetName);
      } catch (e) {
        errors.push(sheetName + ': ' + e.toString());
        Logger.log('Error deleting sheet ' + sheetName + ': ' + e.toString());
      }
    }
  }

  var resultMessage = 'Factory reset complete!\n\n' +
                      'Deleted ' + deletedSheets.length + ' sheets:\n' +
                      deletedSheets.join(', ') + '\n\n' +
                      'Preserved: INTRO';

  if (errors.length > 0) {
    resultMessage += '\n\nErrors:\n' + errors.join('\n');
  }

  ui.alert('Reset Complete', resultMessage, ui.ButtonSet.OK);

  Logger.log('Factory reset complete. Deleted: ' + deletedSheets.length + ' sheets');

  return {
    success: true,
    deletedSheets: deletedSheets,
    errors: errors
  };
}

function pickInitialPage_(lilMode, requestedPage) {
  if (requestedPage === 'invoices') return lilMode ? 'assessments' : 'time';
  if (requestedPage === 'assessments') return lilMode ? 'assessments' : 'time';
  if (['time','contracts','hour-types','annual-views','settings','help','themes','about','bas','expenses','deductions','rate-preview','public-holidays'].indexOf(requestedPage) !== -1) return requestedPage;
  return lilMode ? 'assessments' : 'time';
}
