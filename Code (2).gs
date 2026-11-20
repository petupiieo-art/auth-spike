/**
 * AUTH SPIKE - Option C backend (Google Apps Script, bound to the test spreadsheet).  Revision 4.
 *
 * Deploy as web app: Execute as = Me, Who has access = Anyone.
 * Required Script Property: GOOGLE_CLIENT_ID  (your OAuth Web client ID).
 * Optional Script Property : ENABLE_ACCESS_STATUS = true  -> Option B (NOT approved yet). Leave it UNSET.
 *
 * USERS SHEET SCHEMA (single source of truth: USERS_HEADERS below; setup, authorization and tests all use it)
 *   Email | GoogleSub | Role | EmployeeID | EmploymentStatus | AccessStatus
 *   There is no generic "Status" column. A sheet that still has one is refused (USER_CONFIG_ERROR).
 *
 * IDENTITY MODEL
 *   - Primary identity is Google's stable account ID (the token's "sub"). Users are looked up by GoogleSub.
 *   - Email is used ONLY once, to bind a new account: the first time a verified token arrives whose email matches
 *     a Users row that has no GoogleSub yet, that row is bound to the token's sub. After that, email is never
 *     used to find the user.
 *   - Role and EmployeeID come ONLY from the Users sheet. Nothing the browser sends is trusted for identity.
 *
 * TOKEN VALIDATION (prototype)
 *   Signature verification is DELEGATED to Google's tokeninfo endpoint; this script does not verify the JWT
 *   signature itself. After Google accepts the token, this script checks: issuer is Google, audience equals our
 *   OAuth client ID, token not expired, email_verified is true, and sub is present. Verified results are cached
 *   for at most 5 minutes (keyed by a hash of the token). Authorization is re-read from the sheet on EVERY request.
 *   Google documents tokeninfo as a debugging endpoint; verifying signatures locally is a Phase 1B decision.
 *
 * STATUS CONCEPTS (kept completely separate)
 *   EmploymentStatus ACTIVE | INACTIVE : administrative label for the Admin only. NEVER read by any access
 *                                        decision and never returned to anyone.
 *   AccessStatus     ACTIVE | SUSPENDED: Option A (DEFAULT, this prototype): ignored completely.
 *                                        Option B (needs your approval of Q28): enforced only when the Script
 *                                        Property ENABLE_ACCESS_STATUS is exactly "true". Dormant otherwise.
 *
 * FAIL CLOSED
 *   A duplicate Email or GoogleSub anywhere in Users makes every request fail with USER_CONFIG_ERROR until fixed.
 *
 * Throwaway test code: no business logic, no real data.
 */

var CODE_VERSION = 'auth-spike-r4';
var SHEET_USERS = 'Users';
var SHEET_CLIENTS = 'Clients_Canary';
var USERS_HEADERS = ['Email', 'GoogleSub', 'Role', 'EmployeeID', 'EmploymentStatus', 'AccessStatus'];
var ROLES = ['ADMIN', 'EMPLOYEE'];

// Default-deny policy: an action not listed here does not exist; a role not listed for an action is refused.
var POLICY = {
  whoami:     ['ADMIN', 'EMPLOYEE'],
  myClients:  ['EMPLOYEE'],
  getClient:  ['ADMIN', 'EMPLOYEE'],
  allClients: ['ADMIN']
};

/* ---------- ONE-TIME SETUP (run from the editor) ---------- */

function setupSpike() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var u = ss.getSheetByName(SHEET_USERS) || ss.insertSheet(SHEET_USERS);
  u.clear();
  u.getRange('A1:Z100').clearDataValidations();
  u.getRange('B:B').setNumberFormat('@'); // GoogleSub is a long number: keep it as text
  u.getRange(1, 1, 3, USERS_HEADERS.length).setValues([
    USERS_HEADERS,
    ['PUT_GMAIL_1_HERE@gmail.com', '', 'ADMIN', '', 'ACTIVE', 'ACTIVE'],
    ['PUT_GMAIL_2_HERE@gmail.com', '', 'EMPLOYEE', 'EMP-A', 'ACTIVE', 'ACTIVE']
  ]);
  // Dropdowns help, but invalid values are still accepted (with a warning) so the guide's negative tests can run.
  setList_(u.getRange('C2:C100'), ['ADMIN', 'EMPLOYEE']);
  setList_(u.getRange('E2:E100'), ['ACTIVE', 'INACTIVE']);
  setList_(u.getRange('F2:F100'), ['ACTIVE', 'SUSPENDED']);
  u.getRange('B1').setNote('Filled automatically at first sign-in. After that this is the login identity.');
  u.getRange('E1').setNote('Admin-only label. Never read by the backend, never shown to employees.');
  u.getRange('F1').setNote('Ignored by default (Option A). Enforced only if Script Property ENABLE_ACCESS_STATUS = true (Option B, not approved).');

  var c = ss.getSheetByName(SHEET_CLIENTS) || ss.insertSheet(SHEET_CLIENTS);
  c.clear();
  c.getRange(1, 1, 5, 3).setValues([
    ['ClientID', 'EmployeeID', 'Name'],
    ['CLI-A1', 'EMP-A', 'CANARY-A-Alice'],
    ['CLI-A2', 'EMP-A', 'CANARY-A-Amal'],
    ['CLI-B1', 'EMP-B', 'CANARY-B-Bassem'],
    ['CLI-B2', 'EMP-B', 'CANARY-B-Basma']
  ]);

  // Triggers the one-time permission prompt for outbound requests.
  UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=x', { muteHttpExceptions: true });
  Logger.log('Setup done (' + CODE_VERSION + '). Now edit the Users sheet and set GOOGLE_CLIENT_ID in Script Properties.');
}

function setList_(range, values) {
  range.setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(values, true).setAllowInvalid(true).build()
  );
}

/* ---------- ENTRY POINTS ---------- */

// GET is only a probe of Apps Script's native identity (authentication Option A). Not used for login.
function doGet() {
  return json_({
    ok: true,
    data: {
      version: CODE_VERSION,
      activeUserEmail: Session.getActiveUser().getEmail(),
      note: 'Empty = Apps Script cannot identify this caller natively.'
    }
  });
}

function doPost(e) {
  var reqId = null;
  try {
    // 1. Validate the request before touching any property of it.
    if (!e || !e.postData || typeof e.postData.contents !== 'string' || !e.postData.contents) throw fail_('BAD_REQUEST');
    var req;
    try { req = JSON.parse(e.postData.contents); } catch (x) { throw fail_('BAD_REQUEST'); }
    if (!req || typeof req !== 'object' || Array.isArray(req)) throw fail_('BAD_REQUEST');
    if (typeof req.action !== 'string' || !req.action) throw fail_('BAD_REQUEST');
    reqId = typeof req.requestId === 'string' ? req.requestId.slice(0, 64) : null;
    var payload = normalizePayload_(req.payload);

    // 2. Authenticate, then authorize. Only idToken is read for identity.
    var ident = verifyIdToken_(req.idToken);
    var user = authorize_(ident);        // invalid roles never get past this line

    // 3. Route (default deny).
    var data = route_(req.action, payload, user);
    return json_({ ok: true, requestId: reqId, data: data });
  } catch (err) {
    var code = (err && err.code) ? err.code : 'SERVER_ERROR';
    if (code === 'SERVER_ERROR') console.error(err);
    return json_({ ok: false, requestId: reqId, error: { code: code } });
  }
}

// Always a plain object with own, safe keys only (never an array, string, null or a prototype carrier).
function normalizePayload_(p) {
  var out = {};
  if (!p || typeof p !== 'object' || Array.isArray(p)) return out;
  Object.keys(p).forEach(function (k) {
    if (k === '__proto__' || k === 'constructor' || k === 'prototype') return;
    out[k] = p[k];
  });
  return out;
}

/* ---------- AUTHENTICATION: validate the Google ID token (see header) ---------- */

function verifyIdToken_(token) {
  if (!token || typeof token !== 'string' || token.length > 4096) throw fail_('NO_TOKEN');

  var cache = CacheService.getScriptCache();
  var key = 'tok:' + sha256_(token);
  var hit = cache.get(key);
  if (hit) {
    var c = JSON.parse(hit);
    if (c.exp * 1000 > Date.now()) return c;
  }

  var clientId = PropertiesService.getScriptProperties().getProperty('GOOGLE_CLIENT_ID');
  if (!clientId) throw fail_('SERVER_NOT_CONFIGURED');

  // Signature verification happens at Google; a non-200 answer means the token is not valid.
  var res = UrlFetchApp.fetch(
    'https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(token),
    { muteHttpExceptions: true }
  );
  if (res.getResponseCode() !== 200) throw fail_('INVALID_TOKEN');

  var p = JSON.parse(res.getContentText());
  var okIssuer = p.iss === 'accounts.google.com' || p.iss === 'https://accounts.google.com';
  var okEmail = p.email_verified === true || p.email_verified === 'true';
  if (p.aud !== clientId || !okIssuer || !okEmail || !p.sub || !p.email) throw fail_('INVALID_TOKEN');

  var exp = Number(p.exp);
  if (!(exp * 1000 > Date.now())) throw fail_('TOKEN_EXPIRED');

  var ident = { sub: String(p.sub), email: String(p.email).toLowerCase(), exp: exp };
  var ttl = Math.min(300, Math.floor(exp - Date.now() / 1000));
  if (ttl > 0) cache.put(key, JSON.stringify(ident), ttl);
  return ident;
}

/* ---------- AUTHORIZATION: identity by sub, role from the Users sheet, default deny ---------- */

function accessStatusEnforced_() { // Option B switch. Anything other than the exact string "true" means Option A.
  return PropertiesService.getScriptProperties().getProperty('ENABLE_ACCESS_STATUS') === 'true';
}

function authorize_(ident) {
  var enforceAccess = accessStatusEnforced_();
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (x) { throw fail_('BUSY'); } // first-time binding must not race
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
    if (!sh) throw fail_('USER_CONFIG_ERROR');
    var rows = sh.getDataRange().getValues();

    // Schema: exactly the approved columns; the old generic "Status" column is refused.
    var h = rows[0].map(function (x) { return String(x).trim(); });
    var col = {};
    USERS_HEADERS.forEach(function (n) {
      col[n] = h.indexOf(n);
      if (col[n] === -1) throw fail_('USER_CONFIG_ERROR');
    });
    if (h.indexOf('Status') !== -1) throw fail_('USER_CONFIG_ERROR');

    // Read records and fail closed on duplicates anywhere (never "first match wins").
    var records = [], seenEmail = Object.create(null), seenSub = Object.create(null);
    for (var i = 1; i < rows.length; i++) {
      var rec = {
        i: i,
        email: String(rows[i][col.Email]).trim().toLowerCase(),
        sub: String(rows[i][col.GoogleSub]).trim(),
        role: String(rows[i][col.Role]),                 // exact value: no trimming, no case folding
        employeeId: String(rows[i][col.EmployeeID]).trim()
      };
      if (!rec.email && !rec.sub && !rec.role.trim() && !rec.employeeId) continue; // blank row
      if (rec.email) { if (seenEmail[rec.email]) throw fail_('USER_CONFIG_ERROR'); seenEmail[rec.email] = true; }
      if (rec.sub) { if (seenSub[rec.sub]) throw fail_('USER_CONFIG_ERROR'); seenSub[rec.sub] = true; }
      records.push(rec);
    }

    // Find the user: primary path is the sub; email only binds a not-yet-bound row.
    var match = null, bindNow = false;
    var bySub = records.filter(function (r) { return r.sub && r.sub === ident.sub; });
    if (bySub.length === 1) {
      match = bySub[0];
    } else {
      var byEmail = records.filter(function (r) { return r.email && r.email === ident.email; });
      if (byEmail.length === 0) throw fail_('NOT_AUTHORIZED');
      if (byEmail[0].sub) throw fail_('IDENTITY_MISMATCH'); // that email is bound to a different Google account
      match = byEmail[0];
      bindNow = true;
    }

    // Required fields of the matched user.
    if (!match.email || match.email.indexOf('@') < 1) throw fail_('USER_CONFIG_ERROR');
    if (ROLES.indexOf(match.role) === -1) throw fail_('ROLE_INVALID');
    if (match.role === 'EMPLOYEE' && !match.employeeId) throw fail_('EMPLOYEE_ID_MISSING');

    // Option B only (dormant unless explicitly enabled). EmploymentStatus is never consulted.
    if (enforceAccess) {
      var access = String(rows[match.i][col.AccessStatus]);
      if (access === 'SUSPENDED') throw fail_('ACCOUNT_SUSPENDED');
      if (access !== 'ACTIVE') throw fail_('ACCESS_STATUS_INVALID');
    }

    if (bindNow) sh.getRange(match.i + 1, col.GoogleSub + 1).setNumberFormat('@').setValue(ident.sub);

    return {
      email: ident.email,
      sub: ident.sub,
      role: match.role,
      employeeId: match.role === 'EMPLOYEE' ? match.employeeId : '',
      suspensionEnforced: enforceAccess
    };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- ROUTER (default deny) ---------- */

function route_(action, payload, user) {
  if (!Object.prototype.hasOwnProperty.call(POLICY, action)) throw fail_('UNKNOWN_ACTION');
  if (!user || ROLES.indexOf(user.role) === -1 || POLICY[action].indexOf(user.role) === -1) throw fail_('FORBIDDEN');

  switch (action) {
    case 'whoami':
      return {
        email: user.email,
        role: user.role,
        employeeId: user.employeeId,
        suspensionEnforced: !!user.suspensionEnforced,
        ignoredFromRequest: ['role', 'employeeId', 'email'].filter(function (k) {
          return Object.prototype.hasOwnProperty.call(payload, k);
        })
      };

    case 'myClients': // scope comes from the session, never from the payload
      if (!user.employeeId) throw fail_('FORBIDDEN');
      return clients_()
        .filter(function (c) { return c.EmployeeID === user.employeeId; })
        .map(function (c) { return { clientId: c.ClientID, name: c.Name }; });

    case 'getClient': { // explicit allow rules; everything else looks exactly like a missing ID
      var c = clients_().filter(function (x) { return x.ClientID === String(payload.clientId); })[0];
      if (!c) throw fail_('NOT_FOUND');
      if (user.role === 'ADMIN') return { clientId: c.ClientID, name: c.Name };
      if (user.role === 'EMPLOYEE' && user.employeeId && c.EmployeeID === user.employeeId) {
        return { clientId: c.ClientID, name: c.Name };
      }
      throw fail_('NOT_FOUND');
    }

    case 'allClients':
      return clients_().map(function (c) {
        return { clientId: c.ClientID, employeeId: c.EmployeeID, name: c.Name };
      });

    default:
      throw fail_('UNKNOWN_ACTION');
  }
}

/* ---------- helpers ---------- */

function clients_() {
  var rows = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CLIENTS).getDataRange().getValues();
  var h = rows[0];
  return rows.slice(1).map(function (r) {
    var o = {};
    h.forEach(function (name, i) { o[name] = String(r[i]); });
    return o;
  });
}

function fail_(code) {
  var e = new Error(code);
  e.code = code;
  return e;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function sha256_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8)
    .map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}
