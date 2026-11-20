// Offline logic test for backend/Code.gs (Apps Script services are mocked).
// Run from the repository root:   node tests/logic-test.js
// It tests authorization and routing logic only. It does NOT test Google sign-in, tokeninfo, CORS, spreadsheet
// sharing or phones; those are the manual tests in the guide.
const fs = require('fs'), vm = require('vm'), crypto = require('crypto'), path = require('path');
const CODE_PATH = path.join(__dirname, '..', 'backend', 'Code.gs');
const SRC = fs.readFileSync(CODE_PATH, 'utf8');
const SHA = crypto.createHash('sha256').update(SRC).digest('hex');
const EXPECTED_HEADERS = ['Email','GoogleSub','Role','EmployeeID','EmploymentStatus','AccessStatus'];
const now = () => Math.floor(Date.now() / 1000);
const tk = o => Object.assign({ aud:'CID', iss:'accounts.google.com', email_verified:'true', exp: now()+3000 }, o);
const TOKENS = {
  'tok-admin': tk({ sub:'900', email:'gmail1@gmail.com' }),
  'tok-emp': tk({ sub:'111', email:'gmail2@gmail.com' }),
  'tok-emp-renamed': tk({ sub:'111', email:'gmail2-renamed@gmail.com' }),
  'tok-emp-other-sub': tk({ sub:'222', email:'gmail2@gmail.com' }),
  'tok-stranger': tk({ sub:'333', email:'stranger@gmail.com' }),
  'tok-wrong-aud': tk({ aud:'OTHER', sub:'111', email:'gmail2@gmail.com' }),
  'tok-wrong-iss': tk({ iss:'evil.example.com', sub:'111', email:'gmail2@gmail.com' }),
  'tok-expired': tk({ exp: now()-10, sub:'111', email:'gmail2@gmail.com' }),
  'tok-unverified': tk({ email_verified:'false', sub:'111', email:'gmail2@gmail.com' }),
  'tok-no-sub': tk({ sub:'', email:'gmail2@gmail.com' })
};

function makeCtx(sheetsInit, props, opts) {
  opts = opts || {};
  const sheets = sheetsInit, cache = {}, written = {};
  const rangeObj = (name) => ({
    setValues(v) { written[name] = v; return this; }, setNumberFormat() { return this; }, clearDataValidations() { return this; },
    setDataValidation() { return this; }, setNote() { return this; }, setValue() { return this; }
  });
  const newSheet = name => ({ clear() {}, getRange: () => rangeObj(name) });
  const sheetFromRows = rows => ({ getDataRange: () => ({ getValues: () => rows.map(r => r.slice()) }),
    getRange: (r, c) => ({ setNumberFormat() { return this; }, setValue(v) { rows[r-1][c-1] = v; } }) });
  const wrapped = {};
  Object.keys(sheets).forEach(n => { wrapped[n] = sheetFromRows(sheets[n]); });
  const ctx = vm.createContext({
    console, JSON, Math, Date, Number, String, Object, Array, Error,
    Logger: { log() {} },
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ({ getSheetByName: n => wrapped[n], insertSheet: n => newSheet(n) }),
      newDataValidation: () => ({ requireValueInList() { return this; }, setAllowInvalid() { return this; }, build() { return {}; } })
    },
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
    CacheService: { getScriptCache: () => ({ get: k => cache[k] || null, put: (k, v) => { cache[k] = v; } }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => ({ GOOGLE_CLIENT_ID: 'CID', ...props })[k] || null }) },
    UrlFetchApp: { fetch: url => { const t = decodeURIComponent(String(url).split('id_token=')[1]);
      return TOKENS[t] ? { getResponseCode: () => 200, getContentText: () => JSON.stringify(TOKENS[t]) }
                       : { getResponseCode: () => 400, getContentText: () => '{}' }; } },
    Utilities: { DigestAlgorithm: { SHA_256: 1 }, Charset: { UTF_8: 1 },
      computeDigest: (a, s) => Array.from(crypto.createHash('sha256').update(s).digest()).map(b => b > 127 ? b - 256 : b) },
    ContentService: { MimeType: { JSON: 1 }, createTextOutput: s => ({ text: s, setMimeType() { return this; } }) },
    Session: { getActiveUser: () => ({ getEmail: () => '' }) }
  });
  vm.runInContext(SRC, ctx);
  return { ctx, sheets, written };
}

// ---- Fail fast if the backend under test is not the expected revision ----
console.log('Testing backend/Code.gs  sha256=' + SHA);
const probe = makeCtx({}, {});
if (typeof probe.ctx.CODE_VERSION !== 'string' || !Array.isArray(probe.ctx.USERS_HEADERS)) {
  console.error('\nSTOP: this backend/Code.gs is an OLD revision (no CODE_VERSION / USERS_HEADERS).');
  console.error('Replace it with the latest delivered Code.gs and run again.');
  process.exit(2);
}
console.log('CODE_VERSION=' + probe.ctx.CODE_VERSION);
const H = probe.ctx.USERS_HEADERS.slice();

// ---- Test helpers ----
let pass = 0, fail = 0, group = '';
const failures = [], perGroup = {};
function G(name) { group = name; perGroup[name] = { pass: 0, fail: 0 }; }
function t(name, fn) {
  let ok = false, msg = '';
  try { ok = fn() === true; if (!ok) msg = 'assertion false'; } catch (e) { msg = 'threw: ' + e.message; }
  if (ok) { pass++; perGroup[group].pass++; } else { fail++; perGroup[group].fail++; failures.push(`[${group}] ${name} (${msg})`); }
}
const USER = (o) => H.map(n => (o[n] !== undefined ? o[n] : (n === 'EmploymentStatus' || n === 'AccessStatus') ? 'ACTIVE' : ''));
const ADMIN = () => USER({ Email:'gmail1@gmail.com', Role:'ADMIN' });
const EMP = (o) => USER(Object.assign({ Email:'gmail2@gmail.com', Role:'EMPLOYEE', EmployeeID:'EMP-A' }, o || {}));
const CLIENTS = () => [['ClientID','EmployeeID','Name'],['CLI-A1','EMP-A','CANARY-A-Alice'],['CLI-A2','EMP-A','CANARY-A-Amal'],
  ['CLI-B1','EMP-B','CANARY-B-Bassem'],['CLI-B2','EMP-B','CANARY-B-Basma'],['CLI-X','','CANARY-BLANK-OWNER']];
function env(userRows, props, headers) { // returns { ctx, users, clients }
  const users = [(headers || H).slice(), ...userRows.map(r => r.slice())];
  const e = makeCtx({ Users: users, Clients_Canary: CLIENTS() }, props || {});
  e.users = users; return e;
}
function call(e, action, payload, token, extra) {
  const body = JSON.stringify(Object.assign({}, extra || {}, { action, payload, requestId: 'r1', idToken: token }));
  return JSON.parse(e.ctx.doPost({ postData: { contents: body } }).text);
}
function raw(e, contents) { return JSON.parse(e.ctx.doPost(contents === undefined ? {} : { postData: { contents } }).text); }
const code = r => r.ok ? 'OK' : r.error.code;
function throwsCode(fn) { try { fn(); return 'NO_THROW'; } catch (e) { return e.code || 'ERR:' + e.message; } }

// ======================= TESTS =======================
G('1 Schema is identical everywhere');
t('CODE_VERSION is set', () => /^auth-spike-r\d+$/.test(probe.ctx.CODE_VERSION));
t('USERS_HEADERS is exactly the six approved columns', () => JSON.stringify(H) === JSON.stringify(EXPECTED_HEADERS));
t('no generic Status column in USERS_HEADERS', () => H.indexOf('Status') === -1);
t('setupSpike writes exactly USERS_HEADERS as the Users header row', () => {
  const s = makeCtx({}, {}); s.ctx.setupSpike(); return JSON.stringify(s.written.Users[0]) === JSON.stringify(EXPECTED_HEADERS); });
t('setupSpike sample rows have one cell per column', () => {
  const s = makeCtx({}, {}); s.ctx.setupSpike(); return s.written.Users.every(r => r.length === EXPECTED_HEADERS.length); });
t('setupSpike sample rows are ADMIN and EMPLOYEE only, both ACTIVE/ACTIVE', () => {
  const s = makeCtx({}, {}); s.ctx.setupSpike(); const u = s.written.Users;
  return u[1][2] === 'ADMIN' && u[2][2] === 'EMPLOYEE' && u[1][4] === 'ACTIVE' && u[1][5] === 'ACTIVE' && u[2][4] === 'ACTIVE' && u[2][5] === 'ACTIVE'; });
t('legacy single Status column is refused', () => { const h = ['Email','GoogleSub','Role','EmployeeID','EmploymentStatus','AccessStatus','Status'];
  return code(call(env([[...EMP(), 'ACTIVE']], {}, h), 'whoami', {}, 'tok-emp')) === 'USER_CONFIG_ERROR'; });
t('old 5-column sheet (Status instead of the two status columns) is refused', () => {
  const h = ['Email','GoogleSub','Role','EmployeeID','Status']; const e = env([['gmail2@gmail.com','','EMPLOYEE','EMP-A','ACTIVE']], {}, h);
  return code(call(e, 'whoami', {}, 'tok-emp')) === 'USER_CONFIG_ERROR'; });
for (const missing of H) t(`missing header ${missing} is refused`, () => {
  const h = H.filter(n => n !== missing); const idx = H.indexOf(missing);
  const e = env([EMP().filter((_, i) => i !== idx)], {}, h); return code(call(e, 'whoami', {}, 'tok-emp')) === 'USER_CONFIG_ERROR'; });

G('2 Strict role validation (default deny)');
for (const bad of ['admin','Admin','employee','Employee','MANAGER','','EMPLOYEE ',' ADMIN','SUPERUSER','null','undefined','ADMIN,EMPLOYEE']) {
  for (const [a, p] of [['whoami',{}],['getClient',{clientId:'CLI-B1'}],['getClient',{clientId:'CLI-A1'}],['allClients',{}],['myClients',{}]]) {
    t(`role "${bad}" -> ${a} ${p.clientId || ''} is ROLE_INVALID and returns no data`, () => {
      const r = call(env([ADMIN(), EMP({ Role: bad })]), a, p, 'tok-emp'); return !r.ok && r.error.code === 'ROLE_INVALID' && r.data === undefined; });
  }
}
t('ADMIN is accepted', () => code(call(env([ADMIN()]), 'whoami', {}, 'tok-admin')) === 'OK');
t('EMPLOYEE is accepted', () => code(call(env([ADMIN(), EMP()]), 'whoami', {}, 'tok-emp')) === 'OK');
t('route_ itself refuses an unknown role on every action (defense in depth)', () =>
  ['whoami','myClients','getClient','allClients'].every(a => throwsCode(() => probe.ctx.route_(a, { clientId:'CLI-B1' }, { role:'OWNER', employeeId:'EMP-B' })) === 'FORBIDDEN'));
t('route_ refuses a missing user object', () => throwsCode(() => probe.ctx.route_('whoami', {}, null)) === 'FORBIDDEN');
t('route_ getClient: EMPLOYEE without employeeId cannot read an ownerless client', () =>
  throwsCode(() => probe.ctx.route_('getClient', { clientId:'CLI-X' }, { role:'EMPLOYEE', employeeId:'' })) !== 'NO_THROW');

G('3 Duplicates and required fields');
t('duplicate email (both unbound) -> USER_CONFIG_ERROR', () => code(call(env([ADMIN(), EMP(), EMP({ Role:'ADMIN', EmployeeID:'' })]), 'whoami', {}, 'tok-emp')) === 'USER_CONFIG_ERROR');
t('duplicate email (one bound, one unbound) -> USER_CONFIG_ERROR', () => code(call(env([ADMIN(), EMP({ GoogleSub:'111' }), EMP({ Role:'ADMIN', EmployeeID:'' })]), 'whoami', {}, 'tok-emp')) === 'USER_CONFIG_ERROR');
t('duplicate email (different case) -> USER_CONFIG_ERROR', () => code(call(env([ADMIN(), EMP(), EMP({ Email:'GMAIL2@gmail.com' })]), 'whoami', {}, 'tok-emp')) === 'USER_CONFIG_ERROR');
t('duplicate email rows never authorize the first match', () => { const r = call(env([EMP(), EMP({ Role:'ADMIN', EmployeeID:'' })]), 'allClients', {}, 'tok-emp'); return !r.ok && r.error.code === 'USER_CONFIG_ERROR'; });
t('duplicate GoogleSub -> USER_CONFIG_ERROR', () => code(call(env([ADMIN(), EMP({ GoogleSub:'111' }), USER({ Email:'x@y.com', Role:'ADMIN', GoogleSub:'111' })]), 'whoami', {}, 'tok-emp')) === 'USER_CONFIG_ERROR');
t('a duplicate elsewhere fails closed for everyone, admin included', () => code(call(env([ADMIN(), EMP(), EMP()]), 'whoami', {}, 'tok-admin')) === 'USER_CONFIG_ERROR');
t('EMPLOYEE with empty EmployeeID -> EMPLOYEE_ID_MISSING', () => code(call(env([ADMIN(), EMP({ EmployeeID:'' })]), 'myClients', {}, 'tok-emp')) === 'EMPLOYEE_ID_MISSING');
t('EMPLOYEE with blank-space EmployeeID -> EMPLOYEE_ID_MISSING', () => code(call(env([ADMIN(), EMP({ EmployeeID:'   ' })]), 'myClients', {}, 'tok-emp')) === 'EMPLOYEE_ID_MISSING');
t('empty-EmployeeID employee cannot read the ownerless client row', () => code(call(env([ADMIN(), EMP({ EmployeeID:'' })]), 'getClient', { clientId:'CLI-X' }, 'tok-emp')) === 'EMPLOYEE_ID_MISSING');
t('matched row with empty Email -> USER_CONFIG_ERROR', () => code(call(env([ADMIN(), EMP({ Email:'', GoogleSub:'111' })]), 'whoami', {}, 'tok-emp')) === 'USER_CONFIG_ERROR');
t('matched row with malformed Email -> USER_CONFIG_ERROR', () => code(call(env([ADMIN(), EMP({ Email:'not-an-email', GoogleSub:'111' })]), 'whoami', {}, 'tok-emp')) === 'USER_CONFIG_ERROR');
t('blank rows are ignored', () => code(call(env([ADMIN(), USER({}), EMP(), USER({})]), 'whoami', {}, 'tok-emp')) === 'OK');
t('unknown user -> NOT_AUTHORIZED', () => code(call(env([ADMIN()]), 'whoami', {}, 'tok-emp')) === 'NOT_AUTHORIZED');

G('4 Identity is the Google sub; email only binds');
t('first login binds GoogleSub', () => { const e = env([ADMIN(), EMP()]); call(e, 'whoami', {}, 'tok-emp'); return e.users[2][H.indexOf('GoogleSub')] === '111'; });
t('after binding, changing the sheet email does not matter', () => { const e = env([ADMIN(), EMP()]); call(e, 'whoami', {}, 'tok-emp'); e.users[2][0] = 'someone.else@example.com'; return code(call(e, 'whoami', {}, 'tok-emp')) === 'OK'; });
t('after binding, a renamed Google email still works', () => { const e = env([ADMIN(), EMP()]); call(e, 'whoami', {}, 'tok-emp'); return code(call(e, 'whoami', {}, 'tok-emp-renamed')) === 'OK'; });
t('same email, different Google account -> IDENTITY_MISMATCH', () => { const e = env([ADMIN(), EMP({ GoogleSub:'111' })]); return code(call(e, 'whoami', {}, 'tok-emp-other-sub')) === 'IDENTITY_MISMATCH'; });
t('tampered GoogleSub cell -> IDENTITY_MISMATCH', () => code(call(env([ADMIN(), EMP({ GoogleSub:'112' })]), 'whoami', {}, 'tok-emp')) === 'IDENTITY_MISMATCH');
t('unbound row with typo email -> NOT_AUTHORIZED', () => code(call(env([ADMIN(), EMP({ Email:'typo@gmail.com' })]), 'whoami', {}, 'tok-emp')) === 'NOT_AUTHORIZED');
t('cleared GoogleSub with right email binds again', () => { const e = env([ADMIN(), EMP()]); call(e, 'whoami', {}, 'tok-emp'); e.users[2][1] = ''; return code(call(e, 'whoami', {}, 'tok-emp')) === 'OK' && e.users[2][1] === '111'; });
t('deleted row -> NOT_AUTHORIZED', () => { const e = env([ADMIN(), EMP()]); call(e, 'whoami', {}, 'tok-emp'); e.users.splice(2, 1); return code(call(e, 'whoami', {}, 'tok-emp')) === 'NOT_AUTHORIZED'; });
t('stranger -> NOT_AUTHORIZED', () => code(call(env([ADMIN(), EMP()]), 'whoami', {}, 'tok-stranger')) === 'NOT_AUTHORIZED');

G('5 Employee A/B isolation');
t('A sees only A clients', () => JSON.stringify(call(env([ADMIN(), EMP()]), 'myClients', {}, 'tok-emp').data.map(c => c.name)) === '["CANARY-A-Alice","CANARY-A-Amal"]');
t('B sees only B clients', () => JSON.stringify(call(env([ADMIN(), EMP({ EmployeeID:'EMP-B' })]), 'myClients', {}, 'tok-emp').data.map(c => c.name)) === '["CANARY-B-Bassem","CANARY-B-Basma"]');
t('A reads own client', () => code(call(env([ADMIN(), EMP()]), 'getClient', { clientId:'CLI-A1' }, 'tok-emp')) === 'OK');
t('A reading B client = NOT_FOUND', () => code(call(env([ADMIN(), EMP()]), 'getClient', { clientId:'CLI-B1' }, 'tok-emp')) === 'NOT_FOUND');
t('missing client = NOT_FOUND (indistinguishable from B client)', () => code(call(env([ADMIN(), EMP()]), 'getClient', { clientId:'CLI-ZZZ' }, 'tok-emp')) === 'NOT_FOUND');
t('B reading A client = NOT_FOUND', () => code(call(env([ADMIN(), EMP({ EmployeeID:'EMP-B' })]), 'getClient', { clientId:'CLI-A1' }, 'tok-emp')) === 'NOT_FOUND');
t('ownerless client hidden from employee', () => code(call(env([ADMIN(), EMP()]), 'getClient', { clientId:'CLI-X' }, 'tok-emp')) === 'NOT_FOUND');
t('employee cannot call allClients', () => code(call(env([ADMIN(), EMP()]), 'allClients', {}, 'tok-emp')) === 'FORBIDDEN');
t('admin allClients returns all 5', () => call(env([ADMIN(), EMP()]), 'allClients', {}, 'tok-admin').data.length === 5);
t('admin cannot call myClients', () => code(call(env([ADMIN(), EMP()]), 'myClients', {}, 'tok-admin')) === 'FORBIDDEN');
t('admin getClient works on any client', () => code(call(env([ADMIN(), EMP()]), 'getClient', { clientId:'CLI-B1' }, 'tok-admin')) === 'OK');

G('6 Forged browser values are ignored');
t('forged email/role/employeeId (payload and top level)', () => {
  const r = call(env([ADMIN(), EMP()]), 'whoami', { role:'ADMIN', employeeId:'EMP-B', email:'x@y.z' }, 'tok-emp', { role:'ADMIN', employeeId:'EMP-B', email:'x@y.z' });
  return r.data.role === 'EMPLOYEE' && r.data.employeeId === 'EMP-A' && r.data.email === 'gmail2@gmail.com' && r.data.ignoredFromRequest.length === 3; });
t('forged admin role cannot call allClients', () => code(call(env([ADMIN(), EMP()]), 'allClients', { role:'ADMIN' }, 'tok-emp', { role:'ADMIN' })) === 'FORBIDDEN');
t('forged employeeId cannot widen myClients', () => call(env([ADMIN(), EMP()]), 'myClients', { employeeId:'EMP-B' }, 'tok-emp', { employeeId:'EMP-B' }).data.every(c => c.name.indexOf('CANARY-A') === 0));
t('forged email cannot become another user', () => call(env([ADMIN(), EMP()]), 'whoami', { email:'gmail1@gmail.com' }, 'tok-emp', { email:'gmail1@gmail.com' }).data.role === 'EMPLOYEE');

G('7 Inactive label and Option A default');
for (const emps of ['INACTIVE', '', 'banana']) {
  t(`EmploymentStatus "${emps}": login and work allowed`, () => { const e = env([ADMIN(), EMP({ EmploymentStatus: emps })]);
    return code(call(e, 'whoami', {}, 'tok-emp')) === 'OK' && code(call(e, 'myClients', {}, 'tok-emp')) === 'OK' && code(call(e, 'getClient', { clientId:'CLI-A1' }, 'tok-emp')) === 'OK'; });
  t(`EmploymentStatus "${emps}": never appears in any response`, () => { const e = env([ADMIN(), EMP({ EmploymentStatus: emps })]);
    const s = JSON.stringify(['whoami','myClients'].map(a => call(e, a, {}, 'tok-emp'))).toLowerCase(); return !s.includes('employmentstatus') && !s.includes('inactive'); });
}
for (const acc of ['SUSPENDED', '', 'banana', 'INACTIVE']) {
  t(`Option A default: AccessStatus "${acc}" is ignored`, () => code(call(env([ADMIN(), EMP({ AccessStatus: acc })]), 'whoami', {}, 'tok-emp')) === 'OK'); }
t('Option A default: whoami reports suspensionEnforced=false', () => call(env([ADMIN(), EMP()]), 'whoami', {}, 'tok-emp').data.suspensionEnforced === false);
t('Option A default: INACTIVE + SUSPENDED values together still do not block', () => code(call(env([ADMIN(), EMP({ EmploymentStatus:'INACTIVE', AccessStatus:'SUSPENDED' })]), 'myClients', {}, 'tok-emp')) === 'OK');

G('8 Option B is dormant and only runs when explicitly enabled');
const ON = { ENABLE_ACCESS_STATUS: 'true' };
for (const v of ['TRUE', '1', 'yes', 'false', ' true', 'true ']) t(`flag value "${v}" does NOT enable Option B`, () => code(call(env([ADMIN(), EMP({ AccessStatus:'SUSPENDED' })], { ENABLE_ACCESS_STATUS: v }), 'whoami', {}, 'tok-emp')) === 'OK');
t('flag "true": SUSPENDED is blocked', () => code(call(env([ADMIN(), EMP({ AccessStatus:'SUSPENDED' })], ON), 'whoami', {}, 'tok-emp')) === 'ACCOUNT_SUSPENDED');
t('flag "true": ACTIVE works', () => code(call(env([ADMIN(), EMP()], ON), 'whoami', {}, 'tok-emp')) === 'OK');
t('flag "true": INACTIVE employment + ACTIVE access works', () => code(call(env([ADMIN(), EMP({ EmploymentStatus:'INACTIVE' })], ON), 'myClients', {}, 'tok-emp')) === 'OK');
t('flag "true": only AccessStatus decides', () => code(call(env([ADMIN(), EMP({ EmploymentStatus:'ACTIVE', AccessStatus:'SUSPENDED' })], ON), 'whoami', {}, 'tok-emp')) === 'ACCOUNT_SUSPENDED'
  && code(call(env([ADMIN(), EMP({ EmploymentStatus:'INACTIVE', AccessStatus:'SUSPENDED' })], ON), 'whoami', {}, 'tok-emp')) === 'ACCOUNT_SUSPENDED');
t('flag "true": empty AccessStatus denied', () => code(call(env([ADMIN(), EMP({ AccessStatus:'' })], ON), 'whoami', {}, 'tok-emp')) === 'ACCESS_STATUS_INVALID');
t('flag "true": AccessStatus INACTIVE is not valid access', () => code(call(env([ADMIN(), EMP({ AccessStatus:'INACTIVE' })], ON), 'whoami', {}, 'tok-emp')) === 'ACCESS_STATUS_INVALID');
t('flag "true": whoami reports suspensionEnforced=true', () => call(env([ADMIN(), EMP()], ON), 'whoami', {}, 'tok-emp').data.suspensionEnforced === true);

G('9 Request validation and routing');
const E0 = () => env([ADMIN(), EMP()]);
t('no request object', () => code(JSON.parse(E0().ctx.doPost().text)) === 'BAD_REQUEST');
t('no postData', () => code(raw(E0())) === 'BAD_REQUEST');
t('empty body', () => code(raw(E0(), '')) === 'BAD_REQUEST');
t('body "null"', () => code(raw(E0(), 'null')) === 'BAD_REQUEST');
t('body is an array', () => code(raw(E0(), '[]')) === 'BAD_REQUEST');
t('body is a number', () => code(raw(E0(), '5')) === 'BAD_REQUEST');
t('body is a JSON string', () => code(raw(E0(), '"whoami"')) === 'BAD_REQUEST');
t('body is not JSON', () => code(raw(E0(), '{{')) === 'BAD_REQUEST');
t('missing action', () => code(raw(E0(), JSON.stringify({ idToken:'tok-emp' }))) === 'BAD_REQUEST');
t('non-string action', () => code(raw(E0(), JSON.stringify({ action:{ a:1 }, idToken:'tok-emp' }))) === 'BAD_REQUEST');
t('empty action', () => code(raw(E0(), JSON.stringify({ action:'', idToken:'tok-emp' }))) === 'BAD_REQUEST');
for (const a of ['nope','__proto__','constructor','toString','hasOwnProperty','prototype','WHOAMI','whoami ']) t(`unknown action "${a}" -> UNKNOWN_ACTION`, () => code(call(E0(), a, {}, 'tok-emp')) === 'UNKNOWN_ACTION');
for (const [label, p] of [['string','x'],['array',[1,2]],['null',null],['number',7],['missing',undefined]]) t(`payload ${label} is normalized to an empty object`, () => {
  const r = raw(E0(), JSON.stringify({ action:'whoami', idToken:'tok-emp', payload:p })); return r.ok && r.data.ignoredFromRequest.length === 0; });
t('payload __proto__ key is dropped and cannot elevate', () => {
  const r = raw(E0(), '{"action":"whoami","idToken":"tok-emp","payload":{"__proto__":{"role":"ADMIN"},"constructor":{"x":1},"role":"ADMIN"}}');
  return r.ok && r.data.role === 'EMPLOYEE' && r.data.ignoredFromRequest.join() === 'role'; });
t('normalizePayload_ returns a plain object', () => { const o = probe.ctx.normalizePayload_([1]); return Object.getPrototypeOf(o) === Object.prototype || Object.getPrototypeOf(o).constructor !== undefined && Object.keys(o).length === 0; });
t('invalid role never reaches a business endpoint even with a valid action', () => code(call(env([ADMIN(), EMP({ Role:'GUEST' })]), 'getClient', { clientId:'CLI-B1' }, 'tok-emp')) === 'ROLE_INVALID');

G('10 Token validation (tokeninfo is mocked; the real check is your manual test)');
t('no token', () => code(call(E0(), 'whoami', {}, null)) === 'NO_TOKEN');
t('non-string token', () => code(call(E0(), 'whoami', {}, 12345)) === 'NO_TOKEN');
t('garbage token', () => code(call(E0(), 'whoami', {}, 'garbage')) === 'INVALID_TOKEN');
t('wrong audience', () => code(call(E0(), 'whoami', {}, 'tok-wrong-aud')) === 'INVALID_TOKEN');
t('wrong issuer', () => code(call(E0(), 'whoami', {}, 'tok-wrong-iss')) === 'INVALID_TOKEN');
t('expired token', () => code(call(E0(), 'whoami', {}, 'tok-expired')) === 'TOKEN_EXPIRED');
t('unverified email', () => code(call(E0(), 'whoami', {}, 'tok-unverified')) === 'INVALID_TOKEN');
t('missing sub', () => code(call(E0(), 'whoami', {}, 'tok-no-sub')) === 'INVALID_TOKEN');
t('missing GOOGLE_CLIENT_ID property', () => { const e = makeCtx({ Users: [H.slice(), ADMIN()], Clients_Canary: CLIENTS() }, { GOOGLE_CLIENT_ID: '' }); return code(call(e, 'whoami', {}, 'tok-admin')) === 'SERVER_NOT_CONFIGURED'; });

// ---- Report ----
console.log('\nResults by group:');
Object.keys(perGroup).forEach(g => console.log(`  ${g}: ${perGroup[g].pass} passed, ${perGroup[g].fail} failed`));
console.log(`\nTOTAL: ${pass} passed, ${fail} failed`);
if (fail) { console.log('\nFAILED TESTS:'); failures.forEach(f => console.log('  - ' + f)); }
process.exit(fail ? 1 : 0);
