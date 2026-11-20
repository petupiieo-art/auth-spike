# Auth prototype change log

## Revision 4 (CODE_VERSION auth-spike-r4)
- One schema everywhere: `USERS_HEADERS` in Code.gs = Email, GoogleSub, Role, EmployeeID, EmploymentStatus, AccessStatus. Setup, authorization and test fixtures all use it. The old `Status` column is refused (USER_CONFIG_ERROR).
- Strict roles: only ADMIN and EMPLOYEE; everything else ROLE_INVALID before routing. Router refuses unknown roles again (second layer). getClient has explicit allow rules.
- Duplicate Email or GoogleSub anywhere in Users: USER_CONFIG_ERROR (fail closed, never "first match wins"). Matched row needs a valid Email, an exact Role, and for EMPLOYEE a non-empty EmployeeID.
- Option A is the default: AccessStatus is ignored. Option B (suspension) is dormant and runs only if Script Property ENABLE_ACCESS_STATUS is exactly `true`. EmploymentStatus is never read or returned.
- Request validation: null/array/number/string/empty/missing body or missing/non-string action -> BAD_REQUEST; payload normalized to a plain object (unsafe keys dropped); unknown action -> UNKNOWN_ACTION.
- Identity is the Google `sub`; email only binds a new account once (unchanged from revision 3).
- tests/logic-test.js: 177 isolated checks in 10 groups; one failing check cannot crash the run; refuses to run against an outdated Code.gs; prints the Code.gs sha256.
- Folder `frontend/` renamed `docs/`. Guide delivered as PDF in `guides/`.

## Revision 3
- Identity by Google `sub`; suspension removed; EmploymentStatus never returned; token validation documented.

## Revision 2
- Exact role validation, default-deny policy table, explicit getClient rules, blank EmployeeID refused, duplicate rows refused, BAD_REQUEST handling.

## Revision 1
- First prototype (single Status column).
