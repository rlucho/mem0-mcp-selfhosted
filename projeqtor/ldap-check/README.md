# ProjeQtor — how to check the LDAP connection

A runbook for answering "is our ProjeQtor talking to Active Directory, and if not, where does it
break?", plus a drop-in checker that replays ProjeQtor's own login flow step by step.

Everything below was verified against ProjeQtor source (`model/User.php`, `model/Parameter.php`,
`db/maintenance.php`, `tool/config.php`, `tool/projeqtor.php`) and the current user guide.
Your installed version is authoritative — the same code lives in your web folder, so you can
confirm any claim here by opening `model/User.php` on the server and searching for `ldap`.

> No internal host names, domains or account names are recorded in this folder.
> Put those in `config.local.php`, which is git-ignored.

## 0. What ProjeQtor actually does when LDAP is on

`User::authenticate($login, $password)` in `model/User.php`:

| # | Step | Failure returns |
|---|------|-----------------|
| 1 | `function_exists('ldap_connect')` — the **php_ldap extension** must be enabled | `"ldap"` |
| 2 | `ldap_connect(paramLdap_host, paramLdap_port)` | `"ldap"` |
| 3 | `LDAP_OPT_PROTOCOL_VERSION = paramLdap_version`, `LDAP_OPT_REFERRALS = 0` | — |
| 4 | `ldap_bind(paramLdap_search_user, paramLdap_search_pass)` — the **service account** | `"ldap"` |
| 5 | `ldap_search(paramLdap_base_dn, filter)` where `%USERNAME%` in `paramLdap_user_filter` is replaced by the typed login | `"login"` |
| 6 | the search must return **exactly one** entry — 0 or 2+ both fail | `"login"` |
| 7 | `ldap_bind(dn_of_that_entry, typed_password)` | `"login"` |

Two consequences worth internalising:

- **`"ldap"` vs `"login"` on the login screen is a real diagnostic.** `"ldap"` means the problem is
  the extension, the server, or the service account. `"login"` means the connection worked and the
  problem is the base DN, the filter, or the user's own password.
- **A user that does not exist in ProjeQtor yet is auto-created** on a successful LDAP login
  (step 7 onwards): `isLdap=1`, `email` from the entry's `mail`, `resourceName` from `cn`, profile
  from `ldapDefaultProfile`. That is what the `isLdap` column in the HR user-import CSV
  (`name;userName;email;...;isLdap`) is flagging — `isLdap=1` users have **no local password** and
  authenticate only against the directory.

## 1. Read the current configuration (zero risk, do this first)

**In the UI:** *Configuration → Global parameters → Authentication*, section
*LDAP management parameters* — connection with LDAP user (on/off), base dn, host, port, version,
LDAP user, LDAP password, LDAP user filter, plus default profile for LDAP users, message on
creation of new user from LDAP, actions on LDAP user creation, project to allocate automatically.

**In the database (authoritative — the UI masks the password).** Since V3.0.0 these parameters live
in the `parameter` table, not in a file. In phpMyAdmin:

```sql
SELECT parameterCode, parameterValue
FROM   `parameter`                 -- prefix it if paramDbPrefix is set
WHERE  idUser IS NULL AND idProject IS NULL
  AND (parameterCode LIKE 'paramLdap%' OR parameterCode LIKE 'ldap%')
ORDER BY parameterCode;
```

Expect: `paramLdap_allow_login`, `paramLdap_base_dn`, `paramLdap_host`, `paramLdap_port`,
`paramLdap_version`, `paramLdap_search_user`, `paramLdap_search_pass`, `paramLdap_user_filter`,
`ldapDefaultProfile`, `ldapMsgOnUserCreation`.

**No rows at all, or `paramLdap_allow_login = false`** ⇒ LDAP login is simply off; every user
authenticates against ProjeQtor's own `user` table and the directory is never contacted.

Cross-check who would be affected:

```sql
SELECT id, name, email, isLdap FROM `user` WHERE isLdap = 1 AND idle = 0;
```

**Which file holds what** (in the ProjeQtor web folder):

- `files/config/parameters.php` — DB credentials, `logFile`, `logLevel`. **Not** the LDAP values.
  (Older installs: `tool/parameters.php`; if `tool/parametersLocation.php` exists it names the real path.)
- the `parameter` DB table — all LDAP values.

## 2. Check the php_ldap extension

This is the single most common reason a correct configuration still fails, and the hardest to fix
without admin rights: **an extension cannot be enabled from `.htaccess`** — it needs
`extension=ldap` uncommented in `php.ini` *and* an Apache restart.

Quickest check without touching anything: the checker in this folder reports it as step 1.
Otherwise drop a one-liner in the web folder and delete it after:

```php
<?php var_dump(extension_loaded('ldap')); echo php_ini_loaded_file();
```

## 3. Run the checker

`projeqtor_ldap_check.php` reads ProjeQtor's *own* parameters (parameters file → DB) and replays
steps 1–7, reporting OK/FAIL per step with the LDAP error, the AD diagnostic message and the
decoded AD reason code.

1. `cp config.local.php.example config.local.php` and set `token` to a random string.
2. Copy **both** files into the ProjeQtor web folder (next to `index.php`).
3. Open `http://<server>/<projeqtor>/projeqtor_ldap_check.php?token=<your token>`.
4. Enter a real login name (and optionally that user's password) and run.
5. **Delete both files.**

It never writes anything, never changes ProjeQtor's configuration, and never echoes or logs a
password. Without the matching token it returns 404.

Useful extras:

- **`override`** in `config.local.php` lets you test a different host / base DN / filter / service
  account *without* changing ProjeQtor's configuration. Anything you override is flagged in the
  report so you don't mistake a manual test for the live config.
- **"probe base dn / filter candidates"** tries every `candidate_base_dns` × `candidate_filters`
  combination for one username and prints the hit count of each — the fast way to find the
  combination that returns exactly 1.

CLI mode exists (`php projeqtor_ldap_check.php --token=… --user=…`) but **running `php.exe` from a
file share executes on your own workstation, not on the server** — that tests the wrong network
path. On a UNC-mounted XAMPP, use the browser.

## 4. Turn on ProjeQtor's own log

ProjeQtor writes LDAP failures itself, via `traceLog()` (level 2) and `errorLog()` (level 1). In
`files/config/parameters.php`:

```php
$logFile  = '../files/logs/projeqtor_${date}.log';   // ${date} expands to YYYYMMDD
$logLevel = 2;                                        // must be >= 2 to catch traceLog
```

Back the file up first, set it back when you're done, and make sure the log directory is writable
and not web-readable. Messages to grep for: `LDAP connection error`, `LdapBind Error`,
`Ldap not installed on your PHP server`.

## 5. Common Active Directory gotchas

- **Anonymous bind.** AD refuses anonymous searches by default. Leaving `paramLdap_search_user`
  empty makes step 4 an anonymous bind that may "succeed" while step 5 returns nothing. You need a
  service account. Its value may be a full DN *or* a UPN (`svc-account@corp.example`) — PHP passes
  it straight to `ldap_bind`, and AD accepts both.
- **Exactly one match.** Step 6 fails on 2+ matches just as hard as on 0. `(cn=%USERNAME%)` style
  filters are prone to this; `sAMAccountName` is unique per domain and is the usual choice.
- **Base DN coverage.** The base DN must sit *above* the OU that actually holds the accounts.
  `CN=Users,DC=…` only covers the default container — accounts in a custom OU need that OU (or the
  domain root) instead.
- **Plaintext.** Port 389 sends the service-account password and every user password in the clear.
  Prefer `ldaps://` on 636 — put the scheme in the host field; ProjeQtor passes it to
  `ldap_connect` unchanged.
- **Referrals** are already disabled by ProjeQtor (`LDAP_OPT_REFERRALS = 0`), which is what AD needs.
- **AD hides the real reason** behind "invalid credentials". The extended diagnostic carries
  `data <hex>`; the checker decodes it: `525` no such user, `52e` wrong password, `530`/`531` logon
  time/workstation restriction, `532` password expired, `533` account disabled, `701` account
  expired, `773` must change password, `775` locked out.

## 6. What a standalone `ldap_bind` test script does and doesn't prove

The common "testldap.php" pattern (build `user@DOMAIN_FQDN`, `ldap_connect`, `ldap_bind`, search)
is a **UPN simple bind** — a different flow from ProjeQtor's. Run from the ProjeQtor server it does
prove:

- the DC is reachable on that port from that host (firewall path OK), and
- that one user's credentials are valid.

It does **not** prove anything about the service account, the base DN or the user filter — the three
things ProjeQtor actually needs. Note also that such scripts typically:

- pass `array($conn, $conn)` to `ldap_search`, i.e. a parallel search over the same link twice, and
  then `count()` a search-result object — neither does what it looks like;
- accept an **empty password**, which PHP turns into an *unauthenticated* bind that AD reports as
  success — a blank password appears to "log in";
- run `display_errors` on, and are an unauthenticated password form sitting in the web root.

Treat them as one-shot diagnostics and delete them from the web folder immediately after use.

## 7. If LDAP was never configured — what to request from IT

- A **service account** for directory reads: its DN or UPN, and the password.
- The **base DN** that contains the user accounts in scope.
- The **login attribute** users should type (`sAMAccountName` vs `userPrincipalName`), which fixes
  the filter.
- Whether **LDAPS (636)** is available, and whether plaintext 389 is even permitted.
- **php_ldap enabled** on the PHP install plus an Apache restart, if step 2 says it's missing.

Then set `paramLdap_allow_login = true`, `ldapDefaultProfile` to the profile new LDAP users should
get, and decide `ldapMsgOnUserCreation` before the first login creates accounts silently.
