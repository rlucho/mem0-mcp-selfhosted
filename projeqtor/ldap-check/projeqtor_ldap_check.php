<?php
/**
 * ProjeQtor LDAP connection checker
 * ---------------------------------
 * Replays ProjeQtor's OWN LDAP login flow (model/User.php -> User::authenticate())
 * step by step, against ProjeQtor's OWN stored parameters, and reports exactly
 * which step fails.
 *
 * ProjeQtor's flow, for reference:
 *   1. php_ldap extension must exist            -> else login returns "ldap"
 *   2. ldap_connect(host, port)                 -> else "ldap"
 *   3. LDAP_OPT_PROTOCOL_VERSION / REFERRALS=0
 *   4. ldap_bind(search_user, search_pass)      -> else "ldap"   (service account)
 *   5. ldap_search(base_dn, filter)             -> else "login"  (%USERNAME% substituted)
 *   6. exactly ONE entry must match             -> else "login"  (0 or >1 both fail)
 *   7. ldap_bind(found_dn, typed_password)      -> else "login"
 *
 * USAGE (web, the meaningful test - runs ON the ProjeQtor server):
 *   1. copy config.local.php.example -> config.local.php, fill in a token
 *   2. drop both files in the ProjeQtor web folder (next to index.php)
 *   3. open  http://<server>/<projeqtor>/projeqtor_ldap_check.php?token=YOURTOKEN
 *   4. DELETE BOTH FILES when finished
 *
 * USAGE (CLI):
 *   php projeqtor_ldap_check.php --token=YOURTOKEN [--user=someone] [--password=...]
 *   WARNING: running php.exe from a UNC share executes on YOUR workstation, not on
 *   the server - it tests the wrong network path. Use the web mode for the real answer.
 *
 * This script never writes anywhere, never changes ProjeQtor config, and never
 * logs or echoes a password.
 */

// --------------------------------------------------------------------------- config
$CFG = array(
    // REQUIRED - any random string. Without a matching ?token= the page 404s.
    'token'           => '',
    // Optional: explicit path to ProjeQtor's parameters.php. Empty = auto-locate.
    'parameters_file' => '',
    // Optional: ProjeQtor install root. Empty = directory this file sits in.
    'projeqtor_root'  => '',
    // Set false to remove the password field entirely (config inspection only).
    'allow_user_bind' => true,
    // Seconds before a blocked port gives up instead of hanging the page.
    'network_timeout' => 5,
    // Leave empty to use ProjeQtor's live values; fill in to try something else
    // WITHOUT touching ProjeQtor's configuration.
    'override' => array(
        'paramLdap_host'        => '',
        'paramLdap_port'        => '',
        'paramLdap_version'     => '',
        'paramLdap_base_dn'     => '',
        'paramLdap_user_filter' => '',
        'paramLdap_search_user' => '',
        'paramLdap_search_pass' => '',
    ),
    // Probed by the "discovery" button when you do not yet know the right values.
    'candidate_base_dns' => array(),
    'candidate_filters'  => array(
        '(&(objectClass=user)(sAMAccountName=%USERNAME%))',
        '(&(objectCategory=person)(objectClass=user)(sAMAccountName=%USERNAME%))',
        '(sAMAccountName=%USERNAME%)',
        '(userPrincipalName=%USERNAME%)',
    ),
);
if (is_file(__DIR__ . '/config.local.php')) {
    $local = require __DIR__ . '/config.local.php';
    if (is_array($local)) { $CFG = array_replace_recursive($CFG, $local); }
}

$IS_CLI = (PHP_SAPI === 'cli');

// --------------------------------------------------------------------------- helpers
$REPORT = array();
function step($status, $label, $detail = '') {
    global $REPORT;
    $REPORT[] = array('status' => $status, 'label' => $label, 'detail' => $detail);
}
function mask($v) {
    if ($v === null || $v === '') { return '(empty)'; }
    return '(set, ' . strlen($v) . ' chars)';
}
function h($s) { return htmlspecialchars((string)$s, ENT_QUOTES, 'UTF-8'); }

/** AD stuffs the real reason into the diagnostic message as "data <hex>". */
function adReason($diagnostic) {
    if (!$diagnostic) { return ''; }
    if (!preg_match('/data\s+([0-9a-fA-F]+)/', $diagnostic, $m)) { return ''; }
    $codes = array(
        '525' => 'user not found in the directory',
        '52e' => 'invalid credentials (wrong password)',
        '52f' => 'account restriction',
        '530' => 'not permitted to log on at this time',
        '531' => 'not permitted to log on at this workstation',
        '532' => 'password expired',
        '533' => 'account disabled',
        '568' => 'too many context IDs',
        '701' => 'account expired',
        '773' => 'user must reset password',
        '775' => 'account locked out',
    );
    $code = strtolower($m[1]);
    return isset($codes[$code]) ? 'AD data ' . $code . ': ' . $codes[$code]
                                : 'AD data ' . $code . ': unmapped code';
}

function ldapDiag($cnx) {
    $out = array();
    $err = @ldap_error($cnx);
    if ($err) { $out[] = 'ldap_error: ' . $err; }
    if (defined('LDAP_OPT_DIAGNOSTIC_MESSAGE')) {
        $msg = '';
        if (@ldap_get_option($cnx, LDAP_OPT_DIAGNOSTIC_MESSAGE, $msg) && $msg) {
            $out[] = 'diagnostic: ' . trim($msg);
            $reason = adReason($msg);
            if ($reason) { $out[] = $reason; }
        }
    }
    return implode(' | ', $out);
}

/** ProjeQtor calls ldap_connect($host, $port). Handle ldap:// / ldaps:// URIs too. */
function connectLdap($host, $port, $timeout) {
    if ($host === '') { return array(null, 'no host configured'); }
    $isUri = (bool)preg_match('#^ldaps?://#i', $host);
    if ($isUri) {
        $uri = rtrim($host, '/');
        // append the port only when the URI does not already carry one
        if ($port !== '' && !preg_match('#:\d+$#', $uri)) { $uri .= ':' . (int)$port; }
        $cnx = @ldap_connect($uri);
    } else {
        $cnx = ($port === '') ? @ldap_connect($host) : @ldap_connect($host, (int)$port);
    }
    if (!$cnx) { return array(null, 'ldap_connect() returned false'); }
    if (defined('LDAP_OPT_NETWORK_TIMEOUT')) {
        @ldap_set_option($cnx, LDAP_OPT_NETWORK_TIMEOUT, $timeout);
    }
    return array($cnx, '');
}

/** Read ProjeQtor's parameters.php the same way tool/projeqtor.php does. */
function loadParametersFile($file) {
    $__before = array_keys(get_defined_vars());
    include $file;
    $vars = get_defined_vars();
    unset($vars['__before'], $vars['file']);
    return $vars;
}

function locateParametersFile($cfg, $root) {
    if ($cfg['parameters_file'] !== '') {
        return is_file($cfg['parameters_file'])
            ? array($cfg['parameters_file'], 'configured explicitly')
            : array(null, 'configured path not found: ' . $cfg['parameters_file']);
    }
    // ProjeQtor honours tool/parametersLocation.php first
    $loc = $root . '/tool/parametersLocation.php';
    if (is_file($loc)) {
        $parametersLocation = '';
        include $loc;
        if ($parametersLocation && is_file($parametersLocation)) {
            return array($parametersLocation, 'via tool/parametersLocation.php');
        }
        return array(null, 'tool/parametersLocation.php points at a missing file: ' . $parametersLocation);
    }
    foreach (array('/files/config/parameters.php', '/tool/parameters.php') as $rel) {
        if (is_file($root . $rel)) { return array($root . $rel, 'found at ' . ltrim($rel, '/')); }
    }
    return array(null, 'not found under ' . $root);
}

// --------------------------------------------------------------------------- input
if ($IS_CLI) {
    $argv = isset($argv) ? $argv : array();
    $opts = array();
    foreach ($argv as $a) {
        if (preg_match('/^--([a-z_]+)=(.*)$/i', $a, $m)) { $opts[$m[1]] = $m[2]; }
    }
    $token    = isset($opts['token'])    ? $opts['token']    : '';
    $username = isset($opts['user'])     ? $opts['user']     : '';
    $password = isset($opts['password']) ? $opts['password'] : '';
    $discover = isset($opts['discover']);
} else {
    $token    = isset($_REQUEST['token'])   ? $_REQUEST['token']   : '';
    $username = isset($_POST['username'])   ? $_POST['username']   : '';
    $password = isset($_POST['password'])   ? $_POST['password']   : '';
    $discover = isset($_POST['discover']);
}

// --------------------------------------------------------------------------- gate
if ($CFG['token'] === '') {
    $msg = "REFUSING TO RUN: no token configured.\n"
         . "Copy config.local.php.example to config.local.php and set a random 'token',\n"
         . "then call this page with ?token=<that value>.\n";
    if ($IS_CLI) { fwrite(STDERR, $msg); exit(2); }
    header('Content-Type: text/plain; charset=utf-8');
    echo $msg; exit;
}
if (!hash_equals((string)$CFG['token'], (string)$token)) {
    if ($IS_CLI) { fwrite(STDERR, "bad or missing --token\n"); exit(2); }
    header('HTTP/1.1 404 Not Found');
    echo 'Not Found'; exit;
}

// --------------------------------------------------------------------------- checks
$root = $CFG['projeqtor_root'] !== '' ? rtrim($CFG['projeqtor_root'], '/\\') : __DIR__;
$params = array();      // LDAP parameters as ProjeQtor would read them
$paramSource = 'none';

// 1. php_ldap
if (function_exists('ldap_connect')) {
    step('OK', 'php_ldap extension loaded', 'ldap_connect() exists');
} else {
    step('FAIL', 'php_ldap extension NOT loaded',
        'ProjeQtor logs "Ldap not installed on your PHP server" and the login returns "ldap". '
      . 'Enable extension=ldap in php.ini and restart Apache. php.ini in use: ' . php_ini_loaded_file());
}

// 2. locate ProjeQtor's parameters file (holds the DB credentials)
list($pfile, $pwhy) = locateParametersFile($CFG, $root);
if ($pfile) {
    step('OK', 'ProjeQtor parameters file located', $pfile . '  (' . $pwhy . ')');
} else {
    step('WARN', 'ProjeQtor parameters file not located', $pwhy
        . '. Set projeqtor_root or parameters_file in config.local.php, or use the override block.');
}

// 3. read the LDAP parameters out of ProjeQtor's database
if ($pfile) {
    $pv = loadParametersFile($pfile);
    $dbHost = isset($pv['paramDbHost'])     ? $pv['paramDbHost']     : '';
    $dbName = isset($pv['paramDbName'])     ? $pv['paramDbName']     : '';
    $dbUser = isset($pv['paramDbUser'])     ? $pv['paramDbUser']     : '';
    $dbPass = isset($pv['paramDbPassword']) ? $pv['paramDbPassword'] : '';
    $prefix = isset($pv['paramDbPrefix'])   ? $pv['paramDbPrefix']   : '';
    $logFile  = isset($pv['logFile'])  ? $pv['logFile']  : '(not set)';
    $logLevel = isset($pv['logLevel']) ? $pv['logLevel'] : '(not set)';
    step('INFO', 'ProjeQtor log settings (from parameters file)',
        'logFile=' . $logFile . '   logLevel=' . $logLevel
      . '   -- LDAP failures are written by traceLog() at level 2, so logLevel must be >= 2 and logFile must be set.');

    if (!class_exists('PDO')) {
        step('WARN', 'PDO not available', 'Cannot read the parameter table from here; use phpMyAdmin (SQL in the README) or the override block.');
    } else {
        $port = '3306';
        if (strpos($dbHost, ':') !== false) { list($dbHost, $port) = explode(':', $dbHost, 2); }
        try {
            $dsn = 'mysql:host=' . $dbHost . ';port=' . (int)$port . ';dbname=' . $dbName;
            $pdo = new PDO($dsn, $dbUser, $dbPass, array(PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION));
            $sql = 'SELECT parameterCode, parameterValue FROM `' . $prefix . 'parameter` '
                 . 'WHERE idUser IS NULL AND idProject IS NULL '
                 . "AND (parameterCode LIKE 'paramLdap%' OR parameterCode LIKE 'ldap%')";
            foreach ($pdo->query($sql) as $row) {
                $params[$row['parameterCode']] = $row['parameterValue'];
            }
            $paramSource = 'database (' . $prefix . 'parameter)';
            step('OK', 'Read LDAP parameters from ProjeQtor database',
                count($params) . ' parameter row(s) found in ' . $prefix . 'parameter');
        } catch (Exception $e) {
            step('WARN', 'Could not read the ProjeQtor database', $e->getMessage()
                . ' -- fall back to phpMyAdmin (SQL in the README) or the override block.');
        }
    }
    // pre-V3 installs kept the LDAP values in the parameters file itself
    foreach (array('paramLdap_allow_login','paramLdap_base_dn','paramLdap_host','paramLdap_port',
                   'paramLdap_version','paramLdap_search_user','paramLdap_search_pass',
                   'paramLdap_user_filter') as $code) {
        if (!isset($params[$code]) && isset($pv[$code]) && $pv[$code] !== '') {
            $params[$code] = $pv[$code];
            $paramSource = ($paramSource === 'none') ? 'parameters file' : $paramSource . ' + parameters file';
        }
    }
}

// 4. apply overrides
$overridden = array();
foreach ($CFG['override'] as $k => $v) {
    if ($v !== '') { $params[$k] = $v; $overridden[] = $k; }
}
if ($overridden) {
    step('INFO', 'Manual overrides in effect', implode(', ', $overridden)
        . ' -- these are NOT what ProjeQtor uses, they only test alternatives.');
    $paramSource = ($paramSource === 'none') ? 'override only' : $paramSource . ' + override';
}

$get = function ($k, $default = '') use ($params) {
    return (isset($params[$k]) && $params[$k] !== null) ? (string)$params[$k] : $default;
};
$allow   = strtolower($get('paramLdap_allow_login', 'false'));
$host    = $get('paramLdap_host');
$port    = $get('paramLdap_port', '389');
$version = $get('paramLdap_version', '3');
$baseDn  = $get('paramLdap_base_dn');
$filter  = $get('paramLdap_user_filter');
$suser   = $get('paramLdap_search_user');
$spass   = $get('paramLdap_search_pass');

step('INFO', 'Effective LDAP configuration (source: ' . $paramSource . ')',
      "paramLdap_allow_login = " . ($allow !== '' ? $allow : '(empty)')
    . "\nparamLdap_host        = " . ($host ?: '(empty)')
    . "\nparamLdap_port        = " . ($port ?: '(empty)')
    . "\nparamLdap_version     = " . ($version ?: '(empty)')
    . "\nparamLdap_base_dn     = " . ($baseDn ?: '(empty)')
    . "\nparamLdap_user_filter = " . ($filter ?: '(empty)')
    . "\nparamLdap_search_user = " . ($suser ?: '(empty)')
    . "\nparamLdap_search_pass = " . mask($spass)
    . "\nldapDefaultProfile    = " . $get('ldapDefaultProfile', '(not set)')
    . "\nldapMsgOnUserCreation = " . $get('ldapMsgOnUserCreation', '(not set)'));

// 5. static sanity
if ($allow === 'true') {
    step('OK', 'LDAP login is ENABLED', 'paramLdap_allow_login = true');
} else {
    step('WARN', 'LDAP login is DISABLED',
        'paramLdap_allow_login is "' . $allow . '". ProjeQtor authenticates every user against its own '
      . 'database and never touches LDAP. Users flagged isLdap=1 cannot log in while this is false.');
}
if ($filter !== '' && strpos($filter, '%USERNAME%') === false) {
    step('FAIL', 'User filter has no %USERNAME% placeholder',
        'ProjeQtor substitutes %USERNAME% with the typed login. Without it the search matches the wrong set.');
}
if ($host === '' || $baseDn === '' || $filter === '') {
    step('WARN', 'Incomplete configuration', 'host, base dn and user filter are all required for the flow to work.');
}

// 6. live connection
$cnx = null;
if (function_exists('ldap_connect') && $host !== '') {
    list($cnx, $err) = connectLdap($host, $port, $CFG['network_timeout']);
    if ($cnx) {
        step('OK', 'ldap_connect() accepted', $host . ':' . $port
            . '  (note: ldap_connect does not open a socket - the bind below is the real reachability test)');
        @ldap_set_option($cnx, LDAP_OPT_PROTOCOL_VERSION, (int)$version);
        @ldap_set_option($cnx, LDAP_OPT_REFERRALS, 0);
    } else {
        step('FAIL', 'ldap_connect() failed', $err . ' -- ProjeQtor login would return "ldap".');
    }
}

// 7. service-account bind
$bound = false;
if ($cnx) {
    $bindDn = ($suser === '') ? null : $suser;
    $bindPw = ($spass === '') ? null : $spass;
    $bound  = @ldap_bind($cnx, $bindDn, $bindPw);
    if ($bound) {
        step('OK', 'Service bind succeeded',
            ($bindDn === null ? 'ANONYMOUS bind (no search user configured)' : 'as ' . $bindDn)
          . ($bindDn === null ? ' -- Active Directory normally refuses anonymous searches, so expect the search below to fail.' : ''));
    } else {
        step('FAIL', 'Service bind FAILED', ldapDiag($cnx)
            . ' -- this is the step that makes ProjeQtor return "ldap". '
            . 'Check paramLdap_search_user / paramLdap_search_pass, and that the server is reachable on '
            . $host . ':' . $port . ' from THIS machine.');
    }
}

// 8. search for the user
$foundDn = '';
if ($cnx && $bound && $username !== '' && $baseDn !== '' && $filter !== '') {
    $applied = html_entity_decode(str_replace('%USERNAME%', $username, $filter), ENT_COMPAT, 'UTF-8');
    $res = @ldap_search($cnx, $baseDn, $applied);
    if (!$res) {
        step('FAIL', 'ldap_search() failed', 'filter: ' . $applied . ' | base: ' . $baseDn . ' | ' . ldapDiag($cnx)
            . ' -- ProjeQtor login would return "login".');
    } else {
        $entries = @ldap_get_entries($cnx, $res);
        $count = isset($entries['count']) ? (int)$entries['count'] : 0;
        if ($count === 0) {
            step('FAIL', 'Search matched 0 entries', 'filter: ' . $applied . ' | base: ' . $baseDn
                . ' -- ProjeQtor returns "login". Either the base dn does not cover this user, or the filter '
                . 'attribute is wrong for this directory.');
        } elseif ($count > 1) {
            step('FAIL', 'Search matched ' . $count . ' entries',
                'ProjeQtor REQUIRES exactly one match and returns "login" otherwise. Tighten the filter.');
        } else {
            $first = $entries[0];
            $foundDn = $first['dn'];
            $attrs = array();
            foreach (array('cn', 'mail', 'samaccountname', 'userprincipalname', 'displayname') as $a) {
                if (isset($first[$a][0])) { $attrs[] = $a . '=' . $first[$a][0]; }
            }
            step('OK', 'Search matched exactly 1 entry',
                'dn: ' . $foundDn . "\n" . implode("\n", $attrs)
              . "\n-- ProjeQtor would create the account with email=mail[0] and resourceName=cn[0] if it does not exist yet.");
        }
    }
} elseif ($cnx && $bound && $username === '') {
    step('INFO', 'No username supplied', 'Enter a login name to test the search + user bind steps.');
}

// 9. bind as the user
if ($foundDn !== '' && $CFG['allow_user_bind'] && $password !== '') {
    $ok = @ldap_bind($cnx, $foundDn, $password);
    if ($ok) {
        step('OK', 'User bind succeeded', 'ProjeQtor would log this user in.');
    } else {
        step('FAIL', 'User bind FAILED', ldapDiag($cnx) . ' -- ProjeQtor returns "login".');
    }
} elseif ($foundDn !== '' && $password === '') {
    step('INFO', 'No password supplied', 'The user-bind step was skipped. '
        . 'Note: an EMPTY password would produce an unauthenticated bind that Active Directory reports as success - '
        . 'this script refuses it on purpose.');
}

// 10. discovery
if ($discover && $cnx && $bound && $username !== '') {
    $bases = $CFG['candidate_base_dns'];
    if ($baseDn !== '' && !in_array($baseDn, $bases, true)) { array_unshift($bases, $baseDn); }
    $lines = array();
    foreach ($bases as $b) {
        foreach ($CFG['candidate_filters'] as $f) {
            $applied = str_replace('%USERNAME%', $username, $f);
            $r = @ldap_search($cnx, $b, $applied);
            if (!$r) { $lines[] = sprintf('%-3s %s  %s   (search error: %s)', 'ERR', $b, $applied, @ldap_error($cnx)); continue; }
            $e = @ldap_get_entries($cnx, $r);
            $n = isset($e['count']) ? (int)$e['count'] : 0;
            $lines[] = sprintf('%-3d %s  %s%s', $n, $b, $applied, $n === 1 ? '   <== usable' : '');
        }
    }
    step('INFO', 'Discovery: matches per base dn / filter combination',
        $lines ? "hits base-dn  filter\n" . implode("\n", $lines)
               : 'No candidate base dns configured - add them to candidate_base_dns in config.local.php.');
}

if ($cnx) { @ldap_unbind($cnx); }

// --------------------------------------------------------------------------- output
$counts = array('OK' => 0, 'FAIL' => 0, 'WARN' => 0, 'INFO' => 0);
foreach ($REPORT as $r) { $counts[$r['status']]++; }
$summary = sprintf('%d OK / %d FAIL / %d WARN', $counts['OK'], $counts['FAIL'], $counts['WARN']);

if ($IS_CLI) {
    foreach ($REPORT as $r) {
        echo '[' . str_pad($r['status'], 4) . '] ' . $r['label'] . "\n";
        if ($r['detail'] !== '') {
            foreach (explode("\n", $r['detail']) as $line) { echo '        ' . $line . "\n"; }
        }
    }
    echo "\n" . $summary . "\n";
    exit($counts['FAIL'] > 0 ? 1 : 0);
}

header('Content-Type: text/html; charset=utf-8');
?>
<!DOCTYPE html>
<html><head><meta charset="utf-8"><title>ProjeQtor LDAP check</title>
<style>
 body{font:13px/1.5 Consolas,Menlo,monospace;margin:24px;max-width:1100px;color:#222}
 h1{font-size:18px} h2{font-size:14px;margin-top:24px}
 .row{border-left:4px solid #ccc;padding:6px 10px;margin:6px 0;background:#fafafa}
 .OK{border-color:#2e7d32} .FAIL{border-color:#c62828;background:#fff5f5}
 .WARN{border-color:#ef6c00;background:#fffaf3} .INFO{border-color:#1565c0}
 .tag{display:inline-block;min-width:44px;font-weight:bold}
 .OK .tag{color:#2e7d32} .FAIL .tag{color:#c62828} .WARN .tag{color:#ef6c00} .INFO .tag{color:#1565c0}
 pre{margin:4px 0 0 52px;white-space:pre-wrap;color:#444}
 form{margin:16px 0;padding:12px;background:#f0f4f8;border:1px solid #ccd}
 input[type=text],input[type=password]{font:inherit;padding:3px 6px;width:240px}
 .warn{background:#fff3cd;border:1px solid #ffe08a;padding:10px;margin:16px 0}
</style></head><body>
<h1>ProjeQtor LDAP check &mdash; <?php echo h($summary); ?></h1>
<div class="warn"><strong>Delete this file (and config.local.php) from the web folder when you are done.</strong>
It accepts a password and reports directory contents.</div>
<form method="post" action="?token=<?php echo h($token); ?>">
  <label>Login name to test: <input type="text" name="username" value="<?php echo h($username); ?>" autocomplete="off"></label>
  <?php if ($CFG['allow_user_bind']): ?>
  <label>Password: <input type="password" name="password" autocomplete="off"></label>
  <?php endif; ?>
  <label><input type="checkbox" name="discover" value="1" <?php echo $discover ? 'checked' : ''; ?>> probe base dn / filter candidates</label>
  <button type="submit">Run check</button>
</form>
<?php foreach ($REPORT as $r): ?>
<div class="row <?php echo $r['status']; ?>">
  <span class="tag"><?php echo $r['status']; ?></span> <?php echo h($r['label']); ?>
  <?php if ($r['detail'] !== ''): ?><pre><?php echo h($r['detail']); ?></pre><?php endif; ?>
</div>
<?php endforeach; ?>
<h2>How ProjeQtor maps these failures</h2>
<pre style="margin-left:0">"ldap"  on the login screen  = php_ldap missing, connect failed, or the SERVICE bind failed
"login" on the login screen  = search failed, matched 0 or &gt;1 entries, or the USER bind failed
Both are written to ProjeQtor's log by traceLog() - level 2, so logLevel must be &gt;= 2.</pre>
</body></html>
