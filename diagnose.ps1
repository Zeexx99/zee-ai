# Zee AI diagnostics - read-only. Writes diag2-log.txt next to this script.
$ErrorActionPreference = 'Continue'
$log = Join-Path $PSScriptRoot 'diag2-log.txt'
"===== Zee AI diagnostics 2 $(Get-Date) =====" | Out-File $log -Encoding utf8
function W($t) { $t | Out-File $log -Append -Encoding utf8 }

W '--- processes listening on ports 3000 and 3100'
Get-NetTCPConnection -LocalPort 3000,3100 -State Listen -ErrorAction SilentlyContinue | ForEach-Object {
  $p = Get-CimInstance Win32_Process -Filter "ProcessId=$($_.OwningProcess)"
  W ("{0}:{1} pid={2} name={3}" -f $_.LocalAddress, $_.LocalPort, $_.OwningProcess, $p.Name)
  W ("   cmd: " + $p.CommandLine)
  if ($p.ParentProcessId) { $pp = Get-CimInstance Win32_Process -Filter "ProcessId=$($p.ParentProcessId)"; W ("   parent cmd: " + $pp.CommandLine) }
}

foreach ($h in @('localhost:3100','127.0.0.1:3100','[::1]:3100')) {
  W "--- GET https://$h/taskpane.html"
  $b = (& curl.exe -k -s -m 5 "https://$h/taskpane.html") -join "`n"
  W $b.Substring(0, [Math]::Min(700, $b.Length))
}
W '--- GET /src/main.tsx (first 300 chars)'
$m = (& curl.exe -k -s -m 10 "https://localhost:3100/src/main.tsx") -join "`n"
W $m.Substring(0, [Math]::Min(300, $m.Length))
W '--- GET /src/App.tsx status'
W (& curl.exe -k -s -m 20 -o NUL -w "%{http_code} %{size_download}" "https://localhost:3100/src/App.tsx")
W '--- GET office.js status'
W (& curl.exe -s -m 10 -o NUL -w "%{http_code}" "https://appsforoffice.microsoft.com/lib/1/hosted/office.js")

W '--- registered developer add-ins'
Get-ItemProperty 'HKCU:\Software\Microsoft\Office\16.0\WEF\Developer' -ErrorAction SilentlyContinue | Out-String | W
W '--- loopback exemption for WebView'
& CheckNetIsolation.exe LoopbackExempt -s 2>&1 | Select-String -Pattern 'win32webviewhost|cw5n1h2txyewy' -Context 0,1 | Out-String | W
W '===== done ====='
