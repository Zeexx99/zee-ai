# Waits until the GitHub Pages copy of an add-in has been rebuilt after a push.
# Used by update-all.bat. Reads <site>/build.json (written at build time) until its build time is
# newer than -Since. The site address comes from the repo's GitHub remote.
param(
  [Parameter(Mandatory = $true)][string]$RepoDir,
  [Parameter(Mandatory = $true)][string]$Since,
  [string]$Name = "Zee AI",
  [int]$TimeoutSec = 900
)
$ErrorActionPreference = "Stop"
$remote = (& git -C $RepoDir remote get-url origin) 2>$null
if (-not $remote -or $remote -notmatch 'github\.com[:/]+([^/]+)/([^/]+?)(\.git)?/?$') {
  Write-Host "  [!] $Name : no GitHub remote found - skipping the online check."
  exit 2
}
$site = "https://$($Matches[1].ToLower()).github.io/$($Matches[2])/"
$inv = [Globalization.CultureInfo]::InvariantCulture
$sinceUtc = [DateTime]::Parse($Since, $inv, [Globalization.DateTimeStyles]::AdjustToUniversal).AddMinutes(-2)
$deadline = (Get-Date).AddSeconds($TimeoutSec)
Write-Host "  Waiting for $Name at $site (GitHub usually needs 1-3 minutes)..."
while ((Get-Date) -lt $deadline) {
  try {
    $r = Invoke-RestMethod -Uri ($site + "build.json?t=" + [DateTime]::UtcNow.Ticks) -Headers @{ "Cache-Control" = "no-cache" } -TimeoutSec 20
    $raw = $r.build
    if ($raw -is [DateTime]) { $built = $raw.ToUniversalTime() } else { $built = [DateTime]::Parse([string]$raw, $inv, [Globalization.DateTimeStyles]::AdjustToUniversal) }
    if ($built -ge $sinceUtc) {
      Write-Host "  [OK] $Name is live - version $($r.version), built $($built.ToLocalTime().ToString('g'))"
      exit 0
    }
    Write-Host "  ... $Name : GitHub is still building (online copy is from $($built.ToLocalTime().ToString('g')))"
  } catch {
    Write-Host "  ... $Name : not reachable yet ($($_.Exception.Message))"
  }
  Start-Sleep -Seconds 20
}
Write-Host "  [!] $Name : still not live after $([int]($TimeoutSec / 60)) minutes. Open the repo's Actions page on GitHub to see why."
exit 1
