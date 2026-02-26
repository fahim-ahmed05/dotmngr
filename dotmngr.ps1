<#
dotmngr.ps1 — config-driven dotfile manager for Windows.

Config (REQUIRED):
{
  "global": { "mode": "junction", "trash": true, "trashDir": "%LOCALAPPDATA%\\DotLinksTrash" },
  "packages": {
    "nvim": {
      "enabled": true,
      "mode": "junction",
      "items": [
        { "from": "%USERPROFILE%\\Git\\dotfiles\\nvim", "to": "%LOCALAPPDATA%\\nvim" }
      ]
    }
  }
}

Notes:
- This script intentionally DOES NOT support "~" expansion.
- Use environment variables in config instead:
  - %USERPROFILE%, %LOCALAPPDATA%, %APPDATA%, etc.
- Symbolic links may require Admin unless Developer Mode is enabled.
- State stored in: %USERPROFILE%\.config\dotmngr\state.<configName>.json

Modes:
  symlink   - Symbolic link (file/dir)
  junction  - Junction (dir)
  hardlink  - Hard link (file, same volume)
  copy      - Robocopy sync (skips overwriting newer destination via /XO)
  copyOnce  - Copy only if destination doesn’t exist  shortcut  - Windows .lnk 
  shortcut  - Windows .lnk shortcut (file/dir)
Switches:
  -Unlink   Remove all managed links for the selected packages
  -Status   Show a table of all tracked items and whether they are intact
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory=$true)]
  [ValidateNotNullOrEmpty()]
  [string]$ConfigPath,

  [Parameter()]
  [string[]]$Package,

  [Parameter()]
  [switch]$Unlink,

  [Parameter()]
  [switch]$Status
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ---------------- Path + filesystem helpers (approved verbs) ----------------

function Resolve-DotmngrPath {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]$Path
  )

  # Expand %VARS% (and other env vars). We do NOT support "~".
  $expanded = [System.Environment]::ExpandEnvironmentVariables($Path)

  # Resolve if exists; otherwise return expanded.
  try {
    return (Resolve-Path -LiteralPath $expanded).Path
  } catch {
    return $expanded
  }
}

function New-DirectoryIfMissing {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]$Path
  )

  if (!(Test-Path -LiteralPath $Path)) {
    New-Item -ItemType Directory -Path $Path -Force | Out-Null
  }
}

function New-ParentDirectoryIfMissing {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]$Path
  )

  $parent = Split-Path -Parent $Path
  if ($parent) { New-DirectoryIfMissing -Path $parent }
}

function Test-ReparsePoint {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]$Path
  )

  if (!(Test-Path -LiteralPath $Path)) { return $false }
  $item = Get-Item -LiteralPath $Path -Force
  return [bool]($item.Attributes -band [IO.FileAttributes]::ReparsePoint)
}

function Get-ShortcutTarget {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]$Path
  )

  try {
    $wsh = New-Object -ComObject WScript.Shell
    $sc  = $wsh.CreateShortcut($Path)
    return $sc.TargetPath
  } catch {
    return $null
  }
}

function Resolve-WindowStyle {
  # Accepts a string alias ("normal", "default", "minimized", "maximized") or a
  # raw integer (as string or int) and returns the corresponding Win32 nShowCmd int.
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]
    $Value
  )

  $map = @{
    "normal"    = 1
    "maximized" = 3
    "minimized" = 7
  }

  $str = ([string]$Value).Trim().ToLower()
  if ($map.ContainsKey($str)) { return $map[$str] }

  $n = 0
  if ([int]::TryParse($str, [ref]$n)) { return $n }

  throw "Invalid windowStyle value '$Value'. Use a number or one of: normal, minimized, maximized."
}

function New-WindowsShortcut {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]$Path,

    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]$TargetPath,

    [Parameter()]
    [string]$WorkingDirectory = "",

    [Parameter()]
    [string]$Arguments = "",

    [Parameter()]
    [string]$Description = "",

    [Parameter()]
    [string]$IconLocation = "",

    [Parameter()]
    [int]$WindowStyle = 1
  )

  $wsh = New-Object -ComObject WScript.Shell
  $sc  = $wsh.CreateShortcut($Path)
  $sc.TargetPath  = $TargetPath
  $sc.WindowStyle = $WindowStyle
  if ($WorkingDirectory) { $sc.WorkingDirectory = $WorkingDirectory }
  if ($Arguments)        { $sc.Arguments        = $Arguments }
  if ($Description)      { $sc.Description      = $Description }
  if ($IconLocation)     { $sc.IconLocation     = $IconLocation }
  $sc.Save()
}

function Get-LinkTargetPath {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]$Path
  )

  if (!(Test-Path -LiteralPath $Path)) { return $null }
  $item = Get-Item -LiteralPath $Path -Force

  $t = $null
  try { $t = $item.Target } catch { $t = $null }
  if ($null -eq $t) {
    $t = ($item | Select-Object -ExpandProperty Target -ErrorAction SilentlyContinue)
  }
  if ($null -eq $t) { return $null }
  if ($t -is [System.Array]) { $t = $t[0] }

  return (Resolve-DotmngrPath -Path ([string]$t))
}

function Move-ItemToTrashFolder {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]$Path,

    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]$TrashDir
  )

  New-DirectoryIfMissing -Path $TrashDir

  $name  = Split-Path -Leaf $Path
  $stamp = Get-Date -Format "yyyyMMdd-HHmmss"
  $dest  = Join-Path $TrashDir "$name.$stamp"

  Move-Item -LiteralPath $Path -Destination $dest -Force
  return $dest
}

function Remove-ManagedDestination {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]$Path,

    [Parameter(Mandatory=$true)]
    [bool]$UseTrash,

    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]$TrashDir
  )

  if (!(Test-Path -LiteralPath $Path)) { return }

  if ($UseTrash) {
    $moved = Move-ItemToTrashFolder -Path $Path -TrashDir $TrashDir
        Write-Host "    moved to trash: $moved" -ForegroundColor Yellow
  } else {
    Remove-Item -LiteralPath $Path -Recurse -Force
    Write-Host "    removed." -ForegroundColor Yellow
  }
}

function Test-HardlinkMatchesSource {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]$Destination,

    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]$Source
  )

  try {
    $out = cmd /c "fsutil hardlink list `"$Destination`"" 2>$null
    if ($LASTEXITCODE -ne 0) { return $false }

    foreach ($line in ($out -split "`r?`n")) {
      if ([string]::IsNullOrWhiteSpace($line)) { continue }
      $p = Resolve-DotmngrPath -Path $line.Trim()
      if ($p -eq $Source) { return $true }
    }
    return $false
  } catch {
    return $false
  }
}

function Test-ShouldRemoveManagedLink {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]$Destination,

    [Parameter(Mandatory=$true)]
    [pscustomobject]$StateEntry
  )

  # Safe removal: remove only if destination still matches what we managed.
  if (!(Test-Path -LiteralPath $Destination)) { return $false }

  $modeValue = $StateEntry | Select-Object -ExpandProperty mode -ErrorAction SilentlyContinue
  $fromValue = $StateEntry | Select-Object -ExpandProperty from -ErrorAction SilentlyContinue
  
  $mode = if ($modeValue) { ([string]$modeValue).ToLower() } else { "" }
  $from = if ($fromValue) { Resolve-DotmngrPath -Path ([string]$fromValue) } else { "" }

  if ($mode -eq "symlink" -or $mode -eq "junction") {
    if (-not (Test-ReparsePoint -Path $Destination)) { return $false }
    $t = Get-LinkTargetPath -Path $Destination
    return ($t -eq $from)
  }

  if ($mode -eq "hardlink") {
    if (Test-ReparsePoint -Path $Destination) { return $false }
    return (Test-HardlinkMatchesSource -Destination $Destination -Source $from)
  }

  if ($mode -eq "shortcut") {
    $target = Get-ShortcutTarget -Path $Destination
    return ($null -ne $target -and $target -eq $from)
  }

  # We do not auto-remove copies.
  return $false
}

function Invoke-RobocopySafe {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]$Source,

    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]$Destination,

    [Parameter(Mandatory=$true)]
    [string[]]$Arguments
  )

  & robocopy $Source $Destination @Arguments | Out-Null
  $code = $LASTEXITCODE
  if ($code -ge 8) { throw "robocopy failed with exit code $code" }
}

# ---------------- Load config ----------------

$config = (Get-Content -LiteralPath $ConfigPath -Raw) | ConvertFrom-Json
if (-not $config.global)   { throw "Config must contain 'global'." }
if (-not $config.packages) { throw "Config must contain 'packages'." }

# Safely access mode property with diagnostics
$modeProperty = $config.global | Select-Object -ExpandProperty mode -ErrorAction SilentlyContinue
if ($null -eq $modeProperty) {
  Write-Host "WARN: global.mode not found or empty. Available properties in global:" -ForegroundColor Yellow
  $config.global | Get-Member -MemberType NoteProperty | ForEach-Object { Write-Host "  - $($_.Name)" }
  $globalMode = "symlink"
} else {
  $globalMode = ([string]$modeProperty).ToLower()
}

$globalTrash   = [bool]($config.global | Select-Object -ExpandProperty trash -ErrorAction SilentlyContinue)
$globalTrashDirVal = $config.global | Select-Object -ExpandProperty trashDir -ErrorAction SilentlyContinue
$globalTrashDir = if ($globalTrashDirVal) { Resolve-DotmngrPath -Path ([string]$globalTrashDirVal) } else { "" }

# Create trash dir upfront if needed
if ($globalTrash -and $globalTrashDir) {
  New-DirectoryIfMissing -Path $globalTrashDir
}

# ---------------- State file location ----------------

$userProfile = [System.Environment]::GetFolderPath('UserProfile')
$stateDir = Join-Path $userProfile ".config\dotmngr"
New-DirectoryIfMissing -Path $stateDir

$configFull = (Resolve-Path -LiteralPath $ConfigPath).Path
$configBase = Split-Path -LeafBase $configFull
$statePath  = Join-Path $stateDir ("state.{0}.json" -f $configBase)

$state = [pscustomobject]@{
  updated  = $null
  config   = $configFull
  packages = @{}
}

if (Test-Path -LiteralPath $statePath) {
  try {
    $loaded = (Get-Content -LiteralPath $statePath -Raw) | ConvertFrom-Json
    if ($loaded) { 
      $loadedPackages = $loaded | Select-Object -ExpandProperty packages -ErrorAction SilentlyContinue
      if ($loadedPackages) { $state.packages = $loadedPackages }
    }
  } catch {
  Write-Host "WARN: couldn't parse state file, starting fresh: $statePath" -ForegroundColor Yellow
  }
}

function Get-StatePackage {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]$Name
  )

  if (-not $state.packages) { $state.packages = @{} }

  $hasPackage = $false
  foreach ($prop in $state.packages.PSObject.Properties) {
    if ($prop.Name -eq $Name) { $hasPackage = $true; break }
  }
  
  if (-not $hasPackage) {
    $state.packages | Add-Member -MemberType NoteProperty -Name $Name -Value ([pscustomobject]@{
      updated = $null
      links   = [PSCustomObject]@{}
    })
  }
  
  # Return using safe property access
  foreach ($prop in $state.packages.PSObject.Properties) {
    if ($prop.Name -eq $Name) { return $prop.Value }
  }
  
  # Fallback (should not reach here, but safe)
  return [pscustomobject]@{ updated = $null; links = [PSCustomObject]@{} }
}

function Get-StatePackageLinks {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]$Name
  )

  $pkgState = Get-StatePackage -Name $Name
  
  # Check if links property exists, if not create it
  $linksVal = $pkgState | Select-Object -ExpandProperty links -ErrorAction SilentlyContinue
  if ($null -eq $linksVal) { 
    $pkgState | Add-Member -MemberType NoteProperty -Name links -Value ([PSCustomObject]@{}) -Force
    $linksVal = $pkgState | Select-Object -ExpandProperty links
  }
  
  return $linksVal
}

# Build package map
$packagesMap = @{}
foreach ($p in $config.packages.PSObject.Properties) {
  $packagesMap[$p.Name] = $p.Value
}

function Test-PackageEnabled {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]
    [pscustomobject]$PackageObject
  )

  $enabledVal = $PackageObject | Select-Object -ExpandProperty enabled -ErrorAction SilentlyContinue
  if ($null -eq $enabledVal) { return $true }
  return [bool]$enabledVal
}

# Determine packages to run
$selectedPackages = @()
if ($Package -and $Package.Count -gt 0) {
  # Explicit selection (even if disabled)
  $selectedPackages = @($Package)
} else {
  # Default: enabled packages only
  foreach ($k in $packagesMap.Keys) {
    if (Test-PackageEnabled -PackageObject $packagesMap[$k]) {
      $selectedPackages += $k
    }
  }
}

# ---------------- Status mode ----------------

if ($Status) {
  $rows = @()

  foreach ($pkgProp in $state.packages.PSObject.Properties) {
    $pkgName  = $pkgProp.Name
    $pkgLinks = $pkgProp.Value | Select-Object -ExpandProperty links -ErrorAction SilentlyContinue
    if ($null -eq $pkgLinks) { continue }

    foreach ($linkProp in $pkgLinks.PSObject.Properties) {
      $entry = $linkProp.Value
      $toVal   = $entry | Select-Object -ExpandProperty to   -ErrorAction SilentlyContinue
      $fromVal = $entry | Select-Object -ExpandProperty from -ErrorAction SilentlyContinue
      $modeVal = $entry | Select-Object -ExpandProperty mode -ErrorAction SilentlyContinue

      $toPath = if ($toVal) { Resolve-DotmngrPath -Path ([string]$toVal) } else { "" }

      $linkStatus = if ([string]::IsNullOrWhiteSpace($toPath)) {
        "UNKNOWN"
      } elseif (Test-Path -LiteralPath $toPath) {
        "OK"
      } else {
        "MISSING"
      }

      $rows += [pscustomobject]@{
        Package = $pkgName
        Mode    = if ($modeVal) { [string]$modeVal } else { "" }
        Status  = $linkStatus
        To      = $toPath
        From    = if ($fromVal) { Resolve-DotmngrPath -Path ([string]$fromVal) } else { "" }
      }
    }
  }

  if ($rows.Count -eq 0) {
    Write-Host "No tracked items found in state file: $statePath" -ForegroundColor Yellow
  } else {
    $rows | Format-Table -AutoSize -Property Package, Mode, Status, To, From
  }
  return
}

# ---------------- Unlink mode ----------------

if ($Unlink) {
  $unlinkPkgs = @()
  if ($Package -and $Package.Count -gt 0) {
    $unlinkPkgs = @($Package)
  } else {
    $unlinkPkgs = @()
    foreach ($prop in $state.packages.PSObject.Properties) {
      $unlinkPkgs += $prop.Name
    }
  }

  foreach ($pkg in $unlinkPkgs) {
    $hasPkg = $false
    foreach ($prop in $state.packages.PSObject.Properties) {
      if ($prop.Name -eq $pkg) { $hasPkg = $true; break }
    }
    
    if (-not $hasPkg) {
    Write-Host "==> unlink: $pkg (no state)" -ForegroundColor Yellow
      continue
    }

    Write-Host "==> unlink: $pkg" -ForegroundColor Cyan
    $links = Get-StatePackageLinks -Name $pkg

    $toKeys = @()
    foreach ($prop in $links.PSObject.Properties) {
      $toKeys += [PSCustomObject]@{ Name = $prop.Name; Value = $prop.Value }
    }
    
    foreach ($toKeyInfo in $toKeys) {
      $toKey = $toKeyInfo.Name
      $old = $toKeyInfo.Value
      $oldToVal = $old | Select-Object -ExpandProperty to -ErrorAction SilentlyContinue
      $to  = Resolve-DotmngrPath -Path ([string]$oldToVal)

      Write-Host "  - $to" -ForegroundColor White
      if (Test-ShouldRemoveManagedLink -Destination $to -StateEntry $old) {
        Remove-ManagedDestination -Path $to -UseTrash $globalTrash -TrashDir $globalTrashDir
      } else {
        Write-Host "    WARN: not removing (destination no longer matches managed link)." -ForegroundColor Yellow
      }

      $links.PSObject.Properties.Remove($toKey)
    }

    $state.packages.PSObject.Properties.Remove($pkg)
  }

  $state.updated = (Get-Date).ToString("o")
  $state | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $statePath -Encoding UTF8
  Write-Host "==> state saved: $statePath" -ForegroundColor Green
  return
}

# ---------------- Apply mode ----------------

foreach ($pkg in $selectedPackages) {
  if (-not $packagesMap.ContainsKey($pkg)) {
  Write-Host "WARN: package '$pkg' not found in config." -ForegroundColor Yellow
    continue
  }

  $pkgObj = $packagesMap[$pkg]
  if (-not $pkgObj.items) { throw "Package '$pkg' must contain 'items' array." }

  $pkgModeValue = $pkgObj | Select-Object -ExpandProperty mode -ErrorAction SilentlyContinue
  $pkgMode = if ($pkgModeValue) { ([string]$pkgModeValue).ToLower() } else { $globalMode }
  $useTrash = $globalTrash
  $trashDir = $globalTrashDir

  Write-Host "==> package: $pkg" -ForegroundColor Cyan

  $desired = @{} # toResolved -> {to,from,mode}

  foreach ($it in $pkgObj.items) {
    $itModeValue = $it | Select-Object -ExpandProperty mode -ErrorAction SilentlyContinue
    $mode = if ($itModeValue) { ([string]$itModeValue).ToLower() } else { $pkgMode }

    $itTo = $it | Select-Object -ExpandProperty to -ErrorAction SilentlyContinue
    $itFrom = $it | Select-Object -ExpandProperty from -ErrorAction SilentlyContinue
    
    # Validate required properties
    if ([string]::IsNullOrWhiteSpace($itTo) -or [string]::IsNullOrWhiteSpace($itFrom)) {
      Write-Host "  WARN: item missing 'to' or 'from' property, skipping" -ForegroundColor Yellow
      continue
    }
    
    $toExpanded = Resolve-DotmngrPath -Path ([string]$itTo)
    $fromExpanded = [System.Environment]::ExpandEnvironmentVariables([string]$itFrom)

    $fromResolved = $null
    try {
      $fromResolved = (Resolve-Path -LiteralPath $fromExpanded).Path
    } catch {
      Write-Host "  -> $mode : $fromExpanded -> $toExpanded" -ForegroundColor White
      Write-Host "     WARN: source missing, skipping." -ForegroundColor Yellow
      continue
    }

    $desiredEntry = [pscustomobject]@{ to=$toExpanded; from=$fromResolved; mode=$mode }
    if ($mode -eq "shortcut") {
      $scWorkDir  = $it | Select-Object -ExpandProperty workingDirectory -ErrorAction SilentlyContinue
      $scArgs     = $it | Select-Object -ExpandProperty arguments        -ErrorAction SilentlyContinue
      $scDesc     = $it | Select-Object -ExpandProperty description      -ErrorAction SilentlyContinue
      $scIcon     = $it | Select-Object -ExpandProperty iconLocation     -ErrorAction SilentlyContinue
      $scWinStyle = $it | Select-Object -ExpandProperty windowStyle      -ErrorAction SilentlyContinue
      if ($null -ne $scWorkDir)  { $desiredEntry | Add-Member -MemberType NoteProperty -Name workingDirectory -Value ([string]$scWorkDir) }
      if ($null -ne $scArgs)     { $desiredEntry | Add-Member -MemberType NoteProperty -Name arguments        -Value ([string]$scArgs) }
      if ($null -ne $scDesc)     { $desiredEntry | Add-Member -MemberType NoteProperty -Name description      -Value ([string]$scDesc) }
      if ($null -ne $scIcon)     { $desiredEntry | Add-Member -MemberType NoteProperty -Name iconLocation     -Value ([string]$scIcon) }
      if ($null -ne $scWinStyle) { $desiredEntry | Add-Member -MemberType NoteProperty -Name windowStyle      -Value (Resolve-WindowStyle $scWinStyle) }
    }
    $desired[$toExpanded] = $desiredEntry
  }

  # Cleanup managed destinations removed from this package
  $links = Get-StatePackageLinks -Name $pkg
  
  $propNames = @()
  foreach ($prop in $links.PSObject.Properties) {
    $propNames += [PSCustomObject]@{ Name = $prop.Name; Value = $prop.Value }
  }
  
  foreach ($propInfo in $propNames) {
    $toKey = $propInfo.Name
    if ($desired.ContainsKey($toKey)) { continue }

    $old = $propInfo.Value
    $oldToVal = $old | Select-Object -ExpandProperty to -ErrorAction SilentlyContinue
    
    # Silently clean up malformed/corrupted state entries
    if ([string]::IsNullOrWhiteSpace($oldToVal)) {
      $links.PSObject.Properties.Remove($toKey)
      continue
    }
    
    $oldTo = Resolve-DotmngrPath -Path ([string]$oldToVal)

    Write-Host "  cleanup: $oldTo" -ForegroundColor Cyan
    if (Test-ShouldRemoveManagedLink -Destination $oldTo -StateEntry $old) {
      Remove-ManagedDestination -Path $oldTo -UseTrash $useTrash -TrashDir $trashDir
    } else {
      Write-Host "    WARN: not removing (destination no longer matches managed link)." -ForegroundColor Yellow
    }

    $links.PSObject.Properties.Remove($toKey)
  }

  # Apply desired
  foreach ($toKey in $desired.Keys) {
    $it   = $desired[$toKey]
    $mode = $it | Select-Object -ExpandProperty mode -ErrorAction SilentlyContinue
    $from = $it | Select-Object -ExpandProperty from -ErrorAction SilentlyContinue
    $to   = $it | Select-Object -ExpandProperty to -ErrorAction SilentlyContinue

    Write-Host "  -> $mode : $from -> $to" -ForegroundColor White
    New-ParentDirectoryIfMissing -Path $to

    if ($mode -eq "copyonce") {
      if (Test-Path -LiteralPath $to) {
        Write-Host "     destination exists, skipping (copyOnce)." -ForegroundColor Green
        continue
      }

      if (Test-Path -LiteralPath $from -PathType Container) {
        New-DirectoryIfMissing -Path $to
        Invoke-RobocopySafe -Source $from -Destination $to -Arguments @("/E","/R:1","/W:1","/NFL","/NDL")
        Write-Host "     directory copied once." -ForegroundColor Green
      } else {
        $srcDir = Split-Path -Parent $from
        $dstDir = Split-Path -Parent $to
        $name   = Split-Path -Leaf $from
        Invoke-RobocopySafe -Source $srcDir -Destination $dstDir -Arguments @($name,"/R:1","/W:1","/NFL","/NDL")
        Write-Host "     file copied once." -ForegroundColor Green
      }
      continue
    }

    if ($mode -eq "copy") {
      if (Test-Path -LiteralPath $to -and (Test-ReparsePoint -Path $to)) {
        Write-Host "     destination is a link; removing before copy." -ForegroundColor Yellow
        Remove-ManagedDestination -Path $to -UseTrash $useTrash -TrashDir $trashDir
      }

      if (Test-Path -LiteralPath $from -PathType Container) {
        New-DirectoryIfMissing -Path $to
        Invoke-RobocopySafe -Source $from -Destination $to -Arguments @("/E","/XO","/R:1","/W:1","/NFL","/NDL")
        Write-Host "     robocopy sync completed." -ForegroundColor Green
      } else {
        New-ParentDirectoryIfMissing -Path $to
        $srcDir = Split-Path -Parent $from
        $dstDir = Split-Path -Parent $to
        $name   = Split-Path -Leaf $from
        Invoke-RobocopySafe -Source $srcDir -Destination $dstDir -Arguments @($name,"/XO","/R:1","/W:1","/NFL","/NDL")
        Write-Host "     file synced (robocopy)." -ForegroundColor Green
      }
      continue
    }

    # Link modes (tracked)
    $needsCreate = $true

    if (Test-Path -LiteralPath $to) {
      if ($mode -eq "symlink" -or $mode -eq "junction") {
        if (Test-ReparsePoint -Path $to) {
          $target = Get-LinkTargetPath -Path $to
          if ($target -eq $from) {
            Write-Host "     correct link already exists, skipping." -ForegroundColor Green
            $needsCreate = $false
          } else {
            Write-Host "     link points elsewhere ($target), replacing." -ForegroundColor Yellow
            Remove-ManagedDestination -Path $to -UseTrash $useTrash -TrashDir $trashDir
          }
        } else {
          Write-Host "     exists but is not a link, replacing." -ForegroundColor Yellow
          Remove-ManagedDestination -Path $to -UseTrash $useTrash -TrashDir $trashDir
        }
      }
      elseif ($mode -eq "hardlink") {
        if (Test-ReparsePoint -Path $to) {
          Write-Host "     is a reparse link, replacing with hardlink." -ForegroundColor Yellow
          Remove-ManagedDestination -Path $to -UseTrash $useTrash -TrashDir $trashDir
        } else {
          if (Test-HardlinkMatchesSource -Destination $to -Source $from) {
            Write-Host "     correct hardlink already exists, skipping." -ForegroundColor Green
            $needsCreate = $false
          } else {
            Write-Host "     not verified as correct hardlink, replacing." -ForegroundColor Yellow
            Remove-ManagedDestination -Path $to -UseTrash $useTrash -TrashDir $trashDir
          }
        }
      } elseif ($mode -eq "shortcut") {
        $target = Get-ShortcutTarget -Path $to
        if ($null -ne $target -and $target -eq $from) {
          Write-Host "     correct shortcut already exists, skipping." -ForegroundColor Green
          $needsCreate = $false
        } else {
          Write-Host "     shortcut points elsewhere ($target), replacing." -ForegroundColor Yellow
          Remove-ManagedDestination -Path $to -UseTrash $useTrash -TrashDir $trashDir
        }
      } else {
        throw "Unknown mode '$mode' (supported: hardlink, symlink, junction, copy, copyOnce, shortcut)"
      }
    }

    if (-not $needsCreate) {
      $links | Add-Member -MemberType NoteProperty -Name $toKey -Value ([pscustomobject]@{ to=$to; from=$from; mode=$mode; updated=(Get-Date).ToString("o") }) -Force
      continue
    }

    switch ($mode) {
      "hardlink" {
        if (Test-Path -LiteralPath $from -PathType Container) {
          throw "hardlink mode only supports files: $from"
        }
        New-Item -ItemType HardLink -Path $to -Target $from | Out-Null
        Write-Host "     hardlink created." -ForegroundColor Green
        $links | Add-Member -MemberType NoteProperty -Name $toKey -Value ([pscustomobject]@{ to=$to; from=$from; mode=$mode; updated=(Get-Date).ToString("o") }) -Force
      }
      "symlink" {
        New-Item -ItemType SymbolicLink -Path $to -Target $from | Out-Null
        Write-Host "     symlink created." -ForegroundColor Green
        $links | Add-Member -MemberType NoteProperty -Name $toKey -Value ([pscustomobject]@{ to=$to; from=$from; mode=$mode; updated=(Get-Date).ToString("o") }) -Force
      }
      "junction" {
        New-Item -ItemType Junction -Path $to -Target $from | Out-Null
        Write-Host "     junction created." -ForegroundColor Green
        $links | Add-Member -MemberType NoteProperty -Name $toKey -Value ([pscustomobject]@{ to=$to; from=$from; mode=$mode; updated=(Get-Date).ToString("o") }) -Force
      }
      "shortcut" {
        $scParams = @{ Path = $to; TargetPath = $from }
        $scWorkDir  = $it | Select-Object -ExpandProperty workingDirectory -ErrorAction SilentlyContinue
        $scArgs     = $it | Select-Object -ExpandProperty arguments        -ErrorAction SilentlyContinue
        $scDesc     = $it | Select-Object -ExpandProperty description      -ErrorAction SilentlyContinue
        $scIcon     = $it | Select-Object -ExpandProperty iconLocation     -ErrorAction SilentlyContinue
        $scWinStyle = $it | Select-Object -ExpandProperty windowStyle      -ErrorAction SilentlyContinue
        if ($scWorkDir)            { $scParams.WorkingDirectory = $scWorkDir }
        if ($scArgs)               { $scParams.Arguments        = $scArgs }
        if ($scDesc)               { $scParams.Description      = $scDesc }
        if ($scIcon)               { $scParams.IconLocation     = $scIcon }
        if ($null -ne $scWinStyle) { $scParams.WindowStyle      = Resolve-WindowStyle $scWinStyle }
        New-WindowsShortcut @scParams
        Write-Host "     shortcut created." -ForegroundColor Green
        $links | Add-Member -MemberType NoteProperty -Name $toKey -Value ([pscustomobject]@{ to=$to; from=$from; mode=$mode; updated=(Get-Date).ToString("o") }) -Force
      }
      default {
        throw "Unknown mode '$mode' (supported: hardlink, symlink, junction, copy, copyOnce, shortcut)"
      }
    }
  }

  # Stamp package updated time
  $pkgState = Get-StatePackage -Name $pkg
  $pkgState.updated = (Get-Date).ToString("o")
}

# ---------------- Save state ----------------

$state.updated = (Get-Date).ToString("o")
$state.config  = $configFull
$state | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $statePath -Encoding UTF8
Write-Host "==> state saved: $statePath"