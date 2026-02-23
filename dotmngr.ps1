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
  copyOnce  - Copy only if destination doesn’t exist
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory=$true)]
  [ValidateNotNullOrEmpty()]
  [string]$ConfigPath,

  [Parameter()]
  [string[]]$Package,

  [Parameter()]
  [switch]$Unlink
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
    Write-Host "    moved to trash: $moved"
  } else {
    Remove-Item -LiteralPath $Path -Recurse -Force
    Write-Host "    removed."
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

  $mode = ([string]($StateEntry.mode ?? "")).ToLower()
  $from = Resolve-DotmngrPath -Path ([string]($StateEntry.from ?? ""))

  if ($mode -eq "symlink" -or $mode -eq "junction") {
    if (-not (Test-ReparsePoint -Path $Destination)) { return $false }
    $t = Get-LinkTargetPath -Path $Destination
    return ($t -eq $from)
  }

  if ($mode -eq "hardlink") {
    if (Test-ReparsePoint -Path $Destination) { return $false }
    return (Test-HardlinkMatchesSource -Destination $Destination -Source $from)
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

$globalMode    = if ($config.global.mode) { ([string]$config.global.mode).ToLower() } else { "symlink" }
$globalTrash   = [bool]$config.global.trash
$globalTrashDir = Resolve-DotmngrPath -Path ([string]$config.global.trashDir)

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
    if ($loaded -and $loaded.packages) { $state.packages = $loaded.packages }
  } catch {
    Write-Host "WARN: couldn't parse state file, starting fresh: $statePath"
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

  if (-not ($state.packages.PSObject.Properties.Name -contains $Name)) {
    $state.packages | Add-Member -MemberType NoteProperty -Name $Name -Value ([pscustomobject]@{
      updated = $null
      links   = @{}
    })
  }
  return $state.packages.$Name
}

function Get-StatePackageLinks {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]$Name
  )

  $pkgState = Get-StatePackage -Name $Name
  if (-not $pkgState.links) { $pkgState.links = @{} }
  return $pkgState.links
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

  if ($null -eq $PackageObject.enabled) { return $true }
  return [bool]$PackageObject.enabled
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

# ---------------- Unlink mode ----------------

if ($Unlink) {
  $unlinkPkgs = @()
  if ($Package -and $Package.Count -gt 0) {
    $unlinkPkgs = @($Package)
  } else {
    $unlinkPkgs = @($state.packages.PSObject.Properties.Name)
  }

  foreach ($pkg in $unlinkPkgs) {
    if (-not ($state.packages.PSObject.Properties.Name -contains $pkg)) {
      Write-Host "==> unlink: $pkg (no state)"
      continue
    }

    Write-Host "==> unlink: $pkg"
    $links = Get-StatePackageLinks -Name $pkg

    foreach ($toKey in @($links.PSObject.Properties.Name)) {
      $old = $links.$toKey
      $to  = Resolve-DotmngrPath -Path ([string]$old.to)

      Write-Host "  - $to"
      if (Test-ShouldRemoveManagedLink -Destination $to -StateEntry $old) {
        Remove-ManagedDestination -Path $to -UseTrash $globalTrash -TrashDir $globalTrashDir
      } else {
        Write-Host "    WARN: not removing (destination no longer matches managed link)."
      }

      $links.PSObject.Properties.Remove($toKey)
    }

    $state.packages.PSObject.Properties.Remove($pkg)
  }

  $state.updated = (Get-Date).ToString("o")
  $state | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $statePath -Encoding UTF8
  Write-Host "==> state saved: $statePath"
  return
}

# ---------------- Apply mode ----------------

foreach ($pkg in $selectedPackages) {
  if (-not $packagesMap.ContainsKey($pkg)) {
    Write-Host "WARN: package '$pkg' not found in config."
    continue
  }

  $pkgObj = $packagesMap[$pkg]
  if (-not $pkgObj.items) { throw "Package '$pkg' must contain 'items' array." }

  $pkgMode = if ($pkgObj.mode) { ([string]$pkgObj.mode).ToLower() } else { $globalMode }
  $useTrash = $globalTrash
  $trashDir = $globalTrashDir

  Write-Host "==> package: $pkg"

  $desired = @{} # toResolved -> {to,from,mode}

  foreach ($it in $pkgObj.items) {
    $mode = if ($it.mode) { ([string]$it.mode).ToLower() } else { $pkgMode }

    $toExpanded = Resolve-DotmngrPath -Path ([string]$it.to)
    $fromExpanded = [System.Environment]::ExpandEnvironmentVariables([string]$it.from)

    $fromResolved = $null
    try {
      $fromResolved = (Resolve-Path -LiteralPath $fromExpanded).Path
    } catch {
      Write-Host "  -> $mode : $fromExpanded -> $toExpanded"
      Write-Host "     WARN: source missing, skipping."
      continue
    }

    $desired[$toExpanded] = [pscustomobject]@{ to=$toExpanded; from=$fromResolved; mode=$mode }
  }

  # Cleanup managed destinations removed from this package
  $links = Get-StatePackageLinks -Name $pkg
  foreach ($toKey in @($links.PSObject.Properties.Name)) {
    if ($desired.ContainsKey($toKey)) { continue }

    $old = $links.$toKey
    $oldTo = Resolve-DotmngrPath -Path ([string]$old.to)

    Write-Host "  cleanup: $oldTo"
    if (Test-ShouldRemoveManagedLink -Destination $oldTo -StateEntry $old) {
      Remove-ManagedDestination -Path $oldTo -UseTrash $useTrash -TrashDir $trashDir
    } else {
      Write-Host "    WARN: not removing (destination no longer matches managed link)."
    }

    $links.PSObject.Properties.Remove($toKey)
  }

  # Apply desired
  foreach ($toKey in $desired.Keys) {
    $it   = $desired[$toKey]
    $mode = $it.mode
    $from = $it.from
    $to   = $it.to

    Write-Host "  -> $mode : $from -> $to"
    New-ParentDirectoryIfMissing -Path $to

    if ($mode -eq "copyonce") {
      if (Test-Path -LiteralPath $to) {
        Write-Host "     destination exists, skipping (copyOnce)."
        continue
      }

      if (Test-Path -LiteralPath $from -PathType Container) {
        New-DirectoryIfMissing -Path $to
        Invoke-RobocopySafe -Source $from -Destination $to -Arguments @("/E","/R:1","/W:1","/NFL","/NDL")
        Write-Host "     directory copied once."
      } else {
        $srcDir = Split-Path -Parent $from
        $dstDir = Split-Path -Parent $to
        $name   = Split-Path -Leaf $from
        Invoke-RobocopySafe -Source $srcDir -Destination $dstDir -Arguments @($name,"/R:1","/W:1","/NFL","/NDL")
        Write-Host "     file copied once."
      }
      continue
    }

    if ($mode -eq "copy") {
      if (Test-Path -LiteralPath $to -and (Test-ReparsePoint -Path $to)) {
        Write-Host "     destination is a link; removing before copy."
        Remove-ManagedDestination -Path $to -UseTrash $useTrash -TrashDir $trashDir
      }

      if (Test-Path -LiteralPath $from -PathType Container) {
        New-DirectoryIfMissing -Path $to
        Invoke-RobocopySafe -Source $from -Destination $to -Arguments @("/E","/XO","/R:1","/W:1","/NFL","/NDL")
        Write-Host "     robocopy sync completed."
      } else {
        New-ParentDirectoryIfMissing -Path $to
        $srcDir = Split-Path -Parent $from
        $dstDir = Split-Path -Parent $to
        $name   = Split-Path -Leaf $from
        Invoke-RobocopySafe -Source $srcDir -Destination $dstDir -Arguments @($name,"/XO","/R:1","/W:1","/NFL","/NDL")
        Write-Host "     file synced (robocopy)."
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
            Write-Host "     correct link already exists, skipping."
            $needsCreate = $false
          } else {
            Write-Host "     link points elsewhere ($target), replacing."
            Remove-ManagedDestination -Path $to -UseTrash $useTrash -TrashDir $trashDir
          }
        } else {
          Write-Host "     exists but is not a link, replacing."
          Remove-ManagedDestination -Path $to -UseTrash $useTrash -TrashDir $trashDir
        }
      }
      elseif ($mode -eq "hardlink") {
        if (Test-ReparsePoint -Path $to) {
          Write-Host "     is a reparse link, replacing with hardlink."
          Remove-ManagedDestination -Path $to -UseTrash $useTrash -TrashDir $trashDir
        } else {
          if (Test-HardlinkMatchesSource -Destination $to -Source $from) {
            Write-Host "     correct hardlink already exists, skipping."
            $needsCreate = $false
          } else {
            Write-Host "     not verified as correct hardlink, replacing."
            Remove-ManagedDestination -Path $to -UseTrash $useTrash -TrashDir $trashDir
          }
        }
      } else {
        throw "Unknown mode '$mode' (supported: hardlink, symlink, junction, copy, copyOnce)"
      }
    }

    if (-not $needsCreate) {
      $links.$toKey = [pscustomobject]@{ to=$to; from=$from; mode=$mode; updated=(Get-Date).ToString("o") }
      continue
    }

    switch ($mode) {
      "hardlink" {
        if (Test-Path -LiteralPath $from -PathType Container) {
          throw "hardlink mode only supports files: $from"
        }
        New-Item -ItemType HardLink -Path $to -Target $from | Out-Null
        Write-Host "     hardlink created."
        $links.$toKey = [pscustomobject]@{ to=$to; from=$from; mode=$mode; updated=(Get-Date).ToString("o") }
      }
      "symlink" {
        New-Item -ItemType SymbolicLink -Path $to -Target $from | Out-Null
        Write-Host "     symlink created."
        $links.$toKey = [pscustomobject]@{ to=$to; from=$from; mode=$mode; updated=(Get-Date).ToString("o") }
      }
      "junction" {
        New-Item -ItemType Junction -Path $to -Target $from | Out-Null
        Write-Host "     junction created."
        $links.$toKey = [pscustomobject]@{ to=$to; from=$from; mode=$mode; updated=(Get-Date).ToString("o") }
      }
      default {
        throw "Unknown mode '$mode' (supported: hardlink, symlink, junction, copy, copyOnce)"
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