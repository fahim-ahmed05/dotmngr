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
  seed      - Copy only if destination doesn’t exist (file/dir)
  shortcut  - Windows .lnk shortcut (file/dir)
Switches:
  -Unlink   Remove all managed links for the selected packages
  -Status   Show a table of all tracked items and whether they are intact
  -Force    Recreate each managed destination regardless of current state
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
  [switch]$Status,

  [Parameter()]
  [switch]$Force
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
  $stamp = Get-Date -Format "yyyyMMdd-HHmmss-fff"
  $token = [System.Guid]::NewGuid().ToString("N").Substring(0, 8)

  if (Test-Path -LiteralPath $Path -PathType Container) {
    $destName = "{0}.{1}-{2}" -f $name, $stamp, $token
  } else {
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($name)
    $extension = [System.IO.Path]::GetExtension($name)
    if ([string]::IsNullOrEmpty($extension)) {
      $destName = "{0}.{1}-{2}" -f $name, $stamp, $token
    } else {
      $destName = "{0}_{1}-{2}{3}" -f $baseName, $stamp, $token, $extension
    }
  }

  $dest  = Join-Path $TrashDir $destName

  try {
    Move-Item -LiteralPath $Path -Destination $dest
    return $dest
  } catch {
    Write-LogLine -Tag "warn" -Message ("could not move to trash: {0}" -f $_.Exception.Message) -Color Yellow -Indent 4
    return $null
  }
}

function Move-ItemToRecycleBin {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]$Path
  )

  Add-Type -AssemblyName Microsoft.VisualBasic -ErrorAction Stop

  try {
    if (Test-Path -LiteralPath $Path -PathType Container) {
      [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteDirectory(
        $Path,
        [Microsoft.VisualBasic.FileIO.UIOption]::OnlyErrorDialogs,
        [Microsoft.VisualBasic.FileIO.RecycleOption]::SendToRecycleBin,
        [Microsoft.VisualBasic.FileIO.UICancelOption]::ThrowException
      )
    } else {
      [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile(
        $Path,
        [Microsoft.VisualBasic.FileIO.UIOption]::OnlyErrorDialogs,
        [Microsoft.VisualBasic.FileIO.RecycleOption]::SendToRecycleBin,
        [Microsoft.VisualBasic.FileIO.UICancelOption]::ThrowException
      )
    }

    return $true
  } catch {
    Write-LogLine -Tag "error" -Message ("could not send to recycle bin: {0}" -f $_.Exception.Message) -Color Red -Indent 4
    return $false
  }
}

function Remove-ManagedDestination {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]$Path,

    [Parameter(Mandatory=$true)]
    [bool]$UseTrash,

    [Parameter()]
    [string]$TrashDir = "",

    [Parameter()]
    [string]$ManagedMode = "",

    [Parameter()]
    [string]$ManagedSource = ""
  )

  if (!(Test-Path -LiteralPath $Path)) { return $true }

  $managedModeText = ([string]$ManagedMode).ToLower()
  if (
    $managedModeText -eq "hardlink" -and
    -not (Test-Path -LiteralPath $Path -PathType Container) -and
    -not (Test-ReparsePoint -Path $Path)
  ) {
    $sourceResolved = if ([string]::IsNullOrWhiteSpace($ManagedSource)) { "" } else { Resolve-DotmngrPath -Path $ManagedSource }
    if ($sourceResolved -and (Test-HardlinkMatchesSource -Destination $Path -Source $sourceResolved)) {
      try {
        Remove-Item -LiteralPath $Path -Force
        Write-Host "    removed hardlink only." -ForegroundColor Yellow
        return $true
      } catch {
        Write-LogLine -Tag "error" -Message ("failed to remove hardlink safely: {0}" -f $_.Exception.Message) -Color Red -Indent 4
        return $false
      }
    }
  }

  if (Test-ReparsePoint -Path $Path) {
    try {
      if (Test-Path -LiteralPath $Path -PathType Container) {
        cmd /c "rmdir `"$Path`"" | Out-Null
        if ($LASTEXITCODE -ne 0) {
          throw "rmdir failed with exit code $LASTEXITCODE"
        }
      } else {
        Remove-Item -LiteralPath $Path -Force
      }

      Write-Host "    removed link/junction only." -ForegroundColor Yellow
      return $true
    } catch {
      Write-LogLine -Tag "error" -Message ("failed to remove reparse point safely: {0}" -f $_.Exception.Message) -Color Red -Indent 4
      return $false
    }
  }

  if ($UseTrash -and -not [string]::IsNullOrWhiteSpace($TrashDir)) {
    $moved = Move-ItemToTrashFolder -Path $Path -TrashDir $TrashDir
    if ($moved) {
      Write-Host "    moved to trash: $moved" -ForegroundColor Yellow
      return $true
    }
  }

  if (Move-ItemToRecycleBin -Path $Path) {
    Write-Host "    sent to recycle bin." -ForegroundColor Yellow
    return $true
  }

  return $false
}

function Move-DestinationToSourceIfMissing {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]$SourcePath,

    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]$DestinationPath,

    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]$Mode,

    [Parameter()]
    [bool]$UseTrash = $false,

    [Parameter()]
    [string]$TrashDir = ""
  )

  if (Test-Path -LiteralPath $SourcePath) { return $true }
  if (!(Test-Path -LiteralPath $DestinationPath)) { return $false }
  if (Test-ReparsePoint -Path $DestinationPath) { return $false }

  if ($Mode -eq "junction" -and -not (Test-Path -LiteralPath $DestinationPath -PathType Container)) {
    Write-LogLine -Tag "warn" -Message "destination is not a directory, cannot recover missing source for junction mode." -Color Yellow -Indent 4
    return $false
  }

  if ($Mode -eq "hardlink" -and (Test-Path -LiteralPath $DestinationPath -PathType Container)) {
    Write-LogLine -Tag "warn" -Message "destination is a directory, cannot recover missing source for hardlink mode." -Color Yellow -Indent 4
    return $false
  }

  New-ParentDirectoryIfMissing -Path $SourcePath
  try {
    if (Test-Path -LiteralPath $DestinationPath -PathType Container) {
      New-DirectoryIfMissing -Path $SourcePath
      Invoke-RobocopySafe -Source $DestinationPath -Destination $SourcePath -Arguments @("/E","/R:1","/W:1","/NFL","/NDL")
    } else {
      Copy-Item -LiteralPath $DestinationPath -Destination $SourcePath -Force
    }
  } catch {
    Write-LogLine -Tag "warn" -Message ("source recovery copy failed, leaving destination untouched: {0}" -f $_.Exception.Message) -Color Yellow -Indent 4
    return $false
  }

  if (!(Test-Path -LiteralPath $SourcePath)) {
    Write-LogLine -Tag "warn" -Message "source recovery copy did not produce source path, leaving destination untouched." -Color Yellow -Indent 4
    return $false
  }

  Write-LogLine -Tag "recover" -Message ("source missing; copied destination to source: {0}" -f $SourcePath) -Color Yellow -Indent 4

  $archived = Remove-ManagedDestination -Path $DestinationPath -UseTrash $UseTrash -TrashDir $TrashDir -ManagedMode $Mode -ManagedSource $SourcePath
  if (-not $archived -and (Test-Path -LiteralPath $DestinationPath)) {
    Write-LogLine -Tag "warn" -Message "original destination could not be archived after copy; leaving it in place." -Color Yellow -Indent 4
  }

  return $true
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
    if ($null -eq $target) { return $false }
    $targetResolved = Resolve-DotmngrPath -Path ([string]$target)
    return ($targetResolved -eq $from)
  }

  # We do not auto-remove seed destinations.
  return $false
}

function Get-TrackedItemStatus {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]
    [string]$Destination,

    [Parameter()]
    [string]$Source = "",

    [Parameter()]
    [string]$Mode = ""
  )

  if ([string]::IsNullOrWhiteSpace($Destination)) { return "UNKNOWN" }
  if (!(Test-Path -LiteralPath $Destination)) { return "MISSING" }

  $m = ([string]$Mode).ToLower()
  switch ($m) {
    "symlink" {
      if (!(Test-ReparsePoint -Path $Destination)) { return "DRIFTED" }
      $target = Get-LinkTargetPath -Path $Destination
      if ($target -eq $Source) { return "OK" }
      return "DRIFTED"
    }
    "junction" {
      if (!(Test-ReparsePoint -Path $Destination)) { return "DRIFTED" }
      $target = Get-LinkTargetPath -Path $Destination
      if ($target -eq $Source) { return "OK" }
      return "DRIFTED"
    }
    "hardlink" {
      if (Test-ReparsePoint -Path $Destination) { return "DRIFTED" }
      if (Test-HardlinkMatchesSource -Destination $Destination -Source $Source) { return "OK" }
      return "DRIFTED"
    }
    "shortcut" {
      $target = Get-ShortcutTarget -Path $Destination
      if ($null -eq $target) { return "DRIFTED" }
      $targetResolved = Resolve-DotmngrPath -Path ([string]$target)
      if ($targetResolved -eq $Source) { return "OK" }
      return "DRIFTED"
    }
    default {
      # For seed and unknown modes, existence is the best available signal.
      return "OK"
    }
  }
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

function Write-LogLine {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]
    [string]$Tag,

    [Parameter(Mandatory=$true)]
    [string]$Message,

    [Parameter()]
    [string]$Color = "White",

    [Parameter()]
    [int]$Indent = 0
  )

  $pad = " " * [Math]::Max(0, $Indent)
  Write-Host ("{0}[{1}] {2}" -f $pad, $Tag.ToUpper(), $Message) -ForegroundColor $Color
}

function Write-PackageHeader {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]
    [string]$PackageName
  )

  Write-Host ""
  Write-LogLine -Tag "package" -Message $PackageName -Color Cyan
}

function Write-ItemHeader {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]
    [string]$Mode,

    [Parameter(Mandatory=$true)]
    [string]$From,

    [Parameter(Mandatory=$true)]
    [string]$To
  )

  Write-LogLine -Tag "item" -Message $Mode -Color White -Indent 2
  Write-Host ("    FROM: {0}" -f $From) -ForegroundColor DarkGray
  Write-Host ("    TO  : {0}" -f $To) -ForegroundColor DarkGray
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

if ($globalTrash) {
  if ([string]::IsNullOrWhiteSpace([string]$globalTrashDirVal)) {
    $globalTrashDir = Resolve-DotmngrPath -Path "%USERPROFILE%\Trash\dotmngr"
  } else {
    $globalTrashDir = Resolve-DotmngrPath -Path ([string]$globalTrashDirVal)
  }
} else {
  $globalTrashDir = ""
}

# Create trash dir upfront if needed
if ($globalTrash -and -not [string]::IsNullOrWhiteSpace($globalTrashDir)) {
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
  packages = [pscustomobject]@{}
}

if (Test-Path -LiteralPath $statePath) {
  try {
    $loaded = (Get-Content -LiteralPath $statePath -Raw) | ConvertFrom-Json
    if ($loaded) {
      $loadedPackages = $loaded | Select-Object -ExpandProperty packages -ErrorAction SilentlyContinue
      if ($loadedPackages) {
        if ($loadedPackages -is [hashtable]) {
          $state.packages = [pscustomobject]$loadedPackages
        } else {
          $state.packages = $loadedPackages
        }
      }
    }
  } catch {
    Write-LogLine -Tag "warn" -Message ("couldn't parse state file, starting fresh: {0}" -f $statePath) -Color Yellow
  }
}

function Get-StatePackage {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]$Name
  )

  if (-not $state.packages) { $state.packages = [pscustomobject]@{} }

  $pkgProp = $state.packages.PSObject.Properties[$Name]
  if ($null -eq $pkgProp) {
    $state.packages | Add-Member -MemberType NoteProperty -Name $Name -Value ([pscustomobject]@{
      updated = $null
      links   = [PSCustomObject]@{}
    })
    $pkgProp = $state.packages.PSObject.Properties[$Name]
  }

  $pkgState = $pkgProp.Value
  if (-not $pkgState.PSObject.Properties["links"]) {
    $pkgState | Add-Member -MemberType NoteProperty -Name links -Value ([PSCustomObject]@{}) -Force
  }

  return $pkgState
}

function Get-StatePackageLinks {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]$Name
  )

  $pkgState = Get-StatePackage -Name $Name
  return $pkgState.links
}

function Get-StateLinkEntries {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]
    $LinksObject
  )

  $entries = @()
  foreach ($prop in $LinksObject.PSObject.Properties) {
    $entries += [PSCustomObject]@{ Name = $prop.Name; Value = $prop.Value }
  }

  return $entries
}

function Invoke-TrackedEntryCleanup {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]
    [pscustomobject]$StateEntry,

    [Parameter(Mandatory=$true)]
    [bool]$UseTrash,

    [Parameter()]
    [string]$TrashDir = "",

    [Parameter()]
    [ValidateSet("cleanup", "bullet")]
    [string]$PathOutputStyle = "cleanup"
  )

  $oldToVal = $StateEntry | Select-Object -ExpandProperty to -ErrorAction SilentlyContinue
  if ([string]::IsNullOrWhiteSpace($oldToVal)) { return $true }

  $oldTo = Resolve-DotmngrPath -Path ([string]$oldToVal)

  if ($PathOutputStyle -eq "bullet") {
    Write-Host ("  - {0}" -f $oldTo) -ForegroundColor White
  } else {
    Write-LogLine -Tag "cleanup" -Message $oldTo -Color Cyan -Indent 2
  }

  if (Test-ShouldRemoveManagedLink -Destination $oldTo -StateEntry $StateEntry) {
    $entryMode = [string]($StateEntry | Select-Object -ExpandProperty mode -ErrorAction SilentlyContinue)
    $entrySource = [string]($StateEntry | Select-Object -ExpandProperty from -ErrorAction SilentlyContinue)
    $removed = Remove-ManagedDestination -Path $oldTo -UseTrash $UseTrash -TrashDir $TrashDir -ManagedMode $entryMode -ManagedSource $entrySource
    if (-not $removed -and (Test-Path -LiteralPath $oldTo)) {
      Write-LogLine -Tag "warn" -Message "removal failed; keeping state entry." -Color Yellow -Indent 4
      return $false
    }
  } else {
    Write-LogLine -Tag "warn" -Message "not removing (destination no longer matches managed link)." -Color Yellow -Indent 4
  }

  return $true
}

function Remove-StatePackageIfEmpty {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]
    [string]$PackageName,

    [Parameter(Mandatory=$true)]
    $LinksObject
  )

  $linkCount = @($LinksObject.PSObject.Properties).Count
  if ($linkCount -eq 0) {
    $state.packages.PSObject.Properties.Remove($PackageName)
  } else {
    Write-LogLine -Tag "warn" -Message ("state retained for package '{0}' because some items could not be removed." -f $PackageName) -Color Yellow -Indent 2
  }
}

function Invoke-PackageCleanupFromState {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]
    [string]$PackageName,

    [Parameter(Mandatory=$true)]
    [string]$Tag,

    [Parameter(Mandatory=$true)]
    [string]$Reason,

    [Parameter(Mandatory=$true)]
    [bool]$UseTrash,

    [Parameter()]
    [string]$TrashDir = ""
  )

  if (-not $state.packages.PSObject.Properties[$PackageName]) { return }

  Write-Host ""
  Write-LogLine -Tag $Tag -Message ("{0} ({1})" -f $PackageName, $Reason) -Color Cyan

  $pkgLinks = Get-StatePackageLinks -Name $PackageName
  $toKeys = Get-StateLinkEntries -LinksObject $pkgLinks

  foreach ($toKeyInfo in $toKeys) {
    $toKey = $toKeyInfo.Name
    $old = $toKeyInfo.Value

    $removeStateEntry = Invoke-TrackedEntryCleanup -StateEntry $old -UseTrash $UseTrash -TrashDir $TrashDir
    if ($removeStateEntry) {
      $pkgLinks.PSObject.Properties.Remove($toKey)
    }
  }

  Remove-StatePackageIfEmpty -PackageName $PackageName -LinksObject $pkgLinks
}

function Test-ReplaceDestinationRemoval {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]$DestinationPath,

    [Parameter(Mandatory=$true)]
    [string]$ManagedMode,

    [Parameter(Mandatory=$true)]
    [string]$ManagedSource,

    [Parameter(Mandatory=$true)]
    [bool]$UseTrash,

    [Parameter()]
    [string]$TrashDir = ""
  )

  $removed = Remove-ManagedDestination -Path $DestinationPath -UseTrash $UseTrash -TrashDir $TrashDir -ManagedMode $ManagedMode -ManagedSource $ManagedSource
  if (-not $removed -and (Test-Path -LiteralPath $DestinationPath)) {
    Write-LogLine -Tag "error" -Message "replacement skipped because existing destination could not be removed." -Color Red -Indent 4
    return $false
  }

  return $true
}

function Set-TrackedLinkState {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]
    $LinksObject,

    [Parameter(Mandatory=$true)]
    [string]$DestinationPath,

    [Parameter(Mandatory=$true)]
    [string]$SourcePath,

    [Parameter(Mandatory=$true)]
    [string]$Mode
  )

  $LinksObject | Add-Member -MemberType NoteProperty -Name $DestinationPath -Value ([pscustomobject]@{
    to = $DestinationPath
    from = $SourcePath
    mode = $Mode
    updated = (Get-Date).ToString("o")
  }) -Force
}

function Invoke-SeedApplyIfNeeded {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]
    [string]$Mode,

    [Parameter(Mandatory=$true)]
    [string]$SourcePath,

    [Parameter(Mandatory=$true)]
    [string]$DestinationPath
  )

  if ($Mode -ne "seed") { return $false }

  if (Test-Path -LiteralPath $DestinationPath) {
    Write-LogLine -Tag "skip" -Message "destination exists, skipping (seed)." -Color Green -Indent 4
    return $true
  }

  if (Test-Path -LiteralPath $SourcePath -PathType Container) {
    New-DirectoryIfMissing -Path $DestinationPath
    Invoke-RobocopySafe -Source $SourcePath -Destination $DestinationPath -Arguments @("/E","/R:1","/W:1","/NFL","/NDL")
    Write-LogLine -Tag "create" -Message "directory copied once." -Color Green -Indent 4
  } else {
    Copy-Item -LiteralPath $SourcePath -Destination $DestinationPath -Force
    Write-LogLine -Tag "create" -Message "file copied once." -Color Green -Indent 4
  }

  return $true
}

function Get-ApplyDestinationDecision {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]
    [string]$Mode,

    [Parameter(Mandatory=$true)]
    [string]$SourcePath,

    [Parameter(Mandatory=$true)]
    [string]$DestinationPath,

    [Parameter(Mandatory=$true)]
    [bool]$UseTrash,

    [Parameter()]
    [string]$TrashDir = ""
  )

  $result = [pscustomobject]@{
    Proceed = $true
    NeedsCreate = $true
  }

  if (-not (Test-Path -LiteralPath $DestinationPath)) {
    return $result
  }

  if ($Mode -eq "symlink" -or $Mode -eq "junction") {
    if (Test-ReparsePoint -Path $DestinationPath) {
      $target = Get-LinkTargetPath -Path $DestinationPath
      if ($target -eq $SourcePath) {
        Write-LogLine -Tag "skip" -Message "correct link already exists." -Color Green -Indent 4
        $result.NeedsCreate = $false
        return $result
      }

      Write-LogLine -Tag "replace" -Message ("link points elsewhere ({0}), replacing." -f $target) -Color Yellow -Indent 4
      if (-not (Test-ReplaceDestinationRemoval -DestinationPath $DestinationPath -ManagedMode $Mode -ManagedSource $SourcePath -UseTrash $UseTrash -TrashDir $TrashDir)) {
        $result.Proceed = $false
      }
      return $result
    }

    Write-LogLine -Tag "replace" -Message "exists but is not a link, replacing." -Color Yellow -Indent 4
    if (-not (Test-ReplaceDestinationRemoval -DestinationPath $DestinationPath -ManagedMode $Mode -ManagedSource $SourcePath -UseTrash $UseTrash -TrashDir $TrashDir)) {
      $result.Proceed = $false
    }
    return $result
  }

  if ($Mode -eq "hardlink") {
    if (Test-ReparsePoint -Path $DestinationPath) {
      Write-LogLine -Tag "replace" -Message "is a reparse link, replacing with hardlink." -Color Yellow -Indent 4
      if (-not (Test-ReplaceDestinationRemoval -DestinationPath $DestinationPath -ManagedMode $Mode -ManagedSource $SourcePath -UseTrash $UseTrash -TrashDir $TrashDir)) {
        $result.Proceed = $false
      }
      return $result
    }

    if (Test-HardlinkMatchesSource -Destination $DestinationPath -Source $SourcePath) {
      Write-LogLine -Tag "skip" -Message "correct hardlink already exists." -Color Green -Indent 4
      $result.NeedsCreate = $false
      return $result
    }

    Write-LogLine -Tag "replace" -Message "hardlink target mismatch, replacing." -Color Yellow -Indent 4
    if (-not (Test-ReplaceDestinationRemoval -DestinationPath $DestinationPath -ManagedMode $Mode -ManagedSource $SourcePath -UseTrash $UseTrash -TrashDir $TrashDir)) {
      $result.Proceed = $false
    }
    return $result
  }

  if ($Mode -eq "shortcut") {
    $target = Get-ShortcutTarget -Path $DestinationPath
    $targetResolved = if ($null -ne $target) { Resolve-DotmngrPath -Path ([string]$target) } else { $null }
    if ($null -ne $targetResolved -and $targetResolved -eq $SourcePath) {
      Write-LogLine -Tag "skip" -Message "correct shortcut already exists." -Color Green -Indent 4
      $result.NeedsCreate = $false
      return $result
    }

    Write-LogLine -Tag "replace" -Message ("shortcut points elsewhere ({0}), replacing." -f $targetResolved) -Color Yellow -Indent 4
    if (-not (Test-ReplaceDestinationRemoval -DestinationPath $DestinationPath -ManagedMode $Mode -ManagedSource $SourcePath -UseTrash $UseTrash -TrashDir $TrashDir)) {
      $result.Proceed = $false
    }
    return $result
  }

  throw "Unknown mode '$Mode' (supported: hardlink, symlink, junction, seed, shortcut)"
}

function Invoke-CreateTrackedLinkForMode {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]
    [string]$Mode,

    [Parameter(Mandatory=$true)]
    [string]$SourcePath,

    [Parameter(Mandatory=$true)]
    [string]$DestinationPath,

    [Parameter(Mandatory=$true)]
    [string]$ToKey,

    [Parameter(Mandatory=$true)]
    $LinksObject,

    [Parameter(Mandatory=$true)]
    [pscustomobject]$DesiredItem
  )

  switch ($Mode) {
    "hardlink" {
      if (Test-Path -LiteralPath $SourcePath -PathType Container) {
        throw "hardlink mode only supports files: $SourcePath"
      }
      New-Item -ItemType HardLink -Path $DestinationPath -Target $SourcePath | Out-Null
      Write-LogLine -Tag "create" -Message "hardlink created." -Color Green -Indent 4
      Set-TrackedLinkState -LinksObject $LinksObject -DestinationPath $ToKey -SourcePath $SourcePath -Mode $Mode
    }
    "symlink" {
      New-Item -ItemType SymbolicLink -Path $DestinationPath -Target $SourcePath | Out-Null
      Write-LogLine -Tag "create" -Message "symlink created." -Color Green -Indent 4
      Set-TrackedLinkState -LinksObject $LinksObject -DestinationPath $ToKey -SourcePath $SourcePath -Mode $Mode
    }
    "junction" {
      New-Item -ItemType Junction -Path $DestinationPath -Target $SourcePath | Out-Null
      Write-LogLine -Tag "create" -Message "junction created." -Color Green -Indent 4
      Set-TrackedLinkState -LinksObject $LinksObject -DestinationPath $ToKey -SourcePath $SourcePath -Mode $Mode
    }
    "shortcut" {
      $scParams = @{ Path = $DestinationPath; TargetPath = $SourcePath }
      $scWorkDir  = $DesiredItem | Select-Object -ExpandProperty workingDirectory -ErrorAction SilentlyContinue
      $scArgs     = $DesiredItem | Select-Object -ExpandProperty arguments        -ErrorAction SilentlyContinue
      $scDesc     = $DesiredItem | Select-Object -ExpandProperty description      -ErrorAction SilentlyContinue
      $scIcon     = $DesiredItem | Select-Object -ExpandProperty iconLocation     -ErrorAction SilentlyContinue
      $scWinStyle = $DesiredItem | Select-Object -ExpandProperty windowStyle      -ErrorAction SilentlyContinue
      if ($scWorkDir)            { $scParams.WorkingDirectory = [System.Environment]::ExpandEnvironmentVariables($scWorkDir) }
      if ($scArgs)               { $scParams.Arguments        = [System.Environment]::ExpandEnvironmentVariables($scArgs) }
      if ($scDesc)               { $scParams.Description      = $scDesc }
      if ($scIcon)               { $scParams.IconLocation     = [System.Environment]::ExpandEnvironmentVariables($scIcon) }
      if ($null -ne $scWinStyle) { $scParams.WindowStyle      = Resolve-WindowStyle $scWinStyle }
      New-WindowsShortcut @scParams
      Write-LogLine -Tag "create" -Message "shortcut created." -Color Green -Indent 4
      Set-TrackedLinkState -LinksObject $LinksObject -DestinationPath $ToKey -SourcePath $SourcePath -Mode $Mode
    }
    default {
      throw "Unknown mode '$Mode' (supported: hardlink, symlink, junction, seed, shortcut)"
    }
  }
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
      $fromPath = if ($fromVal) { Resolve-DotmngrPath -Path ([string]$fromVal) } else { "" }
      $modeText = if ($modeVal) { [string]$modeVal } else { "" }
      $linkStatus = Get-TrackedItemStatus -Destination $toPath -Source $fromPath -Mode $modeText

      $rows += [pscustomobject]@{
        Package = $pkgName
        Mode    = $modeText
        Status  = $linkStatus
        To      = $toPath
        From    = $fromPath
      }
    }
  }

  if ($rows.Count -eq 0) {
    Write-Host "No tracked items found in state file: $statePath" -ForegroundColor Yellow
  } else {
    $groupedRows = $rows | Sort-Object Package, To | Group-Object Package

    foreach ($pkgGroup in $groupedRows) {
      $pkgRows = @($pkgGroup.Group)
      $okCount = @($pkgRows | Where-Object { $_.Status -eq "OK" }).Count
      $driftedCount = @($pkgRows | Where-Object { $_.Status -eq "DRIFTED" }).Count
      $missingCount = @($pkgRows | Where-Object { $_.Status -eq "MISSING" }).Count
      $otherCount = $pkgRows.Count - $okCount - $driftedCount - $missingCount

      $summary = "items={0} ok={1} drifted={2} missing={3}" -f $pkgRows.Count, $okCount, $driftedCount, $missingCount
      if ($otherCount -gt 0) {
        $summary = "{0} other={1}" -f $summary, $otherCount
      }

      $headerColor = "Cyan"
      if ($missingCount -gt 0) {
        $headerColor = "Red"
      } elseif ($driftedCount -gt 0) {
        $headerColor = "Yellow"
      }

      Write-Host ""
      Write-LogLine -Tag "package" -Message ("{0} ({1})" -f $pkgGroup.Name, $summary) -Color $headerColor

      $pkgRows |
        Sort-Object -Property @{ Expression = {
          switch ($_.Status) {
            "DRIFTED" { 0 }
            "MISSING" { 1 }
            "OK"      { 2 }
            default     { 3 }
          }
        } }, Mode, To |
        Format-Table -AutoSize -Property Mode, Status, To
    }
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
    if (-not $state.packages.PSObject.Properties[$pkg]) {
      Write-LogLine -Tag "unlink" -Message ("{0} (no state)" -f $pkg) -Color Yellow
      continue
    }

    Write-LogLine -Tag "unlink" -Message $pkg -Color Cyan
    $links = Get-StatePackageLinks -Name $pkg

    $toKeys = Get-StateLinkEntries -LinksObject $links

    foreach ($toKeyInfo in $toKeys) {
      $toKey = $toKeyInfo.Name
      $old = $toKeyInfo.Value

      $removeStateEntry = Invoke-TrackedEntryCleanup -StateEntry $old -UseTrash $globalTrash -TrashDir $globalTrashDir -PathOutputStyle "bullet"
      if ($removeStateEntry) {
        $links.PSObject.Properties.Remove($toKey)
      }
    }

    Remove-StatePackageIfEmpty -PackageName $pkg -LinksObject $links
  }

  $state.updated = (Get-Date).ToString("o")
  $state | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $statePath -Encoding UTF8
  Write-Host ""
  Write-LogLine -Tag "save" -Message ("state saved: {0}" -f $statePath) -Color Green
  return
}

# ---------------- Apply mode ----------------

# When running default apply (no explicit -Package), disabled packages should
# be treated as removed and cleaned up from state.
if (-not ($Package -and $Package.Count -gt 0)) {
  foreach ($pkgName in $packagesMap.Keys) {
    if (Test-PackageEnabled -PackageObject $packagesMap[$pkgName]) { continue }
    Invoke-PackageCleanupFromState -PackageName $pkgName -Tag "disable" -Reason "enabled=false; cleaning managed links" -UseTrash $globalTrash -TrashDir $globalTrashDir
  }

  $orphanStatePackages = @()
  foreach ($pkgProp in $state.packages.PSObject.Properties) {
    if (-not $packagesMap.ContainsKey($pkgProp.Name)) {
      $orphanStatePackages += $pkgProp.Name
    }
  }

  foreach ($pkgName in $orphanStatePackages) {
    Invoke-PackageCleanupFromState -PackageName $pkgName -Tag "orphan" -Reason "not in config; cleaning managed links" -UseTrash $globalTrash -TrashDir $globalTrashDir
  }
}


foreach ($pkg in $selectedPackages) {
  if (-not $packagesMap.ContainsKey($pkg)) {
    Write-LogLine -Tag "warn" -Message ("package '{0}' not found in config." -f $pkg) -Color Yellow
    continue
  }

  $pkgObj = $packagesMap[$pkg]
  if (-not $pkgObj.items) { throw "Package '$pkg' must contain 'items' array." }

  $pkgModeValue = $pkgObj | Select-Object -ExpandProperty mode -ErrorAction SilentlyContinue
  $pkgMode = if ($pkgModeValue) { ([string]$pkgModeValue).ToLower() } else { $globalMode }
  $useTrash = $globalTrash
  $trashDir = $globalTrashDir

  Write-PackageHeader -PackageName $pkg

  $desired = @{} # toResolved -> {to,from,mode}

  foreach ($it in $pkgObj.items) {
    $itModeValue = $it | Select-Object -ExpandProperty mode -ErrorAction SilentlyContinue
    $mode = if ($itModeValue) { ([string]$itModeValue).ToLower() } else { $pkgMode }

    $itTo = $it | Select-Object -ExpandProperty to -ErrorAction SilentlyContinue
    $itFrom = $it | Select-Object -ExpandProperty from -ErrorAction SilentlyContinue

    # Validate required properties
    if ([string]::IsNullOrWhiteSpace($itTo) -or [string]::IsNullOrWhiteSpace($itFrom)) {
      Write-LogLine -Tag "warn" -Message "item missing 'to' or 'from' property, skipping" -Color Yellow -Indent 2
      continue
    }

    $toExpanded = Resolve-DotmngrPath -Path ([string]$itTo)
    $fromExpanded = [System.Environment]::ExpandEnvironmentVariables([string]$itFrom)

    $fromResolved = $null
    try {
      $fromResolved = (Resolve-Path -LiteralPath $fromExpanded).Path
    } catch {
      Write-ItemHeader -Mode $mode -From $fromExpanded -To $toExpanded

      $isRecoverableLinkMode = ($mode -eq "symlink" -or $mode -eq "junction" -or $mode -eq "hardlink")
      if ($isRecoverableLinkMode) {
        $recovered = Move-DestinationToSourceIfMissing -SourcePath $fromExpanded -DestinationPath $toExpanded -Mode $mode -UseTrash $useTrash -TrashDir $trashDir
        if ($recovered) {
          try {
            $fromResolved = (Resolve-Path -LiteralPath $fromExpanded).Path
          } catch {
            Write-LogLine -Tag "warn" -Message "recovery attempted, but source still missing, skipping." -Color Yellow -Indent 4
            continue
          }
        } else {
          Write-LogLine -Tag "warn" -Message "source missing, skipping." -Color Yellow -Indent 4
          continue
        }
      } else {
        Write-LogLine -Tag "warn" -Message "source missing, skipping." -Color Yellow -Indent 4
        continue
      }
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

  $propNames = Get-StateLinkEntries -LinksObject $links

  foreach ($propInfo in $propNames) {
    $toKey = $propInfo.Name
    if ($desired.ContainsKey($toKey)) {
      $desiredModeVal = $desired[$toKey] | Select-Object -ExpandProperty mode -ErrorAction SilentlyContinue
      $desiredMode = if ($desiredModeVal) { ([string]$desiredModeVal).ToLower() } else { "" }

      if ($desiredMode -eq "seed") {
        # seed is intentionally untracked; drop any stale tracked state entry.
        Write-LogLine -Tag "untrack" -Message ("{0} (seed mode)" -f $toKey) -Color Cyan -Indent 2
        $links.PSObject.Properties.Remove($toKey)
      }
      continue
    }

    $old = $propInfo.Value
    $removeStateEntry = Invoke-TrackedEntryCleanup -StateEntry $old -UseTrash $useTrash -TrashDir $trashDir
    if ($removeStateEntry) {
      $links.PSObject.Properties.Remove($toKey)
    }
  }

  # Apply desired
  foreach ($toKey in $desired.Keys) {
    $it   = $desired[$toKey]
    $mode = $it | Select-Object -ExpandProperty mode -ErrorAction SilentlyContinue
    $from = $it | Select-Object -ExpandProperty from -ErrorAction SilentlyContinue
    $to   = $it | Select-Object -ExpandProperty to -ErrorAction SilentlyContinue

    try {
      Write-ItemHeader -Mode $mode -From $from -To $to
      New-ParentDirectoryIfMissing -Path $to

      if ($Force -and (Test-Path -LiteralPath $to)) {
        Write-LogLine -Tag "replace" -Message "removing destination before apply." -Color Yellow -Indent 4
        if (-not (Test-ReplaceDestinationRemoval -DestinationPath $to -ManagedMode $mode -ManagedSource $from -UseTrash $useTrash -TrashDir $trashDir)) {
          continue
        }
      }

      if (Invoke-SeedApplyIfNeeded -Mode $mode -SourcePath $from -DestinationPath $to) {
        continue
      }

      # Link modes (tracked)
      $decision = Get-ApplyDestinationDecision -Mode $mode -SourcePath $from -DestinationPath $to -UseTrash $useTrash -TrashDir $trashDir
      if (-not $decision.Proceed) {
        continue
      }
      $needsCreate = [bool]$decision.NeedsCreate

      if (-not $needsCreate) {
        Set-TrackedLinkState -LinksObject $links -DestinationPath $toKey -SourcePath $from -Mode $mode
        continue
      }

      Invoke-CreateTrackedLinkForMode -Mode $mode -SourcePath $from -DestinationPath $to -ToKey $toKey -LinksObject $links -DesiredItem $it
    } catch {
      $errMsg = $_.Exception.Message
      Write-LogLine -Tag "error" -Message ("failed to process item (mode={0}): {1}" -f $mode, $errMsg) -Color Red -Indent 4
      continue
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
Write-Host ""
Write-LogLine -Tag "save" -Message ("state saved: {0}" -f $statePath) -Color Green
