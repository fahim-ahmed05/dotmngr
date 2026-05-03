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
  -Relink   Unlink then re-link the selected packages
  -Status   Show a table of all tracked items and whether they are intact
  -Force    Recreate each managed destination regardless of current state
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)]
  [ValidateNotNullOrEmpty()]
  [string]$ConfigPath,

  [Parameter(ValueFromRemainingArguments = $true)]
  [string[]]$Package,

  [Parameter()]
  [switch]$Unlink,

  [Parameter()]
  [switch]$Relink,

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
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$Path
  )

  # Expand %VARS% (and other env vars). We do NOT support "~".
  $expanded = [System.Environment]::ExpandEnvironmentVariables($Path)

  # GetFullPath normalizes the path instantly without throwing if the file doesn't exist.
  try {
    return [System.IO.Path]::GetFullPath($expanded)
  }
  catch {
    return $expanded
  }
}

function New-DirectoryIfMissing {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
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
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$Path
  )

  $parent = Split-Path -Parent $Path
  if ($parent) { New-DirectoryIfMissing -Path $parent }
}

function Test-ReparsePoint {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$Path
  )

  if (!(Test-Path -LiteralPath $Path)) { return $false }
  return ([System.IO.File]::GetAttributes($Path) -band [System.IO.FileAttributes]::ReparsePoint) -eq [System.IO.FileAttributes]::ReparsePoint
}

function Get-ShortcutTarget {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$Path
  )

  try {
    $wsh = New-Object -ComObject WScript.Shell
    $sc = $wsh.CreateShortcut($Path)
    return $sc.TargetPath
  }
  catch {
    return $null
  }
}

function Resolve-WindowStyle {
  # Accepts a string alias ("normal", "default", "minimized", "maximized") or a
  # raw integer (as string or int) and returns the corresponding Win32 nShowCmd int.
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
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
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$Path,

    [Parameter(Mandatory = $true)]
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
  $sc = $wsh.CreateShortcut($Path)
  $sc.TargetPath = $TargetPath
  $sc.WindowStyle = $WindowStyle
  if ($WorkingDirectory) { $sc.WorkingDirectory = $WorkingDirectory }
  if ($Arguments) { $sc.Arguments = $Arguments }
  if ($Description) { $sc.Description = $Description }
  if ($IconLocation) { $sc.IconLocation = $IconLocation }
  $sc.Save()
}

function Get-LinkTargetPath {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$Path
  )

  if (!(Test-Path -LiteralPath $Path)) { return $null }
  
  try {
    $target = if (Test-Path -LiteralPath $Path -PathType Container) {
      [System.IO.Directory]::ResolveLinkTarget($Path, $false)
    } else {
      [System.IO.File]::ResolveLinkTarget($Path, $false)
    }
    
    if ($null -ne $target) { 
      return (Resolve-DotmngrPath -Path $target.FullName)
    }
    return $null
  }
  catch {
    return $null
  }
}

function Remove-DestinationForOverwrite {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$Path,

    [Parameter(Mandatory = $true)]
    [bool]$UseTrash,

    [Parameter()]
    [string]$TrashDir = ""
  )

  if (!(Test-Path -LiteralPath $Path)) { return $true }

  # Handle reparse points (symlinks/junctions) - remove only the link, not the target
  if (Test-ReparsePoint -Path $Path) {
    try {
      Remove-Item -LiteralPath $Path -Force -ErrorAction Stop
      Write-LogLine -Tag "remove" -Message "reparse point removed." -Color Yellow -Indent 4
      return $true
    }
    catch {
      Write-LogLine -Tag "error" -Message ("failed to remove reparse point: {0}" -f $_.Exception.Message) -Color Red -Indent 4
      return $false
    }
  }

  # Handle regular files/dirs - try trash first, then recycle bin
  if ($UseTrash -and -not [string]::IsNullOrWhiteSpace($TrashDir)) {
    $moved = Move-ItemToTrashFolder -Path $Path -TrashDir $TrashDir
    if ($moved) {
      Write-LogLine -Tag "backup" -Message "moved to trash" -Color Yellow -Indent 4
      return $true
    }
  }

  if (Move-ItemToRecycleBin -Path $Path) {
    Write-LogLine -Tag "backup" -Message "moved to recycle bin" -Color Yellow -Indent 4
    return $true
  }

  Write-LogLine -Tag "error" -Message "could not remove path, skipping." -Color Red -Indent 4
  return $false
}

function Move-ItemToTrashFolder {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$Path,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$TrashDir
  )

  New-DirectoryIfMissing -Path $TrashDir

  $name = Split-Path -Leaf $Path
  $stamp = Get-Date -Format "yyyyMMdd-HHmmss-fff"
  $token = [System.Guid]::NewGuid().ToString("N").Substring(0, 8)

  if (Test-Path -LiteralPath $Path -PathType Container) {
    $destName = "{0}.{1}-{2}" -f $name, $stamp, $token
  }
  else {
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($name)
    $extension = [System.IO.Path]::GetExtension($name)
    if ([string]::IsNullOrEmpty($extension)) {
      $destName = "{0}.{1}-{2}" -f $name, $stamp, $token
    }
    else {
      $destName = "{0}_{1}-{2}{3}" -f $baseName, $stamp, $token, $extension
    }
  }

  $dest = Join-Path $TrashDir $destName

  try {
    Move-Item -LiteralPath $Path -Destination $dest
    return $dest
  }
  catch {
    Write-LogLine -Tag "warn" -Message ("could not move to trash: {0}" -f $_.Exception.Message) -Color Yellow -Indent 4
    return $null
  }
}

function Move-ItemToRecycleBin {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
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
    }
    else {
      [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile(
        $Path,
        [Microsoft.VisualBasic.FileIO.UIOption]::OnlyErrorDialogs,
        [Microsoft.VisualBasic.FileIO.RecycleOption]::SendToRecycleBin,
        [Microsoft.VisualBasic.FileIO.UICancelOption]::ThrowException
      )
    }

    return $true
  }
  catch {
    Write-LogLine -Tag "error" -Message ("could not send to recycle bin: {0}" -f $_.Exception.Message) -Color Red -Indent 4
    return $false
  }
}

function Remove-ManagedDestination {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$Path,

    [Parameter(Mandatory = $true)]
    [bool]$UseTrash,

    [Parameter()]
    [string]$TrashDir = "",

    [Parameter()]
    [string]$ManagedMode = "",

    [Parameter()]
    [string]$ManagedSource = ""
  )

  if (!(Test-Path -LiteralPath $Path)) { return $true }

  # For hardlinks: safe removal only if it matches the source (not a copy)
  $managedModeText = ([string]$ManagedMode).ToLower()
  if (
    $managedModeText -eq "hardlink" -and
    -not (Test-Path -LiteralPath $Path -PathType Container) -and
    -not (Test-ReparsePoint -Path $Path)
  ) {
    $sourceResolved = if ([string]::IsNullOrWhiteSpace($ManagedSource)) { "" } else { Resolve-DotmngrPath -Path $ManagedSource }
    if ($sourceResolved -and (Test-HardlinkMatchesSource -Destination $Path -Source $sourceResolved)) {
      try {
        Remove-Item -LiteralPath $Path -Force -ErrorAction Stop
        Write-LogLine -Tag "remove" -Message "hardlink removed." -Color Yellow -Indent 4
        return $true
      }
      catch {
        Write-LogLine -Tag "error" -Message ("failed to remove hardlink: {0}" -f $_.Exception.Message) -Color Red -Indent 4
        return $false
      }
    }
  }

  # For reparse points (symlinks/junctions): remove only the link
  if (Test-ReparsePoint -Path $Path) {
    try {
      Remove-Item -LiteralPath $Path -Force -ErrorAction Stop
      Write-LogLine -Tag "remove" -Message "reparse point removed." -Color Yellow -Indent 4
      return $true
    }
    catch {
      Write-LogLine -Tag "error" -Message ("failed to remove reparse point: {0}" -f $_.Exception.Message) -Color Red -Indent 4
      return $false
    }
  }

  # For regular files/dirs: use trash or recycle bin
  return (Remove-DestinationForOverwrite -Path $Path -UseTrash $UseTrash -TrashDir $TrashDir)
}

function Move-DestinationToSourceIfMissing {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$SourcePath,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$DestinationPath,

    [Parameter(Mandatory = $true)]
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
      Invoke-RobocopySafe -Source $DestinationPath -Destination $SourcePath -Arguments @("/E", "/R:1", "/W:1", "/NFL", "/NDL")
    }
    else {
      Copy-Item -LiteralPath $DestinationPath -Destination $SourcePath -Force
    }
  }
  catch {
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
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$Destination,

    [Parameter(Mandatory = $true)]
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
  }
  catch {
    return $false
  }
}

function Test-ShouldRemoveManagedLink {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$Destination,

    [Parameter(Mandatory = $true)]
    [pscustomobject]$StateEntry
  )

  # Safe removal: remove only if destination still matches what we managed.
  if (!(Test-Path -LiteralPath $Destination)) { return $false }

  $modeValue = $StateEntry.PSObject.Properties['mode']?.Value
  $fromValue = $StateEntry.PSObject.Properties['from']?.Value

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

  # Unknown modes are treated conservatively.

  # We do not auto-remove seed destinations.
  return $false
}

function Get-TrackedItemStatus {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
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
    # no special handling for transfer modes.
    default {
      # For seed and unknown modes, existence is the best available signal.
      return "OK"
    }
  }
}

function Invoke-RobocopySafe {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$Source,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$Destination,

    [Parameter(Mandatory = $true)]
    [string[]]$Arguments
  )

  & robocopy $Source $Destination @Arguments | Out-Null
  $code = $LASTEXITCODE
  if ($code -ge 8) { throw "robocopy failed with exit code $code" }
}

function Write-LogLine {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [string]$Tag,

    [Parameter(Mandatory = $true)]
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
    [Parameter(Mandatory = $true)]
    [string]$PackageName
  )

  Write-Host ""
  Write-LogLine -Tag "package" -Message $PackageName -Color Cyan
}

function Write-ItemHeader {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [string]$Mode,

    [Parameter(Mandatory = $true)]
    [string]$From,

    [Parameter(Mandatory = $true)]
    [string]$To
  )

  Write-LogLine -Tag "item" -Message $Mode -Color White -Indent 2
  Write-Host ("    FROM: {0}" -f $From) -ForegroundColor DarkGray
  Write-Host ("    TO  : {0}" -f $To) -ForegroundColor DarkGray
}

# ---------------- Load config ----------------

$config = (Get-Content -LiteralPath $ConfigPath -Raw) | ConvertFrom-Json
if (-not $config.global) { throw "Config must contain 'global'." }
if (-not $config.packages) { throw "Config must contain 'packages'." }

# Safely access mode property with diagnostics
$modeProperty =$config.global.PSObject.Properties['mode']?.Value
if ($null -eq $modeProperty) {
  Write-Host "WARN: global.mode not found or empty. Available properties in global:" -ForegroundColor Yellow
  $config.global | Get-Member -MemberType NoteProperty | ForEach-Object { Write-Host "  - $($_.Name)" }
  $globalMode = "symlink"
}
else {
  $globalMode = ([string]$modeProperty).ToLower()
}

$globalTrash =[bool]($config.global.PSObject.Properties['trash']?.Value)
$globalTrashDirVal =$config.global.PSObject.Properties['trashDir']?.Value

if ($globalTrash) {
  if ([string]::IsNullOrWhiteSpace([string]$globalTrashDirVal)) {
    $globalTrashDir = Resolve-DotmngrPath -Path "%USERPROFILE%\Trash\dotmngr"
  }
  else {
    $globalTrashDir = Resolve-DotmngrPath -Path ([string]$globalTrashDirVal)
  }
}
else {
  $globalTrashDir = ""
}

# Create trash dir upfront if needed
if ($globalTrash -and -not [string]::IsNullOrWhiteSpace($globalTrashDir)) {
  New-DirectoryIfMissing -Path $globalTrashDir
}


$userProfile = [System.Environment]::GetFolderPath('UserProfile')
$stateDir = Join-Path $userProfile ".config\dotmngr"
New-DirectoryIfMissing -Path $stateDir

$configFull = (Resolve-Path -LiteralPath $ConfigPath).Path
$configBase = [System.IO.Path]::GetFileNameWithoutExtension($configFull)
$statePath = Join-Path $stateDir ("state.{0}.json" -f $configBase)

$state = [pscustomobject]@{
  updated  = $null
  config   = $configFull
  packages = [pscustomobject]@{}
}

if (Test-Path -LiteralPath $statePath) {
  try {
    $loaded = (Get-Content -LiteralPath $statePath -Raw) | ConvertFrom-Json
    if ($loaded) {
      $loadedPackages = $loaded.PSObject.Properties['packages']?.Value
      if ($loadedPackages) {
        if ($loadedPackages -is [hashtable]) {
          $state.packages = [pscustomobject]$loadedPackages
        }
        else {
          $state.packages = $loadedPackages
        }
      }
    }
  }
  catch {
    Write-LogLine -Tag "warn" -Message ("couldn't parse state file, starting fresh: {0}" -f $statePath) -Color Yellow
  }
}

function Get-StatePackage {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
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
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$Name
  )

  $pkgState = Get-StatePackage -Name $Name
  return $pkgState.links
}

function Get-StateLinkEntries {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
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
    [Parameter(Mandatory = $true)]
    [pscustomobject]$StateEntry,

    [Parameter(Mandatory = $true)]
    [bool]$UseTrash,

    [Parameter()]
    [string]$TrashDir = "",

    [Parameter()]
    [ValidateSet("cleanup", "bullet")]
    [string]$PathOutputStyle = "cleanup"
  )

  $oldToVal = $StateEntry.PSObject.Properties['to']?.Value
  if ([string]::IsNullOrWhiteSpace($oldToVal)) { return $true }

  $oldTo = Resolve-DotmngrPath -Path ([string]$oldToVal)

  if ($PathOutputStyle -eq "bullet") {
    Write-Host ("  - {0}" -f $oldTo) -ForegroundColor White
  }
  else {
    Write-LogLine -Tag "cleanup" -Message $oldTo -Color Cyan -Indent 2
  }

  $entryMode = ([string]($StateEntry.PSObject.Properties['mode']?.Value)).ToLower()
  if ($entryMode -eq "seed") {
    Write-LogLine -Tag "untrack" -Message "seed mode; leaving paths untouched." -Color Cyan -Indent 4
    return $true
  }

  if (Test-ShouldRemoveManagedLink -Destination $oldTo -StateEntry $StateEntry) {
    $entryMode = [string]($StateEntry.PSObject.Properties['mode']?.Value)
    $entrySource = [string]($StateEntry.PSObject.Properties['from']?.Value)
    $removed = Remove-ManagedDestination -Path $oldTo -UseTrash $UseTrash -TrashDir $TrashDir -ManagedMode $entryMode -ManagedSource $entrySource
    if (-not $removed -and (Test-Path -LiteralPath $oldTo)) {
      Write-LogLine -Tag "warn" -Message "removal failed; keeping state entry." -Color Yellow -Indent 4
      return $false
    }
  }
  else {
    Write-LogLine -Tag "warn" -Message "not removing (destination no longer matches managed link)." -Color Yellow -Indent 4
  }

  return $true
}

function Remove-StatePackageIfEmpty {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [string]$PackageName,

    [Parameter(Mandatory = $true)]
    $LinksObject
  )

  $linkCount = @($LinksObject.PSObject.Properties).Count
  if ($linkCount -eq 0) {
    $state.packages.PSObject.Properties.Remove($PackageName)
  }
  else {
    Write-LogLine -Tag "warn" -Message ("state retained for package '{0}' because some items could not be removed." -f $PackageName) -Color Yellow -Indent 2
  }
}

function Invoke-PackageCleanupFromState {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [string]$PackageName,

    [Parameter(Mandatory = $true)]
    [string]$Tag,

    [Parameter(Mandatory = $true)]
    [string]$Reason,

    [Parameter(Mandatory = $true)]
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
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$DestinationPath,

    [Parameter(Mandatory = $true)]
    [string]$ManagedMode,

    [Parameter(Mandatory = $true)]
    [string]$ManagedSource,

    [Parameter(Mandatory = $true)]
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
    [Parameter(Mandatory = $true)]
    $LinksObject,

    [Parameter(Mandatory = $true)]
    [string]$DestinationPath,

    [Parameter(Mandatory = $true)]
    [string]$SourcePath,

    [Parameter(Mandatory = $true)]
    [string]$Mode
  )

  # Check if link already exists and has same properties
  $existing = $LinksObject.PSObject.Properties[$DestinationPath]
  if ($existing) {
    $existingValue = $existing.Value
    if ($existingValue.to -eq $DestinationPath -and
      $existingValue.from -eq $SourcePath -and
      $existingValue.mode -eq $Mode) {
      # No change, return false
      return $false
    }
  }

  # Link is new or changed; update it with new timestamp
  $LinksObject | Add-Member -MemberType NoteProperty -Name $DestinationPath -Value ([pscustomobject]@{
      to      = $DestinationPath
      from    = $SourcePath
      mode    = $Mode
      updated = (Get-Date).ToString("o")
    }) -Force
  
  return $true
}

function Invoke-SeedApplyIfNeeded {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [string]$Mode,

    [Parameter(Mandatory = $true)]
    [string]$SourcePath,

    [Parameter(Mandatory = $true)]
    [string]$DestinationPath
  )

  if ($Mode -ne "seed") { return $false }

  if (Test-Path -LiteralPath $DestinationPath) {
    Write-LogLine -Tag "skip" -Message "destination exists, skipping (seed)." -Color Green -Indent 4
    return $true
  }

  if (Test-Path -LiteralPath $SourcePath -PathType Container) {
    New-DirectoryIfMissing -Path $DestinationPath
    Invoke-RobocopySafe -Source $SourcePath -Destination $DestinationPath -Arguments @("/E", "/R:1", "/W:1", "/NFL", "/NDL")
    Write-LogLine -Tag "create" -Message "directory copied once." -Color Green -Indent 4
  }
  else {
    Copy-Item -LiteralPath $SourcePath -Destination $DestinationPath -Force
    Write-LogLine -Tag "create" -Message "file copied once." -Color Green -Indent 4
  }

  return $true
}

# copy-by-hash and transfer helpers removed

function Get-ApplyDestinationDecision {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [string]$Mode,

    [Parameter(Mandatory = $true)]
    [string]$SourcePath,

    [Parameter(Mandatory = $true)]
    [string]$DestinationPath,

    [Parameter(Mandatory = $true)]
    [bool]$UseTrash,

    [Parameter()]
    [string]$TrashDir = ""
  )

  $result = [pscustomobject]@{
    Proceed     = $true
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
    [Parameter(Mandatory = $true)]
    [string]$Mode,

    [Parameter(Mandatory = $true)]
    [string]$SourcePath,

    [Parameter(Mandatory = $true)]
    [string]$DestinationPath,

    [Parameter(Mandatory = $true)]
    [string]$ToKey,

    [Parameter(Mandatory = $true)]
    $LinksObject,

    [Parameter(Mandatory = $true)]
    [pscustomobject]$DesiredItem
  )

  switch ($Mode) {
    "hardlink" {
      if (Test-Path -LiteralPath $SourcePath -PathType Container) {
        throw "hardlink mode only supports files: $SourcePath"
      }
      New-Item -ItemType HardLink -Path $DestinationPath -Target $SourcePath -ErrorAction Stop | Out-Null
      Write-LogLine -Tag "create" -Message "hardlink created." -Color Green -Indent 4
      Set-TrackedLinkState -LinksObject $LinksObject -DestinationPath $ToKey -SourcePath $SourcePath -Mode $Mode | Out-Null
    }
    "symlink" {
      New-Item -ItemType SymbolicLink -Path $DestinationPath -Target $SourcePath -ErrorAction Stop | Out-Null
      Write-LogLine -Tag "create" -Message "symlink created." -Color Green -Indent 4
      Set-TrackedLinkState -LinksObject $LinksObject -DestinationPath $ToKey -SourcePath $SourcePath -Mode $Mode | Out-Null
    }
    "junction" {
      New-Item -ItemType Junction -Path $DestinationPath -Target $SourcePath -ErrorAction Stop | Out-Null
      Write-LogLine -Tag "create" -Message "junction created." -Color Green -Indent 4
      Set-TrackedLinkState -LinksObject $LinksObject -DestinationPath $ToKey -SourcePath $SourcePath -Mode $Mode | Out-Null
    }
    "shortcut" {
      $scParams = @{ Path = $DestinationPath; TargetPath = $SourcePath }
      $scWorkDir = $DesiredItem.PSObject.Properties['workingDirectory']?.Value
      $scArgs = $DesiredItem.PSObject.Properties['arguments']?.Value
      $scDesc = $DesiredItem.PSObject.Properties['description']?.Value
      $scIcon = $DesiredItem.PSObject.Properties['iconLocation']?.Value
      $scWinStyle = $DesiredItem.PSObject.Properties['windowStyle']?.Value
      if ($scWorkDir) { $scParams.WorkingDirectory = [System.Environment]::ExpandEnvironmentVariables($scWorkDir) }
      if ($scArgs) { $scParams.Arguments = [System.Environment]::ExpandEnvironmentVariables($scArgs) }
      if ($scDesc) { $scParams.Description = $scDesc }
      if ($scIcon) { $scParams.IconLocation = [System.Environment]::ExpandEnvironmentVariables($scIcon) }
      if ($null -ne $scWinStyle) { $scParams.WindowStyle = Resolve-WindowStyle $scWinStyle }
      New-WindowsShortcut @scParams
      Write-LogLine -Tag "create" -Message "shortcut created." -Color Green -Indent 4
      Set-TrackedLinkState -LinksObject $LinksObject -DestinationPath $ToKey -SourcePath $SourcePath -Mode $Mode | Out-Null
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
    [Parameter(Mandatory = $true)]
    [pscustomobject]$PackageObject
  )

  $enabledVal = $PackageObject.PSObject.Properties['enabled']?.Value
  if ($null -eq $enabledVal) { return $true }
  return [bool]$enabledVal
}

# Determine packages to run
$selectedPackages = @()
if ($Package -and $Package.Count -gt 0) {
  # Explicit selection (even if disabled)
  $selectedPackages = @($Package)
}
else {
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
    $pkgName = $pkgProp.Name
    $pkgLinks =$pkgProp.Value.PSObject.Properties['links']?.Value
    if ($null -eq $pkgLinks) { continue }

    foreach ($linkProp in $pkgLinks.PSObject.Properties) {
      $entry = $linkProp.Value
      $toVal = $entry.PSObject.Properties['to']?.Value
      $fromVal = $entry.PSObject.Properties['from']?.Value
      $modeVal = $entry.PSObject.Properties['mode']?.Value

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
  }
  else {
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
      }
      elseif ($driftedCount -gt 0) {
        $headerColor = "Yellow"
      }

      Write-Host ""
      Write-LogLine -Tag "package" -Message ("{0} ({1})" -f $pkgGroup.Name, $summary) -Color $headerColor

      $pkgRows |
      Sort-Object -Property @{ Expression = {
          switch ($_.Status) {
            "DRIFTED" { 0 }
            "MISSING" { 1 }
            "OK" { 2 }
            default { 3 }
          }
        } 
      }, Mode, To |
      Format-Table -AutoSize -Property Mode, Status, To
    }
  }
  return
}

# ----------- Admin elevation helpers --------

function Test-ItemNeedsAdmin {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    $ItemObject,
    
    [Parameter(Mandatory = $true)]
    $PackageObject
  )

  # Check item-level admin flag first
  $itemAdmin = $ItemObject.PSObject.Properties['admin']?.Value
  if ($null -ne $itemAdmin -and [bool]$itemAdmin) { return $true }

  # Check package-level admin flag
  $pkgAdmin = $PackageObject.PSObject.Properties['admin']?.Value
  if ($null -ne $pkgAdmin -and [bool]$pkgAdmin) { return $true }

  return $false
}

function Invoke-AdminElevatedLinks {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [array]$AdminDesiredItems,

    [Parameter(Mandatory = $true)]
    [bool]$UseTrash,

    [Parameter()]
    [string]$TrashDir = "",

    [Parameter(Mandatory = $true)]
    $LinksObject
  )

  Write-Host ""
  Write-LogLine -Tag "admin" -Message ("batching {0} admin items for elevated execution..." -f $AdminDesiredItems.Count) -Color Cyan

  $tempPayloadFile = $null
  $tempResultFile = $null
  try {
    $tempPayloadFile = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "dotmngr_admin_payload_$(Get-Random).json")
    $tempResultFile = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "dotmngr_admin_result_$(Get-Random).json")

    $payloadItems = @()
    foreach ($item in $AdminDesiredItems) {
      $payloadItems += [pscustomobject]@{
        toKey = $item.toKey
        mode  = $item.mode
        from  = $item.from
        to    = $item.to
      }
    }

    $payloadItems | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $tempPayloadFile -Encoding UTF8

    $escapedPayloadFile = $tempPayloadFile.Replace("'", "''")
    $escapedResultFile = $tempResultFile.Replace("'", "''")

    $launcherScript = @"
`$payloadFile = '$escapedPayloadFile'
`$tempResultFile = '$escapedResultFile'

`$results = @{}

function New-ItemSafe {
  param([string]`$Mode, [string]`$Path, [string]`$Target)
  try {
    `$parent = [System.IO.Path]::GetDirectoryName(`$Path)
    if (`$parent -and -not (Test-Path -LiteralPath `$parent)) {
      New-Item -ItemType Directory -Path `$parent -Force -ErrorAction Stop | Out-Null
    }

    switch ((`$Mode).ToLower()) {
      'hardlink' { New-Item -ItemType HardLink -Path `$Path -Target `$Target -Force -ErrorAction Stop | Out-Null }
      'junction' { New-Item -ItemType Junction -Path `$Path -Target `$Target -Force -ErrorAction Stop | Out-Null }
      'symlink' { New-Item -ItemType SymbolicLink -Path `$Path -Target `$Target -Force -ErrorAction Stop | Out-Null }
      'shortcut' {
        `$wsh = New-Object -ComObject WScript.Shell
        `$sc = `$wsh.CreateShortcut(`$Path)
        `$sc.TargetPath = `$Target
        `$sc.Save()
      }
      default { New-Item -ItemType `$Mode -Path `$Path -Target `$Target -Force -ErrorAction Stop | Out-Null }
    }

    return @{ success = `$true; error = `$null }
  }
  catch {
    return @{ success = `$false; error = `$_.Exception.Message }
  }
}

`$items = @(Get-Content -LiteralPath `$payloadFile -Raw | ConvertFrom-Json)
foreach (`$item in `$items) {
  `$results[`$item.toKey] = New-ItemSafe -Mode `$item.mode -Path `$item.to -Target `$item.from
}

`$results | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath `$tempResultFile -Encoding UTF8
"@

    $encodedCommand = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($launcherScript))

    # Elevate and run script
    Write-LogLine -Tag "admin" -Message "requesting admin permission..." -Color Yellow
    $proc = Start-Process -FilePath "pwsh" -ArgumentList @(
      "-NoProfile",
      "-ExecutionPolicy", "Bypass",
      "-EncodedCommand", $encodedCommand
    ) -Verb RunAs -Wait -PassThru

    if ($proc.ExitCode -ne 0) {
      Write-LogLine -Tag "error" -Message "admin elevation failed or was canceled." -Color Red
      return $false
    }

    # Parse results
    if (!(Test-Path -LiteralPath $tempResultFile)) {
      Write-LogLine -Tag "error" -Message "admin script did not produce results file." -Color Red
      return $false
    }

    $results = Get-Content -LiteralPath $tempResultFile -Raw | ConvertFrom-Json
    $allSuccess = $true

    foreach ($toKey in $results.PSObject.Properties.Name) {
      $result = $results.$toKey
      $item = $AdminDesiredItems | Where-Object { $_.toKey -eq $toKey } | Select-Object -First 1

      if ($result.success) {
        Write-LogLine -Tag "create" -Message ("{0} created (elevated)." -f $item.mode) -Color Green -Indent 4
        Set-TrackedLinkState -LinksObject $LinksObject -DestinationPath $toKey -SourcePath $item.from -Mode $item.mode | Out-Null
      }
      else {
        Write-LogLine -Tag "error" -Message ("failed to create {0}: {1}" -f $item.mode, $result.error) -Color Red -Indent 4
        $allSuccess = $false
      }
    }

    # Clean up result file
    Remove-Item -LiteralPath $tempResultFile -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $tempPayloadFile -Force -ErrorAction SilentlyContinue

    return $allSuccess
  }
  finally {
    if ($tempPayloadFile -and (Test-Path -LiteralPath $tempPayloadFile)) {
      Remove-Item -LiteralPath $tempPayloadFile -Force -ErrorAction SilentlyContinue
    }
    if ($tempResultFile -and (Test-Path -LiteralPath $tempResultFile)) {
      Remove-Item -LiteralPath $tempResultFile -Force -ErrorAction SilentlyContinue
    }
  }
}

# Initialize state modification trackers (used in both unlink and apply modes)
$unlinkStateWasModified = $false
$applyStateWasModified = $false

# ---------------- Unlink / Relink mode ----------------

if ($Unlink -or $Relink) {
  $unlinkPkgs = @()
  if ($Package -and $Package.Count -gt 0) {
    $unlinkPkgs = @($Package)
  }
  else {
    $unlinkPkgs = @()
    foreach ($prop in $state.packages.PSObject.Properties) {
      $unlinkPkgs += $prop.Name
    }
  }

  foreach ($pkg in $unlinkPkgs) {
    if (-not $state.packages.PSObject.Properties[$pkg]) {
      $tag = if ($Relink) { "relink" } else { "unlink" }
      Write-LogLine -Tag $tag -Message ("{0} (no state)" -f $pkg) -Color Yellow
      continue
    }

    $tag = if ($Relink) { "relink" } else { "unlink" }
    Write-LogLine -Tag $tag -Message $pkg -Color Cyan
    $links = Get-StatePackageLinks -Name $pkg

    $toKeys = Get-StateLinkEntries -LinksObject $links
    $fileEntries = @()
    $containerEntries = @()

    foreach ($toKeyInfo in $toKeys) {
      $resolvedTo = Resolve-DotmngrPath -Path $toKeyInfo.Name
      if (Test-Path -LiteralPath $resolvedTo -PathType Container) {
        $containerEntries += $toKeyInfo
      }
      else {
        $fileEntries += $toKeyInfo
      }
    }

    $toKeys = @($fileEntries + $containerEntries)

    foreach ($toKeyInfo in $toKeys) {
      $toKey = $toKeyInfo.Name
      $old = $toKeyInfo.Value

      $removeStateEntry = Invoke-TrackedEntryCleanup -StateEntry $old -UseTrash $globalTrash -TrashDir $globalTrashDir -PathOutputStyle "bullet"
      if ($removeStateEntry) {
        $links.PSObject.Properties.Remove($toKey)
        $unlinkStateWasModified = $true
      }
    }

    Remove-StatePackageIfEmpty -PackageName $pkg -LinksObject $links
  }

  if ($unlinkStateWasModified) {
    $state.updated = (Get-Date).ToString("o")
  }
  
  # Only save state immediately if doing -Unlink without -Relink
  if ($Unlink -and -not $Relink) {
    $state | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $statePath -Encoding UTF8
    Write-Host ""
    Write-LogLine -Tag "save" -Message ("state saved: {0}" -f $statePath) -Color Green
    return
  }
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

  $pkgModeValue = $pkgObj.PSObject.Properties['mode']?.Value
  $pkgMode = if ($pkgModeValue) { ([string]$pkgModeValue).ToLower() } else { $globalMode }
  $useTrash = $globalTrash
  $trashDir = $globalTrashDir

  Write-PackageHeader -PackageName $pkg

  $desired = @{} # toResolved -> {to,from,mode}

  foreach ($it in $pkgObj.items) {
    $itModeValue = $it.PSObject.Properties['mode']?.Value
    $mode = if ($itModeValue) { ([string]$itModeValue).ToLower() } else { $pkgMode }

    $itTo = $it.PSObject.Properties['to']?.Value
    $itFrom = $it.PSObject.Properties['from']?.Value

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
    }
    catch {
      Write-ItemHeader -Mode $mode -From $fromExpanded -To $toExpanded

      $isRecoverableLinkMode = ($mode -eq "symlink" -or $mode -eq "junction" -or $mode -eq "hardlink")
      if ($isRecoverableLinkMode) {
        $recovered = Move-DestinationToSourceIfMissing -SourcePath $fromExpanded -DestinationPath $toExpanded -Mode $mode -UseTrash $useTrash -TrashDir $trashDir
        if ($recovered) {
          try {
            $fromResolved = (Resolve-Path -LiteralPath $fromExpanded).Path
          }
          catch {
            Write-LogLine -Tag "warn" -Message "recovery attempted, but source still missing, skipping." -Color Yellow -Indent 4
            continue
          }
        }
        else {
          Write-LogLine -Tag "warn" -Message "source missing, skipping." -Color Yellow -Indent 4
          continue
        }
      }
      else {
        Write-LogLine -Tag "warn" -Message "source missing, skipping." -Color Yellow -Indent 4
        continue
      }
    }

    $itemAdmin = $it.PSObject.Properties['admin']?.Value
    if ($null -eq $itemAdmin) { $itemAdmin = $pkgObj.PSObject.Properties['admin']?.Value }
    $desiredEntry = [pscustomobject]@{ to = $toExpanded; from = $fromResolved; mode = $mode; admin = [bool]$itemAdmin }
    if ($mode -eq "shortcut") {
      $scWorkDir = $it.PSObject.Properties['workingDirectory']?.Value
      $scArgs = $it.PSObject.Properties['arguments']?.Value
      $scDesc = $it.PSObject.Properties['description']?.Value
      $scIcon = $it.PSObject.Properties['iconLocation']?.Value
      $scWinStyle = $it.PSObject.Properties['windowStyle']?.Value
      if ($null -ne $scWorkDir) { $desiredEntry | Add-Member -MemberType NoteProperty -Name workingDirectory -Value ([string]$scWorkDir) }
      if ($null -ne $scArgs) { $desiredEntry | Add-Member -MemberType NoteProperty -Name arguments        -Value ([string]$scArgs) }
      if ($null -ne $scDesc) { $desiredEntry | Add-Member -MemberType NoteProperty -Name description      -Value ([string]$scDesc) }
      if ($null -ne $scIcon) { $desiredEntry | Add-Member -MemberType NoteProperty -Name iconLocation     -Value ([string]$scIcon) }
      if ($null -ne $scWinStyle) { $desiredEntry | Add-Member -MemberType NoteProperty -Name windowStyle      -Value (Resolve-WindowStyle $scWinStyle) }
    }
    $desired[$toExpanded] = $desiredEntry
  }

  # Cleanup managed destinations removed from this package
  $links = Get-StatePackageLinks -Name $pkg
  $packageWasModified = $false

  $propNames = Get-StateLinkEntries -LinksObject $links

  foreach ($propInfo in $propNames) {
    $toKey = $propInfo.Name
    if ($desired.ContainsKey($toKey)) {
      $desiredModeVal =$desired[$toKey].PSObject.Properties['mode']?.Value
      $desiredMode = if ($desiredModeVal) { ([string]$desiredModeVal).ToLower() } else { "" }

      if ($desiredMode -eq "seed") {
        # seed is intentionally untracked; drop any stale tracked state entry.
        Write-LogLine -Tag "untrack" -Message ("{0} (seed mode)" -f $toKey) -Color Cyan -Indent 2
        $links.PSObject.Properties.Remove($toKey)
        $packageWasModified = $true
      }
      continue
    }

    $old = $propInfo.Value
    $removeStateEntry = Invoke-TrackedEntryCleanup -StateEntry $old -UseTrash $useTrash -TrashDir $trashDir
    if ($removeStateEntry) {
      $links.PSObject.Properties.Remove($toKey)
      $packageWasModified = $true
    }
  }

  # Apply desired - partition into admin and non-admin items
  $adminItems = @()
  $regularItems = @()
  
  foreach ($toKey in $desired.Keys) {
    $it = $desired[$toKey]
    $needsAdmin = Test-ItemNeedsAdmin -ItemObject $it -PackageObject $pkgObj
    
    if ($needsAdmin) {
      $adminItems += @{ toKey = $toKey; item = $it }
    }
    else {
      $regularItems += @{ toKey = $toKey; item = $it }
    }
  }

  # Process admin items first (elevated batch)
  if ($adminItems.Count -gt 0) {
    $adminDesired = @()

    foreach ($entry in $adminItems) {
      $toKey = $entry.toKey
      $it = $entry.item
      $mode = $it.PSObject.Properties['mode']?.Value
      $from = $it.PSObject.Properties['from']?.Value
      $to = $it.PSObject.Properties['to']?.Value

      try {
        Write-ItemHeader -Mode $mode -From $from -To $to

        # proceed with normal handling

        # Don't create parent directories here; elevation will handle creation where needed.

        if ($Force -and (Test-Path -LiteralPath $to)) {
          Write-LogLine -Tag "replace" -Message "removing destination before apply." -Color Yellow -Indent 4
          if (-not (Test-ReplaceDestinationRemoval -DestinationPath $to -ManagedMode $mode -ManagedSource $from -UseTrash $useTrash -TrashDir $trashDir)) {
            continue
          }
        }

        if (Invoke-SeedApplyIfNeeded -Mode $mode -SourcePath $from -DestinationPath $to) {
          continue
        }

        # Pre-flight decision
        $decision = Get-ApplyDestinationDecision -Mode $mode -SourcePath $from -DestinationPath $to -UseTrash $useTrash -TrashDir $trashDir
        if (-not $decision.Proceed) {
          Write-LogLine -Tag "skip" -Message "skipped (pre-flight check failed)." -Color Yellow -Indent 4
          continue
        }

        $needsCreate = [bool]$decision.NeedsCreate
        if (-not $needsCreate) {
          if (Set-TrackedLinkState -LinksObject $links -DestinationPath $toKey -SourcePath $from -Mode $mode) {
            $packageWasModified = $true
          }
          continue
        }

        # Queue for elevated creation
        $adminDesired += @{
          toKey = $toKey
          mode  = $mode
          from  = $from
          to    = $to
        }
      }
      catch {
        $errMsg = $_.Exception.Message
        Write-LogLine -Tag "error" -Message ("failed to preflight admin item (mode={0}): {1}" -f $mode, $errMsg) -Color Red -Indent 4
        continue
      }
    }

    if ($adminDesired.Count -gt 0) {
      $adminSuccess = Invoke-AdminElevatedLinks -AdminDesiredItems $adminDesired -UseTrash $useTrash -TrashDir $trashDir -LinksObject $links
      if ($adminSuccess) {
        $packageWasModified = $true
      }
    }
  }

  # Then process regular (non-admin) items
  foreach ($entry in $regularItems) {
    $toKey = $entry.toKey
    $it = $entry.item
    $mode = $it.PSObject.Properties['mode']?.Value
    $from = $it.PSObject.Properties['from']?.Value
    $to = $it.PSObject.Properties['to']?.Value

    try {
      Write-ItemHeader -Mode $mode -From $from -To $to

      # proceed with normal handling

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
        if (Set-TrackedLinkState -LinksObject $links -DestinationPath $toKey -SourcePath $from -Mode $mode) {
          $packageWasModified = $true
        }
        continue
      }

      Invoke-CreateTrackedLinkForMode -Mode $mode -SourcePath $from -DestinationPath $to -ToKey $toKey -LinksObject $links -DesiredItem $it
      $packageWasModified = $true
    }
    catch {
      $errMsg = $_.Exception.Message
      Write-LogLine -Tag "error" -Message ("failed to process item (mode={0}): {1}" -f $mode, $errMsg) -Color Red -Indent 4
      continue
    }
  }

  # Stamp package updated time only if changes occurred
  if ($packageWasModified) {
    $pkgState = Get-StatePackage -Name $pkg
    $pkgState.updated = (Get-Date).ToString("o")
    $applyStateWasModified = $true
  }
}


# ---------------- Save state ----------------

if ($applyStateWasModified -or $unlinkStateWasModified) {
  $state.updated = (Get-Date).ToString("o")
}
$state.config = $configFull
$state | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $statePath -Encoding UTF8
Write-Host ""
Write-LogLine -Tag "save" -Message ("state saved: {0}" -f $statePath) -Color Green
