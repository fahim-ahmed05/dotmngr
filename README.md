# Dot Manager

A simple PowerShell script that manages dotfiles scattered across your system using a JSON manifest. It can create links (symlink/junction/hardlink) or manage one-way file/directory transfers (`push` or `seed`). It also keeps a state file so removed/moved entries in your config don’t leave stale links behind.

## Features

* Supports:

  * `symlink` (file/dir)
  * `junction` (dir)
  * `hardlink` (file)
  * `push` (one-way push via `robocopy`, skip overwriting newer destination)
  * `seed` (copy only if destination doesn’t exist)  
  * `shortcut` (Windows `.lnk` shortcut)
  
* Safe cleanup: If an entry is removed from config, it removes the destination **only if it still matches what the script created**
* State tracking stored under: `~\.config\dotmngr\state.<configName>.json`


## Requirements

* PowerShell (Windows PowerShell 5.1 or PowerShell 7+)
* `robocopy` (built into Windows)

### Symlink permissions

Creating `symlink` may require Admin **unless** Windows **Developer Mode** is enabled:

Settings → Privacy & Security → For Developers → **Developer Mode** → ON

> `junction` and `hardlink` do not require Developer Mode.


## Install

1. Save the script as `dotmngr.ps1`
2. Create a config file like `dotlinks.json`


## Usage

### Apply enabled packages

```powershell
.\dotmngr.ps1 -ConfigPath .\dotlinks.json
```

When `enabled` is switched from `true` to `false`, a normal apply run treats that package as removed: previously managed links are safely cleaned up (only if they still match what dotmngr created), and its state is dropped.

### Apply specific packages

```powershell
.\dotmngr.ps1 -ConfigPath .\dotlinks.json -Package nvim,git
```

> This runs even if `enabled: false`.

### Force reapply all selected items

Removes existing destinations first, then reapplies each item regardless of current state:

```powershell
.\dotmngr.ps1 -ConfigPath .\dotlinks.json -Force
```

You can combine it with package selection:

```powershell
.\dotmngr.ps1 -ConfigPath .\dotlinks.json -Package nvim -Force
```

### Unlink packages

Removes managed links for selected packages:

```powershell
.\dotmngr.ps1 -ConfigPath .\dotlinks.json -Package nvim -Unlink
```

### Unlink everything

```powershell
.\dotmngr.ps1 -ConfigPath .\dotlinks.json -Unlink
```

## Config Structure

Your config **must** use this format:

```json
{
  "global": {
    "mode": "junction",
    "trash": true,
    "trashDir": "%LOCALAPPDATA%\\DotLinksTrash"
  },

  "packages": {
    "nvim": {
      "enabled": true,
      "mode": "junction",
      "items": [
        {
          "from": "%USERPROFILE%\\Git\\dotfiles\\nvim",
          "to": "%LOCALAPPDATA%\\nvim"
        }
      ]
    },
    "winutil": {
      "enabled": true,
      "items": [
        {
          "mode": "shortcut",
          "from": "%USERPROFILE%\\Git\\scripts\\launch.ps1",
          "to": "%USERPROFILE%\\Desktop\\Launch.lnk",
          "workingDirectory": "%USERPROFILE%\\Git\\scripts",
          "arguments": "-NoProfile",
          "description": "Launch my script",
          "iconLocation": "%SYSTEMROOT%\\System32\\WindowsPowerShell\\v1.0\\powershell.exe,0",
          "windowStyle": 1
        }
      ]
    }
  }
}
```

## Environment Variable Support

`dotmngr` supports **Windows-style environment variables only**:

Supported:

* `%USERPROFILE%`
* `%LOCALAPPDATA%`
* `%APPDATA%`
* Any `%ENV_VAR%`

Not supported:

* `~`
* `$env:LOCALAPPDATA`
* `$HOME`

Use `%VAR%` syntax inside JSON.

Example:

```json
{
  "to": "%LOCALAPPDATA%\\nvim"
}
```

## Configuration Reference

### `global`

| Field      | Description                                       |
| ---------- | ------------------------------------------------- |
| `mode`     | Default mode if not specified per package/item    |
| `trash`    | If `true`, replaced files are moved to `trashDir` |
| `trashDir` | Folder where replaced files are stored            |

Behavior notes:

* If `global.trash` is `true` and `global.trashDir` is empty/missing, dotmngr uses `%USERPROFILE%\\Trash\\dotmngr`.
* If `global.trash` is `false`, `trashDir` is ignored and removable non-link items go to Recycle Bin.

### `packages`

Each package:

```json
"packageName": {
  "enabled": true,
  "mode": "junction",
  "items": [
    { "from": "...", "to": "...", "mode": "..." }
  ]
}
```

#### Fields

| Field     | Required | Description             |
| --------- | -------- | ----------------------- |
| `enabled` | No       | Defaults to `true`      |
| `mode`    | No       | Overrides `global.mode` |
| `items`   | Yes      | Array of mappings       |


### Mode Resolution Order

When processing an item:

1. `item.mode`
2. `package.mode`
3. `global.mode`


## Modes Explained

### `symlink`

Creates a symbolic link at `to` pointing to `from`.

* Files and directories supported
* May require Admin unless Developer Mode is enabled

### `junction`

Creates a directory junction at `to` pointing to `from`.

* Directories only (use this for most Windows folder targets)
* No Admin required

### `hardlink`

Creates a hardlink at `to` pointing to `from`.

* Files only (cannot hardlink directories)
* Source and destination must be on the same drive/volume

### `push`

Uses `robocopy` to push changes from `from → to`.

* Uses timestamps + size
* Uses `/XO` so it **won’t overwrite** destination if destination is newer
* Good for “push updates but don’t clobber local edits”

### `seed`

Copies `from → to` only if `to` does not already exist.

* Once `to` exists, it is never changed again by dotmngr
* Useful for “seed a default config” behavior

### `shortcut`

Creates a Windows `.lnk` shortcut at `to` pointing to `from`.

* `to` must end with `.lnk`
* Uses `WScript.Shell` — no Admin required
* Supports additional optional properties per item:

| Property           | Description                                          |
| ------------------ | ---------------------------------------------------- |
| `workingDirectory` | Working directory when launching the shortcut        |
| `arguments`        | Command-line arguments passed to the target          |
| `description`      | Tooltip text shown on the shortcut                   |
| `iconLocation`     | Icon path, e.g. `"notepad.exe,0"` or `"app.exe,1"`  |
| `windowStyle`      | `normal`, `minimized`, `maximized`, or a raw integer (`normal` = 1, `maximized` = 3, `minimized` = 7) |

**Example:**

```json
"shortcuts": {
  "enabled": true,
  "items": [
    {
      "mode": "shortcut",
      "from": "%USERPROFILE%\\Git\\scripts\\launch.ps1",
      "to": "%USERPROFILE%\\Desktop\\Launch.lnk",
      "workingDirectory": "%USERPROFILE%\\Git\\scripts",
      "arguments": "-NoProfile",
      "description": "Launch my script",
      "iconLocation": "%SYSTEMROOT%\\System32\\WindowsPowerShell\\v1.0\\powershell.exe,0",
      "windowStyle": "maximized"
    }
  ]
}
```

## What happens if `to` already exists?

When `-Force` is used:

* Existing `to` is removed first (or moved to trash when enabled) for all modes
* Item is reapplied without using state-based skip behavior

For link modes (`symlink`, `junction`, `hardlink`, `shortcut`):

* If `to` is already a link/hardlink/shortcut pointing to the correct `from` → **skip**
* If `to` exists but is not the correct managed link → it is moved to `trashDir` (if enabled) and replaced
* If `from` is missing but `to` exists as a normal file/folder (`symlink`, `junction`, `hardlink` only), dotmngr moves `to` back to `from`, then creates the link at `to`

For transfer modes:

* `push` updates `to` from `from` via `robocopy`
* `seed` does nothing if `to` exists


## State file behavior

State files live in:

* `~\.config\dotmngr\state.<configName>.json`

Example:

* Config: `dotlinks.json`
* State: `~\.config\dotmngr\state.dotlinks.json`

When you remove an item from your config, dotmngr will attempt to remove the old destination **only if** it still matches what dotmngr previously created. This prevents accidental deletion if you manually replaced that destination with something else.

## Tips

* Prefer `junction` for folders on Windows (it’s the most painless)
* Use `symlink` when you specifically want a true symlink
* Use `hardlink` for single files you want to appear in multiple places on the same drive
* Use `seed` for apps that generate/modify configs and you only want to seed it once
* Use `shortcut` to place `.lnk` files on the Desktop or Start Menu without touching the real application files

## Troubleshooting

### Symlink fails with permission error

Enable Developer Mode or run PowerShell as Administrator.

### Hardlink fails

Hardlinks only work for files on the same drive. If your dotfiles repo is on `D:` and `to` is on `C:`, use `symlink` or `push` instead.

### Robocopy returns weird exit codes

Robocopy uses bitmask exit codes; dotmngr treats codes **>= 8** as failure.

