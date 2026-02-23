# Dot Manager

A simple PowerShell script that manages dotfiles scattered across your system using a JSON manifest. It can create links (symlink/junction/hardlink) or copy files/directories (sync or one-time copy). It also keeps a state file so removed/moved entries in your config don’t leave stale links behind.

## Features

* Config-driven: no need to reorganize your dotfiles repo
* Supports:

  * `symlink` (file/dir)
  * `junction` (dir)
  * `hardlink` (file)
  * `copy` (sync via `robocopy`, skip overwriting newer destination)
  * `copyOnce` (copy only if destination doesn’t exist)
* Safe cleanup:

  * If an entry is removed from config, it removes the destination **only if it still matches what the script created**
* State tracking stored under:

  * `~\.config\dotmngr\state.<configName>.json`


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

Run from PowerShell:

```powershell
.\dotmngr.ps1 -ConfigPath .\dotlinks.json
```

Re-run anytime after changing the config. The script is designed to be idempotent.


## Example config

> Notes:
>
> * `from` should be a full path (supports `~`)
> * `to` should be the target location (supports `~`)
> * `defaults.mode` is used if an item doesn’t specify `mode`
> * `trashDir` is where replaced files/folders get moved (when `trash: true`)

```json
{
  "defaults": {
    "mode": "junction",
    "trash": true,
    "trashDir": "~/Trash/DotLinks"
  },
  "items": [
    {
      "from": "~/Git/dotfiles/nvim",
      "to": "~/AppData/Local/nvim",
      "mode": "junction"
    },
    {
      "from": "~/Git/dotfiles/powershell/profile.ps1",
      "to": "~/Documents/PowerShell/Microsoft.PowerShell_profile.ps1",
      "mode": "symlink"
    },
    {
      "from": "~/Git/dotfiles/git/.gitconfig",
      "to": "~/.gitconfig",
      "mode": "hardlink"
    },
    {
      "from": "~/Git/dotfiles/wezterm/.wezterm.lua",
      "to": "~/.wezterm.lua",
      "mode": "copy"
    },
    {
      "from": "~/Git/dotfiles/app/sample-config.json",
      "to": "~/AppData/Roaming/MyApp/config.json",
      "mode": "copyOnce"
    }
  ]
}
```


## Modes explained

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

### `copy`

Uses `robocopy` to sync `from → to`.

* Uses timestamps + size
* Uses `/XO` so it **won’t overwrite** destination if destination is newer
* Good for “deploy but don’t clobber local edits”

### `copyOnce`

Copies `from → to` only if `to` does not already exist.

* Once `to` exists, it is never changed again by dotmngr
* Useful for “seed a default config” behavior


## What happens if `to` already exists?

For link modes (`symlink`, `junction`, `hardlink`):

* If `to` is already a link/hardlink pointing to the correct `from` → **skip**
* If `to` exists but is not the correct managed link → it is moved to `trashDir` (if enabled) and replaced

For copy modes:

* `copy` syncs via `robocopy`
* `copyOnce` does nothing if `to` exists


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
* Use `copyOnce` for apps that generate/modify configs and you only want to “seed” it once


## Troubleshooting

### Symlink fails with permission error

Enable Developer Mode or run PowerShell as Administrator.

### Hardlink fails

Hardlinks only work for files on the same drive. If your dotfiles repo is on `D:` and `to` is on `C:`, use `symlink` or `copy` instead.

### Robocopy returns weird exit codes

Robocopy uses bitmask exit codes; dotmngr treats codes **>= 8** as failure.

## Support

If you find this helpful, consider supporting.

<a href="https://www.buymeacoffee.com/fahim.ahmed" target="_blank">
  <img src="https://www.buymeacoffee.com/assets/img/custom_images/orange_img.png" 
       alt="Buy Me A Coffee"
       style="height: 41px !important; width: 174px !important; box-shadow: 0px 3px 2px 0px rgba(190, 190, 190, 0.5);" />
</a>