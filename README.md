# teams-cli

<p align="center">
  <img src="./img/DarkMode_Color.svg" alt="teams-cli logo" width="200" />
</p>

A terminal UI for Microsoft Teams powered by [teams-api](https://github.com/fossteams/teams-api).

Current stable version: `v1.0.0`

## Install

### Linux one-line install (from GitHub)

Ubuntu / Debian:

```bash
sudo apt update && sudo apt install -y git curl golang-go nodejs npm && curl -fsSL https://raw.githubusercontent.com/Gamedirection/teams-cli/master/scripts/install.sh | sudo bash
```

Fedora:

```bash
sudo dnf install -y git curl golang nodejs npm && curl -fsSL https://raw.githubusercontent.com/Gamedirection/teams-cli/master/scripts/install.sh | sudo bash
```

Arch / Manjaro:

```bash
sudo pacman -Sy --noconfirm git curl go nodejs npm && curl -fsSL https://raw.githubusercontent.com/Gamedirection/teams-cli/master/scripts/install.sh | sudo bash
```

openSUSE:

```bash
sudo zypper --non-interactive in git curl go nodejs npm && curl -fsSL https://raw.githubusercontent.com/Gamedirection/teams-cli/master/scripts/install.sh | sudo bash
```

Universal (dependencies preinstalled):

```bash
curl -fsSL https://raw.githubusercontent.com/Gamedirection/teams-cli/master/scripts/install.sh | sudo bash
```

The installer:
- clones/updates into `/opt/teams-cli`
- installs `teams-cli` into `/usr/local/bin`
- installs desktop icon from `img/DarkMode_Color.svg`
- installs a `teams-cli.desktop` launcher

## Run

```bash
teams-cli
```

## Development Run

```bash
teams-cli-dev
```

Notes:
- `teams-cli` is the compiled binary (fast startup, no Go toolchain required).
- `teams-cli-dev` runs the source via `go run` (slower startup, good for development).

## Auth Refresh Integration

This repo includes `teams-token-cli` and `term.everything` as git submodules for token refresh on `401 Unauthorized`.

```bash
git submodule update --init --recursive
```

On `401`, teams-cli will try to run the auth refresh app using:
- local binary (`./teams-token/teams-token`)
- Go (`go run .`)
- Node (`yarn start` or `npm run start`, with install if needed)

Auth refresh uses device code flow by default (terminal-only). If that fails, it falls back to the existing Electron flow.
Set `TEAMS_CLI_DISABLE_DEVICE_CODE=1` to skip device code and use the Electron flow immediately.
Set `TEAMS_CLI_TENANT=<tenant>` to override the default `common` tenant for device code flow.

If `term.everything` is available, the Electron flow runs inside the terminal by default.
Set `TEAMS_CLI_DISABLE_TERM_EVERYTHING=1` to fall back to direct Electron launch.

If a `401` still reaches the error page, a `Run auth refresh` button appears for manual refresh.

Build helper (optional, to generate a `term.everything` binary in `term.everything/dist/...`):

```bash
./scripts/build-term-everything.sh
```

## Features

- Teams + channels listing
- Channel read
- DM/chat read (recent first)
- Send messages in channels and chats
- Chat favorites (`f`)
- Private Notes chat auto-detected and grouped into Favorites
- Chat title refresh (`u`)
- Unread marker auto-refresh every minute (toggle with `m`, manual scan with `Shift+M`)
- Compose title shows scanner status (`ON/OFF`), scan progress, and last scan result
- Manual mark unread hotkey (`r`) for selected chat
- Built-in `Settings & Help` chat at the bottom of the tree
- In-app keybinding settings menu in `Settings & Help`:
  - Enter on config row opens your `$VISUAL` / `$EDITOR`
  - Enter on preset row cycles `default -> vim -> emacs -> jk`
  - Enter on binding row captures a new single key
  - `Esc` while capturing resets that action to preset default
- Chat text mode toggle in `Settings & Help`:
  - `Word Wrap` (default)
  - `Scroll` (single-line messages)
  - configurable wrap characters (`20/40/72/80/100/200/400/600/800/1000/custom`)
  - actual visible wrap is capped by current chat pane width
  - wrap mode keeps the same message/author secondary styling as scroll mode
- Theme color customization in `Settings & Help`:
  - `Compose Color` cycles darker blue variants (`midnight`, `navy`, `dark_blue`, `slate`)
  - `Username Color` cycles (`blue`, `yellow`, `green`, `cyan`, `white`)
  - both are stored in encrypted settings
- Message reactions display in chat (`Reactions: ...`)
- Quick react hotkey in chat (`e` adds 👍 to selected message, server + local fallback)
- Reply mode in chat (`r` replies to selected message)
- Mentions in compose:
  - `@name` prefers current chat members, then global contacts
  - `c@name` forces global contacts lookup
  - while typing mention token, use `Up/Down` in compose to cycle suggestions and prefill
  - posted mentions keep `@` prefix
- Custom keybindings via config file: `~/.config/fossteams/teams-cli-keybindings.json`
- Keybinding presets: `default`, `vim`, `emacs`, `jk`
- Encrypted persistence of:
  - favorites
  - updated chat titles
- Encrypted settings files:
  - `~/.config/fossteams/teams-cli-settings.enc`
  - `~/.config/fossteams/teams-cli-settings.key`

## Keybindings

- `Tab`: next pane
- `Shift+Tab`: previous pane
- `i`: focus compose input
- `Enter` (compose): send message
- `Esc` (compose): back to tree
- `f`: toggle favorite for selected/hovered chat
- `u`: refresh chat titles
- `r` (tree pane): mark selected chat unread
- `r` (chat pane): reply to selected message
- `e` (chat pane): react 👍 to selected message
- `m`: toggle 1-minute unread scan on/off
- `Shift+M`: run unread scan immediately
- `Ctrl+R`: reload keybindings config without restarting

## Keybinding Config

`teams-cli` creates keybinding config at:

- `~/.config/fossteams/teams-cli-keybindings.json`

Example:

```json
{
  "preset": "vim",
  "bindings": {
    "react_message": ["e"],
    "reply_message": ["r"],
    "toggle_favorite": ["f"]
  }
}
```

Available actions:
- `toggle_scan`
- `scan_now`
- `mark_unread`
- `toggle_favorite`
- `refresh_titles`
- `focus_compose`
- `reply_message`
- `react_message`
- `reload_keybindings`
- `move_down`
- `move_up`

In-app keybinding editor notes:
- Use `Settings & Help` chat and press `Tab` until chat pane is focused.
- Press `Enter` on:
  - `Open Keybindings Config` to launch your editor
  - `Preset` to cycle presets
  - `Bind ...` rows to set a new single key
- Press `Esc` while binding to reset that action to preset default.
- For multi-key or advanced mappings, edit `teams-cli-keybindings.json` directly.

## Feature Roadmap

Planned messaging/UX improvements:
- TTS playback for messages
- Inline image viewer for chat attachments
- Reactions support (view/add)
- Reply/thread support from CLI
- Better unread detection and sync accuracy

## Packaging Roadmap

Planned distribution formats:
- `apt`
- `dnf`
- `pacman`
- Homebrew
- Chocolatey
- AppImage
- Flatpak

## Related

- [fossteams-frontend](https://github.com/fossteams/fossteams-frontend)
