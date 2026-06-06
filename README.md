# teams-cli

<p align="center">
  <img src="./img/DarkMode_Color.svg" alt="teams-cli logo" width="200" />
</p>

A terminal UI for Microsoft Teams powered by [teams-api](https://github.com/fossteams/teams-api).

Current stable version: `v1.2.0`

<img width="1091" height="641" alt="image" src="https://github.com/user-attachments/assets/8a15d464-74d7-4728-9f74-2b8452c45689" />


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
- clones/updates the repo into `/opt/teams-cli`
- sets ownership to the invoking user (no sudo needed to run)
- installs `teams-cli` into `/usr/local/bin`
- installs desktop icon and `.desktop` launcher
- sets up `teams-token-cli` automatically (yarn, TypeScript, system Electron)

After install, authenticate once:

```bash
cd /opt/teams-cli/teams-token-cli && yarn start
```

Then run:

```bash
teams-cli
```

## Development Run

```bash
go run ./
```

## teams-token Integration

This repo includes `teams-token-cli` as a git submodule for token refresh on `401 Unauthorized`.

```bash
git submodule update --init --recursive
```

On `401`, teams-cli will try to run `teams-token` using:
- local binary (`./teams-token/teams-token`)
- Go (`go run .`)
- Node (`yarn start` or `npm run start`, with install if needed)

If a `401` still reaches the error page, a `Run teams-token` button appears for manual refresh.

## Terminal Auth Attempts (Feb 2026)

We explored "all-in-terminal" auth alternatives and recorded the outcomes:

- **Device Code Flow**: Rejected by Azure AD for the Teams client ID (`5e3ce6c0-2b1f-4285-8d4b-75ee78787346`). Error: `AADSTS70002` ("client must be marked as mobile/public").
- **term.everything + Xwayland**: Electron consistently aborted (futex/SIGABRT) even with sandbox disabled and GPU/dev-shm flags.
- **term.everything + Wayland/Ozone**: Electron failed to initialize Wayland and, in some runs, hit OOM during startup.

### Suggestions For A Terminal-Only Path

1. **Use a device-code-capable client ID**: Register or reuse an Azure AD public client that allows device code. Parameterize the client ID and tenant.
2. **Headless auth helper**: A small CLI that implements device code or PKCE and writes `~/.config/fossteams/token-*.jwt` directly could fully replace Electron.
3. **Token broker**: Run auth in a browser elsewhere and exchange a short-lived code for tokens via a secure local channel.
4. **Fallback-friendly UX**: Keep terminal-first attempts, but auto-fallback to browser/Electron with a clear reason and next steps.

## Features

- Teams + channels listing
- Channel read
- DM/chat read (recent first)
- Send messages in channels and chats
- Chat favorites (`f`)
- Custom chat groups (`g`) — create, rename, delete named groups; move chats between them
- Private Notes chat auto-detected and grouped into Favorites
- Chat title refresh (`u`)
- Unread marker auto-refresh every minute (toggle with `m`, manual scan with `Shift+M`)
- Compose title shows scanner status (`ON/OFF`), scan progress, and last scan result
- Manual mark unread hotkey (`r`) for selected chat
- Keybinding cheatsheet overlay (`?`) — press from tree or chat pane, dismiss with `?` or `Esc`
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
- Theme color customization in `Settings & Help`:
  - `Compose Color` cycles darker blue variants (`midnight`, `navy`, `dark_blue`, `slate`)
  - `Username Color` cycles (`blue`, `yellow`, `green`, `cyan`, `white`)
  - both are stored in encrypted settings
- Message layout:
  - author name displayed above message content
  - send time shown in dark grey next to author
  - blank line between messages
  - "No messages yet." placeholder for empty conversations
- Message reactions display with emoji rendering (👍 ❤️ 😆 😮 😢 😠)
- Reaction picker (`e`) — choose any reaction
- Links and images in messages (`o` to open/view)
  - Images saved to `~/Pictures/teams-cli-screenshots` (configurable in Settings)
  - Terminal image rendering via `chafa` (requires fresh tokens)
- HTML formatting displayed: **bold**, *italic*, underline, ~~strike~~, `code`, blockquotes
- Markdown on send: type `**bold**`, `*italic*`, `` `code` ``, `> quote`, `~~strike~~`
- Send images: type `<img>~/path/to/file.png</img>` in compose
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
  - custom groups
  - updated chat titles
- Encrypted settings files:
  - `~/.config/fossteams/teams-cli-settings.enc`
  - `~/.config/fossteams/teams-cli-settings.key`

## Keybindings

| Key | Context | Action |
|-----|---------|--------|
| `Tab` | global | next pane |
| `Shift+Tab` | global | previous pane |
| `?` | tree / chat | toggle keybinding cheatsheet |
| `i` | tree | focus compose input |
| `f` | tree | toggle favorite for selected chat |
| `g` | tree (chat node) | move chat to group |
| `g` | tree (group header) | rename / delete group |
| `u` | tree | refresh chat titles |
| `r` | tree | mark selected chat unread |
| `m` | tree | toggle 1-minute unread scan |
| `Shift+M` | tree | run unread scan immediately |
| `G` | tree | jump to most recent chat |
| `Ctrl+R` | tree | reload keybindings config |
| `r` | chat | reply to selected message |
| `e` | chat | open reaction picker (👍 ❤️ 😆 😮 😢 😠) |
| `o` | chat | open link or image picker |
| `G` | chat | scroll to bottom (most recent message) |
| `+` / `-` | image view | grow / shrink terminal image |
| `Esc` | image view | close image |
| `Enter` | compose | send message |
| `Esc` | compose | back to tree |

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
- Better unread detection and sync accuracy
- Scroll/pagination for long chat histories

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
