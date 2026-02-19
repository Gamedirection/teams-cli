# teams-cli

<p align="center">
  <img src="./img/DarkMode_Color.svg" alt="teams-cli logo" width="200" />
</p>

A terminal UI for Microsoft Teams powered by [teams-api](https://github.com/fossteams/teams-api).

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
go run ./
```

## teams-token Integration

This repo includes `teams-token` as a git submodule for token refresh on `401 Unauthorized`.

```bash
git submodule update --init --recursive
```

On `401`, teams-cli will try to run `teams-token` using:
- local binary (`./teams-token/teams-token`)
- Go (`go run .`)
- Node (`yarn start` or `npm run start`, with install if needed)

If a `401` still reaches the error page, a `Run teams-token` button appears for manual refresh.

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
- Message reactions display in chat (`Reactions: ...`)
- Quick react hotkey in chat (`e` adds 👍 to selected message)
- Reply mode in chat (`r` replies to selected message)
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
