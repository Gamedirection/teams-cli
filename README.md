# teams-cli

A Command Line Interface (or TUI) to interact with Microsoft Teams
that uses the [teams-api](https://github.com/fossteams/teams-api)
Go package.

## Status

This project is actively evolving, and already supports daily messaging workflows
for Teams channels and chats (including DMs) from a terminal UI.

## Requirements

- [Golang](https://golang.org/)

## Usage

Follow the instructions on how to obtain a token with [teams-token](https://github.com/fossteams/teams-token),
then simply run the following to start the app. Binary releases will appear on this repository as soon as
we have a product with more features.

```bash
go run ./
```

If everything goes well, you should see something like this:
![Teams CLI example](./docs/screenshots/2021-04-13.png)
<img width="1708" height="881" alt="image" src="https://github.com/user-attachments/assets/899b4aa5-f00e-4d6b-85b8-7dd9f3b07080" />


## What works

- Logging in into Teams using the token generated via `teams-token`
- Getting the list of Teams + Channels
- Reading channels
- Reading chats/DMs (most recent chats are loaded first)
- Sending messages in channels and chats/DMs
- Tab/Shift+Tab to cycle focus between panes
- Favorites for chats:
  - Press `f` on a chat to add/remove it from `Chats > Favorites`
  - Your personal `Private Notes` chat is detected and included in Favorites
- Press `u` in the chat tree to refresh chat titles with better author/name resolution

## What doesn't work

- Some Teams desktop features (calls/meetings/media-rich features) are still not implemented

## Keybindings

- `Tab`: focus next pane
- `Shift+Tab`: focus previous pane
- `i`: focus compose input for the selected conversation
- `Enter` (in compose): send message
- `Esc` (in compose): return focus to tree
- `f`: toggle favorite for selected chat
- `u`: refresh chat names/titles

## You might also be interested in

- [fossteams-frontend](https://github.com/fossteams/fossteams-frontend): a Vue based frontend for Microsoft Teams
