# Changelog

All notable updates made during this project iteration are documented here.

## [v1.1.0] - 2026-06-05

### Added
- **Custom chat groups**: organize any chat into named groups stored alongside Favorites.
  - Press `g` on a chat to open a picker: move to existing group, create a new group, or remove from group (returns to Recent).
  - Press `g` on a group header to rename or delete the group (chats return to Recent on delete).
  - Groups persist encrypted in `~/.config/fossteams/teams-cli-settings.enc`.
  - Tree order: Favorites → [Groups in display order] → Recent.
- **Keybinding cheatsheet overlay**: press `?` from tree or chat pane to toggle a centered help modal showing all current keybindings. Dismiss with `?` or `Esc`. Blocked in compose so `?` can still be typed in messages.
- **Message timestamps**: each message now shows send time in dark grey (`HH:MM AM/PM`) next to the author name.
- **"No messages yet." placeholder** when a conversation has no messages.

### Changed
- **Chat message layout redesign**:
  - Author name displayed above message content (header-first, matching standard chat app conventions).
  - Content lines indented 2 spaces under the author header.
  - Blank line between each message for visual separation.
  - Secondary text disabled; author, timestamp, and reactions embedded in the primary line.
- **Scroll-to-bottom on load**: uses `QueueUpdateDraw` to reliably position at the latest message after initial load.
- **Installer improvements** (`scripts/install.sh`):
  - `chown` install dir to the invoking user so `go run` can write `.cache` without `sudo`.
  - Launcher GOCACHE moved from `/opt/teams-cli/.cache` to `$HOME/.cache/teams-cli/go-build`.
  - `teams-token-cli` set up automatically: yarn install, TypeScript upgrade, system electron detection and symlink (`electron41/39/37`), `path.txt` written without trailing newline.
  - Post-install message prints first-time auth instructions.
  - `node` added to dependency check list.

### Fixed
- Running `teams-cli` as a regular user failed with permission denied on `.cache` write under `/opt`.
- `teams-token-cli` build failures on modern Node (TypeScript 4.2 vs undici-types parse errors, missing `rootDir`, deprecated `baseUrl`).
- Electron binary not found after `yarn install` when no pre-built binary is downloaded (resolved via system electron symlink).

---

## [v1.0.0] - 2026-02-19

### Added
- Favorites section at the top of Chats, grouped with existing team/chat structure.
- Favorite toggle via hotkey (`f`) in chat tree and chat contexts.
- Private Notes/self chat discovery and favorite toggling support.
- DM support improvements:
  - included DMs in the chat list,
  - sorted by most recent activity,
  - improved author/title resolution,
  - reply/send-back support.
- One-minute unread scanner with on/off toggle and manual scan trigger.
- Manual unread marking for chats.
- Persistent unread markers integrated with scanner updates.
- Built-in `Settings & Help` chat node at the bottom of chat tree.
- Reply mode for selected messages in chat view (`r` in chat pane).
- Reactions display in message UI and quick reaction hotkey (`e`).
- Mention/tagging support:
  - `@name` for current thread members,
  - `c@name` to force global contacts,
  - `Up/Down` cycling for mention suggestions in compose.
- Configurable keybindings:
  - in-app binding editor,
  - preset cycling (`default`, `vim`, `emacs`, `jk`),
  - runtime reload without app restart.
- Chat text mode controls:
  - `Word Wrap` and `Scroll` modes,
  - wrap width by characters,
  - preset widths: `20, 40, 72, 80, 100, 200, 400, 600, 800, 1000`, plus custom.
- Encrypted settings persistence for:
  - favorite chats,
  - custom chat titles,
  - wrap settings,
  - unread overrides,
  - compose/author color settings.
- Theme customization in settings:
  - compose input highlight color cycling,
  - username color cycling.
- teams-token integration:
  - added as optional submodule at `/teams-token`,
  - automatic 401 auth-refresh attempts,
  - manual `Run teams-token` button on 401 error screen.
- Installation/packaging improvements:
  - installer script targets system paths,
  - CLI command `teams-cli` available after install,
  - desktop launcher/icon install support.

### Changed
- Compose highlight color moved to a darker palette; default now `slate`.
- Wrap mode styling aligned with scroll mode for consistent message/author presentation.
- Settings UI spacing improved so options are easier to scan.
- README expanded with install methods, hotkeys, settings, keybinds, and roadmap details.

### Fixed
- Crash when toggling favorites with `f`.
- Crash paths around unread toggle/scan hotkeys (`m`, `Shift+M`, `r`) across panes.
- Settings/help interaction crashes.
- 401 refresh flow messaging and fallback behavior when teams-token runners are missing.
- Mentions not cycling/selecting correctly in compose.
- Mention lookup fallback issues across DMs/group chats/private notes.
- Word wrap width handling bugs and effective width calculation issues.
- Compose/readability color contrast issues.

### Docs and Repo Hygiene
- README updated repeatedly to reflect new features and controls.
- `.gitignore` updated to include encrypted/key artifacts (`*.enc`, `*.key`).
- `.cache` cleanup guidance applied to reduce accidental tracking.

---

## Prior Baseline (before this iteration)
- Existing Teams/channel listing and basic chat rendering from earlier project history.
- teams-api and UI foundation already in place.
