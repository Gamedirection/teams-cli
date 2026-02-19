# Changelog

All notable updates made during this project iteration are documented here.

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

