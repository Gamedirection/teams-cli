# Changelog

## [v1.2.0] - 2026-06-06 (ongoing)

### Added
- **Emoji reactions**: `e` key opens a picker (👍 ❤️ 😆 😮 😢 😠). Unknown reaction types fall back to 👍.
- **Emoji display in chat**: `<emoji alt="😊">` and `<img class="emoji" alt="😊">` HTML tags now render as actual emoji characters.
- **Clickable links**: `o` key on any message opens a link/image picker. Links open in browser via `xdg-open`. Images offer browser or terminal view.
- **Terminal image viewer**: images rendered with `chafa` in a scrollable popup (Esc to close). Install auto-detected via package manager.
- **Image sending**: type `<img>~/path/to/file.png</img>` in compose to embed a base64 image. Supports PNG, JPG, GIF, WebP, SVG.
- **Image placeholder**: non-emoji `<img>` tags display `🖼️ filename` in chat.
- **Image save folder**: configurable in Settings & Help → `Image Save Folder`. Default: `~/.local/share/teams-cli/images/`. Symlink at `~/Pictures/teams-cli-screenshots`.
- **HTML formatting in received messages**: `<b>`, `<i>`, `<u>`, `<s>`, `<code>`, `<blockquote>`, `<br>` all render with tview markup (bold, italic, underline, strikethrough, green code, gray quote prefix).
- **Links underlined**: `<a href>` renders with tview underline markup.
- **Markdown toggle**: Settings & Help → `Markdown` (ON by default). Toggles markdown rendering in received messages.
- **Markdown on send**: compose markdown converts to HTML before sending — `**bold**`, `*italic*`, `__underline__`, `~~strike~~`, `` `code` ``, ` ```block``` `, `> quote`.
- **`G` in tree pane**: jump to and open the most recently active chat.
- **`G` in chat pane**: scroll to the bottom (most recent message).
- **`chafa` auto-install**: install script detects pacman/apt/dnf/zypper and installs chafa if missing.
- **Image resize**: `+` / `-` in terminal image view to grow/shrink; re-renders chafa at new size. Esc to close.

### Changed
- `textMessage()` rewritten as HTML-to-tview converter (preserves formatting, emoji, images).
- `wrapTextLines()` uses rune-based visible-length measurement so tview markup tags don't break wrapping.
- Message rendering uses `QueueUpdateDraw` for reliable scroll-to-bottom on load.
- Image CDN auth: uses `curl` subprocess with `skypetoken_asm` cookie only — `Authorization: Bearer` header was conflicting and causing 401. Root cause took 8 attempts to isolate (see `docs/TODO.md`).
- Images cached to disk by URL hash — same image won't re-download.

### Fixed
- `wrapTextLines` panic: `strings.LastIndex` returns byte index, used as rune index → slice-out-of-bounds. Fixed with rune-based space search.
- `s.app.Draw()` inside tview callbacks caused deadlock/crash. Removed.
- Private ANSI sequences (`[?25l`, `[?25h`) from chafa output caused `tview.ANSIWriter` panic. Stripped before rendering.
- Launcher missing `cd "$SCRIPT_DIR"` before binary exec — relative paths (teams-token-cli) resolved from wrong directory.
- `app.go` auth refresh looked for `teams-token` dir; renamed to `teams-token-cli`.

---

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
- Chat message layout redesign: author above content, indented lines, blank line between messages.
- Scroll-to-bottom on load uses `QueueUpdateDraw`.
- Installer: `chown` to invoking user, GOCACHE to `$HOME/.cache`, teams-token-cli auto-setup, binary build during install.

### Fixed
- Running `teams-cli` as regular user failed (permission denied on `.cache` under `/opt`).
- teams-token-cli TypeScript build failures on Node v26 (undici-types, rootDir, deprecated baseUrl, SlowBuffer).
- Electron binary missing — resolved via system electron symlink.

---

## [v1.0.0] - 2026-02-19

### Added
- Favorites, unread scanner, DM support, reply mode, reactions, mentions, keybinding editor, wrap settings, encrypted settings, theme colors, teams-token integration.

---

## Prior Baseline
- Teams/channel listing and basic chat rendering from earlier project history.
