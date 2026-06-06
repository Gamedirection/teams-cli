# TODO / Feature Backlog

## High Priority
- [ ] **Token auto-refresh on 401**: currently the "Run teams-token" button tries to launch Electron headlessly (fails). Options:
  - Detect 401, show clear "run: cd /opt/teams-cli/teams-token-cli && yarn start" message
  - Or launch Carbonyl auth in a new terminal window automatically
- [x] **Terminal image rendering**: ✅ SOLVED — uses curl with `skypetoken_asm` cookie (without Authorization: Bearer which was causing 401). Images render via chafa with `+`/`-` resize.
- [ ] **Unread count badges**: show unread message count next to chat name in tree

## Medium Priority
- [ ] **Search**: fuzzy search across chats and messages
- [ ] **Notification sound**: play a sound on new message
- [ ] **Message pagination**: load older messages on scroll-up
- [ ] **`--msg N` flag**: limit messages loaded per conversation (from upstream)
- [ ] **`--no-live` flag**: disable background polling (from upstream)
- [ ] **Graceful Ctrl+C shutdown**: currently may leave goroutines running
- [ ] **Pre-emptive token refresh**: check token expiry at startup, warn if <30 min left
- [ ] **File attachments**: detect `<attachment>` tags in messages, show download option

## Low Priority / Nice to Have
- [ ] **TTS playback**: read messages aloud
- [ ] **Inline image preview via Kitty/iTerm2 protocol**: instead of chafa symbols, use actual pixel image rendering if terminal supports it
- [ ] **Carbonyl auto-refresh**: when 401 occurs, spawn Carbonyl in a new terminal for re-auth without restarting teams-cli
- [ ] **Device code flow**: register a new Azure AD app that allows device code flow — true headless auth with no browser at all
- [ ] **Merge Carbonyl branch**: once headless token refresh is solved, merge `feat/carbonyl-auth` to master

## Packaging
- [ ] `pacman` AUR package
- [ ] Flatpak
- [ ] AppImage
- [ ] Homebrew (macOS)

## Terminal Image Display — Attempts Log
See `docs/carbonyl-notes.md` for Carbonyl auth branch notes.

All image display approaches tried and current status:

| Attempt | Method | Result |
|---------|--------|--------|
| 1 | `Authorization: Bearer <token-skype.jwt>` | 401 — wrong auth type |
| 2 | `Authentication: skypetoken=<token-skype.jwt>` | 401 — JWT not the short token |
| 3 | `api.GetSkypeToken()` → `Authentication: skypetoken=<jwt>` | 401 — GetSkypeToken returns JWT, not short token |
| 4 | `api.GetSkypeToken()` → `Cookie: skypetoken_asm=<jwt>` | 401 — still the JWT, not short token |
| 5 | Direct POST to `/api/authsvc/v1.0/authz` (wrong URL `/authsvc/...` without `/api/`) | 405 Method Not Allowed |
| 6 | Direct POST to `/api/authsvc/v1.0/authz` (correct URL) → `Cookie: skypetoken_asm=<short_token>` | **200 ✓ confirmed via curl** |
| 7 | Same as 6 but Go `http.Client` instead of curl | 401 — Go HTTP client sent both cookie AND `Authorization: Bearer` which overrode cookie auth |
| 8 | `curl` with ONLY `--cookie skypetoken_asm=<short>` + `-H Authentication: skypetoken=<short>`, NO `Authorization: Bearer` | **200 ✓ WORKING** |

**Root cause**: The `Authorization: Bearer <jwt>` header was conflicting with and overriding the `skypetoken_asm` cookie. Removing `Authorization` makes it work. Use `curl` subprocess (not Go http.Client) to match exact curl behavior.
