# Carbonyl Auth Branch — Session Notes (2026-06-05/06)

## What Was Done

Replaced the Electron-based `teams-token-cli` with a terminal-native auth flow using
[Carbonyl](https://github.com/fathyb/carbonyl) + Puppeteer.

### Key changes

- `teams-token-cli/src/main.ts` — full rewrite from Electron API to Puppeteer + Carbonyl
- `teams-token-cli/package.json` — removed `electron`, added `carbonyl`, `puppeteer-core`, `node-pty`, upgraded `jsonwebtoken` to v9
- `teams-token-cli` converted from git submodule to tracked source
- `app.go` — `refreshAuthFromTeamsToken` now checks `teams-token-cli` dir name first (falls back to `teams-token`)
- `teams-cli` launcher — added `cd "$SCRIPT_DIR"` before binary exec so relative paths resolve

### How the Carbonyl auth flow works

1. Node.js spawns Carbonyl via `node-pty` (gives it a real PTY — keyboard, resize, raw mode all work)
2. Carbonyl renders Microsoft login in the terminal. User logs in, completes DUO — no desktop needed.
3. Node.js connects to Carbonyl via Puppeteer CDP (`puppeteer.connect({ browserURL: 'http://localhost:9977' })`)
4. `waitForRedirect()` listens on `page.on('response')` for the 302 `Location` header from Microsoft containing `https://teams.microsoft.com/go#token=...`
5. On capture: navigate to `about:blank` immediately (stops Teams web app loading), extract token hash, process
6. Sequential: Teams token → Skype token → ChatSvcAgg token
7. Tokens saved to `~/.config/fossteams/token-{type}.jwt`, process exits

### What worked
- Carbonyl renders Microsoft login correctly in terminal ✓
- DUO MFA (push notification) works ✓
- OAuth tokens captured via CDP response interception ✓
- node-pty solves keyboard input freezing ✓
- `timeout: 0` + `setDefaultTimeout(0)` prevent Puppeteer 30s kick ✓

### Remaining limitation
- Token auto-refresh (the "Run teams-token" button on 401) cannot run Carbonyl
  headlessly — Carbonyl requires interactive terminal. The button triggers `yarn start`
  which launches Carbonyl but with no TTY attached, so auth hangs.
- Workaround: user must manually run `cd /opt/teams-cli/teams-token-cli && node ./dist/main.js`

---

## Suggestions for Moving Forward

### Priority 1 — Fix headless refresh
The core UX gap: when a 401 occurs mid-session, teams-cli currently can't auto-refresh
with the Carbonyl approach.

**Options:**
1. **Detect 401 and prompt user** — instead of trying to auto-run teams-token, show
   a message: "Tokens expired. Run: `cd /opt/teams-cli/teams-token-cli && node ./dist/main.js`
   then restart." Simple and honest.
2. **Launch in a new terminal window** — detect the user's terminal emulator (kitty, alacritty,
   gnome-terminal, etc.) and spawn Carbonyl in a new window when 401 occurs. More complex
   but seamless.
3. **Shorter token lifetime handling** — pre-emptively refresh tokens before they expire
   (check expiry at startup and warn if <30 min left).

### Priority 2 — Ship the Carbonyl branch
Once headless refresh is resolved (even option 1), merge to master and remove Electron
from the install script entirely. The install script currently still sets up Electron
as a fallback.

### Priority 3 — Upstream features worth porting
From the fossteams upstream (`feat: harden runtime shutdown`, `feat: expand CLI configuration`):
- `--msg <N>` flag to limit messages loaded
- `--no-live` flag to disable background polling
- Graceful Ctrl+C shutdown (currently may leave goroutines running)

### Priority 4 — Token refresh without browser
If Microsoft ever allows it for this client ID: implement device code flow or PKCE
with a new Azure AD app registration. This would make auth fully scriptable with no
browser at all. The upstream notes on Terminal Auth Attempts document why the current
client ID blocks this.
