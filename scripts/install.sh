#!/usr/bin/env bash
set -euo pipefail

REPO_URL="${REPO_URL:-https://github.com/Gamedirection/teams-cli.git}"
INSTALL_DIR="${INSTALL_DIR:-/opt/teams-cli}"
BIN_DIR="${BIN_DIR:-/usr/local/bin}"
ICON_DIR="${ICON_DIR:-/usr/local/share/icons/hicolor/scalable/apps}"
DESKTOP_DIR="${DESKTOP_DIR:-/usr/local/share/applications}"

if [[ "${EUID}" -ne 0 ]]; then
  echo "Please run as root (use sudo)."
  exit 1
fi

REAL_USER="${SUDO_USER:-${USER}}"
REAL_HOME=$(eval echo "~${REAL_USER}")

for dep in git go node; do
  if ! command -v "$dep" >/dev/null 2>&1; then
    echo "Missing dependency: $dep"
    exit 1
  fi
done

# Install chafa (terminal image viewer) if missing
if ! command -v chafa >/dev/null 2>&1; then
  echo "Installing chafa (terminal image viewer)..."
  if command -v pacman >/dev/null 2>&1; then
    pacman -Sy --noconfirm chafa
  elif command -v apt-get >/dev/null 2>&1; then
    apt-get install -y chafa
  elif command -v dnf >/dev/null 2>&1; then
    dnf install -y chafa
  elif command -v zypper >/dev/null 2>&1; then
    zypper --non-interactive in chafa
  else
    echo "Warning: chafa not found and no supported package manager detected. Install chafa manually for terminal image viewing."
  fi
fi

# Clone or update repo
if [[ -d "${INSTALL_DIR}/.git" ]]; then
  git -C "${INSTALL_DIR}" fetch --all --tags
  git -C "${INSTALL_DIR}" checkout master
  git -C "${INSTALL_DIR}" pull --ff-only origin master
else
  rm -rf "${INSTALL_DIR}"
  git clone "${REPO_URL}" "${INSTALL_DIR}"
fi

git -C "${INSTALL_DIR}" submodule update --init --recursive

# Give ownership to real user so teams-cli can write .cache
chown -R "${REAL_USER}:${REAL_USER}" "${INSTALL_DIR}"

# Build binary as real user
echo "Building teams-cli binary..."
su - "${REAL_USER}" -c "cd '${INSTALL_DIR}' && go build -o teams-cli-bin ." || true

# Write launcher: use pre-built binary if present, fall back to go run
cat > "${INSTALL_DIR}/teams-cli" <<'LAUNCHER'
#!/usr/bin/env bash
set -euo pipefail

SCRIPT_PATH="${BASH_SOURCE[0]}"
if command -v readlink >/dev/null 2>&1; then
  SCRIPT_PATH="$(readlink -f "$SCRIPT_PATH" 2>/dev/null || echo "$SCRIPT_PATH")"
fi
SCRIPT_DIR="$(cd "$(dirname "$SCRIPT_PATH")" && pwd)"

if [[ -x "$SCRIPT_DIR/teams-cli-bin" ]]; then
  exec "$SCRIPT_DIR/teams-cli-bin" "$@"
fi

CACHE_DIR="${XDG_CACHE_HOME:-$HOME/.cache}/teams-cli/go-build"
mkdir -p "$CACHE_DIR"
cd "$SCRIPT_DIR"
exec env GOCACHE="$CACHE_DIR" go run . "$@"
LAUNCHER
chmod +x "${INSTALL_DIR}/teams-cli"

# Set up teams-token-cli
TOKEN_DIR="${INSTALL_DIR}/teams-token-cli"

if [[ -d "${TOKEN_DIR}" ]]; then
  # Fix tsconfig for modern TypeScript
  cat > "${TOKEN_DIR}/tsconfig.json" <<'TSCONFIG'
{
    "compilerOptions": {
      "module": "commonjs",
      "noImplicitAny": false,
      "skipLibCheck": true,
      "strictNullChecks": false,
      "sourceMap": true,
      "rootDir": "./src",
      "outDir": "dist",
      "baseUrl": ".",
      "ignoreDeprecations": "6.0",
      "paths": {
        "*": ["node_modules/*"]
      },
      "esModuleInterop": true
    },
    "include": [
      "src/**/*"
    ]
  }
TSCONFIG

  # Install latest TypeScript and deps as real user
  su - "${REAL_USER}" -c "cd '${TOKEN_DIR}' && npm install -g typescript >/dev/null 2>&1 || true"
  su - "${REAL_USER}" -c "cd '${TOKEN_DIR}' && yarn install --silent"
  su - "${REAL_USER}" -c "cd '${TOKEN_DIR}' && yarn add typescript@latest --dev --silent"

  # Find system electron and symlink it
  ELECTRON_BIN=""
  for candidate in electron electron41 electron39 electron37; do
    if command -v "$candidate" >/dev/null 2>&1; then
      ELECTRON_BIN=$(command -v "$candidate")
      break
    fi
  done

  if [[ -n "${ELECTRON_BIN}" ]]; then
    mkdir -p "${TOKEN_DIR}/node_modules/electron/dist"
    ln -sf "${ELECTRON_BIN}" "${TOKEN_DIR}/node_modules/electron/dist/electron"
    printf "electron" > "${TOKEN_DIR}/node_modules/electron/path.txt"
    echo "Linked system electron: ${ELECTRON_BIN}"
  else
    echo "Warning: no system electron found. teams-token-cli may need 'electron' installed."
  fi

  chown -R "${REAL_USER}:${REAL_USER}" "${TOKEN_DIR}"
fi

# Install binary symlink and desktop integration
install -d "${BIN_DIR}" "${ICON_DIR}" "${DESKTOP_DIR}"
ln -sf "${INSTALL_DIR}/teams-cli" "${BIN_DIR}/teams-cli"
install -m 0644 "${INSTALL_DIR}/img/DarkMode_Color.svg" "${ICON_DIR}/teams-cli.svg"

cat > "${DESKTOP_DIR}/teams-cli.desktop" <<'EOF'
[Desktop Entry]
Type=Application
Name=teams-cli
Comment=Terminal UI for Microsoft Teams
Exec=teams-cli
Icon=teams-cli
Terminal=true
Categories=Network;Chat;
EOF

if command -v update-desktop-database >/dev/null 2>&1; then
  update-desktop-database "${DESKTOP_DIR}" >/dev/null 2>&1 || true
fi

echo ""
echo "teams-cli installed."
echo ""
echo "First-time setup: run the following to authenticate:"
echo "  cd ${TOKEN_DIR} && yarn start"
echo ""
echo "Then run: teams-cli"
