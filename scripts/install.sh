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

if ! command -v git >/dev/null 2>&1; then
  echo "Missing dependency: git"
  exit 1
fi

if ! command -v go >/dev/null 2>&1; then
  echo "Missing dependency: go"
  exit 1
fi

if [[ -d "${INSTALL_DIR}/.git" ]]; then
  git -C "${INSTALL_DIR}" fetch --all --tags
  git -C "${INSTALL_DIR}" checkout master
  git -C "${INSTALL_DIR}" pull --ff-only origin master
else
  rm -rf "${INSTALL_DIR}"
  git clone "${REPO_URL}" "${INSTALL_DIR}"
fi

git -C "${INSTALL_DIR}" submodule update --init --recursive

install -d "${BIN_DIR}" "${ICON_DIR}" "${DESKTOP_DIR}"
install -m 0755 "${INSTALL_DIR}/teams-cli" "${BIN_DIR}/teams-cli"
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

echo "teams-cli installed."
echo "Command: teams-cli"
echo "Install dir: ${INSTALL_DIR}"
