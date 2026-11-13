#!/usr/bin/env bash
# Installs the systemd user service and creates the env file with a fresh token.
set -euo pipefail
cd "$(dirname "$0")/.."
ENV_FILE="$HOME/.config/outlook-mcp.env"
UNIT_DIR="$HOME/.config/systemd/user"

if [[ ! -f "$ENV_FILE" ]]; then
  umask 077
  sed "s/^MCP_AUTH_TOKEN=.*/MCP_AUTH_TOKEN=$(openssl rand -hex 32)/" deploy/outlook-mcp.env.example > "$ENV_FILE"
  echo "Created $ENV_FILE with a new token."
fi

mkdir -p "$UNIT_DIR"
sed "s#%h/Workspace/mail-mcp#$PWD#" deploy/outlook-mcp.service > "$UNIT_DIR/outlook-mcp.service"
systemctl --user daemon-reload
systemctl --user enable --now outlook-mcp
# Keep running when you're logged out.
loginctl enable-linger "$USER" || true
systemctl --user --no-pager status outlook-mcp | head -5
TOKEN=$(grep '^MCP_AUTH_TOKEN=' "$ENV_FILE" | cut -d= -f2-)
PORT=$(grep '^MCP_PORT=' "$ENV_FILE" | cut -d= -f2- || echo 8787)
echo
echo "Setup page (on this machine): http://127.0.0.1:${PORT:-8787}/setup/$TOKEN/"
echo "From your laptop over SSH:    ssh -L ${PORT:-8787}:127.0.0.1:${PORT:-8787} $(hostname)  then open the URL above"
