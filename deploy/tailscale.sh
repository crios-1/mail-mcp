#!/usr/bin/env bash
# Exposes the local MCP server over HTTPS with Tailscale.
#   ./deploy/tailscale.sh serve   -> tailnet only (Claude Code / Claude Desktop on your devices)
#   ./deploy/tailscale.sh funnel  -> public internet (required for claude.ai web/mobile connectors)
#   ./deploy/tailscale.sh off
set -euo pipefail
MODE="${1:-serve}"
PORT="${MCP_PORT:-8787}"
case "$MODE" in
  serve)  sudo tailscale serve --bg --https=443 "http://127.0.0.1:$PORT" ;;
  funnel) sudo tailscale funnel --bg --https=443 "http://127.0.0.1:$PORT" ;;
  off)    sudo tailscale serve reset; exit 0 ;;
  *) echo "usage: $0 serve|funnel|off" >&2; exit 1 ;;
esac
HOST=$(tailscale status --self --json | python3 -c 'import json,sys;print(json.load(sys.stdin)["Self"]["DNSName"].rstrip("."))')
TOKEN=$(grep -s '^MCP_AUTH_TOKEN=' "$HOME/.config/outlook-mcp.env" | cut -d= -f2- || true)
echo
echo "Header auth URL : https://$HOST/mcp   (Authorization: Bearer <token>)"
echo "Path auth URL   : https://$HOST/mcp/${TOKEN:-<token>}"
echo "Setup page      : https://$HOST/setup/${TOKEN:-<token>}/"
