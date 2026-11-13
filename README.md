# Outlook Browser MCP

MCP server that drives **Outlook on the web** (Office 365) with a headless Playwright browser.
No Entra app registration needed: it uses your own signed-in browser session.

Tools: `list_mails`, `search_mails` (Outlook search syntax, all folders), `read_mail` (by index from a folder or a search), `delete_mail` (to Deleted Items), `create_draft` (never sends), `queue_status`.
The selectors handle both English and Dutch Outlook.

Transports:

- **stdio** (default): a local client starts the process.
- **HTTP** (`--http`): Streamable HTTP on `127.0.0.1:8787/mcp`, made for exposure through **Tailscale** so you can use it from anywhere.

## 1. Setup (Linux server, e.g. `tower`)

```bash
npm install
npm run browser:install   # Playwright's Chromium (default browser on Linux)
npm run build
npm run login             # visible browser: sign in yourself (password + MFA) once
```

`npm run login` needs a display. **Headless server / SSH only?** Skip it and sign in through the web setup page instead (step 3).
Don't run `npm run login` while the service is running: both would open the same browser profile.
The session is stored in `.browser-profile/<channel>/` (git-ignored, contains your cookies).
When the session expires, the tools fail with "Not logged in"; run `npm run login` again.

## 2. Run as a service

```bash
npm run service:install
```

This creates `~/.config/outlook-mcp.env` (mode 600) with a random `MCP_AUTH_TOKEN`, then installs and starts the systemd **user** unit `outlook-mcp`.
It also enables lingering, so the service keeps running when you're logged out.

```bash
systemctl --user status outlook-mcp
journalctl --user -u outlook-mcp -f
curl -s localhost:8787/healthz
```

## 3. Expose it with Tailscale

Pick one mode:

| Mode | Command | Reachable from | Use for |
|---|---|---|---|
| serve | `./deploy/tailscale.sh serve` | your tailnet only | Claude Code, Claude Desktop on your devices |
| funnel | `./deploy/tailscale.sh funnel` | the public internet | **claude.ai web / mobile** custom connectors |

claude.ai connectors are called from Anthropic's servers, not from your device, so they can't reach a tailnet-only address. They need **funnel**.
With funnel, the token is the only thing protecting your mailbox: keep it secret, and rotate it by editing the env file and restarting the service.

Either mode serves `https://tower.<your-tailnet>.ts.net`. The script prints the URLs.
Funnel must be allowed for the node in the Tailscale admin console (`nodeAttrs` → `funnel`). The first `tailscale funnel` run prints a link to enable it.

### Setup page

Open `https://<host>/setup/<token>/` from any device on your tailnet. Before Tailscale is set up, use `ssh -L 8787:127.0.0.1:8787 tower` and open `http://127.0.0.1:8787/setup/<token>/`.

- **Sign in to Office 365**: shows the server's headless browser and forwards your clicks and keys, so password + MFA work without a screen on the server.
- **Test connection**: lists one inbox mail.
- **Connect Claude**: ready-to-copy commands and URLs.

## 4. Connect Claude

The token goes in **either** a header or the URL path:

- `https://<host>/mcp` with `Authorization: Bearer <token>`
- `https://<host>/mcp/<token>` for clients that can't send headers

**Claude Code** (any machine on your tailnet):

```bash
claude mcp add --transport http outlook https://<host>/mcp --header "Authorization: Bearer <token>"
```

**claude.ai (web, desktop, mobile)**: Settings → Connectors → *Add custom connector*, URL `https://<host>/mcp/<token>`. This needs funnel.

**Local stdio** (same machine, no network):

```bash
claude mcp add outlook-browser -- node /path/to/mail-mcp/dist/index.js
```

Debug with `npm run inspect` (MCP Inspector). Set `OUTLOOK_HEADLESS=false` to watch the browser.

## Configuration

| Env var | Default | |
|---|---|---|
| `MCP_TRANSPORT` | `stdio` | `http` is the same as `--http` |
| `MCP_AUTH_TOKEN` | — | required in HTTP mode, min 24 chars |
| `MCP_HOST` / `MCP_PORT` | `127.0.0.1` / `8787` | keep localhost; Tailscale proxies to it |
| `OUTLOOK_BROWSER_CHANNEL` | `chromium` on Linux, else `chrome` | `chrome`, `msedge`, `chromium` |
| `OUTLOOK_PROFILE_DIR` | `.browser-profile/<channel>` | |
| `OUTLOOK_CONCURRENCY` | `3` | parallel browser tabs |
| `OUTLOOK_MAX_QUEUE` / `OUTLOOK_JOB_TIMEOUT_MS` | `50` / `120000` | |
| `OUTLOOK_HEADLESS` | `true` | |

## Notes

- OWA selectors (ARIA roles) live in `src/outlook.ts`; adjust them when the UI changes.
- Indexes are positional: re-list after a delete.
- Every client shares one browser and one job queue.
- A Microsoft Graph API based server would be more robust, but needs an Entra app registration (IT approval).
