# Outlook Browser MCP

MCP server that drives **Outlook on the web** (Office 365) with a headless Playwright browser (Edge).

Tools: `list_mails`, `read_mail`, `delete_mail` (to Deleted Items), `create_draft` (never sends).

## Setup

```bash
npm install
npm run build
npm run login      # visible browser: sign in yourself (password + MFA) once
```

The session is stored in `.browser-profile/` (git-ignored, contains your cookies).

## Use in Claude Code

```bash
claude mcp add outlook-browser -- node C:/GIT/SDLC/Practice.OutlookMcp/dist/index.js
```

Debug with `npm run inspect` (MCP Inspector). Set `OUTLOOK_HEADLESS=false` to watch the browser.

## Notes

- OWA selectors (ARIA roles) live in `src/outlook.ts`; adjust them when the UI changes.
- Indexes are positional: re-list after a delete.
- A Microsoft Graph API based server would be more robust, but needs an Entra app registration (IT approval).
