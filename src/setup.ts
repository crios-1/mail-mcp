/**
 * Web setup page at /setup/<token>/ : status, connection info, and a remote
 * login. The remote login shows the headless browser as screenshots and forwards
 * your clicks and keys to it, so you can sign in to Office 365 (password + MFA)
 * without a screen on the server. It uses the server's own browser profile, so
 * the session is saved where the MCP tools look for it.
 */
import type { IncomingMessage, ServerResponse } from "node:http";
import type { Page } from "playwright";
import { close, getContext, gotoOutlook, PROFILE_DIR } from "./browser.js";
import { testConnection } from "./tools.js";

let loginPage: Page | undefined;
let lastTest: { ok: boolean; at: string; detail: string } | undefined;

const SPECIAL_KEYS = new Set([
  "Enter", "Backspace", "Tab", "Escape", "Delete", "ArrowUp", "ArrowDown", "ArrowLeft", "ArrowRight", "Home", "End",
]);

function json(res: ServerResponse, status: number, body: unknown): void {
  res.writeHead(status, { "Content-Type": "application/json", "Cache-Control": "no-store" }).end(JSON.stringify(body));
}

async function readBody(req: IncomingMessage): Promise<any> {
  const chunks: Buffer[] = [];
  for await (const c of req) chunks.push(c as Buffer);
  const raw = Buffer.concat(chunks).toString("utf8");
  return raw ? JSON.parse(raw) : {};
}

async function loginState() {
  if (!loginPage || loginPage.isClosed()) return { active: false };
  const p = loginPage;
  const url = p.url();
  const title = await p.title().catch(() => "");
  const text = await p.evaluate(() => document.body?.innerText ?? "").catch(() => "");
  const inbox = await p.getByRole("listbox").count().catch(() => 0);
  return {
    active: true,
    url,
    title,
    text: text.replace(/\s+/g, " ").slice(0, 200),
    onOutlook: /outlook\.(office|office365|live)\.com\/mail|outlook\.cloud\.microsoft\/mail/.test(url),
    inboxVisible: inbox > 0,
  };
}

async function startLogin(): Promise<void> {
  if (loginPage && !loginPage.isClosed()) return;
  const ctx = await getContext();
  loginPage = await ctx.newPage();
  // Microsoft may ask for a passkey / Windows Hello first. That opens a native
  // browser dialog the screenshots can't show. An empty virtual authenticator makes
  // the request fail at once, so Microsoft offers the other methods (app, password).
  const cdp = await ctx.newCDPSession(loginPage);
  await cdp.send("WebAuthn.enable", { enableUI: false });
  await cdp.send("WebAuthn.addVirtualAuthenticator", {
    options: { protocol: "ctap2", transport: "internal", hasResidentKey: true, hasUserVerification: true, isUserVerified: false },
  });
  await gotoOutlook(loginPage);
}

async function stopLogin(): Promise<void> {
  await loginPage?.close().catch(() => undefined);
  loginPage = undefined;
}

async function handleInput(ev: any): Promise<void> {
  const p = loginPage;
  if (!p || p.isClosed()) throw new Error("No login in progress");
  switch (ev.type) {
    case "click":
      await p.mouse.click(Number(ev.x), Number(ev.y));
      break;
    case "scroll":
      await p.mouse.move(Number(ev.x), Number(ev.y));
      await p.mouse.wheel(0, Number(ev.dy));
      break;
    case "key":
      if (!SPECIAL_KEYS.has(ev.key)) throw new Error(`Key not allowed: ${ev.key}`);
      await p.keyboard.press(ev.key);
      break;
    case "text":
      await p.keyboard.insertText(String(ev.text));
      break;
    case "reload":
      await gotoOutlook(p);
      break;
    default:
      throw new Error(`Unknown input: ${ev.type}`);
  }
}

/** parts = path segments after /setup/<token> */
export async function handleSetup(req: IncomingMessage, res: ServerResponse, parts: string[], token: string): Promise<void> {
  const route = `${req.method} /${parts.join("/")}`;
  try {
    switch (route) {
      case "GET /":
        res.writeHead(200, { "Content-Type": "text/html; charset=utf-8", "Cache-Control": "no-store" }).end(page(req, token));
        return;
      case "GET /status":
        return json(res, 200, { login: await loginState(), lastTest, profile: PROFILE_DIR });
      case "GET /screen.jpg": {
        if (!loginPage || loginPage.isClosed()) return json(res, 404, { error: "No login in progress" });
        const img = await loginPage.screenshot({ type: "jpeg", quality: 60, timeout: 10_000 });
        res.writeHead(200, { "Content-Type": "image/jpeg", "Cache-Control": "no-store" }).end(img);
        return;
      }
      case "POST /login/start":
        await startLogin();
        return json(res, 200, await loginState());
      case "POST /login/stop":
        await stopLogin();
        return json(res, 200, { active: false });
      case "POST /input":
        await handleInput(await readBody(req));
        return json(res, 200, { ok: true });
      case "POST /test": {
        try {
          const n = await testConnection();
          lastTest = { ok: true, at: new Date().toISOString(), detail: `Inbox reachable (${n} mail listed)` };
        } catch (e) {
          lastTest = { ok: false, at: new Date().toISOString(), detail: (e as Error).message.split("\n")[0] };
        }
        return json(res, 200, lastTest);
      }
      case "POST /restart-browser":
        await stopLogin();
        await close();
        return json(res, 200, { ok: true });
      default:
        return json(res, 404, { error: "Not found" });
    }
  } catch (e) {
    return json(res, 500, { error: (e as Error).message.split("\n")[0] });
  }
}

function page(req: IncomingMessage, token: string): string {
  const host = (req.headers["x-forwarded-host"] as string) ?? req.headers.host ?? "localhost";
  const proto = (req.headers["x-forwarded-proto"] as string) ?? (/^(localhost|127\.)/.test(host) ? "http" : "https");
  const base = `${proto}://${host}`;
  const esc = (s: string) => s.replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" })[c]!);
  const claudeCode = `claude mcp add --transport http outlook ${base}/mcp --header "Authorization: Bearer ${token}"`;
  return `<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<meta name="robots" content="noindex"><title>Outlook MCP Setup</title>
<style>
:root{--bg:#f6f7f9;--card:#fff;--fg:#1c2024;--mute:#60646c;--line:#e0e1e6;--acc:#0f6cbd;--ok:#1a7f37;--bad:#c62828}
@media (prefers-color-scheme:dark){:root{--bg:#111315;--card:#1b1d20;--fg:#e8e9ea;--mute:#9a9ea5;--line:#2c2f33;--acc:#4fa3ff;--ok:#4ac26b;--bad:#ff6b6b}}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--fg);font:15px/1.5 system-ui,sans-serif}
main{max-width:1000px;margin:0 auto;padding:24px 16px}h1{font-size:22px;margin:0 0 4px}h2{font-size:16px;margin:0 0 12px}
.sub{color:var(--mute);margin:0 0 20px}.card{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:16px;margin-bottom:16px}
button{font:inherit;padding:7px 14px;border-radius:7px;border:1px solid var(--line);background:var(--card);color:var(--fg);cursor:pointer}
button.primary{background:var(--acc);border-color:var(--acc);color:#fff}button:disabled{opacity:.5;cursor:default}
.row{display:flex;gap:8px;flex-wrap:wrap;align-items:center}
code,pre{font:13px/1.45 ui-monospace,monospace}pre{background:var(--bg);border:1px solid var(--line);border-radius:7px;padding:10px;overflow-x:auto;white-space:pre-wrap;word-break:break-all;margin:6px 0 12px}
.ok{color:var(--ok)}.bad{color:var(--bad)}.mute{color:var(--mute)}
#screen{display:block;width:100%;border:1px solid var(--line);border-radius:7px;cursor:crosshair;outline:none;background:#000;min-height:120px}
#screen:focus{border-color:var(--acc)}input[type=text]{flex:1;min-width:0;font:inherit;padding:7px 10px;border-radius:7px;border:1px solid var(--line);background:var(--card);color:var(--fg)}
</style></head><body><main>
<h1>Outlook MCP Setup</h1><p class="sub">Sign in to Office 365, test the connection, and connect Claude.</p>

<section class="card"><h2>1. Status</h2>
<div id="status" class="mute">Loading…</div>
<div class="row" style="margin-top:12px"><button class="primary" id="test">Test connection</button><button id="restart">Restart browser</button></div>
</section>

<section class="card"><h2>2. Sign in to Office 365</h2>
<p class="mute" style="margin-top:0">Opens Outlook in the server's browser. Click the image to focus it, then click and type as usual (paste works). Choose <b>Yes</b> on “Stay signed in?”. Finish when your inbox is visible.</p>
<div class="row"><button class="primary" id="start">Start sign-in</button><button id="stop" disabled>Finish</button><button id="reload" disabled>Reload Outlook</button></div>
<div id="live" hidden style="margin-top:12px">
<img id="screen" tabindex="0" alt="Remote browser">
<div class="row" style="margin-top:8px"><input type="text" id="txt" placeholder="Or type text here and press Send (e.g. a long password)" autocomplete="off"><button id="send">Send</button><button id="enter">Enter ⏎</button></div>
</div>
</section>

<section class="card"><h2>3. Connect Claude</h2>
<b>Claude Code</b><pre>${esc(claudeCode)}</pre>
<b>claude.ai custom connector</b> <span class="mute">(Settings → Connectors; needs Tailscale funnel)</span><pre>${esc(`${base}/mcp/${token}`)}</pre>
<span class="mute">Keep these secret: the token gives access to your mailbox.</span>
</section>
</main>
<script>
const $=id=>document.getElementById(id);
const api=(path,body)=>fetch(path,{method:body===undefined?'GET':'POST',headers:{'Content-Type':'application/json'},body:body===undefined?undefined:JSON.stringify(body)}).then(r=>r.json());
let live=false,timer;
function render(s){
  const t=s.lastTest;
  const test=t?('<span class="'+(t.ok?'ok':'bad')+'">'+(t.ok?'✔ ':'✘ ')+esc(t.detail)+'</span> <span class="mute">('+new Date(t.at).toLocaleTimeString()+')</span>'):'<span class="mute">Not tested yet.</span>';
  const L=s.login;
  const login=L.active?('Sign-in open'+(L.inboxVisible?' <span class="ok">– inbox visible, click Finish</span>':'')+'<br><span class="mute">'+esc(L.title||'(no title)')+' · '+esc(L.url)+'</span>'):'No sign-in open';
  $('status').innerHTML='Connection: '+test+'<br>'+login+'<br><span class="mute">Profile: '+esc(s.profile)+'</span>';
  setLive(s.login.active);
}
function esc(s){return String(s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'})[c])}
function setLive(on){
  if(on===live)return;live=on;$('live').hidden=!on;$('start').disabled=on;$('stop').disabled=!on;$('reload').disabled=!on;
  clearTimeout(timer);if(on)frame();
}
function frame(){
  if(!live)return;const img=new Image();
  img.onload=()=>{$('screen').src=img.src;timer=setTimeout(frame,600)};
  img.onerror=()=>{timer=setTimeout(frame,1500)};
  img.src='screen.jpg?t='+Date.now();
}
const refresh=()=>api('status').then(render).catch(()=>{});
const send=ev=>api('input',ev).then(r=>{if(r.error)alert(r.error)});
function coords(e){const el=$('screen'),r=el.getBoundingClientRect();return{x:Math.round((e.clientX-r.left)*el.naturalWidth/r.width),y:Math.round((e.clientY-r.top)*el.naturalHeight/r.height)}}
$('screen').addEventListener('click',e=>{$('screen').focus();send({type:'click',...coords(e)})});
$('screen').addEventListener('wheel',e=>{e.preventDefault();send({type:'scroll',...coords(e),dy:e.deltaY})},{passive:false});
const SPECIAL=['Enter','Backspace','Tab','Escape','Delete','ArrowUp','ArrowDown','ArrowLeft','ArrowRight','Home','End'];
$('screen').addEventListener('keydown',e=>{
  if(e.ctrlKey||e.metaKey||e.altKey)return;
  if(SPECIAL.includes(e.key)){e.preventDefault();send({type:'key',key:e.key})}
  else if(e.key.length===1){e.preventDefault();send({type:'text',text:e.key})}
});
$('screen').addEventListener('paste',e=>{e.preventDefault();send({type:'text',text:e.clipboardData.getData('text')})});
$('send').onclick=()=>{const v=$('txt').value;if(v){send({type:'text',text:v});$('txt').value=''}};
$('txt').addEventListener('keydown',e=>{if(e.key==='Enter'){e.preventDefault();$('send').click()}});
$('enter').onclick=()=>send({type:'key',key:'Enter'});
$('start').onclick=async()=>{$('start').disabled=true;const r=await api('login/start',{});if(r.error){alert(r.error);$('start').disabled=false}refresh()};
$('stop').onclick=async()=>{await api('login/stop',{});refresh();$('test').click()};
$('reload').onclick=()=>send({type:'reload'});
$('test').onclick=async()=>{$('test').disabled=true;$('test').textContent='Testing…';await api('test',{});$('test').disabled=false;$('test').textContent='Test connection';refresh()};
$('restart').onclick=async()=>{if(confirm('Close the browser? It restarts on the next call.')){await api('restart-browser',{});refresh()}};
refresh();setInterval(refresh,4000);
</script></body></html>`;
}
