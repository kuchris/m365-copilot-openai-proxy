# Microsoft 365 Copilot OpenAI Proxy

Use Microsoft 365 Copilot through OpenAI-compatible clients, local scripts, and coding tools.

This project runs a local FastAPI proxy over M365 Copilot, exposing OpenAI and Anthropic style HTTP endpoints. It supports direct WebSocket connections and an optional browser UI fallback.

No Azure app registration. No admin consent. Sign in with your normal M365 Copilot browser session.

## Why Use This

- Use M365 Copilot from OpenAI-compatible clients
- Works with your existing signed-in Copilot web session
- Runs locally on `127.0.0.1` by default
- Auto-captures and refreshes the short-lived browser token
- Supports persistent Copilot sessions across turns
- Supports OpenAI Chat Completions, OpenAI Responses, and Anthropic Messages style requests

## Quick Start

```powershell
uv sync
uv run copilot-openai-proxy serve --backend websocket
```

The server starts at:

```text
http://127.0.0.1:8000
```

On first run, the proxy opens a dedicated Edge window. Sign in to M365 Copilot there once. The first API request sends a short `Connection initialization` chat in a separate tab to capture the authenticated connection parameters. This initialization appears in Copilot history. Your actual API prompts and replies travel directly between Python and the Copilot WebSocket; they are not typed into the browser.

Credentials stay in memory. They are refreshed on the next request after 45 minutes, or once after an explicit authentication rejection. Keep the dedicated Edge profile available for refresh, and complete any visible sign-in or verification prompt. The first request and refreshes take longer than ordinary requests.

To opt into a single automatic click on a recognized Cloudflare verification checkbox:

```powershell
uv run copilot-openai-proxy serve --backend websocket --auto-verify
```

The proxy waits for normal automatic checks first, then tries at most one visible, named checkbox in a Cloudflare frame belonging to its own tab. It waits for completion and resumes the pending initialization. Image puzzles, login dialogs, unrecognized widgets, and unsuccessful checks still require attention. This does not guarantee that every verification can be completed automatically. Quota errors such as `OutOfCredits` are returned without a refresh loop.

The dedicated Edge profile is stored at:

```text
%USERPROFILE%\.m365-copilot-openai-proxy\edge-profile
```

The existing default remains `substrate` for compatibility. Select the backend explicitly:

| Backend | How messages are sent | Credential handling |
|---|---|---|
| `websocket` | Python connects directly; a new socket per request; incremental streaming | Browser-assisted initialization and automatic refresh; supports opaque personal-account tokens |
| `substrate` | Original direct WebSocket client | JWT capture to `.env`; refresh based on the token's `exp` |
| `browser` | Types each prompt into the website and reads the reply | Uses the browser login; buffered SSE |

Use `--no-launch-edge` when the dedicated debug Edge is already running. Normal requests are serialized. Stateless requests use distinct conversations; `X-M365-Session-Id` retains a conversation until the proxy restarts. Requests can appear in Copilot history in every backend.

### Original JWT Backend

```powershell
uv run copilot-openai-proxy serve --backend substrate
```

This preserves the original company/school-account workflow. It has not been revalidated with a company/school account in the current personal-account investigation. If startup is waiting for a token, click the Copilot message box and type one character without sending. This backend writes `M365_ACCESS_TOKEN` to `.env` and expects a readable JWT, not an encrypted personal-account token.

## Test It

### Browser Bridge (Experimental)

If the direct Substrate connection does not work with your account, try sending
through the signed-in Copilot website:

```powershell
uv run copilot-openai-proxy serve --backend browser
```

Sign in to Copilot in the dedicated Edge window. API requests then use a separate
bridge tab in the same profile. Keep Edge open; the bridge types the request into
Copilot and reads the completed reply. It does not export access tokens, and
token capture/refresh is disabled in this mode. Complete any sign-in or
verification prompts yourself in Edge.

The base URL, model, and API examples below remain the same. Use
`--no-launch-edge` if the dedicated debug Edge is already running, and
`--browser-timeout-seconds 180` for longer replies.

- Requests run one at a time. Ordinary requests start a new Copilot chat.
- Persistent sessions reuse a conversation URL until the proxy restarts.
- Requests and replies appear in your Copilot web chat history.
- Replies are plain text extracted from the page; rich formatting may be lost.
- `stream: true` uses SSE but delivers the answer after generation completes.
- Tool calls and image inputs are not supported. Page layout changes can break
  the bridge. `/healthz` reports browser connectivity, not a successful chat test.
- The bridge refuses to overwrite an unsent draft in its tab. Avoid using that
  tab manually while a request is running.

### Send A Request

```powershell
$body = @{
  model = "m365-copilot"
  messages = @(
    @{ role = "user"; content = "Say hello in one short sentence." }
  )
} | ConvertTo-Json -Depth 10

$r = Invoke-RestMethod `
  -Method Post `
  -Uri "http://127.0.0.1:8000/v1/chat/completions" `
  -ContentType "application/json" `
  -Body $body

$r.choices[0].message.content
```

## Connect A Client

Use these settings for any OpenAI-compatible client:

| Setting | Value |
|---|---|
| Base URL | `http://127.0.0.1:8000/v1` |
| API Key | `dummy` |
| Model | `m365-copilot` |
| Persistent model | `m365-copilot:persist` |

### OpenCode

```powershell
$env:OPENAI_BASE_URL = "http://127.0.0.1:8000"
$env:OPENAI_API_KEY = "dummy"
opencode
```

Select **OpenAI API** as the provider, then use:

```text
m365-copilot
```

For persistent Copilot-side conversation memory:

```text
m365-copilot:persist
```

### Continue

Add this to `~/.continue/config.json`:

```json
{
  "models": [
    {
      "title": "M365 Copilot",
      "provider": "openai",
      "model": "m365-copilot:persist",
      "apiBase": "http://127.0.0.1:8000/v1",
      "apiKey": "dummy"
    }
  ]
}
```

### Claude Code

```powershell
$env:ANTHROPIC_BASE_URL = "http://127.0.0.1:8000"
$env:ANTHROPIC_API_KEY = "dummy"
claude
```

Claude Code note: this proxy does not implement tool use. It can answer general prompts, but agentic features such as file reading, bash, and code editing still require the real Anthropic API.

## Persistent Sessions

By default, requests are stateless from the Copilot side.

To reuse the same Copilot conversation across turns, send a stable header:

```http
X-M365-Session-Id: my-work-session
```

Or use the model suffix:

```text
m365-copilot:persist
```

Header mode is better when your client supports custom headers, because each workspace or coding-agent session can choose its own id. If your client only lets you change the model name, use `m365-copilot:persist`.

If a client uses `m365-copilot:persist` without sending a `user` field, all requests share one default persistent session until the proxy restarts.

## Token Refresh

For `websocket`, the proxy captures connection parameters and browser-generated verification data together. It does not decode encrypted access tokens. Refresh is automatic on demand after `M365_CREDENTIAL_REFRESH_SECONDS` (default 2700 seconds), or once after an explicit authentication rejection before any response text. Ordinary errors and partial responses are not automatically resubmitted. The refresh interval is a conservative cache age, not a claim about Microsoft's token lifetime.

The remaining instructions in this section apply to `substrate`. Its browser JWTs typically expire in about one hour; the proxy reads `exp` and refreshes from the dedicated signed-in Edge window.

Auto-refresh is on by default:

```powershell
uv run copilot-openai-proxy serve
```

Useful controls:

```powershell
uv run copilot-openai-proxy serve --refresh-before-seconds 300
uv run copilot-openai-proxy serve --no-auto-refresh
uv run copilot-openai-proxy serve --no-capture-on-start
uv run copilot-openai-proxy serve --no-launch-edge
```

You can also press `r` in the server console to refresh the token manually.

### Manual Fallback

```powershell
uv run copilot-openai-proxy set-token
```

Then paste a fresh Substrate WebSocket URL:

1. Open the signed-in M365 Copilot Edge window.
2. Open DevTools (`F12`) -> **Network** tab.
3. Filter by `substrate`.
4. Click the WebSocket entry.
5. Go to **Headers** -> right-click the **Request URL** -> **Copy link address**.
6. Paste it into the terminal.

The command extracts `access_token` automatically and writes it to `.env`.

## Token Health

```powershell
Invoke-RestMethod http://127.0.0.1:8000/healthz
Invoke-RestMethod http://127.0.0.1:8000/v1/token/status
```

For `websocket`, these endpoints report credential readiness, time until the next refresh, refresh count, and Edge connectivity. Readiness indicates cached credentials, not a successful live chat. Before the first request, `/healthz` reports `initializing` if Edge is reachable. For `browser`, health reports Edge connectivity and no exported token is required.

Example for `substrate`:

```json
{
  "status": "ok",
  "token": {
    "valid": true,
    "expires_at": "2026-05-14T02:50:53+00:00",
    "seconds_remaining": 4200
  }
}
```

## API Endpoints

| Endpoint | Description |
|---|---|
| `GET /healthz` | Service health plus token status |
| `GET /v1/token/status` | Token validity, expiry time, and seconds remaining |
| `GET /v1/models` | OpenAI-compatible model list |
| `POST /v1/chat/completions` | OpenAI Chat Completions, streaming supported |
| `POST /v1/responses` | OpenAI Responses API, streaming supported |
| `POST /v1/messages` | Anthropic Messages API style endpoint |

## More Examples

### Streaming

```powershell
$body = @{
  model = "m365-copilot"
  stream = $true
  messages = @(@{ role = "user"; content = "hi" })
} | ConvertTo-Json -Depth 10

Invoke-RestMethod `
  -Method Post `
  -Uri "http://127.0.0.1:8000/v1/chat/completions" `
  -ContentType "application/json" `
  -Body $body
```

### Persistent Session

```powershell
$body = @{
  model = "m365-copilot"
  messages = @(
    @{ role = "user"; content = "Remember this code word: sakura. Reply only OK." }
  )
} | ConvertTo-Json -Depth 10

Invoke-RestMethod `
  -Method Post `
  -Uri "http://127.0.0.1:8000/v1/chat/completions" `
  -Headers @{ "X-M365-Session-Id" = "test1" } `
  -ContentType "application/json" `
  -Body $body
```

### Anthropic-Style Messages

```powershell
$body = @{
  model = "m365-copilot"
  system = "Be concise."
  messages = @(@{ role = "user"; content = "hi" })
} | ConvertTo-Json -Depth 10

$r = Invoke-RestMethod `
  -Method Post `
  -Uri "http://127.0.0.1:8000/v1/messages" `
  -ContentType "application/json" `
  -Body $body

$r.content[0].text
```

## Security Notes

- The proxy listens on `127.0.0.1` by default.
- The `substrate` token is stored locally in `.env`. The `websocket` backend keeps its captured tokens and verification data in memory; `browser` does not export them.
- `.env`, `.venv/`, and Python cache files are ignored by Git.
- The proxy does not send your token to any external service besides Microsoft 365 Copilot's own `substrate.office.com` endpoint.
- Anyone who can read your `.env` can use the token until it expires. Treat it like a secret.

## Environment Variables

The `websocket` and `browser` backends do not require a token in `.env`. Environment variables can select the backend and customize refresh timing.

| Variable | Default | Description |
|---|---|---|
| `M365_ACCESS_TOKEN` | optional at startup | Browser WebSocket token. If missing, startup capture can fill `.env`. |
| `M365_TIME_ZONE` | `Asia/Tokyo` | Optional. Time zone sent to Copilot. Usually no need to set this if `Asia/Tokyo` is correct. |
| `M365_MODEL_ALIAS` | `m365-copilot` | Optional. Model name returned by `/v1/models`. Usually no need to change this. |
| `M365_BACKEND` | `substrate` | `websocket` for browser-assisted direct connections, `substrate` for the original JWT client, or `browser` for UI automation. |
| `M365_CDP_PORT` | `9222` | Browser debugging port when creating the app directly. The CLI uses `--cdp-port`. |
| `M365_BROWSER_TIMEOUT_SECONDS` | `120` | Timeout for browser initialization and each direct WebSocket attempt, after leaving the queue. |
| `M365_CREDENTIAL_REFRESH_SECONDS` | `2700` | Maximum cached credential age in `websocket` mode; refreshed on the next request. |
| `M365_BROWSER_AUTO_VERIFY` | `false` | Opt into one recognized verification-checkbox attempt in the owned browser tab; also enabled by `--auto-verify`. |

## Development Checks

```powershell
uv run --locked --extra dev pytest -q
```

An optional DOM fixture test uses the running debug Edge (without sending a Copilot message):

```powershell
$env:M365_TEST_CDP = "1"
uv run --locked --extra dev pytest -q
```

Unit tests do not establish live Copilot availability. Account verification, service changes, and credential renewal still require live testing.

## Limitations

- This is an unofficial local proxy over the browser-facing M365 Copilot API.
- Token refresh depends on a signed-in Edge profile.
- Tool calls are not supported.
- Token usage numbers are placeholders.
- System prompts and prior conversation history are translated into plain text context.

## License

Apache License 2.0. See [LICENSE](LICENSE).

## Token Automation Details

See [TOKEN_REFRESH.md](TOKEN_REFRESH.md) for the deeper Edge CDP refresh notes and alternatives.

## Support

If this project saves you time, please consider giving it a GitHub star. It helps other people find the repo.

[![Star History Chart](https://api.star-history.com/svg?repos=kuchris/m365-copilot-openai-proxy&type=Date)](https://www.star-history.com/#kuchris/m365-copilot-openai-proxy&Date)
