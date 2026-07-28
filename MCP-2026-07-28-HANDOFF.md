# Hand-off: MCP 2026-07-28 spec + `mcp` Python SDK 2.0.0 migration

**Written:** 2026-07-28 · **Author:** Claude Code (remote web session, branch `claude/mcp-update-analysis-o6isvo`)
**For:** a local Claude Code session that has network access to the Proxmox homelab and both self-hosted MCP servers.

---

## 0. How to use this document

This session could only reach **this repository**. It could not reach the running servers on the
Proxmox host. Everything about the live deployments below was reconstructed from mem0 and is
marked as such — **verify before acting on it.**

Infrastructure specifics (container IDs, addresses, gateway config, service names, credentials
locations) are deliberately **not** written into this file, because this repo may be public.
They are stored in mem0 instead. Start by retrieving them:

```
search_memories("MCP 2026-07-28 migration handoff fleet inventory")
search_memories("cadlab-mcp server specs Caddy gateway")
search_memories("mem0 server deployment container")
```

Then work through §4 in order.

---

## 1. What changed upstream

Two separate things landed on 2026-07-28.

### 1.1 The protocol: MCP spec revision `2026-07-28`

Fifth revision, largest change since launch. What matters for us:

- **Stateless core.** The `initialize` / `notifications/initialized` handshake is gone, as is the
  `Mcp-Session-Id` header. Every request now carries its protocol version and client capabilities
  in `_meta`. `server/discover` replaces `initialize` for version negotiation. Servers are now
  plain HTTP endpoints — deployable on serverless/edge, scalable behind a round-robin load
  balancer with no session stickiness.
- **Cross-call state** is now expected to live in **server-minted handles passed as ordinary tool
  arguments**, not in protocol sessions.
- **Extensions are first-class.** Two graduated: `io.modelcontextprotocol/tasks` (long-running work:
  return a handle, client polls `tasks/get`) and MCP Apps (server-rendered UI in a sandboxed iframe).
- **Standard request headers** (SEP-2243): `Mcp-Method` and `Mcp-Name` are required on Streamable
  HTTP POSTs, so gateways can route/authorize **per tool without parsing the JSON-RPC body**.
  Custom headers can also be sourced from tool parameters via `x-mcp-header`.
- **Cacheable list results** (SEP-2549): `tools/list`, `prompts/list`, `resources/list` etc. now
  carry `ttlMs` and `cacheScope`. Servers SHOULD also return tools in a **deterministic order** to
  improve client-side caching and LLM prompt-cache hit rates.
- **MRTR** (Multi Round-Trip Requests) replaces server-initiated `roots/list` / `sampling/createMessage`
  / `elicitation/create`. A tool returns `InputRequiredResult`; the client retries with `inputResponses`.
- **All results carry a required `resultType`** (`"complete"` or `"input_required"`).
- **Deprecated** (12-month window): Roots, Sampling, Logging, HTTP+SSE transport, and OAuth Dynamic
  Client Registration (superseded by Client ID Metadata Documents).
- **Auth hardening:** `iss` validation per RFC 9207; client credentials must be keyed by issuer.

Anthropic additionally announced **MCP tunnels** (research preview): Claude connects to an MCP
server inside a private network with no public endpoint, no inbound firewall rules, and no IP
allowlisting on the origin. See §4.3 — this is the most architecturally significant item for us.

### 1.2 The SDK: `mcp` 2.0.0 on PyPI

Released the same day. A major rework supporting `2026-07-28` **and every earlier revision** from a
single server. `1.x` is now maintenance-only (security patches).

The upstream maintainers' own guidance: **pin `mcp>=1.28,<2` until you have migrated.**

---

## 2. Verified evidence (reproduced in this session — not inferred)

```console
$ python3 -m venv v2 && ./v2/bin/pip install "mcp[cli]"
$ ./v2/bin/python -c "import importlib.metadata as md; print(md.version('mcp'))"
2.0.0

$ ./v2/bin/python -c "from mcp.server.fastmcp import FastMCP"
ModuleNotFoundError: No module named 'mcp.server.fastmcp'

$ ./v2/bin/python -c "from mcp.server.mcpserver import MCPServer"   # OK
```

Introspected v2 signatures:

```
MCPServer.__init__(self, name=None, title=None, description=None, instructions=None,
                   website_url=None, icons=None, version='', auth_server_provider=None,
                   token_verifier=None, *, tools=None, resources=None, extensions=None,
                   debug=False, log_level='INFO', ..., cache_hints=None, subscriptions=None,
                   middleware=None)

MCPServer.run(self, transport: Literal['stdio','sse','streamable-http'] = 'stdio', **kwargs) -> None

MCPServer.tool(self, name=None, title=None, description=None, annotations=None,
               icons=None, meta=None, structured_output=None)

MCPServer.prompt(self, name=None, title=None, description=None, icons=None)
```

Note two traps visible in that constructor:
1. **`host`/`port` are gone from the constructor** — they moved to `run(**kwargs)`.
2. **The second positional parameter is now `title`, not `instructions`.** Any code passing
   `instructions` positionally will silently populate `title`. Always pass it by keyword.

---

## 3. Impact assessment

### 3.1 This repo (`mem0-mcp-selfhosted`) — **P0, currently broken for new installs**

`pyproject.toml` declares `"mcp[cli]>=1.23.0"` with **no upper bound**, and there is **no
`uv.lock`** in the repo. The README installs via:

```
uvx --from git+https://github.com/<owner>/mem0-mcp-selfhosted.git mem0-mcp-selfhosted
```

`uvx` re-resolves dependencies at install time, so **any fresh install or cache refresh from today
onward pulls `mcp` 2.0.0 and the server fails at import.** Existing warm uvx caches keep working
until invalidated.

Breakage inventory in this repo:

| Location | Issue |
|---|---|
| `src/mem0_mcp_selfhosted/server.py:18` | `from mcp.server.fastmcp import FastMCP` → `ModuleNotFoundError` |
| `server.py:38, 149, 183, 440` | `FastMCP` used as a type annotation |
| `server.py:156-170` | `FastMCP("mem0", host=…, port=…, instructions=…)` — host/port must move to `run()` |
| `server.py:501-508` | `server.run(transport=…)` needs host/port plumbed in |
| `tests/unit/test_mcp_protocol.py:98` | `tool.inputSchema` → `tool.input_schema` (fields are snake_case in v2) |
| `tests/unit/test_mcp_protocol.py:116,127,167` | `content_blocks, _ = await srv.call_tool(...)` — v2 returns a `CallToolResult` **directly**, not a tuple |
| `tests/unit/test_server.py:62,377,391` | `srv._tool_manager._tools` private internals — verify these still exist under `mcp.server.mcpserver` |
| `tests/unit/test_server.py:382` | `srv._prompt_manager._prompts` — same |
| `CLAUDE.md:24,27`, `README.md:291` | Docs say "FastMCP orchestrator" |

What does **not** change: `@mcp.tool()` and `@mcp.prompt()` decorators, and the `Annotated[..., Field(...)]`
parameter schemas. All 11 tool bodies are untouched.

### 3.2 The two live servers — *from mem0, unverified*

Both are FastMCP servers running from venvs behind a Caddy header-auth gateway. Because they run
from pinned venvs rather than `uvx`, **they will not break until someone runs `pip install -U`** —
but they are one careless upgrade away from the same failure.

- **mem0 server** — streamable-HTTP behind a header-auth gateway, exposed publicly, reached from
  laptops via a Python stdio↔HTTP bridge (`mem0_proxy.py`). A 2026-06-21 connectivity note records
  `serverInfo mem0 v1.27.2`; that is the **SDK** version FastMCP reports, not this package's version
  (`0.3.2`) — so that host is on the `1.27` line. **Confirm before assuming.**
- **cadlab server** — 9 CAD/3D-printing tools (`cadquery_build`, `openscad_render`, `freecad_script`,
  `slice_bambu`, `measure`, `render_png`, `list_outputs`, `get_file`, `bambu_status`) over CadQuery,
  build123d, OpenSCAD-Manifold, FreeCAD and Bambu Studio CLI. LAN-only, reached via a twin bridge
  (`cad_proxy.py`).

---

## 4. Action plan

### 4.1 P0 — Stop the bleeding (do this first, everywhere)

Pin below the major version on all three Python environments:

**This repo** — `pyproject.toml`:
```diff
-    "mcp[cli]>=1.23.0",
+    "mcp[cli]>=1.23.0,<2",
```

**Both servers** — check what is actually installed, then pin:
```bash
<venv>/bin/python -c "import importlib.metadata as m; print(m.version('mcp'))"
<venv>/bin/pip install "mcp[cli]>=1.23.0,<2"
```
Add the same constraint to whatever requirements file / install script each service uses, so a
future rebuild does not silently jump to 2.x. Restart the service and confirm it still answers.

> Consider committing a `uv.lock` to this repo as well. Without one, `uvx --from git+…` is
> permanently exposed to upstream major bumps — this is the second time dependency drift has bitten
> this project.

### 4.2 P1 — Migrate to `MCPServer`

Do this on a branch, in this repo first (it has tests), then port the same shape to the two servers.

**`server.py` import and annotations:**
```diff
-from mcp.server.fastmcp import FastMCP
+from mcp.server.mcpserver import MCPServer
```
Replace every `FastMCP` annotation with `MCPServer` (lines 38, 149, 183, 440).

**Constructor — drop host/port, keep `instructions` as a keyword:**
```diff
-    host = env("MEM0_HOST", "0.0.0.0")
-    port = int(env("MEM0_PORT", "8081"))
-
-    mcp = FastMCP(
-        "mem0",
-        host=host,
-        port=port,
-        instructions=(...),
-    )
+    mcp = MCPServer(
+        "mem0",
+        instructions=(...),   # MUST stay keyword — positional #2 is now `title`
+    )
```

**Move host/port to `run()` in `run_server()`:**
```diff
+    host = env("MEM0_HOST", "0.0.0.0")
+    port = int(env("MEM0_PORT", "8081"))
     transport = env("MEM0_TRANSPORT", "stdio").lower()

     if transport == "sse":
-        server.run(transport="sse")
+        server.run(transport="sse", host=host, port=port)       # NOTE: SSE is deprecated upstream
     elif transport == "streamable-http":
-        server.run(transport="streamable-http")
+        server.run(transport="streamable-http", host=host, port=port)
     else:
         server.run(transport="stdio")
```

**Fix the tests** per the table in §3.1 — `input_schema`, the `call_tool` return type, and the
private-manager access.

**Full v1 → v2 mapping** (only rows relevant to this codebase):

| v1 | v2 |
|---|---|
| `mcp.server.fastmcp.FastMCP` | `mcp.server.mcpserver.MCPServer` |
| `mcp.server.fastmcp.*` | `mcp.server.mcpserver.*` |
| `FastMCP(name, instructions=…, host=…, port=…)` | `MCPServer(name, instructions=…)`; host/port → `run()` |
| `mcp.get_context()` | inject `ctx: Context` as a handler parameter |
| `tool.inputSchema` | `tool.input_schema` (wire JSON unchanged — Pydantic aliases) |
| `await srv.call_tool(...)` → `(blocks, raw)` | → `CallToolResult` |
| `McpError` | `MCPError(code, message, data=None)` |
| `MCP_*` env vars / `.env` auto-load | no longer read (we use `MEM0_*` + explicit `load_dotenv()`, so unaffected) |

**Behavioural change to be aware of:** in v2, synchronous `def` handlers run on a worker thread
pool. All 11 of our tools are sync. That means genuinely concurrent tool execution — which makes
the `threading.Lock` in `helpers.call_with_graph()` (held for the full 2–20 s of each Memory call,
because `memory.enable_graph` is mutable instance state) the hard serialization point for the whole
server. Migration does not break this, but it will make the contention more visible. Revisiting
that design is a separate piece of work; see §6.

**Dependency risk — test this before cutting over.** v2 replaces `httpx`/`httpx-sse` with
**`httpx2`** and requires `opentelemetry-api`. `mem0ai`, `anthropic`, `neo4j` and `qdrant-client`
all pull `httpx`. Verify the whole set co-installs in a scratch venv *before* touching either live
server:
```bash
python3 -m venv /tmp/v2test && /tmp/v2test/bin/pip install "mcp[cli]>=2" "mem0ai[graph,llms]>=1.0.3" \
  "anthropic>=0.77.0" "neo4j>=5.23.1"
/tmp/v2test/bin/pip check
```

### 4.3 P2 — Opportunities worth taking

Ordered by value to this setup.

**a) MCP tunnels → delete the whole bridge layer.**
Nearly every awkward part of the current architecture exists because Claude clients could not reach
a private-network server: both stdio↔HTTP proxy scripts, the Caddy header-auth gates, the public
exposure of the mem0 endpoint, and the still-unbuilt Cloudflare tunnel for off-LAN cadlab access.
Tunnels are purpose-built for exactly this. If the research preview holds up, it collapses all of
it and lets the public endpoint be retired.
*Caveat:* this likely does **not** solve the corporate-laptop TLS problem. That is client-side
SSL re-signing by the corporate proxy; a tunnel still terminates TLS somewhere. It removes the
`mcp-remote` / `NODE_EXTRA_CA_CERTS` vs `SSL_CERT_FILE` mess, but the corporate CA bundle question
probably survives.

**b) Tasks extension → fixes the timeout class of bugs. Highest value on cadlab.**
mem0's slow path is 2–20 s (graph-enabled `add_memory`); commit `b4129ff` ("eager background Memory
init to avoid MCP client timeout") exists because of it, and Claude Code recently fixed a bug where
per-server `request_timeout_ms` was ignored and long calls died at the 60 s default.
cadlab is worse — `slice_bambu`, `freecad_script`, FEM solves and large `openscad_render` jobs are
minute-scale blocking calls. Tasks converts them to handle-plus-poll, which is the correct shape.

**c) Per-tool authorization at the gateway.**
With `Mcp-Method` / `Mcp-Name` now required on Streamable HTTP POSTs, Caddy can authorize per tool
without parsing the body. Today the cadlab gateway key is all-or-nothing: the same credential that
grants `measure` and `list_outputs` also grants `freecad_script` and `cadquery_build`, which
**execute arbitrary Python on the container**. Splitting those behind a second key is a genuine
security improvement and is cheap.

**d) Cacheable, deterministic tool lists.**
Our tool set is static (11 here, 9 on cadlab). Declaring a long `ttlMs` with `cacheScope` and
emitting tools in deterministic order cuts re-listing round-trips and improves prompt-cache hits at
every session start — non-trivial given how verbose the `Annotated[..., Field(description=…)]`
schemas are.

**e) Structured tool output.**
Every tool here returns `json.dumps(...)` as a `str` via `_mem0_call()`. v2 loosens `outputSchema`
and `structuredContent` to arbitrary JSON Schema 2020-12, and `@mcp.tool()` takes a
`structured_output` flag. Returning typed content instead of a JSON string means the model stops
re-parsing strings, and costs fewer tokens.

**f) MCP Apps on cadlab.**
`render_png` currently returns a static PNG. Apps allows a server-rendered sandboxed UI — a real
inline STL/3MF viewer or a slicer-result panel with layer preview.

**g) OpenTelemetry.** v2 ships tracing middleware that costs nothing until an exporter is
configured. Useful for characterising where the 2–20 s graph calls actually go.

---

## 5. Things that are already correct — do not "fix" them

- **cadlab's `list_outputs` → `get_file` pattern is exactly what the stateless spec prescribes**
  (server-minted handles passed as ordinary tool arguments). It is accidentally spec-aligned.
- **Neither server uses Roots, Sampling, or Logging**, so those deprecations cost nothing. Confirm
  cadlab does not use MCP `logging/setLevel` before ticking this off.
- Statelessness / horizontal scaling is a **mem0** benefit (N replicas work around the graph lock).
  It does **not** apply to cadlab — single container, local file outputs, one printer. Do not
  over-engineer that one.

---

## 6. Open questions for the local session

1. What `mcp` version is actually installed on each server? (mem0 is *believed* to be 1.27.2 from a
   `serverInfo` string — verify.)
2. Does `mcp.server.mcpserver` still expose `_tool_manager` / `_prompt_manager`, or do the unit
   tests need rewriting against a public API?
3. Does `mem0ai` + `anthropic` + `neo4j` + `qdrant-client` co-install cleanly with `httpx2`? This
   gates the entire P1 migration.
4. Is `MEM0_TRANSPORT=sse` still used anywhere? The transport is now formally deprecated; steer
   remaining users to `streamable-http`.
5. Is the per-call graph toggle (`memory.enable_graph` + global lock) worth redesigning now that v2
   makes sync handlers genuinely concurrent? A per-call Memory instance or a graph/no-graph instance
   pair would remove the lock entirely.

---

## 7. Verification checklist

- [ ] `mcp` pinned `<2` in this repo, and on both servers
- [ ] Both services restarted and answering after the pin
- [ ] `uv.lock` committed to this repo (optional but recommended)
- [ ] Scratch-venv dependency check passes with `mcp>=2` (§4.2)
- [ ] `python3 -m pytest tests/ -m "not integration" -v` green after migration
- [ ] Contract tests (`tests/contract/`) green — they validate `mem0ai` internals, unrelated to this
      change, but a regression here means something else moved
- [ ] Fresh `uvx --from git+…` install starts successfully
- [ ] Both servers complete an MCP handshake from a real client after migration
- [ ] Docs updated: `CLAUDE.md:24,27` and `README.md:291` say "FastMCP orchestrator"

---

## 8. Sources

- MCP 2026-07-28 changelog — https://modelcontextprotocol.io/specification/2026-07-28/changelog
- Anthropic announcement — https://claude.com/blog/bringing-mcp-2026-07-28-to-claude
- Python SDK migration guide — https://py.sdk.modelcontextprotocol.io/migration/
- What's new in SDK v2 — https://py.sdk.modelcontextprotocol.io/whats-new/
- `mcp` on PyPI — https://pypi.org/project/mcp/
