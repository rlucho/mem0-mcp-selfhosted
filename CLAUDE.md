# CLAUDE.md

This file provides guidance to Claude Code when working with code in this repository.

## Writing Style (hard rule)

**Never use an em dash (U+2014).** It is the strongest single tell that text was
machine-generated, and the user rejects it outright. The same applies to an en
dash (U+2013) used as spaced sentence punctuation. This covers everything, not
just repo files: chat replies, emails and messages drafted for the user, code
comments, commit messages, PR descriptions, and documentation.

Use instead: commas or parentheses for an aside, a colon before an explanation,
a semicolon or two sentences to join independent clauses, and `to` for a range
in prose ("5 to 10 minutes"). Hyphens in compound words, identifiers and CLI
flags stay fine (`well-formed`, `--json-response`), as do en dashes inside
numeric ranges in tables.

Check any file you wrote or edited (returns nothing when clean):

```bash
grep -rn "$(printf '\u2014')" <file>
```

A `PostToolUse` hook (`.claude/hooks/check-em-dash.sh`) enforces this on every
Write and Edit and reports the offending lines back. For the rest of what makes
prose read as machine-written (LLM vocabulary, sentence shapes, formatting
tells), use the `humanize` skill in `.claude/skills/humanize/`.

## MCP Servers

- **mem0**: Persistent memory across sessions. At the start of each session, `search_memories` for relevant context before asking the user to re-explain anything. Use `add_memory` whenever you discover project architecture, coding conventions, debugging insights, key decisions, or user preferences. Use `update_memory` when prior context changes. Save information like: "This project uses PostgreSQL with Prisma", "Tests run with pytest -v", "Auth uses JWT validated in middleware". When in doubt, save it: future sessions benefit from over-remembering.

## Build & Test Commands

```bash
pip install -e ".[dev]"              # Install with dev dependencies
python3 -m pytest tests/unit/ -v     # Unit tests (mocked, no infra needed)
python3 -m pytest tests/contract/ -v # Contract tests (validates mem0ai internals)
python3 -m pytest tests/integration/ -v  # Integration tests (requires live Qdrant + Neo4j + Ollama)
python3 -m pytest tests/ -v          # All tests
python3 -m pytest tests/ -m "not integration" -v  # Skip integration
python3 -m pytest tests/unit/test_auth.py::TestIsOatToken -v  # Single test class
python3 -m pytest tests/unit/test_auth.py::TestIsOatToken::test_oat_token_detected -v  # Single test
```

## Architecture

Self-hosted MCP server using `mem0ai` as a library. 11 tools (9 memory + 2 graph), `MCPServer` orchestrator (`mcp` SDK v2).

**Module roles:**
- `server.py`: `MCPServer` orchestrator, registers all tools + `memory_assistant` prompt; transport options (host/port/`json_response`) are passed to `run()`, not the constructor
- `config.py`: Env vars → mem0ai `MemoryConfig` dict, handles all 5 graph LLM provider configs
- `auth.py`: 3-tier token fallback: `MEM0_ANTHROPIC_TOKEN` → `~/.claude/.credentials.json` → `ANTHROPIC_API_KEY`
- `llm_anthropic.py`: Custom Anthropic provider registered with mem0ai's `LlmFactory`; handles OAT headers, structured outputs (JSON schema via `output_config`), and tool-call parsing
- `llm_router.py`: `SplitModelGraphLLM` routes by tool name: extraction tools → Gemini, contradiction tools → Claude
- `helpers.py`: `_mem0_call()` error wrapper, `call_with_graph()` threading lock for per-call graph toggle, `safe_bulk_delete()` iterates+deletes individually (never calls `memory.delete_all()`), `patch_graph_sanitizer()` monkey-patches mem0ai's relationship sanitizer for Neo4j compliance
- `graph_tools.py`: Direct Neo4j Cypher queries with lazy driver init
- `__init__.py`: Suppresses mem0ai telemetry before any imports

**Critical implementation details:**
- `memory.delete()` does NOT clean Neo4j nodes (mem0ai bug #3245), so `safe_bulk_delete()` explicitly calls `memory.graph.delete_all(filters)` after
- `memory.enable_graph` is mutable instance state, so `call_with_graph()` holds a `threading.Lock` for the full duration of each Memory call (2-20s)
- Contract tests (`tests/contract/`) validate mem0ai internal API assumptions; if these fail after a mem0ai upgrade, the code needs updating
- `Memory.update()` uses `data=` parameter, not `text=`
- Structured output support requires claude-opus-4/sonnet-4/haiku-4 models; older models fall back to JSON extraction
- mem0ai's `sanitize_relationship_for_cypher()` has gaps (no hyphen handling, no leading-digit check), so `patch_graph_sanitizer()` wraps it at startup to ensure all relationship types match `^[a-zA-Z_][a-zA-Z0-9_]*$`
