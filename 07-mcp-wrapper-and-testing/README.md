# Module 07: The MCP Wrapper, and Testing It For Real

## The Concept

The previous modules are a library. An MCP server adds a thin layer so an agent can call it.
`main.py` has two roles, as in Module 01 of the sibling course:

```bash
python main.py --serve --dir /some/folder   # the server (a client starts this)
python main.py                              # the test: starts itself as a server and talks to it
```

### Tools are the interface; their descriptions are the manual

The agent never sees your code, only each tool's name, docstring and argument types. The SDK
builds the JSON schema from the Python signature, so type hints and defaults *are* the contract:
`dry_run: bool = True` tells the agent the safe thing is the default.

### Errors: only ToolError survives

In `mcp` 2.x, if a tool raises an ordinary exception the client sees `Error executing tool X` and
nothing else. The agent cannot tell "that sheet does not exist" from a crash. The fix is the
`tool()` decorator: it catches your own `SheetError` and re-raises `ToolError(str(e))`, so the
message ("Sheets: S") reaches the agent. Two details:

- `functools.wraps(fn)` is required. Without it the wrapper's signature is `(*args, **kwargs)` and
  the generated schema is wrong (try removing it; you get a validation error on every call).
- Catch **expected** errors only. A genuine bug should stay loud in your logs, not be dressed up as
  a user error.

### The sandbox

The server takes `--dir`. Every path goes through `resolve()`, which joins it to the root,
resolves `..` and symlinks, and refuses anything outside. Path checks belong in one function that
every tool uses; five tools each doing their own check is how one of them ends up without it.

### Never print to stdout

In stdio mode stdout *is* the protocol. A stray `print` corrupts the JSON-RPC stream. Log to
stderr.

## Testing: Do It Over The Real Pipe

Calling the tool functions directly proves your logic. It does not prove the server starts, lists
its tools, builds valid schemas, or that errors survive transport. `demo()` launches the server
as a subprocess with `stdio_client`, then:

1. lists the tools and checks the names,
2. reads a never-calculated workbook and checks the reply says so,
3. runs a dry run (no file created) and a real edit (new file; original unchanged),
4. attacks: a workbook with a chart part is refused, `../etc/passwd` is refused, a missing sheet
   returns a message that lists the real sheets.

Set `PYTHONIOENCODING=utf-8` for the child process: on Windows the default console codepage can
crash a server that prints non-ASCII text.

## Run It

```bash
python main.py
```

## From Here To sheet-mcp

The finished [sheet-mcp](https://github.com/0xmortuex/sheet-mcp) is this plus: a streaming sheet
reader (200k-row sheets), `describe_columns`, `find_cells`, `append_rows` with table support,
`rename_sheet` that rewrites references through the tokenizer, CSV conversion with explicit
encodings, and a test suite that proves round-trip identity cell by cell.
