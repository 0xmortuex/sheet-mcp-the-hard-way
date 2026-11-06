# sheet-mcp — The Hard Way

Build an MCP server that lets an AI agent read and edit Excel files **without damaging them**.
You start from the zip inside an `.xlsx` and end with a server that refuses to save over a file it
would mutilate.

At the end you will have built a small version of [sheet-mcp](https://github.com/0xmortuex/sheet-mcp)
yourself, and you will understand the things most spreadsheet tutorials skip: what a formula cell
really stores, what a library silently drops when it saves, and how to design edit tools an agent
can use safely.

## Why This Exists

Most "edit a spreadsheet with Python" examples are `ws["A1"] = 5; wb.save(...)`. Run that on a real
workbook and you can lose the charts, the pivot tables, the macros, the cached results of every
formula, and never get an error. Hand that power to an AI agent unsupervised and the question
stops being "how do I write a cell" and becomes:

- how do I know what a save will destroy *before* I do it?
- how do I show a formula's value when nothing ever calculated it?
- how do I make an edit the agent can preview, and the user can trust?
- how do I open a file from the internet without being hurt by it?

## Who This Is For

- Developers who have built a "hello world" MCP server and want a real one
- Anyone who has been bitten by openpyxl silently dropping something
- People who wonder what is actually inside an `.xlsx`

**Prerequisites:** Python and a vague memory of Excel. No Excel installation is needed. Install the
dependencies with `pip install -r requirements.txt`.

## Modules

| # | Module | What You'll Build |
|---|--------|-------------------|
| 01 | [An xlsx Is a Zip](01-xlsx-is-a-zip/) | A workbook written and read by hand with `zipfile`: parts, relationships, shared strings, cached values |
| 02 | [openpyxl Basics](02-openpyxl-basics/) | The three ways to open a workbook and their costs; dates, styles, merges |
| 03 | [Formulas vs Values](03-formulas-vs-values/) | Telling "empty" from "never calculated"; shared formulas; why a save drops every cached result; why not to write a calculator |
| 04 | [What a Save Loses](04-what-a-save-loses/) | A part-name inventory, a refusal with `allow_loss`, and zip-bomb / XML-entity defences |
| 05 | [A Safe set_cells](05-safe-set-cells/) | Explicit coercion rules, dry run, guards for merged and array cells, atomic new-file writes, a before/after proof |
| 06 | [Formula Checks](06-formula-checks/) | Tokenizer-based checks: R1C1 pattern outliers, hardcoded numbers, circular references |
| 07 | [The MCP Wrapper and Testing](07-mcp-wrapper-and-testing/) | A three-tool MCP server tested over real stdio, with sandbox and error handling |

Each module contains:
- **`README.md`**: the concepts and design decisions
- **`main.py`**: a working, commented implementation that checks its own claims with `assert`
- **`break-it.md`**: challenges that break your own code, with hints

## How To Use This Course

**Option A: Build along.** Read the modules in order, run each `main.py`, then do the break-it
challenges.

**Option B: Reference.** Each `main.py` stands alone; there are no imports between modules.

**Option C: Use the finished tool.** `uvx sheet-mcp --dir /path/to/spreadsheets`.

## Every Module Is Verified

```bash
python verify.py
```

This runs each `main.py` in a fresh process and fails if any assertion fails. CI does it on Linux
and Windows on every push and weekly, so the code in this course cannot silently stop working when
openpyxl or the MCP SDK change.

## Architecture Overview

```
 AI agent (Claude Code, Claude Desktop, ...)
        │  MCP over stdio: JSON-RPC tool calls              ← Module 07
        ▼
 ┌────────────────────────────────────────────────────────┐
 │ MCP server: tools, sandbox, errors that reach the agent │  ← Module 07
 │   ├── edits:   coercion, dry run, guards, atomic write  │  ← Module 05
 │   ├── checks:  tokenizer, R1C1, cycles                  │  ← Module 06
 │   ├── safety:  part inventory, zip + XML limits         │  ← Module 04
 │   ├── values:  formula text vs cached result            │  ← Module 03
 │   └── files:   openpyxl over the zip of XML             │  ← Module 01, 02
 └────────────────────────────────────────────────────────┘
```

## License

MIT
