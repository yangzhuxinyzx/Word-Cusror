---
metadata:
  openclaw:
    emoji: "📝"
    requires:
      bins: ["node"]
    description: "Control Word-Cursor document editor via MCP — read, edit, and manage Word documents through the running desktop app."
---

# Word-Cursor Document Editor

Control the Word-Cursor AI document editor through MCP tools. Use this skill
when you need to read, edit, format, or manage Word (.docx) documents.

## Prerequisites

The Word-Cursor desktop app must be running. It exposes a local MCP bridge
on port 19527 (configurable via `WORD_CURSOR_PORT` env var).

## MCP Server Setup

Add to your MCP config:

```json
{
  "mcpServers": {
    "word-cursor": {
      "command": "node",
      "args": ["<path-to-word-cursor>/electron/mcp-server.cjs"]
    }
  }
}
```

Or if installed globally: `word-cursor-mcp`

## Available Tools

| Tool | Description |
|------|-------------|
| `word_read_document` | Read current document as indexed DSL JSON |
| `word_insert` | Insert DSL blocks at a position |
| `word_replace` | Find and replace text |
| `word_delete` | Delete text or blocks |
| `word_insert_chart` | Insert chart (bar/line/pie/doughnut/scatter/radar) |
| `word_save` | Save document to disk |
| `word_open` | Open a .docx file |
| `word_list_files` | List workspace files |

## Editing Workflow

1. **Read first**: `word_read_document` to get the document structure with `_i` block indices
2. **Edit**: Use `word_replace`, `word_insert`, or `word_delete` — reference blocks by `_i` index for precision
3. **Verify**: `word_read_document` again to confirm changes (indices shift after edits)
4. **Save**: `word_save` when done

## DSL Block Format

Blocks returned by `word_read_document` look like:

```json
{"_i": 0, "type": "heading", "level": 1, "content": "Document Title"}
{"_i": 1, "type": "paragraph", "content": "Body text here."}
{"_i": 2, "type": "paragraph", "content": [
  {"text": "Bold part", "bold": true},
  " and normal text."
]}
```

For `word_insert`, provide blocks without `_i`:

```json
[
  {"type": "heading", "level": 2, "content": "New Section"},
  {"type": "paragraph", "content": "Section content goes here."}
]
```

## Chart Insertion

Use `word_insert_chart` with labels and datasets as JSON strings:

```
labels: ["Q1", "Q2", "Q3", "Q4"]
datasets: [{"label": "Revenue", "data": [100, 200, 300, 400]}]
```

Supported types: bar, line, pie, doughnut, scatter, radar.

## Tips

- Always read before editing — block indices change after each edit
- Use `blockIndex` parameter in replace/delete for precision when text appears multiple times
- The editor shows changes with diff highlighting so the user can review
- For large edits, batch related changes and re-read between batches
