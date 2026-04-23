# mcp-docx-server

A local MCP server that reads Microsoft Word (`.docx`) files from local disk or Azure DevOps work item attachments and exposes their content as plain text to AI assistants like GitHub Copilot.

## Features

- `read_docx_local(file_path: str) -> str`
  - Reads a local `.docx` file with `python-docx`
  - Returns full text content (headings + paragraphs)
  - Returns readable error strings for missing/unreadable files
- `read_docx_ado(attachment_url: str, pat: str) -> str`
  - Downloads a `.docx` from an Azure DevOps attachment URL
  - Uses PAT-based Basic Auth
  - Parses binary content in memory (no temp files)
  - Returns full text content

## Installation

```bash
pip install -r requirements.txt
```

## Run (stdio mode)

```bash
python server.py
```

## VS Code MCP configuration (`settings.json`)

```json
{
  "mcp": {
    "servers": {
      "docx": {
        "command": "python",
        "args": ["/absolute/path/to/mcp-docx-server/server.py"]
      }
    }
  }
}
```
