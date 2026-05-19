# mcp-docx-server

A local MCP server that reads Microsoft Word (`.docx`) files from local disk or Azure DevOps work item attachments and exposes their content as plain text to AI assistants like GitHub Copilot.

## Features

- `read_docx_local(file_path: str) -> str`
  - Reads a local `.docx` file with `python-docx`
  - Returns full text content (headings + paragraphs)
  - Returns readable error strings for missing/unreadable files
- `read_docx_ado(attachment_url: str) -> str`
  - Downloads a `.docx` from an Azure DevOps attachment URL
  - Authenticates automatically (MSAL device-code, Azure CLI, Git Credential Manager, or PAT fallback)
  - Parses binary content in memory (no temp files)
  - Returns full text content
- `read_docx_from_workitem(workitem_url: str) -> str`
  - Fetches all `.docx` attachments from an Azure DevOps work item URL
  - Authentication is automatic
- `read_docx_bytes(base64_content: str) -> str`
  - Parses a `.docx` from base64-encoded binary content
- `login_ado_start(org: str) -> str` / `login_ado_complete() -> str`
  - Device-code OAuth login for Azure DevOps (no PAT required)
- `store_ado_pat(pat: str) -> str`
  - Save a PAT to Windows Credential Manager as an auth fallback

## Installation

### Option A — Via `uvx` from GitHub (recommended, no cloning needed)

1. Install [`uv`](https://docs.astral.sh/uv/getting-started/installation/) once on your machine:

   **Windows (PowerShell):**
   ```powershell
   powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
   ```
   **Mac/Linux:**
   ```bash
   curl -LsSf https://astral.sh/uv/install.sh | sh
   ```

2. Add to VS Code `settings.json` — `uvx` handles everything else automatically:

   ```json
   {
     "mcp": {
       "servers": {
         "docx": {
           "command": "uvx",
           "args": ["--from", "git+https://github.com/AlexanderErdelyi/mcp-docx-server", "mcp-docx-server"]
         }
       }
     }
   }
   ```

> **First run** downloads and caches the package (~10–20 s). Subsequent starts are instant.
> To get the latest version: `uv cache clean`

---

### Option B — Local clone (for development)

```bash
git clone https://github.com/AlexanderErdelyi/mcp-docx-server
cd mcp-docx-server
pip install -r requirements.txt
python server.py
```

VS Code `settings.json`:

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

## Uninstalling

1. Remove the `"docx"` entry from VS Code `settings.json`
2. Clear the uv cache (optional): `uv cache clean`
