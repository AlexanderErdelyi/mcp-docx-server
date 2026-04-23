from io import BytesIO
from pathlib import Path

import requests
from docx import Document
from mcp.server.fastmcp import FastMCP


mcp = FastMCP("mcp-docx-server")


def _extract_docx_text(document: Document) -> str:
    return "\n".join(paragraph.text for paragraph in document.paragraphs)


@mcp.tool()
def read_docx_local(file_path: str) -> str:
    """
    Read text content from a local .docx file.
    """
    try:
        document = Document(Path(file_path))
        return _extract_docx_text(document)
    except FileNotFoundError:
        return f"Error: File not found: {file_path}"
    except Exception as exc:
        return f"Error: Could not read DOCX file '{file_path}': {exc}"


@mcp.tool()
def read_docx_ado(attachment_url: str, pat: str) -> str:
    """
    Download a .docx attachment from Azure DevOps and return text content.
    """
    if not pat:
        return "Error: Azure DevOps PAT is required."

    try:
        response = requests.get(attachment_url, auth=("", pat), timeout=30)
        response.raise_for_status()
    except requests.RequestException as exc:
        return f"Error: Could not download attachment: {exc}"

    try:
        document = Document(BytesIO(response.content))
        return _extract_docx_text(document)
    except Exception as exc:
        return f"Error: Could not parse downloaded DOCX content: {exc}"


if __name__ == "__main__":
    mcp.run(transport="stdio")
