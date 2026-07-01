# Publishing Guide

Recommended public repository name: `mcp-libreoffice-calc`.

The name is explicit, searchable, and avoids personal/internal wording. It clearly says that this MCP server controls LibreOffice Calc.

## Local Repository

From the project directory:

```bash
cd /home/maxime/.local/share/mcp-libreoffice-calc
git status --short --branch
```

## Code Configuration

Use this block in `~/.code/config.toml`:

```toml
[mcp_servers.libreoffice_calc]
command = "/home/maxime/.local/bin/uv"
args = ["run", "python", "/home/maxime/.local/share/mcp-libreoffice-calc/src/main.py"]
cwd = "/home/maxime/.local/share/mcp-libreoffice-calc"

[mcp_servers.libreoffice_calc.env]
PYTHONPATH = "/home/maxime/.local/share/mcp-libreoffice-calc:/home/maxime/.local/share/mcp-libreoffice-calc/src:/usr/lib/python3.14/site-packages"
```

Restart or reload the MCP client after changing server code or configuration. A running MCP process does not hot-reload newly added tools.

## GitHub Setup

```bash
cd /home/maxime/.local/share/mcp-libreoffice-calc
git init
git add .
git commit -m "Initial LibreOffice Calc MCP server"
gh repo create mcp-libreoffice-calc --public --source=. --remote=origin --push
```

## Cleanup Before Publishing

Before pushing, remove local caches and virtual environments:

```bash
find . -type d -name '__pycache__' -prune -exec rm -rf {} +
find . -type d -name '.venv' -prune -exec rm -rf {} +
```

Make sure `.gitignore` excludes at least:

```gitignore
.venv/
__pycache__/
*.pyc
.env
```

## Public Tools

- `open_calc`: open a Calc or Excel workbook.
- `ensure_calc_document`: reuse an open workbook or open it explicitly.
- `list_sheets`: list workbook sheets.
- `add_sheet` / `remove_sheet`: manage sheets.
- `read_cell` / `write_cell` / `write_range`: read and write cells.
- `set_hyperlink`: add a Ctrl-click hyperlink.
- `link_url_range`: convert URL text in a range into Ctrl-click hyperlinks.
- `apply_calc_template`: apply the `pc_nostalgia` template.
- `auto_fit_sheets`: adjust column widths and row heights.
- `save_document`: save the workbook.

## Template Note

`pc_nostalgia` is intentionally opinionated. It is useful as a demo and a personal workflow template, while the core server tools remain generic enough for users to create their own styles.
