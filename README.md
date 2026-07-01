# MCP LibreOffice Calc

MCP server for controlling LibreOffice Calc workbooks through UNO.

This project focuses on reliable spreadsheet operations against an already-open workbook, or a workbook opened explicitly by the server. It avoids silent file rewrites outside LibreOffice and provides practical tools for writing cells, ranges, hyperlinks, and reusable visual templates.

## Demo

Short usage demo: [assets/demo.mp4](assets/demo.mp4)

The demo walks through a typical MCP flow: opening a workbook, adding a sheet, writing a range, creating Ctrl-click hyperlinks, applying the `pc_nostalgia` template, auto-fitting, and saving.

## Available Tools

- `libreoffice_status`: check the UNO connection and list reachable Calc workbooks.
- `open_calc`: open a Calc or Excel workbook through LibreOffice.
- `ensure_calc_document`: reuse an already-open workbook, or open it explicitly if missing.
- `list_calc_documents`: list Calc workbooks reachable through UNO.
- `list_sheets`: list workbook sheets.
- `add_sheet`: add a sheet if it does not already exist.
- `remove_sheet`: remove a sheet, while refusing to delete the last sheet.
- `activate_sheet`: activate a sheet and select a cell for live visual workflows.
- `read_cell`: read a cell value and formula.
- `write_cell`: write one cell directly.
- `write_range`: write a 2D table from a starting cell.
- `set_hyperlink`: add a Ctrl-click hyperlink to a Calc cell.
- `set_internal_hyperlink`: add a Ctrl-click hyperlink to another sheet/cell.
- `link_url_range`: convert plain-text URLs in a range into Ctrl-click hyperlinks.
- `apply_calc_template`: apply a visual template, including `pc_nostalgia`.
- `create_gpu_purchase_dashboard_live`: build the GPU purchase dashboard through UNO for live demos, including native Calc charts.
- `auto_fit_sheets`: adjust column widths, row heights, and wrapping.
- `save_document`: save the workbook.

Build the live GPU purchase dashboard, including charts:

```text
create_gpu_purchase_dashboard_live(
  path="/path/to/workbook.xlsx",
  dashboard_sheet="Dashboard_live",
  recreate=true,
  pause_seconds=0.25
)
```

## Manual Start

```bash
cd /home/maxime/.local/share/mcp-libreoffice-calc
uv run python src/main.py
```

The server starts or joins a LibreOffice UNO listener on `127.0.0.1:2002`.

## Code Configuration

Example `~/.code/config.toml` entry:

```toml
[mcp_servers.libreoffice_calc]
command = "uv"
args = ["run", "python", "/home/maxime/.local/share/mcp-libreoffice-calc/src/main.py"]
cwd = "/home/maxime/.local/share/mcp-libreoffice-calc"

[mcp_servers.libreoffice_calc.env]
PYTHONPATH = "/home/maxime/.local/share/mcp-libreoffice-calc/src:/usr/lib/python3.14/site-packages"
```

## Notes

LibreOffice must be reachable through UNO to control an already-open workbook. The recommended flow is to call `ensure_calc_document` once, then use `add_sheet`, `write_cell`, `write_range`, `set_hyperlink`, or `link_url_range`.

MCP servers load tool definitions at process startup. Restart or reload your MCP client after changing server code or configuration.

Files copied from the original base project remain covered by their original license. The main server in `src/` has been rewritten for this focused Calc workflow.
