"""MCP tools for controlling LibreOffice Calc through UNO."""

from __future__ import annotations

import re
import subprocess
import sys
import time
from pathlib import Path
from typing import Any

from mcp.server.fastmcp import FastMCP
from pydantic import BaseModel, Field


UNO_SITE_PACKAGES = Path("/usr/lib/python3.14/site-packages")
SOFFICE = "/usr/lib/libreoffice/program/soffice"
HOST = "127.0.0.1"
PORT = 2002

if UNO_SITE_PACKAGES.exists() and str(UNO_SITE_PACKAGES) not in sys.path:
    sys.path.append(str(UNO_SITE_PACKAGES))


mcp = FastMCP("MCP LibreOffice Calc")


WORKFLOW_GUIDANCE = {
    "primary_path": "Use LibreOffice UNO tools first for live workbooks and visible demos.",
    "fallback_path": "Use file-level libraries such as openpyxl only as an explicit fallback when UNO cannot provide the required operation.",
    "extension_rule": "If a needed spreadsheet operation is missing, add it as an MCP tool instead of silently bypassing the MCP server.",
    "live_demo_rule": "For live demos, create, edit, style, and navigate sheets through UNO so the user can see changes happening in LibreOffice.",
}


def _log_startup(message: str) -> None:
    """Write startup diagnostics to stderr so MCP JSON-RPC stdout stays clean."""
    print(f"[mcp-libreoffice-calc] {message}", file=sys.stderr, flush=True)


class CellValue(BaseModel):
    document: str
    sheet: str
    cell: str
    value: Any = Field(description="Raw value returned by LibreOffice")
    formula: str | None = None


class RangeWriteResult(BaseModel):
    document: str
    sheet: str
    start_cell: str
    rows: int
    columns: int


class DocumentHandle(BaseModel):
    title: str
    path: str
    was_already_open: bool


class HyperlinkResult(BaseModel):
    document: str
    sheet: str
    cell: str
    url: str
    label: str


class TemplateResult(BaseModel):
    document: str
    sheet: str
    template: str
    styled_range: str


class SheetActivationResult(BaseModel):
    document: str
    active_sheet: str
    selected_cell: str


class DashboardResult(BaseModel):
    document: str
    sheet: str
    recreated: bool
    sections: list[str]
    charts: list[str]
    save_used: bool


def _import_uno():
    try:
        import uno  # type: ignore
        from com.sun.star.beans import PropertyValue  # type: ignore
    except Exception as exc:  # pragma: no cover - machine dependent
        raise RuntimeError(
            "Unable to load UNO. Check your LibreOffice installation and "
            f"the path {UNO_SITE_PACKAGES}. Error: {exc}"
        ) from exc
    return uno, PropertyValue


def _property(name: str, value: Any):
    _, PropertyValue = _import_uno()
    prop = PropertyValue()
    prop.Name = name
    prop.Value = value
    return prop


def _start_listener() -> None:
    accept = f"socket,host={HOST},port={PORT};urp;StarOffice.ComponentContext"
    _log_startup(f"starting LibreOffice UNO listener on {HOST}:{PORT}")
    subprocess.Popen(
        [
            SOFFICE,
            "--accept=" + accept,
            "--norestore",
            "--nodefault",
            "--nologo",
        ],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        start_new_session=True,
    )


def _desktop(start_if_needed: bool = True):
    uno, _ = _import_uno()
    local_ctx = uno.getComponentContext()
    resolver = local_ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local_ctx
    )
    url = f"uno:socket,host={HOST},port={PORT};urp;StarOffice.ComponentContext"

    last_error: Exception | None = None
    for attempt in range(12):
        try:
            ctx = resolver.resolve(url)
            return ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        except Exception as exc:  # pragma: no cover - LibreOffice dependent
            last_error = exc
            if attempt == 0 and start_if_needed:
                _start_listener()
            time.sleep(0.5)
    raise RuntimeError(f"Unable to connect to UNO on {HOST}:{PORT}: {last_error}")


def _file_url(path: str) -> str:
    uno, _ = _import_uno()
    return uno.systemPathToFileUrl(str(Path(path).expanduser().resolve()))


def _path_from_url(url: str) -> str:
    uno, _ = _import_uno()
    try:
        return uno.fileUrlToSystemPath(url)
    except Exception:
        return url


def _documents() -> list[Any]:
    desktop = _desktop()
    enum = desktop.Components.createEnumeration()
    docs: list[Any] = []
    while enum.hasMoreElements():
        doc = enum.nextElement()
        if hasattr(doc, "Sheets"):
            docs.append(doc)
    return docs


def _find_document(path: str | None = None):
    target = str(Path(path).expanduser().resolve()) if path else None
    for doc in _documents():
        if not target:
            return doc
        if _path_from_url(getattr(doc, "URL", "")) == target:
            return doc
    return None


def _document(path: str | None = None, open_if_missing: bool = False):
    doc = _find_document(path)
    if doc is not None:
        return doc
    if path and open_if_missing:
        opened = _open_calc_document(path)
        if opened is not None:
            return opened
        raise RuntimeError(f"LibreOffice could not open: {path}")
    if path:
        raise RuntimeError(
            "Workbook not found in the open Calc documents. "
            "Use ensure_calc_document(path) if you want to open it automatically."
        )
    raise RuntimeError("No Calc workbook is open through the UNO listener.")


def _open_calc_document(path: str, hidden: bool = False):
    return _desktop().loadComponentFromURL(
        _file_url(path),
        "_blank",
        0,
        (_property("Hidden", hidden), _property("ReadOnly", False)),
    )


def _sheet(doc: Any, name: str):
    sheets = doc.Sheets
    if not sheets.hasByName(name):
        raise ValueError(f"Sheet not found: {name}")
    return sheets.getByName(name)


def _cell_indexes(cell: str) -> tuple[int, int]:
    match = re.fullmatch(r"([A-Za-z]+)([1-9][0-9]*)", cell.strip())
    if not match:
        raise ValueError(f"Invalid cell reference: {cell}")
    letters, row_text = match.groups()
    col = 0
    for char in letters.upper():
        col = col * 26 + ord(char) - ord("A") + 1
    return col - 1, int(row_text) - 1


def _cell_name(col: int, row: int) -> str:
    letters = ""
    col += 1
    while col:
        col, rem = divmod(col - 1, 26)
        letters = chr(ord("A") + rem) + letters
    return f"{letters}{row + 1}"


def _range_indexes(range_name: str) -> tuple[int, int, int, int]:
    parts = [part.strip() for part in range_name.split(":")]
    if len(parts) == 1:
        col, row = _cell_indexes(parts[0])
        return col, row, col, row
    if len(parts) != 2:
        raise ValueError(f"Invalid range reference: {range_name}")
    start_col, start_row = _cell_indexes(parts[0])
    end_col, end_row = _cell_indexes(parts[1])
    return min(start_col, end_col), min(start_row, end_row), max(start_col, end_col), max(start_row, end_row)


def _escape_formula_text(value: str) -> str:
    return value.replace('"', '""')


def _set_pc_nostalgia_cell(cell_obj: Any, *, header: bool = False, title: bool = False) -> None:
    cell_obj.CellBackColor = 0x1F4E78 if header or title else 0x000000
    cell_obj.CharColor = 0xFFFFFF if header or title else 0x00AE00
    cell_obj.CharFontName = "Consolas"
    cell_obj.CharHeight = 14 if title else 10
    cell_obj.CharWeight = 150.0 if header or title else 100.0
    cell_obj.IsTextWrapped = True


def _style_range(range_obj: Any, *, bg: int = 0x000000, fg: int = 0x00AE00, bold: bool = False, size: int = 10) -> None:
    range_obj.CellBackColor = bg
    range_obj.CharColor = fg
    range_obj.CharFontName = "Consolas"
    range_obj.CharHeight = size
    range_obj.CharWeight = 150.0 if bold else 100.0
    range_obj.IsTextWrapped = True


def _write_cell(sheet_obj: Any, addr: str, value: Any):
    cell_obj = sheet_obj.getCellRangeByName(addr)
    _set_cell_value(cell_obj, value)
    return cell_obj


def _write_row(sheet_obj: Any, row_number: int, values: list[Any], start_col: int = 0) -> None:
    for offset, value in enumerate(values):
        _set_cell_value(sheet_obj.getCellByPosition(start_col + offset, row_number - 1), value)


def _set_column_widths(sheet_obj: Any, widths: list[int]) -> None:
    for index, width in enumerate(widths):
        sheet_obj.Columns.getByIndex(index).Width = width


def _activate_sheet(doc: Any, sheet_obj: Any, selected_cell: str = "A1") -> None:
    controller = doc.CurrentController
    try:
        controller.setActiveSheet(sheet_obj)
    except Exception:
        controller.ActiveSheet = sheet_obj
    controller.select(sheet_obj.getCellRangeByName(selected_cell))


def _remove_sheet_charts(sheet_obj: Any) -> None:
    try:
        charts = sheet_obj.Charts
        for name in list(charts.getElementNames()):
            charts.removeByName(name)
    except Exception:
        return


def _add_sheet_chart(
    sheet_obj: Any,
    *,
    name: str,
    data_range: tuple[int, int, int, int],
    x: int,
    y: int,
    width: int,
    height: int,
    chart_type: str = "bar",
    title: str | None = None,
    first_row_as_label: bool = True,
    first_col_as_label: bool = True,
) -> str | None:
    """Add a native LibreOffice chart to a sheet using UNO's Calc chart API."""
    try:
        from com.sun.star.awt import Rectangle  # type: ignore
        from com.sun.star.table import CellRangeAddress  # type: ignore
    except Exception:
        return None

    try:
        charts = sheet_obj.Charts
        if charts.hasByName(name):
            charts.removeByName(name)

        start_col, start_row, end_col, end_row = data_range
        range_address = CellRangeAddress()
        range_address.Sheet = sheet_obj.RangeAddress.Sheet
        range_address.StartColumn = start_col
        range_address.StartRow = start_row
        range_address.EndColumn = end_col
        range_address.EndRow = end_row

        rect = Rectangle()
        rect.X = x
        rect.Y = y
        rect.Width = width
        rect.Height = height
        charts.addNewByName(name, rect, (range_address,), first_col_as_label, first_row_as_label)
        chart_doc = charts.getByName(name).EmbeddedObject
        if title:
            chart_doc.HasMainTitle = True
            chart_doc.Title.String = title
        if chart_type == "pie":
            chart_doc.Diagram = chart_doc.createInstance("com.sun.star.chart.PieDiagram")
        elif chart_type == "line":
            chart_doc.Diagram = chart_doc.createInstance("com.sun.star.chart.LineDiagram")
        else:
            chart_doc.Diagram = chart_doc.createInstance("com.sun.star.chart.BarDiagram")
        return name
    except Exception:
        return None


def _clear_cell_borders(cell_obj: Any) -> None:
    try:
        from com.sun.star.table import BorderLine2  # type: ignore

        empty = BorderLine2()
        cell_obj.TopBorder = empty
        cell_obj.BottomBorder = empty
        cell_obj.LeftBorder = empty
        cell_obj.RightBorder = empty
    except Exception:
        # Borders are visual only; do not fail the tool if UNO varies by version.
        return


def _set_cell_value(cell_obj: Any, value: Any) -> None:
    if isinstance(value, str) and value.startswith("="):
        cell_obj.Formula = value
    elif isinstance(value, bool):
        cell_obj.Value = 1 if value else 0
    elif isinstance(value, (int, float)):
        cell_obj.Value = value
    elif value is None:
        cell_obj.String = ""
    else:
        cell_obj.String = str(value)


def _read_cell_value(cell_obj: Any) -> Any:
    formula = getattr(cell_obj, "Formula", "")
    if formula and formula.startswith("="):
        return cell_obj.Value
    text = cell_obj.String
    if text != "":
        return text
    return cell_obj.Value


@mcp.tool()
def mcp_workflow_guidance() -> dict[str, str]:
    """Return the workflow contract for agents using this MCP server.

    Agents should call this first when they need to understand how this server
    expects spreadsheet work to be performed.
    """
    return WORKFLOW_GUIDANCE


@mcp.tool()
def libreoffice_status() -> dict[str, Any]:
    """Check the UNO connection and list open Calc workbooks."""
    docs = _documents()
    return {
        "uno_host": HOST,
        "uno_port": PORT,
        "workflow_guidance": WORKFLOW_GUIDANCE,
        "calc_documents": [
            {"title": getattr(doc, "Title", ""), "path": _path_from_url(getattr(doc, "URL", ""))}
            for doc in docs
        ],
    }


@mcp.tool()
def open_calc(path: str, hidden: bool = False) -> dict[str, Any]:
    """Open a Calc workbook in LibreOffice through UNO."""
    doc = _open_calc_document(path, hidden=hidden)
    if doc is None:
        raise RuntimeError(f"LibreOffice could not open: {path}")
    return {"title": getattr(doc, "Title", ""), "path": _path_from_url(getattr(doc, "URL", ""))}


@mcp.tool()
def ensure_calc_document(path: str, hidden: bool = False) -> DocumentHandle:
    """Return the workbook if already open; otherwise open it explicitly."""
    doc = _find_document(path)
    already_open = doc is not None
    if doc is None:
        doc = _open_calc_document(path, hidden=hidden)
    if doc is None:
        raise RuntimeError(f"LibreOffice could not open: {path}")
    return DocumentHandle(
        title=getattr(doc, "Title", ""),
        path=_path_from_url(getattr(doc, "URL", "")),
        was_already_open=already_open,
    )


@mcp.tool()
def list_calc_documents() -> list[dict[str, str]]:
    """List Calc workbooks reachable through the UNO listener."""
    return [
        {"title": getattr(doc, "Title", ""), "path": _path_from_url(getattr(doc, "URL", ""))}
        for doc in _documents()
    ]


@mcp.tool()
def list_sheets(path: str | None = None) -> list[str]:
    """List sheets from an open or specified workbook."""
    doc = _document(path)
    return [doc.Sheets.getElementNames()[i] for i in range(len(doc.Sheets.getElementNames()))]


@mcp.tool()
def add_sheet(name: str, path: str | None = None, position: int | None = None) -> list[str]:
    """Add a Calc sheet if it does not already exist."""
    doc = _document(path)
    sheets = doc.Sheets
    if not sheets.hasByName(name):
        insert_at = len(sheets.getElementNames()) if position is None else position
        sheets.insertNewByName(name, insert_at)
    return list_sheets(path)


@mcp.tool()
def remove_sheet(name: str, path: str | None = None, save: bool = True) -> list[str]:
    """Remove a Calc sheet. Refuses to remove the last sheet."""
    doc = _document(path)
    sheets = doc.Sheets
    names = list(sheets.getElementNames())
    if name not in names:
        return names
    if len(names) <= 1:
        raise ValueError("Cannot remove the last sheet from the workbook")
    sheets.removeByName(name)
    if save:
        doc.store()
    return list_sheets(path)


@mcp.tool()
def activate_sheet(
    sheet: str,
    cell: str = "A1",
    path: str | None = None,
    save: bool = False,
) -> SheetActivationResult:
    """Activate a sheet and select a cell through UNO for live visual demos."""
    doc = _document(path)
    target_sheet = _sheet(doc, sheet)
    _activate_sheet(doc, target_sheet, cell)
    if save:
        doc.store()
    return SheetActivationResult(
        document=_path_from_url(getattr(doc, "URL", "")),
        active_sheet=doc.CurrentController.ActiveSheet.Name,
        selected_cell=cell,
    )


@mcp.tool()
def read_cell(sheet: str, cell: str, path: str | None = None) -> CellValue:
    """Read a Calc cell."""
    doc = _document(path)
    col, row = _cell_indexes(cell)
    cell_obj = _sheet(doc, sheet).getCellByPosition(col, row)
    formula = getattr(cell_obj, "Formula", "") or None
    return CellValue(
        document=_path_from_url(getattr(doc, "URL", "")),
        sheet=sheet,
        cell=_cell_name(col, row),
        value=_read_cell_value(cell_obj),
        formula=formula if formula and formula.startswith("=") else None,
    )


@mcp.tool()
def write_cell(sheet: str, cell: str, value: Any, path: str | None = None, save: bool = True) -> CellValue:
    """Write one Calc cell directly. Strings starting with '=' become formulas."""
    doc = _document(path)
    col, row = _cell_indexes(cell)
    cell_obj = _sheet(doc, sheet).getCellByPosition(col, row)
    _set_cell_value(cell_obj, value)
    if save:
        doc.store()
    return read_cell(sheet=sheet, cell=_cell_name(col, row), path=path)


@mcp.tool()
def write_range(sheet: str, start_cell: str, values: list[list[Any]], path: str | None = None, save: bool = True) -> RangeWriteResult:
    """Write a 2D table in Calc from a starting cell."""
    if not values:
        raise ValueError("values cannot be empty")
    width = max(len(row) for row in values)
    if width == 0:
        raise ValueError("values must contain at least one column")

    doc = _document(path)
    start_col, start_row = _cell_indexes(start_cell)
    target_sheet = _sheet(doc, sheet)
    for r_offset, row_values in enumerate(values):
        for c_offset, value in enumerate(row_values):
            _set_cell_value(target_sheet.getCellByPosition(start_col + c_offset, start_row + r_offset), value)
    if save:
        doc.store()
    return RangeWriteResult(
        document=_path_from_url(getattr(doc, "URL", "")),
        sheet=sheet,
        start_cell=_cell_name(start_col, start_row),
        rows=len(values),
        columns=width,
    )


@mcp.tool()
def set_hyperlink(
    sheet: str,
    cell: str,
    url: str,
    label: str | None = None,
    path: str | None = None,
    save: bool = True,
) -> HyperlinkResult:
    """Add a Ctrl-click hyperlink to a Calc cell.

    The link is written with LibreOffice's HYPERLINK formula to stay compatible
    with `.xlsx` files. If `label` is omitted, the URL is used as visible text.
    """
    if not url.startswith(("http://", "https://", "mailto:")):
        raise ValueError("url must start with http://, https://, or mailto:")
    doc = _document(path)
    col, row = _cell_indexes(cell)
    target = _sheet(doc, sheet).getCellByPosition(col, row)
    visible = label if label is not None else url
    target.Formula = f'=HYPERLINK("{_escape_formula_text(url)}";"{_escape_formula_text(visible)}")'
    target.CharColor = 0x00B0F0
    target.CharUnderline = 1
    target.CharFontName = "Consolas"
    if save:
        doc.store()
    return HyperlinkResult(
        document=_path_from_url(getattr(doc, "URL", "")),
        sheet=sheet,
        cell=_cell_name(col, row),
        url=url,
        label=visible,
    )


@mcp.tool()
def set_internal_hyperlink(
    sheet: str,
    cell: str,
    target_sheet: str,
    target_cell: str = "A1",
    label: str | None = None,
    path: str | None = None,
    save: bool = True,
) -> HyperlinkResult:
    """Add a Ctrl-click hyperlink to another sheet/cell in the same workbook."""
    doc = _document(path)
    _sheet(doc, target_sheet)
    col, row = _cell_indexes(cell)
    _cell_indexes(target_cell)
    source_sheet = _sheet(doc, sheet)
    target = source_sheet.getCellByPosition(col, row)
    url = f"#{target_sheet}.{target_cell}"
    visible = label if label is not None else f"{target_sheet}!{target_cell}"
    target.Formula = f'=HYPERLINK("{_escape_formula_text(url)}";"{_escape_formula_text(visible)}")'
    target.CharColor = 0x00B0F0
    target.CharUnderline = 1
    target.CharFontName = "Consolas"
    if save:
        doc.store()
    return HyperlinkResult(
        document=_path_from_url(getattr(doc, "URL", "")),
        sheet=sheet,
        cell=_cell_name(col, row),
        url=url,
        label=visible,
    )


@mcp.tool()
def link_url_range(
    sheet: str,
    cell_range: str,
    path: str | None = None,
    save: bool = True,
) -> dict[str, Any]:
    """Convert plain-text URLs in a range into Ctrl-click hyperlinks.

    Example: `link_url_range(sheet="vendor_purchases", cell_range="K8:K20")`.
    """
    doc = _document(path)
    target_sheet = _sheet(doc, sheet)
    start_col, start_row, end_col, end_row = _range_indexes(cell_range)
    linked: list[str] = []
    for row in range(start_row, end_row + 1):
        for col in range(start_col, end_col + 1):
            cell_obj = target_sheet.getCellByPosition(col, row)
            value = cell_obj.String.strip()
            if value.startswith(("http://", "https://", "mailto:")):
                cell_obj.Formula = f'=HYPERLINK("{_escape_formula_text(value)}";"{_escape_formula_text(value)}")'
                cell_obj.CharColor = 0x00B0F0
                cell_obj.CharUnderline = 1
                cell_obj.CharFontName = "Consolas"
                linked.append(_cell_name(col, row))
    if save:
        doc.store()
    return {"document": _path_from_url(getattr(doc, "URL", "")), "sheet": sheet, "linked_cells": linked}


@mcp.tool()
def apply_calc_template(
    sheet: str,
    template: str = "pc_nostalgia",
    path: str | None = None,
    used_range: str | None = None,
    title_rows: int = 1,
    header_rows: list[int] | None = None,
    visible_rows: int = 45,
    visible_columns: int = 14,
    clear_borders: bool = True,
    save: bool = True,
) -> TemplateResult:
    """Apply a visual Calc template.

    Available template: `pc_nostalgia`.
    Style: black background, green terminal text, blue title/header rows,
    Consolas font, hidden gridlines, and no vertical borders. `header_rows`
    uses 1-based Calc row numbers, for example `[7]`.
    """
    if template != "pc_nostalgia":
        raise ValueError("currently supported template: pc_nostalgia")
    doc = _document(path)
    target_sheet = _sheet(doc, sheet)
    if used_range:
        start_col, start_row, end_col, end_row = _range_indexes(used_range)
    else:
        start_col, start_row = 0, 0
        end_col, end_row = visible_columns - 1, visible_rows - 1
    header_set = {row_number - 1 for row_number in (header_rows or [])}
    for row in range(start_row, end_row + 1):
        for col in range(start_col, end_col + 1):
            cell_obj = target_sheet.getCellByPosition(col, row)
            _set_pc_nostalgia_cell(
                cell_obj,
                header=row in header_set,
                title=row < title_rows,
            )
            if clear_borders:
                _clear_cell_borders(cell_obj)
    try:
        doc.CurrentController.ShowGrid = False
    except Exception:
        pass
    try:
        target_sheet.TabColor = 0x00AE00
    except Exception:
        pass
    if save:
        doc.store()
    return TemplateResult(
        document=_path_from_url(getattr(doc, "URL", "")),
        sheet=sheet,
        template=template,
        styled_range=f"{_cell_name(start_col, start_row)}:{_cell_name(end_col, end_row)}",
    )


@mcp.tool()
def create_gpu_purchase_dashboard_live(
    path: str | None = None,
    dashboard_sheet: str = "Dashboard_live",
    recreate: bool = True,
    pause_seconds: float = 0.25,
    save: bool = True,
) -> DashboardResult:
    """Create a live GPU purchase dashboard through UNO.

    This tool intentionally uses UNO as the primary path so the user can see the
    sheet being removed, recreated, activated, styled, and filled in LibreOffice.
    It does not use openpyxl. If richer unsupported features are needed, add a
    new MCP tool instead of bypassing the server.
    """
    doc = _document(path)
    sheets = doc.Sheets
    recreated = False
    if sheets.hasByName(dashboard_sheet) and recreate:
        if len(sheets.getElementNames()) <= 1:
            raise ValueError("Cannot recreate the only sheet in the workbook")
        sheets.removeByName(dashboard_sheet)
        recreated = True
        if pause_seconds > 0:
            time.sleep(pause_seconds)
    if not sheets.hasByName(dashboard_sheet):
        sheets.insertNewByName(dashboard_sheet, 0)
        recreated = True
        if pause_seconds > 0:
            time.sleep(pause_seconds)

    sheet_obj = sheets.getByName(dashboard_sheet)
    _activate_sheet(doc, sheet_obj, "A1")
    _remove_sheet_charts(sheet_obj)
    _style_range(sheet_obj.getCellRangeByName("A1:R80"))
    _style_range(sheet_obj.getCellRangeByName("A1:R1"), bg=0x1F4E78, fg=0xFFFFFF, bold=True, size=16)
    _write_cell(sheet_obj, "A1", "Live GPU Purchase Dashboard")
    _write_cell(sheet_obj, "A2", "Sources")
    _write_cell(sheet_obj, "B2", "achat_Vendeur_spe, Achat_RTX3090, Achat_RTX4090, Achat_RTX5090, Calcul")
    _write_cell(sheet_obj, "A3", "Workflow")
    _write_cell(sheet_obj, "B3", "Built live through MCP/UNO. Use file-level libraries only as explicit fallback.")
    _write_cell(sheet_obj, "A4", "Guidance")
    _write_cell(sheet_obj, "B4", WORKFLOW_GUIDANCE["primary_path"])
    if pause_seconds > 0:
        time.sleep(pause_seconds)

    sections = ["title", "kpis", "analysis", "ebay_prices", "etb_watchlist", "financial_model", "navigation"]
    _style_range(sheet_obj.getCellRangeByName("A6:C6"), bg=0x1F4E78, fg=0xFFFFFF, bold=True)
    _write_cell(sheet_obj, "A6", "Key Indicators")
    kpis = [
        ("ETB product rows", "=COUNTA('achat_Vendeur_spe'.C8:C53)"),
        ("ETB priced rows", "=COUNT('achat_Vendeur_spe'.G8:G53)"),
        ("Max VRAM found", "=MAX('achat_Vendeur_spe'.E8:E53)"),
        ("Lowest ETB price", "=MIN('achat_Vendeur_spe'.G8:G53)"),
        ("Avg price >=24GB", "=AVERAGEIF('achat_Vendeur_spe'.E8:E53;\">=24\";'achat_Vendeur_spe'.G8:G53)"),
        ("3090 median eBay", "='Achat_RTX3090'.B8"),
        ("4090 median eBay", "='Achat_RTX4090'.B9"),
        ("5090 median eBay", "='Achat_RTX5090'.B8"),
        ("Cash-flow after debt", "='Calcul'.B6"),
        ("Occupation min cash", "='Calcul'.B11"),
    ]
    for row, (label, formula) in enumerate(kpis, 7):
        _write_cell(sheet_obj, f"A{row}", label)
        _write_cell(sheet_obj, f"B{row}", formula)
    if pause_seconds > 0:
        time.sleep(pause_seconds)

    _style_range(sheet_obj.getCellRangeByName("J6:R6"), bg=0x1F4E78, fg=0xFFFFFF, bold=True)
    _write_cell(sheet_obj, "J6", "Analysis")
    analysis = [
        "RTX 3090 remains the lowest-cost VRAM path if listings are clean.",
        "RTX 4090 is faster but still 24 GB; buy only if throughput matters.",
        "RTX 5090 adds 32 GB, but current purchase price is high.",
        "ETB pro cards reduce listing risk but are expensive versus used consumer GPUs.",
        "Watch RTX PRO 6000 96 GB, RTX 6000 Ada 48 GB, RTX A6000 48 GB, RTX 5000 Ada 32 GB.",
    ]
    for row, text in enumerate(analysis, 7):
        _write_cell(sheet_obj, f"J{row}", f"- {text}")
    if pause_seconds > 0:
        time.sleep(pause_seconds)

    _write_row(sheet_obj, 19, ["GPU class", "VRAM GB", "Min EUR", "P25 EUR", "Median EUR", "P75 EUR", "Max EUR", "8-card median EUR"])
    _style_range(sheet_obj.getCellRangeByName("A19:H19"), bg=0x1F4E78, fg=0xFFFFFF, bold=True)
    gpu_rows = [
        ["RTX 3090", 24, "='Achat_RTX3090'.B6", "='Achat_RTX3090'.B7", "='Achat_RTX3090'.B8", "='Achat_RTX3090'.B9", "='Achat_RTX3090'.B10", "='Achat_RTX3090'.C8"],
        ["RTX 4090", 24, "='Achat_RTX4090'.B7", "='Achat_RTX4090'.B8", "='Achat_RTX4090'.B9", "='Achat_RTX4090'.B10", "='Achat_RTX4090'.B11", "='Achat_RTX4090'.D9"],
        ["RTX 5090", 32, "='Achat_RTX5090'.B6", "='Achat_RTX5090'.B7", "='Achat_RTX5090'.B8", "='Achat_RTX5090'.B9", "='Achat_RTX5090'.B10", "='Achat_RTX5090'.E8"],
    ]
    for row, values in enumerate(gpu_rows, 20):
        _write_row(sheet_obj, row, values)
    if pause_seconds > 0:
        time.sleep(pause_seconds)

    _write_row(sheet_obj, 26, ["Rank", "Product", "Family", "VRAM GB", "Price EUR", "EUR / GB", "Verdict"])
    _style_range(sheet_obj.getCellRangeByName("A26:G26"), bg=0x1F4E78, fg=0xFFFFFF, bold=True)
    for output_row, source_row in enumerate([8, 9, 10, 12, 14, 15, 22, 34], 27):
        _write_row(
            sheet_obj,
            output_row,
            [
                output_row - 26,
                f"='achat_Vendeur_spe'.C{source_row}",
                f"='achat_Vendeur_spe'.D{source_row}",
                f"='achat_Vendeur_spe'.E{source_row}",
                f"='achat_Vendeur_spe'.G{source_row}",
                f"=IF(E{output_row}>0;E{output_row}/D{output_row};\"\")",
                f"='achat_Vendeur_spe'.J{source_row}",
            ],
        )
    if pause_seconds > 0:
        time.sleep(pause_seconds)

    source_sheet = _sheet(doc, "achat_Vendeur_spe")
    verdict_counts: dict[str, int] = {}
    for source_row in range(8, 54):
        verdict = source_sheet.getCellByPosition(9, source_row - 1).String
        if verdict:
            verdict_counts[verdict] = verdict_counts.get(verdict, 0) + 1
    _write_cell(sheet_obj, "J25", "ETB verdict mix")
    _style_range(sheet_obj.getCellRangeByName("J25:K26"), bg=0x1F4E78, fg=0xFFFFFF, bold=True)
    _write_row(sheet_obj, 26, ["Verdict", "Count"], start_col=9)
    for row, (verdict, count) in enumerate(sorted(verdict_counts.items()), 27):
        _write_row(sheet_obj, row, [verdict, count], start_col=9)

    _write_row(sheet_obj, 45, ["Metric", "Value", "Unit", "Source"])
    _style_range(sheet_obj.getCellRangeByName("A45:D45"), bg=0x1F4E78, fg=0xFFFFFF, bold=True)
    financial_rows = [
        ("Annual gross revenue", "='Calcul'.B3", "EUR/year", "Calcul!B3"),
        ("Annual electricity", "='Calcul'.B4", "EUR/year", "Calcul!B4"),
        ("Maintenance", "='Hypotheses'.B8", "EUR/year", "Hypotheses!B8"),
        ("Annual debt service", "='Calcul'.B8", "EUR/year", "Calcul!B8"),
        ("Cash-flow after debt", "='Calcul'.B6", "EUR/year", "Calcul!B6"),
        ("Occupation min cash", "='Calcul'.B11", "%", "Calcul!B11"),
    ]
    for row, values in enumerate(financial_rows, 46):
        _write_row(sheet_obj, row, list(values))

    _write_cell(sheet_obj, "J15", "Source navigation")
    _style_range(sheet_obj.getCellRangeByName("J15:M15"), bg=0x1F4E78, fg=0xFFFFFF, bold=True)
    navigation = [
        ("Specialized vendors", "achat_Vendeur_spe"),
        ("RTX 3090 eBay", "Achat_RTX3090"),
        ("RTX 4090 eBay", "Achat_RTX4090"),
        ("RTX 5090 eBay", "Achat_RTX5090"),
        ("Financial model", "Calcul"),
    ]
    for row, (label, target_sheet) in enumerate(navigation, 16):
        cell_obj = _write_cell(sheet_obj, f"J{row}", f'=HYPERLINK("#{target_sheet}.A1";"{label}")')
        cell_obj.CharColor = 0x00B0F0
        cell_obj.CharUnderline = 1

    charts_created: list[str] = []
    chart_specs = [
        {
            "name": "EbayPriceDistribution",
            "data_range": (0, 18, 6, 21),
            "x": 16500,
            "y": 9500,
            "width": 16000,
            "height": 8500,
            "chart_type": "bar",
            "title": "eBay purchase price distribution",
        },
        {
            "name": "EtbEuroPerGb",
            "data_range": (2, 25, 5, 34),
            "x": 16500,
            "y": 18500,
            "width": 16000,
            "height": 8500,
            "chart_type": "bar",
            "title": "Selected ETB cards: EUR per GB",
        },
        {
            "name": "FinancialComponents",
            "data_range": (0, 44, 1, 50),
            "x": 800,
            "y": 33000,
            "width": 15000,
            "height": 8500,
            "chart_type": "bar",
            "title": "Annual financial model components",
        },
        {
            "name": "EtbVerdictDistribution",
            "data_range": (9, 25, 10, 26 + len(verdict_counts)),
            "x": 32500,
            "y": 9500,
            "width": 10000,
            "height": 8500,
            "chart_type": "pie",
            "title": "ETB verdict distribution",
        },
    ]
    for spec in chart_specs:
        chart_name = _add_sheet_chart(sheet_obj, **spec)
        if chart_name:
            charts_created.append(chart_name)
    if pause_seconds > 0:
        time.sleep(pause_seconds)

    _set_column_widths(sheet_obj, [6000, 3800, 3600, 3600, 3600, 3600, 4200, 4200, 1200, 6800, 4200, 4200, 4200, 4200, 4200, 4200, 4200, 4200])
    try:
        doc.CurrentController.ShowGrid = False
    except Exception:
        pass
    try:
        sheet_obj.TabColor = 0x00B0F0
    except Exception:
        pass
    _activate_sheet(doc, sheet_obj, "A1")
    if save:
        doc.store()
    return DashboardResult(
        document=_path_from_url(getattr(doc, "URL", "")),
        sheet=dashboard_sheet,
        recreated=recreated,
        sections=sections,
        charts=charts_created,
        save_used=save,
    )


@mcp.tool()
def auto_fit_sheets(
    path: str | None = None,
    sheets: list[str] | None = None,
    save: bool = True,
    min_column_width: int = 3500,
    max_column_width: int = 16000,
) -> dict[str, Any]:
    """Adjust Calc sheets without letting LibreOffice create overly narrow columns.

    Widths are in hundredths of a millimeter, LibreOffice's UNO unit.
    """
    doc = _document(path)
    names = list(doc.Sheets.getElementNames()) if sheets is None else sheets
    formatted: list[str] = []
    for name in names:
        target_sheet = _sheet(doc, name)
        cursor = target_sheet.createCursor()
        cursor.gotoStartOfUsedArea(False)
        cursor.gotoEndOfUsedArea(True)
        used = cursor.RangeAddress
        if used.EndColumn < used.StartColumn or used.EndRow < used.StartRow:
            continue
        used_range = target_sheet.getCellRangeByPosition(
            used.StartColumn,
            used.StartRow,
            used.EndColumn,
            used.EndRow,
        )
        used_range.IsTextWrapped = True
        for col in range(used.StartColumn, used.EndColumn + 1):
            column = target_sheet.Columns.getByIndex(col)
            column.OptimalWidth = True
            if column.Width < min_column_width:
                column.Width = min_column_width
            elif column.Width > max_column_width:
                column.Width = max_column_width
        for row in range(used.StartRow, used.EndRow + 1):
            target_sheet.Rows.getByIndex(row).OptimalHeight = True
        formatted.append(name)
    if save:
        doc.store()
    return {"formatted_sheets": formatted}


@mcp.tool()
def save_document(path: str | None = None) -> dict[str, str]:
    """Save the Calc workbook."""
    doc = _document(path)
    doc.store()
    return {"saved": _path_from_url(getattr(doc, "URL", ""))}


def main() -> None:
    _log_startup("starting MCP server")
    _log_startup(f"python={sys.version.split()[0]} executable={sys.executable}")
    _log_startup(f"cwd={Path.cwd()}")
    _log_startup(f"uno_site_packages={UNO_SITE_PACKAGES} exists={UNO_SITE_PACKAGES.exists()}")
    _log_startup(f"soffice={SOFFICE} exists={Path(SOFFICE).exists()}")
    _log_startup(f"uno_endpoint={HOST}:{PORT}")
    try:
        docs = _documents()
        _log_startup(f"UNO connection OK; open_calc_documents={len(docs)}")
        for doc in docs:
            title = getattr(doc, "Title", "")
            path = _path_from_url(getattr(doc, "URL", ""))
            _log_startup(f"document title={title!r} path={path!r}")
    except Exception as exc:
        _log_startup(f"UNO connection not ready yet: {exc}")
    mcp.run()


if __name__ == "__main__":
    main()
