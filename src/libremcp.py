"""MCP Libre Perso: outils MCP pour piloter LibreOffice Calc en direct."""

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


mcp = FastMCP("MCP Libre Perso")


class CellValue(BaseModel):
    document: str
    sheet: str
    cell: str
    value: Any = Field(description="Valeur brute renvoyee par LibreOffice")
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


def _import_uno():
    try:
        import uno  # type: ignore
        from com.sun.star.beans import PropertyValue  # type: ignore
    except Exception as exc:  # pragma: no cover - depend de la machine
        raise RuntimeError(
            "Impossible de charger UNO. Verifie libreoffice-fresh/libreoffice-still et "
            f"le chemin {UNO_SITE_PACKAGES}. Erreur: {exc}"
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
        except Exception as exc:  # pragma: no cover - depend de LibreOffice
            last_error = exc
            if attempt == 0 and start_if_needed:
                _start_listener()
            time.sleep(0.5)
    raise RuntimeError(f"Connexion UNO impossible sur {HOST}:{PORT}: {last_error}")


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
        raise RuntimeError(f"LibreOffice n'a pas pu ouvrir: {path}")
    if path:
        raise RuntimeError(
            "Classeur non trouve dans les documents Calc ouverts. "
            "Utilise ensure_calc_document(path) si tu veux l'ouvrir automatiquement."
        )
    raise RuntimeError("Aucun classeur Calc ouvert via le listener UNO.")


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
        raise ValueError(f"Feuille introuvable: {name}")
    return sheets.getByName(name)


def _cell_indexes(cell: str) -> tuple[int, int]:
    match = re.fullmatch(r"([A-Za-z]+)([1-9][0-9]*)", cell.strip())
    if not match:
        raise ValueError(f"Reference cellule invalide: {cell}")
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
        raise ValueError(f"Reference plage invalide: {range_name}")
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


def _clear_cell_borders(cell_obj: Any) -> None:
    try:
        from com.sun.star.table import BorderLine2  # type: ignore

        empty = BorderLine2()
        cell_obj.TopBorder = empty
        cell_obj.BottomBorder = empty
        cell_obj.LeftBorder = empty
        cell_obj.RightBorder = empty
    except Exception:
        # Les bordures sont purement visuelles; ne pas bloquer l'outil si UNO varie selon version.
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
def libreoffice_status() -> dict[str, Any]:
    """Verifier la connexion UNO et lister les classeurs Calc ouverts."""
    docs = _documents()
    return {
        "uno_host": HOST,
        "uno_port": PORT,
        "calc_documents": [
            {"title": getattr(doc, "Title", ""), "path": _path_from_url(getattr(doc, "URL", ""))}
            for doc in docs
        ],
    }


@mcp.tool()
def open_calc(path: str, hidden: bool = False) -> dict[str, Any]:
    """Ouvrir un classeur Calc dans LibreOffice via UNO."""
    doc = _open_calc_document(path, hidden=hidden)
    if doc is None:
        raise RuntimeError(f"LibreOffice n'a pas pu ouvrir: {path}")
    return {"title": getattr(doc, "Title", ""), "path": _path_from_url(getattr(doc, "URL", ""))}


@mcp.tool()
def ensure_calc_document(path: str, hidden: bool = False) -> DocumentHandle:
    """Retourner le classeur s'il est deja ouvert; sinon l'ouvrir explicitement."""
    doc = _find_document(path)
    already_open = doc is not None
    if doc is None:
        doc = _open_calc_document(path, hidden=hidden)
    if doc is None:
        raise RuntimeError(f"LibreOffice n'a pas pu ouvrir: {path}")
    return DocumentHandle(
        title=getattr(doc, "Title", ""),
        path=_path_from_url(getattr(doc, "URL", "")),
        was_already_open=already_open,
    )


@mcp.tool()
def list_calc_documents() -> list[dict[str, str]]:
    """Lister les classeurs Calc accessibles par le listener UNO."""
    return [
        {"title": getattr(doc, "Title", ""), "path": _path_from_url(getattr(doc, "URL", ""))}
        for doc in _documents()
    ]


@mcp.tool()
def list_sheets(path: str | None = None) -> list[str]:
    """Lister les feuilles d'un classeur ouvert ou du classeur indique."""
    doc = _document(path)
    return [doc.Sheets.getElementNames()[i] for i in range(len(doc.Sheets.getElementNames()))]


@mcp.tool()
def add_sheet(name: str, path: str | None = None, position: int | None = None) -> list[str]:
    """Ajouter une feuille Calc si elle n'existe pas encore."""
    doc = _document(path)
    sheets = doc.Sheets
    if not sheets.hasByName(name):
        insert_at = len(sheets.getElementNames()) if position is None else position
        sheets.insertNewByName(name, insert_at)
    return list_sheets(path)


@mcp.tool()
def remove_sheet(name: str, path: str | None = None, save: bool = True) -> list[str]:
    """Supprimer une feuille Calc. Refuse de supprimer la derniere feuille."""
    doc = _document(path)
    sheets = doc.Sheets
    names = list(sheets.getElementNames())
    if name not in names:
        return names
    if len(names) <= 1:
        raise ValueError("Impossible de supprimer la derniere feuille du classeur")
    sheets.removeByName(name)
    if save:
        doc.store()
    return list_sheets(path)


@mcp.tool()
def read_cell(sheet: str, cell: str, path: str | None = None) -> CellValue:
    """Lire une cellule Calc."""
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
    """Ecrire une cellule Calc en direct. Les chaines commencant par '=' deviennent des formules."""
    doc = _document(path)
    col, row = _cell_indexes(cell)
    cell_obj = _sheet(doc, sheet).getCellByPosition(col, row)
    _set_cell_value(cell_obj, value)
    if save:
        doc.store()
    return read_cell(sheet=sheet, cell=_cell_name(col, row), path=path)


@mcp.tool()
def write_range(sheet: str, start_cell: str, values: list[list[Any]], path: str | None = None, save: bool = True) -> RangeWriteResult:
    """Ecrire un tableau 2D dans Calc a partir d'une cellule de depart."""
    if not values:
        raise ValueError("values ne peut pas etre vide")
    width = max(len(row) for row in values)
    if width == 0:
        raise ValueError("values doit contenir au moins une colonne")

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
    """Ajouter un lien cliquable Ctrl+clic dans une cellule Calc.

    Le lien est ecrit avec la formule LibreOffice HYPERLINK pour rester compatible
    avec les fichiers `.xlsx`. Si `label` est omis, l'URL sert de texte visible.
    """
    if not url.startswith(("http://", "https://", "mailto:")):
        raise ValueError("url doit commencer par http://, https:// ou mailto:")
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
def link_url_range(
    sheet: str,
    cell_range: str,
    path: str | None = None,
    save: bool = True,
) -> dict[str, Any]:
    """Transformer les URLs texte d'une plage en liens Ctrl+clic.

    Exemple: `link_url_range(sheet="achat_Vendeur_spe", cell_range="K8:K20")`.
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
    """Appliquer un template visuel Calc.

    Template disponible: `pc_nostalgia`.
    Style: fond noir, texte vert terminal, titres/en-tetes bleus, police Consolas,
    quadrillage masque, sans bordures verticales. `header_rows` utilise des numeros
    de lignes Calc en base 1, par exemple `[7]`.
    """
    if template != "pc_nostalgia":
        raise ValueError("template supporte actuellement: pc_nostalgia")
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
def auto_fit_sheets(
    path: str | None = None,
    sheets: list[str] | None = None,
    save: bool = True,
    min_column_width: int = 3500,
    max_column_width: int = 16000,
) -> dict[str, Any]:
    """Ajuster les feuilles Calc sans laisser LibreOffice creer des colonnes trop etroites.

    Les largeurs sont en centiemes de millimetre, unite UNO de LibreOffice.
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
    """Sauvegarder le classeur Calc."""
    doc = _document(path)
    doc.store()
    return {"saved": _path_from_url(getattr(doc, "URL", ""))}


def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()
