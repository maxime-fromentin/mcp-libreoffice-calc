from types import SimpleNamespace

from src import libremcp


class _RecordingCell:
    def __init__(self):
        object.__setattr__(self, "Formula", "")
        object.__setattr__(self, "Value", 0.0)
        object.__setattr__(self, "String", "")
        object.__setattr__(self, "operations", [])

    def __setattr__(self, name, value):
        self.operations.append((name, value))
        object.__setattr__(self, name, value)


class _FakeSheet:
    def __init__(self):
        self.cells: dict[tuple[int, int], _RecordingCell] = {}

    def getCellByPosition(self, col, row):
        key = (col, row)
        if key not in self.cells:
            self.cells[key] = _RecordingCell()
        return self.cells[key]


class TestWriteRow:
    def test_offsets(self):
        sheet = _FakeSheet()
        libremcp._write_row(sheet, 3, ["a", "b", "c"])
        # row_number=3 -> row index 2
        assert sheet.cells[(0, 2)].String == "a"
        assert sheet.cells[(1, 2)].String == "b"
        assert sheet.cells[(2, 2)].String == "c"

    def test_start_col(self):
        sheet = _FakeSheet()
        libremcp._write_row(sheet, 1, [10, 20, 30], start_col=5)
        # row_number=1 -> row index 0
        assert sheet.cells[(5, 0)].Value == 10
        assert sheet.cells[(6, 0)].Value == 20
        assert sheet.cells[(7, 0)].Value == 30

    def test_empty_values(self):
        sheet = _FakeSheet()
        libremcp._write_row(sheet, 2, [])
        assert sheet.cells == {}


class TestStyleRange:
    def test_defaults(self):
        r = SimpleNamespace()
        libremcp._style_range(r)
        assert r.CellBackColor == 0x000000
        assert r.CharColor == 0x00AE00
        assert r.CharFontName == "Consolas"
        assert r.CharHeight == 10
        assert r.CharWeight == 100.0
        assert r.IsTextWrapped is True

    def test_custom(self):
        r = SimpleNamespace()
        libremcp._style_range(r, bg=0xFF0000, fg=0x0000FF, bold=True, size=14)
        assert r.CellBackColor == 0xFF0000
        assert r.CharColor == 0x0000FF
        assert r.CharFontName == "Consolas"
        assert r.CharHeight == 14
        assert r.CharWeight == 150.0
        assert r.IsTextWrapped is True


class TestSetPcNostalgiaCell:
    def test_default(self):
        c = SimpleNamespace()
        libremcp._set_pc_nostalgia_cell(c)
        assert c.CellBackColor == 0x000000
        assert c.CharColor == 0x00AE00
        assert c.CharFontName == "Consolas"
        assert c.CharHeight == 10
        assert c.CharWeight == 100.0
        assert c.IsTextWrapped is True

    def test_header(self):
        c = SimpleNamespace()
        libremcp._set_pc_nostalgia_cell(c, header=True)
        assert c.CellBackColor == 0x1F4E78
        assert c.CharColor == 0xFFFFFF
        assert c.CharFontName == "Consolas"
        assert c.CharHeight == 10
        assert c.CharWeight == 150.0

    def test_title(self):
        c = SimpleNamespace()
        libremcp._set_pc_nostalgia_cell(c, title=True)
        assert c.CellBackColor == 0x1F4E78
        assert c.CharColor == 0xFFFFFF
        assert c.CharFontName == "Consolas"
        assert c.CharHeight == 14
        assert c.CharWeight == 150.0
