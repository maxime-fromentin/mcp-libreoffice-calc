import pytest
from src import libremcp


class FakeCell:
    def __init__(self):
        self.Formula = ""
        self.Value = 0.0
        self.String = ""
        self.CellBackColor = None
        self.CharColor = None
        self.CharFontName = None
        self.CharHeight = None
        self.CharWeight = None
        self.IsTextWrapped = None
        self.CharUnderline = None


class FakeSheet:
    def __init__(self, name):
        self._name = name
        self.Name = name
        self._cells: dict[tuple[int, int], FakeCell] = {}

    def getCellByPosition(self, col, row):
        key = (col, row)
        if key not in self._cells:
            self._cells[key] = FakeCell()
        return self._cells[key]

    def getCellRangeByName(self, addr):
        return FakeCell()


class FakeSheets:
    def __init__(self):
        self._sheets: dict[str, FakeSheet] = {}
        self._inserts: list[tuple[str, int]] = []

    def hasByName(self, name):
        return name in self._sheets

    def getByName(self, name):
        if name not in self._sheets:
            self._sheets[name] = FakeSheet(name)
        return self._sheets[name]

    def getElementNames(self):
        return list(self._sheets.keys())

    def insertNewByName(self, name, position):
        self._sheets[name] = FakeSheet(name)
        self._inserts.append((name, position))

    def removeByName(self, name):
        self._sheets.pop(name, None)


class FakeDoc:
    def __init__(self):
        self.Sheets = FakeSheets()
        self.URL = ""
        self.store_count = 0

    def store(self):
        self.store_count += 1


@pytest.fixture
def fake_doc(monkeypatch):
    doc = FakeDoc()
    doc.Sheets._sheets["Sheet1"] = FakeSheet("Sheet1")
    monkeypatch.setattr(libremcp, "_document", lambda path=None: doc)
    monkeypatch.setattr(libremcp, "_path_from_url", lambda url: url)
    return doc


class TestWriteCell:
    def test_write_cell_int(self, fake_doc):
        result = libremcp.write_cell(sheet="Sheet1", cell="A1", value=42, save=False)
        assert result.value == 42
        assert result.cell == "A1"
        assert result.sheet == "Sheet1"
        assert result.document == ""
        assert result.formula is None
        assert fake_doc.store_count == 0

    def test_write_cell_float(self, fake_doc):
        result = libremcp.write_cell(sheet="Sheet1", cell="B2", value=3.14, save=False)
        assert result.value == 3.14
        assert result.cell == "B2"

    def test_write_cell_string(self, fake_doc):
        result = libremcp.write_cell(sheet="Sheet1", cell="C3", value="hello", save=False)
        assert result.value == "hello"
        assert result.cell == "C3"
        assert result.formula is None

    def test_write_cell_formula(self, fake_doc):
        result = libremcp.write_cell(
            sheet="Sheet1", cell="A1", value="=SUM(B1:B3)", save=False
        )
        assert result.formula == "=SUM(B1:B3)"
        assert result.cell == "A1"
        # formula not evaluated in the mock; Value stays at default
        assert result.value == 0.0

    def test_write_cell_calls_store_when_save_true(self, fake_doc):
        libremcp.write_cell(sheet="Sheet1", cell="A1", value=1, save=True)
        assert fake_doc.store_count == 1


class TestReadCell:
    def test_read_cell_numeric(self, fake_doc):
        libremcp.write_cell(sheet="Sheet1", cell="A1", value=99, save=False)
        result = libremcp.read_cell(sheet="Sheet1", cell="A1")
        assert result.value == 99
        assert result.cell == "A1"
        assert result.sheet == "Sheet1"
        assert result.formula is None

    def test_read_cell_string(self, fake_doc):
        libremcp.write_cell(sheet="Sheet1", cell="B2", value="test text", save=False)
        result = libremcp.read_cell(sheet="Sheet1", cell="B2")
        assert result.value == "test text"
        assert result.formula is None

    def test_read_cell_formula(self, fake_doc):
        libremcp.write_cell(sheet="Sheet1", cell="C1", value="=1+2", save=False)
        result = libremcp.read_cell(sheet="Sheet1", cell="C1")
        assert result.formula == "=1+2"

    def test_read_cell_normalises_name(self, fake_doc):
        libremcp.write_cell(sheet="Sheet1", cell="AA10", value=7, save=False)
        result = libremcp.read_cell(sheet="Sheet1", cell="AA10")
        assert result.cell == "AA10"
        assert result.value == 7


class TestWriteRange:
    def test_write_range_basic(self, fake_doc):
        result = libremcp.write_range(
            sheet="Sheet1",
            start_cell="A1",
            values=[[1, 2], [3, 4]],
            save=False,
        )
        assert result.rows == 2
        assert result.columns == 2
        assert result.start_cell == "A1"
        assert result.sheet == "Sheet1"
        assert result.document == ""
        assert fake_doc.store_count == 0

    def test_write_range_values_persist(self, fake_doc):
        libremcp.write_range(
            sheet="Sheet1",
            start_cell="B2",
            values=[[1, 2], [3, 4]],
            save=False,
        )
        sheet = fake_doc.Sheets.getByName("Sheet1")
        assert sheet.getCellByPosition(1, 1).Value == 1
        assert sheet.getCellByPosition(2, 1).Value == 2
        assert sheet.getCellByPosition(1, 2).Value == 3
        assert sheet.getCellByPosition(2, 2).Value == 4

    def test_write_range_jagged_raises(self, fake_doc):
        with pytest.raises(ValueError):
            libremcp.write_range(
                sheet="Sheet1",
                start_cell="A1",
                values=[["a"], ["b", "c", "d"]],
                save=False,
            )

    def test_write_range_save_calls_store(self, fake_doc):
        libremcp.write_range(
            sheet="Sheet1", start_cell="A1", values=[[1]], save=True
        )
        assert fake_doc.store_count == 1

    def test_write_range_empty_raises(self, fake_doc):
        with pytest.raises(ValueError):
            libremcp.write_range(sheet="Sheet1", start_cell="A1", values=[])


class TestListSheets:
    def test_list_sheets(self, fake_doc):
        result = libremcp.list_sheets()
        assert result == ["Sheet1"]

    def test_list_sheets_after_add(self, fake_doc):
        fake_doc.Sheets.insertNewByName("Sheet2", 1)
        result = libremcp.list_sheets()
        assert result == ["Sheet1", "Sheet2"]


class TestAddSheet:
    def test_add_new_sheet(self, fake_doc):
        result = libremcp.add_sheet(name="Sheet2")
        assert "Sheet1" in result
        assert "Sheet2" in result
        assert len(result) == 2
        assert fake_doc.Sheets._inserts == [("Sheet2", 1)]

    def test_add_existing_sheet_noop(self, fake_doc):
        result = libremcp.add_sheet(name="Sheet1")
        assert result == ["Sheet1"]
        assert fake_doc.Sheets._inserts == []

    def test_add_sheet_with_position(self, fake_doc):
        result = libremcp.add_sheet(name="Sheet2", position=0)
        assert "Sheet2" in result
        assert fake_doc.Sheets._inserts == [("Sheet2", 0)]

    def test_add_multiple_sheets(self, fake_doc):
        libremcp.add_sheet(name="Sheet2")
        libremcp.add_sheet(name="Sheet3")
        result = libremcp.list_sheets()
        assert "Sheet1" in result
        assert "Sheet2" in result
        assert "Sheet3" in result
        assert len(result) == 3
