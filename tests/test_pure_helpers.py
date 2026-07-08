import pytest
from src import libremcp


class TestCellIndexes:
    def test_basic(self):
        assert libremcp._cell_indexes("A1") == (0, 0)
        assert libremcp._cell_indexes("B3") == (1, 2)
        assert libremcp._cell_indexes("Z1") == (25, 0)
        assert libremcp._cell_indexes("A10") == (0, 9)

    def test_multi_letter(self):
        assert libremcp._cell_indexes("AA1") == (26, 0)
        assert libremcp._cell_indexes("AZ1") == (51, 0)
        assert libremcp._cell_indexes("BA1") == (52, 0)
        assert libremcp._cell_indexes("ZZ1") == (701, 0)

    def test_case_insensitive(self):
        assert libremcp._cell_indexes("a1") == (0, 0)
        assert libremcp._cell_indexes("b3") == (1, 2)
        assert libremcp._cell_indexes("ab12") == libremcp._cell_indexes("AB12")
        assert libremcp._cell_indexes("zz99") == libremcp._cell_indexes("ZZ99")

    def test_whitespace(self):
        assert libremcp._cell_indexes("  A1  ") == (0, 0)
        assert libremcp._cell_indexes("\tB2\n") == (1, 1)
        assert libremcp._cell_indexes("   C5   ") == (2, 4)

    @pytest.mark.parametrize("bad", ["A0", "1A", "", "A"])
    def test_invalid_raises(self, bad):
        with pytest.raises(ValueError):
            libremcp._cell_indexes(bad)


class TestCellName:
    def test_basic(self):
        assert libremcp._cell_name(0, 0) == "A1"
        assert libremcp._cell_name(1, 2) == "B3"
        assert libremcp._cell_name(25, 0) == "Z1"
        assert libremcp._cell_name(0, 9) == "A10"

    def test_multi_letter(self):
        assert libremcp._cell_name(26, 0) == "AA1"
        assert libremcp._cell_name(51, 0) == "AZ1"
        assert libremcp._cell_name(52, 0) == "BA1"
        assert libremcp._cell_name(701, 0) == "ZZ1"

    @pytest.mark.parametrize("cell", ["A1", "Z99", "AA1", "ZZ1"])
    def test_round_trip(self, cell):
        col, row = libremcp._cell_indexes(cell)
        assert libremcp._cell_name(col, row) == cell


class TestRangeIndexes:
    def test_single_cell(self):
        assert libremcp._range_indexes("A1") == (0, 0, 0, 0)
        assert libremcp._range_indexes("B3") == (1, 2, 1, 2)
        assert libremcp._range_indexes("Z99") == (25, 98, 25, 98)

    def test_normal(self):
        assert libremcp._range_indexes("A1:B2") == (0, 0, 1, 1)
        assert libremcp._range_indexes("A1:C3") == (0, 0, 2, 2)
        assert libremcp._range_indexes("B2:D4") == (1, 1, 3, 3)

    def test_reversed_normalizes(self):
        assert libremcp._range_indexes("C3:A1") == (0, 0, 2, 2)
        assert libremcp._range_indexes("D4:A1") == (0, 0, 3, 3)

    def test_spaces(self):
        assert libremcp._range_indexes("  A1 : B2  ") == (0, 0, 1, 1)
        assert libremcp._range_indexes("C3 : A1") == (0, 0, 2, 2)

    def test_three_parts_raises(self):
        with pytest.raises(ValueError):
            libremcp._range_indexes("A1:B2:C3")


class TestEscapeFormulaText:
    def test_plain(self):
        assert libremcp._escape_formula_text("hello") == "hello"
        assert libremcp._escape_formula_text("hello world") == "hello world"

    def test_quotes(self):
        assert libremcp._escape_formula_text('say "hi"') == 'say ""hi""'
        assert libremcp._escape_formula_text('a"b"c') == 'a""b""c'

    def test_empty(self):
        assert libremcp._escape_formula_text("") == ""
