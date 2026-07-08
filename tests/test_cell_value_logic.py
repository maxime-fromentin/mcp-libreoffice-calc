import pytest
from src import libremcp


class RecordingCell:
    """Fake UNO cell that records every attribute set on it."""

    def __init__(self):
        object.__setattr__(self, "Formula", "")
        object.__setattr__(self, "Value", 0.0)
        object.__setattr__(self, "String", "")
        object.__setattr__(self, "operations", [])

    def __setattr__(self, name, value):
        self.operations.append((name, value))
        object.__setattr__(self, name, value)


class TestSetCellValue:
    def test_formula_string(self):
        cell = RecordingCell()
        libremcp._set_cell_value(cell, "=SUM(A1:A2)")
        assert cell.Formula == "=SUM(A1:A2)"
        assert cell.operations == [("Formula", "=SUM(A1:A2)")]

    def test_int(self):
        cell = RecordingCell()
        libremcp._set_cell_value(cell, 42)
        assert cell.Value == 42
        assert cell.operations == [("Value", 42)]

    def test_float(self):
        cell = RecordingCell()
        libremcp._set_cell_value(cell, 3.14)
        assert cell.Value == 3.14
        assert cell.operations == [("Value", 3.14)]

    def test_true_becomes_one(self):
        cell = RecordingCell()
        libremcp._set_cell_value(cell, True)
        assert cell.Value == 1
        assert cell.operations == [("Value", 1)]

    def test_false_becomes_zero(self):
        cell = RecordingCell()
        libremcp._set_cell_value(cell, False)
        assert cell.Value == 0
        assert cell.operations == [("Value", 0)]

    def test_none_becomes_empty_string(self):
        cell = RecordingCell()
        libremcp._set_cell_value(cell, None)
        assert cell.String == ""
        assert cell.operations == [("String", "")]

    def test_plain_string(self):
        cell = RecordingCell()
        libremcp._set_cell_value(cell, "hello")
        assert cell.String == "hello"
        assert cell.operations == [("String", "hello")]


class TestReadCellValue:
    def test_formula_returns_value(self):
        cell = RecordingCell()
        cell.Formula = "=1+1"
        cell.Value = 2.0
        assert libremcp._read_cell_value(cell) == 2.0

    def test_string_returns_string(self):
        cell = RecordingCell()
        cell.String = "hello"
        assert libremcp._read_cell_value(cell) == "hello"

    def test_empty_falls_back_to_value(self):
        cell = RecordingCell()
        cell.Value = 42.0
        # String is "" by default, Formula is "" by default
        assert libremcp._read_cell_value(cell) == 42.0
