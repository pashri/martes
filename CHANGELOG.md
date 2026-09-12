# Changelog

## 1.0.0

### Removed

- Support for Python 3.9 and 3.10, both past end of life. Requires 3.11 or newer.
- Support for pandas 1.x, which does not import against current NumPy. Requires pandas 2.2.2 or newer.

### Changed

- Reading a coordinate past the edge of a frame returns `NaN` instead of raising `KeyError`, and no longer modifies the frame being read. It previously grew the frame in place, which changed its dtypes: an `int64` column became `float64`, so an unrelated cell that had read as `1` afterwards read as `1.0`. Writing past the edge still grows the frame.
- An unknown sheet name now names the sheets the workbook does have.

### Added

- `xl(frame)`, a module-level equivalent of the `.xl` accessor that type checkers can see.
- Excel-style quoted sheet references, including names containing `!` or an apostrophe.
- `Coordinate.from_cell()`, `Coordinate.split_parts()`, `Coordinate.first_digit()`, `Coordinate.iter_letters()`, `Coordinate.fold_letter()`, `ExcelCoordAccessor.padded()`, `ExcelCoordAccessor.padded_index()`, `ExcelCoordAccessor.padded_columns()`, `ExcelCoordAccessor.expand_columns()`, `ExcelCoordAccessor.expand_rows()`, `PandasWorkbook.sheet()`, `PandasWorkbook.read_sheet()` and `PandasWorkbook.split_reference()`.
- Type annotations throughout, and a `py.typed` marker.
- A test suite, and CI across Python 3.11 to 3.14.

### Fixed

- `expand()` was applied through `DataFrame.pipe()`, which passes a copy on current pandas, so the frame never actually grew and reading past the edge raised `KeyError`.
- `expand()` began its column range at the last existing column and overwrote it with `NaN`.
- A coordinate with no row, such as `'E:E'`, was compared against an integer and raised `TypeError`.
- `_range()` referenced an unbound local when a coordinate was neither a cell nor a range.

## 0.1.0

Initial release.
