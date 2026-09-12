# Martes

Read an Excel workbook in Pandas and reference cells by coordinates

## Install

```sh
uv add git+https://github.com/pashri/martes
```

```sh
pip install git+https://github.com/pashri/martes
```

Requires Python 3.11 or newer.

## Documentation

This is it for now.

## Overview

Martes is an extension onto pandas meant to allow access to DataFrame slices by their Excel Coordinates. It's good for reading in data from Excel. It's not good for saving data to Excel.

### How to use it

You can load the extension like this:

```python
from martes import PandasWorkbook
```

You can load your workbook like this:

```python
workbook = PandasWorkbook(filepath)
```

See your worksheets like this:

```python
print(workbook.sheets)  # Returns a list of sheet names
```

Access a sheet like this:

```python
addresses_df = workbook['Addresses']
```

And access data from the worksheet like this:

```python
workbook['Addresses'].xl['A2']
```

```python
workbook['Addresses'].xl['A2:G50']
```

```python
workbook['Addresses'].xl['A:A']
```

```python
workbook['Addresses!A2']
```

```python
workbook['Addresses!A2:G50']
```

```python
workbook['Addresses!A:A']
```

It will always return a DataFrame for a range, and a single cell's value for a cell.

### Sheet names

Sheet names work quoted or unquoted, so a reference copied straight out of Excel is understood as-is. Inside a quoted name, a doubled apostrophe means a literal one.

```python
workbook['My Sheet!A2']
workbook["'My Sheet'!A2"]
workbook["'Q1!Draft'!A2"]     # a name containing an exclamation mark
workbook["'Bob''s Data'!A2"]  # a name containing an apostrophe
```

Asking for a sheet that isn't there tells you which ones are.

### Reading past the end

Excel sheets have blank cells, and pandas drops entirely blank rows and columns when it reads a file. So a coordinate beyond the edge of a sheet reads as `NaN` rather than raising, exactly like a blank cell inside it.

```python
workbook['Addresses'].xl['ZZ999']  # nan
```

Reading never changes the DataFrame you read from. Writing past the edge does grow it, because it has to:

```python
df.xl['E5'] = 99  # df now reaches E5, with NaN in between
```

### Type checking

`df.xl` is registered with pandas at runtime, so type checkers can't see it and will reject `df.xl()`. Use `xl()` instead wherever you need the code to check:

```python
from martes import xl

xl(df)['A2']
xl(df)()  # rename the axes to Excel coordinates
```

Both reach the same accessor. `df.xl` is the nicer one to type in a notebook.

### Why _Martes?_

It's Tuesday

### Contributing

You can contribute with pull requests. Add tests for your new
functionality, and check it before opening one:

```sh
uv sync
uv run pytest
uv run isort --check-only src tests
uv run mypy
uv run pylint src tests
uv run numpydoc lint src/martes/*.py
```
