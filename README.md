# Martes

Read an Excel workbook in Pandas and reference cells by coordinates

## Install

```sh
uv add git+https://github.com/pashri/martes
```

```sh
pip install git+https://github.com/pashri/martes
```

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
