## Introduction
This is a Python program that compares two Excel column and highlight the matching keywords.

## Example
before highlight:

![before highlight](https://github.com/cheffey/excel_text_match_highlighter/blob/main/img/before_highlight.jpg?raw=true)

after highlight:

![after highlight](https://github.com/cheffey/excel_text_match_highlighter/blob/main/img/after_highlight.jpg?raw=true)

## How to use
```
pip install git+https://github.com/cheffey/excel_text_match_highlighter
```
```
from match_highlighter import CellColor
from match_highlighter import ExcelSheet

sheet = ExcelSheet('example.csv')
sheet.compare('col0').with_column('col1', '#FF0000', 10).with_column('col2', '#00FF00', 2)
sheet.compare('match0').with_column('match1', CellColor.purple, 10)
sheet.export('example.xlsx')
```