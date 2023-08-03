from match_highlighter import CellColor
from match_highlighter import ExcelSheet

sheet = ExcelSheet('example.csv')
sheet.compare('col0').with_column('col1', '#FF0000', 10).with_column('col2', '#00FF00', 2)
sheet.compare('match0').with_column('match1', CellColor.purple, 10)
sheet.export('example.xlsx')
