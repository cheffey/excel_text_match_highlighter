from typing import Union

import pandas as pd

from match_highlighter.cell_color import CellColor
from match_highlighter.content_matcher import find_overlap_fragments
from match_highlighter.xl_rich_text import ExcelRichText


class ExcelSheet:
    def __init__(self, path, sheet_name=0):
        self._path = path
        self._sheet_name = sheet_name
        self._df = _load_excel(path, sheet_name)
        self._compare_plans = []
        self._formats = {}

    def compare(self, src_column_name: str) -> '_ComparePlan':
        if src_column_name not in self._df.columns:
            raise Exception(f'{src_column_name} NOT found')
        compare_plan = _ComparePlan(self, src_column_name)
        self._compare_plans.append(compare_plan)
        return compare_plan

    def export(self, output_path):
        if not self._compare_plans:
            raise Exception('no comparison plan found')
        writer = pd.ExcelWriter(output_path, engine='xlsxwriter')
        df = self._df
        export_sheet_name = self._sheet_name or "Sheet1"
        df.to_excel(writer, sheet_name=export_sheet_name, index=False)
        workbook = writer.book
        ws = writer.sheets[export_sheet_name]
        center_wrap_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
        column_title_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'text_wrap': True, 'bold': True})

        used_column_length = len(df.columns)
        for plan_idx, compare_plan in enumerate(self._compare_plans):
            if not compare_plan._targets:
                raise Exception(f"""column "{compare_plan._src_column_name}" must compare with other columns using like:
excel_sheet.compare({compare_plan._src_column_name}).with("target_column_name", "#FF0000", 10)""")
            compare_plan._base_column_idx = used_column_length
            ws.write(0, compare_plan._base_column_idx, f'({plan_idx}){compare_plan._src_column_name}', column_title_format)
            for target_idx, (target_column_name, _, _) in enumerate(compare_plan._targets):
                ws.write(0, compare_plan._base_column_idx + target_idx + 1, f'({plan_idx}){target_column_name}', column_title_format)
            used_column_length += len(compare_plan._targets) + 1

        for data_idx in range(len(df)):
            for compare_plan in self._compare_plans:
                src_value = df.loc[data_idx, compare_plan._src_column_name]
                if pd.isna(src_value):
                    continue
                src_rich_text = ExcelRichText(src_value)
                for target_idx, (target_column_name, color, min_fragment_size) in enumerate(compare_plan._targets):
                    color_format = self._to_format(workbook, color)
                    target_value = df.loc[data_idx, target_column_name]
                    overlap_matches = find_overlap_fragments(src_value, target_value, min_fragment_size)
                    src_rich_text.color(overlap_matches, color_format)
                    target_rich_text = ExcelRichText(target_value).color(overlap_matches, color_format)
                    target_rich_text.write_to(ws, data_idx + 1, compare_plan._base_column_idx + target_idx + 1)
                src_rich_text.write_to(ws, data_idx + 1, compare_plan._base_column_idx)

        self._format_cells(ws, center_wrap_format, len(df), used_column_length)
        writer.close()
        print(f'exported to {output_path}')

    def _to_format(self, workbook, color):
        if color not in self._formats:
            self._formats[color] = workbook.add_format({'color': color, 'bold': True, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
        return self._formats[color]

    def _format_cells(self, ws, formation, rows, cols):
        ws.set_column(0, cols, 40, formation)
        for row in range(rows):
            ws.set_row(row, None, formation)


class _ComparePlan:
    def __init__(self, excel_sheet: ExcelSheet, src_column_name):
        self._excel_sheet = excel_sheet
        self._src_column_name = src_column_name
        self._targets = []
        self._base_column_idx = None

    # params: color takes CellColor or hex str like "#FF0000"
    def with_column(self, column_name: str, color: Union[str, CellColor], min_fragment_size: int) -> '_ComparePlan':
        if min_fragment_size < 1:
            raise Exception('min_fragment_size must be >= 1')
        if column_name not in self._excel_sheet._df.columns:
            raise Exception(f'{column_name} NOT found')
        if isinstance(color, CellColor):
            color = color.value
        elif not isinstance(color, str):
            raise Exception('color must be CellColor or str')
        self._targets.append((column_name, color, min_fragment_size))
        return self

    def export(self, output_path):
        self._excel_sheet.export(output_path)


def _load_excel(path, sheet_name=0):
    if path.endswith('.xlsx'):
        return pd.read_excel(path, sheet_name=sheet_name)
    elif path.endswith('.csv'):
        return pd.read_csv(path)
    else:
        raise Exception('path must be .xlsx or .csv')
