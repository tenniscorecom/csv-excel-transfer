"""csv-excel-transfer の業務フロー。"""

import logging
from collections.abc import Mapping
from pathlib import Path
from typing import cast

from comken import config
from comken.core import delete_files
from comken.core.table.model import Table
from comken.core.table.transfer import Transfer
from comken.exceptions import (
    ComkenError,
    ExcelColumnNotFoundError,
    TransferSourceColumnNotFoundError,
)
from comken.toolbox.csv import CSV
from comken.toolbox.excel import Excel
from comken.toolbox.windows import ExcelCOMHandler

CSV_KEY_COLUMN = "お客様ID"
EXCEL_KEY_COLUMN = "業務用ID"
SHEET_NAME = "Sheet1"
HEADER_ROW = 1

logger = logging.getLogger(__name__)


def _paths() -> tuple[Path, Path, Path, Path]:
    csv_folder = config.FILES.INPUT_CSV_FOLDER
    excel_folder = config.FILES.INPUT_EXCEL_FOLDER
    output_folder = config.FILES.OUTPUT_EXCEL_FOLDER
    input_excel = excel_folder / config.EXCEL.INPUT_NAME
    output_excel = output_folder / f"{config.EXCEL.OUTPUT_PREFIX}{config.EXCEL.INPUT_NAME}"
    west_csv = csv_folder / config.CSV.WEST
    east_csv = csv_folder / config.CSV.EAST
    input_paths = [west_csv, east_csv, input_excel]
    resolved_inputs = [path.resolve() for path in input_paths]
    if len(set(resolved_inputs)) != len(resolved_inputs):
        raise ComkenError("入力 CSV 2本と入力 Excel には、それぞれ別のファイルを指定してください。")
    if output_excel.resolve() in resolved_inputs:
        raise ComkenError(
            "出力先が入力ファイルと同じです。出力フォルダまたは OUTPUT_PREFIX を変更してください。"
        )
    return west_csv, east_csv, input_excel, output_excel


def _merge_csv(paths: tuple[Path, Path], source_columns: list[str]) -> Table:
    """2 つの CSV を読み込み、``CSV_KEY_COLUMN`` で突合した Table を返す。

    1 ファイル内の重複は ``Table.index()`` の ``TableDuplicateKeyError`` に任せる
    （旧実装の「踏んで上書き」は採用しない）。
    2 ファイル間の重複は ``ComkenError`` として送出する。
    """
    lookup: dict[str, dict[str, str]] = {}
    duplicate_counts: dict[str, int] = {}
    read_columns = [CSV_KEY_COLUMN, *source_columns]

    for path in paths:
        with CSV(path, read_only=True) as csv_file:
            table = csv_file.read()
        if len(table) == 0:
            raise ComkenError(f"CSV にデータ行がありません: {path}")
        indexed = table.index(CSV_KEY_COLUMN)
        for key, row in indexed.items():
            if key in lookup:
                duplicate_counts[key] = duplicate_counts.get(key, 1) + 1
            else:
                lookup[key] = row
    if duplicate_counts:
        keys = ", ".join(f"{key}({count} 件)" for key, count in duplicate_counts.items())
        raise ComkenError(
            f"CSV のキー列「{CSV_KEY_COLUMN}」に 2 ファイル間で重複があります: {keys}\n"
            f"対象ファイル: {', '.join(str(path) for path in paths)}"
        )

    return Table(
        read_columns,
        [{column: row.get(column, "") for column in read_columns} for row in lookup.values()],
    )


def run() -> Path:
    """CSV を統合して Excel へ転記し、保存成功後に入力を削除する。"""
    west_csv, east_csv, input_excel, output_excel = _paths()
    configured_mapping = cast(Mapping[str, str], config.TRANSFER_MAPPING)
    source_columns = list(configured_mapping)

    read_password = str(config.EXCEL.READ_PASSWORD)
    write_password = str(config.EXCEL.WRITE_PASSWORD)

    read_table = _merge_csv((west_csv, east_csv), source_columns)

    with Excel(input_excel, read_only=True) as source_book:
        input_rows = source_book.read_computed_rows_as_dicts(
            SHEET_NAME, header_row=HEADER_ROW
        )
        if not input_rows:
            raise ComkenError(f"INPUT Excel にデータ行がありません: {input_excel}")
        headers = list(input_rows[0])
        required_excel = [EXCEL_KEY_COLUMN, *configured_mapping.values()]
        missing_excel = [column for column in required_excel if column not in headers]
        if missing_excel:
            raise ExcelColumnNotFoundError(missing_excel)

        existing_csv = list(read_table.columns)
        missing_csv = [column for column in source_columns if column not in existing_csv]
        if missing_csv:
            raise TransferSourceColumnNotFoundError(missing_csv, existing_csv)

        # Excel の見出し行と一致する write Table を作る。
        # matched_rows() はキーが一致する行だけを返すため、
        # 一致しなかった Excel 行は write_table の初期値のまま残り、
        # result() にそのまま反映される。
        write_table = Table(
            headers,
            [{header: row.get(header, "") for header in headers} for row in input_rows],
        )
        transfer = Transfer(
            read_table,
            write_table,
            configured_mapping,
            read_key=CSV_KEY_COLUMN,
            write_key=EXCEL_KEY_COLUMN,
        )
        matched_rows = 0
        for read_row, write_row in transfer.matched_rows():
            matched_rows += 1
            transfer.apply_mapping(read_row, write_row)
        result_table = transfer.result()

    # Excel へ書き出す。openpyxl が with を抜けた時に保存する。
    column_count = len(result_table.columns)
    row_count = len(result_table) + 1  # +1 for header
    with Excel(output_excel) as destination:
        sheet = destination.sheet(SHEET_NAME)
        if column_count > 0:
            end_column = chr(ord("A") + column_count - 1)
            values = [
                list(result_table.columns),
                *[list(row.values()) for row in result_table.read()],
            ]
            sheet.write_range(f"A1:{end_column}{row_count}", values)

    # パスワード保存は openpyxl ではできないため、COM で別名保存する。
    # 両方とも空欄なら openpyxl だけで完了しているので COM を起動しない。
    if read_password or write_password:
        with ExcelCOMHandler(output_excel) as excel_com:
            excel_com.save_as(output_excel, read_pw=read_password, write_pw=write_password)

    logger.info("転記件数: %d", matched_rows)
    delete_files([west_csv, east_csv, input_excel], missing_ok=True)
    return output_excel


__all__ = ["run"]
