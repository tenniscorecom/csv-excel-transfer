"""csv-excel-transfer の業務フロー。"""

from collections.abc import Mapping
from dataclasses import dataclass
from pathlib import Path
from typing import cast

from comken import Config
from comken.core import delete_files
from comken.exceptions import (
    ComkenError,
    CsvNoDataRowsError,
    CsvRowDuplicateKeyError,
    ExcelColumnNotFoundError,
    TransferSourceColumnNotFoundError,
)
from comken.exceptions.table import TransferMappingError
from comken.toolbox import Transfer
from comken.toolbox.csv import CsvReader
from comken.toolbox.excel import ExcelWriter

CSV_KEY_COLUMN = "お客様ID"
EXCEL_KEY_COLUMN = "業務用ID"
SHEET_NAME = "Sheet1"
HEADER_ROW = 1


class InputExcelNoDataError(ComkenError):
    """入力 Excel に転記対象のデータ行がない。"""

    def __init__(self, path: Path) -> None:
        super().__init__(f"INPUT Excel にデータ行がありません: {path}")


@dataclass(frozen=True)
class TransferResult:
    """1回の実行結果。"""

    output_path: Path
    matched_rows: int


def validate_config(settings: Config) -> None:
    """必須設定を処理開始前にすべて参照して検証する。"""
    settings.FILES.OUTPUT_EXCEL_FOLDER
    settings.FILES.INPUT_EXCEL_FOLDER
    settings.FILES.INPUT_CSV_FOLDER
    settings.EXCEL.INPUT_NAME
    settings.EXCEL.OUTPUT_PREFIX
    settings.EXCEL.READ_PASSWORD
    settings.EXCEL.WRITE_PASSWORD
    settings.CSV.WEST
    settings.CSV.EAST
    if not settings.TRANSFER_MAPPING:
        raise TransferMappingError


def _paths(settings: Config) -> tuple[Path, Path, Path, Path]:
    csv_folder = Path(settings.FILES.INPUT_CSV_FOLDER)
    excel_folder = Path(settings.FILES.INPUT_EXCEL_FOLDER)
    output_folder = Path(settings.FILES.OUTPUT_EXCEL_FOLDER)
    input_name = str(settings.EXCEL.INPUT_NAME)
    output_prefix = str(settings.EXCEL.OUTPUT_PREFIX)
    input_excel = excel_folder / input_name
    output_excel = output_folder / f"{output_prefix}{input_name}"
    west_csv = csv_folder / str(settings.CSV.WEST)
    east_csv = csv_folder / str(settings.CSV.EAST)
    input_paths = [west_csv, east_csv, input_excel]
    resolved_inputs = [path.resolve() for path in input_paths]
    if len(set(resolved_inputs)) != len(resolved_inputs):
        raise ComkenError("入力 CSV 2本と入力 Excel には、それぞれ別のファイルを指定してください。")
    if output_excel.resolve() in resolved_inputs:
        raise ComkenError(
            "出力先が入力ファイルと同じです。出力フォルダまたは OUTPUT_PREFIX を変更してください。"
        )
    return west_csv, east_csv, input_excel, output_excel


def _merge_csv(paths: tuple[Path, Path], source_columns: list[str]) -> dict[str, dict[str, str]]:
    lookup: dict[str, dict[str, str]] = {}
    duplicate_counts: dict[str, int] = {}
    for path in paths:
        reader = CsvReader(path)
        rows = reader.read_rows([CSV_KEY_COLUMN, *source_columns])
        if not rows:
            raise CsvNoDataRowsError(path)
        indexed = reader.index(CSV_KEY_COLUMN)
        for key, row in indexed.items():
            if key in lookup:
                duplicate_counts[key] = duplicate_counts.get(key, 1) + 1
            else:
                lookup[key] = row
    if duplicate_counts:
        raise CsvRowDuplicateKeyError(
            CSV_KEY_COLUMN,
            duplicate_counts,
            ", ".join(str(path) for path in paths),
        )
    return lookup


def run(settings: Config | None = None) -> TransferResult:
    """CSV を統合して Excel へ転記し、保存成功後に入力を削除する。"""
    actual_settings = settings or Config()
    validate_config(actual_settings)
    west_csv, east_csv, input_excel, output_excel = _paths(actual_settings)
    configured_mapping = cast(Mapping[str, str], actual_settings.TRANSFER_MAPPING)
    source_columns = list(configured_mapping)
    lookup = _merge_csv((west_csv, east_csv), source_columns)

    matched_rows = 0
    with ExcelWriter(input_excel) as source_book:
        source_sheet = source_book.sheet(SHEET_NAME)
        input_rows = source_sheet.read_rows_as_dicts(HEADER_ROW)
        if not input_rows:
            raise InputExcelNoDataError(input_excel)
        headers = list(input_rows[0])
        required_excel = [EXCEL_KEY_COLUMN, *configured_mapping.values()]
        missing_excel = [column for column in required_excel if column not in headers]
        if missing_excel:
            raise ExcelColumnNotFoundError(missing_excel)

        existing_csv = list(next(iter(lookup.values())))
        missing_csv = [column for column in source_columns if column not in existing_csv]
        if missing_csv:
            raise TransferSourceColumnNotFoundError(missing_csv, existing_csv)

        identity_mapping = {header: header for header in headers}
        with ExcelWriter.create(output_excel, sheet_name=SHEET_NAME) as destination_book:
            destination_sheet = destination_book.sheet(SHEET_NAME)
            transfer = Transfer(
                source_sheet,
                destination_sheet,
                identity_mapping,
            )

            def transform(source: dict[str, object]) -> dict[str, object]:
                nonlocal matched_rows
                customer = lookup.get(str(source.get(EXCEL_KEY_COLUMN, "")))
                if customer is None:
                    return source
                matched_rows += 1
                for source_column, destination_column in configured_mapping.items():
                    source[destination_column] = customer[source_column]
                return source

            transfer.run(transform=transform)
            destination_book.save(
                output_excel,
                read_pw=str(actual_settings.EXCEL.READ_PASSWORD),
                write_pw=str(actual_settings.EXCEL.WRITE_PASSWORD),
            )

    delete_files([west_csv, east_csv, input_excel], missing_ok=True)
    return TransferResult(output_path=output_excel, matched_rows=matched_rows)


__all__ = ["TransferResult", "run", "validate_config"]
