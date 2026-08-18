"""
src/settings.py — config.ini の値を読み、処理で使う形に整える

config.ini を読むのはこのファイルだけにする。列名や保存先の組み立てを変えたくなったとき、
直す場所がここに集まる。
"""

import logging
from dataclasses import dataclass
from pathlib import Path

from comken import config
from comken.exceptions import ConfigSectionNotFoundError

from src.exceptions import InvalidHeaderRowError, InvalidOutputPrefixError

logger = logging.getLogger(__name__)


# config.ini のセクション名。1つの処理で複数のセクションへ値を置くため、ここで名前を固定する
SECTION_FILES = "FILES"
SECTION_CSV = "CSV"
SECTION_EXCEL = "EXCEL"
SECTION_MAPPING = "転記_MAPPING"

DEFAULT_CSV_KEY_COLUMN = "お客様ID"


@dataclass(frozen=True)
class SourceLayout:
    """CSV・Excel それぞれのキー列と、転記に使うシート／ヘッダー行。"""

    csv_key_column: str
    excel_sheet: str
    excel_key_column: str
    excel_header_row: int


@dataclass(frozen=True)
class Settings:
    """このツールが使う設定一式。"""

    west_csv_path: Path
    east_csv_path: Path
    input_xlsx_path: Path
    layout: SourceLayout
    mapping: dict[str, str]
    output_prefix: str
    password: str


def load_settings() -> Settings:
    """config.ini を読んで Settings に詰める。

    Raises:
        InvalidHeaderRowError: [EXCEL] HEADER_ROW が1以上の整数でない場合。
        InvalidOutputPrefixError: [EXCEL] OUTPUT_PREFIX が空欄の場合。
        comken.exceptions.ConfigSectionNotFoundError:
            [FILES] / [EXCEL] / [転記_MAPPING] などの必要なセクションがない場合。
    """
    layout = SourceLayout(
        csv_key_column=_csv_key_column(),
        excel_sheet=str(config.EXCEL.SHEET),
        excel_key_column=str(config.EXCEL.KEY_COLUMN),
        excel_header_row=_excel_header_row(),
    )
    # 値は config.ini.example でサンプルを書いている。実際の列名に書き換えて使う
    mapping = config.mapping(SECTION_MAPPING)
    return Settings(
        west_csv_path=Path(config.FILES.WEST_CSV),
        east_csv_path=Path(config.FILES.EAST_CSV),
        input_xlsx_path=Path(config.FILES.INPUT_XLSX),
        layout=layout,
        mapping=mapping,
        output_prefix=_output_prefix(),
        password=str(config.EXCEL.PASSWORD),
    )


def _csv_key_column() -> str:
    """[CSV] KEY_COLUMN を読む。セクションが無い・空なら既定値。"""
    try:
        value = config.CSV.KEY_COLUMN
    except ConfigSectionNotFoundError:
        return DEFAULT_CSV_KEY_COLUMN
    return str(value) or DEFAULT_CSV_KEY_COLUMN


def _excel_header_row() -> int:
    """[EXCEL] HEADER_ROW を読む。1以上の整数。"""
    value = str(config.EXCEL.HEADER_ROW)
    try:
        header_row = int(value)
    except ValueError as error:
        raise InvalidHeaderRowError(value) from error
    if header_row < 1:
        raise InvalidHeaderRowError(value)
    return header_row


def _output_prefix() -> str:
    """[EXCEL] OUTPUT_PREFIX を読む。前後の空白だけなら空欄として扱う。"""
    value = str(config.EXCEL.OUTPUT_PREFIX)
    if not value.strip():
        raise InvalidOutputPrefixError()
    return value
