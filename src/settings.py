"""src/settings.py — config.ini の値を読み、処理で使う形に整える

config.ini を読むのはこのファイルだけにする。列名や保存先の組み立てを変えたくなったとき、
直す場所がここに集まる。
"""

from dataclasses import dataclass
from pathlib import Path

from comken import config


@dataclass(frozen=True)
class Settings:
    """このツールが使う設定一式。"""

    west_csv_path: Path
    east_csv_path: Path
    input_xlsx_path: Path
    csv_key_column: str
    excel_sheet: str
    excel_key_column: str
    excel_header_row: int
    mapping: dict[str, str]
    output_prefix: str
    password: str


def load_settings() -> Settings:
    """config.ini を読んで Settings に詰める。

    Raises:
        comken.exceptions.ConfigInvalidValueError: HEADER_ROW が1以上の整数でない、
            OUTPUT_PREFIX が空欄（前後の空白のみ）の場合。
        comken.exceptions.ConfigSectionNotFoundError / ConfigKeyNotFoundError:
            必要なセクションやキーが無い場合。
    """
    # 値は config.ini.example でサンプルを書いている。実際の列名に書き換えて使う
    mapping = config.mapping("転記_MAPPING")
    return Settings(
        west_csv_path=Path(str(config.FILES.WEST_CSV)),
        east_csv_path=Path(str(config.FILES.EAST_CSV)),
        input_xlsx_path=Path(str(config.FILES.INPUT_XLSX)),
        csv_key_column=str(config.CSV.KEY_COLUMN),
        excel_sheet=str(config.EXCEL.SHEET),
        excel_key_column=str(config.EXCEL.KEY_COLUMN),
        excel_header_row=config.int_value("EXCEL.HEADER_ROW", minimum=1),
        mapping=mapping,
        output_prefix=config.text("EXCEL.OUTPUT_PREFIX"),
        password=str(config.EXCEL.PASSWORD),
    )
