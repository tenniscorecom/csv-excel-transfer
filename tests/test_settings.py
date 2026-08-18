"""tests/test_settings.py — config.ini を読む Settings のテスト"""

import sys
from pathlib import Path

import pytest
from comken import config

from src.exceptions import InvalidHeaderRowError, InvalidOutputPrefixError
from src.settings import (
    DEFAULT_CSV_KEY_COLUMN,
    load_settings,
)


@pytest.fixture
def restore_config_singleton():
    """comken.config の共有状態をテスト間で元に戻す."""
    original = config._singleton
    try:
        yield
    finally:
        config._singleton = original


def _write_config(
    tmp_path: Path,
    *,
    west: str,
    east: str,
    input_xlsx: str,
    sheet: str = "Sheet1",
    key_column: str = "業務用ID",
    header_row: int | str = 1,
    output_prefix: str = "最終_",
    password: str = "",
    csv_key_column: str | None = None,
) -> Path:
    text = "[FILES]\n"
    text += f"WEST_CSV = {west}\n"
    text += f"EAST_CSV = {east}\n"
    text += f"INPUT_XLSX = {input_xlsx}\n\n"
    text += "[EXCEL]\n"
    text += f"SHEET = {sheet}\n"
    text += f"KEY_COLUMN = {key_column}\n"
    text += f"HEADER_ROW = {header_row}\n"
    text += f"OUTPUT_PREFIX = {output_prefix}\n"
    text += f"PASSWORD = {password}\n\n"
    if csv_key_column is not None:
        text += "[CSV]\n"
        text += f"KEY_COLUMN = {csv_key_column}\n\n"
    text += "[転記_MAPPING]\n"
    text += "お名前 = 氏名\n"

    config_path = tmp_path / "config.ini"
    config_path.write_text(text, encoding="utf-8")
    return config_path


def test_load_settings_reads_all_values(tmp_path: Path, restore_config_singleton) -> None:
    config_path = _write_config(
        tmp_path,
        west=str(tmp_path / "west.csv"),
        east=str(tmp_path / "east.csv"),
        input_xlsx=str(tmp_path / "input.xlsx"),
        csv_key_column="お客様ID",
    )
    config.read(config_path)

    settings = load_settings()

    assert settings.west_csv_path == tmp_path / "west.csv"
    assert settings.east_csv_path == tmp_path / "east.csv"
    assert settings.input_xlsx_path == tmp_path / "input.xlsx"
    assert settings.layout.csv_key_column == "お客様ID"
    assert settings.layout.excel_sheet == "Sheet1"
    assert settings.layout.excel_key_column == "業務用ID"
    assert settings.layout.excel_header_row == 1
    assert settings.output_prefix == "最終_"
    assert settings.password == ""
    assert settings.mapping == {"お名前": "氏名"}


def test_load_settings_uses_default_csv_key_column_when_csv_section_missing(
    tmp_path: Path, restore_config_singleton
) -> None:
    # [CSV] セクションを省略（既定の "お客様ID" が使われる）
    config_path = _write_config(
        tmp_path,
        west=str(tmp_path / "west.csv"),
        east=str(tmp_path / "east.csv"),
        input_xlsx=str(tmp_path / "input.xlsx"),
    )
    config.read(config_path)

    settings = load_settings()

    assert settings.layout.csv_key_column == DEFAULT_CSV_KEY_COLUMN


def test_load_settings_rejects_blank_output_prefix(
    tmp_path: Path, restore_config_singleton
) -> None:
    config_path = _write_config(
        tmp_path,
        west=str(tmp_path / "west.csv"),
        east=str(tmp_path / "east.csv"),
        input_xlsx=str(tmp_path / "input.xlsx"),
        output_prefix="   ",
    )
    config.read(config_path)

    with pytest.raises(InvalidOutputPrefixError, match="OUTPUT_PREFIX"):
        load_settings()


@pytest.mark.parametrize("header_row", ["あ", "", 0, -1])
def test_load_settings_rejects_invalid_header_row(
    tmp_path: Path, restore_config_singleton, header_row: int | str
) -> None:
    config_path = _write_config(
        tmp_path,
        west=str(tmp_path / "west.csv"),
        east=str(tmp_path / "east.csv"),
        input_xlsx=str(tmp_path / "input.xlsx"),
        header_row=header_row,
    )
    config.read(config_path)

    with pytest.raises(InvalidHeaderRowError, match="HEADER_ROW"):
        load_settings()


def test_load_settings_returns_mapping_as_configured(
    tmp_path: Path, restore_config_singleton
) -> None:
    config_path = _write_config(
        tmp_path,
        west="w",
        east="e",
        input_xlsx="i",
    )
    config_path.write_text(
        config_path.read_text(encoding="utf-8").replace(
            "[転記_MAPPING]\nお名前 = 氏名\n",
            "[転記_MAPPING]\nお名前 = 氏名\n電話番号 = 電話\n",
        ),
        encoding="utf-8",
    )
    config.read(config_path)

    settings = load_settings()

    assert settings.mapping == {"お名前": "氏名", "電話番号": "電話"}


@pytest.mark.skipif(sys.platform != "win32", reason="Windows のみ検証")
def test_default_csv_key_column_is_顧客_id(tmp_path: Path, restore_config_singleton) -> None:
    """既定のキー列が「お客様ID」であることを確認。"""
    assert DEFAULT_CSV_KEY_COLUMN == "お客様ID"
