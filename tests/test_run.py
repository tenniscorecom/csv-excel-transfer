"""tests/test_run.py — run.py のオーケストレーション全体"""

from pathlib import Path

import pytest
from comken import config, dry_run
from comken.exceptions import (
    CsvRowDuplicateKeyError,
    TransferDestinationColumnNotFoundError,
)
from openpyxl import load_workbook

from src.run import run
from tests.conftest import SAMPLE_MAPPING


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
    csv_key_column: str = "お客様ID",
    header_row: int | str = 1,
    output_prefix: str = "最終_",
    password: str = "",
    mapping: dict[str, str] | None = None,
) -> Path:
    if mapping is None:
        mapping = SAMPLE_MAPPING
    text = "[FILES]\n"
    text += f"WEST_CSV = {west}\n"
    text += f"EAST_CSV = {east}\n"
    text += f"INPUT_XLSX = {input_xlsx}\n\n"
    text += "[CSV]\n"
    text += f"KEY_COLUMN = {csv_key_column}\n\n"
    text += "[EXCEL]\n"
    text += f"SHEET = {sheet}\n"
    text += f"KEY_COLUMN = {key_column}\n"
    text += f"HEADER_ROW = {header_row}\n"
    text += f"OUTPUT_PREFIX = {output_prefix}\n"
    text += f"PASSWORD = {password}\n\n"
    text += "[転記_MAPPING]\n"
    for src_col, dst_col in mapping.items():
        text += f"{src_col} = {dst_col}\n"

    config_path = tmp_path / "config.ini"
    config_path.write_text(text, encoding="utf-8")
    return config_path


# ── 西と東のマージ結果が転記に反映される（test_csv_lookup / test_excel_writer の合流観点） ──


def test_run_creates_final_file_and_deletes_source_files(
    tmp_path: Path, make_csv, make_input_book, restore_config_singleton
) -> None:
    west = make_csv(
        tmp_path / "west.csv",
        ["お客様ID", "お名前", "ご住所", "電話番号"],
        [("C001", "山田一郎", "大阪", "06-0000-0001")],
    )
    east = make_csv(
        tmp_path / "east.csv",
        ["お客様ID", "お名前", "ご住所", "電話番号"],
        [("C002", "鈴木三郎", "東京", "03-0000-0002")],
    )
    input_xlsx = make_input_book(
        tmp_path / "作業対象.xlsx",
        [("C001", "", "", ""), ("C002", "", "", "")],
    )
    expected_output = tmp_path / "最終_作業対象.xlsx"

    config_path = _write_config(
        tmp_path, west=str(west), east=str(east), input_xlsx=str(input_xlsx)
    )
    config.read(config_path)

    result = run()

    # 西と東のキーが違うので両方が lookup に乗り、転記件数も反映される
    assert result.output_path == expected_output
    assert result.matched_rows == 2
    assert expected_output.exists()
    # 元の3ファイルは削除済み
    assert not west.exists()
    assert not east.exists()
    assert not input_xlsx.exists()


def test_run_writes_transferred_values_into_output_book(
    tmp_path: Path, make_csv, make_input_book, restore_config_singleton
) -> None:
    """転記結果が workbook のセルに実際に書き込まれていること（値レベルの検証）。"""
    west = make_csv(
        tmp_path / "west.csv",
        ["お客様ID", "お名前", "ご住所", "電話番号"],
        [("C003", "鈴木三郎", "東京", "03-0000-0003")],
    )
    east = make_csv(
        tmp_path / "east.csv",
        ["お客様ID", "お名前"],
        [],
    )
    input_xlsx = make_input_book(
        tmp_path / "作業対象.xlsx",
        [("C003", "", "", "")],
    )

    config_path = _write_config(
        tmp_path, west=str(west), east=str(east), input_xlsx=str(input_xlsx)
    )
    config.read(config_path)

    run()

    expected_output = tmp_path / "最終_作業対象.xlsx"
    workbook = load_workbook(expected_output)
    sheet = workbook["Sheet1"]
    # 2行目がデータ行（C003 = 鈴木三郎）
    assert sheet["B2"].value == "鈴木三郎"
    assert sheet["C2"].value == "東京"
    assert sheet["D2"].value == "03-0000-0003"
    workbook.close()


def test_run_skips_rows_whose_key_is_not_in_lookup(
    tmp_path: Path, make_csv, make_input_book, restore_config_singleton
) -> None:
    """lookup に無いキー（UNKNOWN）はスキップされ、既存のセルはそのまま。"""
    west = make_csv(
        tmp_path / "west.csv",
        ["お客様ID", "お名前", "ご住所", "電話番号"],
        [("C001", "山田一郎", "大阪", "06-0000-0001")],
    )
    east = make_csv(
        tmp_path / "east.csv",
        ["お客様ID", "お名前"],
        [],
    )
    input_xlsx = make_input_book(
        tmp_path / "作業対象.xlsx",
        [("C001", "", "", ""), ("UNKNOWN", "", "", "")],
    )

    config_path = _write_config(
        tmp_path, west=str(west), east=str(east), input_xlsx=str(input_xlsx)
    )
    config.read(config_path)

    result = run()

    # C001 だけがヒットする
    assert result.matched_rows == 1
    expected_output = tmp_path / "最終_作業対象.xlsx"
    workbook = load_workbook(expected_output)
    sheet = workbook["Sheet1"]
    assert sheet["B2"].value == "山田一郎"
    # UNKNOWN 行は touch されない（既存の None のまま）
    assert sheet["B3"].value is None
    workbook.close()


# ── 跨ぎ重複で止まる（test_csv_lookup の例外観点） ──


def test_run_raises_csv_row_duplicate_key_error_when_same_id_appears_in_both(
    tmp_path: Path, make_csv, make_input_book, restore_config_singleton
) -> None:
    west = make_csv(
        tmp_path / "west.csv",
        ["お客様ID", "お名前"],
        [("X1", "西の顧客")],
    )
    east = make_csv(
        tmp_path / "east.csv",
        ["お客様ID", "お名前"],
        [("X1", "東の顧客"), ("X2", "東の別顧客")],
    )
    input_xlsx = make_input_book(
        tmp_path / "作業対象.xlsx",
        [("X1", "", "", "")],
    )

    config_path = _write_config(
        tmp_path, west=str(west), east=str(east), input_xlsx=str(input_xlsx)
    )
    config.read(config_path)

    # comken の CsvRowDuplicateKeyError で停止（index_files が直接投げる）
    with pytest.raises(CsvRowDuplicateKeyError):
        run()

    # マージで止まるので、元3ファイル・最終ファイルとも残らない/作られない
    assert west.exists()
    assert east.exists()
    assert input_xlsx.exists()
    assert not (tmp_path / "最終_作業対象.xlsx").exists()


# ── DRY-RUN / 失敗時の振る舞い ──


def test_run_dry_run_does_not_create_or_delete_any_files(
    tmp_path: Path, make_csv, make_input_book, restore_config_singleton
) -> None:
    west = make_csv(
        tmp_path / "west.csv",
        ["お客様ID", "お名前", "ご住所", "電話番号"],
        [("C001", "山田一郎", "大阪", "06-0000-0001")],
    )
    east = make_csv(
        tmp_path / "east.csv",
        ["お客様ID", "お名前", "ご住所", "電話番号"],
        [("C002", "鈴木三郎", "東京", "03-0000-0002")],
    )
    input_xlsx = make_input_book(
        tmp_path / "作業対象.xlsx",
        [("C001", "", "", ""), ("C002", "", "", "")],
    )
    expected_output = tmp_path / "最終_作業対象.xlsx"

    config_path = _write_config(
        tmp_path, west=str(west), east=str(east), input_xlsx=str(input_xlsx)
    )
    config.read(config_path)

    with dry_run():
        run()

    # DRY-RUN なので、ファイルは何も作られず・消されず
    assert not expected_output.exists()
    assert west.exists()
    assert east.exists()
    assert input_xlsx.exists()


def test_run_dry_run_with_password_does_not_create_or_delete_any_files(
    tmp_path: Path, make_csv, make_input_book, restore_config_singleton
) -> None:
    """dry-run でパスワード付き経路を動かしても、COM を起動せずファイルが増減しない。"""
    west = make_csv(
        tmp_path / "west.csv",
        ["お客様ID", "お名前", "ご住所", "電話番号"],
        [("C001", "山田一郎", "大阪", "06-0000-0001")],
    )
    east = make_csv(
        tmp_path / "east.csv",
        ["お客様ID", "お名前", "ご住所", "電話番号"],
        [("C002", "鈴木三郎", "東京", "03-0000-0002")],
    )
    input_xlsx = make_input_book(
        tmp_path / "作業対象.xlsx",
        [("C001", "", "", ""), ("C002", "", "", "")],
    )

    config_path = _write_config(
        tmp_path,
        west=str(west),
        east=str(east),
        input_xlsx=str(input_xlsx),
        password="dummy-pw",
    )
    config.read(config_path)

    with dry_run():
        run()

    # DRY-RUN なのでファイルは作られず・消されず
    assert not (tmp_path / "最終_作業対象.xlsx").exists()
    assert west.exists()
    assert east.exists()
    assert input_xlsx.exists()


def test_run_does_not_delete_when_transfer_fails(
    tmp_path: Path, make_csv, make_input_book, restore_config_singleton
) -> None:
    """転記中に失敗したら、元ファイルは消さない（再実行できるようにする）。"""
    west = make_csv(
        tmp_path / "west.csv",
        ["お客様ID", "お名前"],
        [("C001", "山田一郎")],
    )
    east = make_csv(
        tmp_path / "east.csv",
        ["お客様ID", "お名前"],
        [],
    )
    # INPUT エクセルの見出しに mapping 先が無いようにして例外を発生させる
    input_xlsx = make_input_book(
        tmp_path / "作業対象.xlsx",
        [("C001", "", "", "")],
    )

    config_path = _write_config(
        tmp_path,
        west=str(west),
        east=str(east),
        input_xlsx=str(input_xlsx),
        mapping={"お名前": "存在しない列"},
    )
    config.read(config_path)

    with pytest.raises(TransferDestinationColumnNotFoundError):
        run()

    # 失敗時は元ファイルが残る（消えていると再実行できない）
    assert west.exists()
    assert east.exists()
    assert input_xlsx.exists()
    assert not (tmp_path / "最終_作業対象.xlsx").exists()
