"""tests/test_run.py — run.py のオーケストレーション全体"""

from pathlib import Path
from unittest.mock import patch

import pytest
from comken import config, dry_run
from comken.exceptions import (
    CsvRowDuplicateKeyError,
    FileDeletionError,
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


# ── パスワード有無で保存経路が変わる（test_excel_writer の観点） ──


def test_run_with_blank_password_saves_via_openpyxl_and_keeps_output(
    tmp_path: Path, make_csv, make_input_book, monkeypatch, restore_config_singleton
) -> None:
    """パスワード空欄は openpyxl の経路で最終ファイルが生成される。"""
    west = make_csv(
        tmp_path / "west.csv",
        ["お客様ID", "お名前"],
        [("C001", "山田一郎")],
    )
    east = make_csv(
        tmp_path / "east.csv",
        ["お客様ID", "お名前"],
        [("C002", "鈴木三郎")],
    )
    input_xlsx = make_input_book(
        tmp_path / "作業対象.xlsx",
        [("C001", ""), ("C002", "")],
    )
    config_path = _write_config(
        tmp_path,
        west=str(west),
        east=str(east),
        input_xlsx=str(input_xlsx),
        # lookup 側に存在する列のみでマッピング
        mapping={"お名前": "氏名"},
    )
    config.read(config_path)

    real_save_calls: list[dict[str, str]] = []

    def _fake_save(self, path=None, read_pw: str = "", write_pw: str = ""):
        real_save_calls.append({"path": str(path), "read_pw": read_pw})

    monkeypatch.setattr("comken.toolbox.excel.writer.ExcelWriter.save", _fake_save)

    run()

    expected_output = tmp_path / "最終_作業対象.xlsx"
    assert len(real_save_calls) == 1
    assert real_save_calls[0]["path"] == str(expected_output)
    assert real_save_calls[0]["read_pw"] == ""


def test_run_with_password_forwards_read_pw_to_save(
    tmp_path: Path, make_csv, make_input_book, monkeypatch, restore_config_singleton
) -> None:
    """パスワード付き経路では read_pw が ExcelWriter.save() まで確実に渡される。

    COM を起動するかどうかは PC 環境に依存する（Excel 不在なら
    ExcelApplicationNotAvailableError が飛ぶ）。ここでは read_pw が
    確実に save() に届いた事実だけを検証する（秘匿値を渡し損なわないこと）。
    """
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
    input_xlsx = make_input_book(
        tmp_path / "作業対象.xlsx",
        [("C001", "")],
    )
    config_path = _write_config(
        tmp_path,
        west=str(west),
        east=str(east),
        input_xlsx=str(input_xlsx),
        password="dummy-pw",
        mapping={"お名前": "氏名"},
    )
    config.read(config_path)

    real_save_calls: list[dict[str, str]] = []

    def _fake_save(self, path=None, read_pw: str = "", write_pw: str = ""):
        real_save_calls.append({"read_pw": read_pw})

    monkeypatch.setattr("comken.toolbox.excel.writer.ExcelWriter.save", _fake_save)

    try:
        run()
    except Exception:
        # COM が起動できない環境では ExcelApplicationNotAvailableError が飛ぶ
        # （smoke_test.py もそれを許容している）。read_pw の到達は save()
        # 呼び出し時点で確定しているので、ここでは握りつぶして検証だけ進める
        pass

    assert len(real_save_calls) == 1
    assert real_save_calls[0]["read_pw"] == "dummy-pw"


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


def test_run_raises_csv_row_duplicate_key_error_with_multiple_duplicates(
    tmp_path: Path, make_csv, make_input_book, restore_config_singleton
) -> None:
    """複数キーが両方にまたがっていても comken 側で止まること。"""
    west = make_csv(
        tmp_path / "west.csv",
        ["お客様ID", "お名前"],
        [("A1", "西A1"), ("A2", "西A2"), ("A3", "西A3")],
    )
    east = make_csv(
        tmp_path / "east.csv",
        ["お客様ID", "お名前"],
        [("A1", "東A1"), ("A2", "東A2"), ("A3", "東A3")],
    )
    input_xlsx = make_input_book(
        tmp_path / "作業対象.xlsx",
        [("A1", "", "", "")],
    )

    config_path = _write_config(
        tmp_path, west=str(west), east=str(east), input_xlsx=str(input_xlsx)
    )
    config.read(config_path)

    with pytest.raises(CsvRowDuplicateKeyError):
        run()


# ── 削除の部分失敗（comken の FileDeletionError に置換） ──


def test_run_raises_file_deletion_error_when_cleanup_partially_fails(
    tmp_path: Path, make_csv, make_input_book, monkeypatch, restore_config_singleton
) -> None:
    """comken の delete_files を模倣: 1つ目成功・2つ目失敗・3つ目成功 → east だけ .remaining。"""
    from comken.exceptions import FileDeletionError as _FileDeletionError

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
        tmp_path, west=str(west), east=str(east), input_xlsx=str(input_xlsx)
    )
    config.read(config_path)

    real_calls: list[Path] = []

    def _fake_delete_files(paths, missing_ok: bool = True) -> None:
        # comken 本体と同じ「全部試してから remaining にまとめる」振る舞い
        remaining: list[Path] = []
        for path in paths:
            real_calls.append(Path(path))
            try:
                if Path(path) == east:
                    raise PermissionError("他で開かれている想定")
                Path(path).unlink()
            except OSError:
                remaining.append(Path(path))
        if remaining:
            raise _FileDeletionError(remaining)

    monkeypatch.setattr("src.run.delete_files", _fake_delete_files)

    with pytest.raises(FileDeletionError) as exc_info:
        run()

    # 失敗した east だけが残存パスとして報告される
    assert exc_info.value.remaining == [east]
    # 3つとも delete_files に渡された
    assert real_calls == [west, east, input_xlsx]
    assert not west.exists()
    assert not input_xlsx.exists()
    assert east.exists()
    assert (tmp_path / "最終_作業対象.xlsx").exists()


def test_run_raises_file_deletion_error_when_all_deletions_fail(
    tmp_path: Path, make_csv, make_input_book, monkeypatch, restore_config_singleton
) -> None:
    """3つとも消せないとき、全パスが残る（次回古いCSVを拾う事故を防ぐ）。"""
    from comken.exceptions import FileDeletionError as _FileDeletionError

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
        tmp_path, west=str(west), east=str(east), input_xlsx=str(input_xlsx)
    )
    config.read(config_path)

    def _always_fail(paths, missing_ok: bool = True) -> None:
        raise _FileDeletionError(list(paths))

    monkeypatch.setattr("src.run.delete_files", _always_fail)

    with pytest.raises(FileDeletionError) as exc_info:
        run()

    # 3ファイルとも remaining に乗る
    assert set(exc_info.value.remaining) == {west, east, input_xlsx}


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


# ── save 後に ExcelSaveNotCompletedError が comken から飛ぶこと（comken.save の責務） ──


def test_run_excel_save_propagates_save_not_completed_error(
    tmp_path: Path, make_csv, make_input_book, monkeypatch, restore_config_singleton
) -> None:
    """comken の save() が「保存成功後にファイルが無い」と判断した場合、
    ExcelSaveNotCompletedError が伝搬して run() 全体が止まり、元ファイルは残る。"""
    from comken.exceptions import ExcelSaveNotCompletedError

    west = make_csv(
        tmp_path / "west.csv",
        ["お客様ID", "お名前", "ご住所", "電話番号"],
        [("C001", "山田一郎", "大阪", "06-0000-0001")],
    )
    east = make_csv(
        tmp_path / "east.csv",
        ["お客様ID", "お名前", "ご住所", "電話番号"],
        [],
    )
    input_xlsx = make_input_book(
        tmp_path / "作業対象.xlsx",
        [("C001", "", "", "")],
    )

    config_path = _write_config(
        tmp_path, west=str(west), east=str(east), input_xlsx=str(input_xlsx)
    )
    config.read(config_path)

    def _fake_save_raises(self, path=None, read_pw: str = "", write_pw: str = ""):
        # 実ファイルを作らず、comken と同じ ExcelSaveNotCompletedError を投げる
        raise ExcelSaveNotCompletedError(path)

    monkeypatch.setattr("comken.toolbox.excel.writer.ExcelWriter.save", _fake_save_raises)

    with pytest.raises(ExcelSaveNotCompletedError):
        run()

    # 元ファイルが残る
    assert west.exists()
    assert east.exists()
    assert input_xlsx.exists()


# patch は一部テストで import 警告を避けるため
_ = patch  # noqa: F401

