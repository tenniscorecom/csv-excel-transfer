"""tests/test_run.py — run.py のオーケストレーション全体"""

from pathlib import Path
from unittest.mock import patch

import pytest
from comken import dry_run

from src.exceptions import (
    InputEqualsOutputError,
    OutputFileMissingError,
    SourceFileDeletionError,
)
from src.run import run
from src.settings import Settings, SourceLayout
from tests.conftest import SAMPLE_MAPPING


def _build_settings(
    west_csv: Path, east_csv: Path, input_xlsx: Path, output_prefix: str = "最終_"
) -> Settings:
    return Settings(
        west_csv_path=west_csv,
        east_csv_path=east_csv,
        input_xlsx_path=input_xlsx,
        layout=SourceLayout(
            csv_key_column="お客様ID",
            excel_sheet="Sheet1",
            excel_key_column="業務用ID",
            excel_header_row=1,
        ),
        mapping=SAMPLE_MAPPING,
        output_prefix=output_prefix,
        password="",
    )


def test_run_creates_final_file_and_deletes_source_files(
    tmp_path: Path, make_csv, make_input_book, monkeypatch
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

    settings = _build_settings(west, east, input_xlsx)
    monkeypatch.setattr("src.run.load_settings", lambda: settings)

    result = run()

    assert result.output_path == expected_output
    assert result.matched_rows == 2
    assert expected_output.exists()
    # 元の3ファイルは削除済み
    assert not west.exists()
    assert not east.exists()
    assert not input_xlsx.exists()


def test_run_rejects_blank_output_prefix_before_transfer(
    tmp_path: Path, make_csv, make_input_book, monkeypatch
) -> None:
    """OUTPUT_PREFIX が空欄のとき、出力パスが INPUT 自身になるので例外で止める。"""
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
        [("C001", "", "", "")],
    )
    settings = _build_settings(west, east, input_xlsx, output_prefix="")
    monkeypatch.setattr("src.run.load_settings", lambda: settings)

    with pytest.raises(InputEqualsOutputError):
        run()

    # 元の3ファイルは残ったまま（消えると業務データを失う）
    assert west.exists()
    assert east.exists()
    assert input_xlsx.exists()


def test_run_rejects_output_that_resolves_to_input_path(
    tmp_path: Path, make_csv, make_input_book, monkeypatch
) -> None:
    """出力先が INPUT と別文字列でも、Path.resolve() で同じ場所を指すなら止める。

    config.ini で ``INPUT_XLSX = dummy/../作業対象.xlsx`` のように ``..`` を含む
    相対パスが書かれていて OUTPUT_PREFIX が空欄だと、組み立てた出力パスは
    ``dummy/作業対象.xlsx`` になり、文字列としては違うが resolve() で同じ場所に
    解決される。文字列比較だけだと ``..`` の有無で擦り抜けるので、resolve() で
    確実に止める。
    """
    sub = tmp_path / "dummy"
    sub.mkdir()
    west = make_csv(
        sub / "west.csv",
        ["お客様ID", "お名前"],
        [("C001", "山田一郎")],
    )
    east = make_csv(
        sub / "east.csv",
        ["お客様ID", "お名前"],
        [("C002", "鈴木三郎")],
    )
    input_xlsx = make_input_book(
        tmp_path / "作業対象.xlsx",
        [("C001", "", "", "")],
    )
    # INPUT を "dummy/../作業対象.xlsx" 形式で参照させる（組み立て側と文字列が違う）
    settings = Settings(
        west_csv_path=west,
        east_csv_path=east,
        input_xlsx_path=Path("dummy/../作業対象.xlsx"),
        layout=SourceLayout(
            csv_key_column="お客様ID",
            excel_sheet="Sheet1",
            excel_key_column="業務用ID",
            excel_header_row=1,
        ),
        mapping=SAMPLE_MAPPING,
        output_prefix="",
        password="",
    )
    monkeypatch.chdir(tmp_path)
    monkeypatch.setattr("src.run.load_settings", lambda: settings)

    with pytest.raises(InputEqualsOutputError):
        run()

    # 元ファイルが残っている（業務データを守る）
    assert input_xlsx.exists()


def test_run_raises_output_file_missing_when_save_silently_failed(
    tmp_path: Path, make_csv, make_input_book, monkeypatch
) -> None:
    """保存処理が例外を投げなくても、ファイルが実在しなければ元ファイルを消さず例外。"""
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
        [("C001", "", "", "")],
    )
    settings = _build_settings(west, east, input_xlsx)
    monkeypatch.setattr("src.run.load_settings", lambda: settings)

    # transfer_and_save が「成功した」と返したのに、最終ファイルを作らないケースを模擬
    with patch("src.run.transfer_and_save", return_value=1):
        with pytest.raises(OutputFileMissingError):
            run()

    # 元3ファイルが消されていないことが最重要
    assert west.exists()
    assert east.exists()
    assert input_xlsx.exists()
    assert not (tmp_path / "最終_作業対象.xlsx").exists()


def test_run_dry_run_does_not_fail_on_output_missing_check(
    tmp_path: Path, make_csv, make_input_book, monkeypatch
) -> None:
    """dry-run は保存しないので、output_path が無い検査で落ちてはいけない。"""
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
        [("C001", "", "", "")],
    )
    settings = _build_settings(west, east, input_xlsx)
    monkeypatch.setattr("src.run.load_settings", lambda: settings)

    # dry-run ではファイルが何もないのは仕様。OutputFileMissingError を投げていないこと
    with dry_run():
        run()

    # 何も作らず、何 consume せず（元ファイルも残ったまま）
    assert not (tmp_path / "最終_作業対象.xlsx").exists()
    assert west.exists()
    assert east.exists()
    assert input_xlsx.exists()


def test_run_raises_source_file_deletion_error_when_cleanup_partially_fails(
    tmp_path: Path, make_csv, make_input_book, monkeypatch
) -> None:
    """_cleanup_sources で 2 つ目以降の削除に失敗した場合、残存パスを全部並べて例外。"""
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
        [("C001", "", "", "")],
    )
    settings = _build_settings(west, east, input_xlsx)
    monkeypatch.setattr("src.run.load_settings", lambda: settings)

    real_calls: list[Path] = []

    def _fake_delete(path: Path, missing_ok: bool = False) -> None:
        # 1つ目（west）は実際に消す、2つ目（east）は権限エラーを起こす
        # 3つ目（input_xlsx）は消す（最終ファイルは保存済み、消せるものは消したい）
        real_calls.append(path)
        if path == east:
            raise PermissionError("他で開かれている想定")
        path.unlink()

    monkeypatch.setattr("src.run.delete_file", _fake_delete)

    with pytest.raises(SourceFileDeletionError) as exc_info:
        run()

    # 失敗した east だけが残存パスとして報告される（消せた west と input_xlsx は載らない）
    assert exc_info.value.remaining == [east]
    # 3つとも delete_file に渡されたことを確認（途中で止まらず全部試す）
    assert real_calls == [west, east, input_xlsx]
    # 消せたものは実際に消えている
    assert not west.exists()
    assert not input_xlsx.exists()
    # 失敗したものは残っている
    assert east.exists()
    # 最終ファイルは保存済み
    assert (tmp_path / "最終_作業対象.xlsx").exists()


def test_run_raises_source_file_deletion_error_when_all_deletions_fail(
    tmp_path: Path, make_csv, make_input_book, monkeypatch
) -> None:
    """3つとも消せないときは、全パスが並ぶ（次回古いCSVを拾う事故を防ぐ）。"""
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
        [("C001", "", "", "")],
    )
    settings = _build_settings(west, east, input_xlsx)
    monkeypatch.setattr("src.run.load_settings", lambda: settings)

    def _always_fail(path: Path, missing_ok: bool = False) -> None:
        raise PermissionError("他で開かれている想定")

    monkeypatch.setattr("src.run.delete_file", _always_fail)

    with pytest.raises(SourceFileDeletionError) as exc_info:
        run()

    # 残った3ファイルが全部並ぶ
    assert set(exc_info.value.remaining) == {west, east, input_xlsx}
    # 3ファイルとも残っている
    assert west.exists()
    assert east.exists()
    assert input_xlsx.exists()
    # 最終ファイルは保存済み
    assert (tmp_path / "最終_作業対象.xlsx").exists()


def test_run_dry_run_does_not_create_or_delete_any_files(
    tmp_path: Path, make_csv, make_input_book, monkeypatch
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
        [("C001", "", "", "")],
    )
    expected_output = tmp_path / "最終_作業対象.xlsx"

    settings = _build_settings(west, east, input_xlsx)
    monkeypatch.setattr("src.run.load_settings", lambda: settings)

    with dry_run():
        run()

    # DRY-RUN なので、ファイルは何も作られず・消されず
    assert not expected_output.exists()
    assert west.exists()
    assert east.exists()
    assert input_xlsx.exists()


def test_run_does_not_delete_when_transfer_fails(
    tmp_path: Path, make_csv, make_input_book, monkeypatch
) -> None:
    """転記中に失敗したら、元ファイルは消さない（再実行できるようにする）。"""
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
    # INPUT エクセルの見出しに mapping 先が無いようにして例外を発生させる
    input_xlsx = make_input_book(
        tmp_path / "作業対象.xlsx",
        [("C001", "", "", "")],
    )

    bad_mapping = {"お名前": "存在しない列"}
    settings = Settings(
        west_csv_path=west,
        east_csv_path=east,
        input_xlsx_path=input_xlsx,
        layout=SourceLayout(
            csv_key_column="お客様ID",
            excel_sheet="Sheet1",
            excel_key_column="業務用ID",
            excel_header_row=1,
        ),
        mapping=bad_mapping,
        output_prefix="最終_",
        password="",
    )
    monkeypatch.setattr("src.run.load_settings", lambda: settings)

    with pytest.raises(Exception):
        run()

    # 失敗時は元ファイルが残る（消えていると再実行できない）
    assert west.exists()
    assert east.exists()
    assert input_xlsx.exists()
    assert not (tmp_path / "最終_作業対象.xlsx").exists()


def test_run_dry_run_with_password_does_not_create_or_delete_any_files(
    tmp_path: Path, make_csv, make_input_book, monkeypatch
) -> None:
    """パスワード付き経路を dry-run で動かしても、COM を起動せずファイルが増減しない。

    一時ファイル（``.x作業対象.xlsx.XXXXX`` 形式）が tmp_path に残らないことを
    特に確認する。COM が起動できない環境でも通る（dry-run 中は COM を呼ばないため）。
    """
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

    settings = Settings(
        west_csv_path=west,
        east_csv_path=east,
        input_xlsx_path=input_xlsx,
        layout=SourceLayout(
            csv_key_column="お客様ID",
            excel_sheet="Sheet1",
            excel_key_column="業務用ID",
            excel_header_row=1,
        ),
        mapping=SAMPLE_MAPPING,
        output_prefix="最終_",
        password="dummy-pw",  # パスワード付き経路を強制する
    )
    monkeypatch.setattr("src.run.load_settings", lambda: settings)

    with dry_run():
        run()

    # DRY-RUN なので、ファイルは何も作られず・消されず
    assert not (tmp_path / "最終_作業対象.xlsx").exists()
    assert west.exists()
    assert east.exists()
    assert input_xlsx.exists()
    # パスワード経路で生成される一時ファイルも残らない（COM を起動しないため）
    # パスワード経路の一時ファイル名は `.{最終_作業対象.xlsx.}<random>.xlsx` 形式
    leaked = list(tmp_path.glob(".最終_作業対象.xlsx.*"))
    assert leaked == [], f"一時ファイルが残っています: {leaked}"
