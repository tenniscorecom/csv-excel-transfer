"""tests/test_csv_lookup.py — 西CSV・東CSV のマージ"""

from pathlib import Path

import pytest

from src.csv_lookup import merge_lookups
from src.exceptions import CustomerIdDuplicateAcrossCsvError


def test_merge_combines_unique_ids_from_both_csvs(tmp_path: Path, make_csv) -> None:
    west = make_csv(
        tmp_path / "west.csv",
        ["お客様ID", "お名前", "ご住所"],
        [("W001", "山田一郎", "大阪"), ("W002", "佐藤二郎", "京都")],
    )
    east = make_csv(
        tmp_path / "east.csv",
        ["お客様ID", "お名前", "ご住所"],
        [("E001", "鈴木三郎", "東京")],
    )

    merged = merge_lookups(west, east, "お客様ID")

    assert set(merged) == {"W001", "W002", "E001"}
    assert merged["W001"]["お名前"] == "山田一郎"
    assert merged["E001"]["ご住所"] == "東京"


def test_merge_raises_when_same_id_appears_in_both(tmp_path: Path, make_csv) -> None:
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

    with pytest.raises(CustomerIdDuplicateAcrossCsvError) as exc_info:
        merge_lookups(west, east, "お客様ID")

    # メッセージに代表 + 総件数が含まれる
    assert "X1" in str(exc_info.value)
    assert "1 件" in str(exc_info.value)


def test_merge_raises_when_multiple_duplicates_exist(tmp_path: Path, make_csv) -> None:
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

    with pytest.raises(CustomerIdDuplicateAcrossCsvError) as exc_info:
        merge_lookups(west, east, "お客様ID")

    assert "他" not in str(exc_info.value)
    assert "重複数: 3 件" in str(exc_info.value)


def test_merge_message_includes_other_count_after_five_duplicates(tmp_path: Path, make_csv) -> None:
    west = make_csv(
        tmp_path / "west.csv",
        ["お客様ID", "お名前"],
        [(f"A{i}", f"西A{i}") for i in range(6)],
    )
    east = make_csv(
        tmp_path / "east.csv",
        ["お客様ID", "お名前"],
        [(f"A{i}", f"東A{i}") for i in range(6)],
    )

    with pytest.raises(CustomerIdDuplicateAcrossCsvError) as exc_info:
        merge_lookups(west, east, "お客様ID")

    message = str(exc_info.value)
    assert "他 1 件" in message
    assert "重複数: 6 件" in message
