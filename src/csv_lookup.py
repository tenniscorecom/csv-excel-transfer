"""
src/csv_lookup.py — 西CSV・東CSV を 1つの lookup 辞書へマージする

CsvReader.index() は {キー: 行} の形式で 1つの CSV を引く。同じお客様IDが
ファイル内で重複していたら例外で止まる。西と東を単純に足し合わせると、
**両方に同じお客様IDがあるとき**にどちらを採用したか分からなくなる。
そこでマージ後に同じキーがあれば例外で止める（西を優先・東を優先のどちらでも、
業務的に「正しい方」が決まっていない以上は止まる方が安全）。
"""

import logging
from pathlib import Path

from comken.toolbox.csv import CsvReader

from src.exceptions import CustomerIdDuplicateAcrossCsvError

logger = logging.getLogger(__name__)


def merge_lookups(west_path: Path, east_path: Path, key_column: str) -> dict[str, dict[str, str]]:
    """西CSVと東CSVを `index()` で引き、1つの辞書にまとめて返す。

    同じ `key_column` の値が両方にあったら例外で止める。住所で西と東を
    切り替える要件は今は無いので、重複が見つかった瞬間に止めている
    （要件が現実になったら、例外を出すロジックを「ルールに従って片方を選ぶ」へ
    切り替える。docs/仕様書.md 参照）。

    Args:
        west_path: 西CSVのパス。
        east_path: 東CSVのパス。
        key_column: キー列（既定: お客様ID）。

    Returns:
        `{キー: 行データ}` の辞書。キーは文字列（CSVの値）として扱う。

    Raises:
        FileNotFoundError: 西CSV・東CSVのいずれかが無い場合。
        CustomerIdDuplicateAcrossCsvError: 両方に同じキーがあった場合。
    """
    # CsvReader.index() は同じCSV内でキーが重複していると例外で止まる。
    # ここでは CSV 単位では重複が無い前提で、両方の index() を足し合わせる
    west = CsvReader(west_path).index(key_column)
    east = CsvReader(east_path).index(key_column)

    logger.info("西CSV: %d 件、東CSV: %d 件", len(west), len(east))

    duplicates: list[str] = []
    merged: dict[str, dict[str, str]] = dict(west)
    for key, row in east.items():
        if key in merged:
            duplicates.append(key)
            continue
        merged[key] = row

    if duplicates:
        # 5件まで代表として添える。総件数も別行で出して、件数の全体感を伝える
        raise CustomerIdDuplicateAcrossCsvError(duplicates)

    logger.info("マージ後の件数: %d 件", len(merged))
    return merged
