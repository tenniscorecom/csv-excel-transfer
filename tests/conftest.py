"""tests/conftest.py — テスト共通の前準備

INPUT エクセルの雛形・サンプル CSV を ``tmp_path`` に作る fixture を置く。
"""

from collections.abc import Iterable
from pathlib import Path

import pytest
from openpyxl import Workbook

INPUT_HEADERS = ["業務用ID", "氏名", "住所", "電話番号"]
SHEET = "Sheet1"
SAMPLE_MAPPING = {
    "お名前": "氏名",
    "ご住所": "住所",
    "電話番号": "電話番号",
}


@pytest.fixture
def make_input_book():
    """ヘッダーと任意の行を持つ INPUT エクセルを作る。"""

    def _make(path: Path, rows: Iterable[Iterable[object]]) -> Path:
        workbook = Workbook()
        worksheet = workbook.active
        assert worksheet is not None
        worksheet.title = SHEET
        worksheet.append(INPUT_HEADERS)
        for row in rows:
            worksheet.append(list(row))
        workbook.save(path)
        workbook.close()
        return path

    return _make


@pytest.fixture
def make_csv():
    """ヘッダーと任意の行を持つ CSV を作る。"""

    def _make(path: Path, headers: Iterable[str], rows: Iterable[Iterable[object]]) -> Path:
        with path.open("w", encoding="utf-8", newline="") as file:
            file.write(",".join(headers) + "\n")
            for row in rows:
                file.write(",".join(str(value) for value in row) + "\n")
        return path

    return _make
