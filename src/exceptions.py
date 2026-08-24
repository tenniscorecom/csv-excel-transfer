"""プロジェクト固有の例外。

``CSVNoDataRowsError`` と ``CSVRowDuplicateKeyError`` は comken から削除されたため、
このプロジェクトで再定義する。両方とも ``ComkenError`` を継承し、
``main.py`` の ``except ComkenError`` で受け止められるようにする。
"""

from comken.exceptions import ComkenError


class CSVNoDataRowsError(ComkenError):
    """CSV にデータ行がない（見出しは存在するが本体が空）。

    発生箇所: src.run._merge_csv
    """

    def __init__(self, path) -> None:
        super().__init__(f"CSV にデータ行がありません: {path}")


class CSVRowDuplicateKeyError(ComkenError):
    """2 つの CSV をマージした結果、キー列の値に重複がある。

    発生箇所: src.run._merge_csv
    """

    def __init__(self, key_column: str, duplicate_counts: dict[str, int], paths: str) -> None:
        keys = ", ".join(f"{key}({count} 件)" for key, count in duplicate_counts.items())
        super().__init__(
            f"CSV のキー列「{key_column}」に 2 ファイル間で重複があります: {keys}\n"
            f"対象ファイル: {paths}"
        )
