"""
src/exceptions.py — このツール固有のエラー

main.py が ComkenError でまとめて受け取れるよう、comken の例外体系にぶら下げる。
メッセージには「何が・どこで・どうすればよいか」を書く（docs/ERRORS.md も更新すること）。
"""

from pathlib import Path

from comken.exceptions import ComkenError

DUPLICATE_SAMPLE_LIMIT = 5


class InvalidOutputPrefixError(ComkenError):
    """[EXCEL] OUTPUT_PREFIX が空欄の場合。"""

    def __init__(self) -> None:
        super().__init__(
            "config.ini の [EXCEL] OUTPUT_PREFIX を空欄にすることはできません。\n"
            "OUTPUT_PREFIX に「最終_」のような出力ファイル名の接頭辞を設定してください。\n"
            "空欄では出力先が INPUT エクセル自身になり、転記後に元ファイルを削除する処理で"
            "成果物まで消えてしまうためです。"
        )


class InputEqualsOutputError(ComkenError):
    """出力パスが入力パス（INPUT エクセル）と一致してしまう場合。

    ``[EXCEL] OUTPUT_PREFIX`` が空欄、または指定した値から組み立てた出力パスが
    INPUT エクセル自身を指してしまう場合に送出する。転記後に元ファイルを削除する
    処理で、保存した成果物まで消えてしまうためである。

    config.ini を読む段階（`InvalidOutputPrefixError`）で止められない構成や、
    相対パス・大文字小文字の違いで文字列比較だけだと擦り抜けるケースを
    ``Path.resolve()`` で検出して止めるための最終砦。
    """

    def __init__(self, input_path: Path, output_path: Path) -> None:
        super().__init__(
            f"出力先が入力ファイル自身と解決されてしまいます。\n"
            f"  入力: {input_path}\n"
            f"  出力: {output_path}\n"
            "config.ini の [EXCEL] OUTPUT_PREFIX に値を入れてください（例: 最終_）。\n"
            "空欄や、組み立てたパスが INPUT エクセル自身になる値だと、"
            "転記後に元ファイルを削除する処理で保存した成果物まで消えてしまうためです。"
        )


class OutputFileMissingError(ComkenError):
    """最終ファイル（保存先）の保存が、実ファイルとして残っていない場合。

    openpyxl と Excel COM（``DisplayAlerts=False``）は保存に失敗しても
    例外を投げないことがある。最終ファイルを消す前に最後の砦として
    ファイルの存在を確かめ、無ければ元ファイルを消す前に例外で停止する。
    """

    def __init__(self, output_path: Path) -> None:
        super().__init__(
            f"最終ファイルの保存が完了していない可能性があります: {output_path}\n"
            "openpyxl または Excel COM が保存に失敗しても、DisplayAlerts=False のときは"
            "静かに成功したように見えることがあります。\n"
            "出力先に既に同じ名前のファイルが他で開かれていないか、ディスクの空き容量があるか"
            "を確認し、閉じてから再実行してください。"
        )


class SourceFileDeletionError(ComkenError):
    """元ファイル（西CSV・東CSV・INPUT エクセル）の削除が一部でも失敗した場合。

    最終ファイルは既に保存済みなので、消せたものはそのまま消す方針としたうえで、
    残ったファイルのパスを全て並べて報告する。古いCSV が次回実行で拾われるのを
    防ぐため、バッチとしては失敗扱いとする。
    """

    def __init__(self, remaining: list[Path]) -> None:
        self.remaining: list[Path] = list(remaining)
        lines = "\n".join(f"  - {path}" for path in remaining)
        super().__init__(
            "元ファイルの削除が一部失敗しました。最終ファイルは保存済みですが、"
            "次回実行で古いCSVを拾う事故を防ぐため、バッチは失敗扱いにします。\n"
            "残っているファイル:\n"
            f"{lines}\n"
            "ファイルが他で開かれていないか、読み取り専用になっていないかを確認して"
            "手動で削除するか、もう一度実行してください。"
        )


class InvalidHeaderRowError(ComkenError):
    """[EXCEL] HEADER_ROW が1以上の整数でない場合。"""

    def __init__(self, value: object) -> None:
        super().__init__(
            "config.ini の [EXCEL] HEADER_ROW は1以上の整数で指定してください。"
            f"設定値: {value}\n"
            "見出しが1行目なら1、2行目なら2のように修正してください。"
        )


class CustomerIdDuplicateAcrossCsvError(ComkenError):
    """西CSVと東CSVの両方に同じお客様IDがあった場合。

    どちらを採用したか分からないまま転記が進むと、結果が静かにブレる。
    例外で止まって、原因を確定させてから動かす。
    """

    def __init__(self, duplicates: list[str]) -> None:
        # 件数が多いとログが読めなくなるので、代表的な5件だけ添える
        sample_ids = duplicates[:DUPLICATE_SAMPLE_LIMIT]
        sample = ", ".join(sample_ids)
        other_count = len(duplicates) - len(sample_ids)
        extra = f"（他 {other_count} 件）" if other_count else ""
        total = len(duplicates)
        super().__init__(
            "西CSVと東CSVの両方に同じお客様IDが見つかりました。\n"
            f"重複したお客様ID: {sample}{extra}\n"
            f"重複数: {total} 件\n"
            "西と東で同じお客様IDの行がないか確認し、片方に寄せるか、"
            "西/東で住所を分けるなどの運用ルールを決めてください。"
        )
