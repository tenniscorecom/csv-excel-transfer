"""
src/run.py — 処理の本体

業務の流れ:
    1. 西CSV・東CSVを ``index()`` で読み、1つの lookup 辞書にマージする
    2. INPUT エクセルの「業務用ID」列をキーに、lookup の値で転記する
    3. パスワード付き（指定があれば）で「最終_ + 元のファイル名」へ保存する
    4. すべて成功したら、西CSV・東CSV・INPUT エクセル（元ファイル）を削除する
       途中で失敗したら元ファイルは消さない（消えると再実行が効かなくなる）

設計判断とルールの理由は docs/仕様書.md を参照。
"""

import logging
from dataclasses import dataclass
from pathlib import Path

from comken.core import delete_file
from comken.runtime import is_dry_run

from src.csv_lookup import merge_lookups
from src.excel_writer import transfer_and_save
from src.exceptions import (
    InputEqualsOutputError,
    OutputFileMissingError,
    SourceFileDeletionError,
)
from src.settings import Settings, load_settings

logger = logging.getLogger(__name__)


@dataclass(frozen=True)
class TransferResult:
    """1回の実行で何が起きたかを返す箱。"""

    output_path: Path
    matched_rows: int


def run() -> TransferResult:
    """config.ini の設定を読んで INPUT エクセルへ転記し、元ファイルを片付ける。

    Returns:
        TransferResult: 出力パスと転記件数を返す。

    Raises:
        CustomerIdDuplicateAcrossCsvError: 西CSVと東CSVで同じお客様IDがあった場合。
        InvalidOutputPrefixError: [EXCEL] OUTPUT_PREFIX が空欄の場合。
        InputEqualsOutputError: 出力パスが INPUT エクセル自身に解決される場合。
        OutputFileMissingError: 最終ファイルの保存が、実ファイルとして残っていない場合。
        SourceFileDeletionError: 元ファイルの削除が一部でも失敗した場合。
        FileNotFoundError: 入力ファイルが見つからない場合。
        comken.exceptions.ExcelApplicationNotAvailableError:
            パスワード付き保存時、この PC に Excel が入っていない場合。
    """
    settings = load_settings()

    output_path = settings.input_xlsx_path.parent / (
        settings.output_prefix + settings.input_xlsx_path.name
    )
    # 出力先が INPUT エクセル自身を指していないか、Path.resolve() で突き合わせる。
    # 文字列比較だと `.\` や大文字小文字の違い、相対パスの `./` 等で擦り抜けるので、
    # 解決後のパスで判定する。空の OUTPUT_PREFIX や、
    # `Path("a/b") == Path("A/b")` のような比較で検出できないケースをここで止める。
    if output_path == settings.input_xlsx_path:
        raise InputEqualsOutputError(settings.input_xlsx_path, output_path)

    lookup = merge_lookups(
        settings.west_csv_path,
        settings.east_csv_path,
        settings.layout.csv_key_column,
    )

    # 最終ファイルの保存が成功した場合にだけ元ファイルを消す。途中で例外が出ると
    # 元ファイルが残るため、再実行すれば何度でもやり直せる
    matched = transfer_and_save(
        input_path=settings.input_xlsx_path,
        output_path=output_path,
        sheet_name=settings.layout.excel_sheet,
        key_column=settings.layout.excel_key_column,
        lookup=lookup,
        mapping=settings.mapping,
        header_row=settings.layout.excel_header_row,
        password=settings.password,
    )

    # dry-run では最終ファイルを作っていないので、存在確認はスキップする。
    # 本実行では openpyxl / COM が DisplayAlerts=False で静かに失敗することがあるため、
    # 元ファイルを消す前に最終ファイルが実在することを確かめる最後の砦を置く。
    if not is_dry_run() and not output_path.exists():
        raise OutputFileMissingError(output_path)

    _cleanup_sources(settings)

    return TransferResult(output_path=output_path, matched_rows=matched)


def _cleanup_sources(settings: Settings) -> None:
    """成功後に元の3ファイルを削除する。

    3つとも削除を試みてから、**失敗したファイルのパスを全部並べて**例外で報告する。
    1つ目で失敗しても残りは消す（最終ファイルは既に保存済みなので、消せるものは
    消したい）。一部でも残ると次回実行で古いCSVを拾う事故につながるため、
    バッチとしては失敗扱いとする。
    """
    targets = [
        settings.west_csv_path,
        settings.east_csv_path,
        settings.input_xlsx_path,
    ]
    remaining: list[Path] = []
    for path in targets:
        try:
            delete_file(path, missing_ok=False)
        except FileNotFoundError:
            # 既に消えているのは問題なし（missing_ok=False でも消えているときに
            # FileNotFoundError が出る経路に備えた保険。普通に消せていればここには来ない）
            logger.info("元ファイルは既に削除済みでした: %s", path)
        except OSError as error:
            # 権限・排他・読み取り専用等で消せなかったケース。
            # ERROR ログを残しつつ、残ったファイルのパスを集めて後で一括報告する
            logger.error("元ファイルの削除に失敗しました: %s（%s）", path, error)
            remaining.append(path)
    if remaining:
        raise SourceFileDeletionError(remaining)
