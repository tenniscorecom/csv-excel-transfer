"""src/run.py — 処理の本体

業務の流れ:
    1. 西CSV・東CSVを index_files() で1つの lookup 辞書にマージする
    2. INPUT エクセルの「業務用ID」列をキーに、lookup の値で転記する
    3. パスワード付き（指定があれば）で「最終_ + 元のファイル名」へ保存する
       （最終ファイルの存在確認は comken の save() が自分で行う）
    4. すべて成功したら、西CSV・東CSV・INPUT エクセル（元ファイル）を削除する
       途中で失敗したら元ファイルは消さない（消えると再実行が効かなくなる）

設定の読み取りは `comken.config` を直接使う。
main.py で `config.require(...)` が必須項目をまとめて確かめたあと、
ここでは型変換つきのアクセサで値を取り出す。

設計判断とルールの理由は docs/仕様書.md を参照。
"""

from dataclasses import dataclass
from pathlib import Path

from comken import config
from comken.core import delete_files
from comken.exceptions import ConfigKeyNotFoundError
from comken.toolbox.csv import index_files
from comken.toolbox.excel import ExcelWriter


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
        comken.exceptions.ComkenError: 設定不備・CSV の重複キー・保存や削除の失敗など。
            個別の例外と対処は docs/ERRORS.md を参照。
    """
    # 出力パスは config.text("EXCEL.OUTPUT_PREFIX") が空欄を弾くため、
    # 接頭辞は必ず1文字以上で、出力ファイル名は入力と必ず別名になる
    output_path = config.FILES.INPUT_XLSX.parent / (
        config.text("EXCEL.OUTPUT_PREFIX") + config.FILES.INPUT_XLSX.name
    )

    lookup = index_files(
        [config.FILES.WEST_CSV, config.FILES.EAST_CSV],
        config.CSV.KEY_COLUMN,
    )

    # 転記から外したい行があれば lookup 自体を絞り込んでおく。
    # SKIP_COLUMN と SKIP_VALUES はどちらも設定が無ければ何もしない。
    try:
        skip_column = config.text("EXCEL.SKIP_COLUMN", allow_empty=True)
    except ConfigKeyNotFoundError:
        skip_column = ""
    try:
        skip_values_str = config.text("EXCEL.SKIP_VALUES", allow_empty=True)
    except ConfigKeyNotFoundError:
        skip_values_str = ""
    if skip_column and skip_values_str:
        skip_values = {value.strip() for value in skip_values_str.split(",") if value.strip()}
        lookup = {
            key: row
            for key, row in lookup.items()
            if str(row.get(skip_column, "")) not in skip_values
        }

    # ブックを開く → シートを取って転記 → 保存（パスワード有無は ExcelWriter に任せる）
    # save() の read_pw="" はパスワード無し経路。分岐・dry-run・保存後の存在確認は comken 側で行う
    with ExcelWriter(config.FILES.INPUT_XLSX) as writer:
        sheet = writer.sheet(config.EXCEL.SHEET)
        matched = sheet.transfer_by_mapping(
            key_col=config.EXCEL.KEY_COLUMN,
            lookup=lookup,
            mapping=config.mapping("転記_MAPPING"),
            header_row=config.int_value("EXCEL.HEADER_ROW", minimum=1),
        )
        writer.save(path=output_path, read_pw=str(config.EXCEL.PASSWORD))

    # 最終ファイルの保存が成功した場合にだけ元ファイルを消す。途中で例外が出ると
    # 元ファイルが残るため、再実行すれば何度でもやり直せる
    delete_files(
        [config.FILES.WEST_CSV, config.FILES.EAST_CSV, config.FILES.INPUT_XLSX],
        missing_ok=True,
    )

    return TransferResult(output_path=output_path, matched_rows=matched)
