"""
src/excel_writer.py — INPUT エクセルに転記し、最終ファイルを保存する

2 つの保存経路を用意する:

1. パスワードが空欄のとき:
   ExcelWriter（openpyxl）で転記 → ``f.save(path=最終パス)`` で完了。
   openpyxl はパスワードを付ける機能がないため、COM は使わない。
   Excel が入っていない環境でも動かせる。

2. パスワードが指定されているとき:
   ExcelWriter で転記 → ``f.save(path=一時ファイル)`` で**出力先と同じフォルダの
   一時ファイル**へ保存 → 一時ファイルを ``ExcelComHandler`` で開いて
   ``save_as(最終パス, read_pw=パスワード)`` → 一時ファイルを削除。
   openpyxl がパスワードを保存できないので、最終ファイルへの書込みだけ
   COM 経由で行う（openpyxl の内容へ上書きするわけではない）。

いずれの経路でも、出力先に同名のファイルが既にあれば上書きされる。
"""

import logging
import tempfile
from pathlib import Path

from comken.runtime import dry_run_log, is_dry_run
from comken.toolbox.excel import ExcelWriter
from comken.toolbox.windows.handler import ExcelComHandler

logger = logging.getLogger(__name__)


def transfer_and_save(
    input_path: Path,
    output_path: Path,
    sheet_name: str,
    key_column: str,
    lookup: dict[str, dict[str, str]],
    mapping: dict[str, str],
    header_row: int,
    password: str,
) -> int:
    """INPUT エクセルに転記し、最終ファイルを保存する。転記件数を返す。

    Args:
        input_path: 転記対象の INPUT エクセル。
        output_path: 最終ファイルの保存先（``最終_ + 元のファイル名``）。
        sheet_name: 転記対象シート名。
        key_column: キー列（INPUT エクセルの「業務用ID」など）。
        lookup: {キー: 行データ} の転記元辞書。西と東をマージ済み。
        mapping: {転記元の列名: 転記先の列名} の対応表。
        header_row: INPUT エクセルの見出し行番号（1始まり）。
        password: 開封パスワード。空文字なら COM を使わない経路で保存する。

    Returns:
        キー列の値が lookup にヒットした行数。

    Raises:
        comken.exceptions.ExcelColumnNotFoundError: キー列や mapping 先の列が見出しにない場合。
        comken.exceptions.ExcelApplicationNotAvailableError:
            パスワード付き保存時、この PC に Excel が入っていない場合。
    """
    if password:
        return _transfer_and_save_with_password(
            input_path=input_path,
            output_path=output_path,
            sheet_name=sheet_name,
            key_column=key_column,
            lookup=lookup,
            mapping=mapping,
            header_row=header_row,
            password=password,
        )
    return _transfer_and_save_openpyxl(
        input_path=input_path,
        output_path=output_path,
        sheet_name=sheet_name,
        key_column=key_column,
        lookup=lookup,
        mapping=mapping,
        header_row=header_row,
    )


def _transfer_and_save_openpyxl(
    input_path: Path,
    output_path: Path,
    sheet_name: str,
    key_column: str,
    lookup: dict[str, dict[str, str]],
    mapping: dict[str, str],
    header_row: int,
) -> int:
    """openpyxl だけで完結する経路（パスワードなし）。"""
    with ExcelWriter(input_path) as f:
        sheet = f.sheet(sheet_name)
        matched = sheet.transfer_by_mapping(
            key_col=key_column,
            lookup=lookup,
            mapping=mapping,
            header_row=header_row,
        )
        f.save(path=output_path)
    logger.info("最終ファイルを保存しました（パスワードなし）: %s", output_path)
    return matched


def _transfer_and_save_with_password(
    input_path: Path,
    output_path: Path,
    sheet_name: str,
    key_column: str,
    lookup: dict[str, dict[str, str]],
    mapping: dict[str, str],
    header_row: int,
    password: str,
) -> int:
    """パスワード付き保存: openpyxl で転記 → 一時ファイルへ保存 → COM で別名保存。

    一時ファイルは出力先と同じフォルダに置く（COM がUNCパスを不安定に扱うため）。
    例外が起きたときも一時ファイルは必ず削除する。

    dry-run のときは COM を起動しないため、**一切の一時ファイルを作らない**。
    代わりに「何件一致するか」「最終ファイルをどこに保存するか」だけをログに出して
    終わる。転記件数を見たいので、``ExcelWriter`` の ``transfer_by_mapping`` は
    そのまま実行する（``save`` は dry-run でログだけ出してスキップされる）。
    """
    if is_dry_run():
        # dry-run では COM を起動できない（一時ファイルを「Excel で開く」段階で
        # 失敗する）。COM を避けるため一時ファイル自体を作らず、予定だけログへ出す。
        with ExcelWriter(input_path) as f:
            sheet = f.sheet(sheet_name)
            matched = sheet.transfer_by_mapping(
                key_col=key_column,
                lookup=lookup,
                mapping=mapping,
                header_row=header_row,
            )
        # パスワードは秘匿値なので、ログには出さない（パスだけ）
        dry_run_log(
            "パスワード付きで最終ファイルを保存する予定: %s（転記予定 %d 件）",
            output_path,
            matched,
        )
        return matched

    tmp_path = _reserve_tmp_path(output_path)
    try:
        # 1) openpyxl で転記し、同じフォルダの一時ファイルへ保存する
        with ExcelWriter(input_path) as f:
            sheet = f.sheet(sheet_name)
            matched = sheet.transfer_by_mapping(
                key_col=key_column,
                lookup=lookup,
                mapping=mapping,
                header_row=header_row,
            )
            f.save(path=tmp_path)

        # 2) 一時ファイルを COM で開き、最終パスへ「開封パスワード」を付けて別名保存する
        #    openpyxl が出力した内容をそのまま書き直すだけ（内容は触らない）
        with ExcelComHandler(tmp_path) as com:
            com.save_as(output_path, read_pw=password)
        # パスワードは秘匿値なので、ログには出さない
        logger.info("最終ファイルを保存しました（パスワード付き）: %s", output_path)
        return matched
    finally:
        # 一時ファイルは業務ファイルではないので dry-run の有無に関わらず必ず消す。
        # comken.delete_file() は dry-run 中にログだけでスキップするため、ここでは
        # Path.unlink() を直接呼んで確実に削除する（comken の ExcelWriter.save() も
        # 同じ流儀で一時ファイルを片付ける）。
        tmp_path.unlink(missing_ok=True)


def _reserve_tmp_path(output_path: Path) -> Path:
    """出力先と同じフォルダ・拡張子の一時ファイル名を確保して返す。

    COM は拡張子で形式を判断するため、出力先と同じ拡張子を保つ。
    NamedTemporaryFile で名前だけ確保して即座に閉じ、呼び出し側がパスから
    ファイルを作成できる状態にする（COM は同名ファイルへの上書きを避けるため、
    この時点ではまだファイルが無いほうが扱いやすい）。
    """
    tmp = tempfile.NamedTemporaryFile(  # noqa: SIM115
        dir=output_path.parent,
        prefix=f".{output_path.name}.",
        suffix=output_path.suffix,
        delete=False,
    )
    tmp_path = Path(tmp.name)
    tmp.close()
    return tmp_path
