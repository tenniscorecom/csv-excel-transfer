"""main.py — エントリポイント

このプロジェクトの入口。`実行.bat` か `python main.py` で実行する。
処理の本体は src/ 以下に書き、ここでは「実行 → エラーの受け止め」だけを行う。
"""

import logging

from comken import config, debug, dry_run, setup_logging
from comken.exceptions import ComkenError

from src.run import run

logger = logging.getLogger(__name__)


def main() -> None:
    # 設定の読み取りも処理も src/run.py に書く。ここは呼ぶだけにしておく
    run()


if __name__ == "__main__":
    # 単体で動かすので、ログの出力先をここで用意する（コンソールと logs/YYYY-MM-DD.log）。
    setup_logging()
    try:
        # config.ini に必要な項目がそろっているかを最初に確かめる。
        # 途中まで動いてから足りないと分かるより、動き出す前に全部まとめて出す。
        # 使う項目を増やしたらここにも足す（消しても動くが、エラーが遅くなる）
        config.require(
            "RUN.DRY_RUN",
            "RUN.DEBUG",
            "FILES.WEST_CSV",
            "FILES.EAST_CSV",
            "FILES.INPUT_XLSX",
            "CSV.KEY_COLUMN",
            "EXCEL.SHEET",
            "EXCEL.KEY_COLUMN",
            "EXCEL.HEADER_ROW",
            "EXCEL.OUTPUT_PREFIX",
            "EXCEL.PASSWORD",
        )
        # [転記_MAPPING] は config.require() の形式（"SECTION.KEY"）で書けないので、
        # ここで別途存在を確かめる。セクションが無ければ ConfigSectionNotFoundError が飛ぶ。
        # 空セクションも不可（中身が無いと転記が0件のまま動く）
        mapping = config.mapping("転記_MAPPING")
        if not mapping:
            logger.error("[転記_MAPPING] セクションに転記元→転記先の対応が1行も書かれていません")
            raise SystemExit(1)

        # config.ini の [RUN] DRY_RUN で切り替える。True の間は書き込み・移動・保存を
        # せず、何をするつもりかだけログに出す。コードを触らずに試せるので、
        # 本番前の確認を非エンジニアだけで回せる。
        #
        # True のまま戻し忘れると「毎日成功しているのに何も出力されない」状態になり、
        # 終了コードも 0 なのでスケジューラからは正常に見える。気づけるように
        # 実行のたび WARNING を出す（INFO は流し読みされるので警告にする）。
        if config.RUN.DRY_RUN:
            logger.warning(
                "DRY-RUN で実行します。ファイルは書き込まれません"
                "（本番で動かすなら config.ini の [RUN] DRY_RUN を False にする）"
            )

        # config.ini の [RUN] DEBUG で切り替える。True だと @measure を付けた
        # メソッドの出入りを DEBUG ログへ出す。外部待ちでバッチが止まったときに
        # True にして再実行すると、ログの末尾が「DEBUG ○○: 開始」の行で止まるので、
        # どこで止まったかが分かる。普段は False のままでよい。
        #
        # True のままでも業務は正常に動く（ログが増えるだけ）ので dry-run ほど
        # 危険ではないが、ログが膨らみ続けるので検証後は False に戻す。
        # 戻し忘れに気づきやすいよう、True のときは INFO を1行出す。
        if config.RUN.DEBUG:
            logger.info(
                "DEBUG モードで実行します。"
                "各メソッドの開始/完了ログが出ます（検証後 False に戻してください）"
            )

        with dry_run(config.RUN.DRY_RUN), debug(config.RUN.DEBUG):
            main()
    except ComkenError as e:
        # comken のエラーはメッセージに対処法が入っている（docs/ERRORS.md も参照）
        logger.error("処理を中断しました: %s", e)
        raise
    except Exception:
        logger.error("予期しないエラーが発生しました", exc_info=True)
        raise
