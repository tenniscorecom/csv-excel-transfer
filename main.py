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

        # dry-run / debug は `with dry_run():` / `with debug():` の context manager
        # で囲む形のほうが、コードを読む人に「ここが境界」と一目で分かる。
        # 本番で動かすには `with dry_run():` / `with debug():` を外すだけで
        # True 固定になる（=実書き込み）。
        with dry_run(), debug():
            main()
    except ComkenError as e:
        # comken のエラーはメッセージに対処法が入っている（docs/ERRORS.md も参照）
        logger.error("処理を中断しました: %s", e)
        raise
    except Exception:
        logger.error("予期しないエラーが発生しました", exc_info=True)
        raise
