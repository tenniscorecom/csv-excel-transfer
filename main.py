"""csv-excel-transfer のエントリポイント。"""

import logging

from comken import comken_logger, debug, dry_run
from comken.exceptions import ComkenError

from src.run import run

logger = logging.getLogger(__name__)


def main() -> None:
    run()


if __name__ == "__main__":
    comken_logger.setup_local_logging()
    try:
        with dry_run(), debug():
            main()
    except ComkenError as error:
        logger.error("処理を中断しました: %s", error)
        raise
    except Exception:
        logger.error("予期しないエラーが発生しました", exc_info=True)
        raise
