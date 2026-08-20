"""csv-excel-transfer のエントリポイント。"""

import logging

from comken import Config, debug, dry_run, setup_logging
from comken.exceptions import ComkenError

from src.run import run, validate_config

logger = logging.getLogger(__name__)


def main() -> None:
    settings = Config()
    validate_config(settings)
    run(settings)


if __name__ == "__main__":
    setup_logging()
    try:
        with dry_run(), debug():
            main()
    except ComkenError as error:
        logger.error("処理を中断しました: %s", error)
        raise
    except Exception:
        logger.error("予期しないエラーが発生しました", exc_info=True)
        raise
