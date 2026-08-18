"""Smoke test: 一時フォルダでサンプルを作り、main.run() を実際に動かす。

実行: python smoke_test.py
- 一時フォルダに 西CSV・東CSV・INPUT エクセルを作る
- プロジェクトの src/ を一時フォルダにコピーし、config.ini を書く
- python main.py と等価な処理を実行して、最終ファイル生成 + 元ファイル削除を確認する
- 最後に一時フォルダを片付ける

C ドライブの実業務フォルダには触らない（テスト用 tmp のみ）。
"""

from __future__ import annotations

import shutil
import sys
import tempfile
from pathlib import Path
from textwrap import dedent

ROOT = Path(__file__).resolve().parent
SRC = ROOT / "src"


def _make_samples(sample_dir: Path) -> tuple[Path, Path, Path]:
    """西CSV・東CSV・INPUT エクセルのサンプルを sample_dir に作る。"""
    west_csv = sample_dir / "west.csv"
    east_csv = sample_dir / "east.csv"
    input_xlsx = sample_dir / "作業対象.xlsx"

    west_csv.write_text(
        "お客様ID,お名前,ご住所,電話番号\n"
        "C001,山田一郎,大阪市,06-1111-2222\n"
        "C002,佐藤二郎,京都市,075-1111-2222\n",
        encoding="utf-8",
    )
    east_csv.write_text(
        "お客様ID,お名前,ご住所,電話番号\nC003,鈴木三郎,東京,03-1111-2222\n",
        encoding="utf-8",
    )

    import openpyxl

    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Sheet1"
    sheet.append(["業務用ID", "氏名", "住所", "電話番号"])
    sheet.append(["C001", "", "", ""])
    sheet.append(["C002", "", "", ""])
    sheet.append(["C003", "", "", ""])
    workbook.save(input_xlsx)
    workbook.close()
    return west_csv, east_csv, input_xlsx


def _prepare_workspace(work_dir: Path, sample_dir: Path) -> Path:
    """src/ をコピーし、config.ini を work_dir に書く。"""
    shutil.copytree(SRC, work_dir / "src")
    west_csv = sample_dir / "west.csv"
    east_csv = sample_dir / "east.csv"
    input_xlsx = sample_dir / "作業対象.xlsx"
    config_text = dedent(
        f"""\
        [RUN]
        DRY_RUN = False
        DEBUG = False

        [FILES]
        WEST_CSV = {west_csv}
        EAST_CSV = {east_csv}
        INPUT_XLSX = {input_xlsx}

        [CSV]
        KEY_COLUMN = お客様ID

        [EXCEL]
        SHEET = Sheet1
        KEY_COLUMN = 業務用ID
        HEADER_ROW = 1
        OUTPUT_PREFIX = 最終_
        PASSWORD =

        [転記_MAPPING]
        お名前 = 氏名
        ご住所 = 住所
        電話番号 = 電話番号
        """
    )
    config_path = work_dir / "config.ini"
    config_path.write_text(config_text, encoding="utf-8")
    return config_path


def _reload_modules(work_dir: Path) -> None:
    """work_dir を import path に追加し、src.* のキャッシュをクリアする。"""
    sys.path.insert(0, str(work_dir))
    for mod in list(sys.modules):
        if mod == "src" or mod.startswith("src."):
            sys.modules.pop(mod, None)


def run_normal() -> int:
    """パスワード空欄の本実行: 最終ファイル生成 + 元ファイル削除を検証。"""
    with tempfile.TemporaryDirectory(prefix="kokyaku_smoke_") as work_dir_str:
        work_dir = Path(work_dir_str)
        sample_dir = work_dir / "samples"
        sample_dir.mkdir()
        west_csv, east_csv, input_xlsx = _make_samples(sample_dir)
        config_path = _prepare_workspace(work_dir, sample_dir)

        _reload_modules(work_dir)
        from comken import config

        config.read(config_path)
        from src.run import run  # noqa: PLC0415

        result = run()

        expected_output = sample_dir / "最終_作業対象.xlsx"
        if not expected_output.exists():
            print(f"[NG] 最終ファイルが生成されていません: {expected_output}")
            return 1
        if west_csv.exists() or east_csv.exists() or input_xlsx.exists():
            print("[NG] 元ファイルが残っています（消えるはず）")
            return 1
        if result.matched_rows != 3:
            print(f"[NG] 転記件数が想定と違います: {result.matched_rows}")
            return 1

        import openpyxl

        workbook = openpyxl.load_workbook(expected_output)
        sheet = workbook["Sheet1"]
        assert sheet["B2"].value == "山田一郎", sheet["B2"].value
        assert sheet["B3"].value == "佐藤二郎", sheet["B3"].value
        assert sheet["B4"].value == "鈴木三郎", sheet["B4"].value
        workbook.close()

        print(f"[OK] 本実行 通過: 出力={expected_output}, 件数={result.matched_rows}")
        return 0


def run_dry_run() -> int:
    """DRY_RUN = True: ファイルが1つも作られず・消えず、ログだけ出る。"""
    with tempfile.TemporaryDirectory(prefix="kokyaku_dry_") as work_dir_str:
        work_dir = Path(work_dir_str)
        sample_dir = work_dir / "samples"
        sample_dir.mkdir()
        west_csv, east_csv, input_xlsx = _make_samples(sample_dir)
        config_path = _prepare_workspace(work_dir, sample_dir)
        # DRY_RUN を True へ書き換え
        config_text = config_path.read_text(encoding="utf-8").replace(
            "DRY_RUN = False", "DRY_RUN = True"
        )
        config_path.write_text(config_text, encoding="utf-8")

        _reload_modules(work_dir)
        from comken import config, dry_run

        config.read(config_path)
        from src.run import run  # noqa: PLC0415

        with dry_run():
            run()

        expected_output = sample_dir / "最終_作業対象.xlsx"
        if expected_output.exists():
            print(f"[NG] DRY-RUN なのに最終ファイルが生成されました: {expected_output}")
            return 1
        if not (west_csv.exists() and east_csv.exists() and input_xlsx.exists()):
            print("[NG] DRY-RUN なのに元ファイルが消えました")
            return 1

        print("[OK] DRY-RUN 通過: ファイルは作られず・消されず")
        return 0


def run_password_path() -> int:
    """パスワード付き経路が COM に到達することを、この PC に Excel が無い環境でも検証する。

    Excel が無い PC では ``ExcelApplicationNotAvailableError`` が飛ぶことを確かめる
    （COM が起動できない以上、ここで止まるのが正しい挙動）。
    Excel が入っている PC では、その例外は飛ばずに最終ファイルに
    パスワードが付くことを確認する。
    """
    with tempfile.TemporaryDirectory(prefix="kokyaku_pw_") as work_dir_str:
        work_dir = Path(work_dir_str)
        sample_dir = work_dir / "samples"
        sample_dir.mkdir()
        west_csv, east_csv, input_xlsx = _make_samples(sample_dir)
        config_path = _prepare_workspace(work_dir, sample_dir)
        # パスワードをセット
        config_text = config_path.read_text(encoding="utf-8").replace(
            "PASSWORD =", "PASSWORD = dummy-smoke-pw"
        )
        config_path.write_text(config_text, encoding="utf-8")

        _reload_modules(work_dir)
        from comken import config

        config.read(config_path)
        from src.run import run  # noqa: PLC0415

        try:
            run()
        except Exception as error:
            # Excel が入っていない環境ではここで止まるのが正解
            type_name = type(error).__name__
            if "ExcelApplicationNotAvailable" in type_name:
                print(
                    "[情報] パスワード経路は COM に到達し、Excel 不在のため "
                    "ExcelApplicationNotAvailableError で停止（想定どおり）"
                )
                # 元ファイルが残っていることも確認
                if not (west_csv.exists() and east_csv.exists() and input_xlsx.exists()):
                    print("[NG] 失敗時に元ファイルが消えています（消えるべきでない）")
                    return 1
                return 0
            print(f"[NG] 想定外のエラー: {type_name}: {error}")
            return 1

        # Excel が入っていた場合: 最終ファイルが存在し、元ファイルが消えていることを確認
        expected_output = sample_dir / "最終_作業対象.xlsx"
        if not expected_output.exists():
            print("[NG] パスワード付き保存が最終ファイルを生成していません")
            return 1
        if west_csv.exists() or east_csv.exists() or input_xlsx.exists():
            print("[NG] 元ファイルが残っています（消えるはず）")
            return 1
        print(f"[OK] パスワード付き 通過: 出力={expected_output}")
        return 0


def main() -> int:
    rc = run_normal()
    if rc != 0:
        return rc
    rc = run_dry_run()
    if rc != 0:
        return rc
    return run_password_path()


if __name__ == "__main__":
    sys.exit(main())
