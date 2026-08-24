# csv-excel-transfer

西日本・東日本の顧客 CSV を統合し、INPUT Excel の「業務用ID」に一致する顧客情報を
転記する業務ツールです。出力保存が成功した場合だけ、元の CSV 2 ファイルと Excel を
削除します。

## 設定

- `[FILES]`: `OUTPUT_EXCEL_FOLDER`、`INPUT_EXCEL_FOLDER`、`INPUT_CSV_FOLDER`
- `[EXCEL]`: `INPUT_NAME`、`OUTPUT_PREFIX`、`SHEET_NAME`、`HEADER_ROW`、`KEY_COLUMN`、`READ_PASSWORD`、`WRITE_PASSWORD`
- `[CSV]`: `WEST`、`EAST`、`KEY_COLUMN`
- `[TRANSFER_MAPPING]`: CSV 列名 = Excel 列名

`WEST` / `EAST` は `INPUT_CSV_FOLDER` 配下、`INPUT_NAME` は
`INPUT_EXCEL_FOLDER` 配下のファイル名です。出力先は
`OUTPUT_EXCEL_FOLDER / (OUTPUT_PREFIX + INPUT_NAME)` です。

## 実行と検証

```powershell
実行.bat
python -m ruff check .
python -m ruff format --check .
python -m pytest tests
python smoke_test.py
```

詳細は [使い方](docs/使い方.md)、[仕様書](docs/仕様書.md)、
[エラー対応](docs/ERRORS.md) を参照してください。
