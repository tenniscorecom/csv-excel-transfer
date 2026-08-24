# csv-excel-transfer エラー対応

| エラー | 対処 |
|---|---|
| `ConfigSectionNotFoundError` / `ConfigKeyNotFoundError` | `config.ini.example` と設定名を照合する |
| `CSVNoDataRowsError` | 西・東 CSV にヘッダー下のデータがあるか確認する |
| `CSVRowDuplicateKeyError` | 西・東を通じて「お客様ID」が一意になるよう修正する |
| `TableDuplicateKeyError` | CSV ファイル内の「お客様ID」が一意になるよう修正する（旧 CsvReader の「踏んで上書き」は廃止） |
| `ExcelColumnNotFoundError` | Excel 1行目の「業務用ID」と転記先列名を確認する |
| `TransferSourceColumnNotFoundError` | CSV の転記元列名と config.ini のマッピング左側を照合する |
| `InputExcelNoDataError` | `Sheet1` の見出し下にデータ行を追加する |
| `TransferDestinationMultipleMatchError` | 転記先（INPUT Excel）の「業務用ID」を一意にする |
| `ExcelApplicationNotAvailableError` | Excel がインストールされている環境で実行する |
| `ExcelSaveNotCompletedError` | 出力先、空き容量、同名ファイルが開かれていないか確認する |
| `FileDeletionError` | 保存済み出力を確認し、残った入力ファイルを管理者へ報告する |

解決しない場合は、エラー全文と入力ファイル名、実行時刻を管理者へ共有してください。