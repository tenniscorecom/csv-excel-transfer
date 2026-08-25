# csv-excel-transfer エラー対応

| エラー | 対処 |
|---|---|
| `ConfigSectionNotFoundError` / `ConfigKeyNotFoundError` | `config.ini.example` と設定名を照合する |
| `ComkenError`（CSV にデータ行がありません） | 西・東 CSV にヘッダー下のデータがあるか確認する |
| `TableDuplicateKeyError` | 西・東を通じて「お客様ID」が一意になるよう修正する（旧 CsvReader の「踏んで上書き」は廃止）。`Table.concat()` で先に連結してから `Table.index()` で検査するため、ファイル内の重複も 2 ファイル間の重複も同じ例外で止まる |
| `ExcelColumnNotFoundError` | Excel 1行目の「業務用ID」と転記先列名を確認する |
| `TransferSourceColumnNotFoundError` | CSV の転記元列名と config.ini のマッピング左側を照合する |
| `ComkenError`（INPUT Excel にデータ行がありません） | `Sheet1` の見出し下にデータ行を追加する |
| `TransferDestinationMultipleMatchError` | 転記先（INPUT Excel）の「業務用ID」を一意にする |
| `ExcelApplicationNotAvailableError` | Excel がインストールされている環境で実行する |
| `ExcelSaveValidationError` | 出力先、空き容量、同名ファイルが開かれていないか確認する（元ファイルは保持される） |
| `FileDeletionError` | 保存済み出力を確認し、残った入力ファイルを管理者へ報告する |
| `ComkenError`（入力 CSV 2本と入力 Excel には、それぞれ別のファイルを指定） | 入力3ファイルがすべて別の実体を指すよう設定する |
| `ComkenError`（出力先が入力ファイルと同じ） | 出力フォルダまたは OUTPUT_PREFIX を変更する |
| `TransferMappingError` | `[TRANSFER_MAPPING]` に転記元列と転記先列を指定する |

解決しない場合は、エラー全文と入力ファイル名、実行時刻を管理者へ共有してください。