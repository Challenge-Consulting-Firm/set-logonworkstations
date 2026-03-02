# Set-LogonWorkstations

CSVファイルを基に、Active Directory ユーザーの `LogonWorkstations` 属性（ログオン可能なコンピューター）を一括設定する PowerShell スクリプトです。

## 機能

- CSV から対象ユーザーとコンピューター一覧を読み込み、`LogonWorkstations` を設定
- AD 上のユーザー・コンピューターの存在チェック
- タイムスタンプ付き実行ログの出力
- 実行結果サマリの表示（成功数・失敗数・スキップ数）
- 失敗した対象の一覧を CSV で出力

## 前提条件

- Windows Server または RSAT がインストールされた Windows PC
- ActiveDirectory PowerShell モジュール
- 対象ユーザー・コンピューターの属性を変更できる権限

## CSV 形式

ヘッダーは `SamAccountName,Computers` とし、複数 PC はセミコロン(`;`)で区切ります。

```csv
SamAccountName,Computers
taro.yamada,PC001;PC002;PC003
hanako.sato,NB-12345
```

## 使い方

```powershell
# 実行
.\Set-LogonWorkstations.ps1 -CsvPath .\users.csv

# ドライラン（変更内容の確認のみ）
.\Set-LogonWorkstations.ps1 -CsvPath .\users.csv -WhatIf

# ログ出力先を指定
.\Set-LogonWorkstations.ps1 -CsvPath .\users.csv -LogDir C:\Logs
```

### パラメーター

| パラメーター | 必須 | 説明 |
| --- | --- | --- |
| `-CsvPath` | はい | 入力 CSV ファイルのパス |
| `-LogDir` | いいえ | ログ出力先ディレクトリ（既定: スクリプトと同じフォルダ） |
| `-WhatIf` | いいえ | 実際には変更せず、処理内容を表示 |

## 出力ファイル

| ファイル | 説明 |
| --- | --- |
| `Set-LogonWorkstations_yyyyMMdd_HHmmss.log` | 実行ログ（毎回出力） |
| `Set-LogonWorkstations_Failures_yyyyMMdd_HHmmss.csv` | 失敗一覧（失敗時のみ出力） |

### 失敗一覧 CSV の形式

```csv
SamAccountName,Computer,Reason
taro.yamada,PC999,コンピューターが存在しない
jiro.tanaka,PC001,ユーザーが存在しない
```

## 実行結果サマリの例

```text
========== 実行結果サマリ ==========
処理対象ユーザー数 : 10
成功 (ユーザー/PC) : 15
失敗 (ユーザー/PC) : 2
スキップ (空行)    : 0
=====================================
```
