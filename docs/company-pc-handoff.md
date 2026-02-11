# Excel Deep Inspector - 社内PC実行ハンドオフガイド

## 概要

自宅PCで開発・テスト済みのAnalysis Package v2ツール群を社内PCで実行するための手順。
全スクリプトはPowerShell 5.1で動作し、追加インストールは不要。

## 前提環境

- **OS**: Windows 10/11
- **PowerShell**: 5.1（標準搭載）
- **Excel**: COM操作可能なライセンス済みExcel
- **Git**: GitHub リポジトリへのアクセス（コード取得用）

## リポジトリ取得

```powershell
git clone https://github.com/{org}/excel-deep-inspector.git
cd excel-deep-inspector
```

## ファイル構成

```
analysis-package/
  Generate-AnalysisPackage.ps1   # メインスクリプト（COM使用）
  Parse-VBAModules.ps1           # v2パーサー（COM不要）
  Build-CrossReference.ps1       # v2クロスリファレンス（COM不要）
```

## 実行方法

### A. 統合実行（推奨）

Generate-AnalysisPackage.ps1がv1抽出→v2解析→クロスリファレンスを順次実行する。

```powershell
powershell.exe -ExecutionPolicy Bypass -File "analysis-package\Generate-AnalysisPackage.ps1" -FilePath "C:\path\to\target.xlsm"
```

出力先: `C:\path\to\target_analysis_package\` に自動作成。

### B. 個別実行（デバッグ・再解析用）

v1出力（.basファイル）が既にある場合、v2パーサーのみ再実行できる。

```powershell
# 1. v2パーサー単独実行
powershell.exe -ExecutionPolicy Bypass -File "analysis-package\Parse-VBAModules.ps1" `
  -InputDir "C:\path\to\target_analysis_package\01_VBA" `
  -OutputDir "C:\path\to\target_analysis_package\01_VBA\v2_output"

# 2. クロスリファレンス生成
powershell.exe -ExecutionPolicy Bypass -File "analysis-package\Build-CrossReference.ps1" `
  -AnalysisDir "C:\path\to\target_analysis_package"
```

## 出力ファイル一覧

### v1出力（COM必要）
| パス | 内容 |
|------|------|
| `01_VBA/*.bas` | VBAソースコード（ヘッダー付き） |
| `01_VBA/cell_references.json` | セル参照（v1: 重複あり） |
| `01_VBA/procedures.json` | プロシージャ一覧 |
| `02_formulas/*_formulas.json` | シート別数式マップ |
| `02_formulas/named_ranges.json` | 名前付き範囲 |
| `04_structure/sheet_list.json` | シート一覧 |
| `04_structure/controls.json` | コントロール→マクロ紐付け |
| `04_structure/conditional_formats.json` | 条件付き書式 |

### v2出力（COM不要）
| パス | 内容 |
|------|------|
| `01_VBA/v2_output/cell_references_v2.json` | セル参照v2（重複排除・コンテキスト付き） |
| `01_VBA/v2_output/table_references.json` | ListObject/ListColumn参照 |
| `01_VBA/v2_output/event_triggers.json` | イベントトリガー（WS/ActiveX/UserForm） |
| `01_VBA/v2_output/form_calls.json` | フォーム呼び出しチェーン |
| `01_VBA/v2_output/sheet_codenames.json` | CodeName→DisplayNameマッピング（COM生成） |

### v2クロスリファレンス（COM不要）
| パス | 内容 |
|------|------|
| `05_cross_reference/ui_to_vba.json` | UI操作→VBA→セル影響の完全チェーン |
| `05_cross_reference/data_flow.json` | テーブル列/セル参照のデータフロー |

## 社内AIへの投入手順

### 推奨: コンテキスト最大化パターン

1. `00_overview.md` を最初に投入（全体把握）
2. 調査対象のシートに関連する以下を追加投入:
   - `05_cross_reference/ui_to_vba.json`（UIからの影響パス）
   - `05_cross_reference/data_flow.json`（データの流れ）
   - 関連する `.bas` ファイル
3. 質問例:
   - 「ボタンXを押すとどのセルが変更されますか?」
   - 「テーブルYのZ列はどのプロシージャから書き込まれますか?」
   - 「シートAを修正する場合の影響範囲は?」

### 軽量パターン（トークン節約）

1. `00_overview.md` のみ投入
2. 質問に応じて必要なJSONを追加

## トラブルシューティング

### VBAProject access may be disabled
- Excel のトラストセンター → マクロ設定 → 「VBAプロジェクトオブジェクトモデルへのアクセスを信頼する」を有効化

### Parse-VBAModules.ps1 で .Count エラー
- PS5.1 StrictMode互換済み。エラーが出る場合はPSバージョン確認: `$PSVersionTable.PSVersion`

### Build-CrossReference.ps1 で controls.json not found
- Generate-AnalysisPackage.ps1 を先に実行してv1出力を生成する必要がある
- 個別実行時は `-AnalysisDir` パスに `04_structure/controls.json` が存在することを確認

## Codex CLI用プロンプトテンプレート

社内PCでCodex CLIを使って機能追加・修正を行う場合のプロンプト:

```
Excel Deep Inspector の analysis-package/ 配下のPowerShellスクリプトを修正してください。

対象ファイル:
- Parse-VBAModules.ps1: VBAテキスト解析エンジン
- Build-CrossReference.ps1: クロスリファレンス生成
- Generate-AnalysisPackage.ps1: メイン統合スクリプト

制約:
- PowerShell 5.1 互換必須（Set-StrictMode -Version Latest）
- [PSCustomObject]@{} 内に Where-Object {} / ForEach-Object {} / if () {} を書かない（パーサー誤認）
- 配列は常に @() で囲む（.Count がスカラーで失敗するため）
- Excel COM操作はGenerate-AnalysisPackage.ps1内のみ。Parse/Buildスクリプトは COM不要
```
