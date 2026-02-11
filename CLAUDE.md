# Excel Deep Inspector

VBA + 数式 + UIを統合的に分析し、シニアエンジニアと同等以上の分析精度を実現するツール群。

## コンポーネント

### screenshot-tool/
PowerShellスクリーンショットキャプチャツール（ゼロインストール配布用）
- スマート領域検出（書式のみのセル除外）
- 固定行/列検出+キャプチャ
- 座標オーバーレイ（行番号・列番号描画）
- 1シート1ファイル、<10MB

### analysis-package/
PS「分析パッケージ」生成ツール
- VBA抽出 + セル参照パース
- 数式抽出 + 名前付き範囲
- コントロール→マクロ紐付け
- 構造情報JSON出力

## 技術制約
- PowerShell 5.1（Windows標準）のみ使用
- .NET Framework（System.Drawing）のみ使用（追加インストール不可）
- 社内AIに投入: 1シート1ファイル、<10MB
- 対象: .xlsm（VBA含むExcelファイル）

## 環境
- 配布先: 社内各部署PC（PowerShellのみ）
- 分析先: 社内AI（Azure ChatGPTラッパー）
- 開発: 私用PC（Claude Code MAX）/ 社内PC（Codex CLI）
