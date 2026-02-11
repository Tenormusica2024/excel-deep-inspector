<#
.SYNOPSIS
    Excel VBA Analysis Package Generator
.DESCRIPTION
    Excelファイル(.xlsm/.xlsx)から以下の情報を一括抽出し、
    AIによる分析が可能な「分析パッケージ」フォルダを生成する:
      - VBAソースコード抽出 + セル参照の正規表現パース
      - 全シートの数式マッピング + 名前付き範囲
      - コントロール(ボタン等)→マクロ紐付け
      - シート構造情報（固定行/列、条件付き書式）
      - AI向け分析指示書（概要Markdown）

    配布先のPCではPowerShell 5.1のみで動作（追加インストール不要）。
    出力は社内AI（Azure ChatGPTラッパー等）に投入可能な形式。

.PARAMETER FilePath
    対象のExcelファイルパス
.PARAMETER OutputDir
    分析パッケージ出力先（未指定時はファイルと同階層に作成）
.PARAMETER IncludeScreenshots
    スクリーンショットツールも同時実行するか（デフォルト: false）
.EXAMPLE
    .\Generate-AnalysisPackage.ps1 -FilePath "C:\tools\estimate.xlsm"
#>
param(
    [Parameter(Mandatory=$true)]
    [string]$FilePath,

    [Parameter(Mandatory=$false)]
    [string]$OutputDir,

    [Parameter(Mandatory=$false)]
    [switch]$IncludeScreenshots
)

# ============================================================
# 初期化
# ============================================================
$ErrorActionPreference = "Continue"

$resolvedPath = Resolve-Path $FilePath -ErrorAction SilentlyContinue
if ($null -eq $resolvedPath) {
    Write-Host "[ERROR] File not found: $FilePath" -ForegroundColor Red
    exit 1
}
$FilePath = $resolvedPath.Path
$fileName = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
$fileExt = [System.IO.Path]::GetExtension($FilePath).ToLower()

if (-not $OutputDir) {
    $parentDir = Split-Path $FilePath -Parent
    $OutputDir = Join-Path $parentDir "${fileName}_analysis_package"
}

# サブディレクトリ作成
$dirs = @{
    VBA            = Join-Path $OutputDir "01_VBA"
    Formulas       = Join-Path $OutputDir "02_formulas"
    Screenshots    = Join-Path $OutputDir "03_screenshots"
    Structure      = Join-Path $OutputDir "04_structure"
    CrossReference = Join-Path $OutputDir "05_cross_reference"
    V2Output       = Join-Path $OutputDir "01_VBA\v2_output"
}
foreach ($d in $dirs.Values) {
    if (-not (Test-Path $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
}

Write-Host "================================================================" -ForegroundColor Cyan
Write-Host " Excel Analysis Package Generator" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "Input : $FilePath"
Write-Host "Output: $OutputDir"
Write-Host ""

# ============================================================
# VBAセル参照パース用の正規表現パターン
# ============================================================
# L1: 静的解析（正規表現で抽出可能なパターン）
$cellRefPatterns = @(
    # Range("A1") / Range("A1:B10") / Range("Sheet1!A1")
    @{ Name = "Range_Literal";  Pattern = 'Range\(\s*"([^"]+)"\s*\)';  Group = 1 },
    # Cells(row, col) - 数値リテラルのみ
    @{ Name = "Cells_Literal";  Pattern = 'Cells\(\s*(\d+)\s*,\s*(\d+)\s*\)';  Group = 0 },
    # [A1] / [Sheet1!A1:B10] - ブラケット記法
    @{ Name = "Bracket_Ref";    Pattern = '\[([A-Z]+\d+(?::[A-Z]+\d+)?)\]'; Group = 1 },
    # Worksheets("name") / Sheets("name")
    @{ Name = "Sheet_Ref";      Pattern = '(?:Worksheets|Sheets)\(\s*"([^"]+)"\s*\)'; Group = 1 },
    # .Offset(row, col)
    @{ Name = "Offset_Ref";     Pattern = '\.Offset\(\s*(-?\d+)\s*,\s*(-?\d+)\s*\)'; Group = 0 },
    # 名前付き範囲参照: Range("name") - 英数字とアンダースコアのみの名前
    @{ Name = "Named_Range";    Pattern = 'Range\(\s*"([A-Za-z_]\w+)"\s*\)'; Group = 1 },
    # ActiveSheet / ActiveCell
    @{ Name = "Active_Ref";     Pattern = '(ActiveSheet|ActiveCell|ActiveWorkbook)'; Group = 1 },
    # Rows / Columns 参照
    @{ Name = "RowCol_Ref";     Pattern = '(?:\.Rows|\.Columns)\(\s*"?([^")\s]+)"?\s*\)'; Group = 1 }
)

# ============================================================
# ヘルパー関数
# ============================================================

function Convert-ColumnNumberToLetter {
    param([int]$ColumnNumber)
    $result = ""
    while ($ColumnNumber -gt 0) {
        $ColumnNumber--
        $result = [char](65 + ($ColumnNumber % 26)) + $result
        $ColumnNumber = [Math]::Floor($ColumnNumber / 26)
    }
    return $result
}

function Extract-VBACellReferences {
    <#
    .DESCRIPTION
        VBAソースコードからセル参照パターンを正規表現で抽出する（L1解析）。
        各マッチに行番号、パターン名、マッチ文字列、抽出グループを付与。
    #>
    param(
        [string]$VBACode,
        [string]$ModuleName
    )

    $references = @()
    $lines = $VBACode -split "`r?`n"

    for ($lineNum = 0; $lineNum -lt $lines.Count; $lineNum++) {
        $line = $lines[$lineNum]

        # コメント行はスキップ
        $trimmed = $line.TrimStart()
        if ($trimmed.StartsWith("'") -or $trimmed.StartsWith("Rem ")) { continue }

        foreach ($pat in $cellRefPatterns) {
            $matches = [regex]::Matches($line, $pat.Pattern)
            foreach ($m in $matches) {
                $ref = @{
                    Module    = $ModuleName
                    Line      = $lineNum + 1
                    Pattern   = $pat.Name
                    Match     = $m.Value
                    Context   = $line.Trim()
                }
                # グループ値を抽出
                if ($pat.Group -gt 0 -and $m.Groups.Count -gt $pat.Group) {
                    $ref.Extracted = $m.Groups[$pat.Group].Value
                }
                elseif ($pat.Group -eq 0) {
                    # Cells(r,c)等：全グループを連結
                    $groups = @()
                    for ($g = 1; $g -lt $m.Groups.Count; $g++) {
                        $groups += $m.Groups[$g].Value
                    }
                    $ref.Extracted = $groups -join ","
                }
                $references += $ref
            }
        }
    }

    return ,$references
}

function Extract-SubProcedures {
    <#
    .DESCRIPTION
        VBAコードからSub/Functionプロシージャの一覧を抽出する。
        名前、種別、アクセス修飾子、引数、開始行を返す。
    #>
    param(
        [string]$VBACode,
        [string]$ModuleName
    )

    $procedures = @()
    $lines = $VBACode -split "`r?`n"

    for ($i = 0; $i -lt $lines.Count; $i++) {
        $line = $lines[$i].Trim()

        # Sub/Function宣言を検出
        if ($line -match '^\s*(Public\s+|Private\s+)?(Sub|Function)\s+(\w+)\s*\(([^)]*)\)') {
            $procedures += @{
                Module    = $ModuleName
                Access    = if ($Matches[1]) { $Matches[1].Trim() } else { "Public" }
                Type      = $Matches[2]
                Name      = $Matches[3]
                Arguments = $Matches[4].Trim()
                Line      = $i + 1
            }
        }
    }

    return ,$procedures
}

# ============================================================
# メイン処理
# ============================================================
$excel = $null
$wb = $null
$sheetCodeNames = @()

try {
    Write-Host "[1/6] Excel COM starting..." -ForegroundColor Yellow
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    Write-Host "[2/6] Opening: $(Split-Path $FilePath -Leaf)" -ForegroundColor Yellow
    $wb = $excel.Workbooks.Open($FilePath, $false, $true)  # ReadOnly

    $sheetCount = $wb.Worksheets.Count

    # ===========================================================
    # [A] VBAコード抽出
    # ===========================================================
    Write-Host "[3/6] VBA extraction..." -ForegroundColor Yellow

    $allVBAReferences = @()
    $allProcedures = @()
    $vbaModuleCount = 0

    try {
        $vbProject = $wb.VBProject
        $components = $vbProject.VBComponents

        foreach ($comp in $components) {
            $codeModule = $comp.CodeModule
            $lineCount = $codeModule.CountOfLines

            if ($lineCount -gt 0) {
                $vbaCode = $codeModule.Lines(1, $lineCount)
                $moduleName = $comp.Name

                # コンポーネント種別
                $compType = switch ($comp.Type) {
                    1 { "StandardModule" }
                    2 { "ClassModule" }
                    3 { "UserForm" }
                    100 { "Document" }  # ThisWorkbook / Sheet
                    default { "Unknown($($comp.Type))" }
                }

                # VBAソース保存
                $safeModuleName = $moduleName -replace '[\\/:*?"<>|]', '_'
                $vbaFilePath = Join-Path $dirs.VBA "${safeModuleName}.bas"
                # ヘッダーコメント付きで保存
                $header = "' Module: $moduleName`r`n' Type: $compType`r`n' Lines: $lineCount`r`n' ---`r`n"
                ($header + $vbaCode) | Set-Content -Path $vbaFilePath -Encoding UTF8
                $vbaModuleCount++

                Write-Host "    $compType : $moduleName ($lineCount lines)"

                # セル参照抽出
                $refs = Extract-VBACellReferences -VBACode $vbaCode -ModuleName $moduleName
                if ($refs.Count -gt 0) {
                    $allVBAReferences += $refs
                }

                # プロシージャ抽出
                $procs = Extract-SubProcedures -VBACode $vbaCode -ModuleName $moduleName
                if ($procs.Count -gt 0) {
                    $allProcedures += $procs
                }
            }
        }

        Write-Host "    VBA modules: $vbaModuleCount, Cell refs: $($allVBAReferences.Count), Procedures: $($allProcedures.Count)" -ForegroundColor Green
    }
    catch {
        Write-Host "    [WARN] VBA extraction failed: $($_.Exception.Message)" -ForegroundColor Yellow
        Write-Host "    -> VBProject access may be disabled" -ForegroundColor Yellow
    }

    # セル参照JSON保存
    $cellRefJson = @{
        TotalReferences = $allVBAReferences.Count
        # Group-ObjectのPropertyにスクリプトブロックを使用（ハッシュテーブルのキーアクセスを確実にする）
        Patterns = ($allVBAReferences | Group-Object -Property { $_.Pattern } | ForEach-Object {
            @{ Pattern = $_.Name; Count = $_.Count }
        })
        References = $allVBAReferences
    }
    $cellRefJson | ConvertTo-Json -Depth 5 | Set-Content -Path (Join-Path $dirs.VBA "cell_references.json") -Encoding UTF8

    # プロシージャJSON保存
    $procJson = @{
        TotalProcedures = $allProcedures.Count
        Procedures = $allProcedures
    }
    $procJson | ConvertTo-Json -Depth 5 | Set-Content -Path (Join-Path $dirs.VBA "procedures.json") -Encoding UTF8

    # ===========================================================
    # [B] 数式抽出 + 名前付き範囲
    # ===========================================================
    Write-Host "[4/6] Formula extraction..." -ForegroundColor Yellow

    $allNamedRanges = @()
    try {
        foreach ($name in $wb.Names) {
            $allNamedRanges += @{
                Name      = $name.Name
                RefersTo  = $name.RefersTo
                Visible   = $name.Visible
            }
        }
    }
    catch {
        Write-Host "    [WARN] Named range extraction failed: $($_.Exception.Message)" -ForegroundColor Yellow
    }

    # 名前付き範囲JSON保存
    @{
        Count = $allNamedRanges.Count
        NamedRanges = $allNamedRanges
    } | ConvertTo-Json -Depth 3 | Set-Content -Path (Join-Path $dirs.Formulas "named_ranges.json") -Encoding UTF8
    Write-Host "    Named ranges: $($allNamedRanges.Count)"

    # 各シートの数式抽出
    for ($sIdx = 1; $sIdx -le $sheetCount; $sIdx++) {
        $sheet = $wb.Worksheets.Item($sIdx)
        $sheetName = $sheet.Name
        $safeSheetName = $sheetName -replace '[\\/:*?"<>|]', '_'

        Write-Host "    Sheet [$sheetName]..."

        $usedRange = $sheet.UsedRange
        if ($null -eq $usedRange) {
            Write-Host "      -> empty (skipped)"
            continue
        }

        $startRow = $usedRange.Row
        $startCol = $usedRange.Column
        $rowCount = $usedRange.Rows.Count
        $colCount = $usedRange.Columns.Count

        # 値と数式を一括取得
        $values = $usedRange.Value2
        $formulas = $usedRange.Formula

        $isSingle = ($rowCount -eq 1 -and $colCount -eq 1)
        $formulaCells = @()

        if ($isSingle) {
            $formula = "$formulas"
            if ($formula.StartsWith("=")) {
                $addr = (Convert-ColumnNumberToLetter $startCol) + "$startRow"
                $formulaCells += @{
                    Address = $addr
                    Formula = $formula
                    Value   = $values
                }
            }
        }
        else {
            for ($r = 1; $r -le $rowCount; $r++) {
                for ($c = 1; $c -le $colCount; $c++) {
                    $formula = $formulas[$r, $c]
                    if ($null -ne $formula -and "$formula".StartsWith("=")) {
                        $absRow = $startRow + $r - 1
                        $absCol = $startCol + $c - 1
                        $addr = (Convert-ColumnNumberToLetter $absCol) + "$absRow"
                        $formulaCells += @{
                            Address = $addr
                            Formula = "$formula"
                            Value   = $values[$r, $c]
                        }
                    }
                }
            }
        }

        # シート別数式JSON保存
        @{
            SheetName = $sheetName
            SheetIndex = $sIdx
            UsedRange = @{
                Start = "R${startRow}C${startCol}"
                Rows = $rowCount
                Cols = $colCount
            }
            FormulaCount = $formulaCells.Count
            Formulas = $formulaCells
        } | ConvertTo-Json -Depth 4 | Set-Content -Path (Join-Path $dirs.Formulas "${safeSheetName}_formulas.json") -Encoding UTF8

        Write-Host "      Formulas: $($formulaCells.Count)"
    }

    # ===========================================================
    # [C] 構造情報（シート一覧、コントロール、条件付き書式）
    # ===========================================================
    Write-Host "[5/6] Structure info extraction..." -ForegroundColor Yellow

    $sheetInfos = @()
    $allControls = @()
    $allCondFormats = @()

    for ($sIdx = 1; $sIdx -le $sheetCount; $sIdx++) {
        $sheet = $wb.Worksheets.Item($sIdx)
        $sheetName = $sheet.Name

        # シートをアクティブ化（固定行取得に必要）
        $sheet.Activate()
        Start-Sleep -Milliseconds 100

        # 固定行/列
        $frozenRows = 0
        $frozenCols = 0
        try {
            $frozenRows = [int]$excel.ActiveWindow.SplitRow
            $frozenCols = [int]$excel.ActiveWindow.SplitColumn
        }
        catch {}

        $sheetInfos += @{
            Name = $sheetName
            Index = $sIdx
            FrozenRows = $frozenRows
            FrozenCols = $frozenCols
            Visible = $sheet.Visible
        }

        # コントロール（ボタン等）→マクロ紐付け
        try {
            foreach ($shape in $sheet.Shapes) {
                $ctrl = @{
                    Sheet    = $sheetName
                    Name     = $shape.Name
                    Type     = $shape.Type
                    OnAction = ""
                    Left     = [Math]::Round($shape.Left, 1)
                    Top      = [Math]::Round($shape.Top, 1)
                    Width    = [Math]::Round($shape.Width, 1)
                    Height   = [Math]::Round($shape.Height, 1)
                }
                try { $ctrl.OnAction = $shape.OnAction } catch {}
                $allControls += $ctrl
            }
        }
        catch {}

        # 条件付き書式
        try {
            $usedRange = $sheet.UsedRange
            if ($null -ne $usedRange) {
                foreach ($cf in $usedRange.FormatConditions) {
                    $condInfo = @{
                        Sheet     = $sheetName
                        Type      = $cf.Type
                        Priority  = $cf.Priority
                        AppliesTo = ""
                    }
                    # COM parameterized propertyは引数付きで呼ぶ（引数なしだとプロパティメタデータが返る）
                    try { $condInfo.AppliesTo = $cf.AppliesTo.Address($true, $true) } catch {}
                    try { $condInfo.Formula1 = $cf.Formula1 } catch {}
                    try { $condInfo.Formula2 = $cf.Formula2 } catch {}
                    try { $condInfo.Operator = $cf.Operator } catch {}
                    $allCondFormats += $condInfo
                }
            }
        }
        catch {}

        Write-Host "    [$sheetName] Frozen:$frozenRows/$frozenCols Controls:$($allControls.Count) CondFmt:$($allCondFormats.Count)"
    }

    # 構造情報JSON保存
    @{
        FileName = Split-Path $FilePath -Leaf
        SheetCount = $sheetCount
        Sheets = $sheetInfos
    } | ConvertTo-Json -Depth 3 | Set-Content -Path (Join-Path $dirs.Structure "sheet_list.json") -Encoding UTF8

    @{
        TotalControls = $allControls.Count
        Controls = $allControls
    } | ConvertTo-Json -Depth 3 | Set-Content -Path (Join-Path $dirs.Structure "controls.json") -Encoding UTF8

    @{
        TotalCondFormats = $allCondFormats.Count
        ConditionalFormats = $allCondFormats
    } | ConvertTo-Json -Depth 3 | Set-Content -Path (Join-Path $dirs.Structure "conditional_formats.json") -Encoding UTF8

    # ===========================================================
    # [C2] Sheet CodeName → DisplayNameマッピング
    # VBA内の shDash.Range("...") をシート表示名に解決するための対応表
    # ===========================================================
    Write-Host "    Extracting sheet codenames..." -ForegroundColor Gray
    $sheetCodeNames = @()
    for ($sIdx = 1; $sIdx -le $sheetCount; $sIdx++) {
        $sheet = $wb.Worksheets.Item($sIdx)
        $codeName = ""
        try { $codeName = $sheet.CodeName } catch {}
        $sheetCodeNames += @{
            CodeName    = $codeName
            DisplayName = $sheet.Name
            Index       = $sIdx
            Visible     = [int]$sheet.Visible
        }
    }
    # v2_outputディレクトリにも保存（Parse-VBAModules.ps1が参照）
    if (-not (Test-Path $dirs.V2Output)) { New-Item -ItemType Directory -Path $dirs.V2Output -Force | Out-Null }
    @{
        TotalSheets    = $sheetCodeNames.Count
        SheetCodeNames = $sheetCodeNames
    } | ConvertTo-Json -Depth 3 | Set-Content -Path (Join-Path $dirs.V2Output "sheet_codenames.json") -Encoding UTF8
    Write-Host "    Sheet codenames: $($sheetCodeNames.Count)" -ForegroundColor Green

    # ===========================================================
    # [D] AI向け分析指示書（概要Markdown）
    # ===========================================================
    Write-Host "[6/6] Generating overview..." -ForegroundColor Yellow

    # テーブル行を事前構築（here-string内のforeachは改行が入らないため）
    $sheetTableRows = ($sheetInfos | ForEach-Object {
        "| $($_.Index) | $($_.Name) | $($_.FrozenRows) | $($_.FrozenCols) |"
    }) -join "`r`n"

    $procTableRows = ($allProcedures | ForEach-Object {
        "| $($_.Module) | $($_.Type) | $($_.Name) | $($_.Line) |"
    }) -join "`r`n"

    $refSummaryLines = (($allVBAReferences | Group-Object -Property { $_.Pattern }) | ForEach-Object {
        "- **$($_.Name)**: $($_.Count) refs"
    }) -join "`r`n"

    $overviewContent = @"
# Analysis Package: $(Split-Path $FilePath -Leaf)

Generated: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
Tool: Excel Analysis Package Generator v1.0

## File Overview

- **File**: $(Split-Path $FilePath -Leaf)
- **Format**: $fileExt
- **Sheets**: $sheetCount
- **VBA Modules**: $vbaModuleCount
- **Named Ranges**: $($allNamedRanges.Count)
- **Controls (Buttons etc.)**: $($allControls.Count)
- **Conditional Formats**: $($allCondFormats.Count)

## Sheet List

| # | Name | Frozen Rows | Frozen Cols |
|---|------|-------------|-------------|
$sheetTableRows

## VBA Procedures

| Module | Type | Name | Line |
|--------|------|------|------|
$procTableRows

## VBA Cell Reference Summary

Total references found: $($allVBAReferences.Count)

$refSummaryLines

## Analysis Instructions

This package contains structured data extracted from the Excel file above.
Please analyze the following aspects:

1. **VBA Code Analysis**: Review ``01_VBA/*.bas`` files for business logic and cell references
2. **Formula Dependencies**: Check ``02_formulas/*_formulas.json`` for inter-cell and cross-sheet dependencies
3. **UI-Data Binding**: Use ``05_cross_reference/ui_to_vba.json`` for complete UI trigger-to-cell impact chains
4. **Data Flow Tracing**: Use ``05_cross_reference/data_flow.json`` to trace table column read/write patterns
5. **Impact Assessment**: For any proposed modification, trace all affected cells, formulas, and VBA code paths
6. **Risk Identification**: Flag circular references, hardcoded values, and undocumented dependencies
7. **Sheet CodeName Resolution**: Use ``01_VBA/v2_output/sheet_codenames.json`` to resolve VBA codenames to display names

## Package Structure

```
01_VBA/              - VBA source code + cell reference analysis (v1)
01_VBA/v2_output/    - v2: cell refs, table refs, events, form calls, sheet codenames
02_formulas/         - Formula maps per sheet + named ranges
03_screenshots/      - UI screenshots with coordinate overlays
04_structure/        - Sheet list, controls, conditional formats
05_cross_reference/  - v2: UI-to-VBA chains, data flow maps
```
"@

    $overviewContent | Set-Content -Path (Join-Path $OutputDir "00_overview.md") -Encoding UTF8

    Write-Host ""
    Write-Host "Analysis package generated successfully!" -ForegroundColor Green
    Write-Host "  VBA Modules  : $vbaModuleCount"
    Write-Host "  Cell Refs    : $($allVBAReferences.Count)"
    Write-Host "  Procedures   : $($allProcedures.Count)"
    Write-Host "  Named Ranges : $($allNamedRanges.Count)"
    Write-Host "  Controls     : $($allControls.Count)"
    Write-Host "  Cond. Formats: $($allCondFormats.Count)"

    # スクリーンショット実行
    if ($IncludeScreenshots) {
        Write-Host ""
        Write-Host "Running screenshot capture..." -ForegroundColor Yellow
        $screenshotScript = Join-Path (Split-Path $PSScriptRoot) "screenshot-tool\Excel-ScreenshotCapture.ps1"
        if (Test-Path $screenshotScript) {
            & $screenshotScript -FilePath $FilePath -OutputDir $dirs.Screenshots
        }
        else {
            Write-Host "[WARN] Screenshot script not found: $screenshotScript" -ForegroundColor Yellow
        }
    }
}
catch {
    Write-Host "[FATAL] $($_.Exception.Message)" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor DarkRed
}
finally {
    if ($null -ne $wb) { try { $wb.Close($false) } catch {} }
    if ($null -ne $excel) { try { $excel.Quit() } catch {} }
    if ($null -ne $wb) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null }
    if ($null -ne $excel) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    Write-Host ""
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host " Output: $OutputDir" -ForegroundColor Cyan
    Write-Host "================================================================" -ForegroundColor Cyan
}

# ===========================================================
# [POST-COM] v2パーサー実行（COM不要・テキスト解析のみ）
# .basファイルが存在すれば、COM終了後でも実行可能
# ===========================================================
$basFiles = @(Get-ChildItem -Path $dirs.VBA -Filter "*.bas" -ErrorAction SilentlyContinue)
if ($basFiles.Count -gt 0) {
    Write-Host ""
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host " v2 Analysis (COM-independent)" -ForegroundColor Cyan
    Write-Host "================================================================" -ForegroundColor Cyan

    # Parse-VBAModules.ps1: セル参照v2 + テーブル参照 + イベント + フォーム呼び出し
    $parseScript = Join-Path $PSScriptRoot "Parse-VBAModules.ps1"
    if (Test-Path $parseScript) {
        Write-Host "[v2-1] Running Parse-VBAModules.ps1..." -ForegroundColor Yellow
        try {
            # SheetCodeNameMapを構築（COM抽出結果から）
            $sheetCodeNameHashtable = @{}
            foreach ($sc in $sheetCodeNames) {
                if ($sc.CodeName) {
                    $sheetCodeNameHashtable[$sc.CodeName] = $sc.DisplayName
                }
            }
            $parseParams = @{
                InputDir  = $dirs.VBA
                OutputDir = $dirs.V2Output
            }
            if ($sheetCodeNameHashtable.Count -gt 0) {
                $parseParams.SheetCodeNameMap = $sheetCodeNameHashtable
            }
            & $parseScript @parseParams
        }
        catch {
            Write-Host "[WARN] Parse-VBAModules.ps1 failed: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
    else {
        Write-Host "[WARN] Parse-VBAModules.ps1 not found: $parseScript" -ForegroundColor Yellow
    }

    # Build-CrossReference.ps1: UI→VBAチェーン + データフロー
    $crossRefScript = Join-Path $PSScriptRoot "Build-CrossReference.ps1"
    if (Test-Path $crossRefScript) {
        Write-Host "[v2-2] Running Build-CrossReference.ps1..." -ForegroundColor Yellow
        try {
            & $crossRefScript -AnalysisDir $OutputDir
        }
        catch {
            Write-Host "[WARN] Build-CrossReference.ps1 failed: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
    else {
        Write-Host "[WARN] Build-CrossReference.ps1 not found: $crossRefScript" -ForegroundColor Yellow
    }

    Write-Host ""
    Write-Host "v2 analysis complete." -ForegroundColor Green
}
else {
    Write-Host ""
    Write-Host "[INFO] No .bas files found - skipping v2 analysis" -ForegroundColor Gray
}
