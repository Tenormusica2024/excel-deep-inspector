<#
.SYNOPSIS
    Excel シート スクリーンショット キャプチャツール（座標オーバーレイ付き）
.DESCRIPTION
    Excelファイルの各シートをスクリーンショットとして保存する。
    - スマート領域検出: 値/数式のあるセルのみを撮影対象とし、書式のみのセルを除外
    - 固定行/列検出: ウィンドウ枠の固定（Freeze Panes）を検出し、ヘッダー領域を識別
    - 座標オーバーレイ: 行番号・列名（A,B,C...）を画像上に描画
    - 1シート1ファイル: 社内AI投入用（<10MB）
.PARAMETER FilePath
    対象のExcelファイルパス（.xlsx / .xlsm）
.PARAMETER OutputDir
    スクリーンショット出力先ディレクトリ（未指定時はファイルと同階層に作成）
.PARAMETER DPI
    キャプチャ解像度のスケーリング（デフォルト: 1.0 = 画面解像度）
.EXAMPLE
    .\Excel-ScreenshotCapture.ps1 -FilePath "C:\path\to\workbook.xlsm"
    .\Excel-ScreenshotCapture.ps1 -FilePath "C:\path\to\workbook.xlsm" -OutputDir "C:\output"
#>
param(
    [Parameter(Mandatory=$true)]
    [string]$FilePath,

    [Parameter(Mandatory=$false)]
    [string]$OutputDir,

    [Parameter(Mandatory=$false)]
    [double]$DPI = 1.0
)

# ============================================================
# アセンブリ読み込み
# ============================================================
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms

# ============================================================
# 定数
# ============================================================
# Excel CopyPicture の定数
$xlScreen = 1           # XlPictureAppearance.xlScreen
$xlBitmap = 2           # XlCopyPictureFormat.xlBitmap
$xlPrinter = 2          # XlPictureAppearance.xlPrinter

# 座標オーバーレイの設定
$OVERLAY_MARGIN_LEFT = 40     # 行番号用の左マージン（px）
$OVERLAY_MARGIN_TOP = 25      # 列名用の上マージン（px）
$OVERLAY_FONT_NAME = "Consolas"
$OVERLAY_FONT_SIZE = 9
$OVERLAY_BG_COLOR = [System.Drawing.Color]::FromArgb(240, 240, 240)   # ヘッダー背景
$OVERLAY_BORDER_COLOR = [System.Drawing.Color]::FromArgb(180, 180, 180)
$OVERLAY_TEXT_COLOR = [System.Drawing.Color]::FromArgb(60, 60, 60)

# ファイルサイズ上限: 10MB
$MAX_FILE_SIZE_BYTES = 10 * 1024 * 1024

# ============================================================
# ヘルパー関数
# ============================================================

function Convert-ColumnNumberToLetter {
    <#
    .DESCRIPTION
        列番号（1始まり）をExcel列名（A, B, ..., Z, AA, AB, ...）に変換
    #>
    param([int]$ColumnNumber)

    $result = ""
    while ($ColumnNumber -gt 0) {
        $ColumnNumber--
        $result = [char](65 + ($ColumnNumber % 26)) + $result
        $ColumnNumber = [Math]::Floor($ColumnNumber / 26)
    }
    return $result
}

function Get-MeaningfulRange {
    <#
    .DESCRIPTION
        UsedRange内で「値または数式を持つセル」のみの範囲を特定する。
        書式（背景色等）のみのセルを除外し、実データ範囲を返す。

        戻り値: @{ FirstRow, LastRow, FirstCol, LastCol } （1始まり）
        データなしの場合は $null を返す
    #>
    param(
        $Sheet
    )

    $usedRange = $Sheet.UsedRange
    if ($null -eq $usedRange) { return $null }

    # UsedRangeの境界を取得
    $startRow = $usedRange.Row
    $startCol = $usedRange.Column
    $endRow = $startRow + $usedRange.Rows.Count - 1
    $endCol = $startCol + $usedRange.Columns.Count - 1

    Write-Host "  UsedRange: R${startRow}C${startCol}:R${endRow}C${endCol} ($($usedRange.Rows.Count) rows x $($usedRange.Columns.Count) cols)"

    # 値と数式を一括取得（パフォーマンス最適化）
    # Value2: 書式なし生値（日付もシリアル値）、Formulaとの比較で数式有無を判定
    $values = $usedRange.Value2
    $formulas = $usedRange.Formula

    # 1セルのみの場合は配列ではなくスカラーが返る
    $isSingle = ($usedRange.Rows.Count -eq 1 -and $usedRange.Columns.Count -eq 1)

    # 意味のあるセルの境界を探索
    $minRow = [int]::MaxValue
    $maxRow = [int]::MinValue
    $minCol = [int]::MaxValue
    $maxCol = [int]::MinValue
    $hasMeaningfulCell = $false

    $totalRows = $usedRange.Rows.Count
    $totalCols = $usedRange.Columns.Count

    if ($isSingle) {
        # 1セルの場合
        $val = $values
        $formula = $formulas
        $hasValue = ($null -ne $val -and "$val" -ne "")
        # 数式判定: Formula文字列が"="で始まるか
        $hasFormula = ($null -ne $formula -and "$formula".StartsWith("="))

        if ($hasValue -or $hasFormula) {
            $minRow = 1; $maxRow = 1; $minCol = 1; $maxCol = 1
            $hasMeaningfulCell = $true
        }
    }
    else {
        # 複数セルの場合: 行・列をスキャンして意味のある境界を検出
        for ($r = 1; $r -le $totalRows; $r++) {
            for ($c = 1; $c -le $totalCols; $c++) {
                $val = $values[$r, $c]
                $formula = $formulas[$r, $c]

                $hasValue = ($null -ne $val -and "$val" -ne "")
                $hasFormula = ($null -ne $formula -and "$formula".StartsWith("="))

                if ($hasValue -or $hasFormula) {
                    $hasMeaningfulCell = $true
                    if ($r -lt $minRow) { $minRow = $r }
                    if ($r -gt $maxRow) { $maxRow = $r }
                    if ($c -lt $minCol) { $minCol = $c }
                    if ($c -gt $maxCol) { $maxCol = $c }
                }
            }
        }
    }

    if (-not $hasMeaningfulCell) {
        Write-Host "  -> 意味のあるセルなし（書式のみ）" -ForegroundColor Yellow
        return $null
    }

    # UsedRange内のオフセットから実際のシート座標に変換
    $result = @{
        FirstRow = $startRow + $minRow - 1
        LastRow  = $startRow + $maxRow - 1
        FirstCol = $startCol + $minCol - 1
        LastCol  = $startCol + $maxCol - 1
    }

    $firstColLetter = Convert-ColumnNumberToLetter $result.FirstCol
    $lastColLetter = Convert-ColumnNumberToLetter $result.LastCol
    Write-Host "  -> 実データ範囲: ${firstColLetter}$($result.FirstRow):${lastColLetter}$($result.LastRow)"

    return $result
}

function Get-FrozenPaneInfo {
    <#
    .DESCRIPTION
        固定行/列（Freeze Panes）の情報を取得する。
        戻り値: @{ FrozenRows, FrozenCols, HasFreeze }
    #>
    param(
        $ExcelApp
    )

    $window = $ExcelApp.ActiveWindow
    $frozenRows = $window.SplitRow
    $frozenCols = $window.SplitColumn
    $hasFreeze = ($frozenRows -gt 0 -or $frozenCols -gt 0)

    if ($hasFreeze) {
        Write-Host "  固定行/列検出: 行=$frozenRows, 列=$frozenCols" -ForegroundColor Cyan
    }

    return @{
        FrozenRows = [int]$frozenRows
        FrozenCols = [int]$frozenCols
        HasFreeze  = $hasFreeze
    }
}

function Get-CellDimensions {
    <#
    .DESCRIPTION
        指定範囲の各列幅・行高さをピクセル単位で取得する。
        ExcelのWidthプロパティはポイント単位（1pt = 96/72 px @ 96DPI）。

        戻り値: @{ ColumnWidths = @(px...); RowHeights = @(px...); TotalWidth; TotalHeight }
    #>
    param(
        $Sheet,
        [int]$FirstRow,
        [int]$LastRow,
        [int]$FirstCol,
        [int]$LastCol,
        [double]$ScaleFactor = 1.0
    )

    # ポイント→ピクセル変換係数（96 DPI基準）
    $ptToPx = 96.0 / 72.0 * $ScaleFactor

    $colWidths = @()
    $totalWidth = 0
    for ($c = $FirstCol; $c -le $LastCol; $c++) {
        # .Width はポイント単位（読み取り専用）
        $widthPt = $Sheet.Columns($c).Width
        $widthPx = [Math]::Round($widthPt * $ptToPx)
        $colWidths += $widthPx
        $totalWidth += $widthPx
    }

    $rowHeights = @()
    $totalHeight = 0
    for ($r = $FirstRow; $r -le $LastRow; $r++) {
        $heightPt = $Sheet.Rows($r).Height
        $heightPx = [Math]::Round($heightPt * $ptToPx)
        $rowHeights += $heightPx
        $totalHeight += $heightPx
    }

    return @{
        ColumnWidths = $colWidths
        RowHeights   = $rowHeights
        TotalWidth   = $totalWidth
        TotalHeight  = $totalHeight
    }
}

function Capture-RangeToImage {
    <#
    .DESCRIPTION
        Excel範囲をクリップボード経由でBitmapとしてキャプチャする。
        CopyPicture メソッドでセル内容を画像化し、クリップボードから取得。

        戻り値: [System.Drawing.Bitmap] or $null
    #>
    param(
        $Range
    )

    try {
        # クリップボードをクリア（戻り値をOut-Nullで抑制）
        [System.Windows.Forms.Clipboard]::Clear() | Out-Null
        Start-Sleep -Milliseconds 100

        # CopyPicture: xlScreen(1) + xlBitmap(2) でビットマップとしてコピー
        # COM戻り値を抑制しないとPowerShellの関数戻り値に混入する
        $Range.CopyPicture($xlScreen, $xlBitmap) | Out-Null
        Start-Sleep -Milliseconds 300

        # クリップボードから画像を取得
        $img = [System.Windows.Forms.Clipboard]::GetImage()

        if ($null -eq $img) {
            Write-Host "    [WARN] クリップボードから画像取得失敗。リトライ..." -ForegroundColor Yellow
            Start-Sleep -Milliseconds 500
            $img = [System.Windows.Forms.Clipboard]::GetImage()
        }

        if ($null -eq $img) {
            Write-Host "    [ERROR] 画像キャプチャ失敗" -ForegroundColor Red
            return $null
        }

        return $img
    }
    catch {
        Write-Host "    [ERROR] CopyPicture例外: $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }
}

function Add-CoordinateOverlay {
    <#
    .DESCRIPTION
        キャプチャした画像に行番号・列名の座標オーバーレイを描画する。
        左マージンに行番号、上マージンに列名を配置。

        戻り値: [System.Drawing.Bitmap] オーバーレイ付き新画像
    #>
    param(
        [System.Drawing.Image]$SourceImage,
        [int]$FirstRow,
        [int]$FirstCol,
        [int]$LastRow,
        [int]$LastCol,
        [double[]]$ColumnWidths,
        [double[]]$RowHeights
    )

    $srcWidth = $SourceImage.Width
    $srcHeight = $SourceImage.Height

    # オーバーレイ付き画像のサイズ
    $newWidth = $srcWidth + $OVERLAY_MARGIN_LEFT
    $newHeight = $srcHeight + $OVERLAY_MARGIN_TOP

    $bitmap = New-Object System.Drawing.Bitmap($newWidth, $newHeight)
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)

    # 背景を白で塗りつぶし
    $graphics.Clear([System.Drawing.Color]::White)

    # フォント・ブラシ・ペン
    $font = New-Object System.Drawing.Font($OVERLAY_FONT_NAME, $OVERLAY_FONT_SIZE)
    $textBrush = New-Object System.Drawing.SolidBrush($OVERLAY_TEXT_COLOR)
    $bgBrush = New-Object System.Drawing.SolidBrush($OVERLAY_BG_COLOR)
    $borderPen = New-Object System.Drawing.Pen($OVERLAY_BORDER_COLOR, 1)
    $format = New-Object System.Drawing.StringFormat
    $format.Alignment = [System.Drawing.StringAlignment]::Center
    $format.LineAlignment = [System.Drawing.StringAlignment]::Center

    # ヘッダー背景を描画（左マージン + 上マージン）
    $graphics.FillRectangle($bgBrush, 0, 0, $OVERLAY_MARGIN_LEFT, $newHeight)
    $graphics.FillRectangle($bgBrush, 0, 0, $newWidth, $OVERLAY_MARGIN_TOP)

    # 元画像を配置
    $graphics.DrawImage($SourceImage, $OVERLAY_MARGIN_LEFT, $OVERLAY_MARGIN_TOP, $srcWidth, $srcHeight)

    # 列名（A, B, C, ...）を描画
    # CopyPicture の実際のピクセル幅とExcelの列幅計算値にズレが生じるため、
    # 実画像幅に対して列幅の比率でスケーリングする
    $totalCalcWidth = ($ColumnWidths | Measure-Object -Sum).Sum
    $scaleX = if ($totalCalcWidth -gt 0) { $srcWidth / $totalCalcWidth } else { 1.0 }

    $xOffset = $OVERLAY_MARGIN_LEFT
    for ($i = 0; $i -lt $ColumnWidths.Count; $i++) {
        $colWidth = $ColumnWidths[$i] * $scaleX
        $colLetter = Convert-ColumnNumberToLetter ($FirstCol + $i)

        $rect = New-Object System.Drawing.RectangleF($xOffset, 0, $colWidth, $OVERLAY_MARGIN_TOP)
        $graphics.DrawString($colLetter, $font, $textBrush, $rect, $format)

        # 列区切り線
        $graphics.DrawLine($borderPen, [float]$xOffset, 0, [float]$xOffset, [float]$newHeight)

        $xOffset += $colWidth
    }

    # 行番号（1, 2, 3, ...）を描画
    $totalCalcHeight = ($RowHeights | Measure-Object -Sum).Sum
    $scaleY = if ($totalCalcHeight -gt 0) { $srcHeight / $totalCalcHeight } else { 1.0 }

    $yOffset = $OVERLAY_MARGIN_TOP
    for ($i = 0; $i -lt $RowHeights.Count; $i++) {
        $rowHeight = $RowHeights[$i] * $scaleY
        $rowNum = $FirstRow + $i

        $rect = New-Object System.Drawing.RectangleF(0, $yOffset, $OVERLAY_MARGIN_LEFT, $rowHeight)
        $graphics.DrawString("$rowNum", $font, $textBrush, $rect, $format)

        # 行区切り線
        $graphics.DrawLine($borderPen, 0, [float]$yOffset, [float]$newWidth, [float]$yOffset)

        $yOffset += $rowHeight
    }

    # 外枠
    $graphics.DrawRectangle($borderPen, 0, 0, $newWidth - 1, $newHeight - 1)
    # ヘッダーとコンテンツの境界線
    $graphics.DrawLine($borderPen, [float]$OVERLAY_MARGIN_LEFT, 0, [float]$OVERLAY_MARGIN_LEFT, [float]$newHeight)
    $graphics.DrawLine($borderPen, 0, [float]$OVERLAY_MARGIN_TOP, [float]$newWidth, [float]$OVERLAY_MARGIN_TOP)

    # リソース解放
    $font.Dispose()
    $textBrush.Dispose()
    $bgBrush.Dispose()
    $borderPen.Dispose()
    $format.Dispose()
    $graphics.Dispose()

    return $bitmap
}

function Save-ImageWithSizeCheck {
    <#
    .DESCRIPTION
        画像をPNG保存し、10MB超の場合はJPEGで再保存（品質調整）。

        戻り値: 保存パス
    #>
    param(
        [System.Drawing.Bitmap]$Image,
        [string]$OutputPath
    )

    # まずPNGで保存
    $Image.Save($OutputPath, [System.Drawing.Imaging.ImageFormat]::Png)

    $fileSize = (Get-Item $OutputPath).Length

    if ($fileSize -gt $MAX_FILE_SIZE_BYTES) {
        Write-Host "    PNG=${fileSize}bytes (>10MB) -> JPEG変換" -ForegroundColor Yellow

        # JPEGパスに変更
        $jpegPath = [System.IO.Path]::ChangeExtension($OutputPath, ".jpg")

        # JPEG品質パラメータ（85から開始、必要に応じて下げる）
        $quality = 85
        while ($quality -ge 50) {
            $encoder = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() |
                Where-Object { $_.MimeType -eq "image/jpeg" }
            $encoderParams = New-Object System.Drawing.Imaging.EncoderParameters(1)
            $encoderParams.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter(
                [System.Drawing.Imaging.Encoder]::Quality, [long]$quality
            )

            $Image.Save($jpegPath, $encoder, $encoderParams)
            $encoderParams.Dispose()

            $jpegSize = (Get-Item $jpegPath).Length
            if ($jpegSize -le $MAX_FILE_SIZE_BYTES) {
                # PNGを削除してJPEGを返す
                Remove-Item $OutputPath -ErrorAction SilentlyContinue
                Write-Host "    JPEG(Q=$quality)=${jpegSize}bytes -> OK" -ForegroundColor Green
                return $jpegPath
            }

            $quality -= 10
        }

        Write-Host "    [WARN] JPEG Q=50でも10MB超。そのまま保存。" -ForegroundColor Yellow
        Remove-Item $OutputPath -ErrorAction SilentlyContinue
        return $jpegPath
    }

    $sizeMB = [Math]::Round($fileSize / 1MB, 2)
    Write-Host "    PNG=${sizeMB}MB -> OK" -ForegroundColor Green
    return $OutputPath
}

# ============================================================
# メイン処理
# ============================================================

# 入力検証
$resolvedPath = Resolve-Path $FilePath -ErrorAction SilentlyContinue
if ($null -eq $resolvedPath) {
    Write-Host "[ERROR] ファイルが見つかりません: $FilePath" -ForegroundColor Red
    exit 1
}
$FilePath = $resolvedPath.Path

# 出力ディレクトリ
if (-not $OutputDir) {
    $parentDir = Split-Path $FilePath -Parent
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
    $OutputDir = Join-Path $parentDir "${baseName}_screenshots"
}

if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
}

Write-Host "================================================================" -ForegroundColor Cyan
Write-Host " Excel Screenshot Capture Tool (with Coordinate Overlay)" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "Input : $FilePath"
Write-Host "Output: $OutputDir"
Write-Host ""

# Excel COM起動
$excel = $null
$wb = $null

try {
    Write-Host "[1/4] Excel起動中..." -ForegroundColor Yellow
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $true   # CopyPictureにはScreenUpdating=trueが必要

    Write-Host "[2/4] ファイルを開いています: $(Split-Path $FilePath -Leaf)" -ForegroundColor Yellow
    $wb = $excel.Workbooks.Open($FilePath, $false, $true)  # UpdateLinks=false, ReadOnly=true

    $sheetCount = $wb.Worksheets.Count
    Write-Host "[3/4] シート数: $sheetCount" -ForegroundColor Yellow
    Write-Host ""

    # メタ情報を記録（構造情報JSON向け）
    $metaInfo = @{
        FileName   = Split-Path $FilePath -Leaf
        SheetCount = $sheetCount
        Sheets     = @()
    }

    # 各シートを処理
    for ($sheetIdx = 1; $sheetIdx -le $sheetCount; $sheetIdx++) {
        $sheet = $wb.Worksheets.Item($sheetIdx)
        $sheetName = $sheet.Name

        Write-Host "--- Sheet $sheetIdx/$sheetCount : [$sheetName] ---" -ForegroundColor Cyan

        # シートをアクティブ化（CopyPicture / SplitRow 等に必要）
        $sheet.Activate()
        Start-Sleep -Milliseconds 200

        # 固定行/列の検出
        $freezeInfo = Get-FrozenPaneInfo -ExcelApp $excel

        # 意味のある範囲の検出
        $dataRange = Get-MeaningfulRange -Sheet $sheet

        if ($null -eq $dataRange) {
            Write-Host "  スキップ（データなし）" -ForegroundColor DarkGray
            $metaInfo.Sheets += @{
                Name = $sheetName
                Index = $sheetIdx
                Status = "skipped_no_data"
            }
            continue
        }

        # セル寸法の取得
        $dims = Get-CellDimensions -Sheet $sheet `
            -FirstRow $dataRange.FirstRow -LastRow $dataRange.LastRow `
            -FirstCol $dataRange.FirstCol -LastCol $dataRange.LastCol `
            -ScaleFactor $DPI

        Write-Host "  画像推定サイズ: $($dims.TotalWidth) x $($dims.TotalHeight) px"

        # キャプチャ対象範囲を取得
        $firstColLetter = Convert-ColumnNumberToLetter $dataRange.FirstCol
        $lastColLetter = Convert-ColumnNumberToLetter $dataRange.LastCol
        $rangeAddress = "${firstColLetter}$($dataRange.FirstRow):${lastColLetter}$($dataRange.LastRow)"

        Write-Host "  キャプチャ範囲: $rangeAddress"

        $captureRange = $sheet.Range($rangeAddress)

        # 画像キャプチャ
        Write-Host "  キャプチャ中..."
        $img = Capture-RangeToImage -Range $captureRange

        if ($null -eq $img) {
            Write-Host "  [ERROR] キャプチャ失敗 -> スキップ" -ForegroundColor Red
            # 検出情報はキャプチャ失敗時でもAI分析に有用なので記録する
            $metaInfo.Sheets += @{
                Name = $sheetName
                Index = $sheetIdx
                Status = "capture_failed"
                Range = $rangeAddress
                FrozenRows = $freezeInfo.FrozenRows
                FrozenCols = $freezeInfo.FrozenCols
                EstimatedWidth = $dims.TotalWidth
                EstimatedHeight = $dims.TotalHeight
            }
            continue
        }

        Write-Host "  キャプチャ成功: $($img.Width) x $($img.Height) px"

        # 座標オーバーレイの描画
        Write-Host "  座標オーバーレイ描画中..."
        $overlayImg = Add-CoordinateOverlay `
            -SourceImage $img `
            -FirstRow $dataRange.FirstRow `
            -FirstCol $dataRange.FirstCol `
            -LastRow $dataRange.LastRow `
            -LastCol $dataRange.LastCol `
            -ColumnWidths $dims.ColumnWidths `
            -RowHeights $dims.RowHeights

        # 元画像を解放
        $img.Dispose()

        # 安全なファイル名生成（シート名の特殊文字を除去）
        $safeSheetName = $sheetName -replace '[\\/:*?"<>|]', '_'
        $fileName = "Sheet${sheetIdx}_${safeSheetName}_${rangeAddress -replace ':', '-'}.png"
        $outputPath = Join-Path $OutputDir $fileName

        # 保存（サイズチェック付き）
        Write-Host "  保存中..."
        $savedPath = Save-ImageWithSizeCheck -Image $overlayImg -OutputPath $outputPath

        # オーバーレイ画像を解放
        $overlayImg.Dispose()

        # 固定行ヘッダーの別途キャプチャ（固定行がデータ範囲と重ならない場合）
        if ($freezeInfo.HasFreeze -and $freezeInfo.FrozenRows -gt 0) {
            $frozenLastRow = $freezeInfo.FrozenRows
            # 固定行がデータ範囲に含まれていない場合のみ別途キャプチャ
            if ($frozenLastRow -lt $dataRange.FirstRow) {
                Write-Host "  固定ヘッダー行を別途キャプチャ中 (行 1-$frozenLastRow)..."

                $frozenColLetter = Convert-ColumnNumberToLetter $dataRange.LastCol
                $frozenAddress = "A1:${frozenColLetter}${frozenLastRow}"
                $frozenRange = $sheet.Range($frozenAddress)

                $frozenImg = Capture-RangeToImage -Range $frozenRange
                if ($null -ne $frozenImg) {
                    $frozenDims = Get-CellDimensions -Sheet $sheet `
                        -FirstRow 1 -LastRow $frozenLastRow `
                        -FirstCol 1 -LastCol $dataRange.LastCol `
                        -ScaleFactor $DPI

                    $frozenOverlay = Add-CoordinateOverlay `
                        -SourceImage $frozenImg `
                        -FirstRow 1 -FirstCol 1 `
                        -LastRow $frozenLastRow -LastCol $dataRange.LastCol `
                        -ColumnWidths $frozenDims.ColumnWidths `
                        -RowHeights $frozenDims.RowHeights

                    $frozenImg.Dispose()

                    $frozenFileName = "Sheet${sheetIdx}_${safeSheetName}_frozen_header.png"
                    $frozenOutputPath = Join-Path $OutputDir $frozenFileName
                    $frozenSavedPath = Save-ImageWithSizeCheck -Image $frozenOverlay -OutputPath $frozenOutputPath
                    $frozenOverlay.Dispose()

                    Write-Host "  固定ヘッダー保存: $(Split-Path $frozenSavedPath -Leaf)" -ForegroundColor Green
                }
            }
        }

        # メタ情報に記録
        $metaInfo.Sheets += @{
            Name = $sheetName
            Index = $sheetIdx
            Status = "captured"
            Range = $rangeAddress
            ImageFile = Split-Path $savedPath -Leaf
            FrozenRows = $freezeInfo.FrozenRows
            FrozenCols = $freezeInfo.FrozenCols
            ImageWidth = $dims.TotalWidth + $OVERLAY_MARGIN_LEFT
            ImageHeight = $dims.TotalHeight + $OVERLAY_MARGIN_TOP
        }

        Write-Host "  完了: $(Split-Path $savedPath -Leaf)" -ForegroundColor Green
        Write-Host ""
    }

    # メタ情報JSONを保存
    $metaJsonPath = Join-Path $OutputDir "_capture_meta.json"
    $metaInfo | ConvertTo-Json -Depth 5 | Set-Content -Path $metaJsonPath -Encoding UTF8
    Write-Host "[4/4] メタ情報保存: _capture_meta.json" -ForegroundColor Yellow

}
catch {
    Write-Host "[FATAL] $($_.Exception.Message)" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor DarkRed
}
finally {
    # クリーンアップ
    Write-Host ""
    Write-Host "クリーンアップ中..."

    if ($null -ne $wb) {
        try { $wb.Close($false) } catch {}
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
    }

    # COMオブジェクトの解放
    if ($null -ne $wb) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null }
    if ($null -ne $excel) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    Write-Host ""
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host " 完了! 出力先: $OutputDir" -ForegroundColor Cyan
    Write-Host "================================================================" -ForegroundColor Cyan
}
