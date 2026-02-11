<#
.SYNOPSIS
    座標オーバーレイ描画機能の単体テスト。
    ダミー画像を使ってOverlay描画ロジックを検証する。
#>

Add-Type -AssemblyName System.Drawing

# 定数（メインスクリプトと同じ値）
$OVERLAY_MARGIN_LEFT = 40
$OVERLAY_MARGIN_TOP = 25
$OVERLAY_FONT_NAME = "Consolas"
$OVERLAY_FONT_SIZE = 9
$OVERLAY_BG_COLOR = [System.Drawing.Color]::FromArgb(240, 240, 240)
$OVERLAY_BORDER_COLOR = [System.Drawing.Color]::FromArgb(180, 180, 180)
$OVERLAY_TEXT_COLOR = [System.Drawing.Color]::FromArgb(60, 60, 60)

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

# ダミーExcelスクリーンショットを生成（10列 x 13行のグリッド）
$srcWidth = 800
$srcHeight = 350
$srcBitmap = New-Object System.Drawing.Bitmap($srcWidth, $srcHeight)
$srcGraphics = [System.Drawing.Graphics]::FromImage($srcBitmap)
$srcGraphics.Clear([System.Drawing.Color]::White)

# セルグリッドを描画（ダミー）
$gridPen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(200, 200, 200), 1)
$cellFont = New-Object System.Drawing.Font("Segoe UI", 8)
$cellBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::Black)

# 10列、各80px幅
$colWidths = @(40, 100, 60, 60, 80, 80, 70, 90, 50, 70)
# 13行、各27px高さ（ヘッダー2行 + データ10行 + 合計1行）
$rowHeights = @(27, 22, 22, 22, 22, 22, 22, 22, 22, 22, 22, 22, 27)

# ヘッダー行背景
$headerFill = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(217, 225, 242))
$xPos = 0
foreach ($w in $colWidths) {
    $srcGraphics.FillRectangle($headerFill, $xPos, 0, $w, 27)
    $srcGraphics.FillRectangle($headerFill, $xPos, 27, $w, 22)
    $xPos += $w
}

# グリッド線
$xPos = 0
foreach ($w in $colWidths) {
    $xPos += $w
    $srcGraphics.DrawLine($gridPen, [float]$xPos, 0, [float]$xPos, [float]$srcHeight)
}
$yPos = 0
foreach ($h in $rowHeights) {
    $yPos += $h
    $srcGraphics.DrawLine($gridPen, 0, [float]$yPos, [float]$srcWidth, [float]$yPos)
}

# ダミーテキスト
$headers = @("No", "Name", "Dept", "Grade", "Base", "Role", "Comm", "Total", "MoM", "Note")
$xPos = 2
for ($c = 0; $c -lt $headers.Count; $c++) {
    $srcGraphics.DrawString($headers[$c], $cellFont, $cellBrush, [float]$xPos, 5.0)
    $xPos += $colWidths[$c]
}

$srcGraphics.Dispose()
$headerFill.Dispose()
$gridPen.Dispose()
$cellFont.Dispose()
$cellBrush.Dispose()

Write-Host "Dummy source image created: ${srcWidth} x ${srcHeight} px"

# === 座標オーバーレイ描画 ===
$firstRow = 1
$firstCol = 1
$lastRow = 13
$lastCol = 10

$newWidth = $srcWidth + $OVERLAY_MARGIN_LEFT
$newHeight = $srcHeight + $OVERLAY_MARGIN_TOP

$bitmap = New-Object System.Drawing.Bitmap($newWidth, $newHeight)
$graphics = [System.Drawing.Graphics]::FromImage($bitmap)
$graphics.Clear([System.Drawing.Color]::White)

$font = New-Object System.Drawing.Font($OVERLAY_FONT_NAME, $OVERLAY_FONT_SIZE)
$textBrush = New-Object System.Drawing.SolidBrush($OVERLAY_TEXT_COLOR)
$bgBrush = New-Object System.Drawing.SolidBrush($OVERLAY_BG_COLOR)
$borderPen = New-Object System.Drawing.Pen($OVERLAY_BORDER_COLOR, 1)
$format = New-Object System.Drawing.StringFormat
$format.Alignment = [System.Drawing.StringAlignment]::Center
$format.LineAlignment = [System.Drawing.StringAlignment]::Center

# ヘッダー背景
$graphics.FillRectangle($bgBrush, 0, 0, $OVERLAY_MARGIN_LEFT, $newHeight)
$graphics.FillRectangle($bgBrush, 0, 0, $newWidth, $OVERLAY_MARGIN_TOP)

# 元画像を配置
$graphics.DrawImage($srcBitmap, $OVERLAY_MARGIN_LEFT, $OVERLAY_MARGIN_TOP, $srcWidth, $srcHeight)

# 列名（A, B, C, ...）
$totalCalcWidth = ($colWidths | Measure-Object -Sum).Sum
$scaleX = if ($totalCalcWidth -gt 0) { $srcWidth / $totalCalcWidth } else { 1.0 }

$xOffset = $OVERLAY_MARGIN_LEFT
for ($i = 0; $i -lt $colWidths.Count; $i++) {
    $colWidth = $colWidths[$i] * $scaleX
    $colLetter = Convert-ColumnNumberToLetter ($firstCol + $i)
    $rect = New-Object System.Drawing.RectangleF($xOffset, 0, $colWidth, $OVERLAY_MARGIN_TOP)
    $graphics.DrawString($colLetter, $font, $textBrush, $rect, $format)
    $graphics.DrawLine($borderPen, [float]$xOffset, 0, [float]$xOffset, [float]$newHeight)
    $xOffset += $colWidth
}

# 行番号
$totalCalcHeight = ($rowHeights | Measure-Object -Sum).Sum
$scaleY = if ($totalCalcHeight -gt 0) { $srcHeight / $totalCalcHeight } else { 1.0 }

$yOffset = $OVERLAY_MARGIN_TOP
for ($i = 0; $i -lt $rowHeights.Count; $i++) {
    $rowHeight = $rowHeights[$i] * $scaleY
    $rowNum = $firstRow + $i
    $rect = New-Object System.Drawing.RectangleF(0, $yOffset, $OVERLAY_MARGIN_LEFT, $rowHeight)
    $graphics.DrawString("$rowNum", $font, $textBrush, $rect, $format)
    $graphics.DrawLine($borderPen, 0, [float]$yOffset, [float]$newWidth, [float]$yOffset)
    $yOffset += $rowHeight
}

# 外枠と境界線
$graphics.DrawRectangle($borderPen, 0, 0, $newWidth - 1, $newHeight - 1)
$graphics.DrawLine($borderPen, [float]$OVERLAY_MARGIN_LEFT, 0, [float]$OVERLAY_MARGIN_LEFT, [float]$newHeight)
$graphics.DrawLine($borderPen, 0, [float]$OVERLAY_MARGIN_TOP, [float]$newWidth, [float]$OVERLAY_MARGIN_TOP)

# クリーンアップ
$font.Dispose()
$textBrush.Dispose()
$bgBrush.Dispose()
$borderPen.Dispose()
$format.Dispose()
$graphics.Dispose()
$srcBitmap.Dispose()

# 保存
$outputPath = Join-Path $PSScriptRoot "test_overlay_result.png"
$bitmap.Save($outputPath, [System.Drawing.Imaging.ImageFormat]::Png)
$bitmap.Dispose()

$fileSize = (Get-Item $outputPath).Length
$sizeMB = [Math]::Round($fileSize / 1MB, 3)
Write-Host "Overlay test image saved: $outputPath (${sizeMB} MB)" -ForegroundColor Green
