<#
.SYNOPSIS
    スクリーンショットツールのテスト用Excelファイルを生成する
.DESCRIPTION
    テスト用.xlsmファイルを作成:
    - 複数シート / VBAマクロ / 固定行 / 数式 / 書式のみ空セル / 条件付き書式 / 名前付き範囲
#>
param(
    [string]$OutputPath
)

if (-not $OutputPath) {
    $OutputPath = Join-Path $PSScriptRoot "TestWorkbook_HR_Estimate.xlsm"
}

Write-Host "Test workbook generation starting..."

$excel = $null
$wb = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $wb = $excel.Workbooks.Add()

    # === Sheet1: Estimate Data ===
    $s1 = $wb.Worksheets.Item(1)
    $s1.Name = [string]"EstimateData"

    # Headers (row 1)
    $hdr = @("No", "Name", "Dept", "Grade", "BaseSalary", "RoleAllow", "Commute", "Total", "MoM", "Note")
    for ($c = 1; $c -le $hdr.Count; $c++) {
        $cell = $s1.Cells.Item(1, $c)
        $cell.Value = [string]$hdr[$c - 1]
        $cell.Font.Bold = $true
        $cell.Interior.Color = [int]14277081  # 0xD9E1F2 light blue
    }

    # Sub-header (row 2)
    $sub = @("", "", "", "", "(JPY)", "(JPY)", "(JPY)", "(JPY)", "(%)", "")
    for ($c = 1; $c -le $sub.Count; $c++) {
        $cell = $s1.Cells.Item(2, $c)
        $cell.Value = [string]$sub[$c - 1]
        $cell.Interior.Color = [int]14868186  # 0xE2EFDA light green
    }

    # Sample data (10 rows, starting row 3)
    $nameList = @("Tanaka", "Yamada", "Suzuki", "Sato", "Takahashi",
                  "Watanabe", "Ito", "Nakamura", "Kobayashi", "Kato")
    $deptList = @("Sales", "HR", "Tech", "Sales", "Admin",
                  "Tech", "HR", "Sales", "Tech", "Admin")
    $gradeList = @("M1", "S3", "E2", "M2", "S2", "E3", "S1", "M1", "E1", "S3")
    $baseList = @(350000, 280000, 320000, 380000, 260000, 340000, 250000, 350000, 300000, 280000)
    $roleList = @(50000, 0, 30000, 60000, 0, 35000, 0, 50000, 20000, 0)
    $commuteList = @(15000, 20000, 12000, 18000, 10000, 25000, 8000, 15000, 22000, 10000)

    for ($r = 0; $r -lt 10; $r++) {
        $row = $r + 3
        $s1.Cells.Item($row, 1).Value = [double]($r + 1)
        $s1.Cells.Item($row, 2).Value = [string]$nameList[$r]
        $s1.Cells.Item($row, 3).Value = [string]$deptList[$r]
        $s1.Cells.Item($row, 4).Value = [string]$gradeList[$r]
        $s1.Cells.Item($row, 5).Value = [double]$baseList[$r]
        $s1.Cells.Item($row, 6).Value = [double]$roleList[$r]
        $s1.Cells.Item($row, 7).Value = [double]$commuteList[$r]

        # Total formula: =E{row}+F{row}+G{row}
        $s1.Cells.Item($row, 8).Formula = "=E$row+F$row+G$row"

        # MoM formula
        $s1.Cells.Item($row, 9).Formula = "=ROUND((H$row-H${row}*0.98)/H${row}*100,1)"
    }

    # Number format for salary columns
    for ($r = 3; $r -le 12; $r++) {
        for ($c = 5; $c -le 8; $c++) {
            $s1.Cells.Item($r, $c).NumberFormat = "#,##0"
        }
        $s1.Cells.Item($r, 9).NumberFormat = "0.0"
    }

    # Summary row (row 13)
    $s1.Cells.Item(13, 4).Value = [string]"Total"
    $s1.Cells.Item(13, 4).Font.Bold = $true
    $s1.Cells.Item(13, 5).Formula = "=SUM(E3:E12)"
    $s1.Cells.Item(13, 6).Formula = "=SUM(F3:F12)"
    $s1.Cells.Item(13, 7).Formula = "=SUM(G3:G12)"
    $s1.Cells.Item(13, 8).Formula = "=SUM(H3:H12)"
    for ($c = 5; $c -le 8; $c++) {
        $s1.Cells.Item(13, $c).NumberFormat = "#,##0"
        $s1.Cells.Item(13, $c).Font.Bold = $true
        $s1.Cells.Item(13, $c).Interior.Color = [int]16573654  # 0xFCE4D6 light orange
    }

    # Format-only empty cells (smart area detection test)
    # Rows 15-20, columns A-J: background color but NO value
    for ($r = 15; $r -le 20; $r++) {
        for ($c = 1; $c -le 10; $c++) {
            $s1.Cells.Item($r, $c).Interior.Color = [int]15921906  # 0xF2F2F2 light gray
        }
    }

    # Freeze panes (rows 1-2)
    # Select()はVisible=falseで問題が起きるため、SplitRowで直接設定
    $s1.Activate()
    Start-Sleep -Milliseconds 300
    try {
        $excel.ActiveWindow.SplitRow = 2
        $excel.ActiveWindow.SplitColumn = 0
        $excel.ActiveWindow.FreezePanes = $true
        Write-Host "  Freeze panes set (rows 1-2)"
    }
    catch {
        Write-Host "  [WARN] Freeze panes failed: $($_.Exception.Message)" -ForegroundColor Yellow
    }

    # Auto-fit columns
    $s1.Columns.Item("A:J").AutoFit() | Out-Null

    # Conditional formatting: MoM > 2 -> red
    $condRange = $s1.Range("I3:I12")
    $cond = $condRange.FormatConditions.Add(1, 5, "2")  # xlCellValue=1, xlGreater=5
    $cond.Interior.Color = [int]16552080  # 0xFC9090 red
    $cond.Font.Color = [int]10027008     # 0x990000 dark red

    Write-Host "  Sheet1 (EstimateData) created"

    # === Sheet2: Dept Summary ===
    $s2 = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count))
    $s2.Name = [string]"DeptSummary"

    $deptHeaders = @("Dept", "Count", "TotalCost", "AvgSalary")
    for ($c = 1; $c -le $deptHeaders.Count; $c++) {
        $cell = $s2.Cells.Item(1, $c)
        $cell.Value = [string]$deptHeaders[$c - 1]
        $cell.Font.Bold = $true
        $cell.Interior.Color = [int]14277081
    }

    $depts = @("Sales", "HR", "Tech", "Admin")
    for ($d = 0; $d -lt $depts.Count; $d++) {
        $row = $d + 2
        $s2.Cells.Item($row, 1).Value = [string]$depts[$d]
        # Cross-sheet formulas referencing EstimateData
        $s2.Cells.Item($row, 2).Formula = "=COUNTIF(EstimateData!C:C,A$row)"
        $s2.Cells.Item($row, 3).Formula = "=SUMIF(EstimateData!C:C,A$row,EstimateData!H:H)"
        $s2.Cells.Item($row, 4).Formula = "=IF(B$row>0,C$row/B$row,0)"
        $s2.Cells.Item($row, 3).NumberFormat = "#,##0"
        $s2.Cells.Item($row, 4).NumberFormat = "#,##0"
    }

    # Named ranges
    $wb.Names.Add("DeptMaster", $s2.Range("A2:A5"))
    $wb.Names.Add("TotalCost", $s2.Range("C2:C5"))

    $s2.Columns.Item("A:D").AutoFit() | Out-Null
    Write-Host "  Sheet2 (DeptSummary) created"

    # === Sheet3: Settings ===
    $s3 = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count))
    $s3.Name = [string]"Settings"

    $s3.Cells.Item(1, 1).Value = [string]"SettingName"
    $s3.Cells.Item(1, 2).Value = [string]"Value"
    $s3.Cells.Item(1, 1).Font.Bold = $true
    $s3.Cells.Item(1, 2).Font.Bold = $true

    $settings = @(
        @("CalcMonth", "2026-02"),
        @("TaxRate", "0.1"),
        @("InsuranceRate", "0.15"),
        @("CommuteLimit", "30000"),
        @("Author", "BP_User")
    )
    for ($i = 0; $i -lt $settings.Count; $i++) {
        $s3.Cells.Item($i + 2, 1).Value = [string]$settings[$i][0]
        $s3.Cells.Item($i + 2, 2).Value = [string]$settings[$i][1]
    }

    $s3.Columns.Item("A:B").AutoFit() | Out-Null
    Write-Host "  Sheet3 (Settings) created"

    # === VBA Module ===
    Write-Host "  Adding VBA module..."
    try {
        $vbProject = $wb.VBProject
        $vbModule = $vbProject.VBComponents.Add(1)  # vbext_ct_StdModule
        $vbModule.Name = "Module1"

        # VBA code as a single-line string to avoid encoding issues
        $vbaLines = @(
            "' Estimate Tool Main Macros",
            "' Talent Palette Integration",
            "",
            "Sub UpdateEstimate()",
            "    Dim ws As Worksheet",
            "    Set ws = ThisWorkbook.Worksheets(""EstimateData"")",
            "    Dim taxRate As Double",
            "    taxRate = ThisWorkbook.Worksheets(""Settings"").Range(""B3"").Value",
            "    Dim lastRow As Long",
            "    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row",
            "    Dim i As Long",
            "    For i = 3 To lastRow",
            "        ws.Cells(i, 8).Value = ws.Cells(i, 5).Value + ws.Cells(i, 6).Value + ws.Cells(i, 7).Value",
            "    Next i",
            "    MsgBox ""Estimate updated"", vbInformation",
            "End Sub",
            "",
            "Sub ExportToPalette()",
            "    Dim ws As Worksheet",
            "    Set ws = ThisWorkbook.Worksheets(""EstimateData"")",
            "    Dim exportRange As Range",
            "    Set exportRange = ws.Range(""A1:H"" & ws.Cells(ws.Rows.Count, 2).End(xlUp).Row)",
            "    Dim filePath As String",
            "    filePath = ThisWorkbook.Path & ""\export_"" & Format(Now, ""yyyymmdd"") & "".csv""",
            "    Dim fNum As Integer",
            "    fNum = FreeFile",
            "    Open filePath For Output As #fNum",
            "    Dim r As Long, c As Long",
            "    For r = 1 To exportRange.Rows.Count",
            "        Dim line As String",
            "        line = """"",
            "        For c = 1 To exportRange.Columns.Count",
            "            If c > 1 Then line = line & "",""",
            "            line = line & exportRange.Cells(r, c).Text",
            "        Next c",
            "        Print #fNum, line",
            "    Next r",
            "    Close #fNum",
            "    MsgBox ""Export complete: "" & filePath, vbInformation",
            "End Sub",
            "",
            "Sub FormatSheet()",
            "    Dim ws As Worksheet",
            "    Set ws = ActiveSheet",
            "    With ws.Range(""A1:J1"")",
            "        .Font.Bold = True",
            "        .Interior.Color = RGB(217, 225, 242)",
            "        .Borders(xlEdgeBottom).LineStyle = xlContinuous",
            "    End With",
            "    ws.Columns(""A:J"").AutoFit",
            "End Sub"
        )
        $vbaCode = $vbaLines -join "`r`n"
        $vbModule.CodeModule.AddFromString($vbaCode)
        Write-Host "  VBA Module1 added successfully" -ForegroundColor Green
    }
    catch {
        Write-Host "  [WARN] VBA module add failed (VBProject access may be disabled): $($_.Exception.Message)" -ForegroundColor Yellow
        Write-Host "  -> Enable: Trust Center > Macro Settings > Trust access to VBA project object model" -ForegroundColor Yellow
    }

    # === Save as .xlsm ===
    # xlOpenXMLWorkbookMacroEnabled = 52
    $wb.SaveAs($OutputPath, 52)
    Write-Host ""
    Write-Host "Test workbook created: $OutputPath" -ForegroundColor Green
    Write-Host "  Sheet1: EstimateData (frozen rows, formulas, conditional format, empty formatted cells)"
    Write-Host "  Sheet2: DeptSummary (cross-sheet formulas, named ranges)"
    Write-Host "  Sheet3: Settings (master data)"
    Write-Host "  VBA: Module1 (UpdateEstimate, ExportToPalette, FormatSheet)"
}
catch {
    Write-Host "[ERROR] $($_.Exception.Message)" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor DarkRed
}
finally {
    if ($null -ne $wb) { try { $wb.Close($false) } catch {} }
    if ($null -ne $excel) { try { $excel.Quit() } catch {} }
    if ($null -ne $wb) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null }
    if ($null -ne $excel) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
