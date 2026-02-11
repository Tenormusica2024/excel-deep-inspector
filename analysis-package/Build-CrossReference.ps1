<#
.SYNOPSIS
    VBA Cross-Reference Builder - Excel Deep Inspector
.DESCRIPTION
    Parse-VBAModules.ps1の出力 + controls.json を統合して
    ui_to_vba.json（UI操作→VBA→セル影響チェーン）と
    data_flow.json（テーブル列のデータフロー追跡）を生成する。

    入力ファイル（すべてJSON）:
    - cell_references_v2.json  （Parse-VBAModules.ps1出力）
    - table_references.json    （Parse-VBAModules.ps1出力）
    - event_triggers.json      （Parse-VBAModules.ps1出力）
    - form_calls.json          （Parse-VBAModules.ps1出力）
    - controls.json            （Generate-AnalysisPackage.ps1出力、04_structure/）

.PARAMETER AnalysisDir
    分析パッケージのルートディレクトリ（{workbook}_analysis_package/）
.PARAMETER V2Dir
    Parse-VBAModules.ps1のv2出力ディレクトリ（省略時は01_VBA/v2_output/）
.PARAMETER OutputDir
    クロスリファレンスJSON出力先（省略時は05_cross_reference/）
#>
param(
    [Parameter(Mandatory = $true)]
    [string]$AnalysisDir,

    [string]$V2Dir,
    [string]$OutputDir
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if (-not $V2Dir) { $V2Dir = Join-Path $AnalysisDir "01_VBA\v2_output" }
if (-not $OutputDir) { $OutputDir = Join-Path $AnalysisDir "05_cross_reference" }
if (-not (Test-Path $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null }

Write-Host "================================================================" -ForegroundColor Cyan
Write-Host " VBA Cross-Reference Builder" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan

# ============================================================
# 入力ファイル読み込み
# ============================================================

function Read-JsonSafe {
    <# JSONファイルを読み込み。存在しない場合はデフォルト値を返す #>
    param([string]$Path, $Default = $null)
    if (Test-Path $Path) {
        $raw = Get-Content -Path $Path -Encoding UTF8 -Raw
        return ($raw | ConvertFrom-Json)
    }
    Write-Warning "File not found: $Path"
    return $Default
}

$cellRefs      = Read-JsonSafe (Join-Path $V2Dir "cell_references_v2.json")
$tableRefs     = Read-JsonSafe (Join-Path $V2Dir "table_references.json")
$eventTriggers = Read-JsonSafe (Join-Path $V2Dir "event_triggers.json")
$formCalls     = Read-JsonSafe (Join-Path $V2Dir "form_calls.json")

# controls.jsonは04_structure/にある
$controlsPath = Join-Path $AnalysisDir "04_structure\controls.json"
$controls     = Read-JsonSafe $controlsPath

# procedures.jsonは01_VBA/にある（v1出力）
$proceduresPath = Join-Path $AnalysisDir "01_VBA\procedures.json"
$procedures   = Read-JsonSafe $proceduresPath

$inputCellCount  = @($cellRefs.References).Count
$inputTableCount = @($tableRefs.Tables).Count
$inputEventCount = @($eventTriggers.Events).Count
$inputFormCount  = @($formCalls.Calls).Count
$inputCtrlCount  = @($controls.Controls).Count
Write-Host "Input files loaded:" -ForegroundColor Green
Write-Host "  Cell References:  $inputCellCount" -ForegroundColor Gray
Write-Host "  Table References: $inputTableCount tables" -ForegroundColor Gray
Write-Host "  Event Triggers:   $inputEventCount" -ForegroundColor Gray
Write-Host "  Form Calls:       $inputFormCount" -ForegroundColor Gray
Write-Host "  Controls:         $inputCtrlCount" -ForegroundColor Gray
Write-Host ""

# ============================================================
# ヘルパー関数群
# PS5.1制約: [PSCustomObject]@{} ハッシュリテラル内に { } スクリプトブロックを
# 書くとパーサーがカーリーブレースを誤認するため、全て事前計算する
# ============================================================

function Get-ProcedureImpact {
    <# 指定プロシージャが影響するセル参照・テーブル参照・フォーム呼び出しを収集 #>
    param(
        [string]$Module,
        [string]$Procedure
    )

    # このプロシージャ内のセル参照（foreach展開でWhere-Object回避）
    $cells = @()
    foreach ($ref in @($cellRefs.References)) {
        if ($ref.Module -eq $Module -and $ref.Procedure -eq $Procedure) {
            $cells += $ref
        }
    }

    # このプロシージャ内のテーブル列アクセス（ProcedureAccessから個別のAccessType取得）
    $tableCols = @()
    foreach ($table in @($tableRefs.Tables)) {
        foreach ($col in @($table.ColumnsAccessed)) {
            foreach ($pa in @($col.ProcedureAccess)) {
                if ($pa.Procedure -eq $Procedure) {
                    $tableCols += [PSCustomObject]@{
                        TableName  = $table.TableName
                        ColumnName = $col.ColumnName
                        AccessType = $pa.AccessType
                    }
                    break
                }
            }
        }
    }

    # このプロシージャからのフォーム呼び出し
    $forms = @()
    foreach ($fc in @($formCalls.Calls)) {
        if ($fc.Module -eq $Module -and $fc.Procedure -eq $Procedure) {
            $forms += $fc
        }
    }

    return @{
        CellsAffected  = $cells
        TablesAffected  = $tableCols
        FormsTriggered  = $forms
    }
}

function Convert-ImpactToCellsArray {
    <# Impactのセル参照を出力用フォーマットに変換 #>
    param($Impact)
    $result = @()
    if (-not $Impact) { return $result }
    foreach ($cell in @($Impact.CellsAffected)) {
        $result += [PSCustomObject]@{
            Address    = $cell.RawAddress
            Sheet      = $cell.SheetCodeName
            AccessType = $cell.AccessType
        }
    }
    return $result
}

function Convert-ImpactToTablesArray {
    <# Impactのテーブル参照を出力用フォーマットに変換 #>
    param($Impact)
    $result = @()
    if (-not $Impact) { return $result }
    foreach ($tbl in @($Impact.TablesAffected)) {
        $result += [PSCustomObject]@{
            Table      = $tbl.TableName
            Column     = $tbl.ColumnName
            AccessType = $tbl.AccessType
        }
    }
    return $result
}

function Convert-ImpactToFormsArray {
    <# Impactのフォーム呼び出しからFormName配列を取得 #>
    param($Impact)
    $result = @()
    if (-not $Impact) { return $result }
    foreach ($fc in @($Impact.FormsTriggered)) {
        if ($fc.FormName -and $fc.FormName -notin $result) { $result += $fc.FormName }
    }
    return $result
}

function Build-FormSubChains {
    <# フォーム呼び出しチェーンを展開（1階層） #>
    param($Impact)
    $subChains = @()
    if (-not $Impact) { return $subChains }
    if (@($Impact.FormsTriggered).Count -eq 0) { return $subChains }

    foreach ($fc in @($Impact.FormsTriggered)) {
        $formName = $fc.FormName
        # フォーム内のUserFormイベントをforeach展開で収集
        $formEvents = @()
        foreach ($evt in @($eventTriggers.Events)) {
            if ($evt.Module -eq $formName -and $evt.Type -eq "USERFORM_EVENT") {
                $formEvents += $evt
            }
        }

        $formEventDetails = @()
        foreach ($fe in $formEvents) {
            $feImpact = Get-ProcedureImpact -Module $fe.Module -Procedure $fe.ProcedureName

            # セル参照をRead/Write別に分類
            $cellsRead = @()
            $cellsWritten = @()
            foreach ($c in @($feImpact.CellsAffected)) {
                $label = "$($c.SheetCodeName):$($c.RawAddress)"
                if ($c.AccessType -eq "Read") { $cellsRead += $label }
                if ($c.AccessType -eq "Write") { $cellsWritten += $label }
            }
            # テーブル参照をRead/Write別に分類
            $tablesRead = @()
            $tablesWritten = @()
            foreach ($t in @($feImpact.TablesAffected)) {
                $label = "$($t.TableName).$($t.ColumnName)"
                if ($t.AccessType -match "Read") { $tablesRead += $label }
                if ($t.AccessType -match "Write") { $tablesWritten += $label }
            }

            $formEventDetails += [PSCustomObject]@{
                Event         = $fe.ProcedureName
                EventName     = $fe.EventName
                ControlName   = $fe.ControlName
                Description   = $fe.Description
                CellsRead     = @($cellsRead)
                CellsWritten  = @($cellsWritten)
                TablesRead    = @($tablesRead)
                TablesWritten = @($tablesWritten)
            }
        }

        $subChains += [PSCustomObject]@{
            Form   = $formName
            Events = @($formEventDetails)
        }
    }
    return $subChains
}

# ============================================================
# ui_to_vba.json 生成
# ============================================================

Write-Host "Building UI-to-VBA chains..." -ForegroundColor Yellow

$chains = @()
$chainId = 0

# --- ソース1: FormControl OnAction ---
foreach ($ctrl in @($controls.Controls)) {
    if (-not $ctrl.OnAction -or $ctrl.OnAction -eq "") { continue }

    $chainId++
    $onAction = $ctrl.OnAction
    # OnAction形式: "Workbook.xlsm!SubName" or "Workbook.xlsm!Module.SubName" or "SubName"
    $procName = $onAction
    if ($procName -match '!(.+)$') { $procName = $Matches[1] }
    # Module.SubName形式の場合、SubNameだけ取得
    $moduleName = $null
    if ($procName -match '^(.+)\.(.+)$') {
        $moduleName = $Matches[1]
        $procName = $Matches[2]
    }

    # procedures.jsonからモジュール名を解決
    if (-not $moduleName -and $procedures) {
        foreach ($p in @($procedures.Procedures)) {
            if ($p.Name -eq $procName) {
                $moduleName = $p.Module
                break
            }
        }
    }

    # 影響範囲を収集
    $impact = $null
    if ($moduleName) {
        $impact = Get-ProcedureImpact -Module $moduleName -Procedure $procName
    }

    # 事前計算（PSCustomObject外で全配列構築）
    $cellsArr  = @(Convert-ImpactToCellsArray -Impact $impact)
    $tablesArr = @(Convert-ImpactToTablesArray -Impact $impact)
    $formsArr  = @(Convert-ImpactToFormsArray -Impact $impact)
    $subChainsArr = @(Build-FormSubChains -Impact $impact)

    $chains += [PSCustomObject]@{
        Id            = "chain_{0:D3}" -f $chainId
        TriggerType   = "FormControl_OnAction"
        TriggerSource = [PSCustomObject]@{
            ControlName = $ctrl.Name
            Sheet       = $ctrl.Sheet
            ControlType = $ctrl.Type
            OnAction    = $ctrl.OnAction
        }
        EntryPoint    = [PSCustomObject]@{
            Module    = $moduleName
            Procedure = $procName
        }
        CellsAffected  = $cellsArr
        TablesAffected  = $tablesArr
        FormsTriggered  = $formsArr
        SubChains       = $subChainsArr
    }
}

# --- ソース2: ActiveXイベント ---
$activexEvents = @()
foreach ($evt in @($eventTriggers.Events)) {
    if ($evt.Type -eq "ACTIVEX_EVENT") { $activexEvents += $evt }
}

foreach ($axEvt in $activexEvents) {
    $chainId++
    $impact = Get-ProcedureImpact -Module $axEvt.Module -Procedure $axEvt.ProcedureName

    # controls.jsonからActiveXコントロールの詳細を取得
    $ctrlInfo = $null
    foreach ($c in @($controls.Controls)) {
        if ($c.Name -eq $axEvt.ControlName) {
            if ($null -eq $axEvt.SheetDisplayName -or $c.Sheet -eq $axEvt.SheetDisplayName) {
                $ctrlInfo = $c
                break
            }
        }
    }

    # 事前計算（SubChains展開含む）
    $cellsArr     = @(Convert-ImpactToCellsArray -Impact $impact)
    $tablesArr    = @(Convert-ImpactToTablesArray -Impact $impact)
    $formsArr     = @(Convert-ImpactToFormsArray -Impact $impact)
    $subChainsArr = @(Build-FormSubChains -Impact $impact)

    # TriggerSourceの値を事前計算（if/elseをPSCustomObject外に）
    $triggerSheet = $axEvt.SheetDisplayName
    $triggerType  = $null
    if ($ctrlInfo) {
        $triggerSheet = $ctrlInfo.Sheet
        $triggerType  = $ctrlInfo.Type
    }

    $chains += [PSCustomObject]@{
        Id            = "chain_{0:D3}" -f $chainId
        TriggerType   = "ActiveX_Event"
        TriggerSource = [PSCustomObject]@{
            ControlName = $axEvt.ControlName
            Sheet       = $triggerSheet
            ControlType = $triggerType
            EventName   = $axEvt.EventName
        }
        EntryPoint    = [PSCustomObject]@{
            Module    = $axEvt.Module
            Procedure = $axEvt.ProcedureName
            Line      = $axEvt.Line
        }
        Description    = $axEvt.Description
        CellsAffected  = $cellsArr
        TablesAffected  = $tablesArr
        FormsTriggered  = $formsArr
        SubChains       = $subChainsArr
    }
}

# --- ソース3: ワークシートイベント ---
$wsEvents = @()
foreach ($evt in @($eventTriggers.Events)) {
    if ($evt.Type -eq "WORKSHEET_EVENT") { $wsEvents += $evt }
}

foreach ($wsEvt in $wsEvents) {
    $chainId++
    $impact = Get-ProcedureImpact -Module $wsEvt.Module -Procedure $wsEvt.ProcedureName

    # 事前計算（SubChains展開含む）
    $cellsArr     = @(Convert-ImpactToCellsArray -Impact $impact)
    $tablesArr    = @(Convert-ImpactToTablesArray -Impact $impact)
    $formsArr     = @(Convert-ImpactToFormsArray -Impact $impact)
    $subChainsArr = @(Build-FormSubChains -Impact $impact)

    $chains += [PSCustomObject]@{
        Id            = "chain_{0:D3}" -f $chainId
        TriggerType   = "Worksheet_Event"
        TriggerSource = [PSCustomObject]@{
            EventName   = $wsEvt.EventName
            Sheet       = $wsEvt.SheetDisplayName
            SheetModule = $wsEvt.Module
        }
        EntryPoint    = [PSCustomObject]@{
            Module    = $wsEvt.Module
            Procedure = $wsEvt.ProcedureName
            Line      = $wsEvt.Line
        }
        Description    = $wsEvt.Description
        CellsAffected  = $cellsArr
        TablesAffected  = $tablesArr
        FormsTriggered  = $formsArr
        SubChains       = $subChainsArr
    }
}

# --- ソース3b: ワークブックイベント ---
$wbEvents = @()
foreach ($evt in @($eventTriggers.Events)) {
    if ($evt.Type -eq "WORKBOOK_EVENT") { $wbEvents += $evt }
}

foreach ($wbEvt in $wbEvents) {
    $chainId++
    $impact = Get-ProcedureImpact -Module $wbEvt.Module -Procedure $wbEvt.ProcedureName

    # 事前計算
    $cellsArr     = @(Convert-ImpactToCellsArray -Impact $impact)
    $tablesArr    = @(Convert-ImpactToTablesArray -Impact $impact)
    $formsArr     = @(Convert-ImpactToFormsArray -Impact $impact)
    $subChainsArr = @(Build-FormSubChains -Impact $impact)

    $chains += [PSCustomObject]@{
        Id            = "chain_{0:D3}" -f $chainId
        TriggerType   = "Workbook_Event"
        TriggerSource = [PSCustomObject]@{
            EventName   = $wbEvt.EventName
            SheetModule = $wbEvt.Module
        }
        EntryPoint    = [PSCustomObject]@{
            Module    = $wbEvt.Module
            Procedure = $wbEvt.ProcedureName
            Line      = $wbEvt.Line
        }
        Description    = $wbEvt.Description
        CellsAffected  = $cellsArr
        TablesAffected  = $tablesArr
        FormsTriggered  = $formsArr
        SubChains       = $subChainsArr
    }
}

# --- ソース4: StandardModuleの公開Sub（OnAction未マッピングだが直接呼び出し可能） ---
$mappedProcNames = @()
foreach ($chain in $chains) {
    if ($chain.EntryPoint.Procedure) {
        $mappedProcNames += $chain.EntryPoint.Procedure
    }
}

if ($procedures) {
    # Publicで未マッピングのSubを収集
    $unmappedPublicSubs = @()
    foreach ($p in @($procedures.Procedures)) {
        if ($p.Access -ne "Private" -and $p.Type -eq "Sub" -and $p.Name -notin $mappedProcNames) {
            $unmappedPublicSubs += $p
        }
    }

    foreach ($sub in $unmappedPublicSubs) {
        $impact = Get-ProcedureImpact -Module $sub.Module -Procedure $sub.Name

        # フォームを呼び出すか、セルに影響があるSubのみチェーン化
        $hasCells  = @($impact.CellsAffected).Count -gt 0
        $hasTables = @($impact.TablesAffected).Count -gt 0
        $hasForms  = @($impact.FormsTriggered).Count -gt 0
        if (-not ($hasCells -or $hasTables -or $hasForms)) { continue }

        $chainId++

        # 事前計算
        $cellsArr     = @(Convert-ImpactToCellsArray -Impact $impact)
        $tablesArr    = @(Convert-ImpactToTablesArray -Impact $impact)
        $formsArr     = @(Convert-ImpactToFormsArray -Impact $impact)
        $subChainsArr = @(Build-FormSubChains -Impact $impact)

        $chains += [PSCustomObject]@{
            Id            = "chain_{0:D3}" -f $chainId
            TriggerType   = "PublicSub_Unmapped"
            TriggerSource = [PSCustomObject]@{
                Note = "Public Sub (OnAction unmapped, possibly assigned to Dashboard button)"
            }
            EntryPoint    = [PSCustomObject]@{
                Module    = $sub.Module
                Procedure = $sub.Name
                Line      = $sub.Line
            }
            CellsAffected  = $cellsArr
            TablesAffected  = $tablesArr
            FormsTriggered  = $formsArr
            SubChains       = $subChainsArr
        }
    }
}

# ui_to_vba.json 出力
$countFormControl = 0
$countActiveX     = 0
$countWsEvent     = 0
$countWbEvent     = 0
$countUnmapped    = 0
foreach ($ch in @($chains)) {
    switch ($ch.TriggerType) {
        "FormControl_OnAction" { $countFormControl++ }
        "ActiveX_Event"        { $countActiveX++ }
        "Worksheet_Event"      { $countWsEvent++ }
        "Workbook_Event"       { $countWbEvent++ }
        "PublicSub_Unmapped"   { $countUnmapped++ }
    }
}

$uiToVbaOutput = [PSCustomObject]@{
    TotalChains  = @($chains).Count
    ChainsByType = [PSCustomObject]@{
        FormControl_OnAction = $countFormControl
        ActiveX_Event        = $countActiveX
        Worksheet_Event      = $countWsEvent
        Workbook_Event       = $countWbEvent
        PublicSub_Unmapped   = $countUnmapped
    }
    Chains = @($chains)
}

$uiToVbaJson = $uiToVbaOutput | ConvertTo-Json -Depth 15
$uiToVbaJson | Out-File -FilePath (Join-Path $OutputDir "ui_to_vba.json") -Encoding utf8
$chainCount = @($chains).Count
Write-Host "[OK] ui_to_vba.json ($chainCount chains)" -ForegroundColor Green

# ============================================================
# data_flow.json 生成
# ============================================================

Write-Host ""
Write-Host "Building data flow map..." -ForegroundColor Yellow

$flows = @()

foreach ($table in @($tableRefs.Tables)) {
    foreach ($col in @($table.ColumnsAccessed)) {
        # (ListRows.Count)はメタ操作なのでスキップ
        if ($col.ColumnName -eq "(ListRows.Count)") { continue }

        # ProcedureAccessからプロシージャ単位でRead/Write分類
        $writers = @()
        $readers = @()
        foreach ($pa in @($col.ProcedureAccess)) {
            if ($pa.AccessType -eq "Write") {
                $writers += $pa.Procedure
            }
            if ($pa.AccessType -eq "Read") {
                $readers += $pa.Procedure
            }
        }

        # Sheet値を事前計算（PSCustomObject外でif/else実行）
        $sheetName = $table.SheetCodeName
        if ($table.SheetDisplayName) { $sheetName = $table.SheetDisplayName }

        # Select-Objectパイプラインも事前計算
        $uniqueWriters = @($writers | Select-Object -Unique)
        $uniqueReaders = @($readers | Select-Object -Unique)

        $flows += [PSCustomObject]@{
            Table   = $table.TableName
            Sheet   = $sheetName
            Column  = $col.ColumnName
            Writers = $uniqueWriters
            Readers = $uniqueReaders
        }
    }
}

# セル参照からもデータフローを生成（テーブル列以外のセル参照）
$cellFlows = @()
$cellRefsByAddress = @{}
foreach ($ref in @($cellRefs.References)) {
    $key = "$($ref.SheetCodeName):$($ref.RawAddress)"
    if (-not $cellRefsByAddress.ContainsKey($key)) {
        $cellRefsByAddress[$key] = @{
            SheetCodeName = $ref.SheetCodeName
            Address       = $ref.RawAddress
            IsNamedRange  = $ref.IsNamedRange
            Writers       = @()
            Readers       = @()
            Selectors     = @()
        }
    }
    $entry = $cellRefsByAddress[$key]
    $procLabel = "$($ref.Module).$($ref.Procedure)"
    switch ($ref.AccessType) {
        "Write"  { if ($procLabel -notin $entry.Writers)   { $entry.Writers   += $procLabel } }
        "Read"   { if ($procLabel -notin $entry.Readers)   { $entry.Readers   += $procLabel } }
        "Select" { if ($procLabel -notin $entry.Selectors) { $entry.Selectors += $procLabel } }
    }
}

foreach ($key in $cellRefsByAddress.Keys) {
    $entry = $cellRefsByAddress[$key]
    $cellFlows += [PSCustomObject]@{
        Type         = "CellReference"
        Sheet        = $entry.SheetCodeName
        Address      = $entry.Address
        IsNamedRange = $entry.IsNamedRange
        Writers      = @($entry.Writers)
        Readers      = @($entry.Readers)
        Selectors    = @($entry.Selectors)
    }
}

$flowCount     = @($flows).Count
$cellFlowCount = @($cellFlows).Count

$dataFlowOutput = [PSCustomObject]@{
    TotalTableColumnFlows = $flowCount
    TotalCellFlows        = $cellFlowCount
    TableColumnFlows      = @($flows)
    CellFlows             = @($cellFlows)
}

$dataFlowJson = $dataFlowOutput | ConvertTo-Json -Depth 10
$dataFlowJson | Out-File -FilePath (Join-Path $OutputDir "data_flow.json") -Encoding utf8
Write-Host "[OK] data_flow.json ($flowCount table flows, $cellFlowCount cell flows)" -ForegroundColor Green

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host " Cross-Reference Build Complete" -ForegroundColor Cyan
$totalFlows = $flowCount + $cellFlowCount
Write-Host "  UI-to-VBA Chains: $chainCount" -ForegroundColor White
Write-Host "  Data Flows:       $totalFlows" -ForegroundColor White
Write-Host "================================================================" -ForegroundColor Cyan
