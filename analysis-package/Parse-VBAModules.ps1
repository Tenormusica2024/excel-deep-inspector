<#
.SYNOPSIS
    VBA Module Static Analyzer v2 - Excel Deep Inspector
.DESCRIPTION
    .basファイルから構造化データを抽出する静的解析エンジン。
    Excel COM不要 - テキストベースの正規表現解析のみ。

    出力:
    - cell_references.json (v2) - 重複排除・シート/プロシージャコンテキスト付き
    - table_references.json     - ListObject/ListColumn参照
    - event_triggers.json       - ワークシート/ActiveX/UserFormイベント

.PARAMETER InputDir
    .basファイルが格納されたディレクトリパス
.PARAMETER OutputDir
    JSON出力先ディレクトリパス（省略時はInputDirと同じ）
.PARAMETER SheetCodeNameMap
    シートコードネーム→表示名のハッシュテーブル（省略時はモジュールヘッダーから自動推定）
#>
param(
    [Parameter(Mandatory = $true)]
    [string]$InputDir,

    [string]$OutputDir,

    [hashtable]$SheetCodeNameMap
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if (-not $OutputDir) { $OutputDir = $InputDir }

# ============================================================
# SECTION 1: モジュールヘッダーパーサー
# ============================================================
# .basファイル先頭のヘッダーコメントからモジュール情報を抽出
# フォーマット:
#   ' Module: FormOpenInvoice
#   ' Type: UserForm
#   ' Lines: 248
#   ' ---

function Parse-ModuleHeader {
    param([string[]]$Lines)

    $header = @{
        ModuleName = ""
        ModuleType = ""  # "Document" / "UserForm" / "StandardModule" / "ClassModule"
        LineCount  = 0
        CodeStartLine = 0  # ヘッダー後の実コード開始行（0-indexed）
    }

    for ($i = 0; $i -lt [Math]::Min($Lines.Count, 10); $i++) {
        $line = $Lines[$i]
        if ($line -match "^'\s*Module:\s*(.+)$") {
            $header.ModuleName = $Matches[1].Trim()
        }
        elseif ($line -match "^'\s*Type:\s*(.+)$") {
            $header.ModuleType = $Matches[1].Trim()
        }
        elseif ($line -match "^'\s*Lines:\s*(\d+)") {
            $header.LineCount = [int]$Matches[1]
        }
        elseif ($line -match "^'\s*---\s*$") {
            $header.CodeStartLine = $i + 1
            break
        }
    }

    return $header
}

# ============================================================
# SECTION 2: プロシージャ境界追跡エンジン
# ============================================================
# 各行がどのSub/Function内にあるかを追跡する状態マシン。
# VBAの行継続（末尾 _）にも対応。

function Get-ProcedureBoundaries {
    param(
        [string[]]$Lines,
        [int]$CodeStartLine = 0
    )

    $boundaries = @()
    $currentProc = $null

    for ($i = $CodeStartLine; $i -lt $Lines.Count; $i++) {
        $line = $Lines[$i]
        $trimmed = $line.TrimStart()

        # Sub/Function/Property宣言の検出
        if ($trimmed -match '^(Public\s+|Private\s+|Friend\s+)?(Static\s+)?(Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)\s*\(') {
            $currentProc = @{
                Access    = if ($Matches[1]) { $Matches[1].Trim() } else { "Public" }
                Type      = $Matches[3].Trim()
                Name      = $Matches[4]
                StartLine = $i + 1  # 1-indexed（VBAの行番号に合わせる）
                EndLine   = -1
            }
        }

        # End Sub/Function/Property の検出
        if ($trimmed -match '^End\s+(Sub|Function|Property)\b' -and $null -ne $currentProc) {
            $currentProc.EndLine = $i + 1  # 1-indexed
            $boundaries += [PSCustomObject]$currentProc
            $currentProc = $null
        }
    }

    return ,$boundaries
}

function Get-ProcedureAtLine {
    <# 指定行番号（1-indexed）がどのプロシージャ内にあるかを返す #>
    param(
        [PSObject[]]$Boundaries,
        [int]$LineNumber
    )

    foreach ($proc in $Boundaries) {
        if ($LineNumber -ge $proc.StartLine -and $LineNumber -le $proc.EndLine) {
            return $proc
        }
    }
    return $null
}

# ============================================================
# SECTION 3: セル参照パーサー v2
# ============================================================
# v1からの改善点:
# - 重複排除（1マッチ=1分類）
# - シートコンテキスト捕捉（shXxx. / Me. / Worksheets("X").）
# - プロシージャ名記録
# - AccessType判定（Read/Write/Select/Delete）
# - 確信度スコア

function Get-AccessType {
    <# セル参照のアクセスタイプを文脈から推定 #>
    param(
        [string]$Line,        # 行全体
        [int]$MatchEnd        # マッチ終了位置
    )

    $suffix = ""
    if ($MatchEnd -lt $Line.Length) {
        $suffix = $Line.Substring($MatchEnd)
    }

    # If/ElseIf 内の = は比較演算子 → 全てRead扱い
    $trimmedLine = $Line.TrimStart()
    $isComparison = ($trimmedLine -match '^\s*(If|ElseIf)\b')

    # .Select パターン
    if ($suffix -match '^\s*\.Select\b') { return "Select" }
    # .Clear / .ClearContents / .Delete パターン
    if ($suffix -match '^\s*\.(Clear\w*|Delete)\b') { return "Delete" }

    if (-not $isComparison) {
        # .Value = expr（セル値への書き込み）
        if ($suffix -match '^\s*\.Value\d?\s*=\s*(?!=)') { return "Write" }
        # 直接代入: Range("X") = expr
        if ($suffix -match '^\s*=\s*(?!=)') { return "Write" }
        # .DataBodyRange(n).Value = expr（テーブル列への書き込み）
        if ($suffix -match '\.DataBodyRange\(\w+\)\.Value\d?\s*=\s*(?!=)') { return "Write" }
    }

    return "Read"
}

function Test-IsCellAddress {
    <# 文字列がExcelセルアドレスパターン（A1, B10, AA100等）にマッチするか判定 #>
    param([string]$Text)
    return $Text -match '^[A-Z]{1,3}\d+(?::[A-Z]{1,3}\d+)?$'
}

function Extract-CellReferencesV2 {
    param(
        [string[]]$Lines,
        [string]$ModuleName,
        [string]$ModuleType,
        [PSObject[]]$ProcBoundaries,
        [string[]]$KnownSheetCodeNames,
        [hashtable]$CodeNameToDisplay
    )

    $references = @()
    $refId = 0

    # シートコードネームの正規表現パターンを構築
    # shDash, shInvoice, shTemp, Me, ActiveSheet 等
    # @()で囲んで常に配列化（PS5.1で要素1つの場合のスカラー化を防止）
    $sheetPrefixes = @($KnownSheetCodeNames)

    for ($i = 0; $i -lt $Lines.Count; $i++) {
        $line = $Lines[$i]
        $trimmed = $line.TrimStart()
        $lineNum = $i + 1  # 1-indexed

        # コメント行はスキップ
        if ($trimmed.StartsWith("'") -or $trimmed.StartsWith("Rem ")) { continue }

        # 所属プロシージャを特定
        $proc = Get-ProcedureAtLine -Boundaries $ProcBoundaries -LineNumber $lineNum
        $procName = if ($null -ne $proc) { $proc.Name } else { "(module-level)" }
        $procType = if ($null -ne $proc) { $proc.Type } else { "" }

        # このラインで既にマッチした文字位置を記録（重複排除用）
        $matchedPositions = @{}

        # -----------------------------------------------------------
        # パターン1: sheetCodeName.Range("X") - 最高優先度
        # -----------------------------------------------------------
        if (@($sheetPrefixes).Count -gt 0) {
            $sheetPattern = '(?<prefix>' + ($sheetPrefixes -join '|') + ')\.Range\(\s*"(?<addr>[^"]+)"\s*\)'
            $matches1 = [regex]::Matches($line, $sheetPattern)
            foreach ($m in $matches1) {
                $refId++
                $prefix = $m.Groups['prefix'].Value
                $addr = $m.Groups['addr'].Value
                $displayName = if ($CodeNameToDisplay -and $CodeNameToDisplay.ContainsKey($prefix)) {
                    $CodeNameToDisplay[$prefix]
                } else { $null }
                $isNamed = -not (Test-IsCellAddress $addr)
                $accessType = Get-AccessType -Line $line -MatchEnd ($m.Index + $m.Length)

                $references += [PSCustomObject]@{
                    Id               = "ref_{0:D3}" -f $refId
                    Pattern          = "SHEET_RANGE"
                    Module           = $ModuleName
                    Procedure        = $procName
                    ProcedureType    = $procType
                    Line             = $lineNum
                    SheetCodeName    = $prefix
                    SheetDisplayName = $displayName
                    RawAddress       = $addr
                    IsNamedRange     = $isNamed
                    Confidence       = 0.95
                    AccessType       = $accessType
                    Context          = $trimmed
                }
                # この位置をマッチ済みとして記録
                $matchedPositions[$m.Index] = $m.Length
            }
        }

        # -----------------------------------------------------------
        # パターン2: Me.Range("X") - Document/UserFormモジュール内
        # -----------------------------------------------------------
        $meRangeMatches = [regex]::Matches($line, 'Me\.Range\(\s*"(?<addr>[^"]+)"\s*\)')
        foreach ($m in $meRangeMatches) {
            # 既にマッチ済み位置と重複していないか確認
            if ($matchedPositions.ContainsKey($m.Index)) { continue }

            $refId++
            $addr = $m.Groups['addr'].Value
            $isNamed = -not (Test-IsCellAddress $addr)
            # Me = 自モジュール。Document型ならシートコードネーム = モジュール名
            $selfCodeName = if ($ModuleType -eq "Document") { $ModuleName } else { $null }
            $selfDisplay = if ($selfCodeName -and $CodeNameToDisplay -and $CodeNameToDisplay.ContainsKey($selfCodeName)) {
                $CodeNameToDisplay[$selfCodeName]
            } else { $null }
            $accessType = Get-AccessType -Line $line -MatchEnd ($m.Index + $m.Length)

            $references += [PSCustomObject]@{
                Id               = "ref_{0:D3}" -f $refId
                Pattern          = "ME_RANGE"
                Module           = $ModuleName
                Procedure        = $procName
                ProcedureType    = $procType
                Line             = $lineNum
                SheetCodeName    = $selfCodeName
                SheetDisplayName = $selfDisplay
                RawAddress       = $addr
                IsNamedRange     = $isNamed
                Confidence       = 0.90
                AccessType       = $accessType
                Context          = $trimmed
            }
            $matchedPositions[$m.Index] = $m.Length
        }

        # -----------------------------------------------------------
        # パターン3: Worksheets("X").Range("Y") / Sheets("X").Range("Y")
        # -----------------------------------------------------------
        $wsRangeMatches = [regex]::Matches($line, '(?:Worksheets|Sheets)\(\s*"(?<sheet>[^"]+)"\s*\)\.Range\(\s*"(?<addr>[^"]+)"\s*\)')
        foreach ($m in $wsRangeMatches) {
            if ($matchedPositions.ContainsKey($m.Index)) { continue }

            $refId++
            $sheetName = $m.Groups['sheet'].Value
            $addr = $m.Groups['addr'].Value
            $isNamed = -not (Test-IsCellAddress $addr)
            $accessType = Get-AccessType -Line $line -MatchEnd ($m.Index + $m.Length)

            $references += [PSCustomObject]@{
                Id               = "ref_{0:D3}" -f $refId
                Pattern          = "WORKSHEETS_RANGE"
                Module           = $ModuleName
                Procedure        = $procName
                ProcedureType    = $procType
                Line             = $lineNum
                SheetCodeName    = $null
                SheetDisplayName = $sheetName
                RawAddress       = $addr
                IsNamedRange     = $isNamed
                Confidence       = 0.95
                AccessType       = $accessType
                Context          = $trimmed
            }
            $matchedPositions[$m.Index] = $m.Length
        }

        # -----------------------------------------------------------
        # パターン4: ActiveSheet.Range("X")
        # -----------------------------------------------------------
        $asRangeMatches = [regex]::Matches($line, 'ActiveSheet\.Range\(\s*"(?<addr>[^"]+)"\s*\)')
        foreach ($m in $asRangeMatches) {
            if ($matchedPositions.ContainsKey($m.Index)) { continue }

            $refId++
            $addr = $m.Groups['addr'].Value
            $isNamed = -not (Test-IsCellAddress $addr)
            $accessType = Get-AccessType -Line $line -MatchEnd ($m.Index + $m.Length)

            $references += [PSCustomObject]@{
                Id               = "ref_{0:D3}" -f $refId
                Pattern          = "ACTIVESHEET_RANGE"
                Module           = $ModuleName
                Procedure        = $procName
                ProcedureType    = $procType
                Line             = $lineNum
                SheetCodeName    = $null
                SheetDisplayName = $null
                RawAddress       = $addr
                IsNamedRange     = $isNamed
                Confidence       = 0.40
                AccessType       = $accessType
                Context          = $trimmed
            }
            $matchedPositions[$m.Index] = $m.Length
        }

        # -----------------------------------------------------------
        # パターン5: 単独 Range("X") - プレフィックスなし
        # 上位パターンで捕捉されなかった残りのRange()のみ
        # -----------------------------------------------------------
        $plainRangeMatches = [regex]::Matches($line, '(?<![.\w])Range\(\s*"(?<addr>[^"]+)"\s*\)')
        foreach ($m in $plainRangeMatches) {
            # 重複チェック: この位置が上位パターンの一部として既にマッチされていないか
            $alreadyMatched = $false
            foreach ($pos in $matchedPositions.Keys) {
                if ($m.Index -ge $pos -and $m.Index -lt ($pos + $matchedPositions[$pos])) {
                    $alreadyMatched = $true
                    break
                }
            }
            if ($alreadyMatched) { continue }

            $refId++
            $addr = $m.Groups['addr'].Value
            $isNamed = -not (Test-IsCellAddress $addr)
            $patternName = if ($isNamed) { "RANGE_NAMED" } else { "RANGE_LITERAL" }
            $confidence = if ($isNamed) { 0.80 } else { 0.85 }
            $accessType = Get-AccessType -Line $line -MatchEnd ($m.Index + $m.Length)

            $references += [PSCustomObject]@{
                Id               = "ref_{0:D3}" -f $refId
                Pattern          = $patternName
                Module           = $ModuleName
                Procedure        = $procName
                ProcedureType    = $procType
                Line             = $lineNum
                SheetCodeName    = $null
                SheetDisplayName = $null
                RawAddress       = $addr
                IsNamedRange     = $isNamed
                Confidence       = $confidence
                AccessType       = $accessType
                Context          = $trimmed
            }
            $matchedPositions[$m.Index] = $m.Length
        }

        # -----------------------------------------------------------
        # パターン6: sheetCodeName.Cells(X, Y)
        # -----------------------------------------------------------
        if (@($sheetPrefixes).Count -gt 0) {
            $sheetCellsPattern = '(?<prefix>' + ($sheetPrefixes -join '|') + ')\.Cells\(\s*(?<row>[^,]+?)\s*,\s*(?<col>[^)]+?)\s*\)'
            $scMatches = [regex]::Matches($line, $sheetCellsPattern)
            foreach ($m in $scMatches) {
                $refId++
                $prefix = $m.Groups['prefix'].Value
                $row = $m.Groups['row'].Value.Trim()
                $col = $m.Groups['col'].Value.Trim()
                $displayName = if ($CodeNameToDisplay -and $CodeNameToDisplay.ContainsKey($prefix)) {
                    $CodeNameToDisplay[$prefix]
                } else { $null }
                # 両方数値ならリテラル（高確信度）、変数含むなら低確信度
                $bothNumeric = ($row -match '^\d+$') -and ($col -match '^\d+$')
                $confidence = if ($bothNumeric) { 0.90 } else { 0.50 }
                $accessType = Get-AccessType -Line $line -MatchEnd ($m.Index + $m.Length)

                $references += [PSCustomObject]@{
                    Id               = "ref_{0:D3}" -f $refId
                    Pattern          = "SHEET_CELLS"
                    Module           = $ModuleName
                    Procedure        = $procName
                    ProcedureType    = $procType
                    Line             = $lineNum
                    SheetCodeName    = $prefix
                    SheetDisplayName = $displayName
                    RawAddress       = "Cells($row, $col)"
                    IsNamedRange     = $false
                    Confidence       = $confidence
                    AccessType       = $accessType
                    Context          = $trimmed
                }
                $matchedPositions[$m.Index] = $m.Length
            }
        }

        # -----------------------------------------------------------
        # パターン7: Me.Cells(X, Y)
        # -----------------------------------------------------------
        $meCellsMatches = [regex]::Matches($line, 'Me\.Cells\(\s*(?<row>[^,]+?)\s*,\s*(?<col>[^)]+?)\s*\)')
        foreach ($m in $meCellsMatches) {
            if ($matchedPositions.ContainsKey($m.Index)) { continue }

            $refId++
            $row = $m.Groups['row'].Value.Trim()
            $col = $m.Groups['col'].Value.Trim()
            $selfCodeName = if ($ModuleType -eq "Document") { $ModuleName } else { $null }
            $selfDisplay = if ($selfCodeName -and $CodeNameToDisplay -and $CodeNameToDisplay.ContainsKey($selfCodeName)) {
                $CodeNameToDisplay[$selfCodeName]
            } else { $null }
            $bothNumeric = ($row -match '^\d+$') -and ($col -match '^\d+$')
            $confidence = if ($bothNumeric) { 0.85 } else { 0.50 }
            $accessType = Get-AccessType -Line $line -MatchEnd ($m.Index + $m.Length)

            $references += [PSCustomObject]@{
                Id               = "ref_{0:D3}" -f $refId
                Pattern          = "ME_CELLS"
                Module           = $ModuleName
                Procedure        = $procName
                ProcedureType    = $procType
                Line             = $lineNum
                SheetCodeName    = $selfCodeName
                SheetDisplayName = $selfDisplay
                RawAddress       = "Cells($row, $col)"
                IsNamedRange     = $false
                Confidence       = $confidence
                AccessType       = $accessType
                Context          = $trimmed
            }
            $matchedPositions[$m.Index] = $m.Length
        }

        # -----------------------------------------------------------
        # パターン8: 単独 Cells(X, Y) - プレフィックスなし
        # -----------------------------------------------------------
        $plainCellsMatches = [regex]::Matches($line, '(?<![.\w])Cells\(\s*(?<row>[^,]+?)\s*,\s*(?<col>[^)]+?)\s*\)')
        foreach ($m in $plainCellsMatches) {
            $alreadyMatched = $false
            foreach ($pos in $matchedPositions.Keys) {
                if ($m.Index -ge $pos -and $m.Index -lt ($pos + $matchedPositions[$pos])) {
                    $alreadyMatched = $true
                    break
                }
            }
            if ($alreadyMatched) { continue }

            $refId++
            $row = $m.Groups['row'].Value.Trim()
            $col = $m.Groups['col'].Value.Trim()
            $bothNumeric = ($row -match '^\d+$') -and ($col -match '^\d+$')
            $patternName = if ($bothNumeric) { "CELLS_LITERAL" } else { "CELLS_VARIABLE" }
            $confidence = if ($bothNumeric) { 0.85 } else { 0.50 }
            $accessType = Get-AccessType -Line $line -MatchEnd ($m.Index + $m.Length)

            $references += [PSCustomObject]@{
                Id               = "ref_{0:D3}" -f $refId
                Pattern          = $patternName
                Module           = $ModuleName
                Procedure        = $procName
                ProcedureType    = $procType
                Line             = $lineNum
                SheetCodeName    = $null
                SheetDisplayName = $null
                RawAddress       = "Cells($row, $col)"
                IsNamedRange     = $false
                Confidence       = $confidence
                AccessType       = $accessType
                Context          = $trimmed
            }
        }
    }

    return ,$references
}

# ============================================================
# SECTION 4: テーブル参照パーサー
# ============================================================
# ListObject（構造化テーブル）とListColumn参照を検出。
# Set InvoiceTable = shInvoice.ListObjects("InvoiceTable")
# InvoiceTable.ListColumns("Status").DataBodyRange(r).Value

function Extract-TableReferences {
    param(
        [string[]]$Lines,
        [string]$ModuleName,
        [PSObject[]]$ProcBoundaries,
        [string[]]$KnownSheetCodeNames,
        [hashtable]$CodeNameToDisplay
    )

    # Phase 1: ListObjectsバインドの検出
    # パターン: Set {var} = {sheet}.ListObjects("TableName")
    $bindings = @()
    # テーブル名 → シートコードネームのマッピング
    $tableToSheet = @{}

    for ($i = 0; $i -lt $Lines.Count; $i++) {
        $line = $Lines[$i]
        $trimmed = $line.TrimStart()
        $lineNum = $i + 1

        if ($trimmed.StartsWith("'") -or $trimmed.StartsWith("Rem ")) { continue }

        $proc = Get-ProcedureAtLine -Boundaries $ProcBoundaries -LineNumber $lineNum

        # Set {var} = {prefix}.ListObjects("TableName")
        # prefix: shXxx / Me / Worksheets("X") / ActiveSheet
        if ($trimmed -match 'Set\s+(?<var>\w+)\s*=\s*(?<prefix>[\w.]+?)\.ListObjects\(\s*"(?<table>[^"]+)"\s*\)') {
            $prefix = $Matches['prefix']
            $tableName = $Matches['table']
            $varName = $Matches['var']
            $procName = if ($null -ne $proc) { $proc.Name } else { "(module-level)" }

            # プレフィックスからシート解決
            $sheetCodeName = $null
            $sheetDisplayName = $null
            if ($prefix -eq "Me" -and $ModuleName) {
                $sheetCodeName = $ModuleName
            }
            elseif ($KnownSheetCodeNames -contains $prefix) {
                $sheetCodeName = $prefix
            }
            if ($sheetCodeName -and $CodeNameToDisplay -and $CodeNameToDisplay.ContainsKey($sheetCodeName)) {
                $sheetDisplayName = $CodeNameToDisplay[$sheetCodeName]
            }

            $bindings += [PSCustomObject]@{
                Module    = $ModuleName
                Procedure = $procName
                Line      = $lineNum
                Variable  = $varName
                TableName = $tableName
                SheetCodeName    = $sheetCodeName
                SheetDisplayName = $sheetDisplayName
                Context   = $trimmed
            }

            # テーブル→シートマッピングを記録
            if ($sheetCodeName -and -not $tableToSheet.ContainsKey($tableName)) {
                $tableToSheet[$tableName] = @{
                    CodeName    = $sheetCodeName
                    DisplayName = $sheetDisplayName
                }
            }
        }
    }

    # Phase 2: ListColumn参照の検出
    # パターン: {var}.ListColumns("ColumnName")
    # パターン: {table}.ListColumns("ColumnName").DataBodyRange({idx}).Value
    $columnAccesses = @()

    for ($i = 0; $i -lt $Lines.Count; $i++) {
        $line = $Lines[$i]
        $trimmed = $line.TrimStart()
        $lineNum = $i + 1

        if ($trimmed.StartsWith("'") -or $trimmed.StartsWith("Rem ")) { continue }

        $proc = Get-ProcedureAtLine -Boundaries $ProcBoundaries -LineNumber $lineNum
        $procName = if ($null -ne $proc) { $proc.Name } else { "(module-level)" }

        # .ListColumns("ColumnName") の検出
        $colMatches = [regex]::Matches($line, '\.ListColumns\(\s*"(?<col>[^"]+)"\s*\)')
        foreach ($m in $colMatches) {
            $colName = $m.Groups['col'].Value

            # アクセスタイプ判定: .DataBodyRange(x).Value = expr → Write
            # ただし If/ElseIf 内の = は比較演算子（Read）
            $suffix = $line.Substring($m.Index + $m.Length)
            $accessType = "Read"
            $isComparison = ($trimmed -match '^\s*(If|ElseIf)\b')
            if (-not $isComparison) {
                if ($suffix -match '\.DataBodyRange\([^)]+\)\.Value\d?\s*=\s*(?!=)') {
                    $accessType = "Write"
                }
                elseif ($suffix -match '\.DataBodyRange\s*=\s*(?!=)') {
                    $accessType = "Write"
                }
            }

            # どのテーブルに属するか推定（同じプロシージャ内のバインドから）
            $tableName = $null
            foreach ($b in $bindings) {
                if ($b.Module -eq $ModuleName -and $b.Procedure -eq $procName) {
                    $tableName = $b.TableName
                    break
                }
            }

            $columnAccesses += [PSCustomObject]@{
                Module     = $ModuleName
                Procedure  = $procName
                Line       = $lineNum
                ColumnName = $colName
                AccessType = $accessType
                TableName  = $tableName
                Context    = $trimmed
            }
        }

        # .ListRows.Count の検出
        if ($trimmed -match '\.ListRows\.Count') {
            $tableName = $null
            foreach ($b in $bindings) {
                if ($b.Module -eq $ModuleName -and $b.Procedure -eq $procName) {
                    $tableName = $b.TableName
                    break
                }
            }
            $columnAccesses += [PSCustomObject]@{
                Module     = $ModuleName
                Procedure  = $procName
                Line       = $lineNum
                ColumnName = "(ListRows.Count)"
                AccessType = "Read"
                TableName  = $tableName
                Context    = $trimmed
            }
        }
    }

    return @{
        Bindings       = $bindings
        ColumnAccesses = $columnAccesses
        TableToSheet   = $tableToSheet
    }
}

# ============================================================
# SECTION 5: イベントトリガーパーサー
# ============================================================
# Worksheet_Change, ActiveXイベント（axCollapse_Click等）、
# UserFormイベント（btCreateInvoice_Click等）を検出。

function Extract-EventTriggers {
    param(
        [string[]]$Lines,
        [string]$ModuleName,
        [string]$ModuleType,
        [PSObject[]]$ProcBoundaries,
        [hashtable]$CodeNameToDisplay
    )

    $events = @()

    foreach ($proc in $ProcBoundaries) {
        $eventInfo = $null

        # ワークシートイベント: Worksheet_Change, Worksheet_Activate, etc.
        if ($proc.Name -match '^Worksheet_(?<event>\w+)$') {
            $eventInfo = [PSCustomObject]@{
                Type          = "WORKSHEET_EVENT"
                Module        = $ModuleName
                SheetDisplayName = if ($CodeNameToDisplay -and $CodeNameToDisplay.ContainsKey($ModuleName)) {
                    $CodeNameToDisplay[$ModuleName]
                } else { $null }
                ControlName   = $null
                ControlType   = $null
                EventName     = $Matches['event']
                ProcedureName = $proc.Name
                Line          = $proc.StartLine
                Access        = $proc.Access
            }
        }
        # ワークブックイベント: Workbook_Open, Workbook_BeforeClose, etc.
        elseif ($proc.Name -match '^Workbook_(?<event>\w+)$') {
            $eventInfo = [PSCustomObject]@{
                Type          = "WORKBOOK_EVENT"
                Module        = $ModuleName
                SheetDisplayName = $null
                ControlName   = $null
                ControlType   = $null
                EventName     = $Matches['event']
                ProcedureName = $proc.Name
                Line          = $proc.StartLine
                Access        = $proc.Access
            }
        }
        # Document型モジュール内の非Worksheet_*/非Workbook_*イベント → ActiveXイベント候補
        elseif ($ModuleType -eq "Document" -and $proc.Access -eq "Private" -and $proc.Name -match '^(?<ctrl>\w+)_(?<event>\w+)$') {
            $ctrlName = $Matches['ctrl']
            $eventName = $Matches['event']
            # Worksheet_ / Workbook_ でなければActiveXイベント候補
            if ($ctrlName -ne "Worksheet" -and $ctrlName -ne "Workbook") {
                $eventInfo = [PSCustomObject]@{
                    Type          = "ACTIVEX_EVENT"
                    Module        = $ModuleName
                    SheetDisplayName = if ($CodeNameToDisplay -and $CodeNameToDisplay.ContainsKey($ModuleName)) {
                        $CodeNameToDisplay[$ModuleName]
                    } else { $null }
                    ControlName   = $ctrlName
                    ControlType   = $null  # controls.jsonとの照合で後から設定
                    EventName     = $eventName
                    ProcedureName = $proc.Name
                    Line          = $proc.StartLine
                    Access        = $proc.Access
                }
            }
        }
        # UserForm型モジュール内のイベント
        elseif ($ModuleType -eq "UserForm" -and $proc.Access -eq "Private" -and $proc.Name -match '^(?<ctrl>\w+)_(?<event>\w+)$') {
            $ctrlName = $Matches['ctrl']
            $eventName = $Matches['event']
            $eventInfo = [PSCustomObject]@{
                Type          = "USERFORM_EVENT"
                Module        = $ModuleName
                SheetDisplayName = $null
                ControlName   = $ctrlName
                ControlType   = $null
                EventName     = $eventName
                ProcedureName = $proc.Name
                Line          = $proc.StartLine
                Access        = $proc.Access
            }
        }

        if ($null -ne $eventInfo) {
            # プロシージャ本体からコメントを抽出（最初のコメント行 = 説明）
            $description = ""
            for ($j = $proc.StartLine; $j -lt [Math]::Min($proc.StartLine + 5, $Lines.Count); $j++) {
                $bodyLine = $Lines[$j].TrimStart()
                if ($bodyLine.StartsWith("'")) {
                    $description = $bodyLine.TrimStart("'").Trim()
                    break
                }
            }
            $eventInfo | Add-Member -NotePropertyName "Description" -NotePropertyValue $description

            $events += $eventInfo
        }
    }

    return ,$events
}

# ============================================================
# SECTION 6: フォーム呼び出し検出
# ============================================================
# FormXxx.Show パターンを検出してフォーム遷移を追跡

function Extract-FormCalls {
    param(
        [string[]]$Lines,
        [string]$ModuleName,
        [PSObject[]]$ProcBoundaries
    )

    $formCalls = @()

    for ($i = 0; $i -lt $Lines.Count; $i++) {
        $line = $Lines[$i]
        $trimmed = $line.TrimStart()
        $lineNum = $i + 1

        if ($trimmed.StartsWith("'") -or $trimmed.StartsWith("Rem ")) { continue }

        # FormXxx.Show / Load FormXxx
        if ($trimmed -match '(?:(?<form>\w+)\.Show|Load\s+(?<form2>\w+))') {
            $formName = if ($Matches['form']) { $Matches['form'] } else { $Matches['form2'] }
            # Formで始まる名前、またはufrmで始まる名前をフォームとして判定
            if ($formName -match '^(Form|ufrm|frm|frmMain)') {
                $proc = Get-ProcedureAtLine -Boundaries $ProcBoundaries -LineNumber $lineNum
                $procName = if ($null -ne $proc) { $proc.Name } else { "(module-level)" }

                $formCalls += [PSCustomObject]@{
                    Module        = $ModuleName
                    Procedure     = $procName
                    Line          = $lineNum
                    FormName      = $formName
                    CallType      = if ($trimmed -match '\.Show') { "Show" } else { "Load" }
                    Context       = $trimmed
                }
            }
        }
    }

    return ,$formCalls
}

# ============================================================
# SECTION 7: メインオーケストレーター
# ============================================================

function Invoke-VBAAnalysis {
    param(
        [string]$InputDir,
        [string]$OutputDir,
        [hashtable]$SheetCodeNameMap
    )

    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host " VBA Module Static Analyzer v2" -ForegroundColor Cyan
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host "Input : $InputDir"
    Write-Host "Output: $OutputDir"
    Write-Host ""

    # .basファイルを収集
    $basFiles = Get-ChildItem -Path $InputDir -Filter "*.bas" -File
    if ($basFiles.Count -eq 0) {
        Write-Warning ".basファイルが見つかりません: $InputDir"
        return
    }
    Write-Host "Found $($basFiles.Count) .bas files" -ForegroundColor Green

    # Phase 1: 全モジュールのヘッダーを解析してメタデータ収集
    $modules = @()
    $knownSheetCodeNames = @()
    $codeNameToDisplay = @{}

    foreach ($file in $basFiles) {
        $content = Get-Content -Path $file.FullName -Encoding UTF8
        $header = Parse-ModuleHeader -Lines $content

        if (-not $header.ModuleName) {
            # ヘッダーがない場合はファイル名から推定
            $header.ModuleName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
        }

        $modules += @{
            File       = $file
            Lines      = $content
            Header     = $header
        }

        # Document型モジュール → シートコードネーム候補
        if ($header.ModuleType -eq "Document") {
            $knownSheetCodeNames += $header.ModuleName
        }
    }

    # 外部提供のシートコードネームマッピングがあれば使用
    if ($SheetCodeNameMap) {
        $codeNameToDisplay = $SheetCodeNameMap
    }
    else {
        # 自動推定: Document型モジュール名をそのままキーとして登録（DisplayNameはnull）
        foreach ($cn in $knownSheetCodeNames) {
            $codeNameToDisplay[$cn] = $null
        }
    }

    Write-Host "Sheet CodeNames detected: $($knownSheetCodeNames -join ', ')" -ForegroundColor Yellow

    # Phase 2: 各モジュールを解析
    $allCellRefs = @()
    $allTableBindings = @()
    $allColumnAccesses = @()
    $allTableToSheet = @{}
    $allEvents = @()
    $allFormCalls = @()
    $allProcedures = @()

    foreach ($mod in $modules) {
        $lines = $mod.Lines
        $name = $mod.Header.ModuleName
        $type = $mod.Header.ModuleType
        $codeStart = $mod.Header.CodeStartLine

        Write-Host "  Parsing: $name ($type)" -ForegroundColor Gray

        # プロシージャ境界を取得
        $procBounds = Get-ProcedureBoundaries -Lines $lines -CodeStartLine $codeStart
        foreach ($p in $procBounds) {
            $p | Add-Member -NotePropertyName "Module" -NotePropertyValue $name -Force
            $allProcedures += $p
        }

        # セル参照パース（v2）
        $cellRefs = Extract-CellReferencesV2 `
            -Lines $lines `
            -ModuleName $name `
            -ModuleType $type `
            -ProcBoundaries $procBounds `
            -KnownSheetCodeNames $knownSheetCodeNames `
            -CodeNameToDisplay $codeNameToDisplay
        $allCellRefs += $cellRefs

        # テーブル参照パース
        $tableResult = Extract-TableReferences `
            -Lines $lines `
            -ModuleName $name `
            -ProcBoundaries $procBounds `
            -KnownSheetCodeNames $knownSheetCodeNames `
            -CodeNameToDisplay $codeNameToDisplay
        $allTableBindings += $tableResult.Bindings
        $allColumnAccesses += $tableResult.ColumnAccesses
        foreach ($key in $tableResult.TableToSheet.Keys) {
            if (-not $allTableToSheet.ContainsKey($key)) {
                $allTableToSheet[$key] = $tableResult.TableToSheet[$key]
            }
        }

        # イベントトリガーパース
        $events = Extract-EventTriggers `
            -Lines $lines `
            -ModuleName $name `
            -ModuleType $type `
            -ProcBoundaries $procBounds `
            -CodeNameToDisplay $codeNameToDisplay
        $allEvents += $events

        # フォーム呼び出しパース
        $formCalls = Extract-FormCalls `
            -Lines $lines `
            -ModuleName $name `
            -ProcBoundaries $procBounds
        $allFormCalls += $formCalls
    }

    # Phase 3: 結果を集約してJSON出力

    # --- cell_references.json (v2) ---
    $patternSummary = @{}
    foreach ($ref in $allCellRefs) {
        $p = $ref.Pattern
        if ($patternSummary.ContainsKey($p)) { $patternSummary[$p]++ }
        else { $patternSummary[$p] = 1 }
    }
    $cellRefsOutput = [PSCustomObject]@{
        Version         = "2.0"
        TotalReferences = @($allCellRefs).Count
        PatternSummary  = $patternSummary
        References      = @($allCellRefs)
    }
    $cellRefsJson = $cellRefsOutput | ConvertTo-Json -Depth 10
    $cellRefsJson | Out-File -FilePath (Join-Path $OutputDir "cell_references_v2.json") -Encoding utf8
    Write-Host "  [OK] cell_references_v2.json ($(@($allCellRefs).Count) refs)" -ForegroundColor Green

    # --- table_references.json ---
    # テーブルごとに集約（@()で常に配列化 - PS5.1スカラー化防止）
    $tableNames = @($allTableBindings | Select-Object -ExpandProperty TableName -Unique)
    $tableOutput = @()
    foreach ($tName in $tableNames) {
        if (-not $tName) { continue }
        $tBindings = @($allTableBindings | Where-Object { $_.TableName -eq $tName })
        $tColumns = @($allColumnAccesses | Where-Object { $_.TableName -eq $tName })

        # 列ごとのアクセス集約
        $columnSummary = @{}
        foreach ($col in $tColumns) {
            $cn = $col.ColumnName
            if (-not $columnSummary.ContainsKey($cn)) {
                $columnSummary[$cn] = @{
                    ColumnName = $cn
                    AccessTypes = @()
                    Procedures = @()
                }
            }
            if ($col.AccessType -notin $columnSummary[$cn].AccessTypes) {
                $columnSummary[$cn].AccessTypes += $col.AccessType
            }
            if ($col.Procedure -notin $columnSummary[$cn].Procedures) {
                $columnSummary[$cn].Procedures += $col.Procedure
            }
        }

        $sheetInfo = $allTableToSheet[$tName]
        $tableOutput += [PSCustomObject]@{
            TableName        = $tName
            SheetCodeName    = if ($sheetInfo) { $sheetInfo.CodeName } else { $null }
            SheetDisplayName = if ($sheetInfo) { $sheetInfo.DisplayName } else { $null }
            BoundIn          = @($tBindings | Select-Object Module, Procedure, Line, Variable, Context)
            ColumnsAccessed  = @($columnSummary.Values | ForEach-Object {
                [PSCustomObject]@{
                    ColumnName  = $_.ColumnName
                    AccessType  = ($_.AccessTypes -join "/")
                    Procedures  = $_.Procedures
                }
            })
        }
    }
    $uniqueColCount = @($allColumnAccesses | Select-Object -Property ColumnName -Unique).Count
    $tableRefsOutput = [PSCustomObject]@{
        TotalTables         = @($tableOutput).Count
        TotalColumnsAccessed = $uniqueColCount
        Tables              = @($tableOutput)
    }
    $tableRefsJson = $tableRefsOutput | ConvertTo-Json -Depth 10
    $tableRefsJson | Out-File -FilePath (Join-Path $OutputDir "table_references.json") -Encoding utf8
    Write-Host "  [OK] table_references.json ($($tableOutput.Count) tables)" -ForegroundColor Green

    # --- event_triggers.json ---
    $eventSummary = @{}
    foreach ($evt in $allEvents) {
        $t = $evt.Type
        if ($eventSummary.ContainsKey($t)) { $eventSummary[$t]++ }
        else { $eventSummary[$t] = 1 }
    }
    $eventsOutput = [PSCustomObject]@{
        TotalEvents = @($allEvents).Count
        Summary     = $eventSummary
        Events      = @($allEvents)
    }
    $eventsJson = $eventsOutput | ConvertTo-Json -Depth 10
    $eventsJson | Out-File -FilePath (Join-Path $OutputDir "event_triggers.json") -Encoding utf8
    Write-Host "  [OK] event_triggers.json ($(@($allEvents).Count) events)" -ForegroundColor Green

    # --- form_calls.json（補助データ - クロスリファレンス生成用）---
    $formCallsOutput = [PSCustomObject]@{
        TotalCalls = @($allFormCalls).Count
        Calls      = @($allFormCalls)
    }
    $formCallsJson = $formCallsOutput | ConvertTo-Json -Depth 10
    $formCallsJson | Out-File -FilePath (Join-Path $OutputDir "form_calls.json") -Encoding utf8
    Write-Host "  [OK] form_calls.json ($(@($allFormCalls).Count) calls)" -ForegroundColor Green

    Write-Host ""
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host " Analysis Complete" -ForegroundColor Cyan
    Write-Host "  Cell References: $(@($allCellRefs).Count)" -ForegroundColor White
    Write-Host "  Tables:          $(@($tableOutput).Count)" -ForegroundColor White
    Write-Host "  Events:          $(@($allEvents).Count)" -ForegroundColor White
    Write-Host "  Form Calls:      $(@($allFormCalls).Count)" -ForegroundColor White
    Write-Host "================================================================" -ForegroundColor Cyan

    return @{
        CellReferences = $allCellRefs
        TableBindings  = $allTableBindings
        ColumnAccesses = $allColumnAccesses
        Events         = $allEvents
        FormCalls      = $allFormCalls
        Procedures     = $allProcedures
    }
}

# ============================================================
# SECTION 8: スタンドアロン実行エントリーポイント
# ============================================================

if ($InputDir) {
    if (-not (Test-Path $InputDir)) {
        Write-Error "InputDir not found: $InputDir"
        exit 1
    }
    if (-not (Test-Path $OutputDir)) {
        New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
    }

    $result = Invoke-VBAAnalysis `
        -InputDir $InputDir `
        -OutputDir $OutputDir `
        -SheetCodeNameMap $SheetCodeNameMap

    Write-Host ""
    Write-Host "Output directory: $OutputDir" -ForegroundColor Yellow
}
