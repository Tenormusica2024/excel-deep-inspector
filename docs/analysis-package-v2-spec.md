# Analysis Package v2 仕様書
**作成日**: 2026-02-11
**目的**: 社内AI（Azure ChatGPTラッパー）がVBA + 数式 + UIの統合分析を実行するための、構造化データ出力仕様

---

## 前提条件・制約

- **実行環境**: 社内PC（PowerShell 5.1のみ、Python不可）
- **AI送信先**: 社内AI（Azure ChatGPTラッパー）- テキストのみ、画像入力不可
- **ファイルサイズ**: 1シート1ファイル、<10MB
- **開発**: 自宅PC（Claude Code MAX）で仕様策定 → GitHub経由 → 社内PC（Codex CLI）で実装
- **コード転送**: GitHubリポジトリ経由のテキストベースのみ（ファイル直接送付不可）

---

## 出力フォルダ構造（v2）

```
{ワークブック名}_analysis_package/
  00_overview.md                    # v1と同様 + v2追加情報
  01_VBA/
    {モジュール名}.bas              # v1と同様（VBAソースコード）
    procedures.json                 # v1と同様（プロシージャ一覧）
    cell_references.json            # ★v2: 大幅改善（後述）
    table_references.json           # ★v2: 新規（ListObject/ListColumn参照）
    event_triggers.json             # ★v2: 新規（ワークシートイベント + ActiveXイベント）
  02_formulas/
    {シート名}_formulas.json        # v1と同様
    named_ranges.json               # v1と同様
  03_screenshots/
    {シート名}.png                  # v1と同様
  04_structure/
    controls.json                   # v1と同様
    sheet_codenames.json            # ★v2: 新規（コードネーム→表示名マッピング）
    activex_event_map.json          # ★v2: 新規（ActiveXコントロール→VBAイベントハンドラ）
  05_cross_reference/               # ★v2: 新規フォルダ
    ui_to_vba.json                  # UI操作→VBAコード→セル影響の完全マッピング
    data_flow.json                  # テーブル列のデータフロー（読み書き追跡）
```

---

## ★ 新規/改善ファイル仕様

### 1. sheet_codenames.json（新規）

**目的**: VBAコードネーム（例: `shDash`）とExcelの表示シート名（例: `Dashboard`）のマッピング。
VBAコード内の `shDash.Range("A1")` が実際にどのシートを操作しているかをAIが理解するために必須。

**取得方法（PowerShell COM）**:
```powershell
# Workbook.VBProject.VBComponents からCodeNameとSheet.Nameを取得
foreach ($sheet in $Workbook.Sheets) {
    # $sheet.CodeName = VBAコードネーム（例: "shDash"）
    # $sheet.Name = 表示名（例: "Dashboard"）
}
```

**出力フォーマット**:
```json
{
  "Mappings": [
    {
      "CodeName": "shDash",
      "DisplayName": "Dashboard",
      "SheetType": "Worksheet",
      "Index": 1,
      "VBAModuleName": "shDash",
      "HasCode": false
    },
    {
      "CodeName": "shInvoice",
      "DisplayName": "Invoice",
      "SheetType": "Worksheet",
      "Index": 2,
      "VBAModuleName": "shInvoice",
      "HasCode": true
    },
    {
      "CodeName": "shTemp",
      "DisplayName": "Template",
      "SheetType": "Worksheet",
      "Index": 4,
      "VBAModuleName": "shTemp",
      "HasCode": false
    },
    {
      "CodeName": "shMaster",
      "DisplayName": "Master",
      "SheetType": "Worksheet",
      "Index": 5,
      "VBAModuleName": "shMaster",
      "HasCode": false
    },
    {
      "CodeName": "shCalc",
      "DisplayName": "Calculation",
      "SheetType": "Worksheet",
      "Index": 6,
      "VBAModuleName": "shCalc",
      "HasCode": false
    }
  ],
  "TotalSheets": 7
}
```

**AIへの効果**: `shDash.Range("PDFfolder")` → 「DashboardシートのPDFfolderセルを参照」と即座に解釈可能。

---

### 2. cell_references.json（v2改善版）

**v1の問題点**:
- Range_LiteralとNamed_Rangeで同じ参照を二重カウント（16件中8件が重複）
- シートコンテキスト（`shDash.`）が未捕捉
- プロシージャ名が未記録
- Cells()パターン未検出

**v2の改善点**:
1. パターン分類の明確化（重複排除）
2. シートコンテキストの捕捉（CodeName + 解決後のDisplayName）
3. プロシージャ名の記録
4. Cells()パターンの追加
5. 確信度スコアの付与
6. 名前付き範囲の実アドレス解決

**検出パターン一覧（優先順位順）**:

| パターンID | 正規表現（概要） | 確信度 | 例 |
|-----------|----------------|--------|-----|
| `SHEET_RANGE` | `{sheetVar}\.Range\("..."\)` | 0.95 | `shDash.Range("PDFfolder")` |
| `SHEET_CELLS` | `{sheetVar}\.Cells\(..,..\)` | 0.90 | `shTemp.Cells(1, 1)` |
| `WORKSHEETS_RANGE` | `Worksheets\(".."\)\.Range\(".."\)` | 0.95 | `Worksheets("Sheet1").Range("A1")` |
| `ME_RANGE` | `Me\.Range\(".."\)` | 0.90 | `Me.Range("A1")` |
| `RANGE_LITERAL` | `Range\("([A-Z]+\d+..)"\)` | 0.85 | `Range("A1")` |
| `RANGE_NAMED` | `Range\("([^"]+)"\)` ※非セルアドレス | 0.80 | `Range("PDFfolder")` |
| `CELLS_LITERAL` | `Cells\(\d+,\s*\d+\)` | 0.85 | `Cells(1, 3)` |
| `CELLS_VARIABLE` | `Cells\(\w+,\s*\w+\)` | 0.50 | `Cells(i, 3)` |
| `ACTIVESHEET_RANGE` | `ActiveSheet\.Range\(".."\)` | 0.40 | `ActiveSheet.Range("A1")` |

**重複排除ルール**:
- `Range("A1")` → セルアドレスパターン(`[A-Z]+\d+`)にマッチ → `RANGE_LITERAL`のみ
- `Range("PDFfolder")` → セルアドレスパターンにマッチしない → `RANGE_NAMED`のみ
- `shDash.Range("X")` → `SHEET_RANGE`を優先（下位パターンでは重複登録しない）

**出力フォーマット**:
```json
{
  "Version": "2.0",
  "TotalReferences": 12,
  "References": [
    {
      "Id": "ref_001",
      "Pattern": "SHEET_RANGE",
      "Module": "FormOpenInvoice",
      "Procedure": "btCreateInvoice_Click",
      "ProcedureType": "Sub",
      "Line": 83,
      "SheetCodeName": "shDash",
      "SheetDisplayName": "Dashboard",
      "RawAddress": "PDFfolder",
      "ResolvedAddress": "Dashboard!$M$19",
      "IsNamedRange": true,
      "Confidence": 0.95,
      "AccessType": "Read",
      "Context": "PDFFolderPath = ThisWorkbook.Path & \"\\\" & shDash.Range(\"PDFfolder\").Value"
    },
    {
      "Id": "ref_002",
      "Pattern": "SHEET_RANGE",
      "Module": "FormOpenInvoice",
      "Procedure": "btCreateInvoice_Click",
      "ProcedureType": "Sub",
      "Line": 84,
      "SheetCodeName": "shDash",
      "SheetDisplayName": "Dashboard",
      "RawAddress": "Excelfolder",
      "ResolvedAddress": "Dashboard!$M$20",
      "IsNamedRange": true,
      "Confidence": 0.95,
      "AccessType": "Read",
      "Context": "ExcelFolderPath = ThisWorkbook.Path & \"\\\" & shDash.Range(\"Excelfolder\").Value"
    },
    {
      "Id": "ref_003",
      "Pattern": "SHEET_RANGE",
      "Module": "FormOpenInvoice",
      "Procedure": "btCreateInvoice_Click",
      "ProcedureType": "Sub",
      "Line": 116,
      "SheetCodeName": "shDash",
      "SheetDisplayName": "Dashboard",
      "RawAddress": "TODAY",
      "ResolvedAddress": "Dashboard!$C$1",
      "IsNamedRange": true,
      "Confidence": 0.95,
      "AccessType": "Read",
      "Context": ".ListColumns(\"Invoice Date\").DataBodyRange(InvRowId).Value = shDash.Range(\"TODAY\").Value"
    },
    {
      "Id": "ref_004",
      "Pattern": "SHEET_RANGE",
      "Module": "FormOpenInvoice",
      "Procedure": "btCreateInvoice_Click",
      "ProcedureType": "Sub",
      "Line": 130,
      "SheetCodeName": "shTemp",
      "SheetDisplayName": "Template",
      "RawAddress": "B12",
      "ResolvedAddress": "Template!$B$12",
      "IsNamedRange": false,
      "Confidence": 0.95,
      "AccessType": "Read",
      "Context": "InvFileName = InvNum & \"_\" & shTemp.Range(\"B12\").Value"
    },
    {
      "Id": "ref_005",
      "Pattern": "RANGE_LITERAL",
      "Module": "MainInvoiceGenerator",
      "Procedure": "view_invoices",
      "ProcedureType": "Sub",
      "Line": 33,
      "SheetCodeName": null,
      "SheetDisplayName": null,
      "RawAddress": "A1",
      "ResolvedAddress": null,
      "IsNamedRange": false,
      "Confidence": 0.85,
      "AccessType": "Select",
      "Context": "Range(\"A1\").Select",
      "Note": "先行行の shInvoice.Select から Invoice シートと推定可能"
    }
  ],
  "PatternSummary": {
    "SHEET_RANGE": 8,
    "RANGE_LITERAL": 3,
    "RANGE_NAMED": 0,
    "CELLS_VARIABLE": 0,
    "CELLS_LITERAL": 1
  }
}
```

**AccessType判定ルール**:
- `.Value = expr` (左辺) → `"Write"`
- `= ...Range("X").Value` (右辺) → `"Read"`
- `.Select` → `"Select"`
- `.Clear` / `.Delete` → `"Delete"`
- 判定不能 → `"Unknown"`

---

### 3. table_references.json（新規）

**目的**: VBAコード内のListObject（構造化テーブル）とListColumn参照を検出。
Invoice_Generatorでは `InvoiceTable`, `CustomerTable` がデータの中核であり、
`ListColumns("Status").DataBodyRange(r).Value` のような参照パターンが頻出する。

**検出パターン**:

| パターン | 正規表現（概要） | 例 |
|---------|----------------|-----|
| `LISTOBJECT_BIND` | `ListObjects\("(\w+)"\)` | `shInvoice.ListObjects("InvoiceTable")` |
| `LISTCOLUMN_ACCESS` | `ListColumns\("(.+?)"\)\.DataBodyRange` | `InvoiceTable.ListColumns("Status").DataBodyRange(r)` |
| `LISTROWS_COUNT` | `\.ListRows\.Count` | `InvoiceTable.ListRows.Count` |
| `LISTCOLUMN_LARGE` | `WorksheetFunction\.Large\(.*ListColumns` | `Large(.ListColumns("Invoice Number").DataBodyRange, 1)` |

**出力フォーマット**:
```json
{
  "Tables": [
    {
      "TableName": "InvoiceTable",
      "SheetCodeName": "shInvoice",
      "SheetDisplayName": "Invoice",
      "BoundIn": [
        {
          "Module": "FormOpenInvoice",
          "Procedure": "UserForm_Initialize",
          "Line": 32,
          "Variable": "InvoiceTable",
          "Context": "Set InvoiceTable = shInvoice.ListObjects(\"InvoiceTable\")"
        },
        {
          "Module": "FormOpenInvoice",
          "Procedure": "btCreateInvoice_Click",
          "Line": 78,
          "Variable": "InvoiceTable",
          "Context": "Set InvoiceTable = shInvoice.ListObjects(\"InvoiceTable\")"
        },
        {
          "Module": "shInvoice",
          "Procedure": "axHide_Click",
          "Line": 33,
          "Variable": "InvoiceTable",
          "Context": "Set InvoiceTable = Me.ListObjects(\"InvoiceTable\")"
        },
        {
          "Module": "shInvoice",
          "Procedure": "Worksheet_Change",
          "Line": 65,
          "Variable": "InvoiceTable",
          "Context": "Set InvoiceTable = Me.ListObjects(\"InvoiceTable\")"
        }
      ],
      "ColumnsAccessed": [
        {
          "ColumnName": "Status",
          "AccessType": "Read",
          "Procedures": ["UserForm_Initialize", "axHide_Click", "btCreateInvoice_Click"]
        },
        {
          "ColumnName": "Company",
          "AccessType": "Read",
          "Procedures": ["UserForm_Initialize", "btCreateEmail_Click"]
        },
        {
          "ColumnName": "Final Invoiced Amount",
          "AccessType": "Read",
          "Procedures": ["UserForm_Initialize"]
        },
        {
          "ColumnName": "Invoice Number",
          "AccessType": "ReadWrite",
          "Procedures": ["btCreateInvoice_Click"]
        },
        {
          "ColumnName": "Invoice Date",
          "AccessType": "Write",
          "Procedures": ["btCreateInvoice_Click"]
        },
        {
          "ColumnName": "Agreed Amount Total",
          "AccessType": "Read",
          "Procedures": ["btCreateInvoice_Click"]
        },
        {
          "ColumnName": "Customer",
          "AccessType": "Read",
          "Procedures": ["btCreateEmail_Click"]
        }
      ]
    },
    {
      "TableName": "CustomerTable",
      "SheetCodeName": "shMaster",
      "SheetDisplayName": "Master",
      "BoundIn": [
        {
          "Module": "FormOpenInvoice",
          "Procedure": "btCreateEmail_Click",
          "Line": 190,
          "Variable": "CustomerTable",
          "Context": "Set CustomerTable = shMaster.ListObjects(\"CustomerTable\")"
        }
      ],
      "ColumnsAccessed": [
        {
          "ColumnName": "Customer",
          "AccessType": "Read",
          "Procedures": ["btCreateEmail_Click"]
        },
        {
          "ColumnName": "Company",
          "AccessType": "Read",
          "Procedures": ["btCreateEmail_Click"]
        },
        {
          "ColumnName": "Email",
          "AccessType": "Read",
          "Procedures": ["btCreateEmail_Click"]
        }
      ]
    }
  ],
  "TotalTables": 2,
  "TotalColumnsAccessed": 10
}
```

**AIへの効果**: テーブル中心のデータフローが一目瞭然。「InvoiceTableのStatus列は3つのプロシージャで読み取られている」等の分析が可能。

---

### 4. event_triggers.json（新規）

**目的**: ワークシートイベント（Worksheet_Change, Worksheet_Activate等）とActiveXコントロールイベント（axCollapse_Click等）を検出。
これらはUIから「不可視」なトリガーであり、ユーザーが気づかないまま重要な処理を実行している場合が多い。

**検出パターン**:

| イベントタイプ | パターン | 例 |
|-------------|---------|-----|
| `WORKSHEET_EVENT` | `Private Sub Worksheet_{EventName}` | `Worksheet_Change(ByVal Target As Range)` |
| `WORKBOOK_EVENT` | `Private Sub Workbook_{EventName}` | `Workbook_Open()` |
| `ACTIVEX_EVENT` | `Private Sub {ControlName}_{EventName}` ※Document型モジュール内 | `axCollapse_Click()` |
| `USERFORM_EVENT` | `Private Sub {ControlName}_{EventName}` ※UserForm型モジュール内 | `btCreateInvoice_Click()` |

**ActiveXイベントの検出ロジック**:
1. Document型モジュール（shXxx.bas）内の`Private Sub xxx_Click/Change/...`を検出
2. コントロール名部分を controls.json のActiveXコントロール（Type=12）と照合
3. マッチすればActiveXイベントハンドラとして登録

**出力フォーマット**:
```json
{
  "Events": [
    {
      "Type": "ACTIVEX_EVENT",
      "Module": "shInvoice",
      "SheetDisplayName": "Invoice",
      "ControlName": "axCollapse",
      "ControlType": "ActiveX Checkbox",
      "EventName": "Click",
      "ProcedureName": "axCollapse_Click",
      "Line": 9,
      "Description": "Hide/unhide columns that show low-level details",
      "CellsAffected": [],
      "TablesAffected": [],
      "UIEffect": "Outline.ShowLevels ColumnLevels変更（列の表示/非表示）"
    },
    {
      "Type": "ACTIVEX_EVENT",
      "Module": "shInvoice",
      "SheetDisplayName": "Invoice",
      "ControlName": "axHide",
      "ControlType": "ActiveX Checkbox",
      "EventName": "Click",
      "ProcedureName": "axHide_Click",
      "Line": 23,
      "Description": "Hide/unhide the rows where invoices are already paid",
      "CellsAffected": [],
      "TablesAffected": ["InvoiceTable"],
      "UIEffect": "Status='Paid'の行を非表示/表示切替"
    },
    {
      "Type": "WORKSHEET_EVENT",
      "Module": "shInvoice",
      "SheetDisplayName": "Invoice",
      "ControlName": null,
      "ControlType": null,
      "EventName": "Change",
      "ProcedureName": "Worksheet_Change",
      "Line": 57,
      "Description": "Orange cellに値入力 → FormSearch起動",
      "TriggerCondition": "InvoiceTable最終行の1列目に値入力（Target範囲の判定）",
      "CellsAffected": ["InvoiceTable最終行1列目"],
      "TablesAffected": ["InvoiceTable"],
      "FormsTriggered": ["FormSearch"]
    }
  ],
  "TotalEvents": 3,
  "Summary": {
    "ACTIVEX_EVENT": 2,
    "WORKSHEET_EVENT": 1,
    "WORKBOOK_EVENT": 0,
    "USERFORM_EVENT": 0
  }
}
```

---

### 5. activex_event_map.json（新規）

**目的**: controls.json でOnAction=""となっているActiveXコントロールと、実際のVBAイベントハンドラの紐付け。
フォームコントロール（OnActionプロパティで接続）とは異なり、ActiveXコントロールはシートモジュール内のイベントプロシージャで接続されるため、別途マッピングが必要。

**生成ロジック**:
1. controls.json から Type=12（ActiveX）のコントロールを抽出
2. そのコントロールが配置されたシートのCodeNameを特定
3. 対応するDocument型VBAモジュール内で `Private Sub {ControlName}_{EventName}` を検索
4. マッチした結果をマッピング

**出力フォーマット**:
```json
{
  "Mappings": [
    {
      "ControlName": "axCollapse",
      "ControlType": 12,
      "Sheet": "Invoice",
      "SheetCodeName": "shInvoice",
      "Position": { "Left": 21.5, "Top": 158.5, "Width": 105.5, "Height": 30 },
      "EventHandlers": [
        {
          "EventName": "Click",
          "ProcedureName": "axCollapse_Click",
          "Module": "shInvoice",
          "Line": 9
        }
      ]
    },
    {
      "ControlName": "axHide",
      "ControlType": 12,
      "Sheet": "Invoice",
      "SheetCodeName": "shInvoice",
      "Position": { "Left": 153, "Top": 159, "Width": 107, "Height": 30 },
      "EventHandlers": [
        {
          "EventName": "Click",
          "ProcedureName": "axHide_Click",
          "Module": "shInvoice",
          "Line": 23
        }
      ]
    }
  ],
  "UnmappedControls": [],
  "TotalMapped": 2,
  "TotalUnmapped": 0
}
```

---

### 6. ui_to_vba.json（新規 - クロスリファレンス）

**目的**: UI操作（ボタンクリック、セル変更、チェックボックス切替等）から始まる処理チェーンの完全マッピング。
「このボタンを押したら何が起きるか」をAIが即座に理解するための接着剤データ。

**生成ロジック**:
1. controls.json + activex_event_map.json からUI起点を収集
2. event_triggers.json からワークシートイベント起点を収集
3. 各起点からVBAプロシージャを特定
4. cell_references.json + table_references.json からセル/テーブル影響を集約
5. さらにFormの呼び出し（FormXxx.Show）を追跡してチェーンを延長

**出力フォーマット**:
```json
{
  "Chains": [
    {
      "Id": "chain_001",
      "TriggerType": "FormControl_OnAction",
      "TriggerSource": {
        "ControlName": "Graphic 6",
        "Sheet": "Invoice",
        "ControlDescription": "Type=28 (Graphic), ダッシュボードへ戻るボタン"
      },
      "EntryPoint": {
        "OnAction": "Invoice_Generator.xlsm!return_dashboard",
        "Module": "MainInvoiceGenerator",
        "Procedure": "return_dashboard",
        "Line": 37
      },
      "CellsAffected": [
        {
          "Address": "A1",
          "Sheet": null,
          "AccessType": "Select",
          "Note": "shDash.Select後のRange(\"A1\").Select"
        }
      ],
      "TablesAffected": [],
      "SheetsNavigated": ["Dashboard"],
      "FormsTriggered": [],
      "Impact": "Invoiceシート → Dashboardシートへのナビゲーション"
    },
    {
      "Id": "chain_002",
      "TriggerType": "FormControl_OnAction",
      "TriggerSource": {
        "ControlName": "Dashboard上のボタン（OnAction=create_invoice）",
        "Sheet": "Dashboard",
        "ControlDescription": "請求書作成ボタン"
      },
      "EntryPoint": {
        "OnAction": "create_invoice",
        "Module": "MainInvoiceGenerator",
        "Procedure": "create_invoice",
        "Line": 23
      },
      "CellsAffected": [],
      "TablesAffected": [],
      "SheetsNavigated": [],
      "FormsTriggered": ["FormOpenInvoice"],
      "SubChains": [
        {
          "Form": "FormOpenInvoice",
          "Events": [
            {
              "Event": "UserForm_Initialize",
              "TablesRead": ["InvoiceTable"],
              "ColumnsRead": ["Status", "Company", "Final Invoiced Amount"]
            },
            {
              "Event": "btCreateInvoice_Click",
              "TablesWritten": ["InvoiceTable"],
              "ColumnsWritten": ["Invoice Number", "Invoice Date", "Status"],
              "CellsRead": [
                "Dashboard!PDFfolder",
                "Dashboard!Excelfolder",
                "Dashboard!TODAY",
                "Template!B12"
              ],
              "ExternalEffects": [
                "PDF出力: {PDFFolderPath}\\{InvFileName}.pdf",
                "Excel出力: {ExcelFolderPath}\\{InvFileName}.xlsx"
              ]
            },
            {
              "Event": "btCreateEmail_Click",
              "TablesRead": ["CustomerTable", "InvoiceTable"],
              "ColumnsRead": ["Customer", "Company", "Email"],
              "ExternalEffects": [
                "Outlook下書きメール作成（PDF添付）"
              ]
            }
          ]
        }
      ],
      "Impact": "請求書の作成 → PDF/Excel出力 → InvoiceTableステータス更新 → メール作成"
    },
    {
      "Id": "chain_003",
      "TriggerType": "ActiveX_Event",
      "TriggerSource": {
        "ControlName": "axCollapse",
        "Sheet": "Invoice",
        "ControlDescription": "ActiveX Checkbox - 列の折りたたみ"
      },
      "EntryPoint": {
        "Module": "shInvoice",
        "Procedure": "axCollapse_Click",
        "Line": 9
      },
      "CellsAffected": [],
      "TablesAffected": [],
      "SheetsNavigated": [],
      "FormsTriggered": [],
      "Impact": "Invoice列の詳細表示/折りたたみ切替（Outline.ShowLevels）"
    },
    {
      "Id": "chain_004",
      "TriggerType": "ActiveX_Event",
      "TriggerSource": {
        "ControlName": "axHide",
        "Sheet": "Invoice",
        "ControlDescription": "ActiveX Checkbox - 支払済み行の非表示"
      },
      "EntryPoint": {
        "Module": "shInvoice",
        "Procedure": "axHide_Click",
        "Line": 23
      },
      "CellsAffected": [],
      "TablesAffected": ["InvoiceTable"],
      "SheetsNavigated": [],
      "FormsTriggered": [],
      "Impact": "Status=Paidの行を非表示/表示切替"
    },
    {
      "Id": "chain_005",
      "TriggerType": "Worksheet_Event",
      "TriggerSource": {
        "EventName": "Worksheet_Change",
        "Sheet": "Invoice",
        "TriggerCondition": "InvoiceTable最終行の1列目（オレンジセル）に値入力"
      },
      "EntryPoint": {
        "Module": "shInvoice",
        "Procedure": "Worksheet_Change",
        "Line": 57
      },
      "CellsAffected": [
        {
          "Address": "InvoiceTable最終行1列目",
          "AccessType": "ReadWrite",
          "Note": "入力値を読み取り → FormSearch起動 → 元に戻す"
        }
      ],
      "TablesAffected": ["InvoiceTable"],
      "FormsTriggered": ["FormSearch"],
      "Impact": "顧客名入力 → 検索フォーム表示 → 検索結果選択"
    }
  ],
  "TotalChains": 5,
  "CoverageNote": "Dashboard上の一部ボタン（add_data_to_master, edit_view_master_data, view_invoices）のOnActionが controls.json で空のため、chain未生成。DashboardのGroup型コントロールのOnAction取得改善が必要。"
}
```

---

### 7. data_flow.json（新規）

**目的**: テーブル列単位でのデータフロー追跡。「この列に誰が書き込み、誰が読み取るか」を示す。

**出力フォーマット**:
```json
{
  "Flows": [
    {
      "Table": "InvoiceTable",
      "Sheet": "Invoice",
      "Column": "Status",
      "Writers": [
        {
          "Module": "FormOpenInvoice",
          "Procedure": "btCreateInvoice_Click",
          "Line": 153,
          "Value": "\"Created\""
        }
      ],
      "Readers": [
        {
          "Module": "FormOpenInvoice",
          "Procedure": "UserForm_Initialize",
          "Line": 40,
          "Condition": "= \"\" でオープン請求書判定"
        },
        {
          "Module": "shInvoice",
          "Procedure": "axHide_Click",
          "Line": 40,
          "Condition": "= \"Paid\" で支払済み行判定"
        }
      ],
      "FormulaReaders": [],
      "Note": "ステータスのライフサイクル: '' → 'Created' → 'Paid'（Paid設定箇所は未検出 - 手動入力の可能性）"
    },
    {
      "Table": "InvoiceTable",
      "Sheet": "Invoice",
      "Column": "Invoice Number",
      "Writers": [
        {
          "Module": "FormOpenInvoice",
          "Procedure": "btCreateInvoice_Click",
          "Line": 117,
          "Value": "MAX(既存番号) + 1 による自動採番"
        }
      ],
      "Readers": [
        {
          "Module": "FormOpenInvoice",
          "Procedure": "btCreateInvoice_Click",
          "Line": 116,
          "Purpose": "次の請求書番号算出のためにLARGE関数で最大値取得"
        }
      ],
      "FormulaReaders": [],
      "Note": "自動採番（連番）- VBAのWorksheetFunction.Large()で最大値+1を算出"
    }
  ],
  "TotalFlows": 10
}
```

---

## PowerShell実装方針

### パーサーの設計指針

すべてのパーサーはPowerShell 5.1で実装し、以下の原則に従う:

1. **正規表現ベースの静的解析**（ASTパーサーは不使用 - PS5.1制約）
2. **プロシージャ境界追跡**: 各行がどのSub/Function内にあるかを追跡する状態マシン
3. **シートコンテキスト解決**: `shXxx.` プレフィックスを sheet_codenames.json で解決
4. **2パス処理**:
   - Pass 1: 個別ファイル解析（cell_references, table_references, event_triggers）
   - Pass 2: クロスリファレンス生成（ui_to_vba, data_flow） ← Pass 1の出力を結合

### プロシージャ境界追跡の擬似コード

```
state = "OUTSIDE"
currentModule = ""
currentProcedure = ""
currentProcedureType = ""

for each line in VBAソース:
  if line matches "^(Public |Private )?(Sub|Function) (\w+)":
    state = "INSIDE"
    currentProcedure = $matches[3]
    currentProcedureType = $matches[2]

  if state == "INSIDE":
    # この行でのパターンマッチング実行
    # currentModule, currentProcedure をコンテキストとして記録

  if line matches "^End (Sub|Function)":
    state = "OUTSIDE"
    currentProcedure = ""
```

### シートコンテキスト解決の擬似コード

```
# sheet_codenames.json を事前にロード
$codeNameMap = @{ "shDash" = "Dashboard"; "shInvoice" = "Invoice"; ... }

# 行内の shXxx. プレフィックスを検出
if line matches "(\w+)\.Range\(" or "(\w+)\.Cells\(" or "(\w+)\.ListObjects\(":
  $prefix = $matches[1]
  if $codeNameMap.ContainsKey($prefix):
    $sheetDisplayName = $codeNameMap[$prefix]
  elif $prefix == "Me":
    $sheetDisplayName = $codeNameMap[$currentModule]  # Document型モジュール名から解決
  elif $prefix == "ActiveSheet":
    $sheetDisplayName = null  # 実行時に決まるため静的解析では未確定
```

---

## AI送信時のパッケージング戦略

### 推奨送信順序（社内AIへの投入順）

1. **00_overview.md** + **sheet_codenames.json** → 全体構造の把握
2. **ui_to_vba.json** → 操作フローの理解（最も価値が高い）
3. **01_VBA/*.bas** → ソースコード（必要に応じて）
4. **cell_references.json** + **table_references.json** → 詳細な参照マッピング
5. **data_flow.json** → データライフサイクルの理解
6. **02_formulas/*.json** → 数式依存関係
7. **event_triggers.json** + **activex_event_map.json** → 不可視トリガー

### ファイルサイズ見積（Invoice_Generator相当の規模）

| ファイル | 推定サイズ | 備考 |
|---------|-----------|------|
| sheet_codenames.json | ~0.5KB | シート数に比例 |
| cell_references.json (v2) | ~3KB | v1の16件→v2で12件（重複除去） |
| table_references.json | ~4KB | テーブル数×列数に比例 |
| event_triggers.json | ~2KB | イベント数に比例 |
| activex_event_map.json | ~1.5KB | ActiveXコントロール数に比例 |
| ui_to_vba.json | ~6KB | チェーン数×深さに比例 |
| data_flow.json | ~4KB | テーブル列数に比例 |
| **v2追加分合計** | **~21KB** | 10MBリミットに対して余裕 |

---

## 実装ロードマップ

### Phase 1: 基盤データ（sheet_codenames + 改善cell_references）
- sheet_codenames.json 生成ロジック
- cell_references.json v2パーサー（重複排除、シート/プロシージャコンテキスト）
- プロシージャ境界追跡エンジン

### Phase 2: テーブル・イベント解析
- table_references.json パーサー
- event_triggers.json パーサー
- activex_event_map.json 生成ロジック（controls.json + VBAモジュール照合）

### Phase 3: クロスリファレンス生成
- ui_to_vba.json 生成ロジック（Phase 1 + Phase 2の出力を統合）
- data_flow.json 生成ロジック
- コールチェーン追跡（Sub A → Form B → Sub C）

### Phase 4: テスト・検証
- Invoice_Generator.xlsm での実行テスト
- 出力JSONの妥当性検証
- 社内AIへの投入テスト（適切に解釈されるか）
