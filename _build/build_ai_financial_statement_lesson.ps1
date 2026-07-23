param(
    [string]$OutputDocx = ""
)

$ErrorActionPreference = "Stop"

$buildDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$teachingDir = Split-Path -Parent $buildDir
if ([string]::IsNullOrWhiteSpace($OutputDocx)) {
    $OutputDocx = Join-Path $teachingDir "AI_Financial_Statements_Analysis_Lesson_Plan.docx"
}
$OutputDocx = [System.IO.Path]::GetFullPath($OutputDocx)

function Color-Value([string]$hex) {
    $hex = $hex.TrimStart('#')
    $r = [Convert]::ToInt32($hex.Substring(0, 2), 16)
    $g = [Convert]::ToInt32($hex.Substring(2, 2), 16)
    $b = [Convert]::ToInt32($hex.Substring(4, 2), 16)
    return $r + (256 * $g) + (65536 * $b)
}

$NAVY = Color-Value "0B2545"
$BLUE = Color-Value "2E74B5"
$DARKBLUE = Color-Value "1F4D78"
$LIGHTBLUE = Color-Value "E8EEF5"
$PALEBLUE = Color-Value "F3F8FC"
$LIGHTGRAY = Color-Value "F2F4F7"
$MIDGRAY = Color-Value "D9DEE6"
$TEXTGRAY = Color-Value "4B5563"
$GREEN = Color-Value "E9F5EC"
$GOLD = Color-Value "FFF4CC"
$RED = Color-Value "FDECEC"
$WHITE = Color-Value "FFFFFF"
$BLACK = Color-Value "000000"

$word = $null
$doc = $null
$wordProcessId = 0
$existingWordProcessIds = @()

try {
    $existingWordProcessIds = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0

    if (-not ("WordWindowProcessFSA" -as [type])) {
        Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class WordWindowProcessFSA {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
"@
    }
    if ($null -ne $word.Hwnd -and [int64]$word.Hwnd -ne 0) {
        [uint32]$pid = 0
        [void][WordWindowProcessFSA]::GetWindowThreadProcessId([IntPtr]$word.Hwnd, [ref]$pid)
        $wordProcessId = [int]$pid
    }
    if ($wordProcessId -eq 0) {
        Start-Sleep -Milliseconds 500
        $newWordProcess = Get-Process -Name WINWORD -ErrorAction SilentlyContinue |
            Where-Object { $existingWordProcessIds -notcontains $_.Id } |
            Sort-Object StartTime -Descending |
            Select-Object -First 1
        if ($null -ne $newWordProcess) { $wordProcessId = $newWordProcess.Id }
    }

    $doc = $word.Documents.Add()
    $doc.PageSetup.PageWidth = $word.InchesToPoints(8.5)
    $doc.PageSetup.PageHeight = $word.InchesToPoints(11)
    $doc.PageSetup.TopMargin = $word.InchesToPoints(1.0)
    $doc.PageSetup.BottomMargin = $word.InchesToPoints(1.0)
    $doc.PageSetup.LeftMargin = $word.InchesToPoints(1.0)
    $doc.PageSetup.RightMargin = $word.InchesToPoints(1.0)
    $doc.PageSetup.HeaderDistance = $word.InchesToPoints(0.49)
    $doc.PageSetup.FooterDistance = $word.InchesToPoints(0.49)

    # Compact reference guide preset, resolved into Word styles.
    $normal = $doc.Styles.Item("Normal")
    $normal.Font.Name = "Calibri"
    $normal.Font.Size = 11
    $normal.Font.Color = $BLACK
    $normal.ParagraphFormat.SpaceBefore = 0
    $normal.ParagraphFormat.SpaceAfter = 6
    $normal.ParagraphFormat.LineSpacingRule = 5
    $normal.ParagraphFormat.LineSpacing = 15

    foreach ($spec in @(
        @{Name="Heading 1"; Size=16; Color=$BLUE; Before=18; After=10},
        @{Name="Heading 2"; Size=13; Color=$BLUE; Before=14; After=7},
        @{Name="Heading 3"; Size=12; Color=$DARKBLUE; Before=10; After=5}
    )) {
        $style = $doc.Styles.Item($spec.Name)
        $style.Font.Name = "Calibri"
        $style.Font.Size = $spec.Size
        $style.Font.Bold = -1
        $style.Font.Color = $spec.Color
        $style.ParagraphFormat.SpaceBefore = $spec.Before
        $style.ParagraphFormat.SpaceAfter = $spec.After
        $style.ParagraphFormat.KeepWithNext = -1
    }
    $doc.Styles.Item("List Bullet").Font.Name = "Calibri"
    $doc.Styles.Item("List Bullet").Font.Size = 11
    $doc.Styles.Item("List Bullet").ParagraphFormat.SpaceAfter = 4
    $doc.Styles.Item("List Number").Font.Name = "Calibri"
    $doc.Styles.Item("List Number").Font.Size = 11
    $doc.Styles.Item("List Number").ParagraphFormat.SpaceAfter = 4
    $script:PageBreakBeforeNext = $false

    function Add-Paragraph {
        param(
            [string]$Text,
            [string]$Style = "Normal",
            [bool]$Bold = $false,
            [bool]$Italic = $false,
            [double]$Size = 0,
            [int]$Color = -1,
            [int]$After = -1,
            [int]$Alignment = -1
        )
        $p = $doc.Paragraphs.Add()
        $p.Range.InsertBefore($Text)
        $p.Range.Style = $doc.Styles.Item($Style)
        $p.Format.LeftIndent = 0
        $p.Format.RightIndent = 0
        $p.Format.FirstLineIndent = 0
        $p.Alignment = 0
        $p.Format.PageBreakBefore = 0
        if ($Bold) { $p.Range.Font.Bold = -1 }
        if ($Italic) { $p.Range.Font.Italic = -1 }
        if ($Size -gt 0) { $p.Range.Font.Size = $Size }
        if ($Color -ge 0) { $p.Range.Font.Color = $Color }
        if ($After -ge 0) { $p.Format.SpaceAfter = $After }
        if ($Alignment -ge 0) { $p.Alignment = $Alignment }
        if ($script:PageBreakBeforeNext) {
            $p.Format.PageBreakBefore = -1
            $script:PageBreakBeforeNext = $false
        }
        return $p
    }

    function Add-LabeledParagraph {
        param([string]$Label, [string]$Text, [int]$Fill = -1)
        $p = $doc.Paragraphs.Add()
        $p.Range.InsertBefore("$Label $Text")
        $r = $p.Range
        $r.Style = $doc.Styles.Item("Normal")
        $labelRange = $doc.Range($r.Start, $r.Start + $Label.Length)
        $labelRange.Font.Bold = -1
        $labelRange.Font.Color = $DARKBLUE
        if ($Fill -ge 0) {
            $r.Shading.BackgroundPatternColor = $Fill
            $p.Format.LeftIndent = 6
            $p.Format.RightIndent = 6
            $p.Format.SpaceBefore = 4
            $p.Format.SpaceAfter = 8
        }
        return $p
    }

    function Add-Bullets {
        param([string[]]$Items)
        $start = $doc.Content.End - 1
        foreach ($item in $Items) {
            [void](Add-Paragraph -Text $item -Style "Normal")
        }
        $end = $doc.Content.End - 1
        $range = $doc.Range($start, $end)
        $template = $word.ListGalleries.Item(1).ListTemplates.Item(1)
        $range.ListFormat.ApplyListTemplateWithLevel($template, $false, 0, 0, 1)
        foreach ($p in $range.Paragraphs) {
            $p.Format.SpaceAfter = 4
            $p.Format.LineSpacingRule = 5
            $p.Format.LineSpacing = 15
        }
    }

    function Add-Numbers {
        param([string[]]$Items)
        $start = $doc.Content.End - 1
        foreach ($item in $Items) {
            [void](Add-Paragraph -Text $item -Style "Normal")
        }
        $end = $doc.Content.End - 1
        $range = $doc.Range($start, $end)
        $template = $word.ListGalleries.Item(2).ListTemplates.Item(1)
        $range.ListFormat.ApplyListTemplateWithLevel($template, $false, 0, 0, 1)
        foreach ($p in $range.Paragraphs) {
            $p.Format.SpaceAfter = 4
            $p.Format.LineSpacingRule = 5
            $p.Format.LineSpacing = 15
        }
    }

    function Add-CodeBlock {
        param([string]$Text)
        $p = $doc.Paragraphs.Add()
        $blockText = $Text -replace "`r`n", "`v"
        $blockText = $blockText -replace "`n", "`v"
        $p.Range.InsertBefore($blockText)
        $p.Range.Style = $doc.Styles.Item("Normal")
        $p.Range.Font.Name = "Consolas"
        $p.Range.Font.Size = 9
        $p.Range.Font.Color = $NAVY
        $p.Range.Shading.BackgroundPatternColor = $LIGHTGRAY
        $p.Format.LeftIndent = 8
        $p.Format.RightIndent = 8
        $p.Format.SpaceBefore = 4
        $p.Format.SpaceAfter = 8
        $p.Format.LineSpacingRule = 0
        return $p
    }

    function Add-Table {
        param(
            [string[]]$Headers,
            [object[]]$Rows,
            [double[]]$Widths,
            [int]$HeaderFill = -1,
            [double]$FontSize = 9.5
        )
        $range = $doc.Range($doc.Content.End - 1, $doc.Content.End - 1)
        $table = $doc.Tables.Add($range, $Rows.Count + 1, $Headers.Count)
        $table.AllowAutoFit = $false
        $table.PreferredWidthType = 3
        $table.PreferredWidth = 468
        $table.LeftPadding = 6
        $table.RightPadding = 6
        $table.TopPadding = 4
        $table.BottomPadding = 4
        $table.Rows.AllowBreakAcrossPages = 0
        $table.Rows.Item(1).HeadingFormat = -1
        $table.Borders.Enable = 1
        $table.Borders.OutsideColor = $MIDGRAY
        $table.Borders.InsideColor = $MIDGRAY
        if ($HeaderFill -lt 0) { $HeaderFill = $LIGHTBLUE }
        for ($c = 1; $c -le $Headers.Count; $c++) {
            $table.Columns.Item($c).Width = $word.InchesToPoints($Widths[$c - 1])
            $cell = $table.Cell(1, $c)
            $cell.Range.Text = $Headers[$c - 1]
            $cell.Range.Style = $doc.Styles.Item("Normal")
            $cell.Range.Font.Name = "Calibri"
            $cell.Range.Font.Size = $FontSize
            $cell.Range.Font.Bold = -1
            $cell.Range.Font.Color = $NAVY
            $cell.Shading.BackgroundPatternColor = $HeaderFill
            $cell.VerticalAlignment = 1
        }
        for ($r = 0; $r -lt $Rows.Count; $r++) {
            $rowData = $Rows[$r]
            if ($rowData.Count -eq 1 -and $rowData[0] -is [System.Array]) {
                $rowData = $rowData[0]
            }
            for ($c = 0; $c -lt $Headers.Count; $c++) {
                $cell = $table.Cell($r + 2, $c + 1)
                $cell.Range.Text = [string]$rowData[$c]
                $cell.Range.Style = $doc.Styles.Item("Normal")
                $cell.Range.Font.Name = "Calibri"
                $cell.Range.Font.Size = $FontSize
                $cell.Range.Font.Color = $BLACK
                $cell.Range.Font.Bold = 0
                if (($r % 2) -eq 1) { $cell.Shading.BackgroundPatternColor = $LIGHTGRAY }
                $cell.VerticalAlignment = 1
            }
        }
        $after = $doc.Paragraphs.Add()
        $after.Format.SpaceAfter = 4
        return $table
    }

    function Add-PageBreak {
        $script:PageBreakBeforeNext = $true
    }

    # Running furniture.
    $header = $doc.Sections.Item(1).Headers.Item(1).Range
    $header.Text = "FINC 310 | AI Financial Statement Analysis"
    $header.Font.Name = "Calibri"
    $header.Font.Size = 8.5
    $header.Font.Color = $TEXTGRAY
    $header.ParagraphFormat.Alignment = 0

    $footer = $doc.Sections.Item(1).Footers.Item(1).Range
    $footer.Text = "Instructor lesson plan  |  "
    $footer.Font.Name = "Calibri"
    $footer.Font.Size = 8.5
    $footer.Font.Color = $TEXTGRAY
    $footer.ParagraphFormat.Alignment = 2
    [void]$footer.Collapse(0)
    [void]$doc.Sections.Item(1).Footers.Item(1).Range.Fields.Add($doc.Sections.Item(1).Footers.Item(1).Range, 33)

    # PAGE 1: Workshop-agenda opening.
    [void](Add-Paragraph -Text "INSTRUCTOR PLAYBOOK" -Bold $true -Size 10 -Color $BLUE -After 2)
    [void](Add-Paragraph -Text "AI Financial Statements and Financial Statement Analysis" -Bold $true -Size 25 -Color $NAVY -After 6)
    [void](Add-Paragraph -Text "A three-lab lesson built around prediction, AI-assisted analysis, independent verification, explanation, and decision-making" -Size 13 -Color $TEXTGRAY -After 14)

    [void](Add-Table -Headers @("120 MINUTES", "3 AI LABS", "ONE FINANCIAL STORY") -Rows @(
        ,@("Core classroom format", "Mechanics -> evidence -> judgment", "P&G FY2025 audited data")
    ) -Widths @(2.05, 2.05, 2.40) -HeaderFill $NAVY -FontSize 10)
    $metricTable = $doc.Tables.Item($doc.Tables.Count)
    for ($c = 1; $c -le 3; $c++) { $metricTable.Cell(1,$c).Range.Font.Color = $WHITE }

    [void](Add-LabeledParagraph -Label "Teaching thesis:" -Text "AI is most valuable when it makes student reasoning visible and testable. Students must predict before prompting, reconcile every output to the statements, and own the final judgment." -Fill $PALEBLUE)
    [void](Add-Paragraph -Text "The instructional engine" -Style "Heading 2")
    [void](Add-Table -Headers @("PREDICT", "ASK AI", "VERIFY", "EXPLAIN", "DECIDE") -Rows @(
        ,@("Commit before calculating", "Request an auditable draft", "Recompute and reconcile", "Translate numbers into meaning", "Make a defensible recommendation")
    ) -Widths @(1.18,1.22,1.22,1.30,1.58) -FontSize 8.8)
    [void](Add-Paragraph -Text "Case anchor: The Procter & Gamble Company, fiscal year ended June 30, 2025. Units are USD millions except per-share amounts." -Italic $true -Size 9.5 -Color $TEXTGRAY)

    Add-PageBreak

    # PAGE 2: lesson outcomes and design.
    [void](Add-Paragraph -Text "1. Lesson Purpose and Outcomes" -Style "Heading 1")
    [void](Add-Paragraph -Text "This lesson progresses from transaction mechanics to performance diagnosis. The first lab establishes how the three statements connect. The second asks whether accounting improvement is supported by cash. The third decomposes ROE and asks whether the result is operationally sustainable.")
    [void](Add-Paragraph -Text "Learning objectives" -Style "Heading 2")
    Add-Bullets @(
        "Trace operating, investing, and financing transactions through the income statement, balance sheet, and statement of cash flows.",
        "Use common-size, trend, and cash-conversion measures to identify the economic drivers behind reported results.",
        "Calculate and interpret profit margin, total asset turnover, equity multiplier, and DuPont return on equity.",
        "Distinguish a correct calculation from a strong financial interpretation.",
        "Use AI to draft analysis while verifying source values, equations, signs, labels, and assumptions.",
        "Communicate a concise decision supported by reconciled financial evidence."
    )
    [void](Add-Paragraph -Text "The three-lab progression" -Style "Heading 2")
    [void](Add-Table -Headers @("Lab", "Finance question", "AI's role", "Student control") -Rows @(
        ,@("1. Three-Statement Detective", "Where does each transaction appear?", "Draft the mapping and reconciliation", "Accounting equation and cash bridge"),
        ,@("2. P&G Earnings Quality", "What changed, and did cash support it?", "Rank drivers and draft the memo", "Common-size and cash-flow recomputation"),
        ,@("3. DuPont Boardroom", "Why is ROE high, and is it sustainable?", "Stress-test scenarios and challenge assumptions", "Identity check and recommendation")
    ) -Widths @(1.38,1.95,1.55,1.62) -FontSize 9.2)
    [void](Add-LabeledParagraph -Label "Success standard:" -Text "A polished AI answer receives no credit unless the team's verification evidence is visible." -Fill $GOLD)

    Add-PageBreak

    # PAGE 3: preparation and guardrails.
    [void](Add-Paragraph -Text "2. Preparation, Materials, and AI Guardrails" -Style "Heading 1")
    [void](Add-Paragraph -Text "Instructor preparation" -Style "Heading 2")
    Add-Bullets @(
        "Provide PG_FinancialStory_Student.xlsx or distribute the data tables in this playbook.",
        "Keep PG_FinancialStory_Instructor.xlsx and the solution manual private.",
        "Confirm that each team has an AI assistant, Excel or a calculator, and one shared submission document.",
        "Assign teams of three or four and display the five-step workflow throughout class.",
        "Decide whether students may use Python. The core lesson requires only AI plus Excel or a calculator."
    )
    [void](Add-Paragraph -Text "Suggested team roles" -Style "Heading 2")
    [void](Add-Table -Headers @("Role", "Responsibility") -Rows @(
        ,@("Finance lead", "Builds equations and checks financial meaning."),
        ,@("AI operator", "Enters the approved prompt and preserves the exact output."),
        ,@("Verification lead", "Recomputes values and tests identities independently."),
        ,@("Communicator", "Turns the verified results into the team's decision and briefing.")
    ) -Widths @(1.55,4.95) -FontSize 9.5)
    [void](Add-Paragraph -Text "AI use contract" -Style "Heading 2")
    Add-Numbers @(
        "Record a prediction before opening the AI tool.",
        "Give AI only the instructor-approved data; do not allow it to invent replacement figures.",
        "Ask for concise auditable rationale, equations, method, and limitations - not hidden chain-of-thought.",
        "Verify at least two central values and every accounting identity used in the recommendation.",
        "Document one assumption, error, or weakness found in the AI output.",
        "Make the final recommendation in the team's own words."
    )
    [void](Add-LabeledParagraph -Label "Required AI verification statement:" -Text '"We used [tool] to draft [task]. We verified [values/identities] using [method]. The AI assumed or misstated [item]. We corrected or qualified it by [action]. Our final judgment is [conclusion]."' -Fill $PALEBLUE)

    Add-PageBreak

    # PAGE 4: run of show.
    [void](Add-Paragraph -Text "3. Standard 120-Minute Run of Show" -Style "Heading 1")
    [void](Add-Table -Headers @("Time", "Phase", "Instructor move", "Student evidence") -Rows @(
        ,@("0-8", "Launch", "Show three statements and ask: Which one tells the truth?", "Initial claim and reason"),
        ,@("8-33", "Lab 1", "Release the six transactions; enforce prediction before AI.", "Balanced statements and AI audit"),
        ,@("33-43", "Debrief 1", "Surface four common sign/classification errors.", "Corrected transaction map"),
        ,@("43-73", "Lab 2", "Release P&G two-year data and the earnings-quality question.", "Driver table and analyst memo"),
        ,@("73-78", "Reset", "Five-minute break or retrieval check.", "One-sentence cash insight"),
        ,@("78-108", "Lab 3", "Release DuPont inputs and scenario choices.", "ROE identity and board recommendation"),
        ,@("108-116", "Share-out", "Call on teams with different conclusions.", "60-second briefings"),
        ,@("116-120", "Exit ticket", "Collect one equation, one insight, one AI risk.", "Individual response")
    ) -Widths @(0.62,0.90,3.05,1.93) -FontSize 8.6)
    [void](Add-Paragraph -Text "Opening hook" -Style "Heading 2")
    [void](Add-CodeBlock -Text "A company reports higher net income, lower operating cash flow, and a higher ROE. Is the company financially stronger?`r`n`r`nWrite YES, NO, or NOT ENOUGH INFORMATION. Give one reason. Do not calculate yet.")
    [void](Add-Paragraph -Text "Use the opening vote again at the end. A strong student response changes from a one-variable answer to a conditional answer that separates profitability, liquidity, cash generation, and leverage.")
    [void](Add-Paragraph -Text "Compression and extension" -Style "Heading 2")
    Add-Bullets @(
        "75-minute version: use 20 minutes per lab, provide completed calculations for Lab 2, and require only one team presentation.",
        "150-minute version: add peer review after Labs 2 and 3 and require teams to revise their AI prompt after identifying a weakness.",
        "Two-class version: complete Labs 1-2 on day one; begin day two with retrieval practice and finish Lab 3 plus presentations."
    )

    Add-PageBreak

    # LAB 1.
    [void](Add-Paragraph -Text "4. Lab 1 - Three-Statement Detective" -Style "Heading 1")
    [void](Add-LabeledParagraph -Label "Finance question:" -Text "How can profitable activity, investment, and financing change the three statements without violating the accounting equation?" -Fill $PALEBLUE)
    [void](Add-Paragraph -Text "Time: 25 minutes | Recommended groups: 3-4 | Deliverable: one reconciled transaction map")
    [void](Add-Paragraph -Text "Opening balance sheet and transactions" -Style "Heading 2")
    [void](Add-Table -Headers @("Opening assets", '$mm', "Opening liabilities and equity", '$mm') -Rows @(
        ,@("Cash", "200", "Accounts payable", "100"),
        ,@("Accounts receivable", "100", "Debt", "200"),
        ,@("Inventory", "150", "Shareholders' equity", "450"),
        ,@("Net PP&E", "300", "Total L + E", "750"),
        ,@("Total assets", "750", "", "")
    ) -Widths @(2.30,0.70,2.75,0.75) -FontSize 9.3)
    Add-Numbers @(
        'Purchase $80 of inventory on account.',
        'Sell inventory for $140 cash; the inventory cost is $70.',
        'Record $20 of depreciation expense.',
        'Purchase $60 of equipment for cash.',
        'Borrow $50 in cash through long-term debt.',
        'Pay $30 of accounts payable in cash. Ignore taxes.'
    )
    [void](Add-Paragraph -Text "Student workflow" -Style "Heading 2")
    [void](Add-Table -Headers @("Minutes", "Action") -Rows @(
        ,@("0-5", "Predict each transaction's statement effects without AI."),
        ,@("5-12", "Ask AI for a three-statement mapping and ending statements."),
        ,@("12-20", "Verify the accounting equation, net income, and cash-flow reconciliation."),
        ,@("20-25", "Explain the most important AI error risk and submit the map.")
    ) -Widths @(0.85,5.65) -FontSize 9.3)

    [void](Add-Paragraph -Text "Lab 1 AI Prompt and Instructor Key" -Style "Heading 1")
    [void](Add-Paragraph -Text "Student AI prompt" -Style "Heading 2")
    [void](Add-CodeBlock -Text "Act as a junior financial reporting analyst. Using only the opening balance sheet and six transactions supplied by the instructor, map each transaction to the income statement, balance sheet, and statement of cash flows under the indirect method. Show debit/credit logic only when it clarifies the statement effect. Then prepare ending statements and test: Assets = Liabilities + Equity and Beginning Cash + Net Change in Cash = Ending Cash. Include a concise REM block: Reasoning Summary, Equations, Method, and Limitations. Do not treat borrowing as revenue, CapEx as an immediate expense, or depreciation as a cash payment.")
    [void](Add-Paragraph -Text "Instructor answer" -Style "Heading 2")
    [void](Add-Table -Headers @("Measure", "Ending value or result") -Rows @(
        ,@("Revenue / COGS / depreciation", '$140 / $70 / $20'),
        ,@("Net income", '$50'),
        ,@("Cash flow from operations", '$110'),
        ,@("Cash flow from investing", '$(60)'),
        ,@("Cash flow from financing", '$50'),
        ,@("Net increase in cash / ending cash", '$100 / $300'),
        ,@("Ending assets", '$900'),
        ,@("Ending liabilities / equity", '$400 / $500')
    ) -Widths @(3.10,3.40) -FontSize 9.5)
    [void](Add-LabeledParagraph -Label "Reconciliation:" -Text 'CFO = $50 net income + $20 depreciation - $10 increase in inventory + $50 increase in accounts payable = $110. Ending assets of $900 equal liabilities of $400 plus equity of $500.' -Fill $GREEN)
    [void](Add-Paragraph -Text "Errors to surface in debrief" -Style "Heading 2")
    Add-Bullets @(
        'Expensing the $60 equipment purchase immediately instead of recording PP&E and an investing cash outflow.',
        'Treating the $50 borrowing as revenue instead of financing cash flow and debt.',
        "Showing depreciation as a cash outflow even though it is added back under the indirect method.",
        "Using the wrong sign for the increase in accounts payable in the CFO reconciliation."
    )

    # LAB 2.
    [void](Add-Paragraph -Text "5. Lab 2 - P&G Earnings Quality Lab" -Style "Heading 1")
    [void](Add-LabeledParagraph -Label "Finance question:" -Text "P&G's FY2025 profitability improved, but did the cash-flow evidence tell the same story?" -Fill $PALEBLUE)
    [void](Add-Paragraph -Text "Time: 30 minutes | Deliverable: a driver table and a 100-word analyst memo")
    [void](Add-Paragraph -Text "Approved data" -Style "Heading 2")
    [void](Add-Table -Headers @('Metric ($mm)', "FY2025", "FY2024") -Rows @(
        ,@("Net sales", "84,284", "84,039"),
        ,@("Cost of products sold", "41,164", "40,848"),
        ,@("SG&A", "22,669", "23,305"),
        ,@("Intangible impairment", "0", "1,341"),
        ,@("Operating income (EBIT)", "20,451", "18,545"),
        ,@("Net earnings", "16,065", "14,974"),
        ,@("Cash flow from operations", "17,817", "19,846"),
        ,@("Capital expenditures", "3,773", "3,322")
    ) -Widths @(3.50,1.50,1.50) -FontSize 9.4)
    [void](Add-Paragraph -Text "Required calculations" -Style "Heading 2")
    Add-Bullets @(
        "Sales growth and net-income growth.",
        "Gross margin, SG&A margin, EBIT margin, and net margin for both years.",
        "Basis-point change in each margin.",
        "CFO / net income and cash FCF = CFO - CapEx for both years.",
        "One operating explanation, one nonrecurring explanation, and one cash-flow concern."
    )
    [void](Add-Paragraph -Text "Student workflow" -Style "Heading 2")
    [void](Add-Table -Headers @("Minutes", "Action") -Rows @(
        ,@("0-4", "Predict whether FY2025 represents stronger earnings quality."),
        ,@("4-12", "Use AI to calculate and rank the apparent performance drivers."),
        ,@("12-22", "Recompute ratios in Excel and trace each numerator and denominator."),
        ,@("22-30", "Write a conditional conclusion and document one AI weakness.")
    ) -Widths @(0.85,5.65) -FontSize 9.3)

    [void](Add-Paragraph -Text "Lab 2 AI Prompt and Instructor Key" -Style "Heading 1")
    [void](Add-Paragraph -Text "Student AI prompt" -Style "Heading 2")
    [void](Add-CodeBlock -Text "Analyze P&G FY2025 versus FY2024 using only the instructor data table. Calculate sales growth, net-income growth, gross margin, SG&A margin, EBIT margin, net margin, CFO/net income, and cash free cash flow (CFO - CapEx). Show margin changes in basis points. Rank the drivers of EBIT-margin change, distinguish recurring operations from the absence of the FY2024 impairment, and assess whether cash conversion supports the profit improvement. Provide a concise REM block and a draft 100-word memo. State what cannot be concluded from only two years.")
    [void](Add-Paragraph -Text "Instructor answer bank" -Style "Heading 2")
    [void](Add-Table -Headers @("Measure", "FY2025", "FY2024", "Interpretation") -Rows @(
        ,@("Sales growth", "0.29%", "-", "Essentially flat top line"),
        ,@("Gross margin", "51.16%", "51.39%", "Down about 24 bps"),
        ,@("SG&A margin", "26.90%", "27.73%", "Improved about 84 bps"),
        ,@("EBIT margin", "24.26%", "22.07%", "Improved about 220 bps"),
        ,@("Net margin", "19.06%", "17.82%", "Improved about 124 bps"),
        ,@("CFO / net income", "1.109x", "1.325x", "Cash conversion weakened"),
        ,@("Cash FCF", '$14,044', '$16,524', "Down about 15.0%")
    ) -Widths @(1.55,1.05,1.05,2.85) -FontSize 8.8)
    [void](Add-LabeledParagraph -Label "Model conclusion:" -Text 'FY2025 reported profitability improved, helped by lower SG&A and by the nonrecurrence of a $1,341 million impairment. However, CFO and cash FCF declined while net income increased. That divergence merits investigation, but two years alone do not prove poor earnings quality.' -Fill $GREEN)
    [void](Add-Paragraph -Text "Debrief questions" -Style "Heading 2")
    Add-Bullets @(
        "Why can EBIT margin improve while gross margin declines?",
        "Should the absence of an impairment be called growth, improvement, or normalization?",
        "Is CFO greater than net income sufficient evidence of high earnings quality?",
        "Which additional notes or working-capital details would you request before investing?"
    )

    # LAB 3.
    [void](Add-Paragraph -Text "6. Lab 3 - DuPont Boardroom" -Style "Heading 1")
    [void](Add-LabeledParagraph -Label "Finance question:" -Text "Is P&G's high ROE primarily an operating achievement, an efficiency achievement, or a leverage effect?" -Fill $PALEBLUE)
    [void](Add-Paragraph -Text "Time: 30 minutes | Deliverable: one-slide board recommendation and 60-second briefing")
    [void](Add-Paragraph -Text "Approved inputs" -Style "Heading 2")
    [void](Add-Table -Headers @('Input ($mm)', "Amount") -Rows @(
        ,@("Net sales FY2025", "84,284"),
        ,@("Net earnings FY2025", "16,065"),
        ,@("Total assets FY2025 / FY2024", "125,231 / 122,370"),
        ,@("Total equity FY2025 / FY2024", "52,284 / 50,559")
    ) -Widths @(3.75,2.75) -FontSize 9.5)
    [void](Add-Paragraph -Text "Required equations" -Style "Heading 2")
    [void](Add-CodeBlock -Text "Profit Margin = Net Earnings / Net Sales`r`nTotal Asset Turnover = Net Sales / Average Total Assets`r`nEquity Multiplier = Average Total Assets / Average Total Equity`r`nROE = Profit Margin x Asset Turnover x Equity Multiplier")
    [void](Add-Paragraph -Text "Board scenarios" -Style "Heading 2")
    [void](Add-Table -Headers @("Scenario", "Change", "Board question") -Rows @(
        ,@("Deleverage", "Equity multiplier falls to 2.00x", "How much ROE is sacrificed for lower financial risk?"),
        ,@("Efficiency", "Asset turnover rises to 0.72x", "Can more sales be generated without adding assets?"),
        ,@("Margin pressure", "Profit margin falls to 18.0%", "How exposed is ROE to pricing or cost pressure?")
    ) -Widths @(1.20,2.10,3.20) -FontSize 9.1)
    [void](Add-Paragraph -Text "Student workflow" -Style "Heading 2")
    [void](Add-Table -Headers @("Minutes", "Action") -Rows @(
        ,@("0-5", "Predict which factor contributes most to ROE."),
        ,@("5-12", "Use AI to calculate the base case and three scenarios."),
        ,@("12-22", "Verify the DuPont identity and challenge the AI's interpretation."),
        ,@("22-30", "Recommend one way to sustain ROE without increasing leverage.")
    ) -Widths @(0.85,5.65) -FontSize 9.3)

    [void](Add-Paragraph -Text "Lab 3 AI Prompt and Instructor Key" -Style "Heading 1")
    [void](Add-Paragraph -Text "Student AI prompt" -Style "Heading 2")
    [void](Add-CodeBlock -Text "Act as a board finance analyst. Calculate P&G's three-factor DuPont ROE using FY2025 net sales and net earnings and average FY2024-FY2025 assets and equity. Reconcile the product to direct ROE. Then calculate ROE under three separate scenarios: equity multiplier 2.00x, asset turnover 0.72x, and profit margin 18.0%, holding other base factors constant. Explain which factor drives the level of ROE, why leverage can magnify both returns and risk, and recommend one operating action that could sustain ROE without more leverage. Include concise REM and limitations.")
    [void](Add-Paragraph -Text "Instructor answer" -Style "Heading 2")
    [void](Add-Table -Headers @("Factor or scenario", "Result", "Meaning") -Rows @(
        ,@("Profit margin", "19.06%", "Strong earnings retained per sales dollar"),
        ,@("Asset turnover", "0.6808x", "Less than one sales dollar per asset dollar"),
        ,@("Equity multiplier", "2.4076x", "Material leverage magnifies operating returns"),
        ,@("Base DuPont ROE", "31.24%", "Matches direct net earnings / average equity"),
        ,@("Deleverage to 2.00x", "25.96%", "Lower leverage lowers both ROE and financial risk"),
        ,@("Turnover improves to 0.72x", "33.04%", "Efficiency can raise ROE without more leverage"),
        ,@("Margin falls to 18.0%", "29.49%", "ROE remains sensitive to operating margin")
    ) -Widths @(2.35,1.05,3.10) -FontSize 8.8)
    [void](Add-LabeledParagraph -Label "Interpretation standard:" -Text "Do not say the equity multiplier 'contributes the most' merely because 2.4076 is numerically larger than 0.6808 or 0.1906. The factors use different scales. Instead, describe the economic roles of margin, efficiency, and leverage, and use scenarios to assess sensitivity." -Fill $GOLD)
    [void](Add-Paragraph -Text "AI misconceptions to challenge" -Style "Heading 2")
    Add-Bullets @(
        "High ROE always means the company is operationally superior.",
        "A higher equity multiplier is unambiguously good.",
        "Ending balances are interchangeable with average balances.",
        "A DuPont identity explains causality rather than decomposing an observed result."
    )

    # Assessment.
    [void](Add-Paragraph -Text "7. Assessment and Discussion" -Style "Heading 1")
    [void](Add-Paragraph -Text "Common rubric for all three labs" -Style "Heading 2")
    [void](Add-Table -Headers @("Criterion", "Weight", "Evidence of mastery") -Rows @(
        ,@("Financial accuracy", "40%", "Correct values, signs, classifications, and identities"),
        ,@("Interpretation and decision", "25%", "Conclusion follows from the verified evidence"),
        ,@("AI verification", "20%", "Prompt, checks, correction, and limitation are documented"),
        ,@("Communication", "15%", "Concise, readable, decision-oriented presentation")
    ) -Widths @(1.65,0.75,4.10) -FontSize 9.3)
    [void](Add-Paragraph -Text "Whole-class synthesis questions" -Style "Heading 2")
    Add-Bullets @(
        "Which statement is best for measuring profitability, liquidity, and cash generation? Why is the answer different for each question?",
        "How can net income rise while operating cash flow falls?",
        "When does a ratio reveal a problem, and when does it merely identify a question for further investigation?",
        "Which conclusion in today's work depended most on judgment rather than calculation?",
        "What did AI do well, and where did human financial knowledge change the answer?"
    )
    [void](Add-Paragraph -Text "Individual exit ticket" -Style "Heading 2")
    [void](Add-CodeBlock -Text "1. Write one equation you can now explain, not merely calculate.`r`n2. State one insight about P&G that is supported by at least two statements.`r`n3. Name one AI output that required verification and how you checked it.`r`n4. Return to the opening question: Is the company financially stronger? Give a conditional answer.")
    [void](Add-LabeledParagraph -Label "Fast grading option:" -Text "Score each response 0-2 for financial accuracy, evidence, and AI verification. Six points total; five points indicates mastery." -Fill $PALEBLUE)

    # Student-facing one-page checklist.
    [void](Add-Paragraph -Text "Appendix A - Student AI Analysis Checklist" -Style "Heading 1")
    [void](Add-Paragraph -Text "Use this checklist before submitting any AI-assisted financial statement analysis.")
    [void](Add-Table -Headers @("Check", "Student evidence") -Rows @(
        ,@("Prediction recorded", "Our initial expectation appears before the AI output."),
        ,@("Source controlled", "Every input came from the instructor dataset or cited statement."),
        ,@("Formula visible", "The equation and substitutions can be inspected."),
        ,@("Signs verified", "Expenses, CapEx, working capital, and debt are classified correctly."),
        ,@("Identity passed", "The accounting, cash, FCF, or DuPont reconciliation works."),
        ,@("Units and periods match", "USD millions, fiscal year, percentages, and averages are consistent."),
        ,@("AI weakness documented", "We identified and corrected or qualified at least one issue."),
        ,@("Meaning explained", "We translated the calculation into economic language."),
        ,@("Decision owned", "The recommendation is in our words and follows from evidence."),
        ,@("Limitation stated", "We named information that could change the conclusion.")
    ) -Widths @(1.85,4.65) -FontSize 9.4)
    [void](Add-Paragraph -Text "REM response frame" -Style "Heading 2")
    [void](Add-Table -Headers @("Element", "What to report") -Rows @(
        ,@("Reasoning Summary", "The financial question and the direction of the evidence."),
        ,@("Equations", "The exact relationships used in the analysis."),
        ,@("Method", "Source, period, calculations, verification, and corrections."),
        ,@("Limitations", "Data gaps, simplifying assumptions, non-GAAP definitions, or model risk.")
    ) -Widths @(1.65,4.85) -FontSize 9.5)

    [void](Add-Paragraph -Text "Appendix B - Data Sources and Instructor Notes" -Style "Heading 1")
    [void](Add-Table -Headers @("Source", "Use in lesson") -Rows @(
        ,@("P&G FY2025 Form 10-K, p. 36", "Income statement values and margin analysis"),
        ,@("P&G FY2025 Form 10-K, p. 37", "FY2025 and FY2024 balance-sheet values"),
        ,@("P&G FY2025 Form 10-K, p. 39", "Operating cash flow, D&A, working capital, and CapEx"),
        ,@("P&G FY2025 Form 10-K, p. 30", "Company-defined adjusted free cash flow"),
        ,@("PG_FinancialStory_Student.xlsx", "Student source and calculation workbook"),
        ,@("PG_FinancialStory_Instructor.xlsx", "Instructor calculations and audit checks")
    ) -Widths @(2.70,3.80) -FontSize 9.4)
    [void](Add-Paragraph -Text "Instructor cautions" -Style "Heading 2")
    Add-Bullets @(
        "P&G's adjusted free cash flow is a company-defined non-GAAP measure. It must be labeled and reconciled rather than silently substituted for textbook FCF.",
        "Use consolidated net earnings consistently with total shareholders' equity in the base DuPont exercise.",
        "A two-year comparison can identify divergence but cannot establish a durable trend or causality.",
        "Negative operating working capital can be a feature of a strong consumer-products model; analyze the business mechanism before labeling it a weakness.",
        "The goal is not agreement with the AI. The goal is a transparent, verified, and financially meaningful conclusion."
    )
    [void](Add-LabeledParagraph -Label "Package location:" -Text "PG_Financial_Story_Package contains the student workbook, source 10-K, model prompt, solution workbook, visualizations, and grading materials used by this lesson." -Fill $PALEBLUE)

    $doc.Fields.Update() | Out-Null
    $doc.Repaginate()
    $doc.SaveAs2($OutputDocx, 16)
    Write-Output $OutputDocx
}
finally {
    if ($null -ne $doc) {
        try { $doc.Close([ref]$false) } catch { }
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) } catch { }
    }
    if ($null -ne $word) {
        try { $word.Quit([ref]$false) } catch { }
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) } catch { }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    if ($wordProcessId -gt 0) {
        Stop-Process -Id $wordProcessId -Force -ErrorAction SilentlyContinue
    }
}
