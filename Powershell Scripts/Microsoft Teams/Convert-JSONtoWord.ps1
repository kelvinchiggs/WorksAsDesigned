<#
.SYNOPSIS
  Converts Teams audit JSON (quickWins/findings/appendix) into a Word report using PSWriteWord,
  then post-processes with Word COM to generate a TOC and apply conditional row shading.

.REQUIREMENTS
  - PowerShell 5.1+
  - PSWriteWord module
  - Optional (recommended): Microsoft Word installed (for TOC + row shading)

.EXAMPLE
  .\Convert-TeamsAuditJsonToWord.ps1 -JsonPath "C:\Temp\TeamsAudit_ReportData.json" -OutDocx "C:\Temp\TeamsAudit_Report.docx" -CustomerName "Contoso" -TenantName "contoso.onmicrosoft.com"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$JsonPath,

    [Parameter(Mandatory)]
    [string]$OutDocx,

    [string]$CustomerName = "Customer",
    [string]$TenantName   = "Tenant",
    [string]$PreparedBy   = $env:USERNAME,
    [ValidateSet("1.0","1.1","1.2","2.0")]
    [string]$ReportVersion = "1.0",

    # If you only want Amber/Red findings in the detailed section
    [switch]$OnlyAmberRed,

    # If you only want rows where CurrentValue differs from BestPractice
    [switch]$OnlyDiffs
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Ensure-Module {
    param([Parameter(Mandatory)][string]$Name)
    if (-not (Get-Module -ListAvailable -Name $Name)) {
        Install-Module $Name -Scope CurrentUser -Force
    }
    Import-Module $Name -ErrorAction Stop
}

function Get-RagCounts {
    param([Parameter(Mandatory)]$Items)

    $norm = $Items | ForEach-Object {
        $rag = $_.RAGStatus
        if (-not $rag) { $rag = $_.RAG }  # fallback
        [pscustomobject]@{ RAG = [string]$rag }
    }

    $g = $norm | Group-Object RAG
    $get = { param($name) ($g | Where-Object Name -eq $name | Select-Object -ExpandProperty Count -ErrorAction SilentlyContinue) }

    [pscustomobject]@{
        Red   = (& $get 'Red')   ?? 0
        Amber = (& $get 'Amber') ?? 0
        Green = (& $get 'Green') ?? 0
        Total = $Items.Count
    }
}

function Try-AddTocAndShadingWithWord {
    param(
        [Parameter(Mandatory)][string]$DocxPath
    )

    # Try COM automation; if Word isn't installed, we just skip gracefully.
    try {
        $word = New-Object -ComObject Word.Application
    } catch {
        Write-Warning "Microsoft Word COM automation not available. Skipping TOC + conditional shading."
        return
    }

    $wdCollapseEnd = 0
    $wdFieldTOC    = 13
    $wdAutoFitContent = 1

    # Light, readable shading colours (Excel-like)
    function Get-WordColor([int]$r,[int]$g,[int]$b) {
        # Word expects BGR integer
        return ($b -shl 16) -bor ($g -shl 8) -bor $r
    }
    $colorRedLight   = Get-WordColor 255 199 206   # light red
    $colorAmberLight = Get-WordColor 255 235 156   # light amber

    try {
        $word.Visible = $false
        $doc = $word.Documents.Open($DocxPath)

        # 1) Insert TOC at the marker line "%%TOC%%"
        $range = $doc.Content
        $find = $range.Find
        $find.ClearFormatting() | Out-Null
        $find.Text = "%%TOC%%"
        $found = $find.Execute()

        if ($found) {
            # Replace marker with a real TOC
            $range.Text = ""  # remove marker text
            $range.Collapse($wdCollapseEnd) | Out-Null
            $doc.TablesOfContents.Add($range, $true, 1, 3, $true, "", $true, $true) | Out-Null
            # Update TOC
            if ($doc.TablesOfContents.Count -ge 1) { $doc.TablesOfContents.Item(1).Update() | Out-Null }
        } else {
            Write-Warning "TOC marker '%%TOC%%' not found. Skipping TOC insertion."
        }

        # 2) Conditional row shading for all tables that contain a RAG column
        foreach ($tbl in $doc.Tables) {
            $rows = $tbl.Rows.Count
            $cols = $tbl.Columns.Count
            if ($rows -lt 2 -or $cols -lt 1) { continue }

            # Identify the RAG column index by header text (row 1)
            $ragCol = $null
            for ($c = 1; $c -le $cols; $c++) {
                $hdr = $tbl.Cell(1,$c).Range.Text
                $hdr = ($hdr -replace "[\r\a]+","").Trim()
                if ($hdr -match '^(RAGStatus|RAG)$') { $ragCol = $c; break }
            }
            if (-not $ragCol) { continue }

            # Shade data rows (skip header row 1)
            for ($r = 2; $r -le $rows; $r++) {
                $val = $tbl.Cell($r,$ragCol).Range.Text
                $val = ($val -replace "[\r\a]+","").Trim()

                if ($val -eq 'Red') {
                    $tbl.Rows.Item($r).Shading.BackgroundPatternColor = $colorRedLight
                } elseif ($val -eq 'Amber') {
                    $tbl.Rows.Item($r).Shading.BackgroundPatternColor = $colorAmberLight
                }
            }

            # Auto-fit table (keeps it looking clean)
            $tbl.AutoFitBehavior($wdAutoFitContent) | Out-Null
        }

        $doc.Fields.Update() | Out-Null
        $doc.Save()
        $doc.Close()
    }
    finally {
        $word.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

# -------------------------
# MAIN
# -------------------------

Ensure-Module -Name PSWriteWord

if (-not (Test-Path -LiteralPath $JsonPath)) {
    throw "JSON file not found: $JsonPath"
}

$raw  = Get-Content -LiteralPath $JsonPath -Raw
$data = $raw | ConvertFrom-Json -Depth 200

# Basic validation
if (-not $data.findings) { throw "JSON does not contain 'findings'." }

$findings = @($data.findings)
$quickWins = @($data.quickWins)

if ($OnlyAmberRed) {
    $findings = $findings | Where-Object { $_.RAGStatus -in @('Red','Amber') }
}
if ($OnlyDiffs) {
    $findings = $findings | Where-Object { [string]$_.CurrentValue -ne [string]$_.BestPractice }
}

$countsAll = Get-RagCounts -Items $data.findings
$countsSel = Get-RagCounts -Items $findings

# ---- Create Word document (Well-Architected style layout)
$doc = New-WordDocument

# Cover page
Add-WordText -WordDocument $doc -Text "Microsoft Teams Well-Architected Review" -HeadingType Heading1
Add-WordText -WordDocument $doc -Text "$CustomerName — $TenantName" -HeadingType Heading2
Add-WordText -WordDocument $doc -Text "Report Version: $ReportVersion"
Add-WordText -WordDocument $doc -Text "Prepared by: $PreparedBy"
Add-WordText -WordDocument $doc -Text "Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Add-WordText -WordDocument $doc -Text ""
Add-WordText -WordDocument $doc -Text "CONFIDENTIAL — For internal use and authorised stakeholders only."
Add-WordPageBreak -WordDocument $doc

# Table of Contents placeholder
Add-WordText -WordDocument $doc -Text "Table of Contents" -HeadingType Heading1
Add-WordText -WordDocument $doc -Text "%%TOC%%"  # marker for COM automation to replace with real TOC
Add-WordPageBreak -WordDocument $doc

# Executive Summary
Add-WordText -WordDocument $doc -Text "Executive Summary" -HeadingType Heading1
Add-WordText -WordDocument $doc -Text "This report assesses Microsoft Teams configuration against recommended practices, highlighting control gaps and remediation actions prioritised by risk."

$execTable = @(
    [pscustomobject]@{ Metric = "Total Findings (All)";        Value = $countsAll.Total }
    [pscustomobject]@{ Metric = "Red (All)";                   Value = $countsAll.Red }
    [pscustomobject]@{ Metric = "Amber (All)";                 Value = $countsAll.Amber }
    [pscustomobject]@{ Metric = "Green (All)";                 Value = $countsAll.Green }
    [pscustomobject]@{ Metric = "Total Findings (In Scope)";   Value = $countsSel.Total }
    [pscustomobject]@{ Metric = "Red (In Scope)";              Value = $countsSel.Red }
    [pscustomobject]@{ Metric = "Amber (In Scope)";            Value = $countsSel.Amber }
    [pscustomobject]@{ Metric = "Green (In Scope)";            Value = $countsSel.Green }
)

Add-WordTable -WordDocument $doc -DataTable $execTable -Design LightList -AutoFit Window
Add-WordText -WordDocument $doc -Text ""

# Quick wins
if ($quickWins.Count -gt 0) {
    Add-WordText -WordDocument $doc -Text "Priority Remediation (Quick Wins)" -HeadingType Heading2
    Add-WordText -WordDocument $doc -Text "The following items offer high risk reduction with relatively low implementation effort."

    $qwTable = $quickWins | Select-Object `
        Domain, SecurityDomain, PolicyName, Setting, CurrentValue, BestPractice, RAGStatus, RiskLevel, RiskScore, Recommendation

    Add-WordTable -WordDocument $doc -DataTable $qwTable -Design LightList -AutoFit Window
} else {
    Add-WordText -WordDocument $doc -Text "Priority Remediation (Quick Wins)" -HeadingType Heading2
    Add-WordText -WordDocument $doc -Text "No quick wins were provided in the input JSON."
}

Add-WordPageBreak -WordDocument $doc

# Detailed findings grouped by Domain
Add-WordText -WordDocument $doc -Text "Detailed Findings" -HeadingType Heading1
Add-WordText -WordDocument $doc -Text "Findings are grouped by capability domain. Remediation should prioritise Red then Amber items."

$domains = $findings | Group-Object Domain | Sort-Object Name
foreach ($d in $domains) {
    Add-WordText -WordDocument $doc -Text $d.Name -HeadingType Heading2

    $rows = $d.Group |
        Sort-Object @{Expression="RAGStatus";Descending=$false}, @{Expression="RiskScore";Descending=$true} |
        Select-Object PolicyName, Setting, CurrentValue, BestPractice, RAGStatus, RiskLevel, RiskScore, Recommendation

    Add-WordTable -WordDocument $doc -DataTable $rows -Design LightList -AutoFit Window
    Add-WordText -WordDocument $doc -Text ""
}

# Appendix
Add-WordPageBreak -WordDocument $doc
Add-WordText -WordDocument $doc -Text "Appendix" -HeadingType Heading1

if ($data.appendix -and $data.appendix.officialLinks) {
    Add-WordText -WordDocument $doc -Text "Official References" -HeadingType Heading2

    $links = @($data.appendix.officialLinks) | ForEach-Object {
        [pscustomobject]@{ Link = [string]$_ }
    }

    Add-WordTable -WordDocument $doc -DataTable $links -Design LightList -AutoFit Window
} else {
    Add-WordText -WordDocument $doc -Text "Official References" -HeadingType Heading2
    Add-WordText -WordDocument $doc -Text "No official links were provided in the input JSON."
}

# Save PSWriteWord output
Save-WordDocument -WordDocument $doc -FilePath $OutDocx
Write-Host "Created base report: $OutDocx"

# Post-process: TOC + conditional shading (Red/Amber)
Try-AddTocAndShadingWithWord -DocxPath $OutDocx
Write-Host "Completed post-processing (TOC + shading where available). Output: $OutDocx"
