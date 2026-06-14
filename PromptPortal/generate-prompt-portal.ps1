<#
.SYNOPSIS
  Refreshes the Analysis Prompts Portal (index.html) from the latest content of your 8 .docx prompt files.

  Run this file from inside the PromptPortal folder.
  It will automatically read the prompts from the sibling ../Analysis folder.
  This makes it easy to move/rename the Analysis folder frequently.
#>

$ErrorActionPreference = 'Stop'

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Support moving the Analysis folder frequently.
# Run this .ps1 from inside the PromptPortal folder.
# It will automatically pick the prompts from the sibling ../Analysis folder.
$portalDir = $scriptDir
$analysisRoot = Join-Path (Split-Path -Parent $portalDir) 'Analysis'

$indexPath = Join-Path $portalDir 'index.html'
$templatePath = Join-Path $portalDir 'portal-template.html'

if (-not (Test-Path $analysisRoot)) {
    Write-Warning "Could not locate Analysis folder at $analysisRoot"
    Write-Host "Make sure PromptPortal and Analysis are siblings under the same parent folder." -ForegroundColor Yellow
    $analysisRoot = Split-Path -Parent $portalDir   # fallback for old nested layout
}

function Get-DocxPlainText([string]$Path) {
    if (-not (Test-Path $Path)) { return "[FILE NOT FOUND: $Path]" }
    try {
        Add-Type -AssemblyName 'System.IO.Compression.FileSystem' -ErrorAction SilentlyContinue | Out-Null
        $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
        $entry = $zip.GetEntry('word/document.xml')
        if (-not $entry) { $zip.Dispose(); return "[NO DOCUMENT XML]" }
        $reader = New-Object System.IO.StreamReader($entry.Open(), [System.Text.Encoding]::UTF8)
        $xml = $reader.ReadToEnd()
        $reader.Dispose(); $zip.Dispose()

        $text = [regex]::Replace($xml, '<w:p[^>]*>', "`n`n")
        $text = [regex]::Replace($text, '<[^>]+>', '')
        $text = $text -replace '&amp;', '&' -replace '&lt;', '<' -replace '&gt;', '>' `
                      -replace '&quot;', '"' -replace '&#39;', "'" `
                      -replace '&#x201C;', '"' -replace '&#x201D;', '"' `
                      -replace '&#x2018;', "'" -replace '&#x2019;', "'"
        $text = [regex]::Replace($text, ' {2,}', ' ').Trim()
        $text = [regex]::Replace($text, "`n{3,}", "`n`n")
        if ($text.Length -lt 15) {
            $matches = [regex]::Matches($xml, '<w:t[^>]*>([^<]*)</w:t>')
            $parts = foreach ($m in $matches) { $m.Groups[1].Value }
            $text = ($parts -join ' ').Trim()
            $text = [regex]::Replace($text, '\s+', ' ')
        }
        return $text
    } catch {
        return "[EXTRACTION ERROR: $($_.Exception.Message)]"
    }
}

function EscapeJsTemplate([string]$s) {
    if ([string]::IsNullOrEmpty($s)) { return '' }
    $s = $s -replace '\\', '\\'
    $s = $s -replace '`', '\`'
    $s = $s -replace '\${', '\${'
    return $s
}

$entries = @(
    @{ id='indian_market_news'; title='Indian Market News'; region='India'; emoji='🇮🇳'; subtitle='Stocks with positive catalysts & developments'; rel='India\Indian_Market_News.docx' }
    @{ id='indian_results_week'; title='Indian Results Week'; region='India'; emoji='🇮🇳'; subtitle='Weekly results & earnings analysis'; rel='India\Indian_Results_Week.docx' }
    @{ id='indian_stock_analysis'; title='Indian Stock Analysis'; region='India'; emoji='🇮🇳'; subtitle='In-depth stock & sector research'; rel='India\Indian_Stock_Analysis.docx' }
    @{ id='indian_stock_predictions'; title='Indian Stock Predictions'; region='India'; emoji='🇮🇳'; subtitle='Forward-looking price & trend forecasts'; rel='India\Indian_Stock_Predictions.docx' }
    @{ id='usa_market_news'; title='USA Market News'; region='USA'; emoji='🇺🇸'; subtitle='Stocks with positive catalysts & developments'; rel='USA\USA_Market_News.docx' }
    @{ id='usa_results_week'; title='USA Results Week'; region='USA'; emoji='🇺🇸'; subtitle='Weekly results & earnings analysis'; rel='USA\USA_Results_Week.docx' }
    @{ id='usa_stock_analysis'; title='USA Stock Analysis'; region='USA'; emoji='🇺🇸'; subtitle='In-depth stock & sector research'; rel='USA\USA_Stock_Analysis.docx' }
    @{ id='usa_stock_predictions'; title='USA Stock Predictions'; region='USA'; emoji='🇺🇸'; subtitle='Forward-looking price & trend forecasts'; rel='USA\USA_Stock_Predictions.docx' }
)

# Prefix used in the portal for "Copy .docx path" buttons and footer.
# Using ../Analysis so the whole Analysis folder (with prompts) can be moved frequently.
$sourcePrefix = '../Analysis'

Write-Host "Extracting prompts..." -ForegroundColor Cyan
$promptTexts = @{}
foreach ($e in $entries) {
    $p = Join-Path $analysisRoot $e.rel
    $txt = Get-DocxPlainText $p
    $promptTexts[$e.id] = $txt
    Write-Host ("  {0,-28} {1,6} chars" -f $e.title, $txt.Length) -ForegroundColor Green
}

# Build the JS snippets that will be injected
# We want to emit in the HTML:
#   const PROMPTS = {
#     'id1': `long prompt text...`,
#     ...
#   };
$lines = @('const PROMPTS = {')
for ($i = 0; $i -lt $entries.Count; $i++) {
    $e = $entries[$i]
    $comma = if ($i -lt ($entries.Count - 1)) { ',' } else { '' }
    $safe = EscapeJsTemplate $promptTexts[$e.id]
    # Emit:   'id': `escaped text`,
    # Use single-quoted PS string for the skeleton + insert the already-escaped $safe
    $line = "  '$($e.id)': ``$safe``$comma"
    $lines += $line
}
$lines += '};'
$promptsJs = $lines -join [Environment]::NewLine

$lines = @('const PROMPT_META = [')
for ($i = 0; $i -lt $entries.Count; $i++) {
    $e = $entries[$i]
    $comma = if ($i -lt ($entries.Count - 1)) { ',' } else { '' }
    $fullRel = "$sourcePrefix/$($e.rel -replace '\\\\','/')"
    $line = "  { id: '$($e.id)', title: '$($e.title)', region: '$($e.region)', emoji: '$($e.emoji)', subtitle: '$($e.subtitle)', file: '$fullRel' }$comma"
    $lines += $line
}
$lines += '];'
$metaJs = $lines -join [Environment]::NewLine

$generated = Get-Date -Format 'yyyy-MM-dd HH:mm'

if (-not (Test-Path $templatePath)) {
    Write-Error "Template not found at $templatePath. Make sure portal-template.html is next to this script."
    exit 1
}

$html = Get-Content -Path $templatePath -Raw -Encoding UTF8
$html = $html.Replace('__PROMPTS_JS__', $promptsJs)
$html = $html.Replace('__META_JS__', $metaJs)
$html = $html.Replace('__DATE__', $generated)

$html | Out-File -FilePath $indexPath -Encoding UTF8 -Force

Write-Host ""
Write-Host "✅ Portal refreshed successfully!" -ForegroundColor Green
Write-Host "   $indexPath" -ForegroundColor White
Write-Host ""
Write-Host "Prompts picked from: $analysisRoot (../Analysis relative to PromptPortal)" -ForegroundColor Gray
Write-Host ""
Write-Host "Open it with:" -ForegroundColor Gray
Write-Host "   start-process `"$indexPath`"" -ForegroundColor Cyan
Write-Host "   (or just double-click the index.html in File Explorer)" -ForegroundColor Gray
Write-Host ""
Write-Host "Date behavior (live in browser):" -ForegroundColor Gray
Write-Host "   • India prompts include 'Current date (IST): YYYY-MM-DD'" -ForegroundColor Gray
Write-Host "   • USA  prompts include 'Current date (CST): YYYY-MM-DD'" -ForegroundColor Gray
Write-Host "   (computed from your system clock using proper market timezones)" -ForegroundColor Gray
Write-Host ""
Write-Host "Tip: You can move the Analysis folder freely. Just keep PromptPortal as its sibling and re-run this script after changes." -ForegroundColor Gray
