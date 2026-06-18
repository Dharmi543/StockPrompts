<#
.SYNOPSIS
  Refreshes the Analysis Prompts Portal (index.html) from the latest content of your .docx prompt files.

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

    Add-Type -AssemblyName 'System.IO.Compression.FileSystem' -ErrorAction SilentlyContinue | Out-Null

    $zip = $null
    $maxRetries = 5
    $baseDelay = 300  # ms

    for ($attempt = 0; $attempt -lt $maxRetries; $attempt++) {
        try {
            $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
            break
        } catch {
            $msg = $_.Exception.Message
            if ($attempt -eq ($maxRetries - 1)) {
                # Final failure - give a helpful message
                if ($msg -like "*being used by another process*" -or $msg -like "*access*denied*" -or $msg -like "*cannot access*") {
                    return "[FILE IN USE] The .docx file is currently locked (probably open in Microsoft Word or being synced by OneDrive).`n`nClose the document in Word, wait a few seconds for OneDrive sync to finish, then re-run generate-prompt-portal.ps1."
                }
                return "[EXTRACTION ERROR] Could not open the file after $maxRetries attempts.`nDetails: $msg`n`nTip: Close Microsoft Word completely and make sure no other program has the file open."
            }
            # Exponential backoff
            $delay = $baseDelay * [Math]::Pow(1.6, $attempt)
            Start-Sleep -Milliseconds ([int]$delay)
        }
    }

    if ($null -eq $zip) {
        return "[EXTRACTION ERROR] Failed to open file after retries."
    }

    try {
        $entry = $zip.GetEntry('word/document.xml')
        if (-not $entry) { 
            $zip.Dispose()
            return "[NO DOCUMENT XML]" 
        }

        $reader = New-Object System.IO.StreamReader($entry.Open(), [System.Text.Encoding]::UTF8)
        $xmlContent = $reader.ReadToEnd()
        $reader.Dispose()
        $zip.Dispose()

        # Proper XML parsing with namespace
        $doc = [xml]$xmlContent
        $nsMgr = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
        $nsMgr.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")

        $paragraphs = @()
        $body = $doc.SelectSingleNode("//w:body", $nsMgr)
        $searchRoot = if ($body) { $body } else { $doc }

        $paraNodes = $searchRoot.SelectNodes(".//w:p", $nsMgr)

        foreach ($p in $paraNodes) {
            $textParts = $p.SelectNodes(".//w:t | .//w:delText", $nsMgr) | ForEach-Object { $_.InnerText }
            $paraText = ($textParts -join "").Trim()
            if ($paraText) { $paragraphs += $paraText }
        }

        $result = $paragraphs -join "`n`n"

        if ([string]::IsNullOrWhiteSpace($result) -or $result.Length -lt 30) {
            $allText = $doc.SelectNodes("//w:t | //w:delText", $nsMgr) | ForEach-Object { $_.InnerText }
            $result = ($allText -join " ").Trim()
            $result = [regex]::Replace($result, '\s+', ' ')
        }

        $result = $result -replace ' {2,}', ' '
        $result = $result -replace "`n{3,}", "`n`n"
        return $result.Trim()
    }
    catch {
        if ($zip) { $zip.Dispose() }
        return "[EXTRACTION ERROR] $($_.Exception.Message)"
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
    @{ id='indian_market_news'; title='Indian Market News'; region='India'; emoji=''; subtitle='Stocks with positive catalysts & developments'; rel='India\Indian_Market_News.docx' }
    @{ id='indian_results_week'; title='Indian Results Week'; region='India'; emoji=''; subtitle='Weekly results & earnings analysis'; rel='India\Indian_Results_Week.docx' }
    @{ id='indian_stock_analysis'; title='Indian Stock Analysis'; region='India'; emoji=''; subtitle='In-depth stock & sector research'; rel='India\Indian_Stock_Analysis.docx' }
    @{ id='indian_stock_predictions'; title='Indian Stock Predictions'; region='India'; emoji=''; subtitle='Forward-looking price & trend forecasts'; rel='India\Indian_Stock_Predictions.docx' }
    @{ id='indian_mutual_funds'; title='Indian Mutual Funds'; region='India'; emoji=''; subtitle='Mutual fund schemes, performance & recommendations'; rel='India\Indian_Mutual_Funds.docx' }
    @{ id='usa_market_news'; title='USA Market News'; region='USA'; emoji=''; subtitle='Stocks with positive catalysts & developments'; rel='USA\USA_Market_News.docx' }
    @{ id='usa_results_week'; title='USA Results Week'; region='USA'; emoji=''; subtitle='Weekly results & earnings analysis'; rel='USA\USA_Results_Week.docx' }
    @{ id='usa_stock_analysis'; title='USA Stock Analysis'; region='USA'; emoji=''; subtitle='In-depth stock & sector research'; rel='USA\USA_Stock_Analysis.docx' }
    @{ id='usa_stock_predictions'; title='USA Stock Predictions'; region='USA'; emoji=''; subtitle='Forward-looking price & trend forecasts'; rel='USA\USA_Stock_Predictions.docx' }
    @{ id='usa_analyst_ratings'; title='USA Analyst Ratings'; region='USA'; emoji=''; subtitle='Analyst ratings, price targets & recommendations'; rel='USA\USA_Analyst_Ratings.docx' }
)

# Prefix used in the portal for "Copy .docx path" buttons and footer.
# Using ../Analysis so the whole Analysis folder (with prompts) can be moved frequently.
$sourcePrefix = '../Analysis'

Write-Host "Extracting prompts from ../Analysis (using robust XML parser)..." -ForegroundColor Cyan
$promptTexts = @{}
foreach ($e in $entries) {
    $p = Join-Path $analysisRoot $e.rel
    $txt = Get-DocxPlainText $p
    $promptTexts[$e.id] = $txt

    $preview = if ($txt.Length -gt 0) { 
        ($txt.Substring(0, [Math]::Min(90, $txt.Length)) -replace "`r?`n", " ").Trim() + "..."
    } else { "[empty]" }

    $status = if ($txt.StartsWith("[")) { "ERROR" } else { "OK" }
    Write-Host ("  [{0}] {1,-28} {2,6} chars  | {3}" -f $status, $e.title, $txt.Length, $preview) -ForegroundColor $(if ($status -eq "OK") { "Green" } else { "Red" })
}

# Summary of any extraction problems
$failed = $entries | Where-Object { $promptTexts[$_.id].StartsWith("[") }
if ($failed.Count -gt 0) {
    Write-Host ""
    Write-Host "⚠️  EXTRACTION ISSUES DETECTED for $($failed.Count) file(s):" -ForegroundColor Yellow
    foreach ($f in $failed) {
        Write-Host "   - $($f.title)" -ForegroundColor Yellow
    }
    Write-Host ""
    Write-Host "Most common cause: The .docx file is open in Microsoft Word or locked by OneDrive." -ForegroundColor Yellow
    Write-Host "Fix: Close the file(s) completely in Word, wait 3-5 seconds, then run this script again." -ForegroundColor Yellow
    Write-Host ""
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
