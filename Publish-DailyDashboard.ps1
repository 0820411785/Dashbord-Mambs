param(
    [Parameter(Mandatory = $true)]
    [string]$InputPath,

    [Parameter(Mandatory = $true)]
    [string]$Date,

    [string]$Title = "",
    [string]$PortalRoot = ".\Daily_Dashboards"
)

$ErrorActionPreference = "Stop"

function ConvertTo-HtmlText {
    param($Value)
    return [System.Net.WebUtility]::HtmlEncode("$Value")
}

function Convert-ToNumber {
    param($Value)
    if ($null -eq $Value) { return 0.0 }
    if ($Value -is [double] -or $Value -is [int] -or $Value -is [decimal]) { return [double]$Value }
    $text = "$Value".Trim().Replace(",", ".")
    $number = 0.0
    if ([double]::TryParse($text, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$number)) {
        return $number
    }
    return 0.0
}

function Convert-ToOptionalNumber {
    param($Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [double] -or $Value -is [int] -or $Value -is [decimal]) { return [double]$Value }
    $text = "$Value".Trim().Replace(",", ".")
    if ([string]::IsNullOrWhiteSpace($text)) { return $null }

    $number = 0.0
    if ([double]::TryParse($text, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$number)) {
        return $number
    }
    return $null
}

function Get-DashboardDtpt {
    param(
        [string]$PortalPath,
        $Entry
    )

    if ($null -eq $Entry) { return $null }
    $dashboardPath = Join-Path $PortalPath $Entry.Path
    if (-not (Test-Path -LiteralPath $dashboardPath -PathType Leaf)) { return $null }

    $html = Get-Content -LiteralPath $dashboardPath -Raw
    $match = [regex]::Match($html, "<span>\s*DTPT de la journee\s*</span>\s*<strong>([^<]*)</strong>", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if (-not $match.Success) { return $null }

    $valueText = [System.Net.WebUtility]::HtmlDecode($match.Groups.Item(1).Value).Trim()
    return Convert-ToOptionalNumber $valueText
}

function Get-SiteTotals {
    param(
        [string]$PortalPath,
        [array]$Entries
    )

    $siteTotals = @{}
    foreach ($entry in @($Entries)) {
        if ($null -eq $entry) { continue }

        $dashboardFolder = Split-Path $entry.Path -Parent
        if ([string]::IsNullOrWhiteSpace($dashboardFolder)) { continue }

        $siteCsvPath = Join-Path (Join-Path $PortalPath $dashboardFolder) "index_BySite.csv"
        if (-not (Test-Path -LiteralPath $siteCsvPath -PathType Leaf)) { continue }

        foreach ($site in @(Import-Csv -Path $siteCsvPath)) {
            $siteId = "$($site.'Site ID')".Trim()
            if ([string]::IsNullOrWhiteSpace($siteId)) { continue }

            if (-not $siteTotals.ContainsKey($siteId)) {
                $siteTotals[$siteId] = [PSCustomObject]@{
                    SiteId = $siteId
                    SiteName = "$($site.'Site Name')".Trim()
                    Incidents = 0
                    DownHours = 0.0
                    DaysImpacted = 0
                    LastSeen = $entry.Date
                }
            }

            $item = $siteTotals[$siteId]
            if ([string]::IsNullOrWhiteSpace($item.SiteName) -and -not [string]::IsNullOrWhiteSpace("$($site.'Site Name')".Trim())) {
                $item.SiteName = "$($site.'Site Name')".Trim()
            }
            $item.Incidents += [int](Convert-ToNumber $site.Incidents)
            $item.DownHours = [Math]::Round($item.DownHours + (Convert-ToNumber $site.'Down Hours'), 2)
            $item.DaysImpacted += 1
            if ($entry.Date -gt $item.LastSeen) { $item.LastSeen = $entry.Date }
        }
    }

    return @(
        $siteTotals.Values |
            Sort-Object @{ Expression = "DownHours"; Descending = $true }, @{ Expression = "Incidents"; Descending = $true }, SiteId
    )
}

function New-SiteRowsHtml {
    param(
        [array]$SiteRows,
        [switch]$IncludeMonthColumns
    )

    $rows = @()
    foreach ($site in @($SiteRows)) {
        $searchText = "$($site.SiteId) $($site.SiteName) $($site.Incidents) $($site.DownHours) $($site.DaysImpacted) $($site.LastSeen)"
        if ($IncludeMonthColumns) {
            $rows += "<tr data-search='$(ConvertTo-HtmlText $searchText)'><td>$(ConvertTo-HtmlText $site.SiteId)</td><td>$(ConvertTo-HtmlText $site.SiteName)</td><td>$(ConvertTo-HtmlText $site.Incidents)</td><td>$(ConvertTo-HtmlText $site.DownHours)</td><td>$(ConvertTo-HtmlText $site.DaysImpacted)</td><td>$(ConvertTo-HtmlText $site.LastSeen)</td></tr>"
        } else {
            $rows += "<tr data-search='$(ConvertTo-HtmlText $searchText)'><td>$(ConvertTo-HtmlText $site.SiteId)</td><td>$(ConvertTo-HtmlText $site.SiteName)</td><td>$(ConvertTo-HtmlText $site.DownHours)</td></tr>"
        }
    }

    if ($rows.Count -eq 0) {
        $colspan = if ($IncludeMonthColumns) { 6 } else { 4 }
        $rows += "<tr><td colspan='$colspan'>Aucune donnee site disponible.</td></tr>"
    }

    return $rows
}

$dateValue = [datetime]::ParseExact($Date, "yyyy-MM-dd", [System.Globalization.CultureInfo]::InvariantCulture)
$dayFolderName = $dateValue.ToString("yyyy-MM-dd")
$portalPath = [System.IO.Path]::GetFullPath($PortalRoot)
$dayPath = Join-Path $portalPath $dayFolderName

if ([string]::IsNullOrWhiteSpace($Title)) {
    $Title = "Incidents du " + $dateValue.ToString("dd/MM/yyyy")
}

New-Item -ItemType Directory -Force -Path $dayPath | Out-Null

$dashboardPath = Join-Path $dayPath "index.html"
& "$PSScriptRoot\Generate-IncidentDashboard.ps1" -InputPath $InputPath -OutputPath $dashboardPath

$manifestPath = Join-Path $portalPath "dashboards.csv"
if (-not (Test-Path -LiteralPath $manifestPath)) {
    "Date,Title,Path,Source,GeneratedAt" | Set-Content -Path $manifestPath -Encoding UTF8
}

$existing = @()
if (Test-Path -LiteralPath $manifestPath) {
    $existing = @(Import-Csv -Path $manifestPath)
}

$relativePath = "$dayFolderName/index.html"
$newEntry = [PSCustomObject]@{
    Date = $dayFolderName
    Title = $Title
    Path = $relativePath
    Source = [System.IO.Path]::GetFileName($InputPath)
    GeneratedAt = (Get-Date).ToString("yyyy-MM-dd HH:mm")
}

$updated = @($existing | Where-Object { $_.Date -ne $dayFolderName })
$updated += $newEntry
$updated = @($updated | Sort-Object Date -Descending)
$updated | Export-Csv -Path $manifestPath -NoTypeInformation -Encoding UTF8

$dashboardCount = $updated.Count
$latestEntry = if ($dashboardCount -gt 0) { $updated | Select-Object -First 1 } else { $null }
$latestLink = if ($null -ne $latestEntry) { "<a href='$(ConvertTo-HtmlText $latestEntry.Path)'>$(ConvertTo-HtmlText $latestEntry.Title)</a>" } else { "Aucun dashboard" }
$latestGenerated = if ($null -ne $latestEntry) { $latestEntry.GeneratedAt } else { "" }
$latestSource = if ($null -ne $latestEntry) { $latestEntry.Source } else { "" }
$latestDateLabel = if ($null -ne $latestEntry) {
    ([datetime]::ParseExact($latestEntry.Date, "yyyy-MM-dd", [System.Globalization.CultureInfo]::InvariantCulture)).ToString("dd/MM/yyyy")
} else {
    ""
}

$latestDateValue = if ($null -ne $latestEntry) {
    [datetime]::ParseExact($latestEntry.Date, "yyyy-MM-dd", [System.Globalization.CultureInfo]::InvariantCulture)
} else {
    $null
}
$latestDaySuffix = if ([string]::IsNullOrWhiteSpace($latestDateLabel)) { "" } else { " ($latestDateLabel)" }

# --- CALCUL DYNAMIQUE DU MOIS EN COURS (JUILLET) ---
$currentMonthEntries = @()
$currentMonthLabel = "juillet 2026"
if ($null -ne $latestDateValue) {
    $currentMonthLabel = $latestDateValue.ToString("MMMM yyyy", [System.Globalization.CultureInfo]::GetCultureInfo("fr-FR"))
    $currentMonthEntries = @(
        $updated | Where-Object {
            $entryDate = [datetime]::ParseExact($_.Date, "yyyy-MM-dd", [System.Globalization.CultureInfo]::InvariantCulture)
            $entryDate.Year -eq $latestDateValue.Year -and $entryDate.Month -eq $latestDateValue.Month
        }
    )
}

# --- MODIFICATION ICI : On génère uniquement les lignes HTML pour le mois actif sur le portail principal ---
$rows = @()
foreach ($entry in $currentMonthEntries) {
    $searchText = "$($entry.Date) $($entry.Title) $($entry.Source) $($entry.GeneratedAt)"
    $rows += "<tr data-search='$(ConvertTo-HtmlText $searchText)'><td>$(ConvertTo-HtmlText $entry.Date)</td><td><a href='$(ConvertTo-HtmlText $entry.Path)'>$(ConvertTo-HtmlText $entry.Title)</a></td><td>$(ConvertTo-HtmlText $entry.Source)</td><td>$(ConvertTo-HtmlText $entry.GeneratedAt)</td></tr>"
}

$dtptValues = @(
    foreach ($entry in $currentMonthEntries) {
        $value = Get-DashboardDtpt $portalPath $entry
        if ($null -ne $value) { $value }
    }
)
$monthlyDdtpAverage = if ($dtptValues.Count -gt 0) {
    [Math]::Round((($dtptValues | Measure-Object -Average).Average), 2)
} else {
    ""
}
$monthlyDdtpDays = $dtptValues.Count
$monthlyDdtpSuffix = if ([string]::IsNullOrWhiteSpace($currentMonthLabel)) { "" } else { " ($currentMonthLabel)" }

# --- CUMUL SITES DU MOIS EN COURS (JUILLET) ---
$currentMonthSiteTotalsRows = Get-SiteTotals $portalPath $currentMonthEntries
$currentMonthSiteRows = New-SiteRowsHtml $currentMonthSiteTotalsRows -IncludeMonthColumns
$currentMonthSiteCount = $currentMonthSiteTotalsRows.Count
$currentMonthTotalDown = [Math]::Round((@($currentMonthSiteTotalsRows) | Measure-Object -Property DownHours -Sum).Sum, 2)

$latestSiteTotalsRows = Get-SiteTotals $portalPath @($latestEntry)
$latestSiteRows = New-SiteRowsHtml $latestSiteTotalsRows
$latestSiteCount = $latestSiteTotalsRows.Count
$totalDownForLatestDay = [Math]::Round((@($latestSiteTotalsRows) | Measure-Object -Property DownHours -Sum).Sum, 2)

# --- ARCHIVES : GENERATION DE LA PAGE COMPLEMENTAIRE JUIN ---
$pastMonthValue = if ($null -ne $latestDateValue) { $latestDateValue.AddMonths(-1) } else { $null }
$pastMonthLabel = if ($null -ne $pastMonthValue) { $pastMonthValue.ToString("MMMM yyyy", [System.Globalization.CultureInfo]::GetCultureInfo("fr-FR")) } else { "juin 2026" }

$pastEntries = @()
if ($null -ne $pastMonthValue) {
    $pastEntries = @(
        $updated | Where-Object {
            $entryDate = [datetime]::ParseExact($_.Date, "yyyy-MM-dd", [System.Globalization.CultureInfo]::InvariantCulture)
            $entryDate.Year -eq $pastMonthValue.Year -and $entryDate.Month -eq $pastMonthValue.Month
        }
    )
}
$pastSiteTotalsRows = Get-SiteTotals $portalPath $pastEntries
$pastSiteRows = New-SiteRowsHtml $pastSiteTotalsRows -IncludeMonthColumns
$pastSiteCount = $pastSiteTotalsRows.Count

# Generation du fichier dedie pour l'archive
$archiveHtmlFile = "archive_" + ($pastMonthValue.ToString("yyyy_MM")) + ".html"
$archiveHtmlPath = Join-Path $portalPath $archiveHtmlFile

$archivePageHtml = @"
<!doctype html>
<html lang="fr">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Archives - $pastMonthLabel</title>
<style>
:root { color-scheme: light; --ink: #172033; --muted: #5b6678; --line: #d9e1ec; --panel: #fff; --page: #f5f7fb; --brand: #475569; --brand-dark: #334155; }
body { margin: 0; font-family: Arial, Helvetica, sans-serif; color: var(--ink); background: var(--page); }
header { background: linear-gradient(135deg, var(--brand-dark), var(--brand)); color: white; padding: 26px 34px; position: relative; }
h1 { margin: 0 0 6px; font-size: 24px; }
h2 { margin: 20px 0 14px; font-size: 18px; }
main { padding: 24px 34px 42px; }
.back-btn { position: absolute; right: 34px; top: 26px; background: white; color: var(--brand-dark); border: none; padding: 10px 16px; border-radius: 6px; font-weight: bold; text-decoration: none; font-size: 14px; }
.toolbar { display: grid; grid-template-columns: minmax(220px, 1fr) auto; gap: 10px; margin-bottom: 12px; }
input, button { border: 1px solid var(--line); border-radius: 6px; color: var(--ink); font: inherit; min-height: 40px; padding: 8px 10px; background: white; }
button { cursor: pointer; font-weight: 700; color: var(--brand-dark); }
.panel { background: var(--panel); border: 1px solid var(--line); border-radius: 6px; overflow-x: auto; }
.count { color: var(--muted); font-size: 12px; margin: 0 0 10px; }
table { width: 100%; border-collapse: collapse; font-size: 14px; }
th { text-align: left; background: #e2e8f0; color: var(--ink); padding: 12px 14px; }
td { border-bottom: 1px solid var(--line); padding: 12px 14px; }
tbody tr:nth-child(even) { background: #f8fafc; }
</style>
</head>
<body>
<header>
<h1>Archives Historiques - Cumul $pastMonthLabel</h1>
<a href="index.html" class="back-btn"><- Retour au Portail</a>
</header>
<main>
<h2>Synthese des outages par site ($pastMonthLabel)</h2>
<div class="toolbar">
<input id="past-site-search" type="search" placeholder="Rechercher un site, un nom...">
<button id="past-site-clear" type="button">Reinitialiser</button>
</div>
<p id="past-site-count" class="count">$pastSiteCount / $pastSiteCount</p>
<div class="panel">
<table id="past-site-totals-table">
<thead><tr><th>Site ID</th><th>Nom du site</th><th>Incidents cumules</th><th>Down cumule (h)</th><th>Jours impactes</th><th>Derniere journee</th></tr></thead>
<tbody>
$($pastSiteRows -join "`n")
</tbody>
</table>
</div>
</main>
<script>
(function () {
    function normalize(value) { return (value || '').toString().toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, ''); }
    var table = document.getElementById('past-site-totals-table');
    var searchInput = document.getElementById('past-site-search');
    var clearButton = document.getElementById('past-site-clear');
    var countLabel = document.getElementById('past-site-count');
    var rows = Array.prototype.slice.call(table.tBodies[0].rows);

    function applySearch() {
        var query = normalize(searchInput.value);
        var visible = 0;
        rows.forEach(function (row) {
            var show = normalize(row.dataset.search).indexOf(query) !== -1;
            row.hidden = !show;
            if (show) { visible += 1; }
        });
        countLabel.textContent = visible + ' / ' + rows.length;
    }
    searchInput.addEventListener('input', applySearch);
    clearButton.addEventListener('click', function () { searchInput.value = ''; applySearch(); searchInput.focus(); });
}());
</script>
</body>
</html>
"@
[System.IO.File]::WriteAllText($archiveHtmlPath, $archivePageHtml, [System.Text.Encoding]::UTF8)

# --- PAGE DU PORTAIL PRINCIPAL (INDEX) ---
$portalHtml = @"
<!doctype html>
<html lang="fr">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Portail des dashboards incidents</title>
<style>
:root { color-scheme: light; --ink: #172033; --muted: #5b6678; --line: #d9e1ec; --panel: #fff; --page: #f5f7fb; --brand: #0f766e; --brand-dark: #115e59; }
* { box-sizing: border-box; }
body { margin: 0; font-family: Arial, Helvetica, sans-serif; color: var(--ink); background: var(--page); }
header { background: linear-gradient(135deg, var(--brand-dark), var(--brand)); color: white; padding: 26px 34px; display: flex; justify-content: flex-start; align-items: center; position: relative; }
h1 { margin: 0 0 6px; font-size: 28px; line-height: 1.15; }
h2 { margin: 0 0 14px; font-size: 18px; color: var(--ink); }
main { padding: 24px 34px 42px; }
a { color: var(--brand-dark); font-weight: 700; text-decoration: none; }
a:hover { text-decoration: underline; }
.meta { opacity: .9; font-size: 13px; }
.kpis { display: grid; grid-template-columns: repeat(auto-fit, minmax(190px, 1fr)); gap: 12px; margin-bottom: 18px; }
.kpi { background: var(--panel); border: 1px solid var(--line); border-radius: 6px; padding: 14px; min-height: 82px; }
.kpi span { display: block; color: var(--muted); font-size: 12px; margin-bottom: 8px; }
.kpi strong { display: block; color: #111827; font-size: 22px; line-height: 1.15; overflow-wrap: anywhere; }
.toolbar { display: grid; grid-template-columns: minmax(220px, 1fr) auto; gap: 10px; margin-bottom: 12px; }
input, button { border: 1px solid var(--line); border-radius: 6px; color: var(--ink); font: inherit; min-height: 40px; padding: 8px 10px; background: white; }
button { cursor: pointer; font-weight: 700; color: var(--brand-dark); }
.panel { background: var(--panel); border: 1px solid var(--line); border-radius: 6px; overflow-x: auto; margin-bottom: 22px; }
.count { color: var(--muted); font-size: 12px; margin: 0 0 10px; }
table { width: 100%; border-collapse: collapse; font-size: 14px; }
th { text-align: left; background: #e8eef6; color: var(--ink); position: sticky; top: 0; z-index: 1; }
th, td { border-bottom: 1px solid var(--line); padding: 12px 14px; vertical-align: top; }
tbody tr:nth-child(even) { background: #f8fafc; }
.lnk-archive { position: absolute; right: 34px; background: white; color: var(--brand-dark); padding: 10px 16px; border-radius: 6px; font-weight: bold; font-size: 14px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
.lnk-archive:hover { background: #f1f5f9; text-decoration: none; }
@media (max-width: 720px) { main { padding: 18px; } header { padding: 20px; flex-direction: column; align-items: flex-start; } .lnk-archive { position: static; margin-top: 15px; } .toolbar { grid-template-columns: 1fr; } }
</style>
</head>
<body>
<header>
<div>
<h1>Portail des dashboards incidents</h1>
<div class="meta">Selectionner une journee pour consulter le dashboard correspondant.</div>
</div>
<a class="lnk-archive" href="$archiveHtmlFile" target="_blank">📋 Consulter Archives ($pastMonthLabel)</a>
</header>
<main>
<section class="kpis">
<div class="kpi"><span>Dashboards publies</span><strong>$dashboardCount</strong></div>
<div class="kpi"><span>Derniere journee</span><strong>$latestLink</strong></div>
<div class="kpi"><span>Fichier source</span><strong>$(ConvertTo-HtmlText $latestSource)</strong></div>
<div class="kpi"><span>Derniere generation</span><strong>$(ConvertTo-HtmlText $latestGenerated)</strong></div>
<div class="kpi"><span>DDTP moyen du mois$monthlyDdtpSuffix</span><strong>$monthlyDdtpAverage</strong></div>
<div class="kpi"><span>Jours DDTP suivis$monthlyDdtpSuffix</span><strong>$monthlyDdtpDays</strong></div>
<div class="kpi"><span>Sites du jour$latestDaySuffix</span><strong>$latestSiteCount</strong></div>
<div class="kpi"><span>Down du jour$latestDaySuffix</span><strong>$totalDownForLatestDay h</strong></div>
<div class="kpi"><span>Sites cumules $currentMonthLabel</span><strong>$currentMonthSiteCount</strong></div>
<div class="kpi"><span>Down cumule $currentMonthLabel</span><strong>$currentMonthTotalDown h</strong></div>
</section>

<h2>Dashboards journaliers</h2>
<div class="toolbar">
<input id="portal-search" type="search" placeholder="Rechercher une date, un fichier, un titre">
<button id="portal-clear" type="button">Reinitialiser</button>
</div>
<p id="portal-count" class="count">$($currentMonthEntries.Count) / $($currentMonthEntries.Count)</p>
<div class="panel">
<table id="dashboards-table">
<thead><tr><th>Date</th><th>Dashboard</th><th>Fichier source</th><th>Generation</th></tr></thead>
<tbody>
$($rows -join "`n")
</tbody>
</table>
</div>

<h2>Outages du jour par site$latestDaySuffix</h2>
<div class="toolbar">
<input id="site-search" type="search" placeholder="Rechercher un site, un nom, une date">
<button id="site-clear" type="button">Reinitialiser</button>
</div>
<p id="site-count" class="count">$latestSiteCount / $latestSiteCount</p>
<div class="panel">
<table id="site-totals-table">
<thead><tr><th>Site ID</th><th>Nom du site</th><th>Incidents</th><th>Down (h)</th></tr></thead>
<tbody>
$($latestSiteRows -join "`n")
</tbody>
</table>
</div>

<h2>Cumul des down par site - $currentMonthLabel</h2>
<div class="toolbar">
<input id="current-site-search" type="search" placeholder="Rechercher un site, un nom...">
<button id="current-site-clear" type="button">Reinitialiser</button>
</div>
<p id="current-site-count" class="count">$currentMonthSiteCount / $currentMonthSiteCount</p>
<div class="panel">
<table id="current-site-totals-table">
<thead><tr><th>Site ID</th><th>Nom du site</th><th>Incidents cumules</th><th>Down cumule (h)</th><th>Jours impactes</th><th>Derniere journee</th></tr></thead>
<tbody>
$($currentMonthSiteRows -join "`n")
</tbody>
</table>
</div>
</main>
<script>
(function () {
    function normalize(value) {
        return (value || '').toString().toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    }

    function wireSearch(tableId, inputId, clearId, countId) {
        var table = document.getElementById(tableId);
        if (!table) { return; }

        var searchInput = document.getElementById(inputId);
        var clearButton = document.getElementById(clearId);
        var countLabel = document.getElementById(countId);
        var rows = Array.prototype.slice.call(table.tBodies[0].rows);

        function applySearch() {
            var query = normalize(searchInput.value);
            var visible = 0;
            rows.forEach(function (row) {
                var show = normalize(row.dataset.search).indexOf(query) !== -1;
                row.hidden = !show;
                if (show) { visible += 1; }
            });
            countLabel.textContent = visible + ' / ' + rows.length;
        }

        searchInput.addEventListener('input', applySearch);
        clearButton.addEventListener('click', function () {
            searchInput.value = '';
            applySearch();
            searchInput.focus();
        });
    }

    wireSearch('dashboards-table', 'portal-search', 'portal-clear', 'portal-count');
    wireSearch('site-totals-table', 'site-search', 'site-clear', 'site-count');
    wireSearch('current-site-totals-table', 'current-site-search', 'current-site-clear', 'current-site-count');
}());
</script>
</body>
</html>
"@

[System.IO.File]::WriteAllText((Join-Path $portalPath "index.html"), $portalHtml, [System.Text.Encoding]::UTF8)

Write-Output "Dashboard journalier cree: $dashboardPath"
Write-Output "Portail mis a jour: $(Join-Path $portalPath "index.html")"