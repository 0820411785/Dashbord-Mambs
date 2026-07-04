param(
    [string]$InputPath = ".\To Brice_Template Incident_12 (2).xlsb",
    [string]$OutputPath = "",
    [string]$SheetName = "'19-May Outage List (2)$'",
    [switch]$CreateExcel
)

$ErrorActionPreference = "Stop"

function Convert-ToNumber {
    param($Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [double] -or $Value -is [int] -or $Value -is [decimal]) { return [double]$Value }
    $text = "$Value".Trim().Replace(",", ".")
    $number = 0.0
    if ([double]::TryParse($text, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$number)) {
        return $number
    }
    return $null
}

function Convert-ToDateValue {
    param($Value)
    if ($null -eq $Value -or "$Value".Trim() -eq "") { return $null }
    if ($Value -is [datetime]) { return $Value }

    $text = "$Value".Trim()
    $cultures = @(
        [System.Globalization.CultureInfo]::GetCultureInfo("fr-FR"),
        [System.Globalization.CultureInfo]::GetCultureInfo("en-US"),
        [System.Globalization.CultureInfo]::InvariantCulture
    )

    foreach ($culture in $cultures) {
        $date = [datetime]::MinValue
        if ([datetime]::TryParse($text, $culture, [System.Globalization.DateTimeStyles]::AssumeLocal, [ref]$date)) {
            return $date
        }
    }
    return $null
}

function Convert-DurationToHours {
    param($Value)
    if ($null -eq $Value -or "$Value".Trim() -eq "") { return 0.0 }
    if ($Value -is [datetime]) { return $Value.TimeOfDay.TotalHours }
    if ($Value -is [timespan]) { return $Value.TotalHours }
    if ($Value -is [double] -or $Value -is [decimal] -or $Value -is [int]) {
        # Excel stores durations as fractions of one day.
        return [double]$Value * 24.0
    }

    $text = "$Value".Trim()
    $parts = $text.Split(":")
    if ($parts.Count -ge 2) {
        $hours = Convert-ToNumber $parts[0]
        $minutes = Convert-ToNumber $parts[1]
        $seconds = if ($parts.Count -ge 3) { Convert-ToNumber $parts[2] } else { 0 }
        if ($null -ne $hours -and $null -ne $minutes -and $null -ne $seconds) {
            return $hours + ($minutes / 60.0) + ($seconds / 3600.0)
        }
    }
    $number = Convert-ToNumber $text
    if ($null -ne $number) { return $number }
    return 0.0
}

function Add-Worksheet {
    param($Workbook, [string]$Name)
    $sheet = $Workbook.Worksheets.Add()
    $sheet.Name = $Name
    return $sheet
}

function Write-Table {
    param(
        $Sheet,
        [int]$Row,
        [int]$Column,
        [array]$Headers,
        [array]$Rows
    )

    for ($i = 0; $i -lt $Headers.Count; $i++) {
        $Sheet.Cells.Item($Row, $Column + $i).Value2 = $Headers[$i]
        $Sheet.Cells.Item($Row, $Column + $i).Font.Bold = $true
        $Sheet.Cells.Item($Row, $Column + $i).Interior.Color = 15773696
    }

    $currentRow = $Row + 1
    foreach ($item in $Rows) {
        for ($i = 0; $i -lt $Headers.Count; $i++) {
            $value = $item.($Headers[$i])
            if ($value -is [datetime]) {
                $Sheet.Cells.Item($currentRow, $Column + $i).Value2 = $value
                $Sheet.Cells.Item($currentRow, $Column + $i).NumberFormat = "dd/mm/yyyy hh:mm"
            } else {
                $Sheet.Cells.Item($currentRow, $Column + $i).Value2 = $value
            }
        }
        $currentRow++
    }
    $range = $Sheet.Range($Sheet.Cells.Item($Row, $Column), $Sheet.Cells.Item([Math]::Max($currentRow - 1, $Row), $Column + $Headers.Count - 1))
    $range.Borders.LineStyle = 1
    $Sheet.Columns.AutoFit() | Out-Null
}

function Add-Chart {
    param(
        $Sheet,
        [string]$Title,
        [int]$Left,
        [int]$Top,
        $SourceRange,
        [int]$ChartType = 51
    )

    $chartObject = $Sheet.ChartObjects().Add($Left, $Top, 420, 240)
    $chart = $chartObject.Chart
    $chart.ChartType = $ChartType
    $chart.SetSourceData($SourceRange)
    $chart.HasTitle = $true
    $chart.ChartTitle.Text = $Title
    $chart.HasLegend = $false
    return $chartObject
}

function ConvertTo-HtmlText {
    param($Value)
    return [System.Net.WebUtility]::HtmlEncode("$Value")
}

function Format-HtmlValue {
    param($Value)
    if ($null -eq $Value) { return "" }
    if ($Value -is [datetime]) {
        return $Value.ToString("dd/MM/yyyy HH:mm", [System.Globalization.CultureInfo]::GetCultureInfo("fr-FR"))
    }
    return "$Value"
}

function New-BarChartSvg {
    param(
        [array]$Rows,
        [string]$LabelProperty,
        [string]$ValueProperty,
        [int]$Width = 760,
        [int]$BarHeight = 28
    )

    $rowsToDraw = @($Rows | Select-Object -First 10)
    if ($rowsToDraw.Count -eq 0) { return "<p>Aucune donnee</p>" }

    $max = ($rowsToDraw | Measure-Object -Property $ValueProperty -Maximum).Maximum
    if ($null -eq $max -or $max -le 0) { $max = 1 }

    $height = 44 + ($rowsToDraw.Count * ($BarHeight + 10))
    $plotLeft = 170
    $plotWidth = $Width - $plotLeft - 90
    $svg = @()
    $svg += "<svg viewBox='0 0 $Width $height' role='img'>"
    $y = 24
    foreach ($row in $rowsToDraw) {
        $label = ConvertTo-HtmlText $row.$LabelProperty
        $value = [double]$row.$ValueProperty
        $barWidth = [Math]::Max(2, [Math]::Round(($value / $max) * $plotWidth, 0))
        $displayValue = ConvertTo-HtmlText ([Math]::Round($value, 2))
        $svg += "<text x='0' y='$($y + 19)' class='axis-label'>$label</text>"
        $svg += "<rect x='$plotLeft' y='$y' width='$barWidth' height='$BarHeight' rx='3' class='bar'></rect>"
        $svg += "<text x='$($plotLeft + $barWidth + 8)' y='$($y + 19)' class='value-label'>$displayValue</text>"
        $y += $BarHeight + 10
    }
    $svg += "</svg>"
    return ($svg -join "`n")
}

function New-PieSvg {
    param(
        [array]$Rows,
        [string]$LabelProperty,
        [string]$ValueProperty
    )

    $total = ($Rows | Measure-Object -Property $ValueProperty -Sum).Sum
    if ($null -eq $total -or $total -le 0) { return "<p>Aucune donnee</p>" }

    $colors = @("#2563eb", "#f59e0b", "#10b981", "#ef4444", "#7c3aed")
    $items = @()
    $index = 0
    foreach ($row in $Rows) {
        $percent = [Math]::Round(([double]$row.$ValueProperty / $total) * 100, 1)
        $label = ConvertTo-HtmlText $row.$LabelProperty
        $value = ConvertTo-HtmlText $row.$ValueProperty
        $color = $colors[$index % $colors.Count]
        $items += "<div class='legend-row'><span class='swatch' style='background:$color'></span><span>$label</span><strong>$value ($percent%)</strong></div>"
        $index++
    }
    return "<div class='legend-only'>$($items -join '')</div>"
}

function New-HtmlTable {
    param(
        [array]$Rows,
        [array]$Headers,
        [int]$Limit = 20,
        [string]$Id = "",
        [string]$CssClass = ""
    )

    $headerLabels = @{
        "#" = "#"
        "Outage Date" = "Date outage"
        "Site ID" = "Site ID"
        "Site Name" = "Nom du site"
        "Outage Start" = "Debut outage"
        "Outage End" = "Fin outage"
        "Alarm Before Down (h)" = "Alarme avant down (h)"
        "Outage Duration (h)" = "Duree outage (h)"
        "ETA (min)" = "ETA (min)"
        "Passive/Active" = "Passive/Active"
        "Fuel Outage" = "Fuel outage"
        "TRB Ref" = "Reference TRB"
        "RCA" = "RCA"
        "Resolution Comment" = "Commentaire resolution"
        "Incidents" = "Incidents"
        "Down Hours" = "Duree down (h)"
        "Hour" = "Heure"
        "Type" = "Type"
    }

    $attributes = @()
    if (-not [string]::IsNullOrWhiteSpace($Id)) { $attributes += "id='$(ConvertTo-HtmlText $Id)'" }
    if (-not [string]::IsNullOrWhiteSpace($CssClass)) { $attributes += "class='$(ConvertTo-HtmlText $CssClass)'" }
    $attributeText = if ($attributes.Count -gt 0) { " " + ($attributes -join " ") } else { "" }

    $html = @()
    $html += "<table$attributeText><thead><tr>"
    foreach ($header in $Headers) {
        $label = if ($headerLabels.ContainsKey($header)) { $headerLabels[$header] } else { $header }
        $html += "<th>$(ConvertTo-HtmlText $label)</th>"
    }
    $html += "</tr></thead><tbody>"
    $rowsToRender = if ($Limit -le 0) { @($Rows) } else { @($Rows | Select-Object -First $Limit) }
    foreach ($row in $rowsToRender) {
        $rowAttributes = @()
        if ($Id -eq "incidents-table") {
            $searchText = @($Headers | ForEach-Object { Format-HtmlValue $row.$_ }) -join " "
            $duration = Convert-ToNumber $row."Outage Duration (h)"
            $durationText = if ($null -eq $duration) { "" } else { $duration.ToString([System.Globalization.CultureInfo]::InvariantCulture) }
            $fuelValue = $row."Fuel Outage"
            $paValue = $row."Passive/Active"
            $rowAttributes += "data-search='$(ConvertTo-HtmlText $searchText)'"
            $rowAttributes += "data-fuel='$(ConvertTo-HtmlText $fuelValue)'"
            $rowAttributes += "data-pa='$(ConvertTo-HtmlText $paValue)'"
            $rowAttributes += "data-duration='$durationText'"
        }
        $rowAttributeText = if ($rowAttributes.Count -gt 0) { " " + ($rowAttributes -join " ") } else { "" }
        $html += "<tr$rowAttributeText>"
        foreach ($header in $Headers) {
            $html += "<td>$(ConvertTo-HtmlText (Format-HtmlValue $row.$header))</td>"
        }
        $html += "</tr>"
    }
    $html += "</tbody></table>"
    return ($html -join "")
}

function Get-RcaKeywords {
    param([array]$Rows)

    $stopWords = @(
        "avec", "dans", "pour", "suite", "niveau", "site", "tenant", "orange",
        "manuel", "prise", "charge", "caus", "cause", "des", "les", "une",
        "sur", "par", "est", "aux", "du", "de", "la", "le", "un", "en", "au",
        "et", "a", "l", "d", "s"
    )
    $counts = @{}
    foreach ($row in $Rows) {
        $text = "$($row.RCA)".ToLowerInvariant()
        $words = [regex]::Matches($text, "\p{L}[\p{L}\p{Nd}]{2,}") | ForEach-Object { $_.Value }
        foreach ($word in $words) {
            if ($stopWords -contains $word) { continue }
            if (-not $counts.ContainsKey($word)) { $counts[$word] = 0 }
            $counts[$word]++
        }
    }

    return @(
        $counts.GetEnumerator() |
            Sort-Object Value -Descending |
            Select-Object -First 10 |
            ForEach-Object {
                [PSCustomObject]@{
                    "Mot cle" = $_.Key
                    "Occurrences" = $_.Value
                }
            }
    )
}

$resolvedInput = (Resolve-Path -LiteralPath $InputPath).Path
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $stamp = Get-Date -Format "yyyyMMdd_HHmm"
    $extension = if ($CreateExcel) { "xlsx" } else { "html" }
    $OutputPath = Join-Path (Split-Path $resolvedInput -Parent) "Incident_Dashboard_$stamp.$extension"
}
$resolvedOutput = [System.IO.Path]::GetFullPath($OutputPath)

$connection = New-Object -ComObject ADODB.Connection
$connection.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$resolvedInput;Extended Properties='Excel 12.0;HDR=NO;IMEX=1';"
$connection.Open()

$incidents = @()
$summaryFromTemplate = [ordered]@{}

try {
    $recordset = New-Object -ComObject ADODB.Recordset
    $recordset.Open("SELECT * FROM [$SheetName]", $connection)
    $rowNumber = 0

    while (-not $recordset.EOF) {
        $rowNumber++

        if ($rowNumber -eq 2) {
            $summaryFromTemplate["Open Non Fuel"] = Convert-ToNumber $recordset.Fields.Item(6).Value
            $summaryFromTemplate["Open Fuel"] = Convert-ToNumber $recordset.Fields.Item(7).Value
            $summaryFromTemplate["Open Total"] = Convert-ToNumber $recordset.Fields.Item(8).Value
            $summaryFromTemplate["DTPT journee"] = Convert-ToNumber $recordset.Fields.Item(11).Value
        }

        if ($rowNumber -ge 14) {
            $number = Convert-ToNumber $recordset.Fields.Item(1).Value
            $siteId = "$($recordset.Fields.Item(3).Value)".Trim()
            $siteName = "$($recordset.Fields.Item(4).Value)".Trim()

            if ($null -ne $number -and $siteId -ne "") {
                $outageStart = Convert-ToDateValue $recordset.Fields.Item(5).Value
                $outageEnd = Convert-ToDateValue $recordset.Fields.Item(6).Value
                $durationHours = Convert-DurationToHours $recordset.Fields.Item(8).Value
                $alarmBeforeHours = Convert-DurationToHours $recordset.Fields.Item(7).Value
                $etaMinutes = Convert-ToNumber $recordset.Fields.Item(10).Value
                $fuelText = "$($recordset.Fields.Item(14).Value)".Trim()

                $incidents += [PSCustomObject]@{
                    "#" = [int]$number
                    "Outage Date" = "$($recordset.Fields.Item(2).Value)".Trim()
                    "Site ID" = $siteId
                    "Site Name" = $siteName
                    "Outage Start" = $outageStart
                    "Outage End" = $outageEnd
                    "Alarm Before Down (h)" = [Math]::Round($alarmBeforeHours, 2)
                    "Outage Duration (h)" = [Math]::Round($durationHours, 2)
                    "ETA (min)" = if ($null -eq $etaMinutes) { "" } else { [Math]::Round($etaMinutes, 0) }
                    "Passive/Active" = "$($recordset.Fields.Item(11).Value)".Trim()
                    "RCA" = "$($recordset.Fields.Item(12).Value)".Trim()
                    "Resolution Comment" = "$($recordset.Fields.Item(13).Value)".Trim()
                    "Fuel Outage" = if ($fuelText -match "^(yes|y|oui)$") { "Oui" } elseif ($fuelText -match "^(no|n|non)$") { "Non" } else { $fuelText }
                    "TRB Ref" = "$($recordset.Fields.Item(17).Value)".Trim()
                }
            }
        }

        $recordset.MoveNext()
    }
    $recordset.Close()
}
finally {
    $connection.Close()
}

if ($incidents.Count -eq 0) {
    throw "Aucun incident exploitable trouve dans la feuille $SheetName."
}

$totalIncidents = $incidents.Count
$uniqueSites = @($incidents | Select-Object -ExpandProperty "Site ID" -Unique).Count
$totalDownHours = [Math]::Round(($incidents | Measure-Object -Property "Outage Duration (h)" -Sum).Sum, 2)
$avgDownHours = [Math]::Round(($incidents | Measure-Object -Property "Outage Duration (h)" -Average).Average, 2)
$maxDown = $incidents | Sort-Object "Outage Duration (h)" -Descending | Select-Object -First 1
$fuelCount = @($incidents | Where-Object { $_."Fuel Outage" -eq "Oui" }).Count
$nonFuelCount = @($incidents | Where-Object { $_."Fuel Outage" -eq "Non" }).Count

$bySite = @(
    $incidents |
        Group-Object "Site ID" |
        ForEach-Object {
            $items = $_.Group
            [PSCustomObject]@{
                "Site ID" = $_.Name
                "Site Name" = ($items | Select-Object -First 1)."Site Name"
                "Incidents" = $items.Count
                "Down Hours" = [Math]::Round(($items | Measure-Object -Property "Outage Duration (h)" -Sum).Sum, 2)
            }
        } |
        Sort-Object { [double]$_."Down Hours" } -Descending
)

$byFuel = @(
    $incidents |
        Group-Object "Fuel Outage" |
        ForEach-Object {
            [PSCustomObject]@{
                "Type" = if ($_.Name -eq "") { "Non precise" } else { $_.Name }
                "Incidents" = $_.Count
                "Down Hours" = [Math]::Round(($_.Group | Measure-Object -Property "Outage Duration (h)" -Sum).Sum, 2)
            }
        } |
        Sort-Object { [double]$_."Incidents" } -Descending
)

$byHour = @(
    $incidents |
        Where-Object { $null -ne $_."Outage Start" } |
        Group-Object { $_."Outage Start".Hour } |
        ForEach-Object {
            [PSCustomObject]@{
                "Hour" = "{0:00}:00" -f [int]$_.Name
                "Incidents" = $_.Count
                "Down Hours" = [Math]::Round(($_.Group | Measure-Object -Property "Outage Duration (h)" -Sum).Sum, 2)
            }
        } |
        Sort-Object "Hour"
)

$byPassiveActive = @(
    $incidents |
        Group-Object "Passive/Active" |
        ForEach-Object {
            [PSCustomObject]@{
                "Passive/Active" = if ($_.Name -eq "") { "Non precise" } else { $_.Name }
                "Incidents" = $_.Count
                "Down Hours" = [Math]::Round(($_.Group | Measure-Object -Property "Outage Duration (h)" -Sum).Sum, 2)
            }
        } |
        Sort-Object { [double]$_."Incidents" } -Descending
)

$durationBuckets = @(
    [PSCustomObject]@{ "Tranche" = "< 30 min"; "Incidents" = @($incidents | Where-Object { $_."Outage Duration (h)" -lt 0.5 }).Count },
    [PSCustomObject]@{ "Tranche" = "30 min - 1 h"; "Incidents" = @($incidents | Where-Object { $_."Outage Duration (h)" -ge 0.5 -and $_."Outage Duration (h)" -lt 1 }).Count },
    [PSCustomObject]@{ "Tranche" = "1 h - 2 h"; "Incidents" = @($incidents | Where-Object { $_."Outage Duration (h)" -ge 1 -and $_."Outage Duration (h)" -lt 2 }).Count },
    [PSCustomObject]@{ "Tranche" = "2 h - 4 h"; "Incidents" = @($incidents | Where-Object { $_."Outage Duration (h)" -ge 2 -and $_."Outage Duration (h)" -lt 4 }).Count },
    [PSCustomObject]@{ "Tranche" = "> 4 h"; "Incidents" = @($incidents | Where-Object { $_."Outage Duration (h)" -ge 4 }).Count }
)

$rcaKeywords = Get-RcaKeywords $incidents

if (-not $CreateExcel) {
    $outputDirectory = Split-Path $resolvedOutput -Parent
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($resolvedOutput)
    $incidentsCsv = Join-Path $outputDirectory "$baseName`_Incidents.csv"
    $sitesCsv = Join-Path $outputDirectory "$baseName`_BySite.csv"
    $hoursCsv = Join-Path $outputDirectory "$baseName`_ByHour.csv"
    $fuelCsv = Join-Path $outputDirectory "$baseName`_Fuel.csv"
    $durationCsv = Join-Path $outputDirectory "$baseName`_Durees.csv"
    $rcaCsv = Join-Path $outputDirectory "$baseName`_RCA.csv"

    $incidents | Export-Csv -Path $incidentsCsv -NoTypeInformation -Encoding UTF8
    $bySite | Export-Csv -Path $sitesCsv -NoTypeInformation -Encoding UTF8
    $byHour | Export-Csv -Path $hoursCsv -NoTypeInformation -Encoding UTF8
    $byFuel | Export-Csv -Path $fuelCsv -NoTypeInformation -Encoding UTF8
    $durationBuckets | Export-Csv -Path $durationCsv -NoTypeInformation -Encoding UTF8
    $rcaKeywords | Export-Csv -Path $rcaCsv -NoTypeInformation -Encoding UTF8

    $siteChart = New-BarChartSvg $bySite "Site ID" "Down Hours"
    $hourChart = New-BarChartSvg $byHour "Hour" "Incidents"
    $hourDownChart = New-BarChartSvg $byHour "Hour" "Down Hours"
    $fuelChart = New-PieSvg $byFuel "Type" "Incidents"
    $paChart = New-BarChartSvg $byPassiveActive "Passive/Active" "Incidents"
    $durationChart = New-BarChartSvg $durationBuckets "Tranche" "Incidents"
    $rcaChart = New-BarChartSvg $rcaKeywords "Mot cle" "Occurrences"
    $topSitesTable = New-HtmlTable $bySite @("Site ID", "Site Name", "Incidents", "Down Hours") 10 "sites-table" "data-table"
    $incidentsTable = New-HtmlTable $incidents @("#", "Outage Date", "Site ID", "Site Name", "Outage Start", "Outage End", "Outage Duration (h)", "Passive/Active", "Fuel Outage", "TRB Ref", "RCA") 0 "incidents-table" "data-table"
    $longestDownText = "$($maxDown.'Site ID') - $($maxDown.'Site Name') ($($maxDown.'Outage Duration (h)') h)"
    $downloadLinks = @(
        "<a href='$(ConvertTo-HtmlText ([System.IO.Path]::GetFileName($incidentsCsv)))'>Incidents CSV</a>",
        "<a href='$(ConvertTo-HtmlText ([System.IO.Path]::GetFileName($sitesCsv)))'>Sites CSV</a>",
        "<a href='$(ConvertTo-HtmlText ([System.IO.Path]::GetFileName($hoursCsv)))'>Heures CSV</a>",
        "<a href='$(ConvertTo-HtmlText ([System.IO.Path]::GetFileName($fuelCsv)))'>Fuel CSV</a>",
        "<a href='$(ConvertTo-HtmlText ([System.IO.Path]::GetFileName($durationCsv)))'>Durees CSV</a>",
        "<a href='$(ConvertTo-HtmlText ([System.IO.Path]::GetFileName($rcaCsv)))'>RCA CSV</a>"
    ) -join ""

    $html = @"
<!doctype html>
<html lang="fr">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Tableau de bord incidents</title>
<style>
:root { color-scheme: light; --ink: #172033; --muted: #5b6678; --line: #d9e1ec; --panel: #fff; --page: #f5f7fb; --brand: #0f766e; --brand-dark: #115e59; --blue: #2563eb; --amber: #d97706; }
* { box-sizing: border-box; }
body { margin: 0; font-family: Arial, Helvetica, sans-serif; color: var(--ink); background: var(--page); }
header { background: linear-gradient(135deg, var(--brand-dark), var(--brand)); color: white; padding: 24px 32px; }
h1 { margin: 0 0 6px; font-size: 28px; line-height: 1.15; }
h2 { margin: 0 0 14px; font-size: 18px; color: var(--ink); }
main { padding: 24px 32px 40px; }
a { color: var(--brand-dark); font-weight: 700; text-decoration: none; }
a:hover { text-decoration: underline; }
.meta { opacity: .9; font-size: 13px; overflow-wrap: anywhere; }
.actions { display: flex; flex-wrap: wrap; gap: 8px; margin: 14px 0 22px; }
.actions a { background: white; border: 1px solid var(--line); border-radius: 6px; color: var(--brand-dark); padding: 8px 10px; font-size: 12px; }
.kpis { display: grid; grid-template-columns: repeat(auto-fit, minmax(170px, 1fr)); gap: 12px; margin-bottom: 22px; }
.kpi { background: var(--panel); border: 1px solid var(--line); border-radius: 6px; padding: 14px; min-height: 82px; }
.kpi span { display: block; font-size: 12px; color: var(--muted); margin-bottom: 8px; }
.kpi strong { display: block; font-size: 24px; line-height: 1.1; color: #111827; overflow-wrap: anywhere; }
.kpi.wide-kpi { grid-column: span 2; }
.grid { display: grid; grid-template-columns: repeat(2, minmax(280px, 1fr)); gap: 18px; }
.panel { background: var(--panel); border: 1px solid var(--line); border-radius: 6px; padding: 16px; overflow-x: auto; }
svg { width: 100%; height: auto; min-width: 420px; }
.bar { fill: var(--blue); }
.axis-label { font-size: 12px; fill: #334155; }
.value-label { font-size: 12px; fill: #0f172a; font-weight: 700; }
.toolbar { display: grid; grid-template-columns: minmax(220px, 1fr) repeat(3, minmax(130px, 170px)) auto; gap: 10px; align-items: center; margin-bottom: 12px; min-width: 760px; }
input, select, button { border: 1px solid var(--line); border-radius: 6px; color: var(--ink); font: inherit; min-height: 38px; padding: 8px 10px; background: white; }
button { cursor: pointer; font-weight: 700; color: var(--brand-dark); }
.count { color: var(--muted); font-size: 12px; margin-bottom: 10px; }
table { width: 100%; border-collapse: collapse; font-size: 12px; }
th { text-align: left; background: #e8eef6; color: var(--ink); position: sticky; top: 0; z-index: 1; }
th, td { border: 1px solid var(--line); padding: 7px 8px; vertical-align: top; }
.data-table th { cursor: pointer; user-select: none; white-space: nowrap; }
tbody tr:nth-child(even) { background: #f8fafc; }
.wide { grid-column: 1 / -1; }
.legend-only { display: grid; gap: 10px; max-width: 520px; }
.legend-row { display: grid; grid-template-columns: 18px 1fr auto; gap: 8px; align-items: center; }
.swatch { width: 14px; height: 14px; border-radius: 3px; display: inline-block; }
@media (max-width: 850px) {
    .grid { grid-template-columns: 1fr; }
    main { padding: 18px; }
    header { padding: 20px; }
    .kpi.wide-kpi { grid-column: auto; }
    svg { min-width: 360px; }
}
</style>
</head>
<body>
<header>
<h1>Tableau de bord journalier des incidents</h1>
<div class="meta">Source : $(ConvertTo-HtmlText $resolvedInput) | Genere le : $(Get-Date -Format "yyyy-MM-dd HH:mm")</div>
</header>
<main>
<nav class="actions" aria-label="Exports CSV">$downloadLinks</nav>
<section class="kpis">
<div class="kpi"><span>Total incidents</span><strong>$totalIncidents</strong></div>
<div class="kpi"><span>Sites impactes</span><strong>$uniqueSites</strong></div>
<div class="kpi"><span>Duree totale de down</span><strong>$totalDownHours h</strong></div>
<div class="kpi"><span>Duree moyenne de down</span><strong>$avgDownHours h</strong></div>
<div class="kpi"><span>Incidents Fuel</span><strong>$fuelCount</strong></div>
<div class="kpi"><span>Incidents Non Fuel</span><strong>$nonFuelCount</strong></div>
<div class="kpi"><span>DTPT de la journee</span><strong>$($summaryFromTemplate["DTPT journee"])</strong></div>
<div class="kpi wide-kpi"><span>Plus longue interruption</span><strong>$longestDownText</strong></div>
</section>
<section class="grid">
<div class="panel"><h2>Top sites par duree de down</h2>$siteChart</div>
<div class="panel"><h2>Incidents par heure</h2>$hourChart</div>
<div class="panel"><h2>Duree de down par heure</h2>$hourDownChart</div>
<div class="panel"><h2>Repartition des durees</h2>$durationChart</div>
<div class="panel"><h2>Fuel vs Non Fuel</h2>$fuelChart</div>
<div class="panel"><h2>Passive / Active</h2>$paChart</div>
<div class="panel"><h2>Top mots RCA</h2>$rcaChart</div>
<div class="panel wide"><h2>Classement des sites</h2>$topSitesTable</div>
<div class="panel wide"><h2>Details des incidents</h2>
<div class="toolbar">
<input id="incident-search" type="search" placeholder="Rechercher site, TRB, RCA">
<select id="fuel-filter"><option value="">Fuel: tous</option><option value="Oui">Oui</option><option value="Non">Non</option><option value="Non precise">Non precise</option></select>
<select id="pa-filter"><option value="">Type: tous</option><option value="Passive">Passive</option><option value="Active">Active</option><option value="Non precise">Non precise</option></select>
<select id="duration-filter"><option value="">Duree: toutes</option><option value="0.5">&gt;= 30 min</option><option value="1">&gt;= 1 h</option><option value="2">&gt;= 2 h</option><option value="4">&gt;= 4 h</option></select>
<button id="clear-filters" type="button">Reinitialiser</button>
</div>
<div id="incident-count" class="count">$totalIncidents / $totalIncidents</div>
$incidentsTable</div>
</section>
</main>
<script>
(function () {
    var incidentTable = document.getElementById('incidents-table');
    if (!incidentTable) { return; }

    var searchInput = document.getElementById('incident-search');
    var fuelFilter = document.getElementById('fuel-filter');
    var paFilter = document.getElementById('pa-filter');
    var durationFilter = document.getElementById('duration-filter');
    var clearButton = document.getElementById('clear-filters');
    var countLabel = document.getElementById('incident-count');
    var rows = Array.prototype.slice.call(incidentTable.tBodies[0].rows);

    function normalize(value) {
        return (value || '').toString().toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    }

    function applyFilters() {
        var query = normalize(searchInput.value);
        var fuel = fuelFilter.value;
        var pa = paFilter.value;
        var minDuration = parseFloat(durationFilter.value || '0');
        var visible = 0;

        rows.forEach(function (row) {
            var matchesSearch = normalize(row.dataset.search).indexOf(query) !== -1;
            var matchesFuel = !fuel || row.dataset.fuel === fuel;
            var matchesPa = !pa || row.dataset.pa === pa;
            var duration = parseFloat(row.dataset.duration || '0');
            var matchesDuration = duration >= minDuration;
            var show = matchesSearch && matchesFuel && matchesPa && matchesDuration;
            row.hidden = !show;
            if (show) { visible += 1; }
        });

        countLabel.textContent = visible + ' / ' + rows.length;
    }

    [searchInput, fuelFilter, paFilter, durationFilter].forEach(function (control) {
        control.addEventListener('input', applyFilters);
        control.addEventListener('change', applyFilters);
    });

    clearButton.addEventListener('click', function () {
        searchInput.value = '';
        fuelFilter.value = '';
        paFilter.value = '';
        durationFilter.value = '';
        applyFilters();
        searchInput.focus();
    });

    document.querySelectorAll('.data-table').forEach(function (table) {
        Array.prototype.slice.call(table.tHead.rows[0].cells).forEach(function (header, index) {
            header.title = 'Trier';
            header.addEventListener('click', function () {
                var body = table.tBodies[0];
                var tableRows = Array.prototype.slice.call(body.rows);
                var direction = header.dataset.direction === 'asc' ? -1 : 1;
                tableRows.sort(function (a, b) {
                    var av = a.cells[index].textContent.trim();
                    var bv = b.cells[index].textContent.trim();
                    var an = parseFloat(av.replace(',', '.'));
                    var bn = parseFloat(bv.replace(',', '.'));
                    if (!isNaN(an) && !isNaN(bn)) { return (an - bn) * direction; }
                    return av.localeCompare(bv, 'fr', { numeric: true, sensitivity: 'base' }) * direction;
                });
                tableRows.forEach(function (row) { body.appendChild(row); });
                header.dataset.direction = direction === 1 ? 'asc' : 'desc';
            });
        });
    });
}());
</script>
</body>
</html>
"@

    Set-Content -Path $resolvedOutput -Value $html -Encoding UTF8

    Write-Output "Tableau de bord cree: $resolvedOutput"
    Write-Output "CSV cree: $incidentsCsv"
    Write-Output "CSV cree: $sitesCsv"
    Write-Output "CSV cree: $hoursCsv"
    Write-Output "CSV cree: $fuelCsv"
    Write-Output "CSV cree: $durationCsv"
    Write-Output "CSV cree: $rcaCsv"
    Write-Output "Incidents: $totalIncidents"
    Write-Output "Sites impactes: $uniqueSites"
    Write-Output "Duree totale de down: $totalDownHours"
    return
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$workbook = $excel.Workbooks.Add()

try {
    while ($workbook.Worksheets.Count -gt 1) {
        $workbook.Worksheets.Item(1).Delete()
    }

    $dashboard = $workbook.Worksheets.Item(1)
    $dashboard.Name = "Dashboard"
    $incidentsSheet = Add-Worksheet $workbook "Incidents"
    $siteSheet = Add-Worksheet $workbook "By Site"
    $hourSheet = Add-Worksheet $workbook "By Hour"
    $fuelSheet = Add-Worksheet $workbook "Fuel"
    $paSheet = Add-Worksheet $workbook "Passive Active"

    $dashboard.Cells.Item(1, 1).Value2 = "Daily Incident Dashboard"
    $dashboard.Cells.Item(1, 1).Font.Bold = $true
    $dashboard.Cells.Item(1, 1).Font.Size = 20
    $dashboard.Cells.Item(2, 1).Value2 = "Source"
    $dashboard.Cells.Item(2, 2).Value2 = $resolvedInput
    $dashboard.Cells.Item(3, 1).Value2 = "Generated"
    $dashboard.Cells.Item(3, 2).Value2 = (Get-Date).ToString("yyyy-MM-dd HH:mm")

    $kpis = @(
        [PSCustomObject]@{ "KPI" = "Total incidents"; "Value" = $totalIncidents },
        [PSCustomObject]@{ "KPI" = "Sites impacted"; "Value" = $uniqueSites },
        [PSCustomObject]@{ "KPI" = "Total down hours"; "Value" = $totalDownHours },
        [PSCustomObject]@{ "KPI" = "Average down hours"; "Value" = $avgDownHours },
        [PSCustomObject]@{ "KPI" = "Longest down site"; "Value" = "$($maxDown.'Site ID') - $($maxDown.'Site Name') ($($maxDown.'Outage Duration (h)') h)" },
        [PSCustomObject]@{ "KPI" = "Fuel incidents"; "Value" = $fuelCount },
        [PSCustomObject]@{ "KPI" = "Non fuel incidents"; "Value" = $nonFuelCount },
        [PSCustomObject]@{ "KPI" = "Template DTPT journee"; "Value" = $summaryFromTemplate["DTPT journee"] }
    )
    Write-Table $dashboard 5 1 @("KPI", "Value") $kpis

    Write-Table $incidentsSheet 1 1 @("#", "Outage Date", "Site ID", "Site Name", "Outage Start", "Outage End", "Alarm Before Down (h)", "Outage Duration (h)", "ETA (min)", "Passive/Active", "RCA", "Resolution Comment", "Fuel Outage", "TRB Ref") $incidents
    Write-Table $siteSheet 1 1 @("Site ID", "Site Name", "Incidents", "Down Hours") $bySite
    Write-Table $hourSheet 1 1 @("Hour", "Incidents", "Down Hours") $byHour
    Write-Table $fuelSheet 1 1 @("Type", "Incidents", "Down Hours") $byFuel
    Write-Table $paSheet 1 1 @("Passive/Active", "Incidents", "Down Hours") $byPassiveActive

    if ($bySite.Count -gt 0) {
        $last = [Math]::Min($bySite.Count + 1, 11)
        Add-Chart $dashboard "Top sites by down hours" 20 210 ($siteSheet.Range("A1:D$last")) 51 | Out-Null
    }
    if ($byHour.Count -gt 0) {
        $last = $byHour.Count + 1
        Add-Chart $dashboard "Incidents by hour" 470 210 ($hourSheet.Range("A1:B$last")) 4 | Out-Null
    }
    if ($byFuel.Count -gt 0) {
        $last = $byFuel.Count + 1
        Add-Chart $dashboard "Fuel vs non fuel" 20 480 ($fuelSheet.Range("A1:B$last")) 5 | Out-Null
    }
    if ($byPassiveActive.Count -gt 0) {
        $last = $byPassiveActive.Count + 1
        Add-Chart $dashboard "Passive / Active" 470 480 ($paSheet.Range("A1:B$last")) 51 | Out-Null
    }

    $dashboard.Columns.AutoFit() | Out-Null
    $workbook.SaveAs($resolvedOutput, 51)
}
finally {
    $workbook.Close($true)
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}

Write-Output "Dashboard created: $resolvedOutput"
Write-Output "Incidents: $totalIncidents"
Write-Output "Sites impacted: $uniqueSites"
Write-Output "Total down hours: $totalDownHours"
