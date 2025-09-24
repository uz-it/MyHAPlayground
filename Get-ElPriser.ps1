param(
    [int]$WeekNumber
)

# ===== Helper: Get start & end date for ISO week =====
function Get-DateRangeFromWeekNumber {
    param(
        [int]$WeekNumber,
        [int]$Year = (Get-Date).Year
    )
    # ISO weeks: find first Thursday of the year, then calculate offset
    $jan4 = Get-Date "$Year-01-04"
    $thursdayOfWeek1 = $jan4.AddDays(3 - (($jan4.DayOfWeek.value__ + 6) % 7))
    $startDate = $thursdayOfWeek1.AddDays(($WeekNumber - 1) * 7 - 3)  # Monday of that week
    $endDate   = $startDate.AddDays(7) # Sunday
    return @{
        Start = $startDate.Date
        End   = $endDate.Date
    }
}

# ===== Date range selection =====
if ($PSBoundParameters.ContainsKey('WeekNumber')) {
    $range = Get-DateRangeFromWeekNumber -WeekNumber $WeekNumber
    $startDate = $range.Start.ToString("yyyy-MM-dd")
    $endDate   = $range.End.ToString("yyyy-MM-dd")
    Write-Host "Using week $WeekNumber ($startDate to $endDate)" -ForegroundColor Cyan
} else {
    $startDate = (Get-Date).AddDays(-6).ToString("yyyy-MM-dd")  # last 7 days incl. today
    $endDate   = (Get-Date).AddDays(1).ToString("yyyy-MM-dd")  # fetch until tomorrow midnight
    Write-Host "Using last 7 days ($startDate to $endDate)" -ForegroundColor Cyan
}

# ===== Dinel tariffs (øre/kWh) med moms =====
$Sommertarif = @{
    Lav   = 9.83
    Høj   = 14.74
    Spids = 38.31
}
$Vintertarif = @{
    Lav   = 9.83
    Høj   = 29.48
    Spids = 88.43
}

[decimal]$elafgift = 90.01  # States afgift (øre/kWh) er med moms
[decimal]$energinet = 16.87
$moms_factor = 0.125

# ===== API URL =====
$url = "https://api.energidataservice.dk/dataset/Elspotprices?start=$startDate&end=$endDate&filter={""PriceArea"":[""DK1""]}&limit=0"

$response = Invoke-RestMethod -Uri $url -Method Get

# ===== Prepare data =====
$prices = $response.records | ForEach-Object {
    $date = $_."HourDK".Substring(0,10)
    $hour = [int]($_."HourDK".Substring(11,2))
    $month = [int](Get-Date $date).Month
    $spotpris = [decimal]$_.SpotPriceDKK #+ 50

    $season = if ($month -ge 4 -and $month -le 9) { "Sommertarif" } else { "Vintertarif" }

    if ($hour -ge 17 -and $hour -le 20) {
        $cat = "Spids"
    } elseif (($hour -ge 6 -and $hour -le 16) -or ($hour -ge 21 -and $hour -lt 24)) {
        $cat = "Høj"
    } else {
        $cat = "Lav"
    }

    [decimal]$TransportbetalingDKK = if ($season -eq "Sommertarif") { $Sommertarif[$cat] } else { $Vintertarif[$cat] }
    #$totalDKK = ($TransportbetalingDKK  + $elafgift )  + ( $spotpris* $moms_factor)
    [decimal]$totalDKK = ($spotpris* $moms_factor) + $TransportbetalingDKK + $elafgift +$energinet
    #Write-host $date, $hour , $TransportbetalingDKK , $elafgift, $energinet, ($spotpris * $moms_factor), $totalDKK
    [PSCustomObject]@{
        Date = $date
        Hour = $hour
        TotalPriceDKK = [math]::Round(($totalDKK /100) ,2)
        Tarif = $cat
        Season = $season
    }
}

# ===== Pivot table: Hours as rows, Dates as columns =====
$hours = 0..23
$dates = ($prices | Select-Object -ExpandProperty Date | Sort-Object -Unique)

$table = foreach ($h in $hours) {
    $row = [ordered]@{Hour = $h.ToString("D2")}
    foreach ($d in $dates) {
        $value = ($prices | Where-Object { $_.Hour -eq $h -and $_.Date -eq $d } | Select-Object -ExpandProperty TotalPriceDKK)
        $row[$d] = if ($value -ne $null) { $value } else { "-" }
    }
    [PSCustomObject]$row
}

# ===== Print coloured pivot table =====
$hourWidth = 6
$colWidth  = 10

$header = "{0,-$hourWidth}" -f "Hour"
Write-Host -NoNewline $header
foreach ($d in $dates) {
    $hdr = "{0,$colWidth}" -f $d
    Write-Host -NoNewline $hdr
}
Write-Host ""

$dayStats = @{}
foreach ($d in $dates) {
    $vals = $table | ForEach-Object { $_.$d } | Where-Object { $_ -ne "-" } | ForEach-Object { [decimal]$_ }
    if ($vals.Count -gt 0) {
        $min = ($vals | Measure-Object -Minimum).Minimum
        $max = ($vals | Measure-Object -Maximum).Maximum
        $mid = ($min + $max) / 2
    } else {
        $min = $null; $max = $null; $mid = $null
    }
    $dayStats[$d] = [PSCustomObject]@{ Min = $min; Max = $max; Mid = $mid }
}

foreach ($row in $table) {
    $hourCell = "{0,-$hourWidth}" -f $row.Hour
    Write-Host -NoNewline $hourCell

    foreach ($d in $dates) {
        $cell = $row.$d
        if ($cell -eq "-") {
            $text = "{0,$colWidth}" -f $cell
            Write-Host -NoNewline $text
            continue
        }

        $stats = $dayStats[$d]
        $num = [double]$cell

        if ($stats.Max -ne $null -and $num -eq $stats.Max) {
            $bg = "Red"; $fg = "White"
        } elseif ($stats.Mid -ne $null -and $num -ge $stats.Mid) {
            $bg = "Yellow"; $fg = "Black"
        } else {
            $bg = "Green"; $fg = "Black"
        }

        $text = "{0,$colWidth}" -f ("{0:N2}" -f $num)
        Write-Host -NoNewline $text -BackgroundColor $bg -ForegroundColor $fg
    }
    Write-Host ""
}
