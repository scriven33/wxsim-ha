# Adjust this path to match where your latest.csv file is found. 
$csvPath = "C:\wxsim\latest.csv"
# Adjust this path for the location of the ouput file. 
$jsonPath = "C:\Inetpub\wwwroot\forecast.json"

# Skip second row (units)
$raw = Get-Content $csvPath
$header = $raw[0]
$data = $raw | Select-Object -Skip 2
$tempCsv = "$env:TEMP\wxsim_cleaned.csv"
Set-Content -Path $tempCsv -Value $header
Add-Content -Path $tempCsv -Value $data

# Import raw CSV
$csv = Import-Csv $tempCsv

# Clean headers
$cleaned = foreach ($row in $csv) {
    $obj = @{}
    foreach ($prop in $row.PSObject.Properties) {
        $name = $prop.Name.Trim()
        $value = $prop.Value
        $obj[$name] = $value
    }
    [pscustomobject]$obj
}

# Condition mapping
function Get-Condition {
    param ($temp, $precip, $humidity, $wind, $hour)

    $isNight = ($hour -lt 6 -or $hour -ge 21)

    if ($temp -lt 1 -and $precip -gt 0) {
        return 'snowy'
    } elseif ($precip -ge 10) {
        return 'pouring'
    } elseif ($precip -ge 2) {
        return 'rainy'
    } elseif ($precip -gt 0) {
        return 'partlycloudy'
    } elseif ($isNight) {
        return 'clear-night'
    } elseif ($humidity -gt 95) {
        return 'fog'
    } elseif ($wind -gt 15) {
        return 'windy'
    } else {
        return 'sunny'
    }
}

# Parse forecast
$forecast = foreach ($row in $cleaned) {
    try {
        $dt = [datetime]::ParseExact($row.'UTC Date/Time'.Trim(), "yyyy-MM-dd_HH:mm_UTC", $null)
        $temp = [double]$row.'Temperature'.Trim()
        $hum = [double]$row.'Rel.Hum.'.Trim()
        $wind = [double]$row.'Wind Spd.'.Trim()
        $gust = [double]$row.'1 hr Gust'.Trim()
        $rain = if ($row.'Tot.Prcp'.Trim() -match '^[\d\.]+$') { [double]$row.'Tot.Prcp'.Trim() } else { 0 }
        $dew = [double]$row.'Dew Pt.'.Trim()
        $pres = [double]$row.'S.L.P.'.Trim()
        $vis = [double]$row.'VIS'.Trim()
        $uv = [double]$row.'UV Index'.Trim()
        $cloud = [double]$row.'Sky Cov'.Trim()
        $apparent = [double]$row.'Heat Ind'.Trim()
        $bearing = [double]$row.'Wind Dir.'.Trim()

        $cond = Get-Condition $temp $rain $hum $wind $dt.Hour

        [pscustomobject]@{
            datetime             = $dt
            temperature          = $temp
            apparent_temperature = $apparent
            humidity             = $hum
            dew_point            = $dew
            pressure             = $pres
            wind_speed           = $wind
            wind_gust_speed      = $gust
            wind_bearing         = $bearing
            visibility           = $vis
            uv_index             = $uv
            cloud_coverage       = $cloud
            precipitation        = $rain
            condition            = $cond
        }
    } catch {
        Write-Warning "Skipping row due to parse error: $_"
    }
}

# Sort
$forecast = $forecast | Sort-Object datetime

# HOURLY
$hourly = $forecast | Select-Object -First 48 | ForEach-Object {
    [pscustomobject]@{
        datetime             = $_.datetime.ToString("s") + "Z"
        temperature          = $_.temperature
        apparent_temperature = $_.apparent_temperature
        humidity             = $_.humidity
        dew_point            = $_.dew_point
        pressure             = $_.pressure
        wind_speed           = $_.wind_speed
        wind_gust_speed      = $_.wind_gust_speed
        wind_bearing         = $_.wind_bearing
        visibility           = $_.visibility
        uv_index             = $_.uv_index
        cloud_coverage       = $_.cloud_coverage
        precipitation        = $_.precipitation
        condition            = $_.condition
    }
}

# DAILY
$daily = @()
$grouped = $forecast | Group-Object { $_.datetime.Date }

foreach ($day in $grouped | Select-Object -First 5) {
    $temps     = $day.Group | Select-Object -ExpandProperty temperature
    $apparents = $day.Group | Select-Object -ExpandProperty apparent_temperature
    $hums      = $day.Group | Select-Object -ExpandProperty humidity
    $dews      = $day.Group | Select-Object -ExpandProperty dew_point
    $winds     = $day.Group | Select-Object -ExpandProperty wind_speed
    $gusts     = $day.Group | Select-Object -ExpandProperty wind_gust_speed
    $pressures = $day.Group | Select-Object -ExpandProperty pressure
    $vis       = $day.Group | Select-Object -ExpandProperty visibility
    $uvs       = $day.Group | Select-Object -ExpandProperty uv_index
    $clouds    = $day.Group | Select-Object -ExpandProperty cloud_coverage
    $precips   = $day.Group | Select-Object -ExpandProperty precipitation

    $avgTemp = ($temps | Measure-Object -Average).Average
    $sumPrecip = ($precips | Measure-Object -Sum).Sum
    $cond = Get-Condition $avgTemp $sumPrecip ($day.Group[0].datetime.Hour)

    $daily += [pscustomobject]@{
        datetime             = $day.Group[0].datetime.Date.ToString("s") + "Z"
        temperature          = [math]::Round(($temps | Measure-Object -Maximum).Maximum, 1)
        templow              = [math]::Round(($temps | Measure-Object -Minimum).Minimum, 1)
        apparent_temperature = [math]::Round(($apparents | Measure-Object -Average).Average, 1)
        humidity             = [math]::Round(($hums | Measure-Object -Average).Average, 0)
        dew_point            = [math]::Round(($dews | Measure-Object -Average).Average, 1)
        wind_speed           = [math]::Round(($winds | Measure-Object -Average).Average, 1)
        wind_gust_speed      = [math]::Round(($gusts | Measure-Object -Maximum).Maximum, 1)
        pressure             = [math]::Round(($pressures | Measure-Object -Average).Average, 1)
        visibility           = [math]::Round(($vis | Measure-Object -Average).Average, 1)
        uv_index             = [math]::Round(($uvs | Measure-Object -Average).Average, 1)
        cloud_coverage       = [math]::Round(($clouds | Measure-Object -Average).Average, 1)
        precipitation        = [math]::Round($sumPrecip, 2)
        condition            = $cond
    }
}

# Output
$output = [pscustomobject]@{
    generated_at         = (Get-Date).ToUniversalTime().ToString("s") + "Z"
    temperature_unit     = "°C"
    pressure_unit        = "hPa"
    wind_speed_unit      = "mph"
    visibility_unit      = "km"
    precipitation_unit   = "mm"
    hourly               = $hourly
    daily                = $daily
}

$utf8NoBomEncoding = New-Object System.Text.UTF8Encoding $false
[System.IO.File]::WriteAllText($jsonPath, ($output | ConvertTo-Json -Depth 4), $utf8NoBomEncoding)
Write-Host "forecast.json created at $jsonPath"
