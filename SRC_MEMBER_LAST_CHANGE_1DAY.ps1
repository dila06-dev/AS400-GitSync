# =========================
# Konfiguration
# =========================
$ApiUrl      = "http://localhost:8085/api/services/SRC_MEMBER_LAST_CHANGE_1DAY"
$BearerToken = "a1CLZIUkn2mqNzuM5IpCRR4mwnRYfTUkGCALnAuR247e7d45"
$CsvPath     = "D:\NET_INFO\analyse\SRC_MEMBER_LAST_CHANGE_1DAY.csv"
$Delimiter   = ";"   # Für deutsches Excel meist besser ";", sonst "," verwenden
$TimestampFormat = "yyyy-MM-dd HH:mm:ss"


# Timeout für PowerShell Request, Sekunden
$RequestTimeoutSec = 3000

$Headers = @{
    Authorization = "Bearer $BearerToken"
    Accept        = "application/json"
}

function Convert-ToCsvReadyObject {
    param (
        [Parameter(Mandatory = $true)]
        [object]$InputObject,

        [Parameter(Mandatory = $true)]
        [string]$Timestamp
    )

    $result = [ordered]@{}
    $result["Zeitstempel"] = $Timestamp

    foreach ($prop in $InputObject.PSObject.Properties) {
        $value = $prop.Value

        if ($null -eq $value) {
            $result[$prop.Name] = $null
        }
        elseif ($value -is [System.Array] -or $value -is [hashtable] -or $value -is [pscustomobject]) {
            $result[$prop.Name] = ($value | ConvertTo-Json -Depth 10 -Compress)
        }
        else {
            $result[$prop.Name] = $value
        }
    }

    return [pscustomobject]$result
}

try {
    $CsvFolder = Split-Path -Path $CsvPath -Parent
    if (-not (Test-Path $CsvFolder)) {
        New-Item -ItemType Directory -Path $CsvFolder -Force | Out-Null
    }

    Write-Host "Rufe API auf: $ApiUrl"

    $Response = Invoke-RestMethod `
        -Uri $ApiUrl `
        -Method Get `
        -Headers $Headers `
        -TimeoutSec $RequestTimeoutSec ` 
        -ErrorAction Stop

    if ($Response -is [System.Array]) {
        $Data = $Response
    }
    elseif ($Response.PSObject.Properties.Name -contains "data") {
        if ($Response.data -is [System.Array]) {
            $Data = $Response.data
        }
        else {
            $Data = @($Response.data)
        }
    }
    else {
        $Data = @($Response)
    }

    if (-not $Data -or $Data.Count -eq 0) {
        Write-Host "Keine Daten von der API zurückgegeben."
        exit
    }

    Write-Host "Datensätze erhalten: $($Data.Count)"

    $CurrentTimestamp = (Get-Date).ToString($TimestampFormat)
    $AppendCsv = (Test-Path $CsvPath) -and ((Get-Item $CsvPath).Length -gt 0)

    $Data |
        ForEach-Object {
            Convert-ToCsvReadyObject -InputObject $_ -Timestamp $CurrentTimestamp
        } |
        Export-Csv `
            -Path $CsvPath `
            -Delimiter $Delimiter `
            -NoTypeInformation `
            -Encoding UTF8 `
            -Append:$AppendCsv

    if ($AppendCsv) {
        Write-Host "Daten wurden an die CSV angehängt: $CsvPath"
    }
    else {
        Write-Host "CSV wurde neu erstellt: $CsvPath"
    }
}
catch {
    $statusCode = $null
    $responseBody = $null

    if ($_.Exception.Response) {
        try {
            $statusCode = [int]$_.Exception.Response.StatusCode
            $stream = $_.Exception.Response.GetResponseStream()

            if ($stream) {
                $reader = New-Object System.IO.StreamReader($stream)
                $responseBody = $reader.ReadToEnd()
                $reader.Close()
            }
        }
        catch {
            $responseBody = "Konnte Fehlerantwort nicht lesen: $($_.Exception.Message)"
        }
    }

    if (-not $responseBody -and $_.ErrorDetails.Message) {
        $responseBody = $_.ErrorDetails.Message
    }

    Write-Error @"
Fehler beim API-Request oder CSV-Schreiben.

HTTP Status: $statusCode
PowerShell Fehler: $($_.Exception.Message)

Server-Antwort:
$responseBody
"@
}