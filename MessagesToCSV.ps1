$CurrentDir = Get-Location

$JsonFile = Get-ChildItem -Path $CurrentDir -Filter *.json | Select-Object -First 1

if (-not $JsonFile) {
    Write-Host "No JSON files found in the current directory." -ForegroundColor Red
    exit
}

$Data = Get-Content -Path $JsonFile.FullName -Raw | ConvertFrom-Json

$Messages = foreach ($Message in $Data.messages) {
    $Timestamp = [DateTimeOffset]::FromUnixTimeMilliseconds($Message.timestamp_ms).ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss")

    $Content = $Message.content
    if ($Content) {
        $Content = $Content -replace "`n", " "
        $Content = $Content -replace "`r", " "
        $Content = $Content -replace "\u00e2\u0080\u0099", "'"
    }
    $Reactions   = if ($Message.reactions) { ($Message.reactions | ForEach-Object { "$($_.reaction) by $($_.actor)" }) -join '; ' } else { $null }
    $Photos      = if ($Message.photos) { ($Message.photos | ForEach-Object { $_.uri }) -join '; ' } else { $null }
    $ShareLink   = $Message.share.link
    $ShareText   = $Message.share.share_text

    [PSCustomObject]@{
        SenderName            = $Message.sender_name
        Timestamp             = $Timestamp
        Content               = $Content
        IsGeoblockedForViewer = $Message.is_geoblocked_for_viewer
        Reactions             = $Reactions
        Photos                = $Photos
        ShareLink             = $ShareLink
        ShareText             = $ShareText
    }
}

$OutputCsvPath = Join-Path -Path $CurrentDir -ChildPath 'output.csv'

$Messages | Export-Csv -Path $OutputCsvPath -NoTypeInformation -Encoding UTF8

Write-Host "Conversion complete! The CSV file 'output.csv' has been created in the current directory."