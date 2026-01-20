# This needs to be updated each month when the path changes
$FilePath = [POST PATH]

$Dates = (Read-Host "Combine which? [MM.DD-MM.DD+1]").split("-")

$reports = @()

# fill $reports with lines from Anchorage csvs from the first date to just before the second date
dir $FilePath | Where-Object {$_.Name -match "Anchorage" -and $_.Name -gt $Dates[0] -and $_.Name -lt $Dates[1]} | % { $reports += Import-CSV $_.FullName }

$reports | export-csv ($FilePath + "\Combined Anchorage.csv") -NoTypeInformation