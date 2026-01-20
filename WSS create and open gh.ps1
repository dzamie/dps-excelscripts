# update these every month/year

$monthFolder = "[MONTH PATH]"
$monthNumber = "01"
$monthName = "January"
$yearNumber  = "26"

$excelDelay = 2

$locs = @(
  @{"name" = "Spokane"
    "base" = "Spokane P&J\"
    "emf" = "EMF\"
    "wss" = "WSS\" }
  @{"name" = "SLC"
    "base" = "SLC P&J\"
    "emf" = "SLC EMF\"
    "wss" = "SLC WSS\" }
  @{"name" = "Anchorage"
    "base" = "Anchorage\"
    "emf" = "EMF\"
    "wss" = "WSS\" }
)

# Tasks:
# 1. Copy "WSS EMF Journal Entry Template Co11 05.31.2023" in each EMF folder and "WSS Template" in each WSS folder, with appropriate month.day.year
#   1. prompt user for day string
#   2.  JE name: "[m].[d].[y] [loc] WSS EMF JE.xlsx"
#      WSS name: "[m].[d].[y] [loc] WSS Upload.xlsx"
# 2. Open, for Spokane then SLC then Anchorage: EMF log, JE template, WSS template
#   (log name is "[loc] [mname] EMF 20[y]")

# Notes:
# Copy-Item can take param hashtable: Path, Destination

$dayNumber = Read-Host "Please enter date/date range"
if($dayNumber.length -gt 2) {
  # if this is a multi-day
  $wss1 = [int]($dayNumber.Substring(0,2)) - 1
  $wss2 = [int]($dayNumber.Substring(3,2)) - 1
  if($wss1 -lt 10) {
    $wss1 = "0${wss1}"
  }
  if($wss2 -lt 10) {
    $wss2 = "0${wss2}"
  }
  $wssNumber = "${wss1}-${wss2}"
} else {
  # single-day
  $wssNumber = [int]($dayNumber) - 1
  if($wssNumber -lt 10) {
    $wssNumber = "0${wssNumber}"
  } else {
    $wssNumber = "${wssNumber}"
  }
}

# like human use, opens the EMF log, then copies and opens the JE and WSS templates

for($i = 0; $i -lt 3; $i++) {
  $baseDir = "${monthFolder}$($locs[$i]["base"])"
  # open EMF log
  Invoke-Item "${baseDir}$($locs[$i]["name"]) ${monthName} EMF 20${yearNumber}.xlsx"
  # EMF JE
  $copyFolder = "${baseDir}$($locs[$i]["emf"])"
  $copyParams = @{
    Path = "${copyFolder}WSS EMF Journal Entry Template Co11 05.31.2023.xlsx"
    Destination = "${copyFolder}${monthNumber}.${dayNumber}.${yearNumber} $($locs[$i]["name"]) WSS EMF JE.xlsx"
  }
  Copy-Item @copyParams
  Start-Sleep -Seconds $excelDelay
  Invoke-Item "$($copyParams["Destination"])"

  # WSS
  $copyFolder = "${baseDir}$($locs[$i]["wss"])"
  $copyParams = @{
    Path = "${copyFolder}WSS Template.xlsx"
    Destination = "${copyFolder}${monthNumber}.${wssNumber}.${yearNumber} $($locs[$i]["name"]) WSS Upload.xlsx"
  }
  Copy-Item @copyParams
  Start-Sleep -Seconds $excelDelay
  Invoke-Item "$($copyParams["Destination"])"
  Start-Sleep -Seconds $excelDelay
}










