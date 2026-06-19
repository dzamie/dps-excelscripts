param(
  [string]$userdate = "overwrite"
)
# ^asks for user input when run alone, but not when called from another script with argument

$date = Get-Date
$monthNumber = $date.toString("MM")
$monthName = $date.toString("MMMM")
$year = $date.toString("yyyy")

# note: as of 02.12, this no longer does wss, only emf

$reportfile = "[REPORT PATH]"
function logThing {
  param([string]$error)
  Add-Content -Path $ReportFile -Value "$((Get-Date).toString()) - WSS create and open.ps1`n${error}`n"
}
try {

$excelDelay = 1

# $yearNumber = $yearNumber.substring(2,2)
# $months = @{
#  "01" = "January"
#  "02" = "February"
#  "03" = "March"
#  "04" = "April"
#  "05" = "May"
#  "06" = "June"
#  "07" = "July"
#  "08" = "August"
#  "09" = "September"
#  "10" = "October"
#  "11" = "November"
#  "12" = "December"
# }
# $monthName = $months[$monthNumber]
$monthFolder = "[VAULT PATH]\${year}\${monthNumber}.${year}"

$locs = @(
  @{"name" = "Spokane"
    "base" = "Spokane P&J"
    "emf" = "EMF"}
  @{"name" = "SLC"
    "base" = "SLC P&J"
    "emf" = "SLC EMF" }
  @{"name" = "Anchorage"
    "base" = "Anchorage"
    "emf" = "EMF" }
)

# Tasks:
# 1. Copy "WSS EMF Journal Entry Template Co11 05.31.2023" in each EMF folder, with appropriate month.day.year
#   1. prompt user for day string
#   2.  JE name: "[m].[d].[y] [loc] WSS EMF JE.xlsx"
# 2. Open, for Spokane then SLC then Anchorage: EMF log, JE template
#   (log name is "[loc] [mname] EMF 20[y]")

# Notes:
# Copy-Item can take param hashtable: Path, Destination

if($userdate -eq "overwrite") {
  $userdate = Read-Host "Please enter FULL date/date range (mm.dd.yy)"

#  $dotcount = ([regex]::matches($userdate, "\.")).count
#  if($dotcount -eq 0) { # same month
#    $fulldate = "${monthNumber}.${userdate}.${yearNumber}"
#  } elseif($dotcount -eq 2) { # different month
#    $fulldate = "${userdate}.${yearNumber}"
#  } else { # different year
#    $fulldate = $userdate
#  }
} else { # full starter already does this
  $fulldate = $userdate
}

# like human use, opens the EMF log, then copies and opens the JE templates

for($i = 0; $i -lt 3; $i++) {
  $baseDir = "${monthFolder}\$($locs[$i]["base"])"
  # open EMF log
  Invoke-Item "${baseDir}\$($locs[$i]["name"]) ${monthName} EMF ${year}.xlsx"

  # EMF JE
  $copyFolder = "${baseDir}\$($locs[$i]["emf"])"
  $copyTarget = @(gci "${copyFolder}\WSS EMF*")[0]
  $copyParams = @{
    Path = $copyTarget.FullName
    Destination = "${copyFolder}\${fulldate} $($locs[$i]["name"]) WSS EMF JE.xlsx"
  }
  Copy-Item @copyParams
  Start-Sleep -Seconds $excelDelay
  Invoke-Item "$($copyParams["Destination"])"
  Start-Sleep -Seconds $excelDelay
}
} catch {
  logThing($_.ToString())
}
