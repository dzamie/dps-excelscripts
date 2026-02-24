# remember to pass "xx" for month and "xxxx" for year
# they determine where files are saved
# the journal log MUST be copied/renamed manually

$30months = @(4,6,9,11)

$psLoc = "[SCRIPT PATH]"
$logFile = "[LOG PATH]\log.txt" # for errors
$resFile = "[RESOURCE PATH]\full auto res.txt"

$scripts = @(
  "auto downloader.ps1",
  "auto wss.ps1",
  "WSS create and open.ps1",
  "99 WSS full starter.ps1"
)
$rdp = "99 Remote Chrome thing.rdp"

# get dates of previous run and of today
# regex instead of indexing in case I change/add things later
$datergx = "(?<=Last ran: )((\d+) ?)+"
$datefmt = "MM dd yyyy"
$resstr = @(Get-Content $resFile)
$datestr = ($resstr | Select-String -Pattern $datergx).matches[0].value
$oldDate = Get-Date $datestr
$newDate = Get-Date

function rangeToString {
  param(
  [DateTime[]]$range
  )
  $old = $range[0]
  $new = $range[1]
  if($old.Year -ne $new.Year) {
    "$($old.ToString('MM.dd.yy'))-$($new.ToString('MM.dd.yy'))"
  } elseif($old.Month -ne $new.Month) {
    "$($old.ToString('MM.dd'))-$($new.ToString('MM.dd.yy'))"
  } elseif($old.Day -ne $new.Day) {
    "$($old.ToString('MM.dd'))-$($new.ToString('dd.yy'))"
  } else {
    "$($old.ToString('MM.dd.yy'))"
  }
}

# WSS files are olddate - newdate-1
# EMF files are olddate+1 - newdate
$wssDates = @($oldDate, $newDate.AddDays(-1))
$emfDates = @($oldDate.AddDays(1), $newDate)
if($emfDates[0] -gt $emfDates[1]) {
  Add-Content -Path $logfile -Value "Full Starter: $($newDate.toString('dd MMM yyyy')): EMF range $($emfDates[0].toString('dd.MM.yy')) - $($emfDates[1].toString('dd.MM.yy'))"
  throw "Error: start date is after end date"
}
$wssStr = rangeToString $wssDates
$emfStr = rangeToString $emfDates

# update file
# going by index to stop once the date line is found
for($i = 0; $i -lt $resstr.Count; $i ++) {
  if($resstr[$i] -match $datergx) {
    $resstr[$i] = "Last ran: $($newDate.ToString($datefmt))"
    break
  }
}
# 
$resstr += @("$($newDate.ToString('dd MMM yyyy')) EMF: ${emfStr}")
$resstr | Set-Content -Path $resFile

# verify
Write-Output "EMF string: ${emfStr}`nWSS string: ${wssStr}"

# run other scripts and rdp
powershell -file "${psloc}\$($scripts[0])" # auto downloader
powershell -file "${psloc}\$($scripts[1])" -userdate $wssStr -month $newDate.toString("MM") -year $newDate.toString("yyyy") # auto wss
powershell -file "${psloc}\$($scripts[2])" -userdate $emfStr -monthnumber $newDate.toString("MM") -yearnumber $newDate.toString("yyyy") # emf starter (wss create and open)
Invoke-Item "${psloc}\${rdp}"

# # below is obsolete due to get-date and external resource file usage
#
# # date string helper
# # turn ??-long integers into 2-long strings
# function dsh {
#   param([int]$num)
#   if($num -lt 10) {
#     "0${num}"
#   } else {
#     "${num}"
#   }
# }
# 
# # takes EMF start/end dates in [d, m, y]
# # returns [EMF date string, WSS date string]
# function makeDateStrings {
#   param(
#     [int[]]$startdate,
#     [int[]]$enddate
#   )
#   $wssstart = @($startdate[0]-1) + @($startdate[1..2])
#   $wssend = @($enddate[0]-1) + @($enddate[1..2])
#   foreach($date in @($wssstart, $wssend)) {
#     if($date[0] -eq 0) { # should be last day of prev month
#       $date[1] -= 1
#       if($date[1] -eq 0) { # should be Dec of prev year
#         $date[2] -= 1
#         $date[1] = 12
#       }
#       if($date[1] -eq 2) { # Feb
#         if($date[2] % 4 -eq 0) { # leap year
#           $date[0] = 29
#         } else {
#           $date[0] = 28
#         }
#       } elseif($30months.contains($date[1])) {
#         $date[0] = 30
#       } else {
#         $date[0] = 31
#       }
#     }
#   }
#   if($startdate[0] -eq $enddate[0]) { # same day
#     $emfout = "$(dsh($startdate[1])).$(dsh($startdate[0])).$($startdate[2] % 100)"
#   } elseif($startdate[1] -eq $enddate[1]) { # same month
#     $emfout = "$(dsh($startdate[1])).$(dsh($startdate[0]))-$(dsh($enddate[0])).$($startdate[2] % 100)"
#   } elseif($startdate[2] -eq $enddate[2]) { # same year
#     $emfout = "$(dsh($startdate[1])).$(dsh($startdate[0]))-$(dsh($enddate[1])).$(dsh($enddate[0])).$($startdate[2] % 100)"
#   } else { # different years
#     $emfout = "$(dsh($startdate[1])).$(dsh($startdate[0])).$($startdate[2] % 100)-$(dsh($enddate[1])).$(dsh($enddate[0])).$($enddate[2] % 100)"
#   }
#   if($wssstart[0] -eq $wssend[0]) { # same day
#     $wssout = "$(dsh($wssstart[1])).$(dsh($wssstart[0])).$($wssstart[2] % 100)"
#   } elseif($wssstart[1] -eq $wssend[1]) { # same month
#     $wssout = "$(dsh($wssstart[1])).$(dsh($wssstart[0]))-$(dsh($wssend[0])).$($wssstart[2] % 100)"
#   } elseif($wssstart[2] -eq $wssend[2]) { # same year
#     $wssout = "$(dsh($wssstart[1])).$(dsh($wssstart[0]))-$(dsh($wssend[1])).$(dsh($wssend[0])).$($wssstart[2] % 100)"
#   } else { # different years
#     $wssout = "$(dsh($wssstart[1])).$(dsh($wssstart[0])).$($wssstart[2] % 100)-$(dsh($wssend[1])).$(dsh($wssend[0])).$($wssend[2] % 100)"
#   }
#   @($emfout, $wssout)
# }
# 
# $userdate = Read-Host "Enter day/day range for EMF"
# $dates = @($userdate.split("-") | % {[int]$_})
# if($dates[0] -eq 1) {
#   $chmonth = "new"
# } elseif($dates[0] -gt $dates[-1]) {
#   $chmonth = "split"
# } else {
#   $chmonth = "same"
# }
# # luckily, EMF only goes forward
# switch($chmonth) {
#   "same" { # stay in the same month as last run
#       $startdate = @($dates[0], $currmonth, $curryear)
#       $enddate = @($dates[-1], $currmonth, $curryear)
#   }
#   "new" { # advance the month, then mark dates
#       if($currmonth -eq 12) {
#         $currmonth -= 11
#         $curryear += 1
#       } else {
#         $currmonth += 1
#       }
#       $startdate = @($dates[0], $currmonth, $curryear)
#       $enddate = @($dates[-1], $currmonth, $curryear)
#   }
#   "split" { # mark start date, advance month, mark end date
#       $startdate = @($dates[0], $currmonth, $curryear)
#       if($currmonth -eq 12) {
#         $currmonth -= 11
#         $curryear += 1
#       } else {
#         $currmonth += 1
#       }
#       $enddate = @($dates[-1], $currmonth, $curryear)
#   }
# }
# $datestrings = makeDateStrings $startdate $enddate

# # update this file
# # (this is from when this was self-editing for the lulz, rather than having an external resource file
# if($chmonth -ne "same") {
#   Add-Content -Path $logfile -value "$((Get-Date).toString()):`nChanged month to ${currmonth}.${curryear}"
# 
#   (Get-Content -Path "${psloc}\$($scripts[3])") |
#   % { $_ -replace "currMonth = \d+", "currMonth = ${currmonth}" `
#   -replace "currYear = \d+", "currYear = ${currYear}" } |
#   Set-Content -Path "${psloc}\$($scripts[3])"
# }