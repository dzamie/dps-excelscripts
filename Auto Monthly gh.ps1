$testmode = $false

function Test-Out {
  param([String] $message)
  if($testmode) { Write-Output $message }
}

# Sorts, filters, and sums transactions akin to manual processing
# Outputs Summary sheet and per-date Transaction sheets
# manual procedure:
# delete Dates not in range
# delete Types not Sale/Refund
# delete Status not Captured/Refunded
# delete (Type Refund, Method Offline)
# sort Date -> Type -> Method -> Spreedly -> Location
# group by (date, type, method, spreedly, company)
### note: could maybe do date-type-method-company-spreedly?
# sum each group's Amount
# highlight (Sale sum - Refund sum) per-spreedly/company

if($testmode) {
  $infile = Get-Item "[TEST FILE PATH]"
} else {
  $infile = Get-Item (Read-Host "Drag in Transaction file").Replace('"', '')
}

$companyFile = "[STATION/COMPANY REF PATH]"
Test-Out "Importing Station list..."
$colist = Import-Excel $companyFile -EndColumn 3
Test-Out "List imported. Creating lookup..."
$colookup = @{}
$colist | ForEach-Object {
  $colookup[[String]$_."AS400 Station ID"] = [String]$_."Company"
}
Test-Out "Lookup created."

# get date range, to handle single vs multi-day reports
# accepts single dates as well, to cut down on typing
if($testmode) {
  $daterange = @((Get-Date 04.30), (Get-Date 05.02))
} else {
  do {
    $daterange = @((Read-Host "Enter date for report, or hyphenated range") -split "-" | Get-Date)
    if($daterange.Count -eq 1) {
      $daterange += $daterange[0] # turn single day into a range of itself
    } elseif($daterange[0] -gt $daterange[1]) {
      Write-Output "Err: end date before start date. Please reenter."
    }
  } while($daterange[0] -gt $daterange[1]) # this probably won't trigger, but in case I mistype...
}

Test-Out "Loading file..."
$raw = Import-Csv $infile
Test-Out "File loaded."

#$filtered = $raw | Where-Object {
#  ((Get-Date $_."Date") -ge $daterange[0]) -and
#  ((Get-Date $_."Date") -le $daterange[1]) -and
#  (($_."Type" -eq "Sale") -or ($_."Type" -eq "Refund")) -and
#  (($_."Status" -eq "Captured") -or ($_."Status" -eq "Refunded")) -and
#  (-not (($_."Type" -eq "Refund") -and ($_."Method" -eq "Offline")))
#}

$groups = @{}
$recCount = 0
$raw | ForEach-Object {
  if(((Get-Date $_."Date") -ge $daterange[0]) -and
  ((Get-Date $_."Date") -le $daterange[1]) -and
  (($_."Type" -eq "Sale") -or ($_."Type" -eq "Refund")) -and
  (($_."Status" -eq "Captured") -or ($_."Status" -eq "Refunded")) -and
  (-not (($_."Type" -eq "Refund") -and ($_."Method" -eq "Offline")))) {
    $recCount ++
    $temp = $_
    $temp | Add-Member -MemberType NoteProperty -Name "Company" -Value ($colookup[[String]$temp.Location])
    if($temp.Company.Length -lt 2) {
      $temp.Company = "0" + $temp.Company
    }
    $dates = $temp.'Permit Valid Dates' -split " - "
    if((Get-Date $dates[0]) -gt (Get-Date $dates[1])) { # end-date before report date - need to move end-date to end of report-date's month
      $dates[1] = Get-Date (Get-Date $dates[0] -Day 1).AddMonths(1).AddDays(-1) -Format "M/d/yyyy"
      $temp.'Permit Valid Dates' = $dates[0] + " - " + $dates[1]
    }
    $key  = (Get-Date $temp.Date -Format "MM/dd") + " "
    $key += $temp.Type.Substring(0,4) + " "
    $key += $temp.Method.Substring(0,3) + " "
    $key += $temp.Company + " "
    $key += $temp.'Spreedly Merchant Account'
    if($groups.ContainsKey($key)) {
      $groups[$key] += $temp
    } else {
      $groups[$key] = @($temp)
    }
  }
}

Test-Out "Filtered. Found $($recCount) records."
$keylist = $groups.Keys | Sort-Object

# sort on date-type-method-company-spreedly
# note: can add spaces between records by inserting @(""). Maybe run ConvertTo-CSV before exporting to excel, just in case.

$outrecords = @{} # date -> psobject[], for the actual output
$recordsums = @() # psobject[], fields Date, Company, Method, Spreedly, Sum

# example key: "05/01 Sale Onl 5 DPSPPAUTH05"

foreach($key in $keylist) {
  if(-not $outrecords.ContainsKey($groups[$key][0].Date)) {
    $outrecords[$groups[$key][0].Date] = @()
  }
  $outrecords[$groups[$key][0].Date] += $groups[$key]
  # in each record, remove the $ from the money string, turn it into a float, then sum them all
  $sum = ($groups[$key] | ForEach-Object {[math]::round([float]$_.Amount.replace("`$",""),2)} | Measure-Object -Sum).Sum
  $sum = [math]::Round($sum, 2) # to avoid float imprecision errors
  $outrecords[$groups[$key][0].Date] += @([PSCustomObject]@{ "Amount" = $sum }) # "autosum" line
  $outrecords[$groups[$key][0].Date] += @("") # blank separator
  if(@($keylist -like $key.Remove(6, 4).Insert(6, "????")).Count -gt 1) { # sale/refund combo
    if($key[6] -eq "S") {
      $refu = ($groups[$key.Remove(6,4).Insert(6,"Refu")] | ForEach-Object {[math]::round([float]$_.Amount.replace("`$",""),2)} | Measure-Object -Sum).Sum
      $refu = [math]::Round($refu, 2)
      $recordsums += @([PSCustomObject]@{
        "Date" = Get-Date $key.Substring(0,5) -Format "M/d/yyyy"
        "Company" = $key.Substring(15,2)
        "Method" = $groups[$key][0].Method
        "Spreedly" = $key.Substring(18)
        "Sum" = $sum - $refu
      })
    }
  } else {
    $recordsums += @([PSCustomObject]@{
      "Date" = Get-Date $key.Substring(0,5) -Format "M/d/yyyy"
      "Company" = $key.Substring(15,2)
      "Method" = $groups[$key][0].Method
      "Spreedly" = $key.Substring(18)
      "Sum" = if($key[6] -eq "S") { $sum } else { -$sum }
    })
  }
}

# parsed and processed: $outrecords is a [date]->[csv-like] that can each be copied over to the template
#                       $recordsums is a csv-like that should work well for a quick summary sheet

# output plan: summary sheet with $recordsums, (refund sheet with all date online refunds grouped together?), copy/paste sheets with $outrecords

$outfile = $infile.PSParentPath + "\auto.xlsx"
if(Test-Path $outfile) {
  Remove-Item $outfile # avoid weird overlaps, since this makes and names new tabs each time
}

$refunds = @()
$groups.Keys | Where-Object {$_ -match "Refu"} | ForEach-Object {
  $refunds += $groups[$_]
}

if($refunds.Count -gt 0) {
  $refunds | Export-Excel $outfile -WorksheetName "Refunds"
}
$outrecords.Keys | ForEach-Object {
  $outrecords[$_] | Export-Excel $outfile -WorksheetName ($_.replace("/","."))
}

# summary goes at the start, but it's added at the end to use -passthru
$outbook = $recordsums | Export-Excel $outfile -WorksheetName "Summary" -AutoSize -MoveToStart -PassThru

# accounting format string: $#,##0.00_);[Red]($#,##0.00)
# use on cols K:O, T:AE
# and col E on summary
$outbook.Workbook.Worksheets | ForEach-Object {
  Set-ExcelRange -Worksheet $_ -range "A:AE" -FontName "Aptos Narrow" # just looks better tbh
  Set-ExcelRange -Worksheet $_ -Range "K:O" -NumberFormat "$#,##0.00_);[Red]($#,##0.00)"
  Set-ExcelRange -Worksheet $_ -Range "T:AE" -NumberFormat "$#,##0.00_);[Red]($#,##0.00)"
  Set-ExcelRange -Worksheet $_ -Range "L:L" -AutoSize # avoid ###### visuals
}
Set-ExcelRange -Worksheet $outbook.Workbook.Worksheets["Summary"] -Range "E:E" -NumberFormat "$#,##0.00_);[Red]($#,##0.00)" -AutoSize

Close-ExcelPackage -ExcelPackage $outbook -Show