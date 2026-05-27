$testmode = $false

$stationFile = "[PARKWHIZ LOOKUP PATH]"
$companyFile = "[STATION/COMPANY LOOKUP PATH]"
Write-Output "Importing Park Whiz list..."
$whizlist = Import-Excel $stationFile
Write-Output "Importing Master Station list..."
$statlist = Import-Excel $companyFile -EndColumn 3
Write-Output "Station List imported!"

$whizlookup = @{}
$whizlist | ForEach-Object {
  $whizlookup[[String]$_."Location Name"] = [String]$_."Station"
}
$statlookup = @{}
$statlist | ForEach-Object {
  $statlookup[[String]$_."AS400 Station ID"] = [String]$_."Company"
}

<#
# Procedure:
# 0. Separate CA/non-CA
#   0. Literally just a where-object
# 1. Perform CA calcs:
#   0. when buyer fee != 0, base price = list price
#   1. when buyer fee == 0, buyer fee -= 1.55, seller fee += 1.55
#   2. if round(seller/base, 2) != -0.03, buyer += 0.55, seller -= 0.55
#   3. if still != -0.03, undo (+1 to buyer, -1 to seller) and add to Error group
#   4. else, DPS Total (for summary) = Add_on + Base + Buyer + Seller (put in Station:Amount hashtable)
# 2. Perform non-CA calcs:
#   0. if round(buyer+list-base, 2) == 0.48, base = list + buyer
#   1. disc = -(buyer+seller+list)
#   2. DPS Total (for summary) = list + discount + buyer + seller (put in Station:Amount hashtable)
# 3. Summarize for easier upload checking:
#   0. Put the two tables together, make 4 sums: Olympus, co12, co05, co11
#   1. Olympus: Station 4592
#   2. co12: Station 4*
#   3. co05: Station 5*
#   4. co11: Station *
# 
# Output file:
# * Summary: banksheet -> total
# * Processed (CA)
# * Processed (Other)
# * Raw (CA)
# * Raw (Other)
# (for multi-days, output file itself can be filtered)
#>

function CoLookup {
  param([String] $whizname)
  if($whizlookup.ContainsKey($whizname)) {
    $statname = $whizlookup[$whizname]
    if($statname -eq "4592") {
      "Oly"
    } else {
      if($statlookup.ContainsKey($statname)) {
        $statlookup[$statname]
      } else {
        -1
    }
    }
  } else {
    -2
  }
}


# function CoLookup {
#   param([String] $whizname)
#   $whizfind = @($whizlist | Where-Object {$_."Location Name" -eq $whizname})
#   if($whizfind) {
#     $statname = $whizfind[0]."Station"
#     if([String]$statname -eq "4592") {
#       "Oly"
#     } else {
#       $statfind = @($statlist | Where-Object {[String]($_."AS400 Station ID") -eq [String]$statname})
#       if($statfind) {
#         $statfind[0]."Company"
#       } else {
#         -1
#       }
#     }
#   } else {
#     -2
#   }
# }

if($testmode) {
  $infile = Get-Item "[TEST FILE PATH]"
} else {
  $infile = Get-Item (Read-Host "Drag in report file").Replace('"', '')
}

Write-Output "Importing reports..."
$reports = Import-Excel $infile
Write-Output "Reports imported! Found $($reports.Count) reports!"
$cali = @()
$noncali = @()
$errReports = @()

$NAReports = @()
$totalcheck = [ordered]@{
  "PS05" = 0
  "PS11" = 0
  "PS12" = 0
  "Olym" = 0
  "#N/A" = 0
}
$lastcheck = 0

foreach($trans in $reports) {
  $percentage = 100 * ($cali.Count + $noncali.Count) / $reports.Count
  $percentage = [math]::Round($percentage, 0)
  if((($percentage % 5) -eq 0) -and $percentage -ne $lastcheck) { Write-Output "Finished $($percentage)% of reports."; $lastcheck = $percentage }
  if($trans."Seller Account Name".contains("Cali")) {
    # 0. when buyer fee != 0, base price = list price
    if($trans."Sum Total Buyer Fee" -ne 0) {
      $trans."Sum Base Price Original" = $trans."Sum List Price"
    }
    # 1. when buyer fee == 0, buyer fee -= 1.55, seller fee += 1.55
    else {
      $trans."Sum Total Buyer Fee" -= 1.55
      $trans."Sum Total Seller Fee" += 1.55
      # 2. if round(seller/base, 2) != -0.03, buyer += 0.55, seller -= 0.55
      if([math]::Round($trans."Sum Total Seller Fee" / $trans."Sum Base Price Original", 2) -ne -0.03) {
        $trans."Sum Total Buyer Fee" += 0.55
        $trans."Sum Total Seller Fee" -= 0.55
      }
      # if still != -0.03, add to Error group
      if([math]::Round($trans."Sum Total Seller Fee" / $trans."Sum Base Price Original", 2) -ne -0.03) {
        $errReports += $trans
      }
    }
    # 4. DPS Total (for summary) = Base + Add_on  + Buyer + Seller (put in Station:Amount hashtable)
    $total = $trans."Sum Base Price Original" + $trans."Sum Discounts" + $trans."Sum Total Buyer Fee" + $trans."Sum Total Seller Fee"

    if((CoLookup -whizname $trans."Location Internal Name") -lt 0) {
      # station not found - probably in parkwhiz but possibly in station list
      # either way, manual intervention needed
      $NAReports += $trans
      $totalcheck["#N/A"] += $total
    } else {
      $totalcheck["PS11"] += $total
    }

    $cali += $trans
  } else {
    # 0. if round(buyer+list-base, 2) == .48, base = list + buyer
    if([math]::Round($trans."Sum Total Buyer Fee" + $trans."Sum List Price" - $trans."Sum Base Price Original", 2) -eq 0.48) {
      $trans."Sum Base Price Original" = $trans."Sum Total Buyer Fee" + $trans."Sum List Price"
    }
    # 1. disc = net-(buyer+seller+list)
    # using Round to avoid "5e-17" everywhere
    $trans."Sum Discounts" = [math]::Round($trans."Sum Net" - ($trans."Sum Total Buyer Fee" + $trans."Sum List Price" + $trans."Sum Total Seller Fee"), 2)
    # 2. DPS Total (for summary) = list + discount + buyer + seller (put in Station:Amount hashtable)
    $total = $trans."Sum List Price" + $trans."Sum Discounts" + $trans."Sum Total Buyer Fee" + $trans."Sum Total Seller Fee"

    switch ([String](CoLookup -whizname $trans."Location Internal Name")) {
      "11" { $totalcheck["PS11"] += $total; break }
      "5" { $totalcheck["PS05"] += $total; break }
      "05" { $totalcheck["PS05"] += $total; break }
      "Oly" { $totalcheck["Olym"] += $total; break }
      "12" { $totalcheck["PS12"] += $total; break }
      default { $totalcheck["#N/A"] += $total; $NAReports += $trans }
    }

    $noncali += $trans
  }
}

$sumout = @()
foreach($key in $totalcheck.Keys) {
  $sumout += [PSCustomObject][Ordered]@{
    Station = $key
    Total = $totalcheck.$key
    }
}

$outfile = $infile.PSParentPath + "\auto.xlsx"

$cali | Export-Excel $outfile -WorksheetName "CA"
$noncali | Export-Excel $outfile -WorksheetName "Non-CA"
if($errReports.count -gt 0) {
  $errReports | Export-Excel $outfile -WorksheetName "Calc Errors"
}
if($NAReports.count -gt 0) {
  $NAReports | Export-Excel $outfile -WorksheetName "Missing Station"
}
$sumout | Export-Excel $outfile -WorksheetName "Summary" -MoveToStart -Numberformat "$#,##0.00" -AutoSize -Show