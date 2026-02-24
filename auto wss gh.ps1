param(
  [string]$userdate = "overwrite",
  [string]$month = "02",
  [string]$year = "2026"
)
# ^asks for user input when run alone, but not when called from another script with argument
# month/year must be changed manually for running alone

$TESTMODE = $false

$ReportFile = "[REPORT PATH]"

# VaultPath is where the files are saved to when not testing
$VaultPath = "[VAULT PATH]"
$VaultSpokane = "${VaultPath}\Spokane P&J\WSS"
$VaultSLC = "${VaultPath}\SLC P&J\SLC WSS"
$VaultAnch = "${VaultPath}\Anchorage\WSS"
$VaultCSV = "${VaultPath}\Postings\[NAME]\WSS Post"

# BasePath is where the raw reports are located
$BasePath = "[XLSX REPORTS PATH]"
$csvName = "csv snips.csv"
$wssName = "WSS Template.xlsx"

$StationList =  @("B500", "B600", "U131", "U130", "5001", "5003", "5002", "5004")
# these are only the $$$ columns, which need to be [float]'d
$CCHeadList =   @("Charge", "Discount Amt", "Tran Fees", "Total")
$DisbHeadList = @("Rent", "Late Fees", "Fee", "Invoicing", "Manual Fees", "Deposit", "Retail", "Insurance", "Services", "Reservations", "Other", "Credit", "Applied Credit", "Rental Tax", "Retail Tax", "Total Tax", "Total")
$EcheckFee = -0.17

# takes the 4-character station name as its parameter
# returns a hashtable of lightly-formatted CC and Disb report
function OpenReports {
  param (
    [string]$Station
  )
  $out = @{
    "Disb" = @()
    "CC" = @()
  }
  # find CC file and Disb file(s)
  # $ccfile is usually only 1 element long, except for month rollovers
  $ccfile = @(GCI "${BasePath}\$Station\Cred*")
  $disbs  = @(GCI "${BasePath}\$Station\Disb*")
  foreach($c in $ccfile) {
    $out."CC" += ($c | Import-Excel -DataOnly)
  }
  foreach($d in $disbs) {
    $out."Disb" += ($d | Import-Excel -DataOnly)
  }
  
  # turn $$$ strings into useful floats
  # CC first
  foreach($rec in $out."CC") {
    foreach($head in $CCHeadList) {
      $rec.$head = $rec.$head.replace("(","-") # negative detection
      $rec.$head = $rec.$head.replace(")","")  # negative cleanup
      $rec.$head = $rec.$head.replace("$","")  # dollarsign removal
      $rec.$head = [float]($rec.$head)
    }
  }
  # now Disbs
  foreach($rec in $out."Disb") {
    foreach($head in $DisbHeadList) {
#      if($rec.$head -eq "-") { # undash the 0s
#        $rec.$head = [float]0 # so it turns out that this isn't necessary
#      } else {
       $rec.$head = $rec.$head.replace("$","") # undollar the signs
#      }
    }
  }
  $out
}

# takes the disbursement list from OpenReports
# returns a nested hashtable: $out[Date][Type] = records[]
function SortDisbs {
  param (
    $DisbIn
  )
  $out = @{}
  foreach($rec in $DisbIn) {
    # set up new record structure for new date
    if(!$out.ContainsKey($rec.Date)) {
      $out.($rec.Date) = @{
        "CC" = @()
        "Echeck" = @()
        "Check" = @()
        "Cash" = @()
        "Corpo" = @()
        "RAC" = @()
      }
    }
    switch -wildcard ($rec.Type) {
      "Check"	{ $dest = "Check"; break }
      "Eche*"	{ $dest = "Echeck"; break }
      "Corp*"	{ $dest = "Corpo"; break }
      "Cash"	{ $dest = "Cash"; break }
      "Repo*"	{ $dest = "RAC"; break }
      default	{ $dest = "CC" }
    }
    $out.($rec.Date).$dest += $rec
  }
  $out
}

# helper function for GenerateWSSSheet
# takes a blank "template" psobject, a disb psobject, and optionally a cc psobject
# returns the template with the relevant fields filled
function PairCCDisb {
  param (
    [Parameter(Mandatory=$true)] $blank,
    [Parameter(Mandatory=$true)] $disb,
    [Parameter(Mandatory=$false)] $cc = $null
  )
  $out = $blank.PSObject.Copy()
  $disbHeads = @("Contract Number", "Date", "Type", "Rent", "Late Fees", "Fee", "Invoicing", "Manual Fees", "Deposit", "Retail", "Insurance", "Services", "Reservations", "Other", "Credit", "Applied Credit", "Rental Tax", "Retail Tax", "Total Tax", "Total")
  $ccHeads = @("Ach ID", "Card Number", "Card Holder", "Tran Date", "Charge", "Discount Amt", "Tran Fees")
  foreach($h in $disbHeads) {
    $out.$h = $disb.$h
  }
  if(! ($cc -eq $null)) {
    foreach($h in $ccHeads) {
      $out.$h = $cc.$h
    }
    $out."CC Total" = $cc."Total"
  }
  $out
}

# helper function for GenerateWSSSheet
# takes a record, a list of target strings, and a list of replacement strings
# returns the record with the targets replaced in all fields
function RecReplace {
  param (
    $record,
    [string[]]$trgt,
    [string[]]$repl
  )
  $recstr = ConvertTo-CSV $record -NoTypeInformation
  foreach($i in 0..($trgt.length-1)) {
    $recstr = $recstr.replace($trgt[$i], $repl[$i])
  }
  $out = ConvertFrom-CSV $recstr
#  $out = $record.PSObject.Copy()
#  foreach($prop in $out.PSObject.Properties.Name) {
#    if($out.$prop.getType() -eq "str".getType()) { # only do this to strings
#      foreach($i in 0..(@($trgt).count - 1)) {
#        $out.$prop = $out.$prop.replace($trgt[$i], $repl[$i])
#      }
#    }
#  }
  $out
}

# takes the CC/Disb hashtable from OpenReports, with Disb formatted via SortDisbs
# returns a records[]-like that can be exported to Excel and a list of rows with lbls
function GenerateWSSSheet {
  param (
    $reports
  )
  $disb = $reports.Disb
  $cc = $reports.CC
  $snips = Import-CSV "${BasePath}\${csvName}"
  $blank = $snips[0]	# blank line
  $sums = $snips[1]	# record group autosums
  $labels = $snips[2]	# GL ACCT labels
  $calcs = $snips[3]	# GL ACCT calculations
  $topline = 3		# Row of first record in group, used for sums/calcs
  $TransTypes = @("CC", "Echeck", "Check", "Corpo", "RAC", "Cash")

  $out = @()
  $lbllist = @()
  # iterate through dates on the outside
  # from oldest to newest (thanks Sort-Object)
  foreach($date in ($reports.Disb.Keys | Sort-Object)) {
    # separating records into same-type blocks
    foreach($type in $TransTypes) {
      # skip over types without any records
      if($reports.Disb.$date.$type.count -eq 0) {
        continue
      }
      $recCt = $reports.Disb.$date.$type.count
      $block = @()
      $block += $blank # start with a blank line
      # then add the records - paired with CC data if applicable
      if($type -eq "CC") {
        $currCC = $reports.CC | Where-Object { $_."Tran Date" -eq $date }
        foreach($i in 0..($recCt - 1)) { # god I hate off-by-one errors
          $block += PairCCDisb $blank $reports.Disb.$date.$type[$i] $currCC[$i]
        }
      } else {
        foreach($rep in $reports.Disb.$date.$type) {
          $block += PairCCDisb $blank $rep
        }
      }
      # then the autosums
      $block += (RecReplace $sums @("!!!", "???") @("${topline}", "$($topline + $recCt - 1)"))
      # and the labels
      $block += $labels
      # make a note of where this is
      $lbllist += @($topline + $recCt + 1)
      # and finally the calculations
      $block += (RecReplace $calcs @("!!!", "???") @("$($topline + $recCt)", "$($topline + $recCt + 2)"))
      # don't forget that Echecks are weird
      if($type -eq "Echeck") {
        $block[-1]."Manual Fees" = $EcheckFee * $recCt
      }
      # then put the block in the output and update topline
      $out += $block
      $topline += $block.count
    }
  }
  @{"CSV" = $out; "Labels" = $lbllist}
}

# takes array of station names and target xlsx filename to write to
# creates and formats target-path workbook
# returns psobject that can be used to populate upload template
function GenerateDisbBook {
  param (
    [string[]] $stationList,
    [string] $outputPath = "${BasePath}\test.xlsx"
  )
  # make a template PSObject so the output can be made easier
  $outTmplt = New-Object -TypeName PSObject -Property @{
    "Date"		= "now";
    "Station"		= "here";
    "AS400 GL ACCT"	= 1337;
    "Amount"		= 0;
    "Memo"		= "dragon!"
  }
  $out = @()
  $stationMgr = @{}
  $exportParams = @{
    Path = $outputPath;
    Calculate = $true;
    FreezeTopRow = $true;
    BoldTopRow = $true
  }
  $book = @() # this will get overwritten

  foreach($station in $stationlist) {
    $stationMgr[$station] = OpenReports $station
    $stationMgr[$station].Disb = SortDisbs $stationMgr[$station].Disb
    $stationMgr[$station] = GenerateWSSSheet $stationMgr[$station]
    # the station now has a "csv" and a "labels" list
    # time to add it to the workbook
    $exportParams["WorksheetName"] = $station
    # previously bugged out when trying to use passthru
    $stationMgr[$station].csv | Export-Excel @exportParams
  }
  # in case something goes wrong, let's make sure everything's there
  # we can save it again later
  $book = Open-ExcelPackage ($exportParams.Path)
  
  #formatting time
  foreach($station in $stationList) {
    $book.$station.Cells["1:1"].Style.Font.Name = "Times New Roman"
    $book.$station.Cells["1:1"].Style.Font.Size = 12
    $book.$station.Cells["D:T"].Style.NumberFormat.Format = "`"`$`"#,##0.00_);[Red]\(`"`$`"#,##0.00\)"
    $book.$station.Cells["Y:AB"].Style.NumberFormat.Format = "`"`$`"#,##0.00_);[Red]\(`"`$`"#,##0.00\)"
    $book.$station.Cells["A:T"].AutoFitColumns()
    foreach($row in $stationMgr[$station].Labels) {
      $book.$station.Cells["E${row}:K${row}"].Style.Fill.PatternType = [OfficeOpenXML.style.ExcelFillStyle]::solid
      $book.$station.Cells["E${row}:K${row}"].Style.Fill.BackgroundColor.SetColor(255, 0, 176, 240)
    }
  }
  [OfficeOpenXml.CalculationExtension]::Calculate($book.Workbook)
  $book.save()
  # formatted and saved. I could open it and do the csv manually if I wanted to.
  # but I don't. so onwards we go.

  $specialMatches = @("Ech*", "Che*", "Corp*", "Cash", "Rep*")
  foreach($station in $stationList) {
    $GlAccts = @(1600, 1804, 1808, 3709, 4100, 1065, 1810)
    # handle 5002's weird thing
    if($station -eq "5002") {
      $GlAccts[0] = 1200
    }
    $cols = @([Char[]]"EFGHIJK")
    foreach($row in $stationMgr[$station].Labels) {
      $type = $book.$station.Cells["C$($row - 2)"].Value
      # check if it's a non-CC block for memo purposes
      if($specialMatches.where({$type -like $_}).count -gt 0) {
        $memo = "${station} ${type}"
      } else {
        $memo = $station
      }
      foreach($i in 0..($cols.count - 1)) {
        # rounding to 2 places to avoid failing comparisons that should work
        try {
          $val = [math]::round($book.$station.Cells["$($cols[$i])$($row+1)"].Value, 2)
        } catch {
          $report =
"$((Get-Date).toString()) - auto wss.ps1
File: $($book.File.Name)
Address: $($cols[$i])$($row+1)
Value: $($book.$station.Cells["$($cols[$i])$($row+1)"].Value)
Type: $($book.$station.Cells["$($cols[$i])$($row+1)"].Value.getType())`n"
          Add-Content -Path $ReportFile -Value $Report
        }
        if($val -ne 0) { # nonzero value boyyyyz letsgo
          $rec = $outTmplt.psobject.copy()
          $rec.Date = $book.$station.Cells["B$($row-2)"].Value
          $rec.Station = $station
          $rec."AS400 GL ACCT" = $GlAccts[$i]
          $rec.Amount = $val
          $rec.Memo = $memo
          if($i -gt 4) { # check for DP
            $rec.Memo += " DP"
          }
          $out += $rec
        }
      }
    }
  }
  $out
}

# takes upload-like psobject array from GenerateDisbBook, and target filepaths
# creates, fills, and crops upload template, and also csv
# returns nothing
function GenerateUploadBook {
  param (
    $records,
    [string]$TmpltPath = "${BasePath}\testUL.xlsx",
    [string]$CsvPath = "${BasePath}\testUL.csv"
  )
  $headers = @("Date", "Station", "AS400 GL ACCT", "Amount", "Memo")
  $cols = @([Char[]]"ABCDE")
  $file = Open-ExcelPackage "${BasePath}\${WSSName}"
  try {
    $file.saveAs($TmpltPath)
  } catch {
$report = 
"$((Get-Date).toString()) - auto wss.ps1
Path: ${TmpltPath}
Path Type: $($TmpltPath.getType())`n"
    Add-Content -Path $ReportFile -Value $report
  }
  $sheet = $file.Workbook.Worksheets[1]
  foreach($i in 0..($records.count - 1)) {
    foreach($j in 0..($cols.count - 1)) {
      $sheet.Cells["$($cols[$j])$($i+2)"].Value = $records[$i].($headers[$j])
    }
  }
  $sheet.DeleteRow($records.count + 2, 999)
  # this updates the text/values based on the formulas
  [OfficeOpenXml.CalculationExtension]::Calculate($sheet)
  $file.save()
  # astext flag so that the Dates are read as mm/dd/yyyy rather than "since 1/1/1900"
  Import-Excel -ExcelPackage $file -AsText @("*Date") | Export-CSV $CsvPath -NoTypeInformation
}

# takes nothing
# returns stations with disbursements in them, organized by area
function GetFileGroups {
  $stations = @{
    "Spokane" = @("B500", "B600");
    "SLC" = @("U131", "U130");
    "Anchorage" = @("5001", "5003", "5002", "5004")
  }
  $out = @{}
  foreach($city in $stations.keys) {
    foreach($station in $stations.$city) {
      if(@(GCI "${BasePath}\${station}\Disb*").count -gt 0) {
        if($out.containsKey($city)) {
          $out.$city += $station
        } else {
          $out[$city] = @($station)
        }
      }
    }
  }
  $out
}

# The Big One
# takes a date string for filenames, and possibly target folders for output
# checks each city/station for disbursements
# generates disbursement sheets and saves to target folders
# generates upload sheets and saves to target folders
# generates upload csvs and saves to target folder
# returns nothing
function AutoWSS {
  param (
    [string]$fileDate = "test",
    [string]$SpokaneFolder = "${BasePath}\Spokane Output",
    [string]$SLCFolder = "${BasePath}\SLC Output",
    [string]$AnchFolder = "${BasePath}\S_Anch Output",
    [string]$CSVFolder = "${BasePath}\CSV Output"
  )
  
  $folders = @{
    "Spokane"	= $SpokaneFolder;
    "SLC"	= $SLCFolder;
    "Anchorage" = $AnchFolder;
    "csv"	= $CSVFolder
  }
  $stations = GetFileGroups
  foreach($city in $stations.keys) {
    $BaseFilename = "${fileDate} ${city} WSS"
    $DisbParams = @{
      StationList = $stations[$city];
      OutputPath  = "$($folders[$city])\${BaseFilename}.xlsx"
    }
    $UploadObject = GenerateDisbBook @DisbParams

    $UploadParams = @{
      records	= $UploadObject;
      TmpltPath	= "$($folders[$city])\${BaseFilename} Upload.xlsx"
      CsvPath	= "$($folders["csv"])\${BaseFilename} Upload.csv"
    }
    GenerateUploadBook @UploadParams
  }
  
  # finds non-Output folders (doesn't have output or a file extension) to delve into
  foreach($stationName in ((gci $BasePath | Where-Object {-not ($_.Name.Contains("Output") -or $_.Name.Contains("."))}).Name)) {
    Move-Item -Path "${BasePath}\${stationName}\*.xlsx" -Destination "${BasePath}\${stationName}\Done"
  }
}

# main commands time
if($TESTMODE) {
  $WssParams = @{
    "fileDate" = "test";
  }
} else {
  $WssParams = @{
    "fileDate" = "${userdate}";
    "SpokaneFolder" = "${VaultSpokane}";
    "SLCFolder" = "${VaultSLC}";
    "AnchFolder" = "${VaultAnch}";
    "CSVFolder" = "${VaultCSV}"
  }
}
if($userdate -eq "overwrite") {
  $userdate = Read-Host "Enter date string"

  # smart filename-finishing via .-counting
  # this does break slightly if the year ever hits 5 digits
  $dotcount = ([regex]::matches($userdate, "\.")).count
  if($dotcount -eq 0) { # days only - assume month and year
    $WssParams["fileDate"] = "${month}.${userdate}.$($year.substring(2,2))"
  } elseif($dotcount -eq 2) { # months and days - assume year
    $WssParams["fileDate"] = "${userdate}.$($year.substring(2,2))"
  } else { # woah nelly
    $WssParams["fileDate"] = $userdate
  }
}
AutoWSS @WssParams