# goal: find unread disb/cc emails, download their attachments to correct folders

$StationCount = 8

# BasePath is where the raw reports will be saved
$BasePath = "[BASE PATH]"
$ReportFile = "[REPORT PATH]\log.txt"
# these are also folder names
$StationList =  @("B500", "B600", "U131", "U130", "5001", "5003", "5002", "5004")

try {

$outlook = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
$ns = $outlook.GetNamespace("MAPI")
$inbox = $ns.GetDefaultFolder(6)
$wss = $inbox.Folders | Where-Object {$_.name.contains("WSS")}
$ccfold = $wss.Folders | Where-Object {$_.name.contains("CC")}
$disbfold = $wss.Folders | Where-Object {$_.name.contains("Disb")}

$quietdisb = @($disbfold.items | Where-Object {$_.unread})
$unreaddisb = @($quietdisb | Where-Object {$_.attachments.count -gt 0})

# cc emails always have attachments, even if they're blank files
$unreadcc = @($ccfold.items | Where-Object {$_.unread}) | Sort-Object -Property ReceivedTime
if($unreadcc.count -gt $stationcount) { # multiple days
  if($unreadcc[0].subject[-9] -ne $unreadcc[-1].subject[-9]) { # multiple months
    # get the last cc reports from the previous month
    $prevmonth = @($unreadcc | Where-Object {$_.subject[-9] -ne $unreadcc[-1].subject[-9]})[-($stationcount)..-1]
    $unreadcc = $prevmonth + $unreadcc[-($stationcount)..-1]
  } else {
    $unreadcc = $unreadcc[-($stationcount)..-1]
  }
}

# download each station's reports
foreach($station in $StationList) {
  $stationdisb = @($unreaddisb | Where-Object {$_.subject.contains("${station} Report")})
  $stationcc = @($unreadcc | Where-Object {$_.subject.contains("${station} Report")})
  foreach($mail in ($stationdisb + $stationcc)) {
    $fname = $mail.Attachments(1).FileName
    $mail.Attachments(1).SaveAsFile("${BasePath}\${station}\${fname}")
  }
}

# mark all as read
foreach($mail in ($quietdisb + $unreadcc)) {
  $mail.unread = $false
}

} catch {
  Add-Content "$((Get-Date).tostring()) - auto downloader.ps1`n${$_}`n"
}