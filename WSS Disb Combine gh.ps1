# This needs to be updated each month when the path changes
$FilePath = "[IMPORT PATH]"

$StationNames = @("B500","B600","U131","U130","5001","5003","5002","5004")
$StationCount = 8

# Goal: combine Disbursement reports, rename tabs, maybe even run the necessary script
# Plan:
# * Prompt user for excluded stations (digit string 1-8)
# * Create list of included stations
# * Get list of downloaded files in creation order
# * Make new Excel instance, file
# * Set Excel to visible
# * Iterate through remaining files in order:
#   - Open file
#   - Copy sheet into final sheet's book
#   - Rename the new-copied sheet
#   - Close file
# * Delete initial end-sheet
# * Run "WSS Format Disb" if possible

# Snippets:
# Get-ChildItem -Path "${DLPath}\Dis*.xlsx" | Sort-Object -Property LastWriteTime (-Descending)
# foreach($file in $InputFiles) {
#   $openfile = $excel.workbooks.open($file.fullname)
#   $file.sheets.item(1).Copy($MergeSheet)
#   $file.Close()
# }
# $Workbook.sheets.item([1-indexed]).name get/set
# note: sheet.copy() throws it into a new book. sheet.copy(sheet2) inserts it *before* sheet2 in the workbook sheet2 is in

# TODO: prompt and filter
"#`tName`t#`tName"
"1`tB500`t5`t5001"
"2`tB600`t6`t5003"
"3`tU131`t7`t5002"
"4`tU130`t8`t5004"
$SkipString = Read-Host "Please enter stations to skip"
# split string into char array, then turn them back into strings to read as ints
$SkipList = ([char[]]$SkipString).ForEach({ [int][string]$_ })

# open excel
$EO = New-Object -ComObject excel.application
# get file list in order
$InputFiles = Get-ChildItem -Path "${FilePath}\Dis*.xlsx" | Sort-Object -Property LastWriteTime
# make book, mark end sheet to be deleted later
$MergeBook = $EO.Workbooks.add()
$EndSheet = $MergeBook.sheets.item("Sheet1")

$nameCounter = 1
for($i = 0; $i -lt $InputFiles.length; $i++) {
  # advance name counter until it finds something not to be skipped
  while($SkipList.contains($nameCounter)) { $nameCounter ++ }
  $openBook = $EO.workbooks.open($InputFiles[$i].fullname)
  $openBook.sheets.item(1).copy($EndSheet)
  $openBook.close()
  $MergeBook.sheets.item($i + 1).name = $StationNames[$nameCounter - 1]
  $nameCounter ++
}

$EndSheet.delete()
if(Test-Path "${FilePath}\merged.xlsx") {
  Remove-Item "${FilePath}\merged.xlsx"
}
$MergeBook.SaveAs("${FilePath}\merged.xlsx")
$MergeBook.close()
$EO.quit()

Get-ChildItem "${FilePath}\Disb*" | Move-Item -Destination "${FilePath}\Done\"

Invoke-Item "${FilePath}\merged.xlsx"