# Purpose

In case I want to refer to my previous code, or my office's cloud goes down, having a backup is always nice. This readme contains brief descriptions of each script. They're not all ExcelScript, but most of them are. Powershell scripts have base filepaths dummied out - it's hardly infallible security, but I might as well put in at least a token effort.

### Auto PBP - ExcelScript

Processes a report from PayByPhone for easy copying into a csv template. No longer used as of Nov. 2025, when the company switched to a different format of reports.

### Auto PBP New - ExcelScript

(currently unfinished) As above, but for the new report format. It's turning out to be a bit more troublesome, which I blame on pivot tables.

### EMF JE Autosum - ExcelScript

Processes a worksheet from a log of WebSelfStorage transaction records, moves some cells around, sorts them and sums the groups, and generates a block to copy into a csv template (and selects the block for quick copying). This is tied to a button on the template log (not found here, obviously) for ease of activation.

### Finish PBP DPS MOR - ExcelScript

Processes a report taken directly from PBP's site, using different calculations and formatting, and fills out the csv template found in the workbook. Also no longer used as of Nov. 2025.

### Gravity Pivot - ExcelScript

Creates a pivot table from report data, then switches to that sheet to write some sums that are useful at a glance. Not terribly complicated.

### WSS Format Disb - ExcelScript

Sorts, groups, and reformats a (usually) multi-sheet workbook of WSS disbursement reports, adding some labels and formulas for the report it will eventually be used for.

### WSS Format Upload - ExcelScript

Once transaction data has been manually added to the output of the previous script, this gathers the combined data into a summary and drops it into blocks on a new sheet, for easy copying into a csv template. It also moves the cursor to the top-left cell of the first block, for a small save in time and effort.

### CSV ANCH Combine - PowerShell

Combines a number of similarly-named csv files into one file, for easier uploading. These files often need to be uploaded in batches, but the uploader only accepts one file per submission, so this saves time and mouse travel distance.

### WSS create and open - PowerShell

Creates copies of template files and opens them all in one go. There is a short delay between each opening to avoid "Excel cannot open a file while a dialog box is open" errors caused by the computer being slow to load each file. It used to have to be much longer before an upgrade.

### WSS Disb Combine - PowerShell

Combines WSS disbursement report files - each one sheet - into one large workbook. This relies on the assumption that they were downloaded in a specific order. After this runs, the file is ready for the Format Disb excelscript above.

### Auto Downloader - PowerShell

Downloads attachments from emails in a specific Outlook folder and sorts them into their appropriate places on the file system. Used to set up Auto WSS.

### Auto WSS - PowerShell

Combines and processes WSS disbursement report and credit card report files, then generates the relevant csvs, names and moves the processed/formatted files to their appropriate folders, and cleans up the download folders used by Auto Downloader. A combination of WSS Format Disb, WSS Format Upload, and a number of previously-manual steps.

### WSS Full Starter - PowerShell

Uses Get-Date and an external resource file to keep track of where files should go and what they should be named, then runs several scripts to automatically do what can be automated, and set up and open files for what currently can't. It then logs each time it runs. Capable of throwing a custom exception if it's run outside of expected conditions.
It used to rely on user input and self-edit rather than use an external resource file; these sections are commented out in case I want to reference them in a later project.

### ahk manager - AutoHotkey

Helpful script to save bits of effort here and there. Features include:
* remaps the Dedicated Copilot Key (ugh) back to the context menu it used to be on better keyboards
* remaps predictable, repetitive mouseclicks during my many, *many* csv uploads to simple taps of the numpad+ key
* automates a long series of mouseclicks, alt-tabs, and copy/pastes for grabbing PBP DPS MOR reports, using pixel recognition for timing, turning an attention-intensive, half-hour task into a 10-minute task that runs by itself. No longer used as of Nov. 2025.
* maps common strings for permit refunds to simple key combinations, saving hundreds - and sometimes thousands - of keypresses a day
