# Purpose

In case I want to refer to my previous code, or my office's cloud goes down, having a backup is always nice. This readme contains brief descriptions of each script. They're not all ExcelScript, but many of them are. Scripts have base filepaths and easily-identifiable information dummied out - it's hardly infallible security, but I might as well put in at least a token effort.

## ExcelScript

### Auto PBP

Processes a report from PayByPhone for easy copying into a csv template. No longer used as of Nov. 2025, when the company switched to a different format of reports.

### Auto PBP Button

Newest of the three, designed to be attached to a button on a template-generating sheet. Uses a Table to counteract PBP constantly shifting the order and number of data report columns, and has early definition of column names in case those start changing, too. Uses dynamic ranges to mostly avoid keeping manual track of the record count and to make sure the resulting .csv doesn't include any empty-looking lines that nonetheless cause problems on upload.

### Auto PBP New

As above, but for the new report format. It's turning out to be a bit more troublesome, which I blame on pivot tables.

### EMF JE Autosum

Processes a worksheet from a log of WebSelfStorage transaction records, moves some cells around, sorts them and sums the groups, and generates a block to copy into a csv template (and selects the block for quick copying). This is tied to a button on the template log (not found here, obviously) for ease of activation.

### EMF JE Autosum New

Same function as EMF JE Autosum, but has more consistent output formatting, automatically detects record dates, is agnostic to the raw records' starting row, and doesn't error out if there's no data to put in the CSV template.

### Finish PBP DPS MOR

Processes a report taken directly from PBP's site, using different calculations and formatting, and fills out the csv template found in the workbook. Also no longer used as of Nov. 2025.

### Gravity Pivot

Creates a pivot table from report data, then switches to that sheet to write some sums that are useful at a glance. Not terribly complicated.

### PBP Template Generator

Processes a 2026-era PBP report, then generates and switches to an export tab, ready to be CSV'd once the links are refreshed. Uses more Excel formulas to take advantage of dynamic ranges so the export tab is always the correct height. Designed to be attached to a button for ease of use and ease of teaching.

### WSS Format Disb

Sorts, groups, and reformats a (usually) multi-sheet workbook of WSS disbursement reports, adding some labels and formulas for the report it will eventually be used for.

### WSS Format Upload

Once transaction data has been manually added to the output of the previous script, this gathers the combined data into a summary and drops it into blocks on a new sheet, for easy copying into a csv template. It also moves the cursor to the top-left cell of the first block, for a small save in time and effort.

## PowerShell

### Auto Downloader

Downloads attachments from emails in a specific Outlook folder and sorts them into their appropriate places on the file system. Used to set up Auto WSS.

### Auto Monthly

With a supplied generated report file and user-specified date range: filters, divides, sums, sorts, and summarizes reports to be easily copied into a template workbook and referenced. Also copies out "refund" transactions to a dedicated sheet, since those, tragically, must be manually entered.

### Auto ParkWhiz

Takes in a report file, divides and edits reports as instructed, then outputs a file containing the divided, edited reports, a summary page, and a page of errors to be manually handled. Also references a station list to create the summary, and turns the lookup table into a hashtable to decrease lookup time by a factor of ~30.

### Auto WSS

Combines and processes WSS disbursement report and credit card report files, then generates the relevant csvs, names and moves the processed/formatted files to their appropriate folders, and cleans up the download folders used by Auto Downloader. A combination of WSS Format Disb, WSS Format Upload, and a number of previously-manual steps.

### CSV ANCH Combine

Combines a number of similarly-named csv files into one file, for easier uploading. These files often need to be uploaded in batches, but the uploader only accepts one file per submission, so this saves time and mouse travel distance.

### Startup

First part of a daily process after logging in. Opens several work-necessary programs, such as web logins, Outlook, Excel, and the AHK Manager. Then triggers Startup AHK to finish the process.

### WSS create and open

Creates copies of template files and opens them all in one go. There is a short delay between each opening to avoid "Excel cannot open a file while a dialog box is open" errors caused by the computer being slow to load each file. It used to have to be much longer before an upgrade.

### WSS Disb Combine

Combines WSS disbursement report files - each one sheet - into one large workbook. This relies on the assumption that they were downloaded in a specific order. After this runs, the file is ready for the Format Disb excelscript above.

### WSS Full Starter

Uses Get-Date and an external resource file to keep track of where files should go and what they should be named, then runs several scripts to automatically do what can be automated, and set up and open files for what currently can't. It then logs each time it runs. Capable of throwing a custom exception if it's run outside of expected conditions.
It used to rely on user input and self-edit rather than use an external resource file; these sections are commented out in case I want to reference them in a later project.

## AutoHotKey

### AHK Manager

Helpful script to save bits of effort here and there. Features include:
* remaps the Dedicated Copilot Key (ugh) back to the context menu it used to be on better keyboards
* organizes several shortcut modes for different tasks, letting the same key (usually numpad+) be used for different things at different times:
  * maps predictable, repetitive mouseclicks during my many, *many* csv uploads to simple taps of the numpad+ key
  * maps filling out predictable forms, taking screenshots, and saving screenshots, for check refunds to simple taps of the numpad+ key, and adds a "jump to step" dialog box on alt+numpad+
  * maps common strings for permit refunds to simple key combinations, saving hundreds - and sometimes thousands - of keypresses a day
  * maps predictable screenshots, copies, and close-tabs to a few key combinations for Journal Entry recording, saving a handful of keypresses and a lot of mouse distance
  * uses Edge's console and a JS script to automatically select items in a form according to a few filters
 
### Startup AHK

Second part of a daily process after logging in. Logs in to two websites, manages window focus, and then watches for new mail before triggering WSS Full Starter.

## Visual Basic

### AutoMail

Simple email template that lets me set a date to include in it. Used for sending bank sheets, sets up a good bit faster than copy/pasting from Sent Items, and requires less mouse precision/ctrl+arrows to fix the dates. Easily adapted to similar but different emails: just change the recipients and body. Still need to attach and send by hand, but it's honestly easier and faster to use the "recent files" part of Outlook's Attach File dropdown, and sending by hand lets me double-check in case something goes wrong.

## JavaScript

### Gravity Helper

Defines a few functions useful for automatically finding and selecting Cash Sales in a Deposits form, based on their overhead string and their date. Also contains similar code, but outside of a function, as a more readable format of the JS script found in the Gravity Helper section of AHK Manager.
