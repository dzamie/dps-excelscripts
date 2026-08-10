^`::ExitApp

F23::AppsKey

^NumpadAdd:: {
  ; Display choice box
  global
  GuiBox.Show()
}

CoordMode "Mouse", "Screen"
CoordMode "Pixel", "Screen"
global StepCt := 1
global ActiveFlag := 0
global MyGui
global StepGui
global InVars

Csv_Active(Button, Info) {
  global
  All_Off(0)
  ActiveFlag |= 1
  StepCt := 1
  Hotkey "NumpadAdd", Csv_Next, "On"
  Hotkey "+NumpadAdd", Csv_416, "On"
  Hotkey "^Space", Csv_Reset, "On"
  Hotkey "``", Csv_Off, "On"
  GuiBox.Hide()
  return true
}

PPt_Active(Button, Info) {
  global
  All_Off(0)
  ActiveFlag |= 2
  Hotkey "^1", PPt_1, "On"
  Hotkey "^2", PPt_2, "On"
  Hotkey "^3", PPt_3, "On"
  Hotkey "^4", PPt_4, "On"
  Hotkey "+^1", PPt_5_1, "On"
  Hotkey "``", PPt_Off, "On"
  GuiBox.Hide()
  return true
}

Chk_Active(Button, Info) {
  global
  All_Off(0)
  ActiveFlag |= 4
  StepCt := 1
  Hotkey "NumpadAdd", Chk_Step, "On"
  Hotkey "!NumpadAdd", Chk_ChangeStep, "On"
  Hotkey "``", Chk_Off, "On"
  GuiBox.Hide()
  Chk_Menu()
  return true
}

JE_Active(Button, Info) {
  global
  All_Off(0)
  ActiveFlag |= 8
  Hotkey("NumpadAdd", JE_Basic, "On")
  Hotkey("+NumpadAdd", JE_Sameday, "On")
  Hotkey("^+NumpadAdd", JE_Multiday, "On")
  Hotkey("``", JE_Off, "On")
  GuiBox.Hide()
  return true
}

Grav_Active(Button, Info) {
  global
  All_Off(0)
  ActiveFlag |= 16
  Hotkey("NumpadAdd", Grav_StatVis, "On")
  Hotkey("``", Grav_Off, "On")
  GuiBox.Hide()
  jsScript := 'let overhead = "OHplace";let startDate = "SDplace";let endDate = "EDplace";const startHigh = ["BLV", "09801", "01904", "PAS"];let dateList = [];for(let d = new Date(startDate); d <= new Date(endDate); d.setDate(d.getDate() + 1)) { dateList.push(d.toLocaleDateString(undefined, {month: "2-digit", day: "2-digit"}));}if(! startHigh.includes(overhead)) { dateList = dateList.slice(0, -1);}for(const date of dateList) { let recordList = Array.prototype.slice.call(document.getElementsByClassName("uir-list-row-tr")); let relevantList = recordList.filter((x) => x.childNodes[4].innerHTML.includes(overhead) && x.childNodes[4].innerHTML.includes(date)); relevantList.forEach((x) => x.childNodes[0].childNodes[0].click());}'
  Grav_Setup()
  return true
}

; Show which script is running
Active_Scripts(Button, Info) {
  global
  GuiBox.Hide()
  CheckString := "CSV: " . ((ActiveFlag & 1 > 0) ? "On" : "Off")
  CheckString .= "`nMonthly: " . ((ActiveFlag & 2 > 0) ? "On" : "Off")
  CheckString .= "`nCheck: " . ((ActiveFlag & 4 > 0) ? "On" : "Off")
  CheckString .= "`nJE: " . ((ActiveFlag & 8 > 0) ? "On" : "Off")
  CheckString .= "`nGravity: " . ((ActiveFlag & 16 > 0) ? "On" : "Off")
;  CheckString .= "`nClicks Eaten: " . ClicksEaten
  MsgBox CheckString
  return true
}


global GuiBox := Gui(, "Script Spawner",)
GuiBox.AddText(, "Run which script?")
GuiBox.AddButton(, "CSV Helper").OnEvent("Click", Csv_Active)
GuiBox.AddButton(, "Monthly Helper").OnEvent("Click", PPt_Active)
GuiBox.AddButton(, "Check Refund Helper").OnEvent("Click", Chk_Active)
GuiBox.AddButton('ym', "Journal Entry Helper").OnEvent("Click", JE_Active)
GuiBox.AddButton(, "Gravity Helper").OnEvent('Click', Grav_Active)
GuiBox.AddButton(, "Active Scripts").OnEvent("Click", Active_Scripts)

; Go through the CSV Upload templates. Default NumpadAdd is for templates where multi-threading starts Off
Csv_Next(foo) {
  global
  if(StepCt = 1) {
    Click 1778, 994
    StepCt++
  } else if(StepCt = 2) {
    Click 453, 395
    Sleep 200
    Click 1111, 612
    Sleep 200
    Click 1369, 720
    Sleep 200
    Click 1238, 860
    Sleep 200
    Click 1778, 994
    StepCt++
  } else if(StepCt = 3) {
    Click 1778, 994
    StepCt++
  } else if(StepCt = 4) {
    Click 1825, 984, 0
    Sleep 200
    Click 1744, 939
    StepCt++
  } else {
    MsgBox "Out of steps! Ctrl-Space to reset."
  }
}

; replaces step 2 for when default multi-threading is On
Csv_416(foo) {
  global
  Click 453, 395
  Sleep 200
  Click 1369, 720
  Sleep 200
  Click 1238, 860
  Sleep 200
  Click 1778, 994
  StepCt++
}

Csv_Reset(foo) {
  global
  StepCt := 1
  MsgBox "Steps reset!"
}

Csv_Off(foo) {
  global
  Hotkey "NumpadAdd", "Off"
  Hotkey "+NumpadAdd", "Off"
  Hotkey "^Space", "Off"
  Hotkey "``", "Off"
  ActiveFlag &= (255-1)
}


PPt_1(foo) {
  Sleep 200
  Send "Customer cancelling permit. Refunding X Monthly Permit minus fees."
  Send "^{Left 5}"
  Send "{Left}"
  Send "{BS}"
}
PPt_2(foo) {
  Sleep 200
  Send "Customer returned access device. Refunding Access Device fee."
}
PPt_3(foo) {
  Sleep 200
  Send "Customer returned access device. Requesting check to refund Access Device fee."
}
PPt_4(foo) {
  Sleep 200
  Send "Customer switching lots. Refunding X Monthly Permit."
  Send "^{Left 3}"
  Send "{Left}"
  Send "{BS}"
}
PPt_5_1(foo) {
  Sleep 200
  Send "Customer reducing permit count by "
  Hotkey "Tab", PPt_5_2, "On"
}
PPt_5_2(foo) {
  Sleep 200
  Send ". Refunding "
  Hotkey "Tab", PPt_5_3, "On"
}
PPt_5_3(foo) {
  Sleep 200
  Send " Monthly Permits minus fees."
  Hotkey "Tab", "Off"
}

PPt_Off(foo) {
  global
  Hotkey "^1", "Off"
  Hotkey "^2", "Off"
  Hotkey "^3", "Off"
  Hotkey "^4", "Off"
  Hotkey "+^1", "Off"
  Hotkey "``", "Off"
  ActiveFlag &= (255-2)
}

; set up vars for check refund
Chk_Menu(*) {
  global MyGui := Gui(,"Variable Input")
  MyGui.Add("Text",, "Station:")
  MyGui.Add("Text",, "Memo:")
  MyGui.Add("Text",, "Amount:")
  MyGui.Add("Text",, "Bill Name:")
  MyGui.Add("Text",, "Folder:")
  MyGui.Add("Radio","Group Checked vType", "Access Device")
  MyGui.Add("Radio",, "Monthly Permit")
  MyGui.Add("Edit", "vStation ym")
  MyGui.Add("Edit", "vMemo")
  MyGui.Add("Edit", "vAmount")
  MyGui.Add("Edit", "vBill")
  MyGui.Add("Edit", "vFolder")
  MyGui.Add("Button", "default", "OK").OnEvent("Click", Chk_Vars)
  MyGui.OnEvent("Close", Chk_Vars)
  MyGui.Show()
}

Chk_Vars(*) {
  global MyGui
  global InVars := MyGui.Submit()
}

Chk_ScreenShot(x, y) {
  Send "#+s"
  Sleep 1500
  Click 20, 270, "Down"
  Sleep 200
  Click x, y, "Up"
  Sleep 2000
  Click 1740, 820
  Sleep 1500
  Send "!{Esc}" ; Windows doesn't like letting you alt-tab, but this works
  Sleep 200
  Send "^s"
}

Chk_Step(*) {
  global StepCt
  global InVars
  Sleep 200
  Sleep 200
  if(StepCt = 1) {
    ; Cash Refund
    Send InVars.Station "{Tab}"
    Sleep 5000
    Click 20, 270, "WU", 5
    Sleep 200
    Click 160, 520
    Sleep 200
    Send "14{Tab 6}"
    Send InVars.Memo
    Click 20, 270, "WD", 5
    Sleep 200
    Sleep 200
    Click 60, 885
    if(InVars.Type = 1) {
      Send "park{Tab}"
    } else {
      Send "mon{Tab}"
    }
    Sleep 2000
    Send "{Tab}"
    Sleep 200
    Send InVars.Memo "{Tab 2}"
    Sleep 500
    Send InVars.Amount "{Tab 3}"
    Sleep 500
    Send InVars.Station "{Tab}"
    Sleep 500
    Click 20, 270, "WD", 3
    Sleep 200
    Sleep 200
    Click 80, 915
  } else if(StepCt = 2) {
    ; CR screenshot
    A_Clipboard := 0
    Click 60, 327, 2
    Sleep 200
    Send "^c"
    ClipWait
    Click 60, 237, 1
    Sleep 200
    RefundID := A_Clipboard
    Chk_ScreenShot(1650, 700)
    Sleep 1000
    Send "CR " RefundID
  } else if(StepCt = 3) {
    ; CR save screenshot
    Send InVars.Folder "{Enter}"
    Sleep 600
    Send "!s"
    Sleep 200
    Sleep 200
    Send "!{F4}"
  } else if(StepCt = 4) {
    ; Deposit part 1
    Send "{Tab}10351{Tab 3}" InVars.Memo
    Sleep 200
    Click 87, 725
  } else if(StepCt = 5) {
    ; Deposit part 2
    Click 87, 725, "WU", 10
    Sleep 200
    Sleep 200
    Sleep 200
    Click 265, 725
    Sleep 200
    Send "{Tab}"
    Sleep 200
    Send InVars.Amount "{Tab}"
    Sleep 200
    Send "24007{Tab 3}"
    Sleep 200
    Send InVars.Station "{Tab}"
    Sleep 1000
    Click 80, 870
  } else if(StepCt = 6) {
    ; Deposit screenshot
    A_Clipboard := 0
    Click 60, 327, 2
    Sleep 200
    Send "^c"
    ClipWait
    Click 60, 237, 1
    Sleep 200
    RefundID := A_Clipboard
    Chk_ScreenShot(1770, 560)
    Sleep 1000
    Send "Deposit " RefundID
  } else if(StepCt = 7) {
    ; Deposit save screenshot
    Send InVars.Folder "{Enter}"
    Sleep 600
    Send "!s"
    Sleep 200
    Sleep 200
    Send "!{F4}"
  } else if(StepCt = 8) {
    ; Bill
    Click 87, 725, "WD", 3
    Sleep 200
    Sleep 200
    Click 1070, 425 ; GEO
    Sleep 200
    Send "DPS Consolidated GEO : Consolidated Parking : US Parking : Other US Parking : WAPKS{Tab}"
    Click 87, 725, "WU", 5
    Sleep 200
    Sleep 200
    Click 1075, 400
    Sleep 200
    Send InVars.Memo
    Sleep 200
    Click 175, 448
    Sleep 200
    Send InVars.Bill
    Click 175, 448, "WD", 5
    Sleep 200
    Sleep 200
    Click 185, 880
    Sleep 200
    Send "park"
    Sleep 200
    Send "{Down}"
    Sleep 200
    Send "{Tab}"
    Sleep 1000
    Send InVars.Station "{Tab 2}"
    Sleep 1000
    Send InVars.Memo "{Tab}"
    Sleep 200
    Send InVars.Amount "{Tab}"
    Click 175, 448, "WD", 3
    Sleep 500
    Click 80, 915
    Sleep 1000
    Click 175, 448, "WD", 3
    Sleep 200
    Sleep 200
    Click 220, 625
    Sleep 200
    Click 335, 905
  } else if(StepCt = 9) {
    ; Bill screenshot
    Chk_ScreenShot(1500, 960)
    Sleep 1000
    Send "Bill " InVars.Bill
  } else if(StepCt = 10) {
    ; Bill save screenshot
    Send InVars.Folder "{Enter}"
    Sleep 600
    Send "!s"
    Sleep 200
    Sleep 200
    Send "!{F4}"
  } else {
    MsgBox("Out of steps!")
  }
  StepCt += 1
}

Chk_ChangeStep(*) {
  global StepGui := Gui(,"Choose new step")
  StepGui.Add("Radio", "Group Checked vNewStep", "1. Cash Refund")
  StepGui.Add("Radio",, "2. CR Screencap")
  StepGui.Add("Radio",, "3. CR Save SC")
  StepGui.Add("Radio",, "4. Deposit part 1")
  StepGui.Add("Radio", "ym", "5. Deposit part 2")
  StepGui.Add("Radio",, "6. Deposit Screencap")
  StepGui.Add("Radio",, "7. Deposit Save SC")
  StepGui.Add("Radio",, "8. Bill")
  StepGui.Add("Radio",, "9. Bill Screencap")
  StepGui.Add("Radio",, "10. Bill Save SC")
  StepGui.Add("Button", "default", "OK").OnEvent("Click", Chk_GetStep)
  StepGui.OnEvent("Close", Chk_GetStep)
  StepGui.Show()
}

Chk_GetStep(*) {
  global StepCt
  global StepGui
  tempvar := StepGui.Submit()
  StepCt := tempvar.NewStep
}

Chk_Off(*) {
  global
  Hotkey "NumpadAdd", "Off"
  Hotkey "!NumpadAdd", "Off"
  Hotkey "``", "Off"
  ActiveFlag &= 255-4
}

; Journal Entry approval automation
; Start with excel focused on right screen, Edge unfocused on left
JE_Basic(*) {
  global
  Sleep 200
  Send "#+s" ; take screenshot
  Sleep 1500
  Click 20, 200, "Down"
  Sleep 200
  Click 1350, 480, "Up"
  Sleep 1000
  Send "^v" ; paste screenshot
  Sleep 300
  A_Clipboard := 0
  Click 75, 250, 2 ; select JE number
  Sleep 200
  Send "^c" ; copy
  ClipWait(, 0)
  Send "^w" ; close Edge tab
  Sleep 400 ; wait for Edge to finish closing tab
  Click 2500, 500 ; refocus Excel
  Hotkey("NumpadAdd", JE_NextCity, "On")
}
; for ANCH, when the next JE is a different station on the same day
JE_Sameday(*) {
  global
  JE_Basic(0)
}
; for multi-days, when the next sheet needs to be switched to without closing
JE_Multiday(*) {
  global
  Sleep 200
  Send "!hot"
  Sleep 200
  Send "{Down 2}{Left 3}{Enter}"
  Sleep 200
  Send "^{PgDn}"
  Hotkey("NumpadAdd", JE_basic, "On")
}
; switching JE workbooks by saving and closing
JE_NextCity(*) {
  global
  JE_multiday(0)
  Sleep 200
  Send "^w"
  Sleep 500
  Send "!s"
}

JE_Off(*) {
  global
  Hotkey("NumpadAdd", "Off")
  Hotkey("+NumpadAdd", "Off")
  Hotkey("^+NumpadAdd", "Off")
  Hotkey("``", "Off")
  ActiveFlag &= 255-8
}

Grav_Setup(*) {
  global DateBox := Gui(, 'Enter Date Range',)
  DateBox.Add('Text',,'Start Date:')
  DateBox.Add('Text',,'End Date:')
  DateBox.Add('Edit', 'vStartDate ym')
  DateBox.Add('Edit', 'vEndDate')
  DateBox.Add('Button', 'default xm', 'OK').OnEvent('Click', Grav_SetDates)
  global StationBox := Gui(, 'Select Station')
  StationBox.Add('Text',,'Seattle')
  StationBox.Add('Button',, "SEA").OnEvent('Click', Grav_SetStat)
  StationBox.Add('Text','ym','West WA')
  StationBox.Add('Button',, 'co12 UDT').OnEvent('Click', Grav_SetStat) ; 04590
  StationBox.Add('Button',, 'KNG+UDT').OnEvent('Click', Grav_SetStat)
  StationBox.Add('Button',, 'BRM').OnEvent('Click', Grav_SetStat)
  StationBox.Add('Button',, 'BLV').OnEvent('Click', Grav_SetStat)
  StationBox.Add('Button',, 'WNL').OnEvent('Click', Grav_SetStat)
  StationBox.Add('Text','ym','PS')
  StationBox.Add('Button',, 'co05').OnEvent('Click', Grav_SetStat) ; 09801
  StationBox.Add('Button',, 'co01').OnEvent('Click', Grav_SetStat) ; 01904
  StationBox.Add('Text','ym','Oregon')
  StationBox.Add('Button',, 'PTD').OnEvent('Click', Grav_SetStat)
  StationBox.Add('Button',, 'EUG').OnEvent('Click', Grav_SetStat)
  StationBox.Add('Text','ym','East WA')
  StationBox.Add('Button',, 'SPK').OnEvent('Click', Grav_SetStat)
  StationBox.Add('Button',, 'CDA').OnEvent('Click', Grav_SetStat)
  StationBox.Add('Text','ym','Cali')
  StationBox.Add('Button',, 'PAS').OnEvent('Click', Grav_SetStat)
  StationBox.Add('Button',, 'SAA+LGB').OnEvent('Click', Grav_SetStat)
  DateBox.Show()
}

Grav_SetDates(*) {
  global DateBox
  DateBox.Hide()
  global DateVars := DateBox.Submit()
  global jsScript
  jsScript := StrReplace(jsScript, 'SDplace', DateVars.StartDate)
  jsScript := StrReplace(jsScript, 'EDplace', DateVars.EndDate)
}

Grav_SetStat(GuiCtrl, *) {
  global StationBox
  StationBox.Hide()
  global jsScript
  station := GuiCtrl.text
  switch {
  case InStr(station, 'co12'): station := '04590'
  case InStr(station, 'KNG'): station := 'KNG'
  case InStr(station, 'co05'): station := '09801'
  case InStr(station, 'co01'): station := '01904'
  case InStr(station, 'SAA'): station := 'SAA'
  }
  A_Clipboard := StrReplace(jsScript, 'OHplace', station)
  Sleep(300)
  Send("+^j")
  Sleep(1000)
  Send('^v')
  Send('{Enter}')
  if(InStr(station, 'KNG')) {
    A_Clipboard := StrReplace(jsScript, 'OHplace', 'UDT')
    Sleep(1000)
    Send('^v')
    Send('{Enter}')
  } else if(InStr(station, 'SAA')) {
    A_Clipboard := StrReplace(jsScript, 'OHplace', 'LGB')
    Sleep(1000)
    Send('^v')
    Send('{Enter}')
  }
  Sleep(500)
  Send("+^j")
  Sleep(500)
  Send('{Home}')
}

Grav_StatVis(*) {
  global StationBox
  StationBox.Show()
}

Grav_Off(*) {
  global
  Hotkey("NumpadAdd", "Off")
  Hotkey("``", "Off")
  ActiveFlag &= 255-16
}


All_Off(foo) {
  try {
    Csv_Off(0)
  }
  try {
    PPt_Off(0)
  }
  try {
    Chk_Off(0)
  }
  try {
    JE_Off(0)
  }
  try {
    Grav_Off(0)
  }
}
