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

Active_Check(Button, Info) {
  global
  GuiBox.Hide()
  CheckString := "CSV: " . ((ActiveFlag & 1 > 0) ? "On" : "Off")
  CheckString .= "`nMonthly: " . ((ActiveFlag & 2 > 0) ? "On" : "Off")
  CheckString .= "`nCheck: " . ((ActiveFlag & 4 > 0) ? "On" : "Off")
;  CheckString .= "`nClicks Eaten: " . ClicksEaten
  MsgBox CheckString
  return true
}
  

global GuiBox := Gui(, "Script Spawner",)
GuiBox.AddText(, "Run which script?")
global CSVBtn := GuiBox.AddButton(, "CSV Helper")
CSVBtn.OnEvent("Click", Csv_Active)
global PPTBtn := GuiBox.AddButton(, "Monthly Helper")
PPTBtn.OnEvent("Click", PPt_Active)
global ChkBtn := GuiBox.AddButton(, "Check Refund Helper")
ChkBtn.OnEvent("Click", Chk_Active)
global ScrBtn := GuiBox.AddButton(, "Active Scripts")
ScrBtn.OnEvent("Click", Active_Check)

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

PPt_Off(foo) {
  global
  Hotkey "^1", "Off"
  Hotkey "^2", "Off"
  Hotkey "^3", "Off"
  Hotkey "^4", "Off"
  Hotkey "``", "Off"
  ActiveFlag &= (255-2)
}

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
    A_Clipboard := ""
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
    A_Clipboard := ""
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
    Send "[geo string]{Tab}"
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
}
