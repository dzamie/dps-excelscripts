^`::ExitApp

F23::AppsKey

/*
global ClicksEaten := 0

!^LButton::Click "Right"
RButton:: {
  global
  ClicksEaten += 1
  return
}
^RButton::{
  global
  ClicksEaten += 1
  return
}
+RButton::{
  global
  ClicksEaten += 1
  return
}
*/

^NumpadAdd:: {
  ; Display choice box
  global
  GuiBox.Show()
}

CoordMode "Mouse", "Screen"
CoordMode "Pixel", "Screen"
global StepCt := 1
; Anch, Bend, Boise, Eug, Hon, Kon, Mont, Ore, SLC
global Cities := [348, 435, 476, 510, 558, 600, 725, 813, 855]
global CityIndex := 1
global MorFound := false
global ActiveFlag := 0

Csv_Active(Button, Info) {
  global
  if(ActiveFlag & 2 > 0) {
    PPt_Off(0)
  } else if(ActiveFlag & 4 > 0) {
    Mor_Off(0)
  }
  ActiveFlag |= 1
  StepCt := 1
  Hotkey "NumpadAdd", Csv_Next, "On"
  Hotkey "^Space", Csv_Reset, "On"
  Hotkey "``", Csv_Off, "On"
  GuiBox.Hide()
  return true
}
PPt_Active(Button, Info) {
  global
  if(ActiveFlag & 1 > 0) {
    Csv_Off(0)
  } else if(ActiveFlag & 4 > 0) {
    Mor_Off(0)
  }
  ActiveFlag |= 2
  Hotkey "^1", PPt_1, "On"
  Hotkey "^2", PPt_2, "On"
  Hotkey "^3", PPt_3, "On"
  Hotkey "^4", PPt_4, "On"
  Hotkey "``", PPt_Off, "On"
  GuiBox.Hide()
  return true
}
Mor_Active(Button, Info) {
  global
  if(ActiveFlag & 1 > 0) {
    Csv_Off(0)
  } else if(ActiveFlag & 2 > 0) {
    PPt_Off(0)
  }
  ActiveFlag |= 4
  CityIndex := 1
  Hotkey "BS", Mor_Copy, "On"
  Hotkey "NumpadAdd", Mor_Next, "On"
  Hotkey "NumpadMult", Mor_Reset, "On"
  Hotkey "+NumpadAdd", Mor_Auto, "On"
  Hotkey "+NumpadMult", Mor_AutoStop, "On"
  Hotkey "``", Mor_Off, "On"
  GuiBox.Hide()
  return true
}
Active_Check(Button, Info) {
  global
  GuiBox.Hide()
  CheckString := "CSV: " . ((ActiveFlag & 1 > 0) ? "On" : "Off")
  CheckString .= "`nMonthly: " . ((ActiveFlag & 2 > 0) ? "On" : "Off")
  CheckString .= "`nPBP: " . ((ActiveFlag & 4 > 0) ? "On" : "Off")
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
global MORBtn := GuiBox.AddButton(, "PBP MOR Helper")
MORBtn.OnEvent("Click", Mor_Active)
global MORBtn := GuiBox.AddButton(, "Active Scripts")
MORBtn.OnEvent("Click", Active_Check)

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

Csv_Reset(foo) {
  global
  StepCt := 1
  MsgBox "Steps reset!"
}

Csv_Off(foo) {
  global
  Hotkey "NumpadAdd", "Off"
  Hotkey "^Space", "Off"
  Hotkey "``", "Off"
  ActiveFlag &= (255-1)
}

Mor_Copy(foo) {
  global
  Click 2100, 270
  Sleep 100
  Click 1975, 268
  Sleep 100
  Send "{Shift down}{Ctrl down}"
  Sleep 100
  Send "{Right}{Up}"
  Sleep 100
  Send "{Down}"
  Sleep 100
  Send "{Shift up}{Ctrl up}"
  Sleep 100
  Send "^c"
  Sleep 100
  Send "{Alt down}"
  Sleep 100
  Send "{Tab}"
  Sleep 100
  Send "{Right}"
  Sleep 100
  Send "{Alt up}"
  Sleep 200
  Send "{Ctrl down}"
  Sleep 100
  Send "v"
  Sleep 100
  Send "{Up}"
  Sleep 100
  Send "{Down}"
  Sleep 100
  Send "{Ctrl up}"
  Sleep 100
  Send "{Down}"
  Sleep 100
  Send "!{Tab}"
  Sleep 200
  Send "{Esc}"
  Sleep 100
  Send "^w"
}

Mor_Next(foo) {
  global
  xf := "xf"
  yf := "yf"
  flag := false
  MorFound := false
  Click 326, 304
  Sleep 100
  Click 326, Cities[CityIndex]
  Sleep 100
  While(!flag) {
    flag := PixelSearch(&xf,&yf,785,300,800,315,0x555555, 64)
    Sleep 200
  }
  Click 732, 304
  Sleep 100
  Click 732, 485 ; Deposit Individual Transaction Detail
  Sleep 200
  Click 1670, 304
  CityIndex := CityIndex + 1
  
  flag := false
  Sleep 100
  While(!flag) { ; looking for reset Generate
    flag := PixelSearch(&xf,&yf, 1590, 293,1600,303,0x629641, 8)
    Sleep 200
  }
  flag := false
  flag := PixelSearch(&xf,&yf,137,520,147,530,0xf3f3f3, 8)
  If(flag) { ; 1+ records found
    MorFound := true
    Click 1328, 372 ; download xslx
    flag := false
    Sleep 200
    While(!flag) { ; looking for excel icon in downloads
		   ; this is WIDE on the x-axis.
      flag := PixelSearch(&xf,&yf,1400,120,1505,140,0x107c41, 16)
      Sleep 200
    }
    Click 1534, 160, 0 ; extra time needed to get it to trigger
    Sleep 500
    Click 1534, 170, 2 ; doubleclick downloaded file
    Sleep 1500
    Click 1328, 135 ; refocus on browser
  } Else {
    MorFound := false
  }
}

Mor_Auto(foo) {
  global
  xf := "xf"
  yf := "yf"
  While(CityIndex <= Cities.Length) {
    Mor_Next(foo)
    Sleep 200
    If(MorFound) {
      flag := false
      While(!flag) { ; waiting for sheet to load
        flag := PixelSearch(&xf,&yf,3292,245,3297,250,0x000000,8)
        Sleep 200
      }
      Sleep 300
      Mor_Copy(foo)
    }
    Sleep 200
  }
  MsgBox "Finished!"
}

Mor_AutoStop(foo) {
  global
  CityIndex := Cities.Length + 1
}

Mor_Reset(foo) {
  global
  CityIndex := 1
  MsgBox "City Index reset to Anchorage"
}

Mor_Off(foo) {
  global
  Hotkey "BS", "Off"
  Hotkey "NumpadAdd", "Off"
  Hotkey "NumpadMult", "Off"
  Hotkey "+NumpadAdd", "Off"
  Hotkey "+NumpadMult", "Off"
  Hotkey "``", "Off"
  ActiveFlag &= (255-4)
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