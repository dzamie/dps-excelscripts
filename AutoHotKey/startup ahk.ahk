#Requires AutoHotkey v2.0

CoordMode('Mouse','Screen')
CoordMode('Pixel','Screen')

LongWait := 5000

Sleep(LongWait) ; wait for Edge to load

; Edge logins
WinActivate("Edge")
WinSetAlwaysOnTop(1, "Edge") ; so focus isn't lost
Sleep(400)
Click('Left', 145, 20) ; tab 1
Sleep(400)
Click('Left', 455, 620) ; NS login
Sleep(400)
Click('Left', 400, 20) ; tab 2
Sleep(400)
Click('Left', 1210, 555) ; PP login
WinSetAlwaysOnTop(0, "Edge") ; focus no longer needed

Sleep(LongWait) ; wait for Excel/Outlook to finish loading

; Excel activation
WinActivate("Excel") ; doesn't need the force-on-top
Sleep(400)

; Outlook folders
WinActivate("Outlook")
WinSetAlwaysOnTop(1, "Outlook") ; so focus isn't lost
Sleep(400)
Click('Left', 70, 810) ; Public Folders
Sleep(500)
Click('Left', 70, 855) ; Public Folders -> Favorites
Sleep(300)
Click('Left', 80, 300) ; WSS auto-reports

; Disb watching
Flag := 0
while (Flag == 0) {
    Sleep(1000)
    Flag := PixelSearch(&x, &y, 245, 345, 255, 355, 0x2065B0, 0x08)
}

Click('Left', 80, 300) ; close the WSS folder
WinSetAlwaysOnTop(0, "Outlook") ; done clicking around here

Sleep(1000) ; hopefully wait for disbs to finish filtering in

; Run auto wss
PS1Loc := "[AUTO WSS PS1 LOCATION]"
Run('powershell.exe -ExecutionPolicy Bypass -File "' . PS1Loc . '"')
Sleep(500)
WinActivate("ahk_class CabinetWClass") ; switch focus to explorer.exe to force Excel to open in background
