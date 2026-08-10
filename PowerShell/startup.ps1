# Morning ritual:
# 0. Open todo.xlsx, explorer.exe:Daily Shortcuts, Edge, and Outlook
# 1. Run ahk manager.ahk
# 2. Navigate to NS, PP logins in Edge
# 3. Briefly focus todo once it loads
# 4. Wait for Outlook to sort disb reports into disb folder
# 5. Run wss full starter.ps1
# 6. Focus explorer.exe

$timeout = 60

# 0-2. opening stuff
$desktop = "[DESKTOP]"
start "${desktop}\Daily Todo 2026.xlsx"
start "${desktop}\Daily Shortcuts\"
start "${desktop}\Daily Shortcuts\ahk manager.ahk"
start outlook.exe

# these URLS are kinda ugly, ain't they?
start msedge.exe "[NS LOGIN PAGE] [PP LOGIN PAGE]"

# begin the auto-login/excel/outlook/auto wss starter
& "[STARTUP AHK.AHK]"
