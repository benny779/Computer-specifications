#SingleInstance force
#NoTrayIcon
SetTitleMatchMode, 2


~VK13::
if WinActive(" - Connected - DameWare Mini Remote Control")
	WinClose, - Connected - DameWare Mini Remote Control
else if WinActive(" - Connecting - DameWare Mini Remote Control")
	WinClose, - Connecting - DameWare Mini Remote Control
else if WinActive(" - Disconnected - DameWare Mini Remote Control")
	WinClose, - Disconnected - DameWare Mini Remote Control
else if WinActive("DameWare Mini Remote Control - Error - ")
	WinClose, DameWare Mini Remote Control - Error - 
else if WinActive("Helpdesk")
	if (A_ThisHotkey=A_PriorHotkey && A_TimeSincePriorHotkey<250)
		WinClose, Helpdesk
return

^VK13::
ExitApp

/* GET VIRTUAL KEY CODE
#InstallKeybdHook
Sleep 5000
KeyHistory
Pause