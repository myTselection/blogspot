; On Screen Keyboard only numeric, for touchscreen
; Windows OnScreenKeyboard shortcut: Win + Ctrl + O
; Author:         myTselection.blogspot.com
; Open keypad when Windows Security Windows Hello PIN popup appears


#SingleInstance,force
#Persistent

k_WindowWidth = 280
k_WindowHeight = 370

Gosub,TRAYMENU


Gui, +LastFound +AlwaysOnTop +ToolWindow +Caption +Border +E0x08000000

k_FontSize = 24
Gui, Font, s%k_FontSize%, Arial

Gui, Add, Button, % "x" 5+(1-1)*90 " y5 w90 h90 v" 1 " gButton", 1
Gui, Add, Button, % "x" 5+(2-1)*90 " y5 w90 h90 v" 2 " gButton", 2
Gui, Add, Button, % "x" 5+(3-1)*90 " y5 w90 h90 v" 3 " gButton", 3

Gui, Add, Button, % "x" 5+(4-3-1)*90 " y95 w90 h90 v" 4 " gButton", 4
Gui, Add, Button, % "x" 5+(5-3-1)*90 " y95 w90 h90 v" 5 " gButton", 5
Gui, Add, Button, % "x" 5+(6-3-1)*90 " y95 w90 h90 v" 6 " gButton", 6

Gui, Add, Button, % "x" 5+(7-6-1)*90 " y185 w90 h90 v" 7 " gButton", 7
Gui, Add, Button, % "x" 5+(8-6-1)*90 " y185 w90 h90 v" 8 " gButton", 8
Gui, Add, Button, % "x" 5+(9-6-1)*90 " y185 w90 h90 v" 9 " gButton", 9

Gui, Font, s08, Arial
Gui, Add, Button, % "x" 5+(7-6-1)*90 " y275 w90 h90 v" ← " gButton", <
Gui, Font, s24, Arial
Gui, Add, Button, % "x" 5+(8-6-1)*90 " y275 w90 h90 v" 0 " gButton", 0
Gui, Font, s10, Arial
Gui, Add, Button, % "x" 5+(9-6-1)*90 " y275 w90 h90 v" Enter " gButton", Enter

SetTimer, OpenKeypad, 500
return


OpenKeypad:
  IfWinActive, Windows Security 
  {
	GoSub ShowKeypad
	WinActivate, ,Windows Security
	winwaitclose, Windows Security
	GoSub HideKeypad
	return
	}
return



ShowKeypad:
	SysGet k_WorkArea, Monitor, 1
	k_WindowY = %k_WorkAreaBottom%
	k_WindowY -= 100
	k_WindowY -= %k_WindowHeight%
	Gui, Show, x1200 y%k_WindowY% w280 h370, AlwaysOnTop NoActivate
return

HideKeypad:
	Gui Cancel
return



TRAYMENU:
Menu,TRAY,NoStandard 
Menu,TRAY,DeleteAll 
Menu,TRAY,Add,&Show,ShowKeypad
Menu,TRAY,Add,&Hide,HideKeypad
Menu,TRAY,Add
Menu,TRAY,Add,E&xit,EXIT
Menu,TRAY,Default,&Show
Menu,Tray,Tip,PulseSecure keypad
return

Button:
	if (A_GuiControl = "<") {
		Send {Backspace}
	} else {
		SendInput, % "{" A_GuiControl "}"
	}
Return

GuiClose:
	Gui Cancel
Return

EXIT:
ExitApp
return