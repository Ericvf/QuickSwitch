#SingleInstance force
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

Title = QuickSwitch brought to you by Fexelein™
Windows = []

Menu, Tray, Icon, Shell32.dll, 206, 1
Menu, Tray, Add, %Title%, About

; ; Add to start-up
; SplitPath, A_Scriptname, , , , OutNameNoExt
; LinkFile=%A_StartupCommon%\%OutNameNoExt%.lnk
; IfNotExist, %LinkFile%
; FileCreateShortcut, %A_ScriptFullPath%, %LinkFile%
; SetWorkingDir, %A_ScriptDir%

; if not A_IsAdmin
; {
;      Run *RunAs "%A_ScriptFullPath%"
; }

GroupAdd, MSIE, ahk_class CabinetWClass
GroupAdd, MSIE, ahk_class EVERYTHING

Loop
{
	fullPath := ""

    WinWaitActive ahk_group MSIE
		WinGet, explorerWindow, ID, A

    WinWaitNotActive     

    ; https://docs.microsoft.com/en-us/windows/desktop/winmsg/about-window-classes
    ; #32770	The class for a dialog box.

    IfWinActive ahk_class #32770
	{
		WinGetClass, windowClass, ahk_id %explorerWindow%
		

		if (windowClass == "CabinetWClass")
		{
			explorerWindowPath := Explorer_GetPath(explorerWindow)
			if FileExist(explorerWindowPath) or SubStr(explorerWindowPath, 1, 5) = "file:" 
			{
				fullPath := explorerWindowPath 
			}
		}
		else if (windowClass == "EVERYTHING")
		{
			ControlGet, Selected_Items,List,Selected,SysListView321, ahk_id %explorerWindow%
			lines := StrSplit(Selected_Items, "`t")
			filePath := lines[2]
			fullPath = %filePath%
		}
		
		if (fullPath != "") and (FileExist(fullPath) or SubStr(fullPath, 1, 5) = "file:")
		{
			ControlGetText, DialogMode, Button1, A
			if (DialogMode == "Select Folder")
			{
				Send ^l
				Sleep, 5
				ControlSetText, Edit2, %fullPath%, A
				ControlSend, Edit2, {Enter}, A
			}
			else{
				ControlSetText, Edit1, %fullPath%, A
				Sleep, 5
				Send !o
			}
		}
	}
}
	
MenuHandler:
	SelectedPath := Windows[A_ThisMenuItemPos - 1]
	ControlSetText, Edit1, %SelectedPath%, A
	Sleep, 5
	Send !o
	Return

About:
	Return

#IfWinActive ahk_class #32770 
#o::
MButton::
	ShowMenu()
Return

ShowMenu() {
	global Title, Windows
	Menu, MyMenu, Add, %Title%, About
	Menu, MyMenu, Default, %Title%
	WinGet, WinList, List, , , Program Manager
	Index := 0
	Array := []

	loop, %WinList% {
		Current := WinList%A_Index%
		WinGetClass class, ahk_id %Current%

		IsDesktop := RegExMatch(class, "Progman|WorkerW")

		Path1 := Explorer_GetPath(Current)
		if !ErrorLevel and !IsDesktop and Path1 <> ""
		{
			Array.Push(Path1)
			WinGetTitle,WinTitle,ahk_id %Current%
			Menu, MyMenu, Add, %WinTitle%, MenuHandler
			
		}
	}

	Windows := Array

	Menu, MyMenu, Show, %A_CaretX%, %A_CaretY%
	Menu, MyMenu, DeleteAll
}

; https://autohotkey.com/board/topic/60985-get-paths-of-selected-items-in-an-explorer-window/
Explorer_GetWindow(hwnd="")
{
	; thanks to jethrow for some pointers here
    WinGet, process, processName, % "ahk_id" hwnd := hwnd? hwnd:WinExist("A")
    WinGetClass class, ahk_id %hwnd%
	
	if (process!="explorer.exe")
		return
	if (class ~= "(Cabinet|Explore)WClass")
	{
		for window in ComObjCreate("Shell.Application").Windows
			if (window.hwnd==hwnd)
				return window
	}
	else if (class ~= "Progman|WorkerW") 
		return "desktop" ; desktop found
}

Explorer_GetPath(hwnd="")
{
	if !(window := Explorer_GetWindow(hwnd))
		return ErrorLevel := "ERROR"
	if (window="desktop")
		return A_Desktop
	path := window.LocationURL
	path := RegExReplace(path, "ftp://.*@","ftp://")
	StringReplace, path, path, file:///
	StringReplace, path, path, /, \, All 
	
	; thanks to polyethene
	Loop
		If RegExMatch(path, "i)(?<=%)[\da-f]{1,2}", hex)
			StringReplace, path, path, `%%hex%, % Chr("0x" . hex), All
		Else Break
	return path
}