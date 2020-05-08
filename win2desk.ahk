; AutoHotkey Version: AutoHotkey 1.1
; Language:           English
; Platform:           Win10
; Author:             Bin Zhang <ahk.oix.cc>
; Short description:  Save and recover Google Chrome Windows 10 virtual desktop assignments
; Last Mod:           2020-05-05
;

#SingleInstance Force ; The script will Reload if launched while already running
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
#KeyHistory 0 ; Ensures user privacy when debugging is not needed
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

DetectHiddenWindows, On
SetTitleMatchMode, 2
Menu, Tray, Icon, % A_WinDir "\system32\netshell.dll", 86 ; Shows a world icon in the system tray
Menu, Tray, Tip, Win2Desk

; Globals
;
dataPath := "data\"
listFile := dataPath . "windows.txt"
backups := 20

if (not InStr(FileExist("C:\My Folder"), "D"))
    FileCreateDir, %dataPath%

refreshWindowList()
SetTimer, AutoSave, 300000

#Include %A_ScriptDir%\browser.ahk
#Include %A_ScriptDir%\desktop_switcher.ahk
#Include %A_ScriptDir%\user_config.ahk
return

AutoSave:
    refreshWindowList()
    saveWindowList()
return

;
;
;
refreshWindowList(notification = false)
{
    global windowList
    if (notification)
        TrayTip, Win2Desk [refresh], % "Acquiring windows information...",, 16
    windowList := getWindowList()
}
;
; Write list of "desktop@title"
;
saveWindowList(refreshList = false, notification = false)
{
    global dataPath, listFile, backups
    global DesktopCount
    global windowList
    updateGlobalVariables()

    if (refreshList or (not (IsObject(windowList) and IsObject(windowList.desktops) and windowList.desktops.Count() > 0))) {
        if (notification)
            TrayTip, Win2Desk [save], % "Acquiring windows information...",, 16
        OutputDebug, [save] need windowList, calling getWindowList()
        refreshWindowList()
    }
    desktops := windowList.desktops
    windows := windowList.windows
    deskCount := desktops.Count()
    windowCount := windows.Length()

    if (deskCount > 1) { ; more than 1 desktops
        Loop, %backups%
        {
            from := dataPath . "windows.bak" . backups - A_Index . ".txt"
            if (A_Index == backups)
                from := listFile
            to := dataPath . "windows.bak" . backups - A_Index + 1 . ".txt"
            FileCopy, %from%, %to%, true
        }
        file := FileOpen(listFile, "w", "UTF-8")
        for idx, window in windows {
            ; line = % desktopNum "@" process "~" sClass "=" title "->" url "`r`n"
            line = % window.desktop "@" window.url "`r`n"
            file.Write(line)
        }
        file.Close()
        if (notification)
            TrayTip, Win2Desk [save], % windowCount . " windows on " . deskCount . " desktops.",, 16
        OutputDebug, [save] desktops: %deskCount% of %DesktopCount% windows: %windowCount%
    } else {
        if (notification)
            TrayTip, Win2Desk [save], % "All " . windowCount . " windows on the same desktop, nothing to save.",, 17
        OutputDebug, [save skipped] desktops: %deskCount% of %DesktopCount% windows: %windowCount%
    }
}

;
;
;
getWindowList()
{
    ; static running
    windows := []
    desktops := {}
    ; winIDList contains a list of windows IDs ordered from the top to the bottom for each desktop.
    WinGet winIDList, list
    Loop % winIDList {
        windowID := % winIDList%A_Index%
        WinGet, process, ProcessName, ahk_id %windowId%
        WinGetClass, sClass, ahk_id %windowId%
        WinGetTitle, title, ahk_id %windowId%
        ; WinGetText, text, ahk_id %windowId%
        if (process = "chrome.exe" and sClass = "Chrome_WidgetWin_1" and title) {
            desktopNum := getDesktopNumber(windowId)
            url := getBrowserUrlbyId(windowId) ; this step is slow
            if (desktopNum > 0 and url) {
                windows.push({ id: windowID, desktop: desktopNum, process: process, class: sClass, title: title, url: url })
                ; OutputDebug, [detected] window id: %windowID%, desktop: %desktopNum%, process: %process%, class: %sClass%, title: %title%, url: %url%
                if (not desktops.HasKey(desktopNum))
                    desktops[desktopNum] := true
            }
        }
    }
    return { desktops: desktops, windows: windows }
}

;
; Read list and apply
;
applyWindowList(refreshList = false, notification = false)
{
    global listFile
    global DesktopCount
    global windowList
    updateGlobalVariables()

    if (refreshList or (not (IsObject(windowList) and IsObject(windowList.desktops) and windowList.desktops.Count() > 0))) {
        if (notification)
            TrayTip, Win2Desk [apply], % "Acquiring windows information...",, 16
        OutputDebug, [apply] need windowList, calling getWindowList()
        refreshWindowList()
    }
    ; desktops := windowList.desktops
    curWindows := windowList.windows
    ; deskCount := desktops.Count()
    ; windowCount := windows.Length()

    windows := []
    desktops := {}
    Loop, Read, %listFile%
    {
        word_array := StrSplit(A_LoopReadLine, "@",, 2)  ; Omits periods.
        desktopNum := word_array[1]
        url := word_array[2]
        for idx, window in curWindows {
            ; hwnd := WinExist(title)
            if (window.url == url) {
                windows.push({ id: window.id, desktop: desktopNum, url: url })
                break
            }
        }
        if (not desktops.HasKey(desktopNum))
            desktops[desktopNum] := true
    }
    deskCount := desktops.Count()
    deskMax := desktops.Length()
    windowCount := 0
    for idx, window in windows {
        if (window.desktop <= DesktopCount) {
            windowCount++
            DllCall(MoveWindowToDesktopNumberProc, UInt, window.id, UInt, window.desktop - 1)
            ; OutputDebug, % "[applied] window: " . window.id . " url: " . window.url . " desktop: " . window.desktop
        }
    }
    if (notification)
        TrayTip, Win2Desk [apply], % windowCount . " windows set to " . deskCount . " desktops" . (deskMax > DesktopCount ? (" (only " DesktopCount . " of " . deskMax .  " desktops available, please create and apply again).") : "."),, (deskMax > DesktopCount ? 17 : 16)
    OutputDebug, [apply] desktops: max %deskMax% of %DesktopCount% windows: %windowCount%
}

;
;
;
getDesktopNumber(windowId)
{
    global DesktopCount, IsWindowOnDesktopNumberProc
    Loop, %DesktopCount% {
        n := A_Index - 1 ; Desktops start at 0, while in script it's 1
        windowIsOnDesktop := DllCall(IsWindowOnDesktopNumberProc, UInt, windowId, UInt, n)
        if (windowIsOnDesktop == 1) {
            return A_Index
        }
    }
    return -1
}

;
;
;
getBrowserUrlbyId(hwnd) {
    global ModernBrowsers
    WinGetClass, sClass, ahk_id %hwnd%
    If sClass not in % ModernBrowsers
        return
    return GetAddressBarUrl(Acc_ObjectFromWindow(hwnd))
}

;
;
;
getUrl(sURL) {
    ; https://mathiasbynens.be/demo/url-regex
    ; @imme_emosol
    ; @^(https?|ftp|chrome-extension)://(-\.)?([^\s/?\.#-]+\.?)+(/[^\s]*)?$@iS
    ; @stephenhay
    ; @(https?|ftp)://(-\.)?([^\s/?\.#-]+\.?)+(/[^\s]*)?$@iS
    ; @^(https?|ftp|chrome-extension)://[^\s/$.?#].[^\s]*$@iS

    ; i - Case-insensitive
    ; S - PCRE performance cache
    ; no ^, so also matches partial like (The Great Suspender support):
    ; chrome-extension://klbibkeccnjlkjkiokjodocebajanakg/suspended.html
    ; #ttl=Get%20the%20URL%20of%20the%20current%20(active)%20browser%20tab%20-%20AutoHotkey
    ; %20Community&pos=0&uri=https://www.autohotkey.com/boards/viewtopic.php?t=3702)
    ;
    ; See https://www.autohotkey.com/docs/misc/RegEx-QuickRef.htm#Options
    RegExMatch(sURL, "iS)(https?|ftp)://(-\.)?([^\s/?\.#-]+\.?)+(/[^\s]*)?$", url)
    return url
}

;
; Gui, Add, ListView, x2 y0 w400 h500, Process Name|Command Line
; for proc in ComObjGet("winmgmts:").ExecQuery("Select * from Win32_Process")
;     LV_Add("", proc.Name, proc.CommandLine)
; Gui, Show,, Process List
; ; Win32_Process: http://msdn.microsoft.com/en-us/library/aa394372.aspx
;

;
; https://referencesource.microsoft.com/#system/services/monitoring/system/diagnosticts/ProcessManager.cs
; https://stackoverflow.com/q/38205375/11384939
; https://wj32.org/wp/2012/12/12/enumwindows-no-longer-finds-metromodern-ui-windows-a-workaround-2/
; https://social.msdn.microsoft.com/Forums/windowsdesktop/en-US/7e25e104-36cb-41ac-8f36-0e4c6b6146a3/finding-hwnd-of-metro-app-using-win32-api?forum=windowsgeneraldevelopmentissues
; https://stackoverflow.com/q/31801402/11384939
; https://stackoverflow.com/a/16975012/11384939
; https://github.com/x64dbg/ScyllaHide
;

;
; ; BOOL EnumWindows(
; ;   WNDENUMPROC lpEnumFunc,
; ;   LPARAM lParam
; ; );
; EnumWindowsHandler := RegisterCallback("EnumWindowsProc", "Fast")
; DllCall("EnumWindows", "Ptr", EnumWindowsHandler, "Ptr", 0)
;
; ; BOOL CALLBACK EnumWindowsProc(
; ;   HWND hwnd,
; ;   LPARAM lParam
; ; );
; EnumWindowsProc(hwnd, lParam)
; {
;     WinGetTitle, title, ahk_id %hwnd%
;     if (title) {
;         OutputDebug, [EnumWindows] hwnd: %hwnd% title: %title%
;     }
;     return true
; }
;

;
; ; BOOL EnumChildWindows(
; ;   HWND        hWndParent,
; ;   WNDENUMPROC lpEnumFunc,
; ;   LPARAM      lParam
; ; );
; EnumChildHandler := RegisterCallback("EnumChildProc", "Fast")
; DllCall("EnumChildWindows", "Ptr", 0, "Ptr", EnumChildHandler, "Ptr", 0)

; ; BOOL CALLBACK EnumChildProc(
; ;   _In_ HWND   hwnd,
; ;   _In_ LPARAM lParam
; ; );
; EnumChildProc(hwnd, lParam)
; {
;     WinGetTitle, title, ahk_id %hwnd%
;     if (title) {
;         OutputDebug, [EnumChild] hwnd: %hwnd% title: %title%
;     }
;     return true
; }
;
