#Requires AutoHotkey v2.0
#SingleInstance Force

; Global state for recursive child enumeration
global gEnumFilter := ""
global gEnumMatches := []
global gEnumCallback := 0

Main() {
    iniPath := EnsureConfig()
    config := LoadConfig(iniPath)
    if !config.Has("ClassNN") && !config.Has("Process") {
        MsgBox "No window configuration found in " iniPath
        return
    }

    classPrompt := InputBox("Enter (part of) the class name to search for:", "Class Lookup")
    if classPrompt.Result != "OK" || classPrompt.Value = "" {
        MsgBox "No class provided. Exiting."
        return
    }

    targetHwnd := GetTargetWindow(config)
    if !targetHwnd {
        MsgBox "Target window not found. Check showClasses.ini"
        return
    }

    matches := FindChildWindows(targetHwnd, classPrompt.Value)

    dateStamp := FormatTime(, "yyyy-MM-dd")
    outPath := A_ScriptDir "\" dateStamp "-output.txt"
    WriteResults(outPath, matches)
    MsgBox Format("Found {1} matching controls. Details written to:`n{2}", matches.Length, outPath)

    MsgBox "Hover any control and press Ctrl+Alt+D to log its class. This script will keep running until you exit it." "`nOutput file: " dateStamp "-cloutput.txt"
}

EnsureConfig() {
    iniPath := A_ScriptDir "\showClasses.ini"
    if !FileExist(iniPath) {
        IniWrite("Qt673QWindowToolSaveBits", iniPath, "Window", "ClassNN")
        IniWrite("Camera Hub.exe", iniPath, "Window", "Process")
    }
    return iniPath
}

LoadConfig(iniPath) {
    config := Map()
    classNN := IniRead(iniPath, "Window", "ClassNN", "")
    process := IniRead(iniPath, "Window", "Process", "")
    if classNN != ""
        config["ClassNN"] := classNN
    if process != ""
        config["Process"] := process
    return config
}

GetTargetWindow(config) {
    search := ""
    if config.Has("ClassNN")
        search .= "ahk_class " config["ClassNN"]
    if config.Has("Process") {
        if search != ""
            search .= " "
        search .= "ahk_exe " config["Process"]
    }
    if search = ""
        return 0
    return WinExist(search)
}

FindChildWindows(hwnd, filterClass) {
    global gEnumFilter, gEnumMatches, gEnumCallback
    gEnumFilter := StrLower(filterClass)
    gEnumMatches := []
    gEnumCallback := CallbackCreate("EnumChildProc", "Fast")
    EnumChildrenRecursive(hwnd)
    CallbackFree(gEnumCallback)
    gEnumCallback := 0
    return gEnumMatches
}

EnumChildrenRecursive(hwnd) {
    global gEnumCallback
    if !gEnumCallback
        return
    DllCall("EnumChildWindows", "ptr", hwnd, "ptr", gEnumCallback, "ptr", 0)
}

EnumChildProc(childHwnd, lParam) {
    global gEnumFilter, gEnumMatches
    class := ""
    title := ""
    try class := WinGetClass("ahk_id " childHwnd)
    try title := WinGetTitle("ahk_id " childHwnd)

    if gEnumFilter = "" || InStr(StrLower(class), gEnumFilter)
        gEnumMatches.Push(Map("HWND", Format("0x{1:X}", childHwnd), "Class", class, "Title", title))

    EnumChildrenRecursive(childHwnd)
    return true
}

WriteResults(path, results) {
    header := "Control Class Scan - " FormatTime(, "yyyy-MM-dd HH:mm:ss")
    if results.Length = 0 {
        FileDelete(path)
        FileAppend(header "`nNo matches found.`n", path, "UTF-8")
        return
    }
    lines := []
    lines.Push(header)
    for item in results {
        line := Format("Class: {1}`tHWND: {2}`tTitle: {3}", item["Class"], item["HWND"], item["Title"])
        lines.Push(line)
    }
    FileDelete(path)
    FileAppend(JoinLines(lines) "`n", path, "UTF-8")
}

CaptureUnderCursor(*) {
    KeyWait("Ctrl", "T0.1") ; ensure modifier state is clean after hotkey
    MouseGetPos(&mx, &my, &winHwnd, &ctrlHwnd)
    target := ctrlHwnd ? ctrlHwnd : winHwnd
    if !target {
        MsgBox "Could not determine the hovered control."
        return
    }

    class := ""
    title := ""
    try class := WinGetClass("ahk_id " target)
    try title := WinGetTitle("ahk_id " winHwnd)

    dateStamp := FormatTime(, "yyyy-MM-dd")
    path := A_ScriptDir "\" dateStamp "-cloutput.txt"
    timestamp := FormatTime(, "HH:mm:ss")
    line := Format("{1}`tClass: {2}`tWindow: {3}`tPos: ({4}, {5})", timestamp, class, title, mx, my)
    FileAppend(line "`n", path, "UTF-8")
    MsgBox "Captured class " class "`nLogged to " path
}

JoinLines(arr) {
    out := ""
    for index, value in arr {
        if index > 1
            out .= "`n"
        out .= value
    }
    return out
}

^!d::CaptureUnderCursor()

Main()
