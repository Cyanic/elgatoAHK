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

    MsgBox "Hover any control and press Ctrl+Alt+D to log its AutomationId and ClassName." "`nOutput file: " dateStamp "-cloutput.txt"
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
    uia := GetUIAutomation()
    if !uia {
        MsgBox "UI Automation interface not available on this system."
        return
    }

    MouseGetPos(&mx, &my, &winHwnd, &ctrlInfo, 2)
    winHwnd := NormalizeHwnd(winHwnd)
    ctrlHwnd := NormalizeHwnd(ctrlInfo)
    target := ctrlHwnd ? ctrlHwnd : winHwnd
    if !target {
        MsgBox "Could not determine the hovered control."
        return
    }

    winClass := ""
    try winClass := WinGetClass("ahk_id " winHwnd)

    element := UIAElementFromHandle(uia, target)
    if !element
        element := UIAElementFromPoint(uia, mx, my)

    automationId := ""
    className := ""
    if element {
        automationId := UIAGetProperty(element, 30011)
        className := UIAGetProperty(element, 30012)
        UIARelease(element)
    }

    if automationId = ""
        automationId := "<none>"
    if className = ""
        className := "<none>"

    dateStamp := FormatTime(, "yyyy-MM-dd")
    path := A_ScriptDir "\" dateStamp "-cloutput.txt"
    timestamp := FormatTime(, "HH:mm:ss")

    line := Format("{1}`tAutomationId: {2}`tClassName: {3}`tWinClass: {4}`tPos: ({5}, {6})", timestamp, automationId, className, winClass, mx, my)
    FileAppend(line "`n", path, "UTF-8")
    MsgBox "Captured AutomationId: " automationId "`nClassName: " className "`nLogged to " path
}

NormalizeHwnd(value) {
    if value is Integer
        return value
    if value is String {
        if RegExMatch(value, "^0x[0-9A-Fa-f]+$")
            return value + 0
        return 0
    }
    return 0
}

UIAElementFromHandle(uia, hwnd) {
    elementPtr := 0
    hr := ComCall(7, uia, "ptr", hwnd, "ptr*", &elementPtr)
    return hr = 0 ? elementPtr : 0
}

UIAElementFromPoint(uia, x, y) {
    elementPtr := 0
    point := Buffer(8, 0)
    NumPut("int", x, point, 0)
    NumPut("int", y, point, 4)
    hr := ComCall(8, uia, "int64", NumGet(point, 0, "int64"), "ptr*", &elementPtr)
    return hr = 0 ? elementPtr : 0
}

UIAGetProperty(elementPtr, propertyId) {
    if !elementPtr
        return ""
    variant := Buffer(24, 0)
    vtable := NumGet(elementPtr, 0, "ptr")
    fn := NumGet(vtable, (42 + (propertyId = 30012 ? 0 : 0)) * A_PtrSize, "ptr") ; index 42 assumed
    hr := DllCall(fn, "ptr", elementPtr, "int", propertyId, "ptr", variant.Ptr, "int")
    if hr != 0
        return ""
    value := UIAVariantToText(variant)
    DllCall("OleAut32\VariantClear", "ptr", variant.Ptr)
    return value
}

UIARelease(elementPtr) {
    if !elementPtr
        return
    vtable := NumGet(elementPtr, 0, "ptr")
    fn := NumGet(vtable, 2 * A_PtrSize, "ptr")
    DllCall(fn, "ptr", elementPtr, "uint")
}

UIAVariantToText(varBuf) {
    vt := NumGet(varBuf, 0, "ushort")
    switch vt {
        case 0, 1:
            return ""
        case 8:
            bstr := NumGet(varBuf, 8, "ptr")
            return bstr ? StrGet(bstr, "UTF-16") : ""
        case 3:
            return Format("{1}", NumGet(varBuf, 8, "int"))
        case 7:
            return Format("{1}", NumGet(varBuf, 8, "double"))
        default:
            return ""
    }
}

GetUIAutomation() {
    static uia := ""
    if IsObject(uia)
        return uia

    try uia := ComObject("UIAutomationClient.CUIAutomation")
    catch {
        try uia := ComObject("{FF48DBA4-60EF-4201-AA87-54103EEF594E}", "{30CBE57D-D9D0-452A-AB13-7AC5AC4825EE}")
        catch {
            uia := ""
        }
    }
    return uia
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
