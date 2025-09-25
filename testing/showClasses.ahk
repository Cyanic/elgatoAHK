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

    MouseGetPos(&mx, &my, &winHwnd, &ctrlInfo)
    winHwnd := NormalizeHwnd(winHwnd)
    if !winHwnd {
        MsgBox "Could not determine the hovered control."
        return
    }

    ctrlHwnd := NormalizeHwnd(ctrlInfo, winHwnd)
    if !ctrlHwnd {
        ctrlHwnd := NormalizeHwnd(WindowFromPoint(mx, my), winHwnd)
        if ctrlHwnd = winHwnd
            ctrlHwnd := 0
    }
    if ctrlHwnd {
        static GA_ROOT := 2 ; GetAncestor(..., GA_ROOT)
        root := DllCall("GetAncestor", "ptr", ctrlHwnd, "uint", GA_ROOT, "ptr")
        if root != winHwnd
            ctrlHwnd := 0
    }

    winClass := ""
    try winClass := WinGetClass("ahk_id " winHwnd)

    element := 0
    if ctrlHwnd
        element := UIAElementFromHandle(uia, ctrlHwnd)
    if !element
        element := UIAElementFromPoint(uia, mx, my)
    if !element
        element := UIAElementFromHandle(uia, winHwnd)

    automationId := ""
    className := ""
    if element {
        automationId := UIAGetProperty(element, 30011)
        className := UIAGetProperty(element, 30012)
        UIARelease(element)
    }

    if automationId = ""
        automationId := "<none>"
    if className = "" || className = "#32769" {
        nativeClass := GetWindowClassName(ctrlHwnd ? ctrlHwnd : winHwnd)
        if nativeClass != ""
            className := nativeClass
    }
    if className = ""
        className := "<none>"

    dateStamp := FormatTime(, "yyyy-MM-dd")
    path := A_ScriptDir "\" dateStamp "-cloutput.txt"
    timestamp := FormatTime(, "HH:mm:ss")

    line := Format("{1}`tAutomationId: {2}`tClassName: {3}`tWinClass: {4}`tWinHWND: {5}`tCtrlHWND: {6}`tPos: ({7}, {8})",
        timestamp,
        automationId,
        className,
        winClass,
        Format("0x{1:X}", winHwnd),
        ctrlHwnd ? Format("0x{1:X}", ctrlHwnd) : "<none>",
        mx,
        my)
    FileAppend(line "`n", path, "UTF-8")
    MsgBox "Captured AutomationId: " automationId "`nClassName: " className "`nLogged to " path
}

NormalizeHwnd(value, baseWin := 0) {
    if value is Integer
        return value
    if value is String {
        if RegExMatch(value, "^0x[0-9A-Fa-f]+$")
            return value + 0
        if baseWin && value != "" {
            try {
                hwnd := ControlGetHwnd(value, "ahk_id " baseWin)
                if hwnd
                    return hwnd
            }
        }
        return 0
    }
    return 0
}

WindowFromPoint(x, y) {
    point := Buffer(8, 0)
    NumPut("int", x, point, 0)
    NumPut("int", y, point, 4)
    return DllCall("WindowFromPoint", "int64", NumGet(point, 0, "int64"), "ptr")
}

GetWindowClassName(hwnd) {
    if !hwnd
        return ""
    buf := Buffer(512 * 2, 0)
    len := DllCall("GetClassNameW", "ptr", hwnd, "ptr", buf, "int", 512)
    return len ? StrGet(buf, "UTF-16") : ""
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
    variantSize := (A_PtrSize = 8) ? 24 : 16
    static GET_CURRENT_PROPERTY_VALUE := 10
    static GET_CURRENT_PROPERTY_VALUE_EX := 11

    attempts := [
        {index: GET_CURRENT_PROPERTY_VALUE, extra: []},
        {index: GET_CURRENT_PROPERTY_VALUE_EX, extra: [true]}
    ]

    for attempt in attempts {
        variant := Buffer(variantSize, 0)
        params := [attempt.index, elementPtr, "int", propertyId]
        for param in attempt.extra {
            params.Push("int")
            params.Push(param)
        }
        params.Push("ptr")
        params.Push(variant.Ptr)

        hr := ComCall(params*)
        if hr = 0 {
            value := UIAVariantToText(variant)
            DllCall("OleAut32\VariantClear", "ptr", variant.Ptr)
            if value != ""
                return value
        }
    }

    return ""
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
