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

    CoordMode("Mouse", "Screen") ; ensure cursor coordinates are in virtual screen space
    MouseGetPos(&logicalX, &logicalY, &winHwnd, &ctrlInfo)
    ; DPI virtualization can skew logical mouse coordinates on secondary
    ; displays. Capture the physical cursor when available so we can try both
    ; coordinate spaces for hit-testing.
    physX := logicalX
    physY := logicalY
    hasPhysical := GetPhysicalCursorPos(&physX, &physY)
    mx := hasPhysical ? physX : logicalX
    my := hasPhysical ? physY : logicalY
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

    candidates := []
    if ctrlHwnd
        candidates.Push(Map("Kind", "Handle", "Value", ctrlHwnd))
    candidates.Push(Map("Kind", "Point", "X", logicalX, "Y", logicalY))
    if hasPhysical && (physX != logicalX || physY != logicalY)
        candidates.Push(Map("Kind", "Point", "X", physX, "Y", physY))
    candidates.Push(Map("Kind", "Handle", "Value", winHwnd))

    automationId := ""
    className := ""
    uiaClass := ""
    fallbackAuto := ""
    fallbackClass := ""

    for candidate in candidates {
        element := 0
        pointX := candidate.Has("X") ? candidate["X"] : mx
        pointY := candidate.Has("Y") ? candidate["Y"] : my
        switch candidate["Kind"] {
            case "Handle":
                element := UIAElementFromHandle(uia, candidate["Value"])
            case "Point":
                element := UIAElementFromPoint(uia, pointX, pointY)
        }
        if !element
            continue

        refined := UIARefineElementAtPoint(uia, element, pointX, pointY)
        if refined && refined != element {
            UIARelease(element)
            element := refined
        }

        details := UIACollectElementDetails(uia, element, pointX, pointY)
        candAuto := details.Has("Auto") ? details["Auto"] : ""
        candClass := details.Has("Class") ? details["Class"] : ""
        UIARelease(element)

        if candClass != "" {
            if candClass != winClass
                uiaClass := candClass
            else if uiaClass = ""
                uiaClass := candClass
        }
        if fallbackClass = "" && candClass != ""
            fallbackClass := candClass
        if fallbackAuto = "" && candAuto != ""
            fallbackAuto := candAuto

        if candClass != "" && candClass != "#32769" && candClass != winClass {
            automationId := (candAuto != "") ? candAuto : automationId
            className := candClass
            break
        }

        if automationId = "" && candAuto != ""
            automationId := candAuto
        if className = "" && candClass != ""
            className := candClass
    }

    if className = "" && fallbackClass != ""
        className := fallbackClass
    if automationId = "" && fallbackAuto != ""
        automationId := fallbackAuto

    if automationId = ""
        automationId := "<none>"
    if className = "" || className = "#32769" {
        nativeClass := GetWindowClassName(ctrlHwnd ? ctrlHwnd : winHwnd)
        if nativeClass != ""
            className := nativeClass
    }
    if (className = "" || className = winClass) && (uiaClass != "")
        className := uiaClass
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

GetPhysicalCursorPos(&x, &y) {
    point := Buffer(8, 0)
    if !DllCall("GetCursorPos", "ptr", point.Ptr)
        return false
    x := NumGet(point, 0, "int")
    y := NumGet(point, 4, "int")
    return true
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
    ; Try passing POINT by value (preferred) and fall back to pointer form if
    ; the provider rejects it. Both paths are wrapped so UIA quirks on specific
    ; monitors/windows don't crash the script.
    try {
        hr := ComCall(8, uia, "int64", NumGet(point, 0, "int64"), "ptr*", &elementPtr)
        if hr = 0
            return elementPtr
    } catch {
    }
    try {
        hr := ComCall(8, uia, "ptr", point.Ptr, "ptr*", &elementPtr)
        if hr = 0
            return elementPtr
    } catch {
    }
    return 0
}

UIARefineElementAtPoint(uia, elementPtr, x, y) {
    if !uia || !elementPtr
        return 0

    refined := UIAHitTestElement(uia, elementPtr, x, y)
    return refined ? refined : elementPtr
}

UIAHitTestElement(uia, elementPtr, x, y) {
    if !elementPtr
        return 0

    rect := UIAGetBoundingRect(elementPtr)
    if rect && !UIAPointInRect(rect, x, y)
        return 0

    children := UIAFindChildren(uia, elementPtr)
    if !children
        return elementPtr

    count := UIAElementArrayLength(children)
    loop count {
        child := UIAElementArrayGet(children, A_Index - 1)
        if !child
            continue
        childRect := UIAGetBoundingRect(child)
        if childRect && UIAPointInRect(childRect, x, y) {
            deeper := UIAHitTestElement(uia, child, x, y)
            if deeper {
                if deeper != child
                    UIARelease(child)
                UIARelease(children)
                return deeper
            }
        }
        UIARelease(child)
    }
    UIARelease(children)
    return elementPtr
}

UIACollectElementDetails(uia, elementPtr, x, y) {
    info := Map("Auto", "", "Class", "")
    if !elementPtr
        return info

    auto := UIAGetProperty(elementPtr, 30011)
    class := UIAGetProperty(elementPtr, 30012)
    info["Auto"] := auto
    info["Class"] := class

    if (auto != "" && class != "")
        return info

    if auto = "" || class = "" || class = "#32769" {
        deeper := UIAFindDescendantWithAutomation(uia, elementPtr, x, y)
        if IsObject(deeper) {
            if deeper.Has("Auto") && deeper["Auto"] != ""
                info["Auto"] := deeper["Auto"]
            if deeper.Has("Class") && deeper["Class"] != ""
                info["Class"] := deeper["Class"]
        }
    }

    return info
}

UIAFindDescendantWithAutomation(uia, elementPtr, x, y, limit := 128) {
    if !uia || !elementPtr
        return 0

    queue := []
    children := UIAFindChildren(uia, elementPtr)
    if children {
        count := UIAElementArrayLength(children)
        Loop count {
            child := UIAElementArrayGet(children, A_Index - 1)
            if child
                queue.Push(child)
        }
        UIARelease(children)
    }

    processed := 0
    result := 0
    while queue.Length {
        element := queue.RemoveAt(1)
        processed += 1

        auto := UIAGetProperty(element, 30011)
        class := UIAGetProperty(element, 30012)
        rect := UIAGetBoundingRect(element)
        matches := true
        if rect
            matches := UIAPointInRect(rect, x, y)

        if matches && (auto != "" || (class != "" && class != "#32769")) {
            result := Map("Auto", auto, "Class", class)
            UIARelease(element)
            break
        }

        if processed < limit {
            children := UIAFindChildren(uia, element)
            if children {
                count := UIAElementArrayLength(children)
                Loop count {
                    child := UIAElementArrayGet(children, A_Index - 1)
                    if child
                        queue.Push(child)
                }
                UIARelease(children)
            }
        }

        UIARelease(element)
        if processed >= limit
            break
    }

    ; Release any remaining enqueued elements.
    while queue.Length {
        ptr := queue.Pop()
        UIARelease(ptr)
    }

    return result
}

UIAFindChildren(uia, elementPtr) {
    if !uia || !elementPtr
        return 0
    cond := UIAGetTrueCondition(uia)
    if !cond
        return 0
    children := 0
    try {
        if ComCall(7, elementPtr, "int", 2, "ptr", cond, "ptr*", &children) = 0 && children
            return children
    } catch {
    }
    if children
        UIARelease(children)
    return 0
}

UIAElementArrayLength(arrayPtr) {
    if !arrayPtr
        return 0
    length := 0
    try ComCall(4, arrayPtr, "int*", &length)
    return length
}

UIAElementArrayGet(arrayPtr, index) {
    if !arrayPtr
        return 0
    element := 0
    try ComCall(5, arrayPtr, "int", index, "ptr*", &element)
    return element
}

UIAGetTrueCondition(uia) {
    static cond := 0
    if cond
        return cond
    try {
        if ComCall(22, uia, "ptr*", &cond) = 0 && cond
            return cond
    } catch {
    }
    cond := 0
    return 0
}

UIAGetBoundingRect(elementPtr) {
    if !elementPtr
        return 0
    variantSize := (A_PtrSize = 8) ? 24 : 16
    static GET_CURRENT_PROPERTY_VALUE := 10
    static GET_CURRENT_PROPERTY_VALUE_EX := 11
    rectVariant := Buffer(variantSize, 0)

    attempt := ComCall(GET_CURRENT_PROPERTY_VALUE, elementPtr, "int", 30001, "ptr", rectVariant.Ptr)
    if attempt != 0
        attempt := ComCall(GET_CURRENT_PROPERTY_VALUE_EX, elementPtr, "int", 30001, "int", true, "ptr", rectVariant.Ptr)
    if attempt != 0 {
        UIATryVariantClear(rectVariant)
        return 0
    }

    vt := NumGet(rectVariant, 0, "ushort")
    static VT_ARRAY := 0x2000
    static VT_R8 := 5
    if vt != (VT_ARRAY | VT_R8) {
        UIATryVariantClear(rectVariant)
        return 0
    }

    psa := NumGet(rectVariant, 8, "ptr")
    if !psa {
        UIATryVariantClear(rectVariant)
        return 0
    }

    dataPtr := NumGet(psa, (A_PtrSize = 8) ? 16 : 12, "ptr")
    if !dataPtr {
        UIATryVariantClear(rectVariant)
        return 0
    }

    left := NumGet(dataPtr, 0, "double")
    top := NumGet(dataPtr, 8, "double")
    width := NumGet(dataPtr, 16, "double")
    height := NumGet(dataPtr, 24, "double")
    UIATryVariantClear(rectVariant)

    return Map("x", left, "y", top, "w", width, "h", height)
}

UIAPointInRect(rect, x, y) {
    if !IsObject(rect)
        return false
    if x < rect["x"] || y < rect["y"]
        return false
    if x > rect["x"] + rect["w"] || y > rect["y"] + rect["h"]
        return false
    return true
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
            UIATryVariantClear(variant)
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

UIATryVariantClear(varBuf) {
    if !IsObject(varBuf)
        ptr := varBuf
    else
        ptr := varBuf.Ptr
    try DllCall("OleAut32\\VariantClear", "ptr", ptr)
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
