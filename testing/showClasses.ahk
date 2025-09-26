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
    info := GatherCursorUIAInfo(true, false)
    if !IsObject(info)
        return

    dateStamp := FormatTime(, "yyyy-MM-dd")
    path := A_ScriptDir "\" dateStamp "-cloutput.txt"
    timestamp := FormatTime(, "HH:mm:ss")

    line := Format("{1}`tAutomationId: {2}`tClassName: {3}`tWinClass: {4}`tWinHWND: {5}`tCtrlHWND: {6}`tPos: ({7}, {8})",
        timestamp,
        info["AutomationId"],
        info["ClassName"],
        info["WinClass"],
        Format("0x{1:X}", info["WinHWND"]),
        info["CtrlHWND"] ? Format("0x{1:X}", info["CtrlHWND"]) : "<none>",
        info["PhysicalX"],
        info["PhysicalY"])
    FileAppend(line "`n", path, "UTF-8")
    MsgBox "Captured AutomationId: " info["AutomationId"] "`nClassName: " info["ClassName"] "`nLogged to " path
}

CaptureAutomationDebug(*) {
    info := GatherCursorUIAInfo(true, true)
    if !IsObject(info)
        return

    dateStamp := FormatTime(, "yyyy-MM-dd")
    path := A_ScriptDir "\" dateStamp "-uia-debug.txt"
    timestamp := FormatTime(, "yyyy-MM-dd HH:mm:ss")

    lines := []
    lines.Push("=== UIA Debug Capture ===")
    lines.Push("Timestamp: " timestamp)
    lines.Push(Format("LogicalPos: ({}, {})  PhysicalPos: ({}, {})", info["LogicalX"], info["LogicalY"], info["PhysicalX"], info["PhysicalY"]))
    lines.Push(Format("Window: hwnd={} class={} title={}", Format("0x{1:X}", info["WinHWND"]), info["WinClass"], info.Has("WinTitle") ? info["WinTitle"] : ""))
    lines.Push(Format("WindowRect: left={} top={} width={} height={}"
        , info.Has("WinLeft") ? info["WinLeft"] : "?"
        , info.Has("WinTop") ? info["WinTop"] : "?"
        , info.Has("WinWidth") ? info["WinWidth"] : "?"
        , info.Has("WinHeight") ? info["WinHeight"] : "?"))
    lines.Push(Format("ControlHWND: {}", info["CtrlHWND"] ? Format("0x{1:X}", info["CtrlHWND"]) : "<none>"))
    lines.Push(Format("Resolved: AutoId='{}' Class='{}' Source={} Candidate={} UIAClass={} ",
        info["AutomationId"], info["ClassName"], info.Has("SelectedSource") ? info["SelectedSource"] : "",
        info.Has("SelectedCandidate") ? info["SelectedCandidate"] : "",
        info.Has("UIAClass") ? info["UIAClass"] : ""))

    detail := info.Has("Details") ? info["Details"] : 0
    if IsObject(detail) && detail.Count {
        lines.Push("-- Resolved Element Details --")
        lines.Push(Format("  Name='{}' ControlType={} ({}) Localized='{}' Framework='{}'"
            , detail.Has("Name") ? detail["Name"] : ""
            , detail.Has("ControlTypeId") ? detail["ControlTypeId"] : ""
            , detail.Has("ControlType") ? detail["ControlType"] : ""
            , detail.Has("LocalizedControlType") ? detail["LocalizedControlType"] : ""
            , detail.Has("FrameworkId") ? detail["FrameworkId"] : ""))
        if detail.Has("Rect") && IsObject(detail["Rect"]) {
            rect := detail["Rect"]
            lines.Push(Format("  Rect=({}, {}, {}, {})", rect["x"], rect["y"], rect["w"], rect["h"]))
        }
    }

    if info.Has("Notes") && info["Notes"].Length {
        lines.Push("-- Notes --")
        for note in info["Notes"]
            lines.Push("  " note)
    }

    if info.Has("Candidates") && IsObject(info["Candidates"]) {
        lines.Push("-- Candidates --")
        idx := 0
        for cand in info["Candidates"] {
            idx += 1
            line := Format("  [{}] Kind={} Label={} Element={} Refined={} Used={}"
                , idx
                , cand.Has("Kind") ? cand["Kind"] : ""
                , cand.Has("Label") ? cand["Label"] : ""
                , cand.Has("ElementFound") ? (cand["ElementFound"] ? "Yes" : "No") : "Unknown"
                , cand.Has("Refined") ? (cand["Refined"] ? "Yes" : "No") : "No"
                , cand.Has("Used") ? (cand["Used"] ? "Yes" : "No") : "No")
            lines.Push(line)
            if cand.Has("Point")
                lines.Push(Format("    Point=({}, {})", cand["Point"]["x"], cand["Point"]["y"]))
            if cand.Has("Handle")
                lines.Push(Format("    Handle={}", Format("0x{1:X}", cand["Handle"])))
            if cand.Has("Details") {
                det := cand["Details"]
                lines.Push(Format("    AutoId='{}' Class='{}' Source={}"
                    , det.Has("Auto") ? det["Auto"] : ""
                    , det.Has("Class") ? det["Class"] : ""
                    , det.Has("Source") ? det["Source"] : ""))
                if det.Has("Name") || det.Has("LocalizedControlType") || det.Has("FrameworkId")
                    lines.Push(Format("    Name='{}' Localized='{}' Framework='{}'"
                        , det.Has("Name") ? det["Name"] : ""
                        , det.Has("LocalizedControlType") ? det["LocalizedControlType"] : ""
                        , det.Has("FrameworkId") ? det["FrameworkId"] : ""))
                if det.Has("ControlType")
                    lines.Push(Format("    ControlType={} ({})"
                        , det.Has("ControlTypeId") ? det["ControlTypeId"] : ""
                        , det["ControlType"]))
                if det.Has("Rect") && IsObject(det["Rect"]) {
                    rect := det["Rect"]
                    lines.Push(Format("    Rect=({}, {}, {}, {})", rect["x"], rect["y"], rect["w"], rect["h"]))
                }
            }
            if cand.Has("Error")
                lines.Push("    Error=" cand["Error"])
        }
    }

    lines.Push("")
    FileAppend(JoinLines(lines) "`n", path, "UTF-8")
    MsgBox "UIA debug written to: " path
}

GatherCursorUIAInfo(showMessages := true, collectDebug := false) {
    info := Map()
    notes := []
    debugCandidates := collectDebug ? [] : ""

    uia := GetUIAutomation()
    if !uia {
        if showMessages
            MsgBox "UI Automation interface not available on this system."
        return 0
    }

    CoordMode("Mouse", "Screen")
    MouseGetPos(&logicalX, &logicalY, &winHwnd, &ctrlInfo)
    physX := logicalX
    physY := logicalY
    hasPhysical := GetPhysicalCursorPos(&physX, &physY)
    mx := hasPhysical ? physX : logicalX
    my := hasPhysical ? physY : logicalY

    winHwnd := NormalizeHwnd(winHwnd)
    if !winHwnd {
        if showMessages
            MsgBox "Could not determine the hovered control."
        return 0
    }

    ctrlHwnd := NormalizeHwnd(ctrlInfo, winHwnd)
    if !ctrlHwnd {
        ctrlHwnd := NormalizeHwnd(WindowFromPoint(mx, my), winHwnd)
        if ctrlHwnd = winHwnd
            ctrlHwnd := 0
    }
    if ctrlHwnd {
        static GA_ROOT := 2
        root := DllCall("GetAncestor", "ptr", ctrlHwnd, "uint", GA_ROOT, "ptr")
        if root != winHwnd {
            if collectDebug
                notes.Push("Control hwnd rejected because root mismatch")
            ctrlHwnd := 0
        }
    }

    winClass := ""
    winTitle := ""
    try winClass := WinGetClass("ahk_id " winHwnd)
    try winTitle := WinGetTitle("ahk_id " winHwnd)

    winLeft := 0
    winTop := 0
    winWidth := 0
    winHeight := 0
    try WinGetPos(&winLeft, &winTop, &winWidth, &winHeight, "ahk_id " winHwnd)

    candidates := []
    candidates.Push(Map("Kind", "Point", "X", logicalX, "Y", logicalY, "Label", "Logical point"))
    if hasPhysical && (physX != logicalX || physY != logicalY)
        candidates.Push(Map("Kind", "Point", "X", physX, "Y", physY, "Label", "Physical point"))
    if ctrlHwnd
        candidates.Push(Map("Kind", "Handle", "Value", ctrlHwnd, "Label", "Control handle"))
    candidates.Push(Map("Kind", "Handle", "Value", winHwnd, "Label", "Window handle"))

    automationId := ""
    className := ""
    uiaClass := ""
    fallbackAuto := ""
    fallbackClass := ""
    preferredClass := ""
    fallbackAutoDetails := 0
    fallbackClassDetails := 0
    preferredClassDetails := 0
    selectedDetails := 0
    selectedSource := ""
    selectedCandidate := ""
    chosenIndex := 0
    bestDetails := 0
    bestScore := -1
    bestCandidateLabel := ""

    hitContext := Map("WinLeft", winLeft, "WinTop", winTop)

    bestDebugIndex := 0

    for candidate in candidates {
        element := 0
        pointX := candidate.Has("X") ? candidate["X"] : mx
        pointY := candidate.Has("Y") ? candidate["Y"] : my

        candDebug := Map()
        if collectDebug {
            candDebug["Kind"] := candidate["Kind"]
            if candidate.Has("Label")
                candDebug["Label"] := candidate["Label"]
            if candidate["Kind"] = "Point"
                candDebug["Point"] := Map("x", pointX, "y", pointY)
            else if candidate["Kind"] = "Handle"
                candDebug["Handle"] := candidate["Value"]
        }

        switch candidate["Kind"] {
            case "Handle":
                element := UIAElementFromHandle(uia, candidate["Value"])
            case "Point":
                element := UIAElementFromPoint(uia, pointX, pointY)
        }

        if collectDebug
            candDebug["ElementFound"] := element ? true : false

        if !element {
            if collectDebug {
                candDebug["Error"] := "No element"
                debugCandidates.Push(candDebug)
            }
            continue
        }

        refined := UIARefineElementAtPoint(uia, element, pointX, pointY, hitContext)
        if refined && refined != element {
            if collectDebug
                candDebug["Refined"] := true
            UIARelease(element)
            element := refined
        } else if collectDebug {
            candDebug["Refined"] := false
        }

        details := UIACollectElementDetails(uia, element, pointX, pointY, hitContext)
        if !details.Has("Source")
            details["Source"] := candidate.Has("Label") ? candidate["Label"] : candidate["Kind"]
        candAuto := details.Has("Auto") ? details["Auto"] : ""
        candClass := details.Has("Class") ? details["Class"] : ""
        score := UIACandidateScore(details)

        if collectDebug {
            candDebug["Details"] := details
            candDebug["Score"] := score
            debugCandidates.Push(candDebug)
        }

        UIARelease(element)

        if score > bestScore {
            bestScore := score
            bestDetails := details
            bestCandidateLabel := candidate.Has("Label") ? candidate["Label"] : candidate["Kind"]
            selectedCandidate := bestCandidateLabel
            selectedSource := details.Has("Source") ? details["Source"] : ""
            if collectDebug
                bestDebugIndex := debugCandidates.Length
        }

        if candClass != "" {
            if candClass != winClass
                uiaClass := candClass
            else if uiaClass = ""
                uiaClass := candClass
        }
        if fallbackClass = "" && candClass != "" {
            fallbackClass := candClass
            fallbackClassDetails := details
        }
        if fallbackAuto = "" && candAuto != "" {
            fallbackAuto := candAuto
            fallbackAutoDetails := details
        }

        if candClass != "" && candClass != "#32769" && candClass != winClass {
            if preferredClass = "" {
                preferredClass := candClass
                preferredClassDetails := details
            }
        }
    }

    if IsObject(bestDetails) {
        if bestDetails.Has("Auto") && bestDetails["Auto"] != ""
            automationId := bestDetails["Auto"]
        if bestDetails.Has("Class") && bestDetails["Class"] != ""
            className := bestDetails["Class"]
        selectedDetails := bestDetails
        if selectedSource = ""
            selectedSource := bestDetails.Has("Source") ? bestDetails["Source"] : ""
        if bestCandidateLabel != ""
            selectedCandidate := bestCandidateLabel
    }

    if className = "" && preferredClass != "" {
        className := preferredClass
        if !selectedDetails
            selectedDetails := preferredClassDetails
        if selectedSource = ""
            selectedSource := preferredClassDetails && preferredClassDetails.Has("Source") ? preferredClassDetails["Source"] : ""
    }
    if className = "" && fallbackClass != "" {
        className := fallbackClass
        if !selectedDetails
            selectedDetails := fallbackClassDetails
        if selectedSource = ""
            selectedSource := fallbackClassDetails && fallbackClassDetails.Has("Source") ? fallbackClassDetails["Source"] : ""
    }
    if automationId = "" && fallbackAuto != "" {
        automationId := fallbackAuto
        if !selectedDetails
            selectedDetails := fallbackAutoDetails
        if selectedSource = ""
            selectedSource := fallbackAutoDetails && fallbackAutoDetails.Has("Source") ? fallbackAutoDetails["Source"] : ""
    }

    if automationId = ""
        automationId := "<none>"
    nativeUsed := false
    if className = "" || className = "#32769" {
        nativeClass := GetWindowClassName(ctrlHwnd ? ctrlHwnd : winHwnd)
        if nativeClass != "" {
            className := nativeClass
            nativeUsed := true
            if selectedSource = ""
                selectedSource := "Native"
        }
    }
    if (className = "" || className = winClass) && (uiaClass != "") {
        className := uiaClass
        if selectedSource = ""
            selectedSource := "FallbackUIA"
    }
    if className = ""
        className := "<none>"

    if collectDebug {
        chosenIndex := bestDebugIndex
        if chosenIndex && chosenIndex <= debugCandidates.Length
            debugCandidates[chosenIndex]["Used"] := true
    }

    if nativeUsed
        notes.Push("Used native class fallback")
    if selectedSource = ""
        selectedSource := nativeUsed ? "Native" : "Unknown"

    info["AutomationId"] := automationId
    info["ClassName"] := className
    info["WinClass"] := winClass
    info["WinTitle"] := winTitle
    info["WinHWND"] := winHwnd
    info["CtrlHWND"] := ctrlHwnd
    info["WinLeft"] := winLeft
    info["WinTop"] := winTop
    info["WinWidth"] := winWidth
    info["WinHeight"] := winHeight
    info["LogicalX"] := logicalX
    info["LogicalY"] := logicalY
    info["PhysicalX"] := mx
    info["PhysicalY"] := my
    info["UIAClass"] := uiaClass
    info["SelectedSource"] := selectedSource
    info["SelectedCandidate"] := selectedCandidate
    info["Details"] := IsObject(selectedDetails) ? selectedDetails : 0
    info["Notes"] := notes
    if collectDebug
        info["Candidates"] := debugCandidates
    return info
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

UIARefineElementAtPoint(uia, elementPtr, x, y, ctx) {
    if !uia || !elementPtr
        return 0

    refined := UIAHitTestElement(uia, elementPtr, x, y, ctx)
    return refined ? refined : elementPtr
}

UIAHitTestElement(uia, elementPtr, x, y, ctx) {
    if !elementPtr
        return 0

    rect := UIAGetBoundingRect(elementPtr)
    matches := !rect || UIAPointMatchesRect(rect, x, y, ctx)

    children := UIAFindChildren(uia, elementPtr)
    if children {
        count := UIAElementArrayLength(children)
        loop count {
            child := UIAElementArrayGet(children, A_Index - 1)
            if !child
                continue
            deeper := UIAHitTestElement(uia, child, x, y, ctx)
            if deeper {
                if deeper != child
                    UIARelease(child)
                UIARelease(children)
                return deeper
            }
            UIARelease(child)
        }
        UIARelease(children)
    }
    return matches ? elementPtr : 0
}

UIACollectElementDetails(uia, elementPtr, x, y, ctx) {
    info := UIAGetDirectElementInfo(elementPtr)
    if !IsObject(info)
        info := Map("Auto", "", "Class", "", "Source", "None")

    if (info["Auto"] != "" && info["Class"] != "" && info["Class"] != "#32769")
        return info

    deeper := UIAFindDescendantWithAutomation(uia, elementPtr, x, y, 128, ctx)
    if IsObject(deeper) {
        for key, value in deeper {
            if key = "Source"
                info["Source"] := value
            else if key = "Rect" {
                if (!info.Has("Rect") || !IsObject(info["Rect"])) && IsObject(value)
                    info["Rect"] := value
            } else if value != "" {
                if !info.Has(key) || info[key] = "" || info[key] = "#32769"
                    info[key] := value
            }
        }
    }

    return info
}

UIAFindDescendantWithAutomation(uia, elementPtr, x, y, limit := 256, ctx := 0) {
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
    best := 0
    bestScore := -1
    while queue.Length {
        element := queue.RemoveAt(1)
        processed += 1

        info := UIAGetDirectElementInfo(element)
        if IsObject(info) {
            rect := info.Has("Rect") ? info["Rect"] : ""
            matches := true
            if IsObject(rect)
                matches := UIAPointMatchesRect(rect, x, y, ctx)

            hasData := (info["Auto"] != "") || (info["Class"] != "" && info["Class"] != "#32769")
            if matches && hasData {
                info["Source"] := "Descendant"
                score := UIACandidateScore(info)
                if score > bestScore {
                    best := info
                    bestScore := score
                }
            }
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

    return best
}

UIAGetDirectElementInfo(elementPtr) {
    if !elementPtr
        return 0

    info := Map()
    auto := UIAGetProperty(elementPtr, 30011)
    class := UIAGetProperty(elementPtr, 30012)
    name := UIAGetProperty(elementPtr, 30005)
    controlTypeRaw := UIAGetProperty(elementPtr, 30003)
    localizedType := UIAGetProperty(elementPtr, 30004)
    frameworkId := UIAGetProperty(elementPtr, 30024)

    controlTypeId := ""
    if controlTypeRaw != "" && RegExMatch(controlTypeRaw, "^-?\d+$")
        controlTypeId := controlTypeRaw + 0
    else
        controlTypeId := controlTypeRaw

    info["Auto"] := auto
    info["Class"] := class
    info["Name"] := name
    info["ControlTypeId"] := controlTypeId
    info["ControlType"] := UIAControlTypeToName(controlTypeId)
    info["LocalizedControlType"] := localizedType
    info["FrameworkId"] := frameworkId
    info["Rect"] := UIAGetBoundingRect(elementPtr)
    info["Source"] := "Element"
    return info
}

UIAControlTypeToName(id) {
    static typeMap := ""
    if typeMap = "" {
        typeMap := Map(
            50000, "Button",
            50001, "Calendar",
            50002, "CheckBox",
            50003, "ComboBox",
            50004, "Edit",
            50005, "Hyperlink",
            50006, "Image",
            50007, "ListItem",
            50008, "List",
            50009, "Menu",
            50010, "MenuBar",
            50011, "MenuItem",
            50012, "ProgressBar",
            50013, "RadioButton",
            50014, "ScrollBar",
            50015, "Slider",
            50016, "Spinner",
            50017, "StatusBar",
            50018, "Tab",
            50019, "TabItem",
            50020, "Text",
            50021, "ToolBar",
            50022, "ToolTip",
            50023, "Tree",
            50024, "TreeItem",
            50025, "Custom",
            50026, "Group",
            50027, "Thumb",
            50028, "DataGrid",
            50029, "DataItem",
            50030, "Document",
            50031, "SplitButton",
            50032, "Window",
            50033, "Pane",
            50034, "Header",
            50035, "HeaderItem",
            50036, "Table",
            50037, "TitleBar",
            50038, "Separator",
            50039, "SemanticZoom",
            50040, "AppBar"
        )
    }

    if id is Integer
        return typeMap.Has(id) ? typeMap[id] : Format("ControlType({})", id)
    if id != "" && RegExMatch(id, "^-?\d+$") {
        num := id + 0
        return typeMap.Has(num) ? typeMap[num] : Format("ControlType({})", num)
    }
    return id
}

UIAFindChildren(uia, elementPtr) {
    if !uia || !elementPtr
        return 0
    cond := UIAGetTrueCondition(uia)
    if !cond
        return 0
    children := 0
    try {
        if ComCall(6, elementPtr, "int", 2, "ptr", cond, "ptr*", &children) = 0 && children
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
    try ComCall(3, arrayPtr, "int*", &length)
    return length
}

UIAElementArrayGet(arrayPtr, index) {
    if !arrayPtr
        return 0
    element := 0
    try ComCall(4, arrayPtr, "int", index, "ptr*", &element)
    return element
}

UIAGetTrueCondition(uia) {
    static cond := 0
    if cond
        return cond
    try {
        if ComCall(21, uia, "ptr*", &cond) = 0 && cond
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

UIAPointMatchesRect(rect, x, y, ctx := 0) {
    if !IsObject(rect)
        return false
    if UIAPointInRect(rect, x, y)
        return true
    if IsObject(ctx) {
        localX := x
        localY := y
        if ctx.Has("WinLeft")
            localX := x - ctx["WinLeft"]
        if ctx.Has("WinTop")
            localY := y - ctx["WinTop"]
        if UIAPointInRect(rect, localX, localY)
            return true
    }
    return false
}

UIACandidateScore(info) {
    score := 0
    if info.Has("Auto") && info["Auto"] != ""
        score += 100
    if info.Has("Class") && info["Class"] != "" && info["Class"] != "#32769"
        score += 25
    area := 0
    if info.Has("Rect") && IsObject(info["Rect"]) {
        rect := info["Rect"]
        area := Abs(rect["w"]) * Abs(rect["h"])
        ; Smaller areas get higher scores; avoid division by zero.
        if area > 0
            score += Max(0, 1000 - Min(area, 1000))
        else
            score += 1000
    }
    return score
}

UIAGetProperty(elementPtr, propertyId) {
    if !elementPtr
        return ""
    variantSize := (A_PtrSize = 8) ? 24 : 16
    static GET_CURRENT_PROPERTY_VALUE := 10
    static GET_CURRENT_PROPERTY_VALUE_EX := 11

    attempts := [{ index: GET_CURRENT_PROPERTY_VALUE, extra: [] }, { index: GET_CURRENT_PROPERTY_VALUE_EX, extra: [true] }
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

^!a:: CaptureAutomationDebug()
^!d:: CaptureUnderCursor()

Main()
