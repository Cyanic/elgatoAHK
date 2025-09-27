#Requires AutoHotkey v2.0
#SingleInstance Force

global gWin32EnumContext := 0

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

    filterText := Trim(classPrompt.Value)
    matches := FindChildWindows(targetHwnd, filterText)

    dateStamp := FormatTime(, "yyyy-MM-dd")
    outPath := A_ScriptDir "\" dateStamp "-output.txt"
    WriteResults(outPath, matches, filterText)
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
    filterOriginal := Trim(filterClass)
    filter := StrLower(filterOriginal)
    matches := []

    uia := GetUIAutomation()
    if !IsObject(uia)
        return matches

    root := UIAElementFromHandle(uia, hwnd)
    if !root {
        MsgBox "UI Automation could not bind to the target window."
        return matches
    }

    uiaMatches := []
    if filterOriginal != ""
        uiaMatches := UIAFindClassMatches(uia, root, filter, filterOriginal)
    if !uiaMatches.Length
        uiaMatches := UIATraverseMatches(uia, root, filter)
    nativeMatches := Win32TraverseMatches(hwnd, filter)
    return UIAMergeMatches(uiaMatches, nativeMatches)
}

UIATraverseMatches(uia, rootElement, filter) {
    matches := []
    queue := []
    queue.Push(Map("Element", rootElement, "Depth", 0))

    processed := 0
    maxNodes := 5000

    try {
        while queue.Length {
            current := queue.RemoveAt(1)
            element := current["Element"]
            depth := current["Depth"]

            processed += 1

            details := UIAGetDirectElementInfo(element)
            if IsObject(details) {
                record := UIABuildMatchRecord(details, depth)
                if UIAMatchFilter(record, filter)
                    matches.Push(record)
            }

            if processed < maxNodes {
                children := UIAFindChildren(uia, element)
                if children {
                    count := UIAElementArrayLength(children)
                    Loop count {
                        child := UIAElementArrayGet(children, A_Index - 1)
                        if child
                            queue.Push(Map("Element", child, "Depth", depth + 1))
                    }
                    UIARelease(children)
                }
            }

            UIARelease(element)

            if processed >= maxNodes
                break
        }
    } finally {
        while queue.Length {
            leftover := queue.Pop()
            if IsObject(leftover) {
                element := leftover.Has("Element") ? leftover["Element"] : 0
                if element
                    UIARelease(element)
            } else if leftover {
                UIARelease(leftover)
            }
        }
    }

    return matches
}

UIAFindClassMatches(uia, rootElement, filterLower, filterExact) {
    matches := []
    if filterExact = ""
        return matches

    condObj := ""
    try {
        condObj := uia.CreatePropertyCondition(30012, filterExact)
    } catch {
        condObj := ""
    }
    if !IsObject(condObj)
        return matches

    condPtr := ComObjValue(condObj)
    if !condPtr
        return matches

    elements := 0
    static TREE_SCOPE_DESCENDANTS := 4
    hr := 1
    try {
        hr := ComCall(8, rootElement, "int", TREE_SCOPE_DESCENDANTS, "ptr", condPtr, "ptr*", &elements)
    } catch {
        hr := 1
    }
    if hr != 0 || !elements
        return matches

    count := UIAElementArrayLength(elements)
    loop count {
        elem := UIAElementArrayGet(elements, A_Index - 1)
        if !elem
            continue
        details := UIAGetDirectElementInfo(elem)
        if IsObject(details) {
            record := UIABuildMatchRecord(details, 0)
            if filterLower = "" || UIAMatchFilter(record, filterLower)
                matches.Push(record)
        }
        UIARelease(elem)
    }
    UIARelease(elements)
    return matches
}

Win32TraverseMatches(hwnd, filter) {
    if !hwnd
        return []

    global gWin32EnumContext
    context := Map()
    context["Items"] := []
    context["Filter"] := StrLower(filter)
    gWin32EnumContext := context

    callback := CallbackCreate(Win32EnumProc, "Fast")
    try {
        DllCall("EnumChildWindows", "ptr", hwnd, "ptr", callback, "ptr", 0)
    } finally {
        CallbackFree(callback)
    }

    results := context.Has("Items") ? context["Items"] : []
    gWin32EnumContext := 0
    return results
}

Win32EnumProc(childHwnd, lParam) {
    global gWin32EnumContext
    context := gWin32EnumContext
    if !IsObject(context)
        return true
    results := context["Items"]
    filter := context["Filter"]

    class := GetWindowClassName(childHwnd)
    classLower := StrLower(class)
    title := ""
    try title := WinGetTitle("ahk_id " childHwnd)

    record := Map()
    record["HWNDRaw"] := childHwnd
    record["HWND"] := Format("0x{1:X}", childHwnd)
    record["Class"] := class
    record["UIAClass"] := ""
    record["Type"] := ""
    record["AutomationId"] := ""
    record["Name"] := title
    record["Depth"] := -1

    if filter = "" {
        results.Push(record)
    } else {
        titleLower := StrLower(title)
        if InStr(classLower, filter) || (titleLower != "" && InStr(titleLower, filter))
            results.Push(record)
    }
    return true
}

UIABuildMatchRecord(details, depth) {
    record := Map()

    uiaClass := details.Has("Class") ? Trim(details["Class"]) : ""
    name := details.Has("Name") ? Trim(details["Name"]) : ""
    autoId := details.Has("Auto") ? Trim(details["Auto"]) : ""

    localizedType := details.Has("LocalizedControlType") ? Trim(details["LocalizedControlType"]) : ""
    controlType := details.Has("ControlType") ? Trim(details["ControlType"]) : ""
    typeName := localizedType != "" ? localizedType : controlType

    hwnd := 0
    if details.Has("HWND")
        hwnd := details["HWND"]
    nativeClass := ""
    if hwnd
        nativeClass := Trim(GetWindowClassName(hwnd))

    hwndText := hwnd ? Format("0x{1:X}", hwnd) : ""

    record["HWND"] := hwndText
    record["HWNDRaw"] := hwnd
    record["Class"] := nativeClass
    record["UIAClass"] := uiaClass
    record["Type"] := typeName
    record["AutomationId"] := autoId
    record["Name"] := name
    record["Depth"] := depth

    return record
}

UIAMatchFilter(record, filter) {
    if filter = ""
        return true

    fields := ["Class", "UIAClass", "Type", "AutomationId", "Name"]
    for field in fields {
        value := record.Has(field) ? record[field] : ""
        if value != "" {
            if InStr(StrLower(value), filter)
                return true
        }
    }
    return false
}

UIAMergeMatches(uiaMatches, nativeMatches) {
    merged := []
    seen := Map()

    if IsObject(uiaMatches) {
        for entry in uiaMatches {
            merged.Push(entry)
            if entry.Has("HWNDRaw") {
                hwnd := entry["HWNDRaw"]
                if hwnd
                    seen[hwnd] := true
            }
        }
    }

    if IsObject(nativeMatches) {
        for entry in nativeMatches {
            hwnd := entry.Has("HWNDRaw") ? entry["HWNDRaw"] : 0
            if hwnd && seen.Has(hwnd)
                continue
            merged.Push(entry)
        }
    }

    return merged
}

WriteResults(path, results, searchTerm := "") {
    timestamp := FormatTime(, "yyyy-MM-dd HH:mm:ss")
    filterLabel := searchTerm != "" ? searchTerm : "<none>"
    header := Format("Control Class Scan - {1} | Filter: {2}", timestamp, filterLabel)
    lines := []
    lines.Push(header)
    if results.Length = 0 {
        lines.Push("No matches found.")
    } else {
        for item in results {
            class := item.Has("Class") && item["Class"] != "" ? item["Class"] : "<none>"
            uiaClass := item.Has("UIAClass") && item["UIAClass"] != "" ? item["UIAClass"] : "<none>"
            typeName := item.Has("Type") && item["Type"] != "" ? item["Type"] : "<none>"
            autoId := item.Has("AutomationId") && item["AutomationId"] != "" ? item["AutomationId"] : "<none>"
            ctrlName := item.Has("Name") && item["Name"] != "" ? item["Name"] : "<none>"
            line := Format("Class: {1}`tUIAClass: {2}`tType: {3}`tAutomationId: {4}`tName: {5}`tHWND: {6}", class, uiaClass, typeName, autoId, ctrlName, item["HWND"])
            lines.Push(line)
        }
    }
    FileAppend(JoinLines(lines) "`n`n", path, "UTF-8")
}

CaptureUnderCursor(*) {
    info := GatherCursorUIAInfo(true, false)
    if !IsObject(info)
        return

    dateStamp := FormatTime(, "yyyy-MM-dd")
    path := A_ScriptDir "\" dateStamp "-cloutput.txt"
    timestamp := FormatTime(, "HH:mm:ss")

    rectText := ""
    rectSource := ""
    detail := info.Has("Details") ? info["Details"] : 0
    if IsObject(detail) && detail.Has("Rect") && IsObject(detail["Rect"]) {
        rect := detail["Rect"]
        rectText := Format("`tRect: ({},{},{},{})", rect["x"], rect["y"], rect["w"], rect["h"])
        if detail.Has("RectSource") {
            rectSource := " (" detail["RectSource"] " coords)"
        }
    }

    line := Format("{1}`tAutomationId: {2}`tClassName: {3}`tWinClass: {4}`tWinHWND: {5}`tCtrlHWND: {6}`tPos: ({7}, {8}){9}{10}",
        timestamp,
        info["AutomationId"],
        info["ClassName"],
        info["WinClass"],
        Format("0x{1:X}", info["WinHWND"]),
        info["CtrlHWND"] ? Format("0x{1:X}", info["CtrlHWND"]) : "<none>",
        info["PhysicalX"],
        info["PhysicalY"],
        rectText,
        rectSource)
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
            rectSource := detail.Has("RectSource") ? detail["RectSource"] : ""
            lines.Push(Format("  Rect=({}, {}, {}, {}){}", rect["x"], rect["y"], rect["w"], rect["h"], rectSource ? " [" rectSource "]" : ""))
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
                    rectSource := det.Has("RectSource") ? det["RectSource"] : ""
                    lines.Push(Format("    Rect=({}, {}, {}, {}){}", rect["x"], rect["y"], rect["w"], rect["h"], rectSource ? " [" rectSource "]" : ""))
                }
            }
            if cand.Has("Error")
                lines.Push("    Error=" cand["Error"])
        }
    }

    if info.Has("ChildSnapshots") && IsObject(info["ChildSnapshots"]) && info["ChildSnapshots"].Length {
        lines.Push("-- Descendants (up to 2 levels) --")
        for child in info["ChildSnapshots"] {
            level := child.Has("Level") ? child["Level"] : "?"
            autoId := child.Has("Auto") ? child["Auto"] : ""
            class := child.Has("Class") ? child["Class"] : ""
            name := child.Has("Name") ? child["Name"] : ""
            lines.Push(Format("  L{} AutoId='{}' Class='{}' Name='{}'"
                , level
                , autoId
                , class
                , name))

            if child.Has("Rect") && IsObject(child["Rect"]) {
                rect := child["Rect"]
                rectSource := child.Has("RectSource") ? child["RectSource"] : ""
                lines.Push(Format("    Rect=({}, {}, {}, {}){}"
                    , rect["x"], rect["y"], rect["w"], rect["h"], rectSource ? " [" rectSource "]" : ""))
            }
            if child.Has("ControlType") || child.Has("LocalizedControlType") {
                lines.Push(Format("    ControlType={} ({})"
                    , child.Has("ControlTypeId") ? child["ControlTypeId"] : ""
                    , child.Has("LocalizedControlType") ? child["LocalizedControlType"] : child.Has("ControlType") ? child["ControlType"] : ""))
            }
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
    bestCandidate := 0
    bestChildSnapshots := []

    hitContext := Map("WinLeft", winLeft, "WinTop", winTop, "WinWidth", winWidth, "WinHeight", winHeight)

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
        score := UIACandidateScore(details, pointX, pointY, hitContext)

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
            bestCandidate := UIACloneMap(candidate)
            if IsObject(bestCandidate) {
                bestCandidate["PointX"] := pointX
                bestCandidate["PointY"] := pointY
            }
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

    if IsObject(bestCandidate) {
        pointerX := bestCandidate.Has("PointX") ? bestCandidate["PointX"] : mx
        pointerY := bestCandidate.Has("PointY") ? bestCandidate["PointY"] : my
        descendants := UIASnapshotDescendants(uia, bestCandidate, hitContext, 3, pointerX, pointerY)
        bestChildSnapshots := descendants
        if descendants.Length {
            bestDescendant := UIASelectBestSnapshot(descendants, pointerX, pointerY, hitContext, bestScore, bestDetails)
            if IsObject(bestDescendant) {
                if bestDescendant.Has("Auto") && bestDescendant["Auto"] != ""
                    automationId := bestDescendant["Auto"]
                if bestDescendant.Has("Class") && bestDescendant["Class"] != ""
                    className := bestDescendant["Class"]
                selectedDetails := bestDescendant
                bestDetails := bestDescendant
                if selectedSource = ""
                    selectedSource := bestDescendant.Has("Source") ? bestDescendant["Source"] : "Descendant"
                if bestCandidateLabel != ""
                    selectedCandidate := bestCandidateLabel " -> descendant"
            }
        }
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
    if (className = "" || className = "#32769" || className = "<none>") && IsObject(bestChildSnapshots) {
        for child in bestChildSnapshots {
            childClass := child.Has("Class") ? child["Class"] : ""
            if childClass != "" && childClass != "#32769" {
                className := childClass
                if child.Has("Auto") && child["Auto"] != ""
                    automationId := child["Auto"]
                selectedDetails := child
                if selectedSource = ""
                    selectedSource := child.Has("Source") ? child["Source"] : "Descendant"
                if bestCandidateLabel != ""
                    selectedCandidate := bestCandidateLabel " -> descendant"
                break
            }
        }
    }
    if className = ""
        className := "<none>"
    if className = "#32769"
        className := "<fallback>"

    if IsObject(selectedDetails)
        selectedDetails["Class"] := className

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
        info["ChildSnapshots"] := bestChildSnapshots
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

    if info.Has("Rect") && IsObject(info["Rect"])
        normalized := UIANormalizeRect(info["Rect"], x, y, ctx)
    if IsObject(normalized) {
        if normalized.Has("Screen")
            info["Rect"] := normalized["Screen"]
        if normalized.Has("Local")
            info["RectLocal"] := normalized["Local"]
        if normalized.Has("RectSource")
            info["RectSource"] := normalized["RectSource"]
    }

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
            if info.Has("Rect") && IsObject(info["Rect"]) {
                normalized := UIANormalizeRect(info["Rect"], x, y, ctx)
                if IsObject(normalized) {
                    if normalized.Has("Screen")
                        info["Rect"] := normalized["Screen"]
                    if normalized.Has("Local")
                        info["RectLocal"] := normalized["Local"]
                    if normalized.Has("RectSource")
                        info["RectSource"] := normalized["RectSource"]
                }
            }
            rect := info.Has("Rect") ? info["Rect"] : ""
            matches := true
            if IsObject(rect)
                matches := UIAPointMatchesRect(rect, x, y, ctx)

            hasData := (info["Auto"] != "") || (info["Class"] != "" && info["Class"] != "#32769")
            if matches && hasData {
                info["Source"] := "Descendant"
                score := UIACandidateScore(info, x, y, ctx)
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
    nativeHandle := UIAGetProperty(elementPtr, 30020)

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
    if nativeHandle != ""
        info["HWND"] := nativeHandle + 0
    else
        info["HWND"] := 0
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

UIAPointDistanceToRect(rect, x, y, ctx := 0) {
    if !IsObject(rect)
        return 100000

    points := []
    points.Push([x, y])
    if IsObject(ctx) {
        localX := x - (ctx.Has("WinLeft") ? ctx["WinLeft"] : 0)
        localY := y - (ctx.Has("WinTop") ? ctx["WinTop"] : 0)
        points.Push([localX, localY])
    }

    best := 100000
    for point in points {
        px := point[1]
        py := point[2]
        dx := 0
        dy := 0
        if px < rect["x"]
            dx := rect["x"] - px
        else if px > rect["x"] + rect["w"]
            dx := px - (rect["x"] + rect["w"])
        if py < rect["y"]
            dy := rect["y"] - py
        else if py > rect["y"] + rect["h"]
            dy := py - (rect["y"] + rect["h"])
        dist := Sqrt(dx * dx + dy * dy)
        if dist < best
            best := dist
    }
    return best
}

UIANormalizeRect(rect, x, y, ctx := 0) {
    if !IsObject(rect)
        return 0

    result := Map()
    screenCandidate := Map("x", rect["x"], "y", rect["y"], "w", rect["w"], "h", rect["h"])
    result["Screen"] := screenCandidate
    result["RectSource"] := "Raw"

    if !IsObject(ctx)
        return result

    winLeft := ctx.Has("WinLeft") ? ctx["WinLeft"] : 0
    winTop := ctx.Has("WinTop") ? ctx["WinTop"] : 0
    winWidth := ctx.Has("WinWidth") ? ctx["WinWidth"] : 0
    winHeight := ctx.Has("WinHeight") ? ctx["WinHeight"] : 0

    localCandidate := Map("x", rect["x"] - winLeft, "y", rect["y"] - winTop, "w", rect["w"], "h", rect["h"])
    screenFromLocal := Map("x", localCandidate["x"] + winLeft, "y", localCandidate["y"] + winTop, "w", localCandidate["w"], "h", localCandidate["h"])

    pointerLocalX := x - winLeft
    pointerLocalY := y - winTop

    containsScreen := UIAPointInRect(screenCandidate, x, y)
    containsFromLocal := UIAPointInRect(screenFromLocal, x, y)
    containsLocal := UIAPointInRect(localCandidate, pointerLocalX, pointerLocalY)

    scoreScreen := containsScreen ? 0 : UIAPointDistanceToRect(screenCandidate, x, y)
    scoreFromLocal := containsFromLocal ? 0 : UIAPointDistanceToRect(screenFromLocal, x, y)

    if containsScreen && !containsFromLocal {
        result["Screen"] := screenCandidate
        result["Local"] := localCandidate
        result["RectSource"] := "Screen"
    } else if containsFromLocal && !containsScreen {
        result["Screen"] := screenFromLocal
        result["Local"] := localCandidate
        result["RectSource"] := "Local"
    } else {
        if scoreFromLocal < scoreScreen {
            result["Screen"] := screenFromLocal
            result["Local"] := localCandidate
            result["RectSource"] := containsLocal ? "Local" : "Adjusted"
        } else {
            result["Screen"] := screenCandidate
            result["Local"] := localCandidate
            result["RectSource"] := containsScreen ? "Screen" : "Raw"
        }
    }

    return result
}

UIACloneMap(source) {
    if !IsObject(source)
        return source
    clone := Map()
    for key, value in source
        clone[key] := value
    return clone
}

UIASelectBestSnapshot(snapshots, x, y, ctx, baselineScore := -1, baselineDetails := 0) {
    best := 0
    bestScore := baselineScore
    baseHasAuto := IsObject(baselineDetails) && baselineDetails.Has("Auto") && baselineDetails["Auto"] != ""
    baseHasGoodClass := IsObject(baselineDetails) && baselineDetails.Has("Class") && baselineDetails["Class"] != "" && baselineDetails["Class"] != "#32769"
    for snapshot in snapshots {
        score := UIACandidateScore(snapshot, x, y, ctx)
        hasAuto := snapshot.Has("Auto") && snapshot["Auto"] != ""
        hasGoodClass := snapshot.Has("Class") && snapshot["Class"] != "" && snapshot["Class"] != "#32769"

        prefer := false
        if score > bestScore
            prefer := true
        else if !baseHasAuto && hasAuto
            prefer := true
        else if !baseHasGoodClass && hasGoodClass && score >= bestScore - 500
            prefer := true
        else if !best
            prefer := true

        if prefer {
            bestScore := Max(score, bestScore)
            best := snapshot
            baseHasAuto := hasAuto
            baseHasGoodClass := hasGoodClass
        }
    }
    return best
}

UIAResolveCandidateElement(uia, candidate, ctx) {
    if !uia || !IsObject(candidate)
        return 0
    kind := candidate.Has("Kind") ? candidate["Kind"] : ""
    switch kind {
        case "Handle":
            value := candidate.Has("Value") ? candidate["Value"] : 0
            return value ? UIAElementFromHandle(uia, value) : 0
        case "Point":
            x := candidate.Has("PointX") ? candidate["PointX"] : (candidate.Has("X") ? candidate["X"] : 0)
            y := candidate.Has("PointY") ? candidate["PointY"] : (candidate.Has("Y") ? candidate["Y"] : 0)
            return UIAElementFromPoint(uia, x, y)
    }
    return 0
}

UIASnapshotDescendants(uia, candidate, ctx, depth := 2, pointerX := 0, pointerY := 0, limit := 128) {
    if depth <= 0 || limit <= 0
        return []
    element := UIAResolveCandidateElement(uia, candidate, ctx)
    if !element
        return []
    snapshots := []
    UIAGatherDescendantInfo(uia, element, ctx, depth, 1, snapshots, pointerX, pointerY, &limit)
    UIARelease(element)
    return snapshots
}

UIAGatherDescendantInfo(uia, elementPtr, ctx, remaining, level, outArr, pointerX, pointerY, &limit) {
    if remaining <= 0 || limit <= 0
        return
    children := UIAFindChildren(uia, elementPtr)
    if !children
        return
    count := UIAElementArrayLength(children)
    Loop count {
        child := UIAElementArrayGet(children, A_Index - 1)
        if !child
            continue
        info := UIAGetDirectElementInfo(child)
        if IsObject(info) {
            if info.Has("Rect") && IsObject(info["Rect"]) {
                normalized := UIANormalizeRect(info["Rect"], pointerX, pointerY, ctx)
                if IsObject(normalized) {
                    if normalized.Has("Screen")
                        info["Rect"] := normalized["Screen"]
                    if normalized.Has("Local")
                        info["RectLocal"] := normalized["Local"]
                    if normalized.Has("RectSource")
                        info["RectSource"] := normalized["RectSource"]
                }
            }
            info["Level"] := level
            info["Source"] := "Descendant"
            outArr.Push(info)
            limit -= 1
        }
        if remaining > 1 && limit > 0
            UIAGatherDescendantInfo(uia, child, ctx, remaining - 1, level + 1, outArr, pointerX, pointerY, &limit)
        UIARelease(child)
        if limit <= 0
            break
    }
    UIARelease(children)
}

UIACandidateScore(info, x, y, ctx) {
    score := -100000
    pointScore := 0
    if info.Has("Rect") && IsObject(info["Rect"]) {
        rect := info["Rect"]
        if UIAPointMatchesRect(rect, x, y, ctx) {
            pointScore := 4000
        } else {
            overlap := UIAPointDistanceToRect(rect, x, y, ctx)
            pointScore := Max(-4000, 2000 - overlap)
        }
    }
    score += pointScore

    if info.Has("Auto") && info["Auto"] != ""
        score += 1000
    if info.Has("Class") && info["Class"] != "" && info["Class"] != "#32769"
        score += 500
    if info.Has("ControlType") && info["ControlType"] != ""
        score += 100
    if info.Has("LocalizedControlType") && info["LocalizedControlType"] != ""
        score += 100
    if info.Has("Rect") && IsObject(info["Rect"]) {
        rect := info["Rect"]
        area := Abs(rect["w"]) * Abs(rect["h"])
        if area > 0
            score += Max(0, 2000 - Min(area, 2000))
        else
            score += 2000
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
