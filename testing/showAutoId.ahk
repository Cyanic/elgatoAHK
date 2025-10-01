#Requires AutoHotkey v2.0
#SingleInstance Force
#Include "%A_ScriptDir%\uiaHelpers.ahk"

global gAutoIdMaxNodes := 300000
global gAutoIdDebug := true
global gAutoIdLogPath := ""

Main() {
    iniPath := EnsureConfig()
    config := LoadConfig(iniPath)
    if gAutoIdDebug
        InitAutoIdLog()
    LogAutoId("Configuration loaded: " MapToText(config))
    if !config.Has("ClassNN") && !config.Has("Process") {
        LogAutoId("Config missing ClassNN/Process; aborting")
        MsgBox "No window configuration found in " iniPath
        return
    }

    prompt := InputBox("Enter all or part of the AutomationId to search for:", "AutomationId Lookup")
    if prompt.Result != "OK" || prompt.Value = "" {
        LogAutoId("User cancelled prompt or provided empty input")
        MsgBox "No AutomationId provided. Exiting."
        return
    }

    targetHwnd := GetTargetWindow(config)
    LogAutoId(Format("Target window handle: 0x{1:X}", targetHwnd))
    if !targetHwnd {
        LogAutoId("GetTargetWindow returned 0")
        MsgBox "Target window not found. Check showClasses.ini"
        return
    }

    filter := Trim(prompt.Value)
    LogAutoId("Searching for AutomationId filter: '" filter "'")
    matches := FindAutomationMatches(targetHwnd, filter)

    dateStamp := FormatTime(, "yyyy-MM-dd")
    outPath := A_ScriptDir "\" dateStamp "-autoid-output.txt"
    metadata := Map("Mode", "Automation")
    for key, value in config
        metadata[key] := value
    WriteResults(outPath, matches, filter, metadata)
    LogAutoId(Format("WriteResults complete with {1} matches", matches.Length))
    MsgBox Format("Found {1} matching controls. Details written to:`n{2}", matches.Length, outPath)
}


FindAutomationMatches(hwnd, filterText) {
    matches := []
    uia := GetUIAutomation()
    if !IsObject(uia) {
        LogAutoId("GetUIAutomation returned non-object")
        return matches
    }

    root := UIAElementFromHandle(uia, hwnd)
    if !root {
        LogAutoId("UIAElementFromHandle failed for hwnd " Format("0x{1:X}", hwnd))
        MsgBox "UI Automation could not bind to the target window."
        return matches
    }

    try {
        filterTrim := Trim(filterText)
        filterLower := StrLower(filterTrim)
        LogAutoId("Filter trimmed to '" filterTrim "'")

        if filterTrim != "" {
            matches := UIAExactAutomationMatches(uia, root, filterTrim)
            LogAutoId(Format("UIAExactAutomationMatches returned {1} entries", matches.Length))
        }

        if matches.Length = 0 {
            matches := UIABreadthFirstSearch(uia, root, filterLower)
            LogAutoId(Format("UIABreadthFirstSearch returned {1} entries", matches.Length))
        }

        if matches.Length = 0 {
            matches := UIARawAutomationMatches(uia, root, filterLower)
            LogAutoId(Format("UIARawAutomationMatches returned {1} entries", matches.Length))
        }
    } finally {
        UIARelease(root)
        LogAutoId("Released root element")
    }
    return matches
}

UIAExactAutomationMatches(uia, rootElement, filterExact) {
    matches := []
    if filterExact = ""
        return matches

    cond := ""
    try {
        cond := uia.CreatePropertyCondition(30011, filterExact)
    } catch as err {
        cond := ""
        LogAutoId("CreatePropertyCondition failed: " AutoIdErrorText(err))
    }
    if !IsObject(cond)
        return matches

    condPtr := ComObjValue(cond)
    if !condPtr {
        try ObjRelease(cond)
        return matches
    }

    elements := 0
    static TREE_SCOPE_DESCENDANTS := 4
    hr := 1
    try hr := ComCall(8, rootElement, "int", TREE_SCOPE_DESCENDANTS, "ptr", condPtr, "ptr*", &elements)
    catch as err {
        hr := 1
        LogAutoId("ComCall(QIExact) threw: " AutoIdErrorText(err))
    }
    LogAutoId(Format("Exact match ComCall hr={1} elementsPtr=0x{2:X}", hr, elements))
    if hr != 0 || !elements {
        try ObjRelease(cond)
        if elements
            UIARelease(elements)
        return matches
    }

    try {
        count := UIAElementArrayLength(elements)
        LogAutoId(Format("Exact match element array count={1}", count))
        Loop count {
            elem := UIAElementArrayGet(elements, A_Index - 1)
            if !elem
                continue
            details := UIAGetDirectElementInfo(elem)
            if IsObject(details) {
                record := UIABuildMatchRecord(details, -1)
                matches.Push(record)
                LogAutoId(Format("Exact match found autoId='{1}' name='{2}'", record.Has("AutomationId") ? record["AutomationId"] : "", record.Has("Name") ? record["Name"] : ""))
            }
            UIARelease(elem)
        }
    } finally {
        UIARelease(elements)
        try ObjRelease(cond)
        LogAutoId(Format("UIAExactAutomationMatches finalized with {1} results", matches.Length))
    }

    return matches
}


UIABreadthFirstSearch(uia, rootElement, filterLower) {
    maxNodes := gAutoIdMaxNodes
    LogAutoId(Format("Starting BFS filter='{1}'", filterLower))
    queue := []
    queue.Push(Map("Element", rootElement, "Depth", 0, "Release", false))
    matches := []
    index := 1
    processed := 0
    visited := Map()
    rootKey := UIAPointerKey(rootElement)
    if rootKey
        visited[rootKey] := true

    while index <= queue.Length {
        current := queue[index]
        index += 1
        elem := current["Element"]
        depth := current.Has("Depth") ? current["Depth"] : 0
        releaseElem := current.Has("Release") ? current["Release"] : true
        processed += 1

        info := UIAGetDirectElementInfo(elem)
        if IsObject(info) {
            record := UIABuildMatchRecord(info, depth)
            autoId := record.Has("AutomationId") ? record["AutomationId"] : ""
            if filterLower = "" || InStr(StrLower(autoId), filterLower) {
                matches.Push(record)
                LogAutoId(Format("BFS match depth={1} autoId='{2}' name='{3}'", depth, autoId, record.Has("Name") ? record["Name"] : ""))
            }
        }

        if processed < maxNodes {
            children := UIAFindChildren(uia, elem)
            if children {
                count := UIAElementArrayLength(children)
                LogAutoId(Format("BFS depth {1} adding {2} children", depth, count))
                Loop count {
                    child := UIAElementArrayGet(children, A_Index - 1)
                    if !child
                        continue
                    key := UIAPointerKey(child)
                    if key && visited.Has(key) {
                        UIARelease(child)
                        continue
                    }
                    if key
                        visited[key] := true
                    queue.Push(Map("Element", child, "Depth", depth + 1, "Release", true))
                }
                UIARelease(children)
            }
        }

        if releaseElem
            UIARelease(elem)

        if processed >= maxNodes
            break
    }

    if processed >= maxNodes {
        LogAutoId(Format("BFS stopped at node limit {1}", maxNodes))
        while index <= queue.Length {
            leftover := queue[index]
            index += 1
            if IsObject(leftover) && leftover.Has("Element") {
                extra := leftover["Element"]
                if extra
                    UIARelease(extra)
            }
        }
    }

    LogAutoId(Format("UIABreadthFirstSearch processed {1} nodes, matched {2}", processed, matches.Length))
    return matches
}

UIABuildMatchRecord(details, depth) {
    record := Map()
    automationId := details.Has("Auto") ? Trim(details["Auto"]) : ""
    className := details.Has("Class") ? Trim(details["Class"]) : ""
    name := details.Has("Name") ? Trim(details["Name"]) : ""
    typeName := details.Has("LocalizedControlType") ? Trim(details["LocalizedControlType"]) : ""
    if typeName = ""
        typeName := details.Has("ControlType") ? Trim(details["ControlType"]) : ""

    hwnd := details.Has("HWND") ? details["HWND"] : 0

    record["AutomationId"] := automationId
    record["Class"] := className
    record["UIAClass"] := className
    record["Type"] := typeName
    record["Name"] := name
    record["HWND"] := hwnd ? Format("0x{1:X}", hwnd) : ""
    record["HWNDRaw"] := hwnd
    record["Depth"] := depth

    if details.Has("Rect") && IsObject(details["Rect"])
        record["Rect"] := details["Rect"]

    return record
}

UIARawAutomationMatches(uia, rootElement, filterLower) {
    matches := []
    maxNodes := gAutoIdMaxNodes
    cond := UIAGetTrueCondition(uia)
    if !cond {
        LogAutoId("UIAGetTrueCondition returned 0")
        return matches
    }

    elements := 0
    static TREE_SCOPE_DESCENDANTS := 4
    hr := 1
    try hr := ComCall(8, rootElement, "int", TREE_SCOPE_DESCENDANTS, "ptr", cond, "ptr*", &elements)
    catch as err {
        hr := 1
        LogAutoId("Raw ComCall threw: " AutoIdErrorText(err))
    }
    LogAutoId(Format("Raw match ComCall hr={1} elementsPtr=0x{2:X}", hr, elements))
    if hr != 0 || !elements {
        if elements
            UIARelease(elements)
        LogAutoId("Raw match aborted due to hr or empty result")
        return matches
    }

    processed := 0
    try {
        count := UIAElementArrayLength(elements)
        LogAutoId(Format("Raw match element array count={1}", count))
        Loop count {
            if processed >= maxNodes
                break
            elem := UIAElementArrayGet(elements, A_Index - 1)
            if !elem
                continue
            processed += 1
            details := UIAGetDirectElementInfo(elem)
            if IsObject(details) {
                record := UIABuildMatchRecord(details, -1)
                autoId := record.Has("AutomationId") ? record["AutomationId"] : ""
                if filterLower = "" || InStr(StrLower(autoId), filterLower) {
                    matches.Push(record)
                    LogAutoId(Format("Raw match autoId='{1}' name='{2}'", autoId, record.Has("Name") ? record["Name"] : ""))
                }
            }
            UIARelease(elem)
        }
    } finally {
        UIARelease(elements)
        LogAutoId(Format("UIARawAutomationMatches processed {1} nodes, matched {2}", processed, matches.Length))
    }

    return matches
}


Main()

InitAutoIdLog() {
    if !gAutoIdDebug
        return
    global gAutoIdLogPath
    stamp := FormatTime(, "yyyy-MM-dd_HHmmss")
    gAutoIdLogPath := A_ScriptDir "\" stamp "-autoid-debug.log"
    header := Format("=== AutoId Debug Start {1} ===", stamp)
    FileAppend(header "`n", gAutoIdLogPath, "UTF-8")
}

LogAutoId(message) {
    if !gAutoIdDebug
        return
    global gAutoIdLogPath
    if gAutoIdLogPath = ""
        InitAutoIdLog()
    timestamp := FormatTime(, "HH:mm:ss")
    line := Format("[{1}] {2}", timestamp, message)
    FileAppend(line "`n", gAutoIdLogPath, "UTF-8")
}

MapToText(map) {
    if !IsObject(map)
        return "<none>"
    parts := []
    for key, value in map
        parts.Push(key ":" Trim(value))
    if parts.Length = 0
        return "<empty>"
    return "{" Join(parts, ", ") "}"
}

Join(arr, delimiter := ", ") {
    out := ""
    for index, value in arr {
        if index > 1
            out .= delimiter
        out .= value
    }
    return out
}

AutoIdErrorText(err) {
    if !IsObject(err)
        return err
    parts := []
    if err.Has("Message")
        parts.Push("msg=" err.Message)
    if err.Has("What")
        parts.Push("what=" err.What)
    if err.Has("Extra")
        parts.Push("extra=" err.Extra)
    if err.Has("File")
        parts.Push("file=" err.File)
    if err.Has("Line")
        parts.Push("line=" err.Line)
    if parts.Length = 0
        return err
    return Join(parts, "; ")
}
