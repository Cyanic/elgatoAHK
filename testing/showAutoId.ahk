#Requires AutoHotkey v2.0
#SingleInstance Force
#Include "%A_ScriptDir%\uiaHelpers.ahk"

global gAutoIdMaxNodes := 300000

Main() {
    iniPath := EnsureConfig()
    config := LoadConfig(iniPath)
    if !config.Has("ClassNN") && !config.Has("Process") {
        MsgBox "No window configuration found in " iniPath
        return
    }

    prompt := InputBox("Enter all or part of the AutomationId to search for:", "AutomationId Lookup")
    if prompt.Result != "OK" || prompt.Value = "" {
        MsgBox "No AutomationId provided. Exiting."
        return
    }

    targetHwnd := GetTargetWindow(config)
    if !targetHwnd {
        MsgBox "Target window not found. Check showClasses.ini"
        return
    }

    filter := Trim(prompt.Value)
    matches := FindAutomationMatches(targetHwnd, filter)

    dateStamp := FormatTime(, "yyyy-MM-dd")
    outPath := A_ScriptDir "\" dateStamp "-autoid-output.txt"
    metadata := Map("Mode", "Automation")
    for key, value in config
        metadata[key] := value
    WriteResults(outPath, matches, filter, metadata)
    MsgBox Format("Found {1} matching controls. Details written to:`n{2}", matches.Length, outPath)
}


FindAutomationMatches(hwnd, filterText) {
    matches := []
    uia := GetUIAutomation()
    if !IsObject(uia)
        return matches

    root := UIAElementFromHandle(uia, hwnd)
    if !root {
        MsgBox "UI Automation could not bind to the target window."
        return matches
    }

    try {
        filterTrim := Trim(filterText)
        filterLower := StrLower(filterTrim)

        if filterTrim != ""
            matches := UIAExactAutomationMatches(uia, root, filterTrim)

        if matches.Length = 0
            matches := UIABreadthFirstSearch(uia, root, filterLower)

        if matches.Length = 0
            matches := UIARawAutomationMatches(uia, root, filterLower)
    } finally {
        UIARelease(root)
    }
    return matches
}

UIAExactAutomationMatches(uia, rootElement, filterExact) {
    matches := []
    if filterExact = ""
        return matches

    cond := ""
    try cond := uia.CreatePropertyCondition(30011, filterExact)
    catch {
        cond := ""
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
    catch {
        hr := 1
    }
    if hr != 0 || !elements {
        try ObjRelease(cond)
        if elements
            UIARelease(elements)
        return matches
    }

    try {
        count := UIAElementArrayLength(elements)
        Loop count {
            elem := UIAElementArrayGet(elements, A_Index - 1)
            if !elem
                continue
            details := UIAGetDirectElementInfo(elem)
            if IsObject(details) {
                record := UIABuildMatchRecord(details, -1)
                matches.Push(record)
            }
            UIARelease(elem)
        }
    } finally {
        UIARelease(elements)
        try ObjRelease(cond)
    }

    return matches
}


UIABreadthFirstSearch(uia, rootElement, filterLower) {
    maxNodes := gAutoIdMaxNodes
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
            if filterLower = "" || InStr(StrLower(autoId), filterLower)
                matches.Push(record)
        }

        if processed < maxNodes {
            children := UIAFindChildren(uia, elem)
            if children {
                count := UIAElementArrayLength(children)
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
    filterLower := StrLower(filterLower)
    collector := UIARawAutomationCollectorFactory(matches, filterLower)
    UIAForEachRaw(uia, rootElement, collector, gAutoIdMaxNodes)
    return matches
}

UIARawAutomationCollectorFactory(matches, filterLower) {
    Callback(elem, details, depth) {
        record := UIABuildMatchRecord(details, depth)
        autoId := record.Has("AutomationId") ? record["AutomationId"] : ""
        if filterLower = "" || InStr(StrLower(autoId), filterLower)
            matches.Push(record)
        return true
    }
    return Callback
}


Main()
