#Requires AutoHotkey v2.0
#SingleInstance Force
#Include "%A_ScriptDir%\uiaHelpers.ahk"

global gAutoIdMaxNodes := 75000

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
    WriteResults(outPath, matches, filter)
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
        needle := StrLower(Trim(filterText))
        matches := UIABreadthFirstSearch(uia, root, needle)
    } finally {
        UIARelease(root)
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

UIAPointerKey(ptr) {
    if !ptr
        return 0
    key := 0
    try key := ComObjValue(ptr)
    catch {
        key := ptr
    }
    return key
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

WriteResults(path, results, searchTerm := "") {
    timestamp := FormatTime(, "yyyy-MM-dd HH:mm:ss")
    filterLabel := searchTerm != "" ? searchTerm : "<none>"
    header := Format("AutomationId Scan - {1} | Filter: {2}", timestamp, filterLabel)
    lines := []
    lines.Push(header)
    if results.Length = 0 {
        lines.Push("No matches found.")
    } else {
        for item in results {
            autoId := item.Has("AutomationId") && item["AutomationId"] != "" ? item["AutomationId"] : "<none>"
            className := item.Has("Class") && item["Class"] != "" ? item["Class"] : "<none>"
            typeName := item.Has("Type") && item["Type"] != "" ? item["Type"] : "<none>"
            ctrlName := item.Has("Name") && item["Name"] != "" ? item["Name"] : "<none>"
            hwnd := item.Has("HWND") && item["HWND"] != "" ? item["HWND"] : "<none>"
            rectText := "<none>"
            if item.Has("Rect") && IsObject(item["Rect"]) {
                rect := item["Rect"]
                rectText := Format("({},{},{},{})", rect["x"], rect["y"], rect["w"], rect["h"])
            }
            line := Format("AutomationId: {1}`tClass: {2}`tType: {3}`tName: {4}`tHWND: {5}`tRect: {6}`tDepth: {7}",
                autoId, className, typeName, ctrlName, hwnd, rectText, item.Has("Depth") ? item["Depth"] : "?")
            lines.Push(line)
        }
    }
    FileAppend(JoinLines(lines) "`n`n", path, "UTF-8")
}















Main()
