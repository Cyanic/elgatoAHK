#Requires AutoHotkey v2.0
#SingleInstance Force
#Include "%A_ScriptDir%\uiaHelpers.ahk"

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
    WriteResults(outPath, matches, filterText, config)
    MsgBox Format("Found {1} matching controls. Details written to:`n{2}", matches.Length, outPath)

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
        uiaMatches := UIARawClassMatches(uia, root, filter, filterOriginal)
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

UIARawClassMatches(uia, rootElement, filterLower, filterExact) {
    matches := []
    if filterExact = ""
        return matches

    if filterLower = ""
        filterLower := StrLower(filterExact)

    collector := UIARawClassCollectorFactory(matches, filterLower)
    UIAForEachRaw(uia, rootElement, collector, 1000000)
    return matches
}

UIARawClassCollectorFactory(matches, filterLower) {
    Callback(elem, details, depth) {
        record := UIABuildMatchRecord(details, depth)
        classField := ""
        if record.Has("UIAClass")
            classField := record["UIAClass"]
        if classField = "" && record.Has("Class")
            classField := record["Class"]
        if classField != "" && InStr(StrLower(classField), filterLower)
            matches.Push(record)
        return true
    }
    return Callback
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
    rectInfo := GetWindowRectInfo(childHwnd)
    if IsObject(rectInfo)
        record["Rect"] := rectInfo

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
    if details.Has("Rect") && IsObject(details["Rect"])
        record["Rect"] := details["Rect"]

    return record
}

UIAMatchFilter(record, filter) {
    if filter = ""
        return true

    fields := ["Class", "UIAClass"]
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


GetWindowClassName(hwnd) {
    if !hwnd
        return ""
    buf := Buffer(512 * 2, 0)
    len := DllCall("GetClassNameW", "ptr", hwnd, "ptr", buf, "int", 512)
    return len ? StrGet(buf, "UTF-16") : ""
}

GetWindowRectInfo(hwnd) {
    if !hwnd
        return 0
    rectBuf := Buffer(16, 0)
    if !DllCall("GetWindowRect", "ptr", hwnd, "ptr", rectBuf.Ptr)
        return 0
    left := NumGet(rectBuf, 0, "int")
    top := NumGet(rectBuf, 4, "int")
    right := NumGet(rectBuf, 8, "int")
    bottom := NumGet(rectBuf, 12, "int")
    width := right - left
    height := bottom - top
    return Map("x", left, "y", top, "w", width, "h", height)
}


Main()