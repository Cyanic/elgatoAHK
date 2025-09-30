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
    WriteResults(outPath, matches, filterText, config)
    MsgBox Format("Found {1} matching controls. Details written to:`n{2}", matches.Length, outPath)

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

    walker := ""
    try {
        walker := uia.CreateTreeWalker(uia.RawViewCondition)
    } catch {
        walker := ""
    }
    if !IsObject(walker)
        return matches

    start := rootElement
    releaseStart := false
    try {
        normalized := walker.NormalizeElement(rootElement)
        if normalized {
            start := normalized
            if start != rootElement
                releaseStart := true
        }
    } catch {
    }

    if !start
        return matches

    stack := []
    stack.Push(Map("Element", start, "Depth", 0, "Release", releaseStart))
    visited := Map()

    maxNodes := 1000000
    processed := 0

    while stack.Length {
        current := stack.Pop()
        elem := current["Element"]
        depth := current["Depth"]
        releaseElem := current.Has("Release") ? current["Release"] : true
        if !elem
            continue

        key := 0
        try {
            key := ComObjValue(elem)
        }
        catch {
            key := elem
        }
        if key && visited.Has(key) {
            if releaseElem
                UIARelease(elem)
            continue
        }
        if key
            visited[key] := true

        processed += 1
        details := UIAGetDirectElementInfo(elem)
        if IsObject(details) {
            record := UIABuildMatchRecord(details, depth)
            classField := record.Has("UIAClass") ? record["UIAClass"] : ""
            if classField = ""
                classField := record.Has("Class") ? record["Class"] : ""
            if classField != "" && InStr(StrLower(classField), filterLower)
                matches.Push(record)
        }

        if processed > maxNodes {
            if releaseElem
                UIARelease(elem)
            break
        }

        child := walker.GetFirstChildElement(elem)
        while child {
            stack.Push(Map("Element", child, "Depth", depth + 1, "Release", true))
            sibling := walker.GetNextSiblingElement(child)
            child := sibling
        }

        if releaseElem
            UIARelease(elem)
    }

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

WriteResults(path, results, searchTerm := "", config := 0) {
    timestamp := FormatTime(, "yyyy-MM-dd HH:mm:ss")
    filterLabel := searchTerm != "" ? searchTerm : "<none>"
    classLabel := ConfigValueOrDefault(config, "ClassNN")
    processLabel := ConfigValueOrDefault(config, "Process")
    header := Format("Control Class Scan - {1} | Filter: {2} | ClassNN: {3} | Process: {4}", timestamp, filterLabel, classLabel, processLabel)
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
            hwnd := item.Has("HWND") && item["HWND"] != "" ? item["HWND"] : "<none>"
            location := FormatLocation(item)
            line := Format("Class: {1}`tUIAClass: {2}`tType: {3}`tAutomationId: {4}`tName: {5}`tHWND: {6}`t{7}", class, uiaClass, typeName, autoId, ctrlName, hwnd, location)
            lines.Push(line)
        }
    }
    FileAppend(JoinLines(lines) "`n`n", path, "UTF-8")
}

ConfigValueOrDefault(config, key) {
    if !IsObject(config)
        return "<none>"
    if !config.Has(key)
        return "<none>"
    value := Trim(config[key])
    return value != "" ? value : "<none>"
}

FormatLocation(item) {
    if IsObject(item) && item.Has("Rect") && IsObject(item["Rect"]) {
        rect := item["Rect"]
        x := rect.Has("x") ? Round(rect["x"]) : 0
        y := rect.Has("y") ? Round(rect["y"]) : 0
        w := rect.Has("w") ? Round(rect["w"]) : 0
        h := rect.Has("h") ? Round(rect["h"]) : 0
        return Format("Location:{x: {1} y: {2} w:{3}, h: {4}}", x, y, w, h)
    }
    return "Location:{x: <none> y: <none> w:<none>, h: <none>}"
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

UIAElementFromHandle(uia, hwnd) {
    elementPtr := 0
    hr := ComCall(7, uia, "ptr", hwnd, "ptr*", &elementPtr)
    return hr = 0 ? elementPtr : 0
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

Main()

