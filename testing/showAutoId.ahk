#Requires AutoHotkey v2.0
#SingleInstance Force

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

JoinLines(arr) {
    out := ""
    for index, value in arr {
        if index > 1
            out .= "`n"
        out .= value
    }
    return out
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

UIAGetProperty(elementPtr, propertyId) {
    if !elementPtr
        return ""
    variantSize := (A_PtrSize = 8) ? 24 : 16
    static GET_CURRENT_PROPERTY_VALUE := 10
    static GET_CURRENT_PROPERTY_VALUE_EX := 11

    attempts := [{ index: GET_CURRENT_PROPERTY_VALUE, extra: [] }, { index: GET_CURRENT_PROPERTY_VALUE_EX, extra: [true] }]

    for attempt in attempts {
        variant := Buffer(variantSize, 0)
        params := [attempt.index, elementPtr, "int", propertyId]
        for param in attempt.extra {
            params.Push("int")
            params.Push(param)
        }
        params.Push("ptr")
        params.Push(variant.Ptr)

        try hr := ComCall(params*)
        catch {
            hr := -1
        }
        if hr = 0 {
            value := UIAVariantToText(variant)
            UIATryVariantClear(variant)
            if value != ""
                return value
        }
    }

    return ""
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

UIAElementFromHandle(uia, hwnd) {
    elementPtr := 0
    hr := ComCall(7, uia, "ptr", hwnd, "ptr*", &elementPtr)
    return hr = 0 ? elementPtr : 0
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

UIARelease(elementPtr) {
    if !elementPtr
        return
    vtable := NumGet(elementPtr, 0, "ptr")
    fn := NumGet(vtable, 2 * A_PtrSize, "ptr")
    DllCall(fn, "ptr", elementPtr, "uint")
}

Main()