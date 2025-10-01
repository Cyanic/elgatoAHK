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

GetUIAutomation() {
    static uia := ""
    if IsObject(uia)
        return uia

    try {
        uia := ComObject("UIAutomationClient.CUIAutomation")
    } catch {
        try {
            uia := ComObject("{FF48DBA4-60EF-4201-AA87-54103EEF594E}", "{30CBE57D-D9D0-452A-AB13-7AC5AC4825EE}")
        } catch {
            uia := ""
        }
    }
    return uia
}

UIAElementFromHandle(uia, hwnd) {
    elementPtr := 0
    hr := ComCall(7, uia, "ptr", hwnd, "ptr*", &elementPtr)
    return hr = 0 ? elementPtr : 0
}

UIARelease(elementPtr) {
    if !elementPtr
        return
    vtable := NumGet(elementPtr, 0, "ptr")
    fn := NumGet(vtable, 2 * A_PtrSize, "ptr")
    DllCall(fn, "ptr", elementPtr, "uint")
}

UIAAddRef(elementPtr) {
    if !elementPtr
        return
    vtable := NumGet(elementPtr, 0, "ptr")
    fn := NumGet(vtable, 1 * A_PtrSize, "ptr")
    DllCall(fn, "ptr", elementPtr, "uint")
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

UIACreatePropertyCondition(uia, propertyId, value) {
    if !IsObject(uia)
        return 0
    condPtr := 0
    variant := ""
    try {
        variant := ComValue(8, value)
    } catch {
        variant := ComValue(8, "")
    }
    hr := 1
    try {
        hr := ComCall(23, uia, "int", propertyId, "ptr", variant, "ptr*", &condPtr)
    } catch {
        hr := 1
    }
    if hr != 0 || !condPtr
        return 0
    return ComObject(ComValue(13, condPtr))
}

UIAElementArrayLength(arrayPtr) {
    if !arrayPtr
        return 0
    length := 0
    try {
        ComCall(3, arrayPtr, "int*", &length)
    }
    return length
}

UIAElementArrayGet(arrayPtr, index) {
    if !arrayPtr
        return 0
    element := 0
    try {
        ComCall(4, arrayPtr, "int", index, "ptr*", &element)
    }
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

UIAPointerKey(ptr) {
    if !ptr
        return 0
    key := 0
    try {
        key := ComObjValue(ptr)
    } catch {
        key := ptr
    }
    return key
}

UIAForEachRaw(uia, rootElement, callback, maxNodes := 1000000) {
    if !IsObject(uia) || !rootElement
        return 0
    if !IsObject(callback)
        return 0

    walker := ""
    try {
        walker := uia.CreateTreeWalker(uia.RawViewCondition)
    } catch {
        walker := ""
    }
    if !IsObject(walker)
        return 0

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

    if !start {
        walker := ""
        return 0
    }

    stack := []
    stack.Push(Map("Element", start, "Depth", 0, "Release", releaseStart))
    visited := Map()
    processed := 0
    stop := false

    while stack.Length {
        current := stack.Pop()
        if !IsObject(current)
            continue
        elem := current.Has("Element") ? current["Element"] : 0
        depth := current.Has("Depth") ? current["Depth"] : 0
        releaseElem := current.Has("Release") ? current["Release"] : true
        if !elem
            continue

        key := UIAPointerKey(elem)
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
            continueTraversal := 1
            try {
                continueTraversal := callback.Call(elem, details, depth)
            } catch {
                continueTraversal := 1
            }
            if continueTraversal == 0
                stop := true
        }

        if !stop && processed < maxNodes {
            child := walker.GetFirstChildElement(elem)
            while child {
                stack.Push(Map("Element", child, "Depth", depth + 1, "Release", true))
                sibling := walker.GetNextSiblingElement(child)
                child := sibling
            }
        } else if processed >= maxNodes {
            stop := true
        }

        if releaseElem
            UIARelease(elem)

        if stop
            break
    }

    while stack.Length {
        leftover := stack.Pop()
        if IsObject(leftover) && leftover.Has("Element") {
            extra := leftover["Element"]
            if extra
                UIARelease(extra)
        }
    }

    walker := ""
    return processed
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
    return id != "" ? id : ""
}

UIAGetBoundingRect(elementPtr) {
    if !elementPtr
        return 0
    variantSize := (A_PtrSize = 8) ? 24 : 16
    static GET_CURRENT_PROPERTY_VALUE := 10
    static GET_CURRENT_PROPERTY_VALUE_EX := 11
    rectVariant := Buffer(variantSize, 0)

    attempt := -1
    try {
        attempt := ComCall(GET_CURRENT_PROPERTY_VALUE, elementPtr, "int", 30001, "ptr", rectVariant.Ptr)
    } catch {
        attempt := -1
    }
    if attempt != 0 {
        try {
            attempt := ComCall(GET_CURRENT_PROPERTY_VALUE_EX, elementPtr, "int", 30001, "int", true, "ptr", rectVariant.Ptr)
        } catch {
            attempt := -1
        }
    }
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

    attempts := [Map("Index", GET_CURRENT_PROPERTY_VALUE, "Extra", []), Map("Index", GET_CURRENT_PROPERTY_VALUE_EX, "Extra", [true])]

    for attempt in attempts {
        variant := Buffer(variantSize, 0)
        params := [attempt["Index"], elementPtr, "int", propertyId]
        for param in attempt["Extra"] {
            params.Push("int")
            params.Push(param)
        }
        params.Push("ptr")
        params.Push(variant.Ptr)

        try {
            hr := ComCall(params*)
        } catch {
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
    try {
        DllCall("OleAut32\\VariantClear", "ptr", ptr)
    }
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

WriteResults(path, results, searchTerm := "", config := 0) {
    timestamp := FormatTime(, "yyyy-MM-dd HH:mm:ss")
    filterLabel := searchTerm != "" ? searchTerm : "<none>"
    classLabel := ConfigValueOrDefault(config, "ClassNN")
    processLabel := ConfigValueOrDefault(config, "Process")
    header := Format("Scan - {1} | Filter: {2} | ClassNN: {3} | Process: {4}", timestamp, filterLabel, classLabel, processLabel)
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
            rectText := "<none>"
            if item.Has("Rect") && IsObject(item["Rect"]) {
                rect := item["Rect"]
                rectText := Format("({},{},{},{})", rect["x"], rect["y"], rect["w"], rect["h"])
            }
            location := FormatLocation(item)
            line := Format("Class: {1}`tUIAClass: {2}`tType: {3}`tAutomationId: {4}`t`t`tName: {5}`tHWND: {6}`tRect: {7}`tDepth: {8}`tLocation{9}", class, uiaClass, typeName, autoId, ctrlName, hwnd, rectText, item.Has("Depth") ? item["Depth"] : "?", location)
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
