#Requires AutoHotkey v2.0
#SingleInstance Force
#Include "%A_ScriptDir%\uiaHelpers.ahk"

; Automation identifiers for the viewport target and its scroll container.
global gViewportAutomationId := "qt_scrollarea_viewport"
global gScrollContainerClass := "QScrollArea"
global gDebugEnabled := true
global gAutomationSearchLimit := 150000

DebugLog(message, details := "") {
    if !gDebugEnabled
        return

    try {
        timestamp := FormatTime(, "yyyy-MM-dd HH:mm:ss")
        text := details != "" ? Format("[{1}] {2} | {3}", timestamp, message, details) : Format("[{1}] {2}", timestamp, message)
        FileAppend(text "`n", GetDebugLogPath(), "UTF-8")
    } catch {
        ; Swallow logging errors to keep scroll hotkeys functional.
    }
}

GetDebugLogPath() {
    dateStamp := FormatTime(, "yyyy-MM-dd")
    return A_ScriptDir "\" dateStamp "-scroll-debug.txt"
}

FormatPtr(value) {
    return value ? Format("0x{:X}", value) : "0x0"
}

DebugDescribeConfig(config) {
    if !IsObject(config)
        return "<none>"

    items := []
    for key, val in config {
        items.Push(Format("{1}={2}", key, val))
    }
    return items.Length ? JoinWithComma(items) : "<empty>"
}

DebugDescribeElement(element) {
    if !element
        return "<null>"

    auto := UIAGetProperty(element, 30011)
    name := UIAGetProperty(element, 30005)
    class := UIAGetProperty(element, 30012)
    rect := UIAGetBoundingRect(element)
    rectText := DebugDescribeRect(rect)
    return Format("ptr={1} | AutomationId={2} | Name={3} | Class={4} | Rect={5}", FormatPtr(element), auto != "" ? auto : "<none>", name != "" ? name : "<none>", class != "" ? class : "<none>", rectText)
}

DebugDescribeRect(rect) {
    if !IsObject(rect)
        return "<none>"
    x := rect.Has("x") ? Round(rect["x"]) : "?"
    y := rect.Has("y") ? Round(rect["y"]) : "?"
    w := rect.Has("w") ? Round(rect["w"]) : "?"
    h := rect.Has("h") ? Round(rect["h"]) : "?"
    return Format("x={1}, y={2}, w={3}, h={4}", x, y, w, h)
}

JoinWithComma(items) {
    if !IsObject(items)
        return ""
    output := ""
    for item in items {
        if output != ""
            output .= ", "
        output .= item
    }
    return output
}

ScrollViewport(direction := "down", steps := 1) {
    return ScrollResolved(direction, steps, "viewport")
}

ScrollScrollArea(direction := "down", steps := 1) {
    return ScrollResolved(direction, steps, "scroll area")
}

ScrollResolved(direction, steps, targetMode) {
    direction := StrLower(Trim(direction))
    if direction != "up"
        direction := "down"

    steps := Max(1, Round(Abs(steps)))

    DebugLog("ScrollResolved invoked", Format("direction={1} | steps={2} | targetMode={3}", direction, steps, targetMode))

    uia := GetUIAutomation()
    if !IsObject(uia) {
        DebugLog("UI Automation unavailable")
        MsgBox "UI Automation is not available on this system."
        return false
    }

    iniPath := EnsureConfig()
    config := LoadConfig(iniPath)
    targetDesc := BuildWindowSearchLabel(config)
    DebugLog("Loaded configuration", Format("path={1} | config={2}", iniPath, DebugDescribeConfig(config)))
    hwnd := GetTargetWindow(config)
    if !hwnd {
        DebugLog("Target window not found", Format("expected={1}", targetDesc))
        MsgBox Format("Target window not found. Expected match for:`n{1}`nUpdate showClasses.ini and try again.", targetDesc)
        return false
    }

    root := UIAElementFromHandle(uia, hwnd)
    if !root {
        DebugLog("UIAElementFromHandle failed", Format("hwnd={1}", FormatPtr(hwnd)))
        MsgBox "Failed to bind UI Automation to the target window."
        return false
    }

    element := 0
    success := false
    try {
        element := ResolveTargetElement(uia, root, targetMode)
        if !element {
            DebugLog("ResolveTargetElement returned null", Format("targetMode={1}", targetMode))
            MsgBox Format("Unable to locate the {1} element.", targetMode)
            return false
        }
        DebugLog("Resolved element", Format("targetMode={1} | element={2}", targetMode, DebugDescribeElement(element)))
        success := ScrollElement(element, direction, steps)
    } finally {
        if element
            UIARelease(element)
        UIARelease(root)
    }

    if !success {
        DebugLog("ScrollElement failed", Format("targetMode={1} | direction={2} | steps={3}", targetMode, direction, steps))
        MsgBox Format("The {1} element does not support UI Automation scrolling.", targetMode)
        return false
    }
    DebugLog("ScrollResolved completed", Format("targetMode={1} | direction={2} | steps={3}", targetMode, direction, steps))
    return true
}

ResolveTargetElement(uia, root, mode) {
    DebugLog("ResolveTargetElement", Format("mode={1}", mode))
    switch mode {
        case "viewport":
            return ResolveViewportElement(uia, root)
        case "scroll area":
            return ResolveScrollAreaElement(uia, root)
        default:
            return 0
    }
}

ResolveViewportElement(uia, root) {
    DebugLog("ResolveViewportElement", Format("automationId={1}", gViewportAutomationId))
    element := FindElementByAutomationId(uia, root, gViewportAutomationId)
    if element
        DebugLog("Viewport element located", DebugDescribeElement(element))
    else
        DebugLog("Viewport element missing", Format("automationId={1}", gViewportAutomationId))
    return element
}

ResolveScrollAreaElement(uia, root) {
    viewport := FindElementByAutomationId(uia, root, gViewportAutomationId)
    if !viewport
        return FindScrollAreaFallback(uia, root)

    scrollArea := FindAncestorByClass(uia, viewport, gScrollContainerClass)
    DebugLog("ResolveScrollAreaElement", Format("viewport={1} | ancestorResult={2}", DebugDescribeElement(viewport), scrollArea ? DebugDescribeElement(scrollArea) : "<null>"))
    UIARelease(viewport)
    if scrollArea
        return scrollArea
    return FindScrollAreaFallback(uia, root)
}

FindScrollAreaFallback(uia, root) {
    DebugLog("FindScrollAreaFallback", Format("class={1}", gScrollContainerClass))
    target := FindElementByClassWithSize(uia, root, gScrollContainerClass, 208, 168)
    if target {
        DebugLog("Fallback located size-matched scroll area", DebugDescribeElement(target))
        return target
    }
    alt := FindElementByClassName(uia, root, gScrollContainerClass)
    if alt
        DebugLog("Fallback located class match", DebugDescribeElement(alt))
    else
        DebugLog("Fallback failed to locate QScrollArea instance")
    return alt
}

FindElementByAutomationId(uia, root, automationId) {
    if automationId = ""
        return 0

    cond := 0
    try cond := uia.CreatePropertyCondition(30011, automationId)
    catch {
        cond := 0
    }
    if !IsObject(cond) {
        DebugLog("FindElementByAutomationId condition missing", Format("automationId={1}", automationId))
        return 0
    }

    condPtr := ComObjValue(cond)
    if !condPtr {
        DebugLog("FindElementByAutomationId condPtr null", Format("automationId={1}", automationId))
        try ObjRelease(cond)
        return 0
    }

    element := 0
    static TREE_SCOPE_DESCENDANTS := 4
    try hr := ComCall(5, root, "int", TREE_SCOPE_DESCENDANTS, "ptr", condPtr, "ptr*", &element)
    catch {
        hr := -1
    }

    if hr != 0
        DebugLog("FindElementByAutomationId failure", Format("automationId={1} | hr={2}", automationId, hr))
    else
        DebugLog("FindElementByAutomationId success", Format("automationId={1} | element={2}", automationId, element ? DebugDescribeElement(element) : "<null>"))

    result := (hr = 0 && element) ? element : 0

    try ObjRelease(cond)
    if result
        return result

    DebugLog("FindElementByAutomationId fallback", Format("automationId={1}", automationId))
    return FindElementByAutomationIdFallback(uia, root, automationId)
}

FindElementByClassName(uia, root, className) {
    if className = ""
        return 0

    cond := 0
    try cond := uia.CreatePropertyCondition(30012, className)
    catch {
        cond := 0
    }
    if !IsObject(cond) {
        DebugLog("FindElementByClassName condition missing", Format("className={1}", className))
        return 0
    }

    condPtr := ComObjValue(cond)
    if !condPtr {
        DebugLog("FindElementByClassName condPtr null", Format("className={1}", className))
        try ObjRelease(cond)
        return 0
    }

    element := 0
    static TREE_SCOPE_DESCENDANTS := 4
    try hr := ComCall(5, root, "int", TREE_SCOPE_DESCENDANTS, "ptr", condPtr, "ptr*", &element)
    catch {
        hr := -1
    }

    if hr != 0
        DebugLog("FindElementByClassName failure", Format("className={1} | hr={2}", className, hr))
    else
        DebugLog("FindElementByClassName success", Format("className={1} | element={2}", className, element ? DebugDescribeElement(element) : "<null>"))

    result := (hr = 0 && element) ? element : 0

    try ObjRelease(cond)
    return result
}

FindElementByAutomationIdFallback(uia, root, automationId) {
    if automationId = ""
        return 0

    needle := StrLower(automationId)
    cond := UIAGetTrueCondition(uia)
    if !cond
        return 0

    elements := 0
    static TREE_SCOPE_DESCENDANTS := 4
    hr := 1
    try hr := ComCall(8, root, "int", TREE_SCOPE_DESCENDANTS, "ptr", cond, "ptr*", &elements)
    catch {
        hr := 1
    }
    if hr != 0 || !elements {
        DebugLog("FindElementByAutomationId fallback query failed", Format("automationId={1} | hr={2}", automationId, hr))
        return 0
    }

    found := 0
    processed := 0
    try {
        count := UIAElementArrayLength(elements)
        Loop count {
            if processed >= gAutomationSearchLimit
                break
            elem := UIAElementArrayGet(elements, A_Index - 1)
            if !elem
                continue
            processed += 1
            details := UIAGetDirectElementInfo(elem)
            if IsObject(details) {
                autoId := details.Has("Auto") ? Trim(details["Auto"]) : ""
                if autoId != "" && InStr(StrLower(autoId), needle) {
                    found := elem
                    DebugLog("FindElementByAutomationId fallback match", DebugDescribeElement(elem))
                    break
                }
            }
            UIARelease(elem)
        }
    } finally {
        UIARelease(elements)
    }

    if !found
        DebugLog("FindElementByAutomationId fallback exhausted", Format("automationId={1} | processed={2}", automationId, processed))
    return found
}

FindElementByClassWithSize(uia, root, className, width, height) {
    if className = ""
        return 0

    cond := 0
    try cond := uia.CreatePropertyCondition(30012, className)
    catch {
        cond := 0
    }
    if !IsObject(cond) {
        DebugLog("FindElementByClassWithSize condition missing", Format("className={1}", className))
        return 0
    }

    condPtr := ComObjValue(cond)
    if !condPtr {
        DebugLog("FindElementByClassWithSize condPtr null", Format("className={1}", className))
        try ObjRelease(cond)
        return 0
    }

    elements := 0
    static TREE_SCOPE_DESCENDANTS := 4
    try hr := ComCall(6, root, "int", TREE_SCOPE_DESCENDANTS, "ptr", condPtr, "ptr*", &elements)
    catch {
        hr := -1
    }
    if hr != 0 || !elements {
        DebugLog("FindElementByClassWithSize query failed", Format("className={1} | hr={2}", className, hr))
        try ObjRelease(cond)
        if elements
            UIARelease(elements)
        return 0
    }

    found := 0
    try {
        count := UIAElementArrayLength(elements)
        DebugLog("FindElementByClassWithSize candidates", Format("className={1} | expectedSize={2}x{3} | count={4}", className, width, height, count))
        Loop count {
            elem := UIAElementArrayGet(elements, A_Index - 1)
            if !elem
                continue
            rect := UIAGetBoundingRect(elem)
            DebugLog("ClassWithSize candidate", Format("index={1} | element={2}", A_Index - 1, DebugDescribeElement(elem)))
            if IsObject(rect) {
                if Abs(rect["w"] - width) <= 1 && Abs(rect["h"] - height) <= 1 {
                    found := elem
                    DebugLog("ClassWithSize match", Format("index={1} | element={2}", A_Index - 1, DebugDescribeElement(elem)))
                    break
                }
            }
            UIARelease(elem)
        }
    } finally {
        UIARelease(elements)
    }

    try ObjRelease(cond)
    return found
}

FindAncestorByClass(uia, element, className, maxLevels := 10) {
    if !element || className = ""
        return 0

    walker := ""
    try walker := uia.CreateTreeWalker(uia.RawViewCondition)
    catch {
        walker := ""
    }
    if !IsObject(walker)
        return 0

    current := element
    classLower := StrLower(className)
    DebugLog("FindAncestorByClass", Format("start={1} | className={2} | maxLevels={3}", DebugDescribeElement(current), className, maxLevels))

    result := 0
    try {
        Loop maxLevels {
            parent := 0
            try parent := walker.GetParentElement(current)
            catch {
                parent := 0
            }
            if !parent {
                DebugLog("FindAncestorByClass parent lookup ended", Format("reason=no-parent | current={1}", DebugDescribeElement(current)))
                break
            }

            parentClass := UIAGetProperty(parent, 30012)
            DebugLog("FindAncestorByClass candidate", Format("parent={1} | parentClass={2}", DebugDescribeElement(parent), parentClass))
            if StrLower(parentClass) = classLower {
                if current != element
                    UIARelease(current)
                current := element
                result := parent
                DebugLog("FindAncestorByClass matched", DebugDescribeElement(result))
                break
            }

            if current != element
                UIARelease(current)
            current := parent
        }
    } finally {
        if current != element
            UIARelease(current)
        try {
            if IsObject(walker)
                ObjRelease(walker)
        }
    }

    return result
}

ScrollElement(element, direction, steps) {
    if !element
        return false

    static UIA_ScrollPatternId := 10004
    static SCROLL_NO_AMOUNT := 2
    static SCROLL_SMALL_DECREMENT := 1
    static SCROLL_SMALL_INCREMENT := 4

    pattern := 0
    try hr := ComCall(13, element, "int", UIA_ScrollPatternId, "ptr*", &pattern)
    catch {
        hr := -1
    }
    if hr != 0 || !pattern {
        DebugLog("ScrollElement missing scroll pattern", Format("element={1} | hr={2}", DebugDescribeElement(element), hr))
        return false
    }

    success := true
    try {
        vertAmount := (direction = "up") ? SCROLL_SMALL_DECREMENT : SCROLL_SMALL_INCREMENT
        DebugLog("ScrollElement begin", Format("element={1} | direction={2} | steps={3} | vertAmount={4}", DebugDescribeElement(element), direction, steps, vertAmount))
        Loop steps {
            try callHr := ComCall(9, pattern, "int", SCROLL_NO_AMOUNT, "int", vertAmount)
            catch {
                callHr := -1
            }
            if callHr != 0 {
                DebugLog("ScrollElement step failed", Format("step={1} | hr={2}", A_Index, callHr))
                success := false
                break
            }
            DebugLog("ScrollElement step succeeded", Format("step={1}", A_Index))
            Sleep 60
        }
    } finally {
        UIARelease(pattern)
    }
    DebugLog("ScrollElement completed", Format("element={1} | success={2}", DebugDescribeElement(element), success))
    return success
}

BuildWindowSearchLabel(config) {
    if !IsObject(config)
        return "<none>"

    parts := []
    if config.Has("ClassNN") {
        classToken := "ahk_class " config["ClassNN"]
        parts.Push(classToken)
    }
    if config.Has("Process") {
        processToken := "ahk_exe " config["Process"]
        parts.Push(processToken)
    }

    if parts.Length = 0
        return "<none>"

    return JoinWithSpace(parts)
}

JoinWithSpace(items) {
    if !IsObject(items)
        return ""
    text := ""
    for item in items {
        if text != ""
            text .= " "
        text .= item
    }
    return text
}

; Hotkeys for scrolling the viewport element.
^!F1:: ScrollViewport("up")
^!F2:: ScrollViewport("down")

; Hotkeys for scrolling the containing QScrollArea element.
^!F3:: ScrollScrollArea("up")
^!F4:: ScrollScrollArea("down")
