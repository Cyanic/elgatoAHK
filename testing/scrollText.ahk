#Requires AutoHotkey v2.0
#SingleInstance Force
#Include "%A_ScriptDir%\uiaHelpers.ahk"

; Automation identifiers for the viewport target and its scroll container.
global gViewportAutomationId := "qt_scrollarea_viewport"
global gScrollContainerClass := "QScrollArea"

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

    uia := GetUIAutomation()
    if !IsObject(uia) {
        MsgBox "UI Automation is not available on this system."
        return false
    }

    iniPath := EnsureConfig()
    config := LoadConfig(iniPath)
    hwnd := GetTargetWindow(config)
    if !hwnd {
        MsgBox "Target window not found. Update showClasses.ini and try again."
        return false
    }

    root := UIAElementFromHandle(uia, hwnd)
    if !root {
        MsgBox "Failed to bind UI Automation to the target window."
        return false
    }

    element := 0
    success := false
    try {
        element := ResolveTargetElement(uia, root, targetMode)
        if !element {
            MsgBox Format("Unable to locate the {1} element.", targetMode)
            return false
        }
        success := ScrollElement(element, direction, steps)
    } finally {
        if element
            UIARelease(element)
        UIARelease(root)
    }

    if !success {
        MsgBox Format("The {1} element does not support UI Automation scrolling.", targetMode)
        return false
    }
    return true
}

ResolveTargetElement(uia, root, mode) {
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
    return FindElementByAutomationId(uia, root, gViewportAutomationId)
}

ResolveScrollAreaElement(uia, root) {
    viewport := FindElementByAutomationId(uia, root, gViewportAutomationId)
    if !viewport
        return FindScrollAreaFallback(uia, root)

    scrollArea := FindAncestorByClass(uia, viewport, gScrollContainerClass)
    UIARelease(viewport)
    if scrollArea
        return scrollArea
    return FindScrollAreaFallback(uia, root)
}

FindScrollAreaFallback(uia, root) {
    target := FindElementByClassWithSize(uia, root, gScrollContainerClass, 208, 168)
    if target
        return target
    return FindElementByClassName(uia, root, gScrollContainerClass)
}

FindElementByAutomationId(uia, root, automationId) {
    if automationId = ""
        return 0

    cond := 0
    try cond := uia.CreatePropertyCondition(30011, automationId)
    catch {
        cond := 0
    }
    if !IsObject(cond)
        return 0

    condPtr := ComObjValue(cond)
    if !condPtr
        return 0

    element := 0
    static TREE_SCOPE_DESCENDANTS := 4
    try hr := ComCall(5, root, "int", TREE_SCOPE_DESCENDANTS, "ptr", condPtr, "ptr*", &element)
    catch {
        hr := -1
    }
    return (hr = 0 && element) ? element : 0
}

FindElementByClassName(uia, root, className) {
    if className = ""
        return 0

    cond := 0
    try cond := uia.CreatePropertyCondition(30012, className)
    catch {
        cond := 0
    }
    if !IsObject(cond)
        return 0

    condPtr := ComObjValue(cond)
    if !condPtr
        return 0

    element := 0
    static TREE_SCOPE_DESCENDANTS := 4
    try hr := ComCall(5, root, "int", TREE_SCOPE_DESCENDANTS, "ptr", condPtr, "ptr*", &element)
    catch {
        hr := -1
    }
    return (hr = 0 && element) ? element : 0
}

FindElementByClassWithSize(uia, root, className, width, height) {
    if className = ""
        return 0

    cond := 0
    try cond := uia.CreatePropertyCondition(30012, className)
    catch {
        cond := 0
    }
    if !IsObject(cond)
        return 0

    condPtr := ComObjValue(cond)
    if !condPtr
        return 0

    elements := 0
    static TREE_SCOPE_DESCENDANTS := 4
    try hr := ComCall(6, root, "int", TREE_SCOPE_DESCENDANTS, "ptr", condPtr, "ptr*", &elements)
    catch {
        hr := -1
    }
    if hr != 0 || !elements
        return 0

    found := 0
    try {
        count := UIAElementArrayLength(elements)
        Loop count {
            elem := UIAElementArrayGet(elements, A_Index - 1)
            if !elem
                continue
            rect := UIAGetBoundingRect(elem)
            if IsObject(rect) {
                if Abs(rect["w"] - width) <= 1 && Abs(rect["h"] - height) <= 1 {
                    found := elem
                    break
                }
            }
            UIARelease(elem)
        }
    } finally {
        UIARelease(elements)
    }
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

    Loop maxLevels {
        parent := 0
        try parent := walker.GetParentElement(current)
        catch {
            parent := 0
        }
        if !parent
            break

        parentClass := UIAGetProperty(parent, 30012)
        if StrLower(parentClass) = classLower {
            if current != element
                UIARelease(current)
            return parent
        }

        if current != element
            UIARelease(current)
        current := parent
    }

    if current != element
        UIARelease(current)
    return 0
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
    if hr != 0 || !pattern
        return false

    success := true
    try {
        vertAmount := (direction = "up") ? SCROLL_SMALL_DECREMENT : SCROLL_SMALL_INCREMENT
        Loop steps {
            try callHr := ComCall(9, pattern, "int", SCROLL_NO_AMOUNT, "int", vertAmount)
            catch {
                callHr := -1
            }
            if callHr != 0 {
                success := false
                break
            }
            Sleep 60
        }
    } finally {
        UIARelease(pattern)
    }
    return success
}

; Hotkeys for scrolling the viewport element.
^!F1::ScrollViewport("up")
^!F2::ScrollViewport("down")

; Hotkeys for scrolling the containing QScrollArea element.
^!F3::ScrollScrollArea("up")
^!F4::ScrollScrollArea("down")
