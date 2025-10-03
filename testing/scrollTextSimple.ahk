#Requires AutoHotkey v2.0
#Include "%A_ScriptDir%\uiaHelpers.ahk"

; Function to get the scrollable element
getScrollElement() {
    hwnd := WinExist("ahk_exe Camera Hub.exe")
    if !hwnd {
        MsgBox("Camera Hub not found.")
        return 0
    }

    uia := GetUIAutomation()
    if !IsObject(uia) {
        MsgBox("UI Automation not available.")
        return 0
    }

    root := UIAElementFromHandle(uia, hwnd)
    if !root {
        MsgBox("Unable to bind UI Automation to Camera Hub.")
        return 0
    }

    try {
        scrollEl := FindElementByAutomationIdSimple(uia, root, "qt_scrollarea_viewport")
    } finally {
        UIARelease(root)
    }

    if !scrollEl {
        MsgBox("Scrollable element not found.")
        return 0
    }

    return scrollEl
}

ScrollWithUIA(el, direction := "down") {
    static UIA_ScrollPatternId := 10004
    static SCROLL_NO_AMOUNT := 2
    static SCROLL_SMALL_DECREMENT := 1
    static SCROLL_SMALL_INCREMENT := 4

    if !el {
        return
    }

    pattern := 0
    try hr := ComCall(13, el, "int", UIA_ScrollPatternId, "ptr*", &pattern)
    catch {
        hr := -1
    }

    if hr != 0 || !pattern {
        MsgBox("Scroll pattern not available.")
        return
    }

    direction := StrLower(direction)
    vertAmount := direction = "up" ? SCROLL_SMALL_DECREMENT : SCROLL_SMALL_INCREMENT

    try {
        callHr := ComCall(9, pattern, "int", SCROLL_NO_AMOUNT, "int", vertAmount)
        if callHr != 0 {
            MsgBox(Format("Failed to invoke scroll pattern. (hr={1})", callHr))
        }
    } catch Error as err {
        MsgBox("Failed to invoke scroll pattern:`n" err.Message)
    } finally {
        UIARelease(pattern)
    }
}

FindElementByAutomationIdSimple(uia, root, automationId) {
    if automationId = ""
        return 0

    cond := UIACreatePropertyCondition(uia, 30011, automationId)
    if !IsObject(cond)
        return 0

    condPtr := ComObjValue(cond)
    if !condPtr {
        try ObjRelease(cond)
        return 0
    }

    element := 0
    static TREE_SCOPE_DESCENDANTS := 4
    try hr := ComCall(5, root, "int", TREE_SCOPE_DESCENDANTS, "ptr", condPtr, "ptr*", &element)
    catch {
        hr := -1
    }

    try ObjRelease(cond)

    if hr != 0 || !element {
        if element
            UIARelease(element)
        return 0
    }

    return element
}

; Scroll Up
^!1:: {
    el := getScrollElement()
    if !el {
        return
    }
    try {
        ScrollWithUIA(el, "up")
    } finally {
        UIARelease(el)
    }
}

; Scroll Down
^!2:: {
    el := getScrollElement()
    if !el {
        return
    }
    try {
        ScrollWithUIA(el, "down")
    } finally {
        UIARelease(el)
    }
}
