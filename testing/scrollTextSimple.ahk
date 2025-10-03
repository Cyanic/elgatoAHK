#Requires AutoHotkey v2.0
#SingleInstance Force
#Include UIA.ahk

global LogFile := A_ScriptDir . "\\scroll_debug_" . FormatTime(, "yyyyMMdd_HHmmss") . ".log"

; Function to get the scrollable element
getScrollElement() {
    hwnd := WinExist("ahk_exe Camera Hub.exe")
    if !hwnd {
        MsgBox("Camera Hub not found.")
        return 0
    }

    try {
        root := UIA.ElementFromHandle(hwnd)
    } catch Error as err {
        MsgBox("Unable to bind UI Automation to Camera Hub.`n" err.Message)
        return 0
    }

    scrollEl := 0
    try {
        scrollEl := root.FindElement({ AutomationId: "qt_scrollarea_viewport" })
    } catch {
        scrollEl := 0
    }

    if !scrollEl {
        MsgBox("Scrollable element not found.")
        return 0
    }

    return scrollEl
}

getScrollParentElement() {
    scrollEl := getScrollElement()
    if !scrollEl {
        return 0
    }

    parent := 0
    try {
        parent := scrollEl.Parent
    } catch {
        parent := 0
    }

    if !parent {
        MsgBox("Parent scrollable element not found.")
        return 0
    }

    return parent
}

SendMouseWheelUIA(el, direction := "down", steps := 1) {
    if !el {
        return false
    }

    direction := StrLower(direction)
    if direction != "up"
        direction := "down"

    steps := Max(1, Round(Abs(steps)))

    debugSteps := []
    debugSteps.Push("Direction=" direction)
    debugSteps.Push("Steps=" steps)

    pattern := ResolveScrollPattern(el, debugSteps)
    if !pattern {
        debugSteps.Push("ScrollPattern unresolved")
        LogDebug("SendMouseWheelUIA: " . Join(debugSteps, " | "))
        return SendMouseWheel(el, direction)
    }

    vertAmount := direction = "up" ? 1 : 4
    success := true
    try {
        Loop steps {
            pattern.Scroll(0, vertAmount)
            if steps > 1
                Sleep 60
        }
    } catch Error as err {
        debugSteps.Push("Scroll exception: " err.Message)
        success := false
    }

    if !success {
        debugSteps.Push("Falling back to PostMessage")
        LogDebug("SendMouseWheelUIA: " . Join(debugSteps, " | "))
        return SendMouseWheel(el, direction)
    }

    debugSteps.Push("UIA scroll succeeded")
    LogDebug("SendMouseWheelUIA: " . Join(debugSteps, " | "))
    return true
}

SendMouseWheel(el, direction := "down") {
    if !el {
        return false
    }

    debugSteps := []
    debugSteps.Push("Direction=" direction)
    hwnd := ResolveElementHandle(el)
    if !hwnd {
        ShowElementDebug(el, "WM_MOUSEWHEEL: Native window handle not found.")
        LogDebug("WM_MOUSEWHEEL: Native window handle not found")
        return false
    }

    x := 0
    y := 0
    hasLocation := false
    try {
        loc := el.Location
        if IsObject(loc) {
            x := Round(loc.x + loc.w / 2)
            y := Round(loc.y + loc.h / 2)
            hasLocation := true
            debugSteps.Push("Element location succeeded")
        } else {
            debugSteps.Push("Element location returned non-object")
        }
    } catch Error as err {
        debugSteps.Push("Element location exception: " err.Message)
    }

    delta := direction = "up" ? 120 : -120
    wParam := (delta & 0xFFFF) << 16
    lParam := ((y & 0xFFFF) << 16) | (x & 0xFFFF)

    LogDebug("SendMouseWheel: " . Join(debugSteps, " | "))

    sent := DllCall("User32.dll\PostMessageW", "ptr", hwnd, "uint", 0x020A, "ptr", wParam, "ptr", lParam)

    if !sent {
        ShowElementDebug(el, "WM_MOUSEWHEEL: PostMessage failed.")
        LogDebug("WM_MOUSEWHEEL: PostMessage failed")
        return false
    }
    return true
}

ResolveScrollPattern(el, debugSteps) {
    static UIA_ScrollPatternId := 10004
    current := el
    depth := 0
    while current {
        pattern := 0
        try {
            pattern := current.GetCurrentPattern(UIA_ScrollPatternId)
        } catch Error as err {
            debugSteps.Push("ScrollPattern error depth=" depth ": " err.Message)
            pattern := 0
        }
        if pattern {
            debugSteps.Push("ScrollPattern found depth=" depth)
            return pattern
        }
        try {
            current := current.Parent
        } catch {
            current := 0
        }
        depth += 1
    }
    debugSteps.Push("ScrollPattern not located in ancestry")
    return 0
}

ResolveElementHandle(el) {
    current := el
    Loop {
        if !current
            break
        hwnd := 0
        try {
            hwnd := current.NativeWindowHandle
        } catch {
            hwnd := 0
        }
        if hwnd {
            return hwnd
        }
        nextParent := 0
        try {
            nextParent := current.Parent
        } catch {
            nextParent := 0
        }
        current := nextParent
    }
    return 0
}

ShowElementDebug(el, message) {
    detailLines := []
    detailLines.Push(message)

    try {
        detailLines.Push("AutomationId: " (el.AutomationId != "" ? el.AutomationId : "<none>"))
    } catch {
        detailLines.Push("AutomationId: <error>")
    }

    try {
        detailLines.Push("ClassName: " (el.ClassName != "" ? el.ClassName : "<none>"))
    } catch {
        detailLines.Push("ClassName: <error>")
    }

    try {
        loc := el.Location
        detailLines.Push(Format("Location: x={1}, y={2}, w={3}, h={4}", loc.x, loc.y, loc.w, loc.h))
    } catch {
        detailLines.Push("Location: <error>")
    }

    try {
        valuePattern := el.ValuePattern
        detailLines.Push("ReadOnly: " (valuePattern.IsReadOnly ? "true" : "false"))
    } catch {
        detailLines.Push("ReadOnly: <unknown>")
    }

    patternNames := []
    try {
        for patternName, patternId in UIA.Pattern.OwnProps() {
            try {
                ; Accessing pattern property may throw; use HasPattern equivalent via property lookup
                el.GetPattern(patternId)
                patternNames.Push(patternName)
            } catch {
                continue
            }
        }
    } catch {
        patternNames := []
    }

    if patternNames.Length {
        detailLines.Push("Available Patterns: " . Join(patternNames, ", "))
    } else {
        detailLines.Push("Available Patterns: <none>")
    }

    text := Join(detailLines, "`n")
    try {
        A_Clipboard := text
    }
    catch {
        ; ignore clipboard errors
    }
    MsgBox(Join(detailLines, "`n"))
}

Join(items, delimiter := "") {
    if !IsObject(items)
        return ""
    output := ""
    for index, item in items {
        if index > 1
            output .= delimiter
        output .= item
    }
    return output
}

LogDebug(message) {
    global LogFile
    try {
        timestamp := FormatTime(, "yyyy-MM-dd HH:mm:ss")
        FileAppend(Format("[{1}] {2}`n", timestamp, message), LogFile, "UTF-8")
    } catch {
        ; ignore logging errors
    }
}

; Scroll Up
^!1:: {
    el := getScrollElement()
    if !el {
        return
    }
    KeyWait("Ctrl")
    KeyWait("Alt")
    SendMouseWheelUIA(el, "up")
}

; Scroll Down
^!2:: {
    el := getScrollElement()
    if !el {
        return
    }
    KeyWait("Ctrl")
    KeyWait("Alt")
    SendMouseWheelUIA(el, "down")
}

; Scroll Parent Up
^!3:: {
    el := getScrollParentElement()
    if !el {
        return
    }
    KeyWait("Ctrl")
    KeyWait("Alt")
    SendMouseWheel(el, "up")
}

; Scroll Parent Down
^!4:: {
    el := getScrollParentElement()
    if !el {
        return
    }
    KeyWait("Ctrl")
    KeyWait("Alt")
    SendMouseWheel(el, "down")
}
