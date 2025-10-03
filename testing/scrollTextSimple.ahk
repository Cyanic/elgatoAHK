#Requires AutoHotkey v2.0
#SingleInstance Force
#Include UIA.ahk

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

ScrollWithUIA(el, direction := "down") {
    if !el {
        return false
    }

    try {
        if !el.IsScrollPatternAvailable {
            ShowElementDebug(el, "Scroll pattern not available.")
            return false
        }
    } catch {
        ; fall through to attempting pattern retrieval
    }

    scrollPattern := 0
    try {
        scrollPattern := el.ScrollPattern
    } catch Error as err {
        ShowElementDebug(el, "Failed to retrieve scroll pattern.`n" err.Message)
        return false
    }

    direction := StrLower(direction)
    vertAmount := direction = "up" ? UIA.ScrollAmount.SmallDecrement : UIA.ScrollAmount.SmallIncrement

    try {
        scrollPattern.Scroll(vertAmount)
    } catch Error as err {
        ShowElementDebug(el, "Failed to invoke scroll pattern.`n" err.Message)
        return false
    }

    return true
}

SendMouseWheel(el, direction := "down") {
    if !el {
        return false
    }

    hwnd := ResolveElementHandle(el)
    if !hwnd {
        ShowElementDebug(el, "WM_MOUSEWHEEL: Native window handle not found.")
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
        }
    } catch {
        ; ignore location errors
    }

    if !hasLocation {
        rect := Buffer(16, 0)
        if DllCall("user32\\GetWindowRect", "ptr", hwnd, "ptr", rect.Ptr) {
            left := NumGet(rect, 0, "int")
            top := NumGet(rect, 4, "int")
            right := NumGet(rect, 8, "int")
            bottom := NumGet(rect, 12, "int")
            x := (left + right) // 2
            y := (top + bottom) // 2
            hasLocation := true
        }
    }

    if !hasLocation {
        pt := Buffer(8, 0)
        if DllCall("user32\\GetCursorPos", "ptr", pt.Ptr) {
            x := NumGet(pt, 0, "int")
            y := NumGet(pt, 4, "int")
            hasLocation := true
        }
    }

    delta := direction = "up" ? 120 : -120
    wParam := (delta & 0xFFFF) << 16
    lParam := ((y & 0xFFFF) << 16) | (x & 0xFFFF)

    sent := DllCall("User32.dll\PostMessageW", "ptr", hwnd, "uint", 0x020A, "ptr", wParam, "ptr", lParam)
    if !sent {
        ShowElementDebug(el, "WM_MOUSEWHEEL: PostMessage failed.")
        return false
    }
    return true
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

; Scroll Up
^!1:: {
    el := getScrollElement()
    if !el {
        return
    }
    if !ScrollWithUIA(el, "up")
        SendMouseWheel(el, "up")
}

; Scroll Down
^!2:: {
    el := getScrollElement()
    if !el {
        return
    }
    if !ScrollWithUIA(el, "down")
        SendMouseWheel(el, "down")
}

; Scroll Parent Up
^!3:: {
    el := getScrollParentElement()
    if !el {
        return
    }
    if !ScrollWithUIA(el, "up")
        SendMouseWheel(el, "up")
}

; Scroll Parent Down
^!4:: {
    el := getScrollParentElement()
    if !el {
        return
    }
    if !ScrollWithUIA(el, "down")
        SendMouseWheel(el, "down")
}