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
        scrollEl := root.FindElement({AutomationId: "qt_scrollarea_viewport"})
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

getScrollChildElement() {
    scrollEl := getScrollElement()
    if !scrollEl {
        return 0
    }

    children := []
    try {
        children := scrollEl.GetChildren()
    } catch {
        children := []
    }

    if children.Length = 0 {
        MsgBox("Child element not found.")
        return 0
    }

    return children[1]
}

ScrollWithUIA(el, direction := "down") {
    if !el {
        return
    }

    try {
        if !el.IsScrollPatternAvailable {
            ShowElementDebug(el, "Scroll pattern not available.")
            return
        }
    } catch {
        ; fall through to attempting pattern retrieval
    }

    scrollPattern := 0
    try {
        scrollPattern := el.ScrollPattern
    } catch Error as err {
        ShowElementDebug(el, "Failed to retrieve scroll pattern.`n" err.Message)
        return
    }

    direction := StrLower(direction)
    vertAmount := direction = "up" ? UIA.ScrollAmount.SmallDecrement : UIA.ScrollAmount.SmallIncrement

    try {
        scrollPattern.Scroll(vertAmount)
    } catch Error as err {
        ShowElementDebug(el, "Failed to invoke scroll pattern.`n" err.Message)
    }
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
    ScrollWithUIA(el, "up")
}

; Scroll Down
^!2:: {
    el := getScrollElement()
    if !el {
        return
    }
    ScrollWithUIA(el, "down")
}

; Scroll Child Up
^!5:: {
    el := getScrollChildElement()
    if !el {
        return
    }
    ScrollWithUIA(el, "up")
}

; Scroll Child Down
^!6:: {
    el := getScrollChildElement()
    if !el {
        return
    }
    ScrollWithUIA(el, "down")
}

; Scroll Parent Up
^!3:: {
    el := getScrollParentElement()
    if !el {
        return
    }
    ScrollWithUIA(el, "up")
}

; Scroll Parent Down
^!4:: {
    el := getScrollParentElement()
    if !el {
        return
    }
    ScrollWithUIA(el, "down")
}
