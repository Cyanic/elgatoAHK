#Requires AutoHotkey v2.0
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

ScrollWithUIA(el, direction := "down") {
    if !el {
        return
    }

    try {
        if !el.IsScrollPatternAvailable {
            MsgBox("Scroll pattern not available.")
            return
        }
    } catch {
        ; fall through to attempting pattern retrieval
    }

    scrollPattern := 0
    try {
        scrollPattern := el.ScrollPattern
    } catch Error as err {
        MsgBox("Failed to retrieve scroll pattern.`n" err.Message)
        return
    }

    direction := StrLower(direction)
    vertAmount := direction = "up" ? UIA.ScrollAmount.SmallDecrement : UIA.ScrollAmount.SmallIncrement

    try {
        scrollPattern.Scroll(vertAmount)
    } catch Error as err {
        MsgBox("Failed to invoke scroll pattern.`n" err.Message)
    }
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
