#Requires AutoHotkey v2.0
#Include UIA.ahk

; Function to get the scrollable element
getScrollElement() {
    hwnd := WinExist("ahk_exe Camera Hub.exe")
    if !hwnd {
        MsgBox("Camera Hub not found.")
        return
    }

    root := uia.ElementFromHandle(hwnd)

    ; Try to find element by AutomationId
    scrollEl := root.FindFirst("AutomationId=qt_scrollarea_viewport")
    
    if !scrollEl {
        MsgBox("Scrollable element not found.")
        return
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
        ReleaseComPtr(pattern)
    }
}

ReleaseComPtr(ptr) {
    if !ptr {
        return
    }
    vtable := NumGet(ptr, 0, "ptr")
    release := NumGet(vtable, 2 * A_PtrSize, "ptr")
    DllCall(release, "ptr", ptr, "uint")
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
