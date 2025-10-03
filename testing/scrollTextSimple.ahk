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

; Scroll Up
^!1:: {
    el := getScrollElement()
	MsgBox(el.Dump()) ; Shows available patterns, properties, etc.
    if el && el.HasPattern("Scroll") {
        scroll := el.GetCurrentPattern("Scroll")
        scroll.Scroll(0, -1) ; Scroll up
    } else {
        MsgBox("Scroll pattern not available.")
    }
}

; Scroll Down
^!2:: {
    el := getScrollElement()
	MsgBox(el.Dump()) ; Shows available patterns, properties, etc.
    if el && el.HasPattern("Scroll") {
        scroll := el.GetCurrentPattern("Scroll")
        scroll.Scroll(0, 1) ; Scroll down
    } else {
        MsgBox("Scroll pattern not available.")
    }
}