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

    xScreen := x
    yScreen := y

    client := Buffer(8, 0)
    NumPut("int", xScreen, client, 0)
    NumPut("int", yScreen, client, 4)
    converted := DllCall("User32.dll\ScreenToClient", "ptr", hwnd, "ptr", client.Ptr)
    xClient := NumGet(client, 0, "int")
    yClient := NumGet(client, 4, "int")

    delta := direction = "up" ? 120 : -120
    wParam := (delta & 0xFFFF) << 16
    lParamClient := ((yClient & 0xFFFF) << 16) | (xClient & 0xFFFF)
    lParamScreen := ((yScreen & 0xFFFF) << 16) | (xScreen & 0xFFFF)

    postClient := DllCall("User32.dll\PostMessageW", "ptr", hwnd, "uint", 0x020A, "ptr", wParam, "ptr", lParamClient)
    rootHwnd := DllCall("User32.dll\GetAncestor", "ptr", hwnd, "uint", 2, "ptr")
    postRoot := 0
    if rootHwnd && rootHwnd != hwnd
        postRoot := DllCall("User32.dll\PostMessageW", "ptr", rootHwnd, "uint", 0x020A, "ptr", wParam, "ptr", lParamScreen)

    DebugMouseWheel(el, hwnd, rootHwnd, xScreen, yScreen, xClient, yClient, converted, postClient, postRoot)

    if postClient || postRoot
        return true

    ShowElementDebug(el, "WM_MOUSEWHEEL: PostMessage failed.")
    return false
}

DebugMouseWheel(el, hwnd, rootHwnd, xScreen, yScreen, xClient, yClient, converted, postClient, postRoot) {
    details := []
    details.Push("WM_MOUSEWHEEL diagnostics:")
    details.Push(Format("hwnd: 0x{:X}", hwnd))
    details.Push(Format("root hwnd: 0x{:X}", rootHwnd))
    details.Push(Format("Screen point: ({1}, {2})", xScreen, yScreen))
    details.Push(Format("Client point: ({1}, {2})", xClient, yClient))
    details.Push("ScreenToClient converted: " (converted ? "true" : "false"))
    details.Push("PostMessage (client hwnd): " (postClient ? "true" : "false"))
    if rootHwnd && rootHwnd != hwnd
        details.Push("PostMessage (root hwnd): " (postRoot ? "true" : "false"))
    try {
        native := el.NativeWindowHandle
        details.Push(Format("Element NativeWindowHandle: 0x{:X}", native))
    } catch {
        details.Push("Element NativeWindowHandle: <error>")
    }
    text := Join(details, "`n")
    try A_Clipboard := text
    catch
        ; ignore clipboard errors
    MsgBox(text)
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
    SendMouseWheel(el, "up")
}

; Scroll Down
^!2:: {
    el := getScrollElement()
    if !el {
        return
    }
    SendMouseWheel(el, "down")
}

; Scroll Parent Up
^!3:: {
    el := getScrollParentElement()
    if !el {
        return
    }
    SendMouseWheel(el, "up")
}

; Scroll Parent Down
^!4:: {
    el := getScrollParentElement()
    if !el {
        return
    }
    SendMouseWheel(el, "down")
}
