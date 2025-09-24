#Requires AutoHotkey v2.0
#SingleInstance Force

ScrollCameraHub(direction := "down", steps := 1) {
    static WM_MOUSEWHEEL := 0x020A
    static WHEEL_DELTA := 120
    static metrics := {offsetX: 408, offsetY: 216, width: 208, height: 168}

    hwndMain := WinExist("Camera Hub ahk_class Qt673QWindowToolSaveBits ahk_exe Camera Hub.exe")
    if !hwndMain {
        MsgBox "Camera Hub window not found."
        return false
    }

    hwndScroll := GetChildByClass(hwndMain, "QScrollArea")
    if !hwndScroll {
        MsgBox "Scroll area not found."
        return false
    }

    steps := Max(1, Round(steps))
    delta := (StrLower(direction) = "up") ? WHEEL_DELTA : -WHEEL_DELTA

    WinGetPos(&winX, &winY,,, "ahk_id " hwndMain)
    targetX := winX + metrics.offsetX + metrics.width // 2
    targetY := winY + metrics.offsetY + metrics.height // 2

    Loop steps {
        wParam := (delta & 0xFFFF) << 16
        lParam := ((targetY & 0xFFFF) << 16) | (targetX & 0xFFFF)
        DllCall("PostMessageW", "ptr", hwndScroll, "uint", WM_MOUSEWHEEL, "ptr", wParam, "ptr", lParam)
        Sleep 80
    }
    return true
}

GetChildByClass(parentHwnd, className) {
    found := 0
    buf := Buffer(A_PtrSize)
    callback := CallbackCreate((childHwnd, lParam) => {
        if WinGetClass("ahk_id " childHwnd) = className {
            NumPut("ptr", childHwnd, lParam)
            return false
        }
        return true
    }, "Fast")

    DllCall("EnumChildWindows", "ptr", parentHwnd, "ptr", callback, "ptr", buf)
    CallbackFree(callback)
    return NumGet(buf, "ptr")
}

^!Down::ScrollCameraHub("down", 3)
^!Up::ScrollCameraHub("up", 3)
