#Requires AutoHotkey v2.0
#SingleInstance Force

ScrollCameraHub(direction := "down", steps := 1) {
    static WM_MOUSEWHEEL := 0x020A
    static WHEEL_DELTA := 120
    static metrics := {offsetX: 408, offsetY: 216, width: 208, height: 168}

    hwnd := WinExist("Camera Hub ahk_class Qt673QWindowToolSaveBits ahk_exe Camera Hub.exe")
    if !hwnd {
        MsgBox "Camera Hub window not found."
        return false
    }

    steps := Max(1, Round(steps))
    delta := (StrLower(direction) = "up") ? WHEEL_DELTA : -WHEEL_DELTA

    WinGetPos(&winX, &winY,,, "ahk_id " hwnd)
    targetX := winX + metrics.offsetX + metrics.width / 2
    targetY := winY + metrics.offsetY + metrics.height / 2

    lParam := MakeLParam(targetX, targetY)
    wParam := delta << 16

    Loop steps {
        DllCall("SendMessageW", "ptr", hwnd, "uint", WM_MOUSEWHEEL, "ptr", wParam, "ptr", lParam)
        Sleep 80
    }
    return true
}

MakeLParam(x, y) {
    xWord := Round(x) & 0xFFFF
    yWord := Round(y) & 0xFFFF
    return (yWord << 16) | xWord
}

^!Down::ScrollCameraHub("down", 3)
^!Up::ScrollCameraHub("up", 3)
