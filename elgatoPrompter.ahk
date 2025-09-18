; ========= Elgato Prompter Scroll Speed (AHK v2) =========

#Requires AutoHotkey v2
#SingleInstance Force
#include C:\Cyanic\Tools\UIA-v2-main\Lib\UIA.ahk

; ---- App targeting ----
APP_EXE := "Camera Hub.exe"
WIN_CLASS_RX := "Qt\d+QWindowIcon"

; ---- Behavior tuning ----
BASE_STEP := 1              ; 1% per knob detent
APPLY_DELAY_MS := 40        ; coalesce fast pulses so no detents are dropped (40–90 typical)

; ---- Path control ----
SHOW_PATH_TIP := true

; ---- Files ----
INI := A_ScriptDir "\PrompterSpeed.ini"
DEBUG_LOG := A_ScriptDir "\PrompterDebug.txt"

; ---- State & Globals----
global _pendingSum := 0
global _applyArmed := false
global _UIA_RangeValuePatternId := 10003
global _BrightnessSpinner :=
    "EVHMainWindow.centralWidget.stackedWidget.mainPage.propertySidebar.scrollArea.qt_scrollarea_viewport.sidebarContainer.propertyContainer.propertyGroup_com.elgato.vh-ui.group.picture-camlink4k.contentFrame.propertyGroup.property_kPCBrightness.spinBox_kPCBrightness"

; ---- Hotkeys ----
;F18:: QueuePulse(-1)
;F19:: QueuePulse(+1)
^!d:: QueuePulse(-1)
^!a:: QueuePulse(+1)
^!x:: QuitApp()
^!s:: SaveCalibration()
^!z:: DebugProbe()

; ===== Core accumulator =====
QueuePulse(sign) {
    global _pendingSum, _applyArmed, BASE_STEP, APPLY_DELAY_MS
    _pendingSum += (BASE_STEP * sign)
    if !_applyArmed {
        _applyArmed := true
        SetTimer ApplyAccumulated, -APPLY_DELAY_MS
    }
}

ApplyAccumulated() {
    global _pendingSum, _applyArmed, SHOW_PATH_TIP
    delta := _pendingSum, _pendingSum := 0, _applyArmed := false
    if (delta = 0)
        return

    uiaElement := GetCamHubUiaElement()
    if !uiaElement {
        Tip("Camera Hub window not found.")
        return
    }

    prompterElement := uiaElement.FindElement({ AutomationId: _BrightnessSpinner })
    rvp := prompterElement.GetCurrentPattern(_UIA_RangeValuePatternId) ; RangeValue (10003): exposes numeric Value, with Minimum, Maximum, SmallChange, LargeChange, and SetValue(number).
    if rvp {
        cur := rvp.value            ; numeric
        rvp.SetValue(cur + delta)   ; write a number (UIA clamps if needed)
    }

    if SHOW_PATH_TIP {
        Tip("Δ " delta)
    }
}

QuitApp() {
    MsgBox "exit"
    ExitApp()
}

; ===== Debug & helpers =====
DebugProbe() {
    global DEBUG_LOG
    hwnd := GetCamHubHwnd()
    lines := []
    lines.Push("=== Prompter Debug ===")
    lines.Push("Time: " A_YYYY "-" A_MM "-" A_DD " " A_Hour ":" A_Min ":" A_Sec)
    lines.Push("AHK: v" A_AhkVersion "  (x" (A_PtrSize * 8) ")  Elevated: " (A_IsAdmin ? "Yes" : "No"))
    lines.Push("App hwnd: " (hwnd ? hwnd : "NOT FOUND"))
    lines.Push("Saved point: " GetSavedPointText(hwnd))
    lines.Push("AutomationId: " _BrightnessSpinner)

    txt := JoinLines(lines)
    A_Clipboard := txt
    try FileAppend(txt "`r`n`r`n", DEBUG_LOG)
    MsgBox("Debug copied to clipboard.`nSaved to: " DEBUG_LOG, "Prompter Debug", "OK Iconi")
}

GetSavedPointText(hwnd) {
    global INI
    dx := IniRead(INI, "Spinner", "DX", "")
    dy := IniRead(INI, "Spinner", "DY", "")
    if (dx = "" || dy = "")
        return "MISSING (press Ctrl+Alt+S over the spinner)"
    if hwnd {
        WinGetPos &wx, &wy, &ww, &wh, "ahk_id " hwnd
        return "DX=" dx " DY=" dy "  (screen " (wx + dx) "," (wy + dy) ")"
    } else {
        return "DX=" dx " DY=" dy
    }
}

SaveCalibration() {
    ; TODO save all the AutomationIds in use
    global INI
    hwnd := GetCamHubHwnd()
    if !hwnd {
        Tip("Camera Hub not found.")
        return
    }
    MouseGetPos &mx, &my
    WinGetPos &wx, &wy, &ww, &wh, "ahk_id " hwnd
    dx := mx - wx, dy := my - wy
    IniWrite Round(dx), INI, "Spinner", "DX"
    IniWrite Round(dy), INI, "Spinner", "DY"
    Tip("Saved spinner @ (" Round(dx) "," Round(dy) ")")
}

GetCamHubUiaElement() {
    global APP_EXE
    uiaElement := UIA.ElementFromHandle("ahk_exe " APP_EXE)
    if !uiaElement {
        MsgBox "UIA Failed"
    }
    return uiaElement
}

GetCamHubHwnd() {
    global APP_EXE, WIN_CLASS_RX
    hwnd := WinExist("ahk_exe " APP_EXE)
    if !hwnd
        hwnd := WinExist("ahk_class " WIN_CLASS_RX)
    return hwnd
}

; ---- Small utilities ----
JoinLines(arr) {
    s := ""
    for v in arr
        s .= v "`r`n"
    return RTrim(s, "`r`n")
}

Tip(t) {
    ToolTip t
    SetTimer(() => ToolTip(), -900)
}
