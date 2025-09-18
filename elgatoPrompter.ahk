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
global _pending := Map()   ; controlName => delta
global _applyArmed := false
global _UIA_RangeValuePatternId := 10003
global _BrightnessSpinner :=
    "EVHMainWindow.centralWidget.stackedWidget.mainPage.propertySidebar.scrollArea.qt_scrollarea_viewport.sidebarContainer.propertyContainer.propertyGroup_com.elgato.vh-ui.group.picture-camlink4k.contentFrame.propertyGroup.property_kPCBrightness.spinBox_kPCBrightness"
global _ContrastSpinner :=
    "EVHMainWindow.centralWidget.stackedWidget.mainPage.propertySidebar.scrollArea.qt_scrollarea_viewport.sidebarContainer.propertyContainer.propertyGroup_com.elgato.vh-ui.group.picture-camlink4k.contentFrame.propertyGroup.property_kPCContrast.spinBox_kPCContrast"

; ---- Hotkeys ----
;F18:: QueuePulse(-1)
;F19:: QueuePulse(+1)
^!d:: QueuePulse("brightness", -1)
^!a:: QueuePulse("brightness", +1)
^!e:: QueuePulse("contrast", -1)
^!q:: QueuePulse("contrast", +1)
^!x:: QuitApp()
^!s:: SaveCalibration()
^!z:: DebugProbe()

; ===== Core accumulator =====
QueuePulse(controlName, sign) {
    global _pending, _applyArmed, BASE_STEP, APPLY_DELAY_MS
    if !_pending.Has(controlName)
        _pending[controlName] := 0
    _pending[controlName] += (BASE_STEP * sign)

    if !_applyArmed {
        _applyArmed := true
        SetTimer ApplyAccumulated, -APPLY_DELAY_MS
    }
}

ApplyAccumulated() {
    global _pending, _applyArmed, SHOW_PATH_TIP
    global _BrightnessSpinner, _ContrastSpinner, _UIA_RangeValuePatternId
    _applyArmed := false

    ; Nothing to do?
    if (_pending.Count = 0)
        return

    uiaElement := GetCamHubUiaElement()
    if !uiaElement {
        Tip("Camera Hub window not found.")
        _pending.Clear()
        return
    }

    ; Map control -> AutomationId resolver
    ; Add more controls here in the future if needed.
    ctrlToAutoId := Map(
        "brightness", _BrightnessSpinner,
        "contrast", _ContrastSpinner
    )

    summary := ""
    ; Flush all non-zero deltas and reset them to 0
    for controlName, delta in _pending.Clone() {
        if (delta = 0)
            continue

        if !ctrlToAutoId.Has(controlName) {
            _pending[controlName] := 0
            continue
        }

        autoId := ctrlToAutoId[controlName]
        try {
            elem := uiaElement.FindElement({ AutomationId: autoId })
            if elem {
                rvp := elem.GetCurrentPattern(_UIA_RangeValuePatternId)
                if rvp {
                    cur := rvp.value              ; already numeric from UIA
                    rvp.SetValue(cur + delta)     ; UIA will clamp to [min,max]
                    sign := (delta > 0 ? "+" : "")
                    summary .= (summary ? "`n" : "") controlName " " sign delta " -> " (cur + delta)
                }
            }
        }

        ; Zero out after applying (even if it failed—keeps the queue bounded)
        _pending[controlName] := 0
    }

    ; Optional: tooltip that shows what we just targeted

    if SHOW_PATH_TIP && summary {
        Tip("Applied:`n" summary)
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
    lines.Push("AutomationId (Brightness): " _BrightnessSpinner)
    lines.Push("AutomationId (Contrast): " _ContrastSpinner)

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
