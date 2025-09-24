; ========= Elgato Prompter (AHK v2) =========

#Requires AutoHotkey v2
#SingleInstance Force

#Include %A_ScriptDir%\lib\util.ahk
#Include %A_ScriptDir%\lib\ui_helpers.ahk
#Include %A_ScriptDir%\lib\handlers.ahk
#Include %A_ScriptDir%\lib\config.ahk
#Include *i %A_ScriptDir%\uiashim.ahk
#Include *i %A_MyDocuments%\AutoHotkey\Lib\UIA.ahk
#Include *i %A_AppDataCommon%\AutoHotkey\Lib\UIA.ahk
#Include *i %A_ProgramFiles%\AutoHotkey\Lib\UIA.ahk
#Include *i %A_ScriptDir%\UIA-v2-main\Lib\UIA.ahk
#Include *i %A_ScriptDir%\Lib\UIA.ahk
#Include *i %A_ScriptDir%\UIA.ahk
#Include %A_ScriptDir%\lib\diagnostics.ahk

if EnsureUiaInclude()
    RestartForShim()

if !IsSet(UIA) {
    throw Error("UIA library not found. Place UIA.ahk in a standard library location or alongside the script.")
}

; ---- Configuration defaults (overridden later) ----
APP_EXE := "Camera Hub.exe"
WIN_CLASS_RX := "Qt\d+QWindowIcon" ; Qt673QWindowIcon
APPLY_DELAY_MS := 40        ; coalesce fast pulses so no detents are dropped (40–90 typical)
BASE_STEP := 1              ; default value delta per detent (overridden per control later)
SHOW_PATH_TIP := true
ENABLE_PROBE_SCANS := false
DEBUG_VERBOSE_LOGGING := false
MAX_ANCESTOR_DEPTH := 10
SUBTREE_LIST_LIMIT := 50
SCAN_LIST_LIMIT := 25
SLIDER_SCAN_LIMIT := 10
TOOLTIP_HIDE_DELAY_MS := 1000

; ---- Files ----
INI := GetIniPath()
DEBUG_LOG := A_ScriptDir "\PrompterDebug.txt"

; ---- State & Globals ----
global _pending := Map()   ; controlName => pulse count
global _applyArmed := false
global _UIA_RangeValuePatternId := 10003
global _ControlSpecs := Map()
global _CachedCamHubHwnd := 0

LoadConfigOverrides()

; ---- Hotkeys ----
F13:: QueuePulse("scroll", -1)   ; up / slower
F14:: QueuePulse("scroll", +1)   ; down / faster
F18:: QueuePulse("scrollspeed", -1)
F19:: QueuePulse("scrollspeed", +1)
^!d:: QueuePulse("brightness", -1)
^!a:: QueuePulse("brightness", +1)
^!e:: QueuePulse("contrast", -1)
^!q:: QueuePulse("contrast", +1)
^!x:: QuitApp()
^!s:: SaveCalibration()
^!z:: DebugProbe()


^!u:: (FullDiag(), Tip("Wrote full diagnostics to log"))
^!m:: (DumpMonitors(), Tip("Monitors dumped"))
^!w:: (DumpWindows(), Tip("Windows dumped"))
^!p:: (DumpUnderMouse(), Tip("Under-mouse UIA dumped"))
^!f:: {  ; Quick scan
    u := GetCamHubUiaElement()
    cnt1 := u ? Scan(u, { ClassName: "EVHSpinBox" }, "ClassName:EVHSpinBox", 50) : 0
    cnt2 := u ? Scan(u, { ControlType: "Spinner" }, "ControlType:Spinner", 50) : 0
    cnt3 := u ? Scan(u, { ControlType: "Slider" }, "ControlType:Slider", 50) : 0
    Tip("EVHSpinBox=" cnt1 "  Spinner=" cnt2 "  Slider=" cnt3)
}
^!c:: {  ; “Check candidates”
    global APP_EXE, WIN_CLASS_RX
    picked := GetCamHubHwnd()
    Tip("Picked hwnd=" Format("{:#x}", picked))
    for hwnd in WinGetList() {
        try {
            pid := WinGetPID("ahk_id " hwnd)
            exe := ProcessGetPath(pid)
            if InStr(exe, APP_EXE) {
                cls := WinGetClass("ahk_id " hwnd)
                title := WinGetTitle("ahk_id " hwnd)
                Log(Format("CamHub candidate hwnd={:#x}  cls={}  title={}", hwnd, cls, title))
            }
        }
    }
}

^!g:: ToggleProbeScans()   ; Toggle diagnostic UIA probe scans
^!h:: ShowHelp()