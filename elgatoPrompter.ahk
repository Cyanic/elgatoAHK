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
global _UIA_ScrollId := 10004
global _ControlSpecs := Map()
global _CachedCamHubHwnd := 0
global _BoundControlHotkeys := []

LoadConfigOverrides()
InitializeControlHotkeys()

; ---- Hotkeys ----
; Control hotkeys are configured at runtime from prompter.ini.
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

; Initializes or re-initializes configurable control hotkeys.
InitializeControlHotkeys() {
    global _BoundControlHotkeys
    UnbindControlHotkeys()

    cfg := GetHotkeyConfig()
    defs := [
        Map("ConfigKey", "ScrollUp", "Label", "Scroll slower", "Callback", (*) => QueuePulse("scroll", -1)),
        Map("ConfigKey", "ScrollDown", "Label", "Scroll faster", "Callback", (*) => QueuePulse("scroll", +1)),
        Map("ConfigKey", "ScrollSpeedDown", "Label", "Scroll speed down", "Callback", (*) => QueuePulse("scrollspeed", -1)),
        Map("ConfigKey", "ScrollSpeedUp", "Label", "Scroll speed up", "Callback", (*) => QueuePulse("scrollspeed", +1)),
        Map("ConfigKey", "BrightnessDown", "Label", "Brightness down", "Callback", (*) => QueuePulse("brightness", -1)),
        Map("ConfigKey", "BrightnessUp", "Label", "Brightness up", "Callback", (*) => QueuePulse("brightness", +1)),
        Map("ConfigKey", "ContrastDown", "Label", "Contrast down", "Callback", (*) => QueuePulse("contrast", -1)),
        Map("ConfigKey", "ContrastUp", "Label", "Contrast up", "Callback", (*) => QueuePulse("contrast", +1))
    ]

    for def in defs {
        key := cfg.Has(def["ConfigKey"]) ? cfg[def["ConfigKey"]] : ""
        BindConfiguredHotkey(key, def["Label"], def["Callback"])
    }
}

; Binds a single configured hotkey and tracks it for future unbinding.
BindConfiguredHotkey(key, label, callback) {
    global _BoundControlHotkeys, DEBUG_VERBOSE_LOGGING
    key := Trim(key)
    if (key = "") {
        if DEBUG_VERBOSE_LOGGING
            Log("Hotkey skipped for " label ": no key configured")
        return
    }

    try {
        Hotkey(key, callback, "On")
        _BoundControlHotkeys.Push(key)
    } catch as err {
        Log("Failed to bind hotkey " label " (" key "): " err.Message)
        Tip("Hotkey error: " label)
    }
}

; Turns off all previously bound configurable hotkeys.
UnbindControlHotkeys() {
    global _BoundControlHotkeys
    if !_BoundControlHotkeys.Length
        return
    for key in _BoundControlHotkeys {
        try Hotkey(key, "Off")
    }
    _BoundControlHotkeys := []
}

; Reloads the INI hotkey config and rebinds control hotkeys.
ReloadControlHotkeys() {
    LoadHotkeyConfig()
    InitializeControlHotkeys()
}
