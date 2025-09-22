; ========= Elgato Prompter Scroll Speed (AHK v2) =========

#Requires AutoHotkey v2
#SingleInstance Force

; ---- Path resolution helpers (must run before library includes) ----
global _ResolvedIniPath := ""

ResolvePath(path, reference := "") {
    if (path = "")
        return ""
    path := StrReplace(path, "/", "\")
    if RegExMatch(path, "^[A-Za-z]:\\") || SubStr(path, 1, 2) = "\\\\"
        return path

    baseDir := ""
    if (reference = "") {
        baseDir := A_ScriptDir
    } else if DirExist(reference) {
        baseDir := reference
    } else {
        SplitPath(reference,, &baseDir)
        if !baseDir
            baseDir := A_ScriptDir
    }
    baseDir := RTrim(StrReplace(baseDir, "/", "\"), "\\")
    if !baseDir
        baseDir := A_ScriptDir

    if SubStr(path, 1, 1) = "\\" {
        drive := SubStr(baseDir, 1, 2)
        return drive path
    }
    return baseDir "\" LTrim(path, "\\")
}

GetIniPath() {
    global _ResolvedIniPath
    if (_ResolvedIniPath != "")
        return _ResolvedIniPath

    defaultPath := A_ScriptDir "\prompter.ini"
    path := defaultPath
    try {
        override := IniRead(defaultPath, "Files", "IniPath", "")
        if override {
            candidate := ResolvePath(override, defaultPath)
            if FileExist(candidate)
                path := candidate
        }
    } catch {
        ; ignore and fall back to default
    }
    _ResolvedIniPath := path
    return path
}


#Include *i %A_ScriptDir%\UIA-v2-main\Lib\UIA.ahk
#Include *i %A_ScriptDir%\Lib\UIA.ahk
#Include *i %A_ScriptDir%\UIA.ahk

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
TOOLTIP_HIDE_DELAY_MS := 900

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


; ===== Core accumulator =====
QueuePulse(controlName, sign) {
    global _pending, _applyArmed, APPLY_DELAY_MS
    if !sign
        return

    if !_pending.Has(controlName)
        _pending[controlName] := 0
    _pending[controlName] += sign

    if (_pending[controlName] = 0)
        _pending.Delete(controlName)

    if !_applyArmed {
        _applyArmed := true
        SetTimer ApplyAccumulated, -APPLY_DELAY_MS
    }
}

ApplyAccumulated() {
    global _pending, _applyArmed, SHOW_PATH_TIP, DEBUG_VERBOSE_LOGGING
    _applyArmed := false
    if (_pending.Count = 0)
        return

    uiaElement := GetCamHubUiaElement()
    if !uiaElement {
        Tip("Camera Hub window not found.")
        _pending.Clear()
        return
    }

    ctrlSpecs := GetControlSpecs()
    if !ctrlSpecs || (ctrlSpecs.Count = 0) {
        if DEBUG_VERBOSE_LOGGING
            Log("ApplyAccumulated: control spec map is empty")
        _pending.Clear()
        return
    }

    names := []
    for name in _pending
        names.Push(name)

    summaryLines := []

    for name in names {
        if !ctrlSpecs.Has(name) {
            if DEBUG_VERBOSE_LOGGING
                Log("ApplyAccumulated: missing control spec for " name)
            _pending.Delete(name)
            continue
        }

        pulses := _pending.Has(name) ? _pending[name] : 0
        _pending.Delete(name)
        if !pulses
            continue

        spec := ctrlSpecs[name]
        effectivePulses := spec.Has("Invert") && spec["Invert"] ? -pulses : pulses

        result := Map("Success", false, "Detail", "")
        try result := spec["Handler"].Call(uiaElement, spec, effectivePulses)
        catch as err {
            if DEBUG_VERBOSE_LOGGING
                Log("ApplyAccumulated: handler error for " name " -> " err.Message)
        }

        success := result.Has("Success") ? result["Success"] : !!result
        if success {
            detail := result.Has("Detail") ? result["Detail"] : FormatSigned(effectivePulses)
            summaryLines.Push((spec.Has("DisplayName") ? spec["DisplayName"] : name) " " detail)
        } else {
            if DEBUG_VERBOSE_LOGGING
                Log("ApplyAccumulated: handler returned false for " name)
            summaryLines.Push((spec.Has("DisplayName") ? spec["DisplayName"] : name) " FAILED")
        }
    }

    if SHOW_PATH_TIP && summaryLines.Length
        Tip("Applied:`n" JoinLines(summaryLines))
}

GetControlSpecs() {
    global _ControlSpecs
    return _ControlSpecs
}

; Apply a numeric delta to a RangeValue spinner
ApplyRangeValueDelta(root, spec, pulses, uiRangeValueId := 10003) {
    global DEBUG_VERBOSE_LOGGING
    pulses := Round(pulses)
    if (pulses = 0)
        return HandlerResult(false)

    step := spec.Has("Step") ? spec["Step"] : 1
    delta := pulses * step
    el := ResolveControlElement(root, spec)
    if !el {
        if DEBUG_VERBOSE_LOGGING
            Log("ApplyRangeValueDelta: element not found for " spec["Name"])
        return HandlerResult(false)
    }

    try {
        rvp := el.GetCurrentPattern(uiRangeValueId) ; RangeValuePattern = 10003
        if !rvp
            return HandlerResult(false)
        cur := rvp.value
        rvp.SetValue(cur + delta)
        return HandlerResult(true, FormatSigned(delta))
    } catch as err {
        if DEBUG_VERBOSE_LOGGING
            Log("ApplyRangeValueDelta: pattern error " spec["Name"] " msg=" err.Message)
        InvalidateControlCache(spec)
        return HandlerResult(false)
    }
}

; Apply a delta to the Prompter viewport using ScrollPattern (percent fallback when needed)
ApplyScrollDelta(root, spec, pulses, uiScrollId := 10004) {
    global DEBUG_VERBOSE_LOGGING
    pulses := Round(pulses)
    if (pulses = 0)
        return HandlerResult(false)

    vp := ResolveControlElement(root, spec)
    if !vp {
        if DEBUG_VERBOSE_LOGGING
            Log("ApplyScrollDelta: viewport not found for " spec["Name"])
        return HandlerResult(false)
    }

    sp := 0
    try sp := vp.GetCurrentPattern(uiScrollId) ; ScrollPattern = 10004
    if sp {
        pulsesToSend := Abs(pulses)
        dir := (pulses > 0) ? 4 : 1 ; 4=LargeIncrement, 1=LargeDecrement
        if (pulsesToSend > 0) {
            try {
                Loop pulsesToSend
                    sp.Scroll(0, dir)
                return HandlerResult(true, FormatSigned(pulsesToSend, " step" (pulsesToSend = 1 ? "" : "s")))
            } catch as err {
                if DEBUG_VERBOSE_LOGGING
                    Log("ApplyScrollDelta: Scroll call failed dir=" dir " pulses=" pulsesToSend " msg=" err.Message)
                InvalidateControlCache(spec)
            }
        }

        percentPer := spec.Has("PercentPerStep") ? spec["PercentPerStep"] : 0
        if (percentPer != 0) {
            try {
                cur := sp.VerticalScrollPercent  ; -1 means unsupported
                if (cur >= 0) {
                    deltaPercent := pulses * percentPer
                    newp := ClampPercent(cur + deltaPercent)
                    sp.SetScrollPercent(sp.HorizontalScrollPercent, newp)
                    return HandlerResult(true, FormatSigned(deltaPercent, "%"))
                }
            } catch as err {
                if DEBUG_VERBOSE_LOGGING
                    Log("ApplyScrollDelta: SetScrollPercent failed msg=" err.Message)
                InvalidateControlCache(spec)
            }
        }
    } else if DEBUG_VERBOSE_LOGGING {
        Log("ApplyScrollDelta: ScrollPattern unavailable for " spec["Name"])
    }
    return HandlerResult(false)
}

; Try to resolve the Prompter's scrolling viewport element
FindPrompterViewport(root, spec) {
    autoId := spec.Has("AutoId") ? spec["AutoId"] : ""
    el := autoId ? FindByAutoId(root, autoId) : 0
    if el
        return el

    ; Look for the first QTextBrowser underneath any visible QScrollArea
    try {
        areas := root.FindElements({ ClassName: "QScrollArea" })
        for area in areas {
            try {
                qb := area.FindElement({ ClassName: "QTextBrowser" })
                if qb
                    return qb
            }
        }
    }
    catch as err {
        ; ignore
    }

    ; Fallback: first QTextBrowser anywhere
    try {
        qb := root.FindElement({ ClassName: "QTextBrowser" })
        if qb
            return qb
    }
    catch as err {
    }

    return 0
}

; Helper: find an element by AutomationId
FindByAutoId(root, autoId) {
    global DEBUG_VERBOSE_LOGGING
    if !root || !autoId
        return 0
    try {
        return root.FindElement({ AutomationId: autoId })
    }
    catch as err {
        if DEBUG_VERBOSE_LOGGING
            Log("FindByAutoId: lookup failed autoId=" autoId " msg=" err.Message)
        return 0
    }
}

QuitApp() {
    Tip("EXITING")
    Sleep 1000
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
    specs := GetControlSpecs()
    if specs.Has("brightness")
        lines.Push("AutomationId (Brightness): " (specs["brightness"].Has("AutoId") ? specs["brightness"]["AutoId"] : "(missing)"))
    if specs.Has("contrast")
        lines.Push("AutomationId (Contrast): " (specs["contrast"].Has("AutoId") ? specs["contrast"]["AutoId"] : "(missing)"))

    txt := JoinLines(lines)
    A_Clipboard := txt
    try FileAppend(txt "`r`n`r`n", DEBUG_LOG)
    MsgBox("Debug copied to clipboard.`nSaved to: " DEBUG_LOG, "Prompter Debug", "OK Iconi")
}

ToggleProbeScans() {
    global ENABLE_PROBE_SCANS, INI
    ENABLE_PROBE_SCANS := !ENABLE_PROBE_SCANS
    IniWrite(ENABLE_PROBE_SCANS ? 1 : 0, INI, "Debug", "ProbeScans")
    state := ENABLE_PROBE_SCANS ? "enabled" : "disabled"
    Tip("Probe scans " state)
    Log("ToggleProbeScans: " state)
}

ShowHelp() {
    global INI, DEBUG_LOG
    msg := "Elgato Prompter Hotkeys:`n"
    msg .= "F13/F14  -> Scroll viewport slower/faster`n"
    msg .= "F18/F19  -> Scroll speed spinner down/up`n"
    msg .= "Ctrl+Alt+D/A -> Brightness down/up`n"
    msg .= "Ctrl+Alt+E/Q -> Contrast down/up`n"
    msg .= "Ctrl+Alt+S -> Save calibration point`n"
    msg .= "Ctrl+Alt+Z -> Copy debug info`n"
    msg .= "Ctrl+Alt+U/M/W/P -> Diagnostic dumps`n"
    msg .= "Ctrl+Alt+F -> Quick scan counts`n"
    msg .= "Ctrl+Alt+C -> List candidate windows`n"
    msg .= "Ctrl+Alt+G -> Toggle probe scans`n"
    msg .= "Ctrl+Alt+H -> Show this help`n"
    msg .= "Ctrl+Alt+X -> Exit script`n`n"
    msg .= "Config: " INI "`n"
    msg .= "Log: " DEBUG_LOG
    MsgBox(msg, "Elgato Prompter Help", "OK Iconi")
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
    global _CachedCamHubHwnd
    hwnd := GetCamHubHwnd()
    if !hwnd {
        if _CachedCamHubHwnd {
            _CachedCamHubHwnd := 0
            ; nothing cached to clear
        }
        Log("GetCamHubUiaElement: Camera Hub window not found")
        return
    }

    if (_CachedCamHubHwnd != hwnd) {
        _CachedCamHubHwnd := hwnd
        ; nothing cached to clear
    }

    uiaElement := UIA.ElementFromHandle(hwnd)
    if !uiaElement {
        Log("GetCamHubUiaElement: UIA.ElementFromHandle returned NULL")
    }
    return uiaElement
}

GetCamHubHwnd() {
    global APP_EXE, WIN_CLASS_RX

    ; Prefer a class that matches the configured regex among windows for this exe
    if WIN_CLASS_RX {
        for candidate in WinGetList("ahk_exe " APP_EXE) {
            try {
                cls := WinGetClass("ahk_id " candidate)
                if RegExMatch(cls, WIN_CLASS_RX)
                    return candidate
            }
        }
    }

    ; Fallback to first window owned by the executable
    hwnd := WinExist("ahk_exe " APP_EXE)
    if hwnd
        return hwnd

    ; Final fallback: exact class match if provided
    return WIN_CLASS_RX ? WinExist("ahk_class " WIN_CLASS_RX) : 0
}

; ======== Diagnostics & Logging (AHK v2) ========

; Append a line to our debug log with timestamp
Log(line) {
    global DEBUG_LOG
    ts := FormatTime(, "yyyy-MM-dd HH:mm:ss")
    try FileAppend(ts "  " line "`r`n", DEBUG_LOG)
}

; Dump monitor layout and DPI (per-monitor where supported)
DumpMonitors() {
    Log("=== Monitors ===")
    count := MonitorGetCount()   ; or: SysGet(80)
    Log("Monitor count: " count)
    Loop count {
        i := A_Index
        MonitorGet(i, &l, &t, &r, &b)
        name := ""
        try name := MonitorGetName(i)  ; may fail on older Windows
        Log(Format("[{}] {}x{} @({},{})  Name: {}", i, r - l, b - t, l, t, name))
    }
    hwnd := WinActive("A")
    if hwnd {
        dpi := 0
        try dpi := DllCall("User32\GetDpiForWindow", "ptr", hwnd, "uint")
        if dpi
            Log("Active window DPI: " dpi)
    }
}

; Enumerate all windows that match app exe or Qt class and show which one we picked
DumpWindows() {
    global APP_EXE, WIN_CLASS_RX
    Log("=== Windows ===")
    picked := GetCamHubHwnd()
    for hwnd in WinGetList() {
        title := WinGetTitle("ahk_id " hwnd)
        cls := WinGetClass("ahk_id " hwnd)
        pid := WinGetPID("ahk_id " hwnd)
        exe := ""
        try exe := ProcessGetPath(pid)
        flag := (hwnd = picked) ? " <== PICKED" : ""
        Log(Format("hwnd={:#x}  cls={}  pid={}  exe={}  title={}{}", hwnd, cls, pid, exe, title, flag))
    }
}

; Try to infer elevation mismatch risk: if script is admin and target isn't (or vice versa)
CheckElevationRisk() {
    global APP_EXE
    hwnd := WinExist("ahk_exe " APP_EXE)
    if !hwnd {
        Log("ElevationCheck: target hwnd NOT FOUND")
        return
    }
    pid := WinGetPID("ahk_id " hwnd)
    exe := ""
    try exe := ProcessGetPath(pid)
    Log("ElevationCheck: AHK IsAdmin=" (A_IsAdmin ? "Yes" : "No") "  TargetExe=" exe)
    ; NOTE: Accurate token elevation check for target proc requires COM/WMI or adv API; this is a hint only.
}

; Dump the element under the mouse and its ancestors
DumpUnderMouse() {
    global MAX_ANCESTOR_DEPTH
    try {
        MouseGetPos &mx, &my, &hwin
        el := UIA.ElementFromPoint(mx, my)
        Log("=== UIA UnderMouse ===")
        Log(Format("Screen @({},{})  hwnd={:#x}", mx, my, hwin))
        if !el {
            Log("ElementFromPoint: NULL")
            return
        }
        DumpOne(el, "UNDER_MOUSE")
        Log("Ancestors:")
        p := el
        loop MAX_ANCESTOR_DEPTH {
            p := p.Parent
            if !p
                break
            DumpOne(p, "  ^ parent")
        }
    } catch as e {
        Log("DumpUnderMouse ERROR: " e.Message)
    }
}

; Dump basic fields for one UIA element
DumpOne(el, prefix := "") {
    if !el
        return
    name := ""
    tempclass := ""
    aid := ""
    ctype := ""
    rect := ""
    try name := el.Name
    try tempclass := el.ClassName
    try aid := el.AutomationId
    try ctype := el.ControlType  ; may be enum name or number depending on lib
    try {
        r := el.BoundingRectangle
        rect := Format("[{},{},{},{}]", r.x, r.y, r.w, r.h)
    }
    Log(prefix "  Name=" name "  Class=" tempclass "  AutoId=" aid "  CtrlType=" ctype "  Rect=" rect)
}

; Walk the subtree and list first N descendants
DumpSubtree(uiaRoot, limit := SUBTREE_LIST_LIMIT) {
    global SUBTREE_LIST_LIMIT
    if !uiaRoot {
        Log("DumpSubtree: root is NULL")
        return
    }
    Log("=== UIA Subtree (first " limit ") ===")
    cnt := 0
    for el in uiaRoot.FindElements({}) { ; empty condition == all descendants in this lib
        cnt += 1
        DumpOne(el, "#" cnt)
        if (cnt >= limit)
            break
    }
    Log("Subtree enumerated: " cnt " elements (displayed up to " limit ")")
}

; Count by a condition, list up to M
Scan(uiaRoot, condObj, label := "", limit := SCAN_LIST_LIMIT) {
    global SCAN_LIST_LIMIT
    if !uiaRoot {
        Log("Scan[" label "]: root is NULL")
        return 0
    }
    cnt := 0
    Log("=== Scan " label " ===")
    for el in uiaRoot.FindElements(condObj) {
        cnt += 1
        if (cnt <= limit)
            DumpOne(el, "#" cnt)
    }
    Log("Scan " label ": total=" cnt)
    return cnt
}

; Convenience: run a full diagnostic pass
FullDiag() {
    global SUBTREE_LIST_LIMIT, SCAN_LIST_LIMIT, SLIDER_SCAN_LIMIT
    Log("===== FULL DIAGNOSTIC START =====")
    DumpMonitors()
    DumpWindows()
    CheckElevationRisk()
    hwnd := GetCamHubHwnd()
    if !hwnd {
        Log("No target hwnd; aborting UIA dumps.")
        Log("===== FULL DIAGNOSTIC END =====")
        return
    }
    root := UIA.ElementFromHandle(hwnd)
    if !root {
        Log("UIA.ElementFromHandle returned NULL")
        Log("===== FULL DIAGNOSTIC END =====")
        return
    }
    DumpOne(root, "ROOT")
    DumpSubtree(root, SUBTREE_LIST_LIMIT)
    Scan(root, { ClassName: "EVHSpinBox" }, 'ClassName:"EVHSpinBox"', SCAN_LIST_LIMIT)
    ; Fallback: some builds expose spinners as ControlType: Spinner or Slider
    Scan(root, { ControlType: "Spinner" }, 'ControlType:"Spinner"', SCAN_LIST_LIMIT)
    Scan(root, { ControlType: "Slider" }, 'ControlType:"Slider"', SLIDER_SCAN_LIMIT)
    DumpUnderMouse()
    Log("===== FULL DIAGNOSTIC END =====")
}

; Helper: list of monitors with names (best-effort)
MonitorGetList() {
    arr := []
    Loop MonitorGetCount()
        arr.Push(A_Index)
    return arr
}
; ---- Control helpers ----
ResolveControlElement(root, spec) {
    el := 0
    if spec.Has("Resolver") {
        try el := spec["Resolver"].Call(root, spec)
        catch {
            el := 0
        }
    } else if spec.Has("AutoId") {
        el := FindByAutoId(root, spec["AutoId"])
    }
    return el
}

InvalidateControlCache(spec) {
    ; caching disabled (UIA elements proved unstable), nothing to invalidate
}

HandlerResult(success, detail := "") {
    return Map("Success", !!success, "Detail", detail)
}

FormatSigned(value, suffix := "") {
    if !IsNumber(value)
        return value suffix
    str := (Round(value) = value) ? Format("{:+d}", value) : Format("{:+.2f}", value)
    return str suffix
}

ClampPercent(value) {
    if (value < 0)
        return 0
    if (value > 100)
        return 100
    return value
}
; ---- Small utilities ----
JoinLines(arr) {
    s := ""
    for v in arr
        s .= v "`r`n"
    return RTrim(s, "`r`n")
}

IniReadBool(file, section, key, default := false) {
    defStr := default ? "1" : "0"
    val := IniRead(file, section, key, defStr)
    txt := StrLower(Trim(val))
    return (txt = "1" || txt = "true" || txt = "yes" || txt = "on")
}

IniReadNumber(file, section, key, default) {
    val := Trim(IniRead(file, section, key, default))
    return IsNumber(val) ? val + 0 : default
}

LoadConfigOverrides() {
    global INI, APP_EXE, WIN_CLASS_RX, DEBUG_LOG, BASE_STEP, APPLY_DELAY_MS, SHOW_PATH_TIP
    global MAX_ANCESTOR_DEPTH, SUBTREE_LIST_LIMIT, SCAN_LIST_LIMIT, SLIDER_SCAN_LIMIT, TOOLTIP_HIDE_DELAY_MS
    global DEBUG_VERBOSE_LOGGING, ENABLE_PROBE_SCANS
    global _ControlSpecs

    SplitPath(INI,, &iniDir)

    APP_EXE := IniRead(INI, "App", "Executable", APP_EXE)
    WIN_CLASS_RX := IniRead(INI, "App", "ClassRegex", WIN_CLASS_RX)

    debugLogOverride := IniRead(INI, "Files", "DebugLog", DEBUG_LOG)
    DEBUG_LOG := ResolvePath(debugLogOverride, iniDir ? iniDir : A_ScriptDir)

    APPLY_DELAY_MS := IniReadNumber(INI, "Behavior", "ApplyDelayMs", APPLY_DELAY_MS)
    SHOW_PATH_TIP := IniReadBool(INI, "Behavior", "ShowPathTip", SHOW_PATH_TIP)
    BASE_STEP := IniReadNumber(INI, "Behavior", "BaseStep", BASE_STEP)

    ENABLE_PROBE_SCANS := IniReadBool(INI, "Debug", "ProbeScans", ENABLE_PROBE_SCANS)
    DEBUG_VERBOSE_LOGGING := IniReadBool(INI, "Debug", "VerboseLogging", DEBUG_VERBOSE_LOGGING)

    MAX_ANCESTOR_DEPTH := IniReadNumber(INI, "Diagnostics", "MaxAncestorDepth", MAX_ANCESTOR_DEPTH)
    SUBTREE_LIST_LIMIT := IniReadNumber(INI, "Diagnostics", "SubtreeLimit", SUBTREE_LIST_LIMIT)
    SCAN_LIST_LIMIT := IniReadNumber(INI, "Diagnostics", "ScanLimit", SCAN_LIST_LIMIT)
    SLIDER_SCAN_LIMIT := IniReadNumber(INI, "Diagnostics", "SliderScanLimit", SLIDER_SCAN_LIMIT)
    TOOLTIP_HIDE_DELAY_MS := IniReadNumber(INI, "UI", "TooltipHideDelayMs", TOOLTIP_HIDE_DELAY_MS)

    percentPerStep := IniReadNumber(INI, "Scroll", "PercentPerStep", 5)

    controlMeta := [
        Map("Name", "brightness", "AutoKey", "Brightness", "Label", "Brightness", "Handler", ApplyRangeValueDelta),
        Map("Name", "contrast", "AutoKey", "Contrast", "Label", "Contrast", "Handler", ApplyRangeValueDelta),
        Map("Name", "scrollspeed", "AutoKey", "ScrollSpeed", "Label", "Scroll speed", "Handler", ApplyRangeValueDelta),
        Map("Name", "fontsize", "AutoKey", "FontSize", "Label", "Font size", "Handler", ApplyRangeValueDelta),
        Map("Name", "scroll", "AutoKey", "ScrollViewport", "Label", "Scroll", "Handler", ApplyScrollDelta, "Resolver", FindPrompterViewport, "PercentPerStep", percentPerStep)
    ]

    stepsSection := "ControlSteps"
    invertSection := "ControlInvert"

    specs := Map()
    for meta in controlMeta {
        name := meta["Name"]
        autoId := IniRead(INI, "Automation", meta["AutoKey"], "")
        step := IniReadNumber(INI, stepsSection, name, BASE_STEP)
        invert := IniReadBool(INI, invertSection, name, false)

        spec := Map(
            "Name", name,
            "DisplayName", meta["Label"],
            "Handler", meta["Handler"],
            "Step", step
        )

        if autoId
            spec["AutoId"] := autoId
        if meta.Has("Resolver")
            spec["Resolver"] := meta["Resolver"]
        if meta.Has("PercentPerStep")
            spec["PercentPerStep"] := meta["PercentPerStep"]
        if invert
            spec["Invert"] := true

        specs[name] := spec
    }

    _ControlSpecs := specs
}

Tip(t) {
    global TOOLTIP_HIDE_DELAY_MS
    ToolTip t
    SetTimer(() => ToolTip(), -TOOLTIP_HIDE_DELAY_MS)
}
