; ========= Elgato Prompter Scroll Speed (AHK v2) =========

#Requires AutoHotkey v2
#SingleInstance Force
#include C:\Cyanic\Tools\UIA-v2-main\Lib\UIA.ahk

; ---- App targeting ----
APP_EXE := "Camera Hub.exe"
WIN_CLASS_RX := "Qt\d+QWindowIcon" ; Qt673QWindowIcon

; ---- Behavior tuning ----
BASE_STEP := 1              ; 1% per knob detent
APPLY_DELAY_MS := 40        ; coalesce fast pulses so no detents are dropped (40–90 typical)

; ---- Path control ----
SHOW_PATH_TIP := true

; ---- Debug toggles ----
ENABLE_PROBE_SCANS := false
DEBUG_VERBOSE_LOGGING := false

; ---- Diagnostics & UI ----
MAX_ANCESTOR_DEPTH := 10
SUBTREE_LIST_LIMIT := 50
SCAN_LIST_LIMIT := 25
SLIDER_SCAN_LIMIT := 10
TOOLTIP_HIDE_DELAY_MS := 900

; ---- Files ----
INI := A_ScriptDir "\prompter.ini"
DEBUG_LOG := A_ScriptDir "\PrompterDebug.txt"

; Initialize runtime-configurable toggles
ENABLE_PROBE_SCANS := IniReadBool(INI, "Debug", "ProbeScans", ENABLE_PROBE_SCANS)
DEBUG_VERBOSE_LOGGING := IniReadBool(INI, "Debug", "VerboseLogging", DEBUG_VERBOSE_LOGGING)

; ---- State & Globals----
global _pending := Map()   ; controlName => delta
global _applyArmed := false
global _UIA_RangeValuePatternId := 10003
global _BrightnessSpinner :=
    "EVHMainWindow.centralWidget.stackedWidget.mainPage.propertySidebar.scrollArea.qt_scrollarea_viewport.sidebarContainer.propertyContainer.propertyGroup_com.elgato.vh-ui.group.picture-camlink4k.contentFrame.propertyGroup.property_kPCBrightness.spinBox_kPCBrightness"
global _ContrastSpinner :=
    "EVHMainWindow.centralWidget.stackedWidget.mainPage.propertySidebar.scrollArea.qt_scrollarea_viewport.sidebarContainer.propertyContainer.propertyGroup_com.elgato.vh-ui.group.picture-camlink4k.contentFrame.propertyGroup.property_kPCContrast.spinBox_kPCContrast"
global _ScrollSpeedSpinner :=
    "EVHMainWindow.centralWidget.stackedWidget.mainPage.propertySidebar.scrollArea.qt_scrollarea_viewport.sidebarContainer.prompterContainer.propertyGroup_com.elgato.vh-ui.prompter.group.scrolling.contentFrame.propertyGroup.property_kPRPScrollSpeed.spinBox_kPRPScrollSpeed"
global _FontSizeSpinner :=
    "EVHMainWindow.centralWidget.stackedWidget.mainPage.propertySidebar.scrollArea.qt_scrollarea_viewport.sidebarContainer.prompterContainer.propertyGroup_com.elgato.vh-ui.prompter.group.appearance.contentFrame.propertyGroup.property_kPRPFontSize.spinBox_kPRPFontSize"

global _ScrollViewportAutoId := "qt_scrollarea_viewport"

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
    summary := ""

    ; Drain the coalesced deltas and apply each to its target control
    for controlName, delta in _pending {
        if !delta {
            _pending[controlName] := 0
            continue
        }

        spec := ctrlSpecs.Has(controlName) ? ctrlSpecs[controlName] : 0
        if !spec {
            if DEBUG_VERBOSE_LOGGING
                Log("ApplyAccumulated: missing control spec for " controlName)
            _pending[controlName] := 0
            continue
        }

        ok := false
        ; Attempt to invoke the mapped handler (range/scroll) with the pending delta
        try ok := spec["Handler"].Call(uiaElement, spec["AutoId"], delta)
        catch as err {
            if DEBUG_VERBOSE_LOGGING
                Log("ApplyAccumulated: handler error for " controlName " -> " err.Message)
        }

        if ok {
            sign := (delta > 0 ? "+" : "")
            summary .= (summary ? "`n" : "") controlName " " sign delta
        } else {
            if DEBUG_VERBOSE_LOGGING
                Log("ApplyAccumulated: handler returned false for " controlName)
            summary .= (summary ? "`n" : "") controlName " FAILED"
        }

        _pending[controlName] := 0
    }

    ; Show a tooltip reflecting the applied adjustments when enabled
    if SHOW_PATH_TIP && summary
        Tip("Applied:`n" summary)
}

GetControlSpecs() {
    global _BrightnessSpinner, _ContrastSpinner, _ScrollSpeedSpinner, _FontSizeSpinner
    global _ScrollViewportAutoId
    rangeHandler := ApplyRangeValueDelta
    scrollHandler := ApplyScrollDelta
    ; Each entry: name -> { AutoId, Handler }
    return Map(
        "brightness", Map("AutoId", _BrightnessSpinner, "Handler", rangeHandler),
        "contrast", Map("AutoId", _ContrastSpinner, "Handler", rangeHandler),
        "scrollspeed", Map("AutoId", _ScrollSpeedSpinner, "Handler", rangeHandler),
        "fontsize", Map("AutoId", _FontSizeSpinner, "Handler", rangeHandler),
        "scroll", Map("AutoId", _ScrollViewportAutoId, "Handler", scrollHandler)
    )
}

; Apply a numeric delta to a RangeValue spinner
ApplyRangeValueDelta(root, autoId, delta, uiRangeValueId := 10003) {
    global DEBUG_VERBOSE_LOGGING
    el := FindByAutoId(root, autoId)
    if !el {
        if DEBUG_VERBOSE_LOGGING
            Log("ApplyRangeValueDelta: element not found for autoId=" autoId)
        return false
    }
    try {
        rvp := el.GetCurrentPattern(uiRangeValueId) ; RangeValuePattern = 10003
        if !rvp
            return false
        cur := rvp.value              ; numeric value from UIA
        rvp.SetValue(cur + delta)     ; UIA clamps to [min,max]
        return true
    } catch as err {
        if DEBUG_VERBOSE_LOGGING
            Log("ApplyRangeValueDelta: pattern error autoId=" autoId " msg=" err.Message)
        return false
    }
}

; Apply a delta to the Prompter viewport using ScrollPattern (percent fallback when needed)
ApplyScrollDelta(root, autoId, delta, uiScrollId := 10004) {
    global BASE_STEP
    global DEBUG_VERBOSE_LOGGING
    vp := FindPrompterViewport(root, autoId)
    if !vp {
        if DEBUG_VERBOSE_LOGGING
            Log("ApplyScrollDelta: viewport not found autoId=" autoId)
        return false
    }

    if (delta = 0)
        return false

    ; Try ScrollPattern
    sp := 0
    try sp := vp.GetCurrentPattern(uiScrollId) ; ScrollPattern = 10004
    if sp {
        ; Relative scroll if supported
        pulses := Abs(Round(delta / (BASE_STEP ? BASE_STEP : 1)))
        if (pulses < 1)
            pulses := 1
        dir := (delta > 0) ? 4 : 1 ; 4=LargeIncrement, 1=LargeDecrement
        try {
            Loop pulses
                sp.Scroll(0, dir)
            return true
        } catch as err {
            if DEBUG_VERBOSE_LOGGING
                Log("ApplyScrollDelta: Scroll call failed dir=" dir " pulses=" pulses " msg=" err.Message)
        }
        ; Percent fallback
        try {
            cur := sp.VerticalScrollPercent  ; -1 means unsupported
            if (cur >= 0) {
                newp := cur + delta
                if (newp < 0)
                    newp := 0
                else if (newp > 100)
                    newp := 100
                sp.SetScrollPercent(sp.HorizontalScrollPercent, newp)
                return true
            }
        }
        catch as err {
            if DEBUG_VERBOSE_LOGGING
                Log("ApplyScrollDelta: SetScrollPercent failed msg=" err.Message)
            return false
        }
    }
    else if DEBUG_VERBOSE_LOGGING {
        Log("ApplyScrollDelta: ScrollPattern unavailable for autoId=" autoId)
    }
    return false
}

; Try to resolve the Prompter's scrolling viewport element
FindPrompterViewport(root, autoId) {
    ; 1) Direct AutomationId match
    el := FindByAutoId(root, autoId)
    if el
        return el

    ; 2) Look inside QScrollArea for first QTextBrowser child
    for area in root.FindElements({ ClassName: "QScrollArea" }) {
        try {
            qb := area.FindElement({ ClassName: "QTextBrowser" })
            if qb
                return qb
        }
    }

    ; 3) Fallback: first QTextBrowser anywhere
    try {
        qb := root.FindElement({ ClassName: "QTextBrowser" })
        if qb
            return qb
    }
    catch {
    }

    return 0  ; nothing found
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
    lines.Push("AutomationId (Brightness): " _BrightnessSpinner)
    lines.Push("AutomationId (Contrast): " _ContrastSpinner)

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
    hwnd := GetCamHubHwnd()
    if !hwnd {
        Log("GetCamHubUiaElement: Camera Hub window not found")
        return
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
    global _BrightnessSpinner, _ContrastSpinner, _ScrollSpeedSpinner, _FontSizeSpinner, _ScrollViewportAutoId

    APP_EXE := IniRead(INI, "App", "Executable", APP_EXE)
    WIN_CLASS_RX := IniRead(INI, "App", "ClassRegex", WIN_CLASS_RX)
    DEBUG_LOG := IniRead(INI, "Files", "DebugLog", DEBUG_LOG)

    BASE_STEP := IniReadNumber(INI, "Behavior", "BaseStep", BASE_STEP)
    APPLY_DELAY_MS := IniReadNumber(INI, "Behavior", "ApplyDelayMs", APPLY_DELAY_MS)
    SHOW_PATH_TIP := IniReadBool(INI, "Behavior", "ShowPathTip", SHOW_PATH_TIP)

    ENABLE_PROBE_SCANS := IniReadBool(INI, "Debug", "ProbeScans", ENABLE_PROBE_SCANS)
    DEBUG_VERBOSE_LOGGING := IniReadBool(INI, "Debug", "VerboseLogging", DEBUG_VERBOSE_LOGGING)

    MAX_ANCESTOR_DEPTH := IniReadNumber(INI, "Diagnostics", "MaxAncestorDepth", MAX_ANCESTOR_DEPTH)
    SUBTREE_LIST_LIMIT := IniReadNumber(INI, "Diagnostics", "SubtreeLimit", SUBTREE_LIST_LIMIT)
    SCAN_LIST_LIMIT := IniReadNumber(INI, "Diagnostics", "ScanLimit", SCAN_LIST_LIMIT)
    SLIDER_SCAN_LIMIT := IniReadNumber(INI, "Diagnostics", "SliderScanLimit", SLIDER_SCAN_LIMIT)
    TOOLTIP_HIDE_DELAY_MS := IniReadNumber(INI, "UI", "TooltipHideDelayMs", TOOLTIP_HIDE_DELAY_MS)

    _BrightnessSpinner := IniRead(INI, "Automation", "Brightness", _BrightnessSpinner)
    _ContrastSpinner := IniRead(INI, "Automation", "Contrast", _ContrastSpinner)
    _ScrollSpeedSpinner := IniRead(INI, "Automation", "ScrollSpeed", _ScrollSpeedSpinner)
    _FontSizeSpinner := IniRead(INI, "Automation", "FontSize", _FontSizeSpinner)
    _ScrollViewportAutoId := IniRead(INI, "Automation", "ScrollViewport", _ScrollViewportAutoId)
}

Tip(t) {
    global TOOLTIP_HIDE_DELAY_MS
    ToolTip t
    SetTimer(() => ToolTip(), -TOOLTIP_HIDE_DELAY_MS)
}