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
global _ScrollSpeedSpinner :=
    "EVHMainWindow.centralWidget.stackedWidget.mainPage.propertySidebar.scrollArea.qt_scrollarea_viewport.sidebarContainer.prompterContainer.propertyGroup_com.elgato.vh-ui.prompter.group.scrolling.contentFrame.propertyGroup.property_kPRPScrollSpeed.spinBox_kPRPScrollSpeed"
global _FontSizeSpinner :=
    "EVHMainWindow.centralWidget.stackedWidget.mainPage.propertySidebar.scrollArea.qt_scrollarea_viewport.sidebarContainer.prompterContainer.propertyGroup_com.elgato.vh-ui.prompter.group.appearance.contentFrame.propertyGroup.property_kPRPFontSize.spinBox_kPRPFontSize"

global _ScrollViewportAutoId := "qt_scrollarea_viewport"

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

    for controlName, delta in _pending.Clone() {
        if (delta = 0) {
            _pending[controlName] := 0
            continue
        }
        if !ctrlSpecs.Has(controlName) {
            _pending[controlName] := 0
            continue
        }

        spec := ctrlSpecs[controlName]
        autoId := spec["AutoId"]
        handler := spec["Handler"]

        ok := false
        try ok := handler.Call(uiaElement, autoId, delta)
        catch {
        }

        if ok {
            sign := (delta > 0 ? "+" : "")
            summary .= (summary ? "`n" : "") controlName " " sign delta
        } else {
            summary .= (summary ? "`n" : "") controlName " FAILED"
        }

        _pending[controlName] := 0
    }

    if SHOW_PATH_TIP && summary
        Tip("Applied:`n" summary)
}

GetControlSpecs() {
    global _BrightnessSpinner, _ContrastSpinner, _ScrollSpeedSpinner, _FontSizeSpinner
    global _ScrollViewportAutoId
    ; Each entry: name -> { AutoId, Handler }
    return Map(
        "brightness", Map("AutoId", _BrightnessSpinner, "Handler", (root, id, d) => ApplyRangeValueDelta(root, id, d)),
        "contrast", Map("AutoId", _ContrastSpinner, "Handler", (root, id, d) => ApplyRangeValueDelta(root, id, d)),
        "scrollspeed", Map("AutoId", _ScrollSpeedSpinner, "Handler", (root, id, d) => ApplyRangeValueDelta(root, id, d)),
        "fontsize", Map("AutoId", _FontSizeSpinner, "Handler", (root, id, d) => ApplyRangeValueDelta(root, id, d)),
        "scroll", Map("AutoId", _ScrollViewportAutoId, "Handler", (root, id, d) => ApplyScrollDelta(root, id, d))
    )
}

; Apply a numeric delta to a RangeValue spinner
ApplyRangeValueDelta(root, autoId, delta, uiRangeValueId := 10003) {
    el := FindByAutoId(root, autoId)
    if !el
        return false
    try {
        rvp := el.GetCurrentPattern(uiRangeValueId) ; RangeValuePattern = 10003
        if !rvp
            return false
        cur := rvp.value              ; numeric value from UIA
        rvp.SetValue(cur + delta)     ; UIA clamps to [min,max]
        return true
    } catch {
        return false
    }
}

; Apply a delta to the Prompter viewport using ScrollPattern or wheel fallback
ApplyScrollDelta(root, autoId, delta, uiScrollId := 10004) {
    global SCROLL_PERCENT_PER_STEP, SCROLL_WHEEL_PER_STEP
    vp := FindPrompterViewport(root, autoId)
    if !vp
        return false

    ; Try ScrollPattern
    sp := 0
    try sp := vp.GetCurrentPattern(uiScrollId) ; ScrollPattern = 10004
    if sp {
        ; Relative scroll if supported
        try {
            sp.Scroll(0, (delta > 0) ? 4 : 1) ; 4=LargeIncrement, 1=LargeDecrement
            return true
        } catch {
        }
        ; Percent fallback
        try {
            cur := sp.VerticalScrollPercent  ; -1 means unsupported
            if (cur >= 0) {
                step := BASE_STEP * (delta > 0 ? +1 : -1) ; SCROLL_PERCENT_PER_STEP
                newp := cur + step
                if (newp < 0) newp := 0
                    if (newp > 100) newp := 100
                        sp.SetScrollPercent(sp.HorizontalScrollPercent, newp)
                return true
            }
        }
        catch {
            return false
        }
    }
}

; Try to resolve the Prompter's scrolling viewport element
FindPrompterViewport(root, autoId) {
    ; 1) Direct AutomationId match
    el := FindByAutoId(root, autoId)
    if el
        return el

    ; 2) Look inside QScrollArea for QTextBrowser child
    for area in root.FindElements({ ClassName: "QScrollArea" }) {
        for kid in area.FindElements({}) {
            try if (kid.ClassName = "QTextBrowser")
                return kid
        }
    }

    ; 3) Fallback: first QTextBrowser anywhere
    for kid in root.FindElements({ ClassName: "QTextBrowser" })
        return kid

    return 0  ; nothing found
}

; Helper: find an element by AutomationId
FindByAutoId(root, autoId) {
    if !root
        return 0
    if !autoId
        return 0
    try {
        return root.FindElement({ AutomationId: autoId })
    } catch {
        return 0
    }
}

/*
ApplyAccumulated() {
    global _pending, _applyArmed, SHOW_PATH_TIP
    global _ScrollSpeedSpinner, _FontSizeSpinner, _BrightnessSpinner, _ContrastSpinner, _UIA_RangeValuePatternId
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
        "contrast", _ContrastSpinner,
        "scrollspeed", _ScrollSpeedSpinner,
        "fontsize", _FontSizeSpinner
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
        ; MsgBox "Type of autoId: " Type(autoId) "`nValue: " autoId
        ;ClassName:"EVHSpinBox"
        matches := ""
        for el in uiaElement.FindElements({ ClassName: "EVHSpinBox" })
            matches .= el.Dump() "`n"
        ;MsgBox "All elements with type ClassName: `n`n" matches

        matches := ""
        for el in uiaElement.FindElements({ ClassName: "EVHSpinBox" })
            matches .= el.Dump() "`n"

        if (matches = "") {
            Log("ApplyAccumulated: EVHSpinBox not found. Running quick scan.")
            c1 := Scan(uiaElement, { ClassName: "EVHSpinBox" }, 'ClassName:"EVHSpinBox"', 10)
            c2 := Scan(uiaElement, { ControlType: "Spinner" }, 'ControlType:"Spinner"', 10)
            c3 := Scan(uiaElement, { ControlType: "Slider" }, 'ControlType:"Slider"', 10)
            MsgBox "Quick scan — EVHSpinBox: " c1 "  Spinner: " c2 "  Slider: " c3
        } else {
            ;MsgBox "All elements with type ClassName: `n`n" matches
        }


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
        catch as e {
            Log("FindElement ERROR: " e.Message "  What=" e.What "  Extra=" e.Extra)
            MsgBox "Error in FindElement: " e.Message "`nWhat: " e.What "`nExtra: " e.Extra
            A_Clipboard := "Message: " e.Message "`nWhat: " e.What "`nExtra: " e.Extra
        }

        ; Zero out after applying (even if it failed—keeps the queue bounded)
        _pending[controlName] := 0
    }

    ; Optional: tooltip that shows what we just targeted

    if SHOW_PATH_TIP && summary {
        Tip("Applied:`n" summary)
    }
}
*/

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
    hwnd := GetCamHubHwnd()
    if !hwnd {
        MsgBox "Could not find Camera Hub window."
        return
    }
    uiaElement := UIA.ElementFromHandle(hwnd)
    if !uiaElement {
        MsgBox "UIA Failed"
    }
    return uiaElement
}

GetCamHubHwnd() {
    global APP_EXE
    bestHwnd := 0, bestScore := -999999

    ; Enumerate all top-level windows and pick the best candidate for Camera Hub
    for hwnd in WinGetList() {
        ; try-except guards for odd windows
        try {
            pid := WinGetPID("ahk_id " hwnd)
            exe := ProcessGetPath(pid)
            if !exe || !InStr(Exe, "\Camera Hub.exe")
                continue

            cls := WinGetClass("ahk_id " hwnd)
            title := WinGetTitle("ahk_id " hwnd)

            ; Score the candidate. Prefer QWindowIcon, penalize ToolSaveBits/renderer.
            score := 0
            if InStr(cls, "QWindowIcon")
                score += 1000
            if InStr(cls, "ToolSaveBits") || InStr(cls, "Renderer")
                score -= 1000

            ; Light UIA probe (cheap): presence of a QScrollArea suggests the real UI shell
            root := UIA.ElementFromHandle(hwnd)
            if root {
                try score += (Scan(root, { ClassName: "QScrollArea" }, "probe:QScrollArea", 1) > 0) ? 50 : 0
                ; If we see EVHPrompterTextWidget (render surface), penalize further
                try score -= (Scan(root, { ClassName: "EVHPrompterTextWidget" }, "probe:TextWidget", 1) > 0) ? 50 : 0
            }

            ; Keep the best
            if (score > bestScore)
                bestScore := score, bestHwnd := hwnd

            ; Log candidates for visibility
            Log(Format("CamHub candidate hwnd={:#x} cls={} title={} score={}", hwnd, cls, title, score))
        }
    }

    if bestHwnd {
        Log(Format("GetCamHubHwnd: picked={:#x} score={}", bestHwnd, bestScore))
        return bestHwnd
    }

    ; Last-resort fallback (shouldn’t be hit anymore)
    Log("GetCamHubHwnd: fallback path hit")
    return WinExist("ahk_exe " APP_EXE)
}

;GetCamHubHwnd() {
;    global APP_EXE, WIN_CLASS_RX
;    ; Prefer the real UI window (Icon), then fall back to exe
;    hwnd := WinExist("ahk_class " WIN_CLASS_RX)
;    if !hwnd
;        hwnd := WinExist("ahk_exe " APP_EXE)
;    return hwnd
;}

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
        loop 10 {
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
DumpSubtree(uiaRoot, N := 50) {
    if !uiaRoot {
        Log("DumpSubtree: root is NULL")
        return
    }
    Log("=== UIA Subtree (first " N ") ===")
    cnt := 0
    for el in uiaRoot.FindElements({}) { ; empty condition == all descendants in this lib
        cnt += 1
        DumpOne(el, "#" cnt)
        if (cnt >= N)
            break
    }
    Log("Subtree enumerated: " cnt " elements (displayed up to " N ")")
}

; Count by a condition, list up to M
Scan(uiaRoot, condObj, label := "", M := 25) {
    if !uiaRoot {
        Log("Scan[" label "]: root is NULL")
        return 0
    }
    cnt := 0
    Log("=== Scan " label " ===")
    for el in uiaRoot.FindElements(condObj) {
        cnt += 1
        if (cnt <= M)
            DumpOne(el, "#" cnt)
    }
    Log("Scan " label ": total=" cnt)
    return cnt
}

; Convenience: run a full diagnostic pass
FullDiag() {
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
    DumpSubtree(root, 50)
    Scan(root, { ClassName: "EVHSpinBox" }, 'ClassName:"EVHSpinBox"', 25)
    ; Fallback: some builds expose spinners as ControlType: Spinner or Slider
    Scan(root, { ControlType: "Spinner" }, 'ControlType:"Spinner"', 25)
    Scan(root, { ControlType: "Slider" }, 'ControlType:"Slider"', 10)
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

Tip(t) {
    ToolTip t
    SetTimer(() => ToolTip(), -900)
}