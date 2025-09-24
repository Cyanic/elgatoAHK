; Captures core diagnostics, copies them to clipboard, and logs.
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
    try FileAppend(
        txt "`r`n`r`n",
        DEBUG_LOG
    )
    MsgBox("Debug copied to clipboard.`nSaved to: " DEBUG_LOG, "Prompter Debug", "OK Iconi")
}

; Toggles UIA probing and persists the flag to the INI.
ToggleProbeScans() {
    global ENABLE_PROBE_SCANS, INI
    ENABLE_PROBE_SCANS := !ENABLE_PROBE_SCANS
    IniWrite(ENABLE_PROBE_SCANS ? 1 : 0, INI, "Debug", "ProbeScans")
    state := ENABLE_PROBE_SCANS ? "enabled" : "disabled"
    Tip("Probe scans " state)
    Log("ToggleProbeScans: " state)
}

; Displays a help dialog listing available hotkeys and file paths.
ShowHelp() {
    global INI, DEBUG_LOG
    cfg := GetHotkeyConfig()
    msg := "Elgato Prompter Hotkeys:`n"
    msg .= Format("{} / {} -> Scroll viewport slower/faster`n", cfg["ScrollUp"], cfg["ScrollDown"])
    msg .= Format("{} / {} -> Scroll speed spinner down/up`n", cfg["ScrollSpeedDown"], cfg["ScrollSpeedUp"])
    msg .= Format("{} / {} -> Brightness down/up`n", cfg["BrightnessDown"], cfg["BrightnessUp"])
    msg .= Format("{} / {} -> Contrast down/up`n", cfg["ContrastDown"], cfg["ContrastUp"])
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

; Formats the stored spinner calibration offsets for display.
GetSavedPointText(hwnd) {
    global INI
    dx := IniRead(INI, "Spinner", "DX", "")
    dy := IniRead(INI, "Spinner", "DY", "")
    if (dx = "" || dy = "")
        return "MISSING (press Ctrl+Alt+S over the spinner)"
    if hwnd {
        WinGetPos(&wx, &wy, &ww, &wh, "ahk_id " hwnd)
        return "DX=" dx " DY=" dy "  (screen " (wx + dx) "," (wy + dy) ")"
    } else {
        return "DX=" dx " DY=" dy
    }
}

; Saves the mouse offset relative to Camera Hub for spinner targeting.
SaveCalibration() {
    global INI
    hwnd := GetCamHubHwnd()
    if !hwnd {
        Tip("Camera Hub not found.")
        return
    }
    MouseGetPos(&mx, &my)
    WinGetPos(&wx, &wy, &ww, &wh, "ahk_id " hwnd)
    dx := mx - wx, dy := my - wy
    IniWrite(Round(dx), INI, "Spinner", "DX")
    IniWrite(Round(dy), INI, "Spinner", "DY")
    Tip("Saved spinner @ (" Round(dx) "," Round(dy) ")")
}

; Logs monitor geometry and active window DPI for troubleshooting.
DumpMonitors() {
    Log("=== Monitors ===")
    count := MonitorGetCount()
    Log("Monitor count: " count)
    Loop count {
        i := A_Index
        MonitorGet(i, &l, &t, &r, &b)
        name := ""
        try name := MonitorGetName(i)
        Log(Format("[{}] {}x{} @({},{})  Name: {}", i, r - l, b - t, l, t, name))
    }
    hwnd := WinActive("A")
    if hwnd {
        dpi := 0
        try dpi := DllCall("User32\\GetDpiForWindow", "ptr", hwnd, "uint")
        if dpi
            Log("Active window DPI: " dpi)
    }
}

; Enumerates windows and highlights the selected Camera Hub candidate.
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

; Records elevation status for the script and target process.
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
}

; Logs UIA element details for the control under the cursor.
DumpUnderMouse() {
    global MAX_ANCESTOR_DEPTH
    try {
        MouseGetPos(&mx, &my, &hwin)
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

; Writes a single UIA element summary with class, AutomationId, and rect.
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
    try ctype := el.ControlType
    try {
        r := el.BoundingRectangle
        rect := Format("[{},{},{},{}]", r.x, r.y, r.w, r.h)
    }
    Log(prefix "  Name=" name "  Class=" tempclass "  AutoId=" aid "  CtrlType=" ctype "  Rect=" rect)
}

; Logs the first N descendants of a UIA root element.
DumpSubtree(uiaRoot, limit := SUBTREE_LIST_LIMIT) {
    global SUBTREE_LIST_LIMIT
    if !uiaRoot {
        Log("DumpSubtree: root is NULL")
        return
    }
    Log("=== UIA Subtree (first " limit ") ===")
    cnt := 0
    for el in uiaRoot.FindElements({}) {
        cnt += 1
        DumpOne(el, "#" cnt)
        if (cnt >= limit)
            break
    }
    Log("Subtree enumerated: " cnt " elements (displayed up to " limit ")")
}

; Runs a UIA search with filters and logs each match up to a limit.
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

; Performs the comprehensive diagnostic routine and logs results.
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
    Scan(root, { ControlType: "Spinner" }, 'ControlType:"Spinner"', SCAN_LIST_LIMIT)
    Scan(root, { ControlType: "Slider" }, 'ControlType:"Slider"', SLIDER_SCAN_LIMIT)
    DumpUnderMouse()
    Log("===== FULL DIAGNOSTIC END =====")
}

; Returns a 1-based array of available monitor indexes.
MonitorGetList() {
    arr := []
    Loop MonitorGetCount()
        arr.Push(A_Index)
    return arr
}