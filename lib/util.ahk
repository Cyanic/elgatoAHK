; Wraps handler outcomes in a consistent success/detail map.
HandlerResult(success, detail := "") {
    return Map("Success", !!success, "Detail", detail)
}

; Formats numeric deltas with sign and optional suffix.
FormatSigned(value, suffix := "") {
    if !IsNumber(value)
        return value suffix
    str := (Round(value) = value) ? Format("{:+d}", value) : Format("{:+.2f}", value)
    return str suffix
}

; Limits scroll percentages to the UIA-valid 0-100 range.
ClampPercent(value) {
    if (value < 0)
        return 0
    if (value > 100)
        return 100
    return value
}

; Joins an array of strings with CRLF delimiters.
JoinLines(arr) {
    s := ""
    for v in arr
        s .= v "`r`n"
    return RTrim(s, "`r`n")
}

; Reads an INI field and normalizes common truthy values.
IniReadBool(file, section, key, default := false) {
    defStr := default ? "1" : "0"
    val := IniRead(file, section, key, defStr)
    txt := StrLower(Trim(val))
    return (txt = "1" || txt = "true" || txt = "yes" || txt = "on")
}

; Reads an INI field and coerces it to a number when possible.
IniReadNumber(file, section, key, default) {
    val := Trim(IniRead(file, section, key, default))
    return IsNumber(val) ? val + 0 : default
}

; Appends a timestamped line to the configured debug log.
Log(line) {
    global DEBUG_LOG
    if !IsSet(DEBUG_LOG) || (Trim(DEBUG_LOG) = "") {
        DEBUG_LOG := A_ScriptDir "\PrompterDebug.txt"
    }
    ts := FormatTime(, "yyyy-MM-dd HH:mm:ss")
    try FileAppend(
        ts "  " line "`r`n",
        DEBUG_LOG
    )
}

; Shows a tooltip briefly for operator feedback.
Tip(t) {
    global TOOLTIP_HIDE_DELAY_MS
    if !IsSet(TOOLTIP_HIDE_DELAY_MS) || (Trim(TOOLTIP_HIDE_DELAY_MS) = "") {
        DEBUG_LOGTOOLTIP_HIDE_DELAY_MS := 1000
    }
    ToolTip(t)
    SetTimer(() => ToolTip(), -TOOLTIP_HIDE_DELAY_MS)
}

; Displays an exit notice and terminates the script.
QuitApp() {
    Tip("EXITING")
    Sleep(1000)
    ExitApp()
}
