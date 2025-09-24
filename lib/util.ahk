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

Tip(t) {
    global TOOLTIP_HIDE_DELAY_MS
    if !IsSet(TOOLTIP_HIDE_DELAY_MS) || (Trim(TOOLTIP_HIDE_DELAY_MS) = "") {
        DEBUG_LOGTOOLTIP_HIDE_DELAY_MS := 1000
    }
    ToolTip(t)
    SetTimer(() => ToolTip(), -TOOLTIP_HIDE_DELAY_MS)
}

QuitApp() {
    Tip("EXITING")
    Sleep(1000)
    ExitApp()
}