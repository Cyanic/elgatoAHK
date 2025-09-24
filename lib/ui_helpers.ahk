; Locates the scrolling text viewport within Camera Hub using heuristics.
FindPrompterViewport(root, spec) {
    autoId := spec.Has("AutoId") ? spec["AutoId"] : ""
    el := autoId ? FindByAutoId(root, autoId) : 0
    if el
        return el

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
    }

    try {
        qb := root.FindElement({ ClassName: "QTextBrowser" })
        if qb
            return qb
    }
    catch as err {
    }

    return 0
}

; Helper that safely resolves a UIA element by AutomationId.
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

; Resolves a UIA element via resolver callback or AutomationId.
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

; Retrieves and caches the UIA root for the Camera Hub window.
GetCamHubUiaElement() {
    global _CachedCamHubHwnd
    hwnd := GetCamHubHwnd()
    if !hwnd {
        if _CachedCamHubHwnd {
            _CachedCamHubHwnd := 0
        }
        Log("GetCamHubUiaElement: Camera Hub window not found")
        return
    }

    if (_CachedCamHubHwnd != hwnd)
        _CachedCamHubHwnd := hwnd

    uiaElement := UIA.ElementFromHandle(hwnd)
    if !uiaElement
        Log("GetCamHubUiaElement: UIA.ElementFromHandle returned NULL")
    return uiaElement
}

; Finds the target Camera Hub window handle by exe and class.
GetCamHubHwnd() {
    global APP_EXE, WIN_CLASS_RX

    if WIN_CLASS_RX {
        for candidate in WinGetList("ahk_exe " APP_EXE) {
            try {
                cls := WinGetClass("ahk_id " candidate)
                if RegExMatch(cls, WIN_CLASS_RX)
                    return candidate
            }
        }
    }

    hwnd := WinExist("ahk_exe " APP_EXE)
    if hwnd
        return hwnd

    return WIN_CLASS_RX ? WinExist("ahk_class " WIN_CLASS_RX) : 0
}
