; Locates the scrolling text viewport within Camera Hub using heuristics.
FindPrompterViewport(root, spec) {
    el := LocatePrompterViewportInRoot(root, spec)
    if el
        return el

    if spec.Has("ToolClassRegex") {
        el := FindPrompterViewportInToolWindow(spec["ToolClassRegex"], spec)
        if el
            return el
    }

    return 0
}

LocatePrompterViewportInRoot(root, spec) {
    if !root
        return 0

    autoId := spec.Has("AutoId") ? spec["AutoId"] : ""
    if autoId {
        el := FindByAutoId(root, autoId)
        if el
            return el
    }

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

FindPrompterViewportInToolWindow(classRx, spec) {
    global APP_EXE, DEBUG_VERBOSE_LOGGING
    for hwnd in WinGetList("ahk_exe " APP_EXE) {
        try {
            cls := WinGetClass("ahk_id " hwnd)
            if !RegExMatch(cls, classRx)
                continue
            toolRoot := UIA.ElementFromHandle(hwnd)
            if !toolRoot {
                if DEBUG_VERBOSE_LOGGING
                    Log("FindPrompterViewport: UIA root NULL for tool hwnd=" Format("{:#x}", hwnd))
                continue
            }
            vp := LocatePrompterViewportInRoot(toolRoot, spec)
            if vp
                return vp
        } catch as err {
            if DEBUG_VERBOSE_LOGGING
                Log("FindPrompterViewport: tool window scan failed msg=" err.Message)
        }
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
