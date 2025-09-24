; Accumulates detent input for a control and schedules processing.
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
        SetTimer(ApplyAccumulated, -APPLY_DELAY_MS)
    }
}

; Processes queued pulses against Camera Hub UI elements.
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

; Adjusts a spinner/slider via UIA RangeValue pattern.
ApplyRangeValueDelta(root, spec, pulses, uiRangeValueId := _UIA_RangeValuePatternId) {
    global DEBUG_VERBOSE_LOGGING, _UIA_RangeValuePatternId
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
        rvp := el.GetCurrentPattern(uiRangeValueId)
        if !rvp
            return HandlerResult(false)
        cur := rvp.value
        rvp.SetValue(cur + delta)
        return HandlerResult(true, FormatSigned(delta))
    } catch as err {
        if DEBUG_VERBOSE_LOGGING
            Log("ApplyRangeValueDelta: pattern error " spec["Name"] " msg=" err.Message)
        return HandlerResult(false)
    }
}

; Scrolls the prompter viewport using UIA Scroll pattern fallbacks.
ApplyScrollDelta(root, spec, pulses, uiScrollId := _UIA_ScrollId) {
    global DEBUG_VERBOSE_LOGGING, _UIA_ScrollId
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
    try sp := vp.GetCurrentPattern(uiScrollId)
    if sp {
        pulsesToSend := Abs(pulses)
        dir := (pulses > 0) ? 4 : 1
        if (pulsesToSend > 0) {
            try {
                Loop pulsesToSend
                    sp.Scroll(0, dir)
                return HandlerResult(true, FormatSigned(pulsesToSend, " step" (pulsesToSend = 1 ? "" : "s")))
            } catch as err {
                if DEBUG_VERBOSE_LOGGING
                    Log("ApplyScrollDelta: Scroll call failed dir=" dir " pulses=" pulsesToSend " msg=" err.Message)
            }
        }

        percentPer := spec.Has("PercentPerStep") ? spec["PercentPerStep"] : 0
        if (percentPer != 0) {
            try {
                cur := sp.VerticalScrollPercent
                if (cur >= 0) {
                    deltaPercent := pulses * percentPer
                    newp := ClampPercent(cur + deltaPercent)
                    sp.SetScrollPercent(sp.HorizontalScrollPercent, newp)
                    return HandlerResult(true, FormatSigned(deltaPercent, "%"))
                }
            } catch as err {
                if DEBUG_VERBOSE_LOGGING
                    Log("ApplyScrollDelta: SetScrollPercent failed msg=" err.Message)
            }
        }
    } else if DEBUG_VERBOSE_LOGGING {
        Log("ApplyScrollDelta: ScrollPattern unavailable for " spec["Name"])
    }
    return HandlerResult(false)
}