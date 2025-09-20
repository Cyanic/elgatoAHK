# TODO

## Must Fix
- [FIXED] `elgatoPrompter.ahk:409`: use the correct casing for `exe` when calling `InStr` so candidate windows are scored instead of always skipped.
- [FIXED] `elgatoPrompter.ahk:174-198`: declare `BASE_STEP` as `global` (or pass it in) before using it in the scroll percent fallback; otherwise the step stays zero and no scroll adjustment occurs.
- [FIXED] `elgatoPrompter.ahk:182-205`: honor the magnitude of `delta` in `ApplyScrollDelta` so multiple queued scroll pulses result in matching UI scroll steps instead of a single adjustment.

## Nice to Have
- [FIXED] `elgatoPrompter.ahk:174-205`: either initialize and use `SCROLL_PERCENT_PER_STEP` / `SCROLL_WHEEL_PER_STEP` with a proper wheel fallback, or remove the unused globals to simplify the code.
- [FIXED] `elgatoPrompter.ahk:386-395`: replace the blocking `MsgBox` calls in `GetCamHubUiaElement` with non-modal `Tip`/`Log` notifications so diagnostics hotkeys don’t interrupt the user when the app is closed.
- [FIXED] `elgatoPrompter.ahk:403-436`: guard the expensive UIA `Scan` probes behind a debug flag to reduce routine logging and keep window selection lightweight.
- [FIXED] `elgatoPrompter.ahk:175-207`: update the comment describing a “wheel fallback” so it matches the current implementation.
- [FIXED] `elgatoPrompter.ahk:392-407`: reduce redundant tooltips/logs when the Camera Hub window is missing (e.g., keep only logging in one location) to avoid repeated popups during polling.
- [FIXED] `elgatoPrompter.ahk:18-19`: make `ENABLE_PROBE_SCANS` configurable at runtime (create INI toggle) so diagnostics can be enabled without editing the script.
