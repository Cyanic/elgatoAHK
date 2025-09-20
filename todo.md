# TODO

## Must Fix
- [FIXED] `elgatoPrompter.ahk:409`: use the correct casing for `exe` when calling `InStr` so candidate windows are scored instead of always skipped.
- [FIXED] `elgatoPrompter.ahk:174-198`: declare `BASE_STEP` as `global` (or pass it in) before using it in the scroll percent fallback; otherwise the step stays zero and no scroll adjustment occurs.

## Nice to Have
- [FIXED] `elgatoPrompter.ahk:174-205`: either initialize and use `SCROLL_PERCENT_PER_STEP` / `SCROLL_WHEEL_PER_STEP` with a proper wheel fallback, or remove the unused globals to simplify the code.
- [FIXED] `elgatoPrompter.ahk:386-395`: replace the blocking `MsgBox` calls in `GetCamHubUiaElement` with non-modal `Tip`/`Log` notifications so diagnostics hotkeys don’t interrupt the user when the app is closed.
- [FIXED] `elgatoPrompter.ahk:403-436`: guard the expensive UIA `Scan` probes behind a debug flag to reduce routine logging and keep window selection lightweight.
