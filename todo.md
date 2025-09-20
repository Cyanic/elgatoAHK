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
- [FIXED] `elgatoPrompter.ahk:668-672`: trim the INI value before lowercasing in `IniReadBool` so trailing spaces don’t break parsing.
- [FIXED] `elgatoPrompter.ahk:214-223`: narrow the QTextBrowser search inside `FindPrompterViewport` to avoid enumerating entire subtrees.
- [FIXED] INI/Configuration: The script uses hard-coded values for spinner AutomationIds and file paths. Consider optionally reading these from the INI file for easier adjustment without editing the script.
- [FIXED] Magic Numbers: Constants are defined at the top, which is good! Consider allowing BASE_STEP or APPLY_DELAY_MS to be set from the INI file for dynamic tuning.
- [FIXED] For deeper diagnostics - log the nature of failures (e.g., which GetCurrentPattern failed with which code), only when a debug flag is set.
- [FIXED] Logging which control lookups are missing if debug/diagnostic mode is enabled, to help with field mapping.
- [FIXED] For even greater clarity around lambdas in GetControlSpecs, explicitly define named handler functions first, aiding readability and debugging.
- [FIXED] Add a Help hotkey listing all control hotkeys and configuration files in a message box, for new users.
- [FIXED] Add a Help hotkey listing all control hotkeys and configuration files in as a comment in the .ini file, for new users.
