# scrollText TODO

## Fixes
- Release COM objects from `uia.CreatePropertyCondition` in `FindElementByAutomationId`, `FindElementByClassName`, and `FindElementByClassWithSize` to prevent leaks.
- Release COM tree walkers created in `FindAncestorByClass` to avoid handle leaks during repeated use.
- Harden direction handling so unexpected tokens do not silently default to down-scroll.

## Enhancements
- Cache the resolved config and window handle between scroll invocations; refresh only when invalid.
- Replace blocking `MsgBox` calls with non-modal notifications (e.g., `TrayTip`/logging) so hotkeys do not interrupt foreground work.
- Expose scroll cadence tuning (delay, step size) and optional horizontal scrolling support.
- Make the QScrollArea fallback match configurable (dimensions/tolerance) or enrich matching criteria to survive UI/layout changes.
