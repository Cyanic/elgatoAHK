# TODO

## showClasses.ahk
- Wrap `CallbackFree(gEnumCallback)` in a `try…finally` so the callback is freed even if enumeration raises an error.
- Trim user input and INI-derived strings before building the window search query to avoid whitespace mismatches.
- Strip inline comments (text after `;`) from INI values before using them in `WinExist` searches.
- Guard `FileDelete`/`FileAppend` calls so missing or locked files do not throw and abort logging.

## showClasses.ini
- Move comments onto their own lines (or rely on the loader stripping them) so `ClassNN` and `Process` values stay clean tokens.
