# TODO

## showClasses.ahk
- Allow the Win32 traversal fallback to run even when UIA isn't available so native-only windows still show up.
- Release the root UIA element after each search pass to avoid leaking COM references.
- Trim user input and INI-derived strings before building the window search query to avoid whitespace mismatches.
- Strip inline comments (text after `;`) from INI values before using them in `WinExist` searches.
- Surface a friendly note if the UIA traversal hits its safety limit (currently 5000 nodes) so large trees are obvious.
- Replace the UIA traversal queue's `RemoveAt(1)` usage with an index-based iteration to avoid O(n²) behavior on large trees.

## showClasses.ini
- Document that only one window definition should remain active at a time to prevent ambiguous targeting.
- Move comments onto their own lines (or rely on the loader stripping them) so `ClassNN` and `Process` values stay clean tokens.
