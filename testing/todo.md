# TODO

## showClasses.ahk
- Trim user input and INI-derived strings before building the window search query to avoid whitespace mismatches.
- Strip inline comments (text after `;`) from INI values before using them in `WinExist` searches.
- Surface a friendly note if the UIA traversal hits its safety limit (currently 5000 nodes) so large trees are obvious.

## showClasses.ini
- Move comments onto their own lines (or rely on the loader stripping them) so `ClassNN` and `Process` values stay clean tokens.
