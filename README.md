# DOCX & XLSX Lite

Two tiny native macOS apps I built to open Word (`.docx`) and Excel (`.xlsx`) files without installing a full office suite. One Swift file each, no dependencies, under 1 MB per app.

| | App | What it does |
|---|---|---|
| <img src="docx-lite/icon.png" width="64"> | [**Lettore DOCX**](docx-lite/) | Reads `.docx`: text, formatting, tables, images, links. Search (⌘F), zoom, copy, print / save as PDF. Read-only: the original file is never changed. |
| <img src="xlsx-lite/icon.png" width="64"> | [**Lettore XLSX**](xlsx-lite/) | Opens `.xlsx`: all sheets, cell styles, number and date formats, column filters. Double-click a cell to edit it, undo, and save a copy. |

The interface is in Italian.

## Requirements

- Mac with Apple Silicon, macOS 13 or later
- Xcode Command Line Tools (`xcode-select --install`) to compile

## Build

```bash
git clone https://github.com/lucasflati12/docx-xlsx-lite.git
cd docx-xlsx-lite
bash docx-lite/build.sh   # → ~/Desktop/Lettore DOCX.app
bash xlsx-lite/build.sh   # → ~/Desktop/Lettore XLSX.app
```

Pass a path to build somewhere else, e.g. `bash docx-lite/build.sh "/Applications/Lettore DOCX.app"`. The apps are signed ad hoc for your own Mac. Then right-click a file → **Open With** to use them.

## Limits

- No exact Word layout: floating objects, headers/footers, charts and equations may look different than in Microsoft Word.
- Formulas are not recalculated in the app (Excel recalculates them when you open the saved copy).
- Password-protected files are not supported.

## Tests (Lettore DOCX)

```bash
cd docx-lite
python3 test-fixtures.py /tmp/lettoredocx-tests
xcrun swiftc -D DEBUG -swift-version 5 main.swift ReaderTests.swift -o /tmp/lettoredocx-tests/ReaderTests
/tmp/lettoredocx-tests/ReaderTests --smoke-test /tmp/lettoredocx-tests/rich.docx
```

## License

MIT, see [LICENSE](LICENSE).

---

Independent project, not affiliated with or endorsed by Microsoft. Microsoft, Word and Excel are trademarks of Microsoft Corporation, mentioned only to describe file compatibility.
