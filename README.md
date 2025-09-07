# Excel Multi-Phrase Search (VBA + Python Semantic Search)

This project combines Excel VBA with optional Python semantic matching to provide:

- Multi-phrase search across worksheets
- Highlighting of matching words/phrases with distinct colors
- CSV import into a designated worksheet with auto-formatting
- Reset and export utilities
- Optional semantic search with Python for intelligent test case matching

---

## 📂 Project Structure
```
excel-multi-phrase-search/
├─ src/                 # Python source files
│  ├─ sm.py             # main semantic search script
│  └─ semanticmatching.py (helper/alternate implementation)
│
├─ vba-src/             # exported VBA modules
│  ├─ Module1.bas
│  ├─ Sheet1.cls
│  ├─ Sheet2.cls
│  ├─ Sheet3.cls
│  └─ ThisWorkbook.cls
│
├─ examples/            # sample Excel workbooks
│  └─ tests.xlsm
│
├─ bin/                 # (expected location for sm.exe after build)
├─ README.md
├─ .gitignore
└─ .gitattributes
```

> Windows may show `.cls` files as “LaTeX Source File” due to file associations, but they are VBA class modules.

---

## 🚀 Usage

### Import into Excel
1. Open or create a `.xlsm` workbook.
2. Press `Alt+F11` to open the VBA editor.
3. Right-click the project → **Import File…** and select each `.bas`/`.cls` file from `vba-src/`.
4. Ensure **File → Options → Trust Center → Macro Settings → "Trust access to the VBA project object model"** is enabled.

### Export from Excel
Run the included `ExportAllVBA` macro to export all VBA components back into `vba-src/`.

### Running the Tools
- **MultiPhraseSearch** → Highlights and filters rows in *Test Docs* based on comma-separated terms (Instructions!B1).
- **ResetSearch** → Clears highlights, resets filters, and wipes SM values (keeps the header).
- **ImportCsvToTestDocs** → Imports a CSV into the *Test Docs* sheet and preserves SM column.
- **SearchWithPython** → Runs semantic search via `sm.exe` and updates SM column with scores (0.0–1.0) color-coded red → green.

---

## ✨ Features & Improvements

- **Semantic search integration**:
  - Calls `sm.py` (or compiled `sm.exe`) to generate similarity scores.
  - Scores populate column **A (SM)** with conditional formatting.
  - Always runs against the full dataset (ignores active filters).
- **Reset button enhancements**:
  - Clears the search box (Instructions!B1).
  - Wipes SM values but preserves header.
  - Removes conditional formatting and unhides rows.
- **Cleaner utilities**:
  - Unified `QuoteArg` helper for safe command-line calls.
  - Consolidated duplicate functions.
  - Stronger validation for paths and missing binaries.

---

## 🔧 Building the Semantic Matching Executable

This project integrates with a Python backend for semantic search.  
The backend can be compiled into a standalone `sm.exe` so end users don’t need Python installed.

### Requirements
- Python 3.9+ (tested with 3.12)
- Packages:
  ```
  pip install sentence-transformers torch pandas
  ```
- PyInstaller:
  ```
  pip install pyinstaller
  ```

### Build Instructions
From the project root:

```bash
py -m PyInstaller --onefile --name sm src/sm.py
```

This will create:

```
dist/sm.exe
```

Move `sm.exe` into:

```
bin/sm.exe
```

so Excel can locate it.  
If using `examples/tests.xlsm`, the VBA expects `..\bin\sm.exe` relative to the workbook’s folder.

---

## 📖 Instructions Panel (in Excel)

On the **Instructions sheet**:

1. Enter a search term in **B1**.
2. **Search** → highlights and filters matching rows.
3. **Python Search** → runs semantic scoring across all rows (ignores filters).
4. **Reset** → clears SM column, highlights, filters, and search box.
5. **Import** → load a CSV into *Test Docs* (SM column is preserved).

---

## 📝 Notes
- Do not commit compiled binaries (`sm.exe`) — only sources.  
- Use `vba-src/` as the source of truth for macros.  
- For reproducibility, commit after each `ExportAllVBA`.  
