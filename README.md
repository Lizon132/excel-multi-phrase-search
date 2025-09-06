# Excel Multi-Phrase Search (VBA)

This project contains exported VBA modules from an Excel workbook that implements:

- Multi-phrase search across a worksheet
- Highlighting of matching words/phrases with distinct colors
- CSV import into new worksheets with auto-formatting
- Export of all VBA modules for version control

## 📂 Project Structure
```
excel-multi-phrase-search/
├─ vba-src/              # exported VBA source files
│  ├─ Module1.bas        # main code (search, import, formatting, export)
│  ├─ Sheet1.cls         # code-behind for Sheet1
│  ├─ Sheet2.cls         # code-behind for Sheet2
│  ├─ Sheet3.cls         # code-behind for Sheet3
│  └─ ThisWorkbook.cls   # workbook-level events/code
├─ README.md
├─ .gitignore
└─ .gitattributes
```

> Windows may show `.cls` files as "LaTeX Source File" because of file association, but they are VBA class modules.

## 🚀 Usage

### Import into Excel
1. Open or create a `.xlsm` workbook.
2. Press `Alt+F11` to open the VBA editor.
3. Right-click the project → **Import File…** and select each `.bas`/`.cls` file from `vba-src/`.
4. Make sure **File → Options → Trust Center → Macro Settings → "Trust access to the VBA project object model"** is enabled.

### Export from Excel
Run the included `ExportAllVBA` macro to export all VBA components back into `vba-src/`.

### Running the Tools
- **MultiPhraseSearch**: Looks up comma-separated search terms from `Sheet1!C1`, highlights matches in `Sheet2`, and hides non-matching rows.
- **ResetSearch**: Clears formatting and unhides rows.
- **ImportCsvToNewSheetFormatted**: Prompts for a CSV, loads it into a new sheet, and auto-formats `Description`, `Expected Result`, and `*Details` columns with word wrap.

## 📝 Notes
- Keep the workbook (`.xlsm`) out of version control; use the `vba-src/` folder as the source of truth.
- Commit after each export to capture code changes in Git.

## Recent Updates

### ✨ New Features
- Added **semantic search integration** with Python:
  - Calls `sm.py` via VBA to generate similarity scores (0.0–1.0).
  - Writes scores to **SM column (A)** in *Test Docs* sheet.
  - Applies a red → green gradient to visualize match strength.
- Added **filter/visibility guard** so Python search always works on the full dataset (ignores prior Excel filters).
- Added ability to configure **Python.exe path** and **script path** from the **Instructions sheet** (cells B3 and B4).
  - Falls back to constants if fields are left blank.
  - Optional **Browse** buttons to make path selection user-friendly.

### 🛠 Improvements
- Updated **Reset button**:
  - Clears the search box (Instructions!B1).
  - Clears all SM values in column A but keeps the **SM** header.
  - Removes SM conditional formatting.
  - Restores full row visibility and clears filters.
- Hardened path validation and error messages for Python/script.
- Improved `QuoteArg` function for safer command-line argument handling.
- Consolidated utility helpers to avoid duplicate definitions.

### 📖 Instructions Panel
- Text box on **Instructions sheet** now explains:
  1. Enter a search term in **B1**.
  2. Click **Search** → highlights and filters matches.
  3. Click **Python Search** → computes semantic scores for all rows.
  4. Click **Reset** → clears SM values, highlights, and search box.
  5. Click **Import** → load a CSV into *Test Docs* (SM column is preserved).
