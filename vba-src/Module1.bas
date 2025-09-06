Attribute VB_Name = "Module1"
Option Explicit

' ====== CONFIG ======
Private Const INPUT_SHEET As String = "Instructions"     ' where the search terms live
Private Const INPUT_CELL As String = "B1"          ' comma-separated terms
Private Const DATA_SHEET As String = "Test Docs"      ' where the data lives (expanded cases)
' ====================

' Optional: set to True to match whole words only (requires regex)
Private Const WHOLE_WORD_ONLY As Boolean = False

' Public entry point for your button
Public Sub MultiPhraseSearch()
    Dim wsInput As Worksheet, wsData As Worksheet
    Dim dataRng As Range, usedRng As Range
    Dim terms As Variant, colors() As Long
    Dim i As Long, term As String
    Dim r As Range, c As Range
    Dim hadAnyMatch As Boolean
    Dim rowHadMatch As Boolean
    
    On Error GoTo CleanFail
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Set wsInput = ThisWorkbook.Worksheets(INPUT_SHEET)
    Set wsData = ThisWorkbook.Worksheets(DATA_SHEET)
    
    ' Determine the data range on Sheet2 (USED RANGE)
    Set usedRng = wsData.UsedRange
    If usedRng Is Nothing Then GoTo CleanExit
    Set dataRng = usedRng
    
    ' Clear previous formatting/row hiding
    Call ClearSearchFormatting(wsData, dataRng)
    dataRng.EntireRow.Hidden = False
    
    ' Parse search terms from Sheet1!B1
    terms = ParseTerms(wsInput.Range(INPUT_CELL).Value)
    If IsEmpty(terms) Then GoTo CleanExit
    
    ' Assign a random color per term
    ReDim colors(LBound(terms) To UBound(terms))
    Randomize
    For i = LBound(terms) To UBound(terms)
        colors(i) = RandomNiceColor()
    Next i
    
    ' Optional: set up regex object only once if needed
    Dim rx As Object
    If WHOLE_WORD_ONLY Then
        Set rx = CreateObject("VBScript.RegExp")
        rx.Global = True
        rx.IgnoreCase = True
    End If
    
    ' Scan rows and cells; highlight substrings; hide rows without any hits
    hadAnyMatch = False
    For Each r In dataRng.rows
        rowHadMatch = False
        
        For Each c In r.Cells
            If Len(c.Value2) > 0 Then
                ' For each term, find and color matches
                For i = LBound(terms) To UBound(terms)
                    term = terms(i)
                    If Len(term) > 0 Then
                        If WHOLE_WORD_ONLY Then
                            ' Whole-word using regex: \bterm\b with character positions via Execute
                            rowHadMatch = HighlightMatchesRegex(c, term, colors(i), rx) Or rowHadMatch
                        Else
                            ' Simple contains (case-insensitive)
                            rowHadMatch = HighlightMatchesInStr(c, term, colors(i)) Or rowHadMatch
                        End If
                    End If
                Next i
            End If
        Next c
        
        If Not rowHadMatch Then
            r.EntireRow.Hidden = True
        Else
            hadAnyMatch = True
        End If
    Next r
    
    If Not hadAnyMatch Then
        MsgBox "No matches found for: " & Join(terms, ", "), vbInformation
    End If

CleanExit:
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub

CleanFail:
    MsgBox "Error: " & Err.Description, vbExclamation
    Resume CleanExit
End Sub


' Clears prior substring coloring and restores defaults
Private Sub ClearSearchFormatting(ByVal ws As Worksheet, ByVal rng As Range)
    Dim c As Range
    For Each c In rng.Cells
        If Not IsError(c) Then
            ' Reset whole-cell formatting
            c.Font.ColorIndex = xlColorIndexAutomatic
            c.Interior.ColorIndex = xlColorIndexNone

            ' Reset per-character FONT formatting (no Interior on Characters!)
            If Len(c.Value2) > 0 Then
                c.Characters(1, Len(CStr(c.Value2))).Font.ColorIndex = xlColorIndexAutomatic
                c.Characters(1, Len(CStr(c.Value2))).Font.Bold = False
                c.Characters(1, Len(CStr(c.Value2))).Font.Underline = xlUnderlineStyleNone
            End If
        End If
    Next c
End Sub


' Highlight using InStr (case-insensitive), all occurrences; returns True if any match
Private Function HighlightMatchesInStr(ByVal cell As Range, ByVal term As String, ByVal clr As Long) As Boolean
    Dim txt As String
    Dim pos As Long, tLen As Long
    Dim hit As Boolean

    txt = CStr(cell.Value2)
    tLen = Len(term)
    If tLen = 0 Or Len(txt) = 0 Then Exit Function

    pos = InStr(1, txt, term, vbTextCompare)
    Do While pos > 0
        With cell.Characters(pos, tLen).Font
            .Color = clr            ' color the matching text
            .Bold = True            ' (optional) make it bold
            .Underline = xlUnderlineStyleSingle  ' (optional) underline
        End With
        hit = True
        pos = InStr(pos + tLen, txt, term, vbTextCompare)
    Loop

    HighlightMatchesInStr = hit
End Function



' Highlight using whole-word regex; returns True if any match
' Note: finds character positions by counting length of Left(...) due to lack of direct index in VBScript.RegExp
Private Function HighlightMatchesRegex(ByVal cell As Range, ByVal term As String, ByVal clr As Long, ByVal rx As Object) As Boolean
    Dim txt As String, patt As String
    Dim matches As Object, m As Object
    Dim startPos As Long, hit As Boolean

    txt = CStr(cell.Value2)
    If Len(txt) = 0 Then Exit Function

    patt = "\b" & EscapeRegex(term) & "\b"
    rx.Pattern = patt
    Set matches = rx.Execute(txt)
    If matches Is Nothing Then Exit Function

    For Each m In matches
        startPos = CLng(m.FirstIndex) + 1
        With cell.Characters(startPos, CLng(m.Length)).Font
            .Color = clr
            .Bold = True
            .Underline = xlUnderlineStyleSingle
        End With
        hit = True
    Next m

    HighlightMatchesRegex = hit
End Function



' Escape regex special chars in a term
Private Function EscapeRegex(ByVal s As String) As String
    Dim chars As Variant, ch As Variant
    chars = Array("\", ".", "^", "$", "|", "(", ")", "[", "]", "{", "}", "*", "+", "?", "-")
    For Each ch In chars
        s = Replace$(s, CStr(ch), "\" & CStr(ch))
    Next ch
    EscapeRegex = s
End Function


' Parse comma-separated terms from the input cell; trims and removes empties/dupes
Private Function ParseTerms(ByVal raw As String) As Variant
    Dim parts As Variant, out() As String
    Dim i As Long, t As String
    Dim dict As Object
    
    raw = Trim$(raw)
    If Len(raw) = 0 Then Exit Function
    
    parts = Split(raw, ",")
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1 ' TextCompare
    
    For i = LBound(parts) To UBound(parts)
        t = Trim$(CStr(parts(i)))
        If Len(t) > 0 Then
            If Not dict.Exists(t) Then dict.Add t, True
        End If
    Next i
    
    If dict.Count = 0 Then Exit Function
    
    ReDim out(0 To dict.Count - 1)
    i = 0
    Dim k As Variant
    For Each k In dict.Keys
        out(i) = CStr(k)
        i = i + 1
    Next k
    
    ParseTerms = out
End Function


' Generate a bright, dynamic random color
Private Function RandomNiceColor() As Long
    Dim r As Long, g As Long, b As Long
    Dim maxVal As Long, minVal As Long
    
    maxVal = 255
    minVal = 100   ' avoid very dark shades
    
    ' Randomize three channels
    r = Int((maxVal - minVal + 1) * Rnd) + minVal
    g = Int((maxVal - minVal + 1) * Rnd) + minVal
    b = Int((maxVal - minVal + 1) * Rnd) + minVal
    
    ' Boost saturation by making sure one channel is strong
    Select Case Int(3 * Rnd)
        Case 0: r = maxVal
        Case 1: g = maxVal
        Case 2: b = maxVal
    End Select
    
    RandomNiceColor = RGB(r, g, b)
End Function

Public Sub ImportCsvToTestDocs()
    Dim fPath As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Choose a CSV file to import"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "CSV files", "*.csv"
        If .Show <> -1 Then Exit Sub
        fPath = .SelectedItems(1)
    End With

    Dim wbCSV As Workbook, src As Worksheet
    Dim dst As Worksheet
    Application.ScreenUpdating = False

    ' Open the CSV in a temp workbook (robust parsing)
    Set wbCSV = Workbooks.Open(Filename:=fPath, Local:=True)
    Set src = wbCSV.Worksheets(1)

    ' Use existing "Test Docs" sheet
    On Error Resume Next
    Set dst = ThisWorkbook.Worksheets("Test Docs")
    On Error GoTo 0
    If dst Is Nothing Then
        MsgBox "Sheet 'Test Docs' not found.", vbExclamation
        wbCSV.Close SaveChanges:=False
        Exit Sub
    End If

    ' Clear old contents and ensure SM column is in A
    dst.Cells.Clear
    EnsureSMColumn dst, "SM"   ' inserts column A + header if needed

    ' Paste CSV starting at B1 (since A is reserved for SM)
    Dim rowsCount As Long, colsCount As Long
    rowsCount = src.UsedRange.rows.Count
    colsCount = src.UsedRange.Columns.Count
    dst.Range("B1").Resize(rowsCount, colsCount).Value = src.UsedRange.Value

    ' Clean possible BOM in the top-left header
    dst.Range("B1").Value = Replace(CStr(dst.Range("B1").Value), ChrW(&HFEFF), "")

    ' Autofit columns and apply SM formatting
    dst.Columns.AutoFit
    ApplySMFormatting dst, headerRow:=1, smCol:=1   ' 1 = column A

    wbCSV.Close SaveChanges:=False
    Application.ScreenUpdating = True

    MsgBox "CSV imported into 'Test Docs' with SM column ready.", vbInformation
End Sub


Private Sub EnsureSMColumn(ByVal ws As Worksheet, ByVal headerText As String)
    ' Inserts column A with the given header, unless A1 already equals that header.
    If LCase$(CStr(ws.Cells(1, 1).Value)) <> LCase$(headerText) Then
        ws.Columns(1).Insert Shift:=xlToRight
        ws.Cells(1, 1).Value = headerText
    End If
End Sub


' === Dropdown support ===
Public Sub RefreshSheetPicker()
    Dim wsIn As Worksheet: Set wsIn = ThisWorkbook.Worksheets(INPUT_SHEET)
    Dim rngList As Range: Set rngList = wsIn.Range(SHEET_LIST_RANGE)
    Dim ws As Worksheet, i As Long

    rngList.ClearContents
    i = 0
    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible = xlSheetVisible And ws.Name <> INPUT_SHEET Then
            i = i + 1
            rngList.Cells(i, 1).Value = ws.Name
        End If
    Next ws

    With wsIn.Range(SHEET_PICKER_CELL).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:="=" & rngList.Address
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = True
    End With

    wsIn.Range("D1").Value = "Search sheet:"
    ' optional: wsIn.Columns("H").Hidden = True
End Sub

' === Helpers ===

Private Function GetBaseName(ByVal fullPath As String) As String
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    GetBaseName = fso.GetBaseName(fullPath)
End Function



Private Function MakeUniqueSheetName(ByVal wb As Workbook, ByVal base As String) As String
    Dim nameTry As String, n As Long
    nameTry = base
    Do While SheetExists(wb, nameTry)
        n = n + 1
        nameTry = Left$(base, 31 - Len(CStr(n)) - 1) & "_" & CStr(n)
    Loop
    MakeUniqueSheetName = nameTry
End Function



Private Function SanitizeSheetName(ByVal s As String) As String
    ' Remove invalid chars and trim to 31 chars
    Dim bad As Variant, i As Long
    bad = Array("[", "]", ":", "*", "?", "/", "\", Chr(34)) ' "
    For i = LBound(bad) To UBound(bad)
        s = Replace$(s, CStr(bad(i)), "_")
    Next
    s = Trim$(s)
    If Len(s) = 0 Then s = "Import"
    If Len(s) > 31 Then s = Left$(s, 31)
    SanitizeSheetName = s
End Function


Private Function SheetExists(ByVal wb As Workbook, ByVal nm As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(nm)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function


' Call this after your CSV import completes:
'   FormatImportedSheet ActiveSheet
Public Sub FormatImportedSheet(ByVal ws As Worksheet)
    Const HEADER_ROW As Long = 1

    Dim usedRng As Range
    Set usedRng = ws.UsedRange
    If usedRng Is Nothing Then Exit Sub

    Dim lastCol As Long
    lastCol = ws.Cells(HEADER_ROW, ws.Columns.Count).End(xlToLeft).Column

    Dim col As Long
    Dim hdr As String, hdrL As String

    ' 1) Default: turn OFF wrap for all columns, light tidy
    ws.Cells.WrapText = False
    ws.rows.RowHeight = ws.StandardHeight

    ' 2) Walk headers and selectively enable wrapping + widths
    For col = 1 To lastCol
        hdr = CStr(ws.Cells(HEADER_ROW, col).Value2)
        hdrL = LCase$(Trim$(hdr))

        ' Big-text columns: wrap + wider column
        If hdrL = "description" _
           Or hdrL = "expected result" _
           Or EndsWithDetails(hdr) Then

            ws.Columns(col).WrapText = True
            ws.Columns(col).ColumnWidth = 50   ' adjust as you like

        ' Titles/short fields: keep unwrapped, narrower width
        ElseIf hdrL = "title" _
            Or hdrL = "test id" _
            Or hdrL Like "step #" _
            Or hdrL Like "step ##" _
            Or hdrL Like "step ###" Then

                ws.Columns(col).WrapText = False
                ws.C4olumns(col).ColumnWidth = 20

        Else
            ' Everything else: reasonable width
            ws.Columns(col).ColumnWidth = 18
        End If
    Next col

    ' 3) Make header bold and a touch of shading (optional)
    With ws.rows(HEADER_ROW)
        .Font.Bold = True
        .Interior.Color = RGB(242, 242, 242)
    End With

    ' 4) AutoFilter (optional)
    If Not ws.AutoFilterMode Then
        ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(HEADER_ROW, lastCol)).AutoFilter
    End If

    ' 5) Now that widths & wrap are set, autosize row heights so wrapped text shows fully
    ws.rows.AutoFit

    ' 6) Freeze header row (optional)
    ws.Activate
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
End Sub

Private Function EndsWithDetails(ByVal s As String) As Boolean
    s = LCase$(Trim$(s))
    EndsWithDetails = (Right$(s, 7) = "details")
End Function


' Exports all VBA components (Modules/Classes/Forms) to a folder next to the workbook.
Public Sub ExportAllVBA()
    Dim exportPath As String, fso As Object, vbComp As Object, ext As String
    
    exportPath = ThisWorkbook.Path & "\vba-src"
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(exportPath) Then fso.CreateFolder exportPath
    
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Select Case vbComp.Type
            Case 1: ext = ".bas" ' Standard module
            Case 2: ext = ".cls" ' Class module
            Case 3: ext = ".frm" ' UserForm (will also write .frx automatically)
            Case 100: ext = ".cls" ' Document module (e.g., ThisWorkbook/Sheet modules)
            Case Else: ext = ".txt"
        End Select
        vbComp.Export exportPath & "\" & SafeName(vbComp.Name) & ext
    Next vbComp
    
    MsgBox "Exported VBA to: " & exportPath, vbInformation
End Sub

Private Function SafeName(ByVal s As String) As String
    Dim bad As Variant
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    Dim i As Long
    For i = LBound(bad) To UBound(bad)
        s = Replace$(s, bad(i), "_")
    Next
    SafeName = s
End Function


Sub RunSemanticMatching()
    Dim sh As Object
    Dim py As String, script As String, wbPath As String, assertion As String
    Dim inner As String, cmd As String

    ' Raw paths (UNQUOTED here)
    py = "C:\Users\chris\AppData\Local\Programs\Python\Python312\python.exe"
    script = "C:\Users\chris\Documents\excel-multi-phrase-search\semanticmatching.py"
    wbPath = ThisWorkbook.FullName

    ' Get assertion text
    assertion = Sheets("Sheet1").Range("B6").Value   ' <-- change if needed

    ' Build the inner command (quoted piece by piece)
    inner = QuoteArg(py) & " " & _
            QuoteArg(script) & " " & _
            QuoteArg(wbPath) & " " & _
            QuoteArg(assertion)

    ' Run via cmd so we can keep the window open for debugging
    cmd = Environ$("ComSpec") & " /k " & QuoteArg(inner)

    ' Show exactly what we’re about to run (for debugging)
    Debug.Print cmd
    MsgBox cmd, vbInformation, "Command being run"

    Set sh = CreateObject("WScript.Shell")
    sh.Run cmd, 1, True
End Sub



' Button handler: send query to Python, then recolor SM column on DATA_SHEET
Public Sub SearchWithPython()
    Dim wsInstr As Worksheet, wsData As Worksheet
    Dim query As String
    Dim ok As Boolean
    
    Dim pyExeChk As String, pyScriptChk As String
    pyExeChk = GetPythonExePath()
    pyScriptChk = GetPyScriptPath()

    On Error GoTo Fail

    Set wsInstr = ThisWorkbook.Worksheets(INPUT_SHEET)
    Set wsData = ThisWorkbook.Worksheets(DATA_SHEET)
    
    ' NEW: clear filters/hidden rows so Python sees everything
    EnsureAllRowsVisibleUnfiltered wsData

    query = Trim$(CStr(wsInstr.Range(INPUT_CELL).Value))
    If Len(query) = 0 Then
        MsgBox "Enter search text in " & INPUT_SHEET & "!" & INPUT_CELL, vbInformation
        Exit Sub
    End If

    ' Optional sanity checks so we fail fast if paths are wrong
    If Len(pyExeChk) = 0 Or Dir(pyExeChk) = "" Then
        MsgBox "Python.exe path is not set or invalid. Set it on the Instructions sheet (B3).", vbExclamation
        Exit Sub
    End If
    If Len(pyScriptChk) = 0 Or Dir(pyScriptChk) = "" Then
        MsgBox "Semantic script path is not set or invalid. Set it on the Instructions sheet (B4).", vbExclamation
        Exit Sub
    End If

    ' Call your Python script. It should write SM values to column A on DATA_SHEET.
    ok = RunPythonSemanticSimple(query, ThisWorkbook.FullName, DATA_SHEET)
    If Not ok Then GoTo Fail

    ' Re-apply the SM gradient on column A
    ApplySMFormatting wsData, headerRow:=1, smCol:=1

    MsgBox "Semantic search complete. SM values updated.", vbInformation
    Exit Sub

Fail:
    MsgBox "Search failed: " & Err.Description, vbExclamation
End Sub
Private Function RunPythonSemanticSimple(ByVal query As String, _
                                         ByVal workbookPath As String, _
                                         ByVal sheetName As String) As Boolean
    Dim sh As Object, cmd As String, rc As Long
    Dim pyExe As String, pyScript As String

    pyExe = GetPythonExePath()
    pyScript = GetPyScriptPath()

    ' Validate + helpful errors
    If Len(pyExe) = 0 Then
        MsgBox "Please set the Python.exe path on the Instructions sheet (B3).", vbExclamation
        RunPythonSemanticSimple = False: Exit Function
    End If
    If Dir(pyExe) = "" Then
        MsgBox "Python.exe not found: " & pyExe, vbExclamation
        RunPythonSemanticSimple = False: Exit Function
    End If
    If Len(pyScript) = 0 Then
        MsgBox "Please set the semantic script path on the Instructions sheet (B4).", vbExclamation
        RunPythonSemanticSimple = False: Exit Function
    End If
    If Dir(pyScript) = "" Then
        MsgBox "Script not found: " & pyScript, vbExclamation
        RunPythonSemanticSimple = False: Exit Function
    End If

    Set sh = CreateObject("WScript.Shell")

    cmd = """" & pyExe & """ " & _
          """" & pyScript & """ " & _
          "--query " & QuoteArg(query) & " " & _
          "--workbook " & QuoteArg(workbookPath) & " " & _
          "--sheet " & QuoteArg(sheetName)

    rc = sh.Run(cmd, 0, True)  ' 0=hidden, True=wait
    RunPythonSemanticSimple = (rc = 0)
End Function

Private Function QuoteArg(ByVal s As String) As String
    QuoteArg = """" & Replace(s, """", "\""") & """"
End Function

Private Function ExportSheetToTempCsv(ByVal ws As Worksheet) As String
    Dim fso As Object, tmp As String, arr As Variant
    Dim rng As Range, rows As Long, cols As Long
    Dim r As Long, c As Long
    Dim fh As Integer, line As String

    Set rng = ws.UsedRange
    If rng Is Nothing Then Err.Raise 5, , "No data to export."

    rows = rng.rows.Count: cols = rng.Columns.Count
    arr = rng.Value  ' 2D variant

    tmp = BuildTempPath("sheet_export_", ".csv")
    fh = FreeFile
    Open tmp For Output As #fh

    For r = 1 To rows
        line = ""
        For c = 1 To cols
            line = line & IIf(c > 1, ",", "") & CsvEscape(arr(r, c))
        Next c
        Print #fh, line
    Next r

    Close #fh
    ExportSheetToTempCsv = tmp
End Function

Private Function CsvEscape(ByVal v As Variant) As String
    Dim s As String
    s = CStr(v)
    If InStr(s, """") > 0 Or InStr(s, ",") > 0 Or InStr(s, vbLf) > 0 Or InStr(s, vbCr) > 0 Then
        s = """" & Replace(s, """", """""") & """"
    End If
    CsvEscape = s
End Function

Private Function BuildTempPath(ByVal prefix As String, ByVal ext As String) As String
    Dim folder As String
    folder = Environ$("TEMP")
    If Right$(folder, 1) <> "\" Then folder = folder & "\"
    BuildTempPath = folder & prefix & Format(Now, "yyyymmdd_hhnnss_") & CLng(Rnd * 1000000#) & ext
End Function
Private Sub ApplySMFormatting(ByVal ws As Worksheet, _
                              Optional ByVal headerRow As Long = 1, _
                              Optional ByVal smCol As Long = 1)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.Count, smCol).End(xlUp).Row
    If lastRow <= headerRow Then Exit Sub

    Dim rng As Range
    Set rng = ws.Range(ws.Cells(headerRow + 1, smCol), ws.Cells(lastRow, smCol))

    ' Clear prior rules
    rng.FormatConditions.Delete

    ' 1) Stop-if-true rule for blanks / non-numbers / out-of-range (no format)
    Dim colLetter As String
    colLetter = Split(ws.Cells(1, smCol).Address(True, False), "$")(0)
    Dim formulaStop As String
    formulaStop = "=OR(" & colLetter & (headerRow + 1) & "=""""," & _
                       "NOT(ISNUMBER(" & colLetter & (headerRow + 1) & "))," & _
                       colLetter & (headerRow + 1) & "<0," & _
                       colLetter & (headerRow + 1) & ">1)"
    With rng.FormatConditions.Add(Type:=xlExpression, Formula1:=formulaStop)
        .StopIfTrue = True
    End With

    ' 2) 2-color scale for values 0..1 (red -> green)
    Dim cs As ColorScale
    Set cs = rng.FormatConditions.AddColorScale(ColorScaleType:=2)
    With cs.ColorScaleCriteria(1)
        .Type = xlConditionValueNumber
        .Value = 0
        .FormatColor.Color = RGB(255, 0, 0)     ' red at 0.0
    End With
    With cs.ColorScaleCriteria(2)
        .Type = xlConditionValueNumber
        .Value = 1
        .FormatColor.Color = RGB(0, 176, 80)    ' green at 1.0
    End With

    ' Make header bold and center (optional)
    With ws.rows(headerRow)
        .Font.Bold = True
        .VerticalAlignment = xlCenter
    End With
End Sub
' Reset button: clears highlights, SM values (A2:A…), and search box on INPUT_SHEET!B1
Public Sub ResetSearch()
    Dim wsData As Worksheet, wsInstr As Worksheet
    Dim rng As Range, tgt As Range
    Dim lastRowA As Long
    Dim smHeader As String

    On Error GoTo CleanFail
    Application.ScreenUpdating = False

    Set wsData = ThisWorkbook.Worksheets(DATA_SHEET)
    Set wsInstr = ThisWorkbook.Worksheets(INPUT_SHEET)

    ' --- Clear formatting & show all rows on DATA sheet ---
    Set rng = wsData.UsedRange
    If Not rng Is Nothing Then
        ' Remember header so it never gets wiped
        smHeader = CStr(wsData.Cells(1, 1).Value)

        ' Clear in-cell highlights and unhide rows
        Call ClearSearchFormatting(wsData, rng)
        rng.EntireRow.Hidden = False

        ' Clear filters (harmless if none)
        On Error Resume Next
        If wsData.AutoFilterMode Then wsData.ShowAllData
        wsData.AutoFilterMode = False
        On Error GoTo 0

        ' Remove conditional formatting from SM column
        wsData.Columns(1).FormatConditions.Delete

        ' Clear ONLY SM data cells (keep A1)
        lastRowA = wsData.Cells(wsData.rows.Count, 1).End(xlUp).Row
        If lastRowA >= 2 Then
            wsData.Range(wsData.Cells(2, 1), wsData.Cells(lastRowA, 1)).ClearContents
        End If

        ' Ensure header remains (default to "SM" if blank)
        If Len(Trim$(smHeader)) = 0 Then smHeader = "SM"
        wsData.Cells(1, 1).Value = smHeader
    End If

    ' --- Clear the search box on INPUT sheet (B1) robustly ---
    On Error Resume Next
    Set tgt = wsInstr.Range("B1").MergeArea   ' handles merged cells
    If tgt Is Nothing Then Set tgt = wsInstr.Range("B1")
    If wsInstr.ProtectContents Then wsInstr.Unprotect Password:=""
    tgt.ClearContents
    On Error GoTo 0

    ' Put the cursor back into B1
    wsInstr.Activate
    wsInstr.Range("B1").Select

CleanExit:
    Application.ScreenUpdating = True
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "Reset failed: " & Err.Description, vbExclamation
End Sub
' Make sure all rows are visible and no filters are applied
Private Sub EnsureAllRowsVisibleUnfiltered(ByVal ws As Worksheet)
    On Error Resume Next
    ' If an AutoFilter is applied, show all rows
    If ws.AutoFilterMode Then ws.ShowAllData
    ws.AutoFilterMode = False
    On Error GoTo 0

    ' Also unhide any rows hidden by the VBA search
    If Not ws.UsedRange Is Nothing Then
        ws.UsedRange.EntireRow.Hidden = False
    End If
End Sub

' === Replace hard-coded constants with these getters ===

Private Function GetPythonExePath() As String
    On Error Resume Next
    ' Prefer Named Range; fall back to literal cell
    GetPythonExePath = Trim$(CStr(ThisWorkbook.Worksheets(INPUT_SHEET).Range("PythonExePath").Value))
    If Len(GetPythonExePath) = 0 Then
        GetPythonExePath = Trim$(CStr(ThisWorkbook.Worksheets(INPUT_SHEET).Range("B24").Value))
    End If
End Function

Private Function GetPyScriptPath() As String
    On Error Resume Next
    GetPyScriptPath = Trim$(CStr(ThisWorkbook.Worksheets(INPUT_SHEET).Range("PyScriptPath").Value))
    If Len(GetPyScriptPath) = 0 Then
        GetPyScriptPath = Trim$(CStr(ThisWorkbook.Worksheets(INPUT_SHEET).Range("B26").Value))
    End If
End Function
