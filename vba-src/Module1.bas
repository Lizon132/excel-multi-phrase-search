Attribute VB_Name = "Module1"
Option Explicit

' ====== CONFIG ======
Private Const INPUT_SHEET As String = "Instructions"   ' where the search term lives
Private Const INPUT_CELL  As String = "B1"             ' search box
Private Const DATA_SHEET  As String = "Test Docs"      ' data sheet (SM column in A)
' ====================

' Optional: set to True to match whole words only (requires regex)
Private Const WHOLE_WORD_ONLY As Boolean = False


' =========================
'  MULTI-PHRASE (VBA) SEARCH
' =========================
Public Sub MultiPhraseSearch()
    Dim wsInput As Worksheet, wsData As Worksheet
    Dim dataRng As Range, usedRng As Range
    Dim terms As Variant, colors() As Long
    Dim i As Long, term As String
    Dim r As Range, c As Range
    Dim hadAnyMatch As Boolean
    Dim rowHadMatch As Boolean
    Dim rx As Object

    On Error GoTo CleanFail
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Set wsInput = ThisWorkbook.Worksheets(INPUT_SHEET)
    Set wsData = ThisWorkbook.Worksheets(DATA_SHEET)

    Set usedRng = wsData.UsedRange
    If usedRng Is Nothing Then GoTo CleanExit
    Set dataRng = usedRng

    ' Clear previous formatting/row hiding
    ClearSearchFormatting wsData, dataRng
    dataRng.EntireRow.Hidden = False

    ' Parse terms
    terms = ParseTerms(CStr(wsInput.Range(INPUT_CELL).Value))
    If IsEmpty(terms) Then GoTo CleanExit

    ' Assign bright colors, one per term
    ReDim colors(LBound(terms) To UBound(terms))
    Randomize
    For i = LBound(terms) To UBound(terms)
        colors(i) = RandomNiceColor()
    Next i

    If WHOLE_WORD_ONLY Then
        Set rx = CreateObject("VBScript.RegExp")
        rx.Global = True
        rx.IgnoreCase = True
    End If

    hadAnyMatch = False
    For Each r In dataRng.rows
        rowHadMatch = False

        For Each c In r.Cells
            If Len(CStr(c.Value2)) > 0 Then
                For i = LBound(terms) To UBound(terms)
                    term = terms(i)
                    If Len(term) > 0 Then
                        If WHOLE_WORD_ONLY Then
                            rowHadMatch = HighlightMatchesRegex(c, term, colors(i), rx) Or rowHadMatch
                        Else
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


' Clear prior substring formatting (font color/bold/underline) and cell fills
Private Sub ClearSearchFormatting(ByVal ws As Worksheet, ByVal rng As Range)
    Dim c As Range, n As Long
    For Each c In rng.Cells
        If Not IsError(c) Then
            c.Font.ColorIndex = xlColorIndexAutomatic
            c.Interior.ColorIndex = xlColorIndexNone
            n = Len(CStr(c.Value2))
            If n > 0 Then
                With c.Characters(1, n).Font
                    .ColorIndex = xlColorIndexAutomatic
                    .Bold = False
                    .Underline = xlUnderlineStyleNone
                End With
            End If
        End If
    Next c
End Sub

' Highlight using InStr (case-insensitive)
Private Function HighlightMatchesInStr(ByVal cell As Range, ByVal term As String, ByVal clr As Long) As Boolean
    Dim txt As String, pos As Long, tLen As Long, hit As Boolean
    txt = CStr(cell.Value2)
    tLen = Len(term)
    If tLen = 0 Or Len(txt) = 0 Then Exit Function

    pos = InStr(1, txt, term, vbTextCompare)
    Do While pos > 0
        With cell.Characters(pos, tLen).Font
            .Color = clr
            .Bold = True
            .Underline = xlUnderlineStyleSingle
        End With
        hit = True
        pos = InStr(pos + tLen, txt, term, vbTextCompare)
    Loop
    HighlightMatchesInStr = hit
End Function

' Regex whole-word highlighting
Private Function HighlightMatchesRegex(ByVal cell As Range, ByVal term As String, ByVal clr As Long, ByVal rx As Object) As Boolean
    Dim txt As String, patt As String
    Dim matches As Object, m As Object
    Dim startPos As Long, hit As Boolean

    txt = CStr(cell.Value2)
    If Len(txt) = 0 Then Exit Function

    patt = "\b" & EscapeRegex(term) & "\b"
    rx.Pattern = patt
    Set matches = rx.Execute(txt)

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

Private Function EscapeRegex(ByVal s As String) As String
    Dim chars As Variant, ch As Variant
    chars = Array("\", ".", "^", "$", "|", "(", ")", "[", "]", "{", "}", "*", "+", "?", "-")
    For Each ch In chars
        s = Replace$(s, CStr(ch), "\" & CStr(ch))
    Next ch
    EscapeRegex = s
End Function

Private Function ParseTerms(ByVal raw As String) As Variant
    Dim parts As Variant, out() As String
    Dim i As Long, t As String
    Dim dict As Object, k As Variant

    raw = Trim$(raw)
    If Len(raw) = 0 Then Exit Function

    parts = Split(raw, ",")
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1 ' TextCompare

    For i = LBound(parts) To UBound(parts)
        t = Trim$(CStr(parts(i)))
        If Len(t) > 0 Then If Not dict.Exists(t) Then dict.Add t, True
    Next i

    If dict.Count = 0 Then Exit Function
    ReDim out(0 To dict.Count - 1)
    i = 0
    For Each k In dict.Keys
        out(i) = CStr(k)
        i = i + 1
    Next k
    ParseTerms = out
End Function

Private Function RandomNiceColor() As Long
    Dim r As Long, g As Long, b As Long, maxVal As Long, minVal As Long
    maxVal = 255: minVal = 100
    r = Int((maxVal - minVal + 1) * Rnd) + minVal
    g = Int((maxVal - minVal + 1) * Rnd) + minVal
    b = Int((maxVal - minVal + 1) * Rnd) + minVal
    Select Case Int(3 * Rnd)
        Case 0: r = maxVal
        Case 1: g = maxVal
        Case 2: b = maxVal
    End Select
    RandomNiceColor = RGB(r, g, b)
End Function


' =========================
'  CSV IMPORT ? "Test Docs"
' =========================
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

    Dim wbCSV As Workbook, src As Worksheet, dst As Worksheet
    Application.ScreenUpdating = False

    Set wbCSV = Workbooks.Open(Filename:=fPath, Local:=True)
    Set src = wbCSV.Worksheets(1)

    On Error Resume Next
    Set dst = ThisWorkbook.Worksheets(DATA_SHEET)
    On Error GoTo 0
    If dst Is Nothing Then
        MsgBox "Sheet '" & DATA_SHEET & "' not found.", vbExclamation
        wbCSV.Close SaveChanges:=False
        Exit Sub
    End If

    dst.Cells.Clear
    EnsureSMColumn dst, "SM" ' keep SM in A

    ' Paste CSV at B1 (A is SM)
    Dim rowsCount As Long, colsCount As Long
    rowsCount = src.UsedRange.rows.Count
    colsCount = src.UsedRange.Columns.Count
    dst.Range("B1").Resize(rowsCount, colsCount).Value = src.UsedRange.Value

    ' Strip BOM from B1 if present
    dst.Range("B1").Value = Replace(CStr(dst.Range("B1").Value), ChrW(&HFEFF), "")

    dst.Columns.AutoFit
    ApplySMFormatting dst, headerRow:=1, smCol:=1

    wbCSV.Close SaveChanges:=False
    Application.ScreenUpdating = True
    MsgBox "CSV imported into '" & DATA_SHEET & "' and SM is ready.", vbInformation
End Sub

Private Sub EnsureSMColumn(ByVal ws As Worksheet, ByVal headerText As String)
    If LCase$(CStr(ws.Cells(1, 1).Value)) <> LCase$(headerText) Then
        ws.Columns(1).Insert Shift:=xlToRight
        ws.Cells(1, 1).Value = headerText
    End If
End Sub


' =========================
'  OPTIONAL DROPDOWN SUPPORT
' =========================
Public Sub RefreshSheetPicker()
    ' Requires named ranges SHEET_LIST_RANGE and SHEET_PICKER_CELL on Instructions
    Dim wsIn As Worksheet: Set wsIn = ThisWorkbook.Worksheets(INPUT_SHEET)
    Dim rngList As Range: Set rngList = wsIn.Range("SHEET_LIST_RANGE")
    Dim ws As Worksheet, i As Long

    rngList.ClearContents
    i = 0
    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible = xlSheetVisible And ws.Name <> INPUT_SHEET Then
            i = i + 1
            rngList.Cells(i, 1).Value = ws.Name
        End If
    Next ws

    With wsIn.Range("SHEET_PICKER_CELL").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:="=" & rngList.Address
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = True
    End With

    wsIn.Range("D1").Value = "Search sheet:"
End Sub


' =========================
'  MISC HELPERS
' =========================
Private Function SheetExists(ByVal wb As Workbook, ByVal nm As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(nm)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

' Exports all VBA components (Modules/Classes/Forms) to the top-level vba-src folder
Public Sub ExportAllVBA()
    Dim exportPath As String, fso As Object, vbComp As Object, ext As String
    Dim parentPath As String
    
    ' Step up one level from the workbook's folder (examples ? project root)
    parentPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(ThisWorkbook.Path)
    exportPath = parentPath & "\vba-src"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(exportPath) Then fso.CreateFolder exportPath
    
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Select Case vbComp.Type
            Case 1: ext = ".bas" ' Standard module
            Case 2: ext = ".cls" ' Class module
            Case 3: ext = ".frm" ' UserForm (writes .frm + .frx)
            Case 100: ext = ".cls" ' Document module (ThisWorkbook/Sheet modules)
            Case Else: ext = ".txt"
        End Select
        vbComp.Export exportPath & "\" & SafeName(vbComp.Name) & ext
    Next vbComp
    
    MsgBox "Exported VBA to: " & exportPath, vbInformation
End Sub

Private Function SafeName(ByVal s As String) As String
    Dim bad As Variant, i As Long
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For i = LBound(bad) To UBound(bad)
        s = Replace$(s, bad(i), "_")
    Next i
    SafeName = s
End Function

Private Function EndsWithDetails(ByVal s As String) As Boolean
    s = LCase$(Trim$(s))
    EndsWithDetails = (Right$(s, 7) = "details")
End Function


' =========================
'  FORMAT IMPORTED SHEET
' =========================
Public Sub FormatImportedSheet(ByVal ws As Worksheet)
    Const HEADER_ROW As Long = 1

    Dim usedRng As Range
    Dim lastCol As Long, col As Long
    Dim hdr As String, hdrL As String

    Set usedRng = ws.UsedRange
    If usedRng Is Nothing Then Exit Sub

    lastCol = ws.Cells(HEADER_ROW, ws.Columns.Count).End(xlToLeft).Column

    ' 1) defaults
    ws.Cells.WrapText = False
    ws.rows.RowHeight = ws.StandardHeight

    ' 2) header-based width/wrap rules
    For col = 1 To lastCol
        hdr = CStr(ws.Cells(HEADER_ROW, col).Value2)
        hdrL = LCase$(Trim$(hdr))

        If (hdrL = "description") Or (hdrL = "expected result") Or EndsWithDetails(hdrL) Then
            ws.Columns(col).WrapText = True
            ws.Columns(col).ColumnWidth = 50
        ElseIf (hdrL = "title") Or (hdrL = "test id") Or (hdrL Like "step *") Then
            ws.Columns(col).WrapText = False
            ws.Columns(col).ColumnWidth = 20
        Else
            ws.Columns(col).ColumnWidth = 18
        End If
    Next col

    ' 3) header styling
    With ws.rows(HEADER_ROW)
        .Font.Bold = True
        .Interior.Color = RGB(242, 242, 242)
    End With

    ' 4) AutoFilter
    If Not ws.AutoFilterMode Then
        ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(HEADER_ROW, lastCol)).AutoFilter
    End If

    ' 5) row heights
    ws.rows.AutoFit

    ' 6) Freeze header
    ws.Activate
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
End Sub


' =========================
'  PYTHON (EXE) SEMANTIC SEARCH
' =========================

' Preferred: local bin\sm.exe next to workbook; else B24 or named range SemanticExePath
Private Function GetSemanticExePath() As String
    Dim ws As Worksheet, p As String, localBin As String
    Set ws = ThisWorkbook.Worksheets(INPUT_SHEET)

    localBin = ThisWorkbook.Path & "\..\bin\sm.exe"
    If Dir(localBin) <> "" Then
        GetSemanticExePath = localBin
        Exit Function
    End If

    On Error Resume Next
    p = Trim$(CStr(ws.Range("SemanticExePath").Value)) ' named range optional
    If Len(p) = 0 Then p = Trim$(CStr(ws.Range("B24").Value))
    On Error GoTo 0
    GetSemanticExePath = p
End Function

Private Function QuoteArg(ByVal s As String) As String
    ' Double embedded quotes for Windows command line
    QuoteArg = """" & Replace(s, """", """""") & """"
End Function

' Make sure all rows are visible and no filters are applied
Private Sub EnsureAllRowsVisibleUnfiltered(ByVal ws As Worksheet)
    On Error Resume Next
    If ws.AutoFilterMode Then ws.ShowAllData
    ws.AutoFilterMode = False
    On Error GoTo 0
    If Not ws.UsedRange Is Nothing Then ws.UsedRange.EntireRow.Hidden = False
End Sub

' Call sm.exe to compute SM and write to DATA_SHEET!A
Private Function RunSemanticExe(ByVal query As String, _
                                ByVal workbookPath As String, _
                                ByVal sheetName As String) As Boolean
    Dim sh As Object, cmd As String, exePath As String, rc As Long
    exePath = GetSemanticExePath()
    If Len(exePath) = 0 Or Dir(exePath) = "" Then
        MsgBox "sm.exe not found. Place it at: " & ThisWorkbook.Path & "\bin\sm.exe" & vbCrLf & _
               "or set Instructions!B24 (SemanticExePath).", vbExclamation
        RunSemanticExe = False
        Exit Function
    End If

    Set sh = CreateObject("WScript.Shell")
    cmd = """" & exePath & """ " & _
          "--query " & QuoteArg(query) & " " & _
          "--workbook " & QuoteArg(workbookPath) & " " & _
          "--sheet " & QuoteArg(sheetName)
    rc = sh.Run(cmd, 0, True)
    RunSemanticExe = (rc = 0)
End Function

' Button: run semantic search then color SM column
Public Sub SearchWithPython()
    Dim wsInstr As Worksheet, wsData As Worksheet
    Dim query As String, ok As Boolean

    On Error GoTo Fail

    Set wsInstr = ThisWorkbook.Worksheets(INPUT_SHEET)
    Set wsData = ThisWorkbook.Worksheets(DATA_SHEET)

    ' Ensure semantic runs on full dataset
    EnsureAllRowsVisibleUnfiltered wsData

    query = Trim$(CStr(wsInstr.Range(INPUT_CELL).Value))
    If Len(query) = 0 Then
        MsgBox "Enter search text in " & INPUT_SHEET & "!" & INPUT_CELL, vbInformation
        Exit Sub
    End If

    ok = RunSemanticExe(query, ThisWorkbook.FullName, DATA_SHEET)
    If Not ok Then GoTo Fail

    ApplySMFormatting wsData, headerRow:=1, smCol:=1
    MsgBox "Semantic search complete. SM values updated.", vbInformation
    Exit Sub

Fail:
    MsgBox "Search failed: " & Err.Description, vbExclamation
End Sub

' Color-scale 0..1 in SM column, blanks untouched
Private Sub ApplySMFormatting(ByVal ws As Worksheet, _
                              Optional ByVal headerRow As Long = 1, _
                              Optional ByVal smCol As Long = 1)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.Count, smCol).End(xlUp).Row
    If lastRow <= headerRow Then Exit Sub

    Dim rng As Range
    Set rng = ws.Range(ws.Cells(headerRow + 1, smCol), ws.Cells(lastRow, smCol))

    rng.FormatConditions.Delete

    ' stop rule for blanks/non-numbers/out-of-range
    Dim colLetter As String, formulaStop As String
    colLetter = Split(ws.Cells(1, smCol).Address(True, False), "$")(0)
    formulaStop = "=OR(" & colLetter & (headerRow + 1) & "=""""," & _
                  "NOT(ISNUMBER(" & colLetter & (headerRow + 1) & "))," & _
                  colLetter & (headerRow + 1) & "<0," & _
                  colLetter & (headerRow + 1) & ">1)"
    With rng.FormatConditions.Add(Type:=xlExpression, Formula1:=formulaStop)
        .StopIfTrue = True
    End With

    ' red ? green 0..1
    Dim cs As ColorScale
    Set cs = rng.FormatConditions.AddColorScale(ColorScaleType:=2)
    With cs.ColorScaleCriteria(1)
        .Type = xlConditionValueNumber
        .Value = 0
        .FormatColor.Color = RGB(255, 0, 0)
    End With
    With cs.ColorScaleCriteria(2)
        .Type = xlConditionValueNumber
        .Value = 1
        .FormatColor.Color = RGB(0, 176, 80)
    End With

    With ws.rows(headerRow)
        .Font.Bold = True
        .VerticalAlignment = xlCenter
    End With
End Sub


' =========================
'  RESET
' =========================
Public Sub ResetSearch()
    Dim wsData As Worksheet, wsInstr As Worksheet
    Dim rng As Range, tgt As Range
    Dim lastRowA As Long, smHeader As String

    On Error GoTo CleanFail
    Application.ScreenUpdating = False

    Set wsData = ThisWorkbook.Worksheets(DATA_SHEET)
    Set wsInstr = ThisWorkbook.Worksheets(INPUT_SHEET)

    Set rng = wsData.UsedRange
    If Not rng Is Nothing Then
        smHeader = CStr(wsData.Cells(1, 1).Value)

        ClearSearchFormatting wsData, rng
        rng.EntireRow.Hidden = False

        On Error Resume Next
        If wsData.AutoFilterMode Then wsData.ShowAllData
        wsData.AutoFilterMode = False
        On Error GoTo 0

        wsData.Columns(1).FormatConditions.Delete

        lastRowA = wsData.Cells(wsData.rows.Count, 1).End(xlUp).Row
        If lastRowA >= 2 Then
            wsData.Range(wsData.Cells(2, 1), wsData.Cells(lastRowA, 1)).ClearContents
        End If

        If Len(Trim$(smHeader)) = 0 Then smHeader = "SM"
        wsData.Cells(1, 1).Value = smHeader
    End If

    On Error Resume Next
    Set tgt = wsInstr.Range("B1").MergeArea
    If tgt Is Nothing Then Set tgt = wsInstr.Range("B1")
    If wsInstr.ProtectContents Then wsInstr.Unprotect Password:=""
    tgt.ClearContents
    On Error GoTo 0

    wsInstr.Activate
    wsInstr.Range("B1").Select

CleanExit:
    Application.ScreenUpdating = True
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "Reset failed: " & Err.Description, vbExclamation
End Sub


