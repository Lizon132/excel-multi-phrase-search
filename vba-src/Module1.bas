Attribute VB_Name = "Module1"
Option Explicit

' ====== CONFIG ======
Private Const INPUT_SHEET As String = "Sheet1"     ' where the search terms live
Private Const INPUT_CELL As String = "B1"          ' comma-separated terms
Private Const DATA_SHEET As String = "Sheet2"      ' where the data lives (expanded cases)
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
    For Each r In dataRng.Rows
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



' Quick reset button: clears highlights and shows all rows on DATA_SHEET
Public Sub ResetSearch()
    Dim wsData As Worksheet, rng As Range
    On Error GoTo Done
    Application.ScreenUpdating = False
    
    Set wsData = ThisWorkbook.Worksheets(DATA_SHEET)
    Set rng = wsData.UsedRange
    If Not rng Is Nothing Then
        Call ClearSearchFormatting(wsData, rng)
        rng.EntireRow.Hidden = False
    End If

Done:
    Application.ScreenUpdating = True
End Sub


' One-time (or anytime) import: pick your expanded file (CSV or XLSX) and load it into Sheet2
' - If CSV: it will use QueryTables to import
' - If XLSX: it will open the file and copy the first sheet into Sheet2
Public Sub ImportExpandedDataToSheet2()
    Dim fd As FileDialog
    Dim fPath As String
    Dim ws As Worksheet
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Choose expanded test cases file (CSV or XLSX)"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel/CSV", "*.xlsx;*.xls;*.csv", 1
        If .Show <> -1 Then Exit Sub
        fPath = .SelectedItems(1)
    End With
    
    Set ws = ThisWorkbook.Worksheets(DATA_SHEET)
    ws.Cells.Clear
    
    If LCase$(Right$(fPath, 4)) = ".csv" Then
        ' Import CSV into Sheet2
        With ws.QueryTables.Add(Connection:="TEXT;" & fPath, Destination:=ws.Range("A1"))
            .TextFileParseType = xlDelimited
            .TextFileCommaDelimiter = True
            .TextFileOtherDelimiter = False
            .TextFileColumnDataTypes = Array(1)
            .AdjustColumnWidth = True
            .Refresh BackgroundQuery:=False
        End With
    Else
        ' Import first worksheet from the selected workbook
        Dim wb As Workbook
        Dim src As Worksheet
        Set wb = Application.Workbooks.Open(fPath, ReadOnly:=True)
        Set src = wb.Worksheets(1)
        src.UsedRange.Copy ws.Range("A1")
        wb.Close SaveChanges:=False
        ws.Columns.AutoFit
    End If
    Call FormatImportedSheet(ws)
    MsgBox "Data imported to " & DATA_SHEET & ".", vbInformation
End Sub

Option Explicit

' === Import a CSV into a NEW worksheet in THIS workbook (UTF-8 safe) ===
Public Sub ImportCsvToNewSheet()
    Dim fPath As String
    Dim wbCSV As Workbook
    Dim srcWS As Worksheet, dstWS As Worksheet
    Dim srcRng As Range
    Dim baseName As String, newName As String
    Dim appScr As Boolean, appCalc As XlCalculation, appEvt As Boolean
    
    On Error GoTo CleanFail
    appScr = Application.ScreenUpdating
    appCalc = Application.Calculation
    appEvt = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' 1) Pick a CSV file
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Choose a CSV file to import"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "CSV files", "*.csv"
        If .Show <> -1 Then GoTo CleanExit
        fPath = .SelectedItems(1)
    End With

    ' 2) Open the CSV in a temporary workbook with UTF-8 handling
    ' (Works in modern Excel; falls back gracefully on older versions)
    Workbooks.OpenText _
        Filename:=fPath, _
        Origin:=65001, _
        DataType:=xlDelimited, _
        TextQualifier:=xlTextQualifierDoubleQuote, _
        Comma:=True, _
        Other:=False

    Set wbCSV = ActiveWorkbook
    Set srcWS = wbCSV.Worksheets(1)
    Set srcRng = srcWS.UsedRange
    If srcRng Is Nothing Then
        MsgBox "No data found in CSV.", vbInformation
        GoTo CleanExit
    End If

    ' 3) Create a new sheet in THIS workbook and name it after the file
    baseName = GetBaseName(fPath)
    newName = MakeUniqueSheetName(ThisWorkbook, SanitizeSheetName(baseName))
    Set dstWS = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    dstWS.Name = newName

    ' 4) Copy values (fast, no formatting surprises)
    dstWS.Range("A1").Resize(srcRng.Rows.Count, srcRng.Columns.Count).Value = srcRng.Value

    ' 5) Clean up visuals
    dstWS.Columns.AutoFit
    dstWS.Activate
    dstWS.Range("A1").Select

    ' 6) Close the temporary CSV workbook (no save)
    wbCSV.Close SaveChanges:=False

    MsgBox "Imported CSV to sheet: " & dstWS.Name, vbInformation

CleanExit:
    Application.EnableEvents = appEvt
    Application.Calculation = appCalc
    Application.ScreenUpdating = appScr
    Exit Sub

CleanFail:
    MsgBox "Import failed: " & Err.Description, vbExclamation
    On Error Resume Next
    If Not wbCSV Is Nothing Then wbCSV.Close SaveChanges:=False
    On Error GoTo 0
    Resume CleanExit
End Sub

' === Helpers ===

Private Function GetBaseName(ByVal fullPath As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetBaseName = fso.GetBaseName(fullPath)
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

Private Function MakeUniqueSheetName(ByVal wb As Workbook, ByVal base As String) As String
    Dim nameTry As String
    Dim n As Long
    nameTry = base
    Do While SheetExists(wb, nameTry)
        n = n + 1
        nameTry = Left$(base, 31 - Len(CStr(n)) - 1) & "_" & CStr(n)
    Loop
    MakeUniqueSheetName = nameTry
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
    Dim hdr As String

    ' 1) Default: turn OFF wrap for all columns, light tidy
    ws.Cells.WrapText = False
    ws.Rows.RowHeight = ws.StandardHeight

    ' 2) Walk headers and selectively enable wrapping + widths
    For col = 1 To lastCol
        hdr = Trim$(CStr(ws.Cells(HEADER_ROW, col).Value2))

        ' Big-text columns: wrap + wider column
        If LCase$(hdr) = "description" _
           Or LCase$(hdr) = "expected result" _
           Or EndsWithDetails(hdr) Then

            ws.Columns(col).WrapText = True
            ws.Columns(col).ColumnWidth = 50   ' ~50 characters wide (adjust if you like)

        ' Titles/short fields: keep unwrapped, narrower width
        ElseIf LCase$(hdr) = "title" _
           Or hdr Like "Step #"
           Or hdr Like "Step *" _
           Or LCase$(hdr) = "test id" Then

            ws.Columns(col).WrapText = False
            ws.Columns(col).ColumnWidth = 20   ' title-ish width

        Else
            ' Everything else: reasonable auto width
            ws.Columns(col).ColumnWidth = 18
        End If
    Next col

    ' 3) Make header bold and a touch of shading (optional)
    With ws.Rows(HEADER_ROW)
        .Font.Bold = True
        .Interior.Color = RGB(242, 242, 242)
    End With

    ' 4) AutoFilter (optional)
    If Not ws.AutoFilterMode Then ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(HEADER_ROW, lastCol)).AutoFilter

    ' 5) Now that widths & wrap are set, autosize row heights so wrapped text shows fully
    ws.Rows.AutoFit

    ' 6) Freeze header row (optional)
    ws.Activate
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
End Sub

' Helper: does a header end with "Details" (case-insensitive)?
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
