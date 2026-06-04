Attribute VB_Name = "modExcelTasks"
'==============================================================================
' modExcelTasks.bas
' Excel COM Automation Tasks E01-E12 -- VBA Implementation
'
' Each task creates a NEW Excel.Application instance for fair comparison
' with other languages (cross-process marshaling baseline).
'
' All tasks run in background (Visible=False), use performance switches,
' and restore settings in cleanup.
'==============================================================================
Option Explicit

' ---------------------------------------------------------------------------
' Helper: Ensure output directory exists and return its path
' ---------------------------------------------------------------------------
Private Function GetOutDir() As String
    Dim basePath As String
    basePath = ThisWorkbook.Path & "\..\..\out\vba\"

    ' Normalize path
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    basePath = fso.GetAbsolutePathName(basePath)
    If Right(basePath, 1) <> "\" Then basePath = basePath & "\"

    ' Create directory tree if needed
    CreateDirRecursive basePath, fso

    GetOutDir = basePath
    Set fso = Nothing
End Function

Private Sub CreateDirRecursive(ByVal dirPath As String, ByRef fso As Object)
    If fso.FolderExists(dirPath) Then Exit Sub

    Dim parentDir As String
    parentDir = fso.GetParentFolderName(dirPath)
    If Not fso.FolderExists(parentDir) Then
        CreateDirRecursive parentDir, fso
    End If
    fso.CreateFolder dirPath
End Sub

' ---------------------------------------------------------------------------
' Helper: Create a new background Excel.Application instance
' ---------------------------------------------------------------------------
Private Function CreateBackgroundApp() As Excel.Application
    Dim app As Excel.Application
    Set app = New Excel.Application
    app.Visible = False
    app.DisplayAlerts = False
    app.ScreenUpdating = False
    Set CreateBackgroundApp = app
End Function

' ---------------------------------------------------------------------------
' Helper: Clean shutdown of a background Excel.Application
' ---------------------------------------------------------------------------
Private Sub CleanupApp(ByRef app As Excel.Application)
    On Error Resume Next
    If Not app Is Nothing Then
        app.DisplayAlerts = False
        ' Close any remaining workbooks
        Dim wb As Workbook
        For Each wb In app.Workbooks
            wb.Close SaveChanges:=False
        Next wb
        app.Quit
    End If
    Set app = Nothing
    On Error GoTo 0
End Sub

'==============================================================================
' E01 - Workbook Lifecycle
' Create new workbook -> SaveAs .xlsx -> Close -> Reopen -> verify -> Close
'==============================================================================
Public Function E01_WorkbookLifecycle() As Boolean
    E01_WorkbookLifecycle = False

    Dim app As Excel.Application
    Dim wb As Workbook
    Dim wb2 As Workbook
    Dim outDir As String
    Dim filePath As String
    Dim passed As Boolean

    On Error GoTo E01_Error

    outDir = GetOutDir()
    filePath = outDir & "E01_lifecycle.xlsx"

    ' Delete existing file if present
    If Dir(filePath) <> "" Then Kill filePath

    ' Create new background instance
    Set app = CreateBackgroundApp()

    ' Create new workbook
    Set wb = app.Workbooks.Add

    ' Write a marker value
    wb.Worksheets(1).Cells(1, 1).Value2 = "E01-VBA-Lifecycle"

    ' SaveAs .xlsx (xlOpenXMLWorkbook = 51)
    wb.SaveAs Filename:=filePath, FileFormat:=51

    ' Close
    wb.Close SaveChanges:=False
    Set wb = Nothing

    ' Reopen and verify
    Set wb2 = app.Workbooks.Open(filePath)
    Dim readVal As String
    readVal = CStr(wb2.Worksheets(1).Cells(1, 1).Value2)

    passed = (readVal = "E01-VBA-Lifecycle")

    wb2.Close SaveChanges:=False
    Set wb2 = Nothing

    ' Verify file exists on disk
    If Dir(filePath) = "" Then passed = False

    If passed Then
        Debug.Print "  E01 Workbook Lifecycle: PASS"
        E01_WorkbookLifecycle = True
    Else
        Debug.Print "  E01 Workbook Lifecycle: FAIL (read back: " & readVal & ")"
    End If

E01_Cleanup:
    CleanupApp app
    Exit Function

E01_Error:
    Debug.Print "  E01 Workbook Lifecycle: FAIL (Error " & Err.Number & ": " & Err.Description & ")"
    Resume E01_Cleanup
End Function

'==============================================================================
' E02 - Cell Read/Write
' Write string, number, date (as Double serial), boolean -> read back -> verify
'==============================================================================
Public Function E02_CellReadWrite() As Boolean
    E02_CellReadWrite = False

    Dim app As Excel.Application
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim outDir As String
    Dim filePath As String
    Dim allPassed As Boolean

    On Error GoTo E02_Error

    outDir = GetOutDir()
    filePath = outDir & "E02_cellrw.xlsx"
    If Dir(filePath) <> "" Then Kill filePath

    Set app = CreateBackgroundApp()
    Set wb = app.Workbooks.Add
    Set ws = wb.Worksheets(1)

    allPassed = True

    ' Write string
    ws.Cells(1, 1).Value2 = "Hello VBA"
    ' Write number
    ws.Cells(2, 1).Value2 = 42.5
    ' Write date as Double serial (2025-01-15 = serial 45672)
    ws.Cells(3, 1).Value2 = CDbl(DateSerial(2025, 1, 15))
    ' Write boolean
    ws.Cells(4, 1).Value2 = True

    ' Read back and verify
    Dim v1 As Variant, v2 As Variant, v3 As Variant, v4 As Variant
    v1 = ws.Cells(1, 1).Value2
    v2 = ws.Cells(2, 1).Value2
    v3 = ws.Cells(3, 1).Value2
    v4 = ws.Cells(4, 1).Value2

    If CStr(v1) <> "Hello VBA" Then
        Debug.Print "    E02 string: FAIL (got " & CStr(v1) & ")"
        allPassed = False
    End If

    If CDbl(v2) <> 42.5 Then
        Debug.Print "    E02 number: FAIL (got " & CStr(v2) & ")"
        allPassed = False
    End If

    ' Date serial comparison (allow small floating point tolerance)
    Dim expectedSerial As Double
    expectedSerial = CDbl(DateSerial(2025, 1, 15))
    If Abs(CDbl(v3) - expectedSerial) > 0.001 Then
        Debug.Print "    E02 date: FAIL (got " & CStr(v3) & ", expected " & CStr(expectedSerial) & ")"
        allPassed = False
    End If

    If CBool(v4) <> True Then
        Debug.Print "    E02 boolean: FAIL (got " & CStr(v4) & ")"
        allPassed = False
    End If

    ' Save for inspection
    wb.SaveAs Filename:=filePath, FileFormat:=51
    wb.Close SaveChanges:=False
    Set wb = Nothing

    If allPassed Then
        Debug.Print "  E02 Cell Read/Write: PASS"
        E02_CellReadWrite = True
    Else
        Debug.Print "  E02 Cell Read/Write: FAIL (see details above)"
    End If

E02_Cleanup:
    CleanupApp app
    Exit Function

E02_Error:
    Debug.Print "  E02 Cell Read/Write: FAIL (Error " & Err.Number & ": " & Err.Description & ")"
    Resume E02_Cleanup
End Function

'==============================================================================
' E03 - Bulk Range Write
' Idiomatic: 2D array -> Range.Value2 = arr (one shot)
' Naive: double For loop Cells(r,c) = value
' Both produce same result; idiomatic should be significantly faster
'==============================================================================
Public Function E03_BulkRangeWrite() As Boolean
    E03_BulkRangeWrite = False

    Dim app As Excel.Application
    Dim wb As Workbook
    Dim wsIdiomatic As Worksheet
    Dim wsNaive As Worksheet
    Dim outDir As String
    Dim filePath As String

    Const ROWS As Long = 1000
    Const COLS As Long = 10

    On Error GoTo E03_Error

    outDir = GetOutDir()
    filePath = outDir & "E03_bulk.xlsx"
    If Dir(filePath) <> "" Then Kill filePath

    Set app = CreateBackgroundApp()
    Set wb = app.Workbooks.Add
    app.Calculation = xlCalculationManual

    ' --- Idiomatic version: single array assignment ---
    Set wsIdiomatic = wb.Worksheets(1)
    wsIdiomatic.Name = "Idiomatic"

    Dim arr() As Variant
    ReDim arr(1 To ROWS, 1 To COLS)

    Dim r As Long, c As Long
    For r = 1 To ROWS
        For c = 1 To COLS
            arr(r, c) = (r - 1) * COLS + c
        Next c
    Next r

    Dim tStart As Double
    tStart = Timer

    wsIdiomatic.Range( _
        wsIdiomatic.Cells(1, 1), _
        wsIdiomatic.Cells(ROWS, COLS) _
    ).Value2 = arr

    Dim tIdiomatic As Double
    tIdiomatic = Timer - tStart

    ' --- Naive version: cell-by-cell ---
    Set wsNaive = wb.Worksheets.Add(After:=wsIdiomatic)
    wsNaive.Name = "Naive"

    tStart = Timer

    For r = 1 To ROWS
        For c = 1 To COLS
            wsNaive.Cells(r, c).Value2 = (r - 1) * COLS + c
        Next c
    Next r

    Dim tNaive As Double
    tNaive = Timer - tStart

    ' --- Verify: spot-check a few cells ---
    Dim allPassed As Boolean
    allPassed = True

    ' Check corners and middle
    If CLng(wsIdiomatic.Cells(1, 1).Value2) <> 1 Then allPassed = False
    If CLng(wsIdiomatic.Cells(1, 10).Value2) <> 10 Then allPassed = False
    If CLng(wsIdiomatic.Cells(500, 5).Value2) <> 4995 Then allPassed = False
    If CLng(wsIdiomatic.Cells(1000, 10).Value2) <> 10000 Then allPassed = False

    ' Compare idiomatic vs naive
    If CLng(wsNaive.Cells(1, 1).Value2) <> CLng(wsIdiomatic.Cells(1, 1).Value2) Then allPassed = False
    If CLng(wsNaive.Cells(500, 5).Value2) <> CLng(wsIdiomatic.Cells(500, 5).Value2) Then allPassed = False
    If CLng(wsNaive.Cells(1000, 10).Value2) <> CLng(wsIdiomatic.Cells(1000, 10).Value2) Then allPassed = False

    ' Restore calculation
    app.Calculation = xlCalculationAutomatic

    wb.SaveAs Filename:=filePath, FileFormat:=51
    wb.Close SaveChanges:=False
    Set wb = Nothing

    If allPassed Then
        Debug.Print "  E03 Bulk Range Write: PASS"
        E03_BulkRangeWrite = True
    Else
        Debug.Print "  E03 Bulk Range Write: FAIL (verification mismatch)"
    End If
    Debug.Print "    Idiomatic (array):  " & Format(tIdiomatic, "0.0000") & " s"
    Debug.Print "    Naive (cell loop):  " & Format(tNaive, "0.0000") & " s"
    Debug.Print "    Speedup:            " & Format(tNaive / IIf(tIdiomatic > 0, tIdiomatic, 0.0001), "0.0") & "x"

E03_Cleanup:
    CleanupApp app
    Exit Function

E03_Error:
    Debug.Print "  E03 Bulk Range Write: FAIL (Error " & Err.Number & ": " & Err.Description & ")"
    Resume E03_Cleanup
End Function

'==============================================================================
' E04 - Formula & Recalc
' Write formulas -> set Manual calc -> Calculate -> read Value2 -> verify
'==============================================================================
Public Function E04_FormulaRecalc() As Boolean
    E04_FormulaRecalc = False

    Dim app As Excel.Application
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim outDir As String
    Dim filePath As String

    On Error GoTo E04_Error

    outDir = GetOutDir()
    filePath = outDir & "E04_formula.xlsx"
    If Dir(filePath) <> "" Then Kill filePath

    Set app = CreateBackgroundApp()
    Set wb = app.Workbooks.Add
    Set ws = wb.Worksheets(1)

    ' Set manual calculation
    app.Calculation = xlCalculationManual

    ' Write source data
    Dim i As Long
    For i = 1 To 10
        ws.Cells(i, 1).Value2 = i         ' Column A: 1..10
        ws.Cells(i, 2).Value2 = i * 10    ' Column B: 10..100
    Next i

    ' Write formulas in column C: =A1*B1
    For i = 1 To 10
        ws.Cells(i, 3).Formula = "=A" & i & "*B" & i
    Next i

    ' Write a SUM formula
    ws.Cells(11, 3).Formula = "=SUM(C1:C10)"

    ' Force recalculation
    app.Calculate

    ' Read back and verify
    Dim allPassed As Boolean
    allPassed = True

    For i = 1 To 10
        Dim expected As Double
        expected = CDbl(i) * CDbl(i * 10)   ' i * (i*10) = i^2 * 10
        Dim actual As Double
        actual = CDbl(ws.Cells(i, 3).Value2)
        If Abs(actual - expected) > 0.001 Then
            Debug.Print "    E04 row " & i & ": FAIL (expected " & expected & ", got " & actual & ")"
            allPassed = False
        End If
    Next i

    ' Verify SUM: sum of i^2*10 for i=1..10 = 10*(1+4+9+16+25+36+49+64+81+100) = 10*385 = 3850
    Dim sumExpected As Double
    sumExpected = 3850
    Dim sumActual As Double
    sumActual = CDbl(ws.Cells(11, 3).Value2)
    If Abs(sumActual - sumExpected) > 0.001 Then
        Debug.Print "    E04 SUM: FAIL (expected " & sumExpected & ", got " & sumActual & ")"
        allPassed = False
    End If

    ' Restore automatic calculation
    app.Calculation = xlCalculationAutomatic

    wb.SaveAs Filename:=filePath, FileFormat:=51
    wb.Close SaveChanges:=False
    Set wb = Nothing

    If allPassed Then
        Debug.Print "  E04 Formula & Recalc: PASS"
        E04_FormulaRecalc = True
    Else
        Debug.Print "  E04 Formula & Recalc: FAIL (see details above)"
    End If

E04_Cleanup:
    CleanupApp app
    Exit Function

E04_Error:
    Debug.Print "  E04 Formula & Recalc: FAIL (Error " & Err.Number & ": " & Err.Description & ")"
    Resume E04_Cleanup
End Function

'==============================================================================
' E05 - Cell Formatting
' Font.Bold, Font.Size, Font.Color, Interior.Color, NumberFormat,
' HorizontalAlignment, Merge
'==============================================================================
Public Function E05_CellFormatting() As Boolean
    E05_CellFormatting = False

    Dim app As Excel.Application
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim outDir As String
    Dim filePath As String

    On Error GoTo E05_Error

    outDir = GetOutDir()
    filePath = outDir & "E05_format.xlsx"
    If Dir(filePath) <> "" Then Kill filePath

    Set app = CreateBackgroundApp()
    Set wb = app.Workbooks.Add
    Set ws = wb.Worksheets(1)

    ' --- Apply formatting ---

    ' A1: Bold header
    ws.Cells(1, 1).Value2 = "Bold Header"
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.Size = 16
    ws.Cells(1, 1).Font.Color = RGB(0, 0, 255)     ' Blue font

    ' B1: Colored background
    ws.Cells(1, 2).Value2 = "Yellow BG"
    ws.Cells(1, 2).Interior.Color = RGB(255, 255, 0)  ' Yellow fill

    ' A2: Number format
    ws.Cells(2, 1).Value2 = 1234.5678
    ws.Cells(2, 1).NumberFormat = "0.00"

    ' B2: Date format
    ws.Cells(2, 2).Value2 = CDbl(DateSerial(2025, 6, 15))
    ws.Cells(2, 2).NumberFormat = "yyyy-mm-dd"

    ' A3: Center alignment
    ws.Cells(3, 1).Value2 = "Centered"
    ws.Cells(3, 1).HorizontalAlignment = xlCenter  ' -4108

    ' A4:C4: Merged cells
    ws.Range("A4:C4").Merge
    ws.Range("A4").Value2 = "Merged Region"
    ws.Range("A4").HorizontalAlignment = xlCenter
    ws.Range("A4").Font.Bold = True

    ' Save
    wb.SaveAs Filename:=filePath, FileFormat:=51
    wb.Close SaveChanges:=False
    Set wb = Nothing

    ' --- Verify by reopening ---
    Dim allPassed As Boolean
    allPassed = True

    Set wb = app.Workbooks.Open(filePath)
    Set ws = wb.Worksheets(1)

    ' Check bold
    If ws.Cells(1, 1).Font.Bold <> True Then
        Debug.Print "    E05 Bold: FAIL"
        allPassed = False
    End If

    ' Check font size
    If ws.Cells(1, 1).Font.Size <> 16 Then
        Debug.Print "    E05 Font.Size: FAIL (got " & ws.Cells(1, 1).Font.Size & ")"
        allPassed = False
    End If

    ' Check font color (blue = RGB(0,0,255))
    If ws.Cells(1, 1).Font.Color <> RGB(0, 0, 255) Then
        Debug.Print "    E05 Font.Color: FAIL (got " & ws.Cells(1, 1).Font.Color & ")"
        allPassed = False
    End If

    ' Check interior color (yellow = RGB(255,255,0))
    If ws.Cells(1, 2).Interior.Color <> RGB(255, 255, 0) Then
        Debug.Print "    E05 Interior.Color: FAIL (got " & ws.Cells(1, 2).Interior.Color & ")"
        allPassed = False
    End If

    ' Check number format
    If ws.Cells(2, 1).NumberFormat <> "0.00" Then
        Debug.Print "    E05 NumberFormat: FAIL (got " & ws.Cells(2, 1).NumberFormat & ")"
        allPassed = False
    End If

    ' Check center alignment
    If ws.Cells(3, 1).HorizontalAlignment <> xlCenter Then
        Debug.Print "    E05 Alignment: FAIL"
        allPassed = False
    End If

    ' Check merge (A4 should be merged)
    If ws.Range("A4").MergeCells <> True Then
        Debug.Print "    E05 Merge: FAIL"
        allPassed = False
    End If

    wb.Close SaveChanges:=False
    Set wb = Nothing

    If allPassed Then
        Debug.Print "  E05 Cell Formatting: PASS"
        E05_CellFormatting = True
    Else
        Debug.Print "  E05 Cell Formatting: FAIL (see details above)"
    End If

E05_Cleanup:
    CleanupApp app
    Exit Function

E05_Error:
    Debug.Print "  E05 Cell Formatting: FAIL (Error " & Err.Number & ": " & Err.Description & ")"
    Resume E05_Cleanup
End Function

'==============================================================================
' E06 - Row/Column Structure
' Insert/delete rows and columns, AutoFit, set RowHeight/ColumnWidth
'==============================================================================
Public Function E06_RowColumnStructure() As Boolean
    E06_RowColumnStructure = False

    Dim app As Excel.Application
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim outDir As String
    Dim filePath As String

    On Error GoTo E06_Error

    outDir = GetOutDir()
    filePath = outDir & "E06_structure.xlsx"
    If Dir(filePath) <> "" Then Kill filePath

    Set app = CreateBackgroundApp()
    Set wb = app.Workbooks.Add
    Set ws = wb.Worksheets(1)

    ' Write initial data (5 rows x 3 cols)
    Dim r As Long, c As Long
    For r = 1 To 5
        ws.Cells(r, 1).Value2 = "Row" & r
        ws.Cells(r, 2).Value2 = r * 100
        ws.Cells(r, 3).Value2 = "Some longer text for column " & r
    Next r

    ' Insert a row at position 3 (pushes existing row 3+ down)
    ws.Rows(3).Insert
    ws.Cells(3, 1).Value2 = "Inserted Row"
    ws.Cells(3, 2).Value2 = 999
    ws.Cells(3, 3).Value2 = "This row was inserted"

    ' Insert a column at position 2 (pushes existing col 2+ right)
    ws.Columns(2).Insert
    ws.Cells(1, 2).Value2 = "New Col"

    ' Delete row 6 (was originally row 5 before insert)
    ws.Rows(6).Delete

    ' Delete column 4 (originally col 3 = "Some longer text" column, now shifted)
    ws.Columns(4).Delete

    ' AutoFit remaining columns
    ws.Columns("A:C").AutoFit

    ' Set specific row height and column width
    ws.Rows(1).RowHeight = 30
    ws.Columns("A").ColumnWidth = 20

    ' Save
    wb.SaveAs Filename:=filePath, FileFormat:=51

    ' --- Verify ---
    Dim allPassed As Boolean
    allPassed = True

    ' Check inserted row content
    If CStr(ws.Cells(3, 1).Value2) <> "Inserted Row" Then
        Debug.Print "    E06 Inserted row: FAIL (got " & CStr(ws.Cells(3, 1).Value2) & ")"
        allPassed = False
    End If

    ' Check row height
    If Abs(ws.Rows(1).RowHeight - 30) > 0.5 Then
        Debug.Print "    E06 RowHeight: FAIL (got " & ws.Rows(1).RowHeight & ")"
        allPassed = False
    End If

    ' Check column width
    If Abs(ws.Columns("A").ColumnWidth - 20) > 0.5 Then
        Debug.Print "    E06 ColumnWidth: FAIL (got " & ws.Columns("A").ColumnWidth & ")"
        allPassed = False
    End If

    ' Check new column header
    If CStr(ws.Cells(1, 2).Value2) <> "New Col" Then
        Debug.Print "    E06 Inserted col: FAIL"
        allPassed = False
    End If

    wb.Close SaveChanges:=False
    Set wb = Nothing

    If allPassed Then
        Debug.Print "  E06 Row/Column Structure: PASS"
        E06_RowColumnStructure = True
    Else
        Debug.Print "  E06 Row/Column Structure: FAIL (see details above)"
    End If

E06_Cleanup:
    CleanupApp app
    Exit Function

E06_Error:
    Debug.Print "  E06 Row/Column Structure: FAIL (Error " & Err.Number & ": " & Err.Description & ")"
    Resume E06_Cleanup
End Function

'==============================================================================
' E07 - Multi-Worksheet
' Add worksheets, rename, delete one, write cross-sheet formula
'==============================================================================
Public Function E07_MultiWorksheet() As Boolean
    E07_MultiWorksheet = False

    Dim app As Excel.Application
    Dim wb As Workbook
    Dim wsData As Worksheet
    Dim wsSummary As Worksheet
    Dim wsTemp As Worksheet
    Dim outDir As String
    Dim filePath As String

    On Error GoTo E07_Error

    outDir = GetOutDir()
    filePath = outDir & "E07_multisheet.xlsx"
    If Dir(filePath) <> "" Then Kill filePath

    Set app = CreateBackgroundApp()
    Set wb = app.Workbooks.Add

    ' Rename default sheet
    wb.Worksheets(1).Name = "Data"
    Set wsData = wb.Worksheets("Data")

    ' Add Summary sheet
    Set wsSummary = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    wsSummary.Name = "Summary"

    ' Add a temporary sheet (to be deleted)
    Set wsTemp = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    wsTemp.Name = "TempSheet"

    ' Write data to Data sheet
    wsData.Cells(1, 1).Value2 = "Product"
    wsData.Cells(1, 2).Value2 = "Revenue"
    Dim i As Long
    For i = 2 To 6
        wsData.Cells(i, 1).Value2 = "Product " & (i - 1)
        wsData.Cells(i, 2).Value2 = (i - 1) * 1000
    Next i

    ' Write cross-sheet formula in Summary
    wsSummary.Cells(1, 1).Value2 = "Total Revenue"
    wsSummary.Cells(1, 2).Formula = "=SUM('Data'!B2:B6)"

    ' Also reference a single cell
    wsSummary.Cells(2, 1).Value2 = "First Product"
    wsSummary.Cells(2, 2).Formula = "='Data'!A2"

    ' Delete the temporary sheet
    app.DisplayAlerts = False
    wsTemp.Delete
    Set wsTemp = Nothing

    ' Force recalculation
    app.Calculate

    ' --- Verify ---
    Dim allPassed As Boolean
    allPassed = True

    ' Check sheet count (should be 2: Data, Summary)
    If wb.Worksheets.Count <> 2 Then
        Debug.Print "    E07 Sheet count: FAIL (got " & wb.Worksheets.Count & ", expected 2)"
        allPassed = False
    End If

    ' Check cross-sheet SUM formula result: 1000+2000+3000+4000+5000 = 15000
    Dim sumVal As Double
    sumVal = CDbl(wsSummary.Cells(1, 2).Value2)
    If Abs(sumVal - 15000) > 0.001 Then
        Debug.Print "    E07 Cross-sheet SUM: FAIL (got " & sumVal & ", expected 15000)"
        allPassed = False
    End If

    ' Check cross-sheet single cell reference
    Dim refVal As String
    refVal = CStr(wsSummary.Cells(2, 2).Value2)
    If refVal <> "Product 1" Then
        Debug.Print "    E07 Cross-sheet ref: FAIL (got " & refVal & ")"
        allPassed = False
    End If

    ' Check sheet names
    If wb.Worksheets(1).Name <> "Data" Then
        Debug.Print "    E07 Sheet name 1: FAIL (got " & wb.Worksheets(1).Name & ")"
        allPassed = False
    End If
    If wb.Worksheets(2).Name <> "Summary" Then
        Debug.Print "    E07 Sheet name 2: FAIL (got " & wb.Worksheets(2).Name & ")"
        allPassed = False
    End If

    wb.SaveAs Filename:=filePath, FileFormat:=51
    wb.Close SaveChanges:=False
    Set wb = Nothing

    If allPassed Then
        Debug.Print "  E07 Multi-Worksheet: PASS"
        E07_MultiWorksheet = True
    Else
        Debug.Print "  E07 Multi-Worksheet: FAIL (see details above)"
    End If

E07_Cleanup:
    CleanupApp app
    Exit Function

E07_Error:
    Debug.Print "  E07 Multi-Worksheet: FAIL (Error " & Err.Number & ": " & Err.Description & ")"
    Resume E07_Cleanup
End Function

'==============================================================================
' E08 - Data Operations: Find/Replace, Sort, AutoFilter
'==============================================================================
Public Function E08_DataOperations() As Boolean
    E08_DataOperations = False

    Dim app As Excel.Application
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim outDir As String
    Dim filePath As String

    On Error GoTo E08_Error

    outDir = GetOutDir()
    filePath = outDir & "E08_dataops.xlsx"
    If Dir(filePath) <> "" Then Kill filePath

    Set app = CreateBackgroundApp()
    Set wb = app.Workbooks.Add
    Set ws = wb.Worksheets(1)

    ' --- Create sample data (25 rows with headers) ---
    ws.Cells(1, 1).Value2 = "Name"
    ws.Cells(1, 2).Value2 = "Category"
    ws.Cells(1, 3).Value2 = "Amount"
    ws.Cells(1, 4).Value2 = "Region"

    Dim categories As Variant
    categories = Array("Electronics", "Clothing", "Food", "Electronics", "Clothing", _
                       "Food", "Electronics", "Clothing", "Food", "Electronics", _
                       "Clothing", "Food", "Electronics", "Clothing", "Food", _
                       "Electronics", "Clothing", "Food", "Electronics", "Clothing", _
                       "Food", "Electronics", "Clothing", "Food")

    Dim regions As Variant
    regions = Array("North", "South", "East", "West", "North", _
                    "South", "East", "West", "North", "South", _
                    "East", "West", "North", "South", "East", _
                    "West", "North", "South", "East", "West", _
                    "North", "South", "East", "West")

    Dim r As Long
    For r = 2 To 25
        ws.Cells(r, 1).Value2 = "Item " & (r - 1)
        ws.Cells(r, 2).Value2 = categories(r - 2)
        ws.Cells(r, 3).Value2 = (r - 1) * 50 + 100
        ws.Cells(r, 4).Value2 = regions(r - 2)
    Next r

    Dim allPassed As Boolean
    allPassed = True

    ' --- Find/Replace: Replace "Electronics" with "Tech" ---
    Dim dataRange As Range
    Set dataRange = ws.Range("A1:D25")

    ' Count occurrences before replace (for verification)
    Dim replaceCount As Long
    replaceCount = 0
    For r = 2 To 25
        If CStr(ws.Cells(r, 2).Value2) = "Electronics" Then
            replaceCount = replaceCount + 1
        End If
    Next r

    dataRange.Replace What:="Electronics", Replacement:="Tech", LookAt:=xlPart

    ' Verify replacement
    Dim techCount As Long
    techCount = 0
    For r = 2 To 25
        If CStr(ws.Cells(r, 2).Value2) = "Tech" Then
            techCount = techCount + 1
        End If
    Next r

    If techCount <> replaceCount Then
        Debug.Print "    E08 Replace: FAIL (replaced " & techCount & ", expected " & replaceCount & ")"
        allPassed = False
    Else
        Debug.Print "    E08 Replace: OK (" & techCount & " replacements)"
    End If

    ' --- Sort by Amount (column 3) ascending ---
    dataRange.Sort Key1:=ws.Range("C1"), Order1:=xlAscending, Header:=xlYes

    ' Verify sort: Amount in row 2 should be smallest
    Dim prevVal As Double
    Dim sortOk As Boolean
    sortOk = True
    prevVal = CDbl(ws.Cells(2, 3).Value2)
    For r = 3 To 25
        Dim curVal As Double
        curVal = CDbl(ws.Cells(r, 3).Value2)
        If curVal < prevVal Then
            sortOk = False
            Exit For
        End If
        prevVal = curVal
    Next r

    If Not sortOk Then
        Debug.Print "    E08 Sort: FAIL (not in ascending order)"
        allPassed = False
    Else
        Debug.Print "    E08 Sort: OK (ascending by Amount)"
    End If

    ' --- AutoFilter: filter Category = "Tech" ---
    ws.Range("A1:D25").AutoFilter Field:=2, Criteria1:="Tech"

    ' Verify: count visible rows (exclude header)
    Dim visibleCount As Long
    visibleCount = 0
    For r = 2 To 25
        If ws.Rows(r).Hidden = False Then
            visibleCount = visibleCount + 1
        End If
    Next r

    If visibleCount <> techCount Then
        Debug.Print "    E08 AutoFilter: FAIL (visible " & visibleCount & ", expected " & techCount & ")"
        allPassed = False
    Else
        Debug.Print "    E08 AutoFilter: OK (" & visibleCount & " visible rows)"
    End If

    ' Remove AutoFilter
    If ws.AutoFilterMode Then ws.AutoFilterMode = False

    wb.SaveAs Filename:=filePath, FileFormat:=51
    wb.Close SaveChanges:=False
    Set wb = Nothing

    If allPassed Then
        Debug.Print "  E08 Data Operations: PASS"
        E08_DataOperations = True
    Else
        Debug.Print "  E08 Data Operations: FAIL (see details above)"
    End If

E08_Cleanup:
    CleanupApp app
    Exit Function

E08_Error:
    Debug.Print "  E08 Data Operations: FAIL (Error " & Err.Number & ": " & Err.Description & ")"
    Resume E08_Cleanup
End Function

'==============================================================================
' E09 - Chart Generation
' Create clustered column chart from data range using ChartObjects.Add
'==============================================================================
Public Function E09_ChartGeneration() As Boolean
    E09_ChartGeneration = False

    Dim app As Excel.Application
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim outDir As String
    Dim filePath As String

    On Error GoTo E09_Error

    outDir = GetOutDir()
    filePath = outDir & "E09_chart.xlsx"
    If Dir(filePath) <> "" Then Kill filePath

    Set app = CreateBackgroundApp()
    Set wb = app.Workbooks.Add
    Set ws = wb.Worksheets(1)

    ' Create source data for the chart
    ws.Cells(1, 1).Value2 = "Quarter"
    ws.Cells(1, 2).Value2 = "Revenue"
    ws.Cells(1, 3).Value2 = "Profit"

    ws.Cells(2, 1).Value2 = "Q1"
    ws.Cells(2, 2).Value2 = 15000
    ws.Cells(2, 3).Value2 = 3000

    ws.Cells(3, 1).Value2 = "Q2"
    ws.Cells(3, 2).Value2 = 22000
    ws.Cells(3, 3).Value2 = 5500

    ws.Cells(4, 1).Value2 = "Q3"
    ws.Cells(4, 2).Value2 = 18000
    ws.Cells(4, 3).Value2 = 4200

    ws.Cells(5, 1).Value2 = "Q4"
    ws.Cells(5, 2).Value2 = 25000
    ws.Cells(5, 3).Value2 = 7000

    ' Create chart: ChartObjects.Add(Left, Top, Width, Height)
    Dim chartObj As ChartObject
    Set chartObj = ws.ChartObjects.Add(Left:=100, Top:=80, Width:=400, Height:=300)

    ' Configure chart
    With chartObj.Chart
        .SetSourceData Source:=ws.Range("A1:C5")
        .ChartType = xlColumnClustered  ' = 51
        .HasTitle = True
        .ChartTitle.Text = "Quarterly Results"
    End With

    ' Save
    wb.SaveAs Filename:=filePath, FileFormat:=51

    ' --- Verify ---
    Dim allPassed As Boolean
    allPassed = True

    ' Check chart exists
    If ws.ChartObjects.Count < 1 Then
        Debug.Print "    E09 Chart count: FAIL (no charts found)"
        allPassed = False
    Else
        ' Check chart type
        If ws.ChartObjects(1).Chart.ChartType <> xlColumnClustered Then
            Debug.Print "    E09 ChartType: FAIL (got " & ws.ChartObjects(1).Chart.ChartType & ")"
            allPassed = False
        End If

        ' Check chart title
        If ws.ChartObjects(1).Chart.ChartTitle.Text <> "Quarterly Results" Then
            Debug.Print "    E09 ChartTitle: FAIL"
            allPassed = False
        End If
    End If

    wb.Close SaveChanges:=False
    Set wb = Nothing

    If allPassed Then
        Debug.Print "  E09 Chart Generation: PASS"
        E09_ChartGeneration = True
    Else
        Debug.Print "  E09 Chart Generation: FAIL (see details above)"
    End If

E09_Cleanup:
    CleanupApp app
    Exit Function

E09_Error:
    Debug.Print "  E09 Chart Generation: FAIL (Error " & Err.Number & ": " & Err.Description & ")"
    Resume E09_Cleanup
End Function

'==============================================================================
' E10 - Export PDF
' Workbook.ExportAsFixedFormat xlTypePDF(=0), pdfPath
'==============================================================================
Public Function E10_ExportPDF() As Boolean
    E10_ExportPDF = False

    Dim app As Excel.Application
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim outDir As String
    Dim xlPath As String
    Dim pdfPath As String

    On Error GoTo E10_Error

    outDir = GetOutDir()
    xlPath = outDir & "E10_source.xlsx"
    pdfPath = outDir & "E10_export.pdf"
    If Dir(xlPath) <> "" Then Kill xlPath
    If Dir(pdfPath) <> "" Then Kill pdfPath

    Set app = CreateBackgroundApp()
    Set wb = app.Workbooks.Add
    Set ws = wb.Worksheets(1)

    ' Create content to export
    ws.Cells(1, 1).Value2 = "PDF Export Test"
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.Size = 18

    ws.Cells(3, 1).Value2 = "Name"
    ws.Cells(3, 2).Value2 = "Value"
    ws.Cells(3, 1).Font.Bold = True
    ws.Cells(3, 2).Font.Bold = True

    Dim i As Long
    For i = 1 To 10
        ws.Cells(3 + i, 1).Value2 = "Item " & i
        ws.Cells(3 + i, 2).Value2 = i * 123.45
        ws.Cells(3 + i, 2).NumberFormat = "#,##0.00"
    Next i

    ws.Columns("A:B").AutoFit

    ' Save as xlsx first (required for some export scenarios)
    wb.SaveAs Filename:=xlPath, FileFormat:=51

    ' Export as PDF (xlTypePDF = 0)
    wb.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, _
        IgnorePrintAreas:=False

    wb.Close SaveChanges:=False
    Set wb = Nothing

    ' --- Verify: check PDF file exists and is non-empty ---
    Dim allPassed As Boolean
    allPassed = True

    If Dir(pdfPath) = "" Then
        Debug.Print "    E10 PDF file: FAIL (file not found)"
        allPassed = False
    Else
        ' Check file size > 0
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        Dim pdfSize As Long
        pdfSize = fso.GetFile(pdfPath).Size
        Set fso = Nothing

        If pdfSize = 0 Then
            Debug.Print "    E10 PDF size: FAIL (empty file)"
            allPassed = False
        Else
            Debug.Print "    E10 PDF size: " & pdfSize & " bytes"
        End If
    End If

    If allPassed Then
        Debug.Print "  E10 Export PDF: PASS"
        E10_ExportPDF = True
    Else
        Debug.Print "  E10 Export PDF: FAIL (see details above)"
    End If

E10_Cleanup:
    CleanupApp app
    Exit Function

E10_Error:
    Debug.Print "  E10 Export PDF: FAIL (Error " & Err.Number & ": " & Err.Description & ")"
    Resume E10_Cleanup
End Function

'==============================================================================
' E11 - Run VBA Macro
' Define a macro (Sub) that modifies cells, then call it via Application.Run
' VBA is the native macro host -- this is the most natural case.
'==============================================================================
Public Function E11_RunVBAMacro() As Boolean
    E11_RunVBAMacro = False

    Dim app As Excel.Application
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim outDir As String
    Dim filePath As String

    On Error GoTo E11_Error

    outDir = GetOutDir()
    filePath = outDir & "E11_macro.xlsm"
    If Dir(filePath) <> "" Then Kill filePath

    Set app = CreateBackgroundApp()
    Set wb = app.Workbooks.Add
    Set ws = wb.Worksheets(1)

    ' Write initial data
    ws.Cells(1, 1).Value2 = "Before Macro"
    ws.Cells(2, 1).Value2 = 100
    ws.Cells(3, 1).Value2 = 200

    ' Inject VBA code into the workbook's VBA project
    ' This requires "Trust access to the VBA project object model" to be enabled
    Dim vbProj As Object
    Set vbProj = wb.VBProject

    Dim vbComp As Object
    Set vbComp = vbProj.VBComponents.Add(1) ' 1 = vbext_ct_StdModule
    vbComp.Name = "modInjected"

    Dim macroCode As String
    macroCode = "Sub FillWithColors()" & vbCrLf & _
                "    Dim ws As Worksheet" & vbCrLf & _
                "    Set ws = ThisWorkbook.Worksheets(1)" & vbCrLf & _
                "    ws.Cells(1, 1).Value2 = ""Macro Executed""" & vbCrLf & _
                "    ws.Cells(2, 1).Value2 = ws.Cells(2, 1).Value2 * 2" & vbCrLf & _
                "    ws.Cells(3, 1).Value2 = ws.Cells(3, 1).Value2 * 3" & vbCrLf & _
                "    ws.Range(""A1:A3"").Interior.Color = RGB(144, 238, 144)" & vbCrLf & _
                "    ws.Range(""A1:A3"").Font.Bold = True" & vbCrLf & _
                "End Sub"

    vbComp.CodeModule.AddFromString macroCode

    ' Run the injected macro
    app.Run wb.Name & "!modInjected.FillWithColors"

    ' --- Verify ---
    Dim allPassed As Boolean
    allPassed = True

    If CStr(ws.Cells(1, 1).Value2) <> "Macro Executed" Then
        Debug.Print "    E11 Cell A1: FAIL (got " & CStr(ws.Cells(1, 1).Value2) & ")"
        allPassed = False
    End If

    If CDbl(ws.Cells(2, 1).Value2) <> 200 Then  ' 100 * 2
        Debug.Print "    E11 Cell A2: FAIL (got " & CStr(ws.Cells(2, 1).Value2) & ", expected 200)"
        allPassed = False
    End If

    If CDbl(ws.Cells(3, 1).Value2) <> 600 Then  ' 200 * 3
        Debug.Print "    E11 Cell A3: FAIL (got " & CStr(ws.Cells(3, 1).Value2) & ", expected 600)"
        allPassed = False
    End If

    If ws.Range("A1").Interior.Color <> RGB(144, 238, 144) Then
        Debug.Print "    E11 Fill color: FAIL"
        allPassed = False
    End If

    If ws.Range("A1").Font.Bold <> True Then
        Debug.Print "    E11 Bold: FAIL"
        allPassed = False
    End If

    ' Save as .xlsm (macro-enabled = 52)
    wb.SaveAs Filename:=filePath, FileFormat:=52
    wb.Close SaveChanges:=False
    Set wb = Nothing

    If allPassed Then
        Debug.Print "  E11 Run VBA Macro: PASS"
        E11_RunVBAMacro = True
    Else
        Debug.Print "  E11 Run VBA Macro: FAIL (see details above)"
    End If

E11_Cleanup:
    CleanupApp app
    Exit Function

E11_Error:
    Debug.Print "  E11 Run VBA Macro: FAIL (Error " & Err.Number & ": " & Err.Description & ")"
    If Err.Number = 1004 Or Err.Number = -2147188160# Then
        Debug.Print "    NOTE: Enable 'Trust access to VBA project object model' in"
        Debug.Print "          File > Options > Trust Center > Trust Center Settings > Macro Settings"
    End If
    Resume E11_Cleanup
End Function

'==============================================================================
' E12 - Resource Cleanup
' Clean exit with no residual EXCEL.EXE
' Close workbooks, Quit app, Set objects = Nothing
'==============================================================================
Public Function E12_ResourceCleanup() As Boolean
    E12_ResourceCleanup = False

    Dim app As Excel.Application
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim outDir As String
    Dim filePath As String

    On Error GoTo E12_Error

    outDir = GetOutDir()
    filePath = outDir & "E12_cleanup.xlsx"
    If Dir(filePath) <> "" Then Kill filePath

    ' Create a new Excel.Application instance
    Set app = New Excel.Application
    app.Visible = False
    app.DisplayAlerts = False
    app.ScreenUpdating = False

    Debug.Print "    E12 New Excel.Application created (PID via hwnd)"

    ' Do some work
    Set wb = app.Workbooks.Add
    Set ws = wb.Worksheets(1)
    ws.Cells(1, 1).Value2 = "Cleanup Test"
    ws.Cells(2, 1).Value2 = Now()

    ' Add a second workbook
    Dim wb2 As Workbook
    Set wb2 = app.Workbooks.Add
    wb2.Worksheets(1).Cells(1, 1).Value2 = "Second workbook"

    ' Save first workbook
    wb.SaveAs Filename:=filePath, FileFormat:=51

    ' --- Proper cleanup sequence ---
    ' 1. Release worksheet references
    Set ws = Nothing

    ' 2. Close all workbooks without saving
    wb2.Close SaveChanges:=False
    Set wb2 = Nothing

    wb.Close SaveChanges:=False
    Set wb = Nothing

    ' 3. Quit the application
    app.Quit

    ' 4. Release the application object
    Set app = Nothing

    ' --- Verify: file exists on disk ---
    Dim allPassed As Boolean
    allPassed = True

    If Dir(filePath) = "" Then
        Debug.Print "    E12 Output file: FAIL (not found)"
        allPassed = False
    End If

    ' Note: In VBA, we cannot easily verify no residual EXCEL.EXE because
    ' we're running inside Excel ourselves. The cleanup pattern is:
    '   wb.Close False -> app.Quit -> Set app = Nothing
    ' VBA has no ReleaseComObject; setting to Nothing + engine GC handles it.

    If allPassed Then
        Debug.Print "  E12 Resource Cleanup: PASS"
        Debug.Print "    Cleanup sequence: wb.Close(False) -> app.Quit -> Set app = Nothing"
        Debug.Print "    VBA relies on Set = Nothing + engine GC (no ReleaseComObject)"
        E12_ResourceCleanup = True
    Else
        Debug.Print "  E12 Resource Cleanup: FAIL (see details above)"
    End If

    Exit Function

E12_Error:
    Debug.Print "  E12 Resource Cleanup: FAIL (Error " & Err.Number & ": " & Err.Description & ")"
    ' Emergency cleanup
    On Error Resume Next
    If Not wb2 Is Nothing Then wb2.Close False
    If Not wb Is Nothing Then wb.Close False
    If Not app Is Nothing Then app.Quit
    Set ws = Nothing
    Set wb2 = Nothing
    Set wb = Nothing
    Set app = Nothing
    On Error GoTo 0
End Function
