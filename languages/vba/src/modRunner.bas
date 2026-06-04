Attribute VB_Name = "modRunner"
'==============================================================================
' modRunner.bas
' Entry point with Timer-based timing for E01-E12
'
' Usage:
'   - RunAll: runs all 12 tasks sequentially with timing
'   - RunSingle "E01": runs a single task by name
'
' Output goes to the Immediate window (Ctrl+G in VBE).
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' RunAll - Execute all tasks E01-E12 with timing
'------------------------------------------------------------------------------
Public Sub RunAll()
    Dim totalStart As Double
    totalStart = Timer

    Debug.Print "============================================================"
    Debug.Print " VBA Excel COM Automation - E01..E12"
    Debug.Print " Started: " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    Debug.Print "============================================================"
    Debug.Print ""

    RunSingle "E01"
    RunSingle "E02"
    RunSingle "E03"
    RunSingle "E04"
    RunSingle "E05"
    RunSingle "E06"
    RunSingle "E07"
    RunSingle "E08"
    RunSingle "E09"
    RunSingle "E10"
    RunSingle "E11"
    RunSingle "E12"

    Dim totalElapsed As Double
    totalElapsed = Timer - totalStart

    Debug.Print ""
    Debug.Print "============================================================"
    Debug.Print " All tasks completed"
    Debug.Print " Total elapsed: " & Format(totalElapsed, "0.000") & " s"
    Debug.Print " Finished: " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    Debug.Print "============================================================"
End Sub

'------------------------------------------------------------------------------
' RunSingle - Execute a single task by name with timing
'------------------------------------------------------------------------------
Public Sub RunSingle(ByVal taskName As String)
    Dim tStart As Double
    Dim tElapsed As Double

    taskName = UCase(Trim(taskName))
    tStart = Timer

    Select Case taskName
        Case "E01"
            Debug.Print "[E01] Workbook Lifecycle..."
            E01_WorkbookLifecycle

        Case "E02"
            Debug.Print "[E02] Cell Read/Write..."
            E02_CellReadWrite

        Case "E03"
            Debug.Print "[E03] Bulk Range Write..."
            E03_BulkRangeWrite

        Case "E04"
            Debug.Print "[E04] Formula & Recalc..."
            E04_FormulaRecalc

        Case "E05"
            Debug.Print "[E05] Cell Formatting..."
            E05_CellFormatting

        Case "E06"
            Debug.Print "[E06] Row/Column Structure..."
            E06_RowColumnStructure

        Case "E07"
            Debug.Print "[E07] Multi-Worksheet..."
            E07_MultiWorksheet

        Case "E08"
            Debug.Print "[E08] Data Operations..."
            E08_DataOperations

        Case "E09"
            Debug.Print "[E09] Chart Generation..."
            E09_ChartGeneration

        Case "E10"
            Debug.Print "[E10] Export PDF..."
            E10_ExportPDF

        Case "E11"
            Debug.Print "[E11] Run VBA Macro..."
            E11_RunVBAMacro

        Case "E12"
            Debug.Print "[E12] Resource Cleanup..."
            E12_ResourceCleanup

        Case Else
            Debug.Print "Unknown task: " & taskName
            Exit Sub
    End Select

    tElapsed = Timer - tStart

    ' Handle Timer wraparound at midnight
    If tElapsed < 0 Then tElapsed = tElapsed + 86400

    Debug.Print "  Time: " & Format(tElapsed, "0.000") & " s"
    Debug.Print ""
End Sub
