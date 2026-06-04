# run_vba_tests.ps1 - Run VBA E01-E12 tests with file-based logging
# Strategy:
# 1. Kill lingering Excel, create fresh VBA_Runner.xlsm
# 2. Import task modules + inject a RunAllAndLog macro that writes to a file
# 3. Run the macro, then verify output files externally
#
# The injected macro calls each E## sub, catches unhandled errors,
# and also checks that expected output files were created.

$ErrorActionPreference = "Stop"

$srcDir = "D:\Workspace\AI\office-com-automation\languages\vba\src"
# VBA GetOutDir() resolves to ThisWorkbook.Path\..\..\out\vba\ = languages\out\vba\
$outDir = "D:\Workspace\AI\office-com-automation\languages\out\vba"
$logFile = Join-Path $outDir "run_log.txt"
$xlsmPath = Join-Path $srcDir "VBA_Runner.xlsm"

# Ensure output directory exists
if (-not (Test-Path $outDir)) {
    New-Item -ItemType Directory -Path $outDir -Force | Out-Null
}

# Kill any lingering Excel
Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 3

# Remove lock file if exists
$lockFile = Join-Path $srcDir "~`$VBA_Runner.xlsm"
if (Test-Path $lockFile) { Remove-Item $lockFile -Force -ErrorAction SilentlyContinue }

# Clean old output files
Get-ChildItem $outDir -Filter "*.xlsx" -ErrorAction SilentlyContinue | Remove-Item -Force
Get-ChildItem $outDir -Filter "*.pdf" -ErrorAction SilentlyContinue | Remove-Item -Force
if (Test-Path $logFile) { Remove-Item $logFile -Force }

Write-Host "Starting VBA test run..."
Write-Host "Output dir: $outDir"

$excel = $null
$wb = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false

    # Delete existing xlsm to start fresh
    if (Test-Path $xlsmPath) { Remove-Item $xlsmPath -Force }

    # Create new workbook and save as macro-enabled
    $wb = $excel.Workbooks.Add()
    $wb.SaveAs($xlsmPath, 52) # xlOpenXMLWorkbookMacroEnabled = 52

    # Import VBA modules
    $vbProject = $wb.VBProject

    # Remove existing standard modules
    foreach ($comp in $vbProject.VBComponents) {
        if ($comp.Type -eq 1) { # vbext_ct_StdModule = 1
            try { $vbProject.VBComponents.Remove($comp) } catch {}
        }
    }

    Write-Host "Importing modules..."
    $vbProject.VBComponents.Import((Join-Path $srcDir "modExcelTasks.bas")) | Out-Null
    $vbProject.VBComponents.Import((Join-Path $srcDir "modRunner.bas")) | Out-Null

    # Build the injected macro - VBA paths use single backslashes (no escaping needed)
    $vbaOutDir = $outDir

    $logMacroCode = @"
Option Explicit

Public Sub RunAllAndLog()
    Dim logPath As String
    logPath = "$vbaOutDir\run_log.txt"

    Dim fNum As Integer
    fNum = FreeFile
    Open logPath For Output As #fNum

    Print #fNum, "VBA Excel COM Automation - E01..E12"
    Print #fNum, "Started: " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    Print #fNum, "============================================================"
    Print #fNum, ""

    Dim taskNames As Variant
    taskNames = Array("E01", "E02", "E03", "E04", "E05", "E06", "E07", "E08", "E09", "E10", "E11", "E12")

    Dim expectedFiles As Variant
    expectedFiles = Array( _
        "E01_lifecycle.xlsx", _
        "E02_cellrw.xlsx", _
        "E03_bulk.xlsx", _
        "E04_formula.xlsx", _
        "E05_format.xlsx", _
        "E06_structure.xlsx", _
        "E07_multisheet.xlsx", _
        "E08_dataops.xlsx", _
        "E09_chart.xlsx", _
        "E10_export.pdf", _
        "E11_macro.xlsm", _
        "E12_cleanup.xlsx" _
    )

    Dim passCount As Long
    Dim totalStart As Double
    totalStart = Timer
    passCount = 0

    Dim i As Long
    For i = LBound(taskNames) To UBound(taskNames)
        Dim tStart As Double
        tStart = Timer

        Print #fNum, "--- " & taskNames(i) & " ---"

        ' Call each task function and capture its Boolean result
        Dim taskResult As Boolean
        taskResult = False

        Select Case taskNames(i)
            Case "E01": taskResult = E01_WorkbookLifecycle()
            Case "E02": taskResult = E02_CellReadWrite()
            Case "E03": taskResult = E03_BulkRangeWrite()
            Case "E04": taskResult = E04_FormulaRecalc()
            Case "E05": taskResult = E05_CellFormatting()
            Case "E06": taskResult = E06_RowColumnStructure()
            Case "E07": taskResult = E07_MultiWorksheet()
            Case "E08": taskResult = E08_DataOperations()
            Case "E09": taskResult = E09_ChartGeneration()
            Case "E10": taskResult = E10_ExportPDF()
            Case "E11": taskResult = E11_RunVBAMacro()
            Case "E12": taskResult = E12_ResourceCleanup()
        End Select

        Dim tElapsed As Double
        tElapsed = Timer - tStart
        If tElapsed < 0 Then tElapsed = tElapsed + 86400

        ' Check expected output file
        Dim outFile As String
        outFile = "$vbaOutDir\" & expectedFiles(i)

        Dim fileOK As Boolean
        fileOK = (Dir(outFile) <> "")

        If Not taskResult Then
            Print #fNum, "  => FAIL (task returned False)"
        ElseIf Not fileOK Then
            Print #fNum, "  => FAIL (output file not found: " & expectedFiles(i) & ")"
        Else
            Print #fNum, "  => PASS"
            passCount = passCount + 1
        End If

        Print #fNum, "  Time: " & Format(tElapsed, "0.000") & " s"
        Print #fNum, ""
    Next i

    Dim totalElapsed As Double
    totalElapsed = Timer - totalStart
    If totalElapsed < 0 Then totalElapsed = totalElapsed + 86400

    Print #fNum, "============================================================"
    Print #fNum, "Total: " & passCount & "/12 passed, " & Format(totalElapsed, "0.000") & " s"
    Print #fNum, "Finished: " & Format(Now, "yyyy-mm-dd hh:nn:ss")

    Close #fNum
End Sub
"@

    # Add the logging module
    $logModule = $vbProject.VBComponents.Add(1) # vbext_ct_StdModule
    $logModule.Name = "modLogRunner"
    $logModule.CodeModule.AddFromString($logMacroCode)

    # Save
    $wb.Save()

    Write-Host "Running RunAllAndLog macro (this may take a few minutes)..."
    $excel.Run("RunAllAndLog")

    Write-Host "Macro completed."
    $wb.Save()

} catch {
    Write-Host "ERROR: $_"
    Write-Host $_.Exception.Message
    Write-Host $_.ScriptStackTrace
} finally {
    if ($wb) {
        try { $wb.Close($false) } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null } catch {}
    }
    if ($excel) {
        try { $excel.Quit() } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {}
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

# Display results
Write-Host ""
if (Test-Path $logFile) {
    Write-Host "========== VBA TEST RESULTS =========="
    Get-Content $logFile
    Write-Host "======================================"
} else {
    Write-Host "WARNING: Log file not created"
    Write-Host "Checking output files manually..."
    $expected = @("E01_lifecycle.xlsx","E02_cellrw.xlsx","E03_bulk.xlsx","E04_formula.xlsx","E05_format.xlsx","E06_structure.xlsx","E07_multisheet.xlsx","E08_dataops.xlsx","E09_chart.xlsx","E10_export.pdf","E11_macro.xlsm","E12_cleanup.xlsx")
    foreach ($f in $expected) {
        $fp = Join-Path $outDir $f
        if (Test-Path $fp) {
            Write-Host "  $f - EXISTS ($($(Get-Item $fp).Length) bytes)"
        } else {
            Write-Host "  $f - MISSING"
        }
    }
}

# Cleanup lingering Excel
Start-Sleep -Seconds 3
Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
