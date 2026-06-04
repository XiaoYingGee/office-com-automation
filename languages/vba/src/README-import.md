# Importing and Running the VBA Modules

## Prerequisites

- Windows with Microsoft Excel (any version supporting VBA; 2016+ recommended)
- Macro execution must be enabled
- For E11 (VBA macro injection): enable **Trust access to the VBA project object model**

## Step 1: Enable Macros

1. Open Excel
2. Go to **File > Options > Trust Center > Trust Center Settings**
3. Under **Macro Settings**, select **Enable all macros**
4. Check **Trust access to the VBA project object model** (needed for E11)
5. Click OK and restart Excel

## Step 2: Create a Host Workbook

1. Open Excel and create a new blank workbook
2. **Save As** type `.xlsm` (Excel Macro-Enabled Workbook), e.g., `VBA_Runner.xlsm`
3. Place it at: `D:\Workspace\AI\office-com-automation\languages\vba\src\VBA_Runner.xlsm`

> The output directory will be resolved relative to `ThisWorkbook.Path`, so the
> workbook location matters. With the path above, outputs go to
> `D:\Workspace\AI\office-com-automation\out\vba\`.

## Step 3: Import the .bas Modules

1. Press **Alt+F11** to open the VBA Editor (VBE)
2. In the **Project Explorer** (left pane), right-click on your project (e.g., `VBAProject (VBA_Runner.xlsm)`)
3. Select **Import File...**
4. Navigate to `D:\Workspace\AI\office-com-automation\languages\vba\src\`
5. Select `modExcelTasks.bas` and click **Open**
6. Repeat for `modRunner.bas`

Both modules should now appear under **Modules** in the Project Explorer.

## Step 4: Run the Tasks

### Run All Tasks

1. In the VBE, press **Ctrl+G** to open the Immediate window
2. Type `RunAll` and press **Enter**
3. Watch the results appear in the Immediate window

### Run a Single Task

In the Immediate window, type:

```
RunSingle "E01"
```

Or directly call any task procedure:

```
E05_CellFormatting
```

### Run from a Ribbon Button (optional)

1. Go to **Developer > Macros**
2. Select `RunAll` or `RunSingle`
3. Click **Run**

## Step 5: Check Output

Output files are written to: `D:\Workspace\AI\office-com-automation\out\vba\`

Expected output files:

| Task | File | Format |
|------|------|--------|
| E01 | `E01_lifecycle.xlsx` | .xlsx |
| E02 | `E02_cellrw.xlsx` | .xlsx |
| E03 | `E03_bulk.xlsx` | .xlsx |
| E04 | `E04_formula.xlsx` | .xlsx |
| E05 | `E05_format.xlsx` | .xlsx |
| E06 | `E06_structure.xlsx` | .xlsx |
| E07 | `E07_multisheet.xlsx` | .xlsx |
| E08 | `E08_dataops.xlsx` | .xlsx |
| E09 | `E09_chart.xlsx` | .xlsx |
| E10 | `E10_source.xlsx` + `E10_export.pdf` | .xlsx + .pdf |
| E11 | `E11_macro.xlsm` | .xlsm |
| E12 | `E12_cleanup.xlsx` | .xlsx |

## Command-Line Execution (Headless)

You can run the tasks from the command line without opening Excel interactively.

### Method 1: Excel /e with Auto_Open

Add this to a module in the workbook:

```vba
Sub Auto_Open()
    RunAll
    ThisWorkbook.Close SaveChanges:=False
    Application.Quit
End Sub
```

Then run:

```cmd
"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE" /e "D:\Workspace\AI\office-com-automation\languages\vba\src\VBA_Runner.xlsm"
```

### Method 2: VBScript Launcher

Create `run_vba.vbs`:

```vbscript
Dim xlApp, wb
Set xlApp = CreateObject("Excel.Application")
xlApp.Visible = False
xlApp.DisplayAlerts = False

Set wb = xlApp.Workbooks.Open("D:\Workspace\AI\office-com-automation\languages\vba\src\VBA_Runner.xlsm")
xlApp.Run "RunAll"
wb.Close False
xlApp.Quit

Set wb = Nothing
Set xlApp = Nothing
```

Run with: `cscript run_vba.vbs`

> Note: VBScript output goes to Excel's Immediate window buffer, not the console.
> To capture output to console, modify the tasks to use a file logger or
> `WScript.Echo` via a callback.

## Troubleshooting

| Problem | Solution |
|---------|----------|
| "Macros have been disabled" | Enable macros in Trust Center settings |
| E11 fails with error 1004 | Enable "Trust access to VBA project object model" |
| "Path not found" error | Ensure the host workbook is saved at the expected location |
| Timer shows negative values | Midnight wraparound; the code handles this automatically |
| Residual EXCEL.EXE processes | Each task creates and cleans up its own instance; check Task Manager if issues persist |

## Exporting Modules (for version control)

To re-export modules after making changes:

1. In VBE, right-click a module in Project Explorer
2. Select **Export File...**
3. Save as `.bas` to `languages/vba/src/`

This keeps the plain-text `.bas` files in sync with git.
