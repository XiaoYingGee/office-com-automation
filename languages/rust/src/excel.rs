// excel.rs - RAII wrapper around Excel.Application built on the com.rs Dispatch layer.
//
// Mirrors the C++ ExcelApp class: hides the application on construction, restores
// settings and calls Quit() on Drop.

use windows::core::Result;
use windows::Win32::System::Variant::VARIANT;

use crate::com::{bstr, empty, i4, range_address, cell_address, Dispatch, Variant};

pub struct ExcelApp {
    app: Dispatch,
}

impl ExcelApp {
    pub fn new() -> Result<Self> {
        let app = Dispatch::create("Excel.Application")?;
        app.put("Visible", Variant::Bool(false))?;
        app.put("DisplayAlerts", Variant::Bool(false))?;
        app.put("ScreenUpdating", Variant::Bool(false))?;
        Ok(Self { app })
    }

    /// Access the underlying Excel.Application dispatch (for Calculate, Run, etc.).
    pub fn app(&self) -> &Dispatch {
        &self.app
    }

    pub fn add_workbook(&self) -> Result<Dispatch> {
        self.app
            .get_dispatch("Workbooks")?
            .call_dispatch("Add", &[empty()])
    }

    pub fn open_workbook(&self, path: &str) -> Result<Dispatch> {
        self.app
            .get_dispatch("Workbooks")?
            .call_dispatch("Open", &[bstr(path)])
    }

    pub fn get_sheet(&self, wb: &Dispatch, index: i32) -> Result<Dispatch> {
        wb.get_dispatch("Worksheets")?
            .get_dispatch_with_args("Item", &[i4(index)])
    }

    pub fn get_sheet_by_name(&self, wb: &Dispatch, name: &str) -> Result<Dispatch> {
        wb.get_dispatch("Worksheets")?
            .get_dispatch_with_args("Item", &[bstr(name)])
    }

    pub fn add_sheet(&self, wb: &Dispatch) -> Result<Dispatch> {
        wb.get_dispatch("Worksheets")?
            .call_dispatch("Add", &[empty(), empty(), empty(), empty()])
    }

    pub fn get_range(&self, ws: &Dispatch, address: &str) -> Result<Dispatch> {
        ws.get_dispatch_with_args("Range", &[bstr(address)])
    }

    pub fn get_cell(&self, ws: &Dispatch, row: i32, col: i32) -> Result<Dispatch> {
        let addr = cell_address(row, col);
        self.get_range(ws, &addr)
    }

    /// Assign a 2D SAFEARRAY VARIANT to a range in a single COM call.
    pub fn bulk_write(
        &self,
        ws: &Dispatch,
        start_row: i32,
        start_col: i32,
        rows: i32,
        cols: i32,
        sa_variant: VARIANT,
    ) -> Result<()> {
        let addr = range_address(start_row, start_col, start_row + rows - 1, start_col + cols - 1);
        let rng = self.get_range(ws, &addr)?;
        rng.put_raw("Value2", sa_variant)
    }

    /// SaveAs .xlsx (xlOpenXMLWorkbook = 51).
    pub fn save_as_xlsx(&self, wb: &Dispatch, path: &str) -> Result<()> {
        wb.call("SaveAs", &[bstr(path), i4(51)])?;
        Ok(())
    }

    /// The application's top-level window handle (used by E12 to find the PID).
    pub fn hwnd(&self) -> Result<i64> {
        Ok(self.app.get("Hwnd")?.as_f64().unwrap_or(0.0) as i64)
    }
}

impl Drop for ExcelApp {
    fn drop(&mut self) {
        let _ = self.app.put("ScreenUpdating", Variant::Bool(true));
        let _ = self.app.put("DisplayAlerts", Variant::Bool(true));
        let _ = self.app.call("Quit", &[]);
    }
}
