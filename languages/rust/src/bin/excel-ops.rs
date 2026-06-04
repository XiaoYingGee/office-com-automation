use std::io::{Read, Write};

use excel_com::com::ComGuard;
use excel_com::ops::{dispatch, ExcelError, ErrorCategory, OpRequest, OpResponse};

fn main() {
    let mut buf = String::new();
    if std::io::stdin().read_to_string(&mut buf).is_err() {
        let resp = OpResponse::err(ExcelError {
            category: ErrorCategory::InvalidArg,
            code: 0,
            message: "failed to read stdin".into(),
            hint: None,
        });
        let s = serde_json::to_string(&resp).unwrap();
        let _ = std::io::stdout().write_all(s.as_bytes());
        return;
    }

    let resp = match serde_json::from_str::<OpRequest>(&buf) {
        Ok(req) => {
            match ComGuard::new() {
                Ok(_com) => dispatch::execute(req),
                Err(e) => OpResponse::err(ExcelError {
                    category: ErrorCategory::ComError,
                    code: e.code().0 as i64,
                    message: format!("CoInitialize failed: {}", e.message()),
                    hint: None,
                }),
            }
        }
        Err(e) => OpResponse::err(ExcelError {
            category: ErrorCategory::InvalidArg,
            code: 0,
            message: format!("bad request json: {e}"),
            hint: None,
        }),
    };

    let s = serde_json::to_string(&resp).unwrap();
    let _ = std::io::stdout().write_all(s.as_bytes());
}
