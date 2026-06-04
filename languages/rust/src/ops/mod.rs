pub mod dispatch;

use serde::{Deserialize, Serialize};

#[derive(Debug, Serialize, Deserialize, PartialEq)]
pub struct OpRequest {
    pub op: String,
    pub path: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub target: Option<Target>,
    #[serde(default)]
    pub params: serde_json::Value,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub save_as: Option<SaveAs>,
}

#[derive(Debug, Serialize, Deserialize, PartialEq)]
pub struct Target {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sheet: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub range: Option<String>,
}

#[derive(Debug, Serialize, Deserialize, PartialEq)]
pub struct SaveAs {
    pub path: String,
    pub format: String,
}

#[derive(Debug, Serialize, Deserialize, PartialEq)]
pub struct OpResponse {
    pub ok: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub result: Option<serde_json::Value>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub error: Option<ExcelError>,
}

impl OpResponse {
    pub fn ok(result: serde_json::Value) -> Self {
        Self { ok: true, result: Some(result), error: None }
    }
    pub fn err(error: ExcelError) -> Self {
        Self { ok: false, result: None, error: Some(error) }
    }

    /// Validates that the response fields are internally consistent.
    /// Returns `Err` if:
    /// - `ok == true` but an error is present
    /// - `ok == false` but a result is present
    /// - `ok == false` but no error is present
    pub fn validate(&self) -> Result<(), String> {
        if self.ok && self.error.is_some() {
            return Err("ok response must not carry an error".into());
        }
        if !self.ok && self.result.is_some() {
            return Err("error response must not carry a result".into());
        }
        if !self.ok && self.error.is_none() {
            return Err("error response must carry an error".into());
        }
        Ok(())
    }
}

#[derive(Debug, Serialize, Deserialize, PartialEq)]
pub struct ExcelError {
    pub category: ErrorCategory,
    pub code: i64,
    pub message: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hint: Option<String>,
}

#[derive(Debug, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "PascalCase")]
pub enum ErrorCategory {
    FileNotFound,
    FileLocked,
    InvalidArg,
    RangeParseError,
    SheetNotFound,
    UnsupportedFormat,
    ComError,
    MacroTrustDisabled,
    Timeout,
    Unknown,
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn op_request_roundtrips() {
        let json = r#"{"op":"cell.write","path":"x.xlsx","target":{"sheet":"Sheet1","range":"A1"},"params":{"value":"hi","kind":"string"}}"#;
        let req: OpRequest = serde_json::from_str(json).unwrap();
        assert_eq!(req.op, "cell.write");
        let back = serde_json::to_string(&req).unwrap();
        let req2: OpRequest = serde_json::from_str(&back).unwrap();
        assert_eq!(req, req2);

        // Minimal request: target/save_as must be absent from serialized output
        let minimal: OpRequest = serde_json::from_str(r#"{"op":"range.read","path":"x.xlsx"}"#).unwrap();
        assert!(minimal.target.is_none());
        assert!(minimal.save_as.is_none());
        let minimal_json = serde_json::to_string(&minimal).unwrap();
        assert!(!minimal_json.contains("\"target\""), "target must be omitted when None");
        assert!(!minimal_json.contains("\"save_as\""), "save_as must be omitted when None");
    }

    #[test]
    fn op_response_serializes_ok_and_err() {
        let ok = OpResponse::ok(serde_json::json!({"value":"hi"}));
        let ok_str = serde_json::to_string(&ok).unwrap();
        assert!(ok_str.contains("\"ok\":true"));
        // Structural round-trip check for ok response
        let ok_rt: OpResponse = serde_json::from_str(&ok_str).unwrap();
        assert!(ok_rt.ok);
        assert!(ok_rt.result.is_some());
        assert!(ok_rt.error.is_none());

        let err = OpResponse::err(ExcelError {
            category: ErrorCategory::FileNotFound,
            code: 2,
            message: "nope".into(),
            hint: None,
        });
        let s = serde_json::to_string(&err).unwrap();
        assert!(s.contains("\"ok\":false"));
        assert!(s.contains("FileNotFound"));
        // Structural round-trip check for err response
        let err_rt: OpResponse = serde_json::from_str(&s).unwrap();
        assert!(!err_rt.ok);
        assert!(err_rt.error.is_some());
        assert!(err_rt.result.is_none());
    }

    #[test]
    fn op_response_rejects_inconsistent() {
        let json = r#"{"ok":true,"error":{"category":"Unknown","code":0,"message":"x"}}"#;
        let resp: OpResponse = serde_json::from_str(json).unwrap();
        assert!(resp.validate().is_err(), "ok=true with error present must fail validate()");
    }
}
