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
    }

    #[test]
    fn op_response_serializes_ok_and_err() {
        let ok = OpResponse::ok(serde_json::json!({"value":"hi"}));
        assert!(serde_json::to_string(&ok).unwrap().contains("\"ok\":true"));
        let err = OpResponse::err(ExcelError {
            category: ErrorCategory::FileNotFound,
            code: 2,
            message: "nope".into(),
            hint: None,
        });
        let s = serde_json::to_string(&err).unwrap();
        assert!(s.contains("\"ok\":false"));
        assert!(s.contains("FileNotFound"));
    }
}
