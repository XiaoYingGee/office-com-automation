use serde::{Deserialize, Serialize};
use serde_json::Value;
use std::path::Path;

/// Matches capability.schema.json (with assert.read as an object per architecture doc)
#[derive(Debug, Deserialize, Serialize)]
pub struct Capability {
    pub id: String,
    pub name: String,
    pub desc: String,
    pub op: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub param_path: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sample: Option<Value>,
    pub verify: Verify,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub errors: Option<Vec<String>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub com_ref: Option<String>,
    pub support: Support,
}

#[derive(Debug, Deserialize, Serialize)]
pub struct Verify {
    pub action: Action,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub reopen: Option<bool>,
    pub assert: Assert,
}

#[derive(Debug, Deserialize, Serialize)]
pub struct Action {
    pub op: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub target: Option<ActionTarget>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub params: Option<Value>,
}

#[derive(Debug, Deserialize, Serialize)]
pub struct ActionTarget {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sheet: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub range: Option<String>,
}

#[derive(Debug, Deserialize, Serialize)]
pub struct Assert {
    pub read: AssertRead,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub expect: Option<Value>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tol: Option<f64>,
}

#[derive(Debug, Deserialize, Serialize)]
pub struct AssertRead {
    pub op: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub target: Option<ActionTarget>,
}

#[derive(Debug, Deserialize, Serialize)]
pub struct Support {
    pub vba: String,
    pub cpp: String,
    pub rust: String,
    pub openxml: String,
}

/// Extract all ```json ... ``` fenced blocks from a markdown string.
fn extract_json_blocks(markdown: &str) -> Vec<&str> {
    let mut blocks = Vec::new();
    let mut rest = markdown;
    while let Some(start) = rest.find("```json") {
        let after_fence = &rest[start + 7..];
        // skip optional newline after ```json
        let content_start = if after_fence.starts_with('\n') {
            &after_fence[1..]
        } else if after_fence.starts_with("\r\n") {
            &after_fence[2..]
        } else {
            after_fence
        };
        if let Some(end) = content_start.find("```") {
            let block = content_start[..end].trim();
            blocks.push(block);
            // advance past the closing ```
            let consumed = (start + 7) + (after_fence.len() - content_start.len()) + end + 3;
            rest = &rest[consumed..];
        } else {
            break;
        }
    }
    blocks
}

/// Load all capabilities from *.md files in `dir`.
pub fn load_dir(dir: &Path) -> std::io::Result<Vec<Capability>> {
    let mut caps = Vec::new();
    let entries = std::fs::read_dir(dir)?;
    let mut paths: Vec<_> = entries
        .filter_map(|e| e.ok())
        .filter(|e| {
            e.path()
                .extension()
                .and_then(|x| x.to_str())
                .map(|x| x.eq_ignore_ascii_case("md"))
                .unwrap_or(false)
        })
        .map(|e| e.path())
        .collect();
    paths.sort(); // deterministic ordering
    for path in paths {
        let content = std::fs::read_to_string(&path)?;
        for block in extract_json_blocks(&content) {
            match serde_json::from_str::<Capability>(block) {
                Ok(cap) => caps.push(cap),
                Err(e) => {
                    eprintln!("WARN: skipping malformed json block in {}: {e}", path.display());
                }
            }
        }
    }
    Ok(caps)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn extract_blocks_works() {
        let md = "# Title\n\nSome prose.\n\n```json\n{\"id\":\"X\"}\n```\n\nMore text.\n\n```json\n{\"id\":\"Y\"}\n```\n";
        let blocks = extract_json_blocks(md);
        assert_eq!(blocks.len(), 2);
        assert!(blocks[0].contains("\"X\""));
        assert!(blocks[1].contains("\"Y\""));
    }
}
