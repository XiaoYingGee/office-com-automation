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
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub params: Option<serde_json::Value>,
}

#[derive(Debug, Deserialize, Serialize)]
pub struct Support {
    pub vba: String,
    pub cpp: String,
    pub rust: String,
    pub openxml: String,
}

/// Extract all ```json ... ``` fenced blocks from a markdown string.
/// The opening fence info-string is matched case-insensitively and leading/trailing
/// whitespace is ignored (e.g. "```JSON ", "```Json\t" all match).
fn extract_json_blocks(markdown: &str) -> Vec<&str> {
    let mut blocks = Vec::new();
    let mut rest = markdown;
    // Scan for opening ``` fences followed by an info-string that equals "json"
    // (case-insensitive, trimmed).
    while let Some(fence_start) = rest.find("```") {
        let after_backticks = &rest[fence_start + 3..];
        // Read the info string: everything up to the next newline.
        let newline_pos = after_backticks.find('\n');
        let (info_raw, after_fence) = if let Some(nl) = newline_pos {
            (&after_backticks[..nl], &after_backticks[nl + 1..])
        } else {
            // No newline found — can't be a valid fenced block; stop.
            break;
        };
        // Normalise: strip a possible leading \r from info_raw (for \r\n line endings)
        let info = info_raw.trim_end_matches('\r').trim();
        if !info.eq_ignore_ascii_case("json") {
            // Not a json fence — advance past this ``` and keep looking.
            rest = &rest[fence_start + 3..];
            continue;
        }
        // We're inside a ```json block. Find the closing ```.
        if let Some(end) = after_fence.find("```") {
            let block = after_fence[..end].trim();
            blocks.push(block);
            // Advance past the closing ```
            let consumed = fence_start + 3 + info_raw.len() + 1 + end + 3;
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

    #[test]
    fn extract_blocks_case_insensitive() {
        // Uppercase and mixed-case info strings should be recognised.
        let md = "```JSON\n{\"id\":\"A\"}\n```\n\n```Json \n{\"id\":\"B\"}\n```\n";
        let blocks = extract_json_blocks(md);
        assert_eq!(blocks.len(), 2, "expected 2 blocks, got: {blocks:?}");
        assert!(blocks[0].contains("\"A\""));
        assert!(blocks[1].contains("\"B\""));
    }
}
