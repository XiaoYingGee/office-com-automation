//! capctl — capability orchestrator
//!
//! Usage:
//!   capctl verify --backend <name> --backend-cmd <exe> --reference-cmd <exe>
//!                 --catalog <dir> --out <matrix.md>

use std::io::Write;
use std::path::{Path, PathBuf};
use std::process::{Command, Stdio};

use excel_com::ops::catalog::{load_dir, Capability};
use excel_com::ops::{OpRequest, OpResponse, SaveAs, Target};

// ── CLI parsing ──────────────────────────────────────────────────────────────

struct Args {
    backend: String,
    backend_cmd: String,
    reference_cmd: String,
    catalog: PathBuf,
    out: PathBuf,
}

fn parse_args() -> Result<Args, String> {
    let raw: Vec<String> = std::env::args().collect();
    // expect: capctl verify --backend X --backend-cmd Y --reference-cmd Z --catalog D --out F
    if raw.len() < 2 || raw[1] != "verify" {
        return Err("Usage: capctl verify --backend <name> --backend-cmd <exe> --reference-cmd <exe> --catalog <dir> --out <file>".into());
    }
    let mut backend = None;
    let mut backend_cmd = None;
    let mut reference_cmd = None;
    let mut catalog = None;
    let mut out = None;

    let mut i = 2usize;
    while i < raw.len() {
        match raw[i].as_str() {
            "--backend" => { backend = raw.get(i + 1).cloned(); i += 2; }
            "--backend-cmd" => { backend_cmd = raw.get(i + 1).cloned(); i += 2; }
            "--reference-cmd" => { reference_cmd = raw.get(i + 1).cloned(); i += 2; }
            "--catalog" => { catalog = raw.get(i + 1).cloned(); i += 2; }
            "--out" => { out = raw.get(i + 1).cloned(); i += 2; }
            other => return Err(format!("Unknown argument: {other}")),
        }
    }

    let backend = backend.ok_or("--backend required")?;
    let backend_lower = backend.to_ascii_lowercase();
    const VALID_BACKENDS: &[&str] = &["vba", "cpp", "rust", "openxml"];
    if !VALID_BACKENDS.contains(&backend_lower.as_str()) {
        return Err(format!(
            "unknown --backend {:?}; expected one of: {}",
            backend,
            VALID_BACKENDS.join(", ")
        ));
    }
    Ok(Args {
        backend,
        backend_cmd: backend_cmd.ok_or("--backend-cmd required")?,
        reference_cmd: reference_cmd.ok_or("--reference-cmd required")?,
        catalog: PathBuf::from(catalog.ok_or("--catalog required")?),
        out: PathBuf::from(out.ok_or("--out required")?),
    })
}

// ── Subprocess helpers ───────────────────────────────────────────────────────

fn run_op(exe: &str, req: &OpRequest) -> Result<OpResponse, String> {
    let json = serde_json::to_string(req).map_err(|e| e.to_string())?;
    let mut child = Command::new(exe)
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::inherit())
        .spawn()
        .map_err(|e| format!("failed to spawn {exe}: {e}"))?;
    child
        .stdin
        .take()
        .ok_or_else(|| "failed to acquire stdin pipe".to_string())?
        .write_all(json.as_bytes())
        .map_err(|e| e.to_string())?;
    let out = child.wait_with_output().map_err(|e| e.to_string())?;
    let stdout = String::from_utf8_lossy(&out.stdout).into_owned();
    if !out.status.success() {
        return Err(format!("{exe} exited {}: {}", out.status, stdout.trim()));
    }
    serde_json::from_str::<OpResponse>(&stdout)
        .map_err(|e| format!("failed to parse response from {exe}: {e}\nraw: {stdout}"))
}

// ── Result types ─────────────────────────────────────────────────────────────

#[derive(Debug)]
enum CapResult {
    Pass,
    Lossy { got: serde_json::Value, expected: serde_json::Value },
    BackendFail { category: String },
    ReadFail { category: String },
}

impl CapResult {
    fn symbol(&self) -> &'static str {
        match self {
            CapResult::Pass => "✅",
            CapResult::Lossy { .. } => "⚠️",
            CapResult::BackendFail { .. } | CapResult::ReadFail { .. } => "❌",
        }
    }
    fn footnote(&self, id: &str) -> Option<String> {
        match self {
            CapResult::Pass => None,
            CapResult::Lossy { got, expected } => {
                Some(format!("**{id}**: ⚠️ read-back mismatch: got `{got}` expected `{expected}`"))
            }
            CapResult::BackendFail { category } => {
                Some(format!("**{id}**: ❌ backend error category `{category}`"))
            }
            CapResult::ReadFail { category } => {
                Some(format!("**{id}**: ❌ reference-reader error category `{category}`"))
            }
        }
    }
}

// ── Comparison helper ────────────────────────────────────────────────────────

fn values_match(got: &serde_json::Value, expect: &serde_json::Value, tol: f64) -> bool {
    use serde_json::Value::*;
    match (got, expect) {
        // Both null → match
        (Null, Null) => true,
        (Number(a), Number(b)) => {
            let af = a.as_f64().unwrap_or(f64::NAN);
            let bf = b.as_f64().unwrap_or(f64::NAN);
            (af - bf).abs() <= tol
        }
        (Bool(a), Bool(b)) => a == b,
        (String(a), String(b)) => a == b,
        _ => got == expect,
    }
}

// ── Verification logic ───────────────────────────────────────────────────────

fn verify_one(
    cap: &Capability,
    backend_cmd: &str,
    reference_cmd: &str,
    workdir: &Path,
) -> CapResult {
    // Each capability gets its own temp file; remove it so cell.write creates fresh.
    let safe_id = cap.id.replace(['/', '\\', ':', '*', '?', '"', '<', '>', '|'], "_");
    let tmp_path = workdir.join(format!("{safe_id}.xlsx"));
    let _ = std::fs::remove_file(&tmp_path);
    let tmp_str = tmp_path.to_string_lossy().replace('\\', "/");

    // Build write request
    let action = &cap.verify.action;
    let write_target = action.target.as_ref().map(|t| Target {
        sheet: t.sheet.clone(),
        range: t.range.clone(),
    });
    let write_req = OpRequest {
        op: action.op.clone(),
        path: tmp_str.clone(),
        target: write_target,
        params: action.params.clone().unwrap_or(serde_json::Value::Object(Default::default())),
        save_as: Some(SaveAs {
            path: tmp_str.clone(),
            format: "xlsx".into(),
        }),
    };

    // Spawn backend
    let write_resp = match run_op(backend_cmd, &write_req) {
        Ok(r) => r,
        Err(e) => {
            return CapResult::BackendFail { category: format!("spawn-error: {e}") };
        }
    };
    if !write_resp.ok {
        let cat = write_resp
            .error
            .as_ref()
            .map(|e| format!("{:?}", e.category))
            .unwrap_or_else(|| "Unknown".into());
        return CapResult::BackendFail { category: cat };
    }

    // Build read request
    // Stateless backends persist on write and the assert read opens the file fresh,
    // so reopen is implicitly honored; the field is retained for future stateful backends.
    let assert = &cap.verify.assert;
    let read_target = assert.read.target.as_ref().map(|t| Target {
        sheet: t.sheet.clone(),
        range: t.range.clone(),
    });
    let read_req = OpRequest {
        op: assert.read.op.clone(),
        path: tmp_str.clone(),
        target: read_target,
        params: assert.read.params.clone().unwrap_or(serde_json::Value::Object(Default::default())),
        save_as: None,
    };

    // Spawn reference reader
    let read_resp = match run_op(reference_cmd, &read_req) {
        Ok(r) => r,
        Err(e) => {
            return CapResult::ReadFail { category: format!("spawn-error: {e}") };
        }
    };
    if !read_resp.ok {
        let cat = read_resp
            .error
            .as_ref()
            .map(|e| format!("{:?}", e.category))
            .unwrap_or_else(|| "Unknown".into());
        return CapResult::ReadFail { category: cat };
    }

    // Extract result.value from the range.read response:
    // range.read returns {"kind":..., "value": <actual>}
    let got_value = read_resp
        .result
        .as_ref()
        .and_then(|r| r.get("value"))
        .cloned()
        .unwrap_or(serde_json::Value::Null);

    let expect = match &assert.expect {
        Some(e) => e.clone(),
        None => return CapResult::Pass, // no assertion defined
    };
    let tol = assert.tol.unwrap_or(1e-9);

    if values_match(&got_value, &expect, tol) {
        CapResult::Pass
    } else {
        CapResult::Lossy { got: got_value, expected: expect }
    }
}

// ── Matrix generation ────────────────────────────────────────────────────────

// TODO: multi-backend aggregation — currently regenerates from scratch filling only --backend's column; a future step should MERGE per-backend runs into one matrix.
fn generate_matrix(
    caps: &[Capability],
    results: &[(String, CapResult)], // (id, result) for the run backend
    backend_col: &str,
) -> String {
    // Map id -> result symbol for quick lookup
    let result_map: std::collections::HashMap<&str, &CapResult> = results
        .iter()
        .map(|(id, r)| (id.as_str(), r))
        .collect();

    // Column order: vba, cpp, rust, openxml
    let cols = ["vba", "cpp", "rust", "openxml"];
    let col_headers = ["VBA", "C++", "Rust", "OpenXML"];
    let backend_lower = backend_col.to_ascii_lowercase();

    let mut out = String::new();
    out.push_str("# Capability Support Matrix\n\n");
    out.push_str("> Generated by `capctl verify`. ✅ full · ⚠️ partial/lossy · ❌ unsupported · ⬜ untested\n\n");
    out.push_str("| id | VBA | C++ | Rust | OpenXML |\n");
    out.push_str("|---|:---:|:---:|:---:|:---:|\n");

    for cap in caps {
        out.push_str(&format!("| {} ", cap.id));
        for (col_key, _col_hdr) in cols.iter().zip(col_headers.iter()) {
            if *col_key == backend_lower {
                let sym = result_map
                    .get(cap.id.as_str())
                    .map(|r| r.symbol())
                    .unwrap_or("⬜");
                out.push_str(&format!("| {sym} "));
            } else {
                // Keep existing support value from catalog as-is (all ⬜ for now)
                let existing = match *col_key {
                    "vba" => cap.support.vba.as_str(),
                    "cpp" => cap.support.cpp.as_str(),
                    "rust" => cap.support.rust.as_str(),
                    "openxml" => cap.support.openxml.as_str(),
                    _ => "⬜",
                };
                out.push_str(&format!("| {existing} "));
            }
        }
        out.push_str("|\n");
    }

    // Footnotes
    let footnotes: Vec<String> = results
        .iter()
        .filter_map(|(id, r)| r.footnote(id))
        .collect();
    if !footnotes.is_empty() {
        out.push_str("\n## Notes\n\n");
        for note in &footnotes {
            out.push_str(&format!("- {note}\n"));
        }
    }

    out
}

// ── Entry point ──────────────────────────────────────────────────────────────

fn main() {
    let args = match parse_args() {
        Ok(a) => a,
        Err(e) => {
            eprintln!("ERROR: {e}");
            std::process::exit(1);
        }
    };

    // Load capabilities
    let caps = match load_dir(&args.catalog) {
        Ok(c) => c,
        Err(e) => {
            eprintln!("ERROR: failed to load catalog from {}: {e}", args.catalog.display());
            std::process::exit(1);
        }
    };
    if caps.is_empty() {
        eprintln!("WARN: no capabilities found in {}", args.catalog.display());
    }

    // Create temp workdir for this run
    let workdir = std::env::temp_dir().join("capctl").join(&args.backend);
    if let Err(e) = std::fs::create_dir_all(&workdir) {
        eprintln!("ERROR: cannot create workdir {}: {e}", workdir.display());
        std::process::exit(1);
    }

    // Run verification for each capability
    let mut results: Vec<(String, CapResult)> = Vec::new();
    for cap in &caps {
        eprintln!("  verifying {} ...", cap.id);
        let result = verify_one(cap, &args.backend_cmd, &args.reference_cmd, &workdir);
        eprintln!("    → {}", result.symbol());
        results.push((cap.id.clone(), result));
    }

    // Generate matrix markdown
    let matrix = generate_matrix(&caps, &results, &args.backend);

    // Write output
    if let Some(parent) = args.out.parent() {
        let _ = std::fs::create_dir_all(parent);
    }
    match std::fs::write(&args.out, &matrix) {
        Ok(_) => eprintln!("matrix written to {}", args.out.display()),
        Err(e) => {
            eprintln!("ERROR: failed to write matrix to {}: {e}", args.out.display());
            std::process::exit(1);
        }
    }
}
