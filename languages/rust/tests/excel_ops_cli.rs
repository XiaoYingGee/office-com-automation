use std::process::{Command, Stdio};
use std::io::Write;

fn run_op(json: &str) -> String {
    let mut child = Command::new(env!("CARGO_BIN_EXE_excel-ops"))
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .spawn()
        .unwrap();
    child.stdin.take().unwrap().write_all(json.as_bytes()).unwrap();
    let out = child.wait_with_output().unwrap();
    String::from_utf8(out.stdout).unwrap()
}

#[test]
fn write_then_reference_read_roundtrip() {
    let dir = std::env::temp_dir().join("capfound");
    std::fs::create_dir_all(&dir).unwrap();
    let path = dir.join("rt.xlsx");
    let _ = std::fs::remove_file(&path);
    let p = path.to_string_lossy().replace('\\', "/");
    let w = format!(
        r#"{{"op":"cell.write","path":"{p}","target":{{"sheet":"Sheet1","range":"A1"}},"params":{{"value":"hi","kind":"string"}},"save_as":{{"path":"{p}","format":"xlsx"}}}}"#
    );
    let wr = run_op(&w);
    assert!(wr.contains("\"ok\":true"), "write resp: {wr}");
    let r = format!(
        r#"{{"op":"range.read","path":"{p}","target":{{"sheet":"Sheet1","range":"A1"}},"params":{{}}}}"#
    );
    let rr = run_op(&r);
    assert!(rr.contains("hi"), "read resp: {rr}");
}

#[test]
fn write_number_roundtrips() {
    let dir = std::env::temp_dir().join("capfound");
    std::fs::create_dir_all(&dir).unwrap();
    let path = dir.join("num.xlsx");
    let _ = std::fs::remove_file(&path);
    let p = path.to_string_lossy().replace('\\', "/");
    let w = format!(
        r#"{{"op":"cell.write","path":"{p}","target":{{"range":"B2"}},"params":{{"value":42.5,"kind":"number"}}}}"#
    );
    assert!(run_op(&w).contains("\"ok\":true"));
    let r = format!(
        r#"{{"op":"range.read","path":"{p}","target":{{"range":"B2"}},"params":{{}}}}"#
    );
    let rr = run_op(&r);
    assert!(rr.contains("42.5"), "read resp: {rr}");
}

#[test]
fn write_csv_uses_csv_format() {
    let dir = std::env::temp_dir().join("capfound");
    std::fs::create_dir_all(&dir).unwrap();
    let path = dir.join("out.csv");
    let _ = std::fs::remove_file(&path);
    let p = path.to_string_lossy().replace('\\', "/");
    let w = format!(
        r#"{{"op":"cell.write","path":"{p}","target":{{"range":"A1"}},"params":{{"value":"hello","kind":"string"}}}}"#
    );
    let wr = run_op(&w);
    assert!(wr.contains("\"ok\":true"), "write resp: {wr}");
    assert!(path.exists(), "csv file should exist at {}", path.display());
}

#[test]
fn read_formula_property() {
    let dir = std::env::temp_dir().join("capfound");
    std::fs::create_dir_all(&dir).unwrap();
    let path = dir.join("formula_prop.xlsx");
    let _ = std::fs::remove_file(&path);
    let p = path.to_string_lossy().replace('\\', "/");
    // Write a formula
    let w = format!(
        r#"{{"op":"cell.write","path":"{p}","target":{{"range":"A1"}},"params":{{"value":"=2+3","kind":"formula"}}}}"#
    );
    let wr = run_op(&w);
    assert!(wr.contains("\"ok\":true"), "write formula resp: {wr}");
    // Read back with property=formula
    let r = format!(
        r#"{{"op":"range.read","path":"{p}","target":{{"range":"A1"}},"params":{{"property":"formula"}}}}"#
    );
    let rr = run_op(&r);
    assert!(rr.contains("=2+3"), "formula read resp: {rr}");
    assert!(rr.contains("\"ok\":true"), "formula read resp not ok: {rr}");
}

#[test]
fn write_bulk_then_read_corner() {
    let dir = std::env::temp_dir().join("capfound");
    std::fs::create_dir_all(&dir).unwrap();
    let path = dir.join("bulk.xlsx");
    let _ = std::fs::remove_file(&path);
    let p = path.to_string_lossy().replace('\\', "/");
    // Write 2x2 bulk to A5:B6
    let w = format!(
        r#"{{"op":"range.write_bulk","path":"{p}","target":{{"range":"A5:B6"}},"params":{{"values":[[10,20],[30,40]]}}}}"#
    );
    let wr = run_op(&w);
    assert!(wr.contains("\"ok\":true"), "bulk write resp: {wr}");
    // Read A6 (second row, first col = 30)
    let r = format!(
        r#"{{"op":"range.read","path":"{p}","target":{{"range":"A6"}},"params":{{}}}}"#
    );
    let rr = run_op(&r);
    assert!(rr.contains("30"), "A6 should be 30: {rr}");
}
