use std::process::Command;

#[test]
fn capctl_fills_rust_column_for_cell_io() {
    let backend = env!("CARGO_BIN_EXE_excel-ops");
    let capctl = env!("CARGO_BIN_EXE_capctl");
    let catalog = concat!(env!("CARGO_MANIFEST_DIR"), "/../../spec/capabilities/catalog");
    let out = std::env::temp_dir().join("capctl_matrix.md");
    let status = Command::new(capctl)
        .args(["verify","--backend","rust","--backend-cmd",backend,"--reference-cmd",backend,
               "--catalog",catalog,"--out",out.to_str().unwrap()])
        .status().unwrap();
    assert!(status.success(), "capctl exited non-zero");
    let m = std::fs::read_to_string(&out).unwrap();
    // The three cell-write capabilities should be ✅ in the Rust column.
    assert!(m.contains("CELL-WRITE-STRING"), "matrix missing rows:\n{m}");
    // crude check: the Rust column for the string row is ✅ (no ❌/⚠️ in cell-write rows)
    assert!(!m.contains("❌"), "unexpected failures in matrix:\n{m}");
}
