use std::process::Command;

fn row<'a>(m: &'a str, id: &str) -> &'a str {
    m.lines().find(|l| l.contains(id)).unwrap_or("")
}

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
    // The three cell-write capabilities should be present in the matrix.
    assert!(m.contains("CELL-WRITE-STRING"), "matrix missing rows:\n{m}");
    // Per-row checks: each cell-write capability must be ✅ in the Rust column.
    for id in ["CELL-WRITE-STRING", "CELL-WRITE-NUMBER", "CELL-WRITE-BOOL"] {
        let r = row(&m, id);
        assert!(r.contains("✅"), "{id} not ✅: {r}");
        assert!(!r.contains("❌") && !r.contains("⚠️"), "{id} regressed: {r}");
    }
    // The non-rust columns for CELL-WRITE-STRING should remain untested (⬜).
    let r = row(&m, "CELL-WRITE-STRING");
    assert!(r.contains("⬜"), "other columns should be untested: {r}");
}
