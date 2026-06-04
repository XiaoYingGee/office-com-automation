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

    // All 9 capabilities must be present in the matrix.
    for id in [
        "CELL-WRITE-STRING", "CELL-WRITE-NUMBER", "CELL-WRITE-BOOL", "CELL-WRITE-FORMULA",
        "CELL-READ-FORMULA", "CELL-READ-TEXT",
        "RANGE-WRITE-BULK", "RANGE-CLEAR-CONTENTS", "RANGE-COPY-VALUES",
    ] {
        assert!(m.contains(id), "matrix missing row for {id}:\n{m}");
    }

    // Cell-write capabilities: must be ✅ in the Rust column.
    for id in ["CELL-WRITE-STRING", "CELL-WRITE-NUMBER", "CELL-WRITE-BOOL", "CELL-WRITE-FORMULA"] {
        let r = row(&m, id);
        assert!(r.contains("✅"), "{id} not ✅: {r}");
        assert!(!r.contains("❌"), "{id} has ❌ (regressed): {r}");
    }

    // Formula-read capability: must be ✅ (formula string round-trip).
    {
        let r = row(&m, "CELL-READ-FORMULA");
        assert!(r.contains("✅"), "CELL-READ-FORMULA not ✅: {r}");
        assert!(!r.contains("❌"), "CELL-READ-FORMULA has ❌ (regressed): {r}");
    }

    // CELL-READ-TEXT: locale-sensitive — Excel's Range.Text reflects the display format, which can
    // vary by locale.  On English locale we expect ✅, but ⚠️ is also acceptable (lossy read-back
    // due to format string differences).  ❌ (hard failure) is not acceptable.
    {
        let r = row(&m, "CELL-READ-TEXT");
        assert!(
            r.contains("✅") || r.contains("⚠️"),
            "CELL-READ-TEXT should be ✅ or ⚠️ (locale-sensitive), got: {r}"
        );
        assert!(!r.contains("❌"), "CELL-READ-TEXT has ❌ (unexpected hard failure): {r}");
    }

    // Range capabilities: must be ✅.
    for id in ["RANGE-WRITE-BULK", "RANGE-CLEAR-CONTENTS", "RANGE-COPY-VALUES"] {
        let r = row(&m, id);
        assert!(r.contains("✅"), "{id} not ✅: {r}");
        assert!(!r.contains("❌"), "{id} has ❌ (regressed): {r}");
    }

    // The non-rust columns for any capability row should remain untested (⬜).
    let r = row(&m, "CELL-WRITE-STRING");
    assert!(r.contains("⬜"), "other columns should be untested (⬜): {r}");
}
