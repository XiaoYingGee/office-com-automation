"""
Excel Automation Benchmark — 4-backend comparison

Backends:
  1. pywin32  — Python COM (cross-process IPC per property)
  2. vba      — Python + VBA Application.Run (in-process, zero IPC)
  3. rust     — Rust COM exe (backend-protocol, one spawn per op)
  4. openxml  — .NET OpenXML exe (backend-protocol, file-level, no Excel)

Usage:
  python benchmark.py --all
  python benchmark.py --backends pywin32,vba
  python benchmark.py --backends rust,openxml --rust-exe path/to/excel-ops.exe --openxml-exe path/to/ExcelOps.exe

Benchmark tasks (matching project's Cell I/O domain + bulk writes):
  B1: cell.write (single cell, various types)
  B2: range.write_bulk (100x10 array)
  B3: range.read (single cell after write)
  B4: range.clear
  B5: inspect (full workbook structure)
  B6: batch actions (10 mixed writes)

Output: Markdown table + JSON results file
"""
import argparse
import json
import os
import subprocess
import sys
import tempfile
import time

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

# ---------------------------------------------------------------------------
# Backend-protocol runner (Rust / OpenXML)
# ---------------------------------------------------------------------------

class ProtocolBackend:
    """Drives a backend-protocol exe (stdin JSON -> stdout JSON, one spawn per op)."""

    def __init__(self, name, exe_path):
        self.name = name
        self.exe_path = exe_path
        if not os.path.isfile(exe_path):
            raise FileNotFoundError(f"{name} exe not found: {exe_path}")

    def run_op(self, op_request):
        """Run one OpRequest, return (ok, elapsed_ms, response_dict)."""
        payload = json.dumps(op_request, ensure_ascii=False)
        t0 = time.perf_counter()
        proc = subprocess.run(
            [self.exe_path],
            input=payload,
            capture_output=True,
            text=True,
            timeout=30,
        )
        elapsed = (time.perf_counter() - t0) * 1000
        if proc.returncode != 0 and not proc.stdout.strip():
            return False, elapsed, {"ok": False, "error": {"message": proc.stderr.strip()}}
        try:
            resp = json.loads(proc.stdout)
        except Exception:
            resp = {"ok": False, "error": {"message": f"bad stdout: {proc.stdout[:200]}"}}
        return resp.get("ok", False), elapsed, resp


# ---------------------------------------------------------------------------
# COM Backend runner (pywin32 / VBA)
# ---------------------------------------------------------------------------

class COMBackend:
    """Drives pywin32 or VBA backend via the excel_editor.py engine (stateful session)."""

    def __init__(self, name, visible=False):
        self.name = name
        self.visible = visible
        self.engine = None

    def start(self, filepath):
        # Import from skill scripts
        skill_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "scripts")
        sys.path.insert(0, skill_dir)
        from excel_editor import ExcelCOM, ExcelVBA

        if self.name == "vba":
            self.engine = ExcelVBA(visible=self.visible)
        else:
            self.engine = ExcelCOM(visible=self.visible)

        if os.path.isfile(filepath):
            self.engine.open(filepath)
        else:
            self.engine.create(filepath)

    def stop(self):
        if self.engine:
            try:
                self.engine.close()
            except Exception:
                pass
            self.engine = None


# ---------------------------------------------------------------------------
# Benchmark Tasks
# ---------------------------------------------------------------------------

def gen_bulk_data(rows=100, cols=10):
    return [[f"R{r}C{c}" if c == 0 else r * cols + c for c in range(cols)] for r in range(rows)]


def run_benchmark_protocol(backend, workdir):
    """Run benchmarks for a backend-protocol exe."""
    results = {}
    filepath = os.path.join(workdir, f"bench_{backend.name}.xlsx")
    # Ensure clean
    if os.path.isfile(filepath):
        os.remove(filepath)

    # B1: cell.write (5 cells, various types)
    times = []
    test_values = [
        ("string", "Hello"),
        ("number", 42),
        ("number", 3.14159),
        ("bool", True),
        ("formula", "=1+1"),
    ]
    for i, (kind, value) in enumerate(test_values):
        req = {
            "op": "cell.write",
            "path": filepath,
            "target": {"sheet": "Sheet1", "range": f"A{i+1}"},
            "params": {"kind": kind, "value": value},
        }
        ok, ms, _ = backend.run_op(req)
        if ok:
            times.append(ms)
    results["B1_cell_write"] = {
        "ops": len(times),
        "total_ms": round(sum(times), 1),
        "avg_ms": round(sum(times) / len(times), 1) if times else 0,
    }

    # B2: range.write_bulk (100x10)
    bulk_data = gen_bulk_data(100, 10)
    req = {
        "op": "range.write_bulk",
        "path": filepath,
        "target": {"sheet": "Sheet1", "range": "B1:K100"},
        "params": {"values": bulk_data},
    }
    ok, ms, _ = backend.run_op(req)
    results["B2_write_bulk_100x10"] = {"ops": 1, "total_ms": round(ms, 1), "ok": ok}

    # B3: range.read (single cell)
    req = {
        "op": "range.read",
        "path": filepath,
        "target": {"sheet": "Sheet1", "range": "A1"},
        "params": {"property": "value2"},
    }
    ok, ms, resp = backend.run_op(req)
    results["B3_range_read"] = {"ops": 1, "total_ms": round(ms, 1), "ok": ok}

    # B4: range.clear
    req = {
        "op": "range.clear",
        "path": filepath,
        "target": {"sheet": "Sheet1", "range": "B1:K100"},
        "params": {"mode": "contents"},
    }
    ok, ms, _ = backend.run_op(req)
    results["B4_range_clear"] = {"ops": 1, "total_ms": round(ms, 1), "ok": ok}

    # B5: inspect
    req = {
        "op": "inspect",
        "path": filepath,
        "params": {},
    }
    ok, ms, _ = backend.run_op(req)
    results["B5_inspect"] = {"ops": 1, "total_ms": round(ms, 1), "ok": ok}

    # B6: batch writes (10 individual cell.write calls)
    times = []
    for i in range(10):
        req = {
            "op": "cell.write",
            "path": filepath,
            "target": {"sheet": "Sheet1", "range": f"M{i+1}"},
            "params": {"kind": "number", "value": i * 100},
        }
        ok, ms, _ = backend.run_op(req)
        if ok:
            times.append(ms)
    results["B6_batch_10_writes"] = {
        "ops": len(times),
        "total_ms": round(sum(times), 1),
        "avg_ms": round(sum(times) / len(times), 1) if times else 0,
    }

    # B7: set_format (format 5 cells)
    times = []
    for i in range(5):
        req = {
            "op": "set_format",
            "path": filepath,
            "target": {"sheet": "Sheet1", "range": f"A{i+1}"},
            "params": {"bold": True, "font_size": 14, "bg_color": 65535},
        }
        ok, ms, _ = backend.run_op(req)
        if ok:
            times.append(ms)
    results["B7_set_format"] = {
        "ops": len(times),
        "total_ms": round(sum(times), 1),
        "avg_ms": round(sum(times) / len(times), 1) if times else 0,
    }

    # B8: row.insert (insert 5 rows)
    req = {
        "op": "row.insert",
        "path": filepath,
        "target": {"sheet": "Sheet1"},
        "params": {"row": 1, "count": 5},
    }
    ok, ms, _ = backend.run_op(req)
    results["B8_insert_rows"] = {"ops": 1, "total_ms": round(ms, 1), "ok": ok}

    # B9: sheet.add + sheet.rename + sheet.delete
    times = []
    req = {"op": "sheet.add", "path": filepath, "params": {"name": "BenchSheet"}}
    ok, ms, _ = backend.run_op(req)
    if ok: times.append(ms)
    req = {"op": "sheet.rename", "path": filepath, "target": {"sheet": "BenchSheet"}, "params": {"name": "Renamed"}}
    ok, ms, _ = backend.run_op(req)
    if ok: times.append(ms)
    req = {"op": "sheet.delete", "path": filepath, "target": {"sheet": "Renamed"}, "params": {}}
    ok, ms, _ = backend.run_op(req)
    if ok: times.append(ms)
    results["B9_sheet_ops"] = {
        "ops": len(times),
        "total_ms": round(sum(times), 1),
        "avg_ms": round(sum(times) / len(times), 1) if times else 0,
    }

    # Cleanup
    if os.path.isfile(filepath):
        os.remove(filepath)

    return results


def run_benchmark_com(backend, workdir):
    """Run benchmarks for a COM-based backend (pywin32/vba)."""
    results = {}
    filepath = os.path.join(workdir, f"bench_{backend.name}.xlsx")
    if os.path.isfile(filepath):
        os.remove(filepath)

    backend.start(filepath)
    engine = backend.engine

    try:
        # B1: cell.write (5 cells, various types)
        test_values = [
            ("auto", "Hello"),
            ("auto", 42),
            ("auto", 3.14159),
            ("auto", True),
            ("formula", "=1+1"),
        ]
        times = []
        for i, (kind, value) in enumerate(test_values):
            t0 = time.perf_counter()
            engine.write_cell("Sheet1", f"A{i+1}", value, kind)
            elapsed = (time.perf_counter() - t0) * 1000
            times.append(elapsed)
        results["B1_cell_write"] = {
            "ops": len(times),
            "total_ms": round(sum(times), 1),
            "avg_ms": round(sum(times) / len(times), 1),
        }

        # B2: write_range (100x10)
        bulk_data = gen_bulk_data(100, 10)
        # Convert to tuple-of-tuples for COM
        bulk_tuple = tuple(tuple(row) for row in bulk_data)
        t0 = time.perf_counter()
        if backend.name == "vba":
            actions = [{"action": "write_range", "sheet": "Sheet1", "range": "B1:K100", "values": bulk_data}]
            engine.execute_actions(actions)
        else:
            engine.write_range("Sheet1", "B1:K100", bulk_tuple)
        elapsed = (time.perf_counter() - t0) * 1000
        results["B2_write_bulk_100x10"] = {"ops": 1, "total_ms": round(elapsed, 1), "ok": True}

        # B3: read cell
        t0 = time.perf_counter()
        if backend.name == "vba":
            actions = [{"action": "read_cell", "sheet": "Sheet1", "cell": "A1"}]
            engine.execute_actions(actions)
        else:
            engine.read_cell("Sheet1", "A1")
        elapsed = (time.perf_counter() - t0) * 1000
        results["B3_range_read"] = {"ops": 1, "total_ms": round(elapsed, 1), "ok": True}

        # B4: clear range
        t0 = time.perf_counter()
        if backend.name == "vba":
            actions = [{"action": "clear_range", "sheet": "Sheet1", "range": "B1:K100"}]
            engine.execute_actions(actions)
        else:
            engine.clear_range("Sheet1", "B1:K100")
        elapsed = (time.perf_counter() - t0) * 1000
        results["B4_range_clear"] = {"ops": 1, "total_ms": round(elapsed, 1), "ok": True}

        # B5: inspect
        t0 = time.perf_counter()
        engine.inspect()
        elapsed = (time.perf_counter() - t0) * 1000
        results["B5_inspect"] = {"ops": 1, "total_ms": round(elapsed, 1), "ok": True}

        # B6: batch 10 writes
        t0 = time.perf_counter()
        if backend.name == "vba":
            actions = [
                {"action": "write_cell", "sheet": "Sheet1", "cell": f"M{i+1}", "value": i * 100}
                for i in range(10)
            ]
            engine.execute_actions(actions)
        else:
            for i in range(10):
                engine.write_cell("Sheet1", f"M{i+1}", i * 100)
        elapsed = (time.perf_counter() - t0) * 1000
        results["B6_batch_10_writes"] = {
            "ops": 10,
            "total_ms": round(elapsed, 1),
            "avg_ms": round(elapsed / 10, 1),
        }

        # B7: set_format (5 cells)
        t0 = time.perf_counter()
        if backend.name == "vba":
            actions = [
                {"action": "set_format", "sheet": "Sheet1", "range": f"A{i+1}",
                 "bold": True, "font_size": 14, "bg_color": 65535}
                for i in range(5)
            ]
            engine.execute_actions(actions)
        else:
            for i in range(5):
                engine.set_format("Sheet1", f"A{i+1}", bold=True, font_size=14, bg_color=65535)
        elapsed = (time.perf_counter() - t0) * 1000
        results["B7_set_format"] = {
            "ops": 5,
            "total_ms": round(elapsed, 1),
            "avg_ms": round(elapsed / 5, 1),
        }

        # B8: insert rows
        t0 = time.perf_counter()
        if backend.name == "vba":
            engine.execute_actions([{"action": "insert_rows", "sheet": "Sheet1", "row": 1, "count": 5}])
        else:
            engine.insert_rows("Sheet1", 1, 5)
        elapsed = (time.perf_counter() - t0) * 1000
        results["B8_insert_rows"] = {"ops": 1, "total_ms": round(elapsed, 1), "ok": True}

        # B9: sheet ops (add + rename + delete)
        t0 = time.perf_counter()
        if backend.name == "vba":
            engine.execute_actions([
                {"action": "add_sheet", "name": "BenchSheet"},
                {"action": "rename_sheet", "sheet": "BenchSheet", "new_name": "Renamed"},
                {"action": "delete_sheet", "sheet": "Renamed"},
            ])
        else:
            engine.add_sheet("BenchSheet")
            engine.rename_sheet("BenchSheet", "Renamed")
            engine.delete_sheet("Renamed")
        elapsed = (time.perf_counter() - t0) * 1000
        results["B9_sheet_ops"] = {
            "ops": 3,
            "total_ms": round(elapsed, 1),
            "avg_ms": round(elapsed / 3, 1),
        }

        # Save before close
        engine.save()

    finally:
        backend.stop()
        if os.path.isfile(filepath):
            os.remove(filepath)

    return results


# ---------------------------------------------------------------------------
# Report Generation
# ---------------------------------------------------------------------------

def generate_report(all_results, output_dir):
    """Generate markdown report + JSON file."""
    benchmarks = ["B1_cell_write", "B2_write_bulk_100x10", "B3_range_read",
                  "B4_range_clear", "B5_inspect", "B6_batch_10_writes",
                  "B7_set_format", "B8_insert_rows", "B9_sheet_ops"]
    labels = {
        "B1_cell_write": "Cell Write (5 cells)",
        "B2_write_bulk_100x10": "Bulk Write (100x10)",
        "B3_range_read": "Read Cell",
        "B4_range_clear": "Clear Range (1000 cells)",
        "B5_inspect": "Inspect Workbook",
        "B6_batch_10_writes": "Batch 10 Writes",
        "B7_set_format": "Format 5 Cells (bold+size+color)",
        "B8_insert_rows": "Insert 5 Rows",
        "B9_sheet_ops": "Sheet Add+Rename+Delete",
    }

    backends = list(all_results.keys())

    lines = []
    lines.append("# Excel Automation Benchmark Results\n")
    lines.append(f"> Date: {time.strftime('%Y-%m-%d %H:%M')}")
    lines.append(f"> Backends: {', '.join(backends)}")
    lines.append("")
    lines.append("## Latency (ms)\n")

    # Header
    header = "| Benchmark | " + " | ".join(backends) + " |"
    sep = "|-----------|" + "|".join(["------:" for _ in backends]) + "|"
    lines.append(header)
    lines.append(sep)

    for bm in benchmarks:
        row = f"| {labels.get(bm, bm)} |"
        for b in backends:
            data = all_results[b].get(bm, {})
            ms = data.get("total_ms", "—")
            if isinstance(ms, (int, float)) and ms > 0:
                row += f" {ms:.0f}ms |"
            elif data.get("note"):
                row += f" {data['note']} |"
            else:
                row += " — |"
        lines.append(row)

    lines.append("")
    lines.append("## Notes\n")
    lines.append("- **pywin32**: Python → cross-process COM IPC (one call per property)")
    lines.append("- **vba**: Python → Application.Run → in-process VBA (zero IPC for object model)")
    lines.append("- **rust**: Standalone exe, Excel COM, one process spawn per operation")
    lines.append("- **openxml**: Standalone exe, direct .xlsx file manipulation, no Excel process")
    lines.append("")
    lines.append("### Architecture\n")
    lines.append("```")
    lines.append("pywin32:  Python ──COM IPC (per property)──→ Excel process")
    lines.append("vba:      Python ──App.Run (1x)──→ Excel process内 VBA execution")
    lines.append("rust:     [spawn] ──COM──→ Excel process ──close ── (per op)")
    lines.append("openxml:  [spawn] ──file I/O──→ .xlsx (no Excel)")
    lines.append("```")

    report_md = "\n".join(lines)

    # Write files
    md_path = os.path.join(output_dir, "benchmark-results.md")
    json_path = os.path.join(output_dir, "benchmark-results.json")

    with open(md_path, "w", encoding="utf-8") as f:
        f.write(report_md)
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(all_results, f, indent=2, ensure_ascii=False)

    print(f"\nResults written to:")
    print(f"  {md_path}")
    print(f"  {json_path}")
    print(f"\n{report_md}")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Excel Automation Benchmark")
    parser.add_argument("--all", action="store_true", help="Run all available backends")
    parser.add_argument("--backends", default="pywin32,vba",
                        help="Comma-separated backends to test (pywin32,vba,rust,openxml)")
    parser.add_argument("--rust-exe", help="Path to Rust excel-ops.exe")
    parser.add_argument("--openxml-exe", help="Path to OpenXML ExcelOps.exe")
    parser.add_argument("--headed", action="store_true", help="Show Excel window")
    parser.add_argument("--output-dir", default=None, help="Output directory for results")
    parser.add_argument("--warmup", type=int, default=1, help="Warmup rounds (default 1)")
    parser.add_argument("--rounds", type=int, default=3, help="Measurement rounds (default 3)")

    args = parser.parse_args()

    backends_to_run = [b.strip() for b in args.backends.split(",")]
    if args.all:
        backends_to_run = ["pywin32", "vba", "rust", "openxml"]

    output_dir = args.output_dir or os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "..", "..", "benchmarks", "results"
    )
    os.makedirs(output_dir, exist_ok=True)

    workdir = tempfile.mkdtemp(prefix="excel_bench_")
    print(f"Workdir: {workdir}")
    print(f"Backends: {backends_to_run}")
    print(f"Rounds: {args.warmup} warmup + {args.rounds} measured\n")

    all_results = {}

    for bname in backends_to_run:
        print(f"{'='*60}")
        print(f"  Running: {bname}")
        print(f"{'='*60}")

        round_results = []
        total_rounds = args.warmup + args.rounds

        for r in range(total_rounds):
            is_warmup = r < args.warmup
            label = f"warmup {r+1}" if is_warmup else f"round {r+1-args.warmup}"
            print(f"  [{bname}] {label}...", end=" ", flush=True)

            try:
                if bname in ("pywin32", "vba"):
                    backend = COMBackend(bname, visible=args.headed)
                    result = run_benchmark_com(backend, workdir)
                elif bname == "rust":
                    exe = args.rust_exe
                    if not exe:
                        exe = os.path.join(
                            os.path.dirname(os.path.abspath(__file__)),
                            "..", "..", "languages", "rust", "target", "release", "excel-ops.exe"
                        )
                    backend = ProtocolBackend("rust", exe)
                    result = run_benchmark_protocol(backend, workdir)
                elif bname == "openxml":
                    exe = args.openxml_exe
                    if not exe:
                        exe = os.path.join(
                            os.path.dirname(os.path.abspath(__file__)),
                            "..", "..", "languages", "openxml", "bin", "Release", "net10.0", "ExcelOps.exe"
                        )
                    backend = ProtocolBackend("openxml", exe)
                    result = run_benchmark_protocol(backend, workdir)
                else:
                    print(f"Unknown backend: {bname}")
                    continue

                print("done")
                if not is_warmup:
                    round_results.append(result)
            except Exception as e:
                print(f"ERROR: {e}")
                continue

        # Average measured rounds
        if round_results:
            averaged = {}
            for key in round_results[0]:
                totals = [r[key].get("total_ms", 0) for r in round_results if key in r]
                ops_list = [r[key].get("ops", 0) for r in round_results if key in r]
                avg_total = sum(totals) / len(totals) if totals else 0
                avg_ops = round(sum(ops_list) / len(ops_list)) if ops_list else 0
                averaged[key] = {
                    "ops": avg_ops,
                    "total_ms": round(avg_total, 1),
                }
                if avg_ops > 1 and avg_total > 0:
                    averaged[key]["avg_ms"] = round(avg_total / avg_ops, 1)
                # Preserve notes
                if "note" in round_results[0].get(key, {}):
                    averaged[key]["note"] = round_results[0][key]["note"]
            all_results[bname] = averaged
        else:
            print(f"  [{bname}] No successful rounds!")

    if all_results:
        generate_report(all_results, output_dir)
    else:
        print("No results to report.")

    # Cleanup workdir
    try:
        import shutil
        shutil.rmtree(workdir, ignore_errors=True)
    except Exception:
        pass


if __name__ == "__main__":
    main()
