// main.rs - CLI entry point for Excel COM automation tasks E01-E12.
//
// Usage: excel-com[.exe] [E01] [E02] ... [E12]
//   No args  = run all tasks.
//   Specify task names (case-insensitive) to run a subset.

mod com;
mod excel;
mod tasks;

use std::env;
use std::path::PathBuf;
use std::time::Instant;

use com::ComGuard;

type TaskFn = fn(&str) -> windows::core::Result<bool>;

struct TaskEntry {
    name: &'static str,
    func: TaskFn,
}

fn task_registry() -> Vec<TaskEntry> {
    vec![
        TaskEntry { name: "E01", func: tasks::run_e01 },
        TaskEntry { name: "E02", func: tasks::run_e02 },
        TaskEntry { name: "E03", func: tasks::run_e03 },
        TaskEntry { name: "E04", func: tasks::run_e04 },
        TaskEntry { name: "E05", func: tasks::run_e05 },
        TaskEntry { name: "E06", func: tasks::run_e06 },
        TaskEntry { name: "E07", func: tasks::run_e07 },
        TaskEntry { name: "E08", func: tasks::run_e08 },
        TaskEntry { name: "E09", func: tasks::run_e09 },
        TaskEntry { name: "E10", func: tasks::run_e10 },
        TaskEntry { name: "E11", func: tasks::run_e11 },
        TaskEntry { name: "E12", func: tasks::run_e12 },
    ]
}

fn determine_output_dir() -> PathBuf {
    // Prefer <exe_dir>/out; fall back to <cwd>/out.
    if let Ok(exe) = env::current_exe() {
        if let Some(dir) = exe.parent() {
            let candidate = dir.join("out");
            if candidate.exists() {
                return candidate;
            }
            // exe dir exists but no out/ yet — still prefer it for fresh runs
            let _ = std::fs::create_dir_all(&candidate);
            if candidate.exists() {
                return candidate;
            }
        }
    }
    env::current_dir().unwrap_or_else(|_| PathBuf::from(".")).join("out")
}

fn main() {
    let output_dir = determine_output_dir();
    let _ = std::fs::create_dir_all(&output_dir);
    let out_str = output_dir.to_string_lossy().into_owned();

    println!("Output directory: {}", out_str);
    println!("{}", "=".repeat(60));

    let all_tasks = task_registry();

    // Select tasks
    let args: Vec<String> = env::args().skip(1).collect();
    let tasks_to_run: Vec<&TaskEntry> = if args.is_empty() {
        all_tasks.iter().collect()
    } else {
        let wanted: Vec<String> = args.iter().map(|a| a.to_uppercase()).collect();
        all_tasks
            .iter()
            .filter(|t| wanted.iter().any(|w| w == t.name))
            .collect()
    };

    if tasks_to_run.is_empty() {
        println!("No valid tasks specified. Available: E01-E12");
        std::process::exit(1);
    }

    // Initialize COM once for all tasks
    let _com = match ComGuard::new() {
        Ok(g) => g,
        Err(e) => {
            eprintln!("CoInitializeEx failed: {}", e.message());
            std::process::exit(1);
        }
    };

    struct TaskResult {
        name: &'static str,
        passed: bool,
        ms: f64,
    }
    let mut results: Vec<TaskResult> = Vec::new();

    for task in &tasks_to_run {
        println!();
        println!("--- {} ---", task.name);

        let timer = Instant::now();
        let passed = match (task.func)(&out_str) {
            Ok(p) => p,
            Err(e) => {
                println!("  COM ERROR: hr={:?} {}", e.code(), e.message());
                false
            }
        };
        let elapsed = timer.elapsed().as_secs_f64() * 1000.0;
        results.push(TaskResult { name: task.name, passed, ms: elapsed });
        println!(
            "  => {}  ({:.1} ms)",
            if passed { "PASS" } else { "FAIL" },
            elapsed
        );
    }

    // Summary table
    println!();
    println!("{}", "=".repeat(60));
    println!("Task     Result   Time (ms)");
    println!("{}", "-".repeat(30));

    let mut pass_count = 0;
    let mut total_ms = 0.0;
    for r in &results {
        println!(
            "{:<9}{:<9}{:.1}",
            r.name,
            if r.passed { "PASS" } else { "FAIL" },
            r.ms
        );
        if r.passed {
            pass_count += 1;
        }
        total_ms += r.ms;
    }
    println!("{}", "-".repeat(30));
    println!(
        "Total: {}/{} passed, {:.1} ms total",
        pass_count,
        results.len(),
        total_ms
    );

    std::process::exit(if pass_count == results.len() { 0 } else { 1 });
}
