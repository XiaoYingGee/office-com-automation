// main.cpp - CLI entry point for Excel COM automation tasks E01-E12
//
// Usage: excel_com.exe [E01] [E02] ... [E12]
//   No args = run all tasks.
//   Specify task names to run specific tasks.

#include "excel_com.h"
#include "tasks.h"

#include <iostream>
#include <filesystem>
#include <map>
#include <vector>
#include <string>
#include <functional>
#include <algorithm>
#include <locale>
#include <codecvt>

namespace fs = std::filesystem;

// Task registry
struct TaskEntry {
    std::wstring name;
    std::function<bool(const std::wstring&)> fn;
};

static std::vector<TaskEntry> GetTaskRegistry() {
    return {
        { L"E01", RunE01 },
        { L"E02", RunE02 },
        { L"E03", RunE03 },
        { L"E04", RunE04 },
        { L"E05", RunE05 },
        { L"E06", RunE06 },
        { L"E07", RunE07 },
        { L"E08", RunE08 },
        { L"E09", RunE09 },
        { L"E10", RunE10 },
        { L"E11", RunE11 },
        { L"E12", RunE12 },
    };
}

static std::wstring ToUpper(const std::wstring& s) {
    std::wstring result = s;
    for (auto& ch : result) ch = towupper(ch);
    return result;
}

int wmain(int argc, wchar_t* argv[]) {
    // Determine output directory: <exe_dir>\..\out  or  <src_dir>\out
    fs::path exePath = fs::path(argv[0]).parent_path();
    fs::path outputDir = exePath / "out";
    // If running from src/ directly, use src/out
    if (!fs::exists(outputDir)) {
        outputDir = fs::current_path() / "out";
    }
    fs::create_directories(outputDir);
    std::wstring outStr = outputDir.wstring();

    std::wcout << L"Output directory: " << outStr << std::endl;
    std::wcout << std::wstring(60, L'=') << std::endl;

    auto allTasks = GetTaskRegistry();

    // Determine which tasks to run
    std::vector<TaskEntry> tasksToRun;
    if (argc > 1) {
        for (int i = 1; i < argc; ++i) {
            std::wstring arg = ToUpper(argv[i]);
            for (auto& t : allTasks) {
                if (t.name == arg) {
                    tasksToRun.push_back(t);
                    break;
                }
            }
        }
    } else {
        tasksToRun = allTasks;
    }

    if (tasksToRun.empty()) {
        std::wcout << L"No valid tasks specified. Available: E01-E12" << std::endl;
        return 1;
    }

    // Initialize COM once for all tasks
    CoInitGuard comGuard;

    struct Result {
        std::wstring name;
        bool passed;
        double ms;
    };
    std::vector<Result> results;

    for (auto& task : tasksToRun) {
        std::wcout << std::endl;
        std::wcout << L"--- " << task.name << L" ---" << std::endl;

        HiResTimer timer;
        bool passed = false;
        try {
            passed = task.fn(outStr);
        }
        catch (_com_error& e) {
            std::wcout << L"  COM ERROR: " << (e.ErrorMessage() ? e.ErrorMessage() : L"unknown")
                       << std::endl;
            _bstr_t desc = e.Description();
            if (desc.length() > 0) {
                std::wcout << L"  Description: " << static_cast<const wchar_t*>(desc) << std::endl;
            }
        }
        catch (std::exception& e) {
            std::wcout << L"  EXCEPTION: " << e.what() << std::endl;
        }
        catch (...) {
            std::wcout << L"  UNKNOWN EXCEPTION" << std::endl;
        }
        double elapsed = timer.elapsedMs();
        results.push_back({ task.name, passed, elapsed });
        std::wcout << L"  => " << (passed ? L"PASS" : L"FAIL")
                   << L"  (" << elapsed << L" ms)" << std::endl;
    }

    // Summary table
    std::wcout << std::endl;
    std::wcout << std::wstring(60, L'=') << std::endl;
    std::wcout << L"Task     Result   Time (ms)" << std::endl;
    std::wcout << std::wstring(30, L'-') << std::endl;

    int passCount = 0;
    double totalMs = 0;
    for (auto& r : results) {
        std::wcout << r.name;
        // Pad to 9 chars
        for (size_t i = r.name.size(); i < 9; ++i) std::wcout << L' ';
        std::wcout << (r.passed ? L"PASS" : L"FAIL");
        std::wcout << L"     " << r.ms << std::endl;
        if (r.passed) passCount++;
        totalMs += r.ms;
    }
    std::wcout << std::wstring(30, L'-') << std::endl;
    std::wcout << L"Total: " << passCount << L"/" << results.size()
               << L" passed, " << totalMs << L" ms total" << std::endl;

    return (passCount == static_cast<int>(results.size())) ? 0 : 1;
}
