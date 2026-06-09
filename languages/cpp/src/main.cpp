// main.cpp - Backend-protocol entry point for the C++ COM Excel backend.
//
// Process model (see docs/backend-protocol.md §1):
//   1. Read ONE OpRequest JSON object from stdin (until EOF).
//   2. Initialize COM, dispatch the op, build ONE OpResponse JSON object.
//   3. Write the OpResponse to stdout and exit 0.
//
// Exit code is 0 whenever a response was produced (even ok:false). stderr is
// used for diagnostics only.

#include "excel_com.h"
#include "ops.h"

#include <cstdio>
#include <io.h>
#include <fcntl.h>
#include <iostream>
#include <string>

using json = nlohmann::json;

// Emit a {"ok":false,"error":{...}} response to stdout.
static void emit_error(const char* category, long long code, const std::string& message) {
    json resp = {
        {"ok", false},
        {"error", {{"category", category}, {"code", code}, {"message", message}}},
    };
    std::string s = resp.dump();
    fwrite(s.data(), 1, s.size(), stdout);
    fflush(stdout);
}

int main() {
    // Read all of stdin as raw bytes (binary mode avoids CRLF translation).
    _setmode(_fileno(stdin), _O_BINARY);
    _setmode(_fileno(stdout), _O_BINARY);

    std::string input;
    {
        char buf[8192];
        size_t n;
        while ((n = fread(buf, 1, sizeof(buf), stdin)) > 0) {
            input.append(buf, n);
        }
    }

    json req;
    try {
        req = json::parse(input);
    } catch (const std::exception& e) {
        emit_error("InvalidArg", 0, std::string("bad request json: ") + e.what());
        return 0;
    }

    // Initialize COM for this invocation.
    try {
        CoInitGuard comGuard;
        json resp = dispatch_op(req);
        std::string s = resp.dump();
        fwrite(s.data(), 1, s.size(), stdout);
        fflush(stdout);
    } catch (const _com_error& e) {
        HRESULT hr = e.Error();
        std::string msg;
        _bstr_t desc = e.Description();
        if (desc.length() > 0) msg = static_cast<const char*>(desc);
        else msg = "CoInitialize failed";
        emit_error("ComError", static_cast<long long>(static_cast<int>(hr)), msg);
    } catch (const std::exception& e) {
        emit_error("ComError", 0, std::string("CoInitialize failed: ") + e.what());
    }

    return 0;
}
