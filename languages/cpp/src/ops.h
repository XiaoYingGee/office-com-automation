#pragma once
// ops.h - Backend-protocol op dispatch entry point for the C++ COM backend.

#include "json.hpp"

// Execute one OpRequest (parsed JSON) and return one OpResponse JSON object.
// Never throws: all errors are converted to {"ok":false,"error":{...}}.
// COM must already be initialized on the calling thread.
nlohmann::json dispatch_op(const nlohmann::json& req);
