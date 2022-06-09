#pragma once
#include <emscripten.h>

void throwError(const std::string& msg)
{
    EM_ASM_({ throw UTF8ToString($0); }, msg.c_str());
};
