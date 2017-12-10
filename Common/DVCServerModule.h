#pragma once

#include <vector>

typedef std::vector<char> buffer_t;

DWORD OpenDynamicChannel(HANDLE *phFile);
DWORD CallFunctionOverRDP(buffer_t & buffer);
