#pragma once

#include <ShlObj.h>
#include <string>
void LogMessage(const std::wstring& message); // Log Message 
extern "C" __declspec(dllimport) int __cdecl DIE_VB_ScanFile(const wchar_t* pwszFileName, int nFlags, const wchar_t* pwszDatabase, wchar_t* pwszBuffer, int nBufferSize);
std::wstring GetFileInformation(const std::wstring& filePath, const std::wstring& pwszDatabase);
