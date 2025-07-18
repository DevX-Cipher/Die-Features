#pragma once
#include <string>
#include <sstream>
#include <map>

// Forward declaration of your class
class CFileDetailsShellExt;

// Correct function pointer typedef
typedef void (*MetadataHandlerFunc)(CFileDetailsShellExt*, const std::wstring&, std::wstringstream&);

// Declare the handler map
extern const std::map<std::wstring, MetadataHandlerFunc> metadataHandlers;

// Utility function
std::wstring GetFileExtension(const std::wstring& filename);

// Handler declarations
void HandleExe(CFileDetailsShellExt* self, const std::wstring& file, std::wstringstream& tip);
void HandleDll(CFileDetailsShellExt* self, const std::wstring& file, std::wstringstream& tip);
