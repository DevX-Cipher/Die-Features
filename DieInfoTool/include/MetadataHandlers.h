#pragma once
#include <string>
#include <sstream>
#include <map>

// Function pointer type for extension-specific metadata handlers
class MetadataHandler;

typedef void (*MetadataHandlerFunc)(MetadataHandler*, const std::wstring&, std::wstringstream&);

class MetadataHandler
{
public:
  // Main metadata appenders
  bool AppendVersionInfo(const std::wstring& filename, std::wstringstream& sTooltip);
  bool AppendFileSize(const std::wstring& filename, std::wstringstream& sTooltip);
	bool AppendDateCreated(const std::wstring& filename, std::wstringstream& sTooltip);
  bool AppendSystemFileType(const std::wstring& filename, std::wstringstream& sTooltip);
  bool GetDateModifiedViaPropertyStore(const std::wstring& filename, std::wstringstream& sTooltip);
  bool AppendMetadataForExtension(const std::wstring& filename, std::wstringstream& sTooltip);

  // Utility
  static std::wstring GetFileExtension(const std::wstring& filename);

private:
  static const std::map<std::wstring, MetadataHandlerFunc> metadataHandlers;
  static void HandleTxt(MetadataHandler* self, const std::wstring& file, std::wstringstream& tip);
  static void HandleExe(MetadataHandler* self, const std::wstring& file, std::wstringstream& tip);
  static void HandleDll(MetadataHandler* self, const std::wstring& file, std::wstringstream& tip);
};
