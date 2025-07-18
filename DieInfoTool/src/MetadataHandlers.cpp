#include "MetadataHandlers.h"
#include "CFileDetailsShellExt.h" // Include full class definition

void HandleExe(CFileDetailsShellExt* self, const std::wstring& file, std::wstringstream& tip)
{
  self->AppendVersionInfo(file, tip);
  self->AppendShellProperties(file, tip);
}

void HandleDll(CFileDetailsShellExt* self, const std::wstring& file, std::wstringstream& tip)
{
  self->AppendShellProperties(file, tip);
}

const std::map<std::wstring, MetadataHandlerFunc> metadataHandlers = {
  { L".exe", &HandleExe },
  { L".dll", &HandleDll },
};

std::wstring GetFileExtension(const std::wstring& filename)
{
  size_t pos = filename.rfind(L'.');
  if (pos != std::wstring::npos)
    return filename.substr(pos);
  return L"";
}
