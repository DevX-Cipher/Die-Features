#include "MetadataHandlers.h"
#include "CFileDetailsShellExt.h"

#include <Windows.h>
#include <ShObjIdl.h>
#include <Propsys.h>
#include <Propvarutil.h>
#include <memory>
#include <cmath>

const std::map<std::wstring, MetadataHandlerFunc> MetadataHandler::metadataHandlers = {
  { L".exe", &MetadataHandler::HandleExe },
  { L".dll", &MetadataHandler::HandleDll },
};

// Dispatcher
bool MetadataHandler::AppendMetadataForExtension(const std::wstring& filename, std::wstringstream& sTooltip)
{
  std::wstring extension = GetFileExtension(filename);
  std::map<std::wstring, MetadataHandlerFunc>::const_iterator it = metadataHandlers.find(extension);
  if (it != metadataHandlers.end())
  {
    it->second(this, filename, sTooltip);
    return true;
  }
  return false;
}

// Utility
std::wstring MetadataHandler::GetFileExtension(const std::wstring& filename)
{
  size_t pos = filename.rfind(L'.');
  if (pos != std::wstring::npos)
    return filename.substr(pos);
  return L"";
}

// Internal handlers
void MetadataHandler::HandleExe(MetadataHandler* self, const std::wstring& file, std::wstringstream& tip)
{
  self->AppendVersionInfo(file, tip);
  self->AppendShellProperties(file, tip);
}

void MetadataHandler::HandleDll(MetadataHandler* self, const std::wstring& file, std::wstringstream& tip)
{
  self->AppendShellProperties(file, tip);
}

bool MetadataHandler::AppendVersionInfo(const std::wstring& filename, std::wstringstream& sTooltip)
{
  DWORD dwHandle = 0;
  DWORD dwSize = GetFileVersionInfoSizeW(filename.c_str(), &dwHandle);
  if (dwSize == 0)
    return false;

  std::unique_ptr<BYTE[]> pVersionInfoBuffer(new BYTE[dwSize]);
  if (!GetFileVersionInfoW(filename.c_str(), 0, dwSize, pVersionInfoBuffer.get()))
    return false;

  LPVOID lpBuffer = nullptr;
  UINT uLen = 0;
  bool bFirstItem = true;

  auto AppendInfo = [&](LPCWSTR key, LPCWSTR label) {
    if (VerQueryValueW(pVersionInfoBuffer.get(), key, &lpBuffer, &uLen) && uLen > 0)
    {
      if (!bFirstItem) sTooltip << L"\n";
      sTooltip << label << static_cast<wchar_t*>(lpBuffer);
      bFirstItem = false;
    }
    };
  AppendInfo(L"\\StringFileInfo\\040904B0\\FileDescription", L"Description: ");
  AppendInfo(L"\\StringFileInfo\\040904B0\\CompanyName", L"Company: ");
  AppendInfo(L"\\StringFileInfo\\040904B0\\FileVersion", L"Version: ");

  return true;
}

bool MetadataHandler::AppendShellProperties(const std::wstring& filename, std::wstringstream& sTooltip)
{
  IShellItem* pShellItem = nullptr;
  IPropertyStore* pPropertyStore = nullptr;
  PROPVARIANT pv;
  HRESULT hr = SHCreateItemFromParsingName(filename.c_str(), nullptr, IID_PPV_ARGS(&pShellItem));
  bool success = false;

  if (SUCCEEDED(hr))
  {
    hr = pShellItem->BindToHandler(nullptr, BHID_PropertyStore, IID_PPV_ARGS(&pPropertyStore));
    if (SUCCEEDED(hr))
    {
      PropVariantInit(&pv);
      hr = pPropertyStore->GetValue(PKEY_DateCreated, &pv);
      if (SUCCEEDED(hr) && pv.vt == VT_FILETIME)
      {
        FILETIME ftLocal;
        SYSTEMTIME st;
        WCHAR szDateTime[256];

        if (FileTimeToLocalFileTime(&pv.filetime, &ftLocal) &&
          FileTimeToSystemTime(&ftLocal, &st))
        {
          if (GetDateFormatW(LOCALE_USER_DEFAULT, DATE_SHORTDATE, &st, nullptr, szDateTime, ARRAYSIZE(szDateTime)) > 0)
          {
            WCHAR szTime[100];
            if (GetTimeFormatW(LOCALE_USER_DEFAULT, 0, &st, nullptr, szTime, ARRAYSIZE(szTime)) > 0)
            {
              if (!sTooltip.str().empty()) sTooltip << L"\n";
              sTooltip << L"Date created: " << szDateTime << L" " << szTime;
              success = true;
            }
          }
        }
      }
      PropVariantClear(&pv);

      PropVariantInit(&pv);
      hr = pPropertyStore->GetValue(PKEY_Size, &pv);
      if (SUCCEEDED(hr) && pv.vt == VT_UI8)
      {
        ULONGLONG fileSize = pv.uhVal.QuadPart;
        WCHAR szSize[100];

        if (fileSize < 1024)
          swprintf_s(szSize, ARRAYSIZE(szSize), L"%llu bytes", fileSize);
        else if (fileSize < 1024 * 1024)
          swprintf_s(szSize, ARRAYSIZE(szSize), L"%llu KB", fileSize / 1024);
        else if (fileSize < 1024ULL * 1024 * 1024)
        {
          double mb = fileSize / (1024.0 * 1024.0);
          mb = std::floor(mb * 10.0) / 10.0;
          swprintf_s(szSize, ARRAYSIZE(szSize), L"%.1f MB", mb);
        }
        else
        {
          double gb = fileSize / (1024.0 * 1024.0 * 1024.0);
          gb = std::floor(gb * 100.0) / 100.0;
          swprintf_s(szSize, ARRAYSIZE(szSize), L"%.2f GB", gb);
        }

        if (!sTooltip.str().empty()) sTooltip << L"\n";
        sTooltip << L"Size: " << szSize;
        success = true;
      }
      PropVariantClear(&pv);
    }
  }

  if (!success)
  {
    WIN32_FILE_ATTRIBUTE_DATA fileInfo;
    if (GetFileAttributesEx(filename.c_str(), GetFileExInfoStandard, &fileInfo))
    {
      FILETIME ftLocal;
      SYSTEMTIME st;
      WCHAR szDateTime[256];

      if (FileTimeToLocalFileTime(&fileInfo.ftCreationTime, &ftLocal) &&
        FileTimeToSystemTime(&ftLocal, &st))
      {
        if (GetDateFormatW(LOCALE_USER_DEFAULT, DATE_SHORTDATE, &st, nullptr, szDateTime, ARRAYSIZE(szDateTime)) > 0)
        {
          WCHAR szTime[100];
          if (GetTimeFormatW(LOCALE_USER_DEFAULT, 0, &st, nullptr, szTime, ARRAYSIZE(szTime)) > 0)
          {
            if (!sTooltip.str().empty()) sTooltip << L"\n";
            sTooltip << L"Date created: " << szDateTime << L" " << szTime;
            success = true;
          }
        }
      }

      ULONGLONG fileSize = ((ULONGLONG)fileInfo.nFileSizeHigh << 32) | fileInfo.nFileSizeLow;
      WCHAR szSize[100];

      if (fileSize < 1024)
        swprintf_s(szSize, ARRAYSIZE(szSize), L"%llu bytes", fileSize);
      else if (fileSize < 1024 * 1024)
        swprintf_s(szSize, ARRAYSIZE(szSize), L"%llu KB", fileSize / 1024);
      else if (fileSize < 1024ULL * 1024 * 1024)
      {
        double mb = fileSize / (1024.0 * 1024.0);
        mb = std::floor(mb * 10.0) / 10.0;
        swprintf_s(szSize, ARRAYSIZE(szSize), L"%.1f MB", mb);
      }
      else
      {
        double gb = fileSize / (1024.0 * 1024.0 * 1024.0);
        gb = std::floor(gb * 100.0) / 100.0;
        swprintf_s(szSize, ARRAYSIZE(szSize), L"%.2f GB", gb);
      }

      if (!sTooltip.str().empty()) sTooltip << L"\n";
      sTooltip << L"Size: " << szSize;
      success = true;
    }
  }

  if (pPropertyStore) pPropertyStore->Release();
  if (pShellItem) pShellItem->Release();
  return success;
}