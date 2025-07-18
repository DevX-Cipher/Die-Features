#include "MetadataHandlers.h"
#include "CFileDetailsShellExt.h"

#include <Windows.h>
#include <ShObjIdl.h>
#include <Propsys.h>
#include <Propvarutil.h>
#include <memory>
#include <cmath>
#include <iomanip>

const std::map<std::wstring, MetadataHandlerFunc> MetadataHandler::metadataHandlers = {
  { L".exe", &MetadataHandler::HandleExe },
  { L".dll", &MetadataHandler::HandleDll },
	{ L".txt", &MetadataHandler::HandleTxt },
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

std::wstring MetadataHandler::GetFileExtension(const std::wstring& filename)
{
  size_t pos = filename.rfind(L'.');
  if (pos != std::wstring::npos)
    return filename.substr(pos);
  return L"";
}

void MetadataHandler::HandleExe(MetadataHandler* self, const std::wstring& file, std::wstringstream& tip)
{
  bool versionInfoAdded = self->AppendVersionInfo(file, tip);
  bool datecreateAdded = self->AppendDateCreated(file, tip);
	bool fileSizeAdded = self->AppendFileSize(file, tip);

  if (!versionInfoAdded && !datecreateAdded && !fileSizeAdded)
  {
    tip << L"No metadata could be extracted from the executable.\n";
  }
}

void MetadataHandler::HandleDll(MetadataHandler* self, const std::wstring& file, std::wstringstream& tip)
{
  bool versionInfoAdded = self->AppendVersionInfo(file, tip);
  bool fileSizeAdded = self->AppendFileSize(file, tip);
  bool datecreateAdded = self->AppendDateCreated(file, tip);


  if (!versionInfoAdded && !fileSizeAdded && !datecreateAdded)
  {
    tip << L"No metadata could be extracted from the DLL.\n";
  }
}

void MetadataHandler::HandleTxt(MetadataHandler* self, const std::wstring& file, std::wstringstream& tip)
{
  bool typeInfoAdded = self->AppendSystemFileType(file, tip);
  bool fileSizeAdded = self->AppendFileSize(file, tip);
  bool datemodAdded = self->GetDateModifiedViaPropertyStore(file, tip);

  if (!typeInfoAdded && !datemodAdded && !fileSizeAdded)
  {
    tip << L"No metadata could be extracted from the text file.\n";
  }
}

bool MetadataHandler::AppendSystemFileType(const std::wstring& filename, std::wstringstream& sTooltip)
{
  IShellItem* pShellItem = nullptr;
  IPropertyStore* pPropertyStore = nullptr;
  PROPVARIANT pv;
  bool success = false;

  HRESULT hr = SHCreateItemFromParsingName(filename.c_str(), nullptr, IID_PPV_ARGS(&pShellItem));
  if (SUCCEEDED(hr) && pShellItem)
  {
    hr = pShellItem->BindToHandler(nullptr, BHID_PropertyStore, IID_PPV_ARGS(&pPropertyStore));
    if (SUCCEEDED(hr) && pPropertyStore)
    {
      PropVariantInit(&pv);
      hr = pPropertyStore->GetValue(PKEY_ItemTypeText, &pv);
      if (SUCCEEDED(hr) && pv.vt == VT_LPWSTR && pv.pwszVal && wcslen(pv.pwszVal) > 0)
      {
        if (!sTooltip.str().empty()) sTooltip << L"\n";
        sTooltip << L"Type: " << pv.pwszVal;
        success = true;
      }
      PropVariantClear(&pv);
    }
  }

  if (pPropertyStore) pPropertyStore->Release();
  if (pShellItem) pShellItem->Release();
  if (!success)
  {
    SHFILEINFO sfi = { 0 };
    if (SHGetFileInfo(filename.c_str(), 0, &sfi, sizeof(sfi), SHGFI_TYPENAME))
    {
      if (!sTooltip.str().empty()) sTooltip << L"\n";
      sTooltip << L"Type: " << sfi.szTypeName;
      success = true;
    }
  }

  return success;
}

bool MetadataHandler::GetDateModifiedViaPropertyStore(const std::wstring& filename, std::wstringstream& sTooltip)
{
  IShellItem* pShellItem = nullptr;
  IPropertyStore* pPropertyStore = nullptr;
  PROPVARIANT pv;
  bool success = false;

  HRESULT hr = SHCreateItemFromParsingName(filename.c_str(), nullptr, IID_PPV_ARGS(&pShellItem));
  if (SUCCEEDED(hr) && pShellItem)
  {
    hr = pShellItem->BindToHandler(nullptr, BHID_PropertyStore, IID_PPV_ARGS(&pPropertyStore));
    if (SUCCEEDED(hr) && pPropertyStore)
    {
      PropVariantInit(&pv);
      hr = pPropertyStore->GetValue(PKEY_DateModified, &pv);
      if (SUCCEEDED(hr) && pv.vt == VT_FILETIME)
      {
        FILETIME localFileTime;
        if (FileTimeToLocalFileTime(&pv.filetime, &localFileTime))
        {
          SYSTEMTIME sysTime;
          if (FileTimeToSystemTime(&localFileTime, &sysTime))
          {
            std::wstringstream ss;
            ss << std::setfill(L'0')
              << L"Date modified: "
              << std::setw(2) << sysTime.wDay << L"/"
              << std::setw(2) << sysTime.wMonth << L"/"
              << sysTime.wYear << L" "
              << std::setw(2) << ((sysTime.wHour % 12 == 0) ? 12 : sysTime.wHour % 12)
              << L":" << std::setw(2) << sysTime.wMinute
              << L" " << (sysTime.wHour >= 12 ? L"PM" : L"AM");

            if (!sTooltip.str().empty()) sTooltip << L"\n";
            sTooltip << ss.str();
            success = true;
          }
        }
      }
      PropVariantClear(&pv);
    }
  }

  if (pPropertyStore) pPropertyStore->Release();
  if (pShellItem) pShellItem->Release();
  return success;
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

bool MetadataHandler::AppendDateCreated(const std::wstring& filename, std::wstringstream& sTooltip)
{
  IShellItem* pShellItem = nullptr;
  IPropertyStore* pPropertyStore = nullptr;
  PROPVARIANT pv;
  bool success = false;

  HRESULT hr = SHCreateItemFromParsingName(filename.c_str(), nullptr, IID_PPV_ARGS(&pShellItem));
  if (SUCCEEDED(hr) && pShellItem)
  {
    hr = pShellItem->BindToHandler(nullptr, BHID_PropertyStore, IID_PPV_ARGS(&pPropertyStore));
    if (SUCCEEDED(hr) && pPropertyStore)
    {
      PropVariantInit(&pv);
      hr = pPropertyStore->GetValue(PKEY_DateCreated, &pv);
      if (SUCCEEDED(hr) && pv.vt == VT_FILETIME)
      {
        FILETIME ftLocal;
        SYSTEMTIME st;
        WCHAR szDateTime[256], szTime[100];

        if (FileTimeToLocalFileTime(&pv.filetime, &ftLocal) &&
          FileTimeToSystemTime(&ftLocal, &st) &&
          GetDateFormatW(LOCALE_USER_DEFAULT, DATE_SHORTDATE, &st, nullptr, szDateTime, ARRAYSIZE(szDateTime)) > 0 &&
          GetTimeFormatW(LOCALE_USER_DEFAULT, 0, &st, L"h:mm tt", szTime, ARRAYSIZE(szTime)) > 0)
        {
          if (!sTooltip.str().empty()) sTooltip << L"\n";
          sTooltip << L"Date created: " << szDateTime << L" " << szTime;
          success = true;
        }
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
      WCHAR szDateTime[256], szTime[100];

      if (FileTimeToLocalFileTime(&fileInfo.ftCreationTime, &ftLocal) &&
        FileTimeToSystemTime(&ftLocal, &st) &&
        GetDateFormatW(LOCALE_USER_DEFAULT, DATE_SHORTDATE, &st, nullptr, szDateTime, ARRAYSIZE(szDateTime)) > 0 &&
        GetTimeFormatW(LOCALE_USER_DEFAULT, 0, &st, L"h:mm tt", szTime, ARRAYSIZE(szTime)) > 0)
      {
        if (!sTooltip.str().empty()) sTooltip << L"\n";
        sTooltip << L"Date created: " << szDateTime << L" " << szTime;
        success = true;
      }
    }
  }

  if (pPropertyStore) pPropertyStore->Release();
  if (pShellItem) pShellItem->Release();
  return success;
}

bool MetadataHandler::AppendFileSize(const std::wstring& filename, std::wstringstream& sTooltip)
{
  IShellItem* pShellItem = nullptr;
  IPropertyStore* pPropertyStore = nullptr;
  PROPVARIANT pv;
  bool success = false;

  HRESULT hr = SHCreateItemFromParsingName(filename.c_str(), nullptr, IID_PPV_ARGS(&pShellItem));
  if (SUCCEEDED(hr) && pShellItem)
  {
    hr = pShellItem->BindToHandler(nullptr, BHID_PropertyStore, IID_PPV_ARGS(&pPropertyStore));
    if (SUCCEEDED(hr) && pPropertyStore)
    {
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
