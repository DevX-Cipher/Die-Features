#include <Shlwapi.h>
#include <propvarutil.h>
#include <propsys.h>
#include <shobjidl.h>
#include <memory>
#include <fstream>
#include <sstream>
#include <exception>
#include <vector>

#include "CFileDetailsShellExt.h"
#include "die.h"

#pragma comment(lib, "Shlwapi.lib")
#pragma comment(lib, "Propsys.lib")
#pragma comment(lib, "Version.lib")

const CLSID CLSID_TxtInfoShlExt =
{ 0xD4E6F7A1, 0x9C3B, 0x4F8D, { 0xBE, 0x2E, 0x1A, 0x2C, 0x3D, 0x4E, 0x5F, 0x60 } };

CFileDetailsShellExt::CFileDetailsShellExt() : m_cRef(1)
{
}

STDMETHODIMP CFileDetailsShellExt::QueryInterface(REFIID riid, void** ppvObject)
{
  if (!ppvObject) return E_INVALIDARG;
  *ppvObject = nullptr;

  if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IPersistFile))
  {
    *ppvObject = static_cast<IPersistFile*>(this);
  }
  else if (IsEqualIID(riid, IID_IQueryInfo))
  {
    *ppvObject = static_cast<IQueryInfo*>(this);
  }
  else
  {
    return E_NOINTERFACE;
  }

  AddRef();
  return S_OK;
}

STDMETHODIMP_(ULONG) CFileDetailsShellExt::AddRef()
{
  return InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) CFileDetailsShellExt::Release()
{
  ULONG ulRef = InterlockedDecrement(&m_cRef);
  if (ulRef == 0)
  {
    delete this;
  }
  return ulRef;
}

STDMETHODIMP CFileDetailsShellExt::GetClassID(CLSID* pClassID)
{
  if (!pClassID) return E_INVALIDARG;
  *pClassID = CLSID_TxtInfoShlExt;
  return S_OK;
}

STDMETHODIMP CFileDetailsShellExt::Load(LPCOLESTR wszFilename, DWORD dwMode)
{
  UNREFERENCED_PARAMETER(dwMode);
  m_sFilename = wszFilename;
  return S_OK;
}

std::wstring CFileDetailsShellExt::GetDllFolderPath()
{
  wchar_t path[MAX_PATH] = { 0 };
  HMODULE hModule = NULL;

  if (GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS,
    reinterpret_cast<LPCWSTR>(&CFileDetailsShellExt::GetDllFolderPath), &hModule))
  {
    if (GetModuleFileNameW(hModule, path, MAX_PATH))
    {
      std::wstring fullPath(path);
      size_t pos = fullPath.find_last_of(L"\\/");
      if (pos != std::wstring::npos)
      {
        FreeLibrary(hModule);
        return fullPath.substr(0, pos);
      }
    }
    FreeLibrary(hModule);
  }
  return L"";
}

std::wstring CFileDetailsShellExt::GetFileInformation(const std::wstring& filePath, const std::wstring& pwszDatabase)
{
  try
  {
    std::wstring dllFolderPath = GetDllFolderPath();
    std::wstring databasePath = dllFolderPath.empty() ? L"db" : dllFolderPath + L"\\db";
    std::wstring mutableDatabasePath = databasePath;

    // Load database once
    static bool dbLoaded = false;
    static std::wstring lastDbPath;
    if (!dbLoaded || lastDbPath != mutableDatabasePath)
    {
      DIE_LoadDatabaseW(&mutableDatabasePath[0]);
      dbLoaded = true;
      lastDbPath = mutableDatabasePath;
    }
    std::ifstream file(filePath, std::ios::binary);
    if (!file)
    {
      return L"Failed to open file for scanning.";
    }

    std::vector<char> fileData((std::istreambuf_iterator<char>(file)), std::istreambuf_iterator<char>());
    if (fileData.empty())
    {
      return L"File is empty or unreadable.";
    }

    // Heuristic-only scan, explicitly avoiding recursive scan
    constexpr unsigned int scanFlags = DIE_HEURISTICSCAN;

    wchar_t* resultPtr = DIE_ScanMemoryW(fileData.data(), static_cast<int>(fileData.size()), scanFlags, &mutableDatabasePath[0]);
    std::wstring fullResult = resultPtr ? resultPtr : L"";
    if (resultPtr)
    {
      DIE_FreeMemoryW(resultPtr);
    }

    std::wistringstream stream(fullResult);
    std::wstring line;
    std::wstring filtered;

    while (std::getline(stream, line))
    {
      line.erase(0, line.find_first_not_of(L" \t"));
      if (line.find(L"(Heur)Packer:") == 0)
      {
        filtered += line + L"\n";
        break; 
      }
    }

    return filtered;
  }
  catch (const std::exception& e)
  {
    size_t convertedChars = 0;
    wchar_t wWhat[1024] = { 0 };
    mbstowcs_s(&convertedChars, wWhat, e.what(), strlen(e.what()) + 1);
    return L"Exception occurred while retrieving file information: " + std::wstring(wWhat);
  }
  catch (...)
  {
    return L"An unknown exception occurred while retrieving file information";
  }
}



bool EndsWith(const std::wstring& str, const std::wstring& suffix)
{
  if (str.length() < suffix.length())
    return false;
  return _wcsicmp(str.c_str() + str.length() - suffix.length(), suffix.c_str()) == 0;
}

bool AppendVersionInfo(const std::wstring& filename, std::wstringstream& sTooltip)
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

bool AppendShellProperties(const std::wstring& filename, std::wstringstream& sTooltip)
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


STDMETHODIMP CFileDetailsShellExt::GetInfoTip(DWORD dwFlags, LPWSTR* ppwszTip)
{
  UNREFERENCED_PARAMETER(dwFlags);

  if (!EndsWith(m_sFilename, L".exe") && !EndsWith(m_sFilename, L".dll"))
    return E_FAIL;

  LPMALLOC pMalloc = nullptr;
  if (FAILED(SHGetMalloc(&pMalloc)))
    return E_FAIL;

  std::wstringstream sTooltip;

  if (EndsWith(m_sFilename, L".exe"))
  {
    AppendVersionInfo(m_sFilename, sTooltip);
    AppendShellProperties(m_sFilename, sTooltip);
  }
  else
  {
    if (AppendShellProperties(m_sFilename, sTooltip))
    {
      
    }
    else
    {
     
    }
  }
  std::wstring dbFolder = GetDllFolderPath();
  std::wstring scanResult = GetFileInformation(m_sFilename, dbFolder);
  if (!scanResult.empty())
  {
    if (!sTooltip.str().empty()) sTooltip << L"\n";
    sTooltip << scanResult;
  }

  std::wstring tooltipText = sTooltip.str();
  ULONG cbText = static_cast<ULONG>((tooltipText.length() + 1) * sizeof(wchar_t));
  *ppwszTip = static_cast<LPWSTR>(pMalloc->Alloc(cbText));

  if (*ppwszTip == nullptr)
  {
    pMalloc->Release();
    return E_OUTOFMEMORY;
  }

  wcscpy_s(*ppwszTip, tooltipText.length() + 1, tooltipText.c_str());
  pMalloc->Release();
  return S_OK;
}

STDMETHODIMP CFileDetailsShellExt::CreateInstance(REFIID riid, void** ppv)
{
  if (!ppv) return E_POINTER;

  CFileDetailsShellExt* pExt = new CFileDetailsShellExt();
  if (!pExt) return E_OUTOFMEMORY;

  HRESULT hr = pExt->QueryInterface(riid, ppv);
  pExt->Release();
  return hr;
}