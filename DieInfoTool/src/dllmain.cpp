#include <windows.h>
#include <shlobj.h>
#include <tchar.h>
#include <VersionHelpers.h>

#include "dllmain.h"
#include "ClassFactory.h"
#include "CFileDetailsShellExt.h"

const TCHAR* g_fileTypes[] = {
    TEXT(".exe"),
    TEXT("dllfile"),
    TEXT(".txt"),
};

BOOL WINAPI DllMain(HINSTANCE hInstance, DWORD dwReason, LPVOID /*lpReserved*/)
{
  if (dwReason == DLL_PROCESS_ATTACH)
  {
    g_hModule = hInstance;
    DisableThreadLibraryCalls(hInstance);
  }
  return TRUE;
}

STDAPI DllCanUnloadNow(void)
{
  return (g_cDllRef == 0) ? S_OK : S_FALSE;
}

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv)
{
  if (rclsid != CLSID_TxtInfoShlExt)
    return CLASS_E_CLASSNOTAVAILABLE;

  ClassFactory* factory = new ClassFactory();
  if (!factory) return E_OUTOFMEMORY;

  HRESULT hr = factory->QueryInterface(riid, ppv);
  factory->Release();
  return hr;
}

HRESULT RegisterShellExtensionForFileType(const TCHAR* fileType, const TCHAR* clsid)
{
  TCHAR keyPath[MAX_PATH];
  _stprintf_s(keyPath, MAX_PATH, TEXT("%s\\shellex\\{00021500-0000-0000-C000-000000000046}"), fileType);

  HKEY hKey;
  LONG lRet = RegCreateKeyEx(HKEY_CLASSES_ROOT, keyPath,
    0, NULL, REG_OPTION_NON_VOLATILE, KEY_SET_VALUE, NULL, &hKey, NULL);

  if (lRet != ERROR_SUCCESS)
    return HRESULT_FROM_WIN32(lRet);

  lRet = RegSetValueEx(hKey, NULL, 0, REG_SZ,
    (const BYTE*)clsid, (DWORD)(_tcslen(clsid) + 1) * sizeof(TCHAR));

  RegCloseKey(hKey);
  return HRESULT_FROM_WIN32(lRet);
}

HRESULT UnregisterShellExtensionForFileType(const TCHAR* fileType)
{
  TCHAR keyPath[MAX_PATH];
  _stprintf_s(keyPath, MAX_PATH, TEXT("%s\\shellex\\{00021500-0000-0000-C000-000000000046}"), fileType);

  LONG lRet = RegDeleteKey(HKEY_CLASSES_ROOT, keyPath);
  return HRESULT_FROM_WIN32(lRet);
}


STDAPI DllRegisterServer(void)
{
  const TCHAR* szCLSID = TEXT("{D4E6F7A1-9C3B-4F8D-BE2E-1A2C3D4E5F60}");
  const TCHAR* szDesc = TEXT("CTxtInfoShlExt extension");

  // Register in Approved Shell Extensions
  HKEY hKey;
  LONG lRet = RegOpenKeyEx(HKEY_LOCAL_MACHINE,
    TEXT("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Shell Extensions\\Approved"),
    0, KEY_SET_VALUE, &hKey);

  if (lRet != ERROR_SUCCESS)
    return E_ACCESSDENIED;

  lRet = RegSetValueEx(hKey, szCLSID, 0, REG_SZ,
    (const BYTE*)szDesc, (DWORD)(_tcslen(szDesc) + 1) * sizeof(TCHAR));
  RegCloseKey(hKey);

  if (lRet != ERROR_SUCCESS)
    return E_ACCESSDENIED;

  // Register CLSID
  HKEY hClsidKey;
  lRet = RegCreateKeyEx(HKEY_CLASSES_ROOT,
    TEXT("CLSID\\{D4E6F7A1-9C3B-4F8D-BE2E-1A2C3D4E5F60}"),
    0, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hClsidKey, NULL);

  if (lRet != ERROR_SUCCESS)
    return E_ACCESSDENIED;

  const TCHAR* szFriendlyName = TEXT("CTxtInfoShlExt");
  RegSetValueEx(hClsidKey, NULL, 0, REG_SZ,
    (const BYTE*)szFriendlyName, (DWORD)(_tcslen(szFriendlyName) + 1) * sizeof(TCHAR));

  HKEY hInprocKey;
  lRet = RegCreateKeyEx(hClsidKey, TEXT("InprocServer32"),
    0, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hInprocKey, NULL);

  if (lRet == ERROR_SUCCESS)
  {
    TCHAR szModule[MAX_PATH];
    GetModuleFileName(g_hModule, szModule, MAX_PATH);

    RegSetValueEx(hInprocKey, NULL, 0, REG_SZ,
      (const BYTE*)szModule, (DWORD)(_tcslen(szModule) + 1) * sizeof(TCHAR));

    const TCHAR* szThreadingModel = TEXT("Apartment");
    RegSetValueEx(hInprocKey, TEXT("ThreadingModel"), 0, REG_SZ,
      (const BYTE*)szThreadingModel, (DWORD)(_tcslen(szThreadingModel) + 1) * sizeof(TCHAR));

    RegCloseKey(hInprocKey);
  }

  RegCloseKey(hClsidKey);

  // Register for each file type
  for (int i = 0; i < _countof(g_fileTypes); ++i)
  {
    HRESULT hr = RegisterShellExtensionForFileType(g_fileTypes[i], szCLSID);
    if (FAILED(hr))
      return hr;
  }

  return S_OK;
}

STDAPI DllUnregisterServer(void)
{
  const TCHAR* szCLSID = TEXT("{D4E6F7A1-9C3B-4F8D-BE2E-1A2C3D4E5F60}");

  // Remove from Approved Shell Extensions
  HKEY hKey;
  LONG lRet = RegOpenKeyEx(HKEY_LOCAL_MACHINE,
    TEXT("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Shell Extensions\\Approved"),
    0, KEY_SET_VALUE, &hKey);

  if (lRet == ERROR_SUCCESS)
  {
    RegDeleteValue(hKey, szCLSID);
    RegCloseKey(hKey);
  }

  // Unregister for each file type
  for (int i = 0; i < _countof(g_fileTypes); ++i)
  {
    UnregisterShellExtensionForFileType(g_fileTypes[i]);
  }

  // Remove CLSID entries
  RegDeleteKey(HKEY_CLASSES_ROOT,
    TEXT("CLSID\\{D4E6F7A1-9C3B-4F8D-BE2E-1A2C3D4E5F60}\\InprocServer32"));
  RegDeleteKey(HKEY_CLASSES_ROOT,
    TEXT("CLSID\\{D4E6F7A1-9C3B-4F8D-BE2E-1A2C3D4E5F60}"));

  return S_OK;
}


