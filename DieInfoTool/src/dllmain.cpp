#include <windows.h>
#include <shlobj.h>
#include <tchar.h>
#include <VersionHelpers.h>

#include "dllmain.h"
#include "ClassFactory.h"
#include "CFileDetailsShellExt.h"



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

STDAPI DllRegisterServer(void)
{
  const TCHAR* szCLSID = TEXT("{D4E6F7A1-9C3B-4F8D-BE2E-1A2C3D4E5F60}");
  const TCHAR* szDesc = TEXT("CTxtInfoShlExt extension");

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

  HKEY hClsidKey;
  lRet = RegCreateKeyEx(HKEY_CLASSES_ROOT,
    TEXT("CLSID\\{D4E6F7A1-9C3B-4F8D-BE2E-1A2C3D4E5F60}"),
    0, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hClsidKey, NULL);

  if (lRet == ERROR_SUCCESS)
  {
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
  }

  HKEY hExeInfoTipKey;
  lRet = RegCreateKeyEx(HKEY_CLASSES_ROOT,
    TEXT(".exe\\shellex\\{00021500-0000-0000-C000-000000000046}"),
    0, NULL, REG_OPTION_NON_VOLATILE, KEY_SET_VALUE, NULL, &hExeInfoTipKey, NULL);

  if (lRet == ERROR_SUCCESS)
  {
    RegSetValueEx(hExeInfoTipKey, NULL, 0, REG_SZ,
      (const BYTE*)szCLSID, (DWORD)(_tcslen(szCLSID) + 1) * sizeof(TCHAR));
    RegCloseKey(hExeInfoTipKey);
  }

  HKEY hDllInfoTipKey;
  lRet = RegCreateKeyEx(HKEY_CLASSES_ROOT,
    TEXT("dllfile\\shellex\\{00021500-0000-0000-C000-000000000046}"),
    0, NULL, REG_OPTION_NON_VOLATILE, KEY_SET_VALUE, NULL, &hDllInfoTipKey, NULL);

  if (lRet == ERROR_SUCCESS)
  {
    RegSetValueEx(hDllInfoTipKey, NULL, 0, REG_SZ,
      (const BYTE*)szCLSID, (DWORD)(_tcslen(szCLSID) + 1) * sizeof(TCHAR));
    RegCloseKey(hDllInfoTipKey);
  }

  return S_OK;
}

STDAPI DllUnregisterServer(void)
{
  const TCHAR* szCLSID = TEXT("{D4E6F7A1-9C3B-4F8D-BE2E-1A2C3D4E5F60}");

  HKEY hKey;
  LONG lRet = RegOpenKeyEx(HKEY_LOCAL_MACHINE,
    TEXT("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Shell Extensions\\Approved"),
    0, KEY_SET_VALUE, &hKey);

  if (lRet == ERROR_SUCCESS)
  {
    RegDeleteValue(hKey, szCLSID);
    RegCloseKey(hKey);
  }

  RegDeleteKey(HKEY_CLASSES_ROOT,
    TEXT(".exe\\shellex\\{00021500-0000-0000-C000-000000000046}"));

  RegDeleteKey(HKEY_CLASSES_ROOT,
    TEXT("dllfile\\shellex\\{00021500-0000-0000-C000-000000000046}"));

  RegDeleteKey(HKEY_CLASSES_ROOT,
    TEXT("CLSID\\{D4E6F7A1-9C3B-4F8D-BE2E-1A2C3D4E5F60}\\InprocServer32"));
  RegDeleteKey(HKEY_CLASSES_ROOT,
    TEXT("CLSID\\{D4E6F7A1-9C3B-4F8D-BE2E-1A2C3D4E5F60}"));

  return S_OK;
}

