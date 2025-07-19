#include <windows.h>
#include <shlobj.h>
#include <string>
#include <comdef.h>
#include "MyFilePreviewHandler.h"
#include "Reg.h"

const CLSID CLSID_PreviewHandler =
{ 0x78A573CA, 0x297E, 0x4D9F, { 0xA5, 0xFC, 0x7F, 0x6E, 0x5E, 0xEA, 0x6F, 0xC9 } };
const GUID APPID_PreviewHandler =
{ 0x2992DE27, 0x3526, 0x48C5, { 0xB7, 0x65, 0xE5, 0x52, 0x78, 0xEC, 0xBE, 0x9D } };

HINSTANCE   g_hInst = NULL;
long        g_cDllRef = 0;


class MyFilePreviewHandlerFactory : public IClassFactory
{
public:
  HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
  {
    if (ppvObject == nullptr)
      return E_INVALIDARG;

    *ppvObject = nullptr;
    if (riid == IID_IUnknown || riid == IID_IClassFactory)
      *ppvObject = static_cast<IClassFactory*>(this);
    else
      return E_NOINTERFACE;

    AddRef();
    return S_OK;
  }

  ULONG STDMETHODCALLTYPE AddRef() override
  {
    return InterlockedIncrement(&refCount);
  }

  ULONG STDMETHODCALLTYPE Release() override
  {
    ULONG count = InterlockedDecrement(&refCount);
    if (count == 0)
      delete this;
    return count;
  }

  HRESULT STDMETHODCALLTYPE CreateInstance(IUnknown* pUnkOuter, REFIID riid, void** ppvObject) override
  {
    if (pUnkOuter != nullptr)
      return CLASS_E_NOAGGREGATION;

    MyFilePreviewHandler* handler = new MyFilePreviewHandler();
    HRESULT hr = handler->QueryInterface(riid, ppvObject);
    handler->Release();
    return hr;
  }

  HRESULT STDMETHODCALLTYPE LockServer(BOOL fLock) override
  {
    return S_OK;
  }

private:
  LONG refCount = 1;
};

BOOL APIENTRY DllMain(HMODULE hModule, DWORD dwReason, LPVOID lpReserved)
{
  switch (dwReason)
  {
  case DLL_PROCESS_ATTACH:
    g_hInst = hModule;
    DisableThreadLibraryCalls(hModule);
    break;
  case DLL_THREAD_ATTACH:
  case DLL_THREAD_DETACH:
  case DLL_PROCESS_DETACH:
    break;
  }
  return TRUE;
}

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, void** ppv)
{
  HRESULT hr = CLASS_E_CLASSNOTAVAILABLE;

  if (IsEqualCLSID(CLSID_PreviewHandler, rclsid))
  {
    hr = E_OUTOFMEMORY;

    MyFilePreviewHandlerFactory* pClassFactory = new MyFilePreviewHandlerFactory();
    if (pClassFactory)
    {
      hr = pClassFactory->QueryInterface(riid, ppv);
      pClassFactory->Release();
    }
  }

  return hr;
}

STDAPI DllCanUnloadNow(void)
{
  return g_cDllRef > 0 ? S_FALSE : S_OK;
}

STDAPI DllRegisterServer(void)
{
  HRESULT hr;

  wchar_t szModule[MAX_PATH];
  if (GetModuleFileName(g_hInst, szModule, ARRAYSIZE(szModule)) == 0)
  {
    hr = HRESULT_FROM_WIN32(GetLastError());
    return hr;
  }

  hr = RegisterInprocServer(szModule, CLSID_PreviewHandler,
    L"CppShellExtPreviewHandler.PreviewHandler Class",
    L"Apartment", APPID_PreviewHandler);
  if (SUCCEEDED(hr))
  {
    hr = RegisterShellExtPreviewHandler(L".exe", CLSID_PreviewHandler,
      L"RecipePreviewHandler");
  }

  return hr;
}

STDAPI DllUnregisterServer(void)
{
  HRESULT hr = S_OK;

  wchar_t szModule[MAX_PATH];
  if (GetModuleFileName(g_hInst, szModule, ARRAYSIZE(szModule)) == 0)
  {
    hr = HRESULT_FROM_WIN32(GetLastError());
    return hr;
  }

  hr = UnregisterInprocServer(CLSID_PreviewHandler,
    APPID_PreviewHandler);
  if (SUCCEEDED(hr))
  {
    hr = UnregisterShellExtPreviewHandler(L".exe",
      CLSID_PreviewHandler);
  }

  return hr;
}