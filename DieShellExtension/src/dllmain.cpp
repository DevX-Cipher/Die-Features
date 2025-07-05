#include "pch.h"
#include "ClassFactory.h"


HINSTANCE g_hInst = NULL;
long        g_cDllRef = 0;



BOOL APIENTRY DllMain(HMODULE hModule, DWORD dwReason, LPVOID lpReserved)
{
    switch (dwReason)
    {
    case DLL_PROCESS_ATTACH:
        g_hInst = hModule;
        DisableThreadLibraryCalls(hModule);
        LogMessage(L"DllMain: DLL_PROCESS_ATTACH");
        break;

    case DLL_PROCESS_DETACH:
        LogMessage(L"DllMain: DLL_PROCESS_DETACH");
        break;
    }
    return TRUE;
}

STDAPI DllCanUnloadNow(void)
{
    LogMessage(L"DllCanUnloadNow called");
    return g_cDllRef > 0 ? S_FALSE : S_OK;
}

HRESULT DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv) {
    LogMessage(L"DllGetClassObject called."); // Log a message

    HRESULT hr = CLASS_E_CLASSNOTAVAILABLE;

    if (IsEqualCLSID(CLSID_DieMenu, rclsid))
    {
        hr = E_OUTOFMEMORY;

        ClassFactory* pClassFactory = new ClassFactory();
        if (pClassFactory)
        {
            hr = pClassFactory->QueryInterface(riid, ppv);
            pClassFactory->Release();
        }
    }

    return hr;
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

    // Register the component.
    hr = RegisterInprocServer(szModule, CLSID_DieMenu,
        L"CulcShellExtContextMenuHandler.ShellEx Class",
        L"Apartment");
    if (SUCCEEDED(hr))
    {
        // Register the context menu handler. The context menu handler is 
        // associated with the .cpp file class.
        hr = RegisterShellExtContextMenuHandler(L"*",
            CLSID_DieMenu,
            L"CulcShellExtContextMenuHandler.ShellEx");
    }
    if (SUCCEEDED(hr))
    {
        hr = RegisterShellExtContextMenuHandler(L".lnk",
            CLSID_DieMenu,
            L"CulcShellExtContextMenuHandler.FileContextMenuExt");
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

    // Unregister the component.
    hr = UnregisterInprocServer(CLSID_DieMenu);
    if (SUCCEEDED(hr))
    {
        // Unregister the context menu handler.
        hr = UnregisterShellExtContextMenuHandler(L"*",
            CLSID_DieMenu);
    }
    if (SUCCEEDED(hr))
    {
        hr = RegisterShellExtContextMenuHandler(L".lnk",
            CLSID_DieMenu,
            L"CulcShellExtContextMenuHandler.FileContextMenuExt");
    }

    return hr;
}
