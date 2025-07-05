#include "pch.h"
#include "ShellEx.h"
#include <VersionHelpers.h>
#include <strsafe.h>
#include <comutil.h>
#include <memory>
#include "resource.h"

#pragma comment( lib, "comsuppwd.lib")

extern HINSTANCE g_hInst;
extern long g_cDllRef;

#define IDM_DISPLAY             0

DieShellMenu::DieShellMenu(void) :m_cRef(1),

    m_pszMenuText(L"Detect It Easy"),
    m_pszVerb("Die"),
    m_pwszVerb(L"Die"),
    m_pszVerbCanonicalName("Die"),
    m_pwszVerbCanonicalName(L"Die"),
    m_pszVerbHelpText("Die.log"),
    m_pwszVerbHelpText(L"Die"),
    m_nFiles(0)
    
{
    InterlockedIncrement(&g_cDllRef);
    // Load the bitmap for the menu item. 
    m_hMenuBmp = LoadImage(g_hInst, MAKEINTRESOURCE(IDB_OK),
        IMAGE_BITMAP, 0, 0, LR_DEFAULTSIZE | LR_LOADTRANSPARENT);
    if (m_hMenuBmp == NULL) {
        DWORD dwError = GetLastError();
        LPVOID lpMsgBuf;
        FormatMessage(
            FORMAT_MESSAGE_ALLOCATE_BUFFER |
            FORMAT_MESSAGE_FROM_SYSTEM |
            FORMAT_MESSAGE_IGNORE_INSERTS,
            NULL,
            dwError,
            MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
            (LPTSTR)&lpMsgBuf,
            0, NULL);
        MessageBox(NULL, (LPCTSTR)lpMsgBuf, TEXT("Error"), MB_OK | MB_ICONERROR);
        LocalFree(lpMsgBuf);
    }
}

DieShellMenu::~DieShellMenu(void)
{
    LogMessage(L"DieShellMenu destructor called");
    if (m_hMenuBmp)
    {
        DeleteObject(m_hMenuBmp);
        m_hMenuBmp = NULL;
    }
    InterlockedDecrement(&g_cDllRef);
    
   
}


IFACEMETHODIMP_(ULONG) DieShellMenu::AddRef()
{
    return InterlockedIncrement(&m_cRef);
}
IFACEMETHODIMP_(ULONG) DieShellMenu::Release()
{
    ULONG cRef = InterlockedDecrement(&m_cRef);
    if (0 == cRef)
    {
        delete this;
    }

    return cRef;
}
IFACEMETHODIMP DieShellMenu::Initialize(LPCITEMIDLIST pidlFolder, LPDATAOBJECT pDataObj, HKEY hKeyProgID)
{
    if (NULL == pDataObj)
    {
        return E_INVALIDARG;
    }

    HRESULT hr = E_FAIL;

    FORMATETC fe = { CF_HDROP, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
    STGMEDIUM stm;

    // The pDataObj pointer contains the objects being acted upon. In this 
    // example, we get an HDROP handle for enumerating the selected files and 
    // folders.
    if (SUCCEEDED(pDataObj->GetData(&fe, &stm)))
    {
        // Get an HDROP handle.
        HDROP hDrop = static_cast<HDROP>(GlobalLock(stm.hGlobal));
        if (hDrop != NULL)
        {
            // Determine how many files are involved in this operation. This 
            // code sample displays the custom context menu item when only 
            // one file is selected. 
            UINT nFiles = DragQueryFile(hDrop, 0xFFFFFFFF, NULL, 0);
            if (nFiles == 1)
            {
                // Get the path of the file.
                if (0 != DragQueryFile(hDrop, 0, m_szSelectedFile,
                    ARRAYSIZE(m_szSelectedFile)))
                {
                    hr = S_OK;
                }
            }

            GlobalUnlock(stm.hGlobal);
        }

        ReleaseStgMedium(&stm);
    }

    // If any value other than S_OK is returned from the method, the context 
    // menu item is not displayed.
    return hr;
}
IFACEMETHODIMP DieShellMenu::QueryInterface(REFIID riid, void** ppv)
{
    static const QITAB qit[] =
    {
        QITABENT(DieShellMenu, IContextMenu),
        QITABENT(DieShellMenu, IShellExtInit),
        { 0 },
    };
    return QISearch(this, qit, riid, ppv);
}

IFACEMETHODIMP DieShellMenu::QueryContextMenu(
    HMENU hMenu, UINT indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags)
{
    if (CMF_DEFAULTONLY & uFlags)
    {
        return MAKE_HRESULT(SEVERITY_SUCCESS, 0, USHORT(0));
    }

    MENUITEMINFO mii = { sizeof(mii) };
    mii.fMask = MIIM_STRING | MIIM_FTYPE | MIIM_ID | MIIM_STATE | MIIM_BITMAP;
    mii.wID = idCmdFirst + IDM_DISPLAY;
    mii.fType = MFT_STRING;
    size_t len = wcslen(m_pszMenuText) + 1;
    LPWSTR pszMenuTextCopy = new WCHAR[len];
    wcscpy_s(pszMenuTextCopy, len, m_pszMenuText);
    mii.dwTypeData = pszMenuTextCopy;
    mii.fState = MFS_ENABLED;

    mii.hbmpItem = static_cast<HBITMAP>(m_hMenuBmp);

    if (!InsertMenuItem(hMenu, indexMenu, TRUE, &mii))
    {
        delete[] pszMenuTextCopy;  // Delete the copy if insertion fails
        return HRESULT_FROM_WIN32(GetLastError());
    }

    // Add a separator.
    MENUITEMINFO sep = { sizeof(sep) };
    sep.fMask = MIIM_TYPE;
    sep.fType = MFT_SEPARATOR;
    if (!InsertMenuItem(hMenu, indexMenu + 1, TRUE, &sep))
    {
        delete[] pszMenuTextCopy;  // Delete the copy if insertion fails
        return HRESULT_FROM_WIN32(GetLastError());
    }
    delete[] pszMenuTextCopy;

    return MAKE_HRESULT(SEVERITY_SUCCESS, 0, USHORT(IDM_DISPLAY + 1));
}


IFACEMETHODIMP DieShellMenu::InvokeCommand(LPCMINVOKECOMMANDINFO pici)
{
    BOOL fUnicode = FALSE;

    HRESULT hr = S_OK;  // Declare hr

    if (pici->cbSize == sizeof(CMINVOKECOMMANDINFOEX))
    {
        if (pici->fMask & CMIC_MASK_UNICODE)
        {
            fUnicode = TRUE;
        }
    }

    if (!fUnicode && HIWORD(pici->lpVerb))
    {
        // Is the verb supported by this context menu extension?
        if (StrCmpIA(pici->lpVerb, m_pszVerb) == 0)
        {
            //OnVerbDisplayFileName(pici->hwnd);
        }
        else
        {
            // If the verb is not recognized by the context menu handler, it 
            // must return E_FAIL to allow it to be passed on to the other 
            // context menu handlers that might implement that verb.
            return E_FAIL;
        }
    }

    // For the Unicode case, if the high-order word is not zero, the 
    // command's verb string is in lpcmi->lpVerbW. 
    else if (fUnicode && HIWORD(((CMINVOKECOMMANDINFOEX*)pici)->lpVerbW))
    {
        // Is the verb supported by this context menu extension?
        if (StrCmpIW(((CMINVOKECOMMANDINFOEX*)pici)->lpVerbW, m_pwszVerb) == 0)
        {
            //OnVerbDisplayFileName(pici->hwnd);
        }
        else
        {
            // If the verb is not recognized by the context menu handler, it 
            // must return E_FAIL to allow it to be passed on to the other 
            // context menu handlers that might implement that verb.
            return E_FAIL;
        }
    }

    // If the command cannot be identified through the verb string, then 
    // check the identifier offset.
    else
    {

        // Handle the identifier offset case
        if (LOWORD(pici->lpVerb) == IDM_DISPLAY)
        {
            // Get the path of the selected file
            TCHAR szSelectedFile[MAX_PATH];
            wcscpy_s(szSelectedFile, m_szSelectedFile);

            // Check if the selected file is a shortcut
            if (PathMatchSpec(szSelectedFile, L"*.lnk"))
            {
                // Resolve the shortcut to get the actual file
                CComPtr<IShellLink> spShellLink;
                hr = spShellLink.CoCreateInstance(CLSID_ShellLink);
                if (SUCCEEDED(hr))
                {
                    CComQIPtr<IPersistFile> spPersistFile;
                    hr = spShellLink->QueryInterface(IID_PPV_ARGS(&spPersistFile));  // Use QueryInterface
                    if (SUCCEEDED(hr))
                    {
                        hr = spPersistFile->Load(szSelectedFile, STGM_READ);
                        if (SUCCEEDED(hr))
                        {
                            hr = spShellLink->Resolve(NULL, SLR_NO_UI);
                            if (SUCCEEDED(hr))
                            {
                                hr = spShellLink->GetPath(szSelectedFile, MAX_PATH, NULL, 0);
                            }
                        }
                    }
                }
            }

            // Get the path to the DLL
            TCHAR szModule[MAX_PATH];
            GetModuleFileName(g_hInst, szModule, ARRAYSIZE(szModule));

            // Remove the DLL name to get the directory
            PathRemoveFileSpec(szModule);

            // Append the program name to get the full path
            PathAppend(szModule, TEXT("die.exe"));

            // Launch the program with the selected file
            ShellExecute(
                pici->hwnd,
                NULL,
                szModule,
                szSelectedFile,
                NULL,
                SW_SHOWNORMAL
            );

            return S_OK;
        }
    }
    return E_FAIL;
}

IFACEMETHODIMP DieShellMenu::GetCommandString(UINT_PTR idCommand,
    UINT uFlags, UINT* pwReserved, LPSTR pszName, UINT cchMax)
{	
    HRESULT hr = E_INVALIDARG;

    if (idCommand == IDM_DISPLAY)
    {
        switch (uFlags)
        {
        case GCS_HELPTEXTW:
            // Only useful for pre-Vista versions of Windows that have a 
            // Status bar.
            hr = StringCchCopy(reinterpret_cast<PWSTR>(pszName), cchMax,
                m_pwszVerbHelpText);
            break;

        case GCS_VERBW:
            // GCS_VERBW is an optional feature that enables a caller to 
            // discover the canonical name for the verb passed in through 
            // idCommand.
            hr = StringCchCopy(reinterpret_cast<PWSTR>(pszName), cchMax,
                m_pwszVerbCanonicalName);
            break;

        default:
            hr = S_OK;
        }
    }

    // If the command (idCommand) is not supported by this context menu 
    // extension handler, return E_INVALIDARG.

    return hr;
}

