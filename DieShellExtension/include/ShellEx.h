
#include "pch.h"


class DieShellMenu :
    public IShellExtInit,
    public IContextMenu
{
public:
   
   

    // IUnknown
    IFACEMETHODIMP QueryInterface(REFIID riid, void** ppv);
    IFACEMETHODIMP_(ULONG) AddRef();
    IFACEMETHODIMP_(ULONG) Release();

    // IShellExtInit
    IFACEMETHODIMP Initialize(LPCITEMIDLIST pidlFolder, LPDATAOBJECT pDataObj, HKEY hKeyProgID);
    

    // IContextMenu
    IFACEMETHODIMP QueryContextMenu(HMENU hMenu, UINT index, UINT uIDFirst, UINT uIDLast, UINT uFlags);
    IFACEMETHODIMP InvokeCommand(LPCMINVOKECOMMANDINFO pici);
    IFACEMETHODIMP GetCommandString(UINT_PTR idCommand, UINT uFlags, UINT* pwReserved, LPSTR pszName, UINT cchMax);

    DieShellMenu(void);

    

protected:
    ~DieShellMenu(void);

private:
   
    wchar_t m_szSelectedFile[MAX_PATH];
    PCWSTR m_pszMenuText;
   
    HANDLE m_hMenuBmp;
    PCSTR m_pszVerb;
    PCWSTR m_pwszVerb;
    PCSTR m_pszVerbCanonicalName;
    PCWSTR m_pwszVerbCanonicalName;
    PCSTR m_pszVerbHelpText;
    PCWSTR m_pwszVerbHelpText;
    LPWSTR m_pwszLogFileName;
	 
	UINT m_nFiles;
	
 
    long m_cRef;
    
};

