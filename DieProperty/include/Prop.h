#ifndef CSHELLEXT_H
#define CSHELLEXT_H

#include <windows.h>
#include <shlobj.h>

// CShellExt: Implements a shell extension that provides property sheet functionality.
class CShellExt : public IShellExtInit, public IShellPropSheetExt {
protected:
   DWORD m_ObjRefCount; // Reference count for COM object

public:
    TCHAR m_szFile[MAX_PATH]; // File path
    CShellExt(); // Constructor
    virtual ~CShellExt(); // Destructor
    STDMETHOD(QueryInterface) (REFIID riid, LPVOID* ppvObj); // Query for an interface on a COM object
        bool m_bCanDelete; // Add this line

    STDMETHOD_(ULONG, AddRef) (); // Increment the reference count of a COM object
    STDMETHOD_(ULONG, Release) (); // Decrement the reference count of a COM object
    STDMETHOD(Initialize) (LPCITEMIDLIST pidlFolder, LPDATAOBJECT pDataObject, HKEY hProgID); // Initialize the shell extension
    STDMETHOD(AddPages) (LPFNSVADDPROPSHEETPAGE pfnAddPage, LPARAM lParam); // Add property pages to the shell extension
    STDMETHOD(ReplacePage) (UINT uPageID, LPFNSVADDPROPSHEETPAGE lpfnReplaceWith, LPARAM lParam); // Replace property pages of the shell extension
};
#endif // CSHELLEXT_H