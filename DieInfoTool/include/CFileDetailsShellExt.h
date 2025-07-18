#pragma once

#include <shlobj.h>   // Required for IPersistFile, IQueryInfo, SHGetMalloc, etc.
#include <string>     // Required for std::wstring
#include <windows.h>  // Required for HMODULE, GetModuleHandleExW, GetModuleFileNameW, InterlockedIncrement/Decrement
#include <propkey.h>  // Required for PKEY_DateCreated, PKEY_Size (even if just for forward declaration context)

// Forward declaration for the external function.
// This is a placeholder. You would replace this with the actual declaration
// from your DIE_VB library. Ensure the calling convention (__stdcall) is correct.

// External declaration of the COM CLSID for the shell extension.
extern const CLSID CLSID_TxtInfoShlExt;

// CFileDetailsShellExt class declaration.
// This class implements IPersistFile and IQueryInfo interfaces for a shell extension.
class CFileDetailsShellExt : public IPersistFile, public IQueryInfo
{
public:
  // Constructor.
  CFileDetailsShellExt();

  // IUnknown methods (implemented in the .cpp file).
  STDMETHODIMP QueryInterface(REFIID riid, void** ppvObject);
  STDMETHODIMP_(ULONG) AddRef();
  STDMETHODIMP_(ULONG) Release();

  // IPersist method (implemented in the .cpp file).
  STDMETHODIMP GetClassID(CLSID* pClassID);

  // IPersistFile methods (some implemented, some returning E_NOTIMPL).
  STDMETHODIMP IsDirty() { return E_NOTIMPL; } // Not implemented for this shell extension.
  STDMETHODIMP Load(LPCOLESTR pszFileName, DWORD dwMode); // Loads the file path.
  STDMETHODIMP Save(LPCOLESTR, BOOL) { return E_NOTIMPL; } // Not implemented.
  STDMETHODIMP SaveCompleted(LPCOLESTR) { return E_NOTIMPL; } // Not implemented.
  STDMETHODIMP GetCurFile(LPOLESTR*) { return E_NOTIMPL; } // Not implemented.
  static STDMETHODIMP CreateInstance(REFIID riid, void** ppv);
  // IQueryInfo methods (implemented in the .cpp file).
  STDMETHODIMP GetInfoTip(DWORD dwFlags, LPWSTR* ppwszTip); // Generates the info tip (tooltip).
  STDMETHODIMP GetInfoFlags(DWORD* pdwFlags) { *pdwFlags = 0; return S_OK; } // Returns default flags.

  // Static helper method to get the DLL's folder path.
  static std::wstring GetDllFolderPath();

  // Private helper method to get file information using an external scanner.
  std::wstring GetFileInformation(const std::wstring& filePath, const std::wstring& pwszDatabase);

private:
  LONG m_cRef;        // Reference count for COM object lifetime management.
  std::wstring m_sFilename; // Stores the path of the file being queried.
};