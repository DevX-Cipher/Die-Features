#pragma once

#include <windows.h>
#include <shlobj.h>
#include "CFileDetailsShellExt.h"

class ClassFactory : public IClassFactory
{
public:
  ClassFactory() : m_refCount(1) { InterlockedIncrement(&g_cDllRef); }
  ~ClassFactory() { InterlockedDecrement(&g_cDllRef); }

  // IUnknown
  STDMETHODIMP QueryInterface(REFIID riid, void** ppv)
  {
    if (!ppv) return E_POINTER;
    if (riid == IID_IUnknown || riid == IID_IClassFactory)
    {
      *ppv = static_cast<IClassFactory*>(this);
      AddRef();
      return S_OK;
    }
    *ppv = nullptr;
    return E_NOINTERFACE;
  }

  STDMETHODIMP_(ULONG) AddRef() { return InterlockedIncrement(&m_refCount); }
  STDMETHODIMP_(ULONG) Release()
  {
    ULONG res = InterlockedDecrement(&m_refCount);
    if (res == 0) delete this;
    return res;
  }

  STDMETHODIMP CreateInstance(IUnknown* pUnkOuter, REFIID riid, void** ppv)
  {
    if (pUnkOuter) return CLASS_E_NOAGGREGATION;
    return CFileDetailsShellExt::CreateInstance(riid, ppv);
  }

  STDMETHODIMP LockServer(BOOL fLock)
  {
    if (fLock)
      InterlockedIncrement(&g_cDllRef);
    else
      InterlockedDecrement(&g_cDllRef);
    return S_OK;
  }

private:
  LONG m_refCount;
}; 