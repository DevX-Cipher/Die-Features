// MyFilePreviewHandler.h
#pragma once

#include <shlobj.h>    
#include <atlbase.h>   
#include <vector>      
#include <string>      

extern "C" __declspec(dllimport) int __cdecl DIE_VB_ScanFile(const wchar_t* pwszFileName, int nFlags, const wchar_t* pwszDatabase, wchar_t* pwszBuffer, int nBufferSize);

std::wstring GetFileInformation(const std::wstring& filePath, const std::wstring& pwszDatabase);

class MyFilePreviewHandler :
  public IUnknown,
  public IInitializeWithStream,
  public IPreviewHandler,
  public IObjectWithSite 
{
public:
  virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override;
  virtual ULONG STDMETHODCALLTYPE AddRef() override;
  virtual ULONG STDMETHODCALLTYPE Release() override;

  virtual HRESULT STDMETHODCALLTYPE Initialize(IStream* pstream, DWORD grfMode) override;

  virtual HRESULT STDMETHODCALLTYPE SetWindow(HWND hwnd, const RECT* prc) override;
  virtual HRESULT STDMETHODCALLTYPE SetRect(const RECT* prc) override;
  virtual HRESULT STDMETHODCALLTYPE DoPreview() override;
  virtual HRESULT STDMETHODCALLTYPE Unload() override;
  virtual HRESULT STDMETHODCALLTYPE SetFocus() override;
  virtual HRESULT STDMETHODCALLTYPE QueryFocus(HWND* phwnd) override;
  virtual HRESULT STDMETHODCALLTYPE TranslateAccelerator(MSG* pmsg) override;

  virtual HRESULT STDMETHODCALLTYPE SetSite(IUnknown* pUnkSite) override;
  virtual HRESULT STDMETHODCALLTYPE GetSite(REFIID riid, void** ppvSite) override;

  MyFilePreviewHandler();
  ~MyFilePreviewHandler();

private:

  std::wstring filePath;
  LONG m_cRef;                
  HWND m_hwndParent;          
  HWND m_hwndEdit;            
  RECT m_rcView;              
  IStream* stream;
  IUnknown* site;
};
