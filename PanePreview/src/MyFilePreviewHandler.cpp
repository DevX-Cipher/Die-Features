#include "MyFilePreviewHandler.h"
#include <windows.h>
#include <commctrl.h>
#include <vector>
#include <string>
#include <locale>
#include <cwctype>
#include <codecvt>
#include <memory>
#include <ShlObj.h> 
#include <fstream>     
#include <ctime>       

std::wstring GetFileInformation(const std::wstring& filePath, const std::wstring& pwszDatabase)
{
  try {
    const int nBufferSize = 10000;
    std::wstring sBuffer(nBufferSize, L' '); 

    int nResult = DIE_VB_ScanFile(filePath.c_str(), 0, pwszDatabase.c_str(), &sBuffer[0], nBufferSize - 1);

    std::wstring result;
    if (nResult > 0) {
      sBuffer.resize(nResult);
      result = sBuffer;
    }
    else {
      result = L"Scan failed or no result.";
    }

    return result;
  }
  catch (...) {
    return L"Exception occurred while retrieving file information";
  }
}

MyFilePreviewHandler::MyFilePreviewHandler() : m_cRef(1), m_hwndParent(nullptr), m_hwndEdit(nullptr), stream(nullptr), site(nullptr) {}

MyFilePreviewHandler::~MyFilePreviewHandler()
{
  Unload();
}

HRESULT MyFilePreviewHandler::SetRect(const RECT* prc)
{
  if (!prc)
    return E_INVALIDARG;
  m_rcView = *prc;
  if (m_hwndEdit)
  {
    if (!MoveWindow(m_hwndEdit, m_rcView.left, m_rcView.top, m_rcView.right - m_rcView.left, m_rcView.bottom - m_rcView.top, TRUE))
      return E_FAIL;
  }
  return S_OK;
}

HRESULT MyFilePreviewHandler::QueryInterface(REFIID riid, void** ppvObject)
{
  if (ppvObject == nullptr)
    return E_INVALIDARG;

  *ppvObject = nullptr;
  if (riid == IID_IUnknown || riid == IID_IPreviewHandler)
    *ppvObject = static_cast<IPreviewHandler*>(this);
  else if (riid == IID_IInitializeWithStream)
    *ppvObject = static_cast<IInitializeWithStream*>(this);
  else if (riid == IID_IObjectWithSite)
    *ppvObject = static_cast<IObjectWithSite*>(this);
  else
    return E_NOINTERFACE;

  AddRef();
  return S_OK;
}

ULONG MyFilePreviewHandler::AddRef()
{
  return InterlockedIncrement(&m_cRef);
}

ULONG MyFilePreviewHandler::Release()
{
  ULONG count = InterlockedDecrement(&m_cRef);
  if (count == 0)
  {
    delete this;
  }
  return count;
}

HRESULT MyFilePreviewHandler::Initialize(IStream* pStream, DWORD grfMode)
{
  if (stream)
    stream->Release();
  stream = pStream;
  if (stream)
    stream->AddRef();
  return S_OK;
}

HRESULT MyFilePreviewHandler::SetWindow(HWND hwnd, const RECT* prc)
{
  if (!prc)
    return E_INVALIDARG;
  m_hwndParent = hwnd;
  m_rcView = *prc;
  return S_OK;
}

HRESULT MyFilePreviewHandler::DoPreview()
{
  if (!m_hwndParent || !stream)
    return E_FAIL;

  STATSTG stat;
  HRESULT hr = stream->Stat(&stat, STATFLAG_DEFAULT);
  if (FAILED(hr))
    return hr;

  ULONGLONG size = stat.cbSize.QuadPart;
  std::wstring content;

  if (size >= sizeof(wchar_t)) {
    std::vector<wchar_t> buffer(size / sizeof(wchar_t));
    ULONG bytesRead = 0;
    hr = stream->Read(buffer.data(), static_cast<ULONG>(size), &bytesRead);
    if (SUCCEEDED(hr)) {
      size_t charsRead = bytesRead / sizeof(wchar_t);
      content.assign(buffer.begin(), buffer.begin() + charsRead);
    }
  }
  // Get path of DLL to find the 'db' folder
  TCHAR szDllPath[MAX_PATH];
  GetModuleFileName((HINSTANCE)&__ImageBase, szDllPath, MAX_PATH);
  std::wstring wsDllPath(szDllPath);
  size_t pos = wsDllPath.find_last_of(L"\\");
  std::wstring databasePath = wsDllPath.substr(0, pos + 1) + L"db";
  OutputDebugStringW((L"Database path set to: " + databasePath + L"\n").c_str());

  // Transform content using the database
  content = GetFileInformation(content, databasePath);

  OutputDebugStringW((L"Content previewed: " + content.substr(0, 100) + L"\n").c_str());
  m_hwndEdit = CreateWindowExW(0, L"EDIT", nullptr,
    WS_CHILD | WS_VISIBLE | WS_VSCROLL | ES_MULTILINE | ES_READONLY,
    m_rcView.left, m_rcView.top,
    m_rcView.right - m_rcView.left,
    m_rcView.bottom - m_rcView.top,
    m_hwndParent, nullptr, GetModuleHandleW(nullptr), nullptr);

  if (!m_hwndEdit)
    return E_FAIL;

  if (!SetWindowTextW(m_hwndEdit, content.c_str())) {
    DestroyWindow(m_hwndEdit);
    m_hwndEdit = nullptr;
    return E_FAIL;
  }

  return S_OK;
}


HRESULT MyFilePreviewHandler::Unload()
{
  if (m_hwndEdit)
  {
    DestroyWindow(m_hwndEdit);
    m_hwndEdit = nullptr;
  }
  if (stream)
  {
    stream->Release();
    stream = nullptr;
  }
  return S_OK;
}

HRESULT MyFilePreviewHandler::SetFocus()
{
  if (m_hwndEdit)
    ::SetFocus(m_hwndEdit);
  return S_OK;
}

HRESULT MyFilePreviewHandler::QueryFocus(HWND* phwnd)
{
  if (!phwnd)
    return E_INVALIDARG;
  *phwnd = m_hwndEdit;
  return S_OK;
}

HRESULT MyFilePreviewHandler::TranslateAccelerator(MSG* pmsg)
{
  return S_FALSE;
}

HRESULT MyFilePreviewHandler::SetSite(IUnknown* pUnkSite)
{
  if (site)
    site->Release();
  site = pUnkSite;
  if (site)
    site->AddRef();
  return S_OK;
}

HRESULT MyFilePreviewHandler::GetSite(REFIID riid, void** ppvSite)
{
  if (!ppvSite)
    return E_INVALIDARG;
  if (!site)
    return E_FAIL;
  return site->QueryInterface(riid, ppvSite);
}