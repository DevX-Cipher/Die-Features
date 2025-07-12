#include "TaskbarHelper.h"

TaskbarHelper* TaskbarHelper::GetInstance() {
  static TaskbarHelper instance;
  return &instance;
}

bool TaskbarHelper::Initialize(HWND hwnd) {
  return GetInstance()->InitializeInternal(hwnd);
}

void TaskbarHelper::Uninitialize() {
  TaskbarHelper* instance = GetInstance();
  instance->~TaskbarHelper();
}

bool TaskbarHelper::InitializeInternal(HWND hwnd) {
  if (hWnd != nullptr) {
    return false;
  }
  hWnd = hwnd;

  HRESULT hr = CoInitialize(nullptr);
  comInitialized = SUCCEEDED(hr);
  if (!comInitialized && hr != S_OK && hr != S_FALSE) {
    return false;
  }

  hr = CoCreateInstance(CLSID_TaskbarList, nullptr, CLSCTX_INPROC_SERVER,
    IID_PPV_ARGS(&pTaskbarList));
  if (SUCCEEDED(hr)) {
    hr = pTaskbarList->HrInit();
    if (FAILED(hr)) {
      pTaskbarList.Reset();
    }
  }
  else {
    pTaskbarList.Reset();
  }
  return pTaskbarList != nullptr;
}

TaskbarHelper::~TaskbarHelper() {
  if (pTaskbarList) {
    pTaskbarList->SetProgressState(hWnd, TBPF_NOPROGRESS);
    pTaskbarList->SetOverlayIcon(hWnd, nullptr, nullptr);
    pTaskbarList.Reset();
  }
  if (comInitialized) {
    CoUninitialize();
  }
  hWnd = nullptr;
}

bool TaskbarHelper::IsAvailable() {
  return GetInstance()->pTaskbarList != nullptr;
}

void TaskbarHelper::SetProgressValue(int value, int max) {
  auto* instance = GetInstance();
  if (instance->pTaskbarList) {
    instance->pTaskbarList->SetProgressValue(instance->hWnd, value, max);
  }
}

void TaskbarHelper::SetProgressState(TBPFLAG state) {
  auto* instance = GetInstance();
  if (instance->pTaskbarList) {
    instance->pTaskbarList->SetProgressState(instance->hWnd, state);
  }
}

void TaskbarHelper::SetOverlayIcon(HICON hIcon, LPCWSTR description) {
  auto* instance = GetInstance();
  if (instance->pTaskbarList) {
    instance->pTaskbarList->SetOverlayIcon(instance->hWnd, hIcon, description);
  }
}

void TaskbarHelper::ClearOverlayIcon() {
  auto* instance = GetInstance();
  if (instance->pTaskbarList) {
    instance->pTaskbarList->SetOverlayIcon(instance->hWnd, nullptr, nullptr);
  }
}

void TaskbarHelper::FlashTaskbar(bool flash, int count) {
  auto* instance = GetInstance();
  if (instance->hWnd) {
    FLASHWINFO flashInfo = { sizeof(FLASHWINFO) };
    flashInfo.hwnd = instance->hWnd;
    flashInfo.dwFlags = flash ? (FLASHW_TRAY | FLASHW_TIMERNOFG) : FLASHW_STOP;
    flashInfo.uCount = count;
    flashInfo.dwTimeout = 0;
    FlashWindowEx(&flashInfo);
  }
}

bool TaskbarHelper::AddThumbnailButtons(THUMBBUTTON* buttons, UINT count) {
  auto* instance = GetInstance();
  if (instance->pTaskbarList && count <= 7) {
    HRESULT hr = instance->pTaskbarList->ThumbBarAddButtons(instance->hWnd, count, buttons);
    return SUCCEEDED(hr);
  }
  return false;
}

bool TaskbarHelper::UpdateThumbnailButtons(THUMBBUTTON* buttons, UINT count) {
  auto* instance = GetInstance();
  if (instance->pTaskbarList && count <= 7) {
    HRESULT hr = instance->pTaskbarList->ThumbBarUpdateButtons(instance->hWnd, count, buttons);
    return SUCCEEDED(hr);
  }
  return false;
}

void TaskbarHelper::SetThumbnailTooltip(LPCWSTR tooltip) {
  auto* instance = GetInstance();
  if (instance->pTaskbarList) {
    instance->pTaskbarList->SetThumbnailTooltip(instance->hWnd, tooltip);
  }
}

void TaskbarHelper::SetThumbnailClip(RECT* rect) {
  auto* instance = GetInstance();
  if (instance->pTaskbarList) {
    instance->pTaskbarList->SetThumbnailClip(instance->hWnd, rect);
  }
}

void TaskbarHelper::ClearThumbnailClip() {
  auto* instance = GetInstance();
  if (instance->pTaskbarList) {
    instance->pTaskbarList->SetThumbnailClip(instance->hWnd, nullptr);
  }
}

void TaskbarHelper::ShowToastNotification(const WCHAR* message, const WCHAR* appId) {
  auto* instance = GetInstance();
  ComPtr<IToastNotificationManagerStatics> toastManager;
  HRESULT hr = RoGetActivationFactory(
    Wrappers::HStringReference(RuntimeClass_Windows_UI_Notifications_ToastNotificationManager).Get(),
    IID_PPV_ARGS(&toastManager));
  if (FAILED(hr)) return;

  ComPtr<IToastNotifier> notifier;
  hr = toastManager->CreateToastNotifierWithId(
    Wrappers::HStringReference(appId).Get(), &notifier);
  if (FAILED(hr)) return;

  ComPtr<IXmlDocument> toastXml;
  hr = toastManager->GetTemplateContent(ToastTemplateType_ToastText01, &toastXml);
  if (FAILED(hr)) return;

  ComPtr<IXmlNodeList> textNodes;
  hr = toastXml->GetElementsByTagName(Wrappers::HStringReference(L"text").Get(), &textNodes);
  if (FAILED(hr)) return;

  ComPtr<IXmlNode> textNode;
  hr = textNodes->Item(0, &textNode);
  if (FAILED(hr)) return;

  ComPtr<IXmlText> text;
  hr = toastXml->CreateTextNode(Wrappers::HStringReference(message).Get(), &text);
  if (FAILED(hr)) return;

  ComPtr<IXmlNode> textAsNode;
  hr = text.As(&textAsNode);
  if (FAILED(hr)) return;

  ComPtr<IXmlNode> appendedText;
  hr = textNode->AppendChild(textAsNode.Get(), &appendedText);
  if (FAILED(hr)) return;

  ComPtr<IToastNotificationFactory> toastFactory;
  hr = RoGetActivationFactory(
    Wrappers::HStringReference(RuntimeClass_Windows_UI_Notifications_ToastNotification).Get(),
    IID_PPV_ARGS(&toastFactory));
  if (FAILED(hr)) return;

  ComPtr<IToastNotification> toast;
  hr = toastFactory->CreateToastNotification(toastXml.Get(), &toast);
  if (FAILED(hr)) return;

  hr = notifier->Show(toast.Get());
}

bool TaskbarHelper::AddJumpListTasks(const JumpListTask* tasks, UINT count) {
  auto* instance = GetInstance();
  ComPtr<ICustomDestinationList> pDestList;
  HRESULT hr = CoCreateInstance(CLSID_DestinationList, nullptr, CLSCTX_INPROC_SERVER,
    IID_PPV_ARGS(&pDestList));
  if (FAILED(hr)) return false;

  UINT maxSlots;
  ComPtr<IObjectArray> pRemoved;
  hr = pDestList->BeginList(&maxSlots, IID_PPV_ARGS(&pRemoved));
  if (FAILED(hr)) return false;

  ComPtr<IObjectCollection> pCollection;
  hr = CoCreateInstance(CLSID_EnumerableObjectCollection, nullptr, CLSCTX_INPROC_SERVER,
    IID_PPV_ARGS(&pCollection));
  if (FAILED(hr)) return false;

  for (UINT i = 0; i < count; ++i) {
    ComPtr<IShellLink> pLink;
    hr = CoCreateInstance(CLSID_ShellLink, nullptr, CLSCTX_INPROC_SERVER,
      IID_PPV_ARGS(&pLink));
    if (FAILED(hr)) return false;

    pLink->SetPath(tasks[i].exePath);
    pLink->SetArguments(tasks[i].arguments);
    pLink->SetDescription(tasks[i].description);

    hr = pCollection->AddObject(pLink.Get());
    if (FAILED(hr)) return false;
  }

  ComPtr<IObjectArray> pTaskArray;
  hr = pCollection.As(&pTaskArray);
  if (FAILED(hr)) return false;

  hr = pDestList->AddUserTasks(pTaskArray.Get());
  if (FAILED(hr)) return false;

  hr = pDestList->CommitList();
  return SUCCEEDED(hr);
}