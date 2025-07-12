#pragma once
#include <windows.h>
#include <shlobj.h>
#include <propkey.h>
#include <propvarutil.h>
#include <combaseapi.h>
#include <wrl.h>
#include <roapi.h>
#include <windows.ui.notifications.h>
#include <wrl/client.h>

using namespace Microsoft::WRL;
using namespace ABI::Windows::UI::Notifications;
using namespace ABI::Windows::Data::Xml::Dom;

class TaskbarHelper {
public:
  // Structure for jump list tasks
  struct JumpListTask {
    LPCWSTR exePath;
    LPCWSTR arguments;
    LPCWSTR description;
  };

  // Initialize the helper with a window handle
  static bool Initialize(HWND hwnd);
  // Cleanup resources
  static void Uninitialize();
  // Check if taskbar functionality is available
  static bool IsAvailable();

  // Progress bar functions
  static void SetProgressValue(int value, int max);
  static void SetProgressState(TBPFLAG state);

  // Overlay icon functions
  static void SetOverlayIcon(HICON hIcon, LPCWSTR description);
  static void ClearOverlayIcon();

  // Taskbar flashing
  static void FlashTaskbar(bool flash, int count);

  // Thumbnail toolbar buttons
  static bool AddThumbnailButtons(THUMBBUTTON* buttons, UINT count);
  static bool UpdateThumbnailButtons(THUMBBUTTON* buttons, UINT count);

  // Thumbnail tooltip and clip
  static void SetThumbnailTooltip(LPCWSTR tooltip);
  static void SetThumbnailClip(RECT* rect);
  static void ClearThumbnailClip();

  // Toast notification
  static void ShowToastNotification(const WCHAR* message, const WCHAR* appId);

  // Jump list tasks
  static bool AddJumpListTasks(const JumpListTask* tasks, UINT count);

private:
  TaskbarHelper() = default;
  ~TaskbarHelper();

  // Prevent copying
  TaskbarHelper(const TaskbarHelper&) = delete;
  TaskbarHelper& operator=(const TaskbarHelper&) = delete;

  // Singleton instance
  static TaskbarHelper* GetInstance();

  // Internal initialization
  bool InitializeInternal(HWND hwnd);

  // Member variables
  HWND hWnd = nullptr;
  ComPtr<ITaskbarList3> pTaskbarList;
  bool comInitialized = false;
};