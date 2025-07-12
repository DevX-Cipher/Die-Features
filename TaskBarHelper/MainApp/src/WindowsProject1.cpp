#include "framework.h"
#include "WindowsProject1.h"
#include <TaskbarHelper.h>

#define MAX_LOADSTRING 100
#define IDT_TIMER1 1
#define IDT_ERROR_RESET 2
#define IDM_PAUSE_RESUME 1001
#define IDM_ERROR 1002
#define IDM_STOP 1003
#define IDC_PLAY 1004
#define IDC_PAUSE 1005
#define IDC_STOP 1006
#define IDC_RESTART 1007

// Global Variables:
HINSTANCE hInst;                                // current instance
WCHAR szTitle[MAX_LOADSTRING];                  // The title bar text
WCHAR szWindowClass[MAX_LOADSTRING];            // the main window class name
int progress = 0;                              // Progress bar value (0-100)
bool isPaused = false;                         // Pause state
bool isError = false;                          // Error state
THUMBBUTTON thumbnailButtons[4];               // Thumbnail toolbar buttons (added Restart)
const WCHAR* appId = L"com.example.MyApp";     // Application ID for toast notifications

// Forward declarations of functions included in this code module:
ATOM                MyRegisterClass(HINSTANCE hInstance);
BOOL                InitInstance(HINSTANCE, int);
LRESULT CALLBACK    WndProc(HWND, UINT, WPARAM, LPARAM);
INT_PTR CALLBACK    About(HWND, UINT, WPARAM, LPARAM);

int APIENTRY wWinMain(_In_ HINSTANCE hInstance,
  _In_opt_ HINSTANCE hPrevInstance,
  _In_ LPWSTR    lpCmdLine,
  _In_ int       nCmdShow)
{
  UNREFERENCED_PARAMETER(hPrevInstance);
  UNREFERENCED_PARAMETER(lpCmdLine);

  // Initialize global strings
  LoadStringW(hInstance, IDS_APP_TITLE, szTitle, MAX_LOADSTRING);
  LoadStringW(hInstance, IDC_WINDOWSPROJECT1, szWindowClass, MAX_LOADSTRING);
  MyRegisterClass(hInstance);

  // Perform application initialization
  if (!InitInstance(hInstance, nCmdShow))
  {
    return FALSE;
  }

  HACCEL hAccelTable = LoadAccelerators(hInstance, MAKEINTRESOURCE(IDC_WINDOWSPROJECT1));

  MSG msg;

  // Main message loop
  while (GetMessage(&msg, nullptr, 0, 0))
  {
    if (!TranslateAccelerator(msg.hwnd, hAccelTable, &msg))
    {
      TranslateMessage(&msg);
      DispatchMessage(&msg);
    }
  }

  return (int)msg.wParam;
}

ATOM MyRegisterClass(HINSTANCE hInstance)
{
  WNDCLASSEXW wcex;

  wcex.cbSize = sizeof(WNDCLASSEX);

  wcex.style = CS_HREDRAW | CS_VREDRAW;
  wcex.lpfnWndProc = WndProc;
  wcex.cbClsExtra = 0;
  wcex.cbWndExtra = 0;
  wcex.hInstance = hInstance;
  wcex.hIcon = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_WINDOWSPROJECT1));
  wcex.hCursor = LoadCursor(nullptr, IDC_ARROW);
  wcex.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
  wcex.lpszMenuName = MAKEINTRESOURCEW(IDC_WINDOWSPROJECT1);
  wcex.lpszClassName = szWindowClass;
  wcex.hIconSm = LoadIcon(wcex.hInstance, MAKEINTRESOURCE(IDI_SMALL));

  return RegisterClassExW(&wcex);
}

BOOL InitInstance(HINSTANCE hInstance, int nCmdShow)
{
  hInst = hInstance; // Store instance handle in our global variable

  HWND hWnd = CreateWindowW(szWindowClass, szTitle, WS_OVERLAPPEDWINDOW,
    CW_USEDEFAULT, 0, 800, 600, nullptr, nullptr, hInstance, nullptr);

  if (!hWnd)
  {
    return FALSE;
  }

  // Initialize TaskbarHelper
  if (!TaskbarHelper::Initialize(hWnd))
  {
    return FALSE;
  }

  // Set up timer for progress bar animation (updates every 50ms)
  SetTimer(hWnd, IDT_TIMER1, 50, NULL);

  ShowWindow(hWnd, nCmdShow);
  UpdateWindow(hWnd);

  return TRUE;
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
  static HMENU hMenu = nullptr;

  switch (message)
  {
  case WM_CREATE:
  {
    hMenu = GetMenu(hWnd);

    // Add menu items
    HMENU hSubMenu = GetSubMenu(hMenu, 0); // Assume first menu is "File"
    AppendMenuW(hSubMenu, MF_STRING, IDM_PAUSE_RESUME, L"Pause");
    AppendMenuW(hSubMenu, MF_STRING, IDM_ERROR, L"Simulate Error");
    AppendMenuW(hSubMenu, MF_STRING, IDM_STOP, L"Stop");
    AppendMenuW(hSubMenu, MF_STRING, IDC_RESTART, L"Restart");

    // Set up taskbar features
    if (TaskbarHelper::IsAvailable()) {
      TaskbarHelper::SetThumbnailTooltip(L"Scanner Application");

      // Set thumbnail clip to progress bar area
      RECT clipRect = { 50, 200, 750, 250 };
      TaskbarHelper::SetThumbnailClip(&clipRect);

      // Set up thumbnail toolbar buttons
      ZeroMemory(thumbnailButtons, sizeof(thumbnailButtons));
      thumbnailButtons[0].dwMask = THB_TOOLTIP | THB_FLAGS;
      thumbnailButtons[0].iId = IDC_PLAY;
      wcscpy_s(thumbnailButtons[0].szTip, L"Play");
      thumbnailButtons[0].dwFlags = THBF_ENABLED;

      thumbnailButtons[1].dwMask = THB_TOOLTIP | THB_FLAGS;
      thumbnailButtons[1].iId = IDC_PAUSE;
      wcscpy_s(thumbnailButtons[1].szTip, L"Pause");
      thumbnailButtons[1].dwFlags = THBF_ENABLED;

      thumbnailButtons[2].dwMask = THB_TOOLTIP | THB_FLAGS;
      thumbnailButtons[2].iId = IDC_STOP;
      wcscpy_s(thumbnailButtons[2].szTip, L"Stop");
      thumbnailButtons[2].dwFlags = THBF_ENABLED;

      thumbnailButtons[3].dwMask = THB_TOOLTIP | THB_FLAGS;
      thumbnailButtons[3].iId = IDC_RESTART;
      wcscpy_s(thumbnailButtons[3].szTip, L"Restart");
      thumbnailButtons[3].dwFlags = THBF_ENABLED;

      TaskbarHelper::AddThumbnailButtons(thumbnailButtons, 4);

      // Add jump list tasks
      WCHAR exePath[MAX_PATH];
      GetModuleFileNameW(NULL, exePath, MAX_PATH);
      TaskbarHelper::JumpListTask tasks[] = {
        { exePath, L"/newscan", L"Start New Scan" },
        { exePath, L"/openfile", L"Open File" }
      };
      TaskbarHelper::AddJumpListTasks(tasks, 2);
    }
    else {
      // Disable menu items if taskbar interface is unavailable
      EnableMenuItem(hSubMenu, IDM_PAUSE_RESUME, MF_BYCOMMAND | MF_GRAYED);
      EnableMenuItem(hSubMenu, IDM_ERROR, MF_BYCOMMAND | MF_GRAYED);
      EnableMenuItem(hSubMenu, IDM_STOP, MF_BYCOMMAND | MF_GRAYED);
      EnableMenuItem(hSubMenu, IDC_RESTART, MF_BYCOMMAND | MF_GRAYED);
    }
  }
  break;
  case WM_COMMAND:
  {
    int wmId = LOWORD(wParam);
    if (HIWORD(wParam) == THBN_CLICKED) {
      // Handle thumbnail button clicks
      switch (wmId) {
      case IDC_PLAY:
        if (isPaused) {
          isPaused = false;
          ModifyMenuW(hMenu, IDM_PAUSE_RESUME, MF_BYCOMMAND | MF_STRING, IDM_PAUSE_RESUME, L"Pause");
          wcscpy_s(thumbnailButtons[1].szTip, L"Pause");
          thumbnailButtons[1].dwFlags = THBF_ENABLED;
          TaskbarHelper::UpdateThumbnailButtons(thumbnailButtons, 4);
          TaskbarHelper::SetProgressState(TBPF_NORMAL);
          TaskbarHelper::ShowToastNotification(L"Scan resumed", appId);
          InvalidateRect(hWnd, NULL, TRUE);
        }
        break;
      case IDC_PAUSE:
        if (!isPaused) {
          isPaused = true;
          ModifyMenuW(hMenu, IDM_PAUSE_RESUME, MF_BYCOMMAND | MF_STRING, IDM_PAUSE_RESUME, L"Resume");
          wcscpy_s(thumbnailButtons[1].szTip, L"Resume");
          thumbnailButtons[1].dwFlags = THBF_DISABLED;
          TaskbarHelper::UpdateThumbnailButtons(thumbnailButtons, 4);
          TaskbarHelper::SetProgressState(TBPF_PAUSED);
          TaskbarHelper::ShowToastNotification(L"Scan paused", appId);
          InvalidateRect(hWnd, NULL, TRUE);
        }
        break;
      case IDC_STOP:
        isPaused = false;
        isError = false;
        progress = 0;
        ModifyMenuW(hMenu, IDM_PAUSE_RESUME, MF_BYCOMMAND | MF_STRING, IDM_PAUSE_RESUME, L"Pause");
        wcscpy_s(thumbnailButtons[1].szTip, L"Pause");
        thumbnailButtons[1].dwFlags = THBF_ENABLED;
        TaskbarHelper::UpdateThumbnailButtons(thumbnailButtons, 4);
        TaskbarHelper::SetProgressState(TBPF_NOPROGRESS);
        TaskbarHelper::ShowToastNotification(L"Scan stopped", appId);
        InvalidateRect(hWnd, NULL, TRUE);
        break;
      case IDC_RESTART:
        isPaused = false;
        isError = false;
        progress = 0;
        ModifyMenuW(hMenu, IDM_PAUSE_RESUME, MF_BYCOMMAND | MF_STRING, IDM_PAUSE_RESUME, L"Pause");
        wcscpy_s(thumbnailButtons[1].szTip, L"Pause");
        thumbnailButtons[1].dwFlags = THBF_ENABLED;
        TaskbarHelper::UpdateThumbnailButtons(thumbnailButtons, 4);
        TaskbarHelper::SetProgressState(TBPF_NORMAL);
        TaskbarHelper::ShowToastNotification(L"Scan restarted", appId);
        InvalidateRect(hWnd, NULL, TRUE);
        break;
      }
    }
    else {
      // Handle menu commands
      switch (wmId) {
      case IDM_ABOUT:
        DialogBox(hInst, MAKEINTRESOURCE(IDD_ABOUTBOX), hWnd, About);
        break;
      case IDM_EXIT:
        DestroyWindow(hWnd);
        break;
      case IDM_PAUSE_RESUME:
        isPaused = !isPaused;
        ModifyMenuW(hMenu, IDM_PAUSE_RESUME, MF_BYCOMMAND | MF_STRING, IDM_PAUSE_RESUME, isPaused ? L"Resume" : L"Pause");
        wcscpy_s(thumbnailButtons[1].szTip, isPaused ? L"Resume" : L"Pause");
        thumbnailButtons[1].dwFlags = isPaused ? THBF_DISABLED : THBF_ENABLED;
        TaskbarHelper::UpdateThumbnailButtons(thumbnailButtons, 4);
        TaskbarHelper::SetProgressState(isPaused ? TBPF_PAUSED : TBPF_NORMAL);
        TaskbarHelper::ShowToastNotification(isPaused ? L"Scan paused" : L"Scan resumed", appId);
        InvalidateRect(hWnd, NULL, TRUE);
        break;
      case IDM_ERROR:
        isError = true;
        TaskbarHelper::SetProgressState(TBPF_ERROR);
        TaskbarHelper::FlashTaskbar(true, 5);
        TaskbarHelper::ShowToastNotification(L"Scan error detected", appId);
        InvalidateRect(hWnd, NULL, TRUE);
        SetTimer(hWnd, IDT_ERROR_RESET, 2000, NULL);
        break;
      case IDM_STOP:
        isPaused = false;
        isError = false;
        progress = 0;
        ModifyMenuW(hMenu, IDM_PAUSE_RESUME, MF_BYCOMMAND | MF_STRING, IDM_PAUSE_RESUME, L"Pause");
        wcscpy_s(thumbnailButtons[1].szTip, L"Pause");
        thumbnailButtons[1].dwFlags = THBF_ENABLED;
        TaskbarHelper::UpdateThumbnailButtons(thumbnailButtons, 4);
        TaskbarHelper::SetProgressState(TBPF_NOPROGRESS);
        TaskbarHelper::ShowToastNotification(L"Scan stopped", appId);
        InvalidateRect(hWnd, NULL, TRUE);
        break;
      case IDC_RESTART:
        isPaused = false;
        isError = false;
        progress = 0;
        ModifyMenuW(hMenu, IDM_PAUSE_RESUME, MF_BYCOMMAND | MF_STRING, IDM_PAUSE_RESUME, L"Pause");
        wcscpy_s(thumbnailButtons[1].szTip, L"Pause");
        thumbnailButtons[1].dwFlags = THBF_ENABLED;
        TaskbarHelper::UpdateThumbnailButtons(thumbnailButtons, 4);
        TaskbarHelper::SetProgressState(TBPF_NORMAL);
        TaskbarHelper::ShowToastNotification(L"Scan restarted", appId);
        InvalidateRect(hWnd, NULL, TRUE);
        break;
      default:
        return DefWindowProc(hWnd, message, wParam, lParam);
      }
    }
  }
  break;
  case WM_TIMER:
  {
    if (wParam == IDT_TIMER1 && !isPaused && !isError) {
      // Update progress value
      progress += 2;
      if (progress > 100) {
        progress = 0;
        TaskbarHelper::SetProgressState(TBPF_NORMAL);
        TaskbarHelper::ShowToastNotification(L"Scan cycle reset", appId);
      }

      TaskbarHelper::SetProgressValue(progress, 100);

      WCHAR tooltip[64];
      wsprintfW(tooltip, L"Scan Progress: %d%% %s", progress,
        isPaused ? L"(Paused)" : isError ? L"(Error)" : L"");
      TaskbarHelper::SetThumbnailTooltip(tooltip);

      if (progress == 50) {
        isPaused = true;
        ModifyMenuW(hMenu, IDM_PAUSE_RESUME, MF_BYCOMMAND | MF_STRING, IDM_PAUSE_RESUME, L"Resume");
        wcscpy_s(thumbnailButtons[1].szTip, L"Resume");
        thumbnailButtons[1].dwFlags = THBF_DISABLED;
        TaskbarHelper::UpdateThumbnailButtons(thumbnailButtons, 4);
        TaskbarHelper::SetProgressState(TBPF_PAUSED);
        TaskbarHelper::ShowToastNotification(L"Scan paused at 50%", appId);
      }
      else if (progress == 75) {
        isError = true;
        TaskbarHelper::SetProgressState(TBPF_ERROR);
        TaskbarHelper::FlashTaskbar(true, 5);
        TaskbarHelper::ShowToastNotification(L"Scan error detected", appId);
        SetTimer(hWnd, IDT_ERROR_RESET, 2000, NULL);
      }
      else if (progress == 100) {
        TaskbarHelper::SetProgressState(TBPF_NORMAL);
        TaskbarHelper::FlashTaskbar(true, 5);
        TaskbarHelper::ShowToastNotification(L"Scan completed successfully", appId);
      }
      else {
        TaskbarHelper::SetProgressState(TBPF_NORMAL);
      }

      InvalidateRect(hWnd, NULL, TRUE);
    }
    else if (wParam == IDT_ERROR_RESET) {
      KillTimer(hWnd, IDT_ERROR_RESET);
      if (!isPaused) {
        isError = false;
        TaskbarHelper::SetProgressState(TBPF_NORMAL);
        TaskbarHelper::ShowToastNotification(L"Scan error cleared", appId);
        InvalidateRect(hWnd, NULL, TRUE);
      }
    }
  }
  break;
  case WM_SIZE:
  {
    // Update thumbnail clip on window resize
    RECT clipRect = { 50, 200, LOWORD(lParam) - 50, 250 };
    TaskbarHelper::SetThumbnailClip(&clipRect);
    InvalidateRect(hWnd, NULL, TRUE);
  }
  break;
  case WM_PAINT:
  {
    PAINTSTRUCT ps;
    HDC hdc = BeginPaint(hWnd, &ps);

    // Get client area dimensions
    RECT clientRect;
    GetClientRect(hWnd, &clientRect);

    // Draw progress bar background
    RECT barRect = { 50, 200, clientRect.right - 50, 250 };
    HBRUSH bgBrush = CreateSolidBrush(RGB(200, 200, 200));
    FillRect(hdc, &barRect, bgBrush);
    DeleteObject(bgBrush);

    // Draw progress bar fill
    int fillWidth = (barRect.right - barRect.left) * progress / 100;
    RECT fillRect = { 50, 200, 50 + fillWidth, 250 };
    COLORREF fillColor = isPaused ? RGB(255, 255, 0) :
      isError ? RGB(255, 0, 0) : RGB(0, 128, 255);
    HBRUSH fillBrush = CreateSolidBrush(fillColor);
    FillRect(hdc, &fillRect, fillBrush);
    DeleteObject(fillBrush);

    // Draw border
    HBRUSH borderBrush = CreateSolidBrush(RGB(0, 0, 0));
    FrameRect(hdc, &barRect, borderBrush);
    DeleteObject(borderBrush);

    // Draw percentage text
    WCHAR text[64];
    wsprintfW(text, L"Progress: %d%% %s", progress,
      isPaused ? L"(Paused)" : isError ? L"(Error)" : L"");
    SetTextColor(hdc, RGB(0, 0, 0));
    SetBkMode(hdc, TRANSPARENT);
    TextOutW(hdc, 50, 170, text, lstrlenW(text));

    EndPaint(hWnd, &ps);
  }
  break;
  case WM_DESTROY:
    TaskbarHelper::Uninitialize();
    KillTimer(hWnd, IDT_TIMER1);
    KillTimer(hWnd, IDT_ERROR_RESET);
    PostQuitMessage(0);
    break;
  default:
    return DefWindowProc(hWnd, message, wParam, lParam);
  }
  return 0;
}

INT_PTR CALLBACK About(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
  UNREFERENCED_PARAMETER(lParam);
  switch (message)
  {
  case WM_INITDIALOG:
    return (INT_PTR)TRUE;

  case WM_COMMAND:
    if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL)
    {
      EndDialog(hDlg, LOWORD(wParam));
      return (INT_PTR)TRUE;
    }
    break;
  }
  return (INT_PTR)FALSE;
}