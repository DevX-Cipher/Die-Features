#pragma once
#include <iostream>
#include <Windows.h>
#include <Shlobj.h>
#include <fstream>
#include <chrono>
#include <thread>

VOID WINAPI ServiceMain(DWORD argc, LPTSTR* argv);
VOID WINAPI ServiceCtrlHandler(DWORD ctrl);
DWORD WINAPI MonitorDownloadFolder(LPVOID lpParam);
void LogMessage(const std::wstring& message);
void ShowToastNotification(const std::wstring& fileName, const std::wstring& fileInfo);

SERVICE_STATUS g_ServiceStatus;
SERVICE_STATUS_HANDLE g_ServiceStatusHandle;
std::wstring g_DownloadFolderPath;