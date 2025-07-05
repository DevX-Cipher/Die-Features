#include <iostream>
#include <Windows.h>
#include <Shlobj.h>
#include <fstream>
#include <chrono>
#include <thread>
#include "DieService.h"
#include "wintoastlib.h"
#include "ToastNotification.h"

using namespace WinToastLib;

/**
 * @brief Entry point for the application.
 *
 * This function defines the service name and entry point, creates a service table,
 * and starts the service control dispatcher. It returns 0 upon successful execution
 * or 1 if an error occurs.
 *
 * @return int indicating the exit status of the application.
 */

int main()
{
		// Define service name and entry point
		LPCWSTR serviceName = L"DiEService";
		SERVICE_TABLE_ENTRY ServiceTable[] =
		{
						{ const_cast<LPWSTR>(serviceName), ServiceMain },
						{ NULL, NULL }
		};

		// Start the service control dispatcher
		if (!StartServiceCtrlDispatcher(ServiceTable))
		{
				std::cerr << "StartServiceCtrlDispatcher failed, error: " << GetLastError() << std::endl;
				return 1;
		}

		return 0;
}

/**
 * @brief ServiceMain is the entry point for the service.
 *
 * This function initializes the service status, registers the service control handler,
 * retrieves the user's download folder path, creates a thread to monitor the download folder,
 * and enters the main service loop. The service loop performs service work while the service
 * is running and updates the service status accordingly. When the service stops, it sets
 * the service state to SERVICE_STOPPED.
 *
 * @param argc The number of command-line arguments.
 * @param argv An array of command-line arguments.
 */

VOID WINAPI ServiceMain(DWORD argc, LPTSTR* argv)
{
		// Register the service control handler
		g_ServiceStatusHandle = RegisterServiceCtrlHandler(L"DiEService", ServiceCtrlHandler);
		if (!g_ServiceStatusHandle)
		{
				std::cerr << "RegisterServiceCtrlHandler failed, error: " << GetLastError() << std::endl;
				return;
		}

		// Initialize service status
		g_ServiceStatus.dwServiceType = SERVICE_WIN32_OWN_PROCESS;
		g_ServiceStatus.dwCurrentState = SERVICE_START_PENDING;
		g_ServiceStatus.dwControlsAccepted = SERVICE_ACCEPT_STOP;
		g_ServiceStatus.dwWin32ExitCode = NO_ERROR;
		g_ServiceStatus.dwServiceSpecificExitCode = 0;
		g_ServiceStatus.dwCheckPoint = 0;
		g_ServiceStatus.dwWaitHint = 0;

		// Report service status to the service control manager
		if (!SetServiceStatus(g_ServiceStatusHandle, &g_ServiceStatus))
		{
				std::cerr << "SetServiceStatus failed, error: " << GetLastError() << std::endl;
				return;
		}

		// Get the path to the user's download folder
		PWSTR downloadFolderPath = nullptr;
		HRESULT hr = SHGetKnownFolderPath(FOLDERID_Downloads, 0, NULL, &downloadFolderPath);
		if (SUCCEEDED(hr))
		{
				g_DownloadFolderPath = downloadFolderPath;
				CoTaskMemFree(downloadFolderPath);
		}
		else
		{
				LogMessage(L"Failed to retrieve user's Download Folder.");
				return;
		}

		// Create a thread to monitor the user's download folder
		HANDLE hMonitorThread = CreateThread(NULL, 0, MonitorDownloadFolder, NULL, 0, NULL);
		if (hMonitorThread == NULL)
		{
				std::cerr << "Failed to create monitoring thread, error: " << GetLastError() << std::endl;
				return;
		}
		CloseHandle(hMonitorThread); // We don't need to keep the thread handle

		// Update service status to indicate service is running
		g_ServiceStatus.dwCurrentState = SERVICE_RUNNING;
		if (!SetServiceStatus(g_ServiceStatusHandle, &g_ServiceStatus))
		{
				std::cerr << "SetServiceStatus failed, error: " << GetLastError() << std::endl;
				return;
		}

		// Main service loop
		while (g_ServiceStatus.dwCurrentState == SERVICE_RUNNING)
		{
				// Perform service work here
				Sleep(1000); // Example: Sleep for 1 second
		}

		// Service has stopped
		g_ServiceStatus.dwCurrentState = SERVICE_STOPPED;
		if (!SetServiceStatus(g_ServiceStatusHandle, &g_ServiceStatus))
		{
				std::cerr << "SetServiceStatus failed, error: " << GetLastError() << std::endl;
				return;
		}
}

/**
 * @brief ServiceCtrlHandler is the control handler function for the service.
 *
 * This function handles control requests sent to the service, particularly SERVICE_CONTROL_STOP,
 * to initiate the service stopping process. Upon receiving the stop control request,
 * it sets the service state to SERVICE_STOP_PENDING, performs cleanup and stop routines, and then
 * sets the service state to SERVICE_STOPPED.
 *
 * @param ctrl The control code sent to the service.
 */

VOID WINAPI ServiceCtrlHandler(DWORD ctrl)
{
		switch (ctrl)
		{
		case SERVICE_CONTROL_STOP:
				g_ServiceStatus.dwCurrentState = SERVICE_STOP_PENDING;
				// Perform cleanup and stop routine here
				g_ServiceStatus.dwCurrentState = SERVICE_STOPPED;
				break;
		}
}

/**
 * @brief MonitorDownloadFolder monitors a specified folder for changes and takes actions based on them.
 *
 * This function continuously monitors the specified folder for file changes, particularly for .exe files being downloaded.
 * Upon detecting a completed download, it logs the event, retrieves the executable and database paths, fetches file information,
 * and shows a toast notification. It utilizes a map to keep track of file sizes and update them accordingly.
 *
 * @param lpParam Pointer to optional parameters for the function (unused in this implementation).
 * @return DWORD indicating the result of the function execution.
 */

DWORD WINAPI MonitorDownloadFolder(LPVOID lpParam)
{
		HANDLE hDirectory = CreateFile(g_DownloadFolderPath.c_str(), FILE_LIST_DIRECTORY, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, NULL, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, NULL);
		if (hDirectory != INVALID_HANDLE_VALUE)
		{
				BYTE buffer[4096]; // Buffer to hold change notifications
				DWORD bytesReturned;
				std::map<std::wstring, DWORD> fileSizes;
				while (ReadDirectoryChangesW(hDirectory, &buffer, sizeof(buffer), TRUE, FILE_NOTIFY_CHANGE_FILE_NAME | FILE_NOTIFY_CHANGE_DIR_NAME | FILE_NOTIFY_CHANGE_ATTRIBUTES | FILE_NOTIFY_CHANGE_SIZE | FILE_NOTIFY_CHANGE_LAST_WRITE | FILE_NOTIFY_CHANGE_LAST_ACCESS | FILE_NOTIFY_CHANGE_CREATION | FILE_NOTIFY_CHANGE_SECURITY, &bytesReturned, NULL, NULL))
				{
						FILE_NOTIFY_INFORMATION* fileInfo = reinterpret_cast<FILE_NOTIFY_INFORMATION*>(buffer);
						while (fileInfo)
						{
								// Check if the changed file is an .exe file
								std::wstring fileName(fileInfo->FileName, fileInfo->FileNameLength / sizeof(wchar_t));
								if (fileName.length() > 4 && fileName.substr(fileName.length() - 4) == L".exe")
								{
										std::wstring filePath = g_DownloadFolderPath + L"\\" + fileName;
										HANDLE hFile = CreateFile(filePath.c_str(), GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
										if (hFile != INVALID_HANDLE_VALUE)
										{
												DWORD fileSize = GetFileSize(hFile, NULL);
												CloseHandle(hFile);
												if (fileSizes.count(fileName) > 0)
												{
														if (fileSizes[fileName] == fileSize)
														{
																Sleep(5000);
																// File size has not changed, so the download is likely complete
																//ShowToastNotification(fileName);
																std::wstring logMessage = L"Download complete: " + fileName + L" (" + std::to_wstring(fileSize) + L" bytes)";
																LogMessage(logMessage.c_str());

																wchar_t exePath[MAX_PATH];
																if (GetModuleFileName(NULL, exePath, MAX_PATH) == 0) {
																		LogMessage(L"Failed to get executable path.");
																		return 0; // or handle error appropriately
																}
																PathRemoveFileSpec(exePath);
																std::wstring exePathMessage = L"Executable Path: " + std::wstring(exePath);
																LogMessage(exePathMessage.c_str());

																std::wstring pwszDatabase = std::wstring(exePath) + L"\\db";
																std::wstring dbPathMessage = L"Database Path: " + pwszDatabase;
																LogMessage(dbPathMessage.c_str());

																// Get file information
																std::wstring fileInfo = GetFileInformation(filePath, pwszDatabase);
																if (fileInfo.empty()) {
																		LogMessage(L"Failed to get file information.");
																		return 0; // or handle error appropriately
																}

																// Show file information in a toast notification
																ShowToastNotification(fileName, fileInfo);

																fileSizes.erase(fileName);

														}
														else
														{
																// Update the recorded file size
																fileSizes[fileName] = fileSize;
																std::wstring logMessage = L"File size updated: " + fileName + L" (" + std::to_wstring(fileSize) + L" bytes)";
																LogMessage(logMessage.c_str());
														}
												}
												else
												{
														// This is a new file, so record its size
														fileSizes[fileName] = fileSize;
														std::wstring logMessage = L"New file detected: " + fileName + L" (" + std::to_wstring(fileSize) + L" bytes)";
														LogMessage(logMessage.c_str());
												}
										}
								}

								// Move to the next notification
								if (fileInfo->NextEntryOffset == 0)
										break;
								fileInfo = reinterpret_cast<FILE_NOTIFY_INFORMATION*>(reinterpret_cast<BYTE*>(fileInfo) + fileInfo->NextEntryOffset);
						}
				}
				CloseHandle(hDirectory);
		}
		else
		{
				LogMessage(L"Failed to open Download Folder for monitoring.");
		}

		return 0;
}

/**
 * @brief LogMessage logs a message to a file.
 *
 * This function appends the specified message to a log file located at "C:\\ServiceLog.txt".
 * If the file cannot be opened, it prints an error message to stderr.
 *
 * @param message The message to be logged.
 */

void LogMessage(const std::wstring& message)
{
		std::wofstream logFile(L"C:\\ServiceLog.txt", std::ios::app);
		if (logFile.is_open())
		{
				logFile << message << std::endl;
				logFile.close();
		}
		else
		{
				std::cerr << "Failed to open log file." << std::endl;
		}
}

/**
 * @brief ShowToastNotification displays a toast notification.
 *
 * This function creates and configures a toast notification using WinToast library,
 * and then displays it. If the notification fails to initialize or show, it logs an error message.
 *
 * @param fileName The name of the file triggering the notification.
 */

void ShowToastNotification(const std::wstring& fileName, const std::wstring& fileInfo)
{
		// Create a WinToast object
		WinToast::instance()->setAppName(L"DiEService");
		const auto aumi = WinToast::configureAUMI(L"company", L"wintoast", L"wintoastexample", L"20201012");
		WinToast::instance()->setAppUserModelId(aumi);

		WinToast::WinToastError error;
		if (!WinToast::instance()->initialize(&error)) {
				LogMessage(L"Failed to initialize WinToast. Error number: ");
				return;
		}

		// Create and configure the toast notification
		WinToastTemplate templ = WinToastTemplate(WinToastTemplate::Text04);
		templ.setTextField(L"File System Watcher", WinToastTemplate::FirstLine);
		templ.setTextField(fileInfo, WinToastTemplate::SecondLine);

		// Add a button to the toast notification
		templ.addAction(L"Open");

		// Display the toast notification
		MyWinToastHandler* handler = new MyWinToastHandler;
		if (!WinToast::instance()->showToast(templ, handler)) {
				LogMessage(L"Failed to show toast notification.");
		}
}
