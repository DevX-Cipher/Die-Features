
#include "wintoastlib.h"
#include <regex> 
#include <string>
#include <sstream>

extern "C" __declspec(dllimport) int __cdecl DIE_VB_ScanFile(const wchar_t* pwszFileName, int nFlags, const wchar_t* pwszDatabase, wchar_t* pwszBuffer, int nBufferSize);

class MyWinToastHandler : public WinToastLib::IWinToastHandler {
public:
		void toastActivated() const override {
				printf("Toast activated\n");
		}

		// Needs to be Looked at 
		void toastActivated(int actionIndex) const override {
				if (actionIndex == 0) {  // The index of the "Open" button is 0
						// Get the path of the current executable
						wchar_t buffer[MAX_PATH];
						GetModuleFileNameW(NULL, buffer, MAX_PATH);
						std::wstring::size_type pos = std::wstring(buffer).find_last_of(L"\\/");
						std::wstring currentDir = std::wstring(buffer).substr(0, pos);

						// Log the current directory
						LogMessage(L"Current directory: " + currentDir);

						// Construct the path of die.exe
						std::wstring diePath = currentDir + L"\\die.exe";

						// Log the path of die.exe
						LogMessage(L"Path of die.exe: " + diePath);

						// Open die.exe
						HINSTANCE hInst = ShellExecuteW(NULL, L"open", diePath.c_str(), NULL, NULL, SW_SHOWNORMAL);

						// Check if ShellExecute was successful
						if ((int)hInst <= 32) {
								DWORD errorCode = GetLastError();
								std::wstringstream ss;
								ss << L"Failed to open die.exe, error code: " << errorCode;
								LogMessage(ss.str());
						}
						else {
								LogMessage(L"Successfully opened die.exe");
						}
				}
		}

		void toastDismissed(WinToastDismissalReason state) const override {
				switch (state) {
				case UserCanceled:
						printf("Toast dismissed by user\n");
						break;
				case ApplicationHidden:
						printf("Toast dismissed by app\n");
						break;
				case TimedOut:
						printf("Toast dismissed after timeout\n");
						break;
				default:
						printf("Toast dismissed\n");
						break;
				}
		}

		void toastFailed() const override {
				printf("Toast failed\n");
		}
};


std::wstring GetFileInformation(const std::wstring& filePath, const std::wstring& pwszDatabase)
{
		try {
				const int nBufferSize = 10000;
				std::wstring sBuffer(nBufferSize, L' ');  // Allocate buffer
				//Log before calling DIE_VB_ScanFile
				std::wstring logMessageBefore = L"Calling DIE_VB_ScanFile with filePath: " + filePath + L" and database: " + pwszDatabase;
				LogMessage(logMessageBefore.c_str());

				// Use DIE_VB_ScanFile to get the file information
				int nResult = DIE_VB_ScanFile(filePath.c_str(), 0, pwszDatabase.c_str(), &sBuffer[0], nBufferSize - 1);

				// Log after calling DIE_VB_ScanFile
				std::wstring logMessageAfter = L"DIE_VB_ScanFile returned with result: " + std::to_wstring(nResult);
				LogMessage(logMessageAfter.c_str());

				std::wstring result;
				if (nResult > 0) {
						// Trim the buffer to the actual length of the result
						sBuffer.resize(nResult);
						// Extract all occurrences of Protection results from sBuffer
						std::wregex regex(L"Protection:\\s*(.*?)\\s*\\n");
						std::wsregex_iterator it(sBuffer.begin(), sBuffer.end(), regex);
						std::wsregex_iterator end;
						while (it != end) {
								result += it->str(1); // Append the matched Protection result
								result += L"\n"; // Add newline character to separate each Protection result
								++it;
						}
						if (result.empty()) {
								result = L"No Protection results found.";
						}
						// Log the extracted Protection results
						std::wstring logMessageProtection = L"Extracted Protection results:\n" + result;
						LogMessage(logMessageProtection.c_str());
				}
				else {
						result = L"Scan failed or no result.";
				}

				return result;
		}
		catch (...) {
				// Handle the exception here if needed
				return L"Exception occurred while retrieving file information";
		}
}