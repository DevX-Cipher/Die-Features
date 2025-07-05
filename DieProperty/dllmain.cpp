#include <Windows.h>
#include <ShlObj.h>
#include <ShObjIdl.h>
#include <Shlwapi.h>
#include <atlbase.h>
#include <shellapi.h>
#include <fstream>
#include <iostream>
#include <vector>
#include <string>
#include <sstream>
#include <iomanip>

#include "resource.h"
#include "ClassFactory.h"
#include "dllmain.h"

HRESULT RegisterShellExtension(HMODULE hModule);
HRESULT UnregisterShellExtension();

HINSTANCE   g_hInst;
UINT        g_DllRefCount;

/**
 * @brief Logs a message to a file on the desktop.
 *
 * This function writes the specified message to a log file located on the desktop.
 * If the log file does not exist, it is created. If it exists, the message is appended to it.
 * Error messages are printed to the standard error stream if any operation fails.
 *
 * @param message The message to be logged.
 */

void LogMessage(const std::wstring& message) {
#ifdef _DEBUG
		wchar_t desktopPath[MAX_PATH]; // Buffer to store the desktop path
		// Attempt to retrieve the desktop directory path
		if (SUCCEEDED(SHGetFolderPathW(NULL, CSIDL_DESKTOPDIRECTORY, NULL, 0, desktopPath))) {
				std::wstring logFilePath = std::wstring(desktopPath) + L"\\ServiceLog.txt"; // Construct the log file path
				std::wofstream logFile(logFilePath, std::ios::app); // Open the log file in append mode
				if (logFile.is_open()) { // Check if the file was opened successfully
						logFile << message << std::endl; // Write the message to the log file
						logFile.close(); // Close the log file
				}
				else {
						std::cerr << "Failed to open log file." << std::endl; // Print error message to standard error stream
				}
		}
		else {
				std::cerr << "Failed to get desktop directory." << std::endl; // Print error message to standard error stream
		}
#endif
}

/**
 * @brief Entry point for DLL initialization and cleanup.
 *
 * The system calls this function when the DLL is loaded or unloaded.
 * It performs initialization tasks when the DLL is attached to a process and cleanup tasks when detached.
 * The function logs attachment and detachment events for diagnostic purposes.
 *
 * @param hinstDLL Handle to the DLL module.
 * @param fdwReason Reason code for the DLL operation.
 * @param lpvReserved Reserved parameter.
 * @return BOOL Returns TRUE if the operation succeeds, otherwise FALSE.
 */

BOOL WINAPI DllMain(HINSTANCE hinstDLL, DWORD fdwReason, LPVOID lpvReserved) {
		switch (fdwReason) {
		case DLL_PROCESS_ATTACH:
				// Initialize DLL
				g_hInst = (HINSTANCE)hinstDLL; // Store the instance handle
				LogMessage(L"DLL attached."); // Log a message indicating DLL attachment
				break;
		case DLL_PROCESS_DETACH:
				// Cleanup DLL
				LogMessage(L"DLL detached."); // Log a message indicating DLL detachment
				break;
		case DLL_THREAD_ATTACH:
				// Do thread-specific initialization
				break;
		case DLL_THREAD_DETACH:
				// Do thread-specific cleanup
				break;
		}
		return TRUE; // Operation succeeded
}

/**
 * @brief The CLSID for the property sheet handler COM object.
 *
 * This constant defines the CLSID (Class Identifier) for the property sheet handler COM object.
 * It is used to identify the COM object when registering or accessing it.
 */

const CLSID CLSID_MyPropertySheetHandler = { 0x05B54455, 0x9F28, 0x436D, { 0x9E, 0x1F, 0x18, 0x48, 0x4C, 0x5A, 0xB7, 0x9A } };

/**
 * @brief Retrieves a class factory object for the specified CLSID.
 *
 * This function is called by the system to retrieve a class factory object for the specified CLSID.
 * It checks if the specified CLSID matches the CLSID of the property sheet handler COM object.
 * If a match is found, it creates a new instance of the class factory and queries the requested interface.
 *
 * @param rclsid The CLSID of the requested COM object.
 * @param riid The interface identifier.
 * @param ppv Address of the pointer variable that receives the interface pointer requested in riid.
 * @return HRESULT Returns S_OK if successful, otherwise an error code indicating failure.
 */

HRESULT DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv) {
		LogMessage(L"DllGetClassObject called."); // Log a message

		*ppv = NULL; // Initialize output parameter

		// Check if the requested CLSID matches the CLSID of the property sheet handler COM object
		if (!IsEqualCLSID(rclsid, CLSID_MyPropertySheetHandler))
				return CLASS_E_CLASSNOTAVAILABLE; // The requested CLSID is not supported

		// Create a new instance of the class factory
		CClassFactory* pClassFactory = new CClassFactory();
		if (pClassFactory == NULL)
				return E_OUTOFMEMORY; // Memory allocation failed

		// Query the requested interface for the class factory object
		HRESULT hResult = pClassFactory->QueryInterface(riid, ppv);

		// Only release the object if QueryInterface failed
		if (FAILED(hResult)) {
				LogMessage(L"DllGetClassObject: QueryInterface failed. Releasing CClassFactory object."); // Log message indicating QueryInterface failure
				pClassFactory->Release(); // Release the reference to the class factory
		}

		return hResult; // Return the result of the query
}

/**
 * @brief Registers the COM object as a shell extension.
 *
 * This function is called to register the COM object as a shell extension.
 * It calls the RegisterShellExtension function to perform the registration.
 * It logs the outcome of the registration operation for diagnostic purposes.
 *
 * @return HRESULT Returns S_OK if registration succeeds, otherwise an error code indicating failure.
 */

HRESULT WINAPI DllRegisterServer() {
		LogMessage(L"DllRegisterServer called."); // Log a message
		HRESULT hr = RegisterShellExtension(g_hInst); // Register the shell extension
		if (SUCCEEDED(hr)) {
				LogMessage(L"RegisterShellExtension succeeded."); // Log success
		}
		else {
				LogMessage(L"RegisterShellExtension failed."); // Log failure
		}
		return hr; // Return the result of the registration operation
}

/**
 * @brief Unregisters the COM object as a shell extension.
 *
 * This function is called to unregister the COM object as a shell extension.
 * It calls the UnregisterShellExtension function to perform the unregistration.
 * It logs the outcome of the unregistration operation for diagnostic purposes.
 *
 * @return HRESULT Returns S_OK if unregistration succeeds, otherwise an error code indicating failure.
 */

HRESULT WINAPI DllUnregisterServer() {
		LogMessage(L"DllUnregisterServer called."); // Log a message
		HRESULT hr = UnregisterShellExtension(); // Unregister the shell extension
		return hr; // Return the result of the unregistration operation
}

/**
 * @brief Registers the COM object as a shell extension.
 *
 * This function is responsible for registering the COM object as a shell extension in the Windows registry.
 * It creates necessary registry keys and sets appropriate values to associate the COM object with file types.
 * Logging messages are generated to provide diagnostic information during the registration process.
 *
 * @param hModule Handle to the DLL module containing the COM object.
 * @return HRESULT Returns S_OK if registration succeeds, otherwise an error code indicates failure.
 */

HRESULT RegisterShellExtension(HMODULE hModule) {
		HRESULT hr = S_OK; // Initialize result to success
		HKEY hKey = nullptr; // Handle to registry key
		DWORD dwDisposition; // Variable to receive disposition of the registry key

		// Define the path of the DLL module
		wchar_t szModule[MAX_PATH];
		GetModuleFileName(hModule, szModule, ARRAYSIZE(szModule));
		LogMessage(szModule); // Log the DLL module path

		// Create a registry key under HKEY_CLASSES_ROOT\CLSID\{05B54455-9F28-436D-9E1F-18484C5AB79A}
		hr = RegCreateKeyEx(HKEY_CLASSES_ROOT, L"CLSID\\{05B54455-9F28-436D-9E1F-18484C5AB79A}", 0, nullptr, REG_OPTION_NON_VOLATILE, KEY_WRITE, nullptr, &hKey, &dwDisposition);
		if (hr != ERROR_SUCCESS) {
				LogMessage(L"Failed to create registry key for CLSID."); // Log failure
				return hr; // Return error code
		}

		// Set the default value to "My Property Sheet Handler"
		hr = RegSetValueEx(hKey, nullptr, 0, REG_SZ, reinterpret_cast<const BYTE*>(L"My Property Sheet Handler"), sizeof(L"My Property Sheet Handler"));
		if (hr != ERROR_SUCCESS) {
				LogMessage(L"Failed to set default value for CLSID key."); // Log failure
				RegCloseKey(hKey); // Close the registry key
				return hr; // Return error code
		}
		RegCloseKey(hKey); // Close the registry key

		// Create a registry key under HKEY_CLASSES_ROOT\CLSID\{05B54455-9F28-436D-9E1F-18484C5AB79A}\InProcServer32
		hr = RegCreateKeyEx(HKEY_CLASSES_ROOT, L"CLSID\\{05B54455-9F28-436D-9E1F-18484C5AB79A}\\InprocServer32", 0, nullptr, REG_OPTION_NON_VOLATILE, KEY_WRITE, nullptr, &hKey, &dwDisposition);
		if (hr != ERROR_SUCCESS) {
				LogMessage(L"Failed to create registry key for InProcServer32."); // Log failure
				return hr; // Return error code
		}

		// Set the default value to the path of the DLL module
		hr = RegSetValueEx(hKey, nullptr, 0, REG_SZ, reinterpret_cast<const BYTE*>(szModule), (lstrlen(szModule) + 1) * sizeof(wchar_t));
		if (hr != ERROR_SUCCESS) {
				LogMessage(L"Failed to set default value for InProcServer32 key."); // Log failure
				RegCloseKey(hKey); // Close the registry key
				return hr; // Return error code
		}

		// Set the ThreadingModel value to "Apartment"
		hr = RegSetValueEx(hKey, L"ThreadingModel", 0, REG_SZ, reinterpret_cast<const BYTE*>(L"Apartment"), (wcslen(L"Apartment") + 1) * sizeof(wchar_t));
		if (hr != ERROR_SUCCESS) {
				LogMessage(L"Failed to set ThreadingModel value for InProcServer32 key."); // Log failure
				RegCloseKey(hKey); // Close the registry key
				return hr; // Return error code
		}
		RegCloseKey(hKey); // Close the registry key

		// Create registry entries to associate the property sheet handler with .txt files
		hr = RegCreateKeyEx(HKEY_CLASSES_ROOT, L"*\\shellex\\PropertySheetHandlers\\{05B54455-9F28-436D-9E1F-18484C5AB79A}", 0, nullptr, REG_OPTION_NON_VOLATILE, KEY_WRITE, nullptr, &hKey, &dwDisposition);
		if (hr != ERROR_SUCCESS) {
				LogMessage(L"Failed to create registry key for .txt files."); // Log failure
				return hr; // Return error code
		}

		// Set the default value to "My Property Sheet Handler"
		hr = RegSetValueEx(hKey, nullptr, 0, REG_SZ, reinterpret_cast<const BYTE*>(L"My Property Sheet Handler"), sizeof(L"My Property Sheet Handler"));
		if (hr != ERROR_SUCCESS) {
				LogMessage(L"Failed to set default value for .txt files key."); // Log failure
				RegCloseKey(hKey); // Close the registry key
				return hr; // Return error code
		}
		RegCloseKey(hKey); // Close the registry key

		return S_OK; // Return success
}

/**
 * @brief Unregisters the COM object as a shell extension.
 *
 * This function is responsible for unregistering the COM object as a shell extension from the Windows registry.
 * It deletes the registry keys associated with the COM object's registration.
 * Logging messages are generated to provide diagnostic information during the unregistering process.
 *
 * @return HRESULT Returns S_OK if unregistration succeeds, otherwise an error code indicating failure.
 */

HRESULT UnregisterShellExtension() {
		HRESULT hr = S_OK; // Initialize result to success

		// Delete registry entries for the property sheet handler under CLSID
		hr = RegDeleteTree(HKEY_CLASSES_ROOT, L"CLSID\\{05B54455-9F28-436D-9E1F-18484C5AB79A}");
		if (hr != ERROR_SUCCESS) {
				LogMessage(L"Failed to delete registry key for CLSID."); // Log failure
				return hr; // Return error code
		}

		// Delete registry entries for the property sheet handler under .txt files
		hr = RegDeleteTree(HKEY_CLASSES_ROOT, L"*\\shellex\\PropertySheetHandlers\\{05B54455-9F28-436D-9E1F-18484C5AB79A}");
		if (hr != ERROR_SUCCESS) {
				LogMessage(L"Failed to delete registry key for .txt files."); // Log failure
				return hr; // Return error code
		}

		return S_OK; // Return success
}