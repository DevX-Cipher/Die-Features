#include <windows.h>
#include <atlbase.h>
#include <shlobj.h>
#include <shobjidl.h>
#include <shellapi.h>
#include <shlwapi.h>
#include <vector>
#include <string>
#include "Prop.h"
#include <sstream>
#include <iomanip>
#include "resource.h"
#include "ClassFactory.h"
#include "dllmain.h"
#include <regex>
#include <numeric>

INT_PTR CALLBACK MyPropPageDlgProc(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam);


CShellExt::CShellExt()
{
    m_ObjRefCount = 1;
}
CShellExt::~CShellExt() {}


/**
 * @brief Releases a COM object.
 *
 * Decrements the reference count of the COM object. If the reference count reaches zero,
 * the object is deleted.
 *
 * @return DWORD The updated reference count after releasing the object.
 */
STDMETHODIMP_(DWORD) CShellExt::Release(void) {
		if (--m_ObjRefCount == 0) {
				LogMessage(L"Release: Deleting the object."); // Log message indicating deletion
				LogMessage(L"Deleting CShellExt object. File path: " + std::wstring(m_szFile));

				delete this; // Delete the object
				return 0; // Return 0 to indicate object deletion
		}
		return m_ObjRefCount; // Return updated reference count
}

/**
  * @brief Increments the reference count of a COM object.
  *
  * Increments the reference count of the COM object to indicate that this object
  * has another pointer to it.
  *
  * @return DWORD The updated reference count after adding a reference.
  */
STDMETHODIMP_(DWORD) CShellExt::AddRef(void) {
		LogMessage(L"AddRef: Reference count incremented to " + std::to_wstring(m_ObjRefCount)); // Log message indicating reference count increment
		return ++m_ObjRefCount; // Increment reference count and return updated count
}

/**
  * @brief Queries for an interface on a COM object.
  *
  * Queries the specified interface on the COM object and returns a pointer to that interface.
  *
  * @param riid The interface identifier.
  * @param ppReturn Address of the pointer variable that receives the interface pointer requested in riid.
  * @return HRESULT Returns S_OK if the interface is supported, otherwise E_NOINTERFACE.
  */
STDMETHODIMP CShellExt::QueryInterface(REFIID riid, LPVOID* ppReturn)
{
		*ppReturn = NULL; // Initialize output parameter

		// Check if requested interface matches any supported interfaces
		if (IsEqualIID(riid, IID_IUnknown))
				*ppReturn = this; // Return pointer to IUnknown interface
		else if (IsEqualIID(riid, IID_IClassFactory))
				*ppReturn = (IClassFactory*)this; // Return pointer to IClassFactory interface
		else if (IsEqualIID(riid, IID_IShellExtInit))
				*ppReturn = (IShellExtInit*)this; // Return pointer to IShellExtInit interface
		else if (IsEqualIID(riid, IID_IShellPropSheetExt))
				*ppReturn = (IShellPropSheetExt*)this; // Return pointer to IShellPropSheetExt interface

		// If requested interface is supported, add a reference and return S_OK
		if (*ppReturn)
		{
				LPUNKNOWN pUnk = (LPUNKNOWN)(*ppReturn); // Cast to LPUNKNOWN
				pUnk->AddRef(); // Increment reference count
				return S_OK; // Return success
		}

		return E_NOINTERFACE; // Requested interface not supported, return failure
}

/**
 * @brief Initializes the shell extension with data from the data object.
 *
 * This method initializes the shell extension with data retrieved from the data object.
 * It logs messages indicating the initialization process and any errors encountered.
 *
 * @param pidlFolder Pointer to an ITEMIDLIST structure that specifies the parent folder's location.
 * @param pDataObject Pointer to an IDataObject interface from which the initialization data is retrieved.
 * @param hProgID Handle to the registry key for the application.
 * @return HRESULT Returns S_OK if initialization succeeds, otherwise returns an error code.
 */
STDMETHODIMP CShellExt::Initialize(LPCITEMIDLIST pidlFolder, LPDATAOBJECT pDataObject, HKEY hProgID)
{
		LogMessage(L"Initialize method called."); // Log a message indicating the method call

		FORMATETC fmt = { CF_HDROP, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL }; // Specify the format of the data to be retrieved
		STGMEDIUM stg = { TYMED_HGLOBAL }; // Storage medium for the data
		HDROP hDrop;

		// Look for CF_HDROP data in the data object.
		if (FAILED(pDataObject->GetData(&fmt, &stg))) // Attempt to retrieve the data from the data object
		{
				// Failed to retrieve data. Log an error message and return E_INVALIDARG.
				LogMessage(L"Error: Failed to get data from data object."); // Log an error message
				return E_INVALIDARG; // Return an error code
		}

		// Get a pointer to the actual data.
		hDrop = (HDROP)GlobalLock(stg.hGlobal); // Lock the global memory handle to get a pointer to the data

		// Ensure that the pointer to the data is valid.
		if (NULL == hDrop)
		{
				ReleaseStgMedium(&stg); // Release the storage medium
				LogMessage(L"Error: Failed to get a pointer to the actual data."); // Log an error message
				return E_INVALIDARG; // Return an error code
		}

		// Sanity check: ensure there is at least one filename.
		UINT uNumFiles = DragQueryFile(hDrop, 0xFFFFFFFF, NULL, 0); // Get the number of files dropped
		HRESULT hr = S_OK;

		if (0 == uNumFiles)
		{
				GlobalUnlock(stg.hGlobal); // Unlock the global memory handle
				ReleaseStgMedium(&stg); // Release the storage medium
				LogMessage(L"Error: No filenames found."); // Log an error message
				return E_INVALIDARG; // Return an error code
		}

		// In this simple sample, take the first file. In reality, you'd probably want to store them all for later use.
		UINT uRet = DragQueryFile(hDrop, 0, m_szFile, MAX_PATH);
		if (0 == uRet)
		{
				LogMessage(L"Error: DragQueryFile failed."); // Log an error message
				hr = E_INVALIDARG;
		}
		else
		{
				LogMessage(L"DragQueryFile succeeded. File path: " + std::wstring(m_szFile)); // Log the file path
				m_ObjRefCount = 1;
		}

		GlobalUnlock(stg.hGlobal); // Unlock the global memory handle
		ReleaseStgMedium(&stg); // Release the storage medium

		LogMessage(L"Initialize method completed."); // Log a message indicating completion
		return hr; // Return the result of the operation
}

/**
 * @brief Adds a property page to the property sheet.
 *
 * This method adds a property page to the property sheet using the provided function pointer.
 * It initializes the property sheet page structure, creates the property sheet page, and adds it to the property sheet.
 * Messages are logged to indicate the progress and any errors encountered during the process.
 *
 * @param pfnAddPage Pointer to the function used to add a page to the property sheet.
 * @param lParam Additional parameter passed to the function when adding the page.
 * @return HRESULT Returns S_OK if the page is successfully added, otherwise returns an error code.
 */
STDMETHODIMP CShellExt::AddPages(LPFNSVADDPROPSHEETPAGE pfnAddPage, LPARAM lParam) {
		LogMessage(L"AddPages called. Object instance: " + std::to_wstring((uintptr_t)this));
		LogMessage(L"AddPages called. Adding property page."); // Log a message indicating method call
		LogMessage(L"Current value of m_szFile: " + std::wstring(m_szFile));

		PROPSHEETPAGE psp; // Property sheet page structure
		HPROPSHEETPAGE hPage; // Handle to the property sheet page 
		// Initialize the property sheet page structure.
		ZeroMemory(&psp, sizeof(psp)); // Clear the memory for the structure
		LogMessage(L"Property sheet page structure initialized."); // Log a message
		psp.dwSize = sizeof(psp); // Set the size of the structure
		psp.dwFlags = PSP_USEREFPARENT | PSP_USETITLE; // Set flags for the property sheet page
		psp.hInstance = g_hInst; // Set the instance handle
		psp.pszTemplate = MAKEINTRESOURCE(IDD_MYPROPERTYPAGE); // Set the resource template for the page
		psp.pszIcon = NULL; // Set the icon for the page
		psp.pszTitle = L"DIE"; // Set the title for the page
		psp.pfnDlgProc = MyPropPageDlgProc; // Set the dialog procedure for the page
		psp.lParam = reinterpret_cast<LPARAM>(this);
		AddRef(); // Increment the reference count

		psp.pfnCallback = NULL; // Set the callback function for the page
		psp.pcRefParent = (UINT*)&m_ObjRefCount; // Set a reference count for the parent object

		// Create the property sheet page and get a handle to it.
		hPage = CreatePropertySheetPage(&psp); // Create the property sheet page

		if (hPage == NULL) {
				LogMessage(L"Error: Failed to create property sheet page."); // Log an error message
				return E_OUTOFMEMORY; // Return an error code
		}

		// Call the function specified by pfnAddPage to add the page to the property sheet.
		if (!(*pfnAddPage)(hPage, lParam)) {
				DestroyPropertySheetPage(hPage); // Destroy the property sheet page
				LogMessage(L"Error: Failed to add page to property sheet."); // Log an error message
				return E_FAIL; // Return an error code
		}
		LogMessage(L"m_szFile after AddPages: " + std::wstring(m_szFile));
		LogMessage(L"AddPages method completed."); // Log a message indicating completion
		return S_OK; // Return success
}

/**
* @brief Replaces property pages of the shell extension.
*
* Replaces existing property pages of the shell extension with new ones.
*
* @param uPageID Identifier of the page to be replaced.
* @param lpfnReplaceWith Function pointer to the function that replaces a property sheet page.
* @param lParam Value to be passed to the lpfnReplaceWith function.
* @return HRESULT Returns S_OK if property pages are successfully replaced, otherwise an error code.
*/
STDMETHODIMP CShellExt::ReplacePage(UINT uPageID, LPFNSVADDPROPSHEETPAGE  lpfnReplaceWith, LPARAM lParam) {
		return E_NOTIMPL;
}

/**
 * @brief Dialog procedure for the property page.
 *
 * This function processes messages sent to the property page dialog.
 * It handles initialization, control notifications, and destruction messages.
 * Relevant messages are logged to provide insight into the dialog's behavior.
 *
 * @param hDlg Handle to the dialog box window.
 * @param message The message identifier.
 * @param wParam Additional message information.
 * @param lParam Additional message information.
 * @return INT_PTR Returns TRUE if the message was handled, otherwise FALSE.
 */
INT_PTR CALLBACK MyPropPageDlgProc(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam) {

		switch (message) {
		case WM_INITDIALOG:
		{

				if (g_pShellExt == nullptr) {
						g_pShellExt = reinterpret_cast<CShellExt*>(lParam);

				}

				LogMessage(L"WM_INITDIALOG received. Object instance: " + std::to_wstring((uintptr_t)g_pShellExt));
				LogMessage(L"m_szFile in WM_INITDIALOG: " + std::wstring(g_pShellExt->m_szFile));

				LogMessage(L"WM_INITDIALOG received."); // Log a message indicating initialization
				SetWindowText(hDlg, reinterpret_cast<LPCTSTR>(lParam)); // Set window title to the file path

				std::wstring filePath = g_pShellExt->m_szFile;
				LogMessage(L"File path: " + filePath); // Log the file path

				TCHAR szDllPath[MAX_PATH];
				GetModuleFileName((HINSTANCE)&__ImageBase, szDllPath, MAX_PATH);
				std::wstring wsDllPath(szDllPath); // Convert to std::wstring for easier manipulation      
				size_t pos = wsDllPath.find_last_of(L"\\"); // Find the last backslash (the start of the filename)   
				std::wstring databasePath = wsDllPath.substr(0, pos + 1) + L"db"; // Construct the database path by replacing the filename with "db"
				LogMessage(L"Database Path: " + databasePath);

				std::wstring customInfo = GetFileInformation(filePath, databasePath); // Call your custom function to retrieve information

				// Set text of IDC_FILEINFO textbox with the retrieved information
				HWND hFileInfoTextbox = GetDlgItem(hDlg, IDC_FILEINFO); // Get handle to IDC_FILEINFO textbox
				SetWindowText(hFileInfoTextbox, customInfo.c_str()); // Set text of the textbox

				return TRUE;
		}

		case WM_COMMAND:
				LogMessage(L"WM_COMMAND received."); // Log a message indicating command reception
				switch (LOWORD(wParam)) {
				case IDOK:
				case IDCANCEL:
						LogMessage(L"IDOK or IDCANCEL received."); // Log a message indicating button press
						EndDialog(hDlg, LOWORD(wParam)); // Close the dialog with the specified result
						return TRUE; // Indicate that the message was handled
				}
				break;

		case WM_DESTROY:
				g_pShellExt->Release(); // Decrement the reference count
				LogMessage(L"WM_DESTROY received."); // Log a message indicating destruction
				break;
		}

		return FALSE; // Indicate that the message was not handled
}

std::wstring GetFileInformation(const std::wstring& filePath, const std::wstring& pwszDatabase)
{
		try {
				const int nBufferSize = 10000;
				std::wstring sBuffer(nBufferSize, L' ');  // Allocate buffer

				// Use DIE_VB_ScanFile to get the file information
				int nResult = DIE_VB_ScanFile(filePath.c_str(), 0, pwszDatabase.c_str(), &sBuffer[0], nBufferSize - 1);

				std::wstring result;
				if (nResult > 0) {
						// Trim the buffer to the actual length of the result
						sBuffer.resize(nResult);
						result = sBuffer;
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
