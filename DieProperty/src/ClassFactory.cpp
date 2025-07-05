
#include <windows.h>
#include <shlobj.h>

#include "dllmain.h"
#include "ClassFactory.h"
#include "Prop.h"

CShellExt* g_pShellExt = NULL;

/**
 * @brief Constructs a CClassFactory object.
 *
 * Initializes the reference count for the object to 1 and increments the global DLL reference count.
 * This constructor is responsible for managing the lifetime of the COM class factory object.
 */
CClassFactory::CClassFactory() {
	m_ObjRefCount = 1; // Initialize object reference count to 1
	g_DllRefCount++; // Increment global DLL reference count
}

/**
 * @brief Destructs a CClassFactory object.
 *
 * Decrements the global DLL reference count when the object is destroyed.
 * This destructor is responsible for cleaning up resources associated with the COM class factory object.
 */
CClassFactory::~CClassFactory() {
	g_DllRefCount--; // Decrement global DLL reference count
}

/**
 * @brief Queries for an interface on the class factory object.
 *
 * This method implements the QueryInterface method of the IUnknown interface for the class factory object.
 * It returns a pointer to the requested interface if supported by the object and increments the reference count.
 *
 * @param riid The interface identifier.
 * @param ppReturn Address of the pointer variable that receives the interface pointer requested in riid.
 * @return HRESULT Returns S_OK if the interface is supported, otherwise E_NOINTERFACE.
 */
STDMETHODIMP CClassFactory::QueryInterface(REFIID riid, LPVOID* ppReturn) {
		*ppReturn = NULL; // Initialize output parameter

		// Check if requested interface matches any supported interfaces
		if (IsEqualIID(riid, IID_IUnknown))
				*ppReturn = this; // Return pointer to IUnknown interface
		else if (IsEqualIID(riid, IID_IClassFactory))
				*ppReturn = (IClassFactory*)this; // Return pointer to IClassFactory interface

		// If requested interface is supported, add a reference and return S_OK
		if (*ppReturn) {
				LPUNKNOWN pUnk = (LPUNKNOWN)(*ppReturn); // Cast to LPUNKNOWN
				pUnk->AddRef(); // Increment reference count
				return S_OK; // Return success
		}

		return E_NOINTERFACE; // Requested interface not supported, return failure
}

/**
 * @brief Increments the object reference count.
 *
 * Increments the object reference count and returns the updated count.
 *
 * @return DWORD Returns the updated object reference count.
 */
STDMETHODIMP_(DWORD) CClassFactory::AddRef()
{
		return ++m_ObjRefCount; // Increment and return object reference count
}

/**
 * @brief Decrements the object reference count and deletes the object if the count reaches 0.
 *
 * Decrements the object reference count and deletes the object if the count reaches 0.
 * Returns the updated reference count.
 *
 * @return DWORD Returns the updated object reference count or 0 if the object is deleted.
 */
STDMETHODIMP_(DWORD) CClassFactory::Release()
{
		--m_ObjRefCount;

		// Log the current reference count
	// std::wstring message = L"Release: Reference count from CClassFactory is now " + std::to_wstring(m_ObjRefCount);
	// LogMessage(message);

		if (m_ObjRefCount == 0)
		{
				LogMessage(L"Release: Deleting the object."); // Log message indicating deletion
				delete this; // Delete object if reference count reaches 0
				return 0;
		}

		return m_ObjRefCount; // Return updated object reference count
}

/**
 * @brief Creates an instance of the COM object.
 *
 * Creates an instance of the COM object and queries for the specified interface.
 *
 * @param pUnknown Pointer to controlling IUnknown interface (should be NULL).
 * @param riid The interface identifier.
 * @param ppObject Address of the pointer variable that receives the interface pointer requested in riid.
 * @return HRESULT Returns S_OK if the object is created successfully, otherwise E_OUTOFMEMORY.
 */
STDMETHODIMP CClassFactory::CreateInstance(LPUNKNOWN pUnknown, REFIID riid, LPVOID* ppObject)
{

		*ppObject = NULL; // Initialize output parameter
		if (pUnknown != NULL)
				return CLASS_E_NOAGGREGATION; // Aggregation not supported

		// Create an instance of the COM object
		CShellExt* pShellExt = new CShellExt();
		if (NULL == pShellExt)
				return E_OUTOFMEMORY; // Failed to allocate memory

		g_pShellExt = pShellExt;

		HRESULT hResult = pShellExt->QueryInterface(riid, ppObject); // Query interface for the requested interface

		if (FAILED(hResult)) {
				pShellExt->Release();
				delete pShellExt; // Delete the object to avoid memory leak
				return hResult;
		}


		return hResult; // Return result of querying interface
}
