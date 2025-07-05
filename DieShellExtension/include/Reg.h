/****************************** Module Header ******************************\
Module Name:  Reg.h
Project:      CppShellExtContextMenuHandler
Copyright (c) Microsoft Corporation.

The file declares reusable helper functions to register and unregister
in-process COM components and shell context menu handlers in the registry.

RegisterInprocServer - register the in-process component in the registry.
UnregisterInprocServer - unregister the in-process component in the registry.
RegisterShellExtContextMenuHandler - register the context menu handler.
UnregisterShellExtContextMenuHandler - unregister the context menu handler.

This source is subject to the Microsoft Public License.
See http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
All other rights reserved.

THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND,
EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED
WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
\***************************************************************************/

#pragma once

#include <windows.h>

HRESULT RegisterInprocServer(PCWSTR pszModule, const CLSID& clsid,
PCWSTR pszFriendlyName, PCWSTR pszThreadModel);
HRESULT UnregisterInprocServer(const CLSID& clsid);
HRESULT RegisterShellExtContextMenuHandler(
PCWSTR pszFileType, const CLSID& clsid, PCWSTR pszFriendlyName);
HRESULT UnregisterShellExtContextMenuHandler(
PCWSTR pszFileType, const CLSID& clsid);