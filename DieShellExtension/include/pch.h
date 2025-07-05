// pch.h: This is a precompiled header file.
// Files listed below are compiled only once, improving build performance for future builds.
// This also affects IntelliSense performance, including code completion and many code browsing features.
// However, files listed here are ALL re-compiled if any one of them is updated between builds.
// Do not add files here that you will be updating frequently as this negates the performance advantage.

#ifndef PCH_H
#define PCH_H

// add headers that you want to pre-compile here
#include "framework.h"
#include <windows.h>
#include <unknwn.h>
#include <shlwapi.h>
#include <atlbase.h>
#include <Shobjidl.h>
#include <Shlobj.h>
#include <shellapi.h>
#include <string>
#include "Reg.h"
#include <strsafe.h>
#include <iostream>
#include <fstream>



void LogMessage(const std::wstring& message);
    



#define INITGUID
#include <guiddef.h>
const CLSID CLSID_DieMenu =
 { 0x1AC8EAF1, 0x9A6A, 0x471C, 0xB8, 0x7B, 0xB1, 0xDF, 0x2B, 0x8E, 0xA6, 0x2F } ;

#endif //PCH_H
