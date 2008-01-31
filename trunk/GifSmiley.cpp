// GifSmiley.cpp : Implementation of DLL Exports.

#include "stdafx.h"
#include "resource.h"
#include "GifSmileyCtrl.h"

// The module attribute causes DllMain, DllRegisterServer and DllUnregisterServer to be automatically implemented for you
[ module(dll, uuid = "{518667DE-537D-4790-8E93-48B9CE8EBAD2}", 
		 name = "GifSmiley", 
		 helpstring = "Miranda GifSmiley 1.0 Type Library",
		 resource_name = "IDR_GIFSMILEY", 
		 custom = { "a817e7a1-43fa-11d0-9e44-00aa00b6770a", "{B277F564-065F-469B-96D2-59D455411A7E}"})]

class CGifSmileyModule
{
public:
    BOOL WINAPI DllMain(DWORD dwReason, LPVOID lpReserved) 
    {
        return __super::DllMain(dwReason, lpReserved);
    }
	/*
    BOOL WINAPI DllRegisterServer() 
	{
		return FALSE;  return __super::DllRegisterServer();
	}
	BOOL WINAPI DllUnregisterServer() 
	{		
		return FALSE; //return __super::DllUnregisterServer();
	}
    */
};
		 