// dllmain.cpp : implémentation de DllMain.

#include "stdafx.h"
#include "resource.h"
#include "MAPIoRDPClientPlugin_i.h"
#include "dllmain.h"

CMAPIoRDPClientPluginModule _AtlModule;
HMODULE mapi_hLib;

// Point d'entrée de la DLL
extern "C" BOOL WINAPI DllMain(HINSTANCE hInstance, DWORD dwReason, LPVOID lpReserved)
{
	switch (dwReason)
	{
	case DLL_PROCESS_ATTACH:
		mapi_hLib = LoadLibrary(_T("MAPI32.DLL"));
		break;
	case DLL_THREAD_ATTACH:
		break;
	case DLL_THREAD_DETACH:
		break;
	case DLL_PROCESS_DETACH:
		FreeLibrary(mapi_hLib);
		break;
	}
	return _AtlModule.DllMain(dwReason, lpReserved); 
}
