// MAPIoRDPClientPlugin.cpp : implémentation des exportations de DLL.

//
// Remarque : Informations COM+ 1.0 :
//      N'oubliez pas d'exécuter Microsoft Transaction Explorer pour installer les composants.
//      L'inscription n'est pas automatique. 

#include "stdafx.h"
#include "resource.h"
#include "MAPIoRDPClientPlugin_i.h"
#include "dllmain.h"


using namespace ATL;

// Utilisé pour déterminer si la DLL peut être déchargée par OLE.
STDAPI DllCanUnloadNow(void)
{
			return _AtlModule.DllCanUnloadNow();
	}

// Retourne une fabrique de classes pour créer un objet du type demandé.
_Check_return_
STDAPI DllGetClassObject(_In_ REFCLSID rclsid, _In_ REFIID riid, _Outptr_ LPVOID* ppv)
{
		return _AtlModule.DllGetClassObject(rclsid, riid, ppv);
}

// DllRegisterServer - Ajoute des entrées à la base de registres.
STDAPI DllRegisterServer(void)
{
	// inscrit l'objet, la typelib et toutes les interfaces dans la typelib
	HRESULT hr = _AtlModule.DllRegisterServer();
		return hr;
}

// DllUnregisterServer - Supprime des entrées de la base de registres.
STDAPI DllUnregisterServer(void)
{
	HRESULT hr = _AtlModule.DllUnregisterServer();
		return hr;
}

// DllInstall - Ajoute/supprime des entrées de la base de registres par utilisateur et par ordinateur.
STDAPI DllInstall(BOOL bInstall, _In_opt_  LPCWSTR pszCmdLine)
{
	HRESULT hr = E_FAIL;
	static const wchar_t szUserSwitch[] = L"user";

	if (pszCmdLine != NULL)
	{
		if (_wcsnicmp(pszCmdLine, szUserSwitch, _countof(szUserSwitch)) == 0)
		{
			ATL::AtlSetPerUserRegistration(true);
		}
	}

	if (bInstall)
	{	
		hr = DllRegisterServer();
		if (FAILED(hr))
		{
			DllUnregisterServer();
		}
	}
	else
	{
		hr = DllUnregisterServer();
	}

	return hr;
}


