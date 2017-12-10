// MAPIoRDPServerModule.cpp : définit les fonctions exportées pour l'application DLL.
//

#include "stdafx.h"
#include "MAPIoRDPServerModule.h"


// Il s'agit d'un exemple de variable exportée
MAPIORDPSERVERMODULE_API int nMAPIoRDPServerModule=0;

// Il s'agit d'un exemple de fonction exportée.
MAPIORDPSERVERMODULE_API int fnMAPIoRDPServerModule(void)
{
    return 42;
}

// Il s'agit du constructeur d'une classe qui a été exportée.
// consultez MAPIoRDPServerModule.h pour la définition de la classe
CMAPIoRDPServerModule::CMAPIoRDPServerModule()
{
    return;
}
