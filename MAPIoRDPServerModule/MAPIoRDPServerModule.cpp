// MAPIoRDPServerModule.cpp�: d�finit les fonctions export�es pour l'application DLL.
//

#include "stdafx.h"
#include "MAPIoRDPServerModule.h"


// Il s'agit d'un exemple de variable export�e
MAPIORDPSERVERMODULE_API int nMAPIoRDPServerModule=0;

// Il s'agit d'un exemple de fonction export�e.
MAPIORDPSERVERMODULE_API int fnMAPIoRDPServerModule(void)
{
    return 42;
}

// Il s'agit du constructeur d'une classe qui a �t� export�e.
// consultez MAPIoRDPServerModule.h pour la d�finition de la classe
CMAPIoRDPServerModule::CMAPIoRDPServerModule()
{
    return;
}
