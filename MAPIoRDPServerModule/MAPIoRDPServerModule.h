// Le bloc ifdef suivant est la fa�on standard de cr�er des macros qui facilitent l'exportation 
// � partir d'une DLL. Tous les fichiers contenus dans cette DLL sont compil�s avec le symbole MAPIORDPSERVERMODULE_EXPORTS
// d�fini sur la ligne de commande. Ce symbole ne doit pas �tre d�fini pour un projet
// qui utilisent cette DLL. De cette mani�re, les autres projets dont les fichiers sources comprennent ce fichier consid�rent les fonctions 
// MAPIORDPSERVERMODULE_API comme �tant import�es � partir d'une DLL, tandis que cette DLL consid�re les symboles
// d�finis avec cette macro comme �tant export�s.
#ifdef MAPIORDPSERVERMODULE_EXPORTS
#define MAPIORDPSERVERMODULE_API __declspec(dllexport)
#else
#define MAPIORDPSERVERMODULE_API __declspec(dllimport)
#endif

// Cette classe est export�e de MAPIoRDPServerModule.dll
class MAPIORDPSERVERMODULE_API CMAPIoRDPServerModule {
public:
	CMAPIoRDPServerModule(void);
	// TODO: ajoutez ici vos m�thodes.
};

extern MAPIORDPSERVERMODULE_API int nMAPIoRDPServerModule;

MAPIORDPSERVERMODULE_API int fnMAPIoRDPServerModule(void);
