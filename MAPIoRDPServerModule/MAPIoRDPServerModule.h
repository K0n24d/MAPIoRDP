// Le bloc ifdef suivant est la façon standard de créer des macros qui facilitent l'exportation 
// à partir d'une DLL. Tous les fichiers contenus dans cette DLL sont compilés avec le symbole MAPIORDPSERVERMODULE_EXPORTS
// défini sur la ligne de commande. Ce symbole ne doit pas être défini pour un projet
// qui utilisent cette DLL. De cette manière, les autres projets dont les fichiers sources comprennent ce fichier considèrent les fonctions 
// MAPIORDPSERVERMODULE_API comme étant importées à partir d'une DLL, tandis que cette DLL considère les symboles
// définis avec cette macro comme étant exportés.
#ifdef MAPIORDPSERVERMODULE_EXPORTS
#define MAPIORDPSERVERMODULE_API __declspec(dllexport)
#else
#define MAPIORDPSERVERMODULE_API __declspec(dllimport)
#endif

// Cette classe est exportée de MAPIoRDPServerModule.dll
class MAPIORDPSERVERMODULE_API CMAPIoRDPServerModule {
public:
	CMAPIoRDPServerModule(void);
	// TODO: ajoutez ici vos méthodes.
};

extern MAPIORDPSERVERMODULE_API int nMAPIoRDPServerModule;

MAPIORDPSERVERMODULE_API int fnMAPIoRDPServerModule(void);
