// MAPIoRDPMailto.cpp : définit le point d'entrée pour l'application.
//

#include "stdafx.h"
#include "MAPIoRDPMailto.h"

#include <boost\iostreams\device\back_inserter.hpp>
#include <boost\iostreams\stream.hpp>
#include <boost\archive\text_oarchive.hpp>
#include <boost\archive\text_iarchive.hpp>

#define _MAX_WAIT       60000
#define MAX_MSG_SIZE    0x20000

#include "DVCServerModule.h"
#include "MAPISerializer.hpp"

using namespace boost::serialization;

#define MAX_LOADSTRING 100

// Variables globales :
HINSTANCE hInst;                                // instance actuelle
WCHAR szTitle[MAX_LOADSTRING];                  // Le texte de la barre de titre
WCHAR szWindowClass[MAX_LOADSTRING];            // le nom de la classe de fenêtre principale

// Pré-déclarations des fonctions incluses dans ce module de code :
ATOM                MyRegisterClass(HINSTANCE hInstance);
BOOL                InitInstance(HINSTANCE, int);
LRESULT CALLBACK    WndProc(HWND, UINT, WPARAM, LPARAM);
INT_PTR CALLBACK    About(HWND, UINT, WPARAM, LPARAM);

int APIENTRY wWinMain(_In_ HINSTANCE hInstance,
                     _In_opt_ HINSTANCE hPrevInstance,
                     _In_ LPWSTR    lpCmdLine,
                     _In_ int       nCmdShow)
{
    UNREFERENCED_PARAMETER(hPrevInstance);
    
	lpCmdLine;

	buffer_t buffer;

	{
		boost::iostreams::back_insert_device<buffer_t> inserter(buffer);
		boost::iostreams::stream<boost::iostreams::back_insert_device<buffer_t> > s(inserter);
		boost::archive::text_oarchive ar(s);

		ar << std::string("MAILTO");
		save_c_str(ar, lpCmdLine);
		s.flush();
		s.close();
	}

	{
		DWORD rc = CallFunctionOverRDP(buffer);
		if (rc != SUCCESS_SUCCESS)
		{
			MessageBox(NULL, L"Could not call MAILTO over RDP.", L"MAPIoRDP", MB_OK | MB_ICONERROR);
			return rc;
		}
	}

	{
		boost::iostreams::basic_array_source<char> device(buffer.data(), buffer.size());
		boost::iostreams::stream<boost::iostreams::basic_array_source<char> > s(device);
		boost::archive::text_iarchive ar(s);

		int rc;
		ar >> rc;
		return rc;
	}

/*
    // Initialise les chaînes globales
    LoadStringW(hInstance, IDS_APP_TITLE, szTitle, MAX_LOADSTRING);
    LoadStringW(hInstance, IDC_MAPIORDPMAILTO, szWindowClass, MAX_LOADSTRING);
    MyRegisterClass(hInstance);

    // Effectue l'initialisation de l'application :
    if (!InitInstance (hInstance, nCmdShow))
    {
        return FALSE;
    }

    HACCEL hAccelTable = LoadAccelerators(hInstance, MAKEINTRESOURCE(IDC_MAPIORDPMAILTO));

    MSG msg;

    // Boucle de messages principale :
    while (GetMessage(&msg, nullptr, 0, 0))
    {
        if (!TranslateAccelerator(msg.hwnd, hAccelTable, &msg))
        {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
    }

    return (int) msg.wParam;
*/
}



//
//  FONCTION : MyRegisterClass()
//
//  BUT : inscrit la classe de fenêtre.
//
ATOM MyRegisterClass(HINSTANCE hInstance)
{
    WNDCLASSEXW wcex;

    wcex.cbSize = sizeof(WNDCLASSEX);

    wcex.style          = CS_HREDRAW | CS_VREDRAW;
    wcex.lpfnWndProc    = WndProc;
    wcex.cbClsExtra     = 0;
    wcex.cbWndExtra     = 0;
    wcex.hInstance      = hInstance;
    wcex.hIcon          = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_MAPIORDPMAILTO));
    wcex.hCursor        = LoadCursor(nullptr, IDC_ARROW);
    wcex.hbrBackground  = (HBRUSH)(COLOR_WINDOW+1);
    wcex.lpszMenuName   = MAKEINTRESOURCEW(IDC_MAPIORDPMAILTO);
    wcex.lpszClassName  = szWindowClass;
    wcex.hIconSm        = LoadIcon(wcex.hInstance, MAKEINTRESOURCE(IDI_SMALL));

    return RegisterClassExW(&wcex);
}

//
//   FONCTION : InitInstance(HINSTANCE, int)
//
//   BUT : enregistre le handle de l'instance et crée une fenêtre principale
//
//   COMMENTAIRES :
//
//        Dans cette fonction, nous enregistrons le handle de l'instance dans une variable globale, puis
//        créons et affichons la fenêtre principale du programme.
//
BOOL InitInstance(HINSTANCE hInstance, int nCmdShow)
{
   hInst = hInstance; // Stocke le handle d'instance dans la variable globale

   HWND hWnd = CreateWindowW(szWindowClass, szTitle, WS_OVERLAPPEDWINDOW,
      CW_USEDEFAULT, 0, CW_USEDEFAULT, 0, nullptr, nullptr, hInstance, nullptr);

   if (!hWnd)
   {
      return FALSE;
   }

   ShowWindow(hWnd, nCmdShow);
   UpdateWindow(hWnd);

   return TRUE;
}

//
//  FONCTION : WndProc(HWND, UINT, WPARAM, LPARAM)
//
//  BUT :  traite les messages pour la fenêtre principale.
//
//  WM_COMMAND - traite le menu de l'application
//  WM_PAINT - dessine la fenêtre principale
//  WM_DESTROY - génère un message d'arrêt et retourne
//
//
LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
    switch (message)
    {
    case WM_COMMAND:
        {
            int wmId = LOWORD(wParam);
            // Analyse les sélections de menu :
            switch (wmId)
            {
            case IDM_ABOUT:
                DialogBox(hInst, MAKEINTRESOURCE(IDD_ABOUTBOX), hWnd, About);
                break;
            case IDM_EXIT:
                DestroyWindow(hWnd);
                break;
            default:
                return DefWindowProc(hWnd, message, wParam, lParam);
            }
        }
        break;
    case WM_PAINT:
        {
            PAINTSTRUCT ps;
            HDC hdc = BeginPaint(hWnd, &ps);
            // TODO: ajoutez le code de dessin qui utilise hdc ici...
            EndPaint(hWnd, &ps);
        }
        break;
    case WM_DESTROY:
        PostQuitMessage(0);
        break;
    default:
        return DefWindowProc(hWnd, message, wParam, lParam);
    }
    return 0;
}

// Gestionnaire de messages pour la boîte de dialogue À propos de.
INT_PTR CALLBACK About(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
    UNREFERENCED_PARAMETER(lParam);
    switch (message)
    {
    case WM_INITDIALOG:
        return (INT_PTR)TRUE;

    case WM_COMMAND:
        if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL)
        {
            EndDialog(hDlg, LOWORD(wParam));
            return (INT_PTR)TRUE;
        }
        break;
    }
    return (INT_PTR)FALSE;
}
