// Mapi32.cpp : définit les fonctions SimpleMAPI exportées.
//

#include "stdafx.h"
#include <MAPI.h>
#include <boost\iostreams\device\back_inserter.hpp>
#include <boost\iostreams\stream.hpp>
#include <boost\archive\text_oarchive.hpp>
#include <boost\archive\text_iarchive.hpp>
#include "MAPISerializer.hpp"
#include "Tools.h"

using namespace boost::serialization;
extern HMODULE mapi_hLib;;

template<class ArchiveIn, class ArchiveOut>
void cMAPIAddress(ArchiveIn & ai, ArchiveOut & ao)
{
	LHANDLE         lhSession;
	ULONG_PTR       ulUIParam;
	LPSTR           lpszCaption;
	ULONG           nEditFields;
	LPSTR           lpszLabels;
	ULONG           nRecips;
	lpMapiRecipDesc lpRecips;
	FLAGS           flFlags;
	ULONG           ulReserved;
	ULONG           nNewRecips;
	lpMapiRecipDesc lpNewRecips;

	ai >> lhSession;
	ai >> ulUIParam;
	load_c_str(ai, lpszCaption);
	ai >> nEditFields;
	load_c_str(ai, lpszLabels);
	ai >> nRecips;
	lpRecips = (lpMapiRecipDesc)calloc(nRecips, sizeof(MapiRecipDesc));
	for (ULONG i = 0; i < nRecips; i++)
		ai >> lpRecips[i];
	ai >> flFlags;
	ai >> ulReserved;

	LPMAPIADDRESS lpMAPIAddress;
	lpMAPIAddress = (LPMAPIADDRESS)GetProcAddress(mapi_hLib, "MAPIAddress");

	LPMAPIFREEBUFFER lpMAPIFreeBuffer;
	lpMAPIFreeBuffer = (LPMAPIFREEBUFFER)GetProcAddress(mapi_hLib, "MAPIFreeBuffer");

	ULONG rc;
	if (!lpMAPIAddress || !lpMAPIFreeBuffer)
	{
		rc = MAPI_E_FAILURE;
		ao << rc;
	}
	else
	{
		rc = lpMAPIAddress(lhSession, ulUIParam, lpszCaption, nEditFields, lpszLabels, nRecips, lpRecips, flFlags, ulReserved, &nNewRecips, &lpNewRecips);
		ao << rc;
		ao << nNewRecips;
		for (ULONG i = 0; i < nRecips; i++)
		{
			ao << lpNewRecips[i];
			lpMAPIFreeBuffer(&(lpNewRecips[i]));
		}
		for (ULONG i = 0; i < nRecips; i++)
			FreeMAPIRecipient(&(lpRecips[i]));
		free(lpRecips);
		free(lpszCaption);
		free(lpszLabels);
	}
}

/*
ULONG WINAPI MAPIDeleteMail(
	_In_ LHANDLE   lhSession,
	_In_ ULONG_PTR ulUIParam,
	_In_ LPSTR     lpszMessageID,
	     FLAGS     flFlags,
	     ULONG     ulReserved
)
{
	buffer_t buffer;

	{
		boost::iostreams::back_insert_device<buffer_t> inserter(buffer);
		boost::iostreams::stream<boost::iostreams::back_insert_device<buffer_t> > s(inserter);
		boost::archive::text_oarchive ar(s);

		ar << std::string(__func__);
		ar << lhSession;
		ar << ulUIParam;
		save_c_str(ar, lpszMessageID);
		ar << flFlags;
		ar << ulReserved;
		s.flush();
		s.close();
	}

	{
		DWORD rc = CallFunctionOverRDP(buffer);
		if (rc != SUCCESS_SUCCESS)
		{
			MessageBox(NULL, L"MAPIoRDP", L"Could not call MAPI function over RDP.", MB_OK | MB_ICONERROR);
			return rc;
		}
	}

	{
		boost::iostreams::basic_array_source<char> device(buffer.data(), buffer.size());
		boost::iostreams::stream<boost::iostreams::basic_array_source<char> > s(device);
		boost::archive::text_iarchive ar(s);

		ULONG rc;
		ar >> rc;

		return rc;
	}
}

ULONG WINAPI MAPIDetails(
	_In_ LHANDLE         lhSession,
	_In_ ULONG_PTR       ulUIParam,
	_In_ lpMapiRecipDesc lpRecip,
	_In_ FLAGS           flFlags,
	     ULONG           ulReserved
)
{
	buffer_t buffer;

	{
		boost::iostreams::back_insert_device<buffer_t> inserter(buffer);
		boost::iostreams::stream<boost::iostreams::back_insert_device<buffer_t> > s(inserter);
		boost::archive::text_oarchive ar(s);

		ar << std::string(__func__);
		ar << lhSession;
		ar << ulUIParam;
		ar << *lpRecip;
		ar << flFlags;
		ar << ulReserved;
		s.flush();
		s.close();
	}

	{
		DWORD rc = CallFunctionOverRDP(buffer);
		if (rc != SUCCESS_SUCCESS)
		{
			MessageBox(NULL, L"MAPIoRDP", L"Could not call MAPI function over RDP.", MB_OK | MB_ICONERROR);
			return rc;
		}
	}

	{
		boost::iostreams::basic_array_source<char> device(buffer.data(), buffer.size());
		boost::iostreams::stream<boost::iostreams::basic_array_source<char> > s(device);
		boost::archive::text_iarchive ar(s);

		ULONG rc;
		ar >> rc;

		return rc;
	}
}

ULONG WINAPI MAPIFindNext(
	_In_  LHANDLE   lhSession,
	_In_  ULONG_PTR ulUIParam,
	_In_  LPSTR     lpszMessageType,
	_In_  LPSTR     lpszSeedMessageID,
	_In_  FLAGS     flFlags,
	      ULONG     ulReserved,
	_Out_ LPSTR     lpszMessageID
)
{
	buffer_t buffer;

	{
		boost::iostreams::back_insert_device<buffer_t> inserter(buffer);
		boost::iostreams::stream<boost::iostreams::back_insert_device<buffer_t> > s(inserter);
		boost::archive::text_oarchive ar(s);

		ar << std::string(__func__);
		ar << lhSession;
		ar << ulUIParam;
		save_c_str(ar, lpszMessageType);
		save_c_str(ar, lpszSeedMessageID);
		ar << flFlags;
		ar << ulReserved;
		s.flush();
		s.close();
	}

	{
		DWORD rc = CallFunctionOverRDP(buffer);
		if (rc != SUCCESS_SUCCESS)
		{
			MessageBox(NULL, L"MAPIoRDP", L"Could not call MAPI function over RDP.", MB_OK | MB_ICONERROR);
			return rc;
		}
	}

	{
		boost::iostreams::basic_array_source<char> device(buffer.data(), buffer.size());
		boost::iostreams::stream<boost::iostreams::basic_array_source<char> > s(device);
		boost::archive::text_iarchive ar(s);

		ULONG rc;
		ar >> rc;
		load_c_str(ar, lpszMessageID);

		return rc;
	}
}

ULONG WINAPI MAPILogoff(
	_In_ LHANDLE   lhSession,
	_In_ ULONG_PTR ulUIParam,
	     FLAGS     flFlags,
	     ULONG     ulReserved
)
{
	buffer_t buffer;

	{
		boost::iostreams::back_insert_device<buffer_t> inserter(buffer);
		boost::iostreams::stream<boost::iostreams::back_insert_device<buffer_t> > s(inserter);
		boost::archive::text_oarchive ar(s);

		ar << std::string(__func__);
		ar << lhSession;
		ar << ulUIParam;
		ar << flFlags;
		ar << ulReserved;
		s.flush();
		s.close();
	}

	{
		DWORD rc = CallFunctionOverRDP(buffer);
		if (rc != SUCCESS_SUCCESS)
		{
			MessageBox(NULL, L"MAPIoRDP", L"Could not call MAPI function over RDP.", MB_OK | MB_ICONERROR);
			return rc;
		}
	}

	{
		boost::iostreams::basic_array_source<char> device(buffer.data(), buffer.size());
		boost::iostreams::stream<boost::iostreams::basic_array_source<char> > s(device);
		boost::archive::text_iarchive ar(s);

		ULONG rc;
		ar >> rc;

		return rc;
	}
}

ULONG WINAPI MAPILogon(
	_In_     ULONG_PTR ulUIParam,
	_In_opt_ LPSTR     lpszProfileName,
	_In_opt_ LPSTR     lpszPassword,
	_In_     FLAGS     flFlags,
	         ULONG     ulReserved,
	_Out_    LPLHANDLE lplhSession
)
{
	buffer_t buffer;

	{
		boost::iostreams::back_insert_device<buffer_t> inserter(buffer);
		boost::iostreams::stream<boost::iostreams::back_insert_device<buffer_t> > s(inserter);
		boost::archive::text_oarchive ar(s);

		ar << std::string(__func__);
		ar << ulUIParam;
		save_c_str(ar, lpszProfileName);
		save_c_str(ar, lpszPassword);
		ar << flFlags;
		ar << ulReserved;
		s.flush();
		s.close();
	}

	{
		DWORD rc = CallFunctionOverRDP(buffer);
		if (rc != SUCCESS_SUCCESS)
		{
			MessageBox(NULL, L"MAPIoRDP", L"Could not call MAPI function over RDP.", MB_OK | MB_ICONERROR);
			return rc;
		}
	}

	{
		boost::iostreams::basic_array_source<char> device(buffer.data(), buffer.size());
		boost::iostreams::stream<boost::iostreams::basic_array_source<char> > s(device);
		boost::archive::text_iarchive ar(s);

		ULONG rc;
		ar >> rc;
		ar >> *lplhSession;

		return rc;
	}
}

ULONG WINAPI MAPIReadMail(
	_In_  LHANDLE       lhSession,
	_In_  ULONG_PTR     ulUIParam,
	_In_  LPSTR         lpszMessageID,
	_In_  FLAGS         flFlags,
	ULONG         ulReserved,
	_Out_ lpMapiMessage *lppMessage
)
{
	buffer_t buffer;

	{
		boost::iostreams::back_insert_device<buffer_t> inserter(buffer);
		boost::iostreams::stream<boost::iostreams::back_insert_device<buffer_t> > s(inserter);
		boost::archive::text_oarchive ar(s);

		ar << std::string(__func__);
		ar << lhSession;
		ar << ulUIParam;
		save_c_str(ar, lpszMessageID);
		ar << flFlags;
		ar << ulReserved;
		s.flush();
		s.close();
	}

	{
		DWORD rc = CallFunctionOverRDP(buffer);
		if (rc != SUCCESS_SUCCESS)
		{
			MessageBox(NULL, L"MAPIoRDP", L"Could not call MAPI function over RDP.", MB_OK | MB_ICONERROR);
			return rc;
		}
	}

	{
		boost::iostreams::basic_array_source<char> device(buffer.data(), buffer.size());
		boost::iostreams::stream<boost::iostreams::basic_array_source<char> > s(device);
		boost::archive::text_iarchive ar(s);

		ULONG rc;
		ar >> rc;
		*lppMessage = new MapiMessage();
		SetPointerArray(*lppMessage);
		ar >> **lppMessage;

		return rc;
	}
}
*/

template<class ArchiveIn, class ArchiveOut>
void cMAPIResolveName(ArchiveIn & ai, ArchiveOut & ao)
{
	LHANDLE       lhSession;
	ULONG_PTR     ulUIParam;
	LPSTR         lpszName;
	FLAGS         flFlags;
	ULONG         ulReserved;
	lpMapiRecipDesc *lppRecip;

	ai >> lhSession;
	ai >> ulUIParam;
	load_c_str(ai, lpszName);
	ai >> flFlags;
	ai >> ulReserved;

	LPMAPIRESOLVENAME lpMAPIResolveName;
	lpMAPIResolveName = (LPMAPIRESOLVENAME)GetProcAddress(mapi_hLib, "MAPIResolveName");

	LPMAPIFREEBUFFER lpMAPIFreeBuffer;
	lpMAPIFreeBuffer = (LPMAPIFREEBUFFER)GetProcAddress(mapi_hLib, "MAPIFreeBuffer");

	ULONG rc;
	if (!lpMAPIResolveName || !lpMAPIFreeBuffer)
	{
		rc = MAPI_E_FAILURE;
		ao << rc;
	}
	else
	{
		rc = lpMAPIResolveName(lhSession, ulUIParam, lpszName, flFlags, ulReserved, lppRecip);
		ao << rc;
		if (*lppRecip != NULL)
		{
			ao << **lppRecip;
			lpMAPIFreeBuffer(*lppRecip);
		}
		free(lpszName);
	}
}

template<class ArchiveIn, class ArchiveOut>
void cMAPISaveMail(ArchiveIn & ai, ArchiveOut & ao)
{
	LHANDLE       lhSession;
	ULONG_PTR     ulUIParam;
	lpMapiMessage lpMessage;
	FLAGS         flFlags;
	ULONG         ulReserved;
	LPSTR         lpszMessageID;

	ai >> lhSession;
	ai >> ulUIParam;
	lpMessage = (lpMapiMessage)malloc(sizeof(MapiMessage));
	ai >> *lpMessage;
	ai >> flFlags;
	ai >> ulReserved;
	load_c_str(ai, lpszMessageID);

	LPMAPISAVEMAIL lpMAPISaveMail;
	lpMAPISaveMail = (LPMAPISAVEMAIL)GetProcAddress(mapi_hLib, "MAPISaveMail");

	ULONG rc;
	if (!lpMAPISaveMail)
	{
		rc = MAPI_E_FAILURE;
		ao << rc;
	}
	else
	{
		rc = lpMAPISaveMail(lhSession, ulUIParam, lpMessage, flFlags, ulReserved, lpszMessageID);
		ao << rc;
		free(lpszMessageID);
		cleanup_files(*lpMessage);
		FreeMAPIMessage(lpMessage);
	}
}

// TODO deserialize file content
template<class ArchiveIn, class ArchiveOut>
void cMAPISendDocuments(ArchiveIn & ai, ArchiveOut & ao)
{
	ULONG_PTR       ulUIParam;
	CHAR            delimChar;
	LPSTR           lpszFilePaths;
	LPSTR           lpszFileNames;
	ULONG           ulReserved;

	ai >> ulUIParam;
	ai >> delimChar;
	load_c_str(ai, lpszFilePaths);
	load_c_str(ai, lpszFileNames);
	ai >> ulReserved;

	LPMAPISENDDOCUMENTS lpMAPISendDocuments;
	lpMAPISendDocuments = (LPMAPISENDDOCUMENTS)GetProcAddress(mapi_hLib, "MAPISendDocuments");

	ULONG rc;
	if (!lpMAPISendDocuments)
	{
		rc = MAPI_E_FAILURE;
		ao << rc;
	}
	else
	{
		rc = lpMAPISendDocuments(ulUIParam, &delimChar, lpszFilePaths, lpszFileNames, ulReserved);
		ao << rc;
		free(lpszFilePaths);
		free(lpszFileNames);
		// TODO FILE CLEANUP
		//cleanup_files(Message);
	}
}

template<class ArchiveIn, class ArchiveOut>
void cMAPISendMail(ArchiveIn & ai, ArchiveOut & ao)
{
	LHANDLE         lhSession;
	ULONG_PTR       ulUIParam;
	lpMapiMessage   lpMessage;
	FLAGS           flFlags;
	ULONG           ulReserved;

	ai >> lhSession;
	ai >> ulUIParam;
	lpMessage = (lpMapiMessage)malloc(sizeof(MapiMessage));
	ai >> *lpMessage;
	ai >> flFlags;
	ai >> ulReserved;

	LPMAPISENDMAIL lpMAPISendMail;
	lpMAPISendMail = (LPMAPISENDMAIL)GetProcAddress(mapi_hLib, "MAPISendMail");

	ULONG rc;
	if (!lpMAPISendMail)
	{
		rc = MAPI_E_FAILURE;
		ao << rc;
	}
	else
	{
		rc = lpMAPISendMail(lhSession, ulUIParam, lpMessage, flFlags, ulReserved);
		ao << rc;
		cleanup_files(*lpMessage);
		FreeMAPIMessage(lpMessage);
	}
}

template<class ArchiveIn, class ArchiveOut>
void cMAPISendMailHelper(ArchiveIn & ai, ArchiveOut & ao)
{
	LHANDLE         lhSession;
	ULONG_PTR       ulUIParam;
	lpMapiMessage   lpMessage;
	FLAGS           flFlags;
	ULONG           ulReserved;

	ai >> lhSession;
	ai >> ulUIParam;
	lpMessage = (lpMapiMessage)malloc(sizeof(MapiMessage));
	ai >> *lpMessage;
	ai >> flFlags;
	ai >> ulReserved;

	LPMAPISENDMAIL lpMAPISendMailHelper;
	lpMAPISendMailHelper = (LPMAPISENDMAIL)GetProcAddress(mapi_hLib, "MAPISendMailHelper");

	ULONG rc;
	if (!lpMAPISendMailHelper)
	{
		rc = MAPI_E_FAILURE;
		ao << rc;
	}
	else
	{
		rc = lpMAPISendMailHelper(lhSession, ulUIParam, lpMessage, flFlags, ulReserved);
		ao << rc;
		cleanup_files(*lpMessage);
		FreeMAPIMessage(lpMessage);
	}
}

template<class ArchiveIn, class ArchiveOut>
void cMAPISendMailW(ArchiveIn & ai, ArchiveOut & ao)
{
	LHANDLE         lhSession;
	ULONG_PTR       ulUIParam;
	MapiMessageW    Message;
	FLAGS           flFlags;
	ULONG           ulReserved;

	ai >> lhSession;
	ai >> ulUIParam;
	ai >> Message;
	ai >> flFlags;
	ai >> ulReserved;

	LPMAPISENDMAILW lpMAPISendMailW;
	lpMAPISendMailW = (LPMAPISENDMAILW)GetProcAddress(mapi_hLib, "MAPISendMailW");

	ULONG rc;
	if (!lpMAPISendMailW)
	{
		rc = MAPI_E_FAILURE;
		ao << rc;
	}
	else
	{
		rc = lpMAPISendMailW(lhSession, ulUIParam, &Message, flFlags, ulReserved);
		ao << rc;
		cleanup_files(Message);
		FreeMAPIMessageW(&Message);
	}
}
