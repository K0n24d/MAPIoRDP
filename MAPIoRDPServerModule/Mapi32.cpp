// Mapi32.cpp : définit les fonctions SimpleMAPI exportées.
//

#include "stdafx.h"
#include <MAPI.h>
#include <boost\iostreams\device\back_inserter.hpp>
#include <boost\iostreams\stream.hpp>
#include <boost\archive\text_oarchive.hpp>
#include <boost\archive\text_iarchive.hpp>
#include "Tools.h"

#define _MAX_WAIT       60000
#define MAX_MSG_SIZE    0x20000

#include "DVCServerModule.h"
#include "MAPISerializer.hpp"

#define   MAX_POINTERS         32
#define   MAPI_MESSAGE_TYPE     0
#define   MAPI_RECIPIENT_TYPE   1
#define   MAPI_RECIPIENT_ARRAY  2

typedef struct {
	LPVOID    lpMem;
	UCHAR     memType;
} memTrackerType;

// this can't be right.
memTrackerType    memArray[MAX_POINTERS];

//
// For remembering memory...how ironic.
//
void SetPointerArray(LPVOID ptr, BYTE type)
{
	int i;

	for (i = 0; i<MAX_POINTERS; i++)
	{
		if (memArray[i].lpMem == NULL)
		{
			memArray[i].lpMem = ptr;
			memArray[i].memType = type;
			break;
		}
	}
}
void SetPointerArray(lpMapiMessage ptr)
{
	SetPointerArray(ptr, MAPI_MESSAGE_TYPE);
}
void SetPointerArray(lpMapiRecipDesc ptr)
{
	SetPointerArray(ptr, MAPI_RECIPIENT_TYPE);
}
void SetPointerArray(lpMapiRecipDesc *ptr)
{
	SetPointerArray(ptr, MAPI_RECIPIENT_ARRAY);
}

using namespace boost::serialization;

ULONG WINAPI MAPIAddress(
	_In_  LHANDLE         lhSession,
	_In_  ULONG_PTR       ulUIParam,
	_In_  LPSTR           lpszCaption,
	_In_  ULONG           nEditFields,
	_In_  LPSTR           lpszLabels,
	_In_  ULONG           nRecips,
	_In_  lpMapiRecipDesc lpRecips,
	_In_  FLAGS           flFlags,
	      ULONG           ulReserved,
	_Out_ LPULONG         lpnNewRecips,
	_Out_ lpMapiRecipDesc *lppNewRecips
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
		save_c_str(ar, lpszCaption);
		ar << nEditFields;
		save_c_str(ar, lpszLabels);
		ar << nRecips;
		for (ULONG i = 0; i < nRecips; i++)
			ar << lpRecips[i];
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
		if (rc != SUCCESS_SUCCESS)
			return rc;

		ar >> *lpnNewRecips;
		lppNewRecips = (lpMapiRecipDesc *)calloc(nRecips, sizeof(lpMapiRecipDesc));
		SetPointerArray(lppNewRecips);
		for (ULONG i = 0; i < nRecips; i++)
		{
			lppNewRecips[i] = (lpMapiRecipDesc)malloc(sizeof(MapiRecipDesc));
			SetPointerArray(lppNewRecips[i]);
			ar >> *(lppNewRecips[i]);
		}

		return rc;
	}
}

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
		if (rc != SUCCESS_SUCCESS)
			return rc;

		load_c_str(ar, lpszMessageID);

		return rc;
	}
}

ULONG WINAPI MAPIFreeBuffer(
	_In_ LPVOID pv
)
{
	// _In_ LPVOID pv
	// Pointer to memory allocated by the messaging system. 
	// This pointer is returned by the MAPIReadMail, MAPIAddress,
	// and MAPIResolveName functions.

	int   i;

	if (!pv)
		return SUCCESS_SUCCESS;

	for (i = 0; i<MAX_POINTERS; i++)
	{
		if (pv == memArray[i].lpMem)
		{
			if (memArray[i].memType == MAPI_MESSAGE_TYPE)
			{
				FreeMAPIMessage((lpMapiMessage)pv);
				memArray[i].lpMem = NULL;
			}
			else if (memArray[i].memType == MAPI_RECIPIENT_TYPE)
			{
				FreeMAPIRecipient((lpMapiRecipDesc)pv);
				memArray[i].lpMem = NULL;
			}
			else if (memArray[i].memType == MAPI_RECIPIENT_ARRAY)
			{
				FreeMAPIRecipientArray((lpMapiRecipDesc*)pv);
				memArray[i].lpMem = NULL;
			}
		}
	}

	pv = NULL;
	return SUCCESS_SUCCESS;
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
		if (rc != SUCCESS_SUCCESS)
			return rc;

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
		if (rc != SUCCESS_SUCCESS)
			return rc;

		*lppMessage = (lpMapiMessage) malloc(sizeof(MapiMessage));
		SetPointerArray(*lppMessage);
		ar >> **lppMessage;

		return rc;
	}
}

ULONG WINAPI MAPIResolveName(
	_In_  LHANDLE         lhSession,
	_In_  ULONG_PTR       ulUIParam,
	_In_  LPSTR           lpszName,
	_In_  FLAGS           flFlags,
	      ULONG           ulReserved,
	_Out_ lpMapiRecipDesc *lppRecip
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
		save_c_str(ar, lpszName);
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
		if (rc != SUCCESS_SUCCESS)
			return rc;

		*lppRecip = (lpMapiRecipDesc)malloc(sizeof(MapiRecipDesc));
		SetPointerArray(*lppRecip);
		ar >> **lppRecip;

		return rc;
	}
}

ULONG WINAPI MAPISaveMail(
	_In_ LHANDLE       lhSession,
	_In_ ULONG_PTR     ulUIParam,
	_In_ lpMapiMessage lpMessage,
	_In_ FLAGS         flFlags,
	     ULONG         ulReserved,
	_In_ LPSTR         lpszMessageID
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
		ar << *lpMessage;
		ar << flFlags;
		ar << ulReserved;
		save_c_str(ar, lpszMessageID);
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

ULONG WINAPI MAPISendDocuments(
	_In_ ULONG_PTR ulUIParam,
	_In_ LPSTR     lpszDelimChar,
	_In_ LPSTR     lpszFilePaths,
	_In_ LPSTR     lpszFileNames,
	     ULONG     ulReserved
)
{
	buffer_t buffer;

	{
		boost::iostreams::back_insert_device<buffer_t> inserter(buffer);
		boost::iostreams::stream<boost::iostreams::back_insert_device<buffer_t> > s(inserter);
		boost::archive::text_oarchive ar(s);

		ar << std::string(__func__);
		ar << ulUIParam;
		ar << *lpszDelimChar;
		save_c_str(ar, lpszFilePaths);
		save_c_str(ar, lpszFileNames);
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

template<class lpMapiMessageT>
ULONG WINAPI MAPISendMailT(
	_In_ LPCSTR         funcdname,
	_In_ LHANDLE        lhSession,
	_In_ ULONG_PTR      ulUIParam,
	_In_ lpMapiMessageT lpMessage,
	_In_ FLAGS          flFlags,
	_In_ ULONG          ulReserved
)
{
	buffer_t buffer;

	{
		boost::iostreams::back_insert_device<buffer_t> inserter(buffer);
		boost::iostreams::stream<boost::iostreams::back_insert_device<buffer_t> > s(inserter);
		boost::archive::text_oarchive ar(s);

		ar << std::string(funcdname);
		ar << lhSession;
		ar << ulUIParam;
		ar << *lpMessage;
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

ULONG WINAPI MAPISendMail(
	_In_ LHANDLE        lhSession,
	_In_ ULONG_PTR      ulUIParam,
	_In_ lpMapiMessage  lpMessage,
	_In_ FLAGS          flFlags,
	_In_ ULONG          ulReserved
)
{
	return MAPISendMailT(__func__, lhSession, ulUIParam, lpMessage, flFlags, ulReserved);
}

ULONG WINAPI MAPISendMailHelper(
	_In_ LHANDLE        lhSession,
	_In_ ULONG_PTR      ulUIParam,
	_In_ lpMapiMessageW lpMessage,
	_In_ FLAGS          flFlags,
	_In_ ULONG          ulReserved
)
{
	return MAPISendMailT(__func__, lhSession, ulUIParam, lpMessage, flFlags, ulReserved);
}

ULONG WINAPI MAPISendMailW(
	_In_ LHANDLE        lhSession,
	_In_ ULONG_PTR      ulUIParam,
	_In_ lpMapiMessageW lpMessage,
	_In_ FLAGS          flFlags,
	_In_ ULONG          ulReserved
)
{
	return MAPISendMailT(__func__, lhSession, ulUIParam, lpMessage, flFlags, ulReserved);
}
