#include <windows.h>
#include <wtsapi32.h>
#include <pchannel.h>
#include <crtdbg.h>

#include <MAPI.h>

#pragma comment(lib, "wtsapi32.lib")

#define _MAX_WAIT       60000
#define MAX_MSG_SIZE    0x20000
#define START_MSG_SIZE  4
#define STEP_MSG_SIZE   113

#include "DVCServerModule.h"

/*
*  Open a dynamic channel with the name given in szChannelName.
*  The output file handle can be used in ReadFile/WriteFile calls.
*/
DWORD OpenDynamicChannel(
	HANDLE *phFile)
{
	LPCSTR szChannelName = "MAPIoRDP";
	*phFile = NULL;
	
	HANDLE hWTSHandle = NULL;
	HANDLE hWTSFileHandle;
	PVOID vcFileHandlePtr = NULL;
	DWORD len;
	DWORD rc = ERROR_SUCCESS;

	hWTSHandle = WTSVirtualChannelOpenEx(
		WTS_CURRENT_SESSION,
		(LPSTR)szChannelName,
		WTS_CHANNEL_OPTION_DYNAMIC);
	if (NULL == hWTSHandle)
	{
		rc = GetLastError();
		goto exitpt;
	}

	BOOL fSucc = WTSVirtualChannelQuery(
		hWTSHandle,
		WTSVirtualFileHandle,
		&vcFileHandlePtr,
		&len);
	if (!fSucc)
	{
		rc = GetLastError();
		goto exitpt;
	}
	if (len != sizeof(HANDLE))
	{
		rc = ERROR_INVALID_PARAMETER;
		goto exitpt;
	}

	hWTSFileHandle = *(HANDLE *)vcFileHandlePtr;

	fSucc = DuplicateHandle(
		GetCurrentProcess(),
		hWTSFileHandle,
		GetCurrentProcess(),
		phFile,
		0,
		FALSE,
		DUPLICATE_SAME_ACCESS);
	if (!fSucc)
	{
		rc = GetLastError();
		goto exitpt;
	}

	rc = ERROR_SUCCESS;

exitpt:
	if (vcFileHandlePtr)
	{
		WTSFreeMemory(vcFileHandlePtr);
	}
	if (hWTSHandle)
	{
		WTSVirtualChannelClose(hWTSHandle);
	}

	return rc;
}

DWORD CallFunctionOverRDP(buffer_t & buffer)
{
	DWORD rc;
	HANDLE hFile;

	rc = OpenDynamicChannel(&hFile);
	if (ERROR_SUCCESS != rc)
	{
		if (hFile != NULL)
			CloseHandle(hFile);
		return MAPI_E_FAILURE;
	}

	HANDLE  hEvent;
	hEvent = CreateEvent(NULL, FALSE, FALSE, NULL);

	{
		size_t msgSize = buffer.size();

		OVERLAPPED  Overlapped = { 0 };
		Overlapped.hEvent = hEvent;

		BOOL    fSucc;
		DWORD   dwWritten;
		fSucc = WriteFile(
			hFile,
			buffer.data(),
			msgSize,
			&dwWritten,
			&Overlapped);
		if (!fSucc)
		{
			if (GetLastError() == ERROR_IO_PENDING)
			{
				DWORD dw = WaitForSingleObject(Overlapped.hEvent, _MAX_WAIT);
				_ASSERT(WAIT_OBJECT_0 == dw);
				fSucc = GetOverlappedResult(
					hFile,
					&Overlapped,
					&dwWritten,
					FALSE);
			}
		}

		if (!fSucc)
		{
			DWORD error = GetLastError();
			CloseHandle(hFile);
			return error;
		}

		if (dwWritten != msgSize)
		{
			CloseHandle(hFile);
			return MAPI_E_FAILURE;
		}
	}

	// Message has been send, wait and get reply
	{
		BOOL fSucc;

		BYTE ReadBuffer[CHANNEL_PDU_LENGTH];
		buffer.clear();
		CHANNEL_PDU_HEADER *pHdr;
		pHdr = (CHANNEL_PDU_HEADER*)ReadBuffer;

		DWORD dwRead;
		OVERLAPPED  Overlapped = { 0 };
		DWORD TotalRead = 0;

		do {
			Overlapped.hEvent = hEvent;

			// Read the entire message.
			fSucc = ReadFile(
				hFile,
				ReadBuffer,
				sizeof(ReadBuffer),
				&dwRead,
				&Overlapped);
			if (!fSucc)
			{
				if (GetLastError() == ERROR_IO_PENDING)
				{
					DWORD dw = WaitForSingleObject(Overlapped.hEvent, INFINITE);
					_ASSERT(WAIT_OBJECT_0 == dw);
					fSucc = GetOverlappedResult(
						hFile,
						&Overlapped,
						&dwRead,
						FALSE);
				}
			}

			if (!fSucc)
			{
				DWORD error = GetLastError();
				CloseHandle(hFile);
				return error;
			}

			ULONG packetSize = dwRead - sizeof(*pHdr);
			TotalRead += packetSize;
			PBYTE pData = (PBYTE)(pHdr + 1);

			buffer.reserve(buffer.size() + packetSize);
			buffer.insert(buffer.end(), pData, pData + packetSize);

			_ASSERT(packetSize == pHdr->length);

		} while (0 == (pHdr->flags & CHANNEL_FLAG_LAST));

	}

	CloseHandle(hFile);
	return SUCCESS_SUCCESS;
}
