#include "stdafx.h"
#include <stdlib.h>
#include <malloc.h>
#include "Tools.h"

template<class lpMapiFileDescT>
void FreeMAPIFileT(lpMapiFileDescT pv)
{
	if (!pv)
		return;

	if (pv->lpszPathName != NULL)
		free(pv->lpszPathName);

	if (pv->lpszFileName != NULL)
		free(pv->lpszFileName);
}

template<class lpMapiMessageT>
void FreeMAPIMessageT(lpMapiMessageT pv)
{
	ULONG i;

	if (!pv)
		return;

	if (pv->lpszSubject != NULL)
		free(pv->lpszSubject);

	if (pv->lpszNoteText)
		free(pv->lpszNoteText);

	if (pv->lpszMessageType)
		free(pv->lpszMessageType);

	if (pv->lpszDateReceived)
		free(pv->lpszDateReceived);

	if (pv->lpszConversationID)
		free(pv->lpszConversationID);

	if (pv->lpOriginator)
		FreeMAPIRecipientT(pv->lpOriginator);

	for (i = 0; i<pv->nRecipCount; i++)
	{
		if (&(pv->lpRecips[i]) != NULL)
		{
			FreeMAPIRecipientT(&(pv->lpRecips[i]));
		}
	}

	if (pv->lpRecips != NULL)
	{
		free(pv->lpRecips);
	}

	for (i = 0; i<pv->nFileCount; i++)
	{
		if (&(pv->lpFiles[i]) != NULL)
		{
			FreeMAPIFileT(&(pv->lpFiles[i]));
		}
	}

	if (pv->lpFiles != NULL)
	{
		free(pv->lpFiles);
	}

	free(pv);
	pv = NULL;
}

void FreeMAPIMessage(lpMapiMessage pv)
{
	FreeMAPIMessageT(pv);
}

void FreeMAPIMessageW(lpMapiMessageW pv)
{
	FreeMAPIMessageT(pv);
}

void FreeMAPIRecipientArray(lpMapiRecipDesc *pv)
{
	if (!pv)
		return;

	free(pv);
}

template<class lpMapiRecipDescT>
void FreeMAPIRecipientT(lpMapiRecipDescT pv)
{
	if (!pv)
		return;

	if (pv->lpszName != NULL)
		free(pv->lpszName);

	if (pv->lpszAddress != NULL)
		free(pv->lpszAddress);

	if (pv->lpEntryID != NULL)
		free(pv->lpEntryID);
}

void FreeMAPIRecipient(lpMapiRecipDesc pv)
{
	FreeMAPIRecipientT(pv);
}

void FreeMAPIRecipientW(lpMapiRecipDescW pv)
{
	FreeMAPIRecipientT(pv);
}
