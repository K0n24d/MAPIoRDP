#pragma once

#include <MAPI.h>

void FreeMAPIRecipientArray(lpMapiRecipDesc *pv);
void FreeMAPIRecipient(lpMapiRecipDesc pv);
void FreeMAPIRecipientW(lpMapiRecipDescW pv);
void FreeMAPIMessage(lpMapiMessage pv);
void FreeMAPIMessageW(lpMapiMessageW pv);
