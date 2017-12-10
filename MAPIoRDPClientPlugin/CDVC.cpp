/*
*  Sample "echo" plugin that listens on endpoint "DVC_Sample".
*
*/
#include "stdafx.h"
#include "resource.h"
#include "MAPIoRDPClientPlugin_i.h"
#include <tsvirtualchannels.h>

#include <boost\iostreams\device\back_inserter.hpp>
#include <boost\iostreams\stream.hpp>
#include <boost\archive\text_oarchive.hpp>
#include <boost\archive\text_iarchive.hpp>
#include "DVCServerModule.h"
#include "MAPISerializer.hpp"
#include "Mapi32.hpp"
using namespace boost::serialization;

using namespace ATL;

#define CHECK_QUIT_HR( _x_ )    if(FAILED(hr)) { return hr; }

// CDVCSamplePlugin implements IWTSPlugin.
class ATL_NO_VTABLE CMAPIoRDPClientPlugin :
	public CComObjectRootEx<CComMultiThreadModel>,
	public CComCoClass<CMAPIoRDPClientPlugin, &CLSID_MAPIoRDPClientPlugin>,
	public IWTSPlugin
{
public:

	DECLARE_REGISTRY_RESOURCEID(IDR_MAPIORDPCLIENTPLUGIN)

	BEGIN_COM_MAP(CMAPIoRDPClientPlugin)
		COM_INTERFACE_ENTRY(IWTSPlugin)
	END_COM_MAP()

	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}

	void FinalRelease()
	{
	}

	// IWTSPlugin.
	//
	HRESULT STDMETHODCALLTYPE
		Initialize(IWTSVirtualChannelManager *pChannelMgr);

	HRESULT STDMETHODCALLTYPE
		Connected()
	{
		return S_OK;
	}

	HRESULT STDMETHODCALLTYPE
		Disconnected(DWORD dwDisconnectCode)
	{
		// Prevent C4100 "unreferenced parameter" warnings.
		dwDisconnectCode;

		return S_OK;
	}

	HRESULT STDMETHODCALLTYPE
		Terminated()
	{
		return S_OK;
	}

};

OBJECT_ENTRY_AUTO(__uuidof(MAPIoRDPClientPlugin), CMAPIoRDPClientPlugin)

// CSampleChannelCallback implements IWTSVirtualChannelCallback.
class ATL_NO_VTABLE CMAPIoRDPChannelCallback :
	public CComObjectRootEx<CComMultiThreadModel>,
	public IWTSVirtualChannelCallback
{
	CComPtr<IWTSVirtualChannel> m_ptrChannel;

public:

	BEGIN_COM_MAP(CMAPIoRDPChannelCallback)
		COM_INTERFACE_ENTRY(IWTSVirtualChannelCallback)
	END_COM_MAP()

	VOID SetChannel(IWTSVirtualChannel *pChannel)
	{
		m_ptrChannel = pChannel;
	}

	// IWTSVirtualChannelCallback
	//
/*
	HRESULT STDMETHODCALLTYPE
		OnDataReceived(
			ULONG cbSize,
			__in_bcount(cbSize) BYTE *pBuffer
		)
*/
	HRESULT STDMETHODCALLTYPE
		OnDataReceived(
			ULONG cbSize,
			BYTE  *pBuffer
		)
	{
		buffer_t buffer;
		{
			ULONG rc;
			boost::iostreams::back_insert_device<buffer_t> inserter(buffer);
			boost::iostreams::stream<boost::iostreams::back_insert_device<buffer_t> > so(inserter);
			boost::archive::text_oarchive ao(so);

			boost::iostreams::basic_array_source<char> device((char *)pBuffer, cbSize);
			boost::iostreams::stream<boost::iostreams::basic_array_source<char> > si(device);
			boost::archive::text_iarchive ai(si);

			std::string func;
			ai >> func;

			if (func == "MAILTO")
			{
				PWSTR url;
				load_c_str(ai, url);

				if (wcslen(url) > 0)
				{
					int retValue = (int)ShellExecute(NULL, L"open", url, NULL, NULL, SW_SHOWNORMAL);
					if (retValue > 32)
						rc = S_OK;
					else
						rc = retValue;
				}
				else
					rc = TYPE_E_DLLFUNCTIONNOTFOUND;

				delete url;
				ao << rc;
			}
			else if (func == "MAPIAddress")
				cMAPIAddress(ai, ao);
			else if (func == "MAPIDeleteMail");
			else if (func == "MAPIDetails");
			else if (func == "MAPIFindNext");
			else if (func == "MAPIFreeBuffer");
			else if (func == "MAPILogoff");
			else if (func == "MAPILogon");
			else if (func == "MAPIReadMail");
			else if (func == "MAPIResolveName");
			else if (func == "MAPISaveMail");
			else if (func == "MAPISendDocuments")
				cMAPISendDocuments(ai, ao);
			else if (func == "MAPISendMail")
				cMAPISendMail(ai, ao);
			else if (func == "MAPISendMailHelper")
				cMAPISendMailHelper(ai, ao);
			else if (func == "MAPISendMailW")
				cMAPISendMailW(ai, ao);

			else
				rc = TYPE_E_DLLFUNCTIONNOTFOUND;

			so.flush();
			so.close();
		}

		return m_ptrChannel->Write(buffer.size(), (BYTE *) buffer.data(), NULL);
	}

	HRESULT STDMETHODCALLTYPE
		OnClose()
	{
		return m_ptrChannel->Close();
	}
};


// CSampleListenerCallback implements IWTSListenerCallback.
class ATL_NO_VTABLE CMAPIoRDPListenerCallback :
	public CComObjectRootEx<CComMultiThreadModel>,
	public IWTSListenerCallback
{
public:

	BEGIN_COM_MAP(CMAPIoRDPListenerCallback)
		COM_INTERFACE_ENTRY(IWTSListenerCallback)
	END_COM_MAP()

	HRESULT STDMETHODCALLTYPE
		OnNewChannelConnection(
			__in IWTSVirtualChannel *pChannel,
			__in_opt BSTR data,
			__out BOOL *pbAccept,
			__out IWTSVirtualChannelCallback **ppCallback);
};

// IWTSPlugin::Initialize implementation.
HRESULT
CMAPIoRDPClientPlugin::Initialize(
	__in IWTSVirtualChannelManager *pChannelMgr
)
{
	HRESULT hr;
	CComObject<CMAPIoRDPListenerCallback> *pListenerCallback;
	CComPtr<CMAPIoRDPListenerCallback> ptrListenerCallback;
	CComPtr<IWTSListener> ptrListener;

	// Create an instance of the CSampleListenerCallback object.
	hr = CComObject<CMAPIoRDPListenerCallback>::CreateInstance(&pListenerCallback);
	CHECK_QUIT_HR("CMAPIoRDPListenerCallback::CreateInstance");
	ptrListenerCallback = pListenerCallback;

	// Attach the callback to the "DVC_Sample" endpoint.
	hr = pChannelMgr->CreateListener(
		"MAPIoRDP",
		0,
		(CMAPIoRDPListenerCallback*)ptrListenerCallback,
		&ptrListener);
	CHECK_QUIT_HR("CreateListener");

	return hr;
}

// IWTSListenerCallback::OnNewChannelConnection implementation.
HRESULT
CMAPIoRDPListenerCallback::OnNewChannelConnection(
	__in IWTSVirtualChannel *pChannel,
	__in_opt BSTR data,
	__out BOOL *pbAccept,
	__out IWTSVirtualChannelCallback **ppCallback)
{
	HRESULT hr;
	CComObject<CMAPIoRDPChannelCallback> *pCallback;
	CComPtr<CMAPIoRDPChannelCallback> ptrCallback;

	// Prevent C4100 "unreferenced parameter" warnings.
	data;

	*pbAccept = FALSE;

	hr = CComObject<CMAPIoRDPChannelCallback>::CreateInstance(&pCallback);
	CHECK_QUIT_HR("CMAPIoRDPChannelCallback::CreateInstance");
	ptrCallback = pCallback;

	ptrCallback->SetChannel(pChannel);

	*ppCallback = ptrCallback;
	(*ppCallback)->AddRef();

	*pbAccept = TRUE;

	return hr;
}

#include <cchannel.h>

HRESULT VCAPITYPE VirtualChannelGetInstance(
	_In_    REFIID refiid,
	_Inout_ ULONG  *pNumObjs,
	_Out_   VOID   **ppObjArray
)
{
	if (refiid != IID_IWTSPlugin)
		return E_NOINTERFACE;

	if (*pNumObjs != 1)
	{
		*pNumObjs = 1;
		return S_OK;
	}

	if (ppObjArray == NULL)
		return E_POINTER;

	CComObject<CMAPIoRDPClientPlugin>* pObj;
	HRESULT hr = CComObject<CMAPIoRDPClientPlugin>::CreateInstance(&pObj);
	if (FAILED(hr))
		return hr;

	ppObjArray[0] = pObj;
	pObj->AddRef();
	return S_OK;
}
