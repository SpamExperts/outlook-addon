// Connect.h : Declaration of the CConnect

#pragma once
#include "resource.h"       // main symbols

// func info
static _ATL_FUNC_INFO OnClickButtonInfo =
{
	CC_STDCALL,
	VT_EMPTY,
	2,
	{VT_DISPATCH, VT_BYREF | VT_BOOL}
};


// CConnect
class ATL_NO_VTABLE CConnect : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CConnect, &CLSID_Connect>,
	public IDispatchImpl<AddInDesignerObjects::_IDTExtensibility2, &AddInDesignerObjects::IID__IDTExtensibility2, &AddInDesignerObjects::LIBID_AddInDesignerObjects, 1, 0>,
	public IDispEventSimpleImpl<1, CConnect, &__uuidof(Office::_CommandBarButtonEvents)>
{
	typedef IDispEventSimpleImpl<1, CConnect, &__uuidof(Office::_CommandBarButtonEvents)> _SpamButtonEvents;
public:
	CConnect()
	{
	}

DECLARE_REGISTRY_RESOURCEID(IDR_ADDIN)
DECLARE_NOT_AGGREGATABLE(CConnect)

BEGIN_COM_MAP(CConnect)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(AddInDesignerObjects::IDTExtensibility2)
END_COM_MAP()

BEGIN_SINK_MAP(CConnect)
	SINK_ENTRY_INFO(1, __uuidof(Office::_CommandBarButtonEvents), 0x01, OnSpamButtonClick, &OnClickButtonInfo)
END_SINK_MAP()

	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}
	
	void FinalRelease() 
	{
	}

public:
	// Outlook notification handlers
	void __stdcall OnSpamButtonClick(IDispatch * /*Office::_CommandBarButton**/ Ctrl, VARIANT_BOOL * CancelDefault);

private:
	HRESULT AddCommandBar(void);
	CComPtr<Office::_CommandBarButton> AddButton(CComPtr<Office::CommandBarControls> spBarControls,
												_bstr_t bstrCaption, _bstr_t bstrTipText, _bstr_t bstrTag,
												ULONG_PTR bitmapId, bool bEnabled);

	bool PostData(LPBYTE pHdrData, DWORD dwHdrLen, LPBYTE pMsgData, DWORD dwMsgLen);
	DWORD WriteToFile(LPCTSTR lpszFile, LPBYTE pHdrData, DWORD dwHdrLen, LPBYTE pMsgData, DWORD dwMsgLen);
	bool IsResponseOK(LPCTSTR lpszFile) const;
	CStringA GetInternetHeaders(IMessage* pMsg);
	void RemoveLine(CStringA& strData, const CStringA& strLine);
	void RemoveDoubleBreak(CStringA& strData);

private:
	CString m_strDLLPath;
	CComPtr<Office::_CommandBarButton> m_spSpamButton;

public:
	//IDTExtensibility2 implementation:
	STDMETHOD(OnConnection)(IDispatch * Application, AddInDesignerObjects::ext_ConnectMode ConnectMode, IDispatch *AddInInst, SAFEARRAY **custom);
	STDMETHOD(OnDisconnection)(AddInDesignerObjects::ext_DisconnectMode RemoveMode, SAFEARRAY **custom );
	STDMETHOD(OnAddInsUpdate)(SAFEARRAY **custom );
	STDMETHOD(OnStartupComplete)(SAFEARRAY **custom );
	STDMETHOD(OnBeginShutdown)(SAFEARRAY **custom );

	CComPtr<IDispatch> m_pApplication;
	CComPtr<IDispatch> m_pAddInInstance;
};

OBJECT_ENTRY_AUTO(__uuidof(Connect), CConnect)
