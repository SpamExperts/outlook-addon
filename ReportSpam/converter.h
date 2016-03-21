#pragma once

#define CCSF_SMTP			0x0002 
#define CCSF_NOHEADERS		0x0004  
#define CCSF_NO_MSGID		0x4000
#define CCSF_USE_TNEF		0x0010
#define CCSF_INCLUDE_BCC	0x0020
#define CCSF_8BITHEADERS	0x0040
#define CCSF_USE_RTF		0x0080

//
static const GUID CLSID_IConverterSession = 
{ 0x4e3a7680, 0xb77a, 0x11d0, { 0x9d, 0xa5, 0x0, 0xc0, 0x4f, 0xd6, 0x56, 0x85 } };


//DEFINE_GUID(CLSID_IConverterSession, 0x4e3a7680, 0xb77a, 0x11d0, 0x9d, 0xa5, 0x0, 0xc0, 0x4f, 0xd6, 0x56, 0x85); 
DEFINE_GUID(IID_IConverterSession, 0x4b401570, 0xb77b, 0x11d0, 0x9d, 0xa5, 0x0, 0xc0, 0x4f, 0xd6, 0x56, 0x85);

typedef enum tagENCODINGTYPE {
    IET_BINARY = 0,
    IET_BASE64 = 1,
    IET_UUENCODE = 2,
    IET_QP = 3,
    IET_7BIT = 4,
    IET_8BIT = 5,
    IET_INETCSET = 6,
    IET_UNICODE = 7,
    IET_RFC1522 = 8,
    IET_ENCODED = 9,
    IET_CURRENT = 10,
    IET_UNKNOWN = 11,
    IET_BINHEX40 = 12,
    IET_LAST = 13
} ENCODINGTYPE;

interface DECLSPEC_UUID("4B401570-B77B-11D0-9DA5-00C04FD65685") IConverterSession : public IUnknown
{
    virtual void __stdcall UnknownMethod1();
	virtual HRESULT __stdcall SetEncoding(ENCODINGTYPE et);
    virtual void __stdcall UnknownMethod3()=0;
	virtual HRESULT __stdcall MIMEToMAPI(LPSTREAM pstm, LPMESSAGE pmsg, LPCSTR pszSrcSrv, ULONG ulFlags);
	virtual HRESULT __stdcall MAPIToMIMEStm(LPMESSAGE pmsg, LPSTREAM pstm, ULONG ulFlags);
	virtual void __stdcall UnknownMethod6()=0;
	virtual void __stdcall UnknownMethod7()=0;
	virtual void __stdcall UnknownMethod8()=0;
	virtual void __stdcall UnknownMethod9()=0;
	virtual void __stdcall UnknownMethod10()=0;
	virtual void __stdcall UnknownMethod11()=0;
	virtual void __stdcall UnknownMethod12()=0;
};
