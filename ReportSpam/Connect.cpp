// Connect.cpp : Implementation of CConnect
#include "stdafx.h"
#include "AddIn.h"
#include "Connect.h"

#include "Converter.h"
#include "ProgressDlg.h"
#include <atlpath.h>

extern CAddInModule _AtlModule;

// When run, the Add-in wizard prepared the registry for the Add-in.
// At a later time, if the Add-in becomes unavailable for reasons such as:
//   1) You moved this project to a computer other than which is was originally created on.
//   2) You chose 'Yes' when presented with a message asking if you wish to remove the Add-in.
//   3) Registry corruption.
// you will need to re-register the Add-in by building the ReportSpamSetup project, 
// right click the project in the Solution Explorer, then choose install.


// CConnect
STDMETHODIMP CConnect::OnConnection(IDispatch *pApplication, AddInDesignerObjects::ext_ConnectMode /*ConnectMode*/, IDispatch *pAddInInst, SAFEARRAY ** /*custom*/ )
{
	pApplication->QueryInterface(__uuidof(IDispatch), (LPVOID*)&m_pApplication);
	pAddInInst->QueryInterface(__uuidof(IDispatch), (LPVOID*)&m_pAddInInstance);

	WCHAR wszFile[MAX_PATH] = {_T('\0')};
	::GetModuleFileName(_AtlModule.GetResourceInstance(), wszFile, MAX_PATH);
	m_strDLLPath = wszFile;
	int nPos = m_strDLLPath.ReverseFind(_T('\\'));
	m_strDLLPath.Delete(nPos + 1, m_strDLLPath.GetLength() - nPos);

	// Create toolbar and button
	AddCommandBar();
	return S_OK;
}

//
STDMETHODIMP CConnect::OnDisconnection(AddInDesignerObjects::ext_DisconnectMode /*RemoveMode*/, SAFEARRAY ** /*custom*/ )
{
	// detach from button notifications
	HRESULT hr = _SpamButtonEvents::DispEventUnadvise((IDispatch*)m_spSpamButton);

	m_pApplication = NULL;
	m_pAddInInstance = NULL;
	return S_OK;
}

//
STDMETHODIMP CConnect::OnAddInsUpdate (SAFEARRAY ** /*custom*/ )
{
	return S_OK;
}

//
STDMETHODIMP CConnect::OnStartupComplete (SAFEARRAY ** /*custom*/ )
{
	return S_OK;
}

//
STDMETHODIMP CConnect::OnBeginShutdown (SAFEARRAY ** /*custom*/ )
{
	return S_OK;
}

//
HRESULT CConnect::AddCommandBar(void)
{
	CComPtr<Office::_CommandBars> spCmdBars;
	CComPtr<Office::CommandBar>   spCmdBar;
	CComPtr<Outlook::_Explorer>   spExplorer;

	// Get outlook command bars
	CComQIPtr<Outlook::_Application> spApp(m_pApplication);
	IfFailRH(spApp->ActiveExplorer(&spExplorer));
	if (!spExplorer)
		return E_ACCESSDENIED;

	HRESULT hr = spExplorer->get_CommandBars(&spCmdBars);
	IfFailRH(hr);

	ATLASSERT(spCmdBars);

	// Add an MSOffice toolbar control
	CComVariant vName(L"Report Spam Tool");
	CComPtr<Office::CommandBar> spNewCmdBar;
	CComVariant vPos(1); // Position it below all toolbands
	CComVariant vTemp(VARIANT_TRUE); // menu is temporary
	CComVariant vEmpty(DISP_E_PARAMNOTFOUND, VT_ERROR);
	hr = spCmdBars->Add(vName, vPos, vEmpty, vTemp, &spNewCmdBar);
	IfFailRH(hr);

	// add buttons
	CComPtr<Office::CommandBarControls> spBarControls;
	hr = spNewCmdBar->get_Controls(&spBarControls); // Get command bar controls
	ATLASSERT(spBarControls);

	_bstr_t bsBtnCaption(L"Report Spam");
	m_spSpamButton = AddButton(spBarControls, 
		bsBtnCaption, bsBtnCaption, bsBtnCaption, IDB_REPORT_SPAM, true);

	// make the button visible
	spNewCmdBar->put_Visible(VARIANT_TRUE);

	// attach to the button notifications
	hr = _SpamButtonEvents::DispEventAdvise((IDispatch*)m_spSpamButton);
	return hr;
}

//
CComPtr<Office::_CommandBarButton> CConnect::AddButton(CComPtr<Office::CommandBarControls> spBarControls,
												_bstr_t bstrCaption, _bstr_t bstrTipText, _bstr_t bstrTag,
												ULONG_PTR bitmapId, bool bEnabled)
{
	HBITMAP hBmp;
	CComQIPtr<Office::_CommandBarButton> spCmdButton;
	CComPtr<Office::CommandBarControl> spNewCommandBarControl;
	CComVariant vToolBarType(1);
	CComVariant vShow(VARIANT_TRUE);
	CComVariant vEmpty(DISP_E_PARAMNOTFOUND, VT_ERROR);

	// add an button to the toolbar
	spBarControls->Add(vToolBarType, vEmpty, vEmpty, vEmpty, vShow, &spNewCommandBarControl);
	ATLASSERT(spNewCommandBarControl);
	spCmdButton = spNewCommandBarControl;
	ATLASSERT(spCmdButton);

	// put the image on the button
	hBmp = (HBITMAP)::LoadImage(_AtlModule.GetResourceInstance(), 
		MAKEINTRESOURCE(bitmapId), IMAGE_BITMAP, 0, 0, LR_LOADMAP3DCOLORS);
	::OpenClipboard(NULL);
	::EmptyClipboard();
	::SetClipboardData(CF_BITMAP, (HANDLE)hBmp);
	::CloseClipboard();
	::DeleteObject(hBmp);
	spCmdButton->put_Style(Office::msoButtonIconAndCaption);// set style before setting bitmap
	HRESULT hr = spCmdButton->PasteFace();
	if (FAILED(hr))
		return NULL;

	// put text data on the button
	spCmdButton->put_Visible(VARIANT_TRUE);
	spCmdButton->put_Caption(bstrCaption);
	spCmdButton->put_Enabled(bEnabled);
	spCmdButton->put_TooltipText(bstrTipText);
	spCmdButton->put_Tag(bstrTag);
	return spCmdButton;
}

//
void __stdcall CConnect::OnSpamButtonClick(IDispatch * /*Office::_CommandBarButton**/ Ctrl, VARIANT_BOOL * CancelDefault)
{
	HRESULT hr = S_OK;
	CComPtr<Outlook::_Explorer>  spExplorer;
	CComPtr<Outlook::Selection>  spSelection;
	CComPtr<Outlook::_NameSpace> spNameSpace;
	CComPtr<Outlook::MAPIFolder> spDelFolder;

	long lCount = 0L;
	CComQIPtr<Outlook::_Application> spApp(m_pApplication);
	hr = spApp->ActiveExplorer(&spExplorer);

	// find selection
	if (spExplorer)
	{
		spExplorer->get_Selection(&spSelection);
		if (spSelection)
			spSelection->get_Count(&lCount);
	}

	// amount of selected items in the outlook explorer
	if (lCount == 0L)
	{
		::MessageBoxW(::GetActiveWindow(), L"There are no items in the selection!\r\nPlease select mail item(s).", L"Spam Tool", MB_ICONINFORMATION);
		return;
	}

	hr = spApp->GetNamespace(L"MAPI", &spNameSpace);
	if (spNameSpace)
	{
		hr = spNameSpace->GetDefaultFolder(Outlook::olFolderDeletedItems, &spDelFolder);
	}

	CComPtr<IConverterSession> spConverter;
	hr = spConverter.CoCreateInstance(CLSID_IConverterSession, NULL, CLSCTX_INPROC_SERVER);
	if (!spConverter)
	{
		::MessageBoxW(::GetActiveWindow(), L"Failed to create converter!", L"Spam Tool", MB_ICONERROR);
		return;
	}

	static int si_encode = IET_BASE64;
	if (si_encode > 5)
		si_encode = 1;

	hr = spConverter->SetEncoding((ENCODINGTYPE)si_encode);
	if (!spConverter)
	{
		::MessageBoxW(::GetActiveWindow(), L"Failed to set encoding.", L"Spam Tool", MB_ICONINFORMATION);
	}
	//si_encode++;

	CProgressDlg dlgProg(lCount);

	try
	{
		// Extract data of selected email items
		for (long i = 1L; i <= lCount; ++i)
		{
			if (dlgProg.IsCancelled())
				break;

			CComPtr<IDispatch> spDisp;
			hr = spSelection->Item(CComVariant(i), &spDisp);

			CComQIPtr<Outlook::_MailItem> spItem(spDisp);
			if (!spItem)
				continue;

			CComPtr<IMessage> pMsg;
			hr = spItem->get_MAPIOBJECT((LPUNKNOWN*)&pMsg);
			if (!pMsg)
				continue;

			dlgProg.Step();

			CStringA strHeaders = GetInternetHeaders(pMsg);

			CComPtr<IStream> pStream;
			hr = CreateStreamOnHGlobal(NULL, TRUE, &pStream);

			hr = spConverter->MAPIToMIMEStm((IMessage*)pMsg, (IStream*)pStream, CCSF_SMTP | CCSF_NOHEADERS);
			if (FAILED(hr))
			{
				::MessageBox(0, 0, _T("Failed MAPIToMIMEStm"), 0);
				continue;
			}

			HGLOBAL hStr;
			hr = GetHGlobalFromStream(pStream, &hStr);
			DWORD dwLen = (DWORD)GlobalSize(hStr);

			LPBYTE lpMsg = (LPBYTE)GlobalLock(hStr);
			GlobalUnlock(hStr);

			if (dwLen > 0)
			{
				bool bPosted = PostData((LPBYTE)(LPCSTR)strHeaders, strHeaders.GetLength(), lpMsg, dwLen);
				if (bPosted)
				{
					pStream.Release();
					pMsg.Release();

					// move to Trash folder if response is OK
					if (spDelFolder)
					{
						CComPtr<IDispatch> spDisp;
						hr = spItem->Move(spDelFolder, &spDisp);
					}
				}
			}
		}
	}
	catch (...)
	{
	}

	dlgProg.Done();
}

//
bool CConnect::PostData(LPBYTE pHdrData, DWORD dwHdrLen, LPBYTE pMsgData, DWORD dwMsgLen)
{
	bool bRes = false;
	BOOL bDeleted = FALSE;
	TCHAR szTempFile[MAX_PATH] = {_T('\0')};

	try
	{
		TCHAR szTemp[MAX_PATH] = {_T('\0')};
		GetTempPath(MAX_PATH, szTemp);
		GetTempFileName(szTemp, _T("R_S"), 0, szTempFile);

		DWORD dwWritten = WriteToFile(szTempFile, pHdrData, dwHdrLen, pMsgData, dwMsgLen);
		bRes = (dwWritten == (dwHdrLen + dwMsgLen)); // the mail content written successfully
	}
	catch (...)
	{
		// some error occured on writing
		bRes = false;
	}

	// the mail content written successfully
	if (bRes)
	{
		try
		{
			STARTUPINFO si;
			PROCESS_INFORMATION pi;

			ZeroMemory( &si, sizeof(si) );
			si.cb = sizeof(si);
			ZeroMemory( &pi, sizeof(pi) );

			CString strFile;
			strFile.Format(_T("\"%sreportmail.exe\" \"%s\""), m_strDLLPath, szTempFile);

			DWORD dwFlags = CREATE_NO_WINDOW;
#ifdef _DEBUG
			dwFlags = 0;
#endif

			BOOL bCreate = ::CreateProcess(NULL, (LPWSTR)(LPCTSTR)strFile, NULL, NULL, FALSE, dwFlags, NULL, NULL, &si, &pi);

			// Wait until child process exits.
			WaitForSingleObject( pi.hProcess, INFINITE );

			// Close process and thread handles.
			CloseHandle( pi.hProcess );
			CloseHandle( pi.hThread );

			bRes = IsResponseOK(szTempFile);

#ifdef _DEBUG
			strFile.Format(_T("%s%s.eml"), m_strDLLPath, ATLPath::FindFileName(szTempFile));
			::CopyFile(szTempFile, strFile, FALSE);
#endif
			// Delete temporary file
			bDeleted = ::DeleteFile(szTempFile);
		}
		catch (...)
		{
			::MessageBox(NULL, _T("Failed to POST data!"), _T("Report Spam"), MB_ICONERROR);
		}
	}
	else
	{
		::MessageBox(NULL, _T("Failed to create eml file!"), _T("Report Spam"), MB_ICONERROR);
	}

	return bRes;
}

//
DWORD CConnect::WriteToFile(LPCTSTR lpszFile, LPBYTE pHdrData, DWORD dwHdrLen, LPBYTE pMsgData, DWORD dwMsgLen)
{
	DWORD dwWritten = 0;

	try
	{
		CHandle hFile;

		// Write MIME data to file
		hFile.Attach( ::CreateFile(lpszFile, GENERIC_WRITE, FILE_SHARE_READ, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL) );
		if (hFile.m_h != 0 && hFile.m_h != INVALID_HANDLE_VALUE)
		{
			if (dwHdrLen > 0)
			{
				::WriteFile(hFile, pHdrData, dwHdrLen, &dwWritten, NULL);
				dwHdrLen = dwWritten;
				dwWritten = 0;
			}

			if (dwMsgLen > 0)
				::WriteFile(hFile, pMsgData, dwMsgLen, &dwWritten, NULL);

			dwWritten += dwHdrLen;
		}
	}
	catch (...)
	{
		// some error occured on writing
		::MessageBox(NULL, _T("Failed to write to the file!"), _T("Report Spam"), MB_ICONERROR);
	}
	return dwWritten; // number of bytes written
}

//
bool CConnect::IsResponseOK(LPCTSTR lpszFile) const
{
	bool isOK = false;
	try
	{
		CHandle hFile;

		hFile.Attach( ::CreateFile(lpszFile, GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL) );
		if (hFile.m_h != 0 && hFile.m_h != INVALID_HANDLE_VALUE)
		{
			BOOL byteOK = 0;
			DWORD dwRead = 0;
			::ReadFile(hFile, &byteOK, sizeof(BOOL), &dwRead, NULL);

			isOK = (byteOK == TRUE);
		}
	}
	catch (...)
	{
	}
	return isOK;
}

//
CStringA CConnect::GetInternetHeaders(IMessage* pMsg)
{
	CStringA strHeaders = "";
	LPSPropValue propVal = NULL;
	HRESULT hr = HrGetOneProp(pMsg, PR_TRANSPORT_MESSAGE_HEADERS_A, &propVal);
	if ( SUCCEEDED(hr) && propVal != NULL)
	{
		strHeaders = CStringA(propVal->Value.lpszA);

		RemoveLine(strHeaders, "MIME-Version: ");
		RemoveLine(strHeaders, "Content-Type: ");

		RemoveDoubleBreak(strHeaders);
	}
	return strHeaders;
}

//
void CConnect::RemoveLine(CStringA& strData, const CStringA& strLine)
{
	int nPos = strData.Find(strLine);
	if (nPos > 0)
	{
		int nEndPos = strData.Find('\n', nPos + 1);
		strData.Delete(nPos, nEndPos - nPos + 1);
	}
}

//
void CConnect::RemoveDoubleBreak(CStringA& strData)
{
	static const int sc_breaks = 4;

	int nLen = strData.GetLength();
	if (nLen > sc_breaks)
	{
		char lastChars[sc_breaks] = {'\0'};
		for (int i = 0; i < sc_breaks; ++i)
			lastChars[i] = strData[nLen - sc_breaks + i];

		if (lastChars[0] == lastChars[2] && lastChars[0] == '\r' &&
			lastChars[1] == lastChars[3] && lastChars[1] == '\n')
		{
			strData.Delete(nLen - 2, 2);
		}
	}
}
