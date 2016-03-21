// ProgressDlg.cpp : Implementation of CProgressDlg

#include "stdafx.h"
#include "ProgressDlg.h"


// CProgressDlg

//
CProgressDlg::CProgressDlg(long nCount)
: m_bCancelled(false)
, m_nCount(nCount)
, m_nCurrent(0L)
, m_evStart(FALSE, FALSE)
, m_evDone(FALSE, FALSE)
{
	m_hThread.Attach( ::CreateThread(NULL, 0, ThreadProc, this, 0, &m_dwThreadId) );
	::WaitForSingleObject(m_evStart, INFINITE);
}

//
CProgressDlg::~CProgressDlg()
{
}

//
LRESULT CProgressDlg::OnInitDialog(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	CAxDialogImpl<CProgressDlg>::OnInitDialog(uMsg, wParam, lParam, bHandled);
	bHandled = TRUE;

	GetDlgItem(IDC_PROGRESS_SENT).SendMessage(PBM_SETRANGE32, (WPARAM)0, (LPARAM)m_nCount);
	GetDlgItem(IDC_PROGRESS_SENT).PostMessage(PBM_SETSTEP, (WPARAM)1, 0);

	m_evStart.Set();
	return 1;  // Let the system set the focus
}

//
LRESULT CProgressDlg::OnClickedOK(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
{
	CString strInfo;
	if (!m_bCancelled)
	{
		strInfo.Format(_T("Sending mail %d of %d..."), m_nCurrent, m_nCount);
	}
	else
	{
		strInfo = _T("Canceled... Please wait.");
	}
	SetDlgItemText(IDC_STATIC_STATE, (LPCTSTR)strInfo);
	//EndDialog(wID);
	return 0;
}

//
LRESULT CProgressDlg::OnClickedCancel(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
{
	m_bCancelled = true;
	SetDlgItemText(IDC_STATIC_STATE, _T("Canceled. Please wait..."));
	::WaitForSingleObject(m_evDone, INFINITE);

	EndDialog(wID);
	return 0;
}

//
void CProgressDlg::Step()
{
	::InterlockedIncrement(&m_nCurrent);
	GetDlgItem(IDC_PROGRESS_SENT).PostMessage(PBM_STEPIT);
	PostMessage(WM_COMMAND, MAKEWPARAM(IDOK, BN_CLICKED), NULL);
}

//
void CProgressDlg::Done()
{
	if (!m_bCancelled)
		PostMessage(WM_COMMAND, MAKEWPARAM(IDCANCEL, BN_CLICKED), NULL);
	m_evDone.Set();

	::WaitForSingleObject(m_hThread, INFINITE);
}

//
bool CProgressDlg::IsCancelled(void) const
{
	return m_bCancelled;
}

//
DWORD WINAPI CProgressDlg::ThreadProc(LPVOID lpParam)
{
	CProgressDlg* pDlg = (CProgressDlg*)lpParam;
	pDlg->DoModal();

	return 0;
}
