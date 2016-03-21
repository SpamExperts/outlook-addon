// ProgressDlg.h : Declaration of the CProgressDlg

#pragma once

#include "resource.h"       // main symbols
#include <atlhost.h>


// CProgressDlg

class CProgressDlg : 
	public CAxDialogImpl<CProgressDlg>
{
public:
	CProgressDlg(long nCount);
	~CProgressDlg(void);

	enum { IDD = IDD_PROGRESSDLG };

BEGIN_MSG_MAP(CProgressDlg)
	MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
	COMMAND_HANDLER(IDOK, BN_CLICKED, OnClickedOK)
	COMMAND_HANDLER(IDCANCEL, BN_CLICKED, OnClickedCancel)
	CHAIN_MSG_MAP(CAxDialogImpl<CProgressDlg>)
END_MSG_MAP()

// Handler prototypes:
//  LRESULT MessageHandler(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
//  LRESULT CommandHandler(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
//  LRESULT NotifyHandler(int idCtrl, LPNMHDR pnmh, BOOL& bHandled);

	LRESULT OnInitDialog(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	LRESULT OnClickedOK(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
	LRESULT OnClickedCancel(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);

	void Step(void);
	void Done(void);
	bool IsCancelled(void) const;

private:
	static DWORD WINAPI ThreadProc(LPVOID lpParam);

private:
	bool m_bCancelled;
	long m_nCount, m_nCurrent;

	CEvent  m_evStart, m_evDone;
	CHandle m_hThread;
	DWORD m_dwThreadId;
};
