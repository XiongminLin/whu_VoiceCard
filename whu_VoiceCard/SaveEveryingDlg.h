#pragma once
#include "afxcmn.h"


// CSaveEveryingDlg dialog

class CSaveEveryingDlg : public CDialogEx
{
	DECLARE_DYNAMIC(CSaveEveryingDlg)

public:
	CSaveEveryingDlg(CWnd* pParent = NULL);   // standard constructor
	virtual ~CSaveEveryingDlg();
	BOOL OnInitDialog();
	CDialog * whu_ParentDlg;
// Dialog Data
	enum { IDD = IDD_SAVEEVERYING };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	CProgressCtrl m_Progress;
	afx_msg void OnTimer(UINT_PTR nIDEvent);
	afx_msg LRESULT OnUpdateData(WPARAM wParam, LPARAM lParam);
};
