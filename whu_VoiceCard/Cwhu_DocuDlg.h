#pragma once
#include "afxcmn.h"


// Cwhu_DocuDlg dialog

class Cwhu_DocuDlg : public CDialogEx
{
	DECLARE_DYNAMIC(Cwhu_DocuDlg)

public:
	Cwhu_DocuDlg(CWnd* pParent = NULL);   // standard constructor
	virtual ~Cwhu_DocuDlg();
	CDialog * whu_ParentDlg;
// Dialog Data
	enum { IDD = IDD_DOCUMENT };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	CListCtrl m_FaxList;
	void whu_InitList();
	afx_msg void OnBnClickedBaddinfo();
	CString m_info;
	afx_msg void OnBnClickedBsavedoc();
	CListCtrl m_ListNotice;
};
