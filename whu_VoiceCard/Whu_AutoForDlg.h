#pragma once
#include "afxcmn.h"


// CWhu_AutoForDlg dialog

class CWhu_AutoForDlg : public CDialogEx
{
	DECLARE_DYNAMIC(CWhu_AutoForDlg)

public:
	CWhu_AutoForDlg(CWnd* pParent = NULL);   // standard constructor
	virtual ~CWhu_AutoForDlg();
	CDialog * whu_ParentDlg;
	BOOL OnInitDialog();
// Dialog Data
	enum { IDD = IDD_AUTOFAX };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedOk();
	CListCtrl m_List;
	void whu_InitList();
	CListCtrl m_SorList;
	int m_ClickPos;
	void whu_AddAutoForArr(CStringArray &m_SorArr,CStringArray &m_AddArr);
};
