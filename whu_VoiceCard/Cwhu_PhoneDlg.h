#pragma once
//#include "cadodc.h"
// Cwhu_PhoneDlg dialog
#include "afxwin.h"
#include "afxcmn.h"
class Cwhu_PhoneDlg : public CDialogEx
{
	DECLARE_DYNAMIC(Cwhu_PhoneDlg)

public:
	Cwhu_PhoneDlg(CWnd* pParent = NULL);   // standard constructor
	virtual ~Cwhu_PhoneDlg();

// Dialog Data
	enum { IDD = IDD_PHONE };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedRecordstart();
	afx_msg void OnBnClickedSendnotice();
	CDialog * whu_ParentDlg;
	afx_msg void OnBnClickedListen();

	int  SeparateCallNum(CString m_AllNum,CString *m_SeparateNum);
	CString m_HandInputNum;
	DECLARE_EVENTSINK_MAP()
	afx_msg void OnBnClickedChecknumlen();
	CButton m_CheckNumLen;
	BOOL OnInitDialog();
	afx_msg void OnBnClickedButtonback();
	CListCtrl m_ContactList;
	void whu_IntiList();
};
