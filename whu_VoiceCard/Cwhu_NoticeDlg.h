#pragma once
#include "afxcmn.h"


// Cwhu_NoticeDlg dialog

class Cwhu_NoticeDlg : public CDialogEx
{
	DECLARE_DYNAMIC(Cwhu_NoticeDlg)

public:
	Cwhu_NoticeDlg(CWnd* pParent = NULL);   // standard constructor
	virtual ~Cwhu_NoticeDlg();
	BOOL OnInitDialog();
// Dialog Data
	enum { IDD = IDD_NOTICE };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedRecordstart();
	CDialog * whu_ParentDlg;
	CListCtrl m_ContactList;
	afx_msg void OnBnClickedButtonback();
	void whu_IntiList();
	CString m_NoticeNum;
	afx_msg void OnBnClickedContractadd();
	void whu_SaveContact();
	afx_msg void OnBnClickedSendnotice();
	CStringArray m_NoticeNumArr;
	CStringArray m_NoticeState;
	CString m_filename;
	afx_msg void OnBnClickedListen();
	afx_msg void OnTimer(UINT_PTR nIDEvent);
	void whu_showNoticeState(int whu_NoticeState);
	void whu_SortContactList(CListCtrl &m_List,int colum,int mode);
	afx_msg void OnLvnColumnclickListcontact02(NMHDR *pNMHDR, LRESULT *pResult);
	int NumCount;
	afx_msg LRESULT OnUpdateData(WPARAM wParam, LPARAM lParam);
};
