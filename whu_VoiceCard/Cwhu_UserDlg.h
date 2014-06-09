#pragma once
#include "CvvImage.h"
#include "afxcmn.h"

// Cwhu_UserDlg dialog

class Cwhu_UserDlg : public CDialogEx
{
	DECLARE_DYNAMIC(Cwhu_UserDlg)

public:
	Cwhu_UserDlg(CWnd* pParent = NULL);   // standard constructor
	virtual ~Cwhu_UserDlg();
	CDialog * whu_ParentDlg;
	BOOL OnInitDialog();
	void whu_InitUser();
// Dialog Data
	enum { IDD = IDD_USER };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	CListCtrl m_UserList;
	void Lin_InportExcelToList(CString m_FilePath,CListCtrl &m_List);
	CString Lin_GetEnglishCharacter(int Num);
	void Lin_InitList(CListCtrl &m_List,CStringArray &ColumName);
	void Lin_InsertList(CListCtrl &m_List,CStringArray &ItemArr);
	void Lin_ImportArrToList(CStringArray *m_DutyArr,int Num,CListCtrl &m_List);
	int Lin_ListToArr(CListCtrl &m_List,CStringArray *m_DutyArr);
	afx_msg void OnBnClickedBack();
	afx_msg void OnNMClickList1(NMHDR *pNMHDR, LRESULT *pResult);
	int m_ClickPos;
	afx_msg void OnBnClickedAccept();
	afx_msg void OnBnClickedRefuse();
	afx_msg void OnBnClickedLogoff();
	afx_msg void OnBnClickedChgpwd();
};
