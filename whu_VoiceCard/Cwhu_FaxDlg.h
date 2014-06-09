#pragma once
#include "afxcmn.h"
//#include "cadodc.h"
#include "Cwhu_FaxSettingDlg.h"
#include "afxdtctl.h"
// Cwhu_FaxDlg dialog

class Cwhu_FaxDlg : public CDialogEx
{
	DECLARE_DYNAMIC(Cwhu_FaxDlg)

public:
	Cwhu_FaxDlg(CWnd* pParent = NULL);   // standard constructor
	virtual ~Cwhu_FaxDlg();
	BOOL OnInitDialog();
// Dialog Data
	enum { IDD = IDD_FAX };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	int DocumentToTiff(CString m_Inputfile,CString &m_OutputFile);
	afx_msg void OnBnClickedSendok();
	CDialog * whu_ParentDlg;
	SYSTEMTIME whu_SystemTime;
	afx_msg void OnBnClickedChoosefaxfile();
	CString whu_DocFile;
	CString m_faxNum;
	CStringArray m_FaxNumArr;
	int  SeparateCallNum(CString m_AllNum,CString *m_SeparateNum);
	DECLARE_EVENTSINK_MAP()
	afx_msg void OnBnClickedSendcancle();
	
	void whu_InitContactList();
	afx_msg void OnNMRClickFaxcontactlist(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnEnChangeFaxnumber();
	afx_msg void OnBnClickedButtonaddnum();
	void whu_SaveContact();
	afx_msg void OnBnClickedButtonreadfax();
	afx_msg void OnBnClickedButtonfaxsetting();
	void whu_ShowFaxState(int fax_state);
	afx_msg void OnTimer(UINT_PTR nIDEvent);
	afx_msg void OnLvnColumnclickFaxcontactlist(NMHDR *pNMHDR, LRESULT *pResult);
	void whu_SortContactList(CListCtrl &m_List,int colum,int mode);
	void Lin_ImportArrToList(CStringArray *m_DutyArr,int Num,CListCtrl &m_List);
	void Lin_InitList(CListCtrl &m_List,CStringArray &ColumName);
	void Lin_InsertList(CListCtrl &m_List,CStringArray &ItemArr);
	void whu_ReSize();
public:
	int test;
	CListCtrl m_FaxContactList;
	CStringArray whu_AutoForward;
	afx_msg void OnNMClickFaxcontactlist(NMHDR *pNMHDR, LRESULT *pResult);
	int NumCount;
	afx_msg LRESULT OnUpdateData(WPARAM wParam, LPARAM lParam);

};
