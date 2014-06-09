#pragma once
#include"resource.h"
#include "afxcmn.h"
#include "CRange.h"
#include "CSheets.h"
#include "CWorkbook.h"
#include "CWorkbooks.h"
#include "CApplication.h"
#include "CWorksheet.h"
#include "CWorksheets.h"
//#include "Lin_Excel_List.h"
// Cwhu_ContractDlg dialog

class Cwhu_ContractDlg : public CDialogEx
{
	DECLARE_DYNAMIC(Cwhu_ContractDlg)

public:
	Cwhu_ContractDlg(CWnd* pParent = NULL);   // standard constructor
	virtual ~Cwhu_ContractDlg();
	CDialog * whu_ParentDlg;
	BOOL OnInitDialog();
// Dialog Data
	enum { IDD = IDD_CONTRACT };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	CListCtrl m_List;
	//void whu_InitList();
	int pos;
	int ConHeadNum;
	CString m_TelNum;
	CString m_OfficeNum;
	CString m_FaxNum;
	CString m_GroupFaxNum[100];
	CString m_GroupNoticeNum[100];
	int m_ChoosedPos;
	afx_msg void OnNMRClickLiscontract(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnDial();
	afx_msg void OnSendnotice();
	afx_msg void OnSendfax();
	afx_msg void OnDialoffice();
	afx_msg void OnSendgroupfax();
	CString GetStringFromVariant(_variant_t var);
	afx_msg void OnMyDial();
	afx_msg void OnMySendSingleNotice();
	afx_msg void OnMyDialOffice();
	afx_msg void OnMySendSingleNoticeOffice();
	afx_msg void OnBnClickedButtonsave();
	afx_msg void OnBnClickedButtonback();
	afx_msg void OnNMClickLiscontract(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnBnClickedButtonadd();
	afx_msg void OnBnClickedButtondelete();
	void whu_SaveContact();
	afx_msg void OnLvnColumnclickLiscontract(NMHDR *pNMHDR, LRESULT *pResult);
	void whu_SortContactList(CListCtrl &m_List,int colum,int mode);
	afx_msg void OnBnClickedImport();
	CString Lin_GetEnglishCharacter(int Num);
	void Lin_InitList(CListCtrl &m_List,CStringArray &ColumName);
	void Lin_InsertList(CListCtrl &m_List,CStringArray &ItemArr);
	void Lin_ExportListToExcel(CListCtrl &m_List,CString m_FilePath);
	bool Lin_InportExcelToList(CString m_FilePath,CListCtrl &m_List);
	void Lin_ImportArrToList(CStringArray *m_DutyArr,int Num,CListCtrl &m_List);
	int Lin_ListToArr(CListCtrl &m_List,CStringArray *m_DutyArr);
	bool whu_findfile(CString m_file,CString folder);
	afx_msg void OnBnClickedShowphoto();
	afx_msg void OnBnClickedChangepic();
};
