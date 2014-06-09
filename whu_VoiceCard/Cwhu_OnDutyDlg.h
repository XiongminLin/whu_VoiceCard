#pragma once
#include "CAdodc.h"
#include "CDataGrid.h"
#include "afxcmn.h"
#include "atlcomtime.h"
//#include "CRange.h"
//#include "CSheets.h"
//#include "CWorkbook.h"
//#include "CWorkbooks.h"
//#include "CApplication.h"
//#include "CWorksheet.h"
//#include "CWorksheets.h"
// Cwhu_OnDutyDlg dialog

class Cwhu_OnDutyDlg : public CDialogEx
{
	DECLARE_DYNAMIC(Cwhu_OnDutyDlg)

public:
	Cwhu_OnDutyDlg(CWnd* pParent = NULL);   // standard constructor
	virtual ~Cwhu_OnDutyDlg();
	CDialog * whu_ParentDlg;
	BOOL OnInitDialog();

// Dialog Data
	enum { IDD = IDD_ONDUTY };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	CString m_NewOnDutyLeader;
	CString m_NewOnDutyPerson;
	afx_msg void OnBnClickedSwitchduty();
	CListCtrl m_MyList;
	afx_msg void OnBnClickedDelete();
	afx_msg void OnBnClickedImportduty();
	afx_msg void OnBnClickedBack();
	CString Lin_GetEnglishCharacter(int Num);
	void Lin_InitList(CListCtrl &m_List,CStringArray &ColumName);
	void Lin_InsertList(CListCtrl &m_List,CStringArray &ItemArr);
	void Lin_ExportListToExcel(CListCtrl &m_List,CString m_FilePath);
	bool Lin_InportExcelToList(CString m_FilePath,CListCtrl &m_List);
	void Lin_ImportArrToList(CStringArray *m_DutyArr,int Num,CListCtrl &m_List);
	int Lin_ListToArr(CListCtrl &m_List,CStringArray *m_DutyArr);
	int DutyHeadNum;
	afx_msg void OnLvnColumnclickListduty(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnNMClickListduty(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnBnClickedModify();
	int m_ChoosedPos;
	afx_msg void OnBnClickedAdd();
	COleDateTime m_DutyTime;
	bool whu_findfile(CString m_file,CString folder);
};
