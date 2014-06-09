
// OpenExcelDlg.h : header file
//

#pragma once
#include "afxcmn.h"
#include "Resource.h"
#include "CRange.h"
#include "CSheets.h"
#include "CWorkbook.h"
#include "CWorkbooks.h"
#include "CApplication.h"
#include "CWorksheet.h"
#include "CWorksheets.h"
// COpenExcelDlg dialog
class COpenExcelDlg : public CDialogEx
{
// Construction
public:
	COpenExcelDlg(CWnd* pParent = NULL);	// standard constructor

// Dialog Data
	enum { IDD = IDD_OPENEXCEL_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support


// Implementation
protected:
	HICON m_hIcon;

	// Generated message map functions
	BOOL OnInitDialog();
	//afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	//afx_msg void OnPaint();
	//afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedAddfile();
	void Lin_InitList(CListCtrl &m_List,CStringArray &ColumName);
	void Lin_InsertList(CListCtrl &m_List,CStringArray &ItemArr);
	CListCtrl m_MyList;
	void Lin_ExportListToExcel(CListCtrl &m_List);
	void Lin_InportExcelToList(CString m_FilePath,CListCtrl &m_List);
	afx_msg void OnBnClickedOk();
	CString Lin_GetEnglishCharacter(int Num);
	CString m_DocFile;
	afx_msg void OnBnClickedCancel();
	CDialog * whu_ParentDlg;
};
