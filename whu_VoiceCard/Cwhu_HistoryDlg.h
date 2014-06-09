#pragma once
#include "afxcmn.h"
#include "Cwhu_DetailLog.h"
#include "afxdtctl.h"

// Cwhu_HistoryDlg dialog
typedef struct{
CString time;
CString taskmode;
CString unit;
CString name;
CString number;
CString file;
CString beizhu;

}HistoryData;
class Cwhu_HistoryDlg : public CDialogEx
{
	DECLARE_DYNAMIC(Cwhu_HistoryDlg)

public:
	Cwhu_HistoryDlg(CWnd* pParent = NULL);   // standard constructor
	virtual ~Cwhu_HistoryDlg();
	HistoryData m_data[100];
	void whu_InitList();
	BOOL OnInitDialog();
	int pos;
	void whu_ModifyCheck(CString m_str);
	void whu_AddHistoryList(CString time,CString duty,CString taskmode,CString unit,CString name,CString number,CString file,CString stste,CString beizhu);
private:
	
	
// Dialog Data
	enum { IDD = IDD_HISTORY };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	CListCtrl m_HistoryList;
	afx_msg void OnNMRClickHistorylist(NMHDR *pNMHDR, LRESULT *pResult);
	//afx_msg void OnNMClickHistorylist(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnBnClickedHislook();
	CDialog * whu_ParentDlg;
	afx_msg void OnBnClickedHisedit();
	afx_msg void OnBnClickedHissave();
	afx_msg void OnBnClickedHiscancel();
	afx_msg void OnAddDetailLog();
	afx_msg void OnDtnDatetimechangeDatetimepickerlog(NMHDR *pNMHDR, LRESULT *pResult);
	CDateTimeCtrl m_DateCtl;
	afx_msg void OnBnClickedButtonquit();
	void SwitchLog();
	void SaveLog();
	CStringArray whu_ArrLog[10];
	afx_msg void OnLvnColumnclickHistorylist(NMHDR *pNMHDR, LRESULT *pResult);
	CString whu_ChineseToPinyin(CString chinese);
	char convert(wchar_t n);
	int m_ClickPos;
	void OnSaveLogToFile();
	void whu_ListToTxt(CListCtrl *m_list,CString m_txtfile);
	bool whu_findfile(CString m_file,CString folder);
	afx_msg void OnOpenFile();
	afx_msg void OnShowDetailLog();
	void whu_ReSize();
};
