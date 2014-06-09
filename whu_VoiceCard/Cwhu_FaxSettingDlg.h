#pragma once
//#include "afxcmn.h"
// Cwhu_FaxSettingDlg dialog

class Cwhu_FaxSettingDlg : public CDialogEx
{
	DECLARE_DYNAMIC(Cwhu_FaxSettingDlg)

public:
	Cwhu_FaxSettingDlg(CWnd* pParent = NULL);   // standard constructor
	virtual ~Cwhu_FaxSettingDlg();
	CDialog * whu_ParentDlg;
	BOOL OnInitDialog();
// Dialog Data
	enum { IDD = IDD_FAXSETTING };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedOk();
	afx_msg void OnBnClickedCancel();
	CListCtrl m_ListFrom;
	CListCtrl m_ListTo;
	void whu_InitList();
	CStringArray whu_AutoForward;
	int m_ClickPos;
	afx_msg void OnHdnItemclickListform(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnNMClickListform(NMHDR *pNMHDR, LRESULT *pResult);
	void whu_ShowToList(CString m_FaxNum);//实现根据主对话框的字符串数组变量来显示自动转发的名单//
	afx_msg void OnBnClickedAddcontact();
	void Lin_ImportAutoForArr(CString m_FilePath,CStringArray &m_AutoForward);
	CString Lin_GetEnglishCharacter(int Num);
	void Lin_ExportForArrToExcel(CStringArray &m_AutoForward,CString m_FilePath);
	bool whu_findfile(CString m_file,CString folder);
	afx_msg void OnBnClickedDelete();
	void whu_AddAutoForArr(CStringArray &m_SorArr,CStringArray &m_AddArr);
	//BOOL whu_BAutoForward;
	afx_msg void OnBnClickedFaxAutoForward();
	void m_ShowForwardItem(CString FaxNume);
	afx_msg void OnNMCustomdrawListform(NMHDR *pNMHDR, LRESULT *pResult);
};
