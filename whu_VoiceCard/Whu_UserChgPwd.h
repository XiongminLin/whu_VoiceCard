#pragma once


// CWhu_UserChgPwd dialog

class CWhu_UserChgPwd : public CDialogEx
{
	DECLARE_DYNAMIC(CWhu_UserChgPwd)

public:
	CWhu_UserChgPwd(CWnd* pParent = NULL);   // standard constructor
	virtual ~CWhu_UserChgPwd();
	CDialog * whu_ParentDlg;
	BOOL OnInitDialog();
// Dialog Data
	enum { IDD = IDD_USERCHGPWD };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	CString m_Name;
	CString m_OldPwd;
	CString m_NewPwd;
	CString m_NewPwd02;
	void whu_Init();
	afx_msg void OnBnClickedOk();
};
