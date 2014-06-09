#pragma once


// Cwhu_ChgPwdDlg dialog

class Cwhu_ChgPwdDlg : public CDialogEx
{
	DECLARE_DYNAMIC(Cwhu_ChgPwdDlg)

public:
	Cwhu_ChgPwdDlg(CWnd* pParent = NULL);   // standard constructor
	virtual ~Cwhu_ChgPwdDlg();
	CDialog * whu_ParentDlg;
	BOOL OnInitDialog();
// Dialog Data
	enum { IDD = IDD_DIALOGCHGPWD };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedOk();
	afx_msg void OnBnClickedCancel();
	CString m_Name;
	CString m_OldPwd;
	CString m_NewPwd;
	CString m_NewPwd02;
	void whu_Init();
	int m_ClickPos;
};
