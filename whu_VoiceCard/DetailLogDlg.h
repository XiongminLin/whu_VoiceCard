#pragma once


// CDetailLogDlg dialog

class CDetailLogDlg : public CDialogEx
{
	DECLARE_DYNAMIC(CDetailLogDlg)

public:
	CDetailLogDlg(CWnd* pParent = NULL);   // standard constructor
	virtual ~CDetailLogDlg();
	CDialog * whu_ParentDlg;
	BOOL OnInitDialog();
	void whu_ShowData();
	CString m_Edit;
	int m_ClickPos;
// Dialog Data
	enum { IDD = IDD_LOG };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedCancel();
	afx_msg void OnBnClickedLogopenfile();
	afx_msg void OnBnClickedButton1();
	afx_msg void OnBnClickedOk();
};
