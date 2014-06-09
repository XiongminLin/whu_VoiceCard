#pragma once


// Cwhu_RecvCallDlg dialog

class Cwhu_RecvCallDlg : public CDialogEx
{
	DECLARE_DYNAMIC(Cwhu_RecvCallDlg)

public:
	Cwhu_RecvCallDlg(CWnd* pParent = NULL);   // standard constructor
	virtual ~Cwhu_RecvCallDlg();
	BOOL OnInitDialog();
// Dialog Data
	enum { IDD = IDD_DIALOGRECVCALL };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnTimer(UINT_PTR nIDEvent);
	int Time;
	int X,Y;
	HANDLE hThread;
	CDialog * whu_ParentDlg;
	void whu_ShakeWindow();
	bool whu_PickUp;
	afx_msg void OnBnClickedOk();
};
