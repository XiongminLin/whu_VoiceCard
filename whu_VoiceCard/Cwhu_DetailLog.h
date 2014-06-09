#pragma once


// Cwhu_DetailLog dialog

class Cwhu_DetailLog : public CDialogEx
{
	DECLARE_DYNAMIC(Cwhu_DetailLog)

public:
	Cwhu_DetailLog(CWnd* pParent = NULL);   // standard constructor
	virtual ~Cwhu_DetailLog();
	CDialog * whu_ParentDlg;
// Dialog Data
	enum { IDD = IDD_LOG };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	BOOL OnInitDialog();
	DECLARE_MESSAGE_MAP()
};
