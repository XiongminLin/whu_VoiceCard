#pragma once


// m_PhoneDlg dialog

class m_PhoneDlg : public CDialogEx
{
	DECLARE_DYNAMIC(m_PhoneDlg)

public:
	m_PhoneDlg(CWnd* pParent = NULL);   // standard constructor
	virtual ~m_PhoneDlg();

// Dialog Data
	enum { IDD = IDD_PHONE };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
};
