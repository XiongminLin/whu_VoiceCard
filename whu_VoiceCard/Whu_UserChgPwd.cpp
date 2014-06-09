// Whu_UserChgPwd.cpp : implementation file
//

#include "stdafx.h"
#include "whu_VoiceCard.h"
#include "Whu_UserChgPwd.h"
#include "afxdialogex.h"
#include "whu_VoiceCardDlg.h"

// CWhu_UserChgPwd dialog

IMPLEMENT_DYNAMIC(CWhu_UserChgPwd, CDialogEx)

CWhu_UserChgPwd::CWhu_UserChgPwd(CWnd* pParent /*=NULL*/)
	: CDialogEx(CWhu_UserChgPwd::IDD, pParent)
	, m_Name(_T(""))
	, m_OldPwd(_T(""))
	, m_NewPwd(_T(""))
	, m_NewPwd02(_T(""))
{

}

CWhu_UserChgPwd::~CWhu_UserChgPwd()
{
}

void CWhu_UserChgPwd::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Text(pDX, IDC_EDIT1, m_Name);
	DDX_Text(pDX, IDC_EDIT2, m_OldPwd);
	DDX_Text(pDX, IDC_EDIT6, m_NewPwd);
	DDX_Text(pDX, IDC_EDIT7, m_NewPwd02);
}


BEGIN_MESSAGE_MAP(CWhu_UserChgPwd, CDialogEx)
	ON_BN_CLICKED(IDOK, &CWhu_UserChgPwd::OnBnClickedOk)
END_MESSAGE_MAP()


// CWhu_UserChgPwd message handlers

BOOL CWhu_UserChgPwd::OnInitDialog()
{
	CDialogEx::OnInitDialog();
	//将对话框移至屏幕中间//
	CRect rect;
	GetClientRect(rect);
	int srX=rect.right-rect.left;
	int srY=rect.bottom-rect.top;
	int cx,cy; 
	cx = GetSystemMetrics(SM_CXSCREEN); 
	cy = GetSystemMetrics(SM_CYSCREEN);
	int x = (cx-srX)/2;
	int y = (cy-srY)/2;
	SetWindowPos(NULL, x,y,srX,srY, SWP_NOMOVE||SWP_NOSIZE);
	whu_Init();

	return true;
}
void CWhu_UserChgPwd::whu_Init()
{
	m_Name = ((Cwhu_VoiceCardDlg*)whu_ParentDlg)->m_LogName;
	UpdateData(FALSE);
}

void CWhu_UserChgPwd::OnBnClickedOk()
{
	// TODO: Add your control notification handler code here
	UpdateData(true);
	if (m_NewPwd == m_NewPwd02 &&m_OldPwd == ((Cwhu_VoiceCardDlg*)whu_ParentDlg)->whu_password)
	{
		for (int i=0;i<((Cwhu_VoiceCardDlg*)whu_ParentDlg)->whu_User[0].GetSize();i++)
		{

			if (((Cwhu_VoiceCardDlg*)whu_ParentDlg)->whu_User[0].GetAt(i)==m_Name )
			{
				
				((Cwhu_VoiceCardDlg*)whu_ParentDlg)->whu_User[1].RemoveAt(i);
				((Cwhu_VoiceCardDlg*)whu_ParentDlg)->whu_User[1].InsertAt(i,m_NewPwd);
			}
		}
		AfxMessageBox(_T("密码修改成功"));
		CDialogEx::OnOK();
	}else{
		AfxMessageBox(_T("请检查密码是否正确"));
	}
	
}
