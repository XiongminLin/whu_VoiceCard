// Cwhu_ChgPwdDlg.cpp : implementation file
//

#include "stdafx.h"
#include "whu_VoiceCard.h"
#include "Cwhu_ChgPwdDlg.h"
#include "afxdialogex.h"
#include "Cwhu_UserDlg.h"

// Cwhu_ChgPwdDlg dialog

IMPLEMENT_DYNAMIC(Cwhu_ChgPwdDlg, CDialogEx)

Cwhu_ChgPwdDlg::Cwhu_ChgPwdDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(Cwhu_ChgPwdDlg::IDD, pParent)
	, m_Name(_T(""))
	, m_OldPwd(_T(""))
	, m_NewPwd(_T(""))
	, m_NewPwd02(_T(""))
{

}

Cwhu_ChgPwdDlg::~Cwhu_ChgPwdDlg()
{
}

void Cwhu_ChgPwdDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Text(pDX, ID_NAME, m_Name);
	DDX_Text(pDX, IDC_OLDPWD, m_OldPwd);
	DDX_Text(pDX, IDC_NEWPWD, m_NewPwd);
	DDX_Text(pDX, IDC_NEWPWD2, m_NewPwd02);
}


BEGIN_MESSAGE_MAP(Cwhu_ChgPwdDlg, CDialogEx)
	ON_BN_CLICKED(IDOK, &Cwhu_ChgPwdDlg::OnBnClickedOk)
	ON_BN_CLICKED(IDCANCEL, &Cwhu_ChgPwdDlg::OnBnClickedCancel)
END_MESSAGE_MAP()


// Cwhu_ChgPwdDlg message handlers


void Cwhu_ChgPwdDlg::OnBnClickedOk()
{
	// TODO: Add your control notification handler code here
	UpdateData(true);
	if (m_NewPwd02 == m_NewPwd)
	{
		((Cwhu_UserDlg*)whu_ParentDlg)->m_UserList.SetItemText(m_ClickPos,1,m_Name);
		((Cwhu_UserDlg*)whu_ParentDlg)->m_UserList.SetItemText(m_ClickPos,2,m_NewPwd);
	}
	CDialogEx::OnOK();
}


void Cwhu_ChgPwdDlg::OnBnClickedCancel()
{
	// TODO: Add your control notification handler code here
	CDialogEx::OnCancel();
}
BOOL Cwhu_ChgPwdDlg::OnInitDialog()
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
	m_ClickPos = ((Cwhu_UserDlg*)whu_ParentDlg)->m_ClickPos;
	whu_Init();

	return true;
}
void Cwhu_ChgPwdDlg::whu_Init()
{
	m_Name = ((Cwhu_UserDlg*)whu_ParentDlg)->m_UserList.GetItemText(m_ClickPos,1);
	m_OldPwd = ((Cwhu_UserDlg*)whu_ParentDlg)->m_UserList.GetItemText(m_ClickPos,2);
	UpdateData(FALSE);
}