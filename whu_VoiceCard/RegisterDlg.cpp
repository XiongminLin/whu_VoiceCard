// RegisterDlg.cpp : implementation file
//

#include "stdafx.h"
#include "whu_VoiceCard.h"
#include "RegisterDlg.h"
#include "afxdialogex.h"
#include "whu_VoiceCardDlg.h"

// CRegisterDlg dialog

IMPLEMENT_DYNAMIC(CRegisterDlg, CDialogEx)

CRegisterDlg::CRegisterDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(CRegisterDlg::IDD, pParent)
	, m_Name(_T(""))
	, m_Pwd01(_T(""))
	, m_Pwd02(_T(""))
{

}

CRegisterDlg::~CRegisterDlg()
{
}

void CRegisterDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Text(pDX, IDC_EDIT1, m_Name);
	DDX_Text(pDX, IDC_EDIT2, m_Pwd01);
	DDX_Text(pDX, IDC_EDIT6, m_Pwd02);
}


BEGIN_MESSAGE_MAP(CRegisterDlg, CDialogEx)
	ON_BN_CLICKED(IDOK, &CRegisterDlg::OnBnClickedOk)
END_MESSAGE_MAP()


// CRegisterDlg message handlers


void CRegisterDlg::OnBnClickedOk()
{
	// TODO: Add your control notification handler code here
	UpdateData(true);
	bool m_Check = false;
	for (int i=0;i<((Cwhu_VoiceCardDlg*)whu_ParentDlg)->whu_User[0].GetSize();i++)
	{

		if (((Cwhu_VoiceCardDlg*)whu_ParentDlg)->whu_User[0].GetAt(i) == m_Name)
		{
			m_Check  = true;
		}
	}
	if (m_Check == true)
	{
		AfxMessageBox(_T("该用户已注册"));
	}
	else{
		if (m_Name == _T(""))
		{
			AfxMessageBox(_T("用户名为空"));
		}
		else if (m_Pwd01 == m_Pwd02 && m_Pwd02 !=_T(""))
		{
			int pos = ((Cwhu_VoiceCardDlg*)whu_ParentDlg)->whu_User[0].GetSize();
			int test1 = ((Cwhu_VoiceCardDlg*)whu_ParentDlg)->whu_User[1].GetSize();
			int test2 = ((Cwhu_VoiceCardDlg*)whu_ParentDlg)->whu_User[0].GetSize();
			((Cwhu_VoiceCardDlg*)whu_ParentDlg)->whu_User[0].Add(m_Name);
			((Cwhu_VoiceCardDlg*)whu_ParentDlg)->whu_User[1].Add(m_Pwd01);
			((Cwhu_VoiceCardDlg*)whu_ParentDlg)->whu_User[2].Add(_T("申请注册"));
			((Cwhu_VoiceCardDlg*)whu_ParentDlg)->whu_User[3].Add(_T("普通用户"));
			CDialogEx::OnOK();
			AfxMessageBox(_T("注册成功，请等待管理员审批..."));
			
		}
		else if (m_Pwd01 == m_Pwd02 && m_Pwd02 ==_T(""))
		{
			AfxMessageBox(_T("密码不能为空"));
		}
		else{
			AfxMessageBox(_T("前后密码不一致"));
		}
	}
	
}
BOOL CRegisterDlg::OnInitDialog()
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
	return true;
}