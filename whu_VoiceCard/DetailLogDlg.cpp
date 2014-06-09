// DetailLogDlg.cpp : implementation file
//

#include "stdafx.h"
#include "whu_VoiceCard.h"
#include "DetailLogDlg.h"
#include "afxdialogex.h"
#include "Cwhu_HistoryDlg.h"

// CDetailLogDlg dialog

IMPLEMENT_DYNAMIC(CDetailLogDlg, CDialogEx)

CDetailLogDlg::CDetailLogDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(CDetailLogDlg::IDD, pParent)
{

}

CDetailLogDlg::~CDetailLogDlg()
{
}

void CDetailLogDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}


BEGIN_MESSAGE_MAP(CDetailLogDlg, CDialogEx)
	ON_BN_CLICKED(IDCANCEL, &CDetailLogDlg::OnBnClickedCancel)
	ON_BN_CLICKED(IDC_LOGOPENFILE, &CDetailLogDlg::OnBnClickedLogopenfile)
	ON_BN_CLICKED(IDC_BUTTON1, &CDetailLogDlg::OnBnClickedButton1)
	ON_BN_CLICKED(IDOK, &CDetailLogDlg::OnBnClickedOk)
END_MESSAGE_MAP()


// CDetailLogDlg message handlers


void CDetailLogDlg::OnBnClickedCancel()
{
	// TODO: Add your control notification handler code here
	CDialogEx::OnCancel();
}
BOOL CDetailLogDlg::OnInitDialog()
{
	//接收来自父对话框的数据并显示出来//
	//将对话框移至屏幕中间//
	CDialogEx::OnInitDialog();
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
	//加载历史记录数据//
	whu_ShowData();
	return true;
}
void CDetailLogDlg::whu_ShowData()
{
	m_ClickPos = ((Cwhu_HistoryDlg*)whu_ParentDlg)->m_ClickPos;
	CString m_Str = ((Cwhu_HistoryDlg*)whu_ParentDlg)->m_HistoryList.GetItemText(m_ClickPos,0);
	SetDlgItemText(IDC_LOGNUMBER,m_Str);
	m_Str = ((Cwhu_HistoryDlg*)whu_ParentDlg)->m_HistoryList.GetItemText(m_ClickPos,1);
	SetDlgItemText(IDC_LOGTIME,m_Str);
	m_Str = ((Cwhu_HistoryDlg*)whu_ParentDlg)->m_HistoryList.GetItemText(m_ClickPos,2);
	SetDlgItemText(IDC_LOGDUTY,m_Str);
	m_Str = ((Cwhu_HistoryDlg*)whu_ParentDlg)->m_HistoryList.GetItemText(m_ClickPos,3);
	SetDlgItemText(IDC_LOGEVENT,m_Str);
	m_Str = ((Cwhu_HistoryDlg*)whu_ParentDlg)->m_HistoryList.GetItemText(m_ClickPos,4);
	SetDlgItemText(IDC_LOGUNIT,m_Str);
	m_Str = ((Cwhu_HistoryDlg*)whu_ParentDlg)->m_HistoryList.GetItemText(m_ClickPos,5);
	SetDlgItemText(IDC_LOGNAME,m_Str);
	m_Str = ((Cwhu_HistoryDlg*)whu_ParentDlg)->m_HistoryList.GetItemText(m_ClickPos,6);
	SetDlgItemText(IDC_LOGNUM,m_Str);
	m_Str = ((Cwhu_HistoryDlg*)whu_ParentDlg)->m_HistoryList.GetItemText(m_ClickPos,7);
	SetDlgItemText(IDC_LOGFILE,m_Str);
	m_Str = ((Cwhu_HistoryDlg*)whu_ParentDlg)->m_HistoryList.GetItemText(m_ClickPos,8);
	SetDlgItemText(IDC_LOGSTATE,m_Str);
	m_Edit = ((Cwhu_HistoryDlg*)whu_ParentDlg)->m_HistoryList.GetItemText(m_ClickPos,9);
	SetDlgItemText(IDC_LOGEDIT,m_Edit);
	

}
void CDetailLogDlg::OnBnClickedLogopenfile()
{
	// TODO: Add your control notification handler code here
	CString m_Str;
	GetDlgItemText(IDC_LOGFILE,m_Str);
	ShellExecute(NULL, "open",m_Str,NULL,NULL, SW_SHOWNORMAL );
}


void CDetailLogDlg::OnBnClickedButton1() //重置
{
	// TODO: Add your control notification handler code here
	SetDlgItemText(IDC_LOGEDIT,m_Edit);
	
}


void CDetailLogDlg::OnBnClickedOk()
{
	// TODO: Add your control notification handler code here
	GetDlgItemText(IDC_LOGEDIT,m_Edit);
	((Cwhu_HistoryDlg*)whu_ParentDlg)->m_HistoryList.SetItemText(m_ClickPos,9,m_Edit);
	CDialogEx::OnOK();
}
