// SaveEveryingDlg.cpp : implementation file
//

#include "stdafx.h"
#include "whu_VoiceCard.h"
#include "SaveEveryingDlg.h"
#include "afxdialogex.h"
#include "whu_VoiceCardDlg.h"

// CSaveEveryingDlg dialog

IMPLEMENT_DYNAMIC(CSaveEveryingDlg, CDialogEx)

CSaveEveryingDlg::CSaveEveryingDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(CSaveEveryingDlg::IDD, pParent)
{

}

CSaveEveryingDlg::~CSaveEveryingDlg()
{
}

void CSaveEveryingDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_PROGRESS1, m_Progress);
}


BEGIN_MESSAGE_MAP(CSaveEveryingDlg, CDialogEx)
	ON_WM_TIMER()
	ON_MESSAGE(WM_UPATEDATA, OnUpdateData)
END_MESSAGE_MAP()


// CSaveEveryingDlg message handlers

BOOL CSaveEveryingDlg::OnInitDialog()
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
	m_Progress.SetRange(0,1000);
	m_Progress.SetPos(500);
	SetTimer(1,100,NULL);
	return true;
}

void CSaveEveryingDlg::OnTimer(UINT_PTR nIDEvent)
{
	// TODO: Add your message handler code here and/or call default
	CDialogEx::OnTimer(nIDEvent);
}

LRESULT CSaveEveryingDlg::OnUpdateData(WPARAM wParam, LPARAM lParam)
{
	int pos = m_Progress.GetPos();
	if (pos<989)
	{
		pos += 10;
	}
	m_Progress.SetPos(pos);
	UpdateData(false);
	return true;
}