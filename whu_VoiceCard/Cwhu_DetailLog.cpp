// Cwhu_DetailLog.cpp : implementation file
//

#include "stdafx.h"
#include "whu_VoiceCard.h"
#include "DetailLogDlg.h"
#include "afxdialogex.h"


// Cwhu_DetailLog dialog

IMPLEMENT_DYNAMIC(Cwhu_DetailLog, CDialogEx)

Cwhu_DetailLog::Cwhu_DetailLog(CWnd* pParent /*=NULL*/)
	: CDialogEx(Cwhu_DetailLog::IDD, pParent)
{

}

Cwhu_DetailLog::~Cwhu_DetailLog()
{
}

void Cwhu_DetailLog::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}
BOOL Cwhu_DetailLog::OnInitDialog()
{
	//�������Ը��Ի�������ݲ���ʾ����//
	//���Ի���������Ļ�м�//
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

BEGIN_MESSAGE_MAP(Cwhu_DetailLog, CDialogEx)
END_MESSAGE_MAP()


// Cwhu_DetailLog message handlers
