// Cwhu_RecvCallDlg.cpp : implementation file
//

#include "stdafx.h"
#include "whu_VoiceCard.h"
#include "Cwhu_RecvCallDlg.h"
#include "afxdialogex.h"
#include "math.h"
#include "mmsystem.h"//导入声音头文件//

#pragma comment(lib,"winmm.lib")//导入声音头文件库//

#define PI  3.1415926535

// Cwhu_RecvCallDlg dialog

IMPLEMENT_DYNAMIC(Cwhu_RecvCallDlg, CDialogEx)

DWORD WINAPI threadFunc(LPVOID threadNum)
{
	HMODULE hmod=AfxGetResourceHandle(); 
	HRSRC hSndResource=FindResource(hmod,MAKEINTRESOURCE(IDR_WAVE1),_T("WAVE"));
	HGLOBAL hGlobalMem=LoadResource(hmod,hSndResource);
	LPCTSTR lpMemSound=(LPCSTR)LockResource(hGlobalMem);
	sndPlaySound(lpMemSound,SND_MEMORY);
	FreeResource(hGlobalMem);
	return 0;
}
Cwhu_RecvCallDlg::Cwhu_RecvCallDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(Cwhu_RecvCallDlg::IDD, pParent)
{

}

Cwhu_RecvCallDlg::~Cwhu_RecvCallDlg()
{
}

void Cwhu_RecvCallDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}
BOOL Cwhu_RecvCallDlg::OnInitDialog()
{
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

	whu_PickUp = false;
	SetTimer(2,2000,NULL); //设置响铃间隔时间
	SetTimer(3,10000,NULL); //设置响铃时间
	whu_ShakeWindow();
	return true;
}

BEGIN_MESSAGE_MAP(Cwhu_RecvCallDlg, CDialogEx)
	ON_WM_TIMER()
	ON_BN_CLICKED(IDOK, &Cwhu_RecvCallDlg::OnBnClickedOk)
END_MESSAGE_MAP()


// Cwhu_RecvCallDlg message handlers


void Cwhu_RecvCallDlg::OnTimer(UINT_PTR nIDEvent)
{
	// TODO: Add your message handler code here and/or call default
	if (nIDEvent==1)
	{
		CRect rect;
		GetClientRect(rect);
		int srX=rect.right-rect.left;
		int srY=rect.bottom-rect.top;

		int x,y;

		int n=8;		//直径
		int pos=3;		//步长

		if (Time<0)
		{
			Time++;
			return;
		}

		int m=Time%n;
		if((Time-Time%n)%2==0)
		{
			x=X+pos*m;
			y=(int)(Y-pos*sin(m*PI/n));
			if(m==0)
			{
				X=x;
				Y=y;
			}

		}
		else
		{
			x=X-pos*m;
			y=(int)(Y+3*sin(m*PI/n));
			if(m==0)
			{
				X=x;
				Y=y;
			}
		}
		Time++;

		SetWindowPos(NULL, x,y,srX,srY, SWP_NOMOVE||SWP_NOSIZE);

		if (Time==4*n+1)
		{
			KillTimer(1);
			//		TerminateThread(hThread,1);
		}

	}else if (nIDEvent == 2 && whu_PickUp == false) //每隔几秒都一次窗口//
	{
		whu_ShakeWindow();
	}
	else if (nIDEvent == 3)
	{
		whu_PickUp = true;
	}
	CDialogEx::OnTimer(nIDEvent);
}


void Cwhu_RecvCallDlg::whu_ShakeWindow()
{
	CRect rect1;
	GetWindowRect(rect1);
	X=rect1.left;
	Y=rect1.top;

	Time=-5;
	DWORD threadID;
	hThread=CreateThread(NULL,0,threadFunc,NULL,0,&threadID); //播放声音//

	SetTimer(1,5,NULL);
}

void Cwhu_RecvCallDlg::OnBnClickedOk()
{
	// TODO: Add your control notification handler code here
	whu_PickUp = true;
	//CDialogEx::OnOK();
}
