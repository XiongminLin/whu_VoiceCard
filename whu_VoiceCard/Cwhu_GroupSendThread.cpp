// Cwhu_GroupSendThread.cpp : implementation file
//

#include "stdafx.h"
//#include "whu_VoiceCard.h"
#include "whu_VoiceCardDlg.h"
#include "Cwhu_GroupSendThread.h"
// Cwhu_GroupSendThread
//extern int whu_NumHaveSend;
//extern CString whu_NumberBuf;
//extern CString whu_AllNos[MAX_PATH];
IMPLEMENT_DYNCREATE(Cwhu_GroupSendThread, CWinThread)

Cwhu_GroupSendThread::Cwhu_GroupSendThread()
{
}

Cwhu_GroupSendThread::~Cwhu_GroupSendThread()
{
}

BOOL Cwhu_GroupSendThread::InitInstance()
{
	// TODO:  perform and per-thread initialization here
	return TRUE;
}

int Cwhu_GroupSendThread::ExitInstance()
{
	// TODO:  perform any per-thread cleanup here
	return CWinThread::ExitInstance();
}
void Cwhu_GroupSendThread::whu_SendGroupNotice(UINT wParam,LONG lParam)
{

	
	int Num = lParam;
	int TrkCh = 2;
	whu_NumHaveSend = 0;
	while(whu_NumHaveSend < Num)
	{
		whu_NumberBuf = whu_AllNos[whu_NumHaveSend];
		::PostMessageW((HWND)(GetMainWnd()->GetSafeHwnd()),WM_USER+E_MEG_HAVENOTICETASK,TrkCh,NULL);
		Sleep(2000);
	}
}
void Cwhu_GroupSendThread::whu_SendGroupFax(UINT wParam,LONG lParam)
{
	int Num = lParam;
	int TrkCh = 2;
	whu_NumHaveSend = 0;
	while(whu_NumHaveSend < Num)
	{
		whu_NumberBuf = whu_AllNos[whu_NumHaveSend];
		::PostMessageW((HWND)(GetMainWnd()->GetSafeHwnd()),WM_USER+E_MSG_HAVEFAXTASK,TrkCh,NULL);
		Sleep(2000);
	}
}
BEGIN_MESSAGE_MAP(Cwhu_GroupSendThread, CWinThread)
	ON_THREAD_MESSAGE(E_MSG_HAVEGROUPNOTICE ,whu_SendGroupNotice)
	ON_THREAD_MESSAGE(E_MEG_HAVEGROUPFAX ,whu_SendGroupFax)
END_MESSAGE_MAP()


// Cwhu_GroupSendThread message handlers
