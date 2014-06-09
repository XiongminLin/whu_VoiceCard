#pragma once
#include "stdafx.h"
// Cwhu_GroupSendThread

class Cwhu_GroupSendThread : public CWinThread
{
	DECLARE_DYNCREATE(Cwhu_GroupSendThread)

public:
	Cwhu_GroupSendThread();           // protected constructor used by dynamic creation
	virtual ~Cwhu_GroupSendThread();

public:
	virtual BOOL InitInstance();
	virtual int ExitInstance();

protected:
	afx_msg void whu_SendGroupNotice(UINT wParam,LONG lParam);
	afx_msg void whu_SendGroupFax(UINT wParam,LONG lParam);
	DECLARE_MESSAGE_MAP()
};


