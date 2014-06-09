
// whu_VoiceCard.h : main header file for the PROJECT_NAME application
//

#pragma once

#ifndef __AFXWIN_H__
	#error "include 'stdafx.h' before including this file for PCH"
#endif

#include "resource.h"		// main symbols


// Cwhu_VoiceCardApp:
// See whu_VoiceCard.cpp for the implementation of this class
//

class Cwhu_VoiceCardApp : public CWinApp
{
public:
	Cwhu_VoiceCardApp();

// Overrides
public:
	virtual BOOL InitInstance();
	virtual BOOL ExitInstance();
// Implementation

	DECLARE_MESSAGE_MAP()
};

extern Cwhu_VoiceCardApp theApp;