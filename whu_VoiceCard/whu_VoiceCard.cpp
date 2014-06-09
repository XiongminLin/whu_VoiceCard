
// whu_VoiceCard.cpp : Defines the class behaviors for the application.
//

#include "stdafx.h"
#include "whu_VoiceCard.h"
#include "whu_VoiceCardDlg.h"
#include "CDataGrid.h"
#include "CAdodc.h"
#pragma comment(lib,"SkinPPWTL.lib")
#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// Cwhu_VoiceCardApp

BEGIN_MESSAGE_MAP(Cwhu_VoiceCardApp, CWinApp)
	ON_COMMAND(ID_HELP, &CWinApp::OnHelp)
END_MESSAGE_MAP()


// Cwhu_VoiceCardApp construction

Cwhu_VoiceCardApp::Cwhu_VoiceCardApp()
{
	// support Restart Manager
	m_dwRestartManagerSupportFlags = AFX_RESTART_MANAGER_SUPPORT_RESTART;

	// TODO: add construction code here,
	// Place all significant initialization in InitInstance
}


// The one and only Cwhu_VoiceCardApp object

Cwhu_VoiceCardApp theApp;


// Cwhu_VoiceCardApp initialization

BOOL Cwhu_VoiceCardApp::InitInstance()
{
	// InitCommonControlsEx() is required on Windows XP if an application
	// manifest specifies use of ComCtl32.dll version 6 or later to enable
	// visual styles.  Otherwise, any window creation will fail.
	AfxOleInit();
	INITCOMMONCONTROLSEX InitCtrls;
	InitCtrls.dwSize = sizeof(InitCtrls);
	// Set this to include all the common control classes you want to use
	// in your application.
	InitCtrls.dwICC = ICC_WIN95_CLASSES;

	

	InitCommonControlsEx(&InitCtrls);

	CWinApp::InitInstance();

	LoadLibrary("RICHED20.DLL") ;

	AfxInitRichEdit2();

	skinppLoadSkin(_T("spring.ssk")); //很好看，但颜色过于喜庆，可能不适合90//
	//skinppLoadSkin(_T("AquaOS.ssk"));//还凑合60
	//skinppLoadSkin(_T("MAC.ssk"));//苹果的皮肤，可以80//
	//skinppLoadSkin(_T("AlphaOS.ssk")); //太难看了30//
	//skinppLoadSkin(_T("Aura.ssk")); //还凑合60
	//skinppLoadSkin(_T("avfone.ssk"));//还凑合60
	//skinppLoadSkin(_T("blue.ssk"));//还凑合60
	//skinppLoadSkin(_T("bOzen.ssk"));//还凑合60
	//skinppLoadSkin(_T("DameK UltraBlue.ssk"));//70
	//skinppLoadSkin(_T("default.ssk"));//70
	//skinppLoadSkin(_T("dogmax.ssk"));//65
	//skinppLoadSkin(_T("Gloss.ssk"));//65
	//skinppLoadSkin(_T("Longhorn.ssk"));//65
	//skinppLoadSkin(_T("Noire.ssk"));//65
	//skinppLoadSkin(_T("Phenom.ssk"));//65
	//skinppLoadSkin(_T("Royale.ssk"));//65
	//skinppLoadSkin(_T("SlickOS2.ssk"));//65
	//skinppLoadSkin(_T("Steel.ssk"));//65
	//skinppLoadSkin(_T("UMskin.ssk"));//65
	//skinppLoadSkin(_T("Vista.ssk"));//65
	//skinppLoadSkin(_T("vladstudio.ssk"));//65
	//skinppLoadSkin(_T("xp_corona.ssk"));//65
	//skinppLoadSkin(_T("XP-Luna.ssk"));//65








	AfxEnableControlContainer();

	// Create the shell manager, in case the dialog contains
	// any shell tree view or shell list view controls.
	CShellManager *pShellManager = new CShellManager;

	// Standard initialization
	// If you are not using these features and wish to reduce the size
	// of your final executable, you should remove from the following
	// the specific initialization routines you do not need
	// Change the registry key under which our settings are stored
	// TODO: You should modify this string to be something appropriate
	// such as the name of your company or organization
	SetRegistryKey(_T("Local AppWizard-Generated Applications"));

	Cwhu_VoiceCardDlg dlg;
	m_pMainWnd = &dlg;
	INT_PTR nResponse = dlg.DoModal();
	if (nResponse == IDOK)
	{
		// TODO: Place code here to handle when the dialog is
		//  dismissed with OK
	}
	else if (nResponse == IDCANCEL)
	{
		// TODO: Place code here to handle when the dialog is
		//  dismissed with Cancel
	}



	// Delete the shell manager created above.
	if (pShellManager != NULL)
	{
		delete pShellManager;
	}

	// Since the dialog has been closed, return FALSE so that we exit the
	//  application, rather than start the application's message pump.
	return FALSE;
}

BOOL Cwhu_VoiceCardApp::ExitInstance()
{
	skinppExitSkin();     
	return CWinApp::ExitInstance(); 
}