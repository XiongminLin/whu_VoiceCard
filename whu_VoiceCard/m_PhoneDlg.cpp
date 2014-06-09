// m_PhoneDlg.cpp : implementation file
//

#include "stdafx.h"
#include "whu_VoiceCard.h"
#include "m_PhoneDlg.h"
#include "afxdialogex.h"


// m_PhoneDlg dialog

IMPLEMENT_DYNAMIC(m_PhoneDlg, CDialogEx)

m_PhoneDlg::m_PhoneDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(m_PhoneDlg::IDD, pParent)
{

}

m_PhoneDlg::~m_PhoneDlg()
{
}

void m_PhoneDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}


BEGIN_MESSAGE_MAP(m_PhoneDlg, CDialogEx)
END_MESSAGE_MAP()


// m_PhoneDlg message handlers
