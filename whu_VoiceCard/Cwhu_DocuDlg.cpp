// Cwhu_DocuDlg.cpp : implementation file
//

#include "stdafx.h"
#include "whu_VoiceCard.h"
#include "Cwhu_DocuDlg.h"
#include "afxdialogex.h"
#include "whu_VoiceCardDlg.h"

// Cwhu_DocuDlg dialog

IMPLEMENT_DYNAMIC(Cwhu_DocuDlg, CDialogEx)

Cwhu_DocuDlg::Cwhu_DocuDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(Cwhu_DocuDlg::IDD, pParent)
	, m_info(_T(""))
{

}

Cwhu_DocuDlg::~Cwhu_DocuDlg()
{
}

void Cwhu_DocuDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_LIST1, m_FaxList);
	DDX_Text(pDX, IDC_INFO, m_info);
	DDX_Control(pDX, IDC_LISTNOTICE, m_ListNotice);
}
void Cwhu_DocuDlg::whu_InitList()
{
	LONG lStyle; 
	lStyle = GetWindowLong(m_FaxList.m_hWnd, GWL_STYLE);//��ȡ��ǰ����style 
		lStyle &= ~LVS_TYPEMASK; //�����ʾ��ʽλ 
	lStyle |= LVS_REPORT; //����style 
	SetWindowLong(m_FaxList.m_hWnd, GWL_STYLE, lStyle);//����style 
	DWORD dwStyle = m_FaxList.GetExtendedStyle(); 
	dwStyle |= LVS_EX_FULLROWSELECT;//ѡ��ĳ��ʹ���и�����ֻ������report����listctrl�� 
		dwStyle |= LVS_EX_GRIDLINES;//�����ߣ�ֻ������report����listctrl�� 
		dwStyle |= LVS_EX_CHECKBOXES;//itemǰ����checkbox�ؼ� 
	m_FaxList.SetExtendedStyle(dwStyle); //������չ��� 

	m_FaxList.InsertColumn( 0, "ID", LVCFMT_LEFT, 40 );
	m_FaxList.InsertColumn( 1, _T("����"), LVCFMT_LEFT, 120 ); 
	m_FaxList.InsertColumn( 2, _T("ʱ��"), LVCFMT_LEFT, 80 );
	m_FaxList.InsertColumn( 3, _T("���"), LVCFMT_LEFT, 50 );
	m_FaxList.InsertColumn( 4, _T("����"), LVCFMT_LEFT, 140 );
	m_FaxList.InsertColumn( 5, _T("��ע"), LVCFMT_LEFT, 160 ); 
	int nRow = m_FaxList.InsertItem(0, "1");//������ 
	m_FaxList.InsertItem(1, "2");//������ 
	m_FaxList.InsertItem(2, "3");//������ 
	for (int i=0;i<3;i++)
	{
		m_FaxList.SetItemText(i, 1, _T("faxfile.tif"));//�������� 
		m_FaxList.SetItemText(i, 2, _T("2013.04.13"));//�������� 
		m_FaxList.SetItemText(i, 3, _T("send"));//�������� 
		m_FaxList.SetItemText(i, 4, _T("68756269"));//�������� 
		m_FaxList.SetItemText(i, 5, _T(""));//�������� 
	}



	lStyle = GetWindowLong(m_ListNotice.m_hWnd, GWL_STYLE);//��ȡ��ǰ����style 
	lStyle &= ~LVS_TYPEMASK; //�����ʾ��ʽλ 
	lStyle |= LVS_REPORT; //����style 
	SetWindowLong(m_ListNotice.m_hWnd, GWL_STYLE, lStyle);//����style 
	dwStyle = m_ListNotice.GetExtendedStyle(); 
	dwStyle |= LVS_EX_FULLROWSELECT;//ѡ��ĳ��ʹ���и�����ֻ������report����listctrl�� 
	dwStyle |= LVS_EX_GRIDLINES;//�����ߣ�ֻ������report����listctrl�� 
	dwStyle |= LVS_EX_CHECKBOXES;//itemǰ����checkbox�ؼ� 
	m_ListNotice.SetExtendedStyle(dwStyle); //������չ��� 

	m_ListNotice.InsertColumn( 0, "ID", LVCFMT_LEFT, 40 );
	m_ListNotice.InsertColumn( 1, _T("����"), LVCFMT_LEFT, 120 ); 
	m_ListNotice.InsertColumn( 2, _T("ʱ��"), LVCFMT_LEFT, 80 );
	m_ListNotice.InsertColumn( 3, _T("���"), LVCFMT_LEFT, 50 );
	m_ListNotice.InsertColumn( 4, _T("����"), LVCFMT_LEFT, 140 );
	m_ListNotice.InsertColumn( 5, _T("��ע"), LVCFMT_LEFT, 160 ); 
	 nRow = m_ListNotice.InsertItem(0, "1");//������ 
	m_ListNotice.InsertItem(1, "2");//������ 
	m_ListNotice.InsertItem(2, "3");//������ 
	for (int i=0;i<3;i++)
	{
		m_ListNotice.SetItemText(i, 1, _T("Notice.wav"));//�������� 
		m_ListNotice.SetItemText(i, 2, _T("2013.04.13"));//�������� 
		m_ListNotice.SetItemText(i, 3, _T("send"));//�������� 
		m_ListNotice.SetItemText(i, 4, _T("13349911910"));//�������� 
		m_ListNotice.SetItemText(i, 5, _T(""));//�������� 
	}

	GetDlgItem(IDC_INFO)->ShowWindow(false);
}

BEGIN_MESSAGE_MAP(Cwhu_DocuDlg, CDialogEx)
	ON_BN_CLICKED(IDC_BADDINFO, &Cwhu_DocuDlg::OnBnClickedBaddinfo)
	ON_BN_CLICKED(IDC_BSAVEDOC, &Cwhu_DocuDlg::OnBnClickedBsavedoc)
END_MESSAGE_MAP()


// Cwhu_DocuDlg message handlers


void Cwhu_DocuDlg::OnBnClickedBaddinfo()
{
	// TODO: Add your control notification handler code here
	GetDlgItem(IDC_INFO)->ShowWindow(true);
}


void Cwhu_DocuDlg::OnBnClickedBsavedoc()
{
	// TODO: Add your control notification handler code here

	UpdateData(true);
	m_FaxList.SetExtendedStyle(LVS_EX_CHECKBOXES); 
	CString str; 
	for(int i=0; i<m_FaxList.GetItemCount(); i++) 
	{ 
		if( m_FaxList.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED || m_FaxList.GetCheck(i)) 
		{ 
			m_FaxList.SetItemText(i, 5, m_info);//�������� 
			break;
		} 
	} 
	m_ListNotice.SetExtendedStyle(LVS_EX_CHECKBOXES);
	for(int i=0; i<m_ListNotice.GetItemCount(); i++) 
	{ 
		if( m_ListNotice.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED || m_ListNotice.GetCheck(i)) 
		{ 
			m_ListNotice.SetItemText(i, 5, m_info);//�������� 
			break;
		} 
	} 
	m_info = _T("");
	UpdateData(false);
}
