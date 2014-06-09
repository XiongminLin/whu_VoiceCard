// Whu_AutoForDlg.cpp : implementation file
//

#include "stdafx.h"
#include "whu_VoiceCard.h"
#include "Whu_AutoForDlg.h"
#include "Cwhu_FaxSettingDlg.h"
#include "afxdialogex.h"


// CWhu_AutoForDlg dialog

IMPLEMENT_DYNAMIC(CWhu_AutoForDlg, CDialogEx)

CWhu_AutoForDlg::CWhu_AutoForDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(CWhu_AutoForDlg::IDD, pParent)
{

}

CWhu_AutoForDlg::~CWhu_AutoForDlg()
{
}

void CWhu_AutoForDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_LIST1, m_List);
	DDX_Control(pDX, IDC_LIST2, m_SorList);
}


BEGIN_MESSAGE_MAP(CWhu_AutoForDlg, CDialogEx)
	ON_BN_CLICKED(IDOK, &CWhu_AutoForDlg::OnBnClickedOk)
END_MESSAGE_MAP()


// CWhu_AutoForDlg message handlers


void CWhu_AutoForDlg::OnBnClickedOk()
{
	// TODO: Add your control notification handler code here
	int num = m_List.GetItemCount();
	((Cwhu_FaxSettingDlg*)whu_ParentDlg)->m_ListTo.DeleteAllItems();
	int ID = 0;
	CStringArray m_AddArr;
	m_AddArr.Add(m_SorList.GetItemText(0,7));
	for (int i=0;i<num;i++)
	{
		if(m_List.GetCheck(i)==true)
		{
			char buf[3];
			itoa(ID+1,buf,10);
			((Cwhu_FaxSettingDlg*)whu_ParentDlg)->m_ListTo.InsertItem(i,buf);
			((Cwhu_FaxSettingDlg*)whu_ParentDlg)->m_ListTo.SetItemText(ID,0,buf);
			for (int j =1;j<9;j++)
			{
				CString m_str = m_List.GetItemText(i,j);
				((Cwhu_FaxSettingDlg*)whu_ParentDlg)->m_ListTo.SetItemText(ID,j,m_str);
			}
			m_AddArr.Add(m_List.GetItemText(i,7));
			ID++;
		}
	}
	CString NumCount;
	NumCount.Format(_T("%d"),ID);
	m_AddArr.InsertAt(1,NumCount);
	whu_AddAutoForArr(((Cwhu_FaxSettingDlg*)whu_ParentDlg)->whu_AutoForward,m_AddArr);
	((Cwhu_FaxSettingDlg*)whu_ParentDlg)->m_ListFrom.SetFocus();
	((Cwhu_FaxSettingDlg*)whu_ParentDlg)->m_ListFrom.SetItemState(m_ClickPos,LVIS_SELECTED,LVIS_SELECTED);
	CDialogEx::OnOK();
}
BOOL CWhu_AutoForDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();
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
	m_ClickPos = ((Cwhu_FaxSettingDlg*)whu_ParentDlg)->m_ClickPos;
	whu_InitList();

	return true;
}
void CWhu_AutoForDlg::whu_InitList()
{
	//��ϵ���б��ʼ��//
	LONG lStyle = GetWindowLong(m_SorList.m_hWnd, GWL_STYLE);
	lStyle &= ~LVS_TYPEMASK;
	lStyle |= LVS_REPORT; 
	SetWindowLong(m_SorList.m_hWnd, GWL_STYLE, lStyle);
	DWORD dwStyle = m_SorList.GetExtendedStyle(); 
	dwStyle |= LVS_EX_FULLROWSELECT;
	dwStyle |= LVS_EX_GRIDLINES;
	dwStyle |= LVS_EX_CHECKBOXES;
	m_SorList.SetExtendedStyle(dwStyle); 
	m_SorList.InsertColumn( 0, _T("���"), LVCFMT_LEFT, 60 );
	m_SorList.InsertColumn( 1, _T("����"), LVCFMT_LEFT, 180 ); 
	m_SorList.InsertColumn( 2, _T("��λ"), LVCFMT_LEFT, 180 );
	m_SorList.InsertColumn( 3, _T("����"), LVCFMT_LEFT, 180 );
	m_SorList.InsertColumn( 4, _T("ְ��"), LVCFMT_LEFT, 120 );
	m_SorList.InsertColumn( 5, _T("�ֻ�"), LVCFMT_LEFT, 0 ); 
	m_SorList.InsertColumn( 6, _T("�칫�ҵ绰"), LVCFMT_LEFT, 0 ); 
	m_SorList.InsertColumn( 7, _T("�����"), LVCFMT_LEFT, 180 ); 
	m_SorList.InsertColumn( 8, _T("����"), LVCFMT_LEFT, 0 );
	char buf[3];
	itoa(1,buf,10);
	m_SorList.InsertItem(0,buf);
	m_SorList.SetItemText(0,0,_T("1"));
	m_ClickPos = ((Cwhu_FaxSettingDlg*)whu_ParentDlg)->m_ClickPos;
	for (int j =1;j<9;j++)
	{
		CString m_str = ((Cwhu_FaxSettingDlg*)whu_ParentDlg)->m_ListFrom.GetItemText(m_ClickPos,j);
		m_SorList.SetItemText(0,j,m_str);
	}

	//��ϵ���б��ʼ��//
	lStyle = GetWindowLong(m_List.m_hWnd, GWL_STYLE);
	lStyle &= ~LVS_TYPEMASK; 
	lStyle |= LVS_REPORT; 
	SetWindowLong(m_List.m_hWnd, GWL_STYLE, lStyle);
	dwStyle = m_List.GetExtendedStyle(); 
	dwStyle |= LVS_EX_FULLROWSELECT;
	dwStyle |= LVS_EX_GRIDLINES;
	dwStyle |= LVS_EX_CHECKBOXES;
	m_List.SetExtendedStyle(dwStyle); 
	m_List.InsertColumn( 0, _T("���"), LVCFMT_LEFT, 60 );
	m_List.InsertColumn( 1, _T("����"), LVCFMT_LEFT, 180 ); 
	m_List.InsertColumn( 2, _T("��λ"), LVCFMT_LEFT, 180 );
	m_List.InsertColumn( 3, _T("����"), LVCFMT_LEFT, 180 );
	m_List.InsertColumn( 4, _T("ְ��"), LVCFMT_LEFT, 120 );
	m_List.InsertColumn( 5, _T("�ֻ�"), LVCFMT_LEFT, 0 ); 
	m_List.InsertColumn( 6, _T("�칫�ҵ绰"), LVCFMT_LEFT, 0 ); 
	m_List.InsertColumn( 7, _T("�����"), LVCFMT_LEFT, 180 ); 
	m_List.InsertColumn( 8, _T("����"), LVCFMT_LEFT, 0 );

	//ʵ�ڲ�֪����ʲôԭ��//���ʲ��˸��Ի���Ĺ��б���//��������������m_FaxContactList������Ϊpublic:������������������֪��why��//
	int num = ((Cwhu_FaxSettingDlg*)whu_ParentDlg)->m_ListFrom.GetItemCount();
	for (int i=0;i<num;i++)
	{
		char buf[3];
		itoa(i+1,buf,10);
		m_List.InsertItem(i,buf);
		for (int j =0;j<9;j++)
		{
			CString m_str = ((Cwhu_FaxSettingDlg*)whu_ParentDlg)->m_ListFrom.GetItemText(i,j);
			m_List.SetItemText(i,j,m_str);
		}
	}
	////����ת�����Լ�//
	//CString m_name = m_SorList.GetItemText(0,1);
	//for (int i=m_List.GetItemCount();i>=0;i--)
	//{
	//	CString m_str = m_List.GetItemText(i,1);
	//	if (m_str == m_name)
	//	{
	//		m_List.DeleteItem(i);
	//		break;
	//	}
	//}
	//��������//
	for (int i=0;i<m_List.GetItemCount();i++)
	{
		CString m_Str;
		m_Str.Format(_T("%d"),i+1);
		m_List.SetItemText(i,0,m_Str);
	}
	//ԭ��ѡ�ϵģ����ø�ѡ��Ϊѡ��//
	for (int i=0;i<((Cwhu_FaxSettingDlg*)whu_ParentDlg)->m_ListTo.GetItemCount();i++)
	{
		CString m_Str = ((Cwhu_FaxSettingDlg*)whu_ParentDlg)->m_ListTo.GetItemText(i,1);
		for (int j=0;j<m_List.GetItemCount();j++)
		{
			CString m_Str2 = m_List.GetItemText(j,1);
			if (m_Str == m_Str2)
			{
				m_List.SetCheck(j,true);
			}
		}
	}

}
void CWhu_AutoForDlg::whu_AddAutoForArr(CStringArray &m_SorArr,CStringArray &m_AddArr)
{
	int size = m_SorArr.GetSize();
	CString m_FaxNum = m_AddArr.GetAt(0);
	int NumCount = 0;
	for (int i=0;i<size;i++)
	{
		CString m_str = m_SorArr.GetAt(i);
		if ((m_str == m_FaxNum)&&(i < size-1))
		{
			CString m_str2 = m_SorArr.GetAt(i+1);
			int CharCount = m_str2.GetLength(); //�ַ������鳤����������ģ�˵����������ת����������ı�־λ��////
			if(CharCount<4)
			{
				int NumCount = _ttoi(m_str2);
				m_SorArr.RemoveAt(i,NumCount+2);
				break;
			}
			
		}	
	}
	int m_AddSize = m_AddArr.GetSize();
	if (m_AddSize>2)
	{
		m_SorArr.Append(m_AddArr);
	}

}