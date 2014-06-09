// Cwhu_PhoneDlg.cpp : implementation file
//

#include "stdafx.h"
#include "whu_VoiceCard.h"
#include "Cwhu_PhoneDlg.h"
#include "afxdialogex.h"
#include "whu_VoiceCardDlg.h"
#include "Mmsystem.h"
#include  <Windows.h>
// Cwhu_PhoneDlg dialog

IMPLEMENT_DYNAMIC(Cwhu_PhoneDlg, CDialogEx)

Cwhu_PhoneDlg::Cwhu_PhoneDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(Cwhu_PhoneDlg::IDD, pParent)
	, m_HandInputNum(_T(""))
{
}

Cwhu_PhoneDlg::~Cwhu_PhoneDlg()
{
}

void Cwhu_PhoneDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);

	DDX_Text(pDX, IDC_NUMBER, m_HandInputNum);
	DDX_Control(pDX, IDC_CHECKNUMLEN, m_CheckNumLen);
	DDX_Control(pDX, IDC_LISTCONTACT02, m_ContactList);
}


BEGIN_MESSAGE_MAP(Cwhu_PhoneDlg, CDialogEx)
	ON_BN_CLICKED(IDC_RECORDSTART, &Cwhu_PhoneDlg::OnBnClickedRecordstart)
	ON_BN_CLICKED(IDC_SENDNOTICE, &Cwhu_PhoneDlg::OnBnClickedSendnotice)
	ON_BN_CLICKED(IDC_LISTEN, &Cwhu_PhoneDlg::OnBnClickedListen)
	ON_BN_CLICKED(IDC_CHECKNUMLEN, &Cwhu_PhoneDlg::OnBnClickedChecknumlen)
	ON_BN_CLICKED(IDC_BUTTONBACK, &Cwhu_PhoneDlg::OnBnClickedButtonback)
END_MESSAGE_MAP()


// Cwhu_PhoneDlg message handlers


void Cwhu_PhoneDlg::OnBnClickedRecordstart()
{
	// TODO: Add your control notification handler code here

	((Cwhu_VoiceCardDlg*)whu_ParentDlg)->whu_RecordNotice = true;
	
	SetDlgItemText(IDC_NOTICE,_T("请拿起电话听筒，开始录音，挂电话录音结束..."));                                                                                                                                                                                                                                                                                                                                        

}


void Cwhu_PhoneDlg::OnBnClickedSendnotice()
{
	// TODO: Add your control notification handler code here
	UpdateData(true);
	CString m_CallNum[MAX_PATH] = {_T("")};
	int nu = SeparateCallNum(m_HandInputNum,m_CallNum);
	((Cwhu_VoiceCardDlg*)whu_ParentDlg)->whu_SendNotice(m_CallNum,nu);
	//
}
int  Cwhu_PhoneDlg::SeparateCallNum(CString m_AllNum,CString *m_SeparateNum)
{

	int m_len = m_AllNum.GetLength();
	int pos=-2;
	int n=-1;
	CString m_str=_T("");
	while (m_len>2 )
	{
		pos=m_AllNum.Find(";");
		n++;
		if (pos == -1)
		{
			m_SeparateNum[n]=m_AllNum;
			return n+1;
		}
		m_SeparateNum[n]=m_AllNum.Left(pos);
		m_str=m_AllNum.Right(m_len-pos-1);
		m_AllNum=m_str;
		m_len = m_AllNum.GetLength();
		pos=m_AllNum.Find(";");
		if (pos == -1)
		{
			m_SeparateNum[n+1]=m_AllNum;
			return n+2;
		}
	}

}

void Cwhu_PhoneDlg::OnBnClickedListen()
{
	// TODO: Add your control notification handler code here
	bool m_bRecordNotice=((Cwhu_VoiceCardDlg*)whu_ParentDlg)->whu_RecordNotice;
	CString m_filename = ((Cwhu_VoiceCardDlg*)whu_ParentDlg)->whu_RecordFile;
	if (m_bRecordNotice==false&&m_filename!=_T(""))
	{
		sndPlaySound(m_filename,SND_NOSTOP);
	} 

}
BEGIN_EVENTSINK_MAP(Cwhu_PhoneDlg, CDialogEx)
END_EVENTSINK_MAP()






void Cwhu_PhoneDlg::OnBnClickedChecknumlen()
{
	// TODO: Add your control notification handler code here
	int state = m_CheckNumLen.GetCheck();
	if (state == 0)
	{
		((Cwhu_VoiceCardDlg*)whu_ParentDlg)->whu_NumLength = 8;
	} 
	else if(state == 1)
	{
		((Cwhu_VoiceCardDlg*)whu_ParentDlg)->whu_NumLength = 11;
	}

}
BOOL Cwhu_PhoneDlg::OnInitDialog()
{
	whu_IntiList();
	return true;
}

void Cwhu_PhoneDlg::OnBnClickedButtonback()
{
	// TODO: Add your control notification handler code here
	CDialogEx::OnCancel();
}
void Cwhu_PhoneDlg::whu_IntiList()
{
	//联系人列表初始化
	LONG lStyle; 
	lStyle = GetWindowLong(m_ContactList.m_hWnd, GWL_STYLE);//获取当前窗口style 
	lStyle &= ~LVS_TYPEMASK; //清除显示方式位 
	lStyle |= LVS_REPORT; //设置style 
	SetWindowLong(m_ContactList.m_hWnd, GWL_STYLE, lStyle);//设置style 
	DWORD dwStyle = m_ContactList.GetExtendedStyle();  //这里崩溃了//
	dwStyle |= LVS_EX_FULLROWSELECT;//选中某行使整行高亮（只适用与report风格的listctrl） //
	dwStyle |= LVS_EX_GRIDLINES;//网格线（只适用与report风格的listctrl） //
	dwStyle |= LVS_EX_CHECKBOXES;//item前生成checkbox控件 //
	m_ContactList.SetExtendedStyle(dwStyle); //设置扩展风格 //


	m_ContactList.InsertColumn( 0, _T("编号"), LVCFMT_LEFT, 60 );
	m_ContactList.InsertColumn( 1, _T("姓名"), LVCFMT_LEFT, 120 ); 
	m_ContactList.InsertColumn( 2, _T("单位"), LVCFMT_LEFT, 120 );
	m_ContactList.InsertColumn( 3, _T("部门"), LVCFMT_LEFT, 160 );
	m_ContactList.InsertColumn( 4, _T("职务"), LVCFMT_LEFT, 80 );
	m_ContactList.InsertColumn( 5, _T("手机"), LVCFMT_LEFT, 160 ); 
	m_ContactList.InsertColumn( 6, _T("办公室电话"), LVCFMT_LEFT, 160 ); 

	for (int i=0;i<((Cwhu_VoiceCardDlg*)whu_ParentDlg)->m_ContactListData[0].GetSize();i++)
	{
		char posbuf [10] = _T("");
		itoa(i+1,posbuf,10);
		m_ContactList.InsertItem(i, posbuf);//插入行 
		for (int j=0;j<7;j++)
		{
			CString ss = ((Cwhu_VoiceCardDlg*)whu_ParentDlg)->m_ContactListData[j].GetAt(i);
			m_ContactList.SetItemText(i,j,ss);
		}	
	}
}