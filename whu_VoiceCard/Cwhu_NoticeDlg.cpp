// Cwhu_NoticeDlg.cpp : implementation file
//

#include "stdafx.h"
#include "whu_VoiceCard.h"
#include "Cwhu_NoticeDlg.h"
#include "whu_VoiceCardDlg.h"
#include "afxdialogex.h"
#include "Mmsystem.h"
#include  <Windows.h>

// Cwhu_NoticeDlg dialog

IMPLEMENT_DYNAMIC(Cwhu_NoticeDlg, CDialogEx)

Cwhu_NoticeDlg::Cwhu_NoticeDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(Cwhu_NoticeDlg::IDD, pParent)
	, m_NoticeNum(_T(""))
{

}

Cwhu_NoticeDlg::~Cwhu_NoticeDlg()
{
}

void Cwhu_NoticeDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_LISTCONTACT02, m_ContactList);
	DDX_Text(pDX, IDC_NUMBER, m_NoticeNum);
}


BEGIN_MESSAGE_MAP(Cwhu_NoticeDlg, CDialogEx)
	ON_BN_CLICKED(IDC_RECORDSTART, &Cwhu_NoticeDlg::OnBnClickedRecordstart)
	ON_BN_CLICKED(IDC_BUTTONBACK, &Cwhu_NoticeDlg::OnBnClickedButtonback)
	ON_BN_CLICKED(IDC_CONTRACTADD, &Cwhu_NoticeDlg::OnBnClickedContractadd)
	ON_BN_CLICKED(IDC_SENDNOTICE, &Cwhu_NoticeDlg::OnBnClickedSendnotice)
	ON_BN_CLICKED(IDC_LISTEN, &Cwhu_NoticeDlg::OnBnClickedListen)
	ON_WM_TIMER()
	ON_NOTIFY(LVN_COLUMNCLICK, IDC_LISTCONTACT02, &Cwhu_NoticeDlg::OnLvnColumnclickListcontact02)
	ON_MESSAGE(WM_UPATEDATA, OnUpdateData)
END_MESSAGE_MAP()


// Cwhu_NoticeDlg message handlers


void Cwhu_NoticeDlg::OnBnClickedRecordstart()
{
	CString m_str;
	GetDlgItemText(IDC_RECORDSTART,m_str);
	if (m_str == _T("¼��"))
	{
		// TODO: Add your control notification handler code here
		SetDlgItemText(IDC_RECORDSTART,_T("ȡ��"));
		((Cwhu_VoiceCardDlg*)whu_ParentDlg)->whu_RecordNoticeDone = false;
		GetDlgItem(IDC_RECORDFILE)->SetWindowTextA(_T("������绰¼��֪ͨ...�����ҵ绰..."));
		GetDlgItem(IDC_LISTEN)->EnableWindow(false);
		GetDlgItem(IDC_SENDNOTICE)->EnableWindow(false);
		UpdateData(false);
		SYSTEMTIME whu_SystemTime;
		GetLocalTime(&whu_SystemTime);
		((Cwhu_VoiceCardDlg*)whu_ParentDlg)->whu_NoticeFile.Format(_T("C:\\Whu_Data\\����֪ͨ\\֪ͨ%d��%d��%d��%dʱ%d��.wav\0"),whu_SystemTime.wYear,whu_SystemTime.wMonth,whu_SystemTime.wDay,whu_SystemTime.wHour,whu_SystemTime.wMinute);
		m_filename = ((Cwhu_VoiceCardDlg*)whu_ParentDlg)->whu_NoticeFile;
		((Cwhu_VoiceCardDlg*)whu_ParentDlg)->whu_RecordNotice = true;

	} 
	else //m_str == _T("ȡ��")
	{
		SetDlgItemText(IDC_RECORDSTART,_T("¼��"));

		((Cwhu_VoiceCardDlg*)whu_ParentDlg)->whu_RecordNotice = false;
		GetDlgItem(IDC_RECORDFILE)->SetWindowTextA(_T("����ȡ��"));

	}
	
	
}


void Cwhu_NoticeDlg::OnBnClickedButtonback()
{
	// TODO: Add your control notification handler code here
	//whu_SaveContact();
	CDialogEx::OnCancel();
}
void Cwhu_NoticeDlg::whu_IntiList()
{
	//��ϵ���б��ʼ��
	LONG lStyle; 
	lStyle = GetWindowLong(m_ContactList.m_hWnd, GWL_STYLE);
	lStyle &= ~LVS_TYPEMASK; 
	lStyle |= LVS_REPORT; 
	SetWindowLong(m_ContactList.m_hWnd, GWL_STYLE, lStyle);
	DWORD dwStyle = m_ContactList.GetExtendedStyle(); 
	dwStyle |= LVS_EX_FULLROWSELECT;
	dwStyle |= LVS_EX_GRIDLINES;
	dwStyle |= LVS_EX_CHECKBOXES;
	m_ContactList.SetExtendedStyle(dwStyle); 
	m_ContactList.InsertColumn( 0, _T("���"), LVCFMT_LEFT, 80 );
	m_ContactList.InsertColumn( 1, _T("����"), LVCFMT_LEFT, 150 ); 
	m_ContactList.InsertColumn( 2, _T("��λ"), LVCFMT_LEFT, 150 );
	m_ContactList.InsertColumn( 3, _T("����"), LVCFMT_LEFT, 160 );
	m_ContactList.InsertColumn( 4, _T("ְ��"), LVCFMT_LEFT, 100 );
	m_ContactList.InsertColumn( 5, _T("�ֻ�"), LVCFMT_LEFT, 180); 
	m_ContactList.InsertColumn( 6, _T("�칫�ҵ绰"), LVCFMT_LEFT, 0 ); 
	m_ContactList.InsertColumn( 7, _T("�����"), LVCFMT_LEFT, 0 ); 
	m_ContactList.InsertColumn( 8, _T("����"), LVCFMT_LEFT, 0 );
	m_ContactList.InsertColumn( 9, _T("״̬"), LVCFMT_LEFT, 180 );
	//m_FaxContactList.InsertColumn( 0, _T("���"), LVCFMT_LEFT, 40 );
	//m_FaxContactList.InsertColumn( 1, _T("����"), LVCFMT_LEFT, 120 ); 
	//m_FaxContactList.InsertColumn( 2, _T("��λ"), LVCFMT_LEFT, 120 );
	//m_FaxContactList.InsertColumn( 3, _T("����"), LVCFMT_LEFT, 160 );
	//m_FaxContactList.InsertColumn( 4, _T("ְ��"), LVCFMT_LEFT, 80 );
	//m_FaxContactList.InsertColumn( 5, _T("�����"), LVCFMT_LEFT, 120 );
	//CString p = ((Cwhu_VoiceCardDlg*)whu_ParentDlg)->m_test;
	int num = ((Cwhu_VoiceCardDlg*)whu_ParentDlg)->m_ContactListData[0].GetSize();
	int HeadNum = ((Cwhu_VoiceCardDlg*)whu_ParentDlg)->ConHeadNum;
	for (int i=1;i<num;i++)
	{
		CString buf;
		buf.Format(_T("%d"),i);
		m_ContactList.InsertItem(i-1,buf);
		m_ContactList.SetItemText(i-1,0,buf);
		for (int j =0;j<HeadNum;j++)
		{
			m_ContactList.SetItemText(i-1,j+1,((Cwhu_VoiceCardDlg*)whu_ParentDlg)->m_ContactListData[j].GetAt(i)); 
		}
	}
}
BOOL Cwhu_NoticeDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();
	whu_IntiList();
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
	SetTimer(1,1000,NULL);
	GetDlgItem(IDC_SENDNOTICE)->EnableWindow(false);
	GetDlgItem(IDC_LISTEN)->EnableWindow(false);
	NumCount = 0;
	return true;
}

void Cwhu_NoticeDlg::OnBnClickedContractadd()
{
	// TODO: Add your control notification handler code here
	UpdateData(true);
	if (m_NoticeNum != _T(""))
	{
		int pos = m_ContactList.GetItemCount();
		char buf[3];
		itoa(pos,buf,10);
		CString m_str;
		m_str.Format(_T("%d"),pos+1);
		m_ContactList.InsertItem(pos,m_str);
		m_ContactList.SetItemText(pos,0,m_str);
		for (int j =1;j<9;j++)
		{
			m_ContactList.SetItemText(pos,j,_T("δ֪")); 
		}
		m_ContactList.SetItemText(pos,5,m_NoticeNum);
	}
	
	//

}
void Cwhu_NoticeDlg::whu_SaveContact()
{
	for (int i=0;i<11;i++)
	{
		((Cwhu_VoiceCardDlg*)whu_ParentDlg)->m_ContactListData[i].RemoveAll();
	}
	int colum = m_ContactList.GetHeaderCtrl()->GetItemCount() -1;//״̬�����浽��ϵ����//
	int item = m_ContactList.GetItemCount();
	((Cwhu_VoiceCardDlg*)whu_ParentDlg)->m_MyContactList.DeleteAllItems();
	for (int i=0;i<item;i++)
	{
		CString m_Str;
		m_Str.Format(_T("%d"),i+1);
		((Cwhu_VoiceCardDlg*)whu_ParentDlg)->m_MyContactList.InsertItem(i,m_Str);
		for (int j=0;j<colum;j++)
		{
			CString m_Str = m_ContactList.GetItemText(i,j);
			((Cwhu_VoiceCardDlg*)whu_ParentDlg)->m_ContactListData[j].Add(m_Str);
			((Cwhu_VoiceCardDlg*)whu_ParentDlg)->m_MyContactList.SetItemText(i,j,m_Str);
		}
	}

}

void Cwhu_NoticeDlg::OnBnClickedSendnotice()
{
	// TODO: Add your control notification handler code here
	CString m_str;
	GetDlgItemText(IDC_SENDNOTICE,m_str);
	if (m_str == _T("����"))
	{
		SetDlgItemText(IDC_SENDNOTICE,_T("ȡ��"));
		m_NoticeNumArr.RemoveAll();
		CString m_NumBuf[MAX_PATH];
		bool m_Check = false;
		for(int i=0; i<m_ContactList.GetItemCount(); i++) 
		{
			m_Check = m_ContactList.GetCheck(i);//
			if(m_Check)  //��ѡ�и�ѡ��Ⱥ��//m_int == LVIS_SELECTED||
			{ 
				char buf[20];
				m_ContactList.GetItemText(i,5,buf,12);//�ֻ���
				CString TelNum;
				TelNum.Format(_T("%s"),buf);
				if (TelNum!=_T(""))
				{
					m_NoticeNumArr.Add(TelNum);
					m_ContactList.SetItemText(i,9,_T("׼������"));
					((Cwhu_VoiceCardDlg*)whu_ParentDlg)->whu_NoticeState = NOTICE_PREPARE;
				}else{
					m_ContactList.SetItemText(i,9,_T("����Ϊ��"));
				}

			}
		}
		NumCount = m_NoticeNumArr.GetCount();
		//int num = m_NoticeNumArr.GetCount();
		for (int i=0;i<NumCount;i++)
		{
			m_NumBuf[i] = m_NoticeNumArr.GetAt(i);
		}
		CString m_RecordFile = ((Cwhu_VoiceCardDlg*)whu_ParentDlg)->whu_NoticeFile;
		if (NumCount>0&&m_RecordFile!=_T(""))
		{
			((Cwhu_VoiceCardDlg*)whu_ParentDlg)->whu_SendNotice(m_NumBuf,NumCount);
		}

	} 
	else
	{
		////////////////////////////�¼�ģ��//////////////////////////////////////////////
		if (whu_NumHaveSend < NumCount) //��δ�������
		{
			int value = AfxMessageBox(_T("���ڷ���֪ͨ���Ƿ������"));
			if (value == IDOK)
			{
				SetDlgItemText(IDC_SENDNOTICE,_T("����"));
				GetDlgItem(IDC_SENDNOTICE)->EnableWindow(false);
				//������ȡ������//
				whu_NumHaveSend = NumCount;
				int AtrkCh = 2;
				((Cwhu_VoiceCardDlg*)whu_ParentDlg)->whu_UserChState = USER_IDLE;
				((Cwhu_VoiceCardDlg*)whu_ParentDlg)->whu_TrunkChState = TRK_IDLE;
				SsmHangup(AtrkCh);
				((Cwhu_VoiceCardDlg*)whu_ParentDlg)->whu_ShowCallOutInfo(false);
			}
		}
		
		////////////////////////////�¼�ģ��//////////////////////////////////////////////

	}
	
}

void Cwhu_NoticeDlg::OnBnClickedListen()
{
	// TODO: Add your control notification handler code here
	
	//bool m_bRecordNotice=((Cwhu_VoiceCardDlg*)whu_ParentDlg)->whu_RecordNotice;
	//m_filename = ((Cwhu_VoiceCardDlg*)whu_ParentDlg)->whu_NoticeFile;
	if (m_filename!=_T(""))
	{
		sndPlaySound(m_filename,SND_NOSTOP);
		
	} 
}


void Cwhu_NoticeDlg::OnTimer(UINT_PTR nIDEvent)
{
	// TODO: Add your message handler code here and/or call default
	//((Cwhu_VoiceCardDlg*)whu_ParentDlg)->m_MyLogList
	//whu_showNoticeState();
	//����ж�whu_RecordNoticeDone��������ˣ�����ʾ�����ؼ�//
	if (((Cwhu_VoiceCardDlg*)whu_ParentDlg)->whu_RecordNoticeDone == true)
	{
		//GetDlgItem(IDC_RECORDFILE)->ShowWindow(true);
		GetDlgItem(IDC_RECORDFILE)->SetWindowTextA(m_filename);
		GetDlgItem(IDC_LISTEN)->EnableWindow(true);
		GetDlgItem(IDC_SENDNOTICE)->EnableWindow(true);
		SetDlgItemText(IDC_RECORDSTART,_T("¼��"));
		//UpdateData(false);

	} 
	else
	{
		//GetDlgItem(IDC_RECORDFILE)->ShowWindow(false);
		//GetDlgItem(IDC_LISTEN)->EnableWindow(false);
	}
	////////////////////////////�¼�ģ��//////////////////////////////////////////////
	bool m_SendNoticeDone = false;
	if (NumCount == whu_NumHaveSend && NumCount!=0) //˵����ǰ�����Ѿ������//
	{
		SetDlgItemText(IDC_SENDNOTICE,_T("����"));
	}
	//////////////////////////�¼�ģ��/////////////////////////////////////////////////
	CDialogEx::OnTimer(nIDEvent);
}
void Cwhu_NoticeDlg::whu_showNoticeState(int whu_NoticeState)
{
	//m_filename = ((Cwhu_VoiceCardDlg*)whu_ParentDlg)->whu_NoticeFile;
	//int whu_NoticeState = ((Cwhu_VoiceCardDlg*)whu_ParentDlg)->whu_NoticeState;
	CString m_Num = whu_NumberBuf;
	whu_ContractData m_data = ((Cwhu_VoiceCardDlg*)whu_ParentDlg)->whu_GetContractData(whu_NumberBuf);
	for (int i=0;i<m_ContactList.GetItemCount();i++)
	{
		if (m_Num == m_ContactList.GetItemText(i,5)&&m_data.Name ==m_ContactList.GetItemText(i,1))
		{
			if (whu_NoticeState == NOTICE_PLAYING)
			{
				m_ContactList.SetItemText(i,9,_T("���ڷ���..."));
			}
			else if (whu_NoticeState == NOTICE_HANGUP&&m_ContactList.GetItemText(i,9)==_T("���ڷ���..."))
			{
				m_ContactList.SetItemText(i,9,_T("�޷�ȷ���Ƿ�μ�"));
			}
			else if (whu_NoticeState == NOTICE_ATTEND&&m_ContactList.GetItemText(i,9)==_T("���ڷ���..."))
			{

				m_ContactList.SetItemText(i,9,_T("ȷ�ϲμ�"));

			}
			else if (whu_NoticeState == NOTICE_NOATTEND&&m_ContactList.GetItemText(i,9)==_T("���ڷ���..."))
			{
				m_ContactList.SetItemText(i,9,_T("�޷��μ�"));
			}
			else if (whu_NoticeState == NOTICE_NONE)
			{
				m_ContactList.SetItemText(i,9,_T(""));
			}
			else if (whu_NoticeState == NOTICE_PREPARE)
			{
				m_ContactList.SetItemText(i,9,_T("׼������"));
			}
			else if (whu_NoticeState == NOTICE_BUSY)
			{
				m_ContactList.SetItemText(i,9,_T("δ����"));
			}
			
		}
	}
}
void Cwhu_NoticeDlg::whu_SortContactList(CListCtrl &m_List,int colum,int mode)
{
	//m_ContactListData[colum]
	//���ַ���������//
	CStringArray m_ArrData[20];
	for (int i=0;i<m_List.GetItemCount();i++)
	{
		for (int j=0;j<m_List.GetHeaderCtrl()->GetItemCount();j++)
		{
			m_ArrData[j].Add(m_List.GetItemText(i,j));
		}
	}
	CString m_str1;
	CString m_str2;
	for (int row = 1;row <m_ArrData[0].GetCount();row++)
	{
		for (int i=0;i<m_ArrData[0].GetCount()-1;i++)
		{
			m_str1 = m_ArrData[colum].GetAt(i);
			m_str2 = m_ArrData[colum].GetAt(i+1);
			if (mode == 0) //����
			{
				if (m_str1 > m_str2)
				{
					for (int j=1;j<m_List.GetHeaderCtrl()->GetItemCount();j++)
					{
						CString m_str3 = m_ArrData[j].GetAt(i);
						CString m_str4 = m_ArrData[j].GetAt(i+1);
						m_ArrData[j].RemoveAt(i);
						m_ArrData[j].RemoveAt(i);
						m_ArrData[j].InsertAt(i,m_str3);
						m_ArrData[j].InsertAt(i,m_str4);
						m_str3 = m_ArrData[j].GetAt(i);
						m_str4 = m_ArrData[j].GetAt(i+1);

					}
				}
			}
			else if (mode == 1)//����
			{
				if (m_str1 < m_str2)
				{
					for (int j=1;j<m_List.GetHeaderCtrl()->GetItemCount();j++)
					{
						CString m_str3 = m_ArrData[j].GetAt(i);
						CString m_str4 = m_ArrData[j].GetAt(i+1);
						m_ArrData[j].RemoveAt(i);
						m_ArrData[j].RemoveAt(i);
						m_ArrData[j].InsertAt(i,m_str3);
						m_ArrData[j].InsertAt(i,m_str4);
						m_str3 = m_ArrData[j].GetAt(i);
						m_str4 = m_ArrData[j].GetAt(i+1);

					}
				}
			}

		}
	}
	//��������//
	m_List.DeleteAllItems();
	for (int i=0;i<m_ArrData[0].GetSize();i++)
	{
		char buf[3];
		itoa(i+1,buf,10);
		m_List.InsertItem(i,buf);
		for (int j =0;j<m_List.GetHeaderCtrl()->GetItemCount();j++)
		{
			m_List.SetItemText(i,j,m_ArrData[j].GetAt(i)); 
		}
	}

}

void Cwhu_NoticeDlg::OnLvnColumnclickListcontact02(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMLISTVIEW pNMLV = reinterpret_cast<LPNMLISTVIEW>(pNMHDR);
	// TODO: Add your control notification handler code here
	int colum = pNMLV->iSubItem;
	static int mode = 0;
	whu_SortContactList(m_ContactList,colum,mode);
	if (mode == 0)
	{
		mode = 1;
	}
	else{
		mode = 0;
	}
	*pResult = 0;
}
LRESULT Cwhu_NoticeDlg::OnUpdateData(WPARAM wParam, LPARAM lParam)
{
	int* m_NoticeState = (int*)lParam;
	whu_showNoticeState(*m_NoticeState);
	UpdateData(FALSE);
	return true;
}