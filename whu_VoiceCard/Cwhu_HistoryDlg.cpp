// Cwhu_HistoryDlg.cpp : implementation file
//

#include "stdafx.h"
#include "whu_VoiceCard.h"
#include "whu_VoiceCardDlg.h"
#include "Cwhu_HistoryDlg.h"
#include "afxdialogex.h"
#include  <string>
using namespace std;

// Cwhu_HistoryDlg dialog

IMPLEMENT_DYNAMIC(Cwhu_HistoryDlg, CDialogEx)

Cwhu_HistoryDlg::Cwhu_HistoryDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(Cwhu_HistoryDlg::IDD, pParent)
{

}

Cwhu_HistoryDlg::~Cwhu_HistoryDlg()
{
}

void Cwhu_HistoryDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_HISTORYLIST, m_HistoryList);
	DDX_Control(pDX, IDC_DATETIMEPICKERLOG, m_DateCtl);
}
void Cwhu_HistoryDlg::whu_InitList()
{
	LONG lStyle; 
	lStyle = GetWindowLong(m_HistoryList.m_hWnd, GWL_STYLE);//��ȡ��ǰ����style 
	lStyle &= ~LVS_TYPEMASK; //�����ʾ��ʽλ 
	lStyle |= LVS_REPORT; //����style 
	SetWindowLong(m_HistoryList.m_hWnd, GWL_STYLE, lStyle);//����style 
	DWORD dwStyle = m_HistoryList.GetExtendedStyle(); 
	dwStyle |= LVS_EX_FULLROWSELECT;//ѡ��ĳ��ʹ���и�����ֻ������report����listctrl�� //
	dwStyle |= LVS_EX_GRIDLINES;//�����ߣ�ֻ������report����listctrl�� //
	dwStyle |= LVS_EX_CHECKBOXES;//itemǰ����checkbox�ؼ� //
	m_HistoryList.SetExtendedStyle(dwStyle); //������չ��� //

	m_HistoryList.InsertColumn( 0, "ID", LVCFMT_LEFT, 60 );
	m_HistoryList.InsertColumn( 1, _T("ʱ��"), LVCFMT_LEFT, 200 ); 
	m_HistoryList.InsertColumn( 2, _T("ֵ����"), LVCFMT_LEFT, 80 );
	m_HistoryList.InsertColumn( 3, _T("�¼�"), LVCFMT_LEFT, 160 );
	m_HistoryList.InsertColumn( 4, _T("��λ"), LVCFMT_LEFT, 120 );
	m_HistoryList.InsertColumn( 5, _T("����"), LVCFMT_LEFT, 100 ); 
	m_HistoryList.InsertColumn( 6, _T("����"), LVCFMT_LEFT, 160 ); 
	m_HistoryList.InsertColumn( 7, _T("�ļ�"), LVCFMT_LEFT, 200 ); 
	m_HistoryList.InsertColumn( 8, _T("״̬"), LVCFMT_LEFT, 120 ); 
	m_HistoryList.InsertColumn( 9, _T("��ע"), LVCFMT_LEFT, 160 ); 

	pos = 0;
}
void Cwhu_HistoryDlg::whu_AddHistoryList(CString time,CString duty,CString taskmode,CString unit,CString name,CString number,CString file,CString state,CString beizhu)
{
	char posbuf [10] = _T("");
	itoa(pos+1,posbuf,10);
	m_HistoryList.InsertItem(pos, posbuf);//������ 
	m_HistoryList.SetItemText(pos, 1, time);//�������� 
	m_HistoryList.SetItemText(pos, 2, duty);//�������� 
	m_HistoryList.SetItemText(pos, 3, taskmode);//�������� 
	m_HistoryList.SetItemText(pos, 4,unit);//�������� 
	m_HistoryList.SetItemText(pos, 5,name);//�������� 
	m_HistoryList.SetItemText(pos, 6, number);//�������� 
	m_HistoryList.SetItemText(pos, 7, file);//�������� 
	m_HistoryList.SetItemText(pos, 8, state);//��������
	m_HistoryList.SetItemText(pos, 9, beizhu);//��������

	pos++;
}
void Cwhu_HistoryDlg::whu_ModifyCheck(CString m_str)
{
	m_HistoryList.SetItemText(pos-1, 8, m_str);//��������
}

BEGIN_MESSAGE_MAP(Cwhu_HistoryDlg, CDialogEx)
	ON_NOTIFY(NM_RCLICK, IDC_HISTORYLIST, &Cwhu_HistoryDlg::OnNMRClickHistorylist)
	ON_COMMAND(ID_32833, &Cwhu_HistoryDlg::OnAddDetailLog)
	ON_NOTIFY(DTN_DATETIMECHANGE, IDC_DATETIMEPICKERLOG, &Cwhu_HistoryDlg::OnDtnDatetimechangeDatetimepickerlog)
	ON_BN_CLICKED(IDC_BUTTONQUIT, &Cwhu_HistoryDlg::OnBnClickedButtonquit)
	ON_NOTIFY(LVN_COLUMNCLICK, IDC_HISTORYLIST, &Cwhu_HistoryDlg::OnLvnColumnclickHistorylist)
	ON_COMMAND(ID_32834, &Cwhu_HistoryDlg::OnOpenFile)
	ON_COMMAND(ID_32835, &Cwhu_HistoryDlg::OnShowDetailLog)
END_MESSAGE_MAP()


// Cwhu_HistoryDlg message handlers


void Cwhu_HistoryDlg::OnNMRClickHistorylist(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMITEMACTIVATE pNMItemActivate = reinterpret_cast<LPNMITEMACTIVATE>(pNMHDR);
	// TODO: Add your control notification handler code here

	NM_LISTVIEW* pNMListView = (NM_LISTVIEW*)pNMHDR;
	m_ClickPos = pNMListView->iItem;
	if(pNMListView->iItem != -1 &&m_ClickPos<m_HistoryList.GetItemCount()) 
	{ 
		LVCOLUMN lvcol; 
		char    str[256]; 
		lvcol.mask = LVCF_TEXT; 
		lvcol.pszText = str; 
		lvcol.cchTextMax = 256; 
		m_HistoryList.GetColumn(pNMListView->iSubItem,&lvcol);
		//AfxMessageBox(lvcol.pszText); 
		//lvcol.pszTextΪ����//
		CString m_head = lvcol.pszText;
		if (m_head == _T("�ļ�"))
		{
			DWORD dwPos = GetMessagePos(); 
			CPoint point( LOWORD(dwPos), HIWORD(dwPos) ); 
			CMenu menu; 
			VERIFY( menu.LoadMenu( IDR_MENUSHOWLOG ) ); 
			CMenu* popup = menu.GetSubMenu(0); 
			ASSERT( popup != NULL ); 
			popup->TrackPopupMenu(TPM_LEFTALIGN | TPM_RIGHTBUTTON,
				point.x, point.y, this ); 
		}else{
			DWORD dwPos = GetMessagePos(); 
			CPoint point( LOWORD(dwPos), HIWORD(dwPos) ); 
			CMenu menu; 
			VERIFY( menu.LoadMenu( IDR_MENULOGDETAIL ) ); 
			CMenu* popup = menu.GetSubMenu(0); 
			ASSERT( popup != NULL ); 
			popup->TrackPopupMenu(TPM_LEFTALIGN | TPM_RIGHTBUTTON,
				point.x, point.y, this ); 
		}
		
	} 


	*pResult = 0;
}



BOOL Cwhu_HistoryDlg::OnInitDialog()
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

	whu_ReSize();

	SetWindowText(_T("��־"));
	whu_InitList();
	SwitchLog();//������棬������ʾ�������־//
	return true;
}

void Cwhu_HistoryDlg::OnAddDetailLog()
{
	// TODO: Add your command handler code here
	CDetailLogDlg *m_DlgDetail = new CDetailLogDlg();
	m_DlgDetail->whu_ParentDlg = this;
	m_DlgDetail->Create(IDD_LOG,this);
	m_DlgDetail->ShowWindow(SW_SHOW);
}


void Cwhu_HistoryDlg::OnDtnDatetimechangeDatetimepickerlog(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMDATETIMECHANGE pDTChange = reinterpret_cast<LPNMDATETIMECHANGE>(pNMHDR);
	// TODO: Add your control notification handler code here
	SwitchLog();
	*pResult = 0;
}


void Cwhu_HistoryDlg::OnBnClickedButtonquit()
{
	// TODO: Add your control notification handler code here
	OnSaveLogToFile();
	CDialogEx::OnCancel();
}
void Cwhu_HistoryDlg::OnSaveLogToFile()
{
	CTime time;
	m_DateCtl.GetTime(time);
	CString strTime;
	strTime.Format(_T("%d%s%d%s%d"),time.GetYear(),_T("-"),time.GetMonth(),_T("-"),time.GetDay());
	CString file = _T("");
	file.Format(_T("%s%s%s"),_T("C:\\Whu_Data\\ֵ��ϵͳ��־\\"),strTime,_T(".txt"));
	whu_ListToTxt(&m_HistoryList,file);
}
void Cwhu_HistoryDlg::whu_ListToTxt(CListCtrl *m_list,CString m_txtfile)
{
	//�����ݱ��浽�ĵ�//
	//���統����־���ڣ��ȶ�ȡ//
	CString folder = _T("C:\\Whu_Data\\ֵ��ϵͳ��־");
	if (whu_findfile(m_txtfile,folder) == true)
	{
		CStdioFile m_file(m_txtfile,CFile::modeRead);
		m_file.Close();
		m_file.Remove(m_txtfile);
	}
	//�����µ��ļ�//
	CStdioFile m_file(m_txtfile,CFile::modeCreate|CFile::modeWrite|CFile::modeNoTruncate|CFile::modeRead);
	for (int i=0;i<m_list->GetItemCount();i++) //����//
	{
		CString m_LineStr;
		for (int j =1;j<m_list->GetHeaderCtrl()->GetItemCount();j++) //����//
		{
			CString m_str;
			m_str = m_list->GetItemText(i,j);
			m_str = m_str+_T(" "); //ÿ������֮���Կո�������//
			m_LineStr = m_LineStr+m_str;
		}
		m_LineStr = m_LineStr+_T("\n");
		m_file.WriteString(m_LineStr);
	}
	m_file.Close();
}
bool Cwhu_HistoryDlg::whu_findfile(CString m_file,CString folder)
{
	folder = folder + _T("\\*.*");
	CFileFind filefind;
	bool bworking = filefind.FindFile(folder);
	bool find = false;
	while(bworking)
	{
		bworking = filefind.FindNextFileA();
		CString m_curfile = filefind.GetFilePath();
		if (m_curfile == m_file)
		{
			find =true;
		}
	}
	return find;
}
void Cwhu_HistoryDlg::SwitchLog()
{
	CTime time;
	m_DateCtl.GetTime(time);
	CString strTime;
	strTime.Format(_T("%d%s%d%s%d"),time.GetYear(),_T("-"),time.GetMonth(),_T("-"),time.GetDay());
	CString file = _T("");
	file.Format(_T("%s%s%s"),_T("C:\\Whu_Data\\ֵ��ϵͳ��־\\"),strTime,_T(".txt"));
	//�ļ�����//
	CString folder = _T("C:\\Whu_Data\\ֵ��ϵͳ��־");
	//���ļ�����Ѱ���ļ�//
	if (((Cwhu_VoiceCardDlg*)whu_ParentDlg)->whu_findfile(file,folder))
	{
		((Cwhu_VoiceCardDlg*)whu_ParentDlg)->whu_TxtToList(file,&m_HistoryList);
	}
	else
	{
		m_HistoryList.DeleteAllItems();
	}
	SaveLog();
}
void Cwhu_HistoryDlg::SaveLog()
{
	for (int i=0;i<10;i++)
	{
		whu_ArrLog[i].RemoveAll();
	}
	for (int i=0;i<m_HistoryList.GetItemCount();i++)  //����
	{
		for (int j=0;j<m_HistoryList.GetHeaderCtrl()->GetItemCount();j++) //����
		{
			CString m_str = m_HistoryList.GetItemText(i,j);
			whu_ArrLog[j].InsertAt(i,m_str);
		}	
	}
}

void Cwhu_HistoryDlg::OnLvnColumnclickHistorylist(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMLISTVIEW pNMLV = reinterpret_cast<LPNMLISTVIEW>(pNMHDR);
	// TODO: Add your control notification handler code here
	
	int colum = pNMLV->iSubItem;
	static int mode = 0;
	//���ַ���������m_ContactListData//
	CString m_str1;
	CString m_str2;
	for (int row = 1;row <whu_ArrLog[0].GetCount();row++)
	{
		for (int i=0;i<whu_ArrLog[0].GetCount()-1;i++)
		{
			m_str1 = whu_ArrLog[colum].GetAt(i);
			m_str2 = whu_ArrLog[colum].GetAt(i+1);
			if (mode == 0) //����
			{
				if (m_str1 > m_str2)
				{
					for (int j=1;j<10;j++)
					{
						CString m_str3 = whu_ArrLog[j].GetAt(i);
						CString m_str4 = whu_ArrLog[j].GetAt(i+1);
						whu_ArrLog[j].RemoveAt(i);
						whu_ArrLog[j].RemoveAt(i);
						whu_ArrLog[j].InsertAt(i,m_str3);
						whu_ArrLog[j].InsertAt(i,m_str4);
						m_str3 = whu_ArrLog[j].GetAt(i);
						m_str4 = whu_ArrLog[j].GetAt(i+1);

					}
				}
			}
			else if (mode == 1)//����
			{
				if (m_str1 < m_str2)
				{
					for (int j=1;j<10;j++)
					{
						CString m_str3 = whu_ArrLog[j].GetAt(i);
						CString m_str4 = whu_ArrLog[j].GetAt(i+1);
						whu_ArrLog[j].RemoveAt(i);
						whu_ArrLog[j].RemoveAt(i);
						whu_ArrLog[j].InsertAt(i,m_str3);
						whu_ArrLog[j].InsertAt(i,m_str4);
						m_str3 = whu_ArrLog[j].GetAt(i);
						m_str4 = whu_ArrLog[j].GetAt(i+1);

					}
				}
			}

		}
	}
	//��������//
	m_HistoryList.DeleteAllItems();
	for (int i=0;i<whu_ArrLog[0].GetSize();i++)
	{
		char buf[3];
		itoa(i+1,buf,10);
		m_HistoryList.InsertItem(i,buf);
		for (int j =0;j<10;j++)
		{
			m_HistoryList.SetItemText(i,j,whu_ArrLog[j].GetAt(i)); 
		}
	}

	if (mode == 0)
	{
		mode = 1;
	}
	else{
		mode = 0;
	}
	*pResult = 0;
}
CString Cwhu_HistoryDlg::whu_ChineseToPinyin(CString chinese)
{
	string sChinese = "�Ұ��й�"; // ������ַ���//
	char chr[3];
	wchar_t wchr = 0;

	char* buff = new char[sChinese.length()/2];
	memset(buff, 0x00, sizeof(char)*sChinese.length()/2+1);

	for (int i = 0, j = 0; i < (sChinese.length()/2); ++i) 
	{
		memset(chr, 0x00, sizeof(chr));
		chr[0] = sChinese[j++];
		chr[1] = sChinese[j++];
		chr[2] = '\0';

		// �����ַ��ı��� �磺'��' = 0xced2
		wchr = 0;
		wchr = (chr[0] & 0xff) << 8;
		wchr |= (chr[1] & 0xff);

		buff[i] = convert(wchr);
	}

	return _T("ss");
}
char Cwhu_HistoryDlg::convert(wchar_t n)
{
	if (IN(0xB0A1,0xB0C4,n)) return 'a';
	if (IN(0XB0C5,0XB2C0,n)) return 'b';
	if (IN(0xB2C1,0xB4ED,n)) return 'c';
	if (IN(0xB4EE,0xB6E9,n)) return 'd';
	if (IN(0xB6EA,0xB7A1,n)) return 'e';
	if (IN(0xB7A2,0xB8c0,n)) return 'f';
	if (IN(0xB8C1,0xB9FD,n)) return 'g';
	if (IN(0xB9FE,0xBBF6,n)) return 'h';
	if (IN(0xBBF7,0xBFA5,n)) return 'j';
	if (IN(0xBFA6,0xC0AB,n)) return 'k';
	if (IN(0xC0AC,0xC2E7,n)) return 'l';
	if (IN(0xC2E8,0xC4C2,n)) return 'm';
	if (IN(0xC4C3,0xC5B5,n)) return 'n';
	if (IN(0xC5B6,0xC5BD,n)) return 'o';
	if (IN(0xC5BE,0xC6D9,n)) return 'p';
	if (IN(0xC6DA,0xC8BA,n)) return 'q';
	if (IN(0xC8BB,0xC8F5,n)) return 'r';
	if (IN(0xC8F6,0xCBF0,n)) return 's';
	if (IN(0xCBFA,0xCDD9,n)) return 't';
	if (IN(0xCDDA,0xCEF3,n)) return 'w';
	if (IN(0xCEF4,0xD188,n)) return 'x';
	if (IN(0xD1B9,0xD4D0,n)) return 'y';
	if (IN(0xD4D1,0xD7F9,n)) return 'z';
	return '\0';
}

void Cwhu_HistoryDlg::OnOpenFile()
{
	// TODO: Add your command handler code here
	CString file = m_HistoryList.GetItemText(m_ClickPos,7);
	if (file !=_T(""))
	{
		ShellExecute(NULL, "open", file,NULL,NULL, SW_SHOWNORMAL );
	}
	
	
}


void Cwhu_HistoryDlg::OnShowDetailLog()
{
	// TODO: Add your command handler code here
	CDetailLogDlg *m_DlgDetail = new CDetailLogDlg();
	m_DlgDetail->whu_ParentDlg = this;
	m_DlgDetail->Create(IDD_LOG,this);
	m_DlgDetail->ShowWindow(SW_SHOW);
}
void Cwhu_HistoryDlg::whu_ReSize()
{
	POINT Old;
	int cx,cy; 
	cx = GetSystemMetrics(SM_CXSCREEN); 
	cy = GetSystemMetrics(SM_CYSCREEN);

	CRect rcTemp; 
	rcTemp.BottomRight() = CPoint(cx*11/12, cy*11/12); 
	rcTemp.TopLeft() = CPoint(cx/12, cy/12); 

	CRect rect;  
	GetClientRect(&rect); //ȡ�ͻ�����С  
	Old.x=rect.right-rect.left;
	Old.y=rect.bottom-rect.top;
	MoveWindow(&rcTemp);

	float fsp[2];
	POINT Newp; //��ȡ���ڶԻ���Ĵ�С
	CRect recta;  
	GetClientRect(&recta); //ȡ�ͻ�����С//  
	Newp.x=recta.right-recta.left;
	Newp.y=recta.bottom-recta.top;
	fsp[0]=(float)Newp.x/Old.x;
	fsp[1]=(float)Newp.y/Old.y;
	CRect Rect;
	int woc;
	CPoint OldTLPoint,TLPoint; //���Ͻ�//
	CPoint OldBRPoint,BRPoint; //���½�//
	HWND hwndChild=::GetWindow(m_hWnd,GW_CHILD); //�г����пؼ�  
	while(hwndChild)  
	{  
		woc=::GetDlgCtrlID(hwndChild);//ȡ��ID
		GetDlgItem(woc)->GetWindowRect(Rect);  
		ScreenToClient(Rect);  
		OldTLPoint = Rect.TopLeft();  
		TLPoint.x = long(OldTLPoint.x*fsp[0]);  
		TLPoint.y = long(OldTLPoint.y*fsp[1]);  
		OldBRPoint = Rect.BottomRight();  
		BRPoint.x = long(OldBRPoint.x *fsp[0]);  
		BRPoint.y = long(OldBRPoint.y *fsp[1]); //�߶Ȳ��ɶ��Ŀؼ�����:combBox),��Ҫ�ı��ֵ.//
		Rect.SetRect(TLPoint,BRPoint);  
		GetDlgItem(woc)->MoveWindow(Rect,TRUE);
		hwndChild=::GetWindow(hwndChild, GW_HWNDNEXT);  
	}
	Old=Newp;
	//������������//
	//CFont m_editFont; 
	//m_editFont.CreatePointFont(240,_T("΢���ź�"));
	//GetDlgItem(IDC_CHOOSETIME)->SetFont(&m_editFont);
	//GetDlgItem(IDC_DATETIMEPICKERLOG)->SetFont(&m_editFont);
}