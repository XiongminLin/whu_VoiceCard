// Cwhu_FaxDlg.cpp : implementation file
//

#include "stdafx.h"
#include "whu_VoiceCard.h"
#include "whu_VoiceCardDlg.h"
#include "Cwhu_FaxDlg.h"
#include "afxdialogex.h"
//#include "whu_Global.h"

// Cwhu_FaxDlg dialog

IMPLEMENT_DYNAMIC(Cwhu_FaxDlg, CDialogEx)

Cwhu_FaxDlg::Cwhu_FaxDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(Cwhu_FaxDlg::IDD, pParent)
	, whu_DocFile(_T(""))
	, m_faxNum(_T(""))
{
	
}

Cwhu_FaxDlg::~Cwhu_FaxDlg()
{
}

void Cwhu_FaxDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Text(pDX, IDC_FAXFILE, whu_DocFile);
	DDX_Text(pDX, IDC_FAXNUMBER, m_faxNum);
	DDX_Control(pDX, IDC_FAXCONTACTLIST, m_FaxContactList);
}

int Cwhu_FaxDlg::DocumentToTiff(CString m_Inputfile,CString &m_OutputFile)
{

	CConvertAgent *pConvertEngine = NULL;
	pConvertEngine = new CConvertAgent(); 
	int iRet = -1;
	if(pConvertEngine)
	{
		iRet = pConvertEngine->InitAgent("FaxServer Printer", 300, "Demo","Demo");
		if(iRet == SM_SUCCESS) 
		{
			iRet = pConvertEngine->ConvertDoc(m_Inputfile,m_OutputFile);
		}
	} 

	if(pConvertEngine)
	{
		delete pConvertEngine;
		pConvertEngine = NULL;
	}
	return 0;

}
BEGIN_MESSAGE_MAP(Cwhu_FaxDlg, CDialogEx)
	ON_BN_CLICKED(IDC_SENDOK, &Cwhu_FaxDlg::OnBnClickedSendok)
	ON_BN_CLICKED(IDC_CHOOSEFAXFILE, &Cwhu_FaxDlg::OnBnClickedChoosefaxfile)
	ON_BN_CLICKED(IDC_SENDCANCLE, &Cwhu_FaxDlg::OnBnClickedSendcancle)
	ON_NOTIFY(NM_RCLICK, IDC_FAXCONTACTLIST, &Cwhu_FaxDlg::OnNMRClickFaxcontactlist)
	ON_BN_CLICKED(IDC_BUTTONADDNUM, &Cwhu_FaxDlg::OnBnClickedButtonaddnum)
	ON_BN_CLICKED(IDC_BUTTONreadfax, &Cwhu_FaxDlg::OnBnClickedButtonreadfax)
	ON_BN_CLICKED(IDC_BUTTONFAXSETTING, &Cwhu_FaxDlg::OnBnClickedButtonfaxsetting)
	ON_WM_TIMER()
	ON_NOTIFY(LVN_COLUMNCLICK, IDC_FAXCONTACTLIST, &Cwhu_FaxDlg::OnLvnColumnclickFaxcontactlist)
	ON_NOTIFY(NM_CLICK, IDC_FAXCONTACTLIST, &Cwhu_FaxDlg::OnNMClickFaxcontactlist)
	ON_MESSAGE(WM_UPATEDATA, OnUpdateData)
END_MESSAGE_MAP()


// Cwhu_FaxDlg message handlers


void Cwhu_FaxDlg::OnBnClickedSendok()
{
	// TODO: Add your control notification handler code here
	CString m_str;
	GetDlgItemText(IDC_SENDOK,m_str);
	if (m_str == _T("发送"))
	{
		//得到whu_NumBuf电话//
		m_FaxNumArr.RemoveAll();
		CString m_NumBuf[MAX_PATH];
		bool m_Check = false;
		for(int i=0; i<m_FaxContactList.GetItemCount(); i++) 
		{
			m_Check = m_FaxContactList.GetCheck(i);//
			if(m_Check)  //有选中复选框群发//m_int == LVIS_SELECTED||
			{ 
				char buf[20];
				m_FaxContactList.GetItemText(i,7,buf,12);
				CString faxnum;
				faxnum.Format(_T("%s"),buf);
				if (faxnum!=_T(""))
				{
					m_FaxNumArr.Add(faxnum);
					((Cwhu_VoiceCardDlg*)whu_ParentDlg)->whu_FaxSendState = FAX_PREPARE;
					m_FaxContactList.SetItemText(i,9,_T("准备就绪"));
				}else{
					m_FaxContactList.SetItemText(i,9,_T("号码为空"));
				}
			}
		}
		NumCount = m_FaxNumArr.GetCount();
		for (int i=0;i<NumCount;i++)
		{
			m_NumBuf[i] = m_FaxNumArr.GetAt(i);
		}
		if (NumCount>0&&whu_DocFile!=_T(""))
		{
			((Cwhu_VoiceCardDlg*)whu_ParentDlg)->whu_SendFax(m_NumBuf,whu_DocFile,NumCount);
			GetDlgItem(IDC_SENDOK)->SetWindowTextA(_T("取消"));
		}else{
			AfxMessageBox(_T("传真号码或是发送文件为空"));
		}
	} 
	else
	{
		////////////////////////////新加模块//////////////////////////////////////////////
		if (whu_NumHaveSend < NumCount) //还未发送完毕
		{
			int value = AfxMessageBox(_T("正在发送通知，是否放弃？"));
			if (value == IDOK)
			{
				SetDlgItemText(IDC_SENDOK,_T("发送"));
				((Cwhu_VoiceCardDlg*)whu_ParentDlg)->whu_FaxSendState = FAX_CANCEL;
				for (int ItemNum=0;ItemNum<m_FaxContactList.GetItemCount();ItemNum++)
				{
					if (m_FaxContactList.GetCheck(ItemNum) == true)
					{
						m_FaxContactList.SetItemText(ItemNum,9,_T("任务取消"));
					}
				}
				//下面是取消任务//
				whu_NumHaveSend = NumCount;
				int AtrkCh = 2;
				SsmFaxStop(AtrkCh);
				SsmHangup(AtrkCh);
				((Cwhu_VoiceCardDlg*)whu_ParentDlg)->whu_UserChState = USER_IDLE;
				((Cwhu_VoiceCardDlg*)whu_ParentDlg)->whu_TrunkChState = TRK_IDLE;
				((Cwhu_VoiceCardDlg*)whu_ParentDlg)->whu_ShowFaxOutInfo(false);


			}
		}

		////////////////////////////新加模块//////////////////////////////////////////////


	}
	
}
int  Cwhu_FaxDlg::SeparateCallNum(CString m_AllNum,CString *m_SeparateNum)
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

void Cwhu_FaxDlg::OnBnClickedChoosefaxfile()
{
	// TODO: Add your control notification handler code here
	CString m_pathname;
	CFileDialog cf(1, NULL, NULL, OFN_HIDEREADONLY|OFN_OVERWRITEPROMPT, "Fax File(*.tif;*.doc;*.docx;*.xls;*.xlsx;*.ppt;*.pptx;*.pdf;*.txt;*.html;*.jpg;*bmp)|*.tif;*.doc;*.docx;*.xls;*.xlsx;*.ppt;*.pptx;*.pdf;*.txt;*.html;*.jpg;*bmp|Tif File(*.fax)|*.fax|Tiff File(*.tiff)|*.tiff|All File(*.*)|*.*|",NULL);//可以打开doc、pdf等文件//
	cf.m_ofn.lpstrInitialDir= _T("C:\\Whu_Data\\扫描仪输出文件\\");  

	if(cf.DoModal() == IDOK)
	{
		whu_DocFile=cf.GetPathName();
		((Cwhu_VoiceCardDlg*)whu_ParentDlg)->whu_SendFaxFileTitle = cf.GetFileTitle();
		UpdateData(FALSE);
		GetDlgItem(IDC_BUTTONreadfax)->EnableWindow(true);
		for (int i=0;i<m_FaxContactList.GetItemCount();i++)
		{
			if (m_FaxContactList.GetCheck(i)==true)
			{
				GetDlgItem(IDC_SENDOK)->EnableWindow(true);
			}
		}		
	}
}
BEGIN_EVENTSINK_MAP(Cwhu_FaxDlg, CDialogEx)
	
END_EVENTSINK_MAP()





void Cwhu_FaxDlg::OnBnClickedSendcancle()
{
	// TODO: Add your control notification handler code here
	//whu_SaveContact();

	CDialogEx::OnOK();
}
void Cwhu_FaxDlg::whu_InitContactList()
{
	//联系人列表初始化
	LONG lStyle; 
	lStyle = GetWindowLong(m_FaxContactList.m_hWnd, GWL_STYLE);
	lStyle &= ~LVS_TYPEMASK; 
	lStyle |= LVS_REPORT; 
	SetWindowLong(m_FaxContactList.m_hWnd, GWL_STYLE, lStyle);
	DWORD dwStyle = m_FaxContactList.GetExtendedStyle(); 
	dwStyle |= LVS_EX_FULLROWSELECT;
	dwStyle |= LVS_EX_GRIDLINES;
	dwStyle |= LVS_EX_CHECKBOXES;
	m_FaxContactList.SetExtendedStyle(dwStyle); 
	m_FaxContactList.InsertColumn( 0, _T("编号"), LVCFMT_LEFT, 60 );
	m_FaxContactList.InsertColumn( 1, _T("姓名"), LVCFMT_LEFT, 180 ); 
	m_FaxContactList.InsertColumn( 2, _T("单位"), LVCFMT_LEFT, 180 );
	m_FaxContactList.InsertColumn( 3, _T("部门"), LVCFMT_LEFT, 180 );
	m_FaxContactList.InsertColumn( 4, _T("职务"), LVCFMT_LEFT, 150 );
	m_FaxContactList.InsertColumn( 5, _T("手机"), LVCFMT_LEFT, 0 ); 
	m_FaxContactList.InsertColumn( 6, _T("办公室电话"), LVCFMT_LEFT, 0 ); 
	m_FaxContactList.InsertColumn( 7, _T("传真号"), LVCFMT_LEFT, 180 ); 
	m_FaxContactList.InsertColumn( 8, _T("邮箱"), LVCFMT_LEFT, 0 );
	m_FaxContactList.InsertColumn( 9, _T("状态"), LVCFMT_LEFT, 200 );
	//m_FaxContactList.InsertColumn( 0, _T("编号"), LVCFMT_LEFT, 40 );
	//m_FaxContactList.InsertColumn( 1, _T("姓名"), LVCFMT_LEFT, 120 ); 
	//m_FaxContactList.InsertColumn( 2, _T("单位"), LVCFMT_LEFT, 120 );
	//m_FaxContactList.InsertColumn( 3, _T("部门"), LVCFMT_LEFT, 160 );
	//m_FaxContactList.InsertColumn( 4, _T("职务"), LVCFMT_LEFT, 80 );
	//m_FaxContactList.InsertColumn( 5, _T("传真号"), LVCFMT_LEFT, 120 );
	//CString p = ((Cwhu_VoiceCardDlg*)whu_ParentDlg)->m_test;
	int num = ((Cwhu_VoiceCardDlg*)whu_ParentDlg)->m_ContactListData[0].GetSize();
	int HeadNum = ((Cwhu_VoiceCardDlg*)whu_ParentDlg)->ConHeadNum;
	for (int i=1;i<num;i++)
	{
		CString buf;
		buf.Format(_T("%d"),i);
		m_FaxContactList.InsertItem(i-1,buf);
		m_FaxContactList.SetItemText(i-1,0,buf);
		for (int j =0;j<HeadNum;j++)
		{
			m_FaxContactList.SetItemText(i-1,j+1,((Cwhu_VoiceCardDlg*)whu_ParentDlg)->m_ContactListData[j].GetAt(i)); 
		}
	}
	//int HeadNum = ((Cwhu_VoiceCardDlg*)whu_ParentDlg)->ConHeadNum;
	//Lin_ImportArrToList(((Cwhu_VoiceCardDlg*)whu_ParentDlg)->m_ContactListData,HeadNum,m_FaxContactList);
}


BOOL Cwhu_FaxDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();
	whu_InitContactList();
	//将对话框移至屏幕中间//
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
	SetTimer(1,1000,NULL);
	test = 0;
	GetDlgItem(IDC_BUTTONreadfax)->EnableWindow(false);
	GetDlgItem(IDC_SENDOK)->EnableWindow(false);
	NumCount = 0;
	return true;
}

void Cwhu_FaxDlg::OnNMRClickFaxcontactlist(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMITEMACTIVATE pNMItemActivate = reinterpret_cast<LPNMITEMACTIVATE>(pNMHDR);
	// TODO: Add your control notification handler code here
	*pResult = 0;
}





void Cwhu_FaxDlg::OnBnClickedButtonaddnum()
{
	// TODO: Add your control notification handler code here
	UpdateData(true);
	if (m_faxNum !=_T(""))
	{
		int pos = m_FaxContactList.GetItemCount();
		char buf[3];
		itoa(pos,buf,10);
		CString m_str;
		m_str.Format(_T("%d"),pos+1);
		m_FaxContactList.InsertItem(pos,m_str);
		m_FaxContactList.SetItemText(pos,0,m_str);
		for (int j =2;j<9;j++)
		{
			m_FaxContactList.SetItemText(pos,j,_T("未知")); 
		}
		m_FaxContactList.SetItemText(pos,1,_T("陌生人")); 
		m_FaxContactList.SetItemText(pos,7,m_faxNum); 
	}
	
	
}
void Cwhu_FaxDlg::whu_SaveContact()
{
	for (int i=0;i<11;i++)
	{
		((Cwhu_VoiceCardDlg*)whu_ParentDlg)->m_ContactListData[i].RemoveAll();
	}
	int colum = m_FaxContactList.GetHeaderCtrl()->GetItemCount();
	int item = m_FaxContactList.GetItemCount();
	((Cwhu_VoiceCardDlg*)whu_ParentDlg)->m_MyContactList.DeleteAllItems();
	for (int i=0;i<item;i++)
	{
		CString m_Str;
		m_Str.Format(_T("%d"),i+1);
		((Cwhu_VoiceCardDlg*)whu_ParentDlg)->m_MyContactList.InsertItem(i,m_Str);
		for (int j=0;j<colum;j++)
		{
			CString m_Str = m_FaxContactList.GetItemText(i,j);
			((Cwhu_VoiceCardDlg*)whu_ParentDlg)->m_ContactListData[j].Add(m_Str);
			((Cwhu_VoiceCardDlg*)whu_ParentDlg)->m_MyContactList.SetItemText(i,j,m_Str);
		}
	}

}

void Cwhu_FaxDlg::OnBnClickedButtonreadfax()
{
	// TODO: Add your control notification handler code here
	if (whu_DocFile!=_T(""))
	{
		ShellExecute(NULL, "open", whu_DocFile,NULL,NULL, SW_SHOWNORMAL );
	}
	

}


void Cwhu_FaxDlg::OnBnClickedButtonfaxsetting()
{
	// TODO: Add your control notification handler code here
	Cwhu_FaxSettingDlg * m_FaxSettingDlg = new Cwhu_FaxSettingDlg();
	m_FaxSettingDlg->whu_ParentDlg = this;
	m_FaxSettingDlg->DoModal();
	//m_FaxSettingDlg->Create(IDD_FAXSETTING,this);
	//m_FaxSettingDlg->ShowWindow(true);

}
void Cwhu_FaxDlg::whu_ShowFaxState(int fax_state)
{
	//m_filename = ((Cwhu_VoiceCardDlg*)whu_ParentDlg)->whu_NoticeFile;
	//int whu_FaxSendState = ((Cwhu_VoiceCardDlg*)whu_ParentDlg)->whu_FaxSendState;
	CString m_Num = whu_NumberBuf;
	if (m_Num !=_T(""))
	{
		whu_ContractData m_data = ((Cwhu_VoiceCardDlg*)whu_ParentDlg)->whu_GetContractData(whu_NumberBuf);
		for (int i=0;i<m_FaxContactList.GetItemCount();i++)
		{
			if (m_Num == m_FaxContactList.GetItemText(i,7)&&m_data.Name ==m_FaxContactList.GetItemText(i,1))
			{
				if (fax_state == FAX_SENDING)
				{
					m_FaxContactList.SetItemText(i,9,_T("正在发送..."));
				}
				else if (fax_state == FAX_FAIL) //&&m_FaxContactList.GetItemText(i,9)==_T("正在发送...")
				{
					m_FaxContactList.SetItemText(i,9,_T("发送失败"));
				}
				else if (fax_state == FAX_SUCCESS&&m_FaxContactList.GetItemText(i,9)==_T("正在发送..."))
				{

					m_FaxContactList.SetItemText(i,9,_T("发送成功"));

				}
				else if (fax_state == FAX_PREPARE)
				{
					m_FaxContactList.SetItemText(i,9,_T("准备就绪"));
				}
				else if (fax_state == FAX_CANCEL)
				{
					m_FaxContactList.SetItemText(i,9,_T("任务取消"));
				}else if (fax_state == FAX_NOANSWER)
				{
					m_FaxContactList.SetItemText(i,9,_T("对方无应答"));
				}else if (fax_state == FAX_NONE)
				{
					m_FaxContactList.SetItemText(i,9,_T(""));
				}

			}
		}

	}
	

}

void Cwhu_FaxDlg::OnTimer(UINT_PTR nIDEvent)
{
	// TODO: Add your message handler code here and/or call default
	//whu_ShowFaxState();
	////////////////////////////新加模块//////////////////////////////////////////////
	bool m_SendFaxDone = false;
	if (NumCount == whu_NumHaveSend && NumCount!=0) //说明当前任务已经完成了//
	{
		SetDlgItemText(IDC_SENDOK,_T("发送"));
	}
	//////////////////////////新加模块/////////////////////////////////////////////////
	CDialogEx::OnTimer(nIDEvent);
}


void Cwhu_FaxDlg::OnLvnColumnclickFaxcontactlist(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMLISTVIEW pNMLV = reinterpret_cast<LPNMLISTVIEW>(pNMHDR);
	// TODO: Add your control notification handler code here
	int colum = pNMLV->iSubItem;
	static int mode = 0;
	whu_SortContactList(m_FaxContactList,colum,mode);
	if (mode == 0)
	{
		mode = 1;
	}
	else{
		mode = 0;
	}

	*pResult = 0;
}
void Cwhu_FaxDlg::whu_SortContactList(CListCtrl &m_List,int colum,int mode)
{
	//m_ContactListData[colum]
	//对字符进行排序//
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
			if (mode == 0) //正序
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
			else if (mode == 1)//反序
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
	//插入数据//
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

void Cwhu_FaxDlg::Lin_ImportArrToList(CStringArray *m_DutyArr,int Num,CListCtrl &m_List)
{
	CStringArray m_Head;
	for (int i=0;i<Num;i++)
	{
		m_Head.Add(m_DutyArr[i].GetAt(0));
	}
	m_Head.Add(_T("状态"));
	Lin_InitList(m_List,m_Head);
	for (int i=1;i<m_DutyArr[0].GetSize();i++)
	{
		CStringArray m_ItemArr;
		for (int j=0;j<Num;j++)
		{
			m_ItemArr.Add(m_DutyArr[j].GetAt(i));
		}
		Lin_InsertList(m_List,m_ItemArr);
	}
}
void Cwhu_FaxDlg::Lin_InitList(CListCtrl &m_List,CStringArray &ColumName)
{

	LONG lStyle; 
	lStyle = GetWindowLong(m_List.m_hWnd, GWL_STYLE); 
	lStyle &= ~LVS_TYPEMASK;
	lStyle |= LVS_REPORT; 
	SetWindowLong(m_List.m_hWnd, GWL_STYLE, lStyle);
	DWORD dwStyle = m_List.GetExtendedStyle(); 
	dwStyle |= LVS_EX_FULLROWSELECT;
	dwStyle |= LVS_EX_GRIDLINES;
	dwStyle |= LVS_EX_CHECKBOXES;
	m_List.SetExtendedStyle(dwStyle); 
	m_List.InsertColumn(0,_T("编号"),LVCFMT_LEFT,100);
	for (int i=0;i<ColumName.GetSize();i++)
	{
		m_List.InsertColumn(i+1,ColumName.GetAt(i),LVCFMT_LEFT,100);
	}

}
void Cwhu_FaxDlg::Lin_InsertList(CListCtrl &m_List,CStringArray &ItemArr)
{
	int ArrNum = ItemArr.GetSize();
	if (ArrNum < m_List.GetHeaderCtrl()->GetItemCount()-1)
	{
		return;
	}
	int pos = m_List.GetItemCount();
	CString num;
	num.Format(_T("%d"),pos+1);
	m_List.InsertItem(pos,num);
	m_List.SetItemText(pos,0,num);
	for (int colum=0;colum<m_List.GetHeaderCtrl()->GetItemCount()-1;colum++)
	{
		m_List.SetItemText(pos,colum+1,ItemArr.GetAt(colum));
	}
}
void Cwhu_FaxDlg::whu_ReSize()
{
	POINT Old;
	int cx,cy; 
	cx = GetSystemMetrics(SM_CXSCREEN); 
	cy = GetSystemMetrics(SM_CYSCREEN);

	CRect rcTemp; 
	rcTemp.BottomRight() = CPoint(cx*11/12, cy*11/12); 
	rcTemp.TopLeft() = CPoint(cx/12, cy/12); 

	CRect rect;  
	GetClientRect(&rect); //取客户区大小  
	Old.x=rect.right-rect.left;
	Old.y=rect.bottom-rect.top;
	MoveWindow(&rcTemp);

	float fsp[2];
	POINT Newp; //获取现在对话框的大小
	CRect recta;  
	GetClientRect(&recta); //取客户区大小//  
	Newp.x=recta.right-recta.left;
	Newp.y=recta.bottom-recta.top;
	fsp[0]=(float)Newp.x/Old.x;
	fsp[1]=(float)Newp.y/Old.y;
	CRect Rect;
	int woc;
	CPoint OldTLPoint,TLPoint; //左上角//
	CPoint OldBRPoint,BRPoint; //右下角//
	HWND hwndChild=::GetWindow(m_hWnd,GW_CHILD); //列出所有控件  
	while(hwndChild)  
	{  
		woc=::GetDlgCtrlID(hwndChild);//取得ID
		GetDlgItem(woc)->GetWindowRect(Rect);  
		ScreenToClient(Rect);  
		OldTLPoint = Rect.TopLeft();  
		TLPoint.x = long(OldTLPoint.x*fsp[0]);  
		TLPoint.y = long(OldTLPoint.y*fsp[1]);  
		OldBRPoint = Rect.BottomRight();  
		BRPoint.x = long(OldBRPoint.x *fsp[0]);  
		BRPoint.y = long(OldBRPoint.y *fsp[1]); //高度不可读的控件（如:combBox),不要改变此值.//
		Rect.SetRect(TLPoint,BRPoint);  
		GetDlgItem(woc)->MoveWindow(Rect,TRUE);
		hwndChild=::GetWindow(hwndChild, GW_HWNDNEXT);  
	}
	Old=Newp;
	//下面设置字体//
	//CFont m_editFont; 
	//m_editFont.CreatePointFont(180,_T("宋体"));
	//GetDlgItem()->SetFont(&m_editFont);
	//GetDlgItem()->SetFont(&m_editFont);
}

void Cwhu_FaxDlg::OnNMClickFaxcontactlist(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMITEMACTIVATE pNMItemActivate = reinterpret_cast<LPNMITEMACTIVATE>(pNMHDR);
	// TODO: Add your control notification handler code here
	UpdateData(true);
	if (whu_DocFile!=_T(""))
	{
		GetDlgItem(IDC_SENDOK)->EnableWindow(true);
	}
	
	*pResult = 0;
}
LRESULT Cwhu_FaxDlg::OnUpdateData(WPARAM wParam, LPARAM lParam)
{
	int* m_FaxState = (int*)lParam;
	whu_ShowFaxState(*m_FaxState);
	UpdateData(FALSE);
	return true;
}