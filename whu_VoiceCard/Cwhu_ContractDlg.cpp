// Cwhu_ContractDlg.cpp : implementation file
//

#include "stdafx.h"
#include "whu_VoiceCard.h"
#include "Cwhu_ContractDlg.h"
#include "afxdialogex.h"
#include "whu_VoiceCardDlg.h"
#include  <string>
using namespace std;
// Cwhu_ContractDlg dialog

IMPLEMENT_DYNAMIC(Cwhu_ContractDlg, CDialogEx)

//清除空格
//string trim(string &s)
//{
//	const string &space =" \f\n\t\r\v" ;
//	string r=s.erase(s.find_last_not_of(space)+1);
//	return r.erase(0,r.find_first_not_of(space));
//}

Cwhu_ContractDlg::Cwhu_ContractDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(Cwhu_ContractDlg::IDD, pParent)
{

}

Cwhu_ContractDlg::~Cwhu_ContractDlg()
{
}

void Cwhu_ContractDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_LISCONTRACT, m_List);
}
//void Cwhu_ContractDlg::whu_InitList()
//{
//	LONG lStyle; 
//	lStyle = GetWindowLong(m_List.m_hWnd, GWL_STYLE);
//	lStyle &= ~LVS_TYPEMASK; 
//	lStyle |= LVS_REPORT; 
//	SetWindowLong(m_List.m_hWnd, GWL_STYLE, lStyle);
//	DWORD dwStyle = m_List.GetExtendedStyle(); 
//	dwStyle |= LVS_EX_FULLROWSELECT;
//	dwStyle |= LVS_EX_GRIDLINES;
//	dwStyle |= LVS_EX_CHECKBOXES;
//	m_List.SetExtendedStyle(dwStyle); 
//	m_List.InsertColumn( 0, "ID", LVCFMT_LEFT, 60 );
//	m_List.InsertColumn( 1, _T("姓名"), LVCFMT_LEFT, 120 ); 
//	m_List.InsertColumn( 2, _T("单位"), LVCFMT_LEFT, 120 );
//	m_List.InsertColumn( 3, _T("部门"), LVCFMT_LEFT, 160 );
//	m_List.InsertColumn( 4, _T("职务"), LVCFMT_LEFT, 80 );
//	m_List.InsertColumn( 5, _T("手机"), LVCFMT_LEFT, 120 ); 
//	m_List.InsertColumn( 6, _T("办公室电话"), LVCFMT_LEFT, 120 ); 
//	m_List.InsertColumn( 7, _T("传真号"), LVCFMT_LEFT, 160 ); 
//	m_List.InsertColumn( 8, _T("邮箱"), LVCFMT_LEFT, 160 ); 
//	int num = ((Cwhu_VoiceCardDlg*)whu_ParentDlg)->m_ContactListData[0].GetSize();
//	for (int i=0;i<num;i++)
//	{
//		char buf[3];
//		itoa(i+1,buf,10);
//		m_List.InsertItem(i,buf);
//		for (int j =0;j<9;j++)
//		{
//			m_List.SetItemText(i,j,((Cwhu_VoiceCardDlg*)whu_ParentDlg)->m_ContactListData[j].GetAt(i)); 
//		}
//	}
//}

BEGIN_MESSAGE_MAP(Cwhu_ContractDlg, CDialogEx)
	ON_NOTIFY(NM_RCLICK, IDC_LISCONTRACT, &Cwhu_ContractDlg::OnNMRClickLiscontract)
	//ON_COMMAND(ID_DIAL, &Cwhu_ContractDlg::OnDial)
	//ON_COMMAND(ID_SENDNOTICE, &Cwhu_ContractDlg::OnSendnotice)
	//ON_COMMAND(ID_SENDFAX, &Cwhu_ContractDlg::OnSendfax)
	//ON_COMMAND(ID_DIALOFFICE, &Cwhu_ContractDlg::OnDialoffice)
	//ON_COMMAND(ID_SENDGROUPFAX, &Cwhu_ContractDlg::OnSendgroupfax)
	ON_COMMAND(ID_32790, &Cwhu_ContractDlg::OnMyDial)
	ON_COMMAND(ID_32791, &Cwhu_ContractDlg::OnMySendSingleNotice)
	ON_COMMAND(ID_32797, &Cwhu_ContractDlg::OnMyDialOffice)
	ON_COMMAND(ID_32798, &Cwhu_ContractDlg::OnMySendSingleNoticeOffice)
	ON_BN_CLICKED(IDC_BUTTONsave, &Cwhu_ContractDlg::OnBnClickedButtonsave)
	ON_BN_CLICKED(IDC_BUTTONBACK, &Cwhu_ContractDlg::OnBnClickedButtonback)
	ON_NOTIFY(NM_CLICK, IDC_LISCONTRACT, &Cwhu_ContractDlg::OnNMClickLiscontract)
	ON_BN_CLICKED(IDC_BUTTONADD, &Cwhu_ContractDlg::OnBnClickedButtonadd)
	ON_BN_CLICKED(IDC_BUTTONDELETE, &Cwhu_ContractDlg::OnBnClickedButtondelete)
	ON_NOTIFY(LVN_COLUMNCLICK, IDC_LISCONTRACT, &Cwhu_ContractDlg::OnLvnColumnclickLiscontract)
	ON_BN_CLICKED(IDC_IMPORT, &Cwhu_ContractDlg::OnBnClickedImport)
	ON_BN_CLICKED(IDC_SHOWPHOTO, &Cwhu_ContractDlg::OnBnClickedShowphoto)
	ON_BN_CLICKED(IDC_CHANGEPIC, &Cwhu_ContractDlg::OnBnClickedChangepic)
END_MESSAGE_MAP()


// Cwhu_ContractDlg message handlers


void Cwhu_ContractDlg::OnNMRClickLiscontract(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMITEMACTIVATE pNMItemActivate = reinterpret_cast<LPNMITEMACTIVATE>(pNMHDR);
	// TODO: Add your control notification handler code here
	NM_LISTVIEW* pNMListView = (NM_LISTVIEW*)pNMHDR; 
	m_ChoosedPos = pNMListView->iItem;
	if(pNMListView->iItem != -1) 
	{ 
		CString strtemp; 
		strtemp.Format(_T("单击的是第%d行第%d列"), pNMListView->iItem, pNMListView->iSubItem); 
		LVCOLUMN lvcol; 
		char    str[256]; 
		lvcol.mask = LVCF_TEXT; 
		lvcol.pszText = str; 
		lvcol.cchTextMax = 256; 
		m_List.GetColumn(pNMListView->iSubItem,&lvcol);
		//AfxMessageBox(lvcol.pszText); 
		//lvcol.pszText为列名//
		CString m_head = lvcol.pszText;
		if (m_head == _T("手机"))
		{
			char buf[20];
			m_List.GetItemText(m_ChoosedPos,5,buf,12);
			m_TelNum.Format("%s",buf); //得到当前行的号码//
			if (m_TelNum != _T(""))
			{
				DWORD dwPos = GetMessagePos(); 
				CPoint point( LOWORD(dwPos), HIWORD(dwPos) ); 
				CMenu menu; 
				VERIFY( menu.LoadMenu(IDR_MENUTEL) ); 
				CMenu* popup = menu.GetSubMenu(0); 
				ASSERT( popup != NULL ); 
				popup->TrackPopupMenu(TPM_LEFTALIGN | TPM_RIGHTBUTTON,point.x, point.y, this ); 
			}
		}
		if (m_head == _T("办公室电话"))
		{
			char buf[20];
			m_List.GetItemText(m_ChoosedPos,6,buf,12);
			m_OfficeNum.Format("%s",buf);
			if (m_OfficeNum !=_T(""))
			{
				DWORD dwPos = GetMessagePos(); 
				CPoint point( LOWORD(dwPos), HIWORD(dwPos) ); 
				CMenu menu; 
				VERIFY( menu.LoadMenu(IDR_MENUOFFICE) ); 
				CMenu* popup = menu.GetSubMenu(0); 
				ASSERT( popup != NULL ); 
				popup->TrackPopupMenu(TPM_LEFTALIGN | TPM_RIGHTBUTTON,point.x, point.y, this ); 
			}
		}
		if (m_head == _T("传真号"))
		{
			char buf[20];
			m_List.GetItemText(m_ChoosedPos,7,buf,12);
			m_FaxNum.Format("%s",buf);
			if (m_FaxNum != _T(""))
			{
				DWORD dwPos = GetMessagePos(); 
				CPoint point( LOWORD(dwPos), HIWORD(dwPos) ); 
				CMenu menu; 
				VERIFY( menu.LoadMenu(IDR_MENUFAX) ); 
				CMenu* popup = menu.GetSubMenu(0); 
				ASSERT( popup != NULL ); 
				popup->TrackPopupMenu(TPM_LEFTALIGN | TPM_RIGHTBUTTON,point.x, point.y, this ); 
			}
		}
	} 
	
	*pResult = 0;
}


void Cwhu_ContractDlg::OnDial()
{

}


void Cwhu_ContractDlg::OnSendnotice()
{
	int a=0;
	for(int i=0; i<m_List.GetItemCount(); i++) 
	{
		int m_int = m_List.GetItemState(i, LVIS_SELECTED);
		bool m_b = m_List.GetCheck(i);
		if( m_int == LVIS_SELECTED|| m_b )  //有选中复选框群发//
		{ 
			char buf[20]=_T("");
			m_List.GetItemText(i,5,buf,12);
			m_GroupNoticeNum[a].Format("%s",buf);
			if (m_GroupNoticeNum[a] == _T(""))
			{
				m_List.GetItemText(i,7,buf,9); // 如果找不到手机号，就换办公室电话//
				m_GroupNoticeNum[a].Format("%s",buf);
				if (m_GroupNoticeNum[a] == _T(""))
				{
					CString m_str;
					m_str.Format(_T("%s%d%s"),_T("第"),i+1,_T("行未存储手机和办公室号码"));
					AfxMessageBox(m_str);
				}
				else
				{
					a++;
				}

			}else
			{
				a++;
			}

		}
	}

	if(a==0){    //无选中复选框//当前所在行号码
		m_GroupNoticeNum[0] = m_List.GetItemText(m_ChoosedPos,5);
		if (m_GroupNoticeNum[0] ==_T(""))
		{
			m_GroupNoticeNum[0] = m_List.GetItemText(m_ChoosedPos,7);
			if (m_GroupNoticeNum[0] ==_T(""))
			{
				CString m_str;
				m_str.Format(_T("%s%d%s"),_T("第"),m_ChoosedPos+1,_T("行未存储手机和办公室号码"));
				AfxMessageBox(m_str);
			}
			else{
				a=1;
			}
		}else{
			a=1;
		}
	}
	if (a>0)
	{
		((Cwhu_VoiceCardDlg*)whu_ParentDlg)->whu_SendNotice(m_GroupNoticeNum,a);
	}
}


void Cwhu_ContractDlg::OnSendfax()
{
	// TODO: Add your command handler code here
	if (m_FaxNum!=_T(""))
	{
		CString whu_DocFile = _T("");
		CFileDialog cf(1, NULL, NULL, OFN_HIDEREADONLY|OFN_OVERWRITEPROMPT, "Fax File(*.tif;*.doc;*.docx;*.xls;*.xlsx;*.ppt;*.pptx;*.pdf;*.txt;*.html;*.jpg;*bmp)|*.tif;*.doc;*.docx;*.xls;*.xlsx;*.ppt;*.pptx;*.pdf;*.txt;*.html;*.jpg;*bmp|Tif File(*.fax)|*.fax|Tiff File(*.tiff)|*.tiff|All File(*.*)|*.*|",NULL);//可以打开doc、pdf等文件
		if(cf.DoModal() == IDOK)
		{
			whu_DocFile=cf.GetPathName();
			((Cwhu_VoiceCardDlg*)whu_ParentDlg)->whu_SendFax(&m_FaxNum,whu_DocFile,1);
		}
		
	} 
	else
	{
		AfxMessageBox(_T("未存储传真号码"));
	}

}


void Cwhu_ContractDlg::OnDialoffice()
{
	// TODO: Add your command handler code here

}


void Cwhu_ContractDlg::OnSendgroupfax()
{
	// TODO: Add your command handler code here

	int a = 0;
	CString whu_DocFile = _T("");
	CFileDialog cf(1, NULL, NULL, OFN_HIDEREADONLY|OFN_OVERWRITEPROMPT, "Fax File(*.tif;*.doc;*.docx;*.xls;*.xlsx;*.ppt;*.pptx;*.pdf;*.txt;*.html;*.jpg;*bmp)|*.tif;*.doc;*.docx;*.xls;*.xlsx;*.ppt;*.pptx;*.pdf;*.txt;*.html;*.jpg;*bmp|Tif File(*.fax)|*.fax|Tiff File(*.tiff)|*.tiff|All File(*.*)|*.*|",NULL);//可以打开doc、pdf等文件
	if(cf.DoModal() == IDOK)
	{
		whu_DocFile=cf.GetPathName();
	}
	for(int i=0; i<m_List.GetItemCount(); i++) 
	{ 
		if( m_List.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED|| m_List.GetCheck(i) ) 
		{ 
			char buf[20]=_T("");
			m_List.GetItemText(i,8,buf,9);
			m_GroupFaxNum[a].Format("%s",buf);
			if (m_GroupFaxNum[a] == _T(""))
			{
				CString m_str;
				m_str.Format(_T("%s%d%s"),_T("第"),i+1,_T("行未存储传真号码"));
				AfxMessageBox(m_str);
			}else
			{
				a++;
			}
			
		} 
	}
	if (a>0)
	{
		((Cwhu_VoiceCardDlg*)whu_ParentDlg)->whu_SendFax(m_GroupFaxNum,whu_DocFile,a);
	}

}
CString Cwhu_ContractDlg::GetStringFromVariant(_variant_t var)
{
	return var.vt==VT_NULL?"":(char*)(_bstr_t)var;
}

void Cwhu_ContractDlg::OnMyDial() //手机
{	
	if (m_TelNum !=_T("")){
		string str = m_TelNum.GetBuffer(20);
		//str=trim((string)str);
		m_TelNum = str.c_str();
		((Cwhu_VoiceCardDlg*)whu_ParentDlg)->whu_Dial(m_TelNum);
	} 
	else{
		AfxMessageBox(_T("未存储手机号码"));
	}
}


void Cwhu_ContractDlg::OnMySendSingleNotice() //手机//
{
	if (m_TelNum !=_T("")){
		string str = m_TelNum.GetBuffer(20);
		//str=trim((string)str);
		m_TelNum = str.c_str();
		((Cwhu_VoiceCardDlg*)whu_ParentDlg)->whu_SendNotice(&m_TelNum,1);
	} 
	else{
		AfxMessageBox(_T("未存储手机号码"));
	}
	
}


void Cwhu_ContractDlg::OnMyDialOffice()		//办公室电话//
{
	// TODO: Add your command handler code here
	if (m_OfficeNum!=_T(""))
	{
		((Cwhu_VoiceCardDlg*)whu_ParentDlg)->whu_Dial(m_OfficeNum);
	} 
	else 
	{
		AfxMessageBox(_T("未存储办公室电话"));
	}
}

void Cwhu_ContractDlg::OnMySendSingleNoticeOffice()   //办公室电话//
{
	// TODO: Add your command handler code here
	if (m_OfficeNum !=_T("")){
		string str = m_OfficeNum.GetBuffer(20);
		//str=trim((string)str);
		m_OfficeNum = str.c_str();
		((Cwhu_VoiceCardDlg*)whu_ParentDlg)->whu_SendNotice(&m_OfficeNum,1);
	} 
	else{
		AfxMessageBox(_T("未存储办公室号码"));
	}
}

BOOL Cwhu_ContractDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();
	//whu_InitList();
	ConHeadNum = ((Cwhu_VoiceCardDlg*)whu_ParentDlg)->ConHeadNum;
	Lin_ImportArrToList(((Cwhu_VoiceCardDlg*)whu_ParentDlg)->m_ContactListData,ConHeadNum,m_List);
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
	GetDlgItem(IDC_CHANGEPIC)->EnableWindow(false);
	GetDlgItem(IDC_SHOWPHOTO)->EnableWindow(false);
	return true;
}
void Cwhu_ContractDlg::OnBnClickedButtonsave()
{
	// TODO: Add your control notification handler code here
	CString m_Str;
	if (m_ChoosedPos <m_List.GetItemCount())
	{
		GetDlgItemText(IDC_EDITNAME,m_Str);
		m_List.SetItemText(m_ChoosedPos,1,m_Str);
		GetDlgItemText(IDC_EDITUNIT,m_Str);
		m_List.SetItemText(m_ChoosedPos,2,m_Str);
		GetDlgItemText(IDC_EDITDEPARTMENT,m_Str);
		m_List.SetItemText(m_ChoosedPos,3,m_Str);
		GetDlgItemText(IDC_EDITJOB,m_Str);
		m_List.SetItemText(m_ChoosedPos,4,m_Str);
		GetDlgItemText(IDC_EDITTELPHONE,m_Str);
		m_List.SetItemText(m_ChoosedPos,5,m_Str);
		GetDlgItemText(IDC_EDITOFFICEPHONE,m_Str);
		m_List.SetItemText(m_ChoosedPos,6,m_Str);
		GetDlgItemText(IDC_EDITFAX,m_Str);
		m_List.SetItemText(m_ChoosedPos,7,m_Str);
		GetDlgItemText(IDC_EDITEMAIL,m_Str);
		m_List.SetItemText(m_ChoosedPos,8,m_Str);
	}
	
	

}


void Cwhu_ContractDlg::OnBnClickedButtonback()
{
	// TODO: Add your control notification handler code here
	//whu_SaveContact();
	/*CString m_FilePath = _T("C:\\Whu_Data\\Excel文件\\通讯录.xlsx");
	CString m_Folder = _T("C:\\Whu_Data\\Excel文件");
	if (whu_findfile(m_FilePath,m_Folder)==true)
	{
		CFile m_Temp;
		m_Temp.Remove(m_FilePath);
	}
	Lin_ExportListToExcel(m_List,m_FilePath);*/
	((Cwhu_VoiceCardDlg*)whu_ParentDlg)->ConHeadNum = Lin_ListToArr(m_List,((Cwhu_VoiceCardDlg*)whu_ParentDlg)->m_ContactListData);
	ConHeadNum = ((Cwhu_VoiceCardDlg*)whu_ParentDlg)->ConHeadNum;
	Lin_ImportArrToList(((Cwhu_VoiceCardDlg*)whu_ParentDlg)->m_ContactListData,ConHeadNum,((Cwhu_VoiceCardDlg*)whu_ParentDlg)->m_MyContactList); //这个有问题//下午改
	CDialogEx::OnCancel();
}


void Cwhu_ContractDlg::OnNMClickLiscontract(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMITEMACTIVATE pNMItemActivate = reinterpret_cast<LPNMITEMACTIVATE>(pNMHDR);
	// TODO: Add your control notification handler code here
	NM_LISTVIEW* pNMListView = (NM_LISTVIEW*)pNMHDR;
	m_ChoosedPos = pNMListView->iItem;
	if(pNMListView->iItem != -1) 
	{ 
		CString m_Str = m_List.GetItemText(m_ChoosedPos,1);
		SetDlgItemText(IDC_EDITNAME,m_Str);
		CString m_folder = _T("C:\\Whu_Data\\联系人头像");
		CString m_file = m_folder + _T("\\") + m_Str + _T(".jpg");
		if(whu_findfile(m_file,m_folder) == true)
		{
			SetDlgItemText(IDC_PHOTO,m_Str + _T(".jpg"));
			SetDlgItemText(IDC_SHOWPHOTO,_T("查看"));
			GetDlgItem(IDC_SHOWPHOTO)->EnableWindow(true);
			GetDlgItem(IDC_CHANGEPIC)->EnableWindow(true);
		}else{
			SetDlgItemText(IDC_PHOTO,_T("无头像"));
			GetDlgItem(IDC_CHANGEPIC)->EnableWindow(true);
			GetDlgItem(IDC_SHOWPHOTO)->EnableWindow(false);

		}
		m_Str = m_List.GetItemText(m_ChoosedPos,2);
		SetDlgItemText(IDC_EDITUNIT,m_Str);
		m_Str = m_List.GetItemText(m_ChoosedPos,3);
		SetDlgItemText(IDC_EDITDEPARTMENT,m_Str);
		m_Str = m_List.GetItemText(m_ChoosedPos,4);
		SetDlgItemText(IDC_EDITJOB,m_Str);
		m_Str = m_List.GetItemText(m_ChoosedPos,5);
		SetDlgItemText(IDC_EDITTELPHONE,m_Str);
		m_Str = m_List.GetItemText(m_ChoosedPos,6);
		SetDlgItemText(IDC_EDITOFFICEPHONE,m_Str);
		m_Str = m_List.GetItemText(m_ChoosedPos,7);
		SetDlgItemText(IDC_EDITFAX,m_Str);
		m_Str = m_List.GetItemText(m_ChoosedPos,8);
		SetDlgItemText(IDC_EDITEMAIL,m_Str);



	}
	*pResult = 0;
}


void Cwhu_ContractDlg::OnBnClickedButtonadd()
{
	// TODO: Add your control notification handler code here
	UpdateData(true);
	int pos = m_List.GetItemCount();
	char buf[3];
	itoa(pos,buf,10);
	CString m_Str;
	m_Str.Format(_T("%d"),pos+1);
	m_List.InsertItem(pos,m_Str);
	m_List.SetItemText(pos,0,m_Str);
	GetDlgItemText(IDC_EDITNAME,m_Str);
	m_List.SetItemText(pos,1,m_Str);
	GetDlgItemText(IDC_EDITUNIT,m_Str);
	m_List.SetItemText(pos,2,m_Str);
	GetDlgItemText(IDC_EDITDEPARTMENT,m_Str);
	m_List.SetItemText(pos,3,m_Str);
	GetDlgItemText(IDC_EDITJOB,m_Str);
	m_List.SetItemText(pos,4,m_Str);
	GetDlgItemText(IDC_EDITTELPHONE,m_Str);
	m_List.SetItemText(pos,5,m_Str);
	GetDlgItemText(IDC_EDITOFFICEPHONE,m_Str);
	m_List.SetItemText(pos,6,m_Str);
	GetDlgItemText(IDC_EDITFAX,m_Str);
	m_List.SetItemText(pos,7,m_Str);
	GetDlgItemText(IDC_EDITEMAIL,m_Str);
	m_List.SetItemText(pos,8,m_Str);
}


void Cwhu_ContractDlg::OnBnClickedButtondelete()
{
	// TODO: Add your control notification handler code here
	bool m_Check = false;
	for (int i=m_List.GetItemCount()-1;i>=0;i--)
	{
		bool m_Check = m_List.GetCheck(i);
		if (m_Check == true)
		{
			m_List.DeleteItem(i);
		}
	}
	if (m_Check ==false)
	{
		m_List.DeleteItem(m_ChoosedPos);
	}
	int num = m_List.GetItemCount();
	for (int i=0;i<num;i++)
	{
		CString m_Str;
		m_Str.Format(_T("%d"),i+1);
		m_List.SetItemText(i,0,m_Str);
	}
}
void Cwhu_ContractDlg::whu_SaveContact()
{
	
	int length = ((Cwhu_VoiceCardDlg*)whu_ParentDlg)->m_ContactListData[0].GetSize();
	for (int i=0;i<ConHeadNum;i++)
	{
		((Cwhu_VoiceCardDlg*)whu_ParentDlg)->m_ContactListData[i].RemoveAt(1,length-1);
	}
	int test = ((Cwhu_VoiceCardDlg*)whu_ParentDlg)->m_ContactListData[0].GetSize();
	int colum = m_List.GetHeaderCtrl()->GetItemCount();
	int item = m_List.GetItemCount();
	((Cwhu_VoiceCardDlg*)whu_ParentDlg)->m_MyContactList.DeleteAllItems();
	for (int i=0;i<item;i++)
	{
		CString m_Str;
		m_Str.Format(_T("%d"),i+1);
		((Cwhu_VoiceCardDlg*)whu_ParentDlg)->m_MyContactList.InsertItem(i,m_Str);
		((Cwhu_VoiceCardDlg*)whu_ParentDlg)->m_MyContactList.SetItemText(i,0,m_Str);
		for (int j=1;j<colum;j++)
		{
			CString m_Str = m_List.GetItemText(i,j);
			((Cwhu_VoiceCardDlg*)whu_ParentDlg)->m_ContactListData[j-1].InsertAt(i+1,m_Str);
			((Cwhu_VoiceCardDlg*)whu_ParentDlg)->m_MyContactList.SetItemText(i,j,m_Str);
		}
	}
	//for (int i=0;i<item;i++)
	//{
	//	CString m_Str;
	//	m_Str.Format(_T("%d"),i+1);
	//	((Cwhu_VoiceCardDlg*)whu_ParentDlg)->m_MyContactList.InsertItem(i,m_Str);
	//	((Cwhu_VoiceCardDlg*)whu_ParentDlg)->m_MyContactList.SetItemText(i,0,m_Str);
	//	for (int j=1;j<colum;j++)
	//	{
	//		CString m_Str = m_List.GetItemText(i,j);
	//		((Cwhu_VoiceCardDlg*)whu_ParentDlg)->m_ContactListData[j-1].Add(m_Str);
	//		((Cwhu_VoiceCardDlg*)whu_ParentDlg)->m_MyContactList.SetItemText(i,j,m_Str);
	//	}
	//}

}


void Cwhu_ContractDlg::OnLvnColumnclickLiscontract(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMLISTVIEW pNMLV = reinterpret_cast<LPNMLISTVIEW>(pNMHDR);
	// TODO: Add your control notification handler code here
	int colum = pNMLV->iSubItem;
	static int mode = 0;
	whu_SortContactList(m_List,colum,mode);
	if (mode == 0)
	{
		mode = 1;
	}
	else{
		mode = 0;
	}
	*pResult = 0;
}
void Cwhu_ContractDlg::whu_SortContactList(CListCtrl &m_List,int colum,int mode)
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

void Cwhu_ContractDlg::OnBnClickedImport()
{
	// TODO: Add your control notification handler code here
	CString m_pathname;
	CFileDialog cf(1, NULL, NULL, OFN_HIDEREADONLY|OFN_OVERWRITEPROMPT, "Excel File(*.xlsx;*.xls)|*.xlsx;*.xls",NULL);//
	if(cf.DoModal() == IDOK)
	{
		CString m_DocFile=cf.GetPathName();
		Lin_InportExcelToList(m_DocFile,m_List);
		//((Cwhu_VoiceCardDlg*)whu_ParentDlg)->ConHeadNum = Lin_ListToArr(m_List,((Cwhu_VoiceCardDlg*)whu_ParentDlg)->m_ContactListData);
		//Lin_ImportArrToList(((Cwhu_VoiceCardDlg*)whu_ParentDlg)->m_ContactListData,((Cwhu_VoiceCardDlg*)whu_ParentDlg)->ConHeadNum,((Cwhu_VoiceCardDlg*)whu_ParentDlg)->m_MyContactList);
	}
}


CString Cwhu_ContractDlg::Lin_GetEnglishCharacter(int Num)
{
	CString str2;
	switch(Num){
	case 1:
		str2.Format(_T("A"));
		break;
	case 2:
		str2.Format(_T("B"));
		break;
	case 3:
		str2.Format(_T("C"));
		break;
	case 4:
		str2.Format(_T("D"));
		break;
	case 5:
		str2.Format(_T("E"));
		break;
	case 6:
		str2.Format(_T("F"));
		break;
	case 7:
		str2.Format(_T("G"));
		break;
	case 8:
		str2.Format(_T("H"));
		break;
	case 9:
		str2.Format(_T("I"));
		break;
	case 10:
		str2.Format(_T("J"));
		break;
	case 11:
		str2.Format(_T("K"));
		break;
	case 12:
		str2.Format(_T("L"));
		break;
	case 13:
		str2.Format(_T("M"));
		break;
	case 14:
		str2.Format(_T("N"));
		break;
	case 15:
		str2.Format(_T("O"));
		break;
	case 16:
		str2.Format(_T("P"));
		break;
	case 17:
		str2.Format(_T("Q"));
		break;
	case 18:
		str2.Format(_T("R"));
		break;
	case 19:
		str2.Format(_T("S"));
		break;
	case 20:
		str2.Format(_T("T"));
		break;
	case 21:
		str2.Format(_T("U"));
		break;
	case 22:
		str2.Format(_T("V"));
		break;
	case 23:
		str2.Format(_T("W"));
		break;
	case 24:
		str2.Format(_T("X"));
		break;
	case 25:
		str2.Format(_T("Y"));
		break;
	case 26:
		str2.Format(_T("Z"));
		break;
	default:
		str2.Format(_T("Z"));
		break;
	}
	return str2;
}
bool Cwhu_ContractDlg::Lin_InportExcelToList(CString m_FilePath,CListCtrl &m_List)
{
	CStringArray m_StandarHeadName;
	m_StandarHeadName.Add(_T("姓名"));
	m_StandarHeadName.Add(_T("单位"));
	m_StandarHeadName.Add(_T("部门"));
	m_StandarHeadName.Add(_T("职务"));
	m_StandarHeadName.Add(_T("手机"));
	m_StandarHeadName.Add(_T("办公电话"));
	m_StandarHeadName.Add(_T("传真号"));
	m_StandarHeadName.Add(_T("邮箱"));
	//导入
	CApplication app;
	CWorkbook book;
	CWorkbooks books;
	CWorksheet sheet;
	CWorksheets sheets;
	CRange range;
	LPDISPATCH lpDisp;
	//定义变量//
	COleVariant covOptional((long)
		DISP_E_PARAMNOTFOUND,VT_ERROR);
	if (!app.CreateDispatch(_T("Excel.Application")))
	{
		this->MessageBox(_T("无法创建Excel应用"));
		return false;
	}
	books = app.get_Workbooks();
	//打开Excel，其中m_FilePath为Excel表的路径名//
	lpDisp = books.Open(m_FilePath,covOptional
		,covOptional,covOptional,covOptional
		,covOptional,covOptional,covOptional
		,covOptional,covOptional,covOptional
		,covOptional,covOptional,covOptional
		,covOptional);
	book.AttachDispatch(lpDisp);
	sheets = book.get_Worksheets();
	sheet = sheets.get_Item(COleVariant((short)1));
	CStringArray m_HeadName;
	for (int i=1;i<26;i++)
	{
		CString m_pos = Lin_GetEnglishCharacter(i);
		m_pos = m_pos + _T("1");
		range = sheet.get_Range(COleVariant(m_pos),COleVariant(m_pos));
		//获得单元格的内容
		COleVariant rValue;
		rValue = COleVariant(range.get_Value2());
		//转换成宽字符//
		rValue.ChangeType(VT_BSTR);
		//转换格式，并输出//
		CString m_content = CString(rValue.bstrVal);
		if (m_content!=_T("") )
		{
			if (m_content != m_StandarHeadName.GetAt(i-1))
			{
				AfxMessageBox(_T("通讯录不符合模板要求"));
				book.put_Saved(TRUE);
				app.Quit();
				return false;//说明不符合模板要求//
			}else{
				m_HeadName.Add(m_content);
			}	
		}
	}
	//先删除列表内容//   
	m_List.DeleteAllItems();
	while(m_List.GetHeaderCtrl()->GetItemCount()>0)
	{
		m_List.DeleteColumn(0);
	}
	Lin_InitList(m_List,m_HeadName);	
	CStringArray m_ContentArr;
	for (int ItemNum = 0;ItemNum<1024;ItemNum++)
	{
		for (int j=1;j<m_List.GetHeaderCtrl()->GetItemCount();j++)
		{
			CString m_pos = Lin_GetEnglishCharacter(j);
			CString m_Itempos;
			m_Itempos.Format(_T("%d"),ItemNum+2);
			CString m_str = m_pos+m_Itempos;
			range = sheet.get_Range(COleVariant(m_str),COleVariant(m_str));
			//获得单元格的内容
			COleVariant rValue;
			rValue = COleVariant(range.get_Value2());
			//转换成宽字符//
			rValue.ChangeType(VT_BSTR);
			//转换格式，并输出//
			CString m_content = CString(rValue.bstrVal);
			m_ContentArr.Add(m_content);
		}
		if (m_ContentArr.GetAt(0)!=_T(""))
		{
			Lin_InsertList(m_List,m_ContentArr);
			m_ContentArr.RemoveAll();
		}
		else{
			break;
		}
	}
	book.put_Saved(TRUE);
	app.Quit();
	return true;

}
void Cwhu_ContractDlg::Lin_ExportListToExcel(CListCtrl &m_List,CString m_FilePath)
{
	if (m_List.GetHeaderCtrl()->GetItemCount()>0)
	{
		CApplication app;
		CWorkbooks books;
		CWorkbook book;
		CWorksheets sheets;
		CWorksheet sheet;
		CRange range;
		CRange cols;
		int i = 3;
		CString str1, str2;
		COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
		if( !app.CreateDispatch(_T("Excel.Application")))
		{
			MessageBox(_T("无法创建Excel应用！"));
			return;
		}
		books=app.get_Workbooks();
		book=books.Add(covOptional);
		sheets=book.get_Sheets();
		sheet=sheets.get_Item(COleVariant((short)1));

		//写入表头//
		for (int ColumNum=1;ColumNum<m_List.GetHeaderCtrl()->GetItemCount();ColumNum++)
		{
			str2 = Lin_GetEnglishCharacter(ColumNum);
			str2 = str2 + _T("1");
			LVCOLUMN   lvColumn;   
			TCHAR strChar[256];
			lvColumn.pszText=strChar;   
			lvColumn.cchTextMax=256 ;
			lvColumn.mask   = LVCF_TEXT;
			m_List.GetColumn(ColumNum,&lvColumn);
			CString str=lvColumn.pszText;
			range=sheet.get_Range(COleVariant(str2),COleVariant(str2));
			range.put_Value2(COleVariant(str));

		}
		//获取单元格的位置//
		for (int ColumNum =1;ColumNum<m_List.GetHeaderCtrl()->GetItemCount();ColumNum++)
		{

			for (int ItemNum = 0;ItemNum<m_List.GetItemCount();ItemNum++)
			{
				str2 = Lin_GetEnglishCharacter(ColumNum);
				str1.Format(_T("%d"),ItemNum+2);
				str2 = str2+str1;
				CString str = m_List.GetItemText(ItemNum,ColumNum);
				range=sheet.get_Range(COleVariant(str2),COleVariant(str2));
				range.put_Value2(COleVariant(str));
			}
		}
		cols=range.get_EntireColumn();
		cols.AutoFit();
		//app.put_Visible(TRUE);
		//app.put_UserControl(TRUE);
		book.SaveAs(COleVariant(m_FilePath),covOptional,covOptional,covOptional,covOptional,covOptional,0,covOptional,covOptional,covOptional,covOptional,covOptional);
		app.Quit();
	}else{
		AfxMessageBox(_T("列表为空"));
	}

}
void Cwhu_ContractDlg::Lin_InitList(CListCtrl &m_List,CStringArray &ColumName)
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
	m_List.InsertColumn(0,_T("编号"),LVCFMT_LEFT,60);
	for (int i=0;i<ColumName.GetSize();i++)
	{
		m_List.InsertColumn(i+1,ColumName.GetAt(i),LVCFMT_LEFT,150);
	}

}
void Cwhu_ContractDlg::Lin_InsertList(CListCtrl &m_List,CStringArray &ItemArr)
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
void Cwhu_ContractDlg::Lin_ImportArrToList(CStringArray *m_Arr,int Num,CListCtrl &m_List)
{
	m_List.DeleteAllItems();
	while(m_List.GetHeaderCtrl()->GetItemCount()>0)
	{
		m_List.DeleteColumn(0);
	}
	CStringArray m_Head;
	for (int i=0;i<Num;i++)
	{
		m_Head.Add(m_Arr[i].GetAt(0));
	}
	Lin_InitList(m_List,m_Head);
	for (int i=1;i<m_Arr[0].GetSize();i++)
	{
		CStringArray m_ItemArr;
		for (int j=0;j<Num;j++)
		{
			m_ItemArr.Add(m_Arr[j].GetAt(i));
		}
		Lin_InsertList(m_List,m_ItemArr);
	}
}
bool Cwhu_ContractDlg::whu_findfile(CString m_file,CString folder)
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
int Cwhu_ContractDlg::Lin_ListToArr(CListCtrl &m_List,CStringArray *m_Arr)
{
	int Num = m_List.GetHeaderCtrl()->GetItemCount()-1;
	for (int ColumNum =1;ColumNum<m_List.GetHeaderCtrl()->GetItemCount();ColumNum++)
	{
		LVCOLUMN   lvColumn;   
		TCHAR strChar[256];
		lvColumn.pszText=strChar;   
		lvColumn.cchTextMax=256 ;
		lvColumn.mask   = LVCF_TEXT;
		m_List.GetColumn(ColumNum,&lvColumn);
		CString str=lvColumn.pszText;
		m_Arr[ColumNum-1].RemoveAll();
		m_Arr[ColumNum-1].Add(str);
	}
	for (int i=0;i<m_List.GetItemCount();i++)
	{
		for (int j=1;j<m_List.GetHeaderCtrl()->GetItemCount();j++)
		{
			CString m_str = m_List.GetItemText(i,j);
			m_Arr[j-1].Add(m_str);
		}
	}
	return Num;
}

void Cwhu_ContractDlg::OnBnClickedShowphoto()
{
	// TODO: Add your control notification handler code here
	//CString m_str;
	//GetDlgItemText(IDC_PHOTO,m_str);
	//if (m_str == _T("无头像"))
	//{
	//	CString m_pathname;
	//	CFileDialog cf(1, NULL, NULL, OFN_HIDEREADONLY|OFN_OVERWRITEPROMPT, "JPG File(*.jpg)|*.jpg",NULL);//
	//	if(cf.DoModal() == IDOK)
	//	{
	//		CString m_SrcFile=cf.GetPathName();
	//		CString m_str2;
	//		GetDlgItemText(IDC_EDITNAME,m_str2);
	//		CString m_DestFile = _T("C:\\Whu_Data\\联系人头像\\") + m_str2 + _T(".jpg");
	//		CopyFile(m_SrcFile,m_DestFile,false);
	//		SetDlgItemText(IDC_PHOTO,m_str2 + _T(".jpg"));
	//		SetDlgItemText(IDC_SHOWPHOTO, _T("查看"));
	//		GetDlgItem(IDC_CHANGEPIC)->EnableWindow(true);
	//	}
	//	
	//}
	//else{
	CString m_str2;
	GetDlgItemText(IDC_EDITNAME,m_str2);
	CString m_File = _T("C:\\Whu_Data\\联系人头像\\") + m_str2 + _T(".jpg");
	ShellExecute(NULL, "open",m_File,NULL,NULL, SW_SHOWNORMAL );
	//}
	
}


void Cwhu_ContractDlg::OnBnClickedChangepic()
{
	// TODO: Add your control notification handler code here
	CString m_pathname;
	CFileDialog cf(1, NULL, NULL, OFN_HIDEREADONLY|OFN_OVERWRITEPROMPT, "JPG File(*.jpg)|*.jpg",NULL);//
	if(cf.DoModal() == IDOK)
	{
		CString m_SrcFile=cf.GetPathName();
		CString m_str2;
		GetDlgItemText(IDC_EDITNAME,m_str2);
		CString m_DestFile = _T("C:\\Whu_Data\\联系人头像\\") + m_str2 + _T(".jpg");
		CopyFile(m_SrcFile,m_DestFile,false);
		SetDlgItemText(IDC_PHOTO,m_str2 + _T(".jpg"));
		GetDlgItem(IDC_SHOWPHOTO)->EnableWindow(true);
	}
}
