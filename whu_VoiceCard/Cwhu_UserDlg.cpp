// Cwhu_UserDlg.cpp : implementation file
//

#include "stdafx.h"
#include "whu_VoiceCard.h"
#include "Cwhu_UserDlg.h"
#include "afxdialogex.h"
#include "whu_VoiceCardDlg.h"
#include "Cwhu_ChgPwdDlg.h"
// Cwhu_UserDlg dialog

IMPLEMENT_DYNAMIC(Cwhu_UserDlg, CDialogEx)

Cwhu_UserDlg::Cwhu_UserDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(Cwhu_UserDlg::IDD, pParent)
{

}

Cwhu_UserDlg::~Cwhu_UserDlg()
{
	
}

void Cwhu_UserDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_LIST1, m_UserList);
}
void Cwhu_UserDlg::whu_InitUser()
{

}
BEGIN_MESSAGE_MAP(Cwhu_UserDlg, CDialogEx)
	ON_BN_CLICKED(IDC_BACK, &Cwhu_UserDlg::OnBnClickedBack)
	ON_NOTIFY(NM_CLICK, IDC_LIST1, &Cwhu_UserDlg::OnNMClickList1)
	ON_BN_CLICKED(IDC_ACCEPT, &Cwhu_UserDlg::OnBnClickedAccept)
	ON_BN_CLICKED(IDC_REFUSE, &Cwhu_UserDlg::OnBnClickedRefuse)
	ON_BN_CLICKED(IDC_LOGOFF, &Cwhu_UserDlg::OnBnClickedLogoff)
	ON_BN_CLICKED(IDC_CHGPWD, &Cwhu_UserDlg::OnBnClickedChgpwd)
END_MESSAGE_MAP()


// Cwhu_UserDlg message handlers
BOOL Cwhu_UserDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();
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
	m_ClickPos = -1;
	//SetTimer(1,1000,NULL);
	//CString m_filepath = _T("C:\\Whu_Data\\Excel文件\\用户信息.xlsx");
	//Lin_InportExcelToList(m_filepath,m_UserList);
	Lin_ImportArrToList(((Cwhu_VoiceCardDlg*)whu_ParentDlg)->whu_User,4,m_UserList);
	GetDlgItem(IDC_ACCEPT)->EnableWindow(false);
	GetDlgItem(IDC_REFUSE)->EnableWindow(false);
	GetDlgItem(IDC_LOGOFF)->EnableWindow(false);
	GetDlgItem(IDC_CHGPWD)->EnableWindow(false);
	return true;

}
CString Cwhu_UserDlg::Lin_GetEnglishCharacter(int Num)
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
void Cwhu_UserDlg::Lin_InitList(CListCtrl &m_List,CStringArray &ColumName)
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
		m_List.InsertColumn(i+1,ColumName.GetAt(i),LVCFMT_LEFT,200);
	}

}
void Cwhu_UserDlg::Lin_InsertList(CListCtrl &m_List,CStringArray &ItemArr)
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
void Cwhu_UserDlg::Lin_InportExcelToList(CString m_FilePath,CListCtrl &m_List)
{
	//先删除列表内容//
	//m_List.DeleteAllItems();
	//while(m_List.GetHeaderCtrl()->GetItemCount()>0)
	//{
	//	m_List.DeleteColumn(0);
	//}
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
		return;
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
		if (m_content!=_T(""))
		{
			m_HeadName.Add(m_content);
		}
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

}
void Cwhu_UserDlg::Lin_ImportArrToList(CStringArray *m_DutyArr,int Num,CListCtrl &m_List)
{
	////先删除列表内容//
	//m_List.DeleteAllItems();
	//while(m_List.GetHeaderCtrl()->GetItemCount()>0)
	//{
	//	m_List.DeleteColumn(0);
	//}
	CStringArray m_Head;
	for (int i=0;i<Num;i++)
	{ 
		m_Head.Add(m_DutyArr[i].GetAt(0));
	}
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

void Cwhu_UserDlg::OnBnClickedBack()
{
	// TODO: Add your control notification handler code here
	Lin_ListToArr(m_UserList,((Cwhu_VoiceCardDlg*)whu_ParentDlg)->whu_User);
	CDialogEx::OnCancel();
}
int Cwhu_UserDlg::Lin_ListToArr(CListCtrl &m_List,CStringArray *m_DutyArr)
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
		m_DutyArr[ColumNum-1].RemoveAll();
		m_DutyArr[ColumNum-1].Add(str);
	}
	for (int i=0;i<m_List.GetItemCount();i++)
	{
		for (int j=1;j<m_List.GetHeaderCtrl()->GetItemCount();j++)
		{
			CString m_str = m_List.GetItemText(i,j);
			m_DutyArr[j-1].Add(m_str);
		}
	}
	return Num;
}

void Cwhu_UserDlg::OnNMClickList1(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMITEMACTIVATE pNMItemActivate = reinterpret_cast<LPNMITEMACTIVATE>(pNMHDR);
	// TODO: Add your control notification handler code here
	m_ClickPos = pNMItemActivate->iItem;
	CString m_Name = m_UserList.GetItemText(m_ClickPos,1);
	GetDlgItem(IDC_CHGPWD)->EnableWindow(true);
	if (m_Name == _T("管理员"))
	{
		GetDlgItem(IDC_ACCEPT)->EnableWindow(false);
		GetDlgItem(IDC_REFUSE)->EnableWindow(false);
		GetDlgItem(IDC_LOGOFF)->EnableWindow(false);
		
	}else{
		GetDlgItem(IDC_LOGOFF)->EnableWindow(true);
		UpdateData(true);
		bool m_Check = false;
		for (int i=0;i<m_UserList.GetItemCount();i++)
		{
			if (m_UserList.GetCheck(i)==true)
			{
				m_Check = true;
				CString m_str = m_UserList.GetItemText(i,3);
				if (m_str == _T("申请注册"))
				{
					GetDlgItem(IDC_ACCEPT)->EnableWindow(true);
					GetDlgItem(IDC_REFUSE)->EnableWindow(true);
					GetDlgItem(IDC_LOGOFF)->EnableWindow(false);
					GetDlgItem(IDC_CHGPWD)->EnableWindow(false);
				}
			}
		}
		if (m_Check == false)
		{
			CString m_str = m_UserList.GetItemText(m_ClickPos,3);
			if (m_str == _T("申请注册"))
			{
				GetDlgItem(IDC_ACCEPT)->EnableWindow(true);
				GetDlgItem(IDC_REFUSE)->EnableWindow(true);
				GetDlgItem(IDC_LOGOFF)->EnableWindow(false);
				GetDlgItem(IDC_CHGPWD)->EnableWindow(false);
			}
			else{
				GetDlgItem(IDC_ACCEPT)->EnableWindow(false);
				GetDlgItem(IDC_REFUSE)->EnableWindow(false);
				GetDlgItem(IDC_LOGOFF)->EnableWindow(true);
				GetDlgItem(IDC_CHGPWD)->EnableWindow(true);
			}
		}
	}
	
	*pResult = 0;
}


void Cwhu_UserDlg::OnBnClickedAccept()
{
	// TODO: Add your control notification handler code here
	bool m_check = false;
	for (int i=0;i<m_UserList.GetItemCount();i++)
	{
		
		CString m_str = m_UserList.GetItemText(i,3);
		if (m_UserList.GetCheck(i) == true)
		{
			if ((m_str == _T("申请注册"))&&i!=0)
			{
				m_UserList.SetItemText(i,3,_T("已注册"));
				m_UserList.SetItemText(i,4,_T("普通"));
			}
			m_check = true;
		}
	}
	if (m_check == false)
	{
		CString m_str = m_UserList.GetItemText(m_ClickPos,3);
		if ((m_str == _T("申请注册"))&&m_ClickPos!=0)
		{
			m_UserList.SetItemText(m_ClickPos,3,_T("已注册"));
			m_UserList.SetItemText(m_ClickPos,4,_T("普通"));
		}
	}
	GetDlgItem(IDC_ACCEPT)->EnableWindow(false);
	GetDlgItem(IDC_REFUSE)->EnableWindow(false);
	GetDlgItem(IDC_LOGOFF)->EnableWindow(false);
	GetDlgItem(IDC_CHGPWD)->EnableWindow(false);
}


void Cwhu_UserDlg::OnBnClickedRefuse()
{
	// TODO: Add your control notification handler code here
	bool m_check = false;
	for (int i=0;i<m_UserList.GetItemCount();i++)
	{
		CString m_str = m_UserList.GetItemText(i,3);
		if (m_UserList.GetCheck(i) == true&&m_str == _T("申请注册"))
		{
			m_UserList.DeleteItem(i);
			m_check = true;
		}
	}
	if (m_check == false)
	{
		CString m_str = m_UserList.GetItemText(m_ClickPos,3);
		if (m_str == _T("申请注册"))
		{
			m_UserList.DeleteItem(m_ClickPos);
		}
	}
	for (int i=0;i<m_UserList.GetItemCount();i++)
	{
		CString m_str;
		m_str.Format(_T("%d"),i+1);
		m_UserList.SetItemText(i,0,m_str);
	}
	GetDlgItem(IDC_ACCEPT)->EnableWindow(false);
	GetDlgItem(IDC_REFUSE)->EnableWindow(false);
	GetDlgItem(IDC_LOGOFF)->EnableWindow(false);
	GetDlgItem(IDC_CHGPWD)->EnableWindow(false);
}


void Cwhu_UserDlg::OnBnClickedLogoff()
{
	// TODO: Add your control notification handler code here
	GetDlgItem(IDC_LOGOFF)->EnableWindow(false);
	m_UserList.DeleteItem(m_ClickPos);
	for (int i=0;i<m_UserList.GetItemCount();i++)
	{
		CString m_str;
		m_str.Format(_T("%d"),i+1);
		m_UserList.SetItemText(i,0,m_str);
	}
}


void Cwhu_UserDlg::OnBnClickedChgpwd()
{
	// TODO: Add your control notification handler code here
	Cwhu_ChgPwdDlg m_Dlg;
	m_Dlg.whu_ParentDlg = this;
	m_Dlg.DoModal();
}
