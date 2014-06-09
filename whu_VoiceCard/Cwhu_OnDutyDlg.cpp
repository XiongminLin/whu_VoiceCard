// Cwhu_OnDutyDlg.cpp : implementation file
//

#include "stdafx.h"
#include "whu_VoiceCard.h"
#include "Cwhu_OnDutyDlg.h"
#include "afxdialogex.h"
#include "whu_VoiceCardDlg.h"

// Cwhu_OnDutyDlg dialog

IMPLEMENT_DYNAMIC(Cwhu_OnDutyDlg, CDialogEx)

Cwhu_OnDutyDlg::Cwhu_OnDutyDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(Cwhu_OnDutyDlg::IDD, pParent)
	, m_NewOnDutyLeader(_T(""))
	, m_NewOnDutyPerson(_T(""))
	, m_DutyTime(COleDateTime::GetCurrentTime())
{

}

Cwhu_OnDutyDlg::~Cwhu_OnDutyDlg()
{
}


void Cwhu_OnDutyDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Text(pDX, IDC_NEWLEADER, m_NewOnDutyLeader);
	DDX_Text(pDX, IDC_NEWPERSON, m_NewOnDutyPerson);
	DDX_Control(pDX, IDC_LISTDUTY, m_MyList);
	DDX_DateTimeCtrl(pDX, IDC_DUTYTIME, m_DutyTime);
}


BEGIN_MESSAGE_MAP(Cwhu_OnDutyDlg, CDialogEx)
	ON_BN_CLICKED(IDC_SWITCHDUTY, &Cwhu_OnDutyDlg::OnBnClickedSwitchduty)
	ON_BN_CLICKED(IDC_DELETE, &Cwhu_OnDutyDlg::OnBnClickedDelete)
	ON_BN_CLICKED(IDC_IMPORTDUTY, &Cwhu_OnDutyDlg::OnBnClickedImportduty)
	ON_BN_CLICKED(IDC_BACK, &Cwhu_OnDutyDlg::OnBnClickedBack)
	ON_NOTIFY(LVN_COLUMNCLICK, IDC_LISTDUTY, &Cwhu_OnDutyDlg::OnLvnColumnclickListduty)
	ON_NOTIFY(NM_CLICK, IDC_LISTDUTY, &Cwhu_OnDutyDlg::OnNMClickListduty)
	ON_BN_CLICKED(IDC_MODIFY, &Cwhu_OnDutyDlg::OnBnClickedModify)
	ON_BN_CLICKED(IDC_ADD, &Cwhu_OnDutyDlg::OnBnClickedAdd)
END_MESSAGE_MAP()


// Cwhu_OnDutyDlg message handlers

BOOL Cwhu_OnDutyDlg::OnInitDialog()
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
	DutyHeadNum = ((Cwhu_VoiceCardDlg*)whu_ParentDlg)->DutyHeadNum;
	if (DutyHeadNum>0)
	{
		Lin_ImportArrToList(((Cwhu_VoiceCardDlg*)whu_ParentDlg)->whu_DutyPersons,DutyHeadNum,m_MyList); 
	}else
	{
		AfxMessageBox(_T("值班信息为空，请按照模板导入Excel文件"));
	}
	//Lin_ImportArrToList(((Cwhu_VoiceCardDlg*)whu_ParentDlg)->whu_DutyPersons,DutyHeadNum,m_MyList); 

	/*CString m_FilePath = _T("C:\\Whu_Data\\Excel文件\\值班表.xlsx");
	CString m_Folder = _T("C:\\Whu_Data\\Excel文件");
	if (whu_findfile(m_FilePath,m_Folder)==true)
	{
		Lin_InportExcelToList(m_FilePath,m_MyList);
	}*/
	m_ChoosedPos = -1;
	return true;
}
void Cwhu_OnDutyDlg::OnBnClickedSwitchduty()
{

	// TODO: Add your control notification handler code here
	//UpdateData(true);
	CString m_Leader,m_DutyPerson;
	if (m_ChoosedPos<m_MyList.GetItemCount()&&m_ChoosedPos>=0)
	{
		m_Leader = m_MyList.GetItemText(m_ChoosedPos,5);
		m_DutyPerson = m_MyList.GetItemText(m_ChoosedPos,2);
		((Cwhu_VoiceCardDlg*)whu_ParentDlg)->SetDlgItemTextA(IDC_COMBOLEADER,m_Leader);
		((Cwhu_VoiceCardDlg*)whu_ParentDlg)->SetDlgItemTextA(IDC_DUTYPERSON,m_DutyPerson);
	}

}


void Cwhu_OnDutyDlg::OnBnClickedDelete()
{
	// TODO: Add your control notification handler code here
	bool m_Check = false;
	for (int i=m_MyList.GetItemCount()-1;i>=0;i--)
	{
		bool m_Check = m_MyList.GetCheck(i);
		if (m_Check == true)
		{
			m_MyList.DeleteItem(i);
		}
	}
	if (m_Check ==false)
	{
		m_MyList.DeleteItem(m_ChoosedPos);
	}
	for (int i=0;i<m_MyList.GetItemCount();i++)
	{
		CString m_str;
		m_str.Format(_T("%d"),i+1);
		m_MyList.SetItemText(i,0,m_str);
	}
}


void Cwhu_OnDutyDlg::OnBnClickedImportduty()
{
	// TODO: Add your control notification handler code here
	CString m_pathname;
	CFileDialog cf(1, NULL, NULL, OFN_HIDEREADONLY|OFN_OVERWRITEPROMPT, "Excel File(*.xlsx;*.xls)|*.xlsx;*.xls",NULL);//
	if(cf.DoModal() == IDOK)
	{
		CString m_DocFile=cf.GetPathName();
		Lin_InportExcelToList(m_DocFile,m_MyList);
	}
	((Cwhu_VoiceCardDlg*)whu_ParentDlg)->DutyHeadNum = m_MyList.GetHeaderCtrl()->GetItemCount()-1;
}



void Cwhu_OnDutyDlg::OnBnClickedBack()
{
	// TODO: Add your control notification handler code here
	
	//CString m_FilePath = _T("C:\\Whu_Data\\Excel文件\\值班表.xlsx");
	//CString m_Folder = _T("C:\\Whu_Data\\Excel文件");
	//if (whu_findfile(m_FilePath,m_Folder)==true)
	//{
	//	CFile m_Temp;
	//	m_Temp.Remove(m_FilePath);
	//}
	//Lin_ExportListToExcel(m_MyList,m_FilePath);

	//for (int i =0;i<DutyHeadNum;i++)
	//{
	//	int Length = ((Cwhu_VoiceCardDlg*)whu_ParentDlg)->whu_DutyPersons[i].GetSize()-1;
	//	((Cwhu_VoiceCardDlg*)whu_ParentDlg)->whu_DutyPersons[i].RemoveAt(1,Length);
	//}
	//for (int i =1;i<m_MyList.GetHeaderCtrl()->GetItemCount();i++)
	//{
	//	for (int j=0;j<m_MyList.GetItemCount();j++)
	//	{
	//		CString m_test = m_MyList.GetItemText(j,i);
	//		((Cwhu_VoiceCardDlg*)whu_ParentDlg)->whu_DutyPersons[i-1].Add(m_MyList.GetItemText(j,i));
	//	}
	//}
	((Cwhu_VoiceCardDlg*)whu_ParentDlg)->DutyHeadNum = Lin_ListToArr(m_MyList,((Cwhu_VoiceCardDlg*)whu_ParentDlg)->whu_DutyPersons);

	//DestroyWindow();
	CDialogEx::OnOK();
}
CString Cwhu_OnDutyDlg::Lin_GetEnglishCharacter(int Num)
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
bool Cwhu_OnDutyDlg::Lin_InportExcelToList(CString m_FilePath,CListCtrl &m_List)
{
	CStringArray m_StandarHeadName;
	m_StandarHeadName.Add(_T("时间"));
	m_StandarHeadName.Add(_T("姓名"));
	m_StandarHeadName.Add(_T("值班地点"));
	m_StandarHeadName.Add(_T("值班电话"));
	m_StandarHeadName.Add(_T("带班领导"));

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
				AfxMessageBox(_T("值班表不符合模板要求"));
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
	for (int ItemNum = 0;ItemNum<10000;ItemNum++)
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
void Cwhu_OnDutyDlg::Lin_ExportListToExcel(CListCtrl &m_List,CString m_FilePath)
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
void Cwhu_OnDutyDlg::Lin_InitList(CListCtrl &m_List,CStringArray &ColumName)
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
	m_List.InsertColumn(0,_T("编号"),LVCFMT_LEFT,50);
	m_List.InsertColumn(1,ColumName.GetAt(0),LVCFMT_LEFT,250);
	for (int i=1;i<ColumName.GetSize();i++)
	{
		m_List.InsertColumn(i+1,ColumName.GetAt(i),LVCFMT_LEFT,150);
	}

}
void Cwhu_OnDutyDlg::Lin_InsertList(CListCtrl &m_List,CStringArray &ItemArr)
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

void Cwhu_OnDutyDlg::OnLvnColumnclickListduty(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMLISTVIEW pNMLV = reinterpret_cast<LPNMLISTVIEW>(pNMHDR);
	// TODO: Add your control notification handler code here
	*pResult = 0;
}


void Cwhu_OnDutyDlg::OnNMClickListduty(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMITEMACTIVATE pNMItemActivate = reinterpret_cast<LPNMITEMACTIVATE>(pNMHDR);
	// TODO: Add your control notification handler code here
	m_ChoosedPos = pNMItemActivate->iItem;
	if(m_ChoosedPos != -1) 
	{ 
		//下面设置时间//
		CString m_Str = m_MyList.GetItemText(m_ChoosedPos,1);
		int pos1 = m_Str.Find(_T("年"));
		int pos2 = m_Str.Find(_T("月"));
		int pos3 = m_Str.Find(_T("日"));
		if (pos1 >=0&&pos2>=0&&pos3>=0)
		{
			CString year = m_Str.Mid(0,pos1);
			CString Month = m_Str.Mid(pos1+2,pos2-pos1-2);
			CString Day = m_Str.Mid(pos2+2,pos3-pos2-2);
			m_DutyTime.SetDate(_atoi64(year),_atoi64(Month),_atoi64(Day));
		}else if(pos1==-1&&pos2>=0&&pos3>=0){
			CString Month = m_Str.Mid(0,pos2);
			CString Day = m_Str.Mid(pos2+2,pos3-pos2-2);
			m_DutyTime.SetDate(m_DutyTime.GetYear(),_atoi64(Month),_atoi64(Day));
			
		}else if (pos1==-1&&pos2==-1&&pos3>=0)
		{
			CString Day = m_Str.Mid(0,pos3);
			m_DutyTime.SetDate(m_DutyTime.GetYear(),m_DutyTime.GetMonth(),_atoi64(Day));
			
		}else{
			m_DutyTime.SetDate(m_DutyTime.GetYear(),m_DutyTime.GetMonth(),m_DutyTime.GetDay());
		}
		UpdateData(false);

		m_Str = m_MyList.GetItemText(m_ChoosedPos,2);
		SetDlgItemText(IDC_NEWPERSON,m_Str);

		m_Str = m_MyList.GetItemText(m_ChoosedPos,3);
		SetDlgItemText(IDC_DUTYPLACE,m_Str);

		m_Str = m_MyList.GetItemText(m_ChoosedPos,4);
		SetDlgItemText(IDC_DUTYNUMBER,m_Str);

		m_Str = m_MyList.GetItemText(m_ChoosedPos,5);
		SetDlgItemText(IDC_NEWLEADER,m_Str);

	}
	*pResult = 0;
}


void Cwhu_OnDutyDlg::OnBnClickedModify()
{
	// TODO: Add your control notification handler code here
	//m_Str = m_List.GetItemText(m_ChoosedPos,2);
	//SetDlgItemText(IDC_NEWPERSON,m_Str);

	//m_Str = m_List.GetItemText(m_ChoosedPos,3);
	//SetDlgItemText(IDC_DUTYPLACE,m_Str);

	//m_Str = m_List.GetItemText(m_ChoosedPos,4);
	//SetDlgItemText(IDC_DUTYNUMBER,m_Str);

	//m_Str = m_List.GetItemText(m_ChoosedPos,5);
	//SetDlgItemText(IDC_NEWLEADER,m_Str);
	UpdateData(true);
	CString buf=_T("");
	int week = m_DutyTime.GetDayOfWeek();
	switch(week)
	{
	case 2:
		buf = _T("周一");
		break;
	case 3:
		buf = _T("周二");
		break;
	case 4:
		buf = _T("周三");
		break;
	case 5:
		buf = _T("周四");
		break;
	case 6:
		buf = _T("周五");
		break;
	case 7:
		buf = _T("周六");
		break;
	case 1:
		buf = _T("周日");
		break;
	default:
		break;
	}
	CString m_Str = _T("");
	m_Str.Format(_T("%d年%d月%d日（%s）"),m_DutyTime.GetYear(),m_DutyTime.GetMonth(),m_DutyTime.GetDay(),buf);
	if (m_ChoosedPos <m_MyList.GetItemCount())
	{
		m_MyList.SetItemText(m_ChoosedPos,1,m_Str);
		GetDlgItemText(IDC_NEWPERSON,m_Str);
		m_MyList.SetItemText(m_ChoosedPos,2,m_Str);
		GetDlgItemText(IDC_DUTYPLACE,m_Str);
		m_MyList.SetItemText(m_ChoosedPos,3,m_Str);
		GetDlgItemText(IDC_DUTYNUMBER,m_Str);
		m_MyList.SetItemText(m_ChoosedPos,4,m_Str);
		GetDlgItemText(IDC_NEWLEADER,m_Str);
		m_MyList.SetItemText(m_ChoosedPos,5,m_Str);
	}


}


void Cwhu_OnDutyDlg::OnBnClickedAdd()
{
	// TODO: Add your control notification handler code here
	CStringArray m_Arr;
	UpdateData(true);
	CString buf=_T("");
	int week = m_DutyTime.GetDayOfWeek();
	switch(week)
	{
	case 2:
		buf = _T("周一");
		break;
	case 3:
		buf = _T("周二");
		break;
	case 4:
		buf = _T("周三");
		break;
	case 5:
		buf = _T("周四");
		break;
	case 6:
		buf = _T("周五");
		break;
	case 7:
		buf = _T("周六");
		break;
	case 1:
		buf = _T("周日");
		break;
	default:
		break;
	}
	CString m_Str = _T("");
	m_Str.Format(_T("%d年%d月%d日（%s）"),m_DutyTime.GetYear(),m_DutyTime.GetMonth(),m_DutyTime.GetDay(),buf);
	m_Arr.Add(m_Str);
	GetDlgItemText(IDC_NEWPERSON,m_Str);
	m_Arr.Add(m_Str);
	GetDlgItemText(IDC_DUTYPLACE,m_Str);
	m_Arr.Add(m_Str);
	GetDlgItemText(IDC_DUTYNUMBER,m_Str);
	m_Arr.Add(m_Str);
	GetDlgItemText(IDC_NEWLEADER,m_Str);
	m_Arr.Add(m_Str);
	Lin_InsertList(m_MyList,m_Arr);
}

bool Cwhu_OnDutyDlg::whu_findfile(CString m_file,CString folder)
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
void Cwhu_OnDutyDlg::Lin_ImportArrToList(CStringArray *m_DutyArr,int Num,CListCtrl &m_List)
{
	//先删除列表内容//
	m_List.DeleteAllItems();
	while(m_List.GetHeaderCtrl()->GetItemCount()>0)
	{
		m_List.DeleteColumn(0);
	}
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
int Cwhu_OnDutyDlg::Lin_ListToArr(CListCtrl &m_List,CStringArray *m_DutyArr)
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