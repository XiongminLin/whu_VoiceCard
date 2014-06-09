// Cwhu_FaxSettingDlg.cpp : implementation file
//

#include "stdafx.h"
#include "whu_VoiceCard.h"
#include "Cwhu_FaxSettingDlg.h"
#include "afxdialogex.h"
#include "Cwhu_FaxDlg.h"
#include "Whu_AutoForDlg.h"
#include "whu_VoiceCardDlg.h"
// Cwhu_FaxSettingDlg dialog

IMPLEMENT_DYNAMIC(Cwhu_FaxSettingDlg, CDialogEx)

Cwhu_FaxSettingDlg::Cwhu_FaxSettingDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(Cwhu_FaxSettingDlg::IDD, pParent)
{

}

Cwhu_FaxSettingDlg::~Cwhu_FaxSettingDlg()
{
}

void Cwhu_FaxSettingDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_LISTFORM, m_ListFrom);
	DDX_Control(pDX, IDC_LISTTO, m_ListTo);
	//DDX_Check(pDX, IDC_CHECK1, whu_BAutoForward);
}


BEGIN_MESSAGE_MAP(Cwhu_FaxSettingDlg, CDialogEx)
	ON_BN_CLICKED(IDOK, &Cwhu_FaxSettingDlg::OnBnClickedOk)
	ON_BN_CLICKED(IDCANCEL, &Cwhu_FaxSettingDlg::OnBnClickedCancel)
	ON_NOTIFY(HDN_ITEMCLICK, 0, &Cwhu_FaxSettingDlg::OnHdnItemclickListform)
	ON_NOTIFY(NM_CLICK, IDC_LISTFORM, &Cwhu_FaxSettingDlg::OnNMClickListform)
	ON_BN_CLICKED(IDC_ADDCONTACT, &Cwhu_FaxSettingDlg::OnBnClickedAddcontact)
	ON_BN_CLICKED(IDC_DELETE, &Cwhu_FaxSettingDlg::OnBnClickedDelete)
	//ON_BN_CLICKED(IDC_CHECK1, &Cwhu_FaxSettingDlg::OnBnClickedFaxAutoForward)
	ON_NOTIFY(NM_CUSTOMDRAW, IDC_LISTFORM, &Cwhu_FaxSettingDlg::OnNMCustomdrawListform)
END_MESSAGE_MAP()


// Cwhu_FaxSettingDlg message handlers


void Cwhu_FaxSettingDlg::OnBnClickedOk()
{
	// TODO: Add your control notification handler code here
	//CString m_FilePath = _T("C:\\Whu_Data\\Excel�ļ�\\�����Զ�ת����.xlsx");
	//CString m_Folder = _T("C:\\Whu_Data\\Excel�ļ�");
	//if (whu_findfile(m_FilePath,m_Folder)==true)
	//{
	//	CFile m_Temp;
	//	m_Temp.Remove(m_FilePath);
	//}
	//Lin_ExportForArrToExcel(whu_AutoForward,m_FilePath);
	//((Cwhu_FaxDlg*)whu_ParentDlg)->whu_AutoForward = whu_AutoForward;
	//UpdateData(true);
	//((Cwhu_VoiceCardDlg*)(((Cwhu_FaxDlg*)whu_ParentDlg)->whu_ParentDlg))->whu_BAutoForward = whu_BAutoForward;
	whu_FaxAutoForwardArr.Copy(whu_AutoForward);
	CDialogEx::OnOK();
}
bool Cwhu_FaxSettingDlg::whu_findfile(CString m_file,CString folder)
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

void Cwhu_FaxSettingDlg::OnBnClickedCancel()
{
	// TODO: Add your control notification handler code here
	CDialogEx::OnCancel();
}
BOOL Cwhu_FaxSettingDlg::OnInitDialog()
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

	whu_InitList();
	////�����ǳ�ʼ��whu_AutoForwar
	//CString m_FilePath = _T("C:\\Whu_Data\\Excel�ļ�\\�����Զ�ת����.xlsx");
	//Lin_ImportAutoForArr(m_FilePath,whu_AutoForward);
	whu_AutoForward.Copy(whu_FaxAutoForwardArr);

	//����ת������ɫ����Ϊ��ɫ//

	//int test = whu_AutoForward.GetSize();
	//whu_BAutoForward = ((Cwhu_VoiceCardDlg*)(((Cwhu_FaxDlg*)whu_ParentDlg)->whu_ParentDlg))->whu_BAutoForward ;
	//UpdateData(false);
	return true;
}
void Cwhu_FaxSettingDlg::whu_InitList()
{
	//��ϵ���б��ʼ��
	LONG lStyle; 
	lStyle = GetWindowLong(m_ListFrom.m_hWnd, GWL_STYLE);
	lStyle &= ~LVS_TYPEMASK; 
	lStyle |= LVS_REPORT; 
	SetWindowLong(m_ListFrom.m_hWnd, GWL_STYLE, lStyle);
	DWORD dwStyle = m_ListFrom.GetExtendedStyle(); 
	dwStyle |= LVS_EX_FULLROWSELECT;
	dwStyle |= LVS_EX_GRIDLINES;
	m_ListFrom.SetExtendedStyle(dwStyle); 
	m_ListFrom.InsertColumn( 0, _T("���"), LVCFMT_LEFT, 60 );
	m_ListFrom.InsertColumn( 1, _T("����"), LVCFMT_LEFT, 180 ); 
	m_ListFrom.InsertColumn( 2, _T("��λ"), LVCFMT_LEFT, 180 );
	m_ListFrom.InsertColumn( 3, _T("����"), LVCFMT_LEFT, 180 );
	m_ListFrom.InsertColumn( 4, _T("ְ��"), LVCFMT_LEFT, 120 );
	m_ListFrom.InsertColumn( 5, _T("�ֻ�"), LVCFMT_LEFT, 0 ); 
	m_ListFrom.InsertColumn( 6, _T("�칫�ҵ绰"), LVCFMT_LEFT, 0 ); 
	m_ListFrom.InsertColumn( 7, _T("�����"), LVCFMT_LEFT, 180 ); 
	m_ListFrom.InsertColumn( 8, _T("����"), LVCFMT_LEFT, 0 );

	m_ClickPos = 0;

	//��ϵ���б��ʼ��
	lStyle = GetWindowLong(m_ListTo.m_hWnd, GWL_STYLE);
	lStyle &= ~LVS_TYPEMASK; 
	lStyle |= LVS_REPORT; 
	SetWindowLong(m_ListTo.m_hWnd, GWL_STYLE, lStyle);
	dwStyle = m_ListTo.GetExtendedStyle(); 
	dwStyle |= LVS_EX_FULLROWSELECT;
	dwStyle |= LVS_EX_GRIDLINES;
	dwStyle |= LVS_EX_CHECKBOXES;
	m_ListTo.SetExtendedStyle(dwStyle); 
	m_ListTo.InsertColumn( 0, _T("���"), LVCFMT_LEFT, 60 );
	m_ListTo.InsertColumn( 1, _T("����"), LVCFMT_LEFT, 180 ); 
	m_ListTo.InsertColumn( 2, _T("��λ"), LVCFMT_LEFT, 180 );
	m_ListTo.InsertColumn( 3, _T("����"), LVCFMT_LEFT, 180 );
	m_ListTo.InsertColumn( 4, _T("ְ��"), LVCFMT_LEFT, 120 );
	m_ListTo.InsertColumn( 5, _T("�ֻ�"), LVCFMT_LEFT, 0 ); 
	m_ListTo.InsertColumn( 6, _T("�칫�ҵ绰"), LVCFMT_LEFT, 0 ); 
	m_ListTo.InsertColumn( 7, _T("�����"), LVCFMT_LEFT, 180 ); 
	m_ListTo.InsertColumn( 8, _T("����"), LVCFMT_LEFT, 0 );

	//ʵ�ڲ�֪����ʲôԭ��//���ʲ��˸��Ի���Ĺ��б���//��������������m_FaxContactList������Ϊpublic:������������������֪��why��//
	int num = ((Cwhu_FaxDlg*)whu_ParentDlg)->m_FaxContactList.GetItemCount();
	int test = ((Cwhu_FaxDlg*)whu_ParentDlg)->test;
	for (int i=0;i<num;i++)
	{
		char buf[3];
		itoa(i+1,buf,10);
		m_ListFrom.InsertItem(i,buf);
		for (int j =0;j<9;j++)
		{
			CString m_str = ((Cwhu_FaxDlg*)whu_ParentDlg)->m_FaxContactList.GetItemText(i,j);
			m_ListFrom.SetItemText(i,j,m_str);
		}
	}
	this->SetFocus();
	m_ListFrom.SetFocus();
	m_ListFrom.SetItemState(0,LVIS_FOCUSED | LVIS_SELECTED,LVIS_FOCUSED | LVIS_SELECTED);
	m_ListFrom.Update(0);

}

void Cwhu_FaxSettingDlg::OnHdnItemclickListform(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMHEADER phdr = reinterpret_cast<LPNMHEADER>(pNMHDR);
	// TODO: Add your control notification handler code here
	*pResult = 0;
}


void Cwhu_FaxSettingDlg::OnNMClickListform(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMITEMACTIVATE pNMItemActivate = reinterpret_cast<LPNMITEMACTIVATE>(pNMHDR);
	// TODO: Add your control notification handler code here
	m_ClickPos = pNMItemActivate->iItem;
	CString m_FaxNum = m_ListFrom.GetItemText(m_ClickPos,7);//�õ������
	int size = whu_AutoForward.GetSize();
	int NumCount = 0;
	m_ListTo.DeleteAllItems();
	for (int i=0;i<size;i++)
	{
		CString m_str = whu_AutoForward.GetAt(i);
		if ((m_str == m_FaxNum)&&(i < size-1))
		{
			CString m_str2 = whu_AutoForward.GetAt(i+1);
			int CharCount = m_str2.GetLength(); //�ַ������鳤����������ģ�˵����������ת����������ı�־λ��////
			if(CharCount<4)
			{
				int NumCount = _ttoi(m_str2);
				for (int j=0;j<NumCount;j++)
				{
					CString m_ForFaxNum = whu_AutoForward.GetAt(i+2+j);
					whu_ShowToList(m_ForFaxNum);
				}
			}


		}
	}
	*pResult = 0;
}
void Cwhu_FaxSettingDlg::whu_ShowToList(CString m_FaxNum)
{
	//����whu_AutoForward��m_FaxNum����ʾ�Զ�ת��������//
	int Pos = m_ListTo.GetItemCount();
	CString m_PosStr;
	for (int i=0;i<m_ListFrom.GetItemCount();i++)
	{
		bool find = false;
		for (int j=0;j<9;j++)
		{
			CString m_str = m_ListFrom.GetItemText(i,j);
			if (m_str == m_FaxNum)
			{
				find = true;
				break;
			}
		}
		if (find == true)
		{
			m_PosStr.Format(_T("%d"),Pos+1);
			m_ListTo.InsertItem(Pos,m_PosStr);
			m_ListTo.SetItemText(Pos,0,m_PosStr);
			for (int j=1;j<9;j++)
			{
				m_ListTo.SetItemText(Pos,j,m_ListFrom.GetItemText(i,j));
			}
		}
	}
	//int m_FindPos = whu_AutoForward.

}

void Cwhu_FaxSettingDlg::OnBnClickedAddcontact()
{
	// TODO: Add your control notification handler code here
	CWhu_AutoForDlg * m_FaxAutoForDlg = new CWhu_AutoForDlg();
	m_FaxAutoForDlg->whu_ParentDlg = this;
	m_FaxAutoForDlg->DoModal();

}
void Cwhu_FaxSettingDlg::Lin_ImportAutoForArr(CString m_FilePath,CStringArray &m_AutoForward)
{
	int Number = 0;
	//m_DutyArr��������㹻��//
	//����
	CApplication app;
	CWorkbook book;
	CWorkbooks books;
	CWorksheet sheet;
	CWorksheets sheets;
	CRange range;
	LPDISPATCH lpDisp;
	//�������//
	COleVariant covOptional((long)
		DISP_E_PARAMNOTFOUND,VT_ERROR);
	if (!app.CreateDispatch(_T("Excel.Application")))
	{
		this->MessageBox(_T("�޷�����ExcelӦ��"));
	}
	books = app.get_Workbooks();
	//��Excel������m_FilePathΪExcel���·����//
	lpDisp = books.Open(m_FilePath,covOptional
		,covOptional,covOptional,covOptional
		,covOptional,covOptional,covOptional
		,covOptional,covOptional,covOptional
		,covOptional,covOptional,covOptional
		,covOptional);
	book.AttachDispatch(lpDisp);
	sheets = book.get_Worksheets();
	sheet = sheets.get_Item(COleVariant((short)1));
	
	CStringArray m_ContentArr;
	for (int ItemNum = 2;ItemNum<40;ItemNum++)
	{
		
		for (int ColumNum=1;ColumNum<11;ColumNum++)
		{
			CString m_pos = Lin_GetEnglishCharacter(ColumNum);
			CString m_Itempos;
			m_Itempos.Format(_T("%d"),ItemNum);
			CString m_str = m_pos+m_Itempos;
			range = sheet.get_Range(COleVariant(m_str),COleVariant(m_str));
			//��õ�Ԫ�������
			COleVariant rValue;
			rValue = COleVariant(range.get_Value2());
			//ת���ɿ��ַ�//
			rValue.ChangeType(VT_BSTR);
			//ת����ʽ�������//
			CString m_content = CString(rValue.bstrVal);
			if (m_content!=_T(""))
			{
				m_AutoForward.Add(m_content);
			}
		}
	}
	book.put_Saved(TRUE);
	app.Quit();
}
CString Cwhu_FaxSettingDlg::Lin_GetEnglishCharacter(int Num)
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

void Cwhu_FaxSettingDlg::Lin_ExportForArrToExcel(CStringArray &m_AutoForward,CString m_FilePath)
{

	CApplication app;
	CWorkbooks books;
	CWorkbook book;
	CWorksheets sheets;
	CWorksheet sheet;
	CRange range;
	CRange cols;
	CString str1, str2;
	int m_ArrPos = 0;
	int m_ArrCount = m_AutoForward.GetSize();
	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	if( !app.CreateDispatch(_T("Excel.Application")))
	{
		MessageBox(_T("�޷�����ExcelӦ�ã�"));
		return;
	}
	books=app.get_Workbooks();
	book=books.Add(covOptional);
	sheets=book.get_Sheets();
	sheet=sheets.get_Item(COleVariant((short)1));

	//д���ͷ//
	str2 = _T("A1");
	CString str = _T("Դ����");
	range=sheet.get_Range(COleVariant(str2),COleVariant(str2));
	range.put_Value2(COleVariant(str));
	str2 = _T("B1");
	str = _T("ת���������");
	range=sheet.get_Range(COleVariant(str2),COleVariant(str2));
	range.put_Value2(COleVariant(str));

	for (int j=3;j<11;j++)
	{
		str2 = Lin_GetEnglishCharacter(j);
		str2 = str2 +_T("1");
		str= _T("ת������");
		range=sheet.get_Range(COleVariant(str2),COleVariant(str2));
		range.put_Value2(COleVariant(str));

	}
	CStringArray m_LineStart;
	for (int i=0;i<m_ArrCount;i++)
	{
		CString m_str= m_AutoForward.GetAt(i);
		int CharCount = m_str.GetLength();
		if (CharCount<4) //����Ǻ������//
		{
			m_str.Format(_T("%d"),i-1);
			m_LineStart.Add(m_str);
		}
	}
	int ItemNum = 2;
	int ColumNum = 1;
	for (int i=0;i<m_LineStart.GetSize();i++) 
	{
		int Pos = _ttoi(m_LineStart.GetAt(i));
		CString m_str = m_AutoForward.GetAt(Pos+1);
		int Number = _ttoi(m_str);
		for (int j=0;j<Number+2;j++)
		{
			str2 = Lin_GetEnglishCharacter(ColumNum);
			str1.Format(_T("%d"),ItemNum);
			str2 = str2 +str1;
			CString content = _T("'")+m_AutoForward.GetAt(Pos+j);
			range=sheet.get_Range(COleVariant(str2),COleVariant(str2));
			range.put_Value2(COleVariant(content));
			ColumNum++;
			if (ColumNum>10)
			{
				ColumNum = 3;
				ItemNum++;
			}
		}
		ItemNum++;
		ColumNum = 1;
	}
	cols=range.get_EntireColumn();
	cols.AutoFit();
	//app.put_Visible(TRUE);
	//app.put_UserControl(TRUE);
	book.SaveAs(COleVariant(m_FilePath),covOptional,covOptional,covOptional,covOptional,covOptional,0,covOptional,covOptional,covOptional,covOptional,covOptional);
	app.Quit();


}

void Cwhu_FaxSettingDlg::OnBnClickedDelete()
{
	// TODO: Add your control notification handler code here

	for (int i=m_ListTo.GetItemCount()-1;i>=0;i--)
	{
		if (m_ListTo.GetCheck(i) == true)
		{
			m_ListTo.DeleteItem(i);
		}
	}
	CString m_FaxNum = m_ListFrom.GetItemText(m_ClickPos,7);
	CStringArray m_AddArr;
	m_AddArr.Add(m_FaxNum);
	CString m_NumCount;
	m_NumCount.Format(_T("%d"),m_ListTo.GetItemCount());
	m_AddArr.Add(m_NumCount);
	for (int i=0;i<m_ListTo.GetItemCount();i++)
	{
		CString m_str = m_ListFrom.GetItemText(i,7);
		m_AddArr.Add(m_str);
	}
	int test = m_AddArr.GetSize();
	whu_AddAutoForArr(whu_AutoForward,m_AddArr);
	
}
void Cwhu_FaxSettingDlg::whu_AddAutoForArr(CStringArray &m_SorArr,CStringArray &m_AddArr)
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

//void Cwhu_FaxSettingDlg::OnBnClickedFaxAutoForward()
//{
//	// TODO: Add your control notification handler code here
//
//}
void Cwhu_FaxSettingDlg::m_ShowForwardItem(CString FaxNume)
{
	for (int i=0;i<m_ListFrom.GetItemCount();i++)
	{
		CString m_faxNum = m_ListFrom.GetItemText(i,7);
		if (m_faxNum == FaxNume)
		{
			
		}
	}

}

void Cwhu_FaxSettingDlg::OnNMCustomdrawListform(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMCUSTOMDRAW pNMCD = reinterpret_cast<LPNMCUSTOMDRAW>(pNMHDR);
	// TODO: Add your control notification handler code here
	LPNMLVCUSTOMDRAW  lplvcd = (LPNMLVCUSTOMDRAW)pNMHDR; 
	*pResult  =  CDRF_DODEFAULT;

	switch (lplvcd->nmcd.dwDrawStage) 
	{ 
	case CDDS_PREPAINT : 
		{ 
			*pResult = CDRF_NOTIFYITEMDRAW; 
			return; 
		} 
	case CDDS_ITEMPREPAINT: 
		{ 
			*pResult = CDRF_NOTIFYSUBITEMDRAW; 
			return; 
		} 
	case CDDS_ITEMPREPAINT | CDDS_SUBITEM:
		{ 

			int    nItem = static_cast<int>( lplvcd->nmcd.dwItemSpec );
			COLORREF clrNewTextColor, clrNewBkColor;
			clrNewTextColor=RGB(0,0,0);
			if( nItem%2 ==0 )
			{
				clrNewBkColor = RGB( 255, 255, 0); //ż���б���ɫΪ��ɫ//
			}
			else
			{
				clrNewBkColor = RGB( 255, 0, 0 ); //�����б���ɫΪ��ɫ//
			}

			lplvcd->clrText = clrNewTextColor;
			lplvcd->clrTextBk = clrNewBkColor;

			*pResult = CDRF_DODEFAULT;  
			return; 
		} 
	} 

	*pResult = 0;
}
