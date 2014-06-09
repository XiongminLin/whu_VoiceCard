
// OpenExcelDlg.cpp : implementation file
//

#include "stdafx.h"
#include "OpenExcelDlg.h"
#include "afxdialogex.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CAboutDlg dialog used for App About


// COpenExcelDlg dialog




COpenExcelDlg::COpenExcelDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(COpenExcelDlg::IDD, pParent)
	, m_DocFile(_T(""))
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void COpenExcelDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_LIST2, m_MyList);
	DDX_Text(pDX, IDC_EXCELFILE, m_DocFile);
}

BEGIN_MESSAGE_MAP(COpenExcelDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_ADDFILE, &COpenExcelDlg::OnBnClickedAddfile)
	ON_BN_CLICKED(IDOK, &COpenExcelDlg::OnBnClickedOk)
	ON_BN_CLICKED(IDCANCEL, &COpenExcelDlg::OnBnClickedCancel)
END_MESSAGE_MAP()


// COpenExcelDlg message handlers

BOOL COpenExcelDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// Add "About..." menu item to system menu.

	// IDM_ABOUTBOX must be in the system command range.
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon

	// TODO: Add extra initialization here
	
	//CStringArray m_ColumName ;
	//CString m_str = _T("编号");
	//m_ColumName.Add(m_str);
	//m_str = _T("姓名");
	//m_ColumName.Add(m_str);
	//m_str = _T("单位");
	//m_ColumName.Add(m_str);
	//m_str = _T("部门");
	//m_ColumName.Add(m_str);
	//m_str = _T("职务");
	//m_ColumName.Add(m_str);
	//m_str = _T("手机");
	//m_ColumName.Add(m_str);
	//m_str = _T("办公室号码");
	//m_ColumName.Add(m_str);
	//m_str = _T("传真号");
	//m_ColumName.Add(m_str);
	//m_str = _T("邮箱");
	//m_ColumName.Add(m_str);
	//Lin_InitList(m_MyList,m_ColumName);
	//CStringArray m_ItemArr;
	//for (int i=0;i<m_MyList.GetHeaderCtrl()->GetItemCount()-1;i++)
	//{
	//	CString m_Str;
	//	m_Str.Format(_T("内容%d"),i+1);
	//	m_ItemArr.Add(m_Str);
	//}
	//Lin_InsertList(m_MyList,m_ItemArr);
	return TRUE;  // return TRUE  unless you set the focus to a control
}

//void COpenExcelDlg::OnSysCommand(UINT nID, LPARAM lParam)
//{
//	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
//	{
//		CAboutDlg dlgAbout;
//		dlgAbout.DoModal();
//	}
//	else
//	{
//		CDialogEx::OnSysCommand(nID, lParam);
//	}
//}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

//void COpenExcelDlg::OnPaint()
//{
//	if (IsIconic())
//	{
//		CPaintDC dc(this); // device context for painting
//
//		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);
//
//		// Center icon in client rectangle
//		int cxIcon = GetSystemMetrics(SM_CXICON);
//		int cyIcon = GetSystemMetrics(SM_CYICON);
//		CRect rect;
//		GetClientRect(&rect);
//		int x = (rect.Width() - cxIcon + 1) / 2;
//		int y = (rect.Height() - cyIcon + 1) / 2;
//
//		// Draw the icon
//		dc.DrawIcon(x, y, m_hIcon);
//	}
//	else
//	{
//		CDialogEx::OnPaint();
//	}
//}

// The system calls this function to obtain the cursor to display while the user drags
//  the minimized window.
//HCURSOR COpenExcelDlg::OnQueryDragIcon()
//{
//	return static_cast<HCURSOR>(m_hIcon);
//}

void COpenExcelDlg::Lin_InitList(CListCtrl &m_List,CStringArray &ColumName)
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
void COpenExcelDlg::Lin_InsertList(CListCtrl &m_List,CStringArray &ItemArr)
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
void COpenExcelDlg::OnBnClickedAddfile()
{
	// TODO: Add your control notification handler code here
	CString m_pathname;
	CFileDialog cf(1, NULL, NULL, OFN_HIDEREADONLY|OFN_OVERWRITEPROMPT, "Excel File(*.xlsx;*.xls)|*.xlsx;*.xls",NULL);//
	if(cf.DoModal() == IDOK)
	{
		m_DocFile=cf.GetPathName();
		Lin_InportExcelToList(m_DocFile,m_MyList);
		UpdateData(FALSE);
	}
	

}
void COpenExcelDlg::Lin_ExportListToExcel(CListCtrl &m_List)
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

		app.put_Visible(TRUE);
		app.put_UserControl(TRUE);
	}else{
		AfxMessageBox(_T("列表为空"));
	}
	
}
void COpenExcelDlg::Lin_InportExcelToList(CString m_FilePath,CListCtrl &m_List)
{
	//先删除列表内容//
	m_List.DeleteAllItems();
	while(m_List.GetHeaderCtrl()->GetItemCount()>0)
	{
		m_List.DeleteColumn(0);
	}
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

}
void COpenExcelDlg::OnBnClickedOk()
{
	// TODO: Add your control notification handler code here

	Lin_ExportListToExcel(m_MyList);
}
CString COpenExcelDlg::Lin_GetEnglishCharacter(int Num)
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

void COpenExcelDlg::OnBnClickedCancel()
{
	// TODO: Add your control notification handler code here
	CDialogEx::OnCancel();
}
