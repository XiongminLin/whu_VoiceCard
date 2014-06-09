XDispatch

// whu_VoiceCardDlg.cpp : implementation file
//

#include "stdafx.h"
#include "whu_VoiceCard.h"
#include "whu_VoiceCardDlg.h"
#include "afxdialogex.h"
//#include "column.h"
//#include "columns.h"
//#include "COMDEF.H"
#include  <string>
using namespace std;
#include "Mmsystem.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif
int whu_NumHaveSend;
CString whu_NumberBuf;
CString whu_AllNos[MAX_PATH];
CStringArray whu_FaxAutoForwardArr;
// CAboutDlg dialog used for App About

#define ID_SENDTO_MENU 90000
int nMenuIndex = 0;

string trim(string &s)
{
	const string &space =" \f\n\t\r\v" ;
	string r=s.erase(s.find_last_not_of(space)+1);
	return r.erase(0,r.find_first_not_of(space));
}
class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// Dialog Data
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

// Implementation
protected:
	DECLARE_MESSAGE_MAP()
public:
};

CAboutDlg::CAboutDlg() : CDialogEx(CAboutDlg::IDD)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
	ON_WM_RBUTTONDOWN()
END_MESSAGE_MAP()


// Cwhu_VoiceCardDlg dialog


Cwhu_VoiceCardDlg::Cwhu_VoiceCardDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(Cwhu_VoiceCardDlg::IDD, pParent)
	, whu_BRecord(FALSE)
	, whu_password(_T(""))
	, m_bAutoRecvFax(FALSE)
	, m_checkname001(TRUE)
	, m_checkunite002(TRUE)
	, m_checkdepartment003(TRUE)
	, m_checkjob004(TRUE)
	, m_checktel005(TRUE)
	, m_checkOffNum006(TRUE)
	, m_checkFax007(TRUE)
	, m_checkemail008(TRUE)
	, m_ContactSearch(_T(""))
	, m_ComboxShowLeader(_T(""))
	, whu_BAutoForward(FALSE)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
	whu_LogIn = false;
}

void Cwhu_VoiceCardDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	//DDX_Control(pDX, IDC_TAB, whu_Tab);
	//DDX_Control(pDX, IDC_SHOCKWAVEFLASH, whu_flash);
	DDX_Text(pDX, IDC_LOGNMIMA, whu_password);
	DDX_Control(pDX, IDC_MYCONTACTLIST, m_MyContactList);
	DDX_Control(pDX, IDC_MYLOGLIST, m_MyLogList);
	DDX_Check(pDX, IDC_CHECKRECVFAX, m_bAutoRecvFax);
	DDX_Check(pDX, IDC_CHECK001, m_checkname001);
	DDX_Check(pDX, IDC_CHECK002, m_checkunite002);
	DDX_Check(pDX, IDC_CHECK003, m_checkdepartment003);
	DDX_Check(pDX, IDC_CHECK004, m_checkjob004);
	DDX_Check(pDX, IDC_CHECK005, m_checktel005);
	DDX_Check(pDX, IDC_CHECK006, m_checkOffNum006);
	DDX_Check(pDX, IDC_CHECK007, m_checkFax007);
	DDX_Check(pDX, IDC_CHECK008, m_checkemail008);
	//DDX_Text(pDX, IDC_LOGREMARK, m_LogRemark);
	DDX_Control(pDX, IDC_PROGRESS02, m_ProgressCtrl);
	DDX_Text(pDX, IDC_SEARCH, m_ContactSearch);
	DDX_Control(pDX, IDC_COMBOLEADER, m_ComboBoxLeader);
	DDX_CBString(pDX, IDC_COMBOLEADER, m_ComboxShowLeader);
	//DDX_Control(pDX, IDC_LOGCOMBOBOXEX, whu_LogInComBox);
	DDX_Control(pDX, IDC_LOGCOMBOBOX, m_LogInComBox);
	DDX_Control(pDX, IDC_DATETIMEPICKER1, m_DataCtrl);
	DDX_Check(pDX, IDC_CHECKAUTOFORWARDFAX, whu_BAutoForward);
}

BEGIN_MESSAGE_MAP(Cwhu_VoiceCardDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	//ON_NOTIFY(TCN_SELCHANGE, IDC_TAB, &Cwhu_VoiceCardDlg::OnSelchangeTab)
	ON_BN_CLICKED(IDC_REFUSE, &Cwhu_VoiceCardDlg::OnBnClickedRefuse)
	ON_BN_CLICKED(IDC_ACCEPTFAXORRECORD, &Cwhu_VoiceCardDlg::OnBnClickedAcceptfaxorrecord)
	ON_BN_CLICKED(IDC_LOGIN, &Cwhu_VoiceCardDlg::OnBnClickedLogin)
	ON_BN_CLICKED(IDC_BRECORDMYNOTICE, &Cwhu_VoiceCardDlg::OnBnClickedBrecordmynotice)
	//ON_BN_CLICKED(IDC_TRYLISTEN, &Cwhu_VoiceCardDlg::OnBnClickedTrylisten)
	ON_NOTIFY(NM_RCLICK, IDC_MYCONTACTLIST, &Cwhu_VoiceCardDlg::OnNMRClickMycontactlist)
	ON_COMMAND(ID_32790, &Cwhu_VoiceCardDlg::OnTelPhoneDial)
	ON_COMMAND(ID_32791, &Cwhu_VoiceCardDlg::OnTelPhoneSendSingleNotice)
	ON_COMMAND(ID_32797, &Cwhu_VoiceCardDlg::OnOfficePhoneDial)
	ON_COMMAND(ID_32798, &Cwhu_VoiceCardDlg::OnOfficePhoneSendSingleNotice)
	ON_COMMAND(ID_32802, &Cwhu_VoiceCardDlg::OnSendGroupNotice)
	ON_COMMAND(ID_32803, &Cwhu_VoiceCardDlg::OnSendGroupFax)
	ON_COMMAND(ID_32793, &Cwhu_VoiceCardDlg::OnSendSingleFax)
	ON_BN_CLICKED(IDC_OPENSCANNER, &Cwhu_VoiceCardDlg::OnBnClickedOpenscanner)
	ON_COMMAND(ID_32804, &Cwhu_VoiceCardDlg::OnSeeLog)
	ON_NOTIFY(NM_RCLICK, IDC_MYLOGLIST, &Cwhu_VoiceCardDlg::OnNMRClickMyloglist)
	ON_BN_CLICKED(IDC_CHECK001, &Cwhu_VoiceCardDlg::OnBnClickedCheck001)
	ON_BN_CLICKED(IDC_CHECK002, &Cwhu_VoiceCardDlg::OnBnClickedCheck002)
	ON_BN_CLICKED(IDC_CHECK003, &Cwhu_VoiceCardDlg::OnBnClickedCheck003)
	ON_BN_CLICKED(IDC_CHECK004, &Cwhu_VoiceCardDlg::OnBnClickedCheck004)
	ON_BN_CLICKED(IDC_CHECK005, &Cwhu_VoiceCardDlg::OnBnClickedCheck005)
	ON_BN_CLICKED(IDC_CHECK006, &Cwhu_VoiceCardDlg::OnBnClickedCheck006)
	ON_BN_CLICKED(IDC_CHECK007, &Cwhu_VoiceCardDlg::OnBnClickedCheck007)
	ON_BN_CLICKED(IDC_CHECK008, &Cwhu_VoiceCardDlg::OnBnClickedCheck008)
	ON_WM_CTLCOLOR()
	//ON_COMMAND(ID_32808, &Cwhu_VoiceCardDlg::OnAddLogRemark)
	//ON_BN_CLICKED(IDC_LOGSAVE, &Cwhu_VoiceCardDlg::OnBnClickedLogsave)
	//ON_BN_CLICKED(IDC_LOGCANCEL, &Cwhu_VoiceCardDlg::OnBnClickedLogcancel)
	ON_COMMAND(ID_32809, &Cwhu_VoiceCardDlg::OnOpenLogFile)
	ON_COMMAND(ID_32810, &Cwhu_VoiceCardDlg::OnOpenFilePath)
	ON_COMMAND(ID_32805, &Cwhu_VoiceCardDlg::OnDeleteLog)
	ON_BN_CLICKED(IDC_LOGOUT, &Cwhu_VoiceCardDlg::OnBnClickedLogout)
	ON_WM_TIMER()
	ON_BN_CLICKED(IDC_QUITSYSTEM, &Cwhu_VoiceCardDlg::OnBnClickedQuitsystem)
	ON_COMMAND(ID_32811, &Cwhu_VoiceCardDlg::OnSendMail)
	ON_COMMAND(ID_32796, &Cwhu_VoiceCardDlg::OnSeeTelLog)
	ON_COMMAND(ID_32812, &Cwhu_VoiceCardDlg::OnRefreshLog)
	ON_WM_RBUTTONDBLCLK()
	ON_WM_NCRBUTTONUP()
	ON_WM_RBUTTONUP()
	ON_WM_RBUTTONDOWN()
	ON_COMMAND(ID_32813, &Cwhu_VoiceCardDlg::OnRefreshLog2)
	ON_COMMAND(ID_EDIT_COPY, &Cwhu_VoiceCardDlg::OnEditCopy)
	ON_COMMAND(ID_32794, &Cwhu_VoiceCardDlg::OnSeeFaxLog)
	ON_COMMAND(ID_32799, &Cwhu_VoiceCardDlg::OnSeeOfficeNumLog)
	//ON_BN_CLICKED(IDC_SHOWDUTY, &Cwhu_VoiceCardDlg::OnBnClickedShowduty) //��ʾֵ����Ϣ
	ON_NOTIFY(DTN_DATETIMECHANGE, IDC_DATETIMEPICKER1, &Cwhu_VoiceCardDlg::OnDtnDatetimechangeDatetimepicker1)
	ON_COMMAND(ID_32814, &Cwhu_VoiceCardDlg::OnSaveLogToFile)
	ON_COMMAND(ID_32815, &Cwhu_VoiceCardDlg::OnMenuSeeHistory)
	ON_COMMAND(ID_32820, &Cwhu_VoiceCardDlg::OnMenuRecordFolder)
	ON_COMMAND(ID_32821, &Cwhu_VoiceCardDlg::OnMenuNoticeFolder)
	ON_WM_NCRBUTTONUP()
	ON_WM_RBUTTONDOWN()
	ON_COMMAND(ID_32825, &Cwhu_VoiceCardDlg::OnMenuOpenScaner)
	ON_COMMAND(ID_32826, &Cwhu_VoiceCardDlg::OnScanerFolder)
	ON_COMMAND(ID_32827, &Cwhu_VoiceCardDlg::OnMenuFaxRecvFolder)
	ON_COMMAND(ID_32828, &Cwhu_VoiceCardDlg::OnMenuFaxSendFolder)
	ON_COMMAND(ID_MENUSHOWDUTY, &Cwhu_VoiceCardDlg::OnMenuShowDutyInfo)
	ON_NOTIFY(LVN_COLUMNCLICK, IDC_MYCONTACTLIST, &Cwhu_VoiceCardDlg::OnLvnColumnclickMycontactlist)
	ON_BN_CLICKED(IDC_SHOWDUTY, &Cwhu_VoiceCardDlg::OnBnClickedShowduty)
	ON_BN_CLICKED(IDC_BUTTONSHOWLOG, &Cwhu_VoiceCardDlg::OnBnClickedButtonshowlog)
	ON_BN_CLICKED(IDC_BUTTONFAX, &Cwhu_VoiceCardDlg::OnBnClickedButtonfax)
	ON_BN_CLICKED(IDC_BUTTONUPDATE, &Cwhu_VoiceCardDlg::OnBnClickedButtonupdate)
	ON_BN_CLICKED(IDC_BUTTONCONTACT, &Cwhu_VoiceCardDlg::OnBnClickedButtoncontact)
	ON_BN_CLICKED(IDC_CHECKRECORD, &Cwhu_VoiceCardDlg::OnBnClickedCheckrecord)
	ON_NOTIFY(NM_CUSTOMDRAW, IDC_MYCONTACTLIST, &Cwhu_VoiceCardDlg::OnNMCustomdrawMycontactlist)
	ON_BN_CLICKED(IDC_USER, &Cwhu_VoiceCardDlg::OnBnClickedUser)
	ON_BN_CLICKED(IDC_REGISTER, &Cwhu_VoiceCardDlg::OnBnClickedRegister)
	ON_BN_CLICKED(IDC_CHECKAUTOFORWARDFAX, &Cwhu_VoiceCardDlg::OnBnClickedCheckautoforwardfax)
END_MESSAGE_MAP()


// Cwhu_VoiceCardDlg message handlers

BOOL Cwhu_VoiceCardDlg::OnInitDialog()
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

	CString m_str = _T("�人��ѧ");
	CString m_pinyin = whu_HzTopy.HzToPinYin(m_str);


	whu_LogIn = false;

	int cx,cy; 
	cx = GetSystemMetrics(SM_CXSCREEN); 
	cy = GetSystemMetrics(SM_CYSCREEN);
	CRect rcTemp; 
	rcTemp.BottomRight() = CPoint(cx, cy); 
	rcTemp.TopLeft() = CPoint(0, 0); 

	CRect rect;  
	GetClientRect(&rect); //ȡ�ͻ�����С  
	Old.x=rect.right-rect.left;
	Old.y=rect.bottom-rect.top;
	MoveWindow(&rcTemp);
	
	resize();

	m_BShowCallImg = false;
	m_ProgressCtrl.SetRange(0,1000);
	SetTimer(1,100,NULL);
	SetTimer(2,2000,NULL); // ����绰-����ʱ����//


	whu_InitLogList();
	this->SetWindowText(_T("�人��ˮ��ַ�Ѵֵ��ϵͳ"));
	whu_CallImg= cvLoadImage("C:\\Whu_Data\\��ϵ��ͷ��\\İ����.jpg");
	whu_WelcomePic = cvLoadImage("C:\\Whu_Data\\welcome.jpg");
	whu_PicDown = cvLoadImage(_T("C:\\Whu_Data\\PicDown.jpg"));
	for (int i=0;i<10;i++)
	{
		whu_AllNos[i]=_T("");
	}
	//// TODO: Add extra initialization here
	/*if(whu_InitCtiBoard()==false)
	{
		return false;
	}*/
	m_BShowDutyList = false;
	//�û��������½//
	//whu_LinkToUserSql();
	//for (int i=0;i<whu_User[0].GetSize();i++)
	//{
	//	CString m_str = whu_User[0].GetAt(i);
	//	m_LogInComBox.AddString(m_str);
	//}

	
	//��Ǳ�ڵ�bug//
	int test1 = whu_User[0].GetSize();//6
	int test2 = whu_User[1].GetSize();//366
	int test3 = whu_User[2].GetSize();//366
	int test4 = whu_User[3].GetSize();//366
	//��Ǳ�ڵ�bug//


	//����ֵ���������쵼//
	CString m_DutyFile = _T("C:\\Whu_Data\\Excel�ļ�\\ֵ���.xlsx");
	CString m_Folder = _T("C:\\Whu_Data\\Excel�ļ�");
	if (whu_findfile(m_DutyFile,m_Folder) == true)
	{
		DutyHeadNum = Lin_ImportExcelToArr(m_DutyFile,whu_DutyPersons); //DutyHeadNum��whu_DutyPersons������//
	}
	else
	{
		DutyHeadNum = 0;

	}
	int ssss1 = whu_DutyPersons[0].GetSize();
	int ssss2 = whu_DutyPersons[1].GetSize();
	//whu_LinkToDutySql();//����ֵ����õ�ֵ������,��ɾ��//
	//���ô����쵼//
	int test = whu_DutyPersons[4].GetSize(); 
	for (int i=0;i<whu_DutyPersons[4].GetSize();i++)
	{
		CString m_str = whu_DutyPersons[4].GetAt(i);
		if ( m_ComboBoxLeader.FindStringExact(0,m_str)==CB_ERR&&m_str!=_T(""))
		{
			m_ComboBoxLeader.AddString(m_str);
		}
	}
	//while (m_MyContactList.GetHeaderCtrl()->GetItemCount()>0)
	//{
	//	m_MyContactList.DeleteColumn(0);
	//}
	//m_MyContactList.DeleteAllItems();



	////����պ���ĳ��ֵ�࣬�ͽ�����ΪĬ�ϵ�¼��//
	//GetLocalTime(&whu_SystemTime);
	//for (int i=0;i<whu_DutyPersons[1].GetSize();i++)
	//{
	//	CString time = whu_DutyPersons[1].GetAt(i);
	//	int pos = time.Find(_T("��"));
	//	time = time.Left(pos);
	//	pos = time.Find(_T("��"));
	//	int test3 = time.GetLength();
	//	time = time.Right(time.GetLength()-pos-2);
	//	CString day;
	//	day.Format("%d",whu_SystemTime.wDay);

	//	if(time == day)   
	//	{
	//		SetDlgItemText(IDC_LOGCOMBOBOX,whu_DutyPersons[2].GetAt(i));
	//	}

	//}
	CString m_FilePath = _T("C:\\Whu_Data\\Excel�ļ�\\ͨѶ¼.xlsx");
	m_Folder = _T("C:\\Whu_Data\\Excel�ļ�");
	if (whu_findfile(m_FilePath,m_Folder)==true)
	{
		ConHeadNum = Lin_ImportExcelToArr(_T("C:\\Whu_Data\\Excel�ļ�\\ͨѶ¼.xlsx"),m_ContactListData);
		Lin_ImportArrToList(m_ContactListData,ConHeadNum,m_MyContactList); //m_ContactListData��ֵ�ı��ˣ����ܻ�Ӱ����������//;
	}
	else{
		ConHeadNum = 0;
		CStringArray m_StandarHeadName;
		m_StandarHeadName.Add(_T("����"));
		m_StandarHeadName.Add(_T("��λ"));
		m_StandarHeadName.Add(_T("����"));
		m_StandarHeadName.Add(_T("ְ��"));
		m_StandarHeadName.Add(_T("�ֻ�"));
		m_StandarHeadName.Add(_T("�칫�绰"));
		m_StandarHeadName.Add(_T("�����"));
		m_StandarHeadName.Add(_T("����"));
		Lin_InitList(m_MyContactList,m_StandarHeadName);
	}

	////�����ǳ�ʼ��whu_AutoForwar
	CString m_ExcelFilePath = _T("C:\\Whu_Data\\Excel�ļ�\\�����Զ�ת����.xlsx");
	if (whu_findfile(m_ExcelFilePath,m_Folder)==true)
	{
		Lin_ImportAutoForArr(m_ExcelFilePath,whu_FaxAutoForwardArr);
	}

	CString m_UserFile = _T("C:\\Whu_Data\\Excel�ļ�\\�û���Ϣ.xlsx");
	CString m_UserFolder = _T("C:\\Whu_Data\\Excel�ļ�");
	CString m_UserFolder2 = _T("C:\\Whu_Data\\�����ļ�");
	CString m_UserFile2 = _T("C:\\Whu_Data\\�����ļ�\\�û���Ϣ_����.xlsx");
	if (whu_findfile(m_UserFile,m_UserFolder) == true)
	{
		Lin_ImportExcelToArr(m_UserFile,whu_User);
	}else if(whu_findfile(m_UserFile2,m_UserFolder2) == true)
	{
		Lin_ImportExcelToArr(m_UserFile2,whu_User);

	}else{
		AfxMessageBox(_T("�û���Ϣ�ļ���ʧ��ֻ�������Ա��½"));
		whu_User[0].Add(_T("�û���"));
		whu_User[1].Add(_T("����"));
		whu_User[2].Add(_T("״̬"));
		whu_User[3].Add(_T("Ȩ��"));
		whu_User[0].Add(_T("����Ա"));
		whu_User[1].Add(_T("whsw"));
		whu_User[2].Add(_T("��ע��"));
		whu_User[3].Add(_T("�����û�"));
	}
	for (int i=1;i<whu_User[0].GetSize();i++)
	{
		CString m_str = whu_User[0].GetAt(i);
		CString m_state = whu_User[2].GetAt(i);
		if (m_state == _T("��ע��"))
		{
			m_LogInComBox.AddString(m_str);
		}
	}

	whu_TaskMode=TASK_PHONEDIAL;
	whu_FaxChState = FAX_IDLE;
	whu_TrunkChState = TRK_IDLE;
	whu_UserChState = USER_IDLE;
	whu_RecordNotice =false;
	whu_Time = _T("");
	whu_NumHaveSend = 0;
	whu_GroupSendThread = NULL;
	whu_NumLength = 8;
	whu_AcceptCall = true;
	whu_AcceptFax = false;
	//m_DetailLogDlg.whu_ParentDlg = this;
	whu_RecordFile = _T("");
	m_BPhoneDialDone = false;
	m_BPhoneButtonPress = false;
	m_LogDeleteStr = _T("");
	m_BFlashPlaying = false;
	whu_BAutoForward = false;
	m_BQuitSystem = false;

	return TRUE;  // return TRUE  unless you set the focus to a control
}
void Cwhu_VoiceCardDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void Cwhu_VoiceCardDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// Center icon in client rectangle
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Draw the icon
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}


	if (whu_LogIn == false)
	{
		DrawPicToHDC(whu_WelcomePic,IDC_WELCOME2);
	}
	else if(m_BShowCallImg == true)
	{
		DrawPicToHDC(whu_CallImg,IDC_CALLPIC);
	}
	else
	{
		//DrawPicToHDC(whu_PicDown,IDC_DOWNPIC);
		GetLocalTime(&whu_SystemTime);
		CString buf=_T("");
		switch(whu_SystemTime.wDayOfWeek)
		{
		case 1:
			buf = _T("����һ");
			break;
		case 2:
			buf = _T("���ڶ�");
			break;
		case 3:
			buf = _T("������");
			break;
		case 4:
			buf = _T("������");
			break;
		case 5:
			buf = _T("������");
			break;
		case 6:
			buf = _T("������");
			break;
		case 0:
			buf = _T("������");
			break;
		default:
			break;
		}
		whu_Time.Format(_T("%d��%d��%d�� %s"),whu_SystemTime.wYear,whu_SystemTime.wMonth,whu_SystemTime.wDay,buf);
		SetDlgItemText(IDC_TIME,whu_Time);
	}

}

// The system calls this function to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR Cwhu_VoiceCardDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}
int  Cwhu_VoiceCardDlg::SeparateCallNum(CString m_AllNum,CString *m_SeparateNum)
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
LRESULT Cwhu_VoiceCardDlg::WindowProc(UINT message, WPARAM wParam, LPARAM lParam)
{
	// TODO: Add your specialized code here and/or call the base class

	int		nEventCode; //event ID
	int		nCh;		//channel ID
	int		nCheck;
	if(message > WM_USER)
	{
		nEventCode = message - WM_USER;

		//nCheck = SsmGetChType(0);  // ��ϯͨ����nCheck=2
		//nCheck = SsmGetChType(2); // ����ͨ����nCheck=0
		//nCheck = SsmGetChType(8); // ����ͨ����nCheck=9
		nCh=wParam;
		nCheck = SsmGetChType(nCh);
		if(nCheck == -1)
		{
			//AfxMessageBox("Fail to call SsmGetChType");
		}
		else if(nCheck == 5 || nCheck == 9)
		{
			FaxProc(nEventCode, 8, lParam);
		}
		else if(nCheck==2)
		{
			UserProc(nEventCode, 0, lParam);
		}
		else if(nCheck==0)
		{
			TrunkProc(nEventCode, 2, lParam);
		}
		
	}

	return CDialogEx::WindowProc(message, wParam, lParam);
}
void Cwhu_VoiceCardDlg::FaxProc(UINT event, WPARAM wParam, LPARAM lParam)
{
	int faxch=8;
	int nTrunkCh=2;
	switch (event)
	{
	case E_MSG_SEND_FAX:
		{
			if (whu_FaxChState == FAX_IDLE)
			{
				//�õ�Ҫ���͵Ĵ����ļ�whu_SendFaxFile//
				//do something

				if(SsmFaxStartSend(faxch,whu_SendFaxFile) == -1)	//send
				{
					TrunkProc(E_MSG_FAX_IDLE, nTrunkCh, faxch); 
				}
				else					
				{
					//int m_speed1 = SsmFaxGetSpeed(faxch); //�鿴���淢�͵�����
					//SsmFaxSetMaxSpeed(m_speed1); // ���ô��淢�͵�����
					//whu_FaxAllBites = SsmFaxGetAllBytes(faxch); //����whu_FaxAllBites = 0;
					m_AllFaxPage = fBmp_GetFileAllPage(whu_SendFaxFile); //�޷����öϵ�??????������//
					SetDlgItemText(IDC_SHOWINFO,_T("���ڷ��ʹ���...���Ժ�..."));
					GetDlgItem(IDC_SHOWINFO)->ShowWindow(true);
					GetDlgItem(IDC_STATICTASK)->ShowWindow(true);
					GetDlgItem(IDC_PROGRESS02)->ShowWindow(true);
					m_ProgressCtrl.SetPos(0);
					whu_FaxChState = FAX_CHECK_END; 
					whu_FaxSendState = FAX_SENDING;
					//������ģ��//
					m_DlgFax->PostMessage(WM_UPATEDATA, 0, (LPARAM)&whu_FaxSendState);
					//������ģ��//


				}
			} 		
		break;
		}
	case E_MSG_RCV_FAX:
		if (whu_FaxChState == FAX_IDLE)
		{
			////ֹͣ����������Ƶ
			if (m_BFlashPlaying == true)
			{
				//CRect m_min;
				//m_min.SetRectEmpty();
				//whu_flash.MoveWindow(m_min,TRUE);
				//whu_flash.StopPlay();
				m_BFlashPlaying = false;
			}
			//////�õ�whu_RecvFaxFile���ļ���
			whu_TaskMode = TASK_FAXRECV;
			GetLocalTime(&whu_SystemTime);
			whu_RecvFaxFile.Format(_T("C:\\Whu_Data\\����\\���յĴ���\\�յ�����%d��%d��%d��%dʱ%d��_����%s.tif\0"),whu_SystemTime.wYear,whu_SystemTime.wMonth,whu_SystemTime.wDay,whu_SystemTime.wHour,whu_SystemTime.wMinute,whu_NumberBuf);
			CString time = _T("");
			time.Format(_T("%d-%d-%d-%d:%d"),whu_SystemTime.wYear,whu_SystemTime.wMonth,whu_SystemTime.wDay,whu_SystemTime.wHour,whu_SystemTime.wMinute);
			whu_ContractData m_data = whu_GetContractData(whu_NumberBuf);
			whu_AddLog(time,m_data.DutyPerson,_T("���մ���"),m_data.Unit ,m_data.Name,whu_NumberBuf,whu_RecvFaxFile,_T(""),_T(""));
			
			if(SsmFaxStartReceive(faxch,whu_RecvFaxFile) == -1) //receive
			{
				TrunkProc(E_MSG_FAX_IDLE, nTrunkCh, faxch);
			}
			else
			{
				whu_FaxAllBites = SsmFaxGetAllBytes(faxch);
				SetDlgItemText(IDC_SHOWINFO,_T("���ڽ��մ���...���Ժ�..."));
				GetDlgItem(IDC_SHOWINFO)->ShowWindow(true);
				GetDlgItem(IDC_STATICTASK)->ShowWindow(true);
				GetDlgItem(IDC_PROGRESS02)->ShowWindow(true);
				m_ProgressCtrl.SetPos(0);
				whu_FaxChState = FAX_CHECK_END;
			}	
		}
		break;
	case E_PROC_FaxEnd:
		if(whu_FaxChState == FAX_CHECK_END) 
		{
			whu_NumHaveSend++;
			int m_check = lParam;
			whu_FaxChState = FAX_IDLE;
			whu_TrunkChState = TRK_IDLE;
			if (whu_TaskMode==TASK_FAXDIAL)
			{
				switch(m_check)
				{
				case 1 :
					//����ɹ�//
					TrunkProc(E_MSG_FAX_IDLE, nTrunkCh,faxch);
					//whu_HistoryDlg.whu_ModifyCheck(_T("�ɹ�"));
					whu_ModifyCheck(_T("�ɹ�"));
					whu_ShowFaxOutInfo(false);
					whu_FaxSendState = FAX_SUCCESS;
					//������ģ��//
					m_DlgFax->PostMessage(WM_UPATEDATA, 0, (LPARAM)&whu_FaxSendState);
					//������ģ��//
					//��Ӽ�¼
					break;
				case 2 :
					//����ʧ��//
					TrunkProc(E_MSG_FAX_IDLE, nTrunkCh, faxch);
					//whu_HistoryDlg.whu_ModifyCheck(_T("ʧ��"));
					whu_ModifyCheck(_T("ʧ��"));
					whu_FaxSendState = FAX_FAIL;
					//������ģ��//
					m_DlgFax->PostMessage(WM_UPATEDATA, 0, (LPARAM)&whu_FaxSendState);
					//������ģ��//
					whu_ShowFaxOutInfo(false);
					//��Ӽ�¼
					break;
				}//end of switch
			}
			else if (whu_TaskMode == TASK_FAXRECV)
			{
				switch(m_check)
				{
				case 1 :
					//����ɹ�//
					TrunkProc(E_MSG_FAX_IDLE, nTrunkCh,faxch);
					//whu_HistoryDlg.whu_ModifyCheck(_T("�ɹ�"));
					whu_ModifyCheck(_T("�ɹ�"));
					whu_FaxRecvState = FAX_SUCCESS;
					whu_ShowFaxInInfo(false);

					//��������Զ�ת������//��ģ�������//done
					if (whu_BAutoForward == true)
					{
						CStringArray FaxNumArr;
						int Num  = whu_GetAutoForwardFaxNum(whu_NumberBuf,whu_FaxAutoForwardArr,FaxNumArr);
						for (int i=0;i<FaxNumArr.GetSize();i++)
						{
							whu_AllNos[i] = FaxNumArr.GetAt(i);
						}
						whu_SendFaxFile = whu_RecvFaxFile;
						whu_TaskMode = TASK_FAXDIAL;
						if (whu_GroupSendThread==NULL)
						{
							whu_GroupSendThread=(Cwhu_GroupSendThread*)AfxBeginThread(RUNTIME_CLASS(Cwhu_GroupSendThread)); 
						}
						whu_GroupSendThread->PostThreadMessageA(E_MEG_HAVEGROUPFAX,NULL,(LPARAM)Num);
					}
					break;
				case 2 :
					//����ʧ��//
					TrunkProc(E_MSG_FAX_IDLE, nTrunkCh, faxch);
					//whu_HistoryDlg.whu_ModifyCheck(_T("ʧ��"));
					whu_ModifyCheck(_T("ʧ��"));
					whu_FaxRecvState = FAX_FAIL;
					whu_ShowFaxInInfo(false);
					//��Ӽ�¼
					break;
				}//end of switch
				
			}
			
			
		}
		break;
		// event indicating the hangup of the remote fax machine 
	case E_MSG_OFFLINE:
		{
			if(whu_FaxChState == FAX_CHECK_END)
			{
				SsmFaxStop(faxch);
				whu_FaxChState = FAX_IDLE;
				whu_SendFaxFile = _T("");
				whu_RecvFaxFile = _T("");
			}
		}
		break;
	}
}
void Cwhu_VoiceCardDlg::TrunkProc(UINT event, WPARAM wParam, LPARAM lParam)
{
	int AtrkCh = 2;
	int UserCh =0;
	int FaxCh = 8;
	int nNewState;
	switch(event)
	{
		case E_PROC_AutoDial:
			if (whu_TrunkChState==TRK_AutoDial)
			{
				DWORD dw = (DWORD)lParam; 
				switch(dw) { 
				case DIAL_VOICE://	called party answers the call 
					if (whu_TaskMode==TASK_PHONEDIAL) //��ǰ�ǲ���绰//
					{
						SsmStopSendTone(UserCh);
						SsmTalkWith(AtrkCh, UserCh);
						whu_TrunkChState = TRK_TALKING;
						whu_UserChState = USER_TALKING;
						SetDlgItemText(IDC_SHOWINFO,_T("�绰�ѽ�ͨ...������Ͳ�����绰..."));
						GetDlgItem(IDC_SHOWINFO)->ShowWindow(true);
						//UpdateData(TRUE);
						//if (whu_BRecord == true)
						//{
						//	GetLocalTime(&whu_SystemTime);
						//	whu_RecordFile=_T("");
						//	whu_RecordFile.Format(_T("C:\\Recorded\\����%s_%d.%d.%d.%d.%d.wav\0"),whu_NumberBuf,whu_SystemTime.wYear,whu_SystemTime.wMonth,whu_SystemTime.wDay,whu_SystemTime.wHour,whu_SystemTime.wMinute);
						//	SsmSetRecMixer(AtrkCh,TRUE,0);
						//	SsmSetRecVolume(AtrkCh,0);
						//	SsmRecToFile(UserCh,whu_RecordFile,-1,1,0xffffffff,0xffffffff,1);

						//}
					}
					else if (whu_TaskMode == TASK_PHONENOTICE)
					{
						if (whu_NoticeFile!=_T(""))
						{
							//SetDlgItemText(IDC_SHOWINFO,_T("�绰�ѽ�ͨ�����ڷ���֪ͨ..."));
							//GetDlgItem(IDC_SHOWINFO)->ShowWindow(true);
							whu_TrunkChState = TRK_NOTICEPLAYING;
							whu_NoticeState = NOTICE_PLAYING;
							//������ģ��//
							m_NoticeDlg->PostMessage(WM_UPATEDATA, 0, (LPARAM)&whu_NoticeState);
							//������ģ��//
							SsmPlayFile(AtrkCh,whu_NoticeFile,-1,0,0xffffffff);

						}
						else{
							AfxMessageBox(_T("�����ļ�Ϊ��"));
						}
					}
					else if(whu_TaskMode == TASK_FAXDIAL) //��Ӧ�������﷢,���յ�f2����ʱ����ܷ�//???������
					{
						
						if (SsmTalkWith(FaxCh, AtrkCh)==-1)
						{
							//����ʧ��//
						}
						whu_TrunkChState = TRK_TALKING;
						SsmPickup(AtrkCh);
						CString file = _T("C:\\Whu_Data\\�����ʹ���.wav");
						SsmPlayFile(AtrkCh,file,-1,0,0xffffffff);
						//SendMessage(WM_USER+E_MSG_SEND_FAX, FaxCh, AtrkCh);  //to faxproc    //���ʹ����������//
					}
					break;
				case DIAL_VOICEF2:   //�Է������Զ����յĻ�������Ӧ�û�������//������//
					if (SsmTalkWith(FaxCh, AtrkCh)==-1)
					{
						//����ʧ��//
						int ssss = 0;
						break;
					}
					whu_TrunkChState = TRK_FAXING;
					SendMessage(WM_USER+E_MSG_SEND_FAX, FaxCh, AtrkCh);  //to faxproc
					break;
				case 8://called party answers the call and signal F1 is detected
					{
						int ssss = 0;
						break;
					}
		////////3,4,10,11,12�����쳣///////////////////////

				case 2: //	ringback tone is detected after transmission of called id 
					{
						int ssss = 0;
						break;
					}			
				case 5://	silence on the line after ringback tone is detected.
					{
						int ssss = 0;
						break;
					}
				case 10://	the called party does not answer the call 
					{
						whu_NumHaveSend++;
						if (whu_TaskMode == TASK_FAXDIAL)
						{
							whu_FaxChState = FAX_IDLE;
							whu_ShowFaxOutInfo(false);
							whu_ModifyCheck(_T("�ȴ���ʱ"));
							whu_FaxSendState = FAX_NOANSWER;
							//������ģ��//
							m_DlgFax->PostMessage(WM_UPATEDATA, 0, (LPARAM)&whu_FaxSendState);
							//������ģ��//
						}
						else if (whu_TaskMode == TASK_PHONENOTICE)
						{
							whu_UserChState = USER_IDLE;
							whu_ShowCallOutInfo(false);
							whu_ModifyCheck(_T("�ȴ���ʱ"));
							whu_NoticeState = NOTICE_BUSY;
							//������ģ��//
							m_NoticeDlg->PostMessage(WM_UPATEDATA, 0, (LPARAM)&whu_NoticeState);
							//������ģ��//
						}
						else{
							whu_FaxChState = FAX_IDLE;
							whu_UserChState = USER_IDLE;
							whu_ShowCallOutInfo(false);
							whu_ModifyCheck(_T("�ȴ���ʱ"));
						}
						break;;
					}
				case 4://called party's line is busy
					{
						whu_NumHaveSend++;
						if (whu_TaskMode == TASK_FAXDIAL)
						{
							whu_FaxChState = FAX_IDLE;
							whu_ShowFaxOutInfo(false);
							whu_ModifyCheck(_T("�Է�æ"));
							whu_FaxSendState = FAX_NOANSWER;
							//������ģ��//
							m_DlgFax->PostMessage(WM_UPATEDATA, 0, (LPARAM)&whu_FaxSendState);
							//������ģ��//
						}
						else if (whu_TaskMode == TASK_PHONENOTICE)
						{
							whu_UserChState = USER_IDLE;
							whu_ShowCallOutInfo(false);
							whu_ModifyCheck(_T("�Է�æ"));
							whu_NoticeState = NOTICE_BUSY;
							//������ģ��//
							m_NoticeDlg->PostMessage(WM_UPATEDATA, 0, (LPARAM)&whu_NoticeState);
							//������ģ��//
						}
						else{
							whu_FaxChState = FAX_IDLE;
							whu_UserChState = USER_IDLE;
							whu_ShowCallOutInfo(false);
							whu_ModifyCheck(_T("�Է�æ"));
						}
						break;
					}
				case 3://no dialing tone is detected 
					{
						whu_NumHaveSend++; 
						if (whu_TaskMode == TASK_FAXDIAL)
						{
							whu_FaxChState = FAX_IDLE;
							whu_ShowFaxOutInfo(false);
							whu_ModifyCheck(_T("�޲�����"));
							//whu_CallNotice(6);
						}
						else if (whu_TaskMode == TASK_PHONENOTICE)
						{
							//whu_CallNotice(8);
							whu_UserChState = USER_IDLE;
							whu_ShowCallOutInfo(false);
							whu_ModifyCheck(_T("�޲�����"));
						}else{
							whu_FaxChState = FAX_IDLE;
							whu_UserChState = USER_IDLE;
							whu_ShowCallOutInfo(false);
							whu_ModifyCheck(_T("�޲�����"));
						}

						break;
					}
				case 6://silence on the line after completion of auto-dial 
					{
						whu_NumHaveSend++; 
						if (whu_TaskMode == TASK_FAXDIAL)
						{
							whu_FaxChState = FAX_IDLE;
							whu_ShowFaxOutInfo(false);
							whu_ModifyCheck(_T("�޲�����"));
							//whu_CallNotice(6);
						}
						else if (whu_TaskMode == TASK_PHONENOTICE)
						{
							//whu_CallNotice(8);
							whu_UserChState = USER_IDLE;
							whu_ShowCallOutInfo(false);
							whu_ModifyCheck(_T("�޲�����"));
						}else{
							whu_FaxChState = FAX_IDLE;
							whu_UserChState = USER_IDLE;
							whu_ShowCallOutInfo(false);
							whu_ModifyCheck(_T("�޲�����"));
						}

						break;
					}
				case 11://	failed auto-dial 
					{
						whu_NumHaveSend++; 
						if (whu_TaskMode == TASK_FAXDIAL)
						{
							whu_FaxChState = FAX_IDLE;
							whu_ShowFaxOutInfo(false);
							whu_ModifyCheck(_T("�Զ�����ʧ��"));
							//whu_CallNotice(6);
						}
						else if (whu_TaskMode == TASK_PHONENOTICE)
						{
							//whu_CallNotice(8);
							whu_UserChState = USER_IDLE;
							whu_ShowCallOutInfo(false);
							whu_ModifyCheck(_T("�Զ�����ʧ��"));
						}
						else{
							whu_FaxChState = FAX_IDLE;
							whu_UserChState = USER_IDLE;
							whu_ShowCallOutInfo(false);
							whu_ModifyCheck(_T("�Զ�����ʧ��"));
						}
						break;
					}
				case 12://unallocated number
					{
						whu_NumHaveSend++; 
						if (whu_TaskMode == TASK_FAXDIAL)
						{
							whu_FaxChState = FAX_IDLE;
							whu_ShowFaxOutInfo(false);
							whu_ModifyCheck(_T("�޲�����"));
							//whu_CallNotice(6);
						}
						else if (whu_TaskMode == TASK_PHONENOTICE)
						{
							//whu_CallNotice(8);
							whu_UserChState = USER_IDLE;
							whu_ShowCallOutInfo(false);
							whu_ModifyCheck(_T("�޲�����"));
						}else{
							whu_FaxChState = FAX_IDLE;
							whu_UserChState = USER_IDLE;
							whu_ShowCallOutInfo(false);
							whu_ModifyCheck(_T("�޲�����"));
						}
						//SsmHangup(AtrkCh);
						//SsmClearRxDtmfBuf(AtrkCh);
						//whu_TrunkChState = TRK_IDLE;
						//SsmSendTone(UserCh,1);					//send busy tone
						//whu_UserChState = USER_WAIT_HANGUP;
						//break;
					}
				case 1:
					{
						int ssss = whu_NumHaveSend;
						break;
					}
				case 13:
					{
						int ssss = whu_NumHaveSend;
						break;
					}
				case 14:
					{
						int ssss = whu_NumHaveSend;
						break;
					}
				case 15:
					{
						int ssss = whu_NumHaveSend;
						break;
					}
				case 16:
					{
						int ssss = whu_NumHaveSend;
						break;
					}
				default:
					{
						int ssss = whu_NumHaveSend;
						break;
					}
					
				}
			}
			break;
		case E_CHG_ChState:
			nNewState = (int)(lParam & 0xFFFF);	//new state
			if(nNewState == S_CALL_RINGING)		
			{
				if (whu_UserChState==USER_IDLE&&whu_TrunkChState ==TRK_IDLE)
				{
					
					char m_CallNum[30]={'\0'};
					SsmGetCallerId(AtrkCh,m_CallNum);
					whu_NumberBuf.Format("%s",m_CallNum);//�õ�����//
					SsmPickup(AtrkCh);
					CString file = _T("C:\\Whu_Data\\ͨ���밴1�����밴2.wav");
					SsmPlayFile(AtrkCh,file,-1,0,0xffffffff);
					whu_TrunkChState = TRK_WAITREMOTECHOOOSE;
					//whu_UserChState = USER_WAIT_LOCALPICKUP;
					//whu_TrunkChState = TRK_TALKING;
					//SsmStartRing(UserCh);
					
				} 
				else
				{
					SsmSendTone(AtrkCh,1);
					/*CString time = _T("");
					GetLocalTime(&whu_SystemTime);
					time.Format(_T("%d-%d-%d-%d:%d"),whu_SystemTime.wYear,whu_SystemTime.wMonth,whu_SystemTime.wDay,whu_SystemTime.wHour,whu_SystemTime.wMinute);
					whu_ContractData m_data = whu_GetContractData(whu_NumberBuf);
					whu_AddLog(time,m_data.DutyPerson,_T("����"),m_data.Unit ,m_data.Name,whu_NumberBuf,whu_SendFaxFile,_T("�������"),_T(""));*/

					//�Ӹ���������//
				}				
			}
			else if(nNewState == S_CALL_PENDING)  //�Է��һ�
			{
				if (whu_TaskMode == TASK_PHONENOTICE)
				{
					whu_TrunkChState = TRK_IDLE;
					whu_UserChState = USER_IDLE;
					whu_ShowCallOutInfo(false);
					whu_ModifyCheck(_T("�Է��һ�����ȷ���Ƿ�μ�"));
					whu_NoticeState = NOTICE_HANGUP;
					//������ģ��//
					m_NoticeDlg->PostMessage(WM_UPATEDATA, 0, (LPARAM)&whu_NoticeState);
					//������ģ��//
					SsmHangup(AtrkCh);
					whu_NumHaveSend++;
					if (whu_NumHaveSend == whu_NumOfNum)
					{
						whu_TaskMode = TASK_PHONEDIAL;

					}
				}
				else if(whu_TaskMode == TASK_PHONEDIAL)
				{
					SsmStopRecToFile(AtrkCh);
					SsmSendTone(UserCh,3);//����ϯͨ�����߹���
					SsmStopTalkWith(AtrkCh,UserCh);
					SsmHangup(AtrkCh);
					whu_UserChState = USER_IDLE;//
					whu_TrunkChState = TRK_IDLE;
					whu_ShowCallOutInfo(false);
					//whu_CallNotice(2);
				}
				else if(whu_TaskMode == TASK_PHONERECV){
					SsmStopTalkWith(AtrkCh,UserCh);
					//SsmSendTone(UserCh,3);//����ϯͨ�����߹���
					whu_TrunkChState = TRK_IDLE;
					whu_UserChState = USER_IDLE;
					whu_ShowCallInInfo(false);
					SsmHangup(AtrkCh);
					SsmStopRing(UserCh);
				}else if(whu_TaskMode == TASK_FAXDIAL){
					whu_NumHaveSend++;
					SsmHangup(AtrkCh);
					whu_TrunkChState = TRK_IDLE;
					whu_FaxChState = FAX_IDLE;
					whu_ShowFaxOutInfo(false);
					SsmStopTalkWith(AtrkCh,FaxCh);
					whu_ModifyCheck(_T("�Է���Ӧ��"));
					whu_FaxSendState = FAX_NOANSWER;
					//������ģ��//
					m_DlgFax->PostMessage(WM_UPATEDATA, 0, (LPARAM)&whu_FaxSendState);
					//������ģ��//
				}else if(whu_TaskMode==TASK_FAXRECV) //����Է�������ҵ��绰����ٷ������Ϣô��//
				{
					SsmStopTalkWith(AtrkCh,FaxCh);
					SsmHangup(AtrkCh);
					whu_TrunkChState = TRK_IDLE;
					whu_FaxChState = FAX_IDLE;
					whu_ShowFaxInInfo(false);
					
				}
				
			}
			else if (nNewState == E_CHG_PolarRvrsCount)  //��⼫�Է�ת�źţ��Լ��Է��Ƿ�������ժ�������������Ӫ�����뿪ͨ,����Ǯ//
			{
				/////_T("�������¼�����ǲ������룬��׼ȷ")////////
			}
			break;
		case E_CHG_RcvDTMF:   //���յ��Է��绰��dtmf�ź���//
			{
				DWORD dwDtmf = lParam & 0xFFFF;	//the latest DTMF received
				char temp ;
				temp = (char)dwDtmf;
				if (whu_TaskMode == TASK_PHONENOTICE)
				{
					if (temp == '*')
					{
						whu_TrunkChState = TRK_NOTICEPLAYING;
						//SsmStopPlayFile(AtrkCh);
						SsmPlayFile(AtrkCh,whu_NoticeFile,-1,0,0xffffffff);
					} 
					else if(temp =='#')
					{
						//���ȷ����Ϣ//
						whu_ShowCallOutInfo(false);
						SsmHangup(AtrkCh);
						//whu_HistoryDlg.whu_ModifyCheck(_T("�յ�֪ͨ��ȷ���μ�"));
						whu_ModifyCheck(_T("�յ�֪ͨ��ȷ���μ�"));
						whu_TrunkChState = TRK_IDLE;
						whu_NoticeState = NOTICE_ATTEND;
						//������ģ��//
						m_NoticeDlg->PostMessage(WM_UPATEDATA, 0, (LPARAM)&whu_NoticeState);
						//������ģ��//
						whu_NumHaveSend++;
						if (whu_NumHaveSend == whu_NumOfNum)
						{
							whu_TaskMode = TASK_PHONEDIAL;
						}
					}
					else if (temp =='0')
					{
						whu_ShowCallOutInfo(false);
						SsmHangup(AtrkCh);
						//whu_HistoryDlg.whu_ModifyCheck(_T("�յ�֪ͨ���޷��μ�"));
						whu_ModifyCheck(_T("�յ�֪ͨ���޷��μ�"));
						whu_NoticeState = NOTICE_NOATTEND;
						//������ģ��//
						m_NoticeDlg->PostMessage(WM_UPATEDATA, 0, (LPARAM)&whu_NoticeState);
						//������ģ��//
						whu_TrunkChState = TRK_IDLE;
						whu_NumHaveSend++;
						if (whu_NumHaveSend == whu_NumOfNum)
						{
							whu_TaskMode = TASK_PHONEDIAL;
						}
					}
				} 
				else if(whu_TrunkChState == TRK_WAITREMOTECHOOOSE)
				{
					if (temp == '1')
					{
						//ͨ��
						whu_UserChState = USER_WAIT_LOCALPICKUP;
						whu_TrunkChState = TRK_TALKING;
						whu_TaskMode = TASK_PHONERECV;
						SsmStartRing(UserCh);
						whu_ShowCallInInfo(true);
					} 
					else if(temp == '2')
					{
						whu_ShowFaxInInfo(true);
						whu_TaskMode=TASK_FAXRECV;
						whu_FaxRecvState = FAX_PREPARE;
						UpdateData(false);
						if (m_bAutoRecvFax == true)
						{
							SsmStopPlayFile(AtrkCh);
							whu_AcceptFax = true; //���Խ���
							SsmPickup(AtrkCh);
							SsmPlayFile(AtrkCh,_T("C:\\whu_Data\\�뷢�ʹ���.wav"),-1,0,0xffffffff);
							GetDlgItem(IDC_ACCEPTFAXORRECORD)->ShowWindow(false);
						}
						
					}
				}
				

			}
			break;
		case E_CHG_ToneAnalyze:	
			{
				DWORD  dwToneAnalyze = lParam & 0xFFFF;	//the result of tone analysis
				int m_Tone = dwToneAnalyze;
				if(m_Tone == 7)	//F1 tone 
				{
					if (whu_AcceptFax == true)
					{
						whu_TaskMode = TASK_FAXRECV;
						SsmPickup(AtrkCh);
						SsmTalkWith(AtrkCh,FaxCh);
						whu_TrunkChState = TRK_FAXING;
						char m_Num[20]={'\0'};
						SsmGetCallerId(AtrkCh,m_Num);
						whu_NumberBuf.Format("%s",m_Num);
						SendMessage(WM_USER+E_MSG_RCV_FAX, FaxCh, AtrkCh);
					}
					else
					{
						SsmPlayFile(AtrkCh,_T("C:\\Whu_Data\\ϵͳ��ʾ�ź���\\����δ���������Ժ�.wav"),-1,0,0xffffffff);
					}
				}
				else if(m_Tone == 8)	//F2 tone
				{
					SsmTalkWith(FaxCh, AtrkCh);
					whu_TrunkChState = TRK_FAXING;
					SendMessage(WM_USER+E_MSG_SEND_FAX, FaxCh, AtrkCh);
				}
			}
			break;
		case E_MSG_FAX_IDLE:
			SsmStopTalkWith(AtrkCh,FaxCh);
			HangUp(AtrkCh);
			break;
		case E_MSG_HAVEFAXTASK:
			if (whu_TrunkChState == TRK_IDLE)
			{
				CString time = _T("");
				GetLocalTime(&whu_SystemTime);
				time.Format(_T("%d-%d-%d-%d:%d"),whu_SystemTime.wYear,whu_SystemTime.wMonth,whu_SystemTime.wDay,whu_SystemTime.wHour,whu_SystemTime.wMinute);
				whu_ContractData m_data = whu_GetContractData(whu_NumberBuf);
				whu_AddLog(time,m_data.DutyPerson,_T("���ʹ���"),m_data.Unit ,m_data.Name,whu_NumberBuf,whu_SendFaxFile,_T(""),_T(""));
				whu_ShowFaxOutInfo(true);
				SsmPickup(AtrkCh);
				SsmAutoDial(AtrkCh,whu_NumberBuf);
				whu_TrunkChState = TRK_AutoDial;
			}

			break;
		case E_MEG_HAVENOTICETASK:
			{
				
				if (whu_TrunkChState ==TRK_IDLE)
				{
					CString time = _T("");
					GetLocalTime(&whu_SystemTime);
					time.Format(_T("%d-%d-%d-%d:%d"),whu_SystemTime.wYear,whu_SystemTime.wMonth,whu_SystemTime.wDay,whu_SystemTime.wHour,whu_SystemTime.wMinute);
					whu_ContractData m_data = whu_GetContractData(whu_NumberBuf);
					//whu_HistoryDlg.whu_AddHistoryList(time,m_data.DutyPerson,_T("Ⱥ��֪ͨ"),m_data.Unit ,m_data.Name,whu_NumberBuf,whu_RecordFile,_T(""),_T(""));
					whu_AddLog(time,m_data.DutyPerson,_T("Ⱥ��֪ͨ"),m_data.Unit ,m_data.Name,whu_NumberBuf,whu_NoticeFile,_T(""),_T(""));
					//GetDlgItem(IDC_SHOWINFO)->ShowWindow(true);
					//GetDlgItem(IDC_SHOWINFO)->SetWindowTextA(_T("����֪ͨ��������......"));
					//whu_ShowCallOutInfo(true);
					//whu_NoticeState = NOTICE_PREPARE;
					SsmPickup(AtrkCh);
					SsmAutoDial(AtrkCh,whu_NumberBuf);
					whu_TrunkChState = TRK_AutoDial;
				}
			}
			break;
		case E_PROC_PlayEnd:
			if (whu_TrunkChState == TRK_NOTICEPLAYING)
			{
				//��bug����֪ͨβ����δ����ʱ�������*�������Ļ���֪ͨβ��//
				whu_TrunkChState = TRK_TALKING;
				whu_NoticeFinalFile = _T("C:\\Whu_Data\\ϵͳ��ʾ�ź���\\֪ͨβ��.wav");
				SsmPlayFile(AtrkCh,whu_NoticeFinalFile,-1,0,0xffffffff);
			}
			if (whu_TrunkChState == TRK_BUSY)
			{
				SsmHangup(AtrkCh);
				whu_TrunkChState = TRK_IDLE;
			}
			break;
	}
}
void Cwhu_VoiceCardDlg::UserProc(UINT event, WPARAM wParam, LPARAM lParam)
{
	int ch;
	int AtrkCh=2;
	DWORD dw = (DWORD)lParam; 
	int	n;
	switch (event)
	{
	case E_CHG_HookState:
		ch = wParam;
		if (dw == 0)//station channel hangs up
		{
			SsmStopRing(ch);
			whu_NumberBuf.Empty();
			SsmStopRecToFile(ch);
			whu_ShowCallOutInfo(false);
			whu_ShowCallInInfo(false);
			if (whu_RecordNotice==true)
			{
				whu_RecordNotice = false;
				whu_RecordNoticeDone = true;
				whu_ShowRecordInfo(false);
				//SetDlgItemText(IDC_BRECORDMYNOTICE,_T("¼��"));
				//whu_PhoneDlg.SetDlgItemTextA(IDC_NOTICE,m_str);
			}
			switch (whu_UserChState)
			{
			case USER_GET_PHONE_NUM:
				SsmClearRxDtmfBuf(ch);
				whu_UserChState = USER_IDLE;
				whu_TrunkChState = TRK_IDLE;
				break;

			case USER_WAIT_REMOTE_PICKUP: 
				SsmHangup(AtrkCh);
				whu_TrunkChState = TRK_IDLE;
				SsmClearRxDtmfBuf(AtrkCh);
				SsmClearRxDtmfBuf(ch);
				whu_UserChState = USER_IDLE;
				break ;

			case USER_TALKING:
				SsmHangup(AtrkCh);
				SsmStopTalkWith(AtrkCh, ch);
				SsmClearRxDtmfBuf(AtrkCh);
				whu_TrunkChState = TRK_IDLE;
				whu_UserChState = USER_IDLE;
				break;

			case USER_WAIT_HANGUP:
				SsmStopSendTone(ch);					//stop sending busy tone
				SsmClearRxDtmfBuf(ch);
				whu_UserChState = USER_IDLE;
				whu_TrunkChState = TRK_IDLE;

				break;
			case USER_WAIT_LOCALPICKUP:
				SsmHangup(AtrkCh);
				SsmStopTalkWith(AtrkCh, ch);
				SsmStopRing(ch	);
				SsmStopSendTone(ch);
				SsmClearRxDtmfBuf(ch);
				whu_UserChState = USER_IDLE;
				whu_TrunkChState = TRK_IDLE;

			default:
				whu_UserChState = USER_IDLE;
				whu_TrunkChState = TRK_IDLE;
				break;
			}//end switch
		}
		else if (dw ==1) //station channel picks up
		{   
			int ss = 0;
			if (USER_IDLE == whu_UserChState )
			{	
				//whu_RecordNoticeͨ���绰¼��//
				if (whu_RecordNotice)
				{
					whu_RecordNoticeDone = false;
					SsmStopSendTone(ch);
					SsmPickup(ch);
					if (whu_NoticeFile==_T(""))
					{
						GetLocalTime(&whu_SystemTime);
						whu_NoticeFile.Format(_T("C:\\Whu_Data\\����֪ͨ\\֪ͨ%d��%d��%d��%dʱ%d��.wav\0"),whu_SystemTime.wYear,whu_SystemTime.wMonth,whu_SystemTime.wDay,whu_SystemTime.wHour,whu_SystemTime.wMinute);
					}			
					SsmRecToFile(ch,whu_NoticeFile,-1,1,0xffffffff,0xffffffff,1); //��������¼�������ɹ�����//
					
				}
				else
				{
					SsmSendTone(ch,0);
					whu_NumberBuf=_T("");
					SsmClearRxDtmfBuf(ch);					
					whu_UserChState = USER_GET_PHONE_NUM;
					whu_TaskMode = TASK_PHONEDIAL;
					whu_ShowCallOutInfo(true);
				}

			}
			else if(whu_UserChState == USER_WAIT_LOCALPICKUP)  //���صȴ�����
			{
				
				SsmPickup(AtrkCh);
				SsmPickup(ch);
				SsmStopSendTone(ch);
				SsmTalkWith(AtrkCh,ch);
				
			}
			else if (whu_UserChState == USER_TALKING)
			{
				SsmStopSendTone(ch);
				SsmPickup(ch);
				SsmTalkWith(AtrkCh,ch);
			}
		}
		break;
	case E_CHG_RcvDTMF:
		if (whu_UserChState == USER_GET_PHONE_NUM)
		{
			m_BPhoneButtonPress =true;
			ch=wParam;
			DWORD dwDtmf = lParam & 0xFFFF;	//the latest DTMF received
			char temp ;
			char m_latestchar = '\0';
			m_latestchar = (char)dwDtmf;
			whu_NumberBuf.Insert(100,m_latestchar);
			whu_ShowCallOutInfo(true);
			n = dw >> 16;
			int m_PhoNumLen=whu_NumLength;
			m_PhoNumLen=11;
			m_BPhoneDialDone = false;
			//if (n >= m_PhoNumLen)
			//{
			//	char m_NumBuf[30]={'\0'};
			//	SsmGetDtmfStr(ch,m_NumBuf);		//����ĵ绰����//
			//	whu_NumberBuf.Format("%s",m_NumBuf);
			//	if(whu_TrunkChState!=TRK_IDLE)			//no idle trunk channel available
			//	{
			//		SsmSendTone(ch,1);						// send busy tone	
			//		whu_UserChState = USER_WAIT_HANGUP;
			//	}
			//	else 
			//	{

			//		SsmPickup(AtrkCh);					
			//		//whu_CallNotice(1);
			//		if ( SsmAutoDial(AtrkCh, whu_NumberBuf) ==0)
			//		{							
			//			whu_TrunkChState = TRK_AutoDial;
			//			whu_UserChState = USER_WAIT_REMOTE_PICKUP;
			//			GetLocalTime(&whu_SystemTime);
			//			CString time = _T("");
			//			time.Format(_T("%d��%d��%d�� %d:%d"),whu_SystemTime.wYear,whu_SystemTime.wMonth,whu_SystemTime.wDay,whu_SystemTime.wHour,whu_SystemTime.wMinute);
			//			whu_ContractData m_data = whu_GetContractData(whu_NumberBuf);
			//			//whu_HistoryDlg.whu_AddHistoryList(time,m_data.DutyPerson,_T("�绰����"),m_data.Unit ,m_data.Name,whu_NumberBuf,whu_RecordFile,_T(""),_T(""));
			//			whu_AddLog(time,m_data.DutyPerson,_T("�绰����"),m_data.Unit ,m_data.Name,whu_NumberBuf,whu_RecordFile,_T(""),_T(""));
			//			SetDlgItemText(IDC_SHOWINFO,_T("���ڲ�����..."));
			//			GetDlgItem(IDC_SHOWINFO)->ShowWindow(true);

			//		      }
			//		else 
			//		{
			//			SsmHangup(AtrkCh);
			//			SsmClearRxDtmfBuf(AtrkCh);
			//			SsmSendTone(ch,1);				//send busy tone
			//			whu_UserChState = USER_WAIT_HANGUP;
			//		}
			//	}
			//}

		}
		break;
	default:
		break;
	}
}
bool Cwhu_VoiceCardDlg::whu_InitCtiBoard()
{
	//Initialization of CTI driver
	CString CurPath = _T("C:\\Whu_Data\\Setting");
	char ShIndex[260],ShConfig[260];
	//GetCurrentDirectory(200,CurPath);
	strcpy(ShConfig,CurPath);
	strcpy(ShIndex,CurPath);
	strcat(ShConfig,"\\ShConfig.ini");
	strcat(ShIndex,"\\ShIndex.ini");
	if( SsmStartCti(ShConfig,ShIndex) != 0) 
	{
		CString str1;
		SsmGetLastErrMsg(str1.GetBuffer(200));
		AfxMessageBox(str1, MB_OK) ;
		PostQuitMessage(0);
		return false;
	}
	//��ʼ����ϯͨ��0//
	SsmSetASDT(0, 1);
	whu_UserChState = USER_IDLE;
	//��ʼ������ͨ��2//
	int nDirection;
	if( SsmGetAutoCallDirection(2,&nDirection) == 1 ) //auto connection is supported
	{
		if( nDirection == 1 || nDirection == 2 ) //enable dial
		{
			
			whu_NumberBuf = _T("");
			SsmSetMinVocDtrEnergy( 2, 30000);					  
			whu_TrunkChState = TRK_IDLE;
		}
	}
	// set event callback handle.
	EVENT_SET_INFO EventSet;
	EventSet.dwWorkMode = EVENT_MESSAGE;	
	EventSet.lpHandlerParam = this->GetSafeHwnd();	
	SsmSetEvent(-1, -1, TRUE, &EventSet);
	return true;
}
BOOL Cwhu_VoiceCardDlg::DestroyWindow()
{
	// TODO: Add your specialized code here and/or call the base class
	SsmCloseCti();
	cvReleaseImage(&whu_CallImg);
	cvReleaseImage(&whu_WelcomePic);
	cvReleaseImage(&whu_PicDown);
	return CDialogEx::DestroyWindow();
}
void Cwhu_VoiceCardDlg::whu_SendNotice(CString *m_CallNum,int NumOfNum)
{
	whu_NumHaveSend = 0;
	int nTrunkCh = 2;
	int nUserCh = 0;
	for (int i=0;i<NumOfNum;i++)
	{
		whu_AllNos[i] = m_CallNum[i];
	}
	whu_NumOfNum = NumOfNum;
	if (whu_NoticeFile == _T("")) //whu_RecordFileΪ��ǰ¼�Ƶ��ļ�
	{
		//AfxMessageBox(_T("�����ȵ������\n¼����ť¼��֪ͨ"));
		return ;
	}
	whu_TaskMode = TASK_PHONENOTICE;
	//whu_NumberBuf = whu_AllNos[0];
	if (whu_GroupSendThread==NULL)
	{
		whu_GroupSendThread=(Cwhu_GroupSendThread*)AfxBeginThread(RUNTIME_CLASS(Cwhu_GroupSendThread)); 
	}
	whu_GroupSendThread->PostThreadMessageA(E_MSG_HAVEGROUPNOTICE,NULL,(LPARAM)NumOfNum);
	//SendMessage(WM_USER+E_MEG_HAVENOTICETASK, nTrunkCh, NULL);

}
void Cwhu_VoiceCardDlg::whu_Dial(CString m_Num)
{
	int AtrkCh = 2;
	int ch =0;
	whu_NumberBuf = m_Num;
	whu_TaskMode = TASK_PHONEDIAL;
	if (whu_TrunkChState == TRK_IDLE&&whu_UserChState == USER_IDLE)
	{
		SsmPickup(AtrkCh);
		whu_ShowCallOutInfo(true);
		CButton* pBtn = (CButton*)GetDlgItem(IDC_CHECKRECORD);
		pBtn->SetCheck(1);
		if ( SsmAutoDial(AtrkCh, whu_NumberBuf) ==0)
		{	
			GetLocalTime(&whu_SystemTime);
			CString time = _T("");
			time.Format(_T("%d-%d-%d-%d:%d"),whu_SystemTime.wYear,whu_SystemTime.wMonth,whu_SystemTime.wDay,whu_SystemTime.wHour,whu_SystemTime.wMinute);
			whu_ContractData m_data = whu_GetContractData(whu_NumberBuf);
			whu_RecordFile.Format(_T("C:\\Whu_Data\\¼��\\%d��%d��%d��%dʱ%d��_����%s.wav\0"),whu_SystemTime.wYear,whu_SystemTime.wMonth,whu_SystemTime.wDay,whu_SystemTime.wHour,whu_SystemTime.wMinute,whu_NumberBuf);
			SsmSetRecMixer(AtrkCh,TRUE,0);
			SsmRecToFile(ch,whu_RecordFile,-1,1,0xffffffff,0xffffffff,1); //��������¼�������ɹ�����//
			whu_AddLog(time,m_data.DutyPerson,_T("�绰����"),m_data.Unit ,m_data.Name,whu_NumberBuf,whu_RecordFile,_T(""),_T(""));
			whu_TrunkChState = TRK_AutoDial;
			whu_UserChState = USER_WAIT_REMOTE_PICKUP;
			SetDlgItemText(IDC_SHOWINFO,_T("���ڲ�����..."));
			GetDlgItem(IDC_SHOWINFO)->ShowWindow(true);

		}
		else 
		{
			SsmHangup(AtrkCh);
			SsmClearRxDtmfBuf(AtrkCh);
			SsmSendTone(ch,1);				//send busy tone
			whu_UserChState = USER_WAIT_HANGUP;
		}

	}
	else
	{
		whu_TrunkChState = TRK_IDLE;
		whu_UserChState = USER_IDLE;
		AfxMessageBox(_T("��·æ,���Ժ�..."));

	}

}
void Cwhu_VoiceCardDlg::whu_SendFax(CString *m_CallNum,CString whu_DocFile,int num)
{

	int nTrunkCh = 2;
	int nFaxCh = 8;
	whu_NumOfNum = num;
	for (int i=0;i<num;i++)
	{
		whu_AllNos[i] = m_CallNum[i];
	}

	GetLocalTime(&whu_SystemTime);
	CString whu_FaxFile = _T("");
	//if (whu_NumOfNum>1)
	//{
	//	whu_FaxFile.Format(_T("C:\\Recorded\\Ⱥ������%s_%d.%d.%d.%d.%d.tif\0"),m_CallNum[0],whu_SystemTime.wYear,whu_SystemTime.wMonth,whu_SystemTime.wDay,whu_SystemTime.wHour,whu_SystemTime.wMinute);
	//}else{
		/*whu_FaxFile.Format(_T("C:\\Recorded\\����%s.tif\0"),	whu_SendFaxFile);*/
	//}
	whu_FaxFile.Format(_T("C:\\Whu_Data\\����\\���͵Ĵ���\\%s.tif\0"),whu_SendFaxFileTitle);
	CString m_folder = _T("C:\\Whu_Data\\����\\���͵Ĵ���");
	if (whu_findfile(whu_FaxFile,m_folder)== false)
	{
		DocumentToTiff(whu_DocFile,whu_FaxFile);
	}
	whu_NumberBuf = whu_AllNos[0];
	whu_SendFaxFile = whu_FaxFile;
	whu_TaskMode = TASK_FAXDIAL;
	if (whu_GroupSendThread==NULL)
	{
		whu_GroupSendThread=(Cwhu_GroupSendThread*)AfxBeginThread(RUNTIME_CLASS(Cwhu_GroupSendThread)); 
	}
	whu_GroupSendThread->PostThreadMessageA(E_MEG_HAVEGROUPFAX,NULL,(LPARAM)num);

	//SendMessage(WM_USER+E_MSG_HAVEFAXTASK,nTrunkCh,nFaxCh);
}
void Cwhu_VoiceCardDlg::whu_ShowRecordInfo(bool show)
{
	static bool playing = false;
	if (show == true)
	{

		SetDlgItemText(IDC_SHOWINFO,_T("������Ͳ¼��...�ҵ��绰¼������..."));
		GetDlgItem(IDC_SHOWINFO)->ShowWindow(true);

		//GetDlgItem(IDC_BRECORDMYNOTICE)->ShowWindow(false);
		//GetDlgItem(IDC_TRYLISTEN)->ShowWindow(false);
		GetDlgItem(IDC_OPENSCANNER)->ShowWindow(false);
		GetDlgItem(IDC_CHECKRECVFAX)->ShowWindow(false);
		GetDlgItem(IDC_CHECKAUTOFORWARDFAX)->ShowWindow(false);
		
		GetDlgItem(IDC_MYCONTACTLIST)->ShowWindow(false);
		GetDlgItem(IDC_SHAIXUAN)->ShowWindow(false);
		GetDlgItem(IDC_SEARCH)->ShowWindow(false);
		GetDlgItem(IDC_CHECK001)->ShowWindow(false);
		GetDlgItem(IDC_CHECK002)->ShowWindow(false);
		GetDlgItem(IDC_CHECK003)->ShowWindow(false);
		GetDlgItem(IDC_CHECK004)->ShowWindow(false);
		GetDlgItem(IDC_CHECK005)->ShowWindow(false);
		GetDlgItem(IDC_CHECK006)->ShowWindow(false);
		GetDlgItem(IDC_CHECK007)->ShowWindow(false);
		GetDlgItem(IDC_CHECK008)->ShowWindow(false);
		GetDlgItem(IDC_LOGOUT)->ShowWindow(false);
		//GetDlgItem(IDC_MYLOGLIST)->ShowWindow(false);
		//GetDlgItem(IDC_SHOWDUTY)->ShowWindow(false);
		//GetDlgItem(IDC_DATETIMEPICKER1)->ShowWindow(false);
		//GetDlgItem(IDC_DOWNPIC)->ShowWindow(false);
		//if (playing == false)
		//{
		//	CRect m_rec1;
		//	CRect m_flash ;
		//	//whu_Tab.GetClientRect(&m_flash);
		//	GetDlgItem(IDC_FLASH)->GetWindowRect(&m_flash);
		//	GetDlgItem(IDC_DUTYINFO)->GetClientRect(&m_rec1);
		//	m_flash.left = m_rec1.right+10;
		//	GetDlgItem(IDC_TIME)->GetClientRect(&m_rec1);
		//	m_flash.top = m_rec1.bottom +10;
		//	m_flash.right = m_flash.right -30;

		//	whu_flash.MoveWindow(m_flash,TRUE);
		//	CString strFileName = "C:\\Whu_Data\\RecordNotice.swf";//5.swf��flash�ļ������֣���flash�ļ����ڹ���Ŀ¼�¡�
		//	whu_flash.LoadMovie(0, strFileName);
		//	whu_flash.Play();
		//	playing = true;
		//}

	} 
	else
	{

		//SetDlgItemText(IDC_SHOWINFO,_T("¼�������������ڿ���ѡ����ϵ���Ҽ�����֪ͨ..."));
		GetDlgItem(IDC_SHOWINFO)->ShowWindow(false);
		//GetDlgItem(IDC_BRECORDMYNOTICE)->ShowWindow(true);
		//GetDlgItem(IDC_TRYLISTEN)->ShowWindow(true);
		GetDlgItem(IDC_OPENSCANNER)->ShowWindow(true);
		GetDlgItem(IDC_CHECKRECVFAX)->ShowWindow(true);
		GetDlgItem(IDC_CHECKAUTOFORWARDFAX)->ShowWindow(true);
		GetDlgItem(IDC_MYCONTACTLIST)->ShowWindow(true);
		//GetDlgItem(IDC_MYLOGLIST)->ShowWindow(true);
		GetDlgItem(IDC_SHAIXUAN)->ShowWindow(true);
		GetDlgItem(IDC_SEARCH)->ShowWindow(true);
		GetDlgItem(IDC_CHECK001)->ShowWindow(true);
		GetDlgItem(IDC_CHECK002)->ShowWindow(true);
		GetDlgItem(IDC_CHECK003)->ShowWindow(true);
		GetDlgItem(IDC_CHECK004)->ShowWindow(true);
		GetDlgItem(IDC_CHECK005)->ShowWindow(true);
		GetDlgItem(IDC_CHECK006)->ShowWindow(true);
		GetDlgItem(IDC_CHECK007)->ShowWindow(true);
		GetDlgItem(IDC_CHECK008)->ShowWindow(true);
		GetDlgItem(IDC_LOGOUT)->ShowWindow(true);
		//GetDlgItem(IDC_MYLOGLIST)->ShowWindow(true);
		//GetDlgItem(IDC_SHOWDUTY)->ShowWindow(true);
		//GetDlgItem(IDC_DATETIMEPICKER1)->ShowWindow(true);
		//GetDlgItem(IDC_DOWNPIC)->ShowWindow(true);
		/*if (playing == true)
		{
			CRect m_min;
			m_min.SetRectEmpty();
			whu_flash.MoveWindow(m_min,TRUE);
			whu_flash.StopPlay();
			playing = false;
		}*/
		CString time = _T("");
		GetLocalTime(&whu_SystemTime);
		time.Format(_T("%d-%d-%d-%d:%d"),whu_SystemTime.wYear,whu_SystemTime.wMonth,whu_SystemTime.wDay,whu_SystemTime.wHour,whu_SystemTime.wMinute);
		whu_ContractData m_data = whu_GetContractData(_T("11"));
		//whu_HistoryDlg.whu_AddHistoryList(time,m_data.DutyPerson,_T("¼��֪ͨ"),_T("") ,_T(""),_T(""),whu_RecordFile,_T(""),_T(""));
		whu_AddLog(time,m_data.DutyPerson,_T("¼��֪ͨ"),_T("") ,_T(""),_T(""),whu_NoticeFile,_T(""),_T(""));
	}

}
void Cwhu_VoiceCardDlg::whu_ShowCallInInfo(bool m_b)
{
	whu_BRecord = false;
	static bool playing = false;
	if (m_b == true)
	{
		//whu_Tab.ShowWindow(false);
		GetDlgItem(IDC_CALLGROUP)->ShowWindow(true);
		GetDlgItem(IDC_NAME1)->ShowWindow(true);
		GetDlgItem(IDC_NAME2)->ShowWindow(true);
		GetDlgItem(IDC_DANWEI1)->ShowWindow(true);
		GetDlgItem(IDC_DANWEI2)->ShowWindow(true);
		GetDlgItem(IDC_BUMEN1)->ShowWindow(true);
		GetDlgItem(IDC_BUMEN2)->ShowWindow(true);
		GetDlgItem(IDC_ZHIWU1)->ShowWindow(true);
		GetDlgItem(IDC_ZHIWU2)->ShowWindow(true);
		GetDlgItem(IDC_HAOMA1)->ShowWindow(true);
		GetDlgItem(IDC_HAOMA2)->ShowWindow(true);
		GetDlgItem(IDC_CALLPIC)->ShowWindow(true);
		SetDlgItemText(IDC_ACCEPTFAXORRECORD,_T("¼��"));
		GetDlgItem(IDC_ACCEPTFAXORRECORD)->ShowWindow(true);
		SetDlgItemText(IDC_REFUSE,_T("�ܽ�"));
		GetDlgItem(IDC_REFUSE)->ShowWindow(true);
		SetDlgItemText(IDC_SHOWINFO,_T("������...������...������..."));
		GetDlgItem(IDC_SHOWINFO)->ShowWindow(true);

		GetDlgItem(IDC_BRECORDMYNOTICE)->ShowWindow(false);
		//GetDlgItem(IDC_TRYLISTEN)->ShowWindow(false);
		GetDlgItem(IDC_OPENSCANNER)->ShowWindow(false);
		GetDlgItem(IDC_CHECKRECVFAX)->ShowWindow(false);
		GetDlgItem(IDC_CHECKAUTOFORWARDFAX)->ShowWindow(false);
		GetDlgItem(IDC_MYCONTACTLIST)->ShowWindow(false);
		//GetDlgItem(IDC_MYLOGLIST)->ShowWindow(false);
		GetDlgItem(IDC_SHAIXUAN)->ShowWindow(false);
		GetDlgItem(IDC_CHECK001)->ShowWindow(false);
		GetDlgItem(IDC_CHECK002)->ShowWindow(false);
		GetDlgItem(IDC_CHECK003)->ShowWindow(false);
		GetDlgItem(IDC_CHECK004)->ShowWindow(false);
		GetDlgItem(IDC_CHECK005)->ShowWindow(false);
		GetDlgItem(IDC_CHECK006)->ShowWindow(false);
		GetDlgItem(IDC_CHECK007)->ShowWindow(false);
		GetDlgItem(IDC_CHECK008)->ShowWindow(false);
		GetDlgItem(IDC_SEARCH)->ShowWindow(false);
		//GetDlgItem(IDC_LOGREMARK)->ShowWindow(false);
		GetDlgItem(IDC_SHOWDUTY)->ShowWindow(false);
		//GetDlgItem(IDC_DATETIMEPICKER1)->ShowWindow(false);
		//GetDlgItem(IDC_LOGSAVE)->ShowWindow(false);
		//GetDlgItem(IDC_LOGCANCEL)->ShowWindow(false);
		GetDlgItem(IDC_LOGOUT)->ShowWindow(false);
		GetDlgItem(IDC_BUTTONUPDATE)->ShowWindow(false);
	//	GetDlgItem(IDC_DOWNPIC)->ShowWindow(false);


		//if (m_BFlashPlaying == false)
		//{
		//	CRect m_rec1;
		//	CRect m_flash ;
		//	//whu_Tab.GetClientRect(&m_flash);
		//	GetDlgItem(IDC_FLASH)->GetWindowRect(&m_flash);
		//	GetDlgItem(IDC_DUTYINFO)->GetClientRect(&m_rec1);
		//	m_flash.left = m_rec1.right+10;
		//	GetDlgItem(IDC_TIME)->GetClientRect(&m_rec1);
		//	m_flash.top = m_rec1.bottom +10;
		//	whu_flash.MoveWindow(m_flash,TRUE);
		//	CString strFileName = "C:\\Whu_Data\\callin.swf";//5.swf��flash�ļ������֣���flash�ļ����ڹ���Ŀ¼�¡�
		//	whu_flash.LoadMovie(0, strFileName);
		//	whu_flash.Play();
		//	m_BFlashPlaying = true;

		//	//����һ�����ڲ�����//
		//	//Cwhu_RecvCallDlg *m_DlgRecvCall = new Cwhu_RecvCallDlg();
		//	//m_DlgRecvCall->whu_ParentDlg = this;
		//	//m_DlgRecvCall->Create(IDD_DIALOGRECVCALL,this);
		//	//m_DlgRecvCall->ShowWindow(SW_SHOW);

		//}
		whu_ShowNumInfo(whu_NumberBuf);
		m_BShowCallImg = true;
		CString time = _T("");
		GetLocalTime(&whu_SystemTime);
		time.Format(_T("%d-%d-%d-%d:%d"),whu_SystemTime.wYear,whu_SystemTime.wMonth,whu_SystemTime.wDay,whu_SystemTime.wHour,whu_SystemTime.wMinute);
		whu_ContractData m_data = whu_GetContractData(whu_NumberBuf);
		//whu_HistoryDlg.whu_AddHistoryList(time,m_data.DutyPerson,_T("�绰����"),m_data.Unit ,m_data.Name,whu_NumberBuf,whu_RecordFile,_T(""),_T(""));
		whu_AddLog(time,m_data.DutyPerson,_T("�绰����"),m_data.Unit ,m_data.Name,whu_NumberBuf,whu_RecordFile,_T(""),_T(""));
	} 
	else
	{
		//whu_Tab.ShowWindow(true);
		GetDlgItem(IDC_CALLGROUP)->ShowWindow(false);
		GetDlgItem(IDC_NAME1)->ShowWindow(false);
		GetDlgItem(IDC_NAME2)->ShowWindow(false);
		GetDlgItem(IDC_DANWEI1)->ShowWindow(false);
		GetDlgItem(IDC_DANWEI2)->ShowWindow(false);
		GetDlgItem(IDC_BUMEN1)->ShowWindow(false);
		GetDlgItem(IDC_BUMEN2)->ShowWindow(false);
		GetDlgItem(IDC_ZHIWU1)->ShowWindow(false);
		GetDlgItem(IDC_ZHIWU2)->ShowWindow(false);
		GetDlgItem(IDC_HAOMA1)->ShowWindow(false);
		GetDlgItem(IDC_HAOMA2)->ShowWindow(false);
		GetDlgItem(IDC_ACCEPTFAXORRECORD)->ShowWindow(false);
		GetDlgItem(IDC_REFUSE)->ShowWindow(false);
		GetDlgItem(IDC_SHOWINFO)->ShowWindow(false);
		GetDlgItem(IDC_CALLPIC)->ShowWindow(false);
		GetDlgItem(IDC_BRECORDMYNOTICE)->ShowWindow(true);
		//GetDlgItem(IDC_TRYLISTEN)->ShowWindow(true);
		GetDlgItem(IDC_OPENSCANNER)->ShowWindow(true);
		GetDlgItem(IDC_CHECKRECVFAX)->ShowWindow(true);
		GetDlgItem(IDC_CHECKAUTOFORWARDFAX)->ShowWindow(true);
		GetDlgItem(IDC_MYCONTACTLIST)->ShowWindow(true);
		//GetDlgItem(IDC_MYLOGLIST)->ShowWindow(true);
		GetDlgItem(IDC_SHAIXUAN)->ShowWindow(true);
		GetDlgItem(IDC_CHECK001)->ShowWindow(true);
		GetDlgItem(IDC_CHECK002)->ShowWindow(true);
		GetDlgItem(IDC_CHECK003)->ShowWindow(true);
		GetDlgItem(IDC_CHECK004)->ShowWindow(true);
		GetDlgItem(IDC_CHECK005)->ShowWindow(true);
		GetDlgItem(IDC_CHECK006)->ShowWindow(true);
		GetDlgItem(IDC_CHECK007)->ShowWindow(true);
		GetDlgItem(IDC_CHECK008)->ShowWindow(true);
		GetDlgItem(IDC_SEARCH)->ShowWindow(true);
		GetDlgItem(IDC_LOGOUT)->ShowWindow(true);
		GetDlgItem(IDC_SHOWDUTY)->ShowWindow(true);
		//GetDlgItem(IDC_DATETIMEPICKER1)->ShowWindow(true);
		GetDlgItem(IDC_BUTTONUPDATE)->ShowWindow(true);
		//GetDlgItem(IDC_DOWNPIC)->ShowWindow(true);
		whu_AcceptCall = true;
		m_BShowCallImg = false;
		/*if (m_BFlashPlaying == true)
		{
			CRect m_min;
			m_min.SetRectEmpty();
			whu_flash.MoveWindow(m_min,TRUE);
			whu_flash.StopPlay();
			m_BFlashPlaying = false;
		}*/

	}
	
}
void Cwhu_VoiceCardDlg::whu_ShowCallOutInfo(bool m_b)
{
	whu_BRecord = false;
	static bool playing = false;
	if (m_b)
	{
		GetDlgItem(IDC_CALLGROUP)->ShowWindow(true);
		GetDlgItem(IDC_CALLPIC)->ShowWindow(true);
		GetDlgItem(IDC_NAME1)->ShowWindow(true);
		GetDlgItem(IDC_NAME2)->ShowWindow(true);
		GetDlgItem(IDC_DANWEI1)->ShowWindow(true);
		GetDlgItem(IDC_DANWEI2)->ShowWindow(true);
		GetDlgItem(IDC_BUMEN1)->ShowWindow(true);
		GetDlgItem(IDC_BUMEN2)->ShowWindow(true);
		GetDlgItem(IDC_ZHIWU1)->ShowWindow(true);
		GetDlgItem(IDC_ZHIWU2)->ShowWindow(true);
		GetDlgItem(IDC_HAOMA1)->ShowWindow(true);
		GetDlgItem(IDC_HAOMA2)->ShowWindow(true);
		GetDlgItem(IDC_CHECKRECORD)->ShowWindow(true);
		//SetDlgItemText(IDC_ACCEPTFAXORRECORD,_T("¼��"));
		//GetDlgItem(IDC_ACCEPTFAXORRECORD)->ShowWindow(true);
		SetDlgItemText(IDC_REFUSE,_T("�Ҷ�"));
		GetDlgItem(IDC_REFUSE)->ShowWindow(true);
		GetDlgItem(IDC_BRECORDMYNOTICE)->ShowWindow(false);
		//GetDlgItem(IDC_TRYLISTEN)->ShowWindow(false);
		GetDlgItem(IDC_OPENSCANNER)->ShowWindow(false);
		GetDlgItem(IDC_CHECKRECVFAX)->ShowWindow(false);
		GetDlgItem(IDC_CHECKAUTOFORWARDFAX)->ShowWindow(false);
		GetDlgItem(IDC_MYCONTACTLIST)->ShowWindow(false);
		//GetDlgItem(IDC_MYLOGLIST)->ShowWindow(false);
		GetDlgItem(IDC_SHAIXUAN)->ShowWindow(false);
		GetDlgItem(IDC_CHECK001)->ShowWindow(false);
		GetDlgItem(IDC_CHECK002)->ShowWindow(false);
		GetDlgItem(IDC_CHECK003)->ShowWindow(false);
		GetDlgItem(IDC_CHECK004)->ShowWindow(false);
		GetDlgItem(IDC_CHECK005)->ShowWindow(false);
		GetDlgItem(IDC_CHECK006)->ShowWindow(false);
		GetDlgItem(IDC_CHECK007)->ShowWindow(false);
		GetDlgItem(IDC_CHECK008)->ShowWindow(false);
		GetDlgItem(IDC_SEARCH)->ShowWindow(false);
		//GetDlgItem(IDC_LOGREMARK)->ShowWindow(false);
		//GetDlgItem(IDC_LOGSAVE)->ShowWindow(false);
		//GetDlgItem(IDC_LOGCANCEL)->ShowWindow(false);
		GetDlgItem(IDC_LOGOUT)->ShowWindow(false);
		GetDlgItem(IDC_SHOWDUTY)->ShowWindow(false);
		GetDlgItem(IDC_BUTTONFAX)->ShowWindow(false);
		GetDlgItem(IDC_USER)->ShowWindow(false);
		
		GetDlgItem(IDC_BUTTONSHOWLOG)->ShowWindow(false);
		GetDlgItem(IDC_BUTTONCONTACT)->ShowWindow(false);

		//GetDlgItem(IDC_DATETIMEPICKER1)->ShowWindow(false);
		GetDlgItem(IDC_BUTTONUPDATE)->ShowWindow(false);
	//	GetDlgItem(IDC_DOWNPIC)->ShowWindow(false);
		//if (m_BFlashPlaying == false) //2013��6��13��ע��
		//{
		//	CRect m_rec1;
		//	CRect m_flash ;
		//	GetDlgItem(IDC_FLASH)->GetWindowRect(&m_flash);
		//	GetDlgItem(IDC_DUTYINFO)->GetClientRect(&m_rec1);
		//	m_flash.left = m_rec1.right+10;
		//	GetDlgItem(IDC_TIME)->GetClientRect(&m_rec1);
		//	m_flash.top = m_rec1.bottom +10;
		//	whu_flash.MoveWindow(m_flash,TRUE);
		//	CString strFileName = "C:\\Whu_Data\\callout.swf";//5.swf��flash�ļ������֡�
		//	whu_flash.LoadMovie(0, strFileName);
		//	whu_flash.Play();
		//	m_BFlashPlaying = true;
		//}
		SetDlgItemTextA(IDC_HAOMA2,whu_NumberBuf);
		whu_ShowNumInfo(whu_NumberBuf);
		m_BShowCallImg = true;
	} 
	else
	{
		//whu_Tab.ShowWindow(true);
		GetDlgItem(IDC_CALLPIC)->ShowWindow(false);
		GetDlgItem(IDC_CALLGROUP)->ShowWindow(false);
		GetDlgItem(IDC_NAME1)->ShowWindow(false);
		GetDlgItem(IDC_NAME2)->ShowWindow(false);
		GetDlgItem(IDC_DANWEI1)->ShowWindow(false);
		GetDlgItem(IDC_DANWEI2)->ShowWindow(false);
		GetDlgItem(IDC_BUMEN1)->ShowWindow(false);
		GetDlgItem(IDC_BUMEN2)->ShowWindow(false);
		GetDlgItem(IDC_ZHIWU1)->ShowWindow(false);
		GetDlgItem(IDC_ZHIWU2)->ShowWindow(false);
		GetDlgItem(IDC_HAOMA1)->ShowWindow(false);
		GetDlgItem(IDC_HAOMA2)->ShowWindow(false);
		GetDlgItem(IDC_REFUSE)->ShowWindow(false);
		GetDlgItem(IDC_ACCEPTFAXORRECORD)->ShowWindow(false);
		GetDlgItem(IDC_CHECKRECORD)->ShowWindow(false);
		SetDlgItemText(IDC_SHOWINFO,_T(""));
		GetDlgItem(IDC_SHOWINFO)->ShowWindow(false);

		GetDlgItem(IDC_BRECORDMYNOTICE)->ShowWindow(true);
		//GetDlgItem(IDC_TRYLISTEN)->ShowWindow(true);
		GetDlgItem(IDC_OPENSCANNER)->ShowWindow(true);
		GetDlgItem(IDC_CHECKRECVFAX)->ShowWindow(true);
		GetDlgItem(IDC_CHECKAUTOFORWARDFAX)->ShowWindow(true);
		GetDlgItem(IDC_MYCONTACTLIST)->ShowWindow(true);
		//GetDlgItem(IDC_MYLOGLIST)->ShowWindow(true);
		GetDlgItem(IDC_SHAIXUAN)->ShowWindow(true);
		GetDlgItem(IDC_CHECK001)->ShowWindow(true);
		GetDlgItem(IDC_CHECK002)->ShowWindow(true);
		GetDlgItem(IDC_CHECK003)->ShowWindow(true);
		GetDlgItem(IDC_CHECK004)->ShowWindow(true);
		GetDlgItem(IDC_CHECK005)->ShowWindow(true);
		GetDlgItem(IDC_CHECK006)->ShowWindow(true);
		GetDlgItem(IDC_CHECK007)->ShowWindow(true);
		GetDlgItem(IDC_CHECK008)->ShowWindow(true);
		GetDlgItem(IDC_SEARCH)->ShowWindow(true);
		GetDlgItem(IDC_LOGOUT)->ShowWindow(true);
		GetDlgItem(IDC_SHOWDUTY)->ShowWindow(true);
		//GetDlgItem(IDC_DATETIMEPICKER1)->ShowWindow(true);
		GetDlgItem(IDC_BUTTONUPDATE)->ShowWindow(true);
		GetDlgItem(IDC_BUTTONFAX)->ShowWindow(true);
		GetDlgItem(IDC_USER)->ShowWindow(true);
		GetDlgItem(IDC_BUTTONSHOWLOG)->ShowWindow(true);
		GetDlgItem(IDC_BUTTONCONTACT)->ShowWindow(true);
		//GetDlgItem(IDC_DOWNPIC)->ShowWindow(true);
		//if (m_BFlashPlaying == true)
		//{
		//	CRect m_min;
		//	m_min.SetRectEmpty();
		//	whu_flash.MoveWindow(m_min,TRUE);
		//	whu_flash.StopPlay();
		//	m_BFlashPlaying = false;
		//}
		m_BShowCallImg = false;
	}
}
void Cwhu_VoiceCardDlg::whu_ShowFaxInInfo(bool show)
{
	static bool playing = false;
	if (show == true)
	{
		//whu_Tab.ShowWindow(false);
		//UpdateData(false);
		//if(m_bAutoRecvFax==true){
			//GetDlgItem(IDC_STATICTASK)->ShowWindow(true);
			//GetDlgItem(IDC_PROGRESS02)->ShowWindow(true); //���մ����޷���ʾ����������ȡ����byte��δ֪��
		//}else{
			//GetDlgItem(IDC_STATICTASK)->ShowWindow(false);
			//GetDlgItem(IDC_PROGRESS02)->ShowWindow(false); //���մ����޷���ʾ����������ȡ����byte��δ֪��

		//}
		GetDlgItem(IDC_CALLGROUP)->ShowWindow(true);
		GetDlgItem(IDC_NAME1)->ShowWindow(true);
		GetDlgItem(IDC_NAME2)->ShowWindow(true);
		GetDlgItem(IDC_DANWEI1)->ShowWindow(true);
		GetDlgItem(IDC_DANWEI2)->ShowWindow(true);
		GetDlgItem(IDC_BUMEN1)->ShowWindow(true);
		GetDlgItem(IDC_BUMEN2)->ShowWindow(true);
		GetDlgItem(IDC_ZHIWU1)->ShowWindow(true);
		GetDlgItem(IDC_ZHIWU2)->ShowWindow(true);
		GetDlgItem(IDC_HAOMA1)->ShowWindow(true);
		GetDlgItem(IDC_HAOMA2)->ShowWindow(true);
		GetDlgItem(IDC_CALLPIC)->ShowWindow(true);
		SetDlgItemText(IDC_REFUSE,_T("�ܽ�"));
		GetDlgItem(IDC_REFUSE)->ShowWindow(true);
		SetDlgItemText(IDC_ACCEPTFAXORRECORD,_T("����"));
		GetDlgItem(IDC_ACCEPTFAXORRECORD)->ShowWindow(true);

		GetDlgItem(IDC_BRECORDMYNOTICE)->ShowWindow(false);
		//GetDlgItem(IDC_TRYLISTEN)->ShowWindow(false);
		GetDlgItem(IDC_OPENSCANNER)->ShowWindow(false);
		GetDlgItem(IDC_CHECKRECVFAX)->ShowWindow(false);
		GetDlgItem(IDC_MYCONTACTLIST)->ShowWindow(false);
		//GetDlgItem(IDC_MYLOGLIST)->ShowWindow(false);
		GetDlgItem(IDC_SHAIXUAN)->ShowWindow(false);
		GetDlgItem(IDC_CHECK001)->ShowWindow(false);
		GetDlgItem(IDC_CHECK002)->ShowWindow(false);
		GetDlgItem(IDC_CHECK003)->ShowWindow(false);
		GetDlgItem(IDC_CHECK004)->ShowWindow(false);
		GetDlgItem(IDC_CHECK005)->ShowWindow(false);
		GetDlgItem(IDC_CHECK006)->ShowWindow(false);
		GetDlgItem(IDC_CHECK007)->ShowWindow(false);
		GetDlgItem(IDC_CHECK008)->ShowWindow(false);
		GetDlgItem(IDC_SEARCH)->ShowWindow(false);
		//GetDlgItem(IDC_LOGREMARK)->ShowWindow(false);
		//GetDlgItem(IDC_LOGSAVE)->ShowWindow(false);
		//GetDlgItem(IDC_LOGCANCEL)->ShowWindow(false);
		GetDlgItem(IDC_LOGOUT)->ShowWindow(false);
		GetDlgItem(IDC_SHOWDUTY)->ShowWindow(false);
		//GetDlgItem(IDC_DATETIMEPICKER1)->ShowWindow(false);
		GetDlgItem(IDC_BUTTONUPDATE)->ShowWindow(false);
		GetDlgItem(IDC_BUTTONFAX)->ShowWindow(false);
		GetDlgItem(IDC_USER)->ShowWindow(false);
		GetDlgItem(IDC_BUTTONSHOWLOG)->ShowWindow(false);
		GetDlgItem(IDC_BUTTONCONTACT)->ShowWindow(false);
		//GetDlgItem(IDC_DOWNPIC)->ShowWindow(false);
		//if (m_BFlashPlaying == false)
		//{
		//	CRect m_rec1;
		//	CRect m_flash ;
		//	//whu_Tab.GetClientRect(&m_flash);
		//	GetDlgItem(IDC_FLASH)->GetWindowRect(&m_flash);
		//	GetDlgItem(IDC_DUTYINFO)->GetClientRect(&m_rec1);
		//	m_flash.left = m_rec1.right+10;
		//	GetDlgItem(IDC_TIME)->GetClientRect(&m_rec1);
		//	m_flash.top = m_rec1.bottom +10;
		//	whu_flash.MoveWindow(m_flash,TRUE);
		//	CString strFileName = "C:\\whu_Data\\faxin.swf";//5.swf��flash�ļ������֣���flash�ļ����ڹ���Ŀ¼�¡�
		//	whu_flash.LoadMovie(0, strFileName);
		//	whu_flash.Play();
		//	m_BFlashPlaying = true;
		//}
		whu_ShowNumInfo(whu_NumberBuf);
		m_BShowCallImg = true;
		//CString time = _T("");
		//GetLocalTime(&whu_SystemTime);
		//time.Format(_T("%d��%d��%d�� %d:%d"),whu_SystemTime.wYear,whu_SystemTime.wMonth,whu_SystemTime.wDay,whu_SystemTime.wHour,whu_SystemTime.wMinute);
		//whu_ContractData m_data = whu_GetContractData(whu_NumberBuf);
		////whu_HistoryDlg.whu_AddHistoryList(time,m_data.DutyPerson,_T("���մ���"),m_data.Unit ,m_data.Name,whu_NumberBuf,whu_RecvFaxFile,_T(""),_T(""));
		//whu_AddLog(time,m_data.DutyPerson,_T("���մ���"),m_data.Unit ,m_data.Name,whu_NumberBuf,whu_RecvFaxFile,_T(""),_T(""));
	} 
	else
	{
		//whu_Tab.ShowWindow(true);
		GetDlgItem(IDC_STATICTASK)->ShowWindow(false);
		GetDlgItem(IDC_PROGRESS02)->ShowWindow(false);
		GetDlgItem(IDC_CALLGROUP)->ShowWindow(false);
		GetDlgItem(IDC_NAME1)->ShowWindow(false);
		GetDlgItem(IDC_NAME2)->ShowWindow(false);
		GetDlgItem(IDC_DANWEI1)->ShowWindow(false);
		GetDlgItem(IDC_DANWEI2)->ShowWindow(false);
		GetDlgItem(IDC_BUMEN1)->ShowWindow(false);
		GetDlgItem(IDC_BUMEN2)->ShowWindow(false);
		GetDlgItem(IDC_ZHIWU1)->ShowWindow(false);
		GetDlgItem(IDC_ZHIWU2)->ShowWindow(false);
		GetDlgItem(IDC_HAOMA1)->ShowWindow(false);
		GetDlgItem(IDC_HAOMA2)->ShowWindow(false);
		GetDlgItem(IDC_REFUSE)->ShowWindow(false);
		GetDlgItem(IDC_ACCEPTFAXORRECORD)->ShowWindow(false);
		GetDlgItem(IDC_CALLPIC)->ShowWindow(false);
		GetDlgItem(IDC_SHOWINFO)->ShowWindow(false);
		GetDlgItem(IDC_BRECORDMYNOTICE)->ShowWindow(true);
		//GetDlgItem(IDC_TRYLISTEN)->ShowWindow(true);
		GetDlgItem(IDC_OPENSCANNER)->ShowWindow(true);
		GetDlgItem(IDC_CHECKRECVFAX)->ShowWindow(true);
		GetDlgItem(IDC_MYCONTACTLIST)->ShowWindow(true);
		//GetDlgItem(IDC_MYLOGLIST)->ShowWindow(true);
		GetDlgItem(IDC_SHAIXUAN)->ShowWindow(true);
		GetDlgItem(IDC_CHECK001)->ShowWindow(true);
		GetDlgItem(IDC_CHECK002)->ShowWindow(true);
		GetDlgItem(IDC_CHECK003)->ShowWindow(true);
		GetDlgItem(IDC_CHECK004)->ShowWindow(true);
		GetDlgItem(IDC_CHECK005)->ShowWindow(true);
		GetDlgItem(IDC_CHECK006)->ShowWindow(true);
		GetDlgItem(IDC_CHECK007)->ShowWindow(true);
		GetDlgItem(IDC_CHECK008)->ShowWindow(true);
		GetDlgItem(IDC_SEARCH)->ShowWindow(true);
		GetDlgItem(IDC_LOGOUT)->ShowWindow(true);
		GetDlgItem(IDC_SHOWDUTY)->ShowWindow(true);
		GetDlgItem(IDC_BUTTONFAX)->ShowWindow(true);
		GetDlgItem(IDC_USER)->ShowWindow(true);
		GetDlgItem(IDC_BUTTONSHOWLOG)->ShowWindow(true);
		GetDlgItem(IDC_BUTTONCONTACT)->ShowWindow(true);
		//GetDlgItem(IDC_DATETIMEPICKER1)->ShowWindow(true);
		GetDlgItem(IDC_BUTTONUPDATE)->ShowWindow(true);
		//GetDlgItem(IDC_DOWNPIC)->ShowWindow(true);
		whu_AcceptFax = false;
		m_BShowCallImg = false;
		//if (m_BFlashPlaying == true)
		//{
		//	CRect m_min;
		//	m_min.SetRectEmpty();
		//	whu_flash.MoveWindow(m_min,TRUE);
		//	whu_flash.StopPlay();
		//	m_BFlashPlaying = false;
		//}
		if (whu_RecvFaxFile!=_T(""))
		{
			ShellExecute(NULL, "open", whu_RecvFaxFile,NULL,NULL, SW_SHOWNORMAL );
			//ShellExecute(NULL, "open", _T("C:\\Whu_Data\\����\\���յĴ���"),NULL,NULL, SW_SHOWNORMAL );	
		}
		
	}

}
void Cwhu_VoiceCardDlg::whu_ShowFaxOutInfo(bool show)
{
	static bool playing = false;
	if (show)
	{
		//GetDlgItem(IDC_STATICTASK)->ShowWindow(true);
		//GetDlgItem(IDC_PROGRESS02)->ShowWindow(true);
		GetDlgItem(IDC_CALLGROUP)->ShowWindow(true);
		GetDlgItem(IDC_CALLPIC)->ShowWindow(true);
		GetDlgItem(IDC_NAME1)->ShowWindow(true);
		GetDlgItem(IDC_NAME2)->ShowWindow(true);
		GetDlgItem(IDC_DANWEI1)->ShowWindow(true);
		GetDlgItem(IDC_DANWEI2)->ShowWindow(true);
		GetDlgItem(IDC_BUMEN1)->ShowWindow(true);
		GetDlgItem(IDC_BUMEN2)->ShowWindow(true);
		GetDlgItem(IDC_ZHIWU1)->ShowWindow(true);
		GetDlgItem(IDC_ZHIWU2)->ShowWindow(true);
		GetDlgItem(IDC_HAOMA1)->ShowWindow(true);
		GetDlgItem(IDC_HAOMA2)->ShowWindow(true);
		GetDlgItem(IDC_SHOWINFO)->ShowWindow(true);
		GetDlgItem(IDC_SHOWINFO)->SetWindowTextA(_T("׼���������ȴ��Է�����..."));
		SetDlgItemText(IDC_REFUSE,_T("�Ҷ�"));
		GetDlgItem(IDC_REFUSE)->ShowWindow(true);
		GetDlgItem(IDC_BRECORDMYNOTICE)->ShowWindow(false);
		//GetDlgItem(IDC_TRYLISTEN)->ShowWindow(false);
		GetDlgItem(IDC_OPENSCANNER)->ShowWindow(false);
		GetDlgItem(IDC_CHECKRECVFAX)->ShowWindow(false);
		GetDlgItem(IDC_CHECKAUTOFORWARDFAX)->ShowWindow(false);
		GetDlgItem(IDC_MYCONTACTLIST)->ShowWindow(false);
		//GetDlgItem(IDC_MYLOGLIST)->ShowWindow(false);
		GetDlgItem(IDC_SHAIXUAN)->ShowWindow(false);
		GetDlgItem(IDC_CHECK001)->ShowWindow(false);
		GetDlgItem(IDC_CHECK002)->ShowWindow(false);
		GetDlgItem(IDC_CHECK003)->ShowWindow(false);
		GetDlgItem(IDC_CHECK004)->ShowWindow(false);
		GetDlgItem(IDC_CHECK005)->ShowWindow(false);
		GetDlgItem(IDC_CHECK006)->ShowWindow(false);
		GetDlgItem(IDC_CHECK007)->ShowWindow(false);
		GetDlgItem(IDC_CHECK008)->ShowWindow(false);
		GetDlgItem(IDC_SEARCH)->ShowWindow(false);
		//GetDlgItem(IDC_LOGREMARK)->ShowWindow(false);
		//GetDlgItem(IDC_LOGSAVE)->ShowWindow(false);
		//GetDlgItem(IDC_LOGCANCEL)->ShowWindow(false);
		GetDlgItem(IDC_LOGOUT)->ShowWindow(false);
		GetDlgItem(IDC_SHOWDUTY)->ShowWindow(false);
		GetDlgItem(IDC_BUTTONFAX)->ShowWindow(false);
		GetDlgItem(IDC_USER)->ShowWindow(false);
		GetDlgItem(IDC_BUTTONSHOWLOG)->ShowWindow(false);
		GetDlgItem(IDC_BUTTONCONTACT)->ShowWindow(false);
		//GetDlgItem(IDC_DATETIMEPICKER1)->ShowWindow(false);
		GetDlgItem(IDC_BUTTONUPDATE)->ShowWindow(false);
		//GetDlgItem(IDC_DOWNPIC)->ShowWindow(false);
		//whu_Tab.ShowWindow(false);
		//if (m_BFlashPlaying == false)
		//{
		//	CRect m_rec1;
		//	CRect m_flash ;
		//	//whu_Tab.GetClientRect(&m_flash);
		//	GetDlgItem(IDC_FLASH)->GetWindowRect(&m_flash);
		//	GetDlgItem(IDC_DUTYINFO)->GetClientRect(&m_rec1);
		//	m_flash.left = m_rec1.right+10;
		//	GetDlgItem(IDC_TIME)->GetClientRect(&m_rec1);
		//	m_flash.top = m_rec1.bottom +10;
		//	whu_flash.MoveWindow(m_flash,TRUE);
		//	CString strFileName = "C:\\Whu_Data\\faxout.swf";//5.swf��flash�ļ������֣���flash�ļ����ڹ���Ŀ¼�¡�
		//	whu_flash.LoadMovie(0, strFileName);
		//	whu_flash.Play();
		//	m_BFlashPlaying = true;
		//}
		//SetDlgItemTextA(IDC_HAOMA2,whu_NumberBuf);
		whu_ShowNumInfo(whu_NumberBuf);
		m_BShowCallImg = true;
	} 
	else
	{
		//whu_Tab.ShowWindow(true);
		GetDlgItem(IDC_STATICTASK)->ShowWindow(false);
		GetDlgItem(IDC_PROGRESS02)->ShowWindow(false);
		GetDlgItem(IDC_CALLPIC)->ShowWindow(false);
		GetDlgItem(IDC_CALLGROUP)->ShowWindow(false);
		GetDlgItem(IDC_NAME1)->ShowWindow(false);
		GetDlgItem(IDC_NAME2)->ShowWindow(false);
		GetDlgItem(IDC_DANWEI1)->ShowWindow(false);
		GetDlgItem(IDC_DANWEI2)->ShowWindow(false);
		GetDlgItem(IDC_BUMEN1)->ShowWindow(false);
		GetDlgItem(IDC_BUMEN2)->ShowWindow(false);
		GetDlgItem(IDC_ZHIWU1)->ShowWindow(false);
		GetDlgItem(IDC_ZHIWU2)->ShowWindow(false);
		GetDlgItem(IDC_HAOMA1)->ShowWindow(false);
		GetDlgItem(IDC_HAOMA2)->ShowWindow(false);
		GetDlgItem(IDC_REFUSE)->ShowWindow(false);
		GetDlgItem(IDC_SHOWINFO)->ShowWindow(false);
		GetDlgItem(IDC_STATICTASK)->ShowWindow(false);
		GetDlgItem(IDC_BRECORDMYNOTICE)->ShowWindow(true);
		//GetDlgItem(IDC_TRYLISTEN)->ShowWindow(true);
		GetDlgItem(IDC_OPENSCANNER)->ShowWindow(true);
		GetDlgItem(IDC_CHECKRECVFAX)->ShowWindow(true);
		GetDlgItem(IDC_CHECKAUTOFORWARDFAX)->ShowWindow(true);
		GetDlgItem(IDC_MYCONTACTLIST)->ShowWindow(true);
		//GetDlgItem(IDC_MYLOGLIST)->ShowWindow(true);
		GetDlgItem(IDC_SHAIXUAN)->ShowWindow(true);
		GetDlgItem(IDC_CHECK001)->ShowWindow(true);
		GetDlgItem(IDC_CHECK002)->ShowWindow(true);
		GetDlgItem(IDC_CHECK003)->ShowWindow(true);
		GetDlgItem(IDC_CHECK004)->ShowWindow(true);
		GetDlgItem(IDC_CHECK005)->ShowWindow(true);
		GetDlgItem(IDC_CHECK006)->ShowWindow(true);
		GetDlgItem(IDC_CHECK007)->ShowWindow(true);
		GetDlgItem(IDC_CHECK008)->ShowWindow(true);
		GetDlgItem(IDC_SEARCH)->ShowWindow(true);
		GetDlgItem(IDC_LOGOUT)->ShowWindow(true);
		GetDlgItem(IDC_SHOWDUTY)->ShowWindow(true);
		GetDlgItem(IDC_BUTTONFAX)->ShowWindow(true);
		GetDlgItem(IDC_USER)->ShowWindow(true);
		GetDlgItem(IDC_BUTTONSHOWLOG)->ShowWindow(true);
		GetDlgItem(IDC_BUTTONCONTACT)->ShowWindow(true);
		//GetDlgItem(IDC_DATETIMEPICKER1)->ShowWindow(true);
		GetDlgItem(IDC_BUTTONUPDATE)->ShowWindow(true);
		//GetDlgItem(IDC_DOWNPIC)->ShowWindow(true);
		/*if (m_BFlashPlaying == true)
		{
			CRect m_min;
			m_min.SetRectEmpty();
			whu_flash.MoveWindow(m_min,TRUE);
			whu_flash.StopPlay();
			m_BFlashPlaying = false;
		}*/
		whu_FaxAllBites = 0;
		whu_FaxDoneBites = 0;
		m_BShowCallImg = false;

	}
}
void Cwhu_VoiceCardDlg::OnBnClickedRefuse()
{
	// TODO: Add your control notification handler code here
	CString m_str;
	GetDlgItemText(IDC_REFUSE,m_str);
	if (m_str == _T("�ܽ�"))
	{
		int AtrkCh = 2;
		int UserCh = 0;
		SsmStopRing(UserCh);
		SsmPickup(AtrkCh);
		SsmPlayFile(AtrkCh,_T("C:\\Whu_Data\\����æ.wav"),-1,0,0xffffffff);
		whu_TrunkChState = TRK_BUSY;
		whu_UserChState = USER_IDLE;
		whu_ShowCallInInfo(false);
		whu_ShowFaxInInfo(false);
		whu_AcceptCall = false;
		whu_AcceptFax = false;
	}
	if (m_str == _T("�Ҷ�"))
	{
		int AtrkCh = 2;
		int UserCh = 0;
		whu_UserChState = USER_IDLE;
		whu_TrunkChState = TRK_IDLE;
		whu_NumHaveSend++;
		whu_ModifyCheck(_T("����ȡ��"));
		if (whu_TaskMode == TASK_FAXDIAL)
		{
			
			SsmFaxStop(AtrkCh);
			SsmHangup(AtrkCh);
			whu_ShowFaxOutInfo(false);
			whu_FaxSendState = FAX_CANCEL;
			//������ģ��//
			m_DlgFax->PostMessage(WM_UPATEDATA, 0, (LPARAM)&whu_FaxSendState);
			//������ģ��//
		} 
		else if(whu_TaskMode == TASK_PHONEDIAL||whu_TaskMode ==TASK_PHONENOTICE)
		{
			SsmHangup(AtrkCh);
			whu_ShowCallOutInfo(false);
		}

	}
}
int Cwhu_VoiceCardDlg::DocumentToTiff(CString m_Inputfile,CString &m_OutputFile)
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
void Cwhu_VoiceCardDlg::OnBnClickedAcceptfaxorrecord()
{
	// TODO: Add your control notification handler code here
	int AtrkCh =2;
	int ch = 0;
	CString m_str = _T("");
	GetDlgItemText(IDC_ACCEPTFAXORRECORD,m_str);
	if (m_str == _T("¼��"))
	{
		//SsmStopRing(ch);
		whu_UserChState = USER_WAIT_LOCALPICKUP;
		SetDlgItemText(IDC_ACCEPTFAXORRECORD,_T("ֹͣ¼��"));
		whu_BRecord = true ;
		GetLocalTime(&whu_SystemTime);
		//whu_RecordFile=_T("C:\\Whu_Data\\¼��\\001.wav");
		//CString m_file = _T("");
		//m_file.Format(_T("C:\\Whu_Data\\¼��\\%d��%d��%d��%dʱ%d��_����%s.wav\0"),whu_NumberBuf,whu_SystemTime.wYear,whu_SystemTime.wMonth,whu_SystemTime.wDay,whu_SystemTime.wHour,whu_SystemTime.wMinute);
		//whu_RecordFile.Format(_T("C:\\Whu_Data\\¼��\\%d��%d��%d��%dʱ%d��_����%s.wav\0"),whu_NumberBuf,whu_SystemTime.wYear,whu_SystemTime.wMonth,whu_SystemTime.wDay,whu_SystemTime.wHour,whu_SystemTime.wMinute);
		whu_RecordFile.Format(_T("C:\\Whu_Data\\¼��\\%d��%d��%d��%dʱ%d��_����%s.wav\0"),whu_SystemTime.wYear,whu_SystemTime.wMonth,whu_SystemTime.wDay,whu_SystemTime.wHour,whu_SystemTime.wMinute,whu_NumberBuf);
		SsmSetRecMixer(AtrkCh,TRUE,0);
		SsmRecToFile(ch,whu_RecordFile,-1,1,0xffffffff,0xffffffff,1); //��������¼�������ɹ�����//
		//whu_HistoryDlg.m_HistoryList.SetItemText(whu_HistoryDlg.pos-1, 7, whu_RecordFile);//��������
		m_MyLogList.SetItemText(m_MyLogList.GetItemCount()-1,7,whu_RecordFile);
		whu_SaveLog();

	}
	else if (m_str == _T("ֹͣ¼��"))
	{
		SsmStopRecToFile(ch);
		SetDlgItemText(IDC_ACCEPTFAXORRECORD,_T("¼��"));
		
	}
	else if (m_str == _T("����"))
	{
		whu_AcceptFax = true; //���Խ���
		SsmPickup(AtrkCh);
		SsmPlayFile(AtrkCh,_T("C:\\whu_Data\\�뷢�ʹ���.wav"),-1,0,0xffffffff);
		GetDlgItem(IDC_ACCEPTFAXORRECORD)->ShowWindow(false);
		GetDlgItem(IDC_SHOWINFO)->ShowWindow(true);
		GetDlgItem(IDC_SHOWINFO)->SetWindowTextA(_T("׼���������ȴ��Է����ʹ���..."));
		//GetDlgItem(IDC_STATICTASK)->ShowWindow(true);
		//GetDlgItem(IDC_PROGRESS02)->ShowWindow(true); //���մ����޷���ʾ����������ȡ����byte��δ֪��

	}
	
}
void Cwhu_VoiceCardDlg::whu_ShowNumInfo(CString m_Num)
{

	int m_Total =m_MyContactList.GetItemCount();
	whu_ContractData m_ContractData = whu_GetContractData(m_Num);
	SetDlgItemText(IDC_NAME2,m_ContractData.Name);
	SetDlgItemText(IDC_DANWEI2,m_ContractData.Department);
	SetDlgItemText(IDC_BUMEN2,m_ContractData.Unit);
	SetDlgItemText(IDC_ZHIWU2,m_ContractData.Job);
	SetDlgItemText(IDC_HAOMA2,m_Num);
	whu_CallImg= cvLoadImage(m_ContractData.PictureFile);
	DrawPicToHDC(whu_CallImg,IDC_CALLPIC);
	
}
void Cwhu_VoiceCardDlg::DrawPicToHDC(IplImage *img, UINT ID )
{
	CDC *pDC = GetDlgItem(ID)->GetDC();
	HDC hDC= pDC->GetSafeHdc(); 
	CRect rect; 
	GetDlgItem(ID)->GetClientRect(&rect); 
	CvvImage cimg; 
	cimg.CopyOf(img); 
	cimg.DrawToHDC(hDC,&rect); 
	cimg.Destroy();
	ReleaseDC(pDC); 

}

bool Cwhu_VoiceCardDlg::whu_CheckLogIn(CString name,CString password)
{

	for (int i=0;i<whu_User[0].GetSize();i++)
	{
		if (name == whu_User[0].GetAt(i)&&password == whu_User[1].GetAt(i)&&whu_User[2].GetAt(i)==_T("��ע��"))
		{
			return true;
		}
	}

	return false;
}
whu_ContractData Cwhu_VoiceCardDlg::whu_GetContractData(CString m_num)
{
	//int m_Total = whu_ContractDlg.pos;
	int m_Total = m_MyContactList.GetItemCount();
	whu_ContractData m_ContractData;

	if (m_num!=_T(""))
	{
		for(int i=0;i<m_Total;i++)
		{
			m_ContractData.Department = m_MyContactList.GetItemText(i,3);
			m_ContractData.Unit = m_MyContactList.GetItemText(i,2);
			m_ContractData.Name = m_MyContactList.GetItemText(i,1);
			m_ContractData.Job = m_MyContactList.GetItemText(i,4);
			m_ContractData.TelNum = m_MyContactList.GetItemText(i,5);
			string str = m_ContractData.TelNum.GetBuffer(20);
			str=trim((string)str);
			m_ContractData.TelNum = str.c_str();

			m_ContractData.E_Mail = m_MyContactList.GetItemText(i,8);
			m_ContractData.OfficeNum = m_MyContactList.GetItemText(i,6);
			str = m_ContractData.OfficeNum.GetBuffer(20);
			str=trim((string)str);
			m_ContractData.OfficeNum = str.c_str();

			m_ContractData.FaxNum = m_MyContactList.GetItemText(i,7);
			str = m_ContractData.FaxNum.GetBuffer(20);
			str=trim((string)str);
			m_ContractData.FaxNum = str.c_str();

			//�õ���ϵ�˵�ͷ��//
			CString m_folder = _T("C:\\Whu_Data\\��ϵ��ͷ��");
			CString m_Pic = m_folder + _T("\\") + m_ContractData.Name + _T(".jpg");
			if(whu_findfile(m_Pic,m_folder) != true)
			{
				m_ContractData.PictureFile = _T("C:\\Whu_Data\\��ϵ��ͷ��\\İ����.jpg");
			}else{
				m_ContractData.PictureFile =  m_Pic;
			}

			GetDlgItemText(IDC_DUTYPERSON,m_ContractData.DutyPerson);
			if(m_num == m_ContractData.TelNum||m_num == m_ContractData.OfficeNum||m_num == m_ContractData.FaxNum)
			{
				return m_ContractData;
			}
		}
		whu_ContractData m_data ;
		m_data.Department = _T("δ֪");
		m_data.Unit =  _T("δ֪");
		m_data.Name =  _T("İ����");
		m_data.Job =  _T("δ֪");
		m_data.TelNum =  _T("δ֪");
		m_data.E_Mail =  _T("δ֪");
		m_data.OfficeNum =  _T("δ֪");
		m_data.FaxNum =  _T("δ֪");
		GetDlgItemText(IDC_DUTYPERSON,m_data.DutyPerson);
		m_data.PictureFile = _T("C:\\Whu_Data\\��ϵ��ͷ��\\İ����.jpg");
		return m_data;
	}
	else{
		whu_ContractData m_data ;
		m_data.Department = _T("");
		m_data.Unit =  _T("");
		m_data.Name =  _T("");
		m_data.Job =  _T("");
		m_data.TelNum =  _T("");
		m_data.E_Mail =  _T("");
		m_data.OfficeNum =  _T("");
		m_data.FaxNum =  _T("");
		m_data.PictureFile = _T("C:\\Whu_Data\\��ϵ��ͷ��\\İ����.jpg");
		return m_data;
	}
	
}

void Cwhu_VoiceCardDlg::resize()
{
	float fsp[2];
	POINT Newp; //��ȡ���ڶԻ���Ĵ�С
	CRect recta;  
	GetClientRect(&recta); //ȡ�ͻ�����С  
	Newp.x=recta.right-recta.left;
	Newp.y=recta.bottom-recta.top;
	fsp[0]=(float)Newp.x/Old.x;
	fsp[1]=(float)Newp.y/Old.y;
	CRect Rect;
	int woc;
	CPoint OldTLPoint,TLPoint; //���Ͻ�
	CPoint OldBRPoint,BRPoint; //���½�
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
	m_editFont.CreatePointFont(240,_T("����"));
	//GetDlgItem(IDC_DUTYLEADER)->SetFont(&m_editFont);
	GetDlgItem(IDC_DUTYPERSON)->SetFont(&m_editFont);
	GetDlgItem(IDC_SHOWINFO)->SetFont(&m_editFont);
	GetDlgItem(IDC_NAME2)->SetFont(&m_editFont);
	GetDlgItem(IDC_DANWEI2)->SetFont(&m_editFont);
	GetDlgItem(IDC_BUMEN2)->SetFont(&m_editFont);
	GetDlgItem(IDC_ZHIWU2)->SetFont(&m_editFont);
	GetDlgItem(IDC_HAOMA2)->SetFont(&m_editFont);
	GetDlgItem(IDC_NAME1)->SetFont(&m_editFont);
	GetDlgItem(IDC_DANWEI1)->SetFont(&m_editFont);
	GetDlgItem(IDC_BUMEN1)->SetFont(&m_editFont);
	GetDlgItem(IDC_ZHIWU1)->SetFont(&m_editFont);
	GetDlgItem(IDC_HAOMA1)->SetFont(&m_editFont);
	GetDlgItem(IDC_BRECORDMYNOTICE)->SetFont(&m_editFont);
	GetDlgItem(IDC_COMBOLEADER)->SetFont(&m_editFont);
	GetDlgItem(IDC_STATICTASK)->SetFont(&m_editFont);
	
	//::SetTextColor(GetDlgItem(IDC_SHOWINFO)->,RGB(0,0,255)); 

}


void Cwhu_VoiceCardDlg::OnBnClickedBrecordmynotice()
{
	// TODO: Add your control notification handler code here
	/*CString m_Str =_T("");
	GetDlgItemText(IDC_BRECORDMYNOTICE,m_Str);
	if (m_Str == _T("¼��"))
	{
		whu_RecordNotice = true;
		whu_ShowRecordInfo(true);
		SetDlgItemText(IDC_BRECORDMYNOTICE,_T("ȡ��"));
	}
	else if(m_Str == _T("ȡ��")){
		whu_RecordNotice = false;
		whu_ShowRecordInfo(false);
		SetDlgItemText(IDC_BRECORDMYNOTICE,_T("¼��"));
		m_MyLogList.DeleteItem(m_MyLogList.GetItemCount()-1);
		whu_NoticeFile = _T("");
	}*/
	whu_NoticeState = NOTICE_NONE;
	m_NoticeDlg = new Cwhu_NoticeDlg();
	m_NoticeDlg->whu_ParentDlg = this;
	m_NoticeDlg->DoModal();
	//m_NoticeDlg->Create(IDD_NOTICE,this);
	//m_NoticeDlg->ShowWindow(SW_SHOW);

}


//void Cwhu_VoiceCardDlg::OnBnClickedTrylisten()
//{
//	// TODO: Add your control notification handler code here
//	//CString m_Str;
//	//GetDlgItemText(IDC_TRYLISTEN,m_Str);
//	//if (m_Str ==_T("����")&&whu_NoticeFile!=_T(""))
//	//{
//	//	sndPlaySound(whu_NoticeFile,SND_ASYNC);
//	//	SetDlgItemText(IDC_TRYLISTEN,_T("ֹͣ"));
//	//} 
//	//else
//	//{
//	//	sndPlaySound(NULL,SND_ASYNC);
//	//	SetDlgItemText(IDC_TRYLISTEN,_T("����"));
//	//}
//	Cwhu_RecvCallDlg *m_DlgRecvCall = new Cwhu_RecvCallDlg();
//	m_DlgRecvCall->whu_ParentDlg = this;
//	m_DlgRecvCall->Create(IDD_DIALOGRECVCALL,this);
//	m_DlgRecvCall->ShowWindow(SW_SHOW);
//
//}
void Cwhu_VoiceCardDlg::whu_LinkToSQL()
{

	//��ϵ���б��ʼ��
	LONG lStyle; 
	lStyle = GetWindowLong(m_MyContactList.m_hWnd, GWL_STYLE);
	lStyle &= ~LVS_TYPEMASK; 
	lStyle |= LVS_REPORT; 
	SetWindowLong(m_MyContactList.m_hWnd, GWL_STYLE, lStyle);
	DWORD dwStyle = m_MyContactList.GetExtendedStyle(); 
	dwStyle |= LVS_EX_FULLROWSELECT;
	dwStyle |= LVS_EX_GRIDLINES;
	dwStyle |= LVS_EX_CHECKBOXES;
	m_MyContactList.SetExtendedStyle(dwStyle); 
	m_MyContactList.InsertColumn( 0, _T("���"), LVCFMT_LEFT, 40 );
	m_MyContactList.InsertColumn( 1, _T("����"), LVCFMT_LEFT, 120 ); 
	m_MyContactList.InsertColumn( 2, _T("��λ"), LVCFMT_LEFT, 120 );
	m_MyContactList.InsertColumn( 3, _T("����"), LVCFMT_LEFT, 160 );
	m_MyContactList.InsertColumn( 4, _T("ְ��"), LVCFMT_LEFT, 80 );
	m_MyContactList.InsertColumn( 5, _T("�ֻ�"), LVCFMT_LEFT, 160 ); 
	m_MyContactList.InsertColumn( 6, _T("�칫�ҵ绰"), LVCFMT_LEFT, 160 ); 
	m_MyContactList.InsertColumn( 7, _T("�����"), LVCFMT_LEFT, 120 ); 
	m_MyContactList.InsertColumn( 8, _T("����"), LVCFMT_LEFT, 150 );
	//���ӵ�MS SQL Server  
	_ConnectionPtr pMyConnect("ADODB.Connection");
	pMyConnect.CreateInstance("ADODB.Connection"); 
	//�ҵĵ���
	_bstr_t strConnect = "Provider=SQLOLEDB.1;Server=CHINESE-70D3153\\WHSW;Database=whu;uid=sa;pwd=123456;"; 
	//���ػ�
	//_bstr_t strConnect = "Provider=SQLOLEDB.1;Server=MICROSOF-0FCDA5\\WHSW;Database=whu;uid=whu;pwd=123456;"; 
	try  
	{  
		pMyConnect->Open(strConnect, "", "", adModeUnknown);  
	}  
	catch(_com_error &e)  
	{  
		MessageBox(e.Description(), _T("����"),MB_OK|MB_ICONINFORMATION);  
	}
	CString strSql="select * from contact";//����ִ�е�SQL���  
	_RecordsetPtr pRecordset;  
	HRESULT hTRes;
	hTRes = pRecordset.CreateInstance(_T("ADODB.Recordset"));
	if (SUCCEEDED(hTRes))
	{
		hTRes = pRecordset->Open((LPTSTR)strSql.GetBuffer(130),
			pMyConnect.GetInterfacePtr(),
			adOpenDynamic,adLockPessimistic,adCmdText);
		int i=0;
		char buf[10] = {'\0'};
		while(!(pRecordset->adoEOF))
		{
			//����ѯ�����ʾ���б��ؼ���//
			m_MyContactList.InsertItem(i,GetStringFromVariant(pRecordset->GetCollect("ID"))); 
			//m_MyContactList.InsertItem(i,_itoa(i+1,buf,10));
			m_MyContactList.SetItemText(i,1,GetStringFromVariant(pRecordset->GetCollect("����"))); 
			m_MyContactList.SetItemText(i,2,GetStringFromVariant(pRecordset->GetCollect("��λ"))); 
			m_MyContactList.SetItemText(i,3,GetStringFromVariant(pRecordset->GetCollect("����"))); 
			m_MyContactList.SetItemText(i,4,GetStringFromVariant(pRecordset->GetCollect("ְ��")));
			m_MyContactList.SetItemText(i,5,GetStringFromVariant(pRecordset->GetCollect("�ֻ�")));
			m_MyContactList.SetItemText(i,6,GetStringFromVariant(pRecordset->GetCollect("�칫�ҵ绰")));
			m_MyContactList.SetItemText(i,7,GetStringFromVariant(pRecordset->GetCollect("�����")));
			m_MyContactList.SetItemText(i,8,GetStringFromVariant(pRecordset->GetCollect("����")));
			m_MyContactList.SetItemText(i,9,GetStringFromVariant(pRecordset->GetCollect("ͷ��")));
			m_ContactListData[10].InsertAt(i,GetStringFromVariant(pRecordset->GetCollect("ƴ��")));
			i++;
			if(!(pRecordset->adoEOF))
				pRecordset->MoveNext();
		}
	}
	//������ϵ���б�����ո�//
	static bool save = false;
	if (save == false)
	{
		for (int i=0;i<m_MyContactList.GetItemCount();i++)
		{
			for (int j=0;j<10;j++)
			{
				CString m_str = m_MyContactList.GetItemText(i,j);
				string str = m_str.GetBuffer();
				str=trim((string)str);
				m_str = str.c_str();
				m_ContactListData[j].InsertAt(i,m_str);
			}	
		}
		for (int i=0;i<m_ContactListData[0].GetSize();i++)
		{
			for (int j=0;j<9;j++)
			{
				CString ss = m_ContactListData[j].GetAt(i);
				m_MyContactList.SetItemText(i,j,ss);
			}	
		}
		save = true;
	}
	
}
CString Cwhu_VoiceCardDlg::GetStringFromVariant(_variant_t var)
{
	return var.vt==VT_NULL?"":(char*)(_bstr_t)var;
}


void Cwhu_VoiceCardDlg::OnNMRClickMycontactlist(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMITEMACTIVATE pNMItemActivate = reinterpret_cast<LPNMITEMACTIVATE>(pNMHDR);
	// TODO: Add your control notification handler code here
	NM_LISTVIEW* pNMListView = (NM_LISTVIEW*)pNMHDR; 
	int m_ChoosedPos = pNMListView->iItem;
	int a=0;
	bool m_Check = false;
	for(int i=0; i<m_MyContactList.GetItemCount(); i++) 
	{
		m_Check = m_MyContactList.GetCheck(i);//
		if(m_Check&&pNMListView->iItem != -1)  //��ѡ�и�ѡ��Ⱥ��//m_int == LVIS_SELECTED||
		{ 
			DWORD dwPos = GetMessagePos(); 
			CPoint point( LOWORD(dwPos), HIWORD(dwPos) ); 
			CMenu menu; 
			VERIFY( menu.LoadMenu(IDR_MENUGROUP) ); 
			CMenu* popup = menu.GetSubMenu(0); 
			ASSERT( popup != NULL ); 
			popup->TrackPopupMenu(TPM_LEFTALIGN | TPM_RIGHTBUTTON,point.x, point.y, this ); 
			break;
		}
	}
	if(pNMListView->iItem != -1&&m_Check==false) 
	{ 
		CString strtemp; 
		strtemp.Format(_T("�������ǵ�%d�е�%d��"), pNMListView->iItem, pNMListView->iSubItem); 
		LVCOLUMN lvcol; 
		char    str[256]; 
		lvcol.mask = LVCF_TEXT; 
		lvcol.pszText = str; 
		lvcol.cchTextMax = 256; 
		m_MyContactList.GetColumn(pNMListView->iSubItem,&lvcol);
		//AfxMessageBox(lvcol.pszText); 
		//lvcol.pszTextΪ����//
		CString m_head = lvcol.pszText;
		/*if(m_head == _T("����"))
		{
			char buf[60];
			m_MyContactList.GetItemText(m_ChoosedPos,1,buf,50);
			m_MyContactData.Name.Format("%s",buf);
			if (m_MyContactData.Name != _T(""))
			{
				DWORD dwPos = GetMessagePos(); 
				CPoint point( LOWORD(dwPos), HIWORD(dwPos) ); 
				CMenu menu; 
				VERIFY( menu.LoadMenu(IDR_MENUSEEHISTORY) ); 
				CMenu* popup = menu.GetSubMenu(0); 
				ASSERT( popup != NULL ); 
				popup->TrackPopupMenu(TPM_LEFTALIGN | TPM_RIGHTBUTTON,point.x, point.y, this ); 
			}
		}
		else if(m_head == _T("��λ"))
		{
			char buf[60];
			m_MyContactList.GetItemText(m_ChoosedPos,2,buf,50);
			m_MyContactData.Unit.Format("%s",buf);
			if (m_MyContactData.Unit != _T(""))
			{
				DWORD dwPos = GetMessagePos(); 
				CPoint point( LOWORD(dwPos), HIWORD(dwPos) ); 
				CMenu menu; 
				VERIFY( menu.LoadMenu(IDR_MENUSEEHISTORY) ); 
				CMenu* popup = menu.GetSubMenu(0); 
				ASSERT( popup != NULL ); 
				popup->TrackPopupMenu(TPM_LEFTALIGN | TPM_RIGHTBUTTON,point.x, point.y, this ); 
			}
		}
		else if(m_head == _T("����"))
		{
			char buf[60];
			m_MyContactList.GetItemText(m_ChoosedPos,3,buf,50);
			m_MyContactData.Department.Format("%s",buf);
			if (m_MyContactData.Department != _T(""))
			{
				DWORD dwPos = GetMessagePos(); 
				CPoint point( LOWORD(dwPos), HIWORD(dwPos) ); 
				CMenu menu; 
				VERIFY( menu.LoadMenu(IDR_MENUSEEHISTORY) ); 
				CMenu* popup = menu.GetSubMenu(0); 
				ASSERT( popup != NULL ); 
				popup->TrackPopupMenu(TPM_LEFTALIGN | TPM_RIGHTBUTTON,point.x, point.y, this ); 
			}
		}
		else if(m_head == _T("ְ��"))
		{
			char buf[60];
			m_MyContactList.GetItemText(m_ChoosedPos,4,buf,50);
			m_MyContactData.Job.Format("%s",buf);
			if (m_MyContactData.Job != _T(""))
			{
				DWORD dwPos = GetMessagePos(); 
				CPoint point( LOWORD(dwPos), HIWORD(dwPos) ); 
				CMenu menu; 
				VERIFY( menu.LoadMenu(IDR_MENUSEEHISTORY)); 
				CMenu* popup = menu.GetSubMenu(0); 
				ASSERT( popup != NULL ); 
				popup->TrackPopupMenu(TPM_LEFTALIGN | TPM_RIGHTBUTTON,point.x, point.y, this ); 
			}
		}*/
		if (m_head == _T("�ֻ�"))
		{


			m_MyContactData.ID = m_MyContactList.GetItemText(m_ChoosedPos,0);
			m_MyContactData.Name = m_MyContactList.GetItemText(m_ChoosedPos,1);
			m_MyContactData.Unit = m_MyContactList.GetItemText(m_ChoosedPos,2);
			m_MyContactData.Department = m_MyContactList.GetItemText(m_ChoosedPos,3);
			m_MyContactData.Job = m_MyContactList.GetItemText(m_ChoosedPos,4);
			m_MyContactData.TelNum = m_MyContactList.GetItemText(m_ChoosedPos,5);
			m_MyContactData.OfficeNum = m_MyContactList.GetItemText(m_ChoosedPos,6);
			m_MyContactData.FaxNum = m_MyContactList.GetItemText(m_ChoosedPos,7);

			if (m_MyContactData.TelNum != _T(""))
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
		else if (m_head == _T("�칫�绰"))
		{
			m_MyContactData.ID = m_MyContactList.GetItemText(m_ChoosedPos,0);
			m_MyContactData.Name = m_MyContactList.GetItemText(m_ChoosedPos,1);
			m_MyContactData.Unit = m_MyContactList.GetItemText(m_ChoosedPos,2);
			m_MyContactData.Department = m_MyContactList.GetItemText(m_ChoosedPos,3);
			m_MyContactData.Job = m_MyContactList.GetItemText(m_ChoosedPos,4);
			m_MyContactData.TelNum = m_MyContactList.GetItemText(m_ChoosedPos,5);
			m_MyContactData.OfficeNum = m_MyContactList.GetItemText(m_ChoosedPos,6);
			m_MyContactData.FaxNum = m_MyContactList.GetItemText(m_ChoosedPos,7);
			if (m_MyContactData.OfficeNum !=_T(""))
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
		else if (m_head == _T("�����"))
		{
			char buf[20];
			m_MyContactList.GetItemText(m_ChoosedPos,7,buf,12);
			m_MyContactData.FaxNum.Format("%s",buf);
			if (m_MyContactData.FaxNum != _T(""))
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
		else if(m_head == _T("����"))
		{
			char buf[60];
			m_MyContactList.GetItemText(m_ChoosedPos,8,buf,50);
			m_MyContactData.E_Mail.Format("%s",buf);
			if (m_MyContactData.E_Mail != _T(""))
			{
				DWORD dwPos = GetMessagePos(); 
				CPoint point( LOWORD(dwPos), HIWORD(dwPos) ); 
				CMenu menu; 
				VERIFY( menu.LoadMenu(IDR_MENUEMAIL) ); 
				CMenu* popup = menu.GetSubMenu(0); 
				ASSERT( popup != NULL ); 
				popup->TrackPopupMenu(TPM_LEFTALIGN | TPM_RIGHTBUTTON,point.x, point.y, this ); 
			}
		}
	}
	if(pNMListView->iItem == -1)
	{
		DWORD dwPos = GetMessagePos(); 
		CPoint point( LOWORD(dwPos), HIWORD(dwPos) ); 
		CMenu menu; 
		VERIFY( menu.LoadMenu(IDR_MENUREFRESH) ); 
		CMenu* popup = menu.GetSubMenu(0); 
		ASSERT( popup != NULL ); 
		popup->TrackPopupMenu(TPM_LEFTALIGN | TPM_RIGHTBUTTON,point.x, point.y, this ); 
	}

	*pResult = 0;
}


void Cwhu_VoiceCardDlg::OnTelPhoneDial()
{
	// TODO: Add your command handler code here
	if (m_MyContactData.TelNum !=_T("")){
		string str = m_MyContactData.TelNum.GetBuffer(20);
		str=trim((string)str);
		m_MyContactData.TelNum = str.c_str();
		whu_Dial(m_MyContactData.TelNum);
	} 
	else{
		AfxMessageBox(_T("δ�洢�ֻ�����"));
	}
}


void Cwhu_VoiceCardDlg::OnTelPhoneSendSingleNotice()
{
	// TODO: Add your command handler code here
	if (m_MyContactData.TelNum !=_T("")){
		string str = m_MyContactData.TelNum.GetBuffer(20);
		str=trim((string)str);
		m_MyContactData.TelNum = str.c_str();
		whu_SendNotice(&m_MyContactData.TelNum,1);
	} 
	else{
		AfxMessageBox(_T("δ�洢�ֻ�����"));
	}
}


void Cwhu_VoiceCardDlg::OnOfficePhoneDial()
{
	// TODO: Add your command handler code here
	if (m_MyContactData.OfficeNum!=_T(""))
	{
		string str = m_MyContactData.OfficeNum.GetBuffer(20);
		str=trim((string)str);
		m_MyContactData.OfficeNum = str.c_str();
		whu_Dial(m_MyContactData.OfficeNum);
	} 
	else 
	{
		AfxMessageBox(_T("δ�洢�칫�ҵ绰"));
	}
}


void Cwhu_VoiceCardDlg::OnOfficePhoneSendSingleNotice()
{
	// TODO: Add your command handler code here
	if (m_MyContactData.OfficeNum !=_T("")){
		string str = m_MyContactData.OfficeNum.GetBuffer(20);
		str=trim((string)str);
		m_MyContactData.OfficeNum = str.c_str();
		whu_SendNotice(&m_MyContactData.OfficeNum,1);
	} 
	else{
		AfxMessageBox(_T("δ�洢�칫�Һ���"));
	}
}
void Cwhu_VoiceCardDlg::OnSendSingleFax()
{
	// TODO: Add your command handler code here
	int a = 0;
	CString whu_DocFile = _T("");
	CFileDialog cf(1, NULL, NULL, OFN_HIDEREADONLY|OFN_OVERWRITEPROMPT, "Fax File(*.tif;*.doc;*.docx;*.xls;*.xlsx;*.ppt;*.pptx;*.pdf;*.txt;*.html;*.jpg;*bmp)|*.tif;*.doc;*.docx;*.xls;*.xlsx;*.ppt;*.pptx;*.pdf;*.txt;*.html;*.jpg;*bmp|Tif File(*.fax)|*.fax|Tiff File(*.tiff)|*.tiff|All File(*.*)|*.*|",NULL);//���Դ�doc��pdf���ļ�
	if(cf.DoModal() == IDOK)
	{
		whu_DocFile=cf.GetPathName();
		whu_SendFaxFileTitle = cf.GetFileTitle();
	}
	string str = m_MyContactData.FaxNum.GetBuffer(20);
	str=trim((string)str);
	m_MyContactData.FaxNum = str.c_str();
	whu_SendFax(&m_MyContactData.FaxNum,whu_DocFile,1);
}


void Cwhu_VoiceCardDlg::OnSendGroupNotice()
{
	// TODO: Add your command handler code here
	int a=0;
	for(int i=0; i<m_MyContactList.GetItemCount(); i++) 
	{
		int m_int = m_MyContactList.GetItemState(i, LVIS_SELECTED);
		bool m_b = m_MyContactList.GetCheck(i);
		if( m_int == LVIS_SELECTED|| m_b )  //��ѡ�и�ѡ��Ⱥ��//
		{ 
			char buf[20]=_T("");
			m_MyContactList.GetItemText(i,5,buf,12);
			whu_AllNos[a].Format("%s",buf);
			if (whu_AllNos[a] == _T(""))
			{
				m_MyContactList.GetItemText(i,6,buf,12); // ����Ҳ����ֻ��ţ��ͻ��칫�ҵ绰//
				whu_AllNos[a].Format("%s",buf);
				if (whu_AllNos[a] == _T(""))
				{
					CString m_str;
					m_str.Format(_T("%s%d%s"),_T("��"),i+1,_T("��δ�洢�ֻ��Ͱ칫�Һ���"));
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
	if (a>0)
	{
		whu_SendNotice(whu_AllNos,a);
	}
}


void Cwhu_VoiceCardDlg::OnSendGroupFax()
{
	// TODO: Add your command handler code here
	int a = 0;
	CString whu_DocFile = _T("");
	CFileDialog cf(1, NULL, NULL, OFN_HIDEREADONLY|OFN_OVERWRITEPROMPT, "Fax File(*.tif;*.doc;*.docx;*.xls;*.xlsx;*.ppt;*.pptx;*.pdf;*.txt;*.html;*.jpg;*bmp)|*.tif;*.doc;*.docx;*.xls;*.xlsx;*.ppt;*.pptx;*.pdf;*.txt;*.html;*.jpg;*bmp|Tif File(*.fax)|*.fax|Tiff File(*.tiff)|*.tiff|All File(*.*)|*.*|",NULL);//���Դ�doc��pdf���ļ�
	if(cf.DoModal() == IDOK)
	{
		whu_DocFile=cf.GetPathName();
		whu_SendFaxFileTitle = cf.GetFileTitle();
	}
	for(int i=0; i<m_MyContactList.GetItemCount(); i++) 
	{ 
		if( m_MyContactList.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED|| m_MyContactList.GetCheck(i) ) 
		{ 
			char buf[20]=_T("");
			m_MyContactList.GetItemText(i,7,buf,11);
			whu_AllNos[a].Format("%s",buf);
			if (whu_AllNos[a] == _T(""))
			{
				CString m_str;
				m_str.Format(_T("%s%d%s"),_T("��"),i+1,_T("��δ�洢�������"));
				AfxMessageBox(m_str);
			}else
			{
				a++;
			}

		} 
	}
	if (a>0)
	{
		whu_SendFax(whu_AllNos,whu_DocFile,a);
	}

}




void Cwhu_VoiceCardDlg::OnBnClickedOpenscanner() //��ɨ����//
{
	// TODO: Add your control notification handler code here
	//ShellExecute(NULL, "open", _T("C:\\Program Files\\Canon Electronics\\DRM140\\TouchDR.exe"),NULL,NULL, SW_SHOWNORMAL );
	ShellExecute(NULL, "open", _T("F:\\Canon Electronics\\DRM140\\TouchDR.exe"),NULL,NULL, SW_SHOWNORMAL );
}

void Cwhu_VoiceCardDlg::whu_InitLogList()
{
	LONG lStyle; 
	lStyle = GetWindowLong(m_MyLogList.m_hWnd, GWL_STYLE);//��ȡ��ǰ����style 
	lStyle &= ~LVS_TYPEMASK; //�����ʾ��ʽλ 
	lStyle |= LVS_REPORT; //����style 
	SetWindowLong(m_MyLogList.m_hWnd, GWL_STYLE, lStyle);//����style 
	DWORD dwStyle = m_MyLogList.GetExtendedStyle(); 
	dwStyle |= LVS_EX_FULLROWSELECT;//ѡ��ĳ��ʹ���и�����ֻ������report����listctrl�� //
	dwStyle |= LVS_EX_GRIDLINES;//�����ߣ�ֻ������report����listctrl�� //
	//dwStyle |= LVS_EX_CHECKBOXES;//itemǰ����checkbox�ؼ� //
	m_MyLogList.SetExtendedStyle(dwStyle); //������չ��� //

	m_MyLogList.InsertColumn( 0, "ID", LVCFMT_LEFT, 40 );
	m_MyLogList.InsertColumn( 1, _T("ʱ��"), LVCFMT_LEFT, 90 ); 
	m_MyLogList.InsertColumn( 2, _T("ֵ����"), LVCFMT_LEFT, 50 );
	m_MyLogList.InsertColumn( 3, _T("�¼�"), LVCFMT_LEFT, 50 );
	m_MyLogList.InsertColumn( 4, _T("��λ"), LVCFMT_LEFT, 100 );
	m_MyLogList.InsertColumn( 5, _T("����"), LVCFMT_LEFT, 60 ); 
	m_MyLogList.InsertColumn( 6, _T("����"), LVCFMT_LEFT, 100 ); 
	m_MyLogList.InsertColumn( 7, _T("�ļ�"), LVCFMT_LEFT, 200 ); 
	m_MyLogList.InsertColumn( 8, _T("״̬"), LVCFMT_LEFT, 160 ); 
	m_MyLogList.InsertColumn( 9, _T("��ע"), LVCFMT_LEFT, 160 );
	whu_logchangetime();

}
void Cwhu_VoiceCardDlg::whu_AddLog(CString time,CString duty,CString taskmode,CString unit,CString name,CString number,CString file,CString state,CString beizhu)
{
	//�����жϣ��ж��Ƿ��ǽ������־//
	GetLocalTime(&whu_SystemTime);
	CTime m_time1(whu_SystemTime.wYear,whu_SystemTime.wMonth,whu_SystemTime.wDay,0,0,0); //��ǰ����
	CTime m_time2;
	m_DataCtrl.GetTime(m_time2);  //�û��鿴��־��ʱ��ѡ�������,��һ�����޸�//
	if (m_time1.GetDay()!= m_time2.GetDay()|m_time1.GetMonth()!=m_time2.GetMonth()|m_time1.GetYear()!=m_time2.GetYear())
	{
		m_DataCtrl.SetTime(&m_time1);
		whu_logchangetime();
	}
	//д������//
	int pos = m_MyLogList.GetItemCount();
	char posbuf [10] = _T("");
	itoa(pos+1,posbuf,10);
	m_MyLogList.InsertItem(pos, posbuf);//������// 
	m_MyLogList.SetItemText(pos, 1, time);//�������� //
	m_MyLogList.SetItemText(pos, 2, duty);//�������� //
	m_MyLogList.SetItemText(pos, 3, taskmode);//�������� //
	m_MyLogList.SetItemText(pos, 4,unit);//�������� //
	m_MyLogList.SetItemText(pos, 5,name);//�������� //
	m_MyLogList.SetItemText(pos, 6, number);//�������� //
	m_MyLogList.SetItemText(pos, 7, file);//�������� //
	m_MyLogList.SetItemText(pos, 8, state);//��������//
	m_MyLogList.SetItemText(pos, 9, beizhu);//��������//
	whu_SaveLog();
}
void Cwhu_VoiceCardDlg::whu_ModifyCheck(CString m_str)
{
	////�����жϣ��ж��Ƿ��ǽ������־//
	//GetLocalTime(&whu_SystemTime);
	//CTime m_time1(whu_SystemTime.wYear,whu_SystemTime.wMonth,whu_SystemTime.wDay,0,0,0); //��ǰ����//
	//CTime m_time2;
	//m_DataCtrl.GetTime(m_time2);  //�û��鿴��־��ʱ��ѡ�������//
	//if (m_time1.GetDay()!= m_time2.GetDay()|m_time1.GetMonth()!=m_time2.GetMonth()|m_time1.GetYear()!=m_time2.GetYear())
	//{
	//	m_DataCtrl.SetTime(&m_time1);
	//	whu_logchangetime();
	//}
	//д������//
	int pos = m_MyLogList.GetItemCount();
	m_MyLogList.SetItemText(pos-1, 8, m_str);//��������//
	whu_SaveLog();
}
void Cwhu_VoiceCardDlg::OnSeeLog()
{
	// TODO: Add your command handler code here
	CString m_str[10];
	for (int i=0;i<10;i++)
	{
		m_str[i] = m_MyLogList.GetItemText(m_ClickLogPos,i);
	}
}


void Cwhu_VoiceCardDlg::OnNMRClickMyloglist(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMITEMACTIVATE pNMItemActivate = reinterpret_cast<LPNMITEMACTIVATE>(pNMHDR);
	// TODO: Add your control notification handler code here
	//NM_LISTVIEW* pNMListView = (NM_LISTVIEW*)pNMHDR; 
	//m_ClickLogPos = pNMListView->iItem;
	//if(pNMListView->iItem != -1) 
	//{ 
	//	CString strtemp; 
	//	strtemp.Format(_T("�������ǵ�%d�е�%d��"), pNMListView->iItem, pNMListView->iSubItem); 
	//	LVCOLUMN lvcol; 
	//	char    str[256]; 
	//	lvcol.mask = LVCF_TEXT; 
	//	lvcol.pszText = str; 
	//	lvcol.cchTextMax = 256; 
	//	m_MyLogList.GetColumn(pNMListView->iSubItem,&lvcol);
	//	//AfxMessageBox(lvcol.pszText); 
	//	//lvcol.pszTextΪ����//
	//	CString m_head = lvcol.pszText;

	//	DWORD dwPos = GetMessagePos(); 
	//	CPoint point( LOWORD(dwPos), HIWORD(dwPos) ); 
	//	CMenu menu; 
	//	VERIFY( menu.LoadMenu(IDR_MENULOG) ); 
	//	CMenu* popup = menu.GetSubMenu(0); 
	//	ASSERT( popup != NULL ); 

	//	popup->TrackPopupMenu(TPM_LEFTALIGN | TPM_RIGHTBUTTON,point.x, point.y, this ); 

	//}
	//else{
	//	DWORD dwPos = GetMessagePos(); 
	//	CPoint point( LOWORD(dwPos), HIWORD(dwPos) ); 
	//	CMenu menu; 
	//	VERIFY( menu.LoadMenu(IDR_MENUREFRESH) ); 
	//	CMenu* popup = menu.GetSubMenu(0); 
	//	ASSERT( popup != NULL ); 
	//	popup->TrackPopupMenu(TPM_LEFTALIGN | TPM_RIGHTBUTTON,point.x, point.y, this ); 
	//}
	
	*pResult = 0;


}


void Cwhu_VoiceCardDlg::OnBnClickedCheck001()
{
	// TODO: Add your control notification handler code here
	UpdateData(true);
	static int width = m_MyContactList.GetColumnWidth(1);
	if (m_checkname001)
	{
		m_MyContactList.SetColumnWidth(1,width);
		//AfxMessageBox("11");
	} 
	else
	{
		m_MyContactList.SetColumnWidth(1,0);
	}
}


void Cwhu_VoiceCardDlg::OnBnClickedCheck002()
{
	// TODO: Add your control notification handler code here
	UpdateData(true);
	static int width = m_MyContactList.GetColumnWidth(2);
	if (m_checkunite002)
	{
		m_MyContactList.SetColumnWidth(2,width);
	} 
	else
	{
		m_MyContactList.SetColumnWidth(2,0);
	}
}


void Cwhu_VoiceCardDlg::OnBnClickedCheck003()
{
	// TODO: Add your control notification handler code here
	UpdateData(true);
	static int width = m_MyContactList.GetColumnWidth(3);
	if (m_checkdepartment003)
	{
		m_MyContactList.SetColumnWidth(3,width);
		//AfxMessageBox("11");
	} 
	else
	{
		m_MyContactList.SetColumnWidth(3,0);
	}
}


void Cwhu_VoiceCardDlg::OnBnClickedCheck004()
{
	// TODO: Add your control notification handler code here
	UpdateData(true);
	static int width = m_MyContactList.GetColumnWidth(4);
	if (m_checkjob004)
	{
		m_MyContactList.SetColumnWidth(4,width);
		//AfxMessageBox("11");
	} 
	else
	{
		m_MyContactList.SetColumnWidth(4,0);
	}
}


void Cwhu_VoiceCardDlg::OnBnClickedCheck005()
{
	// TODO: Add your control notification handler code here
	UpdateData(true);
	static int width = m_MyContactList.GetColumnWidth(5);
	if (m_checktel005)
	{
		m_MyContactList.SetColumnWidth(5,width);
		//AfxMessageBox("11");
	} 
	else
	{
		m_MyContactList.SetColumnWidth(5,0);
	}
}


void Cwhu_VoiceCardDlg::OnBnClickedCheck006()
{
	// TODO: Add your control notification handler code here
	UpdateData(true);
	static int width = m_MyContactList.GetColumnWidth(6);
	if (m_checkOffNum006)
	{
		m_MyContactList.SetColumnWidth(6,width);
		//AfxMessageBox("11");
	} 
	else
	{
		m_MyContactList.SetColumnWidth(6,0);
	}
}
void Cwhu_VoiceCardDlg::OnBnClickedCheck007()
{
	// TODO: Add your control notification handler code here
	UpdateData(true);
	static int width = m_MyContactList.GetColumnWidth(7);
	if (m_checkFax007)
	{
		m_MyContactList.SetColumnWidth(7,width);
		//AfxMessageBox("11");
	} 
	else
	{
		m_MyContactList.SetColumnWidth(7,0);
	}
}


void Cwhu_VoiceCardDlg::OnBnClickedCheck008()
{
	// TODO: Add your control notification handler code here
	UpdateData(true);
	static int width = m_MyContactList.GetColumnWidth(8);
	if (m_checkemail008)
	{
		m_MyContactList.SetColumnWidth(8,width);
		//AfxMessageBox("11");
	} 
	else
	{
		m_MyContactList.SetColumnWidth(8,0);
	}
}



//void Cwhu_VoiceCardDlg::OnAddLogRemark()
//{
//	// TODO: Add your command handler code here
//	m_LogRemark = m_MyLogList.GetItemText(m_ClickLogPos,9);
//	UpdateData(false);
//	//������־�б��λ��//
//	CRect m_rect1,m_rect2;
//	GetDlgItem(IDC_LOGREMARK)->GetWindowRect(&m_rect1);
//	GetDlgItem(IDC_MYLOGLIST)->GetWindowRect(&m_rect2);
//	m_rect2.bottom = m_rect1.top - 40;
//	GetDlgItem(IDC_MYLOGLIST)->MoveWindow(m_rect2);
//	GetDlgItem(IDC_LOGREMARK)->ShowWindow(true);
//	GetDlgItem(IDC_LOGSAVE)->ShowWindow(true);
//	GetDlgItem(IDC_LOGCANCEL)->ShowWindow(true);
//	
//}


//void Cwhu_VoiceCardDlg::OnBnClickedLogsave()
//{
//	// TODO: Add your control notification handler code here
//	//������־�б��λ��//
//	GetDlgItem(IDC_MYLOGLIST)->MoveWindow(m_LogRect);
//
//	GetDlgItem(IDC_LOGREMARK)->ShowWindow(false);
//	GetDlgItem(IDC_LOGSAVE)->ShowWindow(false);
//	GetDlgItem(IDC_LOGCANCEL)->ShowWindow(false);
//	UpdateData(true);
//	m_MyLogList.SetItemText(m_ClickLogPos,9,m_LogRemark);
//}


//void Cwhu_VoiceCardDlg::OnBnClickedLogcancel()
//{
//	// TODO: Add your control notification handler code here
//	m_LogRemark = m_MyLogList.GetItemText(m_ClickLogPos,9);
//	GetDlgItem(IDC_LOGREMARK)->SetWindowText(m_LogRemark);
//}


void Cwhu_VoiceCardDlg::OnOpenLogFile()
{
	// TODO: Add your command handler code here
	CString file = m_MyLogList.GetItemText(m_ClickLogPos,7);
	if (file!=_T(""))
	{
		string str = file.GetBuffer(20);
		str=trim((string)str);
		file = str.c_str();
		ShellExecute(NULL, "open",file,NULL,NULL, SW_SHOWNORMAL );
	}
}


void Cwhu_VoiceCardDlg::OnOpenFilePath()
{
	// TODO: Add your command handler code here
	CString file = m_MyLogList.GetItemText(m_ClickLogPos,7);
	if (file!=_T(""))
	{
		string str = file.GetBuffer(20);
		str=trim((string)str);
		file = str.c_str();
		int pos = file.ReverseFind('\\');
		file = file.Left(pos+1);
		ShellExecute(NULL, "open",file,NULL,NULL, SW_SHOWNORMAL );
	}
	
}


void Cwhu_VoiceCardDlg::OnDeleteLog()
{
	// TODO: Add your command handler code here
	if (m_LogName == _T("����Ա"))
	{
		int state = AfxMessageBox(_T("ɾ����¼��"),MB_OK|| MB_OKCANCEL);
		if (state == IDOK)
		{
			//�õ�Ҫɾ����ĳһ�е���������//
			for (int i=1;i<m_MyLogList.GetHeaderCtrl()->GetItemCount();i++)
			{
				m_LogDeleteStr = m_LogDeleteStr + m_MyLogList.GetItemText(m_ClickLogPos,i)+_T(" ");
			}
			m_MyLogList.DeleteItem(m_ClickLogPos);
			int sss = m_MyLogList.GetItemCount();
			whu_ResetListNum(&m_MyLogList);
			whu_SaveLog();
		} 
	}
	else{
		AfxMessageBox(_T("�Բ��𣬵�ǰ�û��޴�Ȩ��,����ϵ����Ա"));
	}
}
void Cwhu_VoiceCardDlg::whu_ResetListNum(CListCtrl *m_MyList)
{
	int num = m_MyList->GetItemCount();
	for (int i=1;i<=num;i++)
	{
		CString m_str;
		m_str.Format(_T("%d"),i);
		m_MyList->SetItemText(i-1,0,m_str);	
	}
}
void Cwhu_VoiceCardDlg::OnBnClickedLogin()
{
	// TODO: Add your control notification handler code here

	m_LogInComBox.GetLBText(m_LogInComBox.GetCurSel(),m_LogName);
	UpdateData(true);
	if (whu_CheckLogIn(m_LogName,whu_password)) 
	{
		
		m_ComboxShowLeader = whu_ShowLeader(m_LogName);
		if (m_LogName == _T("����Ա"))
		{
			GetDlgItem(IDC_USER)->SetWindowTextA(_T("�û�����"));
		}else
		{
			GetDlgItem(IDC_USER)->SetWindowTextA(_T("�޸�����"));
		}
		UpdateData(false);
		whu_LogIn = true;
		if(whu_InitCtiBoard()==false)
		{
			AfxMessageBox(_T("δ��⵽�忨������ȷ���忨�Ƿ��Ѱ�װ"));
		}
		/*CString time = _T("");
		GetLocalTime(&whu_SystemTime);
		time.Format(_T("%d-%d-%d-%d:%d"),whu_SystemTime.wYear,whu_SystemTime.wMonth,whu_SystemTime.wDay,whu_SystemTime.wHour,whu_SystemTime.wMinute);
		whu_AddLog(time,m_LogName,_T("��½"),_T(""),_T(""),_T(""),_T(""),_T(""),_T(""));*/
		SetDlgItemText(IDC_DUTYPERSON,m_LogName);
		GetDlgItem(IDC_LOGOUT)->ShowWindow(true);
		GetDlgItem(IDC_TIME)->ShowWindow(true);
		GetDlgItem(IDC_DUTYLEADER2)->ShowWindow(true);
		GetDlgItem(IDC_COMBOLEADER)->ShowWindow(true);
		GetDlgItem(IDC_DUTYPERSON2)->ShowWindow(true);
		GetDlgItem(IDC_DUTYPERSON)->ShowWindow(true);
		GetDlgItem(IDC_SHOWDUTY)->ShowWindow(true);
		//GetDlgItem(IDC_DATETIMEPICKER1)->ShowWindow(true);
		GetDlgItem(IDC_BUTTONFAX)->ShowWindow(true);
		GetDlgItem(IDC_USER)->ShowWindow(true);
		GetDlgItem(IDC_BUTTONUPDATE)->ShowWindow(true);
		
		//GetDlgItem(IDC_TAB)->ShowWindow(true);
		//GetDlgItem(IDC_SHOCKWAVEFLASH)->ShowWindow(true);

		GetDlgItem(IDC_BRECORDMYNOTICE)->ShowWindow(true);
		//GetDlgItem(IDC_TRYLISTEN)->ShowWindow(true);
		GetDlgItem(IDC_OPENSCANNER)->ShowWindow(true);
		GetDlgItem(IDC_CHECKRECVFAX)->ShowWindow(true);
		GetDlgItem(IDC_CHECKAUTOFORWARDFAX)->ShowWindow(true);
		GetDlgItem(IDC_MYCONTACTLIST)->ShowWindow(true);
		//GetDlgItem(IDC_MYLOGLIST)->ShowWindow(true);
		GetDlgItem(IDC_SHAIXUAN)->ShowWindow(true);
		GetDlgItem(IDC_CHECK001)->ShowWindow(true);
		GetDlgItem(IDC_CHECK002)->ShowWindow(true);
		GetDlgItem(IDC_CHECK003)->ShowWindow(true);
		GetDlgItem(IDC_CHECK004)->ShowWindow(true);
		GetDlgItem(IDC_CHECK005)->ShowWindow(true);
		GetDlgItem(IDC_CHECK006)->ShowWindow(true);
		GetDlgItem(IDC_CHECK007)->ShowWindow(true);
		GetDlgItem(IDC_CHECK008)->ShowWindow(true);
		GetDlgItem(IDC_SEARCH)->ShowWindow(true);
		GetDlgItem(IDC_DUTYINFO)->ShowWindow(true);
		GetDlgItem(IDC_BUTTONSHOWLOG)->ShowWindow(true);
		GetDlgItem(IDC_BUTTONCONTACT)->ShowWindow(true);
		GetDlgItem(IDC_USER)->ShowWindow(true);
		GetDlgItem(IDC_WELCOME2)->ShowWindow(false);
		GetDlgItem(IDC_LOGIN)->ShowWindow(false);
		GetDlgItem(IDC_QUITSYSTEM)->ShowWindow(false);
		GetDlgItem(IDC_LOGNAME2)->ShowWindow(false);
		GetDlgItem(IDC_LOGCOMBOBOX)->ShowWindow(false);
		GetDlgItem(IDC_LOGMIMA2)->ShowWindow(false);
		GetDlgItem(IDC_LOGNMIMA)->ShowWindow(false);
		GetDlgItem(IDC_REGISTER)->ShowWindow(false);
		//m_MyContactList.SetColColor(1,RGB(0,255,255));
		//m_MyContactList.SetColColor(3,RGB(0,255,255));
		//m_MyContactList.SetColColor(5,RGB(0,255,255));
		//m_MyContactList.SetColColor(7,RGB(0,255,255));
		//m_MyContactList.SetBkColor(RGB(80,176,144));
		COLORREF crBkColor = ::GetSysColor(COLOR_3DFACE);
		m_MyContactList.SetTextBkColor(crBkColor);

		if (ConHeadNum == 0)
		{
			if (AfxMessageBox(_T("ͨѶ¼Ϊ�գ��밴��ģ�嵼��ͨ��¼�ļ�"))==IDOK)
			{
				CString m_pathname;
				CFileDialog cf(1, NULL, NULL, OFN_HIDEREADONLY|OFN_OVERWRITEPROMPT, "Excel File(*.xlsx;*.xls)|*.xlsx;*.xls",NULL);//
				CString m_FilePath;
				//����ʵ�ֵ��벻����Ҫ���ͨѶ¼��ʱ�����ٴε����Ի���cf//ʹ����ȡ�ɵİ취����forѭ�����//
				//bool m_check = false;
				//do 
				//{
				//	if(cf.DoModal() == IDOK)
				//	{
				//		m_FilePath=cf.GetPathName();
				//		m_check = whu_ImportContactData(m_FilePath);
				//	}
				//} while( m_check == true);
				for (int i=0;i<10;i++)
				{
					if(cf.DoModal() == IDOK)
					{
						CString m_FilePath=cf.GetPathName();
						if(whu_ImportContactData(m_FilePath)==true){break;}
					}
				}
				
			}
			
			
		}

	} 
	else
	{
		AfxMessageBox(_T("�û������������"));
	}

}
void Cwhu_VoiceCardDlg::OnBnClickedLogout()
{
	// TODO: Add your control notification handler code here


	m_BShowDutyList = false;
	SetDlgItemText(IDC_LOGCOMBOBOX,m_LogName);
	int check = SsmCloseCti();
	//CString m_LogName;
	//m_LogInComBox.GetLBText(m_LogInComBox.GetCurSel(),m_LogName);
	m_LogInComBox.ResetContent();
	for (int i=1;i<whu_User[0].GetSize();i++)
	{
		CString m_str = whu_User[0].GetAt(i);
		CString m_state = whu_User[2].GetAt(i);
		if (m_state == _T("��ע��"))
		{
			m_LogInComBox.AddString(m_str);
		}
	}
	GetDlgItem(IDC_LOGOUT)->ShowWindow(false);
	whu_LogIn = false;
	//GetDlgItem(IDC_DOWNPIC)->ShowWindow(false);
	GetDlgItem(IDC_TIME)->ShowWindow(false);
	GetDlgItem(IDC_DUTYLEADER2)->ShowWindow(false);
	GetDlgItem(IDC_COMBOLEADER)->ShowWindow(false);
	GetDlgItem(IDC_DUTYPERSON2)->ShowWindow(false);
	GetDlgItem(IDC_DUTYPERSON)->ShowWindow(false);
	GetDlgItem(IDC_SHOWDUTY)->ShowWindow(false);
	//GetDlgItem(IDC_TAB)->ShowWindow(true);
	//GetDlgItem(IDC_SHOCKWAVEFLASH)->ShowWindow(false);
	GetDlgItem(IDC_BUTTONFAX)->ShowWindow(false);
	GetDlgItem(IDC_USER)->ShowWindow(false);
	
	GetDlgItem(IDC_BRECORDMYNOTICE)->ShowWindow(false);
	GetDlgItem(IDC_BUTTONUPDATE)->ShowWindow(false);
	//GetDlgItem(IDC_TRYLISTEN)->ShowWindow(false);
	GetDlgItem(IDC_OPENSCANNER)->ShowWindow(false);
	GetDlgItem(IDC_CHECKRECVFAX)->ShowWindow(false);
	GetDlgItem(IDC_CHECKAUTOFORWARDFAX)->ShowWindow(false);
	GetDlgItem(IDC_MYCONTACTLIST)->ShowWindow(false);
	//GetDlgItem(IDC_MYLOGLIST)->ShowWindow(false);
	GetDlgItem(IDC_SHAIXUAN)->ShowWindow(false);
	GetDlgItem(IDC_CHECK001)->ShowWindow(false);
	GetDlgItem(IDC_CHECK002)->ShowWindow(false);
	GetDlgItem(IDC_CHECK003)->ShowWindow(false);
	GetDlgItem(IDC_CHECK004)->ShowWindow(false);
	GetDlgItem(IDC_CHECK005)->ShowWindow(false);
	GetDlgItem(IDC_CHECK006)->ShowWindow(false);
	GetDlgItem(IDC_CHECK007)->ShowWindow(false);
	GetDlgItem(IDC_CHECK008)->ShowWindow(false);
	GetDlgItem(IDC_SEARCH)->ShowWindow(false);
	GetDlgItem(IDC_DUTYINFO)->ShowWindow(false);
	//GetDlgItem(IDC_LOGREMARK)->ShowWindow(false);
	//GetDlgItem(IDC_LOGSAVE)->ShowWindow(false);
	//GetDlgItem(IDC_LOGCANCEL)->ShowWindow(false);
	//GetDlgItem(IDC_DATETIMEPICKER1)->ShowWindow(false);
	GetDlgItem(IDC_BUTTONSHOWLOG)->ShowWindow(false);
	GetDlgItem(IDC_BUTTONCONTACT)->ShowWindow(false);
	

	GetDlgItem(IDC_WELCOME2)->ShowWindow(true);
	GetDlgItem(IDC_LOGIN)->ShowWindow(true);
	GetDlgItem(IDC_QUITSYSTEM)->ShowWindow(true);
	GetDlgItem(IDC_REGISTER)->ShowWindow(true);
	GetDlgItem(IDC_LOGNAME2)->ShowWindow(true);
	GetDlgItem(IDC_LOGCOMBOBOX)->ShowWindow(true);
	GetDlgItem(IDC_LOGMIMA2)->ShowWindow(true);
	GetDlgItem(IDC_LOGNMIMA)->ShowWindow(true);
	GetDlgItem(IDC_LOGCOMBOBOX)->SetWindowTextA(_T(""));
	GetDlgItem(IDC_LOGNMIMA)->SetWindowTextA(_T(""));


	
	//GetLocalTime(&whu_SystemTime);
	//CTime m_time(whu_SystemTime.wYear,whu_SystemTime.wMonth,whu_SystemTime.wDay,0,0,0);
	//m_DataCtrl.SetTime(&m_time);
	//whu_logchangetime();
	//CString time = _T("");
	//GetLocalTime(&whu_SystemTime);
	//time.Format(_T("%d-%d-%d-%d:%d"),whu_SystemTime.wYear,whu_SystemTime.wMonth,whu_SystemTime.wDay,whu_SystemTime.wHour,whu_SystemTime.wMinute);
	//whu_AddLog(time,m_LogName,_T("�˳���¼"),_T(""),_T(""),_T(""),_T(""),_T(""),_T(""));
	OnSaveLogToFile();
	//OnRefreshLog();

}


void Cwhu_VoiceCardDlg::OnTimer(UINT_PTR nIDEvent)
{
	// TODO: Add your message handler code here and/or call default
	int Event = nIDEvent;
	int faxch = 8;
	int TrkCh = 2;
	int ch = 0;
	switch(Event)
	{
	case 1:
		GetLocalTime(&whu_SystemTime);
		if (whu_TaskMode == TASK_FAXDIAL) //���ʹ����ʱ����ʾ����//
		{
			if (whu_FaxSendState == FAX_SENDING)
			{
				whu_FaxAllBites = SsmFaxGetAllBytes(faxch);
				// m_AllFaxPage = fBmp_GetFileAllPage(whu_SendFaxFile); //�õ�������ҳ����//
				long m_DoneFaxPage = SsmFaxGetPages(faxch);
				//int m_speed1 = SsmFaxGetSpeed(faxch); //�鿴���淢�͵�����
				//SsmFaxSetMaxSpeed(m_speed1); // ���ô��淢�͵�����
				if (whu_FaxAllBites >0)
				{
					whu_FaxDoneBites = SsmFaxGetSendBytes(faxch);
					//int a =(int)(((whu_FaxDoneBites/whu_FaxAllBites + m_DoneFaxPage)*1000)/m_AllFaxPage); //���㷨��������//
					//int a =(int) (whu_FaxDoneBites*1000/whu_FaxAllBites);//��ʾ��ҳ�����㷨//
					int a =m_ProgressCtrl.GetPos();
					a++;
					if (a>990)
					{
						a = 990;
					}
					m_ProgressCtrl.SetPos(a);
				}
			}
			else if (whu_FaxSendState == FAX_SUCCESS)
			{
				m_ProgressCtrl.SetPos(990);
			}
		}
		else if (whu_TaskMode == TASK_FAXRECV)
		{
				//���մ���Ľ�����//
			if (whu_FaxRecvState == FAX_PREPARE) //���մ����ʱ��û�취��ʾ���ȣ�������ʾ�յ����ֽ���//
			{
				int pos = m_ProgressCtrl.GetPos();
				pos = pos+3;
				if (pos>990)
				{
					pos = 990;
				}
				m_ProgressCtrl.SetPos(pos);
				/*whu_FaxDoneBites = SsmFaxGetRcvBytes(faxch);
				if (whu_FaxDoneBites >0)
				{

					int a =(int) (whu_FaxDoneBites/100);
					if (a >=990)
					{
						a = 990;
					}
					m_ProgressCtrl.SetPos(a);
					if (whu_FaxSendState == FAX_SUCCESS)
					{
						m_ProgressCtrl.SetPos(999);
					}
				}*/
			}else if (whu_FaxRecvState == FAX_SUCCESS)
			{
				m_ProgressCtrl.SetPos(990);
			}
		}
		
		UpdateData(true);
		whu_SearchContact(m_ContactSearch);

		break;
	case 2: //����绰���
		if (m_BPhoneDialDone == true&&whu_UserChState == USER_GET_PHONE_NUM&&m_BPhoneButtonPress==true)
		{
			m_BPhoneButtonPress = false;
			char m_NumBuf[30]={'\0'};
			SsmGetDtmfStr(ch,m_NumBuf);		//����ĵ绰����//
			whu_NumberBuf.Format("%s",m_NumBuf);
			if(whu_TrunkChState!=TRK_IDLE)			//no idle trunk channel available
			{
				SsmSendTone(ch,1);						// send busy tone	
				whu_UserChState = USER_WAIT_HANGUP;
			}
			else 
			{
				whu_UserChState = USER_IDLE;  //�������if����ϣ������ظ�����//
				whu_Dial(whu_NumberBuf);
			}
		}
		else{
			m_BPhoneDialDone = true; //����绰��ʱ�򣬵���������3���Ժ����°������£�����Ϊ�����������//
		}
		//whu_SaveLog();
		break;
	}
	CDialogEx::OnTimer(nIDEvent);
}
void Cwhu_VoiceCardDlg::whu_SearchContact(CString m_str)
{
	static bool m_Show = true;
	if (m_str!=_T(""))
	{
		m_MyContactList.DeleteAllItems();
		m_Show = false;//˵����ĳ�����ݱ�ɾ����;
		for (int i=1;i<m_ContactListData[0].GetSize();i++)
		{
			bool find = false;
			for (int j =0;j<ConHeadNum;j++)
			{
				CString listdata = m_ContactListData[j].GetAt(i);
				CString m_PinYin = whu_HzTopy.HzToPinYin(listdata);
				int ss = listdata.Find(m_str);
				int ss2 = m_PinYin.Find(m_str);
				if (ss !=-1||ss2!=-1)
				{
					find = true;
				}
			}
			if (find == true)
			{
				CStringArray m_Arr;
				for (int j=0;j<ConHeadNum;j++)
				{
					m_Arr.Add(m_ContactListData[j].GetAt(i));
				}
				Lin_InsertList(m_MyContactList,m_Arr);
			}
		}
	}else {
		if (m_Show == false)
		{
			m_MyContactList.DeleteAllItems();
			while(m_MyContactList.GetHeaderCtrl()->GetItemCount()>0)
			{
				m_MyContactList.DeleteColumn(0);
			}
			Lin_ImportArrToList(m_ContactListData,ConHeadNum,m_MyContactList);
			/*for (int i=1;i<m_ContactListData[0].GetSize();i++)
			{
				char buf[3];
				itoa(i+1,buf,10);
				m_MyContactList.InsertItem(i,buf);
				for (int j =0;j<ConHeadNum;j++)
				{
					m_MyContactList.SetItemText(i,j,m_ContactListData[j].GetAt(i)); 
				}
			}*/
			m_Show = true;

		}
	}
}

void Cwhu_VoiceCardDlg::OnBnClickedQuitsystem()
{
	
	m_SaveEveryingDlg = new CSaveEveryingDlg();
	m_SaveEveryingDlg->whu_ParentDlg = this;
	m_SaveEveryingDlg->Create(IDD_SAVEEVERYING,this);
	m_SaveEveryingDlg->ShowWindow(SW_SHOW);
	m_BQuitSystem = true;
	m_SaveEveryingDlg->PostMessage(WM_UPATEDATA, 0, NULL);
	//m_Dlg->DoModal();
	// TODO: Add your control notification handler code here
	CString m_FilePath = _T("C:\\Whu_Data\\Excel�ļ�\\ͨѶ¼.xlsx");
	CString m_Folder = _T("C:\\Whu_Data\\Excel�ļ�");
	if (m_MyContactList.GetItemCount()>1)
	{
		if (whu_findfile(m_FilePath,m_Folder)==true)
		{
			CFile m_Temp;
			m_Temp.Remove(m_FilePath);
		}
		Lin_ExportListToExcel(m_MyContactList,m_FilePath);
	}
	CString m_backUpFilePath = _T("C:\\Whu_Data\\�����ļ�\\ͨѶ¼_����.xlsx");
	CopyFile(m_FilePath,m_backUpFilePath,false);
	////���ֵ���//
	m_FilePath = _T("C:\\Whu_Data\\Excel�ļ�\\ֵ���.xlsx");
	m_Folder = _T("C:\\Whu_Data\\Excel�ļ�");
	if (DutyHeadNum > 0)
	{
		if (whu_findfile(m_FilePath,m_Folder)==true)
		{
			CFile m_Temp;
			m_Temp.Remove(m_FilePath);
		}
		Lin_ImpoerArrToExcel(whu_DutyPersons,DutyHeadNum,m_FilePath);
	}
	m_backUpFilePath = _T("C:\\Whu_Data\\�����ļ�\\ֵ���_����.xlsx");
	CopyFile(m_FilePath,m_backUpFilePath,false);


	m_FilePath = _T("C:\\Whu_Data\\Excel�ļ�\\�����Զ�ת����.xlsx");
	m_Folder = _T("C:\\Whu_Data\\Excel�ļ�");
	if (whu_findfile(m_FilePath,m_Folder)==true)
	{
		CFile m_Temp;
		m_Temp.Remove(m_FilePath);
	}
	Lin_ExportForArrToExcel(whu_FaxAutoForwardArr,m_FilePath);

	m_backUpFilePath = _T("C:\\Whu_Data\\�����ļ�\\�����Զ�ת����_����.xlsx");
	CopyFile(m_FilePath,m_backUpFilePath,false);

	m_FilePath = _T("C:\\Whu_Data\\Excel�ļ�\\�û���Ϣ.xlsx");
	m_Folder = _T("C:\\Whu_Data\\Excel�ļ�");
	if (whu_findfile(m_FilePath,m_Folder)==true)
	{
		CFile m_Temp;
		m_Temp.Remove(m_FilePath);
	}
	Lin_ImpoerArrToExcel(whu_User,4,m_FilePath);

	m_backUpFilePath = _T("C:\\Whu_Data\\�����ļ�\\�û���Ϣ_����.xlsx");
	CopyFile(m_FilePath,m_backUpFilePath,false);

	m_SaveEveryingDlg->DestroyWindow();
	DestroyWindow();
}


void Cwhu_VoiceCardDlg::OnSendMail()
{
	// TODO: Add your command handler code here
	CString time = _T("");
	time.Format(_T("%d-%d-%d-%d:%d"),whu_SystemTime.wYear,whu_SystemTime.wMonth,whu_SystemTime.wDay,whu_SystemTime.wHour,whu_SystemTime.wMinute);
	
	CString m_pcUserName = _T("512290279");
	CString	m_pcUserPassWord = _T("arthur135060lxm");
	CString m_pcSenderName = _T("������");
	CString m_pcSender = _T("512290279@qq.com");
	//CString m_pcReceiver = m_MyContactData.E_Mail;
	CString m_pcReceiver = _T("linxiongmin@whu.edu.cn");
	CString m_pcTitle = _T("�ʼ�����");
	CString m_pcBody = _T("�ʼ�����");
	CString m_pcIPAddr = _T(""); //������IP
	CString m_pcIPName = _T("qq.com ");
	whu_MailData.m_pcUserName = m_pcUserName.GetBuffer();
	whu_MailData.m_pcUserPassWord = m_pcUserPassWord.GetBuffer();
	whu_MailData.m_pcSenderName = m_pcSenderName.GetBuffer();
	whu_MailData.m_pcSender =m_pcSender.GetBuffer();
	whu_MailData.m_pcReceiver = m_pcReceiver.GetBuffer();
	whu_MailData.m_pcTitle = m_pcTitle.GetBuffer();
	whu_MailData.m_pcBody = m_pcBody.GetBuffer();
	whu_MailData.m_pcIPAddr = m_pcIPAddr.GetBuffer();
	whu_MailData.m_pcIPName = m_pcIPName.GetBuffer();
	
	//whu_MyMail.SendMail(whu_MailData);

	whu_AddLog(time,m_LogName,_T("�����ʼ�"),_T(""),_T(""),m_MyContactData.E_Mail,_T(""),_T(""),_T(""));
	whu_SaveLog();
	ShellExecute(NULL, "open", _T("D:\\Program Files\\Foxmail 7.0\\Foxmail.exe"),NULL,NULL, SW_SHOWNORMAL );
}


void Cwhu_VoiceCardDlg::OnSeeTelLog()
{
	// TODO: Add your command handler code here
	whu_SearchLog(m_MyContactData.TelNum);
}
void Cwhu_VoiceCardDlg::whu_SaveLog()
{

	//�б�����ݲ������ϸ���ʱ���������//

	//���б����ݱ��浽�ַ���������//
	for (int i=0;i<10;i++)
	{
		m_LogListData[i].RemoveAll();
	}
	for (int i=0;i<m_MyLogList.GetItemCount();i++)  //����
	{
		for (int j=0;j<m_MyLogList.GetHeaderCtrl()->GetItemCount();j++) //����
		{
			CString m_str = m_MyLogList.GetItemText(i,j);
			string str = m_str.GetBuffer();
			str=trim((string)str);
			m_str = str.c_str();
			m_LogListData[j].InsertAt(i,m_str);
		}	
	}
	//���б����ݱ��浽�ĵ���//
	//�����жϣ��ж��Ƿ��ǽ������־//
	GetLocalTime(&whu_SystemTime);
	CTime m_time1(whu_SystemTime.wYear,whu_SystemTime.wMonth,whu_SystemTime.wDay,0,0,0); //��ǰ����
	CTime m_time2;
	m_DataCtrl.GetTime(m_time2);  //�û��鿴��־��ʱ��ѡ�������
	if (m_time1.GetDay()== m_time2.GetDay()&&m_time1.GetMonth()==m_time2.GetMonth()&&m_time1.GetYear()==m_time2.GetYear())   //ֻ�е�ǰʱ����û�ѡ��ʱ��һ�£����ܱ�����־//
	{
		OnSaveLogToFile();
	}
}

void Cwhu_VoiceCardDlg::whu_SearchLog(CString m_str)
{
	/*whu_SaveLog();*/
	static bool m_Show = true;
	int m_Delete[1000];
	if (m_str!=_T(""))
	{
		int deletenum = 0;
		for (int i=0;i<m_MyLogList.GetItemCount();i++)
		{
			int num1 = m_LogListData[0].GetSize();
			int num2 = m_MyLogList.GetItemCount();
			bool find = false;
			for (int j =0;j<10;j++)
			{
				CString listdata = m_MyLogList.GetItemText(i,j);
				if (listdata == m_str)
				{
					find = true;
				}
			}
			if (find == false)
			{
				//m_MyLogList.DeleteItem(i);
				m_Show = false;//˵����ĳ�����ݱ�ɾ����;
				m_Delete[deletenum] = i;
				deletenum++;
			}
		}
		for (int i=deletenum-1;i>=0;i--)
		{
			m_MyLogList.DeleteItem(m_Delete[i]);
		}

	}
}

void Cwhu_VoiceCardDlg::OnRefreshLog()
{
	// TODO: Add your command handler code here
	m_MyLogList.DeleteAllItems();
	int num = m_LogListData[0].GetSize();
	for (int i=0;i<m_LogListData[0].GetSize();i++)
	{
		char buf[3];
		itoa(i+1,buf,10);
		m_MyLogList.InsertItem(i,buf);
		for (int j =0;j<10;j++)
		{
			m_MyLogList.SetItemText(i,j,m_LogListData[j].GetAt(i)); 
		}
	}
	//OnSaveLogToFile();
	//GetLocalTime(&whu_SystemTime);
	//CTime m_time(whu_SystemTime.wYear,whu_SystemTime.wMonth,whu_SystemTime.wDay,0,0,0);
	//m_DataCtrl.SetTime(&m_time);
	//whu_logchangetime();
}

void Cwhu_VoiceCardDlg::OnRefreshLog2()
{
	// TODO: Add your command handler code here
	OnRefreshLog();
}


void Cwhu_VoiceCardDlg::OnEditCopy()
{
	// TODO: Add your command handler code here
	int i=0;
}


void Cwhu_VoiceCardDlg::OnSeeFaxLog()
{
	// TODO: Add your command handler code here
	whu_SearchLog(m_MyContactData.FaxNum);
}


void Cwhu_VoiceCardDlg::OnSeeOfficeNumLog()
{
	// TODO: Add your command handler code here
	whu_SearchLog(m_MyContactData.OfficeNum);
}
void Cwhu_VoiceCardDlg::whu_LinkToUserSql()
{
	_ConnectionPtr pMyConnect("ADODB.Connection");
	pMyConnect.CreateInstance("ADODB.Connection");  
	//�ҵĵ���
	_bstr_t strConnect = "Provider=SQLOLEDB.1;Server=CHINESE-70D3153\\WHSW;Database=whu;uid=sa;pwd=123456;"; 
	//���ػ�
	//_bstr_t strConnect = "Provider=SQLOLEDB.1;Server=MICROSOF-0FCDA5\\WHSW;Database=whu;uid=whu;pwd=123456;"; 
	try  
	{  
		pMyConnect->Open(strConnect, "", "", adModeUnknown);  
	}  
	catch(_com_error &e)  
	{  
		MessageBox(e.Description(), _T("����"),MB_OK|MB_ICONINFORMATION);  
	}
	CString strSql="select * from [user]";//����ִ�е�SQL���  
	_RecordsetPtr pRecordset;  
	HRESULT hTRes;
	hTRes = pRecordset.CreateInstance(_T("ADODB.Recordset"));
	if (SUCCEEDED(hTRes))
	{
		hTRes = pRecordset->Open((LPTSTR)strSql.GetBuffer(130),
			pMyConnect.GetInterfacePtr(),
			adOpenDynamic,adLockPessimistic,adCmdText);
		int i=0;
		char buf[10] = {'\0'};
		while(!(pRecordset->adoEOF))
		{
			//�õ��û���//
			CString m_name = GetStringFromVariant(pRecordset->GetCollect(_T("�û���")));
			string str = m_name.GetBuffer();
			str=trim((string)str);
			m_name = str.c_str();
			whu_User[0].InsertAt(i,m_name);

			//�õ��û�����//
			CString m_password = GetStringFromVariant(pRecordset->GetCollect(_T("����")));
			str = m_password.GetBuffer();
			str=trim((string)str);
			m_password = str.c_str();
			whu_User[1].InsertAt(i,m_password);

			i++;
			if(!(pRecordset->adoEOF))
				pRecordset->MoveNext();
		}
	}

}
void Cwhu_VoiceCardDlg::whu_LinkToDutySql()
{

		//////ֵ�����б��ʼ��
		LONG lStyle; 
		lStyle = GetWindowLong(m_MyContactList.m_hWnd, GWL_STYLE);
		lStyle &= ~LVS_TYPEMASK; 
		lStyle |= LVS_REPORT; 
		SetWindowLong(m_MyContactList.m_hWnd, GWL_STYLE, lStyle);
		DWORD dwStyle = m_MyContactList.GetExtendedStyle(); 
		dwStyle |= LVS_EX_FULLROWSELECT;
		dwStyle |= LVS_EX_GRIDLINES;
		dwStyle |= LVS_EX_CHECKBOXES;
		m_MyContactList.SetExtendedStyle(dwStyle); 
		m_MyContactList.InsertColumn( 0, "���", LVCFMT_LEFT, 60 );
		m_MyContactList.InsertColumn( 1, _T("ʱ��"), LVCFMT_LEFT, 220 ); 
		m_MyContactList.InsertColumn( 2, _T("����"), LVCFMT_LEFT, 180 );
		m_MyContactList.InsertColumn( 3, _T("ֵ��ص�"), LVCFMT_LEFT, 100 );
		m_MyContactList.InsertColumn( 4, _T("ֵ��绰"), LVCFMT_LEFT, 100 );
		m_MyContactList.InsertColumn( 5, _T("�����쵼"), LVCFMT_LEFT, 100 ); 
		//���ӵ�MS SQL Server  
		_ConnectionPtr pMyConnect("ADODB.Connection");
		pMyConnect.CreateInstance("ADODB.Connection");  

		//�ҵĵ���
		_bstr_t strConnect = "Provider=SQLOLEDB.1;Server=CHINESE-70D3153\\WHSW;Database=whu;uid=sa;pwd=123456;"; 
		//���ػ�
		//_bstr_t strConnect = "Provider=SQLOLEDB.1;Server=MICROSOF-0FCDA5\\WHSW;Database=whu;uid=whu;pwd=123456;"; 
		try  
		{  
			pMyConnect->Open(strConnect, "", "", adModeUnknown);  
		}  
		catch(_com_error &e)  
		{  
			MessageBox(e.Description(), _T("����"),MB_OK|MB_ICONINFORMATION);  
		}
		CString strSql="select * from onduty";//����ִ�е�SQL���  
		_RecordsetPtr pRecordset;  
		HRESULT hTRes;
		hTRes = pRecordset.CreateInstance(_T("ADODB.Recordset"));
		if (SUCCEEDED(hTRes))
		{
			hTRes = pRecordset->Open((LPTSTR)strSql.GetBuffer(130),
				pMyConnect.GetInterfacePtr(),
				adOpenDynamic,adLockPessimistic,adCmdText);
			int i=0;
			char buf[10] = {'\0'};
			while(!(pRecordset->adoEOF))
			{
				//����ѯ�����ʾ���б��ؼ���//
				m_MyContactList.InsertItem(i,GetStringFromVariant(pRecordset->GetCollect("ID"))); 
				//m_MyContactList.InsertItem(i,_itoa(i+1,buf,10));
				m_MyContactList.SetItemText(i,1,GetStringFromVariant(pRecordset->GetCollect("ʱ��"))); 
				m_MyContactList.SetItemText(i,2,GetStringFromVariant(pRecordset->GetCollect("����"))); 
				m_MyContactList.SetItemText(i,3,GetStringFromVariant(pRecordset->GetCollect("ֵ��ص�"))); 
				m_MyContactList.SetItemText(i,4,GetStringFromVariant(pRecordset->GetCollect("ֵ��绰")));
				m_MyContactList.SetItemText(i,5,GetStringFromVariant(pRecordset->GetCollect("�����쵼")));
				
				////�õ�ֵ��ʱ��//
				//CString m_name = GetStringFromVariant(pRecordset->GetCollect(_T("ʱ��")));
				//string str = m_name.GetBuffer();
				//str=trim((string)str);
				//m_name = str.c_str();
				//whu_DutyPersons[0].InsertAt(i,m_name);

				////�õ��û���//
				//CString m_name = GetStringFromVariant(pRecordset->GetCollect(_T("����")));
				//string str = m_name.GetBuffer();
				//str=trim((string)str);
				//m_name = str.c_str();
				//whu_DutyPersons[1].InsertAt(i,m_name);

				////�õ�ֵ��ص�//
				//CString m_name = GetStringFromVariant(pRecordset->GetCollect(_T("ֵ��ص�")));
				//string str = m_name.GetBuffer();
				//str=trim((string)str);
				//m_name = str.c_str();
				//whu_DutyPersons[2].InsertAt(i,m_name);

				////�õ��칫�ҵ绰//
				//CString m_name = GetStringFromVariant(pRecordset->GetCollect(_T("����")));
				//string str = m_name.GetBuffer();
				//str=trim((string)str);
				//m_name = str.c_str();
				//whu_DutyPersons[3].InsertAt(i,m_name);

				////�õ������쵼//
				//CString m_leader = GetStringFromVariant(pRecordset->GetCollect(_T("�����쵼")));
				//str = m_leader.GetBuffer();
				//str=trim((string)str);
				//m_leader = str.c_str();
				//whu_DutyPersons[4].InsertAt(i,m_leader);
				//m_ComboBoxLeader.AddString(whu_DutyPersons[4].GetAt(i));




				i++;
				if(!(pRecordset->adoEOF))
					pRecordset->MoveNext();
			}
		}
		//������ϵ���б�����ո�//
		for (int i=0;i<m_MyContactList.GetItemCount();i++)
		{
			for (int j=0;j<6;j++)
			{
				CString m_str = m_MyContactList.GetItemText(i,j);
				string str = m_str.GetBuffer();
				str=trim((string)str);
				m_str = str.c_str();
				whu_DutyPersons[j].InsertAt(i,m_str);
			}
		}
		for (int i=0;i<whu_DutyPersons[0].GetSize();i++)
		{
			for (int j=0;j<6;j++)
			{
				CString ss = whu_DutyPersons[j].GetAt(i);
				m_MyContactList.SetItemText(i,j,ss);
			}	
		}
		
	
}
CString Cwhu_VoiceCardDlg::whu_ShowLeader(CString m_DutyPerson)
{
	CString m_Str = _T("�����쵼");
	for (int i=0;i<whu_DutyPersons[1].GetSize();i++)
	{
		CString m_test = whu_DutyPersons[1].GetAt(i);
		if (whu_DutyPersons[1].GetAt(i) == m_DutyPerson)
		{
			return whu_DutyPersons[4].GetAt(i);
		}
	}
	return m_Str;
}

void Cwhu_VoiceCardDlg::OnBnClickedShowduty() //��ʾֵ����Ϣ//
{
	// TODO: Add your control notification handler code here
	//CString m_str;
	//GetDlgItemText(IDC_SHOWDUTY,m_str);
	//if (m_str == _T("ֵ����Ϣ"))
	//{
	//	m_BShowDutyList = true;
	//	SetDlgItemText(IDC_SHOWDUTY,_T("��ʾ��ϵ��"));
	//	while (m_MyContactList.GetHeaderCtrl()->GetItemCount()>0)
	//	{
	//		m_MyContactList.DeleteColumn(0);
	//	}
	//	m_MyContactList.DeleteAllItems();
	//	whu_LinkToDutySql();
	//	GetDlgItem(IDC_SHAIXUAN)->ShowWindow(false);
	//	GetDlgItem(IDC_CHECK001)->ShowWindow(false);
	//	GetDlgItem(IDC_CHECK002)->ShowWindow(false);
	//	GetDlgItem(IDC_CHECK003)->ShowWindow(false);
	//	GetDlgItem(IDC_CHECK004)->ShowWindow(false);
	//	GetDlgItem(IDC_CHECK005)->ShowWindow(false);
	//	GetDlgItem(IDC_CHECK006)->ShowWindow(false);
	//	GetDlgItem(IDC_CHECK007)->ShowWindow(false);
	//	GetDlgItem(IDC_CHECK008)->ShowWindow(false);
	//	GetDlgItem(IDC_SEARCH)->ShowWindow(false);

	//}
	//else{
	//	m_BShowDutyList = false;
	//	SetDlgItemText(IDC_SHOWDUTY,_T("ֵ����Ϣ"));
	//	while (m_MyContactList.GetHeaderCtrl()->GetItemCount()>0)
	//	{
	//		m_MyContactList.DeleteColumn(0);
	//	}
	//	m_MyContactList.DeleteAllItems();
	//	//whu_LinkToSQL(); //����bug

	//	//��ϵ���б��ʼ��
	//	LONG lStyle; 
	//	lStyle = GetWindowLong(m_MyContactList.m_hWnd, GWL_STYLE);
	//	lStyle &= ~LVS_TYPEMASK; 
	//	lStyle |= LVS_REPORT; 
	//	SetWindowLong(m_MyContactList.m_hWnd, GWL_STYLE, lStyle);
	//	DWORD dwStyle = m_MyContactList.GetExtendedStyle(); 
	//	dwStyle |= LVS_EX_FULLROWSELECT;
	//	dwStyle |= LVS_EX_GRIDLINES;
	//	dwStyle |= LVS_EX_CHECKBOXES;
	//	m_MyContactList.SetExtendedStyle(dwStyle); 
	//	m_MyContactList.InsertColumn( 0, _T("���"), LVCFMT_LEFT, 40 );
	//	m_MyContactList.InsertColumn( 1, _T("����"), LVCFMT_LEFT, 120 ); 
	//	m_MyContactList.InsertColumn( 2, _T("��λ"), LVCFMT_LEFT, 120 );
	//	m_MyContactList.InsertColumn( 3, _T("����"), LVCFMT_LEFT, 160 );
	//	m_MyContactList.InsertColumn( 4, _T("ְ��"), LVCFMT_LEFT, 80 );
	//	m_MyContactList.InsertColumn( 5, _T("�ֻ�"), LVCFMT_LEFT, 160 ); 
	//	m_MyContactList.InsertColumn( 6, _T("�칫�ҵ绰"), LVCFMT_LEFT, 160 ); 
	//	m_MyContactList.InsertColumn( 7, _T("�����"), LVCFMT_LEFT, 120 ); 
	//	m_MyContactList.InsertColumn( 8, _T("����"), LVCFMT_LEFT, 150 );
	//	for (int i=0;i<m_ContactListData[0].GetSize();i++)
	//	{
	//		CString m_Str;
	//		m_Str.Format(_T("%d"),i+1);
	//		m_MyContactList.InsertItem(i,m_Str);
	//		for (int j=0;j<9;j++)
	//		{
	//			CString ss = m_ContactListData[j].GetAt(i);
	//			m_MyContactList.SetItemText(i,j,ss);
	//		}	
	//	}

	//	GetDlgItem(IDC_SHAIXUAN)->ShowWindow(true);
	//	GetDlgItem(IDC_CHECK001)->ShowWindow(true);
	//	GetDlgItem(IDC_CHECK002)->ShowWindow(true);
	//	GetDlgItem(IDC_CHECK003)->ShowWindow(true);
	//	GetDlgItem(IDC_CHECK004)->ShowWindow(true);
	//	GetDlgItem(IDC_CHECK005)->ShowWindow(true);
	//	GetDlgItem(IDC_CHECK006)->ShowWindow(true);
	//	GetDlgItem(IDC_CHECK007)->ShowWindow(true);
	//	GetDlgItem(IDC_CHECK008)->ShowWindow(true);
	//	GetDlgItem(IDC_SEARCH)->ShowWindow(true);
	//	
	//}
	
	//Cwhu_OnDutyDlg *m_Dlg = new Cwhu_OnDutyDlg();
	//m_Dlg->whu_ParentDlg = this;
	//m_Dlg->Create(IDD_ONDUTY,this);
	//m_Dlg->ShowWindow(SW_SHOW);

	Cwhu_OnDutyDlg m_DutyDlg;
	m_DutyDlg.whu_ParentDlg = this;
	m_DutyDlg.DoModal();
}


void Cwhu_VoiceCardDlg::OnDtnDatetimechangeDatetimepicker1(NMHDR *pNMHDR, LRESULT *pResult)  //�ı�ʱ�䣬�鿴��Ӧ��־
{
	LPNMDATETIMECHANGE pDTChange = reinterpret_cast<LPNMDATETIMECHANGE>(pNMHDR);
	// TODO: Add your control notification handler code here
	whu_logchangetime();	
	*pResult = 0;
}
bool Cwhu_VoiceCardDlg::whu_findfile(CString m_file,CString folder)
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
void Cwhu_VoiceCardDlg::whu_logchangetime()
{


	CTime time;
	m_DataCtrl.GetTime(time);
	CString strTime;
	strTime.Format(_T("%d%s%d%s%d"),time.GetYear(),_T("-"),time.GetMonth(),_T("-"),time.GetDay());
	CString file = _T("");
	file.Format(_T("%s%s%s"),_T("C:\\Whu_Data\\ֵ��ϵͳ��־\\"),strTime,_T(".txt"));
	//�ļ�����//
	CString folder = _T("C:\\Whu_Data\\ֵ��ϵͳ��־");
	//���ļ�����Ѱ���ļ�//
	if (whu_findfile(file,folder))
	{
		whu_TxtToList(file,&m_MyLogList);
	}
	else
	{
		m_MyLogList.DeleteAllItems();
	}
	whu_SaveLog(); //������䣬����������������־�ļ�¼�ˡ�//
}




void Cwhu_VoiceCardDlg::OnSaveLogToFile()
{
	// TODO: Add your command handler code here
	CTime time;
	m_DataCtrl.GetTime(time);
	CString strTime;
	strTime.Format(_T("%d%s%d%s%d"),time.GetYear(),_T("-"),time.GetMonth(),_T("-"),time.GetDay());
	CString file = _T("");
	file.Format(_T("%s%s%s"),_T("C:\\Whu_Data\\ֵ��ϵͳ��־\\"),strTime,_T(".txt"));
	whu_ListToTxt(&m_MyLogList,file);
}
void Cwhu_VoiceCardDlg::whu_ListToTxt(CListCtrl *m_list,CString m_txtfile)
{
	////�����ݱ��浽�ĵ�//
	////���統����־���ڣ��ȶ�ȡ//
	//CStringArray m_allstr ;
	//CString folder = _T("C:\\Whu_Data\\ֵ��ϵͳ��־");
	//if (whu_findfile(m_txtfile,folder) == true)
	//{
	//	CStdioFile m_file(m_txtfile,CFile::modeRead);
	//	CString m_str = _T("");
	//	while (m_file.ReadString(m_str))
	//	{
	//		m_allstr.Add(m_str);
	//	}
	//	m_file.Close();
	//	m_file.Remove(m_txtfile);
	//}
	////�����µ��ļ�//
	//CStdioFile m_file(m_txtfile,CFile::modeCreate|CFile::modeWrite|CFile::modeNoTruncate|CFile::modeRead);
	//for (int i=0;i<m_list->GetItemCount();i++) //����//
	//{
	//	CString m_LineStr;
	//	for (int j =1;j<m_list->GetHeaderCtrl()->GetItemCount();j++) //����//
	//	{
	//		CString m_str;
	//		m_str = m_list->GetItemText(i,j);
	//		string str = m_str.GetBuffer();
	//		str=trim((string)str);
	//		m_str = str.c_str();
	//		m_str = m_str+_T(" "); //ÿ������֮���Կո�������//
	//		m_LineStr = m_LineStr+m_str;
	//	}
	//	//�Ƚϡ�ɾ���ظ���Ϣ//
	//	bool m_bSame = false;
	//	for (int i=0;i<m_allstr.GetCount();i++)
	//	{
	//		if (m_allstr.GetAt(i)==m_LineStr)
	//		{
	//			m_bSame = true;
	//		}
	//	}
	//	if (m_bSame == false)
	//	{
	//		m_allstr.Add(m_LineStr);
	//	}	
	//}
	////�������Աɾ��ĳ�����ݣ�����ҲҪ����ɾ//
	//int NumOfStr = m_allstr.GetCount();
	//for (int i=0;i<NumOfStr;i++)
	//{
	//	if (m_allstr.GetAt(i)!=m_LogDeleteStr)
	//	{
	//		CString m_str = m_allstr.GetAt(i)+_T("\n");
	//		m_file.WriteString(m_str);
	//	}
	//}
	//m_file.Close();

	//�����ݱ��浽�ĵ�//
	//���統����־���ڣ��ȶ�ȡ//
	CStringArray m_allstr ;
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
			string str = m_str.GetBuffer();
			str=trim((string)str);
			m_str = str.c_str();
			m_str = m_str+_T(" "); //ÿ������֮���Կո�������//
			m_LineStr = m_LineStr+m_str;
		}
		m_LineStr = m_LineStr+_T("\n");
		m_file.WriteString(m_LineStr);
	}
	m_file.Close();
}
void Cwhu_VoiceCardDlg::whu_TxtToList(CString m_txtfile,CListCtrl *m_list)
{
	CStdioFile m_file(m_txtfile,CFile::modeRead);
	CString m_str = _T("");
	int ColumNum = m_list->GetHeaderCtrl()->GetItemCount();
	int RowNum = 0;
	m_list->DeleteAllItems();
	while (m_file.ReadString(m_str)) //��һ�е�����
	{
		char buf[4] = _T("");
		itoa(RowNum+1,buf,10);
		m_list->InsertItem(RowNum,buf);
		CString m_No=_T("");
		m_No.Format(_T("%d"),RowNum+1);
		m_list->SetItemText(RowNum,0,m_No);
		for (int j=1;j<ColumNum;j++) //////
		{
			int pos = m_str.Find(_T(" "));
			CString m_item = m_str.Left(pos);
			m_list->SetItemText(RowNum,j,m_item);
			m_str = m_str.Right(m_str.GetLength()-pos-1);
		}
		RowNum++;
	}
	m_file.Close();

}

void Cwhu_VoiceCardDlg::OnMenuSeeHistory() //�˵�-�Ҽ�-�鿴��¼����Ӧ����//
{
	// TODO: Add your command handler code here
	if (m_MyContactData.Name!=_T(""))
	{
		whu_SearchLog(m_MyContactData.Name); //������������ƥ�����Ϣ
	}
	else if(m_MyContactData.Unit!=_T(""))
	{
		whu_SearchLog(m_MyContactData.Unit); //�����뵥λ��ƥ�����Ϣ
	}
	else if(m_MyContactData.Department!=_T(""))
	{
		whu_SearchLog(m_MyContactData.Department); //�����벿����ƥ�����Ϣ
	}
	else if(m_MyContactData.Job!=_T(""))
	{
		whu_SearchLog(m_MyContactData.Job); //������ְ����ƥ�����Ϣ
	}
}


//void Cwhu_VoiceCardDlg::OnOpenPDFFolder()
//{
//	// TODO: Add your command handler code here
//	ShellExecute(NULL, "open", _T("C:\\Whu_Data\\ɨ����pdf\\"),NULL,NULL, SW_SHOWNORMAL );
//}
//
//
//void Cwhu_VoiceCardDlg::OnMenuFaxRecvFolder()
//{
//	// TODO: Add your command handler code here
//	ShellExecute(NULL, "open", _T("C:\\Whu_Data\\����\\���յĴ���\\"),NULL,NULL, SW_SHOWNORMAL );
//}
//
//
//void Cwhu_VoiceCardDlg::OnMenuFaxSendFolder()
//{
//	// TODO: Add your command handler code here
//	ShellExecute(NULL, "open", _T("C:\\Whu_Data\\����\\���͵Ĵ���\\"),NULL,NULL, SW_SHOWNORMAL );
//}


void Cwhu_VoiceCardDlg::OnMenuRecordFolder()
{
	// TODO: Add your command handler code here
	ShellExecute(NULL, "open", _T("C:\\Whu_Data\\¼��\\"),NULL,NULL, SW_SHOWNORMAL );
}


void Cwhu_VoiceCardDlg::OnMenuNoticeFolder()
{
	// TODO: Add your command handler code here
	ShellExecute(NULL, "open", _T("C:\\Whu_Data\\����֪ͨ\\"),NULL,NULL, SW_SHOWNORMAL );
}




void Cwhu_VoiceCardDlg::OnRButtonDown(UINT nFlags, CPoint point)
{
	// TODO: Add your message handler code here and/or call default
	if (whu_LogIn == true)
	{
		CMenu menu; 
		VERIFY( menu.LoadMenu(IDR_MENUREFRESH) ); 
		CMenu* popup = menu.GetSubMenu(0); 
		ASSERT( popup != NULL ); 
		//CString sss;
		//popup->GetMenuStringA(1,sss,MF_BYPOSITION); //�õ��˵�ѡ�����Ϣ
		popup->TrackPopupMenu(TPM_LEFTALIGN | TPM_RIGHTBUTTON,point.x, point.y, this ); 
		CDialogEx::OnRButtonDown(nFlags, point);
	}

}


void Cwhu_VoiceCardDlg::OnMenuOpenScaner()
{
	// TODO: Add your command handler code here
	ShellExecute(NULL, "open", _T("C:\\Program Files\\Canon Electronics\\DRM140\\TouchDR.exe"),NULL,NULL, SW_SHOWNORMAL );
}


void Cwhu_VoiceCardDlg::OnScanerFolder()
{
	// TODO: Add your command handler code here
	ShellExecute(NULL, "open", _T("C:\\Whu_Data\\ɨ��������ļ�\\"),NULL,NULL, SW_SHOWNORMAL );
}


void Cwhu_VoiceCardDlg::OnMenuFaxRecvFolder()
{
	// TODO: Add your command handler code here
	ShellExecute(NULL, "open", _T("C:\\Whu_Data\\����\\���յĴ���\\"),NULL,NULL, SW_SHOWNORMAL );
}


void Cwhu_VoiceCardDlg::OnMenuFaxSendFolder()
{
	// TODO: Add your command handler code here
	ShellExecute(NULL, "open", _T("C:\\Whu_Data\\����\\���͵Ĵ���\\"),NULL,NULL, SW_SHOWNORMAL );

}


void Cwhu_VoiceCardDlg::OnMenuShowDutyInfo()
{
	// TODO: Add your command handler code here
	//static CString m_strDuty = _T("��ʾֵ����Ϣ");
	////GetDlgItemText(IDC_SHOWDUTY,m_str);
	////CMenu menu; 
	////VERIFY( menu.LoadMenu(IDR_MENUREFRESH) ); 
	////CMenu* popup = menu.GetSubMenu(0); 
	////ASSERT( popup != NULL ); 
	////popup->GetMenuStringA(1,m_str,MF_BYPOSITION); //�õ��˵�ѡ�����Ϣ//
	//if (m_strDuty == _T("��ʾֵ����Ϣ"))
	//{
	//	//popup->ModifyMenuA(1,MF_BYPOSITION,0,_T("��ʾ��ϵ��"));
	//	//popup->GetMenuStringA(1,m_str,MF_BYPOSITION); //�õ��˵�ѡ�����Ϣ//

	//	m_BShowDutyList = true;
	//	//SetDlgItemText(IDC_SHOWDUTY,_T("��ʾ��ϵ��"));
	//	m_strDuty = _T("��ʾ��ϵ��");
	//	while (m_MyContactList.GetHeaderCtrl()->GetItemCount()>0)
	//	{
	//		m_MyContactList.DeleteColumn(0);
	//	}
	//	m_MyContactList.DeleteAllItems();
	//	whu_LinkToDutySql();
	//	GetDlgItem(IDC_SHAIXUAN)->ShowWindow(false);
	//	GetDlgItem(IDC_CHECK001)->ShowWindow(false);
	//	GetDlgItem(IDC_CHECK002)->ShowWindow(false);
	//	GetDlgItem(IDC_CHECK003)->ShowWindow(false);
	//	GetDlgItem(IDC_CHECK004)->ShowWindow(false);
	//	GetDlgItem(IDC_CHECK005)->ShowWindow(false);
	//	GetDlgItem(IDC_CHECK006)->ShowWindow(false);
	//	GetDlgItem(IDC_CHECK007)->ShowWindow(false);
	//	GetDlgItem(IDC_CHECK008)->ShowWindow(false);
	//	GetDlgItem(IDC_SEARCH)->ShowWindow(false);

	//}
	//else{
	//	//popup->ModifyMenuA(1,MF_BYPOSITION,0,_T("��ʾֵ����Ϣ"));
	//	//popup->GetMenuStringA(1,m_str,MF_BYPOSITION); //�õ��˵�ѡ�����Ϣ//
	//	m_BShowDutyList = false;
	//	//SetDlgItemText(IDC_SHOWDUTY,_T("��ʾֵ����Ϣ"));
	//	m_strDuty = _T("��ʾֵ����Ϣ");
	//	while (m_MyContactList.GetHeaderCtrl()->GetItemCount()>0)
	//	{
	//		m_MyContactList.DeleteColumn(0);
	//	}
	//	m_MyContactList.DeleteAllItems();
	//	whu_LinkToSQL();
	//	GetDlgItem(IDC_SHAIXUAN)->ShowWindow(true);
	//	GetDlgItem(IDC_CHECK001)->ShowWindow(true);
	//	GetDlgItem(IDC_CHECK002)->ShowWindow(true);
	//	GetDlgItem(IDC_CHECK003)->ShowWindow(true);
	//	GetDlgItem(IDC_CHECK004)->ShowWindow(true);
	//	GetDlgItem(IDC_CHECK005)->ShowWindow(true);
	//	GetDlgItem(IDC_CHECK006)->ShowWindow(true);
	//	GetDlgItem(IDC_CHECK007)->ShowWindow(true);
	//	GetDlgItem(IDC_CHECK008)->ShowWindow(true);
	//	GetDlgItem(IDC_SEARCH)->ShowWindow(true);

	//}
}


void Cwhu_VoiceCardDlg::OnLvnColumnclickMycontactlist(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMLISTVIEW pNMLV = reinterpret_cast<LPNMLISTVIEW>(pNMHDR);
	// TODO: Add your control notification handler code here
	int colum = pNMLV->iSubItem;
	static int mode = 0;
	whu_SortContactList(colum,mode);
	if (mode == 0)
	{
		mode = 1;
	}
	else{
		mode = 0;
	}
	*pResult = 0;
}
void Cwhu_VoiceCardDlg::whu_SortContactList(int colum,int mode)
{
	//���ַ���������//
	CStringArray m_ArrData[20];
	for (int i=0;i<m_MyContactList.GetItemCount();i++)
	{
		for (int j=0;j<m_MyContactList.GetHeaderCtrl()->GetItemCount();j++)
		{
			m_ArrData[j].Add(m_MyContactList.GetItemText(i,j));
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
					for (int j=1;j<m_MyContactList.GetHeaderCtrl()->GetItemCount();j++)
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
					for (int j=1;j<m_MyContactList.GetHeaderCtrl()->GetItemCount();j++)
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
	m_MyContactList.DeleteAllItems();
	for (int i=0;i<m_ArrData[0].GetSize();i++)
	{
		char buf[3];
		itoa(i+1,buf,10);
		m_MyContactList.InsertItem(i,buf);
		for (int j =0;j<m_MyContactList.GetHeaderCtrl()->GetItemCount();j++)
		{
			m_MyContactList.SetItemText(i,j,m_ArrData[j].GetAt(i)); 
		}
	}
	
}


void Cwhu_VoiceCardDlg::OnBnClickedButtonshowlog()
{
	// TODO: Add your control notification handler code here
	//��ģ̬�Ի���//
	//Cwhu_HistoryDlg *m_DlgLog = new Cwhu_HistoryDlg();
	//m_DlgLog->whu_ParentDlg = this;
	//m_DlgLog->Create(IDD_HISTORY,this);
	//m_DlgLog->ShowWindow(SW_SHOW);
	//ģ̬�Ի���//
	Cwhu_HistoryDlg m_DlgLog;
	m_DlgLog.whu_ParentDlg = this;
	m_DlgLog.DoModal();
}


void Cwhu_VoiceCardDlg::OnBnClickedButtonfax()
{
	// TODO: Add your control notification handler code here
	//ģ̬�Ի���//
	whu_FaxSendState = FAX_NONE;
	m_DlgFax = new Cwhu_FaxDlg();
	m_DlgFax->whu_ParentDlg = this;
	m_DlgFax->Create(IDD_FAX,this);
	m_DlgFax->ShowWindow(SW_SHOW);

	//CWnd* pWnd = FindWindow(NULL, _T("����"));
	//if(NULL == pWnd)
	//{
	//	whu_FaxSendState = FAX_NONE;
	//	m_DlgFax = new Cwhu_FaxDlg();
	//	m_DlgFax->whu_ParentDlg = this;
	//	m_DlgFax->Create(IDD_FAX,this);
	//	m_DlgFax->ShowWindow(SW_SHOW);
	//}
	//else
	//{
	//	pWnd->BringWindowToTop();
	//}	

}


void Cwhu_VoiceCardDlg::OnBnClickedButtonupdate()
{
	// TODO: Add your control notification handler code here
	AfxMessageBox(_T("��ǰ�汾Ϊ���£�"));
}

void Cwhu_VoiceCardDlg::OnBnClickedButtoncontact()
{
	// TODO: Add your control notification handler code here
	//Cwhu_ContractDlg *m_DlgCon = new Cwhu_ContractDlg();
	//m_DlgCon->whu_ParentDlg = this;
	//m_DlgCon->Create(IDD_CONTRACT,this);
	//m_DlgCon->ShowWindow(SW_SHOW);

	Cwhu_ContractDlg m_DlgCon;
	m_DlgCon.whu_ParentDlg = this;
	m_DlgCon.DoModal();

}


void Cwhu_VoiceCardDlg::OnBnClickedCheckrecord()
{
	// TODO: Add your control notification handler code here
	int AtrkCh =2;
	int ch = 0;
	CButton* pBtn = (CButton*)GetDlgItem(IDC_CHECKRECORD);
	int state = pBtn->GetCheck();
	if (state==0) //��ѡ��û��ѡ��//
	{
		SsmStopRecToFile(ch);
		if (whu_findfile(whu_RecordFile,_T("C:\\Whu_Data\\¼��")))
		{
			CStdioFile m_file(whu_RecordFile,CFile::modeRead);
			m_file.Close();
			m_file.Remove(whu_RecordFile);
		}
		GetLocalTime(&whu_SystemTime);
		m_MyLogList.SetItemText(m_MyLogList.GetItemCount()-1,7,_T(""));
		whu_SaveLog();
	}
	else{
		GetLocalTime(&whu_SystemTime);
		whu_RecordFile.Format(_T("C:\\Whu_Data\\¼��\\%d��%d��%d��%dʱ%d��_����%s.wav\0"),whu_SystemTime.wYear,whu_SystemTime.wMonth,whu_SystemTime.wDay,whu_SystemTime.wHour,whu_SystemTime.wMinute,whu_NumberBuf);
		SsmSetRecMixer(AtrkCh,TRUE,0);
		SsmRecToFile(ch,whu_RecordFile,-1,1,0xffffffff,0xffffffff,1); //��������¼�������ɹ�����//
		m_MyLogList.SetItemText(m_MyLogList.GetItemCount()-1,7,whu_RecordFile);
		whu_SaveLog();
	}
}
CString Cwhu_VoiceCardDlg::whu_GetContactPic(CString Name)
{
	//�õ���ϵ�˵�ͷ��//
	CString m_folder = _T("C:\\Whu_Data\\��ϵ��ͷ��");
	CString m_Pic = m_folder + _T("\\") + Name + _T(".jpg");
	if(whu_findfile(m_Pic,m_folder) != true)
	{
		return _T("C:\\Whu_Data\\��ϵ��ͷ��\\İ����.jpg");
	}else{
		return m_Pic;
	}
}
void Cwhu_VoiceCardDlg::Lin_InportExcelToList(CString m_FilePath,CListCtrl &m_List)
{
	//��ɾ���б�����//
	m_List.DeleteAllItems();
	while(m_List.GetHeaderCtrl()->GetItemCount()>0)
	{
		m_List.DeleteColumn(0);
	}
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
		return;
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
	CStringArray m_HeadName;
	for (int i=1;i<26;i++)
	{
		CString m_pos = Lin_GetEnglishCharacter(i);
		m_pos = m_pos + _T("1");
		range = sheet.get_Range(COleVariant(m_pos),COleVariant(m_pos));
		//��õ�Ԫ�������
		COleVariant rValue;
		rValue = COleVariant(range.get_Value2());
		//ת���ɿ��ַ�//
		rValue.ChangeType(VT_BSTR);
		//ת����ʽ�������//
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
			//��õ�Ԫ�������
			COleVariant rValue;
			rValue = COleVariant(range.get_Value2());
			//ת���ɿ��ַ�//
			rValue.ChangeType(VT_BSTR);
			//ת����ʽ�������//
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
void Cwhu_VoiceCardDlg::Lin_ExportListToExcel(CListCtrl &m_List,CString m_FilePath)
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
			MessageBox(_T("�޷�����ExcelӦ�ã�"));
			return;
		}
		books=app.get_Workbooks();
		book=books.Add(covOptional);
		sheets=book.get_Sheets();
		sheet=sheets.get_Item(COleVariant((short)1));

		//д���ͷ//
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
		//��ȡ��Ԫ���λ��//
		for (int ColumNum =1;ColumNum<m_List.GetHeaderCtrl()->GetItemCount();ColumNum++)
		{

			for (int ItemNum = 0;ItemNum<m_List.GetItemCount();ItemNum++)
			{
				str2 = Lin_GetEnglishCharacter(ColumNum);
				str1.Format(_T("%d"),ItemNum+2);
				str2 = str2+str1;
				CString str = m_List.GetItemText(ItemNum,ColumNum);
				str = _T("'")+str;
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
		AfxMessageBox(_T("�б�Ϊ��"));
	}

}
void Cwhu_VoiceCardDlg::Lin_InitList(CListCtrl &m_List,CStringArray &ColumName)
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
	m_List.InsertColumn(0,_T("���"),LVCFMT_LEFT,60);
	for (int i=0;i<ColumName.GetSize();i++)
	{
		m_List.InsertColumn(i+1,ColumName.GetAt(i),LVCFMT_LEFT,160);
	}

}
void Cwhu_VoiceCardDlg::Lin_InsertList(CListCtrl &m_List,CStringArray &ItemArr)
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
int Cwhu_VoiceCardDlg::Lin_ImportExcelToArr(CString m_FilePath,CStringArray *m_DutyArr)//����CString����ĸ���
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
		return -1;
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
	CStringArray m_HeadName;
	for (int i=1;i<26;i++)
	{
		CString m_pos = Lin_GetEnglishCharacter(i);
		m_pos = m_pos + _T("1");
		range = sheet.get_Range(COleVariant(m_pos),COleVariant(m_pos));
		//��õ�Ԫ�������
		COleVariant rValue;
		rValue = COleVariant(range.get_Value2());
		//ת���ɿ��ַ�//
		rValue.ChangeType(VT_BSTR);
		//ת����ʽ�������//
		CString m_content = CString(rValue.bstrVal);
		if (m_content!=_T(""))
		{
			m_HeadName.Add(m_content);
		}
	}
	int test = m_HeadName.GetSize();
	for (int i=0;i<m_HeadName.GetSize();i++)
	{
		m_DutyArr[i].RemoveAll();
		m_DutyArr[i].Add(m_HeadName.GetAt(i));
	}
	//�Դˣ��õ�excel�ĵ�����ͷ��//
	CString m_str = m_DutyArr[0].GetAt(0);
	CStringArray m_ContentArr;
	Number = m_HeadName.GetSize();
	for (int ColumNum = 0;ColumNum<Number;ColumNum++)
	{
		CString m_pos = Lin_GetEnglishCharacter(ColumNum+1);
		for (int ItemNum=0;ItemNum<1024;ItemNum++) //��Ǳ�ڵ�bug//
		{
			CString m_Itempos;
			m_Itempos.Format(_T("%d"),ItemNum+2);
			CString m_str = m_pos+m_Itempos;
			range = sheet.get_Range(COleVariant(m_str),COleVariant(m_str));
			//��õ�Ԫ�������
			COleVariant rValue;
			rValue = COleVariant(range.get_Value2());
			//ת���ɿ��ַ�//
			rValue.ChangeType(VT_BSTR);
			//ת����ʽ�������//
			CString m_content = CString(rValue.bstrVal);
			m_DutyArr[ColumNum].Add(m_content);
		}
	}
	int MaxNum = m_DutyArr[0].GetSize();
	for (int i=0;i<m_DutyArr[0].GetSize();i++)
	{
		if (m_DutyArr[0].GetAt(i) != _T(""))
		{
			MaxNum = i+1;
		}
	}

	for (int j=0;j<Number;j++)
	{
		m_DutyArr[j].RemoveAt(MaxNum,m_DutyArr[j].GetSize() - MaxNum);
	}
	for (int j=0;j<Number;j++)
	{
		int test3333 = m_DutyArr[j].GetSize();
		int sss = 0;
	}
	book.put_Saved(TRUE);
	app.Quit();
	return Number;
}
void Cwhu_VoiceCardDlg::Lin_ImportArrToList(CStringArray *m_DutyArr,int Num,CListCtrl &m_List)
{
	////��ɾ���б�����//
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
CString Cwhu_VoiceCardDlg::Lin_GetEnglishCharacter(int Num)
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
void Cwhu_VoiceCardDlg::Lin_ImpoerArrToExcel(CStringArray *m_DutyArr,int Num,CString m_FilePath)
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
		MessageBox(_T("�޷�����ExcelӦ�ã�"));
		return;
	}
	books=app.get_Workbooks();
	book=books.Add(covOptional);
	sheets=book.get_Sheets();
	sheet=sheets.get_Item(COleVariant((short)1));

	for (int ColumNum =0;ColumNum<Num;ColumNum++)
	{

		for (int ItemNum = 0;ItemNum<m_DutyArr[0].GetSize();ItemNum++)
		{
			str2 = Lin_GetEnglishCharacter(ColumNum+1);
			str1.Format(_T("%d"),ItemNum+1);
			str2 = str2+str1;
			CString str = m_DutyArr[ColumNum].GetAt(ItemNum);
			str = _T("'")+str;
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

}
int Cwhu_VoiceCardDlg::Lin_ListToArr(CListCtrl &m_List,CStringArray *m_DutyArr)
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
void Cwhu_VoiceCardDlg::Lin_ImportAutoForArr(CString m_FilePath,CStringArray &m_AutoForward)
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
void Cwhu_VoiceCardDlg::Lin_ExportForArrToExcel(CStringArray &m_AutoForward,CString m_FilePath)
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
int Cwhu_VoiceCardDlg::whu_GetAutoForwardFaxNum(CString m_FaxNum,CStringArray &whu_SourceArr,CStringArray &m_FaxNumArr)
{
	int size = whu_SourceArr.GetSize();
	int NumCount = 0;
	for (int i=0;i<size;i++)
	{
		CString m_str = whu_SourceArr.GetAt(i);
		if ((m_str == m_FaxNum)&&(i < size-1))
		{
			CString m_str2 = whu_SourceArr.GetAt(i+1);
			int CharCount = m_str2.GetLength(); 
			if(CharCount<4)
			{
				NumCount = _ttoi(m_str2);
				for (int j=i+2;j<i+NumCount+2;j++)
				{
					 m_FaxNumArr.Add(whu_SourceArr.GetAt(j));
				}
				break;
			}

		}	
	}
	return NumCount;
}

void Cwhu_VoiceCardDlg::OnNMCustomdrawMycontactlist(NMHDR *pNMHDR, LRESULT *pResult)
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
				clrNewBkColor = RGB( 255, 255, 0); //ż���б���ɫΪ��ɫ
			}
			else
			{
				clrNewBkColor = RGB( 255, 0, 0 ); //�����б���ɫΪ��ɫ
			}

			lplvcd->clrText = clrNewTextColor;
			lplvcd->clrTextBk = clrNewBkColor;

			*pResult = CDRF_DODEFAULT;  
			return; 
		} 
	} 

	*pResult = 0;
}



void Cwhu_VoiceCardDlg::OnBnClickedUser()
{
	// TODO: Add your control notification handler code here
	if (m_LogName == _T("����Ա"))
	{
		Cwhu_UserDlg m_UserDlg;
		m_UserDlg.whu_ParentDlg = this;
		m_UserDlg.DoModal();
	}else{
		CWhu_UserChgPwd m_Dlg;
		m_Dlg.whu_ParentDlg = this;
		m_Dlg.DoModal();
	}

}


void Cwhu_VoiceCardDlg::OnBnClickedRegister()
{
	// TODO: Add your control notification handler code here
	CRegisterDlg m_Dlg;
	m_Dlg.whu_ParentDlg = this;
	m_Dlg.DoModal();
}
bool Cwhu_VoiceCardDlg::whu_ImportContactData(CString m_File)
{
	CStringArray m_StandarHeadName;
	m_StandarHeadName.Add(_T("����"));
	m_StandarHeadName.Add(_T("��λ"));
	m_StandarHeadName.Add(_T("����"));
	m_StandarHeadName.Add(_T("ְ��"));
	m_StandarHeadName.Add(_T("�ֻ�"));
	m_StandarHeadName.Add(_T("�칫�绰"));
	m_StandarHeadName.Add(_T("�����"));
	m_StandarHeadName.Add(_T("����"));
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
		return false;
	}
	books = app.get_Workbooks();
	//��Excel������m_FilePathΪExcel���·����//
	lpDisp = books.Open(m_File,covOptional
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
		//��õ�Ԫ�������
		COleVariant rValue;
		rValue = COleVariant(range.get_Value2());
		//ת���ɿ��ַ�//
		rValue.ChangeType(VT_BSTR);
		//ת����ʽ�������//
		CString m_content = CString(rValue.bstrVal);
		if (m_content!=_T("") )
		{
			if (m_content != m_StandarHeadName.GetAt(i-1))
			{
				AfxMessageBox(_T("ͨѶ¼������ģ��Ҫ��"));
				book.put_Saved(TRUE);
				app.Quit();
				return false;//˵��������ģ��Ҫ��//
			}else{
				m_HeadName.Add(m_content);
			}	
		}
	}
	//��ɾ���б�����//   
	m_MyContactList.DeleteAllItems();
	while(m_MyContactList.GetHeaderCtrl()->GetItemCount()>0)
	{
		m_MyContactList.DeleteColumn(0);
	}
	Lin_InitList(m_MyContactList,m_HeadName);
	for (int i=0;i<m_HeadName.GetSize();i++)
	{
		m_ContactListData[i].RemoveAll();
		m_ContactListData[i].Add(m_HeadName.GetAt(i));
	}
	CStringArray m_ContentArr;
	for (int ItemNum = 0;ItemNum<1024;ItemNum++)
	{
		for (int j=1;j<m_MyContactList.GetHeaderCtrl()->GetItemCount();j++)
		{
			CString m_pos = Lin_GetEnglishCharacter(j);
			CString m_Itempos;
			m_Itempos.Format(_T("%d"),ItemNum+2);
			CString m_str = m_pos+m_Itempos;
			range = sheet.get_Range(COleVariant(m_str),COleVariant(m_str));
			//��õ�Ԫ�������
			COleVariant rValue;
			rValue = COleVariant(range.get_Value2());
			//ת���ɿ��ַ�//
			rValue.ChangeType(VT_BSTR);
			//ת����ʽ�������//
			CString m_content = CString(rValue.bstrVal);
			m_ContentArr.Add(m_content);
		}
		if (m_ContentArr.GetAt(0)!=_T(""))
		{
			Lin_InsertList(m_MyContactList,m_ContentArr);
			m_ContentArr.RemoveAll();
		}
		else{
			break;
		}
	}
	book.put_Saved(TRUE);
	app.Quit();
	ConHeadNum = m_MyContactList.GetHeaderCtrl()->GetItemCount()-1;
	for (int ColumNum=1;ColumNum<m_MyContactList.GetHeaderCtrl()->GetItemCount();ColumNum++)
	{
		for (int ItemNum=0;ItemNum<m_MyContactList.GetItemCount();ItemNum++)
		{
			CString m_str = m_MyContactList.GetItemText(ItemNum,ColumNum);
			m_ContactListData[ColumNum-1].Add(m_str);
		}
	}
	return true;
}

void Cwhu_VoiceCardDlg::OnBnClickedCheckautoforwardfax()
{
	// TODO: Add your control notification handler code here
	UpdateData(true);
}
