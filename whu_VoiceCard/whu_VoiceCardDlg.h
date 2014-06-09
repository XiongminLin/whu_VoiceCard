
// whu_VoiceCardDlg.h : header file
//

#pragma once
#include "afxcmn.h"
#include "Cwhu_ContractDlg.h"
#include "Cwhu_DocuDlg.h"
#include "Cwhu_FaxDlg.h"
#include "Cwhu_PhoneDlg.h"
#include "Cwhu_OnDutyDlg.h"
#include "Cwhu_UserDlg.h"
#include "Cwhu_HistoryDlg.h"
#include "DetailLogDlg.h"
#include "Cwhu_RecvCallDlg.h"
#include "Cwhu_NoticeDlg.h"
#include "SaveEveryingDlg.h"
#include "Shpa3api.h"
#include "ConvertAgent.h"
#include "PrinterSettings.h"
#include "Cwhu_GroupSendThread.h"
#include "SkinPPWTL.h"
#include "afxwin.h"
#include "shockwaveflash1.h"
#include <cv.h>
#include <highgui.h>
#include "CvvImage.h"
#include "SendMail.h"
#include  "HeaderCtrlCl.h "
#include  "ListCtrlCl.h "
#include "atlcomtime.h"
#include "afxdtctl.h"
#include "BmpApi.h"
#include "HzToPy.h"
#include "Whu_UserChgPwd.h"
#include "RegisterDlg.h"
#define  WM_UPATEDATA WM_USER+20
extern int whu_NumHaveSend;
extern CString whu_NumberBuf;
extern CString whu_AllNos[MAX_PATH];
extern CStringArray whu_FaxAutoForwardArr;
enum APP_USER_STATE {
	USER_IDLE,
	USER_GET_PHONE_NUM,
	USER_WAIT_DIAL_TONE,
	USER_WAIT_REMOTE_PICKUP,
	USER_TALKING,
	USER_WAIT_HANGUP,
	USER_WAIT_LOCALPICKUP,
	USER_RECORDING,
	TRK_IDLE,
	TRK_BUSY,
	TRK_AutoDial,
	TRK_TALKING,
	TRK_WAITREMOTECHOOOSE,
	TRK_NOTICEPLAYING,
	TRK_FAXING,					//faxing...
	TRK_DIALING,
	FAX_IDLE,					//idle
	FAX_CHECK_END,        //faxing...
	TASK_PHONEDIAL,
	TASK_PHONENOTICE,
	TASK_FAXDIAL,
	TASK_FAXRECV,
	TASK_PHONERECV,
	NOTICE_PREPARE,
	NOTICE_PLAYING,
	NOTICE_ATTEND,
	NOTICE_NOATTEND,
	NOTICE_HANGUP,
	NOTICE_NONE,
	NOTICE_BUSY,
	FAX_PREPARE,
	FAX_SENDING,
	FAX_FAIL,
	FAX_SUCCESS,
	FAX_CANCEL,
	FAX_NOANSWER,
	FAX_RECVING,
	FAX_NONE
};

typedef struct{
	CString Name;
	CString Unit;
	CString Department;
	CString Job;
	CString TelNum;
	CString OfficeNum;
	CString FaxNum;
	CString E_Mail;
	CString PictureFile;
	CString DutyPerson;
	CString task;
	CString mark;
	bool TaskSuccess;
	int Priority; //值越大，优先级越高//
	CString ID;

}whu_ContractData;


// Cwhu_VoiceCardDlg dialog
class Cwhu_VoiceCardDlg : public CDialogEx
{
// Construction
public:
	Cwhu_VoiceCardDlg(CWnd* pParent = NULL);	// standard constructor

// Dialog Data
	enum { IDD = IDD_WHU_VOICECARD_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support

// Implementation
protected:
	HICON m_hIcon;

	// Generated message map functions
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
	//CTabCtrl whu_Tab;
	//Cwhu_UserDlg whu_UserDlg;
	//Cwhu_OnDutyDlg whu_OndutyDlg;
	//Cwhu_ContractDlg whu_ContractDlg;
	//Cwhu_FaxDlg whu_FaxDlg;
	//Cwhu_PhoneDlg whu_PhoneDlg;
	//Cwhu_DocuDlg whu_DocumentDlg;
	//Cwhu_HistoryDlg whu_HistoryDlg;
	//afx_msg void OnSelchangeTab(NMHDR *pNMHDR, LRESULT *pResult);
	CFont m_editFont; 
public:
	int  SeparateCallNum(CString m_AllNum,CString *m_SeparateNum);
	//int whu_NumHaveSend;
	//CString whu_NumberBuf;
	//CString whu_AllNos[MAX_PATH];
	int whu_NumOfNum;
	int whu_NumLength;
	POINT Old;
	void resize();
private:
	void		TrunkProc(UINT event, WPARAM wParam, LPARAM lParam); //lin:中继通道回调函数
	void		FaxProc(UINT event, WPARAM wParam, LPARAM lParam);	//lin:传真回调函数
	void        UserProc(UINT event, WPARAM wParam, LPARAM lParam);
	virtual LRESULT WindowProc(UINT message, WPARAM wParam, LPARAM lParam);

	Cwhu_GroupSendThread *whu_GroupSendThread;

	bool whu_InitCtiBoard();

	
public:
	SYSTEMTIME whu_SystemTime;
	CString whu_Time;
	int whu_FaxChState;
	int whu_TrunkChState;
	int whu_UserChState;
	int whu_NoticeState;
	int whu_FaxSendState;
	int whu_FaxRecvState;
	bool whu_RecordNotice;
	CString whu_RecordFile;
	CString whu_NoticeFile; 
	CString whu_NoticeFinalFile;//播放录音后会拨一段提示按*重听按#确认//

	CString whu_SendFaxFile;
	CString whu_RecvFaxFile;
	CString whu_SendFaxFileTitle;
	int whu_TaskMode;
	IplImage *whu_CallImg;
	IplImage *whu_WelcomePic;
	IplImage *whu_PicDown;
	bool m_BShowCallImg;
	bool whu_LogIn;
	bool whu_RecordNoticeDone;
public:
	virtual BOOL DestroyWindow();

	void whu_SendNotice(CString *m_CallNum,int NumOfNum);
	void whu_SendFax(CString *m_CallNum,CString m_SendFilename,int num);
	void whu_Dial(CString m_Num);
	//CShockwaveflash whu_flash;
	void whu_ShowCallInInfo(bool m_b);
	void whu_ShowCallOutInfo(bool m_b);
	void whu_ShowFaxInInfo(bool show);
	void whu_ShowFaxOutInfo(bool show);
	void whu_ShowRecordInfo(bool show);
	int DocumentToTiff(CString m_Inputfile,CString &m_OutputFile);
	afx_msg void OnBnClickedRefuse();
	bool whu_AcceptCall;
	bool whu_AcceptFax;
	BOOL whu_BRecord;
	void whu_ShowNumInfo(CString m_Num);
	whu_ContractData whu_GetContractData(CString m_num);
	whu_ContractData m_MyContactData;
	CString whu_GetContactPic(CString Name);
	void DrawPicToHDC(IplImage *img, UINT ID );
	//void whu_TabInit();
	afx_msg void OnBnClickedAcceptfaxorrecord();
	afx_msg void OnBnClickedLogin();
	//CComboBoxEx *whu_LogInComBox;
	//CComboBoxEx whu_LogInComBox;
	bool whu_CheckLogIn(CString name,CString password);
	CString whu_ShowLeader(CString m_DutyPerson);
	CString whu_password;
	afx_msg void OnBnClickedBrecordmynotice();
	afx_msg void OnBnClickedTrylisten();
	CListCtrl m_MyContactList;
	CListCtrl m_MyLogList;
	//CListCtrlCl m_MyLogList;
	//CListCtrlCl m_MyContactList;
	void whu_LinkToSQL();
	CString GetStringFromVariant(_variant_t var);
	afx_msg void OnNMRClickMycontactlist(NMHDR *pNMHDR, LRESULT *pResult);
	
	afx_msg void OnTelPhoneDial();
	afx_msg void OnTelPhoneSendSingleNotice();
	afx_msg void OnOfficePhoneDial();
	afx_msg void OnOfficePhoneSendSingleNotice();
	afx_msg void OnSendGroupNotice();
	afx_msg void OnSendGroupFax();
	afx_msg void OnSendSingleFax();
	afx_msg void OnBnClickedOpenscanner();

	CStringArray m_LogListData[10];
	void whu_SaveLog();
	void whu_ListToTxt(CListCtrl *m_list,CString m_txtfile);
	void whu_TxtToList(CString m_txtfile,CListCtrl *m_list);
	void whu_SearchLog(CString m_str);
	void whu_InitLogList();
	void whu_AddLog(CString time,CString duty,CString taskmode,CString unit,CString name,CString number,CString file,CString state,CString beizhu);
	void whu_ModifyCheck(CString m_str);
	afx_msg void OnSeeLog();
	afx_msg void OnNMRClickMyloglist(NMHDR *pNMHDR, LRESULT *pResult);
	int m_ClickLogPos;
	//CDetailLogDlg m_DetailLogDlg;
	BOOL m_bAutoRecvFax;
	BOOL m_checkname001;
	BOOL m_checkunite002;
	BOOL m_checkdepartment003;
	BOOL m_checkjob004;
	BOOL m_checktel005;
	BOOL m_checkOffNum006;
	BOOL m_checkFax007;
	BOOL m_checkemail008;
	afx_msg void OnBnClickedCheck001();
	afx_msg void OnBnClickedCheck002();
	afx_msg void OnBnClickedCheck003();
	afx_msg void OnBnClickedCheck004();
	afx_msg void OnBnClickedCheck005();
	afx_msg void OnBnClickedCheck006();
	afx_msg void OnBnClickedCheck007();
	afx_msg void OnBnClickedCheck008();
	//afx_msg void OnAddLogRemark();
	//CString m_LogRemark;
	//afx_msg void OnBnClickedLogsave();
	//afx_msg void OnBnClickedLogcancel();
	afx_msg void OnOpenLogFile();
	afx_msg void OnOpenFilePath();
	afx_msg void OnDeleteLog();
	afx_msg void OnBnClickedLogout();
	CString m_LogName;
	long whu_FaxAllBites;
	long whu_FaxDoneBites;
	long m_AllFaxPage;
	CProgressCtrl m_ProgressCtrl;
	afx_msg void OnTimer(UINT_PTR nIDEvent);
	CString m_ContactSearch;
	CStringArray m_ContactListData[11];
	void whu_SearchContact(CString m_str);
	afx_msg void OnBnClickedQuitsystem();
	sMailInfo whu_MailData;
	CSendMail whu_MyMail;
	afx_msg void OnSendMail();
	afx_msg void OnSeeTelLog();
	afx_msg void OnRefreshLog();
	afx_msg void OnRefreshLog2();
	afx_msg void OnEditCopy();
	afx_msg void OnSeeFaxLog();
	afx_msg void OnSeeOfficeNumLog();
	CStringArray whu_DutyPersons[20];
	void whu_LinkToDutySql();
	CComboBox m_ComboBoxLeader;
	CString m_ComboxShowLeader;
	//afx_msg void OnBnClickedShowduty();
	bool m_BShowDutyList;
	CComboBox m_LogInComBox;
	void whu_LinkToUserSql();
	CStringArray whu_User[4];
	
	afx_msg void OnDtnDatetimechangeDatetimepicker1(NMHDR *pNMHDR, LRESULT *pResult);
	CDateTimeCtrl m_DataCtrl;
	bool whu_findfile(CString m_file,CString folder);
	void whu_logchangetime();
	void whu_ResetListNum(CListCtrl *m_MyList);
	bool m_BPhoneDialDone;
	bool m_BPhoneButtonPress;

	afx_msg void OnSaveLogToFile();
	CString m_LogDeleteStr;
	bool m_BFlashPlaying;
	afx_msg void OnMenuSeeHistory();
	//afx_msg void OnOpenPDFFolder();
	//afx_msg void OnMenuFaxRecvFolder();
	//afx_msg void OnMenuFaxSendFolder();
	afx_msg void OnMenuRecordFolder();
	afx_msg void OnMenuNoticeFolder();
	afx_msg void OnRButtonDown(UINT nFlags, CPoint point);
	afx_msg void OnMenuOpenScaner();
	afx_msg void OnScanerFolder();
	afx_msg void OnMenuFaxRecvFolder();
	afx_msg void OnMenuFaxSendFolder();
	afx_msg void OnMenuShowDutyInfo();
	CRect m_LogRect;
	afx_msg void OnLvnColumnclickMycontactlist(NMHDR *pNMHDR, LRESULT *pResult);
	void whu_SortContactList(int colum,int mode);
	afx_msg void OnBnClickedShowduty();
	afx_msg void OnBnClickedButtonshowlog();
	afx_msg void OnBnClickedButtonfax();
	afx_msg void OnBnClickedButtonupdate();
	afx_msg void OnBnClickedButtoncontact();
	afx_msg void OnBnClickedCheckrecord();
	CString Lin_GetEnglishCharacter(int Num);
	void Lin_InportExcelToList(CString m_FilePath,CListCtrl &m_List);
	void Lin_ExportListToExcel(CListCtrl &m_List,CString m_FilePath);
	void Lin_InitList(CListCtrl &m_List,CStringArray &ColumName);
	void Lin_InsertList(CListCtrl &m_List,CStringArray &ItemArr);
	int Lin_ImportExcelToArr(CString m_FilePath,CStringArray *m_DutyArr);//返回CString数组的个数
	void Lin_ImpoerArrToExcel(CStringArray *m_DutyArr,int Num,CString m_FilePath);
	int DutyHeadNum;
	int ConHeadNum;
	void Lin_ImportArrToList(CStringArray *m_DutyArr,int Num,CListCtrl &m_List);
	int Lin_ListToArr(CListCtrl &m_List,CStringArray *m_DutyArr);
	//CStringArray whu_AutoForward;
	void Lin_ImportAutoForArr(CString m_FilePath,CStringArray &m_AutoForward);
	void Lin_ExportForArrToExcel(CStringArray &m_AutoForward,CString m_FilePath);
	bool whu_ImportContactData(CString m_File);
	CHzToPy whu_HzTopy;
	int whu_GetAutoForwardFaxNum(CString m_FaxNum,CStringArray &whu_SourceArr,CStringArray &m_FaxNumArr);
	//bool whu_BAutoForward;
	afx_msg void OnNMCustomdrawMycontactlist(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnBnClickedUser();
	afx_msg void OnBnClickedRegister();

	Cwhu_FaxDlg *m_DlgFax;
	Cwhu_NoticeDlg *m_NoticeDlg;
	CSaveEveryingDlg *m_SaveEveryingDlg;
	afx_msg void OnBnClickedCheckautoforwardfax();
	BOOL whu_BAutoForward;
	bool m_BQuitSystem;
};
