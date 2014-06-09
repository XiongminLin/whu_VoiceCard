//// CAdodc.h  : Declaration of ActiveX Control wrapper class(es) created by Microsoft Visual C++
//
//#pragma once
//
///////////////////////////////////////////////////////////////////////////////
//// CAdodc
//
//class CAdodc : public CWnd
//{
//protected:
//	DECLARE_DYNCREATE(CAdodc)
//public:
//	CLSID const& GetClsid()
//	{
//		static CLSID const clsid
//			= { 0x67397AA3, 0x7FB1, 0x11D0, { 0xB1, 0x48, 0x0, 0xA0, 0xC9, 0x22, 0xE8, 0x20 } };
//		return clsid;
//	}
//	virtual BOOL Create(LPCTSTR lpszClassName, LPCTSTR lpszWindowName, DWORD dwStyle,
//						const RECT& rect, CWnd* pParentWnd, UINT nID, 
//						CCreateContext* pContext = NULL)
//	{ 
//		return CreateControl(GetClsid(), lpszWindowName, dwStyle, rect, pParentWnd, nID); 
//	}
//
//    BOOL Create(LPCTSTR lpszWindowName, DWORD dwStyle, const RECT& rect, CWnd* pParentWnd, 
//				UINT nID, CFile* pPersist = NULL, BOOL bStorage = FALSE,
//				BSTR bstrLicKey = NULL)
//	{ 
//		return CreateControl(GetClsid(), lpszWindowName, dwStyle, rect, pParentWnd, nID,
//		pPersist, bStorage, bstrLicKey); 
//	}
//
//// Attributes
//public:
//
//// Operations
//public:
//
//	CString get_ConnectionString()
//	{
//		CString result;
//		InvokeHelper(0x1, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
//		return result;
//	}
//	void put_ConnectionString(LPCTSTR newValue)
//	{
//		static BYTE parms[] = VTS_BSTR ;
//		InvokeHelper(0x1, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
//	}
//	CString get_OLEDBString()
//	{
//		CString result;
//		InvokeHelper(0x2, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
//		return result;
//	}
//	void put_OLEDBString(LPCTSTR newValue)
//	{
//		static BYTE parms[] = VTS_BSTR ;
//		InvokeHelper(0x2, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
//	}
//	CString get_OLEDBFile()
//	{
//		CString result;
//		InvokeHelper(0x3, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
//		return result;
//	}
//	void put_OLEDBFile(LPCTSTR newValue)
//	{
//		static BYTE parms[] = VTS_BSTR ;
//		InvokeHelper(0x3, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
//	}
//	CString get_DataSourceName()
//	{
//		CString result;
//		InvokeHelper(0x4, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
//		return result;
//	}
//	void put_DataSourceName(LPCTSTR newValue)
//	{
//		static BYTE parms[] = VTS_BSTR ;
//		InvokeHelper(0x4, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
//	}
//	CString get_OtherAttributes()
//	{
//		CString result;
//		InvokeHelper(0x5, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
//		return result;
//	}
//	void put_OtherAttributes(LPCTSTR newValue)
//	{
//		static BYTE parms[] = VTS_BSTR ;
//		InvokeHelper(0x5, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
//	}
//	long get_ConnectStringType()
//	{
//		long result;
//		InvokeHelper(0x6, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
//		return result;
//	}
//	void put_ConnectStringType(long newValue)
//	{
//		static BYTE parms[] = VTS_I4 ;
//		InvokeHelper(0x6, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
//	}
//	CString get_UserName()
//	{
//		CString result;
//		InvokeHelper(0x7, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
//		return result;
//	}
//	void put_UserName(LPCTSTR newValue)
//	{
//		static BYTE parms[] = VTS_BSTR ;
//		InvokeHelper(0x7, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
//	}
//	CString get_Password()
//	{
//		CString result;
//		InvokeHelper(0x8, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
//		return result;
//	}
//	void put_Password(LPCTSTR newValue)
//	{
//		static BYTE parms[] = VTS_BSTR ;
//		InvokeHelper(0x8, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
//	}
//	long get_Mode()
//	{
//		long result;
//		InvokeHelper(0x9, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
//		return result;
//	}
//	void put_Mode(long newValue)
//	{
//		static BYTE parms[] = VTS_I4 ;
//		InvokeHelper(0x9, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
//	}
//	long get_CursorLocation()
//	{
//		long result;
//		InvokeHelper(0xa, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
//		return result;
//	}
//	void put_CursorLocation(long newValue)
//	{
//		static BYTE parms[] = VTS_I4 ;
//		InvokeHelper(0xa, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
//	}
//	long get_IsolationLevel()
//	{
//		long result;
//		InvokeHelper(0xb, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
//		return result;
//	}
//	void put_IsolationLevel(long newValue)
//	{
//		static BYTE parms[] = VTS_I4 ;
//		InvokeHelper(0xb, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
//	}
//	long get_ConnectionTimeout()
//	{
//		long result;
//		InvokeHelper(0xc, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
//		return result;
//	}
//	void put_ConnectionTimeout(long newValue)
//	{
//		static BYTE parms[] = VTS_I4 ;
//		InvokeHelper(0xc, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
//	}
//	long get_CommandTimeout()
//	{
//		long result;
//		InvokeHelper(0xd, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
//		return result;
//	}
//	void put_CommandTimeout(long newValue)
//	{
//		static BYTE parms[] = VTS_I4 ;
//		InvokeHelper(0xd, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
//	}
//	CString get_RecordSource()
//	{
//		CString result;
//		InvokeHelper(0xe, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
//		return result;
//	}
//	void put_RecordSource(LPCTSTR newValue)
//	{
//		static BYTE parms[] = VTS_BSTR ;
//		InvokeHelper(0xe, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
//	}
//	long get_CursorType()
//	{
//		long result;
//		InvokeHelper(0xf, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
//		return result;
//	}
//	void put_CursorType(long newValue)
//	{
//		static BYTE parms[] = VTS_I4 ;
//		InvokeHelper(0xf, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
//	}
//	long get_LockType()
//	{
//		long result;
//		InvokeHelper(0x10, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
//		return result;
//	}
//	void put_LockType(long newValue)
//	{
//		static BYTE parms[] = VTS_I4 ;
//		InvokeHelper(0x10, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
//	}
//	long get_CommandType()
//	{
//		long result;
//		InvokeHelper(0x11, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
//		return result;
//	}
//	void put_CommandType(long newValue)
//	{
//		static BYTE parms[] = VTS_I4 ;
//		InvokeHelper(0x11, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
//	}
//	long get_CursorOptions()
//	{
//		long result;
//		InvokeHelper(0x12, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
//		return result;
//	}
//	long get_CacheSize()
//	{
//		long result;
//		InvokeHelper(0x13, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
//		return result;
//	}
//	void put_CacheSize(long newValue)
//	{
//		static BYTE parms[] = VTS_I4 ;
//		InvokeHelper(0x13, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
//	}
//	long get_MaxRecords()
//	{
//		long result;
//		InvokeHelper(0x14, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
//		return result;
//	}
//	void put_MaxRecords(long newValue)
//	{
//		static BYTE parms[] = VTS_I4 ;
//		InvokeHelper(0x14, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
//	}
//	long get_BOFAction()
//	{
//		long result;
//		InvokeHelper(0x15, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
//		return result;
//	}
//	void put_BOFAction(long newValue)
//	{
//		static BYTE parms[] = VTS_I4 ;
//		InvokeHelper(0x15, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
//	}
//	long get_EOFAction()
//	{
//		long result;
//		InvokeHelper(0x16, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
//		return result;
//	}
//	void put_EOFAction(long newValue)
//	{
//		static BYTE parms[] = VTS_I4 ;
//		InvokeHelper(0x16, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
//	}
//	CString get_Caption()
//	{
//		CString result;
//		InvokeHelper(DISPID_CAPTION, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
//		return result;
//	}
//	void put_Caption(LPCTSTR newValue)
//	{
//		static BYTE parms[] = VTS_BSTR ;
//		InvokeHelper(DISPID_CAPTION, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
//	}
//	long get_Appearance()
//	{
//		long result;
//		InvokeHelper(DISPID_APPEARANCE, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
//		return result;
//	}
//	void put_Appearance(long newValue)
//	{
//		static BYTE parms[] = VTS_I4 ;
//		InvokeHelper(DISPID_APPEARANCE, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
//	}
//	unsigned long get_BackColor()
//	{
//		unsigned long result;
//		InvokeHelper(DISPID_BACKCOLOR, DISPATCH_PROPERTYGET, VT_UI4, (void*)&result, NULL);
//		return result;
//	}
//	void put_BackColor(unsigned long newValue)
//	{
//		static BYTE parms[] = VTS_UI4 ;
//		InvokeHelper(DISPID_BACKCOLOR, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
//	}
//	unsigned long get_ForeColor()
//	{
//		unsigned long result;
//		InvokeHelper(DISPID_FORECOLOR, DISPATCH_PROPERTYGET, VT_UI4, (void*)&result, NULL);
//		return result;
//	}
//	void put_ForeColor(unsigned long newValue)
//	{
//		static BYTE parms[] = VTS_UI4 ;
//		InvokeHelper(DISPID_FORECOLOR, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
//	}
//	long get_Orientation()
//	{
//		long result;
//		InvokeHelper(0x17, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
//		return result;
//	}
//	void put_Orientation(long newValue)
//	{
//		static BYTE parms[] = VTS_I4 ;
//		InvokeHelper(0x17, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
//	}
//	LPDISPATCH get_Recordset()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x18, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}
//	void putref_Recordset(LPDISPATCH newValue)
//	{
//		static BYTE parms[] = VTS_DISPATCH ;
//		InvokeHelper(0x18, DISPATCH_PROPERTYPUTREF, VT_EMPTY, NULL, parms, newValue);
//	}
//	LPDISPATCH get_Font()
//	{
//		LPDISPATCH result;
//		InvokeHelper(DISPID_FONT, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}
//	void putref_Font(LPDISPATCH newValue)
//	{
//		static BYTE parms[] = VTS_DISPATCH ;
//		InvokeHelper(DISPID_FONT, DISPATCH_PROPERTYPUTREF, VT_EMPTY, NULL, parms, newValue);
//	}
//	BOOL get_Enabled()
//	{
//		BOOL result;
//		InvokeHelper(DISPID_ENABLED, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
//		return result;
//	}
//	void put_Enabled(BOOL newValue)
//	{
//		static BYTE parms[] = VTS_BOOL ;
//		InvokeHelper(DISPID_ENABLED, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
//	}
//	void Refresh()
//	{
//		InvokeHelper(0x19, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
//	}
//	void AboutBox()
//	{
//		InvokeHelper(DISPID_ABOUTBOX, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
//	}
//	void FireErrorInfo(SCODE sc, LPUNKNOWN pUnknown)
//	{
//		static BYTE parms[] = VTS_SCODE VTS_UNKNOWN ;
//		InvokeHelper(0x1a, DISPATCH_METHOD, VT_EMPTY, NULL, parms, sc, pUnknown);
//	}
//
//
//};
#if !defined(AFX_ADODC_H__F4C8512A_1C02_4EC2_9916_FDBFE1B8933C__INCLUDED_)
#define AFX_ADODC_H__F4C8512A_1C02_4EC2_9916_FDBFE1B8933C__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// Machine generated IDispatch wrapper class(es) created by Microsoft Visual C++

// NOTE: Do not modify the contents of this file.  If this class is regenerated by
//  Microsoft Visual C++, your modifications will be overwritten.


// Dispatch interfaces referenced by this interface
class C_Recordset;
class COleFont;

/////////////////////////////////////////////////////////////////////////////
// CAdodc wrapper class

class CAdodc : public CWnd
{
protected:
	DECLARE_DYNCREATE(CAdodc)
public:
	CLSID const& GetClsid()
	{
		static CLSID const clsid
			= { 0x67397aa3, 0x7fb1, 0x11d0, { 0xb1, 0x48, 0x0, 0xa0, 0xc9, 0x22, 0xe8, 0x20 } };
		return clsid;
	}
	virtual BOOL Create(LPCTSTR lpszClassName,
		LPCTSTR lpszWindowName, DWORD dwStyle,
		const RECT& rect,
		CWnd* pParentWnd, UINT nID,
		CCreateContext* pContext = NULL)
	{ return CreateControl(GetClsid(), lpszWindowName, dwStyle, rect, pParentWnd, nID); }

	BOOL Create(LPCTSTR lpszWindowName, DWORD dwStyle,
		const RECT& rect, CWnd* pParentWnd, UINT nID,
		CFile* pPersist = NULL, BOOL bStorage = FALSE,
		BSTR bstrLicKey = NULL)
	{ return CreateControl(GetClsid(), lpszWindowName, dwStyle, rect, pParentWnd, nID,
	pPersist, bStorage, bstrLicKey); }

	// Attributes
public:

	// Operations
public:
	CString GetConnectionString();
	void SetConnectionString(LPCTSTR lpszNewValue);
	CString GetUserName_();
	void SetUserName(LPCTSTR lpszNewValue);
	CString GetPassword();
	void SetPassword(LPCTSTR lpszNewValue);
	long GetMode();
	void SetMode(long nNewValue);
	long GetCursorLocation();
	void SetCursorLocation(long nNewValue);
	long GetConnectionTimeout();
	void SetConnectionTimeout(long nNewValue);
	long GetCommandTimeout();
	void SetCommandTimeout(long nNewValue);
	CString GetRecordSource();
	void SetRecordSource(LPCTSTR lpszNewValue);
	long GetCursorType();
	void SetCursorType(long nNewValue);
	long GetLockType();
	void SetLockType(long nNewValue);
	long GetCommandType();
	void SetCommandType(long nNewValue);
	long GetCacheSize();
	void SetCacheSize(long nNewValue);
	long GetMaxRecords();
	void SetMaxRecords(long nNewValue);
	long GetBOFAction();
	void SetBOFAction(long nNewValue);
	long GetEOFAction();
	void SetEOFAction(long nNewValue);
	CString GetCaption();
	void SetCaption(LPCTSTR lpszNewValue);
	long GetAppearance();
	void SetAppearance(long nNewValue);
	unsigned long GetBackColor();
	void SetBackColor(unsigned long newValue);
	unsigned long GetForeColor();
	void SetForeColor(unsigned long newValue);
	long GetOrientation();
	void SetOrientation(long nNewValue);
	C_Recordset GetRecordset();
	void SetRefRecordset(LPDISPATCH newValue);
	COleFont GetFont();
	void SetRefFont(LPDISPATCH newValue);
	BOOL GetEnabled();
	void SetEnabled(BOOL bNewValue);
	void Refresh();
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_ADODC_H__F4C8512A_1C02_4EC2_9916_FDBFE1B8933C__INCLUDED_)
