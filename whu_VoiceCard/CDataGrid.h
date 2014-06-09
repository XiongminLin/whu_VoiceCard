//// CDataGrid.h  : Declaration of ActiveX Control wrapper class(es) created by Microsoft Visual C++
//
//#pragma once
//
///////////////////////////////////////////////////////////////////////////////
//// CDataGrid
//
//class CColumns;
//#include "Column.h"
//#include "Columns.h"
//#include "COMDEF.H"
//class COleFont;
//class CStdDataFormatsDisp;
//class CPicture;
//class CColumns;
//class CSelBookmarks;
//class CSplits;
//
//class CDataGrid : public CWnd
//{
//protected:
//	DECLARE_DYNCREATE(CDataGrid)
//public:
//	CLSID const& GetClsid()
//	{
//		static CLSID const clsid
//			= { 0xCDE57A43, 0x8B86, 0x11D0, { 0xB3, 0xC6, 0x0, 0xA0, 0xC9, 0xA, 0xEA, 0x82 } };
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
//	long get_AddNewMode()
//	{
//		long result;
//		InvokeHelper(0x1, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
//		return result;
//	}
//	BOOL get_AllowAddNew()
//	{
//		BOOL result;
//		InvokeHelper(0x2, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
//		return result;
//	}
//	void put_AllowAddNew(BOOL newValue)
//	{
//		static BYTE parms[] = VTS_BOOL ;
//		InvokeHelper(0x2, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
//	}
//	BOOL get_AllowArrows()
//	{
//		BOOL result;
//		InvokeHelper(0x3, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
//		return result;
//	}
//	void put_AllowArrows(BOOL newValue)
//	{
//		static BYTE parms[] = VTS_BOOL ;
//		InvokeHelper(0x3, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
//	}
//	BOOL get_AllowDelete()
//	{
//		BOOL result;
//		InvokeHelper(0x4, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
//		return result;
//	}
//	void put_AllowDelete(BOOL newValue)
//	{
//		static BYTE parms[] = VTS_BOOL ;
//		InvokeHelper(0x4, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
//	}
//	BOOL get_AllowRowSizing()
//	{
//		BOOL result;
//		InvokeHelper(0x5, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
//		return result;
//	}
//	void put_AllowRowSizing(BOOL newValue)
//	{
//		static BYTE parms[] = VTS_BOOL ;
//		InvokeHelper(0x5, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
//	}
//	BOOL get_AllowUpdate()
//	{
//		BOOL result;
//		InvokeHelper(0x6, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
//		return result;
//	}
//	void put_AllowUpdate(BOOL newValue)
//	{
//		static BYTE parms[] = VTS_BOOL ;
//		InvokeHelper(0x6, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
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
//	long get_ApproxCount()
//	{
//		long result;
//		InvokeHelper(0x8, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
//		return result;
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
//	VARIANT get_Bookmark()
//	{
//		VARIANT result;
//		InvokeHelper(0x9, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
//		return result;
//	}
//	void put_Bookmark(VARIANT newValue)
//	{
//		static BYTE parms[] = VTS_VARIANT ;
//		InvokeHelper(0x9, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, &newValue);
//	}
//	long get_BorderStyle()
//	{
//		long result;
//		InvokeHelper(DISPID_BORDERSTYLE, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
//		return result;
//	}
//	void put_BorderStyle(long newValue)
//	{
//		static BYTE parms[] = VTS_I4 ;
//		InvokeHelper(DISPID_BORDERSTYLE, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
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
//	short get_Col()
//	{
//		short result;
//		InvokeHelper(0xa, DISPATCH_PROPERTYGET, VT_I2, (void*)&result, NULL);
//		return result;
//	}
//	void put_Col(short newValue)
//	{
//		static BYTE parms[] = VTS_I2 ;
//		InvokeHelper(0xa, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
//	}
//	BOOL get_ColumnHeaders()
//	{
//		BOOL result;
//		InvokeHelper(0xb, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
//		return result;
//	}
//	void put_ColumnHeaders(BOOL newValue)
//	{
//		static BYTE parms[] = VTS_BOOL ;
//		InvokeHelper(0xb, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
//	}
//	BOOL get_CurrentCellModified()
//	{
//		BOOL result;
//		InvokeHelper(0xc, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
//		return result;
//	}
//	void put_CurrentCellModified(BOOL newValue)
//	{
//		static BYTE parms[] = VTS_BOOL ;
//		InvokeHelper(0xc, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
//	}
//	BOOL get_CurrentCellVisible()
//	{
//		BOOL result;
//		InvokeHelper(0xd, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
//		return result;
//	}
//	void put_CurrentCellVisible(BOOL newValue)
//	{
//		static BYTE parms[] = VTS_BOOL ;
//		InvokeHelper(0xd, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
//	}
//	BOOL get_DataChanged()
//	{
//		BOOL result;
//		InvokeHelper(0xe, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
//		return result;
//	}
//	void put_DataChanged(BOOL newValue)
//	{
//		static BYTE parms[] = VTS_BOOL ;
//		InvokeHelper(0xe, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
//	}
//	LPUNKNOWN get_DataSource()
//	{
//		LPUNKNOWN result;
//		InvokeHelper(0x2a, DISPATCH_PROPERTYGET, VT_UNKNOWN, (void*)&result, NULL);
//		return result;
//	}
//	void putref_DataSource(LPUNKNOWN newValue)
//	{
//		static BYTE parms[] = VTS_UNKNOWN ;
//		InvokeHelper(0x2a, DISPATCH_PROPERTYPUTREF, VT_EMPTY, NULL, parms, newValue);
//	}
//	CString get_DataMember()
//	{
//		CString result;
//		InvokeHelper(0x2b, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
//		return result;
//	}
//	void put_DataMember(LPCTSTR newValue)
//	{
//		static BYTE parms[] = VTS_BSTR ;
//		InvokeHelper(0x2b, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
//	}
//	float get_DefColWidth()
//	{
//		float result;
//		InvokeHelper(0x10, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
//		return result;
//	}
//	void put_DefColWidth(float newValue)
//	{
//		static BYTE parms[] = VTS_R4 ;
//		InvokeHelper(0x10, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
//	}
//	BOOL get_EditActive()
//	{
//		BOOL result;
//		InvokeHelper(0x11, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
//		return result;
//	}
//	void put_EditActive(BOOL newValue)
//	{
//		static BYTE parms[] = VTS_BOOL ;
//		InvokeHelper(0x11, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
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
//	CString get_ErrorText()
//	{
//		CString result;
//		InvokeHelper(0x12, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
//		return result;
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
//	VARIANT get_FirstRow()
//	{
//		VARIANT result;
//		InvokeHelper(0x13, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
//		return result;
//	}
//	void put_FirstRow(VARIANT newValue)
//	{
//		static BYTE parms[] = VTS_VARIANT ;
//		InvokeHelper(0x13, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, &newValue);
//	}
//	LPDISPATCH get_HeadFont()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x14, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}
//	void putref_HeadFont(LPDISPATCH newValue)
//	{
//		static BYTE parms[] = VTS_DISPATCH ;
//		InvokeHelper(0x14, DISPATCH_PROPERTYPUTREF, VT_EMPTY, NULL, parms, newValue);
//	}
//	float get_HeadLines()
//	{
//		float result;
//		InvokeHelper(0x15, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
//		return result;
//	}
//	void put_HeadLines(float newValue)
//	{
//		static BYTE parms[] = VTS_R4 ;
//		InvokeHelper(0x15, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
//	}
//	long get_hWnd()
//	{
//		long result;
//		InvokeHelper(DISPID_HWND, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
//		return result;
//	}
//	long get_hWndEditor()
//	{
//		long result;
//		InvokeHelper(0x16, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
//		return result;
//	}
//	short get_LeftCol()
//	{
//		short result;
//		InvokeHelper(0x17, DISPATCH_PROPERTYGET, VT_I2, (void*)&result, NULL);
//		return result;
//	}
//	void put_LeftCol(short newValue)
//	{
//		static BYTE parms[] = VTS_I2 ;
//		InvokeHelper(0x17, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
//	}
//	long get_MarqueeStyle()
//	{
//		long result;
//		InvokeHelper(0x19, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
//		return result;
//	}
//	void put_MarqueeStyle(long newValue)
//	{
//		static BYTE parms[] = VTS_I4 ;
//		InvokeHelper(0x19, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
//	}
//	BOOL get_RecordSelectors()
//	{
//		BOOL result;
//		InvokeHelper(0x1a, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
//		return result;
//	}
//	void put_RecordSelectors(BOOL newValue)
//	{
//		static BYTE parms[] = VTS_BOOL ;
//		InvokeHelper(0x1a, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
//	}
//	BOOL get_RightToLeft()
//	{
//		BOOL result;
//		InvokeHelper(0xfffffd9d, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
//		return result;
//	}
//	void put_RightToLeft(BOOL newValue)
//	{
//		static BYTE parms[] = VTS_BOOL ;
//		InvokeHelper(0xfffffd9d, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
//	}
//	short get_Row()
//	{
//		short result;
//		InvokeHelper(0x1b, DISPATCH_PROPERTYGET, VT_I2, (void*)&result, NULL);
//		return result;
//	}
//	void put_Row(short newValue)
//	{
//		static BYTE parms[] = VTS_I2 ;
//		InvokeHelper(0x1b, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
//	}
//	long get_RowDividerStyle()
//	{
//		long result;
//		InvokeHelper(0x1c, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
//		return result;
//	}
//	void put_RowDividerStyle(long newValue)
//	{
//		static BYTE parms[] = VTS_I4 ;
//		InvokeHelper(0x1c, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
//	}
//	float get_RowHeight()
//	{
//		float result;
//		InvokeHelper(0x1d, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
//		return result;
//	}
//	void put_RowHeight(float newValue)
//	{
//		static BYTE parms[] = VTS_R4 ;
//		InvokeHelper(0x1d, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
//	}
//	long get_ScrollBars()
//	{
//		long result;
//		InvokeHelper(0xfffffde9, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
//		return result;
//	}
//	void put_ScrollBars(long newValue)
//	{
//		static BYTE parms[] = VTS_I4 ;
//		InvokeHelper(0xfffffde9, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
//	}
//	short get_SelEndCol()
//	{
//		short result;
//		InvokeHelper(0x1f, DISPATCH_PROPERTYGET, VT_I2, (void*)&result, NULL);
//		return result;
//	}
//	void put_SelEndCol(short newValue)
//	{
//		static BYTE parms[] = VTS_I2 ;
//		InvokeHelper(0x1f, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
//	}
//	long get_SelLength()
//	{
//		long result;
//		InvokeHelper(0xfffffddc, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
//		return result;
//	}
//	void put_SelLength(long newValue)
//	{
//		static BYTE parms[] = VTS_I4 ;
//		InvokeHelper(0xfffffddc, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
//	}
//	long get_SelStart()
//	{
//		long result;
//		InvokeHelper(0xfffffddd, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
//		return result;
//	}
//	void put_SelStart(long newValue)
//	{
//		static BYTE parms[] = VTS_I4 ;
//		InvokeHelper(0xfffffddd, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
//	}
//	short get_SelStartCol()
//	{
//		short result;
//		InvokeHelper(0x22, DISPATCH_PROPERTYGET, VT_I2, (void*)&result, NULL);
//		return result;
//	}
//	void put_SelStartCol(short newValue)
//	{
//		static BYTE parms[] = VTS_I2 ;
//		InvokeHelper(0x22, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
//	}
//	CString get_SelText()
//	{
//		CString result;
//		InvokeHelper(0xfffffdde, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
//		return result;
//	}
//	void put_SelText(LPCTSTR newValue)
//	{
//		static BYTE parms[] = VTS_BSTR ;
//		InvokeHelper(0xfffffdde, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
//	}
//	short get_Split()
//	{
//		short result;
//		InvokeHelper(0x24, DISPATCH_PROPERTYGET, VT_I2, (void*)&result, NULL);
//		return result;
//	}
//	void put_Split(short newValue)
//	{
//		static BYTE parms[] = VTS_I2 ;
//		InvokeHelper(0x24, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
//	}
//	BOOL get_TabAcrossSplits()
//	{
//		BOOL result;
//		InvokeHelper(0x25, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
//		return result;
//	}
//	void put_TabAcrossSplits(BOOL newValue)
//	{
//		static BYTE parms[] = VTS_BOOL ;
//		InvokeHelper(0x25, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
//	}
//	long get_TabAction()
//	{
//		long result;
//		InvokeHelper(0x26, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
//		return result;
//	}
//	void put_TabAction(long newValue)
//	{
//		static BYTE parms[] = VTS_I4 ;
//		InvokeHelper(0x26, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
//	}
//	CString get_Text()
//	{
//		CString result;
//		InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
//		return result;
//	}
//	void put_Text(LPCTSTR newValue)
//	{
//		static BYTE parms[] = VTS_BSTR ;
//		InvokeHelper(0x0, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
//	}
//	short get_VisibleCols()
//	{
//		short result;
//		InvokeHelper(0x27, DISPATCH_PROPERTYGET, VT_I2, (void*)&result, NULL);
//		return result;
//	}
//	short get_VisibleRows()
//	{
//		short result;
//		InvokeHelper(0x28, DISPATCH_PROPERTYGET, VT_I2, (void*)&result, NULL);
//		return result;
//	}
//	BOOL get_WrapCellPointer()
//	{
//		BOOL result;
//		InvokeHelper(0x29, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
//		return result;
//	}
//	void put_WrapCellPointer(BOOL newValue)
//	{
//		static BYTE parms[] = VTS_BOOL ;
//		InvokeHelper(0x29, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
//	}
//	LPDISPATCH get_DataFormats()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x2c, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}
//	void AboutBox()
//	{
//		InvokeHelper(DISPID_ABOUTBOX, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
//	}
//	LPDISPATCH CaptureImage()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x65, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}
//	void ClearSelCols()
//	{
//		InvokeHelper(0x66, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
//	}
//	void ClearFields()
//	{
//		InvokeHelper(0x67, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
//	}
//	short ColContaining(float X)
//	{
//		short result;
//		static BYTE parms[] = VTS_R4 ;
//		InvokeHelper(0x68, DISPATCH_METHOD, VT_I2, (void*)&result, parms, X);
//		return result;
//	}
//	LPDISPATCH get_Columns()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x69, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}
//	VARIANT GetBookmark(long RowNum)
//	{
//		VARIANT result;
//		static BYTE parms[] = VTS_I4 ;
//		InvokeHelper(0x6a, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms, RowNum);
//		return result;
//	}
//	void HoldFields()
//	{
//		InvokeHelper(0x6b, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
//	}
//	void ReBind()
//	{
//		InvokeHelper(0x6c, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
//	}
//	void Refresh()
//	{
//		InvokeHelper(DISPID_REFRESH, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
//	}
//	VARIANT RowBookmark(short RowNum)
//	{
//		VARIANT result;
//		static BYTE parms[] = VTS_I2 ;
//		InvokeHelper(0x6d, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms, RowNum);
//		return result;
//	}
//	short RowContaining(float Y)
//	{
//		short result;
//		static BYTE parms[] = VTS_R4 ;
//		InvokeHelper(0x6e, DISPATCH_METHOD, VT_I2, (void*)&result, parms, Y);
//		return result;
//	}
//	float RowTop(short RowNum)
//	{
//		float result;
//		static BYTE parms[] = VTS_I2 ;
//		InvokeHelper(0x6f, DISPATCH_METHOD, VT_R4, (void*)&result, parms, RowNum);
//		return result;
//	}
//	void Scroll(long Cols, long Rows)
//	{
//		static BYTE parms[] = VTS_I4 VTS_I4 ;
//		InvokeHelper(0xdc, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Cols, Rows);
//	}
//	LPDISPATCH get_SelBookmarks()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x71, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}
//	short SplitContaining(float X, float Y)
//	{
//		short result;
//		static BYTE parms[] = VTS_R4 VTS_R4 ;
//		InvokeHelper(0x72, DISPATCH_METHOD, VT_I2, (void*)&result, parms, X, Y);
//		return result;
//	}
//	LPDISPATCH get_Splits()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x73, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}
//
//	//CString CDataGrid::GetItem(int ColNum)
//	//{
//	//	CColumns cols=GetColumns();
//	//	VARIANT v_ColNum,v_Value;
//	//	v_ColNum.vt=VT_I2;
//	//	v_ColNum.iVal=ColNum;
//
//	//	CColumn col = cols.GetItem(v_ColNum);
//
//	//	v_Value=col.GetValue();
//	//	return v_Value.bstrVal;
//	//}
//};
#if !defined(AFX_DATAGRID_H__01A425F7_F50F_468E_AFE4_1DD257C7CDCD__INCLUDED_)
#define AFX_DATAGRID_H__01A425F7_F50F_468E_AFE4_1DD257C7CDCD__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// Machine generated IDispatch wrapper class(es) created by Microsoft Visual C++

// NOTE: Do not modify the contents of this file.  If this class is regenerated by
//  Microsoft Visual C++, your modifications will be overwritten.


// Dispatch interfaces referenced by this interface
#include "font.h"
#include "StdDataFormatsDisp.h"
#include "Picture.h"
#include "Column.h"
#include "Columns.h"
#include "COMDEF.H"
#include "SelBookmarks.h"
#include "Splits.h"

class COleFont;
class CStdDataFormatsDisp;
class CPicture;
class CColumns;
class CSelBookmarks;
class CSplits;

/////////////////////////////////////////////////////////////////////////////
// CDataGrid wrapper class

class CDataGrid : public CWnd
{
protected:
	DECLARE_DYNCREATE(CDataGrid)
public:
	CLSID const& GetClsid()
	{
		static CLSID const clsid
			= { 0xcde57a43, 0x8b86, 0x11d0, { 0xb3, 0xc6, 0x0, 0xa0, 0xc9, 0xa, 0xea, 0x82 } };
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
	CString GetItem(int ColNum);
	long GetAddNewMode();
	BOOL GetAllowAddNew();
	void SetAllowAddNew(BOOL bNewValue);
	BOOL GetAllowArrows();
	void SetAllowArrows(BOOL bNewValue);
	BOOL GetAllowDelete();
	void SetAllowDelete(BOOL bNewValue);
	BOOL GetAllowRowSizing();
	void SetAllowRowSizing(BOOL bNewValue);
	BOOL GetAllowUpdate();
	void SetAllowUpdate(BOOL bNewValue);
	long GetAppearance();
	void SetAppearance(long nNewValue);
	long GetApproxCount();
	unsigned long GetBackColor();
	void SetBackColor(unsigned long newValue);
	VARIANT GetBookmark();
	void SetBookmark(const VARIANT& newValue);
	long GetBorderStyle();
	void SetBorderStyle(long nNewValue);
	CString GetCaption();
	void SetCaption(LPCTSTR lpszNewValue);
	short GetCol();
	void SetCol(short nNewValue);
	BOOL GetColumnHeaders();
	void SetColumnHeaders(BOOL bNewValue);
	BOOL GetCurrentCellModified();
	void SetCurrentCellModified(BOOL bNewValue);
	BOOL GetCurrentCellVisible();
	void SetCurrentCellVisible(BOOL bNewValue);
	BOOL GetDataChanged();
	void SetDataChanged(BOOL bNewValue);
	LPUNKNOWN GetDataSource();
	void SetRefDataSource(LPUNKNOWN newValue);
	CString GetDataMember();
	void SetDataMember(LPCTSTR lpszNewValue);
	float GetDefColWidth();
	void SetDefColWidth(float newValue);
	BOOL GetEditActive();
	void SetEditActive(BOOL bNewValue);
	BOOL GetEnabled();
	void SetEnabled(BOOL bNewValue);
	CString GetErrorText();
	COleFont GetFont();
	void SetRefFont(LPDISPATCH newValue);
	unsigned long GetForeColor();
	void SetForeColor(unsigned long newValue);
	VARIANT GetFirstRow();
	void SetFirstRow(const VARIANT& newValue);
	COleFont GetHeadFont();
	void SetRefHeadFont(LPDISPATCH newValue);
	float GetHeadLines();
	void SetHeadLines(float newValue);
	long GetHWnd();
	long GetHWndEditor();
	short GetLeftCol();
	void SetLeftCol(short nNewValue);
	long GetMarqueeStyle();
	void SetMarqueeStyle(long nNewValue);
	BOOL GetRecordSelectors();
	void SetRecordSelectors(BOOL bNewValue);
	BOOL GetRightToLeft();
	void SetRightToLeft(BOOL bNewValue);
	short GetRow();
	void SetRow(short nNewValue);
	long GetRowDividerStyle();
	void SetRowDividerStyle(long nNewValue);
	float GetRowHeight();
	void SetRowHeight(float newValue);
	long GetScrollBars();
	void SetScrollBars(long nNewValue);
	short GetSelEndCol();
	void SetSelEndCol(short nNewValue);
	long GetSelLength();
	void SetSelLength(long nNewValue);
	long GetSelStart();
	void SetSelStart(long nNewValue);
	short GetSelStartCol();
	void SetSelStartCol(short nNewValue);
	CString GetSelText();
	void SetSelText(LPCTSTR lpszNewValue);
	short GetSplit();
	void SetSplit(short nNewValue);
	BOOL GetTabAcrossSplits();
	void SetTabAcrossSplits(BOOL bNewValue);
	long GetTabAction();
	void SetTabAction(long nNewValue);
	CString GetText();
	void SetText(LPCTSTR lpszNewValue);
	short GetVisibleCols();
	short GetVisibleRows();
	BOOL GetWrapCellPointer();
	void SetWrapCellPointer(BOOL bNewValue);
	CStdDataFormatsDisp GetDataFormats();
	CPicture CaptureImage();
	void ClearSelCols();
	void ClearFields();
	short ColContaining(float X);
	CColumns GetColumns();
	VARIANT GetBookmark(long RowNum);
	void HoldFields();
	void ReBind();
	void Refresh();
	VARIANT RowBookmark(short RowNum);
	short RowContaining(float Y);
	float RowTop(short RowNum);
	void Scroll(long Cols, long Rows);
	CSelBookmarks GetSelBookmarks();
	short SplitContaining(float X, float Y);
	CSplits GetSplits();
	DECLARE_EVENTSINK_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_DATAGRID_H__01A425F7_F50F_468E_AFE4_1DD257C7CDCD__INCLUDED_)
