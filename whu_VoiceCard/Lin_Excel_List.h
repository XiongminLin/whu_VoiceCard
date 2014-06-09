#pragma once
#include "CRange.h"
#include "CSheets.h"
#include "CWorkbook.h"
#include "CWorkbooks.h"
#include "CApplication.h"
#include "CWorksheet.h"
#include "CWorksheets.h"
class CLin_Excel_List
{
public:
	CLin_Excel_List(void);
	~CLin_Excel_List(void);
private:
	CString GetEnglishCharacter(int Num);
	void InitList(CListCtrl &m_List,CStringArray &ColumName);
	void InsertList(CListCtrl &m_List,CStringArray &ItemArr);
public:
	void ListToExcel(CListCtrl &m_List,CString m_FilePath);
	void ExcelToList(CString m_FilePath,CListCtrl &m_List);
	void ArrToList(CStringArray *m_DutyArr,int Num,CListCtrl &m_List);
	int ListToArr(CListCtrl &m_List,CStringArray *m_DutyArr);
	int ExcelToArr(CString m_FilePath,CStringArray *m_DutyArr);//返回CString数组的个数
	void ArrToExcel(CStringArray *m_DutyArr,int Num,CString m_FilePath);
};

