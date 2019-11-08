#pragma once
#include <OAIdl.h>

class excel {
private:
public:
	HWND hDlg;
	HRESULT hResult;
	CLSID clsid;
	UINT uArgErr;

	//action
	IDispatch* pXlApplication = NULL;
	IDispatch* pXlApplication_create_id = NULL;
	IDispatch* pXlApplication_read_id = NULL;
	IDispatch* pXlApplication_close_id = NULL;
	IDispatch* pXlApplication_quit_id = NULL;

	//create
	IDispatch* pXlWorkbooks_createtmp = NULL;
	IDispatch* pXlWorkbook_createtmp = NULL;

	//readData
	IDispatch* pXlWorksheets = NULL;
	IDispatch* pXlWorksheet = NULL;
	IDispatch* pXlRangeCells = NULL;
	IDispatch* pXlRangeCell = NULL;
	IDispatch* pXlRangeRange = NULL;


	EXCEPINFO excepInfo;
	DISPID dispid;
	OLECHAR* lpszName;
	DISPPARAMS dispParams = { NULL, NULL, 0, 0 };
	VARIANT variant;


	BOOL excelstart(HWND hDlg);
	BOOL excelcreatenewwork();
	BOOL excelreadfile(OLECHAR* input);
	
	BOOL excelsave();
	BOOL excelclosefile();
	BOOL excelquit();

	BOOL dispatchUnInit(IDispatch* dis);
	HRESULT resultExceptionHandle(HRESULT hResult);
	BOOL setdispParams();

	BOOL excelDataSelect(OLECHAR* input);
	double excelDataRead(TCHAR* outputString);
	BOOL excelDataWrite(OLECHAR* input);
	~excel();
};