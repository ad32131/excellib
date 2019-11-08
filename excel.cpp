#include "excel.h"

BOOL excel::excelstart(HWND hDlg) {;
	this->hDlg = hDlg;
	pXlApplication = NULL;

	hResult = resultExceptionHandle( OleInitialize(NULL) );
	hResult = resultExceptionHandle( CLSIDFromProgID(OLESTR("Excel.Application"), &clsid) );

	hResult = resultExceptionHandle( CoCreateInstance(clsid, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch, (LPVOID*)&pXlApplication) );


	pXlApplication_create_id = pXlApplication;
	pXlApplication_read_id = pXlApplication;
	pXlApplication_close_id = pXlApplication;
	pXlApplication_quit_id = pXlApplication;
	pXlWorksheet = pXlApplication;
	pXlRangeCells = pXlApplication;
	return TRUE;
}

BOOL excel::excelcreatenewwork() {

	



	lpszName = (OLECHAR*)OLESTR("Workbooks");
	setdispParams();
	hResult = resultExceptionHandle(pXlApplication_create_id->GetIDsOfNames(IID_NULL, &lpszName, 1, LOCALE_USER_DEFAULT, &dispid) );

	setdispParams();
	hResult = resultExceptionHandle( pXlApplication_create_id->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYGET, &dispParams, &variant, NULL, NULL) );

	pXlWorkbooks_createtmp = variant.pdispVal;

	lpszName = (OLECHAR*)OLESTR("Add");
	hResult = pXlWorkbooks_createtmp->GetIDsOfNames(IID_NULL, &lpszName, 1, LOCALE_USER_DEFAULT, &dispid);

	setdispParams();
	hResult = resultExceptionHandle( pXlWorkbooks_createtmp->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYGET, &dispParams, &variant, NULL, NULL) );

	pXlWorkbook_createtmp = variant.pdispVal;


	pXlApplication_close_id = pXlWorkbooks_createtmp;
	return TRUE;
}

BOOL excel::excelreadfile(OLECHAR* input) {
	//OLECHAR inputname[MAX_PATH] = OLESTR("C:\\testcc\\Mthetr_Avails.xlsx");
	pXlApplication_read_id = pXlWorkbooks_createtmp;

	setdispParams();
	lpszName = (OLECHAR*)OLESTR("Open");
	

	setdispParams();
	hResult = resultExceptionHandle(pXlApplication_read_id->GetIDsOfNames(IID_NULL, &lpszName, 1, LOCALE_USER_DEFAULT, &dispid) );

	dispParams.cArgs = 1;
	dispParams.cNamedArgs = 1;
	dispParams.rgvarg = new VARIANTARG[1];
	dispParams.rgdispidNamedArgs = new DISPID[1];
	dispParams.rgvarg[0].vt = VT_BSTR;
	dispParams.rgvarg[0].bstrVal = SysAllocString( /*inputname*/ input );
	dispParams.rgdispidNamedArgs[0] = 0;
	VariantInit(&variant);
	hResult = resultExceptionHandle( pXlApplication_read_id->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_METHOD, &dispParams, &variant, &excepInfo, NULL) );

	SysFreeString(dispParams.rgvarg[0].bstrVal);
	delete[] dispParams.rgvarg;
	delete[] dispParams.rgdispidNamedArgs;

	pXlApplication_read_id = variant.pdispVal;

	return TRUE;
}

BOOL excel::excelDataSelect(OLECHAR* input) {
	uArgErr = 0;
	lpszName = (OLECHAR*)OLESTR("Cells");
	hResult = pXlWorksheet->GetIDsOfNames(IID_NULL, &lpszName, 1, LOCALE_USER_DEFAULT, &dispid);



	setdispParams();
	hResult = resultExceptionHandle( pXlWorksheet->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYGET, &dispParams, &variant, NULL, NULL) );
	pXlRangeCells = variant.pdispVal;

	//OLESTR("$A$1")

	lpszName = (OLECHAR*)OLESTR("Range");
	hResult = resultExceptionHandle( pXlWorksheet->GetIDsOfNames(IID_NULL, &lpszName, 1, LOCALE_USER_DEFAULT, &dispid) );

	dispParams.cArgs = 1;
	dispParams.cNamedArgs = 1;
	dispParams.rgvarg = new VARIANTARG[1];
	dispParams.rgdispidNamedArgs = new DISPID[1];
	VariantInit(&dispParams.rgvarg[0]);
	dispParams.rgvarg[0].vt = VT_BSTR;
	//dispParams.rgvarg[0].bstrVal = SysAllocString(OLESTR("$A$1:$C$3")); range
	dispParams.rgvarg[0].bstrVal = SysAllocString(input);
	dispParams.rgdispidNamedArgs[0] = 0;
	VariantInit(&variant);
	hResult = resultExceptionHandle( pXlWorksheet->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYGET, &dispParams, &variant, &excepInfo, &uArgErr) );

	pXlRangeRange = variant.pdispVal;



	
	return TRUE;
}

double excel::excelDataRead(TCHAR* outputString) {
	lpszName = (OLECHAR*)OLESTR("Value");
	hResult = resultExceptionHandle(pXlRangeRange->GetIDsOfNames(IID_NULL, &lpszName, 1, LOCALE_USER_DEFAULT, &dispid));

	dispParams.cArgs = 0;
	dispParams.cNamedArgs = 0;
	dispParams.rgvarg = new VARIANTARG[1];
	dispParams.rgdispidNamedArgs = new DISPID[1];

	VariantInit(&variant);

	hResult = resultExceptionHandle(pXlRangeRange->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYGET, &dispParams, &variant, NULL, NULL));

	
	if (variant.vt == 8) {//string
		memcpy_s(outputString, MAX_PATH, variant.bstrVal, wcslen(variant.bstrVal)*2);
		return 0;
	}
	else if(variant.vt == 5){//double
		return variant.dblVal;
	}

	//MessageBox(hDlg, variant., L"Data", 0);

	delete[] dispParams.rgvarg;
	delete[] dispParams.rgdispidNamedArgs;

	
}



BOOL excel::excelDataWrite(OLECHAR* input) {
	lpszName = (OLECHAR*)OLESTR("Value");
	hResult = resultExceptionHandle( pXlRangeRange->GetIDsOfNames(IID_NULL, &lpszName, 1, LOCALE_USER_DEFAULT, &dispid) );

	dispParams.cArgs = 1;
	dispParams.cNamedArgs = 1;
	dispParams.rgvarg = new VARIANTARG[1];
	dispParams.rgdispidNamedArgs = new DISPID[1];

	dispParams.rgvarg[0].vt = VT_BSTR;
	dispParams.rgvarg[0].bstrVal = SysAllocString(input);
	dispParams.rgvarg[0].cVal;

	dispParams.rgdispidNamedArgs[0] = DISPID_PROPERTYPUT;

	VariantInit(&variant);

	hResult = resultExceptionHandle( pXlRangeRange->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYPUT, &dispParams, &variant, NULL, NULL) );


	delete[] dispParams.rgvarg;
	delete[] dispParams.rgdispidNamedArgs;

	return TRUE;
}

BOOL excel::excelsave() {
	lpszName = (OLECHAR*)OLESTR("Save");
	hResult = resultExceptionHandle( pXlApplication_read_id->GetIDsOfNames(IID_NULL, &lpszName, 1, LOCALE_USER_DEFAULT, &dispid) );
	

	
	dispParams.cArgs = 0;
	dispParams.cNamedArgs = 0;
	dispParams.rgvarg = NULL;
	dispParams.rgdispidNamedArgs = NULL;
	VariantInit(&variant);
	hResult = resultExceptionHandle( pXlApplication_read_id->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_METHOD, &dispParams, &variant, NULL, NULL) );

	return TRUE;
}

BOOL excel::excelclosefile(){

	lpszName = (OLECHAR*)OLESTR("Close");
	hResult = resultExceptionHandle( pXlApplication_close_id->GetIDsOfNames(IID_NULL, &lpszName, 1, LOCALE_USER_DEFAULT, &dispid) );

	setdispParams();
	hResult = resultExceptionHandle( pXlApplication_close_id->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYGET, &dispParams, &variant, NULL, NULL) );

	pXlApplication_close_id = variant.pdispVal;

	return TRUE;
}

BOOL excel::excelquit() {
	hResult = OleInitialize(NULL);
	hResult = CLSIDFromProgID(OLESTR("Excel.Application"), &clsid);

	lpszName = (OLECHAR*)OLESTR("Quit");
	hResult = resultExceptionHandle( pXlApplication_quit_id->GetIDsOfNames(IID_NULL, &lpszName, 1, LOCALE_USER_DEFAULT, &dispid) );

	setdispParams();
	hResult = resultExceptionHandle( pXlApplication_quit_id->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_METHOD, &dispParams, &variant, NULL, NULL) );

	
	dispatchUnInit( pXlApplication );
	dispatchUnInit( pXlApplication_create_id );
	dispatchUnInit( pXlApplication_read_id );
	dispatchUnInit( pXlApplication_close_id );
	dispatchUnInit( pXlApplication_quit_id );

	dispatchUnInit( pXlWorkbooks_createtmp );
	dispatchUnInit( pXlWorkbook_createtmp );
	dispatchUnInit( pXlWorksheets );
	dispatchUnInit( pXlWorksheet );
	dispatchUnInit( pXlRangeCells );

	dispatchUnInit( pXlRangeCell );
	dispatchUnInit( pXlRangeRange );

	OleUninitialize();
	return TRUE;
}


BOOL excel::dispatchUnInit(IDispatch* dis) {
	if (dis && (((short)dis) != -1) ) dis->Release();
	return TRUE;
};

BOOL excel::setdispParams() {
	dispParams.cArgs = 0;
	dispParams.cNamedArgs = 0;
	dispParams.rgvarg = (VARIANTARG*)NULL;
	dispParams.rgdispidNamedArgs = (DISPID*)NULL;
	VariantInit(&variant);

	return TRUE;
}

HRESULT excel::resultExceptionHandle(HRESULT hResult) {
	if (hResult == 0) {
		return hResult;
	}
	else if (hResult == -2147352573) {
		return hResult;
	}
	else {
		
		TCHAR error[MAX_PATH];
		wsprintf(error, L"0x%p CLASS %s ERROR!", hResult , lpszName);
		MessageBox(this->hDlg, error, L"error", 0);
		/*
		excelquit();
		exit(0);
		*/
		return hResult;
	}
};

excel::~excel(){


}