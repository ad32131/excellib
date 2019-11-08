# C++ Excel Lib V 1.0

# usage

## 헤더 선언
	 #include "excel.h"


## 클래스 선언
	excel excel1;

​
## 파일 읽어서 데이터 가져올시
	//데이터 초기화 클래스 생성
	excel1.excelstart();
	excel1.excelreadfile(file_name_input);

	//데이터 읽기
	excel1.excelDataSelect(range_input);
	excel1.excelDataRead(getData);

	​//데이터 수치 읽기
	excel1.excelDataSelect(range_input);
	double result_double = excel1.excelDataRead(0); 

## 데이터쓰기
	excel1.excelDataSelect(range_input);
	excel1.excelDataWrite(writeData);

	//데이터 닫기 클래스 소멸
	excel1.excelclosefile();
	excel1.excelquit();
