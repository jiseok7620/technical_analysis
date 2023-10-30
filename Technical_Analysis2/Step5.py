import pandas as pd
import openpyxl
import os
import win32com.client
import pythoncom
import time

class ebestapi_t1921_cls:
    # 쿼리 상태 초기화
    query_state = 0

    # 데이터 받으면 해당 이벤트로 이동
    def OnReceiveData(self, code):
        ebestapi_t1921_cls.query_state = 1
        print(code, ' : t1921 데이터 수신 완료')
        
    # 실행 시 메세지 및 에러 받음
    def OnReceiveMessage(self, err, msgco, msg):
        print('t1921 에러발생 : ', err)
        print('t1921 메세지 : ', msg)

class t1921_cls:
    def exe_t1921(self, shcode, edate):
        # 쿼리 객체 생성
        object_t1921 = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", ebestapi_t1921_cls)
        
        # Res 파일 등록
        object_t1921.ResFileName = "C:\\eBEST\\xingAPI\\Res\\t1921.res"
        
        # InBlock에 값 설정
        object_t1921.SetFieldData("t1921InBlock", "shcode", 0, shcode)
        object_t1921.SetFieldData("t1921InBlock", "gubun", 0, "1")
        object_t1921.SetFieldData("t1921InBlock", "date", 0, edate)
        object_t1921.SetFieldData("t1921InBlock", "idx", 0, "")
        
        # 데이터 요청 - False : 반복조회, True : 연속조회
        object_t1921.Request(False)
        
        # 10분내 요청한 요청 횟수 취득
        count_limit = object_t1921.GetTRCountLimit("t1921")
        count_request = object_t1921.GetTRCountRequest("t1921")
        print('t1921 10분 당 제한 건수 : ', count_limit)
        print('t1921 10분 내 요청 횟수 : ', count_request)
        
        # 수신 대기
        while ebestapi_t1921_cls.query_state == 0:
            pythoncom.PumpWaitingMessages()
        
        # 연속조회시 개수 가져오기 => 연속 시 1을 붙여줌
        count = object_t1921.GetBlockCount("t1921OutBlock1")
        print('t1921 가져올 데이터의 수 : ',count)
        
        arr_date = [] # 일자
        arr_amt0000 = []
        arr_amt0001 = []
        arr_amt0002 = []
        arr_amt0003 = []
        arr_amt0004 = []
        arr_amt0005 = []
        
        # 필요한 필드 가져오기
        for i in range(count):
            mmdate = object_t1921.GetFieldData("t1921OutBlock1", "mmdate", i) # 일자
            nvolume = object_t1921.GetFieldData("t1921OutBlock1", "nvolume", i) # 신규
            svolume = object_t1921.GetFieldData("t1921OutBlock1", "svolume", i) # 상환
            jvolume = object_t1921.GetFieldData("t1921OutBlock1", "jvolume", i) # 잔고
            price = object_t1921.GetFieldData("t1921OutBlock1", "price", i) # 금액
            gyrate = object_t1921.GetFieldData("t1921OutBlock1", "gyrate", i) # 공여율
            jkrate = object_t1921.GetFieldData("t1921OutBlock1", "jkrate", i) # 잔고율
            
            arr_date.append(mmdate)
            arr_amt0000.append(nvolume)
            arr_amt0001.append(svolume)
            arr_amt0002.append(jvolume)
            arr_amt0003.append(price)
            arr_amt0004.append(gyrate)
            arr_amt0005.append(jkrate)
            
        # 데이터 프레임으로 만들기
        data = pd.DataFrame({'일자' : arr_date, '신규' : arr_amt0000, '상환' : arr_amt0001, 
                             '잔고' : arr_amt0002, '금액' : arr_amt0003, '공여율': arr_amt0004, '잔고율': arr_amt0005
                             })
        
        # 데이터프레임 리턴하기
        return data