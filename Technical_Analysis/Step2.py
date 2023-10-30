'''
Pro2. 신호의 n년치 수정주가 데이터 가져오기
# 데이터
 1) 수정 시가, 고가, 저가, 종가
 2) 거래량
'''

import pandas as pd
import openpyxl
import os
import win32com.client
import pythoncom
import time

class ebestapi_ChartIndex_cls:
    # 쿼리 상태 초기화
    query_state = 0

    # 데이터 받으면 해당 이벤트로 이동
    def OnReceiveData(self, code):
        ebestapi_ChartIndex_cls.query_state = 1
        #print(code, ' : ChartIndex 데이터 수신 완료')
        
    # 실행 시 메세지 및 에러 받음
    def OnReceiveMessage(self, err, msgco, msg):
        print('ChartIndex 에러발생 : ', err)
        #print('ChartIndex 메세지 : ', msg)
        
class ChartIndex_cls:
    def exe_ChartIndex(self, shcode, edate):
        # query_state
        ebestapi_ChartIndex_cls.query_state = 0
        
        
        
        ## ----------------------------------------------------------------------------------------------- ##
        # 쿼리 객체 생성
        object_ChartIndex = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", ebestapi_ChartIndex_cls)
        
        # Res 파일 등록
        object_ChartIndex.ResFileName = "C:\\eBEST\\xingAPI\\Res\\ChartIndex.res"
        
        # InBlock에 값 설정
        object_ChartIndex.SetFieldData("ChartIndexInBlock", "indexname", 0, "")
        object_ChartIndex.SetFieldData("ChartIndexInBlock", "market", 0, "1")
        object_ChartIndex.SetFieldData("ChartIndexInBlock", "period", 0, "2")
        object_ChartIndex.SetFieldData("ChartIndexInBlock", "shcode", 0, shcode)
        object_ChartIndex.SetFieldData("ChartIndexInBlock", "qrycnt", 0, "500")
        object_ChartIndex.SetFieldData("ChartIndexInBlock", "ncnt", 0, "")
        object_ChartIndex.SetFieldData("ChartIndexInBlock", "sdate", 0, "")
        object_ChartIndex.SetFieldData("ChartIndexInBlock", "edate", 0, edate)
        object_ChartIndex.SetFieldData("ChartIndexInBlock", "Isamend", 0, "1") # 수정주가 반영여부(1:반영)
        object_ChartIndex.SetFieldData("ChartIndexInBlock", "Isgab", 0, "1") # 갭보정여부(1:보정)
        object_ChartIndex.SetFieldData("ChartIndexInBlock", "IsReal", 0, "0")
        
        # 데이터 요청
        # 차트인덱스는 RequestService로 요청해야함
        object_ChartIndex.RequestService("ChartIndex", 0)
        
        # 10분내 요청한 요청 횟수 취득
        count_limit = object_ChartIndex.GetTRCountLimit("ChartIndex")
        count_request = object_ChartIndex.GetTRCountRequest("ChartIndex")
        #print('ChartIndex 10분 당 제한 건수 : ', count_limit)
        #print('ChartIndex 10분 내 요청 횟수 : ', count_request)
        
        # 수신 대기
        while ebestapi_ChartIndex_cls.query_state == 0:
            pythoncom.PumpWaitingMessages()
        
        # 연속조회 시 가져올 데이터 개수
        count = object_ChartIndex.GetBlockCount("ChartIndexOutBlock1")
        print('ChartIndex 가져올 데이터의 수 : ',count)
        
        arr_date = []
        arr_open = []
        arr_high = []
        arr_low = []
        arr_close = []
        arr_volume = []
        # 필요한 필드 가져오기
        for i in range(count):
            date = object_ChartIndex.GetFieldData("ChartIndexOutBlock1", "date", i)
            open = object_ChartIndex.GetFieldData("ChartIndexOutBlock1", "open", i)
            high = object_ChartIndex.GetFieldData("ChartIndexOutBlock1", "high", i)
            low = object_ChartIndex.GetFieldData("ChartIndexOutBlock1", "low", i)            
            close = object_ChartIndex.GetFieldData("ChartIndexOutBlock1", "close", i)
            volume = object_ChartIndex.GetFieldData("ChartIndexOutBlock1", "volume", i)
            arr_date.append(date)
            arr_open.append(open)
            arr_high.append(high)
            arr_low.append(low)
            arr_close.append(close)
            arr_volume.append(volume)
            
        # 데이터 프레임으로 만들기
        data = pd.DataFrame({'일자' : arr_date, '시가' : arr_open, '고가' : arr_high,
                             '저가' : arr_low, '종가' : arr_close, '거래량' : arr_volume
                             })
        
        # 인덱스 0 값 지우고, 인덱스 초기화하기
        # 0 일자 0 0 0 0 0 요런값이 나오기 때문에
        data = data.drop(index=0, axis=0)
        data.reset_index(drop=True, inplace=True)
        
        # 데이터프레임 리턴하기
        return data