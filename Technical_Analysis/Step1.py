'''
Pro1. 신호 발굴
# 조건
 1) 유통주식수의 n% 이상
 2) 전일 거래량의 n배 이상
 3) 거래대금이 n억 이상
'''

import pandas as pd
import os
import numpy as np
import datetime

class step1_cls:
    def exe_step1(self, con):
        # 경로에 있는 csv 파일명을 가져와서 배열로 저장
        csv_files_collect = []
        for path, dirs, files in os.walk("F:/JusikData/oneday_csv/onedaydata"):
            csv_files_collect.append(''.join(files))

        # 배열의 첫번째는 값이 없으므로 제거
        del csv_files_collect[0]
        
        # .csv를 빼서 종목명만 집어넣기
        JongMok = []
        for i in csv_files_collect:
            aa = i.replace('.csv','')
            JongMok.append(aa)

        # 배열선언하기
        ddate = [] # 리턴할 일자
        dname = [] # 리턴할 종목명
        dcode = [] # 리턴할 종목코드
        dmarket = [] # 리턴할 시장구분
        dopen = [] # 리턴할 시가
        dhigh = [] # 리턴할 고가
        dlow = [] # 리턴할 저가
        dclose = [] # 리턴할 종가
        dvolume = [] # 리턴할 거래량
        dmoney = [] # 리턴할 거래대금
        dstock = [] # 리턴할 주식수
        
        
        # 오늘 날짜 구하기
        d_today = datetime.date.today()
        nowDate = d_today.strftime('%Y%m%d')
        
        # 실행문
        print('기술적 분석 Step1 시작')
        
        # 종목 수 만큼 for문 돌려서 조건에 맞는 종목 찾기 
        for name in JongMok:
            print('step1_',name, '...진행중')
            
            # 해당 종목의 경로와 데이터 가져오기
            path = "F:/JusikData/oneday_csv/onedaydata/"+name+'/'+name+'.csv'
            data = pd.read_csv(path, encoding='cp949')
            
            if con == 'con1':
                # 조건1 : 유통주식수의 n% 이상
                for i in data.index:
                    try:
                        if data.iloc[i]['거래량'] > data.iloc[i]['상장주식수'] * 0.2 :
                            ddate.append(data.iloc[i]['일자']) # 리턴할 일자
                            dname.append(data.iloc[i]['종목명']) # 리턴할 종목명
                            dcode.append(data.iloc[i]['종목코드']) # 리턴할 종목코드
                            dmarket.append(data.iloc[i]['시장구분']) # 리턴할 시장구분
                            dopen.append(data.iloc[i]['시가']) # 리턴할 시가
                            dhigh.append(data.iloc[i]['고가']) # 리턴할 고가
                            dlow.append(data.iloc[i]['저가']) # 리턴할 저가
                            dclose.append(data.iloc[i]['종가']) # 리턴할 종가
                            dvolume.append(data.iloc[i]['거래량']) # 리턴할 거래량
                            dmoney.append(data.iloc[i]['거래대금']) # 리턴할 거래대금
                            dstock.append(data.iloc[i]['상장주식수']) # 리턴할 주식수
                    except:
                        print(name, '_', str(data.iloc[i]['일자']), '_오류')
                        
            elif con == 'con2':    
                # 조건 2: 전일 거래량의 n배
                for i in data.index:
                    try:
                        if data.iloc[i]['거래량'] > data.iloc[i-1]['거래량'] * 5 :
                            ddate.append(data.iloc[i]['일자']) # 리턴할 일자
                            dname.append(data.iloc[i]['종목명']) # 리턴할 종목명
                            dcode.append(data.iloc[i]['종목코드']) # 리턴할 종목코드
                            dmarket.append(data.iloc[i]['시장구분']) # 리턴할 시장구분
                            dopen.append(data.iloc[i]['시가']) # 리턴할 시가
                            dhigh.append(data.iloc[i]['고가']) # 리턴할 고가
                            dlow.append(data.iloc[i]['저가']) # 리턴할 저가
                            dclose.append(data.iloc[i]['종가']) # 리턴할 종가
                            dvolume.append(data.iloc[i]['거래량']) # 리턴할 거래량
                            dmoney.append(data.iloc[i]['거래대금']) # 리턴할 거래대금
                            dstock.append(data.iloc[i]['상장주식수']) # 리턴할 주식수
                    except:
                        print(name, '_', str(data.iloc[i]['일자']), '_오류')
                        
            elif con == 'con3':            
                # 조건 3: 거래대금이 n억 이상
                for i in data.index:
                    try:
                        if data.iloc[i]['거래대금'] > 5000000000 :
                            ddate.append(data.iloc[i]['일자']) # 리턴할 일자
                            dname.append(data.iloc[i]['종목명']) # 리턴할 종목명
                            dcode.append(data.iloc[i]['종목코드']) # 리턴할 종목코드
                            dmarket.append(data.iloc[i]['시장구분']) # 리턴할 시장구분
                            dopen.append(data.iloc[i]['시가']) # 리턴할 시가
                            dhigh.append(data.iloc[i]['고가']) # 리턴할 고가
                            dlow.append(data.iloc[i]['저가']) # 리턴할 저가
                            dclose.append(data.iloc[i]['종가']) # 리턴할 종가
                            dvolume.append(data.iloc[i]['거래량']) # 리턴할 거래량
                            dmoney.append(data.iloc[i]['거래대금']) # 리턴할 거래대금
                            dstock.append(data.iloc[i]['상장주식수']) # 리턴할 주식수
                    except:
                        print(name, '_', str(data.iloc[i]['일자']), '_오류')
            
        # 리턴할 데이터 셋
        dataset = pd.DataFrame({'일자': ddate, '종목명': dname, '종목코드': dcode, '시장구분': dmarket,
                                '시가': dopen, '고가': dhigh, '저가': dlow, '종가': dclose,
                                '거래량': dvolume, '거래대금': dmoney, '상장주식수': dstock
                                })
    
        # 리턴하기
        return dataset
