from Billionaire.Technical_Analysis3.ebest_server_login import login_cls
from Billionaire.Technical_Analysis3.Step1 import ChartIndex_cls
from Billionaire.Technical_Analysis3.Step2 import trendline_cls
from Billionaire.Technical_Analysis3.Step3 import t1702_cls
from Billionaire.Technical_Analysis3.Step4 import jisu_cls
from Billionaire.Technical_Analysis3.Step5 import trends_cls
from openpyxl.drawing.image import Image
from pandas.io.excel._base import read_excel
import pandas as pd
import time
import os
import openpyxl
import sys
import random
from lxml.html.builder import SUP

class test_cls:
    def exe_test(self):
        # 구할 종목
        jongname = '엘컴텍'
        jongcode = '037950'
        buydate = '20220211'
        
        
        # 서버로그인
        login_cls.exe_login(self)
        
        # 지수 추세 데이터구하기
        if os.path.isfile('F:/JusikData/API/Analy_Technical/지수데이터/kospi_high.xlsx') :
                highpoint_data_kp = pd.read_excel('F:/JusikData/API/Analy_Technical/지수데이터/kospi_high.xlsx', engine='openpyxl')
                lowpoint_data_kp = pd.read_excel('F:/JusikData/API/Analy_Technical/지수데이터/kospi_low.xlsx', engine='openpyxl')
                highpoint_data_kq = pd.read_excel('F:/JusikData/API/Analy_Technical/지수데이터/kosdaq_high.xlsx', engine='openpyxl')
                lowpoint_data_kq = pd.read_excel('F:/JusikData/API/Analy_Technical/지수데이터/kosdaq_low.xlsx', engine='openpyxl')
        else :
            # 코스피(2015~)
            highpoint_data_kp, lowpoint_data_kp = jisu_cls.make_jisu(self, 'kospi', 10)
            
            # 코스닥(2015~) 
            highpoint_data_kq, lowpoint_data_kq = jisu_cls.make_jisu(self, 'kosdaq', 10)
            
            # 엑셀 저장
            highpoint_data_kp.to_excel('F:/JusikData/API/Analy_Technical/지수데이터/kospi_high.xlsx', index=False)
            lowpoint_data_kp.to_excel('F:/JusikData/API/Analy_Technical/지수데이터/kospi_low.xlsx', index=False)
            highpoint_data_kq.to_excel('F:/JusikData/API/Analy_Technical/지수데이터/kosdaq_high.xlsx', index=False)
            lowpoint_data_kq.to_excel('F:/JusikData/API/Analy_Technical/지수데이터/kosdaq_low.xlsx', index=False)
        
        #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
        #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
        #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
        #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
        #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
        
        # 수정주가 가져오기
        if os.path.isfile('F:/JusikData/API/Analy_Technical/Step1/juga_'+jongname+'_'+buydate+'.xlsx') :
            dataset = pd.read_excel('F:/JusikData/API/Analy_Technical/Step1/juga_'+jongname+'_'+buydate+'.xlsx', engine='openpyxl')
        else:
            # 데이터를 두개로 나눠서 각각 넣어줘야 에러가 발생하지 않음
            try:
                # 데이터 1000봉 가져오기
                data = ChartIndex_cls.exe_ChartIndex(self, jongcode, buydate)
                time.sleep(2)
                buydate2 = data.iloc[0]['일자']
                data2 = ChartIndex_cls.exe_ChartIndex(self, jongcode, buydate2)
                
                # 데이터 프레임 타입 int로 바꾸기
                data = data.astype(int)
                data2 = data2.astype(int)
                
                # 데이터 프레임 합치기, 데이터 프레임 중복 제거하기
                dataset = data2.append(data, sort=False)
                dataset = dataset.drop_duplicates(['일자'])
                
            except:
                print('해당종목은 상장폐지되었습니다.')
            
            dataset.to_excel('F:/JusikData/API/Analy_Technical/Step1/juga_'+jongname+'_'+buydate+'.xlsx', index=False)

        #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
        #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
        #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
        #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
        #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
        
        # 최저선 데이터 가져오기
        if os.path.isfile('F:/JusikData/API/Analy_Technical/Step2/low_'+jongname+'_'+buydate+'.xlsx') :
            lowpoint_grape_data = pd.read_excel('F:/JusikData/API/Analy_Technical/Step2/low_'+jongname+'_'+buydate+'.xlsx', engine='openpyxl')
        else:
            try:
                lowpoint_grape_data = trendline_cls.exe_trendline(self, jongname, buydate, dataset, 10)
                lowpoint_grape_data.to_excel('F:/JusikData/API/Analy_Technical/Step2/low_'+jongname+'_'+buydate+'.xlsx', index=False)
            except:
                print('최저선 에러발생.')
            
        #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
        #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
        #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
        #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
        #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
        
        # 수급구하기
        if os.path.isfile('F:/JusikData/API/Analy_Technical/Step3/supply_'+jongname+'_'+buydate+'.xlsx') :
                supply_data = pd.read_excel('F:/JusikData/API/Analy_Technical/Step3/supply_'+jongname+'_'+buydate+'.xlsx', engine='openpyxl')
        else :
            # 끝날짜 구하기
            stdate = dataset.iloc[0]['일자']
            
            # 매수 수급 
            supply_data = t1702_cls.exe_t1702(self, jongcode, buydate, '0', stdate) 
            
            # 엑셀 저장
            supply_data.to_excel('F:/JusikData/API/Analy_Technical/Step3/supply_'+jongname+'_'+buydate+'.xlsx', index=False)
            
        #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
        #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
        #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
        #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
        #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
        
        # 주가의 추세 구하기
        if os.path.isfile('F:/JusikData/API/Analy_Technical/Step4/touch1_high'+jongname+'_'+buydate+'.xlsx') :
                highpoint_data_touch1 = pd.read_excel('F:/JusikData/API/Analy_Technical/Step4/touch1_high.'+jongname+'_'+buydate+'.xlsx', engine='openpyxl')
                lowpoint_data_touch1 = pd.read_excel('F:/JusikData/API/Analy_Technical/Step4/touch1_low.'+jongname+'_'+buydate+'.xlsx', engine='openpyxl')
        else :
            # 첫터치, 둘터치
            highpoint_data_touch1, lowpoint_data_touch1 = trends_cls.make_trends(self, jongname, buydate, 10)
            
            # 엑셀 저장
            highpoint_data_touch1.to_excel('F:/JusikData/API/Analy_Technical/Step4/touch1_high.'+jongname+'_'+buydate+'.xlsx', index=False)
            lowpoint_data_touch1.to_excel('F:/JusikData/API/Analy_Technical/Step4/touch1_low.'+jongname+'_'+buydate+'.xlsx', index=False)
        
        #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
        #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
        #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
        #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
        #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
        
        # 모든 데이터 합치기 : 종합데이터
        # 지수데이터 : highpoint_data_kp, lowpoint_data_kp, highpoint_data_kq, lowpoint_data_kq
        # 기준 : 최저선 데이터 : lowpoint_grape_data
        # 수급데이터 : supply_data
        
        # 1. 지수데이터 날짜의 특수문자 제거
        highpoint_data_kp['일자'] = highpoint_data_kp['일자'].str.replace(pat=r'/', repl= r'', regex=True)
        lowpoint_data_kp['일자'] = lowpoint_data_kp['일자'].str.replace(pat=r'/', repl= r'', regex=True)
        highpoint_data_kq['일자'] = highpoint_data_kq['일자'].str.replace(pat=r'/', repl= r'', regex=True)
        lowpoint_data_kq['일자'] = lowpoint_data_kq['일자'].str.replace(pat=r'/', repl= r'', regex=True)
        
        # 2. 지수데이터를 int 형식으로 변환
        highpoint_data_kp = highpoint_data_kp.astype(int)
        lowpoint_data_kp = lowpoint_data_kp.astype(int)
        highpoint_data_kq = highpoint_data_kq.astype(int)
        lowpoint_data_kq = lowpoint_data_kq.astype(int)

        # 3. 최저선 일자의 수급 확인
        #print(lowpoint_grape_data)
        arrList7 = []
        arrList8 = []
        arrList9 = []
        
        for i in lowpoint_grape_data.index:
            try:
                series3 = supply_data[supply_data['일자'] == lowpoint_grape_data.iloc[i]['둘터치']]
                둘터치1등 = series3.iloc[0]['개인']
                둘터치2등 = series3.iloc[0]['외국인']
                둘터치3등 = series3.iloc[0]['기관']
            except:
                둘터치1등 = '최저'
                둘터치2등 = '최저'
                둘터치3등 = '최저'
                
            arrList7.append(둘터치1등)
            arrList8.append(둘터치2등)
            arrList9.append(둘터치3등)
        
        lowpoint_grape_data['개인'] = arrList7
        lowpoint_grape_data['외국인'] = arrList8
        lowpoint_grape_data['기관'] = arrList9
        
        # 4. 최저선 닿는데에서 지수의 추세 확인하기
        kp_inclhigh_arr = []
        kp_incllow_arr = []
        kq_inclhigh_arr = []
        kq_incllow_arr = []
        
        for i in lowpoint_grape_data.index:
            if lowpoint_grape_data.iloc[i]['둘터치'] == '최저' :
                kp_data_high = highpoint_data_kp[highpoint_data_kp['일자'] <= lowpoint_grape_data.iloc[i]['일자']]
                kp_data_low = lowpoint_data_kp[lowpoint_data_kp['일자'] <= lowpoint_grape_data.iloc[i]['일자']]
                inclination_high = (kp_data_high.iloc[-1]['y값'] - kp_data_high.iloc[-2]['y값']) / (kp_data_high.iloc[-1]['x값'] - kp_data_high.iloc[-2]['x값']) 
                inclination_low = (kp_data_low.iloc[-1]['y값'] - kp_data_low.iloc[-2]['y값']) / (kp_data_low.iloc[-1]['x값'] - kp_data_low.iloc[-2]['x값']) 
                kp_inclhigh_arr.append(inclination_high)
                kp_incllow_arr.append(inclination_low)
                
                kq_data_high = highpoint_data_kq[highpoint_data_kq['일자'] <= lowpoint_grape_data.iloc[i]['일자']]
                kq_data_low = lowpoint_data_kq[lowpoint_data_kq['일자'] <= lowpoint_grape_data.iloc[i]['일자']]
                inclination_high2 = (kq_data_high.iloc[-1]['y값'] - kq_data_high.iloc[-2]['y값']) / (kq_data_high.iloc[-1]['x값'] - kq_data_high.iloc[-2]['x값']) 
                inclination_low2 = (kq_data_low.iloc[-1]['y값'] - kq_data_low.iloc[-2]['y값']) / (kq_data_low.iloc[-1]['x값'] - kq_data_low.iloc[-2]['x값']) 
                kq_inclhigh_arr.append(inclination_high2)
                kq_incllow_arr.append(inclination_low2)
            else :
                kp_data_high = highpoint_data_kp[highpoint_data_kp['일자'] <= lowpoint_grape_data.iloc[i]['둘터치']]
                kp_data_low = lowpoint_data_kp[lowpoint_data_kp['일자'] <= lowpoint_grape_data.iloc[i]['둘터치']]
                inclination_high = (kp_data_high.iloc[-1]['y값'] - kp_data_high.iloc[-2]['y값']) / (kp_data_high.iloc[-1]['x값'] - kp_data_high.iloc[-2]['x값']) 
                inclination_low = (kp_data_low.iloc[-1]['y값'] - kp_data_low.iloc[-2]['y값']) / (kp_data_low.iloc[-1]['x값'] - kp_data_low.iloc[-2]['x값']) 
                kp_inclhigh_arr.append(inclination_high)
                kp_incllow_arr.append(inclination_low)
                
                kq_data_high = highpoint_data_kq[highpoint_data_kq['일자'] <= lowpoint_grape_data.iloc[i]['둘터치']]
                kq_data_low = lowpoint_data_kq[lowpoint_data_kq['일자'] <= lowpoint_grape_data.iloc[i]['둘터치']]
                inclination_high2 = (kq_data_high.iloc[-1]['y값'] - kq_data_high.iloc[-2]['y값']) / (kq_data_high.iloc[-1]['x값'] - kq_data_high.iloc[-2]['x값']) 
                inclination_low2 = (kq_data_low.iloc[-1]['y값'] - kq_data_low.iloc[-2]['y값']) / (kq_data_low.iloc[-1]['x값'] - kq_data_low.iloc[-2]['x값']) 
                kq_inclhigh_arr.append(inclination_high2)
                kq_incllow_arr.append(inclination_low2)
            
        lowpoint_grape_data['kp상'] = kp_inclhigh_arr
        lowpoint_grape_data['kp하'] = kp_incllow_arr
        lowpoint_grape_data['kq상'] = kq_inclhigh_arr
        lowpoint_grape_data['kq하'] = kq_incllow_arr
        
        # 5. 최저선과 최저선 닿는데에서 주가의 추세 확인하기
        date_inclhigh_arr = []
        date_incllow_arr = []
        touch2_inclhigh_arr = []
        touch2_incllow_arr = []
        
        for i in lowpoint_grape_data.index:
            if lowpoint_grape_data.iloc[i]['둘터치'] == '최저' :
                kp_data_high = highpoint_data_kp[highpoint_data_kp['일자'] <= lowpoint_grape_data.iloc[i]['일자']]
                kp_data_low = lowpoint_data_kp[lowpoint_data_kp['일자'] <= lowpoint_grape_data.iloc[i]['일자']]
                inclination_high = (kp_data_high.iloc[-1]['y값'] - kp_data_high.iloc[-2]['y값']) / (kp_data_high.iloc[-1]['x값'] - kp_data_high.iloc[-2]['x값']) 
                inclination_low = (kp_data_low.iloc[-1]['y값'] - kp_data_low.iloc[-2]['y값']) / (kp_data_low.iloc[-1]['x값'] - kp_data_low.iloc[-2]['x값']) 
                date_inclhigh_arr.append(inclination_high)
                date_incllow_arr.append(inclination_low)
                touch2_inclhigh_arr.append('최저')
                touch2_incllow_arr.append('최저')
            else :
                kp_data_high = highpoint_data_kp[highpoint_data_kp['일자'] <= lowpoint_grape_data.iloc[i]['일자']]
                kp_data_low = lowpoint_data_kp[lowpoint_data_kp['일자'] <= lowpoint_grape_data.iloc[i]['일자']]
                inclination_high = (kp_data_high.iloc[-1]['y값'] - kp_data_high.iloc[-2]['y값']) / (kp_data_high.iloc[-1]['x값'] - kp_data_high.iloc[-2]['x값']) 
                inclination_low = (kp_data_low.iloc[-1]['y값'] - kp_data_low.iloc[-2]['y값']) / (kp_data_low.iloc[-1]['x값'] - kp_data_low.iloc[-2]['x값']) 
                date_inclhigh_arr.append(inclination_high)
                date_incllow_arr.append(inclination_low)
                
                kq_data_high = highpoint_data_kq[highpoint_data_kq['일자'] <= lowpoint_grape_data.iloc[i]['둘터치']]
                kq_data_low = lowpoint_data_kq[lowpoint_data_kq['일자'] <= lowpoint_grape_data.iloc[i]['둘터치']]
                inclination_high2 = (kq_data_high.iloc[-1]['y값'] - kq_data_high.iloc[-2]['y값']) / (kq_data_high.iloc[-1]['x값'] - kq_data_high.iloc[-2]['x값']) 
                inclination_low2 = (kq_data_low.iloc[-1]['y값'] - kq_data_low.iloc[-2]['y값']) / (kq_data_low.iloc[-1]['x값'] - kq_data_low.iloc[-2]['x값']) 
                touch2_inclhigh_arr.append(inclination_high2)
                touch2_incllow_arr.append(inclination_low2)
            
        lowpoint_grape_data['일자상'] = date_inclhigh_arr
        lowpoint_grape_data['일자하'] = date_incllow_arr
        lowpoint_grape_data['둘터치상'] = touch2_inclhigh_arr
        lowpoint_grape_data['둘터치하'] = touch2_incllow_arr
        
        # 종합을 엑셀로 저장
        lowpoint_grape_data.to_excel('F:/JusikData/API/Analy_Technical/종합/com_'+jongname+'_'+buydate+'.xlsx', index=False)
        
conn = test_cls()
conn.exe_test()

'''
# 3. 최저선 일자의 수급 확인(일자, 첫터치, 둘터치)
        #print(lowpoint_grape_data)
        arrList1 = []
        arrList2 = []
        arrList3 = []
        arrList4 = []
        arrList5 = []
        arrList6 = []
        arrList7 = []
        arrList8 = []
        arrList9 = []
        
        for i in lowpoint_grape_data.index:
            print(i)
            series = supply_data[supply_data['일자'] == lowpoint_grape_data.iloc[i]['일자']]
            del series['일자'] # 일자는 빼기
            series = series.transpose() # 행열 전환
            series.columns = ['수급'] # 컬럼명 지정하기
            series['일자랭크'] = series['수급'].rank(method = 'first', ascending = False) # 랭크달기
            일자1등 = series[series['일자랭크'] == 1.0].index[0]
            일자2등 = series[series['일자랭크'] == 2.0].index[0]
            일자3등 = series[series['일자랭크'] == 3.0].index[0]
            
            try:
                series2 = supply_data[supply_data['일자'] == lowpoint_grape_data.iloc[i]['첫터치']]
                del series2['일자'] # 일자는 빼기
                series2 = series2.transpose() # 행열 전환
                series2.columns = ['수급'] # 컬럼명 지정하기
                series2['첫터치랭크'] = series2['수급'].rank(method = 'first', ascending = False) # 랭크달기
                첫터치1등 = series2[series2['첫터치랭크'] == 1.0].index[0]
                첫터치2등 = series2[series2['첫터치랭크'] == 2.0].index[0]
                첫터치3등 = series2[series2['첫터치랭크'] == 3.0].index[0]
            except:
                첫터치1등 = '최저'
                첫터치2등 = '최저'
                첫터치3등 = '최저'
                
            try:
                series3 = supply_data[supply_data['일자'] == lowpoint_grape_data.iloc[i]['둘터치']]
                del series3['일자'] # 일자는 빼기
                series3 = series3.transpose() # 행열 전환
                series3.columns = ['수급'] # 컬럼명 지정하기
                series3['둘터치랭크'] = series3['수급'].rank(method = 'first', ascending = False) # 랭크달기
                둘터치1등 = series3[series3['둘터치랭크'] == 1.0].index[0]
                둘터치2등 = series3[series3['둘터치랭크'] == 2.0].index[0]
                둘터치3등 = series3[series3['둘터치랭크'] == 3.0].index[0]
            except:
                둘터치1등 = '최저'
                둘터치2등 = '최저'
                둘터치3등 = '최저'
                
                
            arrList1.append(일자1등)
            arrList2.append(일자2등)
            arrList3.append(일자3등)
            arrList4.append(첫터치1등)
            arrList5.append(첫터치2등)
            arrList6.append(첫터치3등)
            arrList7.append(둘터치1등)
            arrList8.append(둘터치2등)
            arrList9.append(둘터치3등)
        
        lowpoint_grape_data['일자1등'] = arrList1
        lowpoint_grape_data['일자2등'] = arrList2
        lowpoint_grape_data['일자3등'] = arrList3
        lowpoint_grape_data['첫터치1등'] = arrList4
        lowpoint_grape_data['첫터치2등'] = arrList5
        lowpoint_grape_data['첫터치3등'] = arrList6
        lowpoint_grape_data['둘터치1등'] = arrList7
        lowpoint_grape_data['둘터치2등'] = arrList8
        lowpoint_grape_data['둘터치3등'] = arrList9
'''