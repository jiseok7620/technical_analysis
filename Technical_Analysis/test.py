'''
리턴 값 받고 넘겨주고, 엑셀에 필요 데이터 추출
'''

from Billionaire.Technical_Analysis.Step1 import step1_cls
from Billionaire.Technical_Analysis.Step2 import ChartIndex_cls
from Billionaire.Technical_Analysis.Step3 import trendline_cls
from Billionaire.Technical_Analysis.Step4 import t1702_cls
from Billionaire.Technical_Analysis.Step5 import t1921_cls
from Billionaire.Technical_Analysis.ebest_server_login import login_cls
from Billionaire.Technical_Analysis.test2 import test2_cls
from openpyxl.drawing.image import Image
from pandas.io.excel._base import read_excel
import pandas as pd
import time
import os
import openpyxl
import sys
import random

class test_cls:
    def exe_test(self):
        # 서버로그인
        login_cls.exe_login(self)
        
        #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
        #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
        #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
        #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
        #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
        
        # Step1 매수신호 파악하기(거래량 신호)
        # step1_cls.exe_step1(self, con)
        
        # Step1에서 나온 종목명, 종목코드, 매수일자
        '''
        jongname = 'AJ네트웍스'
        jongcode = '095570'
        buydate = '20191220'
        long = 60
        short = 20
        '''
        
        # 테스트 데이터 가져오기
        testdata = pd.read_excel('F:/JusikData/analysis_csv/step/step2/증가감소테스트(1).xlsx', engine='openpyxl')
        
        for i in range(200) :
            # 난수 발생
            n_num = random.randint(0, len(testdata)-1)
            
            # 테스트
            n번째 = n_num
            #n번째 = 2923
            
            jongname = testdata.iloc[n번째]['종목명']
            path1 = "F:/JusikData/oneday_csv/onedaydata/"+jongname+'/'+jongname+'.csv'
            codedata = pd.read_csv(path1, encoding='cp949')
            jongcode = str(codedata[codedata['종목명'] == jongname]['종목코드'][0])
            
            # 코드에 0붙여주기
            if len(jongcode) < 6:
                while True:
                    jongcode = '0' + jongcode 
                    if len(jongcode) == 6:
                        break
                    
            buydate = str(testdata.iloc[n번째]['일자'])
            long = 60
            short = 20
            inc_20 = testdata.iloc[n번째]['20일최대증가율']
            inc_60 = testdata.iloc[n번째]['60일최대증가율']
            inc_120 = testdata.iloc[n번째]['120일최대증가율']
            dec_20 = testdata.iloc[n번째]['20일최대감소율']
            dec_60 = testdata.iloc[n번째]['60일최대감소율']
            dec_120 = testdata.iloc[n번째]['120일최대감소율']
            
            print(n번째)
            print(jongname)
            print(buydate)
    
            #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
            #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
            #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
            #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
            #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
            
            time.sleep(1)
            # Step2 수정주가데이터 가져오기
            if os.path.isfile('F:/JusikData/API/Analy_Technical/Step2/juga_'+jongname+'_'+buydate+'.xlsx') :
                data = pd.read_excel('F:/JusikData/API/Analy_Technical/Step2/juga_'+jongname+'_'+buydate+'.xlsx', engine='openpyxl')
                data1 = data.astype(int)
                data2 = data.astype(int)
            else:
                # 데이터를 두개로 나눠서 각각 넣어줘야 에러가 발생하지 않음
                try:
                    data = ChartIndex_cls.exe_ChartIndex(self, jongcode, buydate)
                except:
                    print('해당종목은 상장폐지되었습니다.')
                    #sys.exit('종료')
                    continue
                data1 = data.astype(int)
                data2 = data.astype(int)
                data.to_excel('F:/JusikData/API/Analy_Technical/Step2/juga_'+jongname+'_'+buydate+'.xlsx', index=False)
    
            #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
            #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
            #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
            #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
            #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
            
            # Step3 매물대, 추세선 데이터 가져오기
            if os.path.isfile('F:/JusikData/API/Analy_Technical/Step3/volpro_'+jongname+'_'+buydate+'.xlsx') :
                data_volpro = pd.read_excel('F:/JusikData/API/Analy_Technical/Step3/volpro_'+jongname+'_'+buydate+'.xlsx', engine='openpyxl')
                data_acquisition = pd.read_excel('F:/JusikData/API/Analy_Technical/Step3/acquisition_'+jongname+'_'+buydate+'.xlsx', engine='openpyxl')
                data_trends = pd.read_excel('F:/JusikData/API/Analy_Technical/Step3/trends_'+jongname+'_'+buydate+'.xlsx', engine='openpyxl')
                data_inclination = pd.read_excel('F:/JusikData/API/Analy_Technical/Step3/inclination_'+jongname+'_'+buydate+'.xlsx', engine='openpyxl')
            else :
                try:
                    data_volpro, data_acquisition, data_trends, inclination_high_long, inclination_low_long, inclination_high_short, inclination_low_short,  = trendline_cls.exe_trendline(self, jongname, buydate, data1, long, short)
                except:
                    continue
                data_inclination = pd.DataFrame({'장기울기상':[inclination_high_long],
                                                     '장기울기하':[inclination_low_long],
                                                     '단기울기상':[inclination_high_short],
                                                     '단기울기하':[inclination_low_short]})
                data_volpro.to_excel('F:/JusikData/API/Analy_Technical/Step3/volpro_'+jongname+'_'+buydate+'.xlsx', index=False)
                data_acquisition.to_excel('F:/JusikData/API/Analy_Technical/Step3/acquisition_'+jongname+'_'+buydate+'.xlsx', index=False)
                data_trends.to_excel('F:/JusikData/API/Analy_Technical/Step3/trends_'+jongname+'_'+buydate+'.xlsx', index=False)
                data_inclination.to_excel('F:/JusikData/API/Analy_Technical/Step3/inclination_'+jongname+'_'+buydate+'.xlsx', index=False)
            #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
            #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
            #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
            #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
            #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
    
            # Step4 수급 구하기 => 개인도 1. 개미 2. 세력 으로 나눠야 할 듯
            '''
            # 1. 현재봉(매수시점봉)의 수급 정보 = 이 시점이 세력이 진짜로 올리는 시점인지 파악
            # 2. 매물대 부근의 수급 정보 = 매물대(저항선)에 물린 세력이 누군지 파악하기
            # 3. 매집봉 자체의 수급 정보 = 매집을 하고있는 주체가 누군지 파악하기
            # 4. 최신 매물대 첫번째 ~ 현재, 두번째 ~ 현재 까지의 수급 = 위 정보를 토대로 수급의 주체 파악하기
            # 매물대 가장 과거의 20일 전부터 현재까지의 수급데이터 모두 구하기
            인덱스    일자    분류        세부분류     거래량    수급정보...
                              일반        구분자
                              매물대        -
                              매집봉      구분자
            '''
            if os.path.isfile('F:/JusikData/API/Analy_Technical/Step4/buy_'+jongname+'_'+buydate+'.xlsx') :
                supply_buy_data = pd.read_excel('F:/JusikData/API/Analy_Technical/Step4/buy_'+jongname+'_'+buydate+'.xlsx', engine='openpyxl')
                supply_sell_data = pd.read_excel('F:/JusikData/API/Analy_Technical/Step4/sell_'+jongname+'_'+buydate+'.xlsx', engine='openpyxl')
                supply_sum_data = pd.read_excel('F:/JusikData/API/Analy_Technical/Step4/sum_'+jongname+'_'+buydate+'.xlsx', engine='openpyxl')
            else :
                if data_volpro.empty:
                    stdate = data.iloc[0]['일자']
                else:
                    stdate = int(data_volpro.iloc[0]['일자'])
                # 매수 수급 
                supply_buy_data = t1702_cls.exe_t1702(self, jongcode, buydate, '1', stdate) # 1 = 매수, 2 = 매도\
                time.sleep(1) # 여기에서도 1초 쉬어줘야함
                # 매도 수급
                supply_sell_data = t1702_cls.exe_t1702(self, jongcode, buydate, '2', stdate)
                # 순매수 수급
                supply_sum_data = supply_buy_data + supply_sell_data
                supply_sum_data['일자'] = supply_sell_data['일자']
                # 엑셀 저장
                supply_buy_data.to_excel('F:/JusikData/API/Analy_Technical/Step4/buy_'+jongname+'_'+buydate+'.xlsx', index=False)
                supply_sell_data.to_excel('F:/JusikData/API/Analy_Technical/Step4/sell_'+jongname+'_'+buydate+'.xlsx', index=False)
                supply_sum_data.to_excel('F:/JusikData/API/Analy_Technical/Step4/sum_'+jongname+'_'+buydate+'.xlsx', index=False)
                
            #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
            #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
            #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
            #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
            #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
            
            # Step5 해당 일자로 부터 n일 전 까지의 신용 융자 현황 확인하기
            #if os.path.isfile('F:/JusikData/API/Analy_Technical/Step5/credit_'+jongname+'_'+buydate+'.xlsx') :
            #    credit_loan_data = pd.read_excel('F:/JusikData/API/Analy_Technical/Step5/credit_'+jongname+'_'+buydate+'.xlsx', engine='openpyxl')
            #else:
            #    credit_loan_data = t1921_cls.exe_t1921(self, jongcode, buydate)
            #    credit_loan_data.to_excel('F:/JusikData/API/Analy_Technical/Step5/credit_'+jongname+'_'+buydate+'.xlsx', index=False)
                
            #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
            #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
            #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
            #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
            #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#    
            
            # 엑셀에 정리하기
            try:
                test2_cls.exe_test2(self, str(jongname), str(buydate), inc_20, inc_60, inc_120, dec_20, dec_60, dec_120)
            except:
                print('엑셀정리 에러발생')
                continue
            
            time.sleep(120)
            
conn = test_cls()
conn.exe_test()