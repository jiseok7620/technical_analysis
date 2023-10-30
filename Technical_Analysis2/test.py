'''
리턴 값 받고 넘겨주고, 엑셀에 필요 데이터 추출
'''

from Billionaire.Technical_Analysis2.Step1 import step1_cls
from Billionaire.Technical_Analysis2.Step2 import ChartIndex_cls
from Billionaire.Technical_Analysis2.Step3_1 import trendline_cls
from Billionaire.Technical_Analysis2.Step3_2 import trendline_cls2
from Billionaire.Technical_Analysis2.Step3_3 import trendline_cls3
from Billionaire.Technical_Analysis2.Step4 import t1702_cls
from Billionaire.Technical_Analysis2.Step5 import t1921_cls
from Billionaire.Technical_Analysis2.ebest_server_login import login_cls
from Billionaire.Technical_Analysis2.test2 import test2_cls
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
        
        # 테스트 데이터 가져오기
        testdata = pd.read_excel('F:/JusikData/API/Analy_Technical/Step1/con1.xlsx', engine='openpyxl')
        
        for i in range(500) :
            # 난수 발생
            n_num = random.randint(0, len(testdata)-1)
            
            # 테스트
            n번째 = n_num
            #n번째 = 9169
            
            jongname = testdata.iloc[n번째]['종목명']
            jongcode = str(testdata.iloc[n번째]['종목코드'])
            
            # 코드에 0붙여주기
            if len(jongcode) < 6:
                while True:
                    jongcode = '0' + jongcode 
                    if len(jongcode) == 6:
                        break
                    
            buydate = str(testdata.iloc[n번째]['일자'])
            if buydate[4:6] == "10":
                enddate = str(int(buydate[0:4])+1) + "04" + buydate[6:8]
            elif buydate[4:6] == "11":
                enddate = str(int(buydate[0:4])+1) + "05" + buydate[6:8]
            elif buydate[4:6] == "12":
                enddate = str(int(buydate[0:4])+1) + "06" + buydate[6:8]
            elif buydate[4:6] == "7":
                enddate = str(int(buydate[0:4])+1) + "01" + buydate[6:8]
            elif buydate[4:6] == "8":
                enddate = str(int(buydate[0:4])+1) + "02" + buydate[6:8]
            elif buydate[4:6] == "9":
                enddate = str(int(buydate[0:4])+1) + "03" + buydate[6:8]
            else :
                enddate = str(int(buydate) + 600)
            
            long = 20
            short = 10
            print(n번째)
            print(jongname)
            print(buydate)
            print(enddate)
    
            #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
            #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
            #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
            #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
            #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
            
            time.sleep(1)
            # Step2 수정주가데이터 가져오기
            if os.path.isfile('F:/JusikData/API/Analy_Technical/Step2/juga_'+jongname+'_'+buydate+'.xlsx') :
                data = pd.read_excel('F:/JusikData/API/Analy_Technical/Step2/juga_'+jongname+'_'+buydate+'.xlsx', engine='openpyxl')
            else:
                # 데이터를 두개로 나눠서 각각 넣어줘야 에러가 발생하지 않음
                try:
                    data = ChartIndex_cls.exe_ChartIndex(self, jongcode, enddate)
                    data = data.astype(int)
                except:
                    print('해당종목은 상장폐지되었습니다.')
                    #sys.exit('종료')
                    continue
                data.to_excel('F:/JusikData/API/Analy_Technical/Step2/juga_'+jongname+'_'+buydate+'.xlsx', index=False)
    
            #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
            #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
            #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
            #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
            #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
            
            # Step3 매물대, 추세선 데이터 가져오기
            if os.path.isfile('F:/JusikData/API/Analy_Technical/Step3/volpro_'+jongname+'_'+buydate+'.xlsx') :
                data_volpro = pd.read_excel('F:/JusikData/API/Analy_Technical/Step3/volpro_'+jongname+'_'+buydate+'.xlsx', engine='openpyxl')
                data_trends = pd.read_excel('F:/JusikData/API/Analy_Technical/Step3/trends_'+jongname+'_'+buydate+'.xlsx', engine='openpyxl')
            else :
                try:
                    data_volpro, data_trends = trendline_cls.exe_trendline(self, jongname, buydate, data, long, short)
                    #data_volpro, data_trends = trendline_cls2.exe_trendline2(self, jongname, enddate, data, long, short)
                    #data_volpro, data_trends = trendline_cls3.exe_trendline3(self, jongname, enddate, data, long, short)
                except:
                   print('에러발생')
                   continue
                #data_volpro.to_excel('F:/JusikData/API/Analy_Technical/Step3/volpro_'+jongname+'_'+buydate+'.xlsx', index=False)
                #data_trends.to_excel('F:/JusikData/API/Analy_Technical/Step3/trends_'+jongname+'_'+buydate+'.xlsx', index=False)
                
            #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
            #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
            #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
            #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
            #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
    
            time.sleep(5)
            
conn = test_cls()
conn.exe_test()