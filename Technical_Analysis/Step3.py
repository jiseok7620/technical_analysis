'''
Pro3. 추세선, 매물대, 매물대 해소 표시하기

'''

import pandas as pd
import numpy as np
import math
import time
import os
import mpl_finance
import matplotlib.pyplot as plt
from cmath import nan
from PIL._imaging import display
from numpy import round

class trendline_cls:
    def exe_trendline(self, name, edate, dataset, long, short):
        ##--------------------------------------------------------------------------------------------------##
        ## 고점, 저점 구하기 ##
        highpoint_bf = [] # 고점보다 작은값들의 수 리스트(당일 이전)
        highpoint_af = [] # 고점보다 작은값들의 수 리스트(당일 이후)
        highpoint_sum = [] # 고점보다 작은값들의 수 합
        highpoint_count_bf = 0 # 고점보다 작은값들의 수(당일 이전)
        highpoint_count_af = 0 # 고점보다 작은값들의 수(당일 이후)
        
        lowpoint_bf = [] # 저점보다 큰값들의 수 리스트(당일 이전)
        lowpoint_af = [] # 저점보다 큰값들의 수 리스트(당일 이후)
        lowpoint_sum = [] # 저점보다 큰값들의 수 합
        lowpoint_count_bf = 0 # 저점보다 큰값들의 수(당일 이전)
        lowpoint_count_af = 0 # 저점보다 큰값들의 수(당일 이후)
        
        # for문 돌리기
        for i in dataset.index:
            # i : 해당 일자의 인덱스
            # 해당 일자 전의 데이터 분석
            for j in range(i) : 
                # 첫 날은 무시
                if i == 0:
                    break
                # 당일의 고가보다 전일의 고가가 더 큰 값이 있으면 for문 빠져나가기(<--'당일' 방향으로 분석)
                if dataset.iloc[i]['고가'] <= dataset.iloc[i-j-1]['고가']:
                    break
                # 당일의 고가보다 전일의 고가가 작으면 개수 세기(<--'당일' 방향으로 분석)
                elif dataset.iloc[i]['고가'] > dataset.iloc[i-j-1]['고가']:
                    highpoint_count_bf += 1
            
            # 해당 일자 이후의 데이터 분석
            for k in range(i,len(dataset)):
                # 마지막 날은 무시
                if i == len(dataset)-1:
                    break
                # 당일 값은 무시
                if i == k :
                    continue
                # 당일의 고가보다 후일의 고가가 더 큰 값이 있으면 for문 빠져나가기('당일'--> 방향으로 분석)
                if dataset.iloc[i]['고가'] <= dataset.iloc[k]['고가']:
                    break
                # 당일의 고가보다 후일의 고가가 작으면 개수 세기('당일'--> 방향으로 분석)
                elif dataset.iloc[i]['고가'] > dataset.iloc[k]['고가']:
                    highpoint_count_af += 1
                    
            # 저점 구하기
            for l in range(i) : 
                if i == 0:
                    break
                if dataset.iloc[i]['저가'] >= dataset.iloc[i-l-1]['저가']:
                    break
                elif dataset.iloc[i]['저가'] < dataset.iloc[i-l-1]['저가']:
                    lowpoint_count_bf += 1
                
            for m in range(i,len(dataset)):
                if i == len(dataset)-1:
                    break
                if i == m :
                    continue
                if dataset.iloc[i]['저가'] >= dataset.iloc[m]['저가']:
                    break
                elif dataset.iloc[i]['저가'] < dataset.iloc[m]['저가']:
                    lowpoint_count_af += 1
            
            # 고점 배열 추가
            highpoint_bf.append(highpoint_count_bf)
            highpoint_af.append(highpoint_count_af)
            highpoint_sum.append(highpoint_count_bf+highpoint_count_af)
            highpoint_count_bf = 0
            highpoint_count_af = 0
            
            # 저점 배열 추가            
            lowpoint_bf.append(lowpoint_count_bf)
            lowpoint_af.append(lowpoint_count_af)
            lowpoint_sum.append(lowpoint_count_bf+lowpoint_count_af) # 다른 방법 : sum_max_column = [bf_max_column[i] + af_max_column[i] for i in range(len(bf_max_column))]
            lowpoint_count_bf = 0
            lowpoint_count_af = 0
            
        # 결과
        highpoint_data = pd.DataFrame({'전수' : highpoint_bf, '후수' : highpoint_af, '총합' : highpoint_sum})
        lowpoint_data = pd.DataFrame({'전수' : lowpoint_bf, '후수' : lowpoint_af, '총합' : lowpoint_sum})
        #print(highpoint_data)
        
        # 내림차순 정렬
        highpoint_data = highpoint_data.sort_values('총합', ascending=False)
        lowpoint_data = lowpoint_data.sort_values('총합', ascending=False)
        
        # 전작은값수나 후작은값수가 n미만 인 것들은 필터
        # 장기 추세
        highpoint_data_long = highpoint_data[highpoint_data['전수'] >= long]
        highpoint_data_long = highpoint_data_long[highpoint_data_long['후수'] >= long]
        lowpoint_data_long = lowpoint_data[lowpoint_data['전수'] >= long]
        lowpoint_data_long = lowpoint_data_long[lowpoint_data_long['후수'] >= long]
        
        # 만약에 long 데이터가 두개도 없다면
        new_high_long = long
        new_low_long = long
        if len(highpoint_data_long) < 2 :
            while len(highpoint_data_long) < 2 :
                new_high_long -= 1
                highpoint_data_long = highpoint_data[highpoint_data['전수'] >= new_high_long]
                highpoint_data_long = highpoint_data_long[highpoint_data_long['후수'] >= new_high_long]
        if len(lowpoint_data_long) < 2 :
            while len(lowpoint_data_long) < 2 :
                new_low_long -= 1
                lowpoint_data_long = lowpoint_data[lowpoint_data['전수'] >= new_low_long]
                lowpoint_data_long = lowpoint_data_long[lowpoint_data_long['후수'] >= new_low_long]
        
        # 단기 추세 
        highpoint_data_short = highpoint_data[highpoint_data['전수'] >= short]
        highpoint_data_short = highpoint_data_short[highpoint_data_short['후수'] >= short]
        lowpoint_data_short = lowpoint_data[lowpoint_data['전수'] >= short]
        lowpoint_data_short = lowpoint_data_short[lowpoint_data_short['후수'] >= short]
        
        # 만약에 short 데이터가 두개도 없다면
        new_high_short = short
        new_low_short = short
        if len(highpoint_data_short) < 2 :
            while len(highpoint_data_short) < 2 :
                new_high_short -= 1
                highpoint_data_short = highpoint_data[highpoint_data['전수'] >= new_high_short]
                highpoint_data_short = highpoint_data_short[highpoint_data_short['후수'] >= new_high_short]
        if len(lowpoint_data_short) < 2 :
            while len(lowpoint_data_short) < 2 :
                new_low_short -= 1
                lowpoint_data_short = lowpoint_data[lowpoint_data['전수'] >= new_low_short]
                lowpoint_data_short = lowpoint_data_short[lowpoint_data_short['후수'] >= new_low_short]
        
        
        ##--------------------------------------------------------------------------------------------------##        
        ## 추세선을 그리기 위해서 일차함수 구하기 ##
        # 고점, 저점의 일자, 인덱스, 가격 가져오기
        highpoint_date_long = []
        highpoint_index_long = []
        highpoint_price_long = []
        
        lowpoint_date_long = []
        lowpoint_index_long = []
        lowpoint_price_long = []
        
        highpoint_date_short = []
        highpoint_index_short = []
        highpoint_price_short = []
        
        lowpoint_date_short = []
        lowpoint_index_short = []
        lowpoint_price_short = []
        
        for i in range(len(highpoint_data_long)) :
            # 인덱스 가져오기
            highidx = highpoint_data_long.index[i]
            
            # 해당 인덱스의 일자 가져오기
            highdate = dataset.loc[highidx]['일자']
            
            # 해당 인덱스의 고가 가져오기
            highprice = dataset.loc[highidx]['고가']
            
            # 리스트에 값 넣기
            highpoint_date_long.append(highdate)
            highpoint_index_long.append(highidx)
            highpoint_price_long.append(highprice)
        
        for i in range(len(lowpoint_data_long)) :
            # 인덱스 가져오기
            lowidx = lowpoint_data_long.index[i]
            
            # 해당 인덱스의 일자 가져오기
            lowdate = dataset.loc[lowidx]['일자']
            
            # 해당 인덱스의 고가 가져오기
            lowprice = dataset.loc[lowidx]['저가']
            
            # 리스트에 값 넣기
            lowpoint_date_long.append(lowdate)
            lowpoint_index_long.append(lowidx)
            lowpoint_price_long.append(lowprice)
            
        for i in range(len(highpoint_data_short)) :
            # 인덱스 가져오기
            highidx = highpoint_data_short.index[i]
            
            # 해당 인덱스의 일자 가져오기
            highdate = dataset.loc[highidx]['일자']
            
            # 해당 인덱스의 고가 가져오기
            highprice = dataset.loc[highidx]['고가']
            
            # 리스트에 값 넣기
            highpoint_date_short.append(highdate)
            highpoint_index_short.append(highidx)
            highpoint_price_short.append(highprice)
        
        for i in range(len(lowpoint_data_short)) :
            # 인덱스 가져오기
            lowidx = lowpoint_data_short.index[i]
            
            # 해당 인덱스의 일자 가져오기
            lowdate = dataset.loc[lowidx]['일자']
            
            # 해당 인덱스의 고가 가져오기
            lowprice = dataset.loc[lowidx]['저가']
            
            # 리스트에 값 넣기
            lowpoint_date_short.append(lowdate)
            lowpoint_index_short.append(lowidx)
            lowpoint_price_short.append(lowprice)
            
        # 결과 데이터프레임 만들기
        highpoint_grape_data_long = pd.DataFrame({'일자' : highpoint_date_long, 'x값' : highpoint_index_long, 'y값' : highpoint_price_long})
        lowpoint_grape_data_long = pd.DataFrame({'일자' : lowpoint_date_long, 'x값' : lowpoint_index_long, 'y값' : lowpoint_price_long})
        highpoint_grape_data_short = pd.DataFrame({'일자' : highpoint_date_short, 'x값' : highpoint_index_short, 'y값' : highpoint_price_short})
        lowpoint_grape_data_short = pd.DataFrame({'일자' : lowpoint_date_short, 'x값' : lowpoint_index_short, 'y값' : lowpoint_price_short})
        
        # 내림차순 정리
        highpoint_grape_data_long = highpoint_grape_data_long.sort_values('일자', ascending=False)
        lowpoint_grape_data_long = lowpoint_grape_data_long.sort_values('일자', ascending=False)
        highpoint_grape_data_short = highpoint_grape_data_short.sort_values('일자', ascending=False)
        lowpoint_grape_data_short = lowpoint_grape_data_short.sort_values('일자', ascending=False)
        
        # 최신값 두개를 가져와서 직선만들기 : x는 인덱스값, y는 고가
        # 첫 번째 값의 좌표
        x1_high_long = highpoint_grape_data_long.iloc[0]['x값']
        y1_high_long = highpoint_grape_data_long.iloc[0]['y값']
        
        x1_low_long = lowpoint_grape_data_long.iloc[0]['x값']
        y1_low_long = lowpoint_grape_data_long.iloc[0]['y값']
        
        x1_high_short = highpoint_grape_data_short.iloc[0]['x값']
        y1_high_short = highpoint_grape_data_short.iloc[0]['y값']
        
        x1_low_short = lowpoint_grape_data_short.iloc[0]['x값']
        y1_low_short = lowpoint_grape_data_short.iloc[0]['y값']
        
        # 두 번째 값의 좌표
        x2_high_long = highpoint_grape_data_long.iloc[1]['x값']
        y2_high_long = highpoint_grape_data_long.iloc[1]['y값']
        
        x2_low_long = lowpoint_grape_data_long.iloc[1]['x값']
        y2_low_long = lowpoint_grape_data_long.iloc[1]['y값']
        
        x2_high_short = highpoint_grape_data_short.iloc[1]['x값']
        y2_high_short = highpoint_grape_data_short.iloc[1]['y값']
        
        x2_low_short = lowpoint_grape_data_short.iloc[1]['x값']
        y2_low_short = lowpoint_grape_data_short.iloc[1]['y값']
        
        # 기울기
        inclination_high_long = (y1_high_long - y2_high_long) / (x1_high_long - x2_high_long) 
        inclination_low_long = (y1_low_long - y2_low_long) / (x1_low_long - x2_low_long) 
        inclination_high_short = (y1_high_short - y2_high_short) / (x1_high_short - x2_high_short) 
        inclination_low_short = (y1_low_short - y2_low_short) / (x1_low_short - x2_low_short) 
        
        # 직선 그래프 데이터 리스트를 만들기
        st_grape_high_long = []
        st_grape_low_long = []
        st_grape_high_short = []
        st_grape_low_short = []
        for i in dataset.index:
            ## 장기 추세선 고점
            if i >= x2_high_long :
                # 그래프그리기시작
                st_y = inclination_high_long * (i-x1_high_long) + y1_high_long
                st_grape_high_long.append(st_y)
            else : 
                # 그래프안그리는 곳 
                st_grape_high_long.append(y2_high_long)
            
            ## 장기 추세선 저점
            if i >= x2_low_long :
                # 그래프그리기시작
                st_y = inclination_low_long * (i-x1_low_long) + y1_low_long
                st_grape_low_long.append(st_y)
            else : 
                # 그래프안그리는 곳 
                st_grape_low_long.append(y2_low_long)
            
            ## 단기 추세선 고점
            if i >= x2_high_short :
                # 그래프그리기시작
                st_y = inclination_high_short * (i-x1_high_short) + y1_high_short
                st_grape_high_short.append(st_y)
            else : 
                # 그래프안그리는 곳 
                st_grape_high_short.append(y2_high_short)
            
            ## 단기 추세선 저점
            if i >= x2_low_short :
                # 그래프그리기시작
                st_y = inclination_low_short * (i-x1_low_short) + y1_low_short
                st_grape_low_short.append(st_y)
            else : 
                # 그래프안그리는 곳 
                st_grape_low_short.append(y2_low_short)
                
        # 결과를 전체 데이터에 넣기
        dataset['high_long'] = st_grape_high_long
        dataset['low_long'] = st_grape_low_long
        dataset['high_short'] = st_grape_high_short
        dataset['low_short'] = st_grape_low_short
        
        
        ##--------------------------------------------------------------------------------------------------##
        ## 매물대 구하기 ##
        # highpoint_grape_data_short # 장기 저점 데이터 사용
        
        # 개별 종목 데이터 가져오기
        path1 = "F:/JusikData/oneday_csv/onedaydata/"+name+'/'+name+'.csv'
        data_jusiksu = pd.read_csv(path1, encoding='cp949')
        
        if data_jusiksu.empty:
            print('해당 종목의 개별데이터가 없습니다.')
        
        arr_volpro_sec = []
        arr_dd_long = []
        arr_jusik_long = []
        arr_trade_long = []
        arr_high_long = []
        arr_per_long = []
        arr_idx_long = []
        
        ## 단기 고점 매물대 가져오기
        for i in highpoint_grape_data_short.index:
            # 구분자 넣기
            arr_volpro_sec.append(i)
            
            # 해당일자를 리스트에 담기
            date = highpoint_grape_data_short.iloc[i]['일자']
            arr_dd_long.append(date)
            
            # 해당일자의 유통주식수 구하기
            jusik = data_jusiksu[data_jusiksu['일자'] == highpoint_grape_data_short.iloc[i]['일자']]['상장주식수'].values[0]
            arr_jusik_long.append(jusik)
        
            # 해당일자의 인덱스
            idx_dd = highpoint_grape_data_short.iloc[i]['x값']
            
            # 해당일자 거래량 구하기
            dd_tr = dataset.iloc[idx_dd]['거래량']
            arr_trade_long.append(dd_tr)
            
            # 해당일자 고가 구하기
            dd_hi = highpoint_grape_data_short.iloc[i]['y값']
            arr_high_long.append(dd_hi)
            
            # 유통주식수 대비 거래량의 %
            jusik_tr = round(dd_tr / jusik * 100, 2)
            arr_per_long.append(jusik_tr)
            
            # 인덱스 넣기
            arr_idx_long.append(idx_dd)
        
        # 결과 데이터프레임 만들기
        data_sale_long = pd.DataFrame({'구분자' : arr_volpro_sec, '인덱스' : arr_idx_long, '일자' : arr_dd_long ,'매물대' : arr_high_long,  
                                     '거래량' : arr_trade_long, '유통주식수' : arr_jusik_long, '비율' : arr_per_long})
        
        
        
        ##-----------------------------------------------------------------------------------------------##
        ## 매물대 해소 판단 ##
        release_arr = []
        result_num = 0
        release_volume_one = [] # 매물대 보다 가격이 높았던 구간의 거래량 보기
        release_volume_all = [] # 매물대 보다 가격이 높았던 구간의 거래량 합
        
        for i in data_sale_long.index:
            idx = int(data_sale_long.iloc[i]['인덱스']) # 해당일자의 인덱스
            
            for j in range(idx+1,len(dataset)-1):
                if dataset.iloc[j]['고가'] >= data_sale_long.iloc[i]['매물대']:
                    vol = dataset.iloc[j]['거래량']
                    release_volume_one.append(vol)
                    result_num += 1
                
            # 이후에 매물대보다 고가가 있었다면 그 개수를 세어서 넣음
            if result_num != 0:
                release_arr.append(result_num)
            else:
                release_arr.append(0)
                  
            # 매물대보다 높았던 가격의 거래량 합을 리스트에 넣기
            lst_sum = sum(release_volume_one)
            release_volume_all.append(lst_sum)
            
            # 변수, 리스트 초기화
            result_num = 0
            release_volume_one = []
            
        #print(release_arr)
        data_sale_long['release'] = release_arr
        data_sale_long['release_vol'] = release_volume_all        
        
        
        ##-----------------------------------------------------------------------------------------------##
        ## 매집봉 확인 ##
        # 해소가 안된 봉의 정보를 가져오기
        data_not_resolved = data_sale_long[data_sale_long['release'] == 0]
        data_not_resolved = data_not_resolved.sort_values('인덱스', ascending=True) # 인덱스컬럼 기준 오름차순 정렬
        data_not_resolved = data_not_resolved.reset_index(drop=True) # 인덱스 초기화
        
        # 해소가 안된 봉 이후의 매집봉이 있는지 확인하기 
        acquisition_class = []
        acquisition_class_price = []
        acquisition_date = []
        acquisition_idx = []
        acquisition_high = []
        acquisition_per = []
        acquisition_vol = []
        acquisition_jusik = []
        for i in data_not_resolved.index:
            if i == len(data_not_resolved) - 1 :
                idx = data_not_resolved.iloc[i]['인덱스']
                
                for j in range(int(idx+1), len(dataset)-1):
                    if dataset.iloc[j]['거래량'] >= dataset.iloc[j-1]['거래량'] * 5 : # 전일 거래량의 5배인 봉을 찾아보기
                        acquisition_class.append(data_not_resolved.iloc[i]['구분자'])
                        acquisition_date.append(int(dataset.iloc[j]['일자']))
                        acquisition_idx.append(j)
                        acquisition_high.append(dataset.iloc[j]['고가'])
                        acquisition_class_price.append(data_not_resolved.iloc[i]['매물대'])
                        acquisition_vol.append(dataset.iloc[j]['거래량'])
                        # 해당일자의 유통주식수 구하기
                        jusik = data_jusiksu[data_jusiksu['일자'] == dataset.iloc[j]['일자']]['상장주식수'].values[0]
                        acquisition_jusik.append(jusik)
                        
                        # 매물대와 고가의 차이를 %로                        
                        per = (data_not_resolved.iloc[i]['매물대'] - dataset.iloc[j]['고가']) / data_not_resolved.iloc[i]['매물대'] * 100
                        acquisition_per.append(per)
            else :
                idx_a = data_not_resolved.iloc[i]['인덱스']
                idx_b = data_not_resolved.iloc[i+1]['인덱스']
    
                for j in range(int(idx_a+1),int(idx_b)):
                    if dataset.iloc[j]['거래량'] >= dataset.iloc[j-1]['거래량'] * 5 : # 전일 거래량의 5배인 봉을 찾아보기
                        acquisition_class.append(i) # 구분자
                        acquisition_date.append(dataset.iloc[j]['일자']) # 일자
                        acquisition_idx.append(j) # 인덱스
                        acquisition_high.append(dataset.iloc[j]['고가']) # 고가
                        acquisition_class_price.append(data_not_resolved.iloc[i]['매물대']) # 비교 매물대
                        acquisition_vol.append(dataset.iloc[j]['거래량'])
                        # 해당일자의 유통주식수 구하기
                        jusik = data_jusiksu[data_jusiksu['일자'] == dataset.iloc[j]['일자']]['상장주식수'].values[0]
                        acquisition_jusik.append(jusik)
                        
                        # 매물대와 고가의 차이를 %로
                        per = (data_not_resolved.iloc[i]['매물대'] - dataset.iloc[j]['고가']) / data_not_resolved.iloc[i]['매물대'] * 100
                        acquisition_per.append(per)
                        
        # 결과 데이터프레임 만들기
        data_acquisition = pd.DataFrame({'구분자' : acquisition_class, '일자' : acquisition_date, '매물대' : acquisition_class_price ,  
                                     '거래량' : acquisition_vol, '상장주식수' : acquisition_jusik,
                                     '고가' : acquisition_high, '매물대대비' : acquisition_per, '인덱스' : acquisition_idx
                                     })
        
        
        
        ##-----------------------------------------------------------------------------------------------##
        # 그래프 그리기 - 해소 안된 매물대 부터 현재까지
        # 표시할 데이터 개수 정하기
        if data_not_resolved.empty:
            start_idx = len(dataset)
        else:
            start_idx = int(data_not_resolved.iloc[0]['인덱스'])
        display_data = dataset[start_idx:]
        display_data = display_data.reset_index(drop=True)
        
        # 캔버스 설정(크기, 배경 설정 등)
        fig = plt.figure(figsize=(17,8)) ## 캔버스 생성
        fig.set_facecolor('white') ## 캔버스 색상 설정
        
        #차트1 : 지수-종목그래프 and 추세선
        ax1 = fig.add_subplot(211)
        ax1.plot(display_data['high_long'], color='red', linewidth=2)
        ax1.plot(display_data['low_long'], color='red', linewidth=2)
        ax1.plot(display_data['high_short'], color='blue', linewidth=1)
        ax1.plot(display_data['low_short'], color='blue', linewidth=1)         
        mpl_finance.candlestick2_ohlc(ax1, display_data['시가'], display_data['고가'], display_data['저가'], display_data['종가'], width=0.5, colorup='r', colordown='b')
        #plt.xticks(visible=False) # 축값없애기
        
        # 차트2 : 거래량차트
        # 거래량 바 차트
        ax2 = fig.add_subplot(212)
        plt.bar(range(len(display_data)), display_data['거래량'])
        
        # 수평선 긋기 = 매물대
        for i in data_sale_long.index:
            if data_sale_long.iloc[i]['release'] >= 1 and data_sale_long.iloc[i]['release_vol'] < data_sale_long.iloc[i]['매물대']:
                if data_sale_long.iloc[i]['비율'] >= 20:
                    color = 'red'
                elif data_sale_long.iloc[i]['비율'] >= 10:
                    color = 'blue'
                elif data_sale_long.iloc[i]['비율'] >= 5:
                    color = 'green'
                else:
                    color = 'violet'
                
                # 수평선의 시작과 끝을 지정하기 위함
                firstidx = int(data_sale_long.iloc[i]['인덱스']) - start_idx
                lastidx = len(display_data)                
                stidx = firstidx / lastidx 
                
                ax1.axhline(data_sale_long.iloc[i]['매물대'], stidx, 1, color=color, linestyle='--', linewidth=0.5) 
                
            elif data_sale_long.iloc[i]['release'] == 0 :
                if data_sale_long.iloc[i]['비율'] >= 20:
                    color = 'red'
                elif data_sale_long.iloc[i]['비율'] >= 10:
                    color = 'blue'
                elif data_sale_long.iloc[i]['비율'] >= 5:
                    color = 'green'
                else:
                    color = 'violet'
                
                # 수평선의 시작과 끝을 지정하기 위함 
                firstidx = int(data_sale_long.iloc[i]['인덱스']) - start_idx
                lastidx = len(display_data)                
                stidx = firstidx / lastidx 
                    
                ax1.axhline(data_sale_long.iloc[i]['매물대'], stidx, 1, color=color, linewidth=0.5)
                
        # 수직선 긋기 = 매집봉
        for i in data_acquisition.index:
            if data_acquisition.iloc[i]['구분자'] % 2 == 1:
                ax1.axvline(int(data_acquisition.iloc[i]['인덱스'])-start_idx, color='dodgerblue', linestyle='--', linewidth=0.5)        
            elif data_acquisition.iloc[i]['구분자'] % 2 == 0:
                ax1.axvline(int(data_acquisition.iloc[i]['인덱스'])-start_idx, color='limegreen', linestyle='--', linewidth=0.5)  
        
        # 그래프 보기
        #plt.show()
        
        # 저장
        plt.savefig('F:/JusikData/API/Analy_Technical/Step3/grape1_'+name+'_'+edate+'.png')
        
        # 메모리 제거
        plt.close()
        
        
        '''
        ##-----------------------------------------------------------------------------------------------##
        # 그래프 그리기2 - 짧은 그래프(그래프 크게 보기)
        # 해소 안된 첫번째, 두번째 에서부터 현재
        # 표시할 데이터 개수 정하기
        if len(data_not_resolved) < 2 :
            start_idx = int(data_not_resolved.iloc[-1]['인덱스'])
        else:
            start_idx = int(data_not_resolved.iloc[-2]['인덱스'])
        display_data = dataset[start_idx:]
        display_data = display_data.reset_index(drop=True)
        
        # 캔버스 설정(크기, 배경 설정 등)
        fig2 = plt.figure(figsize=(20,10)) ## 캔버스 생성
        fig2.set_facecolor('white') ## 캔버스 색상 설정
        
        #차트1 : 지수-종목그래프 and 추세선
        ax1 = fig2.add_subplot(211)
        ax1.plot(display_data['high_long'], color='red', linewidth=1)
        ax1.plot(display_data['low_long'], color='red', linewidth=1)
        ax1.plot(display_data['high_short'], color='blue', linewidth=1)
        ax1.plot(display_data['low_short'], color='blue', linewidth=1)         
        mpl_finance.candlestick2_ohlc(ax1, display_data['시가'], display_data['고가'], display_data['저가'], display_data['종가'], width=0.5, colorup='r', colordown='b')
        #plt.xticks(visible=False) # 축값없애기
        
        # 차트2 : 거래량차트
        # 거래량 바 차트
        ax2 = fig2.add_subplot(212)
        plt.bar(range(len(display_data)), display_data['거래량'])
        
        # 수평선 긋기 = 매물대 => 바로 위에 두개 매물대 정도만 표시
        # 해소가 안된 봉들만 가져오기
        data_sale_long = data_not_resolved.sort_values('매물대', ascending=True)
        for i in range(2):
            if len(data_sale_long) < i+1 :
                break
            
            if data_sale_long.iloc[i]['release'] >= 1 and data_sale_long.iloc[i]['release_vol'] < data_sale_long.iloc[i]['매물대']:
                color = 'red'
                ax1.axhline(data_sale_long.iloc[i]['매물대'], color=color, linestyle='--', linewidth=0.5) 
            else:
                if data_sale_long.iloc[i]['비율'] >= 20:
                    color = 'red'
                elif data_sale_long.iloc[i]['비율'] >= 10:
                    color = 'blue'
                elif data_sale_long.iloc[i]['비율'] >= 5:
                    color = 'green'
                else:
                    color = 'violet'
                    
                ax1.axhline(data_sale_long.iloc[i]['매물대'], color=color, linewidth=0.5)
                
        # 수직선 긋기 = 매집봉
        for i in data_acquisition.index:
            if data_acquisition.iloc[i]['구분자'] == len(data_not_resolved) - 1:
                ax1.axvline(int(data_acquisition.iloc[i]['인덱스'])-start_idx, color='dodgerblue', linestyle='--', linewidth=0.5)        
            elif data_acquisition.iloc[i]['구분자'] == len(data_not_resolved) - 2:
                ax1.axvline(int(data_acquisition.iloc[i]['인덱스'])-start_idx, color='limegreen', linestyle='--', linewidth=0.5) 
        
        
        
        # 그래프 보기
        #plt.show()
        
        # 저장
        plt.savefig('F:/JusikData/API/Analy_Technical/Step3/grape2_'+name+'_'+edate+'.png')
        
        # 메모리 제거
        plt.close()
        '''


        ##-----------------------------------------------------------------------------------------------##
        # 데이터프레임 리턴
        # 1. 매물대 = data_sale_long
        # 2. 매집봉 = data_acquisition
        return data_not_resolved, data_acquisition, dataset, inclination_high_long, inclination_low_long, inclination_high_short, inclination_low_short
    
'''
자체 모듈 테스트
conn = trendline_cls()
dataset =  pd.read_excel('F:/JusikData/API/data/'+'AJ네트웍스'+'_'+'20191220'+'.xlsx', engine='openpyxl')
conn.exe_trendline('AJ네트웍스', '20191220', dataset, 1)
'''