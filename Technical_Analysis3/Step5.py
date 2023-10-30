from pandas.io.excel._base import read_excel
import pandas as pd
import time
import os
import openpyxl
import sys
import random

class trends_cls:
    def make_trends(self, jongname, buydate, long):
        # 데이터 가져오기
        dataset = pd.read_excel('F:/JusikData/API/Analy_Technical/Step1/juga_'+jongname+'_'+buydate+'.xlsx', engine='openpyxl')
        #print(dataset)
        
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
            lowpoint_sum.append(lowpoint_count_bf+lowpoint_count_af)
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
        
        ##--------------------------------------------------------------------------------------------------##        
        ## 추세선을 그리기 위해서 일차함수 구하기 ##
        # 고점, 저점의 일자, 인덱스, 가격 가져오기
        highpoint_date_long = []
        highpoint_index_long = []
        highpoint_price_long = []
        
        lowpoint_date_long = []
        lowpoint_index_long = []
        lowpoint_price_long = []
        
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
            
            # 해당 인덱스의 저가 가져오기
            lowprice = dataset.loc[lowidx]['저가']
            
            # 리스트에 값 넣기
            lowpoint_date_long.append(lowdate)
            lowpoint_index_long.append(lowidx)
            lowpoint_price_long.append(lowprice)
            
        # 결과 데이터프레임 만들기
        highpoint_grape_data_long = pd.DataFrame({'일자' : highpoint_date_long, 'x값' : highpoint_index_long, 'y값' : highpoint_price_long})
        lowpoint_grape_data_long = pd.DataFrame({'일자' : lowpoint_date_long, 'x값' : lowpoint_index_long, 'y값' : lowpoint_price_long})
        
        # 내림차순 정리
        highpoint_grape_data_long = highpoint_grape_data_long.sort_values('일자', ascending=True)
        lowpoint_grape_data_long = lowpoint_grape_data_long.sort_values('일자', ascending=True)

        #print(highpoint_grape_data_long)
        #print(lowpoint_grape_data_long)
        
        return highpoint_grape_data_long, lowpoint_grape_data_long

'''  
conn = jisu_cls()
conn.make_jisu('kospi', 20)
'''