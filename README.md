# technical_analysis
국내 상장 기업 기술적 분석

# 실행 준비
1. make_folder_alldata.py 
   - (필수) http://data.krx.co.kr에서 전종목 시세 csv파일 다운로드
   - csv 파일에 하루 종목 전체 추가하기

# 내용
 - 유통주식수, 전일 거래량, 거래대금 기준으로 종목 추출
 - 수정주가를 이용하여 신호 파악
 - 추세선, 매물대, 매물대 해소 파악
 - 수급으로 투자 판단하기
 - Ebest API를 이용한 데이터 수신
 - 최저선을 파악하여 지지 파악하기
 - 고점과 저점을 전 후 봉수로 파악하여 지지, 저항선 구하기

# 향후 발전
 - KRX의 데이터 다운로드(oneday_csv) => FinanceDataReader 모듈로 대체
 - UI 필요 시 pyqt를 이용하여 제작
