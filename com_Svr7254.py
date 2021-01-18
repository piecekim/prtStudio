import win32com.client
import time
import pandas as pd
import logging
import psycopg2


logger = logging.getLogger()
logger.setLevel(logging.DEBUG)
 
# 콘솔 출력을 지정합니다
ch = logging.StreamHandler()
ch.setLevel(logging.DEBUG)
 
#DB 연동
conn = psycopg2.connect(host='13.124.90.197', dbname='prt_db', user='prt_db', password='1937', port='5432')
cur = conn.cursor()
 
# 연결 여부 체크
objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
bConnect = objCpCybos.IsConnect
if (bConnect == 0):
    print("PLUS가 정상적으로 연결되지 않음. ")
    exit()
 
def inquiry(m_start,m_end):

    cur.execute("SELECT 종목코드 FROM 종목코드")
    cr = cur.fetchall()
    
    for i in cr:
        print(i[0])
        cpSvr7254 = win32com.client.Dispatch("CpSysDib.CpSvr7254")
        cpSvr7254.SetInputValue(0, i[0])       # 종목코드
        cpSvr7254.SetInputValue(1, '0')          # 기간선택 0:기간선택, 1:1개월, ... , 4:6개월
        cpSvr7254.SetInputValue(2, m_start)          # 시작일자: 기간선택구분을 0이 아닐경우 생략
        cpSvr7254.SetInputValue(3, m_end)          # 끝일자: 기간선택구분을 0이 아닐경우 생략
        cpSvr7254.SetInputValue(4, '0')         # 0:순매수 1:비중
        cpSvr7254.SetInputValue(5, '0')         # 투자자
        cpSvr7254.BlockRequest()                # 요청
        
    # ## 대신 API 세팅
    # cpSvr7254 = win32com.client.Dispatch("CpSysDib.CpSvr7254")
    # cpSvr7254.SetInputValue(0, cr)       # 종목코드
    # cpSvr7254.SetInputValue(1, '0')          # 기간선택 0:기간선택, 1:1개월, ... , 4:6개월
    # cpSvr7254.SetInputValue(2, m_start)          # 시작일자: 기간선택구분을 0이 아닐경우 생략
    # cpSvr7254.SetInputValue(3, m_end)          # 끝일자: 기간선택구분을 0이 아닐경우 생략
    # cpSvr7254.SetInputValue(4, '0')         # 0:순매수 1:비중
    # cpSvr7254.SetInputValue(5, '0')         # 투자자
    # cpSvr7254.BlockRequest()                # 요청
 
        Num = cpSvr7254.GetHeaderValue(1)
        # print(Num)
        # print("종목코드 : ",i[0])
        
        for j in range(Num):
            # print("-----------------------------")
            sql = "insert into svr7254 ("
            sql += "    종목코드"
            sql += "    , 일자"
            sql += "    , 개인"
            sql += "    , 외국인"
            sql += "    , 기관계"
            sql += "    , 금융투자"
            sql += "    , 보험"
            sql += "    , 투신"
            sql += "    , 은행"
            sql += "    , 기타금융"
            sql += "    , 연기금"
            sql += "    , 기타법인"
            sql += "    , 기타외인"
            sql += "    , 사모펀드"
            sql += "    , 국가지자체"
            sql += "    , 종가"
            sql += "    , 대비"
            sql += "    , 대비율"
            sql += "    , 거래량"
            sql += ") values ('"
            sql += i[0] + "','" # 종목코드 
            sql += str(cpSvr7254.GetDataValue(0, j))+ "'," # 일자
            sql += str(cpSvr7254.GetDataValue(1, j))+ "," # 개인
            sql += str(cpSvr7254.GetDataValue(2, j))+ "," # 외국인
            sql += str(cpSvr7254.GetDataValue(3, j))+ "," # 기관계
            sql += str(cpSvr7254.GetDataValue(4, j))+ "," # 금융투자
            sql += str(cpSvr7254.GetDataValue(5, j))+ "," # 보험
            sql += str(cpSvr7254.GetDataValue(6, j))+ "," # 투신
            sql += str(cpSvr7254.GetDataValue(7, j))+ "," # 은행
            sql += str(cpSvr7254.GetDataValue(8, j))+ "," # 기타금융
            sql += str(cpSvr7254.GetDataValue(9, j))+ "," # 연기금
            sql += str(cpSvr7254.GetDataValue(10, j))+ "," # 기타법인
            sql += str(cpSvr7254.GetDataValue(11, j))+ "," # 기타외인
            sql += str(cpSvr7254.GetDataValue(12, j))+ "," # 사모펀드
            sql += str(cpSvr7254.GetDataValue(13, j))+ "," # 국가지자체
            sql += str(cpSvr7254.GetDataValue(14, j))+ "," # 종가
            sql += str(cpSvr7254.GetDataValue(15, j))+ "," # 대비
            sql += str(cpSvr7254.GetDataValue(16, j))+ "," # 대비율
            sql += str(cpSvr7254.GetDataValue(17, j)) # 거래량

            sql += ")"
            cur.execute(sql)
            
            conn.commit()

 
# code = itemcode("A027050")
 
inquiry(20170103,20170108)

