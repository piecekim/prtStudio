import win32com.client
import time
import pandas as pd
import logging
import psycopg2
import pymysql

########################################## 추세_투자신탁 ##########################################

logger = logging.getLogger()
logger.setLevel(logging.DEBUG)
 
# 콘솔 출력을 지정합니다
ch = logging.StreamHandler()
ch.setLevel(logging.DEBUG)
 
#DB 연동
conn = psycopg2.connect(host='13.124.90.197', dbname='prt_db', user='prt_db', password='1937', port='5432')
cur = conn.cursor()
 
# # 연결 여부 체크
# objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
# bConnect = objCpCybos.IsConnect
# if (bConnect == 0):
#     print("PLUS가 정상적으로 연결되지 않음. ")
#     exit()

# 추세 투자신탁 테이블에 데이터를 insert 하기 위한 함수
def TodaySellAndBuy(code, inDate): 
     # SQL문 실행
    sql =   " select 투신"
    sql +=  " from prt_studio.svr7254 ss"
    sql +=  " where 1=1"
    sql +=  " and ss.종목코드='" + code + "'"
    sql +=  " and ss.일자='" + inDate + "'"
    cur.execute(sql)

    # 데이터 Fetch
    datas = []
    data = []
    datas = cur.fetchall()
    for data in datas:
        # print(data)

        매수매도량 = 0
        누적합계 = 0
        매집고점 = 0
        매집수량 = 0
        매집저점 = 0
        분산비율 = 0
        추세5일 = 0
        추세20일 = 0
        추세60일 = 0
        추세120일 = 0
        추세240일 = 0

        # 매수매도량 산출
        매수매도량 = data[0]

        # 매집고점, 매집수량, 매집저점 조회

        # 누적합계 산출
        sqlSum = "select "
        sqlSum += "    AA.매집수량 "
        sqlSum += "    , AA.매집고점 " 
        sqlSum += "    , AA.매집저점 "
        sqlSum += "from "
        sqlSum += "("
        sqlSum += "    select "
        sqlSum += "        A.*"
        sqlSum += "    from"
        sqlSum += "        prt_studio.추세_투자신탁 A"
        sqlSum += "    where 1=1"
        sqlSum += "    and A.종목코드 = '" + code + "'"
        sqlSum += "    and A.일자 < '" + inDate + "'"
        sqlSum += "    order by"
        sqlSum += "        A.일자 desc"
        sqlSum += ") AA"
        sqlSum += " LIMIT 1"
        cur.execute(sqlSum)

        # 데이터 Fetch
        dataSum = cur.fetchall()

        if len(dataSum) != 0 :
            매집수량 = dataSum[0][0] + data[0]
            if dataSum[0][1] < 매집수량 :
                매집고점 = 매집수량
            else :
                매집고점 = dataSum[0][1]

            if dataSum[0][2] < 매집수량 :
                매집저점 = dataSum[0][2]
            else :
                매집저점 = 매집수량
            누적합계 = 매집수량 + 매집저점
            if 매집고점 == 0:
                분산비율 = 0
            else :
                분산비율 = (매집수량 / 매집고점) * 100
        else :
            누적합계 = data[0]
            매집고점 = data[0]
            매집수량 = data[0]
            매집저점 = data[0]
            if 매집고점 == 0:
                분산비율 = 0
            else :
                분산비율 = (매집수량 / 매집고점) * 100
            # if dataSum[0] is not None :
            #     print(dataSum[0])
            # if dataSum[1] is not None :
            #     print(dataSum[1])
            # if dataSum[2] is not None :            
            #     print(dataSum[2])
        
        분산비율 = round(분산비율,2)

        sql5Avg =  "select"
        sql5Avg += "    coalesce(AVG(AA.매집수량),0) as 추세5일"
        sql5Avg += "    from "
        sql5Avg += "    ("
        sql5Avg += "        select "
        sql5Avg += "            A.*"
        sql5Avg += "        from"
        sql5Avg += "            prt_studio.추세_투자신탁 A"
        sql5Avg += "        where 1=1"
        sql5Avg += "        and A.종목코드 = '" + code + "'"
        sql5Avg += "        and A.일자 < '" + inDate + "'"
        sql5Avg += "        order by"
        sql5Avg += "            A.일자 desc"
        sql5Avg += "    ) AA"
        sql5Avg += "    LIMIT 5"
        
        cur.execute(sql5Avg)

        # 데이터 Fetch
        data5Avg = cur.fetchall()

        if len(data5Avg) != 0 :
            추세5일 = data5Avg[0][0]
        else :
            추세5일 = 0

        sql20Avg =  "select"
        sql20Avg += "    coalesce(AVG(AA.매집수량),0) as 추세20일"
        sql20Avg += "    from "
        sql20Avg += "    ("
        sql20Avg += "        select "
        sql20Avg += "            A.*"
        sql20Avg += "        from"
        sql20Avg += "            prt_studio.추세_투자신탁 A"
        sql20Avg += "        where 1=1"
        sql20Avg += "        and A.종목코드 = '" + code + "'"
        sql20Avg += "        and A.일자 < '" + inDate + "'"
        sql20Avg += "        order by"
        sql20Avg += "            A.일자 desc"
        sql20Avg += "    ) AA"
        sql20Avg += "    LIMIT 20"
        
        cur.execute(sql20Avg)

        # 데이터 Fetch
        data20Avg = cur.fetchall()

        if len(data20Avg) != 0 :
            추세20일 = data20Avg[0][0]
        else :
            추세20일 = 0

        sql60Avg =  "select"
        sql60Avg += "    coalesce(AVG(AA.매집수량),0) as 추세60일"
        sql60Avg += "    from "
        sql60Avg += "    ("
        sql60Avg += "        select "
        sql60Avg += "            A.*"
        sql60Avg += "        from"
        sql60Avg += "            prt_studio.추세_투자신탁 A"
        sql60Avg += "        where 1=1"
        sql60Avg += "        and A.종목코드 = '" + code + "'"
        sql60Avg += "        and A.일자 < '" + inDate + "'"
        sql60Avg += "        order by"
        sql60Avg += "            A.일자 desc"
        sql60Avg += "    ) AA"
        sql60Avg += "    LIMIT 60"
        
        cur.execute(sql60Avg)

        # 데이터 Fetch
        data60Avg = cur.fetchall()

        if len(data60Avg) != 0 :
            추세60일 = data60Avg[0][0]
        else :
            추세60일 = 0

        sql120Avg =  "select"
        sql120Avg += "    coalesce(AVG(AA.매집수량),0) as 추세120일"
        sql120Avg += "    from "
        sql120Avg += "    ("
        sql120Avg += "        select "
        sql120Avg += "            A.*"
        sql120Avg += "        from"
        sql120Avg += "            prt_studio.추세_투자신탁 A"
        sql120Avg += "        where 1=1"
        sql120Avg += "        and A.종목코드 = '" + code + "'"
        sql120Avg += "        and A.일자 < '" + inDate + "'"
        sql120Avg += "        order by"
        sql120Avg += "            A.일자 desc"
        sql120Avg += "    ) AA"
        sql120Avg += "    LIMIT 120"
        
        cur.execute(sql120Avg)

        # 데이터 Fetch
        data120Avg = cur.fetchall()

        if len(data120Avg) != 0 :
            추세120일 = data120Avg[0][0]
        else :
            추세120일 = 0

        sql240Avg =  "select"
        sql240Avg += "    coalesce(AVG(AA.매집수량),0) as 추세240일"
        sql240Avg += "    from "
        sql240Avg += "    ("
        sql240Avg += "        select "
        sql240Avg += "            A.*"
        sql240Avg += "        from"
        sql240Avg += "            prt_studio.추세_투자신탁 A"
        sql240Avg += "        where 1=1"
        sql240Avg += "        and A.종목코드 = '" + code + "'"
        sql240Avg += "        and A.일자 < '" + inDate + "'"
        sql240Avg += "        order by"
        sql240Avg += "            A.일자 desc"
        sql240Avg += "    ) AA"
        sql240Avg += "    LIMIT 240"
        
        cur.execute(sql240Avg)

        # 데이터 Fetch
        data240Avg = cur.fetchall()

        if len(data240Avg) != 0 :
            추세240일 = data240Avg[0][0]
        else :
            추세240일 = 0

        sqlInsert = "insert into prt_studio.추세_투자신탁 ("
        sqlInsert += "    종목코드"
        sqlInsert += "    , 일자"
        sqlInsert += "    , 매수매도량"
        sqlInsert += "    , 누적합계"
        sqlInsert += "    , 매집고점"
        sqlInsert += "    , 매집수량"
        sqlInsert += "    , 매집저점"
        sqlInsert += "    , 분산비율"
        sqlInsert += "    , 추세5일"
        sqlInsert += "    , 추세20일"
        sqlInsert += "    , 추세60일"
        sqlInsert += "    , 추세120일"
        sqlInsert += "    , 추세240일"
        sqlInsert += ") values ('"
        sqlInsert += code + "','" # 종목코드 
        sqlInsert += inDate + "'," # 일자
        sqlInsert += str(매수매도량)+ "," # 매수매도량
        sqlInsert += str(누적합계)+ "," # 누적합계
        sqlInsert += str(매집고점)+ "," # 매집고점
        sqlInsert += str(매집수량)+ "," # 매집수량
        sqlInsert += str(매집저점)+ "," # 매집저점
        sqlInsert += str(분산비율)+ "," # 분산비율
        sqlInsert += str(round(추세5일, 2))+ "," # 추세5일
        sqlInsert += str(round(추세20일, 2))+ "," # 추세20일
        sqlInsert += str(round(추세60일, 2))+ "," # 추세60일
        sqlInsert += str(round(추세120일, 2))+ "," # 추세120일
        sqlInsert += str(round(추세240일, 2)) # 추세240일
        sqlInsert += ")"

        # print(sqlInsert)

        cur.execute(sqlInsert)
        
        conn.commit()
    # conn.commit()

# svr7254 테이블에서 기준 데이터(코드,일자)를 가지고 오는 함수
def getCodeDateData(code):
    # SQL문 실행
    sql =  "select"
    sql += "    종목코드"
    sql += "    , to_char(일자, 'yyyy-mm-dd') 일자"
    sql += " from"
    sql += "    prt_studio.svr7254"
    sql += " where 1=1"
    # sql += " and 일자 between '20040101' and '20041231' "
    sql += " and 종목코드 = '" + code+ "'"
    sql += " group by "
    sql += "    종목코드, 일자 "
    sql += " order by "
    sql += "    일자 "
    # print(sql)
    
    cur.execute(sql)
    datas = cur.fetchall()
    for data in datas:
        TodaySellAndBuy(data[0], data[1])

sql =  "select 종목코드 from prt_studio.종목코드"

cur.execute(sql)
datas = cur.fetchall()
for data in datas:
    getCodeDateData(data[0])
# TodaySellAndBuy('A005930','20160106')
