import win32com.client
import time
import pandas as pd
import logging
import psycopg2
 
 
# 콘솔 출력을 지정합니다
ch = logging.StreamHandler()
ch.setLevel(logging.DEBUG)

#DB 연동
conn = psycopg2.connect(host='13.124.90.197', dbname='prt_db', user='prt_db', password='1937', port='5432')
cur = conn.cursor()
 
def subCpCodeMgr():
    # 연결 여부 체크
    # objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
    # bConnect = objCpCybos.IsConnect
    # if (bConnect == 0):
    #     print("PLUS가 정상적으로 연결되지 않음. ")
    #     exit()
    
    # 종목코드 리스트 구하기
    objCpCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")
    codeList = objCpCodeMgr.GetStockListByMarket(1) #거래소 0
    codeList2 = objCpCodeMgr.GetStockListByMarket(2) #코스닥 1

    # print("거래소 종목코드", len(codeList))
    for i, code in enumerate(codeList):
    
        secondCode = objCpCodeMgr.GetStockSectionKind(code) # code 에 해당하는 부 구분 코드를 반환한다
        name = objCpCodeMgr.CodeToName(code) #code 에 해당하는 주식/선물/옵션 종목명 을 반환한다
        stdPrice = objCpCodeMgr.GetStockStdPrice(code) #code 에 해당하는 권리락 등으로 인한 기준가를 반환한다
        stdStatu = objCpCodeMgr.GetStockStatusKind(code) #code 에 해당하는 주식상태를 반환한다

        # print(i, code, secondCode, name, stdPrice, stdStatu)
        
        cur.execute("SELECT 종목코드, 종목명 FROM 종목코드 WHERE 종목코드='%s' and 종목구분='0'" % code)
        cr = cur.fetchall()
        
        if(cr != []):
            # 종목명을 비교해서 비교 결과가 다를 경우 update 같을 경우 skip
            # cr 의 값 [('A123456', '테스트')]
            cur.execute("UPDATE 종목코드 SET 주식상태='%s' WHERE 종목코드='%s' and 종목구분='0'" % (stdStatu,code))
            if(cr[0][1] != name):
                cur.execute("UPDATE 종목코드 SET 종목명='%s' WHERE 종목코드='%s'" % (name, code))
        else:
            # 데이터가 없는 것이기 때문에 현재 코드값을 종목코드 테이블에 insert
            try:
                cur.execute("INSERT INTO 종목코드 (종목코드, 종목명, 종목구분) VALUES (%s, %s, %s)", (code, name, 0))

            except Exception as ex: # 에러 종류
                print('에러가 발생 했습니다', ex) # ex는 발생한 에러의 이름을 받아오는 변수
    

    # print("코스닥 종목코드", len(codeList2))
    for i, code in enumerate(codeList2):

        secondCode = objCpCodeMgr.GetStockSectionKind(code)
        name = objCpCodeMgr.CodeToName(code)
        stdPrice = objCpCodeMgr.GetStockStdPrice(code)
        stdStatu = objCpCodeMgr.GetStockStatusKind(code)
        
        #print(i, code, secondCode, stdPrice, name, stdStatu)
        cur.execute("SELECT 종목코드, 종목명 FROM 종목코드 WHERE 종목코드='%s' and 종목구분='1'" % code)
        cr = cur.fetchall()
        
        if(cr != []):
            # 종목명을 비교해서 비교 결과가 다를 경우 update 같을 경우 skip
            # cr 의 값 [('A123456', '테스트')]
            cur.execute("UPDATE 종목코드 SET 주식상태='%s' WHERE 종목코드='%s' and 종목구분='1'" % (stdStatu,code))
            if(cr[0][1] != name):
                cur.execute("UPDATE 종목코드 SET 종목명='%s' WHERE 종목코드='%s'" % (name, code))
        else:
            # 데이터가 없는 것이기 때문에 현재 코드값을 종목코드 테이블에 insert
            try:
                cur.execute("INSERT INTO 종목코드 (종목코드, 종목명, 종목구분) VALUES (%s, %s, %s)", (code, name, 1))

            except Exception as ex: # 에러 종류
                print('에러가 발생 했습니다', ex) # ex는 발생한 에러의 이름을 받아오는 변수

    conn.commit()   

subCpCodeMgr()
 
