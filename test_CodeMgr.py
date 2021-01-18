import win32com.client
import time
import pandas as pd
import logging
import psycopg2
 

def subCpCodeMgr():
    # 연결 여부 체크
    # objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
    # bConnect = objCpCybos.IsConnect
    # if (bConnect == 0):
    #     print("PLUS가 정상적으로 연결되지 않음. ")
    #     exit()
    
    # 종목코드 리스트 구하기
    objCpCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")
    codeList = objCpCodeMgr.GetStockListByMarket(1) #거래소
    codeList2 = objCpCodeMgr.GetStockListByMarket(2) #코스닥
    
    
    print("거래소 종목코드", len(codeList))
    for i, code in enumerate(codeList):
        secondCode = objCpCodeMgr.GetStockSectionKind(code) # code 에 해당하는 부 구분 코드를 반환한다
        name = objCpCodeMgr.CodeToName(code) #code 에 해당하는 주식/선물/옵션 종목명 을 반환한다
        stdPrice = objCpCodeMgr.GetStockStdPrice(code) #code 에 해당하는 권리락 등으로 인한 기준가를 반환한다
        stdStatu = objCpCodeMgr.GetStockStatusKind(code) #code 에 해당하는 주식상태를 반환한다

        print(i, code, secondCode, stdPrice, name, stdStatu)
    
    print("코스닥 종목코드", len(codeList2))
    for i, code in enumerate(codeList2):
        secondCode = objCpCodeMgr.GetStockSectionKind(code)
        name = objCpCodeMgr.CodeToName(code)
        stdPrice = objCpCodeMgr.GetStockStdPrice(code)
        stdStatu = objCpCodeMgr.GetStockStatusKind(code)
        print(i, code, secondCode, stdPrice, name, stdStatu)
    
    print("거래소 + 코스닥 종목코드 ",len(codeList) ,len(codeList2))

subCpCodeMgr()