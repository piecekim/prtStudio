import win32com.client
import time
import pandas as pd
import logging
 
logger = logging.getLogger()
logger.setLevel(logging.DEBUG)
 
# 콘솔 출력을 지정합니다
ch = logging.StreamHandler()
ch.setLevel(logging.DEBUG)
 
# logger.addHandler(ch)
 
# 파일 출력을 지정합니다.
fh = logging.FileHandler(filename="logging_test.log")
fh.setLevel(logging.INFO)
 
# add ch to logger
logger.addHandler(ch)
logger.addHandler(fh)
 
# 연결 여부 체크
objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
bConnect = objCpCybos.IsConnect
if (bConnect == 0):
    print("PLUS가 정상적으로 연결되지 않음. ")
    exit()
 
# def itemcode(m_code):
 
#     # 종목코드 리스트 구하기
#     objCpCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")
#     codeList = objCpCodeMgr.GetStockListByMarket(1) #거래소
#     codeList2 = objCpCodeMgr.GetStockListByMarket(2) #코스닥
    
#     # print("거래소 종목코드", len(codeList))
#     for i, code in enumerate(codeList):
#         secondCode = objCpCodeMgr.GetStockSectionKind(code)  # code에 해당하는 부구분코드
#         name = objCpCodeMgr.CodeToName(code)
#         stdPrice = objCpCodeMgr.GetStockStdPrice(code)  #기준가
#         if(code == m_code):
#             print("* 거래소 종목")
#             print(i, code, secondCode, name, stdPrice)
#             return name
    
#     # print("코스닥 종목코드", len(codeList2))
#     for i, code in enumerate(codeList2):
#         secondCode = objCpCodeMgr.GetStockSectionKind(code)
#         name = objCpCodeMgr.CodeToName(code)
#         stdPrice = objCpCodeMgr.GetStockStdPrice(code)
#         if(code == m_code):
#             print("* 코스닥 종목")
#             print(i, code, secondCode, name, stdPrice)
#             return code
 
#     # print("거래소 + 코스닥 종목코드 ",len(codeList) + len(codeList2))
 
def itemcode(m_name):
 
    # 종목코드 리스트 구하기
    objCpCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")
    codeList = objCpCodeMgr.GetStockListByMarket(1) #거래소
    codeList2 = objCpCodeMgr.GetStockListByMarket(2) #코스닥
 
    itemname = m_name
    
    # print("거래소 종목코드", len(codeList))
    for i, code in enumerate(codeList):
        name = objCpCodeMgr.CodeToName(code)
        if(name == itemname):
            return code
 
    for i, code in enumerate(codeList2):
        name = objCpCodeMgr.CodeToName(code)
        if(name == itemname):
            return code
 
 
def inquiry(m_code,m_start,m_end):
    ## 대신 API 세팅
    cpSvr7254 = win32com.client.Dispatch("CpSysDib.CpSvr7254")
    cpSvr7254.SetInputValue(0, m_code)       # 종목코드
    cpSvr7254.SetInputValue(1, '0')          # 기간선택 0:기간선택, 1:1개월, ... , 4:6개월
    cpSvr7254.SetInputValue(2, m_start)          # 시작일자: 기간선택구분을 0이 아닐경우 생략
    cpSvr7254.SetInputValue(3, m_end)          # 끝일자: 기간선택구분을 0이 아닐경우 생략
    cpSvr7254.SetInputValue(4, '0')         # 0:순매수 1:비중
    cpSvr7254.SetInputValue(5, '0')         # 투자자
    cpSvr7254.BlockRequest()                # 요청
 
    Num = cpSvr7254.GetHeaderValue(1)
    
    print("종목코드 : ",m_code)
    
    # while cpSvr7254.Continue:
    # cpSvr7254.BlockRequest()
    # Num = cpSvr7254.GetHeaderValue(1)
    # print(Num)
    
    for i in range(Num):
        print("-----------------------------")
        print("일자:", cpSvr7254.GetDataValue(0, i))
        print("개인:", cpSvr7254.GetDataValue(1, i))
        print("외국인: ", cpSvr7254.GetDataValue(2, i))
        print("기관계: ", cpSvr7254.GetDataValue(3, i))
        print("금융투자: ", cpSvr7254.GetDataValue(4, i))
        print("보험: ", cpSvr7254.GetDataValue(5, i))
        print("투신: ", cpSvr7254.GetDataValue(6, i))
        print("은행: ", cpSvr7254.GetDataValue(7, i))
        print("기타금융: ", cpSvr7254.GetDataValue(8, i))
        print("연기금 등: ", cpSvr7254.GetDataValue(9, i))
        print("기타법인: ", cpSvr7254.GetDataValue(10, i))
        print("기타외인: ", cpSvr7254.GetDataValue(11, i))  
        print("사모펀드: ", cpSvr7254.GetDataValue(12, i))  
        print("국가지자체: ", cpSvr7254.GetDataValue(13, i))  
        print("종가: ", cpSvr7254.GetDataValue(14, i))  
        print("대비: ", cpSvr7254.GetDataValue(15, i)) 
        print("대비율: ", cpSvr7254.GetDataValue(16, i)) 
        print("거래량: ", cpSvr7254.GetDataValue(17, i))  

 
# code = itemcode("A027050")
 
inquiry('A005930',20031225,20040105)
