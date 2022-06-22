import sys
import pandas as pd
from PyQt5.QtWidgets import *
import win32com.client
# from sqlalchemy import create_engine
 
# 설명: 당일 상승률 상위 200 종목을 가져와 현재가  실시간 조회하는 샘플

# STEP 2: MySQL Connection 연결
# db_connection_str = 'mysql+pymysql://root:dkrkwk18!@localhost/stock'
# db_connection = create_engine(db_connection_str)
# conn = db_connection.connect()
 
# CpEvent: 실시간 이벤트 수신 클래스
class CpEvent:
    def set_params(self, client):
        self.client = client
 
    def OnReceived(self):
        code = self.client.GetHeaderValue(0)  # 초
        name = self.client.GetHeaderValue(1)  # 초
        timess = self.client.GetHeaderValue(18)  # 초
        exFlag = self.client.GetHeaderValue(19)  # 예상체결 플래그
        cprice = self.client.GetHeaderValue(13)  # 현재가
        diff = self.client.GetHeaderValue(2)  # 대비
        cVol = self.client.GetHeaderValue(17)  # 순간체결수량
        vol = self.client.GetHeaderValue(9)  # 거래량
 
        if (exFlag == ord('1')):  # 동시호가 시간 (예상체결)
            print("실시간(예상체결)", name, timess, "*", cprice, "대비", diff, "체결량", cVol, "거래량", vol)
        elif (exFlag == ord('2')):  # 장중(체결)
            print("실시간(장중 체결)", name, timess, cprice, "대비", diff, "체결량", cVol, "거래량", vol)
 
# CpStockCur: 실시간 현재가 요청 클래스
class CpStockCur:
    def Subscribe(self, code):
        self.objStockCur = win32com.client.Dispatch("DsCbo1.StockCur")
        # 주식 시세 데이터를 수신합니다. 
        handler = win32com.client.WithEvents(self.objStockCur, CpEvent)
        self.objStockCur.SetInputValue(0, code)
        handler.set_params(self.objStockCur)
        self.objStockCur.Subscribe()
 
    def Unsubscribe(self):
        self.objStockCur.Unsubscribe()
 
 
# Cp7043 상승률 상위 요청 클래스 - 연속 조회를 통해 200 종목 가져옴
class Cp7043:
    def __init__(self, howmany):
        # 통신 OBJECT 기본 세팅
        self.objRq = win32com.client.Dispatch("CpSysDib.CpSvrNew7043")
        self.objRq.SetInputValue(0, ord('0')) # 거래소 + 코스닥
        self.objRq.SetInputValue(1, ord('2'))  # 상승
        self.objRq.SetInputValue(2, ord('1'))  # 당일
        self.objRq.SetInputValue(3, 21)  # 전일 대비 상위 순
        self.objRq.SetInputValue(4, ord('1'))  # 관리 종목 제외
        self.objRq.SetInputValue(5, ord('0'))  # 거래량 전체
        self.objRq.SetInputValue(6, ord('0'))  # '표시 항목 선택 - '0': 시가대비
        self.objRq.SetInputValue(7, 0)  #  등락율 시작
        self.objRq.SetInputValue(8, 30)  # 등락율 끝
        self.howmany=howmany

    # 실제적인 7043 통신 처리
    def rq7043(self, codes ):
        self.objRq.BlockRequest()
        # 현재가 통신 및 통신 에러 처리
        rqStatus = self.objRq.GetDibStatus()
        rqRet = self.objRq.GetDibMsg1()
        # print("통신상태", rqStatus, rqRet)
        if rqStatus != 0:
            return False
 
        cnt = self.objRq.GetHeaderValue(0) #종목수
        cntTotal  = self.objRq.GetHeaderValue(1)
        # print(cnt, cntTotal)
        
        df_stock = pd.DataFrame(columns=['STOCK_CODE', 'STOCK_NAME', 'OPEN_PRICE', 'UP_DOWN_RATIO',
                                         'DIFF', 'VOLUME', 'DIFFFLAG'])
        
        
        # cnt = self.objRq.GetHeaderValue(cntidx)
        # data = []
        # for i in range(cnt):
        #     dict_item = {k: obj.GetDataValue(j, cnt-1-i) for j, k in data_fields.items()}
        #     dict_item.append(codes)
        #     data.append(dict_item)
        # print(data)
        
        
        
        for i in range(cnt):
            code = self.objRq.GetDataValue(0, i)  # 코드            
            codes.append(code)            
            if len(codes) >  self.howmany:       # 최대 5 종목만,
                break
            name = self.objRq.GetDataValue(1, i)  # 종목명
            now = self.objRq.GetDataValue(2, i) #현재가
            diffflag = self.objRq.GetDataValue(3, i)
            diff = self.objRq.GetDataValue(4, i)
            ratio = self.objRq.GetDataValue(5, i) #상승폭, 등략률
            vol = self.objRq.GetDataValue(6, i)  # 거래량
            
            # new_data = [code, name, now, round(ratio,2), diff, vol, diffflag]
            # df_stock.loc[len(df_stock)] = new_data
    
            new_data = {'STOCK_CODE' : code,
                        'STOCK_NAME' : name,
                        'OPEN_PRICE' : now,
                        'UP_DOWN_RATIO' : round(ratio,2),
                        'DIFF' : diff,
                        'VOLUME' : vol,
                        'DIFFFLAG' : diffflag
                        }
            df_stock = df_stock.append(new_data, ignore_index=True)
        print(df_stock)
        # df_stock.to_sqlw(name='stock_info', con=db_connection, if_exists='replace', index=False)

        
            # print(f'{code}, {name}, {now}, {round(ratio,2)}%, {diff}, {vol}, {diffflag}')
            # print('='*70)
        

    def Request(self, codes):
        self.rq7043(codes)
        # 연속 데이터 조회 - 5 개까지만.
        # while self.objRq.Continue:
        #     self.rq7043(retCode)
        #     # print(f'연속 데이터 조회 개수: {len(retCode)}')
        #     if len(retCode) > self.howmany:
        #         break
 
        # #7043 상승하락 서비스를 통해 받은 상승률 상위 200 종목
        # size = len(retCode)
        # for i in range(size):
        #     print(retCode[i])
        # return True
 
 
# # CpMarketEye : 복수종목 현재가 통신 서비스 - 200 종목 현재가를 조회 함
# class CpMarketEye:
#     def Request(self, codes, rqField):
#         # 연결 여부 체크
#         objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
#         bConnect = objCpCybos.IsConnect
#         if (bConnect == 0):
#             print("PLUS가 정상적으로 연결되지 않음. ")
#             return False
 
#         # 관심종목 객체 구하기
#         objRq = win32com.client.Dispatch("CpSysDib.MarketEye")
#         # 요청 필드 세팅 - 종목코드, 종목명, 시간, 대비부호, 대비, 현재가, 거래량
#         # rqField = [0,17, 1,2,3,4,10]
#         objRq.SetInputValue(0, rqField) # 요청 필드
#         objRq.SetInputValue(1, codes)  # 종목코드 or 종목코드 리스트
#         objRq.BlockRequest()
 
  
#         # 현재가 통신 및 통신 에러 처리
#         rqStatus = objRq.GetDibStatus()
#         rqRet = objRq.GetDibMsg1()
#         print("통신상태", rqStatus, rqRet)
#         if rqStatus != 0:
#             return False
 
#         cnt  = objRq.GetHeaderValue(2)
 
#         for i in range(cnt):
#             rpCode = objRq.GetDataValue(0, i)  # 코드
#             rpName = objRq.GetDataValue(1, i)  # 종목명
#             rpTime= objRq.GetDataValue(2, i)  # 시간
#             rpDiffFlag = objRq.GetDataValue(3, i)  # 대비부호
#             rpDiff = objRq.GetDataValue(4, i)  # 대비
#             rpCur = objRq.GetDataValue(5, i)  # 현재가
#             rpVol = objRq.GetDataValue(6, i)  # 거래량
#             print(rpCode, rpName, rpTime,  rpDiffFlag, rpDiff, rpCur, rpVol)
 
#         return True
 
 

# CpEvent: 실시간 이벤트 수신 클래스
# CpStockCur: 실시간 현재가 요청 클래스


# Cp7043 상승률 상위 요청 클래스 - 연속 조회를 통해 200 종목 가져옴
# CpMarketEye : 복수종목 현재가 통신 서비스 - 200 종목 현재가를 조회 함
  


def btnStart_clicked(howmany):    
    codes = []
    obj7043 = Cp7043(howmany)
    if obj7043.Request(codes) == False:
        return

    print(f"상승종목 개수: {howmany}")

    # # 요청 필드 배열 - 종목코드, 시간, 대비부호 대비, 현재가, 거래량, 종목명
    # rqField = [0, 1, 2, 3, 4, 10, 17]  #요청 필드
    # objMarkeyeye = CpMarketEye()
    # if (objMarkeyeye.Request(codes, rqField) == False):
    #     exit()

    # objCur = []

    # cnt = len(codes)
    # for i in range(cnt):
    #     objCur.append(CpStockCur())
    #     objCur[i].Subscribe(codes[i])

    # print(cnt , "종목 실시간 현재가 요청 시작")
    # isSB = True







if __name__ == "__main__":
    btnStart_clicked(30)
    
    
    

