from PyQt5.QAxContainer import * #키움 api를 실행해서 제어할려면 pyqt를 사용해야 한다. 
from PyQt5.QtCore import *
from matplotlib.style import available #이벤트 루프를 사용하기 위한(동시성 처리를 위해)
# from config.errorCode import *
# from config.kiwoomtype import *
# from config.log_class import *
import pandas as pd
import time
import pymysql
from sqlalchemy import create_engine
import pandas as pd
import sys
from PyQt5.QtWidgets import *
import win32com.client








class CreonAPI(QAxWidget):   # OpenAPI+가 제공하는 메서드를 호출하려면 QAxWidget 클래스의 인스턴스가 필요

cpCodeMgr = win32com.client.Dispatch('CpUtil.CpStockCode')
'''CYBOS에서사용되는주식코드조회작업을함.'''
cpStatus = win32com.client.Dispatch('CpUtil.CpCybos') 
'''각종 상태 반환 (매수, 매도, 조회)'''
cpTradeUtil = win32com.client.Dispatch('CpTrade.CpTdUtil')
'''설명 :주문오브젝트를사용하기위해필요한초기화과정들을수행한다

모든주문오브젝트는사용하기전에, 필수적으로 TradeInit을호출한후에사용할수있다.
전역변수(글로벌변수) 로선언하여사용하여야합니다.'''

cpStock = win32com.client.Dispatch('DsCbo1.StockMst')
	
'''주식종목의 현재가에 관련된 데이터(10차 호가 포함)'''
cpOhlc = win32com.client.Dispatch('CpSysDib.StockChart')
	
'''주식, 업종, ELW의차트데이터를수신합니다.'''
cpBalance = win32com.client.Dispatch('CpTrade.CpTd6033')
"""설명: 계좌별잔고및주문체결평가현황데이터를요청하고수신한다"""
cpCash = win32com.client.Dispatch('CpTrade.CpTdNew5331A')
"""설명: 계좌별매수주문가능금액/수량데이터를요청하고수신한다"""

cpOrder = win32com.client.Dispatch('CpTrade.CpTd0311')  
'''설명: 장내주식/코스닥주식/ELW주문(현금주문) 데이터를요청하고수신한다'''    
    
    
    
    
    def __init__(self):
        
        self._create_kiwoom_instance()
        self._set_signal_slots()
        self.comm_connect()
        
        # 이 두함수는 클라스가 실행되면 무조건 먼저 실행되야 하는 함수 
        
        
        
        # self.on_receive_opw00001()
        self.get_account_info()
        self.detail_account_info()
        
    
    def Cp_util(self):
        objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
        bConnect = objCpCybos.IsConnect
        if (bConnect == 0):
            print("PLUS가 정상적으로 연결되지 않음. ")
            exit()
            
            
            
            
            
    def cp7043(self): #상한가 인풋값
        
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




    # CpEvent: 실시간 이벤트 수신 클래스
    def cpEvent(self, client):
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
        handler = win32com.client.WithEvents(self.objStockCur, cpEvent)
        self.objStockCur.SetInputValue(0, code)
        handler.set_params(self.objStockCur)
        self.objStockCur.Subscribe()
 
    def Unsubscribe(self):
        self.objStockCur.Unsubscribe()








class Ui_class():
    def __init__(self):
        print('UI 클래스입니다.')
        
        
        # self.app = QApplication(sys.argv)
        self.kw = CreonAPI() #키움 클래스를 불러옴
        # self.app.exec_()





class Main(): #메인 클래스
    def __init__(self):
        print('실행할 메인 클래스')
        
        
        Ui_class()  #UI클래스를 불러옴
        
    
    
    
if __name__ == "__main__":
   Main()