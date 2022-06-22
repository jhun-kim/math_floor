# import win32com.client


# # 연결 여부 체크
# objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
# bConnect = objCpCybos.IsConnect
# if (bConnect == 0):
#     print("PLUS가 정상적으로 연결되지 않음. ")
#     exit()
 
 
# # 일자별 object 구하기
# objStockWeek = win32com.client.Dispatch("DsCbo1.StockWeek")
# objStockWeek.SetInputValue(0, 'A005930')   #종목 코드 - 삼성전자
 




 
# def ReqeustData(obj):
#     # 데이터 요청
#     obj.BlockRequest()
 
#     # 통신 결과 확인
#     rqStatus = obj.GetDibStatus()
#     rqRet = obj.GetDibMsg1()
#     print("통신상태", rqStatus, rqRet)
#     if rqStatus != 0:
#         return False
 
#     # 일자별 정보 데이터 처리
#     count = obj.GetHeaderValue(1)  # 데이터 개수
#     for i in range(count):
#         date = obj.GetDataValue(0, i)  # 일자
#         open = obj.GetDataValue(1, i)  # 시가
#         high = obj.GetDataValue(2, i)  # 고가
#         low = obj.GetDataValue(3, i)  # 저가
#         close = obj.GetDataValue(4, i)  # 종가
#         diff = obj.GetDataValue(5, i)  # 종가
#         vol = obj.GetDataValue(6, i)  # 종가
#         # print(date, open, high, low, close, diff, vol)
 
#     return date, open, high, low, close
 
# # 최초 데이터 요청
# ret = ReqeustData(objStockWeek)
# if ret == False:
#     exit()
 
 
# # 연속 데이터 요청
# # 예제는 5번만 연속 통신 하도록 함.
# NextCount = 1
# while objStockWeek.Continue:  #연속 조회처리
#     NextCount+=1;
#     if (NextCount > 5):
#         break
#     ret = ReqeustData(objStockWeek)
#     if ret == False:
#         exit()
 
 











import win32com.client
import sys
from PyQt5.QtWidgets import *
# import win32com.client
import requests
from slack import Post 



# def post_message(token, channel, text):
#         response = requests.post("https://slack.com/api/chat.postMessage",
#         headers={"Authorization": "Bearer "+token},
#         data={"channel": channel,"text": text})
        
 
# myToken = "xoxb-3162802854482-3643669615106-zrrUHSu6qTmcQ361pLUvPvou"



class Daily_ohlc():
    def __init__(self):
        # 연결 여부 체크
        objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
        bConnect = objCpCybos.IsConnect
        if (bConnect == 0):
            print("PLUS가 정상적으로 연결되지 않음. ")
            exit()
        
            
    
        
        
    def ohlc(self, code):    
        # 일자별 object 구하기
        self.objStockWeek = win32com.client.Dispatch("DsCbo1.StockWeek")
        self.objStockWeek.SetInputValue(0, code)   #종목 코드 - 삼성전자
        
        
        # 최초 데이터 요청
        ret = ReqeustData(self.objStockWeek)
        if ret == False:
            exit()
         # 연속 데이터 요청
    # 예제는 5번만 연속 통신 하도록 함.
        NextCount = 1
        while self.objStockWeek.Continue:  #연속 조회처리
            
            NextCount+=1;
            if (NextCount > 5):
                break
            ret = ReqeustData(self.objStockWeek)
            
            if ret == False:
                exit()
        

         
def ReqeustData(obj):
        # 데이터 요청
    obj.BlockRequest()

    # 통신 결과 확인
    rqStatus = obj.GetDibStatus()
    rqRet = obj.GetDibMsg1()
    print("통신상태", rqStatus, rqRet)
    if rqStatus != 0:
        return False

    # 일자별 정보 데이터 처리
    list = []
    count = obj.GetHeaderValue(1)  # 데이터 개수
    for i in range(count):
        info = []
        date = obj.GetDataValue(0, i)  # 일자
        open_ = obj.GetDataValue(1, i)  # 시가
        high = obj.GetDataValue(2, i)  # 고가
        low = obj.GetDataValue(3, i)  # 저가
        close = obj.GetDataValue(4, i)  # 종가
        diff = obj.GetDataValue(5, i)  # 종가
        vol = obj.GetDataValue(6, i)  # 종가
        
        info.append(date)
        info.append(open_)
        info.append(high)
        info.append(low)
        info.append(close)
        info.append(diff)
        info.append(vol)
        
        # list.append(info)
        # post_message(myToken,"#sj", f'{date}, {open_}, {high}, {low}, {close}, {diff}, {vol}')
        
        Post().post_message("#sj", f'{date}, {open_}, {high}, {low}, {close}, {diff}, {vol}')
        
        
    # print(list)
    # return list
        
# 일자별 object 구하기
# objStockWeek = win32com.client.Dispatch("DsCbo1.StockWeek")
# objStockWeek.SetInputValue(0, 'A005930')   #종목 코드 - 삼성전자
# abc = ReqeustData(objStockWeek)
        
        
            
            
if __name__ == "__main__":
    app = QApplication(sys.argv)
    daily = Daily_ohlc()
    daily.ohlc('A005930')
    # myWindow = MyWindow()
    # myWindow.show()
    
    app.exec_()