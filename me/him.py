import sys
import os
import time
import argparse
import subprocess
import abc
# from util.util import util
from PyQt5.QtWidgets import *

import warnings
warnings.simplefilter("ignore", UserWarning)
sys.coinit_flags = 2


import win32com.client
from pywinauto import application







class Creon():
    def __init__(self):
        self.obj_CpUtil_CpCybos = win32com.client.Dispatch('CpUtil.CpCybos') #매수매도 오브젝
        self.obj_CpUtil_CpCodeMgr = win32com.client.Dispatch('CpUtil.CpCodeMgr') #각종코드정보및코드리스트를얻을수있습니다.
        self.obj_CpSysDib_StockChart = win32com.client.Dispatch('CpSysDib.StockChart') # 주식 차트데이터 받기
        self.obj_CpTrade_CpTdUtil = win32com.client.Dispatch('CpTrade.CpTdUtil') # 주문을 위해 초기화 과정 수행
        self.obj_CpSysDib_MarketEye = win32com.client.Dispatch('CpSysDib.MarketEye') # 
        self.obj_CpSysDib_CpSvr7238 = win32com.client.Dispatch('CpSysDib.CpSvr7238')
        self.obj_CpTrade_CpTdNew5331B = win32com.client.Dispatch('CpTrade.CpTdNew5331B') # 계좌벼 매도 주문 가능 금액/수량 데이터 요텅하고 수신
        self.obj_CpTrade_CpTdNew5331A = win32com.client.Dispatch('CpTrade.CpTdNew5331A') # 계좌별 매수 주문 가능 금액/수량 데이터 요청하고 수신
        self.obj_CpSysDib_CpSvr7254 = win32com.client.Dispatch('CpSysDib.CpSvr7254')
        self.obj_CpSysDib_CpSvr8548 = win32com.client.Dispatch('CpSysDib.CpSvr8548')
        self.obj_CpTrade_CpTd0311 = win32com.client.Dispatch('CpTrade.CpTd0311') # 현금 주문 데이터를 요청하고 수신한다. 
        self.obj_CpTrade_CpTd5341 = win32com.client.Dispatch('CpTrade.CpTd5341')
        self.obj_CpTrade_CpTd6033 = win32com.client.Dispatch('CpTrade.CpTd6033') # 계좌별 잔고 및 주문 체결 평가현황데이터를 요청하고 수신한다.
        self.obj_Dscbo1_CpConclusion = win32com.client.Dispatch('Dscbo1.CpConclusion') #미체결 요청하고 수신하는 오브젝
        self.obj_CpTrade_CpTd0322 = win32com.client.Dispatch('CpTrade.CpTd0322')
        self.obj_Dscbo1_StockBid = win32com.client.Dispatch('Dscbo1.StockBid')
        
        # contexts
        self.stockcur_handlers = {}  # 주식/업종/ELW시세 subscribe event handlers
        self.stockbid_handlers = {}  # 주식/ETF/ELW 호가, 호가잔량 subscribe event handlers
        self.orderevent_handler = None

    
    
    
    
    def connect(self, id_, pwd, pwdcert, trycnt=300): 
        
        # os.system('taskkill /IM coStarter* /F /T')
        # os.system('taskkill /IM CpStart* /F /T')
        # os.system('taskkill /IM DibServer* /F /T')
        # os.system('wmic process where "name like \'%coStarter%\'" call terminate')
        # os.system('wmic process where "name like \'%CpStart%\'" call terminate')
        # os.system('wmic process where "name like \'%DibServer%\'" call terminate')
        # time.sleep(5)        

        # app = application.Application()
        # app.start('C:\CREON\STARTER\coStarter.exe /prj:cp /id:{id_} /pwd:{pwd} /pwdcert:{pwdcert} /autostart'.format(
        #             id_=id_, pwd=pwd, pwdcert=pwdcert
        #         ))
        # time.sleep(60)

        if not self.connected():
            app = application.Application()
            
            app.start(
            'C:\\CREON\\STARTER\\coStarter.exe /prj:cp /id:{id_} /pwd:{pwd} /pwdcert:{pwdcert} /autostart'.format(
                    id_=id_, pwd=pwd, pwdcert=pwdcert
                )
            )

        cnt = 0
        while not self.connected():
            if cnt > trycnt:
                return False
            time.sleep(1)
            cnt += 1
        return True





    def connected(self):
        tasklist = subprocess.check_output('TASKLIST')
        if b"DibServer.exe" in tasklist and b"CpStart.exe" in tasklist:
            return self.obj_CpUtil_CpCybos.IsConnect != 0
        return False






    def disconnect(self):
        plist = [
            'coStarter',
            'CpStart',
            'DibServer',
        ]
        for p in plist:
            os.system('wmic process where "name like \'%{}%\'" call terminate'.format(p))
        return True




    
    
    # CpCybos의 LimitRequestRemainTime과 GetLimitRemainCount로타임아웃까지남은시간과남아있는요청개수를얻을수있습니다.
    
    def wait(self):
        remain_time = self.obj_CpUtil_CpCybos.LimitRequestRemainTime
        remain_count = self.obj_CpUtil_CpCybos.GetLimitRemainCount(1)
        if remain_count <= 3:
            time.sleep(remain_time / 1000)






    def request(self, obj, data_fields, header_fields=None, cntidx=0, n=None):
        def process():
            obj.BlockRequest()

            status = obj.GetDibStatus()
            msg = obj.GetDibMsg1()
            if status != 0:
                return None

            cnt = obj.GetHeaderValue(cntidx)
            data = []
            for i in range(cnt):
                dict_item = {k: obj.GetDataValue(j, cnt-1-i) for j, k in data_fields.items()}
                data.append(dict_item)
            return data

        # 연속조회 처리
        data = process()
        while obj.Continue:
            self.wait()
            _data = process()
            if len(_data) > 0:
                data = _data + data
                if n is not None and n <= len(data):
                    break
            else:
                break

        result = {'data': data}
        if header_fields is not None:
            result['header'] = {k: obj.GetHeaderValue(i) for i, k in header_fields.items()}

        return result
    
    
    
    
    
    def get_balance(self):
        """
        매수가능금액
        """
        account_no, account_gflags = self.init_trade()
        self.obj_CpTrade_CpTdNew5331A.SetInputValue(0, account_no)
        self.obj_CpTrade_CpTdNew5331A.BlockRequest()
        v = self.obj_CpTrade_CpTdNew5331A.GetHeaderValue(10)
        return v

    
    
    
    
    
    def get_holdingstocks(self):
        """
        보유종목
        """
        account_no, account_gflags = self.init_trade()
        self.obj_CpTrade_CpTdNew5331B.SetInputValue(0, account_no)
        self.obj_CpTrade_CpTdNew5331B.SetInputValue(3, ord('1')) # 1: 주식, 2: 채권
        self.obj_CpTrade_CpTdNew5331B.BlockRequest()
        cnt = self.obj_CpTrade_CpTdNew5331B.GetHeaderValue(0)
        res = []
        for i in range(cnt):
            item = {
                'code': self.obj_CpTrade_CpTdNew5331B.GetDataValue(0, i),
                'name': self.obj_CpTrade_CpTdNew5331B.GetDataValue(1, i),
                'holdnum': self.obj_CpTrade_CpTdNew5331B.GetDataValue(6, i),
                'buy_yesterday': self.obj_CpTrade_CpTdNew5331B.GetDataValue(7, i),
                'sell_yesterday': self.obj_CpTrade_CpTdNew5331B.GetDataValue(8, i),
                'buy_today': self.obj_CpTrade_CpTdNew5331B.GetDataValue(10, i),
                'sell_today': self.obj_CpTrade_CpTdNew5331B.GetDataValue(11, i),
            }
            res.append(item)
        return res
    
    
    
    
    
    
    def init_trade(self):
        if self.obj_CpTrade_CpTdUtil.TradeInit(0) != 0:
            print("TradeInit failed.", file=sys.stderr)
            return
        account_no = self.obj_CpTrade_CpTdUtil.AccountNumber[0]  # 계좌번호
        account_gflags = self.obj_CpTrade_CpTdUtil.GoodsList(account_no, 1)  # 주식상품 구분
        return account_no, account_gflags






    def order(self, action, code, amount):
        if not code.startswith('A'):
            code = 'A' + code
        account_no, account_gflags = self.init_trade()
        self.obj_CpTrade_CpTd0311.SetInputValue(0, action)  # 1: 매도, 2: 매수
        self.obj_CpTrade_CpTd0311.SetInputValue(1, account_no)  # 계좌번호
        self.obj_CpTrade_CpTd0311.SetInputValue(2, account_gflags[0])  # 상품구분
        self.obj_CpTrade_CpTd0311.SetInputValue(3, code)  # 종목코드
        self.obj_CpTrade_CpTd0311.SetInputValue(4, amount)  # 매수수량
        self.obj_CpTrade_CpTd0311.SetInputValue(8, '03')  # 시장가
        result = self.obj_CpTrade_CpTd0311.BlockRequest()
        if result != 0:
            print('order request failed.', file=sys.stderr)
        status = self.obj_CpTrade_CpTd0311.GetDibStatus()
        msg = self.obj_CpTrade_CpTd0311.GetDibMsg1()
        if status != 0:
            print('order failed. {}'.format(msg), file=sys.stderr)






    def buy(self, code, amount):
        self.order('2', code, amount)
    
    
    

    
        





    def sell(self, code, amount):
        print(f'구매했습니다. {code}, {amount}')
        return self.order('1', code, amount)
    
    
    
    
    
    def get_holdings(self):
        """
        0 - (string) 계좌번호
        1 - (string) 상품관리구분코드
        2 - (long) 요청건수[default:14] - 최대 50개
        3 - (string) 수익률구분코드 - ( "1" : 100% 기준, "2": 0% 기준)
        """
        account_no, account_gflags = self.init_trade()
        self.obj_CpTrade_CpTd6033.SetInputValue(0, account_no)
        self.obj_CpTrade_CpTd6033.SetInputValue(1, account_gflags[0])
        self.obj_CpTrade_CpTd6033.SetInputValue(3, '2')

        header_fields = {
            0: '계좌명',
            1: '결제잔고수량',
            2: '체결잔고수량',
            3: '총평가금액',
            4: '평가손익',
            6: '대출금액',
            7: '수신개수',
            8: '수익율',
        }

        data_fields = {
            0: '종목명',
            1: '신용구분',
            2: '대출일',
            3: '결제잔고수량',
            4: '결제장부단가',
            5: '전일체결수량',
            6: '금일체결수량',
            7: '체결잔고수량',
            9: '평가금액',
            10: '평가손익',
            11: '수익률',
            12: '종목코드',
            13: '주문구분',
            15: '매도가능수량',
            16: '만기일',
            17: '체결장부단가',
            18: '손익단가',
        }

        result = self.request(self.obj_CpTrade_CpTd6033, data_fields, header_fields=header_fields, cntidx=7)
        return result
    
    
    
    
    
    
class EventHandler:
    # 실시간 조회(subscribe)는 최대 400건

    def set_attrs(self, obj, cb):
        self.obj = obj
        self.cb = cb

    @abc.abstractmethod
    def OnReceived(self):
        pass







class StockCurEventHandler(EventHandler):
    def OnReceived(self):
        item = {
            'code': self.obj.GetHeaderValue(0),
            'name': self.obj.GetHeaderValue(1),
            'diffratio': self.obj.GetHeaderValue(2),
            'timestamp': self.obj.GetHeaderValue(3),  # 시간 형태 확인 필요
            'price_open': self.obj.GetHeaderValue(4),
            'price_high': self.obj.GetHeaderValue(5),
            'price_low': self.obj.GetHeaderValue(6),
            'bid_sell': self.obj.GetHeaderValue(7),
            'bid_buy': self.obj.GetHeaderValue(8),
            'cum_volume': self.obj.GetHeaderValue(9),  # 주, 거래소지수: 천주
            'cum_trans': self.obj.GetHeaderValue(10),
            'price': self.obj.GetHeaderValue(13),
            'contract_type': self.obj.GetHeaderValue(14),
            'cum_sell_volume': self.obj.GetHeaderValue(15),
            'cum_buy_volume': self.obj.GetHeaderValue(16),
            'contract_volume': self.obj.GetHeaderValue(17),
            'second': self.obj.GetHeaderValue(18),
            'price_type': chr(self.obj.GetHeaderValue(19)),  # 1: 동시호가시간 예상체결가, 2: 장중 체결가
            'market_flag': chr(self.obj.GetHeaderValue(20)),  # '1': 장전예상체결, '2': 장중, '4': 장후시간외, '5': 장후예상체결
            'premarket_volume': self.obj.GetHeaderValue(21),
            'diffsign': chr(self.obj.GetHeaderValue(22)),
            'LP보유수량':self.obj.GetHeaderValue(23),
            'LP보유수량대비':self.obj.GetHeaderValue(24),
            'LP보유율':self.obj.GetHeaderValue(25),
            '체결상태(호가방식)':self.obj.GetHeaderValue(26),
            '누적매도체결수량(호가방식)':self.obj.GetHeaderValue(27),
            '누적매수체결수량(호가방식)':self.obj.GetHeaderValue(28),
        }
        self.cb(item)







class StockBidEventHandler(EventHandler):
    def OnReceived(self):
        item = {
            'code': self.obj.GetHeaderValue(0),
            'total_offer': self.obj.GetHeaderValue(23),
            'total_bid': self.obj.GetHeaderValue(24),
        }
        for i in range(10):
            item[f'offer_{i+1}'] = self.obj.GetHeaderValue(3 + i)
            item[f'bid_{i+1}'] = self.obj.GetHeaderValue(3 + i + 1)
            item[f'offer_volume_{i+1}'] = self.obj.GetHeaderValue(3 + i + 2)
            item[f'bid_volume_{i+1}'] = self.obj.GetHeaderValue(3 + i + 3)
        self.cb(item)







class OrderEventHandler(EventHandler):
    def OnReceived(self):
        item = {
            '계좌명': self.obj.GetHeaderValue(1),
            'name': self.obj.GetHeaderValue(2),
            '체결수량': self.obj.GetHeaderValue(3),
            '체결가격': self.obj.GetHeaderValue(4),
            '주문번호': self.obj.GetHeaderValue(5),
            '원주문번호': self.obj.GetHeaderValue(6),
            '계좌번호': self.obj.GetHeaderValue(7),
            '상품관리구분코드': self.obj.GetHeaderValue(8),
            '종목코드': self.obj.GetHeaderValue(9),
            '매매구분코드': self.obj.GetHeaderValue(12),
            '체결구분코드': self.obj.GetHeaderValue(14),
            '체결구분코드': self.obj.GetHeaderValue(14),
            '체결구분코드': self.obj.GetHeaderValue(14),
            '현금신용대용구분코드': self.obj.GetHeaderValue(17),
        }
        self.cb(item)

def sql_buy_list():        
    import pymysql
    import pandas as pd        
            
    table_name = 'issues_name_code'

    _db = pymysql.connect(
        host="localhost",
        db='k_score',
        user = "root",
        password= "dkrkwk18!")


    search_sql = f"SELECT * From `{table_name}`"

    cursor = _db.cursor(pymysql.cursors.DictCursor)

    cursor.execute(search_sql)            
    result = cursor.fetchall() 

    data_table = pd.DataFrame(result)
    def final_list_():
        
        #'=30'에서 notnull인 df만들기
        asd = data_table['code']
        final_list = []
        for i in asd :
            final_list.append(i)
        return final_list

    final_list = final_list_()
    return final_list







if __name__ == '__main__':
    app = QApplication(sys.argv)
    
    # parser = argparse.ArgumentParser()
    # parser.add_argument('action', choices=['connect', 'disconnect'])
    # parser.add_argument('gyjh486')
    # parser.add_argument('dkrkw18!')
    # parser.add_argument('dkrkwk18!!')
    # args = parser.parse_args()

    c = Creon()
    # c.connect('gyjh486', 'dkrkw18!', 'dkrkwk18!!')
    
    for i in sql_buy_list():
        from datetime import datetime
        now_ = datetime.now().replace(microsecond=0)
        c.buy(f'{i}', 10)
        print(f'({now_})')        
        

    

    # if args.action == 'connect':
    #     c.connect(args.id, args.pwd, args.pwdcert)
    # elif args.action == 'disconnect':
    #     c.disconnect()


    
    app.exec_()
    