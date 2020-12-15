import sys
from PyQt5.QtWidgets import *
import win32com.client
import pandas as pd
import openpyxl
from openpyxl import Workbook
import time
import numpy
import os
import locale
# 요약: MACD 지표 데이터 실시간 구하기
#     : 차트 OBJECT 를 통해 차트 데이터를 받은 후
#     : 지표 실시간 계산 OBJECT 를 통해 지표 데이터를 계산

# 종목코드 리스트 구하기
objCpCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")
StockItemCodeList1 = objCpCodeMgr.GetStockListByMarket(1) #거래소
StockItemCodeList2 = objCpCodeMgr.GetStockListByMarket(2) #코스닥
StockItemCodeList = objCpCodeMgr.GetStockListByMarket(1) + objCpCodeMgr.GetStockListByMarket(2)

Tuple_nameL1 = [] 
Tuple_nameL2 = []

result_dict = dict()
result_dict_fin = dict()

print("거래소 종목코드", len(StockItemCodeList))
for i, StockItemCode in enumerate(StockItemCodeList):
    secondCodeL1 = objCpCodeMgr.GetStockSectionKind(StockItemCode)
    nameL1 = objCpCodeMgr.CodeToName(StockItemCode)
    stdPriceL1 = objCpCodeMgr.GetStockStdPrice(StockItemCode)
    Tuple_nameL1.append(nameL1)
    #print(i, StockItemCode, secondCode, stdPrice, name)
 
print("코스닥 종목코드", len(StockItemCodeList2))
for i, StockItemCode in enumerate(StockItemCodeList2):
    secondCodeL2 = objCpCodeMgr.GetStockSectionKind(StockItemCode)
    nameL2 = objCpCodeMgr.CodeToName(StockItemCode)
    stdPriceL2 = objCpCodeMgr.GetStockStdPrice(StockItemCode)
    Tuple_nameL2.append(nameL2)
   #print(i, StockItemCode, secondCode, stdPrice, name)
 
print("거래소 + 코스닥 종목코드 ",len(StockItemCodeList) + len(StockItemCodeList2))

class CpEvent:
    def set_params(self, client, objCaller):
        self.client = client
        self.caller = objCaller
 
    def OnReceived(self):
        searchcode = self.client.GetHeaderValue(0)  # 초
        name = self.client.GetHeaderValue(1)  # 초
        timess = self.client.GetHeaderValue(18)  # 초
        exFlag = self.client.GetHeaderValue(19)  # 예상체결 플래그
        cprice = self.client.GetHeaderValue(13)  # 현재가
        diff = self.client.GetHeaderValue(2)  # 대비
        cVol = self.client.GetHeaderValue(17)  # 순간체결수량
        vol = self.client.GetHeaderValue(9)  # 거래량
        open = self.client.GetHeaderValue(4)  # 고가
        high = self.client.GetHeaderValue(5)  # 고가
        low = self.client.GetHeaderValue(6)  # 저가
 
        if (exFlag == ord('1')):  # 동시호가 시간 (예상체결)
            print("실시간(예상체결)", name, timess, "*", cprice, "대비", diff, "체결량", cVol, "거래량", vol)
            return  # 차트는 예상 체결 시간 update 없음.
        elif (exFlag == ord('2')):  # 장중(체결)
            print("실시간(장중 체결)", name, timess, cprice, "대비", diff, "체결량", cVol, "거래량", vol)
 
        # MACD 지표 update 함수 호출
        self.caller.updateMACD(cprice, open, high, low, vol)
 
class CpStockChart: #@@@@@@Cybos  연결
    def __init__(self):
        self.objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
        self.objStockChart = win32com.client.Dispatch("CpSysDib.StockChart")
 
    def Request(self, searchcode, objCaller):
        # 연결 여부 체크
        bConnect = self.objCpCybos.IsConnect
        if (bConnect == 0):
            print("PLUS가 정상적으로 연결되지 않음. ")
            return False
 
        print("프로그램을 시작합니다... by yoonseok")

        

        # 현재가 객체 구하기
        self.objStockChart.SetInputValue(0, searchcode)  # 종목 코드 - 삼성전자
        self.objStockChart.SetInputValue(1, ord('2'))  # 개수로 조회
        self.objStockChart.SetInputValue(4, Day)  # 최근 500일치
        self.objStockChart.SetInputValue(5, [0, 2, 3, 4, 5, 8])  # 날짜,시가,고가,저가,종가,거래량
        self.objStockChart.SetInputValue(6, ord('D'))  # '차트 주기 - 일간 차트 요청
        self.objStockChart.SetInputValue(9, ord('1'))  # 수정주가 사용
        self.objStockChart.BlockRequest()
        
        rqStatus = self.objStockChart.GetDibStatus()
        rqRet = self.objStockChart.GetDibMsg1()
        print("통신상태", rqStatus, rqRet)
        if rqStatus != 0:
            print("통신상태가 올바르지 않습니다.")
 
        # MACD 지표 계산 함수 호출
        objCaller.makeChartSeries(self.objStockChart)
 




class CpStockCur:
    def Subscribe(self, searchcode, objIndex):
        self.objStockCur = win32com.client.Dispatch("DsCbo1.StockCur")
        handler = win32com.client.WithEvents(self.objStockCur, CpEvent)
        self.objStockCur.SetInputValue(0, searchcode)
        handler.set_params(self.objStockCur, objIndex)
        self.objStockCur.Subscribe()
 
    def Unsubscribe(self):
        self.objStockCur.Unsubscribe()
 
 
class MyWindow(QMainWindow):
    CLICK = 0
    ROW = 2
    Auto = 0
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PLUS API TEST")
        self.setGeometry(300, 300, 300, 150)
        self.isSB = False
        self.objCur = []

        

        btnStart = QPushButton("전체 검색", self)
        btnStart.move(20, 20)
        btnStart.clicked.connect(self.btnStart_clicked)
        btnSelect = QPushButton("코드 검색", self)
        btnSelect.move(20, 70)
        btnSelect.clicked.connect(self.btnSelect_clicked)
        btnSave = QPushButton("데이터 저장", self)
        btnSave.move(150,120)
        btnSave.clicked.connect(self.btnSave_clicked)
        btnExit = QPushButton("종료", self)
        btnExit.move(20, 120)
        btnExit.clicked.connect(self.btnExit_clicked)
        ######################################################################################################################
        self.lineEdit = QLineEdit(self)
        self.lineEdit.move (150,70)
        self.lineEdit.returnPressed.connect(self.lineEdit_enter)

        self.lineEditDay = QLineEdit(self)
        self.lineEditDay.move (150,20)
        self.lineEditDay.returnPressed.connect(self.lineEditDay_enter)

        # obj 미리 선언??
    def btnSave_clicked(self) :
        self.StopSubscribe();

        df = pd.DataFrame([result_dict_fin], index=[1])
        df = df.transpose()
        df.to_excel("KOSPILIST.xlsx")

    def lineEdit_enter(self, codetext) :
        self.lineEdit.setText(self.lineEdit.text())
        self.lineEdit.adjustSize()
    def lineEditDay_enter(self, codetext) :
        self.lineEditDay.setText(self.lineEdit.text())
        self.lineEditDay.adjustSize()

    def StopSubscribe(self):
        if self.isSB:
            cnt = len(self.objCur)
            for i in range(cnt):
                self.objCur[i].Unsubscribe()
            print(cnt, "종목 실시간 해지되었음")
        self.isSB = False
 
        self.objCur = []
    
        
    def btnStart_clicked(self):
        self.StopSubscribe();
        
        # 요청 종목
        global Day 
        Day = self.lineEditDay.text()
        searchcode = StockItemCodeList[self.CLICK]
        print(Tuple_nameL1[self.CLICK])
        self.CLICK+=1
        # 지표 계산을 위한 시리즈 선언 - 차트 데이터 수신 받아 데이터를 넣어야 함.
        self.objSeries = win32com.client.Dispatch("CpIndexes.CpSeries")
 
        # 1. 차트 데이터 통신 요청
        self.objChart = CpStockChart()
        if self.objChart.Request(searchcode,  self) == False:
            print("exit")
 
        # 2. macd 지표 만들기
        self.makeMACD()
 
        
       # 3. 현재가 실시간 요청하기
        self.objCur.append(CpStockCur())
        self.objCur[0].Subscribe(searchcode,self)
        
        print("============================")
        print("종목 실시간 현재가 요청 시작")
        self.isSB = True
 
    def btnSelect_clicked(self):
        self.StopSubscribe()
        global Day 
        Day = self.lineEditDay.text()
        searchcode = self.lineEdit.text()
    # 지표 계산을 위한 시리즈 선언 - 차트 데이터 수신 받아 데이터를 넣어야 함.
        self.objSeries = win32com.client.Dispatch("CpIndexes.CpSeries")
 
        # 1. 차트 데이터 통신 요청
        self.objChart = CpStockChart()
        if self.objChart.Request(searchcode,  self) == False:
            print("exit")
 
        # 2. macd 지표 만들기
        self.makeMACD()
 
        
       # 3. 현재가 실시간 요청하기
        self.objCur.append(CpStockCur())
        self.objCur[0].Subscribe(searchcode,self)
        
        print("============================")
        print("종목 실시간 현재가 요청 시작")
        self.isSB = True
 
    def btnExit_clicked(self):
        self.StopSubscribe()
        exit()
 
    # 차트 수신 데이터 --> 시리즈 생성
    # 차트 수신 데이터의 경우 최근 데이터가 맨 앞에 있으나
    # 시리즈는 반대로 넣어야 함.
    # 차트 데이터를 가져와 역순으로 시리즈에 넣는 작업 필요
    def makeChartSeries(self, objStockChart):
        len = objStockChart.GetHeaderValue(3)
 
        print("날짜", "시가", "고가", "저가", "종가", "거래량")
        print("=================================================1")
 
        for i in range(len):
            day = objStockChart.GetDataValue(0, len - i - 1)
            open = objStockChart.GetDataValue(1, len - i - 1)
            high = objStockChart.GetDataValue(2, len - i - 1)
            low = objStockChart.GetDataValue(3, len - i - 1)
            close = objStockChart.GetDataValue(4, len - i - 1)
            vol = objStockChart.GetDataValue(5, len - i - 1)
            print(day, open, high, low, close, vol)
            # objSeries.Add 종가, 시가, 고가, 저가, 거래량, 코멘트
            self.objSeries.Add(close, open, high, low, vol)
        return
 
    # CpIndex 를 이용하여 MACD 지표 계산
    # MACD 는 총 3가지 지표가 들어 있음(MACD, SIGNAL, OSCILLATOR)
    # 최근 데이터는 지표의 맨 마지막 데이터에 들어 있음.
    def makeMACD(self):
        ROW = 2
        write_wb = Workbook()
        # 지표 계산 object
        self.objIndex = win32com.client.Dispatch("CpIndexes.CpIndex")
        self.objIndex.series = self.objSeries
        self.objIndex.put_IndexKind("MACD")     # 계산할 지표: MACD
        self.objIndex.put_IndexDefault("MACD")  # MACD 지표 기본 변수 자동 세팅
 
        print("MACD 변수", self.objIndex.get_Term1(), self.objIndex.get_Term2(), self.objIndex.get_Signal())
        # 지표 데이터 계산 하기
        self.objIndex.Calculate()
        staticsCalculation = []
        cntofIndex = self.objIndex.ItemCount
        print("지표 개수:  ", cntofIndex )
        indexName = ["MACD", "SIGNAL", "OSCILLATOR"]
        for index in range(cntofIndex):
            print(len(indexName[index]))
            cnt = self.objIndex.GetCount(index)
            #for j in range(cnt) :
            #    value = self.objIndex.GetResult(index,j)
            value = self.objIndex.GetResult(index, cnt - 1) # 지표의 가장 최근 값 - 맨 뒤 데이터
            print(indexName[index], value)  # 지표의 최근 값 표시
            staticsCalculation.append(value)
        
            
        
        try :
            if -4 < staticsCalculation[0] < 5 and staticsCalculation[0] > staticsCalculation[1] and abs(staticsCalculation[0]) < abs(staticsCalculation[1]) :
                print("가즈아아아아아아아")
                result_dict[Tuple_nameL1[self.CLICK-1]] = StockItemCodeList[self.CLICK-1]
                result_dict_fin.update(result_dict)
                print(result_dict_fin)
                df = pd.DataFrame([result_dict_fin], index=[1])
                df = df.transpose()
                df.to_excel("MACDgotoZero.xlsx")
            
            else :
                
                print("다른 종목을 검색중입니다....")
                print("현재 검색한 종목 수 : ",(self.CLICK))
                print('현재 기록한 종목 수 : ',len(result_dict_fin))
                
        except :
            print("passing...") 
            pass    




    # 실시간 시세 수신 받아 MACD 계산
    def updateMACD(self, cprice, open, high, low, vol):
        # 지표 데이터 update
        self.objSeries.update(cprice, open, high, low, vol)
        self.objIndex.update()
        cntofIndex = self.objIndex.ItemCount
        print("지표 개수:  ", cntofIndex )
 
        indexName = ["MACD", "SIGNAL", "OSCILLATOR"]
 
        for index in range(cntofIndex):
            cnt = self.objIndex.GetCount(index)
            # print(index , "번째 지표의 데이터 개수", cnt)
            value = self.objIndex.GetResult(index, cnt - 1) # 지표의 가장 최근 값 - 맨 뒤 데이터
            print(indexName[index], value)  # 지표의 최근 값 표시
            print("라인 219")
        return
 
 
if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()
