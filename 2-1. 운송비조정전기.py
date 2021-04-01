import win32com.client
import time
from SAPvariables import *

def SAP_OP():
   
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    if not type(SapGuiAuto) == win32com.client.CDispatch:
        return

    application = SapGuiAuto.GetScriptingEngine
    if not type(application) == win32com.client.CDispatch:
        SapGuiAuto = None
        return

    connection = application.Children(0)
    if not type(connection) == win32com.client.CDispatch:
        application = None
        SapGuiAuto = None
        return

    session = connection.Children(0)
    if not type(session) == win32com.client.CDispatch:
        connection = None
        application = None
        SapGuiAuto = None
        return

    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "/Nzsdfr844"	
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/ctxtPA_BUKRS").text = companyCode	    #회사코드
    session.findById("wnd[0]/usr/ctxtPA_MONTH").text = yearANDmonth		#연도 + 월
    session.findById("wnd[0]/usr/ctxtSO_BUDAT-LOW").text = sttDate  	#시작일
    session.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").text = endDate 	#종료일
    time.sleep(5)
    session.findById("wnd[0]/tbar[1]/btn[8]").press()	                 #실행
    session.findById("wnd[0]/shellcont/shell").selectedRows = "0-50"	 #전체행선택시 행수
    time.sleep(5)
    session.findById("wnd[0]/tbar[1]/btn[18]").press()			      	#전표생성 버튼
    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()		    	#팝업창 "예"
   
SAP_OP()

