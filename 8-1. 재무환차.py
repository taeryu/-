import win32com.client
from SAPvariables import *

#영업외배부 전 ZCOQR443	재무환차손익 결산조정 실행	

def FS_Exchange():

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
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nZCOQR443"
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]/usr/ctxtP_BUKRS").text = companyCode
    session.findById("wnd[0]/usr/txtP_GJAHR").text = thisYear
    session.findById("wnd[0]/usr/txtP_MONAT").text = thisMonth
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/usr/cntlCUSTOM_0100/shellcont/shell").currentCellColumn = ""
    session.findById("wnd[0]/usr/cntlCUSTOM_0100/shellcont/shell").selectedRows = "0-1"  #전체행선택하기 어떻게 함??

FS_Exchange()