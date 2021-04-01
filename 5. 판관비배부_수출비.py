import win32com.client
from SAPvariables import *

#수출비 배부 테스트완료
def SAP_export():  

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
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nKSV5"                           #수출비 티코드
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]/usr/chkRKGA2U-TEST").selected = False                    #테스트실행 해제
    session.findById("wnd[0]/usr/txtRKGA2U-FROM").text = thisMonth
    session.findById("wnd[0]/usr/txtRKGA2U-TO").text = thisMonth
    session.findById("wnd[0]/usr/sub:SAPMKGA2:0101/ctxtRKGA2-KSCYC[0,0]").text = "903AD1"   	 #사이클명
    session.findById("wnd[0]/usr/sub:SAPMKGA2:0101/ctxtRKGA2-KSCYC[0,0]").setFocus()
    session.findById("wnd[0]/usr/sub:SAPMKGA2:0101/ctxtRKGA2-KSCYC[0,0]").caretPosition = "6"
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()

SAP_export()					