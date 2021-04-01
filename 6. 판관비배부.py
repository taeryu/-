import win32com.client
from SAPvariables import *

#판관비 테스트완료
def SAP_OP2():  

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

    session.findById("wnd[0]/tbar[0]/okcd").text = "/nKSU5"					           
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]/usr/chkRKGA2U-TEST").selected = False				             #테스트실행 해제
    session.findById("wnd[0]/usr/sub:SAPMKGA2:0101/ctxtRKGA2-KSCYC[0,0]").text = "903AA1"
    session.findById("wnd[0]/usr/sub:SAPMKGA2:0101/ctxtRKGA2-KSCYC[0,0]").setFocus
    session.findById("wnd[0]/usr/sub:SAPMKGA2:0101/ctxtRKGA2-KSCYC[0,0]").caretPosition = "6"
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
    session.findById("wnd[0]/usr/sub:SAPMKGA2:0101/ctxtRKGA2-KSCYC[0,0]").text = "903AA2"
    session.findById("wnd[0]/usr/sub:SAPMKGA2:0101/ctxtRKGA2-KSCYC[1,0]").text = "903AA3"
    session.findById("wnd[0]/usr/sub:SAPMKGA2:0101/ctxtRKGA2-KSCYC[2,0]").text = "903AA4"
    session.findById("wnd[0]/usr/sub:SAPMKGA2:0101/ctxtRKGA2-KSCYC[5,0]").text = "903AA7"
    session.findById("wnd[0]/usr/sub:SAPMKGA2:0101/ctxtRKGA2-KSCYC[6,0]").text = "903AA8"
    session.findById("wnd[0]/usr/sub:SAPMKGA2:0101/ctxtRKGA2-KSCYC[7,0]").text = "903APA"
    session.findById("wnd[0]/usr/sub:SAPMKGA2:0101/ctxtRKGA2-KSCYC[8,0]").text = "903APB"
    session.findById("wnd[0]/usr/sub:SAPMKGA2:0101/ctxtRKGA2-KSCYC[9,0]").text = "903APD"
    session.findById("wnd[0]/usr/sub:SAPMKGA2:0101/ctxtRKGA2-KSCYC[2,0]").setFocus()
    session.findById("wnd[0]/usr/sub:SAPMKGA2:0101/ctxtRKGA2-KSCYC[2,0]").caretPosition = "5"
    session.findById("wnd[0]").sendVKey (0)				
    session.findById("wnd[0]/tbar[1]/btn[8]").press()	

SAP_OP2()	