import win32com.client
from SAPvariables import *

#영업외배부

def SAP_OtherIncome():

    excelPath = r'#'
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
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nKSU5"
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]/usr/chkRKGA2U-TEST").selected = False
    session.findById("wnd[0]/usr/txtRKGA2U-FROM").text = thisMonth
    session.findById("wnd[0]/usr/txtRKGA2U-TO").text = thisMonth
    session.findById("wnd[0]/usr/txtRKGA2U-GJAHR").text = thisYear
    session.findById("wnd[0]/usr/sub:SAPMKGA2:0101/ctxtRKGA2-KSCYC[0,0]").text = "903AA6"
    session.findById("wnd[0]/usr/sub:SAPMKGA2:0101/ctxtRKGA2-KSCYC[1,0]").text = "903APC"
    session.findById("wnd[0]/usr/sub:SAPMKGA2:0101/ctxtRKGA2-KSCYC[0,0]").setFocus()
    session.findById("wnd[0]/usr/sub:SAPMKGA2:0101/ctxtRKGA2-KSCYC[0,0]").caretPosition = "6"
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()

SAP_OtherIncome()