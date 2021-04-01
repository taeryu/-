import win32com.client
from SAPvariables import *

def SAP_CI():

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
    session.findById("wnd[0]/tbar[0]/okcd").text = "/Nfagll03"	
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/ctxtSO_BUDAT-LOW").text = sttDate
    session.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").text = endDate
    session.findById("wnd[0]/usr/ctxtPA_VARI").text = "/BS3"
    session.findById("wnd[0]/usr/txtPA_NMAX").text = "1000000"
    session.findById("wnd[0]/usr/txtPA_NMAX").setFocus()
    session.findById("wnd[0]/usr/txtPA_NMAX").caretPosition = "10"
    session.findById("wnd[0]/tbar[1]/btn[8]").press()

SAP_CI()