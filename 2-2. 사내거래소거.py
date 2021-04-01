import win32com.client
from SAPvariables import *

 #테스트 미완료 aggregation까지만
def SAP_innerSalesDeduct():

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
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nZCOQR402"
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]/usr/ctxtP_BUKRS").text = companyCode
    session.findById("wnd[0]/usr/ctxtP_MONTH").text = thisMonth
    session.findById("wnd[0]/usr/cmbP_PG").setFocus()
    session.findById("wnd[0]/usr/cmbP_PG").key = "03"
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/tbar[1]/btn[19]").press()
 #aggregation까지

SAP_innerSalesDeduct()