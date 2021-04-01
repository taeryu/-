import win32com.client
from SAPvariables import *

def SAP_innerSales():

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
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nzcoqr410"   #섬산화사내거래조정
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/ctxtP_BUKRS").text = companyCode
    session.findById("wnd[0]/usr/txtP_GJAHR").text = thisYear
    session.findById("wnd[0]/usr/ctxtP_MONTH").text = thisMonth
    session.findById("wnd[0]/usr/ctxtP_PRGRU").text = companyCode
    session.findById("wnd[0]/usr/ctxtP_PRGRU").setFocus()
    session.findById("wnd[0]/usr/ctxtP_PRGRU").caretPosition = "4"
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
    session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectedRows = "0-3"
    session.findById("wnd[0]/tbar[1]/btn[19]").press()
    session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
    session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectedRows = "0-3"
    session.findById("wnd[0]/tbar[1]/btn[13]").press()
    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    session.findById("wnd[0]/usr/radP_R02").setFocus()
    session.findById("wnd[0]/usr/radP_R02").select()
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
    session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectedRows = "0-3"
    session.findById("wnd[0]/tbar[1]/btn[13]").press()
    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()

SAP_innerSales()