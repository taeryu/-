import win32com.client
from SAPvariables import *

#전원경영 테스트완료
def SAP_everybody():  

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
    session.findById("wnd[0]/tbar[0]/okcd").text = "ZCOQR427"
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]/usr/ctxtP_BUKRS").text = companyCode
    session.findById("wnd[0]/usr/txtP_GJAHR").text = thisYear
    session.findById("wnd[0]/usr/ctxtP_MONTH").text = thisMonth
    session.findById("wnd[0]/usr/ctxtS_VKBUR-LOW").setFocus()
    session.findById("wnd[0]/usr/ctxtS_VKBUR-LOW").caretPosition = (0)
    session.findById("wnd[0]/usr/btn%_S_GSBER_%_APP_%-VALU_PUSH").press()
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "3040"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "3060"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").setFocus
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").caretPosition = 4
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").pressToolbarButton ("AGGR")
    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
    session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").pressToolbarButton ("POST")
    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()

SAP_everybody()

