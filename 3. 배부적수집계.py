import win32com.client
from SAPvariables import *

#배부적수집계 테스트 미완료 (해외법인꺼에서 오류)
def SAP_SKF():  

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
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nzcoqr404"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/ctxtP_BUKRS").text = companyCode
    session.findById("wnd[0]/usr/txtP_GJAHR").text = thisYear
    session.findById("wnd[0]/usr/ctxtP_MONTH").text = thisMonth
    
    # #전사배부_사외매출
    session.findById("wnd[0]/usr/ctxtP_STAGR").text = "p00001"
    session.findById("wnd[0]/usr/ctxtP_STAGR").setFocus()
    session.findById("wnd[0]/usr/ctxtP_STAGR").caretPosition = "6"
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/tbar[1]/btn[19]").press()
    session.findById("wnd[0]/tbar[1]/btn[13]").press()
    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    session.findById("wnd[0]/usr/radP_R04").setFocus()
    session.findById("wnd[0]/usr/radP_R04").select()
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/tbar[1]/btn[19]").press()
    session.findById("wnd[0]/tbar[1]/btn[13]").press()
    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    
    #전사외배부_총매출
    session.findById("wnd[0]/usr/ctxtP_STAGR").text = "P00011"
    session.findById("wnd[0]/usr/radP_R01").setFocus()
    session.findById("wnd[0]/usr/radP_R01").select()
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/tbar[1]/btn[19]").press()
    session.findById("wnd[0]/tbar[1]/btn[13]").press()
    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    session.findById("wnd[0]/usr/radP_R04").setFocus()
    session.findById("wnd[0]/usr/radP_R04").select()
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/tbar[1]/btn[19]").press()
    session.findById("wnd[0]/tbar[1]/btn[13]").press()
    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    
    #PU 공통_사외 수출매출
    session.findById("wnd[0]/usr/ctxtP_STAGR").text = "P00012"
    session.findById("wnd[0]/usr/radP_R01").setFocus()
    session.findById("wnd[0]/usr/radP_R01").select()
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/tbar[1]/btn[19]").press()
    session.findById("wnd[0]/tbar[1]/btn[13]").press()
    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    session.findById("wnd[0]/usr/radP_R04").setFocus()
    session.findById("wnd[0]/usr/radP_R04").select()
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/tbar[1]/btn[19]").press()
    session.findById("wnd[0]/tbar[1]/btn[13]").press()
    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()

    #전사배부_사외매출
    session.findById("wnd[0]/usr/ctxtP_STAGR").text = "zco001"
    session.findById("wnd[0]/usr/radP_R01").setFocus()
    session.findById("wnd[0]/usr/radP_R01").select()
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/tbar[1]/btn[19]").press()
    session.findById("wnd[0]/tbar[1]/btn[13]").press()
    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()

    #전사외배부_총매출
    session.findById("wnd[0]/usr/ctxtP_STAGR").text = "ZCO011"
    session.findById("wnd[0]/usr/ctxtP_STAGR").setFocus()
    session.findById("wnd[0]/usr/ctxtP_STAGR").caretPosition = "6"
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/tbar[1]/btn[19]").press()
    session.findById("wnd[0]/tbar[1]/btn[13]").press()
    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()

  
    #해외법인 SKF집계  (P00003 티엔씨 P00004첨단소재 P00005효성화학) ( 연구도 돌려야됨)
    session.findById("wnd[0]/tbar[1]/btn[28]").press()              #기타버튼 클릭
    session.findById("wnd[1]/usr/btn%#AUTOTEXT005").press()         #연결결산 SKF집계 버튼 클릭
    session.findById("wnd[0]/usr/ctxtP_BUKRS").text = companyCode        #회사코드 
    session.findById("wnd[0]/usr/ctxtP_PERBL").text = yearANDmonth     #기간
    session.findById("wnd[0]/usr/ctxtP_STAGR").text = "P00005"      #통계주요지표
    session.findById("wnd[0]/usr/ctxtP_STAGR").setFocus()
    session.findById("wnd[0]/usr/ctxtP_STAGR").caretPosition = "6"
    session.findById("wnd[0]/tbar[1]/btn[8]").press()                   #실행
    session.findById("wnd[1]/tbar[0]/btn[0]").press()                   #이미 기존데이터가 있습나다 = OK 
    session.findById("wnd[0]/tbar[1]/btn[21]").press()                  #저장
    session.findById("wnd[1]/usr/btnBUTTON_1").press()                  #저장하시겠습니까?
    session.findById("wnd[0]/tbar[0]/btn[3]").press()                   #뒤로가기
    session.findById("wnd[0]/tbar[0]/btn[3]").press()                   #뒤로가기 
    session.findById("wnd[0]/usr/ctxtP_STAGR").text = "P00005"          #배부적수 집계 및 POSTING화면에서 SKF코드 입력
    session.findById("wnd[0]/usr/ctxtP_STAGR").setFocus ()
    session.findById("wnd[0]/usr/ctxtP_STAGR").caretPosition = 6
    session.findById("wnd[0]/tbar[1]/btn[8]").press()                    #실행
    session.findById("wnd[0]/tbar[1]/btn[19]").press()
    session.findById("wnd[0]/tbar[1]/btn[13]").press()
    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()

    #해외법인 SKF집계 (연구)
    session.findById("wnd[0]/usr/radP_R04").setFocus()
    session.findById("wnd[0]/usr/radP_R04").select()
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/tbar[1]/btn[19]").press()
    session.findById("wnd[0]/tbar[1]/btn[13]").press()
    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()

SAP_SKF()