import win32com.client
from SAPvariables import *

def SAP_PCA_PL():  

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
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nzcoqr451"
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]/usr/txt$0R-YEAR").text = thisYear
    session.findById("wnd[0]/usr/ctxt$ZBBUKRS").text = companyCode
    session.findById("wnd[0]/usr/txt$0F-RP00").text = thisMonth
    session.findById("wnd[0]/usr/ctxt$PCTR-GR").text = companyCode
    session.findById("wnd[0]/usr/ctxt_PCTR-GR-LOW").setFocus()
    session.findById("wnd[0]/usr/ctxt_PCTR-GR-LOW").caretPosition = (0)
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/shellcont/shell/shellcont[0]/shell").hierarchyHeaderWidth = "371"
    session.findById("wnd[0]/shellcont/shell/shellcont[2]/shell").selectedNode = "000004"
    session.findById("wnd[0]/usr/lbl[55,23]").setFocus()
    session.findById("wnd[0]/usr/lbl[55,23]").caretPosition = "9"
    session.findById("wnd[0]/usr").verticalScrollbar.position = "27"
    session.findById("wnd[0]/shellcont/shell/shellcont[2]/shell").expandNode "000053"
    session.findById("wnd[0]/shellcont/shell/shellcont[2]/shell").selectedNode = "000055"
    session.findById("wnd[0]/shellcont/shell/shellcont[2]/shell").topNode = "000001"
    session.findById("wnd[0]/shellcont/shell/shellcont[2]/shell").expandNode "000060"
    session.findById("wnd[0]/shellcont/shell/shellcont[2]/shell").selectedNode = "000061"
    session.findById("wnd[0]/shellcont/shell/shellcont[2]/shell").topNode = "000001"
    session.findById("wnd[0]/usr/lbl[55,36]").setFocus()
    session.findById("wnd[0]/usr/lbl[55,36]").caretPosition = "18"

SAP_innerinterst()

'''
손익 센터	매치코드에 대한 손익센터내역
90300099	전사소거
90330090	화학_공통_본사
90330091	화학_공통_용연
90332090	PPDH_PU공통
90332091	PPDH_공통_용연
90332098	PPDH_재무관리 차이조정
90332099	PPDH_PU소거
90332101	DH_1_용연
90332103	DH_2_용연
90332190	DH_본사
90332191	DH_공통_용연
90332198	DH_재무관리 차이조정
90332199	DH_재무관리 차이조정
90332201	PP1_용연
90332202	PP2_용연
90332204	PP3_용연
90332290	PP_본사
90332291	PP_공통_용연
90332298	PP_재무관리 차이조정
90332299	PP_재무관리 차이조정
90332390	CNT_본사
90333099	TPA_PU소거
90333101	TPA_용연
90333190	TPA_본사
90334090	필름_PU공통
90334091	필름_공통_구미
90334099	필름_PU소거
90334101	나이론필름_대전
90334102	나이론필름_구미
90334190	나이론필름_본사
90334201	PET필름_구미
90334202	PET필름_1LINE_용연
90334203	PET필름_2LINE_용연
90334290	PET필름_본사
90334291	PET필름_공통_용연
90335099	NEOCHEM PU소거
90335101	NEOCHEM_DF_용연
90335102	NEOCHEM_ECF-1_용연
90335104	NEOCHEM_ECF-2_용연
90335105	NEOCHEM_ECF-3_용연
90335106	NEOCHEM_F2N2_용연
90335107	NEOCHEM_상품_용연
90335190	NEOCHEM_본사
90335191	NEOCHEM_공통_용연
90336090	OPTICAL FILM_PU공통
90336092	OPTICAL FILM_옥산공장공통
90336099	OPTICAL FILM_PU소거
90336101	TAC_용연
90336103	TAC_옥산
90336190	TAC_본사
90336201	코팅_옥산
90336290	코팅_본사
90338101	POK_용연
90338190	POK_본사
90339101	TANK.T_온산
90339102	TANK.T_용연
90339190	TANK.T_본사
90391001	화학_직할
90391002	화학_경영전략실
90391003	화학_지원실
90391004	화학_재무실
90392101	화학_임대/기타_본사
90392106	화학_임대/기타_구미공장
90392109	임대/기타_용연공장
90392112	임대/기타_온산탱크
90392301	전사기타_본사
90392306	화학_전사기타_구미
90392313	전사기타_용연
90392315	전사기타_대전2
90392326	전사기타_옥산
90392328	전사기타_온산탱크
90392342	화학_전사기타_본사_구매
90392343	화학_전사기타_재무본부

'''