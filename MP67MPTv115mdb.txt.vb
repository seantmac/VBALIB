===============
basApp
===============
Option Compare Database
Option Explicit

'RKP/04-24-13/V01
'********** START OPTIONS            **********
'Option Compare Database
'********** END   OPTIONS            **********
'********** START DLL DECLARATIONS   **********
'********** END   DLL DECLARATIONS   **********
'********** START PUBLIC CONSTANTS   **********
''Alias Null Values for each VB data type
'Public Const NULL_INTEGER = -32768
'Public Const NULL_LONG = -2147483648#
'Public Const NULL_SINGLE = -3.402823E+38
'Public Const NULL_DOUBLE = -1.7976931348623E+308
'Public Const NULL_CURRENCY = -922337203685477#
'Public Const NULL_STRING = ""
'Public Const NULL_DATE = #1/1/100#
'Public Const NULL_BYTE = 0
'********** END   PUBLIC CONSTANTS   **********
'********** START PUBLIC VARIABLES   **********
'********** END   PUBLIC VARIABLES   **********
'********** START PRIVATE CONSTANTS  **********
'********** END   PRIVATE CONSTANTS  **********
'********** START PRIVATE VARIABLES  **********
Private mlLastErr       As Long
Private msLastErr       As String
Private msErrSource     As String
Private msErrDesc       As String
'********** END   PRIVATE VARIABLES  **********
'********** START USER DEFINED TYPES **********
Private Type TYPE_TOTALS
    startRow As Integer
    endRow As Integer
    str As String
End Type
'********** END   USER DEFINED TYPES **********

Private Sub GetSAPECCPULPMASTERS()
   'This script runs the SAP Queries for the Pulp MPT
   'it assumes the assigned variant will be run for each query
   'file location and names are stored in the variants

   Dim session           As Object
   Dim sapInstance       As String
   Dim sStartMonthYear   As String
   Dim sFinishMonthYear  As String

   sapInstance = "E01"
   Set session = SAP_GetSession(sapInstance, False)
   If session Is Nothing Then
      MsgBox "No active SAP session found for """ & sapInstance & """." & vbNewLine & "Please log on to """ & sapInstance & """ and try the operation again.", vbExclamation, Application.CurrentProject.Name
   Else

      'Port Mapping Information
      session.findById("wnd[0]").resizeWorkingPane 155, 26, False
      session.findById("wnd[0]/tbar[0]/okcd").text = "SQ00"
      session.findById("wnd[0]/tbar[0]/btn[0]").press
      session.findById("wnd[0]/mbar/menu[5]/menu[0]").select
      session.findById("wnd[1]/usr/radRAD1").select
      session.findById("wnd[1]/tbar[0]/btn[2]").press
      session.findById("wnd[0]/tbar[1]/btn[19]").press
      session.findById("wnd[1]/tbar[0]/btn[29]").press
      session.findById("wnd[2]/usr/subSUB_DYN0500:SAPLSKBH:0600/btnAPP_WL_SING").press
      session.findById("wnd[2]/usr/subSUB_DYN0500:SAPLSKBH:0600/btn600_BUTTON").press
      session.findById("wnd[3]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = "COEREPORTS"
      session.findById("wnd[3]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").caretPosition = 10
      session.findById("wnd[3]/tbar[0]/btn[0]").press
      session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell").selectedRows = "0"
      session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell").doubleClickCurrentCell
      session.findById("wnd[0]/usr/ctxtRS38R-QNUM").text = "DRS_PulpPorts"
      session.findById("wnd[0]/usr/ctxtRS38R-QNUM").SetFocus
      session.findById("wnd[0]/usr/ctxtRS38R-QNUM").caretPosition = 14
      session.findById("wnd[0]/tbar[1]/btn[17]").press
      session.findById("wnd[1]/usr/ctxtRS38R-VARIANT").caretPosition = 0
      session.findById("wnd[1]/tbar[0]/btn[0]").press
      
      session.findById("wnd[0]/usr/ctxtSP$00009-LOW").text = ""
      session.findById("wnd[0]/usr/ctxtSP$00001-LOW").text = ""
      session.findById("wnd[0]/usr/txtSP$00004-LOW").text = ""
      session.findById("wnd[0]/usr/txtSP$00003-LOW").text = ""
      session.findById("wnd[0]/usr/ctxtSP$00005-LOW").text = ""
      session.findById("wnd[0]/usr/ctxtSP$00007-LOW").text = ""
      session.findById("wnd[0]/usr/ctxtSP$00002-LOW").text = ""
      
      session.findById("wnd[0]/tbar[1]/btn[8]").press
      session.findById("wnd[1]/usr/chkRSAQDOWN-COLUMN").Selected = True
      session.findById("wnd[1]/usr/chkRSAQDOWN-COLUMN").SetFocus
      session.findById("wnd[1]/tbar[0]/btn[0]").press
      session.findById("wnd[1]/tbar[0]/btn[0]").press
      session.findById("wnd[0]/tbar[0]/btn[3]").press
      
      'End User Data
      session.findById("wnd[0]/usr/ctxtRS38R-QNUM").text = "DRS_PULPEUSER"
      session.findById("wnd[0]/usr/ctxtRS38R-QNUM").SetFocus
      session.findById("wnd[0]/usr/ctxtRS38R-QNUM").caretPosition = 14
      session.findById("wnd[0]/tbar[1]/btn[17]").press
      session.findById("wnd[1]/usr/ctxtRS38R-VARIANT").caretPosition = 0
      session.findById("wnd[1]/tbar[0]/btn[0]").press
      
      session.findById("wnd[0]/usr/ctxtSP$00009-LOW").text = ""
      session.findById("wnd[0]/usr/ctxtSP$00001-LOW").text = ""
      session.findById("wnd[0]/usr/txtSP$00004-LOW").text = ""
      session.findById("wnd[0]/usr/txtSP$00003-LOW").text = ""
      session.findById("wnd[0]/usr/ctxtSP$00005-LOW").text = ""
      session.findById("wnd[0]/usr/ctxtSP$00007-LOW").text = ""
      session.findById("wnd[0]/usr/ctxtSP$00002-LOW").text = ""
      
      session.findById("wnd[0]/tbar[1]/btn[8]").press
      session.findById("wnd[1]/usr/chkRSAQDOWN-COLUMN").Selected = True
      session.findById("wnd[1]/usr/chkRSAQDOWN-COLUMN").SetFocus
      session.findById("wnd[1]/tbar[0]/btn[0]").press
      session.findById("wnd[1]/tbar[0]/btn[0]").press
      session.findById("wnd[0]/tbar[0]/btn[3]").press
      
      'Grade Data
      session.findById("wnd[0]/usr/ctxtRS38R-QNUM").text = "DRS_ZDMGRADE"
      session.findById("wnd[0]/tbar[1]/btn[8]").press
      session.findById("wnd[0]/usr/txtSP$00003-LOW").text = ""
      session.findById("wnd[0]/usr/ctxtSP$00001-LOW").text = ""
      session.findById("wnd[0]/usr/ctxtSP$00002-LOW").text = "35"
      session.findById("wnd[0]/usr/txt%PATH").text = "\\s02afs01\Towers-050\GSC COE PUBLIC\GSC COE Plan\PulpMPT\Grades.txt"
      session.findById("wnd[0]/usr/txt%PATH").SetFocus
      session.findById("wnd[0]/usr/txt%PATH").caretPosition = 68
      session.findById("wnd[0]/tbar[1]/btn[8]").press
      session.findById("wnd[1]/usr/chkRSAQDOWN-COLUMN").Selected = True
      session.findById("wnd[1]/usr/chkRSAQDOWN-COLUMN").SetFocus
      session.findById("wnd[1]/tbar[0]/btn[0]").press
      session.findById("wnd[1]/tbar[0]/btn[0]").press
      session.findById("wnd[0]/tbar[0]/btn[3]").press
      
      'Customer Data
      session.findById("wnd[0]/usr/ctxtRS38R-QNUM").text = "DRS_ORG_INFO"
      session.findById("wnd[0]/usr/ctxtRS38R-QNUM").SetFocus
      session.findById("wnd[0]/usr/ctxtRS38R-QNUM").caretPosition = 3
      session.findById("wnd[0]/tbar[1]/btn[8]").press
      session.findById("wnd[0]/usr/ctxtSP$00004-LOW").text = ""
      session.findById("wnd[0]/usr/ctxtSP$00003-LOW").text = "35"
      session.findById("wnd[0]/usr/ctxtSP$00001-LOW").text = ""
      session.findById("wnd[0]/usr/ctxtSP$00002-LOW").text = ""
      session.findById("wnd[0]/usr/txt%PATH").text = "\\s02afs01\Towers-050\GSC COE PUBLIC\GSC COE Plan\PulpMPT\PulpCustData.txt"
      session.findById("wnd[0]/usr/txt%PATH").SetFocus
      session.findById("wnd[0]/usr/txt%PATH").caretPosition = 74
      session.findById("wnd[0]/tbar[1]/btn[8]").press
      session.findById("wnd[1]/usr/chkRSAQDOWN-COLUMN").Selected = True
      session.findById("wnd[1]/usr/chkRSAQDOWN-COLUMN").SetFocus
      session.findById("wnd[1]/tbar[0]/btn[0]").press
      session.findById("wnd[1]/tbar[0]/btn[0]").press
      session.findById("wnd[0]/tbar[0]/btn[3]").press
      
      'Materials
      session.findById("wnd[0]/tbar[1]/btn[19]").press
      session.findById("wnd[1]/tbar[0]/btn[29]").press
      session.findById("wnd[2]/usr/subSUB_DYN0500:SAPLSKBH:0600/cntlCONTAINER1_FILT/shellcont/shell").selectedRows = "0"
      session.findById("wnd[2]/usr/subSUB_DYN0500:SAPLSKBH:0600/btnAPP_WL_SING").press
      session.findById("wnd[2]/usr/subSUB_DYN0500:SAPLSKBH:0600/btn600_BUTTON").press
      session.findById("wnd[3]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = "BMOS"
      session.findById("wnd[3]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").caretPosition = 4
      session.findById("wnd[3]/tbar[0]/btn[0]").press
      session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell").selectedRows = "0"
      session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell").doubleClickCurrentCell
      session.findById("wnd[0]/usr/ctxtRS38R-QNUM").text = "DRS_TPL_Matl"
      session.findById("wnd[0]/usr/ctxtRS38R-QNUM").SetFocus
      session.findById("wnd[0]/usr/ctxtRS38R-QNUM").caretPosition = 12
      session.findById("wnd[0]/tbar[1]/btn[8]").press
      session.findById("wnd[0]/usr/btn%_SP$00002_%_APP_%-VALU_PUSH").press
      session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "PRD3"
      session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "PRD4"
      session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").SetFocus
      session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").caretPosition = 4
      session.findById("wnd[1]/tbar[0]/btn[8]").press
      session.findById("wnd[0]/usr/ctxtSP$00003-LOW").text = ""
      session.findById("wnd[0]/usr/ctxtSP$00001-LOW").text = ""
      session.findById("wnd[0]/usr/ctxtSP$00004-LOW").text = "35"
      session.findById("wnd[0]/usr/txtSP$00006-LOW").text = ""
      session.findById("wnd[0]/usr/txtSP$00007-LOW").text = ""
      session.findById("wnd[0]/usr/txtSP$00005-LOW").text = ""
      session.findById("wnd[0]/usr/txtSP$00008-LOW").text = ""
      session.findById("wnd[0]/usr/txtSP$00010-LOW").text = ""
      session.findById("wnd[0]/usr/txtSP$00011-LOW").text = ""
      session.findById("wnd[0]/usr/txtSP$00009-LOW").text = ""
      session.findById("wnd[0]/usr/txtSP$00012-LOW").text = ""
      session.findById("wnd[0]/usr/txt%PATH").text = "\\s02afs01\Towers-050\GSC COE PUBLIC\GSC COE Plan\PulpMPT\Materials.txt"
      session.findById("wnd[0]/usr/txtSP$00012-LOW").SetFocus
      session.findById("wnd[0]/usr/txtSP$00012-LOW").caretPosition = 0
      session.findById("wnd[0]/tbar[1]/btn[8]").press
      session.findById("wnd[1]/usr/chkRSAQDOWN-COLUMN").Selected = True
      session.findById("wnd[1]/usr/chkRSAQDOWN-COLUMN").SetFocus
      session.findById("wnd[1]/tbar[0]/btn[0]").press
      session.findById("wnd[1]/tbar[0]/btn[0]").press
      session.findById("wnd[0]/tbar[0]/btn[3]").press
      session.findById("wnd[0]/tbar[0]/btn[3]").press
   End If
End Sub
Private Sub GetSAPAPOPULPFORECAST()

   Dim session           As Object
   Dim sapInstance       As String
   Dim sStartMonthYear   As String
   Dim sFinishMonthYear  As String


''' DateDiff("m", #8/10/2000#, #9/14/2000#)
''' DateAdd("d", 30, Date)
   sFinishMonthYear = Format(Month(DateAdd("m", 5, Now())), "00") & "/" & Format(Year(DateAdd("m", 5, Now())), "0000")
   sStartMonthYear = Format(Month(DateAdd("m", -6, Now())), "00") & "/" & Format(Year(DateAdd("m", -6, Now())), "0000")
   sapInstance = "A01"
   Set session = SAP_GetSession(sapInstance, False)
   If session Is Nothing Then
      MsgBox "No active SAP session found for """ & sapInstance & """." & vbNewLine & "Please log on to """ & sapInstance & """ and try the operation again.", vbExclamation, Application.CurrentProject.Name
   Else
    
    
      session.findById("wnd[0]/tbar[0]/btn[3]").press   'BACK
      session.findById("wnd[0]/tbar[0]/btn[3]").press   'BACK
      session.findById("wnd[0]/tbar[0]/btn[3]").press   'BACK
      session.findById("wnd[0]/tbar[0]/btn[3]").press   'BACK
      session.findById("wnd[0]/tbar[0]/btn[3]").press   'BACK
   
      session.findById("wnd[0]").resizeWorkingPane 155, 26, False
      session.findById("wnd[0]/tbar[0]/okcd").text = "LISTCUBE"
      session.findById("wnd[0]").sendVKey 0
      session.findById("wnd[0]/usr/ctxtP_DTA").text = "ZDD2LCBK"
      session.findById("wnd[0]/usr/ctxtP_DTA").caretPosition = 8
      session.findById("wnd[0]/tbar[1]/btn[8]").press
      session.findById("wnd[0]/tbar[1]/btn[25]").press
      session.findById("wnd[0]/tbar[1]/btn[26]").press
                                                            '//PRODUCT INFO
      session.findById("wnd[0]/usr/chkS001").Selected = True   'Basis Wt
      session.findById("wnd[0]/usr/chkS002").Selected = True   'Business Planning Code
      session.findById("wnd[0]/usr/chkS003").Selected = True   'Grade
      session.findById("wnd[0]/usr/chkS004").Selected = True   'Grade Group
      session.findById("wnd[0]/usr/chkS005").Selected = True   'Major Product Line
      session.findById("wnd[0]/usr/chkS006").Selected = True   'Product Number
      
                                                            '//CUSTOMER INFO
      session.findById("wnd[0]/usr/chkS017").Selected = True   'Credit Acct
      session.findById("wnd[0]/usr/chkS021").Selected = True   'End User
      session.findById("wnd[0]/usr/chkS023").Selected = True   'Shipto
      session.findById("wnd[0]/usr/chkS025").Selected = True   'Soldto
      session.findById("wnd[0]/usr/chkS020").Selected = True   'Domestic/Export Indicator
      
                                                            '//OTHER INFO
      session.findById("wnd[0]/usr/chkS033").Selected = True   'SalesRepID
      session.findById("wnd[0]/usr/chkS034").Selected = True   'APO Means of Transport (SHIPMODE)
      session.findById("wnd[0]/usr/chkS036").Selected = True   'Record Type
      session.findById("wnd[0]/usr/chkS037").Selected = True   'Request ID
      session.findById("wnd[0]/usr/chkS038").Selected = True   'Calendar Year and Month
      session.findById("wnd[0]/usr/chkS039").Selected = True   'Unit Of Measure
      session.findById("wnd[0]/usr/chkS058").Selected = True   'LT Stats Forecast
      session.findById("wnd[0]/usr/chkS059").Selected = True   'PreConsensus Forecast
      
      session.findById("wnd[0]/usr/chkS059").SetFocus
      session.findById("wnd[0]/tbar[1]/btn[8]").press
      session.findById("wnd[0]/usr/btn%_C002_%_APP_%-VALU_PUSH").press
      session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "FP"
      session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "PS"
      session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").SetFocus
      session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").caretPosition = 0
      session.findById("wnd[1]/tbar[0]/btn[8]").press
      session.findById("wnd[0]/usr/chkL_CE").Selected = False
      session.findById("wnd[0]/usr/chkL_NO").Selected = True
      
      session.findById("wnd[0]/usr/ctxtC038-LOW").text = sStartMonthYear      'e.g. "11/2015"
      session.findById("wnd[0]/usr/ctxtC038-HIGH").text = sFinishMonthYear    'e.g. "04/2016"
      
      session.findById("wnd[0]/usr/txtL_MX").text = ""
      session.findById("wnd[0]/usr/txtL_MX").SetFocus       'NULL OUT MAXIMUM NUMBER OF HITS
      session.findById("wnd[0]/usr/txtL_MX").caretPosition = 11
      session.findById("wnd[0]/tbar[1]/btn[8]").press
      
      session.findById("wnd[0]/tbar[1]/btn[33]").press
      session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").currentCellRow = -1
      session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectColumn "VARIANT"
      session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").contextMenu
      session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectContextMenuItem "&FILTER"
      session.findById("wnd[2]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = "/STM_PULPMPT"
      session.findById("wnd[2]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").SetFocus
      session.findById("wnd[2]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").caretPosition = 12
      session.findById("wnd[2]/tbar[0]/btn[0]").press
      session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "0"
      session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell
      
      session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell -1, "K____022"
      session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "K____022"
      session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu
      session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem "&FILTER"
      session.findById("wnd[1]").sendVKey 4
      session.findById("wnd[2]/usr/lbl[1,3]").caretPosition = 3
      session.findById("wnd[2]").sendVKey 2
      session.findById("wnd[1]/tbar[0]/btn[0]").press
      
      session.findById("wnd[0]/tbar[1]/btn[45]").press
      session.findById("wnd[1]/tbar[0]/btn[0]").press
      session.findById("wnd[1]/usr/ctxtDY_PATH").text = "\\s02afs01\Towers-050\GSC COE PUBLIC\GSC COE Plan\PulpMPT"
      session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "PulpForecast.txt"
      session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 16
      session.findById("wnd[1]/tbar[0]/btn[11]").press
      
      session.findById("wnd[0]/tbar[0]/btn[3]").press   'BACK
      session.findById("wnd[0]/tbar[0]/btn[3]").press   'BACK
      session.findById("wnd[0]/tbar[0]/btn[3]").press   'BACK
      session.findById("wnd[0]/tbar[0]/btn[3]").press   'BACK
      session.findById("wnd[0]/tbar[0]/btn[3]").press   'BACK
        
   End If
        
   MsgBox "SAP APO PULP FORECAST EXTRACT...  Complete.  "

End Sub


Public Function SAP_GetSession(ByVal sapInstance As String, ByVal reset As Boolean) As Object
'**********************************************
'Author  :  RKP
'Date/Ver:  11-09-12/V01
'Input   :
'Output  :
'Comments:
'**********************************************
    On Error GoTo Err_Handler

    Dim sapGuiAuto      As Object
    Dim sapApp          As Object
    Dim sapConn         As Object
    Dim sapGuiComponent As Object
    Dim sapSessionCount As Integer
    Dim childCtr        As Integer
    Dim err619          As Boolean
    Dim sapSessionFound As Boolean
    
    sapSessionFound = False
    Set SAP_GetSession = Nothing
    
    Set sapGuiAuto = GetObject("SAPGUI")
    If sapGuiAuto Is Nothing Then
        'MsgBox "No active SAP session found." & vbNewLine & "Please log on to: """ & sapInstance & """ and try the operation again.", vbExclamation, Application.ThisWorkbook.Name
    Else
        Set sapApp = sapGuiAuto.GetScriptingEngine
        If sapApp Is Nothing Then
            'MsgBox "No active SAP session found." & vbNewLine & "Please log on to: """ & sapInstance & """ and try the operation again.", vbExclamation, Application.ThisWorkbook.Name
        Else
            sapSessionCount = sapApp.Children.count
            For childCtr = 0 To sapSessionCount - 1
                DoEvents
                Set sapGuiComponent = sapApp.Children.ElementAt(childCtr)
                If sapGuiComponent.Children.count > 0 Then
                    If sapGuiComponent.Children.ElementAt(0).Info.systemName = sapInstance Then
                        Set SAP_GetSession = sapGuiComponent.Children.ElementAt(0)
                        err619 = False
                        Debug.Print SAP_GetSession.findById("wnd[0]/usr/txtRSYST-BNAME").text
                        If sapSessionFound Then
                            Set SAP_GetSession = sapGuiComponent.Children.ElementAt(0)
                            If reset Then
                                SAP_GetSession.findById("wnd[0]/tbar[0]/btn[3]").press
                                SAP_GetSession.findById("wnd[0]/tbar[0]/btn[3]").press
                                SAP_GetSession.findById("wnd[0]/tbar[0]/btn[3]").press
                                SAP_GetSession.findById("wnd[0]/tbar[0]/btn[3]").press
                                SAP_GetSession.findById("wnd[0]/tbar[0]/btn[3]").press
                                SAP_GetSession.findById("wnd[0]/tbar[0]/btn[3]").press
                            End If
                        Else
                            Set SAP_GetSession = Nothing
                        End If
                    End If
                End If
            Next
            
'            Set sapConn = sapApp.Children(0)
'            If sapConn Is Nothing Then
'                MsgBox "No active SAP session found." & vbNewLine & "Please log on to: """ & sapInstance & """ and try the operation again.", vbExclamation, Application.ThisWorkbook.Name
'            Else
'                Set SAP_GetSession = sapConn.Children(0)
'            End If
        End If
    End If

Err_Handler:
    mlLastErr = Err.Number
    msLastErr = Err.Description
    'Function1 = mlLastErr
    If Err Then
        If Err.Number = 49 Then 'Bad DLL calling convention
            mlLastErr = 0
            msLastErr = ""
            Resume Next
        Else
            If Err.Number = 619 Then
                err619 = True
                If err619 Then
                    sapSessionFound = True
                    Resume Next
                End If
            Else
                'ProcessMsg Err.Number, Err.Description, "", ""
                'MsgBox Err.Number & " - " & Err.Description
            End If
        End If
    End If

    Exit Function
    Resume
End Function


Public Function UnwindTable(sInputTable As String, sOutputTable As String, iStaticCols As Integer, iUnwindCols As Integer)

Dim sSQL As String
Dim iOutputCounter As Long
Dim iInputCounter As Long
Dim iInputStaticCounter As Long
Dim RS As ADODB.Recordset
Dim rs2 As ADODB.Recordset


'sInputTable = "tbl004MasterData020SeasonalityEdit"
'sOutputTable = "tbl004MasterData021SeasonalityWorking"
'iStaticCols = 2
'iUnwindCols = 12

Set RS = Application.CurrentProject.Connection.Execute("select * from " & sInputTable & " where 0=1")
Set rs2 = Application.CurrentProject.Connection.Execute("select * from " & sOutputTable & " where 0=1")


DoCmd.SetWarnings False

iInputCounter = iStaticCols

DoCmd.RunSQL "Delete * from " & sOutputTable

Do Until iInputCounter = iStaticCols + iUnwindCols
    iOutputCounter = 0
    iInputStaticCounter = 0
    
    sSQL = "INSERT INTO " & sOutputTable & "  ( "
    
    'add output table fields to sql
    Do Until iOutputCounter = iStaticCols + 2
        sSQL = sSQL & rs2.Fields(iOutputCounter).Name & ", "
        iOutputCounter = iOutputCounter + 1
    Loop

    sSQL = Left(sSQL, Len(sSQL) - 2) & " ) SELECT "
    
    'add input static fields to sql
    Do Until iInputStaticCounter = iStaticCols
        sSQL = sSQL & RS.Fields(iInputStaticCounter).Name & ", "
        iInputStaticCounter = iInputStaticCounter + 1
    Loop
    
    sSQL = sSQL & "'" & RS.Fields(iInputCounter).Name & "'" & ", [" & RS.Fields(iInputCounter).Name & "] FROM " & sInputTable & " WHERE [" & RS.Fields(iInputCounter).Name & "] Is Not Null"
    
    '"INSERT INTO " & sOutputTable & "  ( "Plant, Grade, SeasMonth, SeasValue ) SELECT Plant, Grade," & iCounter & ", [" & iCounter & "] FROM " & sInputTable
        
    DoCmd.RunSQL sSQL
    
    
    iInputCounter = iInputCounter + 1
    
Loop

DoCmd.SetWarnings True

End Function

Public Function unwindtabletest()

DoCmd.SetWarnings False
DoCmd.RunSQL "DELETE * FROM tblFreightLeg2;"
UnwindTable "tblFreightLeg2_Xtb", "tblFreightLeg2", 2, 8
DoCmd.SetWarnings True

End Function


Sub AllCodeToDesktop()
''The reference for the FileSystemObject Object is Windows Script Host Object Model
''but it not necessary to add the reference for this procedure.

Dim fs As Object
Dim f As Object
Dim strMod As String
Dim mdl As Object
Dim i As Integer
Dim sfolder As String


Set fs = CreateObject("Scripting.FileSystemObject")
sfolder = "C:\Users\smacder.NAIPAPER\Desktop"  'no trailing \
Set f = fs.CreateTextFile(sfolder & "\" & Replace(CurrentProject.Name, ".", "") & ".txt")

''For each component in the project ...
For Each mdl In VBE.ActiveVBProject.VBComponents
    ''using the count of lines ...
    i = VBE.ActiveVBProject.VBComponents(mdl.Name).codemodule.CountOfLines
    ''put the code in a string ...
    If i > 0 And mdl.Name <> "basUtility" And mdl.Name <> "InterfaceCommon" And mdl.Name <> "modExcelPivot" Then
       strMod = VBE.ActiveVBProject.VBComponents(mdl.Name).codemodule.Lines(1, i)
    End If
    ''and then write it to a file, first marking the start with
    ''some equal signs and the component name.
    f.WriteLine String(15, "=") & vbCrLf & mdl.Name _
        & vbCrLf & String(15, "=") & vbCrLf & strMod
Next

''Close eveything
f.Close
Set fs = Nothing
End Sub


Public Function CreatePivotTable()
   Dim xlWB As Workbook

   Set xlWB = GeneratePivotTable_Generic(, , , 13)
   Set xlWB = GeneratePivotTable_Generic(, , , 14, xlWB)
End Function


Public Function GetVersion() As String
'**********************************************
'Author  :  RKP
'Date/Ver:  03-12-13/V01
'Input   :
'Output  :
'Comments:
'**********************************************
    On Error GoTo Err_Handler

'    Dim version     As String
'    Dim frm         As Form_MainPage
'    Dim ctr         As Integer
    
    GetVersion = "v" & GetScalarValue("SELECT KeyValue FROM tsysSettings WHERE KeyName = 'Version'")
    
    'Set frm = Forms("MainPage")
    'frm.lblVersion = version
'    Set frm = New Form_MainPage
'    frm.lblVersion.Caption = "v" & version
'
'    frm.Refresh
'    frm.Repaint
    
'    For Each frm In Application.Forms
'        Debug.Print frm.Name
'        'frm.Controls("lblVersion").value = version
'        For ctr = 0 To frm.Controls.count - 1
'            If frm.Controls.Item(ctr).value = "lblVersion" Then
'                Debug.Print frm.Controls.Item(ctr).value
'            End If
'        Next
'    Next

Err_Handler:
    mlLastErr = Err.Number
    msLastErr = Err.Description
    'Function1 = mlLastErr
    If Err Then
        If Err.Number = 49 Then 'Bad DLL calling convention
            mlLastErr = 0
            msLastErr = ""
            Resume Next
        Else
            'ProcessMsg Err.Number, Err.Description, "", ""
            MsgBox Err.Number & " - " & Err.Description
        End If
    End If

    Exit Function
    Resume
End Function


Function GetHaversineMiles(lat1Degrees As Double, lon1Degrees As Double, lat2Degrees As Double, lon2Degrees As Double) As Double
    Dim earthSphereRadiusKilometers As Double
    Dim kilometerConversionToMilesFactor As Double
    Dim lat1Radians As Double
    Dim lon1Radians As Double
    Dim lat2Radians As Double
    Dim lon2Radians As Double
    Dim AsinBase As Double
    Dim DerivedAsin As Double
    'Mean radius of the earth (replace with 3443.89849 to get nautical miles)
    earthSphereRadiusKilometers = 6371
    'Convert kilometers into miles
    kilometerConversionToMilesFactor = 0.621371
    'Convert each decimal degree to radians
    lat1Radians = (lat1Degrees / 180) * 3.14159265359
    lon1Radians = (lon1Degrees / 180) * 3.14159265359
    lat2Radians = (lat2Degrees / 180) * 3.14159265359
    lon2Radians = (lon2Degrees / 180) * 3.14159265359
    AsinBase = Sin(Sqr(Sin((lat1Radians - lat2Radians) / 2) ^ 2 + Cos(lat1Radians) * Cos(lat2Radians) * Sin((lon1Radians - lon2Radians) / 2) ^ 2))
    DerivedAsin = (AsinBase / Sqr(-AsinBase * AsinBase + 1))
    'Get distance from [lat1,lon1] to [lat2,lon2]
    'KM:    = Round(2 * DerivedAsin * earthSphereRadiusKilometers, 2)
    'Miles: = Round(2 * DerivedAsin * (earthSphereRadiusKilometers * kilometerConversionToMilesFactor), 2)
    GetHaversineMiles = Round(2 * DerivedAsin * (earthSphereRadiusKilometers * kilometerConversionToMilesFactor), 2)
End Function


Function PRep()  'Print Reports
    DoCmd.OpenReport "repKI", acViewNormal
    DoCmd.OpenReport "repCAPMachineDaysandCost", acViewNormal
End Function

Function PRep01()
    DoCmd.OpenReport "repKI-PROSHEET", acViewNormal
    DoCmd.OpenReport "repKI", acViewNormal
    DoCmd.OpenReport "repCAPMillProductionSummarySGrade", acViewNormal
    'DoCmd.OpenReport "repCAPMillProduction", acViewNormal
    DoCmd.OpenReport "repCAPAvgBW", acViewNormal
    DoCmd.OpenReport "repCAPDemand_NotSatisfied", acViewNormal
End Function

Function PRep02()
    DoCmd.OpenReport "repKI-PROSHEET", acViewNormal
    DoCmd.OpenReport "repKI", acViewNormal
    DoCmd.OpenReport "repCAPMillProductionSummarySGrade", acViewNormal
End Function

Function PRepExec()
    DoCmd.OpenReport "repKI-PROSHEET", acViewNormal
    DoCmd.OpenReport "repKI", acViewNormal
    DoCmd.OpenReport "repCAPMachineDaysandCost", acViewNormal
'    DoCmd.OpenReport "repCAPMillProductionSummarySGrade", acViewNormal
    DoCmd.OpenReport "repCAPMillProduction", acViewNormal
    DoCmd.OpenReport "repTopTen", acViewNormal
    DoCmd.OpenReport "repCAPAvgBW", acViewNormal
    DoCmd.OpenReport "repCAPDemand_NotSatisfied", acViewNormal
End Function

Function PRepExec2Word()
    DoCmd.OutputTo acOutputReport, "repKI-PROSHEET", acFormatRTF, DLookup("RUN_NAME", "tsysRunName", "ID=1") & "-" & Format(Now(), "yyyymmdd-hhmm") & "r00 PROSHEET.RTF", False
    DoCmd.OutputTo acOutputReport, "repKI", acFormatRTF, DLookup("RUN_NAME", "tsysRunName", "ID=1") & "-" & Format(Now(), "yyyymmdd-hhmm") & "r01 Key Indicators.RTF", False
End Function
===============
basDocumentation
===============
Option Compare Database

'''''==========================================================
'''''TableExists function
'''''Takes: The name of a table whose existence we want to test.
'''''Returns: True if table is found, False otherwise.
'''''Description: Test that a table exists by trying to assign
'''''a TableDef to it.
'''''If the table has not been found, an error is generated and the
'''''TableExists function returns False; otherwise it is True.
'''''Created: 26 Nov 2007 by Denis Wright.
'''''Modified: 7 July 2010.
'''''==========================================================
'''Function TableExists(strTable As String) As Boolean
'''  Dim tdf As DAO.TableDef
'''
'''  TableExists = True
'''
'''  On Error GoTo Err_Handle
'''  Set tdf = CurrentDb.TableDefs(strTable)
'''
'''Err_Exit:
'''  Set tdf = Nothing
'''Exit Function
'''
'''Err_Handle:
'''  If Err.Number = 3265 Then 'table not found
'''    'exit without error message; the caller function
'''    'will create the table
'''  Else
'''    MsgBox "Error " & Err.Number & "; " & Err.Description
'''  End If
'''  TableExists = False
'''  Resume Err_Exit
'''
'''End Function
''==================================================================
''CreateztblTableFields function.
''Description: Creates a table to hold field names and
''properties. First tries to delete the table, to
''prevent an error when the SQL statement runs.
''Created 27 Nov 2007 by Denis Wright.
''Updated 21 Dec 2007 by Denis Wright.
''Updated 22 Apr 2008 by Denis Wright -- added Description field
''==================================================================
Function Create_ztblTableFields()
  Dim sSQL As String

  DoCmd.SetWarnings False

  On Error Resume Next
  DoCmd.RunSQL "DROP TABLE ztblTableFields"
  On Error GoTo 0

  sSQL = "CREATE TABLE ztblTableFields " _
    & "( TableFieldID COUNTER, " _
    & "TableName TEXT(50), " _
    & "FieldName TEXT(50), " _
    & "FieldType LONG, " _
    & "FieldRequired INTEGER, " _
    & "FieldDefault TEXT(50), " _
    & "FieldDescription TEXT(255) )"

  DoCmd.RunSQL sSQL
  DoCmd.SetWarnings True

End Function

''===================================================================
''CreateztblIndexes function.
''Description: Creates a table to hold field names and
''index properties. First tries to delete the table, to
''prevent an error when the SQL statement runs.
''Created 27 Nov 2007 by Denis Wright.
''Updated 22 Apr 2007 by Denis Wright -- Added IndexIgnoreNulls field
''===================================================================
Function Create_ztblIndexes()
  Dim sSQL As String

  DoCmd.SetWarnings False

  On Error Resume Next
  DoCmd.RunSQL "DROP TABLE ztblIndexes"
  On Error GoTo 0

  sSQL = "CREATE TABLE ztblIndexes " _
    & "( TableIndexID COUNTER, " _
    & "TableName TEXT(50), " _
    & "IndexName TEXT(50), " _
    & "IndexRequired INTEGER, " _
    & "IndexUnique INTEGER, " _
    & "IndexPrimary INTEGER, " _
    & "IndexForeign INTEGER, " _
    & "IndexIgnoreNulls INTEGER )"

  DoCmd.RunSQL sSQL
  DoCmd.SetWarnings True

End Function


''====================================================================
''ListTableFields function
''Takes: The name of a table to be processed.
''Returns: The name and selected properties of the table's fields.
''Created: 27 Nov 2007 by Denis Wright
''Updated: 22 Apr 2008 by Denis Wright -- added Description property
''====================================================================
Function ListTableFields(strTableName As String)
  Dim dbs As DAO.Database
  Dim tdf As DAO.TableDef
  Dim rst As DAO.Recordset
  Dim fld As DAO.Field

  Set dbs = CurrentDb()

  'define the tabledef and write field properties to ztblTableFields
  Set tdf = dbs.TableDefs(strTableName)
  Set rst = dbs.TableDefs("ztblTableFields").OpenRecordset
  On Error Resume Next
  For Each fld In tdf.Fields
    rst.AddNew
    rst!tableName = tdf.Name
    rst!fieldName = fld.Name
    rst!FieldType = fld.Type
    rst!FieldSize = fld.FieldSize
    rst!FieldRequired = fld.Required
    rst!FieldDefault = fld.DefaultValue
    rst!FieldUpdatable = fld.DataUpdatable
    rst!FieldDescription = fld.Properties("Description")
    rst.Update
  Next fld

  rst.Close
  Set rst = Nothing
  Set tdf = Nothing
  Set dbs = Nothing

  On Error GoTo 0
End Function

''====================================================================
''ListTableIndexes function
''Takes: The name of a table to be processed.
''Returns: The name and selected properties of the table's indexes.
''Created: 27 Nov 2007 by Denis Wright
''====================================================================
Function ListTableIndexes(strTableName As String)
  Dim dbs As DAO.Database
  Dim tdf As DAO.TableDef
  Dim rst As DAO.Recordset
  Dim idx As DAO.Index

  Set dbs = CurrentDb()

  'define the tabledef and write field properties to ztblTableFields
  Set tdf = dbs.TableDefs(strTableName)
  Set rst = dbs.TableDefs("ztblIndexes").OpenRecordset
  For Each idx In tdf.Indexes
    rst.AddNew
    rst!tableName = tdf.Name
    rst!IndexName = idx.Name
    rst!IndexRequired = idx.Required
    rst!IndexUnique = idx.Unique
    rst!IndexPrimary = idx.Primary
    rst!IndexForeign = idx.Foreign
    rst!IndexIgnoreNulls = idx.IgnoreNulls
    rst.Update
  Next idx

  rst.Close
  Set rst = Nothing
  Set tdf = Nothing
  Set dbs = Nothing
End Function



''=======================================================================
''DocumentTables function.
''Description: Checks to see if the two documenter tables exist. If not,
''they are created; otherwise they are cleared.
''The function then loops through the TableDefs collection and documents
''all tables except for system and temporary tables.
''Created: 27 Nov 2007 by Denis Wright.
''Modified: 21 Dec 2007 -- turn off warnings
''=======================================================================
Function DocumentTables()
  Dim dbs As DAO.Database
  Dim tdf As DAO.TableDef

  DoCmd.SetWarnings False

  'check that the documentor tables exist; if not, create them.
  'otherwise, remove contents of documenter tables before re-populating
  If Not TableExists("ztblTableFields") Then
    Call Create_ztblTableFields
  Else
    DoCmd.SetWarnings False
    DoCmd.RunSQL "DELETE * FROM ztblTableFields"
    DoCmd.SetWarnings True
  End If

  If Not TableExists("ztblIndexes") Then
    Call Create_ztblIndexes
  Else
    DoCmd.SetWarnings False
    DoCmd.RunSQL "DELETE * FROM ztblIndexes"
    DoCmd.SetWarnings True
  End If

  Set dbs = CurrentDb()

  'exclude non-system and temp tables, otherwise list field and index properties
  For Each tdf In dbs.TableDefs
    If Left(tdf.Name, 4) <> "Msys" And Left(tdf.Name, 4) <> "~TMP" Then
      Call ListTableFields(tdf.Name)
      Call ListTableIndexes(tdf.Name)
    End If
  Next tdf
  DoCmd.SetWarnings True

  Set dbs = Nothing
End Function

'Requires a reference to the Microsoft DAO 3.6 Object Library
Function WriteSQLForQueries()
  Dim dbs As DAO.Database
  Dim qry As DAO.QueryDef

  DoCmd.SetWarnings False
  Set dbs = CurrentDb()

  'remove contents of table before re-populating
  DoCmd.RunSQL "DELETE * FROM ztblQueries"

  'loop through queries, writing names and SQL to the table
  For Each qry In dbs.QueryDefs
    If Left(qry.Name, 1) <> "~" Then
      DoCmd.RunSQL "INSERT INTO ztblQueries ( QueryName, QuerySQL, QueryType ) " _
        & "SELECT '" & qry.Name & "' AS qName, '" & qry.sql & "' AS qSQL, " & qry.Type & " AS qType;"
    End If
  Next qry
  DoCmd.SetWarnings True

  'clean up references
  Set dbs = Nothing

End Function




Function WriteTableList()
  ''=========================================================''
  ''Writes name of all tables
  ''to ztblTables. Table is cleared first.
  ''=========================================================''
  Dim dbs As DAO.Database
  Dim rst As DAO.Recordset
  Dim tdf As DAO.TableDef
  Dim sSQL As String
  
  sSQL = "DELETE * FROM ztblTables"
  
  'Clear out ztblTables to avoid duplication
  DoCmd.SetWarnings False
  DoCmd.RunSQL sSQL
  DoCmd.SetWarnings True

  Set dbs = CurrentDb()
  Set rst = dbs.TableDefs("ztblTables").OpenRecordset

  'Add table names to ztblTables. The If statement excludes system tables (MSYS*),
  'utility tables used in the documentation process (z*), and
  'temporary tables caused by deletion of tables in other routines (~*)
  For Each tdf In dbs.TableDefs 'exclude system, temp and accessory tables
    If Left(tdf.Name, 4) <> "MSYS" _
      And Left(tdf.Name, 1) <> "z" _
      And Left(tdf.Name, 1) <> "~" Then

      rst.AddNew
      rst!tableName = tdf.Name
      rst.Update
    End If
  Next tdf

  'cleanup
  rst.Close
  Set rst = Nothing
  Set tdf = Nothing
  Set dbs = Nothing
End Function


Function ReferencesInQueries()
  ''=========================================================''
  ''Runs through query names in ztblQueries to see what tables
  ''and queries are referenced. Any that are found are written
  ''out to ztblReferencedQueries
  ''Written 11 Sep 2007 by Denis Wright
  ''Modified 12 Sep 2007 to account for parent and child references
  ''=========================================================''
  Dim dbs As DAO.Database
  Dim rstMain As DAO.Recordset
  Dim rstTable As DAO.Recordset
  Dim rstQuery As DAO.Recordset
  Dim x, Y 'arrays to hold split data
  Dim sCheckParent As String, _
      sCheckChild As String
  Dim bChild As Boolean

  'define recordsets
  Set dbs = CurrentDb()
  Set rstMain = dbs.CreateQueryDef("", "SELECT * FROM ztblQueries").OpenRecordset
  Set rstTable = dbs.CreateQueryDef("", "SELECT TableName FROM ztblTables").OpenRecordset
  Set rstQuery = dbs.CreateQueryDef("", "SELECT QueryName FROM ztblQueries").OpenRecordset

  DoCmd.SetWarnings False
  'clean out existing data
  DoCmd.RunSQL "DELETE * FROM ztblReferencedQueries"

  'loop through all queries, checking against table and query names
  With rstMain
    .MoveFirst
    Do Until .EOF
      bChild = False
      If rstMain!QueryType = 80 Then 'Make-table query
        bChild = True
        x = Split(rstMain!QuerySQL, "FROM")
        Y = Split(x(0), "INTO")
        sCheckChild = Y(1)
        sCheckParent = Y(0)
        ElseIf rstMain!QueryType = 64 Then 'Append query
          bChild = True
          x = Split(rstMain!QuerySQL, "SELECT")
          sCheckChild = x(0)
          sCheckParent = x(1)
          Else
            sCheckParent = rstMain!QuerySQL
      End If

      'loop through all table names
      With rstTable
        .MoveFirst
        Do Until .EOF
          If bChild Then
            Call AddRefRecord(sCheckChild, rstTable!tableName, _
             rstMain!QueryName, "Child")
          End If
            Call AddRefRecord(sCheckParent, rstTable!tableName, _
             rstMain!QueryName, "Parent")
        .MoveNext
        Loop
      End With

      'loop through all query names
      With rstQuery
        .MoveFirst
        Do Until .EOF
          If bChild Then
            Call AddRefRecord(sCheckChild, rstQuery!QueryName, _
             rstMain!QueryName, "Child")
          End If
            Call AddRefRecord(sCheckParent, rstQuery!QueryName, _
             rstMain!QueryName, "Parent")
        .MoveNext
        Loop
      End With
    .MoveNext
    Loop
  End With

  DoCmd.SetWarnings True

  'clean up references
  rstQuery.Close
  rstTable.Close
  rstMain.Close
  Set rstQuery = Nothing
  Set rstTable = Nothing
  Set rstMain = Nothing
  Set dbs = Nothing
End Function


Function AddRefRecord(sTestString As String, sCompareString As String, sQueryName As String, _
  sRelationship As String) As Boolean
  Dim sSQL As String

  'this SQL creates ztblReferencedQueries on the fly if it doesn't exist
  sSQL = "CREATE TABLE ztblReferencedQueries ( " _
    & "RefID COUNTER, " _
    & "ObjectName TEXT(225), " _
    & "RefName TEXT (255), " _
    & "Relationship TEXT(30)); "
  If Not TableExists("ztblReferencedQueries") Then DoCmd.RunSQL (sSQL)
   
  On Error GoTo Err_Handle

  AddRefRecord = False
  If InStr(1, sTestString, sCompareString) > 0 Then
    'add the new record
    DoCmd.SetWarnings False
    sSQL = ""
    sSQL = "INSERT INTO ztblReferencedQueries ( ObjectName,RefName,Relationship ) " _
    & "VALUES ('" & sQueryName & "','" & sCompareString & "'" _
    & ",'" & sRelationship & "')"
    DoCmd.RunSQL sSQL
    
    '& ",'" & sRelationship & "')"
    '& ",'" & sRelationship & "','" & sObjectType & "')"
    '& ",'" & sRelationship & "')"
    
    DoCmd.SetWarnings True
    AddRefRecord = True
  End If
  
Err_Exit:
  Exit Function
Err_Handle:
  Select Case Err.Number
    Case 3192 'Table does not exist
    'create the table, then continue
    DoCmd.SetWarnings False
    DoCmd.RunSQL sSQL
    DoCmd.SetWarnings True
    Resume Next
  Case Else
    MsgBox "Error " & Err.Number & ": " & Err.Description
    Resume Err_Exit
  End Select
End Function



===============
basColors
===============
Option Compare Database

' modSystemColorToRGBColor
' 1999/12/21 From a news.devx.com, vb.general question I posed.
' Hi Larry,
' Declare Sub OleTranslateColor Lib "oleaut32.dll" (ByVal ColorIn As Long, ByVal
' hPal As Long, ByRef RGBColorOut As Long)
'
' And call
' Dim lpColor As Long
' Call OleTranslateColor(vbActiveTitleBar, 0, lpColor)
' lpColor will contain the RGB mapping when the call returns.
' Klaus H. Probst
    
DefLng A-Z
Private Const S_OK = &H0

#If Win64 And VBA7 Then
    Private Declare PtrSafe Function OleTranslateColor Lib "oleaut32.dll" (ByVal ColorIn As Long, ByVal hPal As Long, ByRef RGBColorOut As Long) As Long
#Else
    Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal ColorIn As Long, ByVal hPal As Long, ByRef RGBColorOut As Long) As Long
#End If

Dim miRed As Integer, miGreen As Integer, miBlue As Integer
Const csComma = ", "
    
Public Function ColorCodeToRGBValue(sColorCodeString As String) As Long
' 96/06/17 Return the RGB value from the color code string
' 96/06/18 Handle even in no space after the comma. Larry.
    Dim i As Integer
    Dim cs As String
    cs = Trim$(csComma)
    Dim sValue As String                'modified from another place in SIMPL
    sValue = sColorCodeString
    i = InStr(sValue, cs)               'get the red value
    miRed = Val(Left$(sValue, i - 1))
    sValue = Mid$(sValue, i + 1)
    i = InStr(sValue, cs)               'get the green value
    miGreen = Val(Left$(sValue, i - 1))
    sValue = Mid$(sValue, i + 1)
    miBlue = Val(sValue)                 'and remainder is blue
    ColorCodeToRGBValue = RGB(miRed, miGreen, miBlue)
End Function
Public Function ColorCodeToRGBString(lColorCode As Long) As String
'  96/06/17 Used to store colors in the registry, return: "128, 128, 255"
    Dim sTemp As String
    
    If ColorCodeToRGB(lColorCode, miRed, miGreen, miBlue) Then
        sTemp = miRed & csComma & miGreen & csComma & miBlue
        ColorCodeToRGBString = sTemp
    End If
End Function
Public Function ColorCodeToRGB(lColorCode As Long, iRed As Integer, iGreen As Integer, iBlue As Integer) As Boolean
' 96/01/16 Return the individual colors for lColorCode.
' Enter with:
'   lColorCode contains the color to be converted
'
' Return:
'   iRed      contains the red component
'   iGreen             the green component
'   iBlue              the blue component
'
' 96/07/15 Use Tip 171: Determining RGB Color Values, MSDN July 1996. Larry
    Dim lColor As Long          'Tip 171: Determining RGB Color Values, MSDN July 1996
    lColor = lColorCode         'work long
    iRed = lColor Mod &H100     'get red component
    lColor = lColor \ &H100     'divide
    iGreen = lColor Mod &H100   'get green component
    lColor = lColor \ &H100     'divide
    iBlue = lColor Mod &H100    'get blue component
    
    ColorCodeToRGB = True
End Function

Public Function RGBtoColorCode(iRed As Integer, iGreen As Integer, iBlue As Integer) As Long
' Return the long using iRed, iGreen and iBlue, same as the RGB function but added for completness
    On Error Resume Next
    RGBtoColorCode = RGB(iRed, iGreen, iBlue)
End Function

Public Function RGB_Red(lColorCode As Long) As Integer
' Return the red component of lColorCode
    If ColorCodeToRGB(lColorCode, miRed, miGreen, miBlue) Then
        RGB_Red = miRed
    End If
End Function

Public Function RGB_Green(lColorCode As Long) As Integer
' Return the green component of lColorCode
    If ColorCodeToRGB(lColorCode, miRed, miGreen, miBlue) Then
        RGB_Green = miGreen
    End If
End Function

Public Function RGB_Blue(lColorCode As Long) As Integer
' Return the blue component of lColorCode
    If ColorCodeToRGB(lColorCode, miRed, miGreen, miBlue) Then
        RGB_Blue = miBlue
    End If
End Function

Public Function SystemColorToRGBColor(lSystemColor As Long, lpColor As Long) As Boolean
    If OleTranslateColor(lSystemColor, 0, lpColor) = S_OK Then
        SystemColorToRGBColor = True
    End If
End Function

Private Function ColorLongToWords(ByVal lColour As Long) As String
'Converts an RGB colour number to a string equivalent
'By Si_the_geek, VBForums

  'ensure value is within range for colours
  lColour = lColour And &HFFFFFF

  'convert to separate RGB values
Dim iRed As Integer, iGreen As Integer, iBlue As Integer
  iRed = lColour And &HFF
  iGreen = (lColour \ &H100) And &HFF
  iBlue = lColour \ &H10000

  '"guess" the colour based on these values
Dim sColourName As String
  Select Case iRed
  Case Is > 170     'lots of red
      Select Case iGreen
      Case Is > 170     'lots of green
          Select Case iBlue  'blue:  lots/medium/little
          Case Is > 170:  sColourName = "White"
          Case Is > 85:   sColourName = "Bright Yellow"
          Case Else:      sColourName = "Yellow"
          End Select
      Case Is > 85      'medium green
          Select Case iBlue  'blue:  lots/medium/little
          Case Is > 170:  sColourName = "Pink"
          Case Is > 85:   sColourName = "Magenta"
          Case Else:      sColourName = "Orange"
          End Select
      Case Else         'little green
          Select Case iBlue  'blue:  lots/medium/little
          Case Is > 170:  sColourName = "Purple"
          Case Is > 85:   sColourName = "Dark Pink"
          Case Else:      sColourName = "Red"
          End Select
      End Select

  Case Is > 85      'medium red
      Select Case iGreen
      Case Is > 170     'lots of green
          Select Case iBlue  'blue:  lots/medium/little
          Case Is > 170:  sColourName = "Cyan"
          Case Is > 85:   sColourName = "Green"
          Case Else:      sColourName = "Bright Green"
          End Select
      Case Is > 85      'medium green
          Select Case iBlue  'blue:  lots/medium/little
          Case Is > 170:  sColourName = "Dark Blue"
          Case Is > 85:   sColourName = "Dark Grey"
          Case Else:      sColourName = "Dark Green"
          End Select
      Case Else         'little green
          Select Case iBlue  'blue:  lots/medium/little
          Case Is > 170:  sColourName = "Dark Blue"
          Case Is > 85:   sColourName = "Purple"
          Case Else:      sColourName = "Dark Red"
          End Select
      End Select

  Case Else         'little red
      Select Case iGreen
      Case Is > 170     'lots of green
          Select Case iBlue  'blue:  lots/medium/little
          Case Is > 170:  sColourName = "Cyan"
          Case Is > 85:   sColourName = "Green"
          Case Else:      sColourName = "Bright Green"
          End Select
      Case Is > 85      'medium green
          Select Case iBlue  'blue:  lots/medium/little
          Case Is > 170:  sColourName = "Blue"
          Case Is > 85:   sColourName = "Dark Cyan"
          Case Else:      sColourName = "Dark Green"
          End Select
      Case Else         'little green
          Select Case iBlue  'blue:  lots/medium/little
          Case Is > 170:  sColourName = "Bright Blue"
          Case Is > 85:   sColourName = "Dark Blue"
          Case Else:      sColourName = "Black"
          End Select
      End Select
   End Select

  ColorLongToWords = sColourName

End Function

'Private Sub ConvertIt()
'    Dim lColor As Long
'    Me.Picture1(0).BackColor = GetColorFromCBO()                        'get the system color selected
'    If SystemColorToRGBColor(Me.Picture1(0).BackColor, lColor) Then     'convert it
'        Me.Picture1(1).BackColor = lColor       'into ours
'        Me.Label1(1).Caption = "&&H" & Hex(lColor) & " = RGB(" & ColorCodeToRGBString(lColor) & ")"
'    Else                                        'failed, just in case
'        Me.Picture1(1).BackColor = vbBlack
'        Me.Label1(1).Caption = "&&H" & 0 & " = RGB(0, 0, 0)"
'        MsgBox "Conversion failed!", vbExclamation
'    End If
'End Sub

'Private Function GetColorFromCBO() As Long
'    With Me.Combo1          'get ItemData from the currently selected entry
'        GetColorFromCBO = .ItemData(.ListIndex)
'    End With
'End Function
'
'Private Sub Combo1_Click()
'    ConvertIt               'has changed so convert it
'End Sub

'Private Sub Form_Load()
'    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2 'simple center
'    AddSystemColorsToCBO    'add system color information to the combo box
'End Sub

Private Sub AddSystemColorsToCBO()
' From VB Documentation Topic: Color Constants
    AddToCBO "vbScrollBars", &H80000000           'Scroll bar color
    AddToCBO "vbDesktop", &H80000001              'Desktop color
    AddToCBO "vbActiveTitleBar", &H80000002       'Color of the title bar for the active window
    AddToCBO "vbInactiveTitleBar", &H80000003     'Color of the title bar for the inactive window
    AddToCBO "vbMenuBar", &H80000004              'Menu background color
    AddToCBO "vbWindowBackground", &H80000005     'Window background color
    AddToCBO "vbWindowFrame", &H80000006          'Window frame color
    AddToCBO "vbMenuText", &H80000007             'Color of text on menus
    AddToCBO "vbWindowText", &H80000008           'Color of text in windows
    AddToCBO "vbTitleBarText", &H80000009         'Color of text in caption, size box, and scroll arrow
    AddToCBO "vbActiveBorder", &H8000000A         'Border color of active window
    AddToCBO "vbInactiveBorder", &H8000000B       'Border color of inactive window
    AddToCBO "vbApplicationWorkspace", &H8000000C 'Background color of multiple-document interface (MDI) applications
    AddToCBO "vbHighlight", &H8000000D            'Background color of items selected in a control
    AddToCBO "vbHighlightText", &H8000000E        'Text color of items selected in a control
    AddToCBO "vbButtonFace", &H8000000F           'Color of shading on the face of command buttons
    AddToCBO "vbButtonShadow", &H80000010         'Color of shading on the edge of command buttons
    AddToCBO "vbGrayText", &H80000011             'Grayed (disabled) text
    AddToCBO "vbButtonText", &H80000012           'Text color on push buttons
    AddToCBO "vbInactiveCaptionText", &H80000013  'Color of text in an inactive caption
    AddToCBO "vb3DHighlight", &H80000014          'Highlight color for 3D display elements
    AddToCBO "vb3DDKShadow", &H80000015           'Darkest shadow color for 3D display elements
    AddToCBO "vb3DLight", &H80000016              'Second lightest of the 3D colors after vb3Dhighlight
    AddToCBO "vb3DFace", &H8000000F               'Color of text face
    AddToCBO "vb3DShadow", &H80000010             'Color of text shadow
    AddToCBO "vbInfoText", &H80000017             'Color of text in ToolTips
    AddToCBO "vbInfoBackground", &H80000018       'Background color of ToolTips
    'Me.Combo1.ListIndex = 6     'force first one
End Sub

Private Sub AddToCBO(sItem As String, lItemData As Long)
'    With Me.Combo1          'add the entry to the combo box
'        .AddItem sItem      'item
'        .ItemData(.NewIndex) = lItemData    'value
'    End With
End Sub


Function tacklemytable() As Integer
 Dim intColor As Long
 Dim rst As DAO.Recordset
 Dim dbs As DAO.Database
 Set dbs = CurrentDb
 Set rst = dbs.TableDefs("tlkColorCrayola").OpenRecordset
 On Error Resume Next
  
   rst.MoveFirst
   While rst.EOF = False
      intColor = RGB(rst!R, rst!G, rst!B)
      Debug.Print rst!id & "    " & intColor
      rst.MoveNext
   Wend
   tacklemytable = True
End Function


'Function colormath()
''Convert between Long, RGB, VB, and Web colors
''in VB, a long integer representing color is created from RGB values:
'
''''ERR//''''''Color = B   *  &FF00&     +   G   * &HFF&    +   R
'
''or this without Hex notation
'Color = (256 * 256 * B) + G * 256 + R
'
''the bult-in VB function RGB can calculate the Long value for you
'Color = RGB(R, G, B)
'End Function

Function Color_to_RGB(Color As Long, R As Integer, G As Integer, B As Integer) As Long
'to get the RGB from a long
  
  R = Color And &HFF&
  G = (Color And &HFF00&) \ &H100&
'''ERR//''''''  B = (Color & And &HFF0000) \ &H10000
   'or
  R = Color Mod 256
  G = (Color \ 256) Mod 256
  B = (Color \ 256 \ 256) Mod 256
End Function

Function colorsystemstuff()
'in VB colors > &H80000000 are systems colors - which must be
'interpreted by VB - they are not standard Long color values!
'use the GetSysColor API to return the true long value of a system color

'''Private Declare GetSysColor Lib "user32" ( ByVal nIndex As Long ) As Long
'''iColor = GetSysColor(iColor And &HFFFFFF)

'values of R, G, B can be used to  as BBGGRR to form a hex representation of a color that VB understands
'so, for R = "F0",  G="A3, and B = "2F, the hex representation in VB becomes BBGGRR:

'to get the VB hex string for a color from the Long or RGB
'''ERR//''''''  VBColorHexString = Right$( "000000" & Hex $( Color), 6)
'''ERR//''''''  VBColorHexString = Right$( "000000" & Hex $( R + 256 * (G + 256 * B ), 6)

'Note:  The Internet and other applications use RRGGBB for the hex format of a color
'''ERR//''''''  WebColorHexString = Right$( "000000" & Hex $( B + 256 * (G + 256 * R ), 6)

'to get the Web hex color string from the VB hex color string, swap the first 2 and last 2 character strings
'''ERR//''''''  WebColorHexString = Right$(VBHexColorString, 2) & Mid $(VBColorHexString, 3, 2) & Left$(VBHexColorString, 2)


End Function
===============
Form_frmAbout
===============
Option Compare Database     '** Use database order for string comparisons
Option Explicit             '** All variables must be declared before use

    Const mSYSTEM_RESOURCES = &H0
    Const mUSER_RESOURCES = &H2
    Const mENHANCED_MODE = &H20
    Const mWF_80x87 = &H400
    
    Dim mintWinFlags As Integer

Private Sub Button18_Click()
    Call cmdOK_Click
End Sub

Private Sub cmdOK_Click()
    DoCmd.Close
End Sub

Private Sub Form_Load()
   Me.Caption = "About The Market Pulp Optimizer"
   lblSystem_Name1.Caption = "C-OPT Optimization Software"
   lblSystem_Name2.Caption = "MEMPHIS, TN"
   lblCopyright.Caption = "Copyright, 1997-2013"
   
   lblSystem_Site.Caption = "MEMPHIS"
   lblSystem_Installed_Date.Caption = "JULY 6, 2012"
   lblSW_Version.Caption = "1.000"
   lblDB_Version.Caption = "NO DATABASE VERSION"
End Sub






===============
Form_frmDataCheck
===============
Private Sub btnCreatePriceOverRideTable_Click()
   DoCmd.OpenQuery "qChkPriceMktblPriceOverRides"
End Sub

Private Sub btnEditPriceOverRides_Click()
    DoCmd.OpenQuery "qChkPriceEditOverRides", acViewNormal, acEdit
End Sub

Private Sub btnRelaxSupplyLimitsForDemandThatCantBeMet_Click()
   DoCmd.SetWarnings False
   DoCmd.Hourglass True
   
   DoCmd.OpenQuery "qChkSalesTransferLimitConflict2"
   DoCmd.OpenQuery "qChkSalesTransferLimitConflict3"
   
   DoCmd.SetWarnings True
   DoCmd.Hourglass False
   
   MsgBox "Limits relaxed for lanes supplying demand that can't be met"
   
End Sub

Private Sub btnUpdPricesWithOverRides_Click()
   DoCmd.OpenQuery "qChkPriceUpdWithOverRides"
End Sub

Private Sub CloseCheckForDataIssuesForm_Click()
   DoCmd.Close acForm, Me.Name
End Sub


===============
Form_frmDataTables
===============
Option Compare Database

Private Sub CloseMaintainDataInputs_Click()
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub DemandSetMinEqMax_Click()
    DoCmd.RunSQL ("UPDATE tblDemand SET tblDemand.[Min] = [tblDemand].[Max]")
    Me.Requery
    Me.Refresh
End Sub

Private Sub DemandSetMinEqMaxForKeyAndPref_Click()
    DoCmd.RunSQL ("UPDATE tblDemand SET tblDemand.[Min] = [tblDemand].[Max]WHERE (([tblDemand].[CustClass]='40')) OR (([tblDemand].[CustClass]='50'))")
    Me.Requery
    Me.Refresh
End Sub

Private Sub DemandSetMinToZero_Click()
    DoCmd.RunSQL ("UPDATE tblDemand SET tblDemand.[Min] = 0")
    Me.Requery
    Me.Refresh
End Sub

Private Sub SetMinToPctOfMax_Click()
    DoCmd.OpenQuery ("qDemandSetMinToSpecifiedPctOfMax"), acViewNormal
    Me.Requery
    Me.Refresh
End Sub


===============
Form_frmMain
===============

Private Sub cboWorkingMonth_AfterUpdate()
   DoCmd.SetWarnings False
   DoCmd.RunSQL "UPDATE tbl001WorkingMonth SET tbl001WorkingMonth.WorkingMonth = (cboWorkingMonth)"
   DoCmd.Echo False
   DoCmd.Close acForm, "frmMain"
   DoCmd.OpenForm "frmMain"
   DoCmd.Echo True
   DoCmd.SetWarnings True
   MsgBox "Working Month Changed"
End Sub

Private Sub cmdDataCheck_Click()
    DoCmd.OpenForm "frmDataCheck"
End Sub


Private Sub cmdExcelPivotReports_Click()
   'was once command89
   Dim xlWB As Workbook
   Dim xlWS As Worksheet
   
   DoCmd.SetWarnings False
   
   RunQueriesByPrefix "q410Arch"  'Archive Table Main
   
   Set xlWB = GeneratePivotTable_Generic(, , , 30)      '30 FreightPivot
   Set xlWB = GeneratePivotTable_Generic(, , , 40)      '40 FCSTbyCustMill
   Set xlWB = GeneratePivotTable_Generic(, , , 50)      '50 FCSTbyExportFPMill
   Set xlWB = GeneratePivotTable_Generic(, , , 60)      '60 FCSTbyCountryMill
   Set xlWB = GeneratePivotTable_Generic(, , , 70)      '70 FCSTbyCarrierMill

   DoCmd.SetWarnings True
End Sub


Private Sub cmdHelp_Click()
    Dim sMsg As String
    sMsg = ""
    sMsg = sMsg & "(Preliminary - Before Using This Database) " & vbCrLf
    sMsg = sMsg & "1. ) " & vbCrLf
    sMsg = sMsg & "2. ( ) " & vbCrLf
    sMsg = sMsg & " " & vbCrLf
    sMsg = sMsg & "3. Go back to  Main Page and click Initialize Mapping " & vbCrLf
    sMsg = sMsg & "4. Click Missing Rate Check " & vbCrLf
    sMsg = sMsg & "          (this shows grade/bw's on our machines without a valid rate) " & vbCrLf
    sMsg = sMsg & "5. Click Assign Missing Mfg Loc to clear Missing mapping situations " & vbCrLf
    sMsg = sMsg & "          (on this page only fill in APO Loc for Ports, otherwise system will automatically populate with assigned Mill) " & vbCrLf
    sMsg = sMsg & " " & vbCrLf
   ''sMsg = ""
   sMsg = sMsg & "(Preliminary - Before Using This Database) " & vbCrLf
   sMsg = sMsg & "1. Run SAP Demand Planning Script to update Historical data" & vbCrLf
   sMsg = sMsg & "" & vbCrLf
   sMsg = sMsg & "" & vbCrLf
   sMsg = sMsg & "(Initial Data Import and Mapping - This Database)" & vbCrLf
   sMsg = sMsg & "1. Set new planning cycle by typing over current planning cycle "
   sMsg = sMsg & "     date (you can type in using the format yyyymm)" & vbCrLf
   sMsg = sMsg & "2. Click Import Data" & vbCrLf
   sMsg = sMsg & "3. Click Initialize" & vbCrLf
   sMsg = sMsg & "4. Review Master Data Tables for Missing Information" & vbCrLf
   sMsg = sMsg & "     (exceptions shown at the bottom of each tab)" & vbCrLf
   sMsg = sMsg & "5. Use Adjust Forecast form to view and change forecast" & vbCrLf
   sMsg = sMsg & "     (adjustments can be made at any level)" & vbCrLf
   
   MsgBox sMsg, vbOKOnly, "Demand Planning Basic Steps"
   
   
   
End Sub


Private Sub cmdImportNewData_Click()
    '//MAIN SCREEN -- IMPORT NEW DATA
    Dim sMsg As String
    sMsg = ""
    sMsg = sMsg & "THIS NEEDS TO BE COMPLETED." & vbCrLf
    sMsg = sMsg & "" & vbCrLf
    sMsg = sMsg & "Automate Extracts and Automate Imports." & vbCrLf
    sMsg = sMsg & "Customers, Materials, Grades, Forecast, OTHERS." & vbCrLf
    sMsg = sMsg & "Then... " & vbCrLf
    sMsg = sMsg & "click the Initialize Button. " & vbCrLf
    
    MsgBox sMsg
End Sub


Private Sub cmdInit_Click()
   Dim sMsg As String
   sMsg = ""
   sMsg = sMsg & "********  W A R N I N G  ! ! !  ********" & vbCrLf
   
   sMsg = sMsg & "" & vbCrLf
   sMsg = sMsg & "Initialization Will Restart The Planning Cycle." & vbCrLf
   sMsg = sMsg & "Click CANCEL To Return To Main Page or OK To Proceed." & vbCrLf
   sMsg = sMsg & "[CODE IS COMMENTED OUT UNTIL DATA FEEDS ARE HOOKED UP." & vbCrLf
      
''   If MsgBox(sMsg, vbOKCancel, "Fluff Pulp Master Planning") = vbOK Then
''
''      ''WAS ''''InitializeMapping "frmMain"
''
''      RunQueriesByPrefix ("q010Init0")
''      MsgBox "Initialization Routine -- is still a work in Progress.  Ran q010Init0__ action queries in order. Complete."
''
''   End If
End Sub

Private Sub cmdITBMOSUseOnly_Click()
   'was a btnITBMOSUseOnly
   Dim sPasswordInput As String
   sPasswordInput = ""
   
   sPasswordInput = InputBox("Enter Code To Manage Display Of Access Controls")
   
   If UCase(sPasswordInput) = "SHOW" Then
       ShowNavigationPane
       ShowRibbon
   ElseIf UCase(sPasswordInput) = "HIDE" Then
       HideNavigationPane
       HideRibbon
   End If
End Sub

Private Sub cmdLoadLevelReport_Click()
   DoCmd.OpenReport "srepPMBalance", acViewPreview
End Sub

Private Sub cmdLockTable_Click()
   DoCmd.OpenTable "tblModelSourcingLocks", acViewNormal, acEdit
End Sub

Private Sub cmdMaintainDataTables_Click()
    DoCmd.OpenForm "frmDataTables"
    'DoCmd.Close acForm, "frmMain"
End Sub


Private Sub cmdMaintainMasterData_Click()
    DoCmd.OpenForm "frmMasterData"
    'DoCmd.Close acForm, "frmMain"
End Sub


Private Sub Command50_Click()
DoCmd.OpenForm "reportpage"
DoCmd.Close acForm, "mainpage"
End Sub


Private Sub cmdNewBalancePage_Click()
   DoCmd.OpenForm "frmChanges"
End Sub


Private Sub cmdOptModel_Click()
    '//MAIN SCREEN -- OPT MODEL INTERFACE
    Dim sMsg As String
    sMsg = ""
    sMsg = sMsg & "There is no Opt Model at Present." & vbCrLf
    sMsg = sMsg & "Future plans for this tool will include some of the following:  " & vbCrLf
    sMsg = sMsg & "   Automatic sourcing of all unlocked demands to Minimum Freight Source.  " & vbCrLf
    sMsg = sMsg & "   Optimized Sourcing via an optimization model accounting for capacity and other constraints.  " & vbCrLf
    sMsg = sMsg & "   Etc." & vbCrLf
    sMsg = sMsg & "   " & vbCrLf
    MsgBox sMsg
    
       'OPT RUN
    Dim scmd As String
    Dim sCOPTPathAndFile As String
    Dim db As DAO.Database
    Dim RS As ADODB.Recordset
    Dim dTons As Double
    Dim sHostName As String
    Dim bShiptoSoleSource As Boolean
    sHostName = Environ$("computername")
    Debug.Print sHostName
    
          
DoCmd.SetWarnings False


DoCmd.RunSQL "DELETE * FROM tlkTime"
DoCmd.RunSQL "DELETE * FROM tlkCustLoc"
DoCmd.RunSQL "DELETE * FROM tblDemand"
DoCmd.RunSQL "DELETE * FROM tblDPLimit"
DoCmd.RunSQL "DELETE * FROM tblFPLimit"
DoCmd.RunSQL "DELETE * FROM tblMillLimit"
DoCmd.RunSQL "DELETE * FROM tblMillMachProd"
DoCmd.RunSQL "DELETE * FROM tblSalesLanes"


RunQueriesByPrefix "qryPreOpt0"


DoCmd.SetWarnings False



DoCmd.RunSQL "DELETE * FROM tblDemand IN 'C:\OPTMODELS\MP67\MP67.mdb'"
DoCmd.RunSQL "DELETE * FROM TblDPLimit IN 'C:\OPTMODELS\MP67\MP67.mdb'"
DoCmd.RunSQL "DELETE * FROM TblFPLimit IN 'C:\OPTMODELS\MP67\MP67.mdb'"
DoCmd.RunSQL "DELETE * FROM TblFreightLeg1 IN 'C:\OPTMODELS\MP67\MP67.mdb'"
DoCmd.RunSQL "DELETE * FROM TblFreightLeg2 IN 'C:\OPTMODELS\MP67\MP67.mdb'"
DoCmd.RunSQL "DELETE * FROM TblMachineDays IN 'C:\OPTMODELS\MP67\MP67.mdb'"
DoCmd.RunSQL "DELETE * FROM TblMillLimit IN 'C:\OPTMODELS\MP67\MP67.mdb'"
DoCmd.RunSQL "DELETE * FROM TblMillMachProd IN 'C:\OPTMODELS\MP67\MP67.mdb'"
DoCmd.RunSQL "DELETE * FROM TblSalesLanes IN 'C:\OPTMODELS\MP67\MP67.mdb'"
DoCmd.RunSQL "DELETE * FROM tlkCustLoc IN 'C:\OPTMODELS\MP67\MP67.mdb'"
DoCmd.RunSQL "DELETE * FROM tlkMachine IN 'C:\OPTMODELS\MP67\MP67.mdb'"
DoCmd.RunSQL "DELETE * FROM tlkMaterial IN 'C:\OPTMODELS\MP67\MP67.mdb'"
DoCmd.RunSQL "DELETE * FROM tlktime IN 'C:\OPTMODELS\MP67\MP67.mdb'"



DoCmd.RunSQL "INSERT INTO tblDemand IN 'C:\OPTMODELS\MP67\MP67.mdb' SELECT * FROM tblDemand"
DoCmd.RunSQL "INSERT INTO TblDPLimit IN 'C:\OPTMODELS\MP67\MP67.mdb' SELECT * FROM TblDPLimit"
DoCmd.RunSQL "INSERT INTO TblFPLimit IN 'C:\OPTMODELS\MP67\MP67.mdb' SELECT * FROM TblFPLimit"
DoCmd.RunSQL "INSERT INTO TblFreightLeg1 IN 'C:\OPTMODELS\MP67\MP67.mdb' SELECT * FROM TblFreightLeg1"
DoCmd.RunSQL "INSERT INTO TblFreightLeg2 IN 'C:\OPTMODELS\MP67\MP67.mdb' SELECT * FROM TblFreightLeg2"
DoCmd.RunSQL "INSERT INTO TblMachineDays IN 'C:\OPTMODELS\MP67\MP67.mdb' SELECT * FROM TblMachineDays"
DoCmd.RunSQL "INSERT INTO TblMillLimit IN 'C:\OPTMODELS\MP67\MP67.mdb' SELECT * FROM TblMillLimit"
DoCmd.RunSQL "INSERT INTO TblMillMachProd IN 'C:\OPTMODELS\MP67\MP67.mdb' SELECT * FROM TblMillMachProd"
DoCmd.RunSQL "INSERT INTO TblSalesLanes IN 'C:\OPTMODELS\MP67\MP67.mdb' SELECT * FROM TblSalesLanes"
DoCmd.RunSQL "INSERT INTO TlkCustLoc IN 'C:\OPTMODELS\MP67\MP67.mdb' SELECT * FROM tlkCustLoc"
DoCmd.RunSQL "INSERT INTO tlkMachine IN 'C:\OPTMODELS\MP67\MP67.mdb' SELECT * FROM tlkMachine"
DoCmd.RunSQL "INSERT INTO tlkMaterial IN 'C:\OPTMODELS\MP67\MP67.mdb' SELECT * FROM tlkMaterial"
DoCmd.RunSQL "INSERT INTO tlktime IN 'C:\OPTMODELS\MP67\MP67.mdb' SELECT * FROM tlktime"
'q410ModelPrep296ShiptoSoldtoProdGroupAllowSplits


 

'''   Select Case "MIP"
'''      Case "MIP"
'''         'Do SoleSource Demand Points
'''         DoCmd.RunSQL "UPDATE tsysDefCol IN 'C:\OPTMODELS\IPR2\IPR2.mdb' SET tsysDefCol.BNDBinary = True WHERE (((tsysDefCol.ColType)=""SALESON""))"
'''         'Follow Check Box to determine if SoleSourcing Shipto/SoldTo/Prodgroup
'''         If bShiptoSoleSource = True Then
'''            DoCmd.RunSQL "UPDATE tsysDefCol IN 'C:\OPTMODELS\IPR2\IPR2.mdb' SET tsysDefCol.BNDBinary = True WHERE (((tsysDefCol.ColType)=""SHIPTOON""))"
'''         Else
'''            DoCmd.RunSQL "UPDATE tsysDefCol IN 'C:\OPTMODELS\IPR2\IPR2.mdb' SET tsysDefCol.BNDBinary = False WHERE (((tsysDefCol.ColType)=""SHIPTOON""))"
'''         End If
'''      Case "CONTINUOUS"
'''         DoCmd.RunSQL "UPDATE tsysDefCol IN 'C:\OPTMODELS\IPR2\IPR2.mdb' SET tsysDefCol.BNDBinary = False WHERE (((tsysDefCol.ColType)=""SALESON""))"
'''         DoCmd.RunSQL "UPDATE tsysDefCol IN 'C:\OPTMODELS\IPR2\IPR2.mdb' SET tsysDefCol.BNDBinary = False WHERE (((tsysDefCol.ColType)=""SHIPTOON""))"
'''      Case Else
'''
'''   End Select
   

   

   'Continuous
   '
   'DoCmd.RunSQL "UPDATE tsysDefCol in 'C:\OPTMODELS\IPR2\IPR2.mdb' SET tsysDefCol.BNDBinary = False WHERE (((tsysDefCol.ColType)=""SHIPTOON""))"

   'MsgBox "Data Send Complete"


   DoCmd.SetWarnings True
          
   sCOPTPathAndFile = "C:\Progra~1\BMOS\C-OPT\C-OPTConsole.exe " & _
      "RUN /PRJ MP67 /WORKDIR C:\OPTMODELS\MP67\OUTPUT /Solver CoinMP /Sense MIN -1 /SetRelativeGapTolerance " & CIntlNumber("0.001", 3) & _
      " /SetTimeLimit " & CIntlNumber("40.000", 3) & " 600 /NOPROMPT"
   If InStr(sHostName, "SMACDER") > 0 Then  'same command, different c-opt installation dir.
      sCOPTPathAndFile = "C:\Programs\MATH\C-OPTx64\C-OPTConsole.exe " & _
         "RUN /PRJ MP67 /WORKDIR C:\OPTMODELS\MP67\OUTPUT /Solver CoinMP /Sense MIN -1 /SetRelativeGapTolerance " & CIntlNumber("0.001", 3) & _
         " /SetTimeLimit " & CIntlNumber("40.000", 3) & " 600 /NOPROMPT"
   End If
       
   scmd = sCOPTPathAndFile
       Debug.Print scmd
     
'        If MsgBox("Click YES to Run Opt Model or NO to cancel.  Run will take about 5 minutes.", vbYesNo + vbQuestion) = vbYes Then
      
              'me.runcopt.
              DoCmd.Hourglass (True)
              ExecCmd (scmd)
              DoCmd.Hourglass (False)
'              MsgBox ("C-OPT run complete.")
'        End If


'Apply Results
    DoCmd.SetWarnings False
    DoCmd.RunSQL "select * INTO tmtxCol01_SALES FROM tmtxCol01_SALES in 'C:\OPTMODELS\MP67\MP67.mdb'"

    Set RS = Nothing
    Set RS = Application.CurrentProject.Connection.Execute("select sum(iif(activity is null, 0, activity)) as TOTTONS FROM tmtxCol01_SALES")

    If RS.Fields("TOTTONS").value = 0 Then
        MsgBox "Opt Run Failed, No Results To Import"
    Else
''        DoCmd.OpenQuery "q420ModelResultsMaxSourcingPathWithTons"
''        DoCmd.OpenQuery "q430ModelResults010ApplyOptSourcing"
''        If bDemandLimitToggle = True Then
''            DoCmd.OpenQuery "q430ModelResults020ApplyOptTons"
''            DoCmd.OpenQuery "q430ModelResults022ApplyOptTonsNullActivity"
''        End If
        DoCmd.SetWarnings True
        DoCmd.Echo False
        DoCmd.Close acForm, "FrmMain"
        DoCmd.OpenForm "FrmMain"
        DoCmd.Maximize
        DoCmd.Echo True
''        Forms("optModelInterface").ForceDCToggle.value = bForceDCToggle
''        Forms("optModelInterface").DemandLimitToggle.value = bDemandLimitToggle
''        Forms("optModelInterface").ForceDCOnlyToggle.value = bForceDCOnlyToggle
''        Forms("optModelInterface").chkShiptoSoleSource.value = bShiptoSoleSource

        MsgBox "Optimization Applied"
        'DoCmd.OpenForm "FrmPGOutput"
    End If

'Me.Refresh
End Sub


Private Sub cmdTableMain_Click()
   DoCmd.OpenTable "tblMain"
End Sub

Private Sub Form_Load()
   Dim RS As ADODB.Recordset
   
   DoCmd.Maximize
   Me.lblVersion.Caption = basUtility.GetVersion
   Set RS = Application.CurrentProject.Connection.Execute("SELECT * FROM tbl001WorkingMonth")
   Me.cboWorkingMonth.value = RS.Fields("WorkingMonth").value
   Set RS = Nothing
   Set RS = Application.CurrentProject.Connection.Execute("SELECT * FROM tbl001PlanningCycle")
   Me.txtPlanningCycle.value = RS.Fields("PlanningCycle").value
   Set RS = Nothing
End Sub

Public Sub InitializeMapping(sform)
   'MsgBox "This Action Will Take A Few Minutes To Complete"
''   UnwindTable "tbl004MasterData020SeasonalityEdit", "tbl004MasterData021SeasonalityWorking", 2, 12
''   UnwindTable "q010Init020aaZeroFillBoxForecastEdit", "tbl002RawData030BoxForecastWorking", 2, 12
''   UnwindTable "tbl004MasterData030BoxConstructionEdit", "tbl004MasterData030BoxConstructionWorking", 6, 6
''   DoCmd.SetWarnings True
''
''   RunQueriesByPrefix "q010Init"
''   DoCmd.OpenForm sform
''
''   UnwindTable "tblSeasonalityEdit", "tblSeasonalityWorking", 2, 12
''   RunQueriesByPrefix "q010init"
''   MsgBox "Initialization Complete"
'''''
'''''RunQueriesByPrefix ("q010Init0")
'''''
'''''MsgBox "Initialization Routine -- is still a work in Progress.  Ran q010Init0__ action queries in order. Complete."

End Sub


Private Sub lblShow_Click()
   'code to show the db stuff
   ShowNavigationPane
   ShowRibbon
End Sub


Private Sub lblShow_DblClick(Cancel As Integer)
   'code to HIDE the db stuff
   HideNavigationPane
   HideRibbon
End Sub


Private Sub txtPlanningCycle_AfterUpdate()
   Dim pc As String
   Dim sfolder As String
   Dim sTag As String
   
   DoCmd.SetWarnings False
   DoCmd.RunSQL "UPDATE tbl001PlanningCycle SET tbl001PlanningCycle.planningcycle = (txtPlanningCycle)"
   DoCmd.RunSQL "UPDATE tbl006lkpPlanningCycle SET tbl006lkpPlanningCycle.PlanningCycle = (txtPlanningCycle)"
   DoCmd.SetWarnings True
   
   sTag = Left(txtPlanningCycle, 4) & Format(Right(txtPlanningCycle, 2), "00")
   sfolder = sBasePath & "" & sTag & " Cycle"
   
   If Len(Dir(sfolder, vbDirectory)) = 0 Then
       MkDir sfolder
   End If
   MsgBox "Planning Cycle Changed"
End Sub
===============
Form_frmMasterData
===============
Option Compare Database

Private Sub CloseUpdateTheModelForm_Click()
    DoCmd.Close acForm, Me.Name
End Sub

===============
InterfaceSourcingLocks
===============
'RKP/09-15-11/V01
'********** START OPTIONS            **********
Option Explicit
Option Compare Database
'********** Global   OPTIONS         **********
Global Const sSrcLkCtlTableName As String = "tblSourceLockControlTables"
'********** END   OPTIONS            **********
'********** START DLL DECLARATIONS   **********
'********** END   DLL DECLARATIONS   **********
'********** START PUBLIC CONSTANTS   **********
'********** END   PUBLIC CONSTANTS   **********
'********** START PUBLIC VARIABLES   **********
'********** END   PUBLIC VARIABLES   **********
'********** START PRIVATE CONSTANTS  **********
'********** END   PRIVATE CONSTANTS  **********
'********** START PRIVATE VARIABLES  **********



Public Function ApplyMillSourcingLimitsGeneric()
'//=====================================================================================//
'/|   FUNCTION:  ApplyMillSourcingLimitsGeneric                                         //
'/| PARAMETERS:  -NONE-                                                                 //
'/|    RETURNS:  -NONE-                                                                 //
'/|    PURPOSE:  Run The Generic Sourcing Limits, which                                 //
'/|              creates records in the table of disallowed lanes (closed paths)        //
'/|      USAGE:  ApplyMillSourcingLimitsGeneric2()                                      //
'/|              Ravi uses Macro ModelApplyLocks to call via RunCode,                   //
'/|              after running the action queries that perform the                      //
'/|              "Do Not Move" and "Lock To" Locks                                      //
'/|         BY:  Anjali                                                                 //
'/|       DATE:  12/19/2013                                                             //
'/|    HISTORY:  Adapted from Ravi's CBK Generic NeverLocks,                            //
'/|              which was adapted from BR2 (Brazil) and                                //
'/|              the logic developed by Ray Squires.                                    //
'//=====================================================================================//
    On Error GoTo ApplyMillSourcingLimitsGeneric_Err:
   
    Dim rsSrcLkCtlTables As ADODB.Recordset 'recordsets
    Dim rsSrcLkCtlFields As ADODB.Recordset
    Dim rsSrcLkCtlOutputFields As ADODB.Recordset
    Dim rsSourcingLocks As ADODB.Recordset
    
    Dim sSelectFromTable As String
    Dim sSourceLockTable As String
    Dim sOutputTable As String
    Dim sOutputField As String
    Dim sSelectFromTableField As String
    Dim sSourceLockTableField As String
    Dim sAllPathsWithAttributesField As String
    Dim sSourceLockControlOutputFieldsTable As String
    Dim sSourceLockControlFieldsTable As String
    Dim sOperatorSign As String
    Dim SourceCtr As Long
    
    Dim ctr As Long
    Dim z As Date
    
    Dim sSQL As String 'sql statement parts
    Dim sDELETEsql As String
    Dim sINSERTpart As String
    Dim sSELECTpart As String
    Dim sFROMpart As String
    Dim sWHEREpart As String
    Dim sWHEREconditionPart As String
    Dim sSelectSLIDpart As String
    
       
    DoCmd.SetWarnings False
    z = Now()
    Set rsSrcLkCtlTables = Application.CurrentProject.Connection.Execute("SELECT * FROM " & sSrcLkCtlTableName) ' sSrcLkCtlTableName is a global constant
        sSelectFromTable = rsSrcLkCtlTables.Fields("SelectFromTable").value
        sSourceLockTable = rsSrcLkCtlTables.Fields("SourceLockTable").value
        sOutputTable = rsSrcLkCtlTables.Fields("OutputTable").value
        sSourceLockControlOutputFieldsTable = rsSrcLkCtlTables.Fields("SourceLockControlOutputFieldsTable").value
        sSourceLockControlFieldsTable = rsSrcLkCtlTables.Fields("SourceLockControlFieldsTable").value
    
    Set rsSrcLkCtlOutputFields = Application.CurrentProject.Connection.Execute("SELECT * FROM " & sSourceLockControlOutputFieldsTable)
        
    Set rsSrcLkCtlFields = Application.CurrentProject.Connection.Execute("SELECT * FROM " & sSourceLockControlFieldsTable)
        sSourceLockTableField = rsSrcLkCtlFields.Fields("SourceLockTableField").value
        sAllPathsWithAttributesField = rsSrcLkCtlFields.Fields("AllPathsWithAttributesField").value
   
    Set rsSourcingLocks = Application.CurrentProject.Connection.Execute("SELECT * FROM " & sSourceLockTable & " Order by SLID ")
        sDELETEsql = "Delete * FROM  " & sOutputTable & ";" & vbCrLf  ' clean the output table before the insert
       ' MsgBox "Deleting Old Closed Paths Table"
        DoCmd.RunSQL sDELETEsql
        
        sINSERTpart = "Insert into " & sOutputTable & " ( "
        sSELECTpart = "Select " & vbCrLf
        sFROMpart = "From " & sSelectFromTable & ", " & sSourceLockTable
        sWHEREpart = "Where " & vbCrLf
        
        While Not rsSrcLkCtlOutputFields.EOF
            sINSERTpart = sINSERTpart & rsSrcLkCtlOutputFields.Fields("OutputField") & " , "
            sSELECTpart = sSELECTpart & "   " & sSelectFromTable & "." & rsSrcLkCtlOutputFields.Fields("SelectFromTableField") & " , " & vbCrLf
            rsSrcLkCtlOutputFields.MoveNext
        Wend
        sINSERTpart = sINSERTpart & " SLID ) "   ' there is already an extra comma before the SLID might change code for SLID
        sSelectSLIDpart = ""
                 
        While Not rsSourcingLocks.EOF
            If rsSourcingLocks.Fields("Never").value = False And rsSourcingLocks.Fields("Always").value = False Then
            Else
               If rsSourcingLocks.Fields("Never").value = True And rsSourcingLocks.Fields("Always").value = True Then
               MsgBox ("Always & never can not both be checked at the same time, SLID: " & rsSourcingLocks.Fields("SLID") & " was skipped")
               Else
                  sSelectSLIDpart = "   '" & "srclk " & rsSourcingLocks.Fields("SLID") & "'" & " as SLID "
                  sSelectSLIDpart = sSELECTpart & sSelectSLIDpart
                  Set rsSrcLkCtlFields = Application.CurrentProject.Connection.Execute("SELECT * FROM " & sSourceLockControlFieldsTable)
                  sWHEREconditionPart = sWHEREpart & "   " & sSourceLockTable & "." & "SLID" & " = " & rsSourcingLocks.Fields("SLID") & " AND " & vbCrLf
                  ' Anjali 04/16/2014 - modified code to include Always logic in table Sourcing locks
                  SourceCtr = 0
                  While Not rsSrcLkCtlFields.EOF
                  ' For ctr = 0 To rsSourcingLocks.Fields.Count - 1
                      If rsSourcingLocks.Fields(rsSrcLkCtlFields.Fields("SourceLockTableField").value).value & "" <> "" Then
                         sOperatorSign = " = "
                         If rsSourcingLocks.Fields("Always").value = True Then
                            'Debug.Print rsSrcLkCtlFields.Fields("FieldType")
                            If rsSrcLkCtlFields.Fields("FieldType") = "Source" Then
                              SourceCtr = SourceCtr + 1
                              sOperatorSign = " <> "
                            End If
                         End If
                         sWHEREconditionPart = sWHEREconditionPart & "   " & sSelectFromTable & "." & rsSrcLkCtlFields.Fields("AllPathsWithAttributesField") & sOperatorSign & sSourceLockTable & "." & rsSrcLkCtlFields.Fields("SourceLockTableField") & " AND " & vbCrLf
                      End If
                 ' Next ctr
                  rsSrcLkCtlFields.MoveNext
                  Wend
              
                  sWHEREconditionPart = VBA.Left(VBA.Trim(sWHEREconditionPart), VBA.Len(VBA.Trim(sWHEREconditionPart)) - 7)
                  sSQL = sINSERTpart & vbCrLf & sSelectSLIDpart & vbCrLf & sFROMpart & vbCrLf & sWHEREconditionPart & ";" & vbCrLf
                  
                  If SourceCtr > 1 Then
                  Else
                     'Debug.Print SourceCtr
                     DoCmd.RunSQL sSQL
                     'Debug.Print sSQL
                  End If
                  
                  rsSrcLkCtlFields.Close
                  Set rsSrcLkCtlFields = Nothing
                  sSelectSLIDpart = ""
                  sWHEREconditionPart = ""
               End If
            End If
            rsSourcingLocks.MoveNext
                        
        Wend
    DoCmd.SetWarnings True
   ' MsgBox "New table has been created," & " Run time:  " & Format(Now() - z, "hh:nn:ss")
ApplyMillSourcingLimitsGeneric_Done:
  Exit Function

ApplyMillSourcingLimitsGeneric_Err:
   Select Case Err.Number
      Case 9 'or 13 'subscript out of range
         'Resume Next
      Case 3265
         MsgBox Err.Number & " - " & Err.Description & vbCrLf & " You have a wrong field in your control table "
      Case Else
         MsgBox Err.Number & " - " & Err.Description
   End Select
   Resume ApplyMillSourcingLimitsGeneric_Done
   Resume
End Function



===============
InterfaceCustom
===============
Option Explicit
Option Compare Database

'// Variables //
Global iNbrDecimals As Integer
'// Constants //
Global Const sfrmChangesControlTablePrefix As String = "tsysX"
Global Const sLabelFontName As String = "Calibri"
Global Const iLabelFontSize As String = 11            '// IPG MPT60 likes it at 11; 10 is default
Global Const sBasePath As String = "C:\OPTMODELS\MP67\"
Global Const sUOM As String = "Tons"

Global sCCVolume As String    'Change Candidates Volume Field
Global sFullQuickImport As String


Public Function WorkingMonthChange(sFormName As String, scmbboxdate As String, change As String)
   DoCmd.SetWarnings False
   DoCmd.RunSQL "UPDATE tbl001workingmonth SET tbl001workingmonth.workingmonth = """ & scmbboxdate & """"
   fSetSummaryBoxes sFormName, "Update"
   DoCmd.SetWarnings True
   
   Forms(sFormName).ForecastChangeArea.SourceObject = ""
   Forms(sFormName).VolumeChangeArea.SourceObject = ""
   Forms(sFormName).MoveArea.SourceObject = ""
   Forms(sFormName).DistMoveArea.SourceObject = ""
   Forms(sFormName).ViewOnlyArea.SourceObject = ""
   
   If change <> "FC" Then
      fSetComboBoxes sFormName, "FC", "FullRefresh"
   End If
   If change <> "VC" Then
      fSetComboBoxes sFormName, "VC", "FullRefresh"
   End If
   If change <> "MM" Then
      fSetComboBoxes sFormName, "MM", "FullRefresh"
   End If
   If change <> "VO" Then
      fSetComboBoxes sFormName, "VO", "FullRefresh"
   End If
   If change <> "DM" Then
      fSetComboBoxes sFormName, "DM", "FullRefresh"
   End If
   
   Forms(sFormName).Refresh
   
   On Error Resume Next
   Forms("MainPage")("WorkingMonthCombo").value = scmbboxdate
   
   MsgBox "Working Month Changed"

End Function

Public Sub CleanupForFC(sFormName As String)
'DoCmd.OpenQuery "qintUpdCleanUpForFC"
End Sub

Public Sub CleanupForVC(sFormName As String)
'DoCmd.OpenQuery "qintUpdCleanUpForVC"
End Sub

Public Sub CleanupForMM(sFormName As String)
   DoCmd.OpenQuery "qupdCleanUpMainCurrMill"
   ''?? IF THE CURRENT DISTPOINT IS ANY MILL
   ''   WRITE CURRENT MILL FIELD INTO CURRENT DIST POINT FIELD
   ''   E.G.--MAKE SURE CURR MILL AND CURR DIST PT ARE NOT BOTH MILLS AND UNMATCHED
   DoCmd.OpenQuery "qupdCleanUpMainCurrMillIfDistPtIsMill"
   'DoCmd.OpenQuery "q210CleanUp011AfterMM2"
End Sub

Public Sub CleanupForDM(sFormName As String, bReMapToggle As Boolean)
If sFormName = "fintChanges" Then
    If bReMapToggle = False Then
        ''?? IF THIS WAS REMAPPED TO A DIST POINT THAT *IS* A MILL, UPDATE
        ''   CURRENT MILL FIELD AND CURRENT MACHINE TO DIST POINT CODE AND
        ''   A MACHINE AT THAT MILL
        ''DoCmd.OpenQuery "q210CleanUp020AfterDM"
    Else
        DoCmd.OpenQuery "qupdCleanUpMillMachineAfterDM"
    End If
    ''DoCmd.OpenQuery "q210CleanUp025AfterDMSetK5"
    ''DoCmd.OpenQuery "q210CleanUp026AfterDMSetK6"
    ''DoCmd.OpenQuery "q210CleanUp027AfterDMSetK7"
End If
End Sub

Public Sub InitializeFormPages(sFormName As String)

Dim sSQL As String
iNbrDecimals = 0

   If sFormName = "frmChanges" Then
      iNbrDecimals = 2
      
      Forms(sFormName)("Summary1").Caption = "Summary 1"
      Forms(sFormName)("Summary2").Caption = "Summary 2"
      Forms(sFormName)("Summary3").Caption = "Summary 3"
      Forms(sFormName)("VO").Caption = "ViewOnly"
      Forms(sFormName)("DM").Caption = "DistPtMove"
      Forms(sFormName)("MM").Caption = "MachMove"
      Forms(sFormName)("FC").Caption = "FcastChg"
      Forms(sFormName)("VC").Caption = "VolumeChg"
       
      Forms(sFormName)("Summary1AreaA").SourceObject = "Query.qrepPMBalance" '"query.q910Report015PMBalanceAllMonths"
      Forms(sFormName)("Summary1AreaB").SourceObject = "Query.qrepInvBalance"
      Forms(sFormName)("Summary2AreaA").SourceObject = "Query.qrepMillByDistPoint"
      Forms(sFormName)("Summary2AreaB").SourceObject = "Query.qrepMillByProductProdLine"
      Forms(sFormName)("Summary3AreaA").SourceObject = "Query.qrepMillByCountry"
   
      Forms(sFormName)("Summary1").Visible = True
      Forms(sFormName)("Summary2").Visible = True
      Forms(sFormName)("Summary3").Visible = True
       
      ''sSQL = "Select *  into tblfintlkpDistPointCapability  from q810Lkp012DistPointCapability;"
      ''DoCmd.RunSQL sSQL
   End If

End Sub

Public Sub OpenDashboard(sFormName As String)
   If sFormName = "fintChanges" Then
      DoCmd.OpenForm "fintVolumeDashboard"
   End If
End Sub


===============
InterfaceCommon
===============
Option Explicit
Option Compare Database

'// Variables //
Global iNbrDecimals As Integer
'// Constants //
Global Const sfrmChangesControlTablePrefix As String = "tsysX"
Global Const sLabelFontName As String = "Calibri"
Global Const iLabelFontSize As String = 11            '// IPG MPT60 likes it at 11; 10 is default
Global Const sBasePath As String = "C:\OPTMODELS\MP67\"
Global Const sUOM As String = "Tons"

Global sCCVolume As String    'Change Candidates Volume Field
Global sFullQuickImport As String


Public Function WorkingMonthChange(sFormName As String, scmbboxdate As String, change As String)
   DoCmd.SetWarnings False
   DoCmd.RunSQL "UPDATE tbl001workingmonth SET tbl001workingmonth.workingmonth = """ & scmbboxdate & """"
   fSetSummaryBoxes sFormName, "Update"
   DoCmd.SetWarnings True
   
   Forms(sFormName).ForecastChangeArea.SourceObject = ""
   Forms(sFormName).VolumeChangeArea.SourceObject = ""
   Forms(sFormName).MoveArea.SourceObject = ""
   Forms(sFormName).DistMoveArea.SourceObject = ""
   Forms(sFormName).ViewOnlyArea.SourceObject = ""
   
   If change <> "FC" Then
      fSetComboBoxes sFormName, "FC", "FullRefresh"
   End If
   If change <> "VC" Then
      fSetComboBoxes sFormName, "VC", "FullRefresh"
   End If
   If change <> "MM" Then
      fSetComboBoxes sFormName, "MM", "FullRefresh"
   End If
   If change <> "VO" Then
      fSetComboBoxes sFormName, "VO", "FullRefresh"
   End If
   If change <> "DM" Then
      fSetComboBoxes sFormName, "DM", "FullRefresh"
   End If
   
   Forms(sFormName).Refresh
   
   On Error Resume Next
   Forms("MainPage")("WorkingMonthCombo").value = scmbboxdate
   
   MsgBox "Working Month Changed"

End Function

Public Sub CleanupForFC(sFormName As String)
'DoCmd.OpenQuery "qintUpdCleanUpForFC"
End Sub

Public Sub CleanupForVC(sFormName As String)
'DoCmd.OpenQuery "qintUpdCleanUpForVC"
End Sub

Public Sub CleanupForMM(sFormName As String)
   DoCmd.OpenQuery "qupdCleanUpMainCurrMill"
   ''?? IF THE CURRENT DISTPOINT IS ANY MILL
   ''   WRITE CURRENT MILL FIELD INTO CURRENT DIST POINT FIELD
   ''   E.G.--MAKE SURE CURR MILL AND CURR DIST PT ARE NOT BOTH MILLS AND UNMATCHED
   DoCmd.OpenQuery "qupdCleanUpMainCurrMillIfDistPtIsMill"
   'DoCmd.OpenQuery "q210CleanUp011AfterMM2"
End Sub

Public Sub CleanupForDM(sFormName As String, bReMapToggle As Boolean)
If sFormName = "fintChanges" Then
    If bReMapToggle = False Then
        ''?? IF THIS WAS REMAPPED TO A DIST POINT THAT *IS* A MILL, UPDATE
        ''   CURRENT MILL FIELD AND CURRENT MACHINE TO DIST POINT CODE AND
        ''   A MACHINE AT THAT MILL
        ''DoCmd.OpenQuery "q210CleanUp020AfterDM"
    Else
        DoCmd.OpenQuery "qupdCleanUpMillMachineAfterDM"
    End If
    ''DoCmd.OpenQuery "q210CleanUp025AfterDMSetK5"
    ''DoCmd.OpenQuery "q210CleanUp026AfterDMSetK6"
    ''DoCmd.OpenQuery "q210CleanUp027AfterDMSetK7"
End If
End Sub

Public Sub InitializeFormPages(sFormName As String)

Dim sSQL As String
iNbrDecimals = 0

   If sFormName = "frmChanges" Then
      iNbrDecimals = 2
      
      Forms(sFormName)("Summary1").Caption = "Summary 1"
      Forms(sFormName)("Summary2").Caption = "Summary 2"
      Forms(sFormName)("Summary3").Caption = "Summary 3"
      Forms(sFormName)("VO").Caption = "ViewOnly"
      Forms(sFormName)("DM").Caption = "DistPtMove"
      Forms(sFormName)("MM").Caption = "MachMove"
      Forms(sFormName)("FC").Caption = "FcastChg"
      Forms(sFormName)("VC").Caption = "VolumeChg"
       
      Forms(sFormName)("Summary1AreaA").SourceObject = "Query.qrepPMBalance" '"query.q910Report015PMBalanceAllMonths"
      Forms(sFormName)("Summary1AreaB").SourceObject = "Query.qrepInvBalance"
      Forms(sFormName)("Summary2AreaA").SourceObject = "Query.qrepMillByDistPoint"
      Forms(sFormName)("Summary2AreaB").SourceObject = "Query.qrepMillByProductProdLine"
      Forms(sFormName)("Summary3AreaA").SourceObject = "Query.qrepMillByCountry"
   
      Forms(sFormName)("Summary1").Visible = True
      Forms(sFormName)("Summary2").Visible = True
      Forms(sFormName)("Summary3").Visible = True
       
      ''sSQL = "Select *  into tblfintlkpDistPointCapability  from q810Lkp012DistPointCapability;"
      ''DoCmd.RunSQL sSQL
   End If

End Sub

Public Sub OpenDashboard(sFormName As String)
   If sFormName = "fintChanges" Then
      DoCmd.OpenForm "fintVolumeDashboard"
   End If
End Sub


===============
basUtility
===============
Option Explicit
Option Compare Database

'// Variables //
Global iNbrDecimals As Integer
'// Constants //
Global Const sfrmChangesControlTablePrefix As String = "tsysX"
Global Const sLabelFontName As String = "Calibri"
Global Const iLabelFontSize As String = 11            '// IPG MPT60 likes it at 11; 10 is default
Global Const sBasePath As String = "C:\OPTMODELS\MP67\"
Global Const sUOM As String = "Tons"

Global sCCVolume As String    'Change Candidates Volume Field
Global sFullQuickImport As String


Public Function WorkingMonthChange(sFormName As String, scmbboxdate As String, change As String)
   DoCmd.SetWarnings False
   DoCmd.RunSQL "UPDATE tbl001workingmonth SET tbl001workingmonth.workingmonth = """ & scmbboxdate & """"
   fSetSummaryBoxes sFormName, "Update"
   DoCmd.SetWarnings True
   
   Forms(sFormName).ForecastChangeArea.SourceObject = ""
   Forms(sFormName).VolumeChangeArea.SourceObject = ""
   Forms(sFormName).MoveArea.SourceObject = ""
   Forms(sFormName).DistMoveArea.SourceObject = ""
   Forms(sFormName).ViewOnlyArea.SourceObject = ""
   
   If change <> "FC" Then
      fSetComboBoxes sFormName, "FC", "FullRefresh"
   End If
   If change <> "VC" Then
      fSetComboBoxes sFormName, "VC", "FullRefresh"
   End If
   If change <> "MM" Then
      fSetComboBoxes sFormName, "MM", "FullRefresh"
   End If
   If change <> "VO" Then
      fSetComboBoxes sFormName, "VO", "FullRefresh"
   End If
   If change <> "DM" Then
      fSetComboBoxes sFormName, "DM", "FullRefresh"
   End If
   
   Forms(sFormName).Refresh
   
   On Error Resume Next
   Forms("MainPage")("WorkingMonthCombo").value = scmbboxdate
   
   MsgBox "Working Month Changed"

End Function

Public Sub CleanupForFC(sFormName As String)
'DoCmd.OpenQuery "qintUpdCleanUpForFC"
End Sub

Public Sub CleanupForVC(sFormName As String)
'DoCmd.OpenQuery "qintUpdCleanUpForVC"
End Sub

Public Sub CleanupForMM(sFormName As String)
   DoCmd.OpenQuery "qupdCleanUpMainCurrMill"
   ''?? IF THE CURRENT DISTPOINT IS ANY MILL
   ''   WRITE CURRENT MILL FIELD INTO CURRENT DIST POINT FIELD
   ''   E.G.--MAKE SURE CURR MILL AND CURR DIST PT ARE NOT BOTH MILLS AND UNMATCHED
   DoCmd.OpenQuery "qupdCleanUpMainCurrMillIfDistPtIsMill"
   'DoCmd.OpenQuery "q210CleanUp011AfterMM2"
End Sub

Public Sub CleanupForDM(sFormName As String, bReMapToggle As Boolean)
If sFormName = "fintChanges" Then
    If bReMapToggle = False Then
        ''?? IF THIS WAS REMAPPED TO A DIST POINT THAT *IS* A MILL, UPDATE
        ''   CURRENT MILL FIELD AND CURRENT MACHINE TO DIST POINT CODE AND
        ''   A MACHINE AT THAT MILL
        ''DoCmd.OpenQuery "q210CleanUp020AfterDM"
    Else
        DoCmd.OpenQuery "qupdCleanUpMillMachineAfterDM"
    End If
    ''DoCmd.OpenQuery "q210CleanUp025AfterDMSetK5"
    ''DoCmd.OpenQuery "q210CleanUp026AfterDMSetK6"
    ''DoCmd.OpenQuery "q210CleanUp027AfterDMSetK7"
End If
End Sub

Public Sub InitializeFormPages(sFormName As String)

Dim sSQL As String
iNbrDecimals = 0

   If sFormName = "frmChanges" Then
      iNbrDecimals = 2
      
      Forms(sFormName)("Summary1").Caption = "Summary 1"
      Forms(sFormName)("Summary2").Caption = "Summary 2"
      Forms(sFormName)("Summary3").Caption = "Summary 3"
      Forms(sFormName)("VO").Caption = "ViewOnly"
      Forms(sFormName)("DM").Caption = "DistPtMove"
      Forms(sFormName)("MM").Caption = "MachMove"
      Forms(sFormName)("FC").Caption = "FcastChg"
      Forms(sFormName)("VC").Caption = "VolumeChg"
       
      Forms(sFormName)("Summary1AreaA").SourceObject = "Query.qrepPMBalance" '"query.q910Report015PMBalanceAllMonths"
      Forms(sFormName)("Summary1AreaB").SourceObject = "Query.qrepInvBalance"
      Forms(sFormName)("Summary2AreaA").SourceObject = "Query.qrepMillByDistPoint"
      Forms(sFormName)("Summary2AreaB").SourceObject = "Query.qrepMillByProductProdLine"
      Forms(sFormName)("Summary3AreaA").SourceObject = "Query.qrepMillByCountry"
   
      Forms(sFormName)("Summary1").Visible = True
      Forms(sFormName)("Summary2").Visible = True
      Forms(sFormName)("Summary3").Visible = True
       
      ''sSQL = "Select *  into tblfintlkpDistPointCapability  from q810Lkp012DistPointCapability;"
      ''DoCmd.RunSQL sSQL
   End If

End Sub

Public Sub OpenDashboard(sFormName As String)
   If sFormName = "fintChanges" Then
      DoCmd.OpenForm "fintVolumeDashboard"
   End If
End Sub


===============
Form_frmChanges
===============
Option Explicit
Option Compare Database


Private Sub cmdChangeTableDM_Click()        'was command269
    DoCmd.SetWarnings False
    Me.SubtotalDM.value = "Click Here For Move Subtotal"
    Me.DistMoveArea.SourceObject = ""
    fChangeTablePrep Me.Name, "DM"
    Me.DistMoveArea.SourceObject = "Table." & sfrmChangesControlTablePrefix & "changecandidatesDM"
    Me.Refresh
    HideChangeColumns Me.Name, "DM", "DistMoveArea"
    DoCmd.SetWarnings True
End Sub


Private Sub cmdChangeTableFC_Click()        'was command125
    DoCmd.SetWarnings False
    Me.ForecastChangeArea.SourceObject = ""
    fChangeTablePrep Me.Name, "FC"
    Me.ForecastChangeArea.SourceObject = "Table." & sfrmChangesControlTablePrefix & "changecandidatesFC"
    Me.Refresh
    HideChangeColumns Me.Name, "FC", "ForecastChangeArea"
    DoCmd.SetWarnings True
End Sub


Private Sub cmdChangeTableMM_Click()        'was command95
    DoCmd.SetWarnings False
    Me.SubtotalMM.value = "Click Here For Move Subtotal"
    Me.WtdAvgWidthMM.value = "Click for WtdAvg Width of the Move"
    Me.MoveArea.SourceObject = ""
    fChangeTablePrep Me.Name, "MM"
    Me.MoveArea.SourceObject = "Table." & sfrmChangesControlTablePrefix & "changecandidatesMM"
    Me.Refresh
    HideChangeColumns Me.Name, "MM", "MoveArea"
    DoCmd.SetWarnings True
End Sub


Private Sub cmdChangeTableVC_Click()        'was command39
    DoCmd.SetWarnings False
    Me.VolumeChangeArea.SourceObject = ""
    fChangeTablePrep Me.Name, "VC"
    Me.VolumeChangeArea.SourceObject = "Table." & sfrmChangesControlTablePrefix & "changecandidatesVC"
    Me.Refresh
    HideChangeColumns Me.Name, "VC", "VolumeChangeArea"
    DoCmd.SetWarnings True
End Sub


Private Sub cmdChangeTableVO_Click()        'was command199
    Dim RS As ADODB.Recordset
    On Error GoTo cmdChangeTableVO_Click_Err
    DoCmd.SetWarnings False
    Me.WtdAvgWidth.value = "WtdAvg Width"
    Me.ViewOnlyArea.SourceObject = ""
    fChangeTablePrep Me.Name, "VO"
    Me.ViewOnlyArea.SourceObject = "Table." & sfrmChangesControlTablePrefix & "changecandidatesVO"
    Me.Refresh
    
    sCCVolume = "Plan"   'for MP67 only
    
    'set WtdAvgWidth Textbox
    Set RS = Application.CurrentProject.Connection.Execute("SELECT Round(Sum(CDbl([Width])*[" & sCCVolume & "])/Sum([" & sCCVolume & "]),2) AS WtdAvgWidth FROM " & _
                                                               sfrmChangesControlTablePrefix & "changecandidatesVO;")
    If Not RS.EOF Then
       Me.WtdAvgWidth.value = RS.Fields("WtdAvgWidth").value
    Else
       Me.WtdAvgWidth.value = "N/A"
    End If
    
    HideChangeColumns Me.Name, "VO", "ViewOnlyArea"
    DoCmd.SetWarnings True

cmdChangeTableVO_Click_Err:
   If Err Then
      If Err.Number = -2146697211 Then
'         MsgBox "The web server is not responding to a request for data." & vbNewLine & vbNewLine & "Please try the operation again or report the issue with ""Internet Services"" if the problem persists.", vbExclamation, Application.ActiveWorkbook.Name
      ElseIf Err.Number = -2147217904 Then    '
          Me.WtdAvgWidth.value = "N/A"
      Else
         'Resume
'         MsgBox "An error has occured while accessing data via Web Service." & vbNewLine & vbNewLine & "Error description:" & vbNewLine & Err.Number & " - " & Err.Description, vbExclamation, Application.ActiveWorkbook.Name
      End If
   End If
    
End Sub


Private Sub cmdCloseForm_Click()        'was command21
    DoCmd.Close acForm, [formname]
End Sub


Private Sub cmdExecuteDM_Click()        'was command270
    DoCmd.SetWarnings False
    fUpdateChanges Me.Name, "DM"
    CleanupForDM Me.Name, Me.chkReMapToggle
    DoCmd.SetWarnings True
    cmdChangeTableDM_Click
    If Me.VolumeChangeArea.SourceObject = "" Then
    Else
        cmdChangeTableVC_Click
    End If
    fSetSummaryBoxes Me.Name, "Update"
End Sub


Private Sub cmdExecuteFC_Click()        'was command126
    DoCmd.SetWarnings False
    Me.ForecastChangeArea.SourceObject = ""
    fUpdateChanges Me.Name, "FC"
    CleanupForFC Me.Name
    DoCmd.SetWarnings True
    cmdChangeTableFC_Click   'was Command125_Click
    If Me.VolumeChangeArea.SourceObject = "" Then
    Else
    cmdChangeTableVC_Click
    End If
    fSetSummaryBoxes Me.Name, "Update"
End Sub


Private Sub cmdExecuteMM_Click()        'was command96
    DoCmd.SetWarnings False
    fUpdateChanges Me.Name, "MM"
    CleanupForMM Me.Name
    DoCmd.SetWarnings True
    cmdChangeTableMM_Click              'fmrly call to command95
    If Me.VolumeChangeArea.SourceObject = "" Then
    Else
        cmdChangeTableVC_Click
    End If
    fSetSummaryBoxes Me.Name, "Update"
End Sub


Private Sub cmdExecuteVC_Click()        'was command40
    DoCmd.SetWarnings False
    fUpdateChanges Me.Name, "VC"
    CleanupForVC Me.Name
    DoCmd.SetWarnings True
    cmdChangeTableVC_Click
    fSetSummaryBoxes Me.Name, "Update"
End Sub


Private Sub cmdOpenDashboard_Click()        'was command189
    OpenDashboard Me.Name
End Sub


Private Sub cmdSaveSettingAsDefault_Click()        'was command286    (only on Dist Point Move tab (DM) on 6/19/2012)
    Dim ReMapToggleValue As Integer
    Dim CurrentForm As String
    DoCmd.Echo False
    ReMapToggleValue = Me.chkReMapToggle
    CurrentForm = Me.Name
    DoCmd.OpenForm CurrentForm, acDesign
    Forms(CurrentForm).chkReMapToggle.DefaultValue = ReMapToggleValue
    DoCmd.Save acForm, CurrentForm
    DoCmd.OpenForm CurrentForm, acNormal
    Forms(CurrentForm).chkReMapToggle.value = ReMapToggleValue
    DoCmd.Echo True
End Sub


Private Sub DistMoveArea_Exit(Cancel As Integer)
Dim RS As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim sSQL As String

sSQL = "SELECT Name FROM MSysObjects WHERE Name = """ & sfrmChangesControlTablePrefix & "changecandidatesDM""" & ";"
Set rs2 = Application.CurrentProject.Connection.Execute(sSQL)

If Not rs2.EOF Then
    Set RS = Application.CurrentProject.Connection.Execute("select sum(" & sUOM & ") as MoveSubtotal from " & sfrmChangesControlTablePrefix & "changecandidatesDM where Move = -1")
    Me.SubtotalDM.value = RS.Fields("MoveSubtotal").value
End If


End Sub

Private Sub Form_Close()
    DoCmd.SetWarnings False
    On Error Resume Next
    fDeleteLookupTables Me.Name
    DoCmd.OpenForm Me, acDesign
    Me.ForecastChangeArea.SourceObject = ""
    Me.VolumeChangeArea.SourceObject = ""
    Me.MoveArea.SourceObject = ""
    Me.DistMoveArea.SourceObject = ""
    Me.ViewOnlyArea.SourceObject = ""
    fSetComboBoxes Me.Name, "FC", "Clear"
    fSetComboBoxes Me.Name, "VC", "Clear"
    fSetComboBoxes Me.Name, "MM", "Clear"
    fSetComboBoxes Me.Name, "DM", "Clear"
    fSetComboBoxes Me.Name, "VO", "Clear"
    DoCmd.DeleteObject acTable, sfrmChangesControlTablePrefix & "changecandidatesMM"
    DoCmd.DeleteObject acTable, sfrmChangesControlTablePrefix & "changecandidatesDM"
    DoCmd.DeleteObject acTable, sfrmChangesControlTablePrefix & "changecandidatesVC"
    DoCmd.DeleteObject acTable, sfrmChangesControlTablePrefix & "changecandidatesFC"
    DoCmd.DeleteObject acTable, sfrmChangesControlTablePrefix & "changecandidatesVO"
    DoCmd.SetWarnings True
End Sub


Private Sub Form_Open(Cancel As Integer)
    DoCmd.Maximize
    DoCmd.SetWarnings False
    Me.ForecastChangeArea.SourceObject = ""
    Me.VolumeChangeArea.SourceObject = ""
    Me.MoveArea.SourceObject = ""
    Me.DistMoveArea.SourceObject = ""
    Me.ViewOnlyArea.SourceObject = ""
    fCreateLookupTables Me.Name
    fSetComboBoxes Me.Name, "FC", "FullRefresh"
    fSetComboBoxes Me.Name, "VC", "FullRefresh"
    fSetComboBoxes Me.Name, "MM", "FullRefresh"
    fSetComboBoxes Me.Name, "VO", "FullRefresh"
    fSetComboBoxes Me.Name, "DM", "FullRefresh"
    
    fSetSummaryBoxes Me.Name, "FullRefresh"
    
    InitializeFormPages Me.Name
    
    Me.Refresh
    
    Me.TabCtl0.Pages(0).SetFocus
    
    DoCmd.SetWarnings True
End Sub



Private Sub MoveArea_Exit(Cancel As Integer)
Dim RS As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim sSQL As String
On Error GoTo MAE_ERR

sSQL = "SELECT Name FROM MSysObjects WHERE Name = """ & sfrmChangesControlTablePrefix & "changecandidatesMM""" & ";"
Set rs2 = Application.CurrentProject.Connection.Execute(sSQL)

If Not rs2.EOF Then
    Set RS = Application.CurrentProject.Connection.Execute("select sum(" & sUOM & ") as MoveSubtotal from " & sfrmChangesControlTablePrefix & "changecandidatesMM where Move = -1")
    Me.SubtotalMM.value = RS.Fields("MoveSubtotal").value
    Set RS = Application.CurrentProject.Connection.Execute("SELECT Round(Sum(CDbl([Width])*[" & sCCVolume & "])/Sum([" & sCCVolume & "]),2) AS WtdAvgWidth FROM " & _
                                                               sfrmChangesControlTablePrefix & "changecandidatesMM WHERE Move = -1;")
    If Not RS.EOF Then
       Me.WtdAvgWidthMM.value = RS.Fields("WtdAvgWidth").value
    Else
       Me.WtdAvgWidthMM.value = RS.Fields("WtdAvgWidth").value = "N/A"
    End If
End If

MAE_ERR:
   If Err Then
      If Err.Number = -2146697211 Then
'         MsgBox "The web server is not responding to a request for data." & vbNewLine & vbNewLine & "Please try the operation again or report the issue with ""Internet Services"" if the problem persists.", vbExclamation, Application.ActiveWorkbook.Name
      ElseIf Err.Number = -2147217904 Then    '
          Me.WtdAvgWidthMM.value = "N/A"
      Else
         'Resume
'         MsgBox "An error has occured while accessing data via Web Service." & vbNewLine & vbNewLine & "Error description:" & vbNewLine & Err.Number & " - " & Err.Description, vbExclamation, Application.ActiveWorkbook.Name
      End If
   End If

End Sub





===============
Form_frmSettingsExcelPivots
===============
Option Compare Database
Option Explicit

'RKP/04-26-13/V01
'********** START OPTIONS            **********
'Option Compare Database
'********** END   OPTIONS            **********
'********** START DLL DECLARATIONS   **********
'********** END   DLL DECLARATIONS   **********
'********** START PUBLIC CONSTANTS   **********
'********** END   PUBLIC CONSTANTS   **********
'********** START PUBLIC VARIABLES   **********
'********** END   PUBLIC VARIABLES   **********
'********** START PRIVATE CONSTANTS  **********
'********** END   PRIVATE CONSTANTS  **********
'********** START PRIVATE VARIABLES  **********
Private mlLastErr       As Long
Private msLastErr       As String
Private msErrSource     As String
Private msErrDesc       As String
'********** END   PRIVATE VARIABLES  **********
'********** START USER DEFINED TYPES **********
'********** END   USER DEFINED TYPES **********

'Public Function Function1()
''**********************************************
''Author  :  RKP
''Date/Ver:  04-26-13/V01
''Input   :
''Output  :
''Comments:
''**********************************************
'    On Error GoTo Err_Handler
'
'
'
'Err_Handler:
'    mlLastErr = Err.Number
'    msLastErr = Err.Description
'    'Function1 = mlLastErr
'    If Err Then
'        If Err.Number = 49 Then 'Bad DLL calling convention
'            mlLastErr = 0
'            msLastErr = ""
'            Resume Next
'        Else
'            'ProcessMsg Err.Number, Err.Description, "", ""
'            MsgBox Err.Number & " - " & Err.Description
'        End If
'    End If
'
'    Exit Function
'    Resume
'End Function

Private Sub btnGeneratePivotTable_Click()
'**********************************************
'Author  :  RKP
'Date/Ver:  04-26-13/V01
'Input   :
'Output  :
'Comments:
'**********************************************
    On Error GoTo Err_Handler
    
    Dim listOfRowIDs()  As String
    Dim ctr             As Integer
    Dim xlWB            As Workbook

    If VBA.Trim(GetControlValue(Me.txtRowIDs)) = "" Then
        GeneratePivotTable_Generic , , , VBA.CInt(GetControlValue(Me.RowID))
    Else
        If VBA.InStr(1, GetControlValue(Me.txtRowIDs), ",", vbTextCompare) > 0 Then
            listOfRowIDs = VBA.Split(GetControlValue(Me.txtRowIDs), ",", , vbTextCompare)
            For ctr = 0 To UBound(listOfRowIDs)
                DoEvents
                
                Set xlWB = GeneratePivotTable_Generic(, , , VBA.CInt(listOfRowIDs(ctr)), xlWB)
            Next
        Else
            GeneratePivotTable_Generic , , , GetControlValue(Me.txtRowIDs)
        End If
        
    End If

Err_Handler:
    mlLastErr = Err.Number
    msLastErr = Err.Description
    'Function1 = mlLastErr
    If Err Then
        If Err.Number = 49 Then 'Bad DLL calling convention
            mlLastErr = 0
            msLastErr = ""
            Resume Next
        Else
            'ProcessMsg Err.Number, Err.Description, "", ""
            MsgBox Err.Number & " - " & Err.Description
        End If
    End If

    Exit Sub
    Resume
End Sub

Public Function GetControlValue(ByRef cntrl As Object) As String
'**********************************************
'Author  :  RKP
'Date/Ver:  04-26-13/V01
'Input   :
'Output  :
'Comments:
'**********************************************
    On Error GoTo Err_Handler

    cntrl.SetFocus
    GetControlValue = cntrl & ""

Err_Handler:
    mlLastErr = Err.Number
    msLastErr = Err.Description
    'Function1 = mlLastErr
    If Err Then
        If Err.Number = 49 Then 'Bad DLL calling convention
            mlLastErr = 0
            msLastErr = ""
            Resume Next
        Else
            'ProcessMsg Err.Number, Err.Description, "", ""
            MsgBox Err.Number & " - " & Err.Description
        End If
    End If

    Exit Function
    Resume
End Function
===============
modExcelPivot
===============
Option Compare Database
Option Explicit

'RKP/04-26-13/V01
'********** START OPTIONS            **********
'Option Compare Database
'********** END   OPTIONS            **********
'********** START DLL DECLARATIONS   **********
'********** END   DLL DECLARATIONS   **********
'********** START PUBLIC CONSTANTS   **********
'********** END   PUBLIC CONSTANTS   **********
'********** START PUBLIC VARIABLES   **********
'********** END   PUBLIC VARIABLES   **********
'********** START PRIVATE CONSTANTS  **********
'********** END   PRIVATE CONSTANTS  **********
'********** START PRIVATE VARIABLES  **********
Private mlLastErr       As Long
Private msLastErr       As String
Private msErrSource     As String
Private msErrDesc       As String
'********** END   PRIVATE VARIABLES  **********
'********** START USER DEFINED TYPES **********
'********** END   USER DEFINED TYPES **********

'Public Function Function1()
''**********************************************
''Author  :  RKP
''Date/Ver:  04-26-13/V01
''Input   :
''Output  :
''Comments:
''**********************************************
'    On Error GoTo Err_Handler
'
'
'
'Err_Handler:
'    mlLastErr = Err.Number
'    msLastErr = Err.Description
'    'Function1 = mlLastErr
'    If Err Then
'        If Err.Number = 49 Then 'Bad DLL calling convention
'            mlLastErr = 0
'            msLastErr = ""
'            Resume Next
'        Else
'            'ProcessMsg Err.Number, Err.Description, "", ""
'            MsgBox Err.Number & " - " & Err.Description
'        End If
'    End If
'
'    Exit Function
'    Resume
'End Function

Private Sub btnGeneratePivotTable_Click()
'**********************************************
'Author  :  RKP
'Date/Ver:  04-26-13/V01
'Input   :
'Output  :
'Comments:
'**********************************************
    On Error GoTo Err_Handler
    
    Dim listOfRowIDs()  As String
    Dim ctr             As Integer
    Dim xlWB            As Workbook

    If VBA.Trim(GetControlValue(Me.txtRowIDs)) = "" Then
        GeneratePivotTable_Generic , , , VBA.CInt(GetControlValue(Me.RowID))
    Else
        If VBA.InStr(1, GetControlValue(Me.txtRowIDs), ",", vbTextCompare) > 0 Then
            listOfRowIDs = VBA.Split(GetControlValue(Me.txtRowIDs), ",", , vbTextCompare)
            For ctr = 0 To UBound(listOfRowIDs)
                DoEvents
                
                Set xlWB = GeneratePivotTable_Generic(, , , VBA.CInt(listOfRowIDs(ctr)), xlWB)
            Next
        Else
            GeneratePivotTable_Generic , , , GetControlValue(Me.txtRowIDs)
        End If
        
    End If

Err_Handler:
    mlLastErr = Err.Number
    msLastErr = Err.Description
    'Function1 = mlLastErr
    If Err Then
        If Err.Number = 49 Then 'Bad DLL calling convention
            mlLastErr = 0
            msLastErr = ""
            Resume Next
        Else
            'ProcessMsg Err.Number, Err.Description, "", ""
            MsgBox Err.Number & " - " & Err.Description
        End If
    End If

    Exit Sub
    Resume
End Sub

Public Function GetControlValue(ByRef cntrl As Object) As String
'**********************************************
'Author  :  RKP
'Date/Ver:  04-26-13/V01
'Input   :
'Output  :
'Comments:
'**********************************************
    On Error GoTo Err_Handler

    cntrl.SetFocus
    GetControlValue = cntrl & ""

Err_Handler:
    mlLastErr = Err.Number
    msLastErr = Err.Description
    'Function1 = mlLastErr
    If Err Then
        If Err.Number = 49 Then 'Bad DLL calling convention
            mlLastErr = 0
            msLastErr = ""
            Resume Next
        Else
            'ProcessMsg Err.Number, Err.Description, "", ""
            MsgBox Err.Number & " - " & Err.Description
        End If
    End If

    Exit Function
    Resume
End Function
===============
Form_sfrmMachineLevel
===============
Option Compare Database
Option Explicit

Private Sub Detail_Paint()
   Dim dDayMax As Double
   Dim lMaxBarWidth As Double
   Dim lRedLength As Long
   Dim lBlueLength As Double
   'dDayMax = 360
   'lMaxBarWidth = 7700
    
   'lRedLength = Me.Avail / dDayMax * 5.33 * 1440 + 720
   'lBlueLength = Round(Me.PlanDays / dDayMax * 5.33 * 1440, 0)
   
   'Me.lineMillCapacity.Left = lRedLength
   'Me.txtMillBar.Width = lBlueLength
   
'    ''Me.lineMillPlanTons.Left = lBlueLength + 720
'    ''Me!line_mill_upper.Width = lBlueLength
'    ''Me!line_mill_lower.Width = lBlueLength
'
'    'Me!status_box.Value = "OK"
'    'Me!status_fill.Visible = True
'    'Me!status_fill.Width = 1 * 1440
End Sub
===============
Form_sfrmPMBalance
===============
Option Compare Database
Option Explicit

Private Sub Detail_Paint()
   Dim dDayMax As Double
   Dim lMaxBarWidth As Double
   Dim lRedLength As Long
   Dim lBlueLength As Double
   dDayMax = 360
   lMaxBarWidth = 8780
   
   lRedLength = Me.Avail / dDayMax * 5.33 * 1440 + 720
   lBlueLength = Round(Me.PlanDays / dDayMax * 5.33 * 1440, 0)
   
   If lBlueLength > lMaxBarWidth Then lBlueLength = lMaxBarWidth - 144
   
   'Me.lineMillCapacity.Left = lRedLength
   'Me.txtMillBar.Width = lBlueLength
End Sub
===============
Report_srepPMBalance
===============
Option Compare Database
Option Explicit

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
   Dim dDayMax      As Double
   Dim lMaxBarWidth As Double
   Dim lRedLength   As Long
   Dim lBlueLength  As Double
   Dim sLabelName   As String
   Dim i            As Integer
   
   dDayMax = DLookup("MaxDays", "qdatHorizonDays") 'usually 30 OR 360
   
   For i = 1 To 6
      sLabelName = "lblVal" & i
      Reports(Me.Name)(sLabelName).Caption = CStr(CInt(Round(i * dDayMax / 6, 2)))
   Next i
   
   lMaxBarWidth = 7700
    
   lRedLength = uMIN(VBA.Round(Me.Avail / dDayMax * 5.33 * 1440 + 720, 0), lMaxBarWidth)
   lBlueLength = uMIN(VBA.Round(Me.PlanDays / dDayMax * 5.33 * 1440, 0), lMaxBarWidth)
   
   Me.lineMillCapacity.Left = lRedLength
   Me.txtMillBar.Width = lBlueLength
   
    'Me.lineMillPlanTons.Left = lBlueLength + 720
'    ''Me!line_mill_upper.Width = lBlueLength
'    ''Me!line_mill_lower.Width = lBlueLength
'
'    'Me!status_box.Value = "OK"
'    'Me!status_fill.Visible = True
'    'Me!status_fill.Width = 1 * 1440
End Sub
