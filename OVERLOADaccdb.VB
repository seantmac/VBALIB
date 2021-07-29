===============
basApp
===============
Option Compare Database
Option Explicit

'********** START PRIVATE VARIABLES  **********
Private mlLastErr       As Long
Private msLastErr       As String
Private msErrSource     As String
Private msErrDesc       As String


Public Function CreateTimePeriods()
   Dim i           As Integer
   Dim j           As Integer
   Dim msg         As String
   Dim RS          As DAO.Recordset
   Dim sSQL        As String
   Dim z  As Date
   Dim zz As Date
   Dim iDayofWeekPeriodBoundary As Integer  'e.g. 5 = THURSDAY
   Dim dBegDate As Date
   Dim dEndDate As Date
   CreateTimePeriods = False
   
   iDayofWeekPeriodBoundary = 3 '5 is THURSDAY 3 IS default
   iDayofWeekPeriodBoundary = VBA.CInt(basUtility.GetSetting("DayOfWeekPeriodEnds", "tsysSettings"))   'was hardwired to 5 for Thursday
   
   dBegDate = DateAdd("s", 1, 1 * (Date - DatePart("y", Date - 1)))
   i = 0
   j = 0
   
   DoCmd.SetWarnings False
   sSQL = "DELETE tlkTimePeriod.* FROM tlkTimePeriod "
   DoCmd.RunSQL (sSQL)
   DoCmd.SetWarnings True
   
   sSQL = "SELECT tlkTimePeriod.* FROM tlkTimePeriod "
   Set RS = CurrentDb.OpenRecordset(sSQL)
   'rs.MoveLast
   'rs.MoveFirst
   
   For i = 0 To VBA.CInt(basUtility.GetSetting("MaxTimeHorizonInDays", "tsysSettings"))   'was hardwired to 99
      
      If Weekday(Int(Now()) + i) = iDayofWeekPeriodBoundary Then 'it is THURSDAY
         dEndDate = DateAdd("s", -2, Int(Now()) + i)
         With RS
         .AddNew
         !TimeID = j
         !TimeActive = True
         !TimeName = "P" & j
         !TimeCode = "P" & j
         !Year = Year(dBegDate)
         !Month = Month(dBegDate)
         !Months = Round((dEndDate - dBegDate) / 30, 1)
         !Weeks = Round((dEndDate - dBegDate) / 7, 1)
         !Days = Round((dEndDate - dBegDate), 1)
         !StartingDate = dBegDate
         !EndingDate = dEndDate
         !BlockPointDefinition = VBA.CStr(basUtility.GetSetting("BlockPointDefinition", "tsysSettings"))    'was hardwired to MidPoint
                                           'NOW USUALLY SET TO STARTING TO USE RAVI LOGIC
         .Update
         End With
         dBegDate = DateAdd("s", 1, dEndDate)
         j = j + 1
      End If
   Next i
   
   CreateTimePeriods = True

End Function


Public Function ImportZPPY2Files() As Boolean
Dim i As Integer
Dim oWSH As Object

ImportZPPY2Files = False

'Clean and UnWrap
For i = 1 To 99
    If FileExists("C:\TEMP\ZPPY\ZPPY_LVL2_" & Format(i, "00") & ".txt") Then
        Call CleanPipeDelimTextFile("C:\TEMP\ZPPY\ZPPY_LVL2_" & Format(i, "00") & ".txt")
        Call Unwrap2Lines("C:\TEMP\ZPPY\ZPPY_LVL2_" & Format(i, "00") & "_CLN.txt")
    Else
        Exit For
    End If
Next i

'copy/glom the files all together
Set oWSH = CreateObject("WScript.Shell")
    oWSH.Run ("%comspec% /c copy C:\TEMP\ZPPY\*_Revised.txt C:\TEMP\ZPPY\ZPPY_LEVEL2.txt")

Call WaitSecs(3)

'IMPORT TO TABLE IN ONE STEP
i = TXTImportFromText2("Zppy2Clean", "C:\TEMP\ZPPY\ZPPY_LEVEL2.txt", "|")

'remove text header in contents of table
DoCmd.SetWarnings False
DoCmd.RunSQL "DELETE *  FROM Zppy2Clean  WHERE (((Zppy2Clean.[Req Qty])='Req Qty'))"
DoCmd.SetWarnings True

ImportZPPY2Files = True
End Function


Public Function Tester() As Boolean
   DoCmd.RunSQL ("DELETE tblRESTRbyGBWwidth.* FROM tblRESTRbyGBWwidth; ")         'Restrictions
   DoCmd.RunSQL ("INSERT INTO tblRESTRbyGBWwidth  ( MACH, GRADE, BW, WIDTH, ShipTo ) " & _
                 "SELECT tdatRestrictions.MachineCode, tdatRestrictions.GRADE, tdatRestrictions.BW, tdatRestrictions.WID, tdatRestrictions.ShipTo  " & _
                 "FROM tdatRestrictions IN '\\s02afs01.na.ipaper.com\lenox-275\CST\Supply Planning\Trim Help RHC\RollHelpCentral.mdb'; ")
                 
   DoCmd.RunSQL ("DELETE tblRESTRbyGBW.* FROM tblRESTRbyGBW; ")         'RestrictionsZeroWidth
   DoCmd.RunSQL ("INSERT INTO tblRESTRbyGBW ( Mach, Grade, BW, ShipTo ) " & _
                 "SELECT tdatRESTRZeroWidth.[Restricted PM], tdatRESTRZeroWidth.Grade, tdatRESTRZeroWidth.BW, tdatRESTRZeroWidth.ShipTo " & _
                 "FROM tdatRESTRZeroWidth IN '\\s02afs01.na.ipaper.com\lenox-275\CST\Supply Planning\Trim Help RHC\RollHelpCentral.mdb' ")
End Function


Sub Unwrap2Lines(sPathFileName As String)
'//===============================================================//
'//===============================================================//

    Dim FSO As FileSystemObject   'needs ms scripting runtime reference
    Dim FSOFile As TextStream, FSOFileRevised As TextStream
    Dim filePath As String, FilePathRevised As String
    Dim sLine As String
    Dim iLineCounter As Integer
         
    filePath = sPathFileName
    FilePathRevised = Left(filePath, Len(filePath) - 4) & "_Revised" & Right(filePath, 4)       ' adds "_Revised" to your file name
     
    Set FSO = New FileSystemObject
    
    If FSO.FileExists(filePath) Then
         ' opens the file for reading
        Set FSOFile = FSO.OpenTextFile(filePath, 1, False)
         ' opens "revised" file in write mode
        Set FSOFileRevised = FSO.OpenTextFile(FilePathRevised, 2, True)
        
        iLineCounter = 1
        
        Do While Not FSOFile.AtEndOfStream
            sLine = FSOFile.ReadLine
            If iLineCounter Mod 2 = 1 Then  'ODD
               FSOFileRevised.Write Trim(sLine) & Space(2) & "  |  "    'writes but does not Carriage Return Line Feed
            Else 'EVEN
               FSOFileRevised.Write Trim(sLine) & vbNewLine
            End If
            iLineCounter = iLineCounter + 1
        Loop
        
        FSOFile.Close
        FSOFileRevised.Close
    Else
        MsgBox (filePath & " does not exist")
    End If
    
End Sub


'you have all the runqrybyprefix stuff


===============
Form_frmMain
===============
Option Compare Database
Option Explicit


Private Sub cmdAdjZPPYLevel1_Click()

DoCmd.OpenQuery "drsqZPPYLevel1Adj"

End Sub

Private Sub cmdLevel12Check_Click()

DoCmd.OpenQuery "qrepZPPYLevel2ScopingFromStagedCHECKsummary"

End Sub

Private Sub cmdLiveZPPYCheck_Click()

DoCmd.OpenQuery "drsqLiveZPPYSummary"

End Sub

Private Sub cmdOpenControls_Click()
DoCmd.OpenForm "drsfrmControls"
End Sub

Private Sub cmdOptResults_Click()
DoCmd.OpenQuery "drsqmodresChangeSummary"
End Sub

Private Sub cmdRedoMoveOptions_Click()
   Dim sSQL As String
   
   DoCmd.SetWarnings False
   sSQL = "DELETE tmtx00MOVEOPTIONS.* FROM tmtx00MOVEOPTIONS "
   DoCmd.RunSQL (sSQL)
   
   DoCmd.OpenQuery "qappMOVEOPTIONS"
   
   DoCmd.SetWarnings True
End Sub


Private Sub cmdRedoSnapshots_Click()
   Dim sSQL As String
   Dim i    As Integer
   
   UpdateStatus ("Creating Time Periods Table")
   i = CreateTimePeriods()
   
   DoCmd.SetWarnings False
   sSQL = "DELETE tdatZPPYLevel1.* FROM tdatZPPYLevel1 "
   DoCmd.RunSQL (sSQL)
   sSQL = "DELETE tdatZPPYLevel2.* FROM tdatZPPYLevel2 "
   DoCmd.RunSQL (sSQL)
   
   DoCmd.OpenQuery "qappLocalZPPYLevel1"
   DoCmd.OpenQuery "qappLocalZPPYLevel2"
   
   DoCmd.SetWarnings True
End Sub


Private Sub cmdRedoWorkingTables_Click()
   Dim sSQL As String
   
   DoCmd.SetWarnings False
   sSQL = "DELETE tblWorkingCapacity.* FROM tblWorkingCapacity "
   DoCmd.RunSQL (sSQL)
   sSQL = "DELETE tblWorkingDemand.* FROM tblWorkingDemand "
   DoCmd.RunSQL (sSQL)
   
   DoCmd.OpenQuery "qapptblWorkingCapacity"
   DoCmd.OpenQuery "qapptblWorkingDemand"
   
   DoCmd.SetWarnings True
End Sub


Private Sub cmdRunOptModel_Click()
   
   '#OPT RUN
   Dim sMsg As String
   Dim RS   As DAO.Recordset
   Dim scmd As String
   Dim sCOPTPathAndFile As String
   Dim sSolver As String
   Dim x As String
   Dim db As DAO.Database
''   Dim rs As ADODB.Recordset
   Dim dTons As Double
   Dim sHostName As String
   Dim bShiptoSoleSource As Boolean
   Dim sSQL As String
      
   sHostName = Environ$("computername")
   Debug.Print sHostName
   
   sMsg = ""
   sMsg = sMsg & "The Opt Model is self-contained in the OPT Database." & vbCrLf
   sMsg = sMsg & "Move Options are generated at runtime.    " & vbCrLf
   sMsg = sMsg & "   " & vbCrLf
   
   If MsgBox(sMsg, vbOKCancel, "Optimization Run") <> vbOK Then
      Exit Sub
   End If
   
   
   '** Now, Find C-opt and go run the opt model with C-OPT
   '**
   '**      "RUN /PRJ OVERLOAD /Solver CPLEX /WorkDir C:\OPTMODELS\OVERLOAD\Output /RunModelChecks /GenNativeMPS MPSFREE /GenSOLFixedFile /GenLPFile /GenRLPFile /GenSAVFile /NoUrlLog"
'''         "RUN /PRJ GCF8 /WORKDIR C:\OPTMODELS\GCF8\OUTPUT /Solver CoinMP /Sense MIN -1 /SetRelativeGapTolerance " & CIntlNumber("0.001", 3) /SetTimeLimit " & CIntlNumber("240.000", 3) & " 600 /NOPROMPT /NoUrlLog"
   
   
   sSolver = "CoinMP"
   x = basUtility.GetSetting("Solver", "tsysSettings")
   If x = "COIN" Or x = "CPLEX" Or x = "COINMP" Then
      sSolver = x
   End If
   sCOPTPathAndFile = "C:\Program Files\BMOS\C-OPT\C-OPTConsole.exe " & _
      "RUN /PRJ OVERLOAD /Solver " & sSolver & " /WorkDir C:\OPTMODELS\OVERLOAD\Output /RunModelChecks /GenSOLFixedFile /GenLPFile /GenRLPFile /GenSAVFile /NoUrlLog /NoPrompt "
       '/GenNativeMPS MPSFREE
   If Len(GetCOPTpath()) > 0 Then  'SET THE PATH ACCORDING TO THE USER'S C-OPT INSTALL DIRECTORY
      sCOPTPathAndFile = GetCOPTpath() & "C-OPTConsole.exe " & _
      "RUN /PRJ OVERLOAD /Solver " & sSolver & " /WorkDir C:\OPTMODELS\OVERLOAD\Output /RunModelChecks /GenSOLFixedFile /GenLPFile /GenRLPFile /GenSAVFile /NoUrlLog /NoPrompt "
   End If
   
   
   scmd = sCOPTPathAndFile
   Debug.Print scmd
   
   DoCmd.Hourglass (True)
   ExecCmd (scmd)
   DoCmd.Hourglass (False)
   
   
   DoCmd.Echo True
''   MsgBox "Optimization Complete."
   DoCmd.SetWarnings True
   
   
   '''''open results form -- modeler results OR user results OR both
   ''''DoCmd.OpenForm "frmnavOptResults"
   
    DoCmd.SetWarnings False
   
    sSQL = "Delete drstShedDemand.* From drstShedDemand"
    DoCmd.RunSQL (sSQL)
    DoCmd.OpenQuery "drsqappShedDemand"
    
    DoCmd.OpenQuery "drsqdelrepMovesFromModel"
    DoCmd.OpenQuery "drsqdelrepMovesSmallVolmoves"
    DoCmd.OpenQuery "drsqdelrepMovesSplits"
    
    DoCmd.OpenQuery "drsqapprepMovesFromModel"
    
    DoCmd.OpenQuery "drsqapprepMovesSplits"
    DoCmd.OpenQuery "drsqupdrepMovesMarkSplits"
    DoCmd.OpenQuery "drsqdelrepMovesRemoveMarked"
    
    DoCmd.OpenQuery "drsqapprepMovesSmallVolmoves"
    DoCmd.OpenQuery "drsqupdrepMovesMarkSmallVolume"
    DoCmd.OpenQuery "drsqdelrepMovesRemoveMarked"
   
    DoCmd.SetWarnings True
   
   DoCmd.OpenQuery "drsqmodresChangeSummary"

End Sub


Private Sub cmdRunPrep_Click()
   Dim sSQL As String
   
   DoCmd.SetWarnings False

   '// REFRESH PRIORITIES AND VALID LANES BY ORIG-DEST-GBW-MODES
   sSQL = "DELETE tblDEMANDPriority.* FROM tblDEMANDPriority "
   DoCmd.RunSQL (sSQL)
   sSQL = "DELETE tblLANESbyGBW.* FROM tblLANESbyGBW "
   DoCmd.RunSQL (sSQL)

   
   DoCmd.OpenQuery "qappDEMANDPriority"
   DoCmd.OpenQuery "qappLANESbyGBW"

'''Moved To Weekly Refresh
'''   sSQL = "DELETE drstTLanesAll.* FROM drstTLanesAll "
'''   DoCmd.RunSQL (sSQL)
'''   DoCmd.OpenQuery "drsqappTLanesAll"
   
   '// REDO P01 SYSTEM TEMPLATE / P01 = P0+P1 Latest  P1 Latest
   sSQL = "DELETE tblP01SysTemplateChanges.* FROM tblP01SysTemplateChanges "
   DoCmd.RunSQL (sSQL)
   
   DoCmd.OpenQuery "qappP01SysTemplateChanges"
   
   
   '// REDO WORKING TABLES FOR CAPACITY AND DEMAND
   sSQL = "DELETE tblWorkingCapacity.* FROM tblWorkingCapacity "
   DoCmd.RunSQL (sSQL)
   sSQL = "DELETE tblWorkingDemand.* FROM tblWorkingDemand "
   DoCmd.RunSQL (sSQL)
   
   DoCmd.OpenQuery "qapptblWorkingCapacity"
   DoCmd.OpenQuery "qapptblWorkingDemand"
   
   '// REDO MODEL TABLES FOR TBLCHANGE (BALDELTAS) AND TBLDEMAND
   sSQL = "DELETE tblCHANGE.* FROM tblCHANGE "
   DoCmd.RunSQL (sSQL)
   sSQL = "DELETE tblDEMAND.* FROM tblDEMAND "
   DoCmd.RunSQL (sSQL)
   
   DoCmd.OpenQuery "qappTblCHANGE"
   DoCmd.OpenQuery "qappTblDEMAND"
   
   'override all machine groups with "ANY" to allow moves outside machine groups if tsysSettingsSwitch is ON.
   DoCmd.OpenQuery "qupdCHANGEPMGtoANY"
   DoCmd.OpenQuery "qupdDEMANDPMGtoANY"
   
   
   '// REDO RESTRICTIONS FROM ROLL HELP CENTRAL BUT PUT IN tblRESTRbyGBW AND tblRESTRbyGBWwidth not the tdatRESTR
   DoCmd.RunSQL ("DELETE tblRESTRbyGBWwidth.* FROM tblRESTRbyGBWwidth; ")         'Restrictions
   DoCmd.RunSQL ("INSERT INTO tblRESTRbyGBWwidth  ( MACH, GRADE, BW, WIDTH, ShipTo ) " & _
                 "SELECT tdatRestrictions.MachineCode, tdatRestrictions.GRADE, tdatRestrictions.BW, tdatRestrictions.WID, tdatRestrictions.ShipTo  " & _
                 "FROM tdatRestrictions IN '\\s02afs01.na.ipaper.com\lenox-275\CST\Supply Planning\Trim Help RHC\RollHelpCentral.mdb'; ")
                 
   DoCmd.RunSQL ("DELETE tblRESTRbyGBW.* FROM tblRESTRbyGBW; ")         'RestrictionsZeroWidth
   DoCmd.RunSQL ("INSERT INTO tblRESTRbyGBW ( Mach, Grade, BW, ShipTo ) " & _
                 "SELECT tdatRESTRZeroWidth.[Restricted PM], tdatRESTRZeroWidth.Grade, tdatRESTRZeroWidth.BW, tdatRESTRZeroWidth.ShipTo " & _
                 "FROM tdatRESTRZeroWidth IN '\\s02afs01.na.ipaper.com\lenox-275\CST\Supply Planning\Trim Help RHC\RollHelpCentral.mdb' ")
   
   
   '// REDO FREIGHT TABLE from VALID Lanes By GBW (APPEND NEW AND UPDATE ALL CPTs
   DoCmd.RunSQL ("DELETE tblFREIGHT.* FROM tblFREIGHT; ")         'Restrictions
   DoCmd.OpenQuery "qappNewFreight"
   DoCmd.OpenQuery "qupdNewFreight"
   DoCmd.OpenQuery "drsqupdMissingFreightFromMPT75"  '''DRS
   
   
   '// PRE-MAKE THE MOVEOPTIONS tbl (model remakes on-the-fly every time)
''''   DoCmd.OpenQuery "qdelMOVEOPTIONS"
''''   DoCmd.OpenQuery "qappMOVEOPTIONS"


'''''''Table tblRESTRbyGBWwidth is in C:\OPTMODELS\OVERLOAD\OVERLOAD.accdb
'''''''    qlkxDemandSrcAllowedGBWwidth
'''''''Table tblRESTRbyGBW is in C:\OPTMODELS\OVERLOAD\OVERLOAD.accdb
'''''''    qlkxDemandSrcAllowedGBW
'''''''    qlkxDemandSrcAllowedGBWwidth
'''''''Table tdatRESTRZeroWidth is in C:\OPTMODELS\OVERLOAD\OVERLOAD.accdb
'''''''Table tdatRestrictions is in C:\OPTMODELS\OVERLOAD\OVERLOAD.accdb



   DoCmd.SetWarnings True
   MsgBox "Run Prep Complete."
   
   
End Sub

Private Sub cmdSetTheTimePeriods_Click()
   Dim i As Boolean
   ''If MsgBox("Disabled for testing at the moment.  Pls comment out this msgBox code if you really want to test it out. ") = vbOK Then
   ''Else
      i = CreateTimePeriods()
      MsgBox "Time Periods Updated"
   ''End If
End Sub


Private Sub cmdSettings_Click()
   DoCmd.OpenTable "tsysSettings"
End Sub


Private Sub cmdUploadLanes_Click()

    'DRS/11-13-20 - brought over from TPL to Overload
    Dim sapInstance
    Dim session
    Dim sSAPModel
    Dim sBasePath
    
    sBasePath = "C:\OPTMODELS\Overload"
    sSAPModel = "000"

DoCmd.TransferText acExportDelim, "DrsqmodresTLanesNeededToCreateUpload Export Specification", "DrsqmodresTLanesNeededToCreateUpload", sBasePath & "\Data\NewLanes.csv", False

    sapInstance = "A01"   'sAPOinstance
    Set session = SAP_GetSession(sapInstance, False)
    If session Is Nothing Then
        MsgBox "No active SAP session found for """ & sapInstance & """." & vbNewLine & "Please log on to """ & sapInstance & """ and try the operation again.", vbExclamation, Application.CurrentProject.Name
    Else

        session.findById("wnd[0]/tbar[0]/btn[3]").press
        session.findById("wnd[0]/tbar[0]/btn[3]").press
        session.findById("wnd[0]/tbar[0]/btn[3]").press
        session.findById("wnd[0]").resizeWorkingPane 155, 26, False
        session.findById("wnd[0]/tbar[0]/okcd").text = "ZSP_TLANE_MAT_UP"
        session.findById("wnd[0]/tbar[0]/btn[0]").press
        session.findById("wnd[0]/usr/ctxtSP_MODEL").text = sSAPModel
        session.findById("wnd[0]/usr/ctxtSP_FPATH").text = sBasePath & "\Data\NewLanes.csv"
        session.findById("wnd[0]/usr/ctxtSP_FPATH").SetFocus
        session.findById("wnd[0]/usr/ctxtSP_FPATH").caretPosition = 114
        session.findById("wnd[0]/tbar[1]/btn[8]").press
        'session.findById("wnd[0]/tbar[0]/btn[3]").press
        session.findById("wnd[0]/usr/chkSP_TEST").Selected = True
        'session.findById("wnd[0]/usr/chkSP_TEST").setFocus
        'session.findById("wnd[0]/tbar[1]/btn[8]").press
        'session.findById("wnd[0]/tbar[0]/btn[3]").press
        'session.findById("wnd[0]/tbar[0]/btn[3]").press
    MsgBox "If no warnings uncheck test mode and re-execute"
    End If

End Sub

Private Sub cmdViewCHANGE_Click()
DoCmd.OpenTable "tblCHANGE"
End Sub


Private Sub cmdViewDEMAND_Click()
   DoCmd.OpenTable "tblDEMAND"
End Sub


Private Sub cmdViewMachines_Click()
   DoCmd.OpenTable "tlkMachine"
End Sub


Private Sub cmdViewMills_Click()
   DoCmd.OpenTable "tlkMill"
End Sub


Private Sub cmdViewTimePeriods_Click()
   DoCmd.OpenTable "tlkTimePeriod"
End Sub




Private Sub cmdWeeklyRefresh_Click()
   Dim i    As Boolean
   Dim sSQL As String
   Dim Y As Boolean
   
   
    If MsgBox("This Should Only Be Run To Refresh Time Periods and ZPPY Data" & vbNewLine & "Click Cancel to Stop Update.", vbOKCancel) = vbOK Then
   
           '// REDO TIME PERIODS
           i = CreateTimePeriods()
           DoCmd.SetWarnings False
           '// Set Control Tables to No Impact
                sSQL = "UPDATE drstCapAdj SET drstCapAdj.CapAdj = 0"
                DoCmd.RunSQL (sSQL)
                sSQL = "UPDATE drstNextBlkDaysAdj SET drstNextBlkDaysAdj.NextBlkDaysAdj = -999"
                DoCmd.RunSQL (sSQL)
                sSQL = "DELETE drstModeShiftExclude.* FROM drstModeShiftExclude"
                DoCmd.RunSQL (sSQL)
                sSQL = "DELETE drstComboTypeExclusion.* FROM drstComboTypeExclusion"
                DoCmd.RunSQL (sSQL)
                sSQL = "UPDATE tlkMachine SET tlkMachine.MachineActive = -1"
                DoCmd.RunSQL (sSQL)
                sSQL = "DELETE trepMoves.* FROM trepMoves"
                DoCmd.RunSQL (sSQL)
                sSQL = "Delete drstShedDemand.* From drstShedDemand"
                DoCmd.RunSQL (sSQL)
        
           '// REDO ZPPY LEVEL 1 & 2 SNAPSHOTS (LOCAL TABLES)
           sSQL = "DELETE tdatZPPYLevel1.* FROM tdatZPPYLevel1 "
           DoCmd.RunSQL (sSQL)
           sSQL = "DELETE tdatZPPYLevel2.* FROM tdatZPPYLevel2 "
           DoCmd.RunSQL (sSQL)
        
           DoCmd.OpenQuery "qappLocalZPPYLevel1"
           
           'use this statement to get level 2 from SQL
           DoCmd.OpenQuery "qappLocalZPPYLevel2"
                      
''''           'use these two statements to get zppy from manual text files in C:\temp\zppy
''''           Y = ImportZPPY2Files()
''''           DoCmd.OpenQuery "drsqzppy2cleanappend"
           
           '// IMPORT RATE INFO FROM SAP TXT FILE
           'FLAT FILE FIXED WIDTH BUT 3000 RECORDS OF ALL HISTORY
           '\\s02afs01\Lenox-275\CST\Supply Planning\Integrated Master Planning\DataExtracts\SAP\SAPRateInfo.txt
           'WOULD LIKE TO HIT THIS BUT -->linked to TEXTFILE and ISAM driver chokes
           'SELECT qRates4MachineRates.Grade AS GRADE, qRates4MachineRates.[Basis Weight] AS BW, qRates4MachineRates.Plnt AS MILLABBR, qRates4MachineRates.[Work Ctr] AS MachineCode, Round([qRates4MachineRates].[TPD],2) AS TPD, *
           'FROM qRates4MachineRates IN '\\s02afs01\Lenox-275\CST\Supply Planning\Integrated Master Planning\DataExtracts\SAP\SAPDataCleansing60.mdb';
        
        cmdRunPrep_Click
        
        DoCmd.SetWarnings False
        DoCmd.OpenQuery "drsqdellMatExt"
        DoCmd.OpenQuery "drsqappMatExt"
        DoCmd.OpenQuery "qupdComboTypeInitToXout"
        DoCmd.OpenQuery "qupdComboTypeCurrentZPPY"
        DoCmd.OpenQuery "qupdComboTypeCurrentDemand"
        
        sSQL = "DELETE drstTLanesAll.* FROM drstTLanesAll "
        DoCmd.RunSQL (sSQL)
        DoCmd.OpenQuery "drsqappTLanesAll"
        
        
        DoCmd.SetWarnings True
           MsgBox "Weekly Data Refresh Complete."
    End If
    
   
   
End Sub


Private Sub Form_Load()
   Me.lblVersion.Caption = basUtility.GetVersion
End Sub

===============
Form_frmnavOptResults
===============
Option Compare Database
Option Explicit

Private Sub cmdCloseForm_Click()
    DoCmd.Close acForm, [formname]
    DoCmd.OpenForm "frmMain"
End Sub
