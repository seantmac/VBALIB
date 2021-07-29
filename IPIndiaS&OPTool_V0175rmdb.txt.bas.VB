'15,000 lines of code in INDIA Project



'~~~~~~~~~~~~~~~~~~~~~~~~   4500    lines of code removed  for basUtility   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~   1600    lines of code removed  for modExcelPivot      ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~   1500    lines of code removed  for IFcommon      ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~    500    lines of code removed  for IFcustom      ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



===============+
Form_HelpMainPage
===============
Option Compare Database

Private Sub btnApplyChanges_Click() 'Used to be Command39 - Check mark
DoCmd.Close acForm, "HelpMainPage"
End Sub

Private Sub btnForITUseOnly_Click() 'Used to be Command45 - For IT Use Only
Dim sPasswordInput As String

sPasswordInput = ""

sPasswordInput = InputBox("Enter Code To Manage Display Of Access Controls")

If UCase(sPasswordInput) = "SHOW" Then
    ShowNavigationPane
    ShowRibbon
    btnApplyChanges_Click
ElseIf UCase(sPasswordInput) = "HIDE" Then
    HideNavigationPane
    HideRibbon
    btnApplyChanges_Click
End If

    
End Sub
===============+
Form_ResetWorkingMonthWarning
===============
Option Compare Database

Private Sub Command39_Click()
DoCmd.OpenForm "mainpage"
DoCmd.Close acForm, "resetworkingmonthwarning"
End Sub
===============+
Form_fintVolumeDashboard
===============
Private Sub btnBack_Click() 'Used to be Command32 - Back
DoCmd.Close acForm, [formname]
End Sub

Private Sub Command76_Click()
DoCmd.OpenQuery "qRepTradeAndDirectPurchase2BalanceSummaryContract"
End Sub

Private Sub Command53_Click()
DoCmd.OpenQuery "q910Report015PMBalanceAllMonths"
End Sub

Private Sub Command54_Click()
DoCmd.OpenQuery "q910Report019InvBalanceAllMonths"
End Sub
===============+
Form_MachineDays
===============
Private Sub Command32_Click()
DoCmd.Close acForm, "MachineDays"
End Sub
===============+
Form_MachineRatesAndCosts
===============
Private Sub Command32_Click()
DoCmd.Close acForm, "MachineRatesAndCosts"
End Sub
===============+
Form_FreightRates
===============
Private Sub Command32_Click()
DoCmd.Close acForm, "FreightRates"
End Sub
===============+
Form_MaintainFreightCostData
===============

Private Sub Command128_Click()
DoCmd.SetWarnings False

sSQL = "SELECT Name FROM MSysObjects WHERE (Flags=32 Or Flags=48 Or Flags=64 Or Flags=80) order by Name;"

Set RS = Nothing
Set RS = Application.CurrentProject.Connection.Execute(sSQL)

    While Not RS.EOF
        sQuery = RS.Fields("Name").value
        If Left(sQuery, 10) = "q010Init02" Then
            'Debug.Print sQuery & "Starting"
            DoCmd.OpenQuery sQuery
            'Debug.Print sQuery & "Completed"
        End If
        RS.MoveNext
    Wend

Set RS = Nothing

DoCmd.SetWarnings True
Application.Echo False
DoCmd.Close acForm, "MaintainFreightCostData"
DoCmd.OpenForm "MaintainFreightCostData"
Form_MaintainFreightCostData.FreightRates.SetFocus
DoCmd.Maximize
Application.Echo True
MsgBox "Freight Rates Updated"
End Sub

Private Sub Command129_Click()
DoCmd.SetWarnings False
DoCmd.OpenQuery "q320Misc024UpdateTModeFromHistory"
DoCmd.SetWarnings True
MsgBox "Transit Modes Have Been Updated From History" & vbNewLine & "Any Record Not In History Maintained The Same Mode"
End Sub

Private Sub Command130_Click()
DoCmd.SetWarnings False
DoCmd.OpenQuery "q320Misc026UpdateAllTModeKORYfromSVET"
DoCmd.SetWarnings True
MsgBox "KORY Transit Modes Have Been Updated To Match SVET"
End Sub

Private Sub Command32_Click()
DoCmd.OpenForm "MainPage"
DoCmd.Close acForm, "MaintainFreightCostData"
End Sub


===============+
Form_OptModelInterface
===============
Dim bMap As Boolean



Private Sub Command111_Click()
    On Error GoTo Err_Handler

'Dim mpAppRail As MapPoint.Application
'Dim mpAppTruck As MapPoint.Application
Dim mpApp As Object 'MapPoint.Application
Dim myMap As Object 'MapPoint.Map
'Dim RailMap As MapPoint.Map
'Dim TruckMap As MapPoint.Map



MsgBox "This Will Take About 1 Minute." & vbNewLine & _
"Wait For The Completed Message."

DoCmd.SetWarnings False
sSQL = "SELECT q494ModelMap010COPRByDistPointRail.* INTO tblCOPRByDistPointRail IN '" & sBasePath & "JMAMPT\Reporting\JMAMPTReporting.mdb' FROM      q494ModelMap010COPRByDistPointRail;"
DoCmd.RunSQL "SELECT q494ModelMap010COPRByDistPointRail.* INTO tblCOPRByDistPointRail IN '" & sBasePath & "JMAMPT\Reporting\JMAMPTReporting.mdb' FROM q494ModelMap010COPRByDistPointRail;"
DoCmd.RunSQL "SELECT q494ModelMap010COPRByDistPointTruck.* INTO tblCOPRByDistPointTruck IN '" & sBasePath & "JMAMPT\Reporting\JMAMPTReporting.mdb' FROM q494ModelMap010COPRByDistPointTruck;"
DoCmd.RunSQL "SELECT q494ModelMap010COPRSharedByDistPointRail.* INTO tblCOPRSharedByDistPointRail IN '" & sBasePath & "JMAMPT\Reporting\JMAMPTReporting.mdb' FROM q494ModelMap010COPRSharedByDistPointRail;"
DoCmd.RunSQL "SELECT q494ModelMap010COPRSharedByDistPointTruck.* INTO tblCOPRSharedByDistPointTruck IN '" & sBasePath & "JMAMPT\Reporting\JMAMPTReporting.mdb' FROM q494ModelMap010COPRSharedByDistPointTruck;"
DoCmd.RunSQL "SELECT q494ModelMap010AllByDistPointRail.* INTO tblAllByDistPointRail IN '" & sBasePath & "JMAMPT\Reporting\JMAMPTReporting.mdb' FROM q494ModelMap010AllByDistPointRail;"
DoCmd.RunSQL "SELECT q494ModelMap010AllByDistPointTruck.* INTO tblAllByDistPointTruck IN '" & sBasePath & "JMAMPT\Reporting\JMAMPTReporting.mdb' FROM q494ModelMap010AllByDistPointTruck;"
DoCmd.SetWarnings True

Set mpApp = CreateObject("MapPoint.Application")
If mpApp Is Nothing Then
    MsgBox "MapPoint software may not be installed on your machine, in order to generate maps.", vbInformation, Application.CurrentProject.Name
Else

    'mpApp.Visible = True
    'mpApp.UserControl = True
    Set myMap = mpApp.OpenMap(sBasePath & "JMAMPT\Reporting\ReportMasters\COPRRail.ptm")
    myMap.DataSets(2).UpdateLink
    myMap.DataSets.ZoomTo
    myMap.SaveAs sBasePath & "JMAMPT\Current Cycle\COPRRail.ptm"
    
    Set myMap = mpApp.OpenMap(sBasePath & "JMAMPT\Reporting\ReportMasters\COPRSharedRail.ptm")
    myMap.DataSets(2).UpdateLink
    myMap.DataSets.ZoomTo
    myMap.SaveAs sBasePath & "JMAMPT\Current Cycle\COPRSharedRail.ptm"
    
    Set myMap = mpApp.OpenMap(sBasePath & "JMAMPT\Reporting\ReportMasters\AllRail.ptm")
    myMap.DataSets(2).UpdateLink
    myMap.DataSets.ZoomTo
    myMap.SaveAs sBasePath & "JMAMPT\Current Cycle\AllRail.ptm"
    
    Set myMap = mpApp.OpenMap(sBasePath & "JMAMPT\Reporting\ReportMasters\COPRTruck.ptm")
    myMap.DataSets(2).UpdateLink
    myMap.DataSets.ZoomTo
    myMap.SaveAs sBasePath & "JMAMPT\Current Cycle\COPRTruck.ptm"
    
    Set myMap = mpApp.OpenMap(sBasePath & "JMAMPT\Reporting\ReportMasters\COPRSharedTruck.ptm")
    myMap.DataSets(2).UpdateLink
    myMap.DataSets.ZoomTo
    myMap.SaveAs sBasePath & "JMAMPT\Current Cycle\COPRSharedTruck.ptm"
    
    Set myMap = mpApp.OpenMap(sBasePath & "JMAMPT\Reporting\ReportMasters\AllTruck.ptm")
    myMap.DataSets(2).UpdateLink
    myMap.DataSets.ZoomTo
    myMap.SaveAs sBasePath & "JMAMPT\Current Cycle\AllTruck.ptm"
    
    mpApp.Quit
    
    MsgBox "Maps Created In Current Cycle Folder"
End If

Err_Handler:
    If Err Then
        MsgBox "An error occured while generating map - " & sName & vbNewLine & vbNewLine & "Error msg:" & vbNewLine & Err.Number & " - " & Err.Description & vbNewLine & vbNewLine & "Please make sure you have MapPoint software properly installed on your machine, before trying the operation again.", vbExclamation, Application.CurrentProject.Name
    End If
End Sub

Private Sub Command112_Click()

If MsgBox("Clicking OK will replace the original forecast file with the current working month CDP." & vbNewLine & _
"If necessary, the forecast can be reset back to the current APO forecast by using Import Data on the Main Menu." & vbNewLine & _
"Click OK to replace forecast.", vbOKCancel) = vbOK Then
    DoCmd.SetWarnings False
    DoCmd.RunSQL "delete * from tbl002RawData010APOForecast"
    DoCmd.OpenQuery "q320Misc022AppendWorkingMonthCDPToForecast"
    'InitializeMapping "optmodelinterface"
    MsgBox "Original forecast data has been set to match the current CDP." & vbNewLine & _
    "It is recommended that you Initialize Mapping at this time (from the main menu)."
    DoCmd.SetWarnings True
End If

End Sub

Private Sub Command113_Click()
    On Error GoTo Err_Handler
Dim mpApp As Object 'MapPoint.Application
Dim myMap As Object 'MapPoint.Map
Dim sName As String

sName = InputBox("Enter a Short Name For The Map File")
If sName <> "" Then
    

    
    'MsgBox "This Will Take About 1 Minute." & vbNewLine & _
    '"Wait For The Completed Message."
    
    DoCmd.SetWarnings False
    DoCmd.RunSQL "SELECT q494ModelMap010CustomByDistPoint.* INTO tblCustomByDistPoint IN '" & sBasePath & "JMAMPT\Reporting\JMAMPTReporting.mdb'FROM q494ModelMap010CustomByDistPoint;"
    
    Set mpApp = CreateObject("MapPoint.Application")
    If mpApp Is Nothing Then
        MsgBox "MapPoint software may not be installed on your machine, in order to generate maps.", vbInformation, Application.CurrentProject.Name
    Else
        mpApp.Visible = True
        mpApp.UserControl = True
        Set myMap = mpApp.OpenMap(sBasePath & "JMAMPT\Reporting\ReportMasters\Custom.ptm")
        myMap.DataSets(2).UpdateLink
        myMap.DataSets.ZoomTo
        myMap.SaveAs sBasePath & "JMAMPT\Current Cycle\" & sName & ".ptm"
        
        'mpApp.Quit
        'MappointControl1.OpenMap sbasepath & "JMAMPT\Current Cycle\" & sName & ".ptm"
        'bMap = True
        
        MsgBox "Map Saved In Current Cycle Folder"
    End If
Else
    MsgBox "File Name Required.  Operation Cancelled."
End If

Err_Handler:
    If Err Then
        MsgBox "An error occured while generating map - " & sName & vbNewLine & vbNewLine & "Error msg:" & vbNewLine & Err.Number & " - " & Err.Description & vbNewLine & vbNewLine & "Please make sure you have MapPoint software properly installed on your machine, before trying the operation again.", vbExclamation, Application.CurrentProject.Name
    End If
End Sub

Private Sub Command32_Click()
DoCmd.OpenForm "MainPage"
DoCmd.Close acForm, "OptModelInterface"
bMap = False
End Sub


Private Sub Command48_Click()

    Dim scmd As String
    Dim sCOPTPathAndFile As String
    Dim db As DAO.Database
    Dim RS As ADODB.Recordset
    Dim dTons As Double
    Dim sHostName As String
    sHostName = Environ$("computername")
    Debug.Print sHostName
    
          
DoCmd.SetWarnings False

DoCmd.RunSQL "delete * from tblDmdByPath in 'C:\OPTMODELS\IPR2\IPR2.mdb'"
bForceDCToggle = Me.ForceDCToggle.value
bForceDCOnlyToggle = Me.ForceDCOnlyToggle.value
If bForceDCOnlyToggle = True Then
    DoCmd.RunSQL "insert into tblDmdByPath in 'C:\OPTMODELS\IPR2\IPR2.mdb' select * from q410ModelPrep016AllPathsForForceDCOnly"
ElseIf bForceDCToggle = True Then
    DoCmd.RunSQL "insert into tblDmdByPath in 'C:\OPTMODELS\IPR2\IPR2.mdb' select * from q410ModelPrep014AllPathsForForceDC"
Else
    DoCmd.RunSQL "insert into tblDmdByPath in 'C:\OPTMODELS\IPR2\IPR2.mdb' select * from q410ModelPrep010AllPaths"
End If
DoCmd.RunSQL "delete * from tblDaysConstraints in 'C:\OPTMODELS\IPR2\IPR2.mdb'"

DoCmd.RunSQL "insert into tblDaysConstraints in 'C:\OPTMODELS\IPR2\IPR2.mdb' select * from q410ModelPrep022MachineDaysConstraints"

DoCmd.RunSQL "delete * from tblDmdConstraints in 'C:\OPTMODELS\IPR2\IPR2.mdb'"
bDemandLimitToggle = Me.DemandLimitToggle.value
If bDemandLimitToggle = False Then
DoCmd.RunSQL "insert into tblDmdConstraints in 'C:\OPTMODELS\IPR2\IPR2.mdb' select * from q410ModelPrep020DemandConstraints"
Else
DoCmd.RunSQL "insert into tblDmdConstraints in 'C:\OPTMODELS\IPR2\IPR2.mdb' select * from q410ModelPrep021DemandConstraints2WithLimits"
End If

DoCmd.RunSQL "delete * from tblDistPointConstraints in 'C:\OPTMODELS\IPR2\IPR2.mdb'"
DoCmd.RunSQL "insert into tblDistPointConstraints in 'C:\OPTMODELS\IPR2\IPR2.mdb' select * from q410ModelPrep024DistPointTonsConstraints"

DoCmd.RunSQL "delete * from tblDistPointProdGroupConstraints in 'C:\OPTMODELS\IPR2\IPR2.mdb'"
DoCmd.RunSQL "insert into tblDistPointProdGroupConstraints in 'C:\OPTMODELS\IPR2\IPR2.mdb' select * from q410ModelPrep025DistPointProdGroupTonsConstraints"

DoCmd.RunSQL "delete * from tblMillModeConstraints in 'C:\OPTMODELS\IPR2\IPR2.mdb'"
DoCmd.RunSQL "insert into tblMillModeConstraints in 'C:\OPTMODELS\IPR2\IPR2.mdb' select * from q410ModelPrep026MillModeTonsConstraints"

DoCmd.RunSQL "delete * from tblMillModeProdGroupConstraints in 'C:\OPTMODELS\IPR2\IPR2.mdb'"
DoCmd.RunSQL "insert into tblMillModeProdGroupConstraints in 'C:\OPTMODELS\IPR2\IPR2.mdb' select * from q410ModelPrep027MillModeProdGroupTonsConstraints"

DoCmd.RunSQL "delete * from tblMillProdGroupQualityConstraints in 'C:\OPTMODELS\IPR2\IPR2.mdb'"
DoCmd.RunSQL "insert into tblMillProdGroupQualityConstraints in 'C:\OPTMODELS\IPR2\IPR2.mdb' select * from q410ModelPrep028MillProdGroupQualityTonsConstraints"

'MIP
DoCmd.RunSQL "UPDATE tsysDefCol in 'C:\OPTMODELS\IPR2\IPR2.mdb' SET tsysDefCol.BNDBinary = True WHERE (((tsysDefCol.ColType)=""SALESON""))"
'Continuous
'DoCmd.RunSQL "UPDATE tsysDefCol in 'C:\OPTMODELS\IPR2\IPR2.mdb' SET tsysDefCol.BNDBinary = False WHERE (((tsysDefCol.ColType)=""SALESON""))"


'MsgBox "Data Send Complete"



DoCmd.SetWarnings True
          
       sCOPTPathAndFile = "C:\Progra~1\BMOS\C-OPT\C-OPTConsole.exe " & _
                            "RUN /PRJ IPR2 /WORKDIR C:\OPTMODELS\IPR2\OUTPUT /Solver CoinMP /Sense MIN -1 /SetRelativeGapTolerance " & CIntlNumber("0.001", 3) & " /SetTimeLimit " & CIntlNumber("30.000", 3) & " 300 /NOPROMPT"
       If InStr(sHostName, "SMACDER") > 0 Then
           sCOPTPathAndFile = "C:\Programs\MATH\C-OPTx64\C-OPTConsole.exe " & _
                            "RUN /PRJ IPR2 /WORKDIR C:\OPTMODELS\IPR2\OUTPUT /Solver CoinMP /Sense MIN -1 /SetRelativeGapTolerance " & CIntlNumber("0.001", 3) & " /SetTimeLimit " & CIntlNumber("30.000", 3) & " 300 /NOPROMPT"
       End If
       
       scmd = sCOPTPathAndFile
     
'        If MsgBox("Click YES to Run Opt Model or NO to cancel.  Run will take about 5 minutes.", vbYesNo + vbQuestion) = vbYes Then
      
              'me.runcopt.
              DoCmd.Hourglass (True)
              ExecCmd (scmd)
              DoCmd.Hourglass (False)
'              MsgBox ("C-OPT run complete.")
'        End If


'Apply Results
    DoCmd.SetWarnings False
    DoCmd.RunSQL "select * into tbl420ModelResults010RawData from tmtxCol01_SALES in 'C:\OPTMODELS\IPR2\IPR2.mdb'"
    
    Set RS = Nothing
    Set RS = Application.CurrentProject.Connection.Execute("select sum(iif(activity is null, 0, activity)) as Tons from tbl420ModelResults010RawData")
    
    If RS.Fields("Tons").value = 0 Then
        MsgBox "Opt Run Failed, No Results To Import"
    Else
        DoCmd.OpenQuery "q420ModelResultsMaxSourcingPathWithTons"
        DoCmd.OpenQuery "q430ModelResults010ApplyOptSourcing"
        If bDemandLimitToggle = True Then
            DoCmd.OpenQuery "q430ModelResults020ApplyOptTons"
            DoCmd.OpenQuery "q430ModelResults022ApplyOptTonsNullActivity"
        End If
        DoCmd.SetWarnings True
        DoCmd.Echo False
        DoCmd.Close acForm, "OptModelInterface"
        DoCmd.OpenForm "optModelInterface"
        DoCmd.Maximize
        DoCmd.Echo True
        Forms("optModelInterface").ForceDCToggle.value = bForceDCToggle
        Forms("optModelInterface").DemandLimitToggle.value = bDemandLimitToggle
        Forms("optModelInterface").ForceDCOnlyToggle.value = bForceDCOnlyToggle

        MsgBox "Optimization Sourcing Applied"
    End If

End Sub




Private Sub Command55_Click()
DoCmd.SetWarnings False
bForceDCToggle = Me.ForceDCToggle.value
bDemandLimitToggle = Me.DemandLimitToggle.value
bForceDCOnlyToggle = Me.ForceDCOnlyToggle.value

If bForceDCOnlyToggle = True Then
    DoCmd.OpenQuery "q810Lkp014TotalCostByPath9bMinCostForceDCOnlyBestSupply"
    DoCmd.OpenQuery "q310Modify010ApplyBestSourcingToMain"
ElseIf bForceDCToggle = True Then
    DoCmd.OpenQuery "q810Lkp014TotalCostByPath7MinCostForceDCBestSupply"
    DoCmd.OpenQuery "q310Modify010ApplyBestSourcingToMain"
Else
    DoCmd.OpenQuery "q810Lkp014TotalCostByPath3BestSupply"
    DoCmd.OpenQuery "q310Modify010ApplyBestSourcingToMain"
End If
DoCmd.Echo False
DoCmd.Close acForm, "OptModelInterface"
DoCmd.OpenForm "OptModelInterface"
DoCmd.Maximize
DoCmd.Echo True
Forms("optModelInterface").ForceDCToggle.value = bForceDCToggle
Forms("optModelInterface").ForceDCOnlyToggle.value = bForceDCOnlyToggle
Forms("optModelInterface").DemandLimitToggle.value = bDemandLimitToggle

DoCmd.SetWarnings True
MsgBox "Update Complete"
End Sub

Private Sub Command59_Click()

If MsgBox("Click YES to Modify Default Sourcing or NO to cancel", vbYesNo + vbQuestion) = vbYes Then
    DoCmd.SetWarnings False
    DoCmd.OpenQuery "q310Modify020UpdateDistPointSourcingFromCurrent"
    DoCmd.SetWarnings True
    MsgBox "Distribution Point Mapping Updated"
End If

End Sub


Private Sub Command72_Click()
DoCmd.SetWarnings False
bForceDCToggle = Me.ForceDCToggle.value
bDemandLimitToggle = Me.DemandLimitToggle.value
bForceDCOnlyToggle = Me.ForceDCOnlyToggle.value

DoCmd.OpenQuery "q320Misc014ClearTblMainWorkingMonth"
DoCmd.OpenQuery "q320Misc016AppendTblMainWorkingMonthFromBackup"

DoCmd.Echo False
DoCmd.Close acForm, "OptModelInterface"
DoCmd.OpenForm "OptModelInterface"
DoCmd.Maximize
DoCmd.Echo True
Forms("optModelInterface").ForceDCToggle.value = bForceDCToggle
Forms("optModelInterface").ForceDCOnlyToggle.value = bForceDCOnlyToggle
Forms("optModelInterface").DemandLimitToggle.value = bDemandLimitToggle


DoCmd.SetWarnings True
MsgBox "Prior Sourcing Restored"
End Sub

Private Sub Command88_Click()
MsgBox "If this box is checked, active customer/products on the 'ForceDC' tab will be sold out of a DC." _
& vbNewLine & vbNewLine & "Other customer/products will NOT be allowed to be sourced from a DC." _
& vbNewLine & vbNewLine & "WARNING - If this box is checked and KORY direct shipments are limited it could cause a bad Optimization Run."
End Sub

Private Sub Command89_Click()
MsgBox "If this box is checked, active customer/products on the 'ForceDC' tab will be sold out of a DC." _
& vbNewLine & vbNewLine & "Other customer/products WILL be allowed to be sourced from a DC."
End Sub

Private Sub ForceDCOnlyToggle_AfterUpdate()
If Me.ForceDCOnlyToggle.value = True Then
    Me.ForceDCToggle.value = False
End If
End Sub

Private Sub ForceDCToggle_AfterUpdate()
If Me.ForceDCToggle.value = True Then
    Me.ForceDCOnlyToggle.value = False
End If
End Sub

Private Sub Form_Close()
If bMap = True Then
    MappointControl1.closemap
End If
End Sub

Private Sub Form_Open(Cancel As Integer)
Set RS = Application.CurrentProject.Connection.Execute("select * from tbl001workingmonth")
Me.cmbWorkingMonth.value = RS.Fields("workingmonth").value
Set RS = Nothing
Me.ForceDCOnlyToggle.value = True
Me.ForceDCToggle.value = False
End Sub



Private Sub cmbWorkingMonth_AfterUpdate()
DoCmd.SetWarnings False
DoCmd.RunSQL "UPDATE tbl001workingmonth SET tbl001workingmonth.workingmonth = (cmbWorkingMonth)"
DoCmd.Echo False
DoCmd.Close acForm, "OptModelInterface"
DoCmd.OpenForm "OptModelInterface"
DoCmd.Echo True
DoCmd.SetWarnings True
End Sub
===============+
Form_fNotes
===============
Option Compare Database

Private Sub btnApplyChanges_Click() 'Used to be Command39 - Apply Changes
DoCmd.Close acForm, "fNotes"
End Sub
===============+
Form_OpFcstMainPage
===============














Private Sub Command45_Click()
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

Private Sub Command64_Click()
MsgBox "Disabled"
'MsgBox "If Any Master Planning Excel Report Is Open, Close It, Then Click OK To Proceed"


'DoCmd.SetWarnings False
'DoCmd.RunSQL "SELECT q920ExcelReport010FullDataExtract.* INTO tblFullDataExtract IN '" & sbasepath & "JMAMPT\Reporting\JMAMPTReporting.mdb'FROM q920ExcelReport010FullDataExtract;"
'DoCmd.RunSQL "SELECT q910Report015PMBalanceAllMonths.* INTO tblPMBalanceAllMonths IN '" & sbasepath & "JMAMPT\Reporting\JMAMPTReporting.mdb'FROM q910Report015PMBalanceAllMonths;"

'DoCmd.SetWarnings True

'Dim xlApp As Object
'Set xlApp = CreateObject("Excel.Application")
'xlApp.Visible = True
'xlApp.Workbooks.Open (sbasepath & "JMAMPT\Reporting\ReportMasters\JMAMPT_AllDataMaster.xls")
'xlApp.Run ("Refresh")
'xlApp.displayalerts = False
'xlApp.Quit
'Set xlApp = Nothing

'MsgBox "New Reports Generated In Folder 'Current Cycle'"
End Sub



Private Sub Command66_Click()

'MsgBox "Disabled for Demo Mode"
DoCmd.TransferText acExportDelim, "Tbl710OpFcst020OpForecastMain Export Specification", "tbl710OpFcst020OpForecastMain", sBasePath & "JMAMPT\ReplenishmentData\Export_OpFcst.txt", False
MsgBox "New File Generated In 'ReplenishmentData' Folder"

End Sub

Private Sub Command67_Click()
DoCmd.OpenForm "fintOpFcst"
End Sub

Private Sub Command68_Click()
DoCmd.OpenForm "HelpMainPage"
End Sub













Private Sub Form_Open(Cancel As Integer)


DoCmd.Maximize


End Sub



===============+
Form_fOpFcstNotes
===============
Option Compare Database

Private Sub Command39_Click()
DoCmd.Close acForm, "fOpFcstNotes"
End Sub
===============+
Form_fintOpFcst
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
    DoCmd.SetWarnings False
    Me.ViewOnlyArea.SourceObject = ""
    fChangeTablePrep Me.Name, "VO"
    Me.ViewOnlyArea.SourceObject = "Table." & sfrmChangesControlTablePrefix & "changecandidatesVO"
    Me.Refresh
    HideChangeColumns Me.Name, "VO", "ViewOnlyArea"
    DoCmd.SetWarnings True
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

sSQL = "SELECT Name FROM MSysObjects WHERE Name = """ & sfrmChangesControlTablePrefix & "changecandidatesMM""" & ";"
Set rs2 = Application.CurrentProject.Connection.Execute(sSQL)

If Not rs2.EOF Then
    Set RS = Application.CurrentProject.Connection.Execute("select sum(" & sUOM & ") as MoveSubtotal from " & sfrmChangesControlTablePrefix & "changecandidatesMM where Move = -1")
    Me.SubtotalMM.value = RS.Fields("MoveSubtotal").value
End If
End Sub
===============+
basUtility
===============
'//  //
'/|     MODULE:  basUtility                                         |/
'/|    PURPOSE:  Utility Functions for Access Applications          |/
'/|         BY:  Sean                                               |/
'/|       DATE:  11/30/96                                           |/
'/|    HISTORY:  11/30/96    Initial Release                        |/
'/|              1/11/97     Added Some New Routines                |/
'/|              7/6/2009    UPDATE WITH SOME NEW CODE              |/
'/|              11/1/2012   UPDATE WITH SOME NEW CODE AND COMBINE  |/
'/|                            with modLibrary from Ravi and Cull   |/
'/|                            unused routines.                     |/
'/|              04/02/2015  Update RunQueriesByPrefix and          |/
'/|                            FieldExists by Sean and Anjali       |/
'/|              06/01/2015  Update other stuff                     |/
'/|              06/04/2015  More updates                           |\
'/|                          Added AuditCruise, GetVersion,GetScalar|/
'/|                          ExecSQL, ChangeSettingsTablesTestMode  |/
'/|                          FormatFilePath, FormatDate             |/
'/|                          dateGetPriorMonth and dateMonthLastDay |/
'/|              06/05/2015  Added GetScalar Again STM              |\
'/|              02/20/2016  Added Compact&Repair SubRoutines       |\
'/|                                GetAccessFileVersion SubRoutines |/
'/|                                                                 |/
'/|              04/05/2016  TableExists()  & QueryExists() by Sean |/
'/|              05/04/2016  FileExists()   & FolderExists()        |/
'/|                             and Move GetSetting() back in here  |/
'/|                                                                 |/
'/|                                                                 |/
'/|                                                                 |/
'/|                                                                 |/
'/|                                                                 |/
'/|                                                                 |/
'//  //

'Attribute VB_Name = "basUtility"

'~~~~~~~~~~~~~~~~~~~~~~~~   4500    lines of code removed  for basUtility   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


===============+
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
    Dim xlWB            As Object 'Workbook

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


===============+
Form_fintChanges1
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
    DoCmd.SetWarnings False
    Me.ViewOnlyArea.SourceObject = ""
    fChangeTablePrep Me.Name, "VO"
    Me.ViewOnlyArea.SourceObject = "Table." & sfrmChangesControlTablePrefix & "changecandidatesVO"
    Me.Refresh
    HideChangeColumns Me.Name, "VO", "ViewOnlyArea"
    DoCmd.SetWarnings True
End Sub


Private Sub cmdCloseForm_Click()        'was command21
    DoCmd.Close acForm, [formname]
    DoCmd.OpenForm "MainPage"
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

sSQL = "SELECT Name FROM MSysObjects WHERE Name = """ & sfrmChangesControlTablePrefix & "changecandidatesMM""" & ";"
Set rs2 = Application.CurrentProject.Connection.Execute(sSQL)

If Not rs2.EOF Then
    Set RS = Application.CurrentProject.Connection.Execute("select sum(" & sUOM & ") as MoveSubtotal from " & sfrmChangesControlTablePrefix & "changecandidatesMM where Move = -1")
    Me.SubtotalMM.value = RS.Fields("MoveSubtotal").value
End If
End Sub
===============+
Form_fintChangesASF
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
    DoCmd.SetWarnings False
    Me.ViewOnlyArea.SourceObject = ""
    fChangeTablePrep Me.Name, "VO"
    Me.ViewOnlyArea.SourceObject = "Table." & sfrmChangesControlTablePrefix & "changecandidatesVO"
    Me.Refresh
    HideChangeColumns Me.Name, "VO", "ViewOnlyArea"
    DoCmd.SetWarnings True
End Sub


Private Sub cmdCloseForm_Click()        'was command21
    DoCmd.Close acForm, [formname]
    DoCmd.OpenForm "MainPage"
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

sSQL = "SELECT Name FROM MSysObjects WHERE Name = """ & sfrmChangesControlTablePrefix & "changecandidatesMM""" & ";"
Set rs2 = Application.CurrentProject.Connection.Execute(sSQL)

If Not rs2.EOF Then
    Set RS = Application.CurrentProject.Connection.Execute("select sum(" & sUOM & ") as MoveSubtotal from " & sfrmChangesControlTablePrefix & "changecandidatesMM where Move = -1")
    Me.SubtotalMM.value = RS.Fields("MoveSubtotal").value
End If
End Sub
===============+
Form_fintChangesUDP
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
    DoCmd.SetWarnings False
    Me.ViewOnlyArea.SourceObject = ""
    fChangeTablePrep Me.Name, "VO"
    Me.ViewOnlyArea.SourceObject = "Table." & sfrmChangesControlTablePrefix & "changecandidatesVO"
    Me.Refresh
    HideChangeColumns Me.Name, "VO", "ViewOnlyArea"
    DoCmd.SetWarnings True
End Sub


Private Sub cmdCloseForm_Click()        'was command21
    DoCmd.Close acForm, [formname]
    DoCmd.OpenForm "MainPage"
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

sSQL = "SELECT Name FROM MSysObjects WHERE Name = """ & sfrmChangesControlTablePrefix & "changecandidatesMM""" & ";"
Set rs2 = Application.CurrentProject.Connection.Execute(sSQL)

If Not rs2.EOF Then
    Set RS = Application.CurrentProject.Connection.Execute("select sum(" & sUOM & ") as MoveSubtotal from " & sfrmChangesControlTablePrefix & "changecandidatesMM where Move = -1")
    Me.SubtotalMM.value = RS.Fields("MoveSubtotal").value
End If
End Sub
===============+
Form_fintChangesMP
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
    DoCmd.SetWarnings False
    Me.ViewOnlyArea.SourceObject = ""
    fChangeTablePrep Me.Name, "VO"
    Me.ViewOnlyArea.SourceObject = "Table." & sfrmChangesControlTablePrefix & "changecandidatesVO"
    Me.Refresh
    HideChangeColumns Me.Name, "VO", "ViewOnlyArea"
    DoCmd.SetWarnings True
End Sub


Private Sub cmdCloseForm_Click()        'was command21
    DoCmd.Close acForm, [formname]
    DoCmd.OpenForm "MainPage"
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

sSQL = "SELECT Name FROM MSysObjects WHERE Name = """ & sfrmChangesControlTablePrefix & "changecandidatesMM""" & ";"
Set rs2 = Application.CurrentProject.Connection.Execute(sSQL)

If Not rs2.EOF Then
    Set RS = Application.CurrentProject.Connection.Execute("select sum(" & sUOM & ") as MoveSubtotal from " & sfrmChangesControlTablePrefix & "changecandidatesMM where Move = -1")
    Me.SubtotalMM.value = RS.Fields("MoveSubtotal").value
End If
End Sub


Private Sub SubtotalMM_Exit(Cancel As Integer)
    MoveArea_Exit Cancel
End Sub
===============+
Form_frmSettingsExcelPivots2
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
    Dim xlWB            As Object 'Workbook

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
===============+
Form_fintChangesFlex
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
    DoCmd.SetWarnings False
    Me.ViewOnlyArea.SourceObject = ""
    fChangeTablePrep Me.Name, "VO"
    Me.ViewOnlyArea.SourceObject = "Table." & sfrmChangesControlTablePrefix & "changecandidatesVO"
    Me.Refresh
    HideChangeColumns Me.Name, "VO", "ViewOnlyArea"
    DoCmd.SetWarnings True
End Sub


Private Sub cmdCloseForm_Click()        'was command21
    DoCmd.Close acForm, [formname]
    DoCmd.OpenForm "MainPage"
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

sSQL = "SELECT Name FROM MSysObjects WHERE Name = """ & sfrmChangesControlTablePrefix & "changecandidatesMM""" & ";"
Set rs2 = Application.CurrentProject.Connection.Execute(sSQL)

If Not rs2.EOF Then
    Set RS = Application.CurrentProject.Connection.Execute("select sum(" & sUOM & ") as MoveSubtotal from " & sfrmChangesControlTablePrefix & "changecandidatesMM where Move = -1")
    Me.SubtotalMM.value = RS.Fields("MoveSubtotal").value
End If
End Sub

===============+
Common
===============
Option Explicit
Option Compare Database




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
DoCmd.RunSQL "Delete * from tbl210CalcReq012MainReport;"
UnwindTable "q210CalcReq910ReportExtract1", "tbl210CalcReq012MainReport", 6, 5
DoCmd.SetWarnings True
MsgBox "On Order Data Updated"

End Function

Public Function CreatePivotTable()

Dim xlWB As Workbook

Set xlWB = GeneratePivotTable_Generic(, , , 13)
Set xlWB = GeneratePivotTable_Generic(, , , 14, xlWB)

End Function
===============+
Form_FrmAddToMain
===============
Option Compare Database

Private Sub btnBackToMain_Click()
    DoCmd.Close acForm, [formname]
    DoCmd.OpenForm "MainPage"
End Sub




Private Sub cmdExecuteNewRec_Click()
    RunQueriesByPrefix "q300AddRec0"
    Me.Refresh
    Form_MainPage.UpdateMainFrmChkBoxes (12)
    DoCmd.Close acForm, "Mainpage"
    MsgBox "Added New Rec(s)"
End Sub

===============+
modExcelPivot
===============


'~~~~~~~~~~~~~~~~~~~~~~~~   1600    lines of code removed  for modExcelPivot      ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~




===============+
modMain
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

Public Function Import_SalesHistory_AllFields()
'**********************************************
'Author  :  RKP
'Date/Ver:  04-24-13/V01
'Input   :
'Output  :
'Comments:
'**********************************************
    On Error GoTo Err_Handler
    
    Dim minRowID        As Long
    Dim RS              As ADODB.Recordset
    
    'DoCmd.OpenQuery "q010Import013SalesHistoryAllFields05Merge01Clear"

    DoCmd.OpenQuery "q010Import013SalesHistoryAllFields01Line1Clear"
    DoCmd.OpenQuery "q010Import013SalesHistoryAllFields02Line2Clear"
    
    DoCmd.OpenQuery "q010Import013SalesHistoryAllFields03Line1Append"
    DoCmd.OpenQuery "q010Import013SalesHistoryAllFields04Line2Append"

'    DoCmd.OpenQuery "q010Import013SalesHistoryAllFields03Line1Append02"
'    DoCmd.OpenQuery "q010Import013SalesHistoryAllFields04Line2Append02"
    
    Set RS = Application.CurrentProject.Connection.Execute("SELECT MIN(RowID)-1 AS MinRowID FROM tbl002RawData010SalesHistoryAllFieldsLine1")
    minRowID = RS.Fields("MinRowID").value
    DoCmd.RunSQL "UPDATE tbl002RawData010SalesHistoryAllFieldsLine1 SET AdjRowID = RowID-" & minRowID
    
    Set RS = Application.CurrentProject.Connection.Execute("SELECT MIN(RowID)-1 AS MinRowID FROM tbl002RawData010SalesHistoryAllFieldsLine2")
    minRowID = RS.Fields("MinRowID").value
    DoCmd.RunSQL "UPDATE tbl002RawData010SalesHistoryAllFieldsLine2 SET AdjRowID = RowID-" & minRowID
    
    DoCmd.OpenQuery "q010Import013SalesHistoryAllFields06Merge02"

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

'''Public Function GetVersion() As String
''''**********************************************
''''Author  :  RKP
''''Date/Ver:  03-12-13/V01
''''Input   :
''''Output  :
''''Comments:
''''**********************************************
'''    On Error GoTo Err_Handler
'''
''''    Dim version     As String
''''    Dim frm         As Form_MainPage
''''    Dim ctr         As Integer
'''
'''    GetVersion = "v" & basUtility.GetScalarValue("SELECT KeyValue FROM tbl000Settings WHERE KeyName = 'Version'")
'''
'''    'Set frm = Forms("MainPage")
'''    'frm.lblVersion = version
''''    Set frm = New Form_MainPage
''''    frm.lblVersion.Caption = "v" & version
''''
''''    frm.Refresh
''''    frm.Repaint
'''
''''    For Each frm In Application.Forms
''''        Debug.Print frm.Name
''''        'frm.Controls("lblVersion").value = version
''''        For ctr = 0 To frm.Controls.count - 1
''''            If frm.Controls.Item(ctr).value = "lblVersion" Then
''''                Debug.Print frm.Controls.Item(ctr).value
''''            End If
''''        Next
''''    Next
'''
'''Err_Handler:
'''    mlLastErr = Err.Number
'''    msLastErr = Err.Description
'''    'Function1 = mlLastErr
'''    If Err Then
'''        If Err.Number = 49 Then 'Bad DLL calling convention
'''            mlLastErr = 0
'''            msLastErr = ""
'''            Resume Next
'''        Else
'''            'ProcessMsg Err.Number, Err.Description, "", ""
'''            MsgBox Err.Number & " - " & Err.Description
'''        End If
'''    End If
'''
'''    Exit Function
'''    Resume
'''End Function

Public Function RSM_ImportSalesForecasts_Old()
'**********************************************
'Author  :  RKP
'Date/Ver:  03-12-13/V01
'Input   :
'Output  :
'Comments:
'
'**********************************************
    On Error GoTo Err_Handler

    Dim filePath        As String
    Dim xlApp           As Object 'Excel.Application
    Dim xlWB            As Object 'Excel.Workbook
    Dim xlWS            As Object 'Excel.Worksheet
    Dim lastRow         As Long
    Dim lastCol         As Integer
    Dim RS              As ADODB.Recordset
    Dim sql             As String
    Dim rowCtr          As Long
    Dim colCtr          As Integer
    Dim ret             As Long
    Dim rsmFolder       As String
    Dim fso             As Object
    Dim fsoFiles        As Object
    Dim fsoFile         As Object
    Dim noOfCols        As Integer
    Dim minRowIDRSM     As Long
    Dim minRowIDMain    As Long
    Dim minRowIDDiff    As Long
    Dim rowCount        As Long
    Dim remainder       As Integer
    Dim multiplier      As Integer
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    rsmFolder = basUtility.GetSetting("RSMFolder")
    
    Set fsoFiles = fso.GetFolder(rsmFolder).Files
    
    sql = "DELETE * FROM tbl001RSMData"
    Application.CurrentProject.Connection.Execute sql, ret
    
    'Start RSM/Region Loop
    For Each fsoFile In fsoFiles
        DoEvents
        
        sql = "SELECT * FROM tbl001RSMData WHERE 0 = 1"
        Set RS = New ADODB.Recordset
        RS.LockType = adLockBatchOptimistic
        RS.CursorLocation = adUseClient
        RS.Open sql, Application.CurrentProject.Connection, adOpenKeyset, adLockBatchOptimistic, adCmdText
        
        'BasePath = ""
        'Reports
        filePath = rsmFolder & "\" & fsoFile.Name
        If VBA.Dir(filePath) <> "" Then
            Debug.Print filePath
            'filePath = "C:\OPTMODELS\IND1\Reports\APPM_S&OPTool_RSMView_RSM-W_201304.xlsx"
            Set xlApp = CreateObject("Excel.Application")
            xlApp.Visible = True
            Set xlWB = xlApp.Workbooks.Open(filePath)
            xlWB.worksheets("RSMData").Visible = xlSheetVisible
            Set xlWS = xlWB.worksheets("RSMData")
            xlWS.Activate
            
            lastRow = GetLastRow(xlWB, "RSMData", "A", 1)
            
            noOfCols = 0
            For colCtr = 1 To 100
                DoEvents
                
                If xlWS.cells(1, colCtr).value = "" Then Exit For
                noOfCols = noOfCols + 1
            Next
            
            For rowCtr = 2 To lastRow
                DoEvents
                
                RS.AddNew
                
                For colCtr = 1 To noOfCols '21
                    DoEvents
                    
                    Application.SysCmd acSysCmdSetStatus, "Importing RSM Sales Forecasts from..." & fsoFile.Name & "..." & rowCtr
                    RS.Fields(VBA.CStr(xlWS.cells(1, colCtr))).value = xlWS.cells(rowCtr, colCtr).value
                Next
            Next
            Application.SysCmd acSysCmdSetStatus, "Importing RSM Sales Forecasts from..." & fsoFile.Name & "..." & "Saving data into tool..."
            RS.UpdateBatch adAffectAllChapters
            RS.Close
            Set RS = Nothing
            
            xlWB.Close False
            xlApp.Quit
            
            Set xlWS = Nothing
            Set xlWB = Nothing
            Set xlApp = Nothing
        End If
    Next
    
    'End RSM/Region Loop
    
'    sql = "SELECT MIN(RowID) AS MinRowID FROM tbl001RSMData"
'    Set rs = Application.CurrentProject.Connection.Execute(sql, ret)
'    minRowIDRSM = rs.Fields("MinRowID").value
'
'    sql = "SELECT MIN(RowID) AS MinRowID FROM tbl001Main"
'    Set rs = Application.CurrentProject.Connection.Execute(sql, ret)
'    minRowIDMain = rs.Fields("MinRowID").value
'
'    minRowIDDiff = minRowIDMain - minRowIDRSM
'
'    sql = "SELECT COUNT(*) AS RowCount FROM tbl001Main"
'    Set rs = Application.CurrentProject.Connection.Execute(sql, ret)
'    rowCount = rs.Fields("RowCount").value
'
'    remainder = minRowIDDiff Mod rowCount
'
'    If VBA.Abs(minRowIDDiff) > 0 Then
'        If remainder = 0 Then
'            multiplier = minRowIDDiff / rowCount
'            If multiplier > 0 Then
'                sql = "UPDATE tbl001Main INNER JOIN tbl001RSMData ON tbl001Main.RowID = tbl001RSMData.RowID + " & rowCount * multiplier & " SET tbl001Main.SalesTons = [tbl001RSMData].[SalesTons], tbl001Main.UDPTons = [tbl001RSMData].[SalesTons], tbl001Main.CDPTons = [tbl001RSMData].[SalesTons], tbl001Main.PlanTons = [tbl001RSMData].[SalesTons];"
'                Set rs = Application.CurrentProject.Connection.Execute(sql, ret)
'            Else
'            End If
'        End If
'    Else
'        DoCmd.SetWarnings False
'        DoCmd.OpenQuery "q712OpFcst010UpdateMainForRSMChanges"
'        DoCmd.SetWarnings True
'    End If
    
    DoCmd.SetWarnings False
    DoCmd.OpenQuery "q712OpFcst010UpdateMainForRSMChanges"
    DoCmd.SetWarnings True

    MsgBox "Updates sales forecasts from RSM files have been updated successfully!", vbInformation, Application.CurrentProject.Name

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
    Set fsoFile = Nothing
    Set fsoFiles = Nothing
    Set fso = Nothing
    Application.SysCmd acSysCmdSetStatus, "Ready"

    Exit Function
    Resume
End Function

Public Function RSM_ImportSalesForecasts()
'**********************************************
'Author  :  RKP
'Date/Ver:  03-12-13/V01
'Input   :
'Output  :
'Comments:
'
'**********************************************
    On Error GoTo Err_Handler

    Dim filePath        As String
    Dim xlApp           As Object 'Excel.Application
    Dim xlWB            As Object 'Excel.Workbook
    Dim xlWS            As Object 'Excel.Worksheet
    Dim lastRow         As Long
    Dim lastCol         As Integer
    Dim RS              As ADODB.Recordset
    Dim sql             As String
    Dim rowCtr          As Long
    Dim colCtr          As Integer
    Dim ret             As Long
    Dim rsmFolder       As String
    Dim fso             As Object
    Dim fsoFiles        As Object
    Dim fsoFile         As Object
    Dim noOfCols        As Integer
    Dim minRowIDRSM     As Long
    Dim minRowIDMain    As Long
    Dim minRowIDDiff    As Long
    Dim rowCount        As Long
    Dim remainder       As Integer
    Dim multiplier      As Integer
    Dim startTime       As Date
    
    startTime = VBA.Now()
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    rsmFolder = basUtility.GetSetting("RSMFolder")
    
    Set fsoFiles = fso.GetFolder(rsmFolder).Files
    
    sql = "DELETE * FROM tbl001RSMData"
    Application.CurrentProject.Connection.Execute sql, ret
    
    'Start RSM/Region Loop
    For Each fsoFile In fsoFiles
        DoEvents
        
'        sql = "SELECT * FROM tbl001RSMData WHERE 0 = 1"
'        Set rs = New ADODB.Recordset
'        rs.LockType = adLockBatchOptimistic
'        rs.CursorLocation = adUseClient
'        rs.Open sql, Application.CurrentProject.Connection, adOpenKeyset, adLockBatchOptimistic, adCmdText
        
        'BasePath = ""
        'Reports
        filePath = rsmFolder & "\" & fsoFile.Name
        If VBA.Dir(filePath) <> "" Then
            Debug.Print filePath
            
            Application.SysCmd acSysCmdSetStatus, "Importing RSM Sales Forecasts from..." & fsoFile.Name & "..."
            DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12, "tbl001RSMData", filePath, True, "RSMData"
            
'            Set xlApp = CreateObject("Excel.Application")
'            xlApp.Visible = True
'            Set xlWB = xlApp.Workbooks.Open(filePath)
'            xlWB.workSheets("RSMData").Visible = xlSheetVisible
'            Set xlWS = xlWB.workSheets("RSMData")
'            xlWS.Activate
'
'            lastRow = GetLastRow(xlWB, "RSMData", "A", 1)
'
'            noOfCols = 0
'            For colCtr = 1 To 100
'                DoEvents
'
'                If xlWS.cells(1, colCtr).value = "" Then Exit For
'                noOfCols = noOfCols + 1
'            Next
'
'            For rowCtr = 2 To lastRow
'                DoEvents
'
'                rs.AddNew
'
'                For colCtr = 1 To noOfCols '21
'                    DoEvents
'
'                    Application.SysCmd acSysCmdSetStatus, "Importing RSM Sales Forecasts from..." & fsoFile.Name & "..." & rowCtr
'                    rs.Fields(VBA.CStr(xlWS.cells(1, colCtr))).value = xlWS.cells(rowCtr, colCtr).value
'                Next
'            Next
'            Application.SysCmd acSysCmdSetStatus, "Importing RSM Sales Forecasts from..." & fsoFile.Name & "..." & "Saving data into tool..."
'            rs.UpdateBatch adAffectAllChapters
'            rs.Close
'            Set rs = Nothing
'
'            xlWB.Close False
'            xlApp.Quit
'
'            Set xlWS = Nothing
'            Set xlWB = Nothing
'            Set xlApp = Nothing
        End If
    Next
    
    'End RSM/Region Loop
    
'    sql = "SELECT MIN(RowID) AS MinRowID FROM tbl001RSMData"
'    Set rs = Application.CurrentProject.Connection.Execute(sql, ret)
'    minRowIDRSM = rs.Fields("MinRowID").value
'
'    sql = "SELECT MIN(RowID) AS MinRowID FROM tbl001Main"
'    Set rs = Application.CurrentProject.Connection.Execute(sql, ret)
'    minRowIDMain = rs.Fields("MinRowID").value
'
'    minRowIDDiff = minRowIDMain - minRowIDRSM
'
'    sql = "SELECT COUNT(*) AS RowCount FROM tbl001Main"
'    Set rs = Application.CurrentProject.Connection.Execute(sql, ret)
'    rowCount = rs.Fields("RowCount").value
'
'    remainder = minRowIDDiff Mod rowCount
'
'    If VBA.Abs(minRowIDDiff) > 0 Then
'        If remainder = 0 Then
'            multiplier = minRowIDDiff / rowCount
'            If multiplier > 0 Then
'                sql = "UPDATE tbl001Main INNER JOIN tbl001RSMData ON tbl001Main.RowID = tbl001RSMData.RowID + " & rowCount * multiplier & " SET tbl001Main.SalesTons = [tbl001RSMData].[SalesTons], tbl001Main.UDPTons = [tbl001RSMData].[SalesTons], tbl001Main.CDPTons = [tbl001RSMData].[SalesTons], tbl001Main.PlanTons = [tbl001RSMData].[SalesTons];"
'                Set rs = Application.CurrentProject.Connection.Execute(sql, ret)
'            Else
'            End If
'        End If
'    Else
'        DoCmd.SetWarnings False
'        DoCmd.OpenQuery "q712OpFcst010UpdateMainForRSMChanges"
'        DoCmd.SetWarnings True
'    End If
    
    DoCmd.SetWarnings False
    DoCmd.OpenQuery "q712OpFcst010UpdateMainForRSMChanges"
    DoCmd.SetWarnings True

    MsgBox "Sales Forecasts from RSM files have been updated successfully!" & vbNewLine & "Time Taken: " & VBA.DateDiff("s", startTime, VBA.Now()) & " seconds", vbInformation, Application.CurrentProject.Name

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
    Set fsoFile = Nothing
    Set fsoFiles = Nothing
    Set fso = Nothing
    Application.SysCmd acSysCmdSetStatus, "Ready"

    Exit Function
    Resume
End Function

Public Function GenerateRSMFile()
'**********************************************
'Author  :  RKP
'Date/Ver:  03-18-13/V01
'Input   :
'Output  :
'Comments:
'Application.VBE.ActiveVBProject.Description (V01)
'Application.VBE.ActiveVBProject.FileName (C:\OPTMODELS\IND1\S&OPTool\IPIndiaS&OPTool_V01.12.mdb)
'Application.VBE.ActiveVBProject.Name (IPIndiaS&OPTool)
'**********************************************
    On Error GoTo Err_Handler

    Dim filePath        As String
    Dim xlWB            As Object 'Excel.Workbook
    Dim timeStart       As Date
    
    If MsgBox("This operation will:" & vbNewLine & vbNewLine & "Generate an Excel file for RSM's to enter their forecast." & vbNewLine & vbNewLine & "Base File Path:" & vbNewLine & basUtility.GetSetting("BaseFilePath") & vbNewLine & vbNewLine & "Do you want to continue?", vbOKCancel + vbQuestion, Application.CurrentProject.Name) <> vbOK Then GoTo Err_Handler
    
    timeStart = VBA.Now
    
    'Generate the first pivot table that is used to generate RSMEdit.
    Set xlWB = GeneratePivotTable_Generic(, , , 10)
    
    If xlWB Is Nothing Then
    Else
        GenerateRSMEditCustom xlWB
        
        'Generate the second pivot table that forms RSMPivot.
        Set xlWB = GeneratePivotTable_Generic(, , , 12, xlWB)
        
        GenerateRSMPivotCustom xlWB
        
        xlWB.worksheets("RSMData").Visible = xlSheetHidden
        xlWB.worksheets("RSMEdit").Activate
        xlWB.worksheets("RSMEdit").Range("A6").Select
        
        xlWB.Save

        If Application.Forms("MainPage").chkGenerateRSMRegionFiles.value Then
            GenerateRSMRegionFiles timeStart
        Else
            MsgBox "RSM File was created successfully at:" & vbNewLine & filePath & vbNewLine & vbNewLine & "Time taken: " & VBA.DateDiff("s", timeStart, VBA.Now()) & " seconds", vbInformation, Application.CurrentProject.Name
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
        ElseIf Err.Number = 53 Then 'File not found
            mlLastErr = 0
            msLastErr = ""
            Resume Next
        ElseIf Err.Number = 70 Then 'Permission denied
            mlLastErr = 0
            msLastErr = ""
            MsgBox "Error - " & Err.Description & " (#" & Err.Number & ")" & vbNewLine & vbNewLine & "The RSM file:" & vbNewLine & filePath & vbNewLine & "appears to be in use." & vbNewLine & vbNewLine & "Please close the file and try the operation again.", vbExclamation, Application.CurrentProject.Name
        Else
            'ProcessMsg Err.Number, Err.Description, "", ""
'            If trap = "SETTINGS" Then
'                Resume Next
'            Else
            MsgBox Err.Number & " - " & Err.Description
'            End If
        End If
    End If

    Exit Function
    Resume
End Function

Public Function GenerateRSMRegionFiles(Optional timeStart)
'**********************************************
'Author  :  RKP
'Date/Ver:  03-18-13/V01
'Input   :
'Output  :
'Comments:
'**********************************************
    On Error GoTo Err_Handler

    Dim filePath        As String
    Dim xlWB            As Object 'Excel.Workbook
    Dim DataQueryName   As String
    Dim sql             As String
    Dim rsRSMRegion     As ADODB.Recordset
    Dim ret             As Long
    Dim ctrRSMRegion    As Integer
    
    If VBA.IsMissing(timeStart) Then
        If MsgBox("This operation will:" & vbNewLine & vbNewLine & "Generate an Excel file for RSM's to enter their forecast." & vbNewLine & vbNewLine & "Do you want to continue?", vbOKCancel + vbQuestion, Application.CurrentProject.Name) <> vbOK Then GoTo Err_Handler
        timeStart = VBA.Now
    End If
    
    sql = "SELECT DISTINCT [RSMRegion] FROM [" & basUtility.GetSetting("RSMQueryName") & "]"
    Set rsRSMRegion = Application.CurrentProject.Connection.Execute(sql, ret)
    
    ctrRSMRegion = 0
    While Not rsRSMRegion.EOF
        DoEvents
        
        ctrRSMRegion = ctrRSMRegion + 1
        
        DataQueryName = "SELECT * FROM [" & basUtility.GetSetting("RSMQueryName") & "] WHERE RSMRegion = '" & rsRSMRegion("RSMRegion") & "'"
        filePath = FormatFilePath(basUtility.GetSetting("BaseFilePath")) & basUtility.GetSetting("ReportsFolder") & "APPM_S&OPTool_RSMView_" & rsRSMRegion.Fields("RSMRegion").value & "_" & Application.CurrentProject.Connection.Execute("SELECT PlanningCycle FROM tbl001PlanningCycle").Fields("PlanningCycle").value & ".xlsx"
        
        'Generate the first pivot table that is used to generate RSMEdit.
        Set xlWB = GeneratePivotTable_Generic(, , , 14, , DataQueryName, filePath)
        
        If xlWB Is Nothing Then
        Else
            GenerateRSMEditCustom xlWB
            
            'Generate the second pivot table that forms RSMPivot.
            Set xlWB = GeneratePivotTable_Generic(, , , 16, xlWB, , filePath)
            
            GenerateRSMPivotCustom xlWB
            
            xlWB.worksheets("RSMData").Visible = xlSheetHidden
            xlWB.worksheets("RSMEdit").Visible = xlSheetVisible
            xlWB.worksheets("RSMEdit").Activate
            xlWB.worksheets("RSMEdit").Range("A6").Select
            xlWB.worksheets("RSMPivot").Visible = xlSheetVisible
            
            xlWB.Save

        End If
        
        rsRSMRegion.MoveNext
    Wend
    
    MsgBox ctrRSMRegion & " RSM files, along with the RSM Master file were created successfully at:" & vbNewLine & basUtility.GetSetting("BaseFilePath") & basUtility.GetSetting("ReportsFolder") & vbNewLine & vbNewLine & "Time taken: " & VBA.DateDiff("s", timeStart, VBA.Now) & " seconds", vbInformation, Application.CurrentProject.Name
    
Err_Handler:
    mlLastErr = Err.Number
    msLastErr = Err.Description
    'Function1 = mlLastErr
    If Err Then
        If Err.Number = 49 Then 'Bad DLL calling convention
            mlLastErr = 0
            msLastErr = ""
            Resume Next
        ElseIf Err.Number = 53 Then 'File not found
            mlLastErr = 0
            msLastErr = ""
            Resume Next
        ElseIf Err.Number = 70 Then 'Permission denied
            mlLastErr = 0
            msLastErr = ""
            MsgBox "Error - " & Err.Description & " (#" & Err.Number & ")" & vbNewLine & vbNewLine & "The RSM file:" & vbNewLine & filePath & vbNewLine & "appears to be in use." & vbNewLine & vbNewLine & "Please close the file and try the operation again.", vbExclamation, Application.CurrentProject.Name
        Else
            'ProcessMsg Err.Number, Err.Description, "", ""
            MsgBox Err.Number & " - " & Err.Description
        End If
    End If

    Exit Function
    Resume
End Function

Public Function GetLastColumnInRSMEdit(ByRef xlWB As Object) As Long
'**********************************************
'Author  :  RKP
'Date/Ver:  03-22-13/V01
'Input   :
'Output  :
'Comments:
'**********************************************
    On Error GoTo Err_Handler

    Dim rowCtr      As Long
    Dim lastRow     As Long
    Dim ws          As Object 'Worksheet
    
    Set ws = xlWB.worksheets("RSMEdit")
    For rowCtr = 1 To 1000
        DoEvents
        
        If ws.Range("A" & rowCtr).value <> "" Then lastRow = ws.Range("A" & rowCtr).Row
    Next
    
    GetLastColumnInRSMEdit = lastRow

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

Public Function GenerateRSMEditCustom(ByRef xlWB As Object)
'**********************************************
'Author  :  RKP
'Date/Ver:  03-12-13/V01
'Input   :
'Output  :
'Comments:
'**********************************************
    On Error GoTo Err_Handler
    
    Dim wsName          As String
    Dim ws              As Object 'Worksheet
    Dim lastRowEdit     As Long
    Dim lastRowData     As Long
    Dim colBaseKey      As String
    Dim colMonth        As String
    Dim colSalesTons    As String
    Dim colASFTons      As String

    wsName = "RSMPivot"
    
    xlWB.sheets(wsName).Activate
    xlWB.sheets(wsName).Range("A1").Select
    xlWB.sheets(wsName).Range("A1").FormulaR1C1 = "APPM S&OP Tool - RSM View (" & xlWB.worksheets("RSMData").Range("C2").value & " Planning Cycle)"
    xlWB.sheets(wsName).Range("A1").Select
    xlWB.sheets(wsName).Range("A1").Font.Bold = True
    
    ApplyFormat_Font_Generic xlWB.sheets(wsName).Range("A1"), 100
    
    'xlWB.Worksheets("RSMPivot").Select
    xlWB.worksheets(wsName).cells.Select
    xlWB.Application.Selection.Copy
    
    wsName = "RSMEdit"
    
    Worksheet_Action xlWB, actionCreate, wsName
    
    xlWB.worksheets(wsName).Select
    xlWB.Application.Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    xlWB.worksheets(wsName).Range("A1").Select
    
    Set ws = xlWB.worksheets(wsName)
    'ws.Name = wsName
    
    'ws.Range("A1:Z100000").value = ws.Range("A1:Z100000").value
    lastRowEdit = GetLastColumnInRSMEdit(xlWB)
    
    ws.cells.Select
    xlWB.Application.Selection.Style = "Comma"
    xlWB.Application.Selection.NumberFormat = "_(* #,##0.0_);_(* (#,##0.0);_(* ""-""??_);_(@_)"
    
    ws.Columns("D:D").Select
    xlWB.Application.Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
    
    xlWB.sheets(wsName).cells.Select
    xlWB.sheets(wsName).cells.EntireColumn.AutoFit
    
    xlWB.sheets(wsName).cells.Select
    xlWB.sheets(wsName).cells.EntireColumn.AutoFit
    xlWB.sheets(wsName).Range("G3").value = ""
    xlWB.sheets(wsName).Range("H3").value = ""
    
    'TODO:
    'Add this here...
    'PivotTable_AdjustMonths
    
    'Move Month over to Sales column
    xlWB.sheets(wsName).Range("I4").value = "'" & xlWB.sheets(wsName).Range("G4").value
    xlWB.sheets(wsName).Range("G4").value = ""
    xlWB.sheets(wsName).Range("L4").value = "'" & xlWB.sheets(wsName).Range("J4").value
    xlWB.sheets(wsName).Range("J4").value = ""
    xlWB.sheets(wsName).Range("O4").value = "'" & xlWB.sheets(wsName).Range("M4").value
    xlWB.sheets(wsName).Range("M4").value = ""
    'Anjali Code to extend 6 months
    xlWB.sheets(wsName).Range("R4").value = "'" & xlWB.sheets(wsName).Range("P4").value
    xlWB.sheets(wsName).Range("P4").value = ""
    xlWB.sheets(wsName).Range("U4").value = "'" & xlWB.sheets(wsName).Range("S4").value
    xlWB.sheets(wsName).Range("S4").value = ""
    xlWB.sheets(wsName).Range("X4").value = "'" & xlWB.sheets(wsName).Range("V4").value
    xlWB.sheets(wsName).Range("V4").value = ""
    
    'Hide Hist columns
    xlWB.sheets(wsName).Columns("E:F").Select
    xlWB.sheets(wsName).Columns("E:F").EntireColumn.Hidden = True
    'xlWB.Sheets(wsName).Columns("I:I").Select
    xlWB.sheets(wsName).Columns("J:J").EntireColumn.Hidden = True
    'xlWB.Sheets(wsName).Columns("L:L").Select
    xlWB.sheets(wsName).Columns("M:M").EntireColumn.Hidden = True
    'Anjali Code to extend 6 months
    xlWB.sheets(wsName).Columns("P:P").EntireColumn.Hidden = True
    xlWB.sheets(wsName).Columns("S:S").EntireColumn.Hidden = True
    xlWB.sheets(wsName).Columns("V:V").EntireColumn.Hidden = True
    
    
    xlWB.sheets(wsName).Range("A1").Select
    xlWB.sheets(wsName).Range("A1").FormulaR1C1 = "APPM - RSM View"
    xlWB.sheets(wsName).Range("A1").Select
    xlWB.sheets(wsName).Range("A1").FormulaR1C1 = "APPM S&OP Tool - RSM View (" & xlWB.worksheets("RSMData").Range("C2").value & " Planning Cycle)"
    xlWB.sheets(wsName).Range("A1").Select
    xlWB.sheets(wsName).Range("A1").Font.Bold = True
    With xlWB.sheets(wsName).Range("A1").Font
        .Name = "Calibri"
        .Size = 14
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    'Anjali Modified A:O TO A:X
    xlWB.sheets(wsName).Range("A2:X5").Select
    xlWB.sheets(wsName).Range(xlWB.sheets(wsName).Range("A2:X5"), xlWB.sheets(wsName).Range("A2:X5").End(xlDown)).Select
    xlWB.sheets(wsName).Range("A2:X" & lastRowEdit).Select
    xlWB.sheets(wsName).Range("A2:X" & lastRowEdit).Borders(xlDiagonalDown).LineStyle = xlNone
    xlWB.sheets(wsName).Range("A2:X" & lastRowEdit).Borders(xlDiagonalUp).LineStyle = xlNone
    With xlWB.sheets(wsName).Range("A2:X" & lastRowEdit).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With xlWB.sheets(wsName).Range("A2:X" & lastRowEdit).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With xlWB.sheets(wsName).Range("A2:X" & lastRowEdit).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With xlWB.sheets(wsName).Range("A2:X" & lastRowEdit).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With xlWB.sheets(wsName).Range("A2:X" & lastRowEdit).Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With xlWB.sheets(wsName).Range("A2:X" & lastRowEdit).Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    xlWB.sheets(wsName).Range("I4").Font.Bold = True
    xlWB.sheets(wsName).Range("L4").Font.Bold = True
    xlWB.sheets(wsName).Range("O4").Font.Bold = True
    'Anjali - Added column R, U & X
    xlWB.sheets(wsName).Range("R4").Font.Bold = True
    xlWB.sheets(wsName).Range("U4").Font.Bold = True
    xlWB.sheets(wsName).Range("X4").Font.Bold = True
    
    
    xlWB.sheets(wsName).Range("A5:X5").Font.Bold = True
    
    xlWB.sheets(wsName).Columns("G:X").EntireColumn.AutoFit

    xlWB.sheets(wsName).Select
    xlWB.sheets(wsName).Rows("6:6").Select
    xlWB.Windows(1).FreezePanes = True
    
    xlWB.sheets(wsName).Range("A1").Select

    xlWB.sheets(wsName).Range("A1").Select
    
    'Used to generate RSMEdit by Cat1
    'ApplyFormula_EditTotals_V01 xlWB, lastRowEdit
    
    xlWB.sheets("RSMEdit").Select
    xlWB.sheets("RSMEdit").Columns("A:A").Select
    xlWB.Application.Selection.ColumnWidth = 10.71
    xlWB.sheets("RSMEdit").Columns("B:B").Select
    xlWB.Application.Selection.ColumnWidth = 12.15
    xlWB.sheets("RSMEdit").Columns("C:C").Select
    xlWB.Application.Selection.ColumnWidth = 28
    xlWB.sheets("RSMEdit").Columns("D:D").Select
    xlWB.Application.Selection.ColumnWidth = 5.71
    xlWB.sheets("RSMEdit").Columns("G:X").Select
    xlWB.Application.Selection.ColumnWidth = 9
    
    'Used to generates RSMEdit by GSM
    ApplyFormula_EditTotals xlWB, lastRowEdit
    

    
    
    Set ws = xlWB.worksheets("RSMData")
    
    lastRowData = GetLastRow(xlWB, "RSMData", "A", 2)
    
    colBaseKey = Excel_sCol(GetColumnNo(ws, 1, "RSMBaseKey"))
    colMonth = Excel_sCol(GetColumnNo(ws, 1, "Month"))
    colSalesTons = Excel_sCol(GetColumnNo(ws, 1, "SalesTons"))
    colASFTons = Excel_sCol(GetColumnNo(ws, 1, "ASFTons"))
    
    'ws.Range("T2:T" & lastRowData).Formula = "=VLOOKUP(B2,RSMEdit!$F:$O,VLOOKUP(D2,MonthLookupTable,3,FALSE),FALSE)/VLOOKUP(B2,RSMEdit!$F:$O,VLOOKUP(D2,MonthLookupTable,2,FALSE),FALSE)*R2"
    ws.Range(colSalesTons & "2:" & colSalesTons & lastRowData).Formula = "=VLOOKUP(" & colBaseKey & "2,RSMEdit!$F:$X,VLOOKUP(" & colMonth & "2,MonthLookupTable,3,FALSE),FALSE)/VLOOKUP(" & colBaseKey & "2,RSMEdit!$F:$X,VLOOKUP(" & colMonth & "2,MonthLookupTable,2,FALSE),FALSE)*" & colASFTons & "2"
    
    'GenerateRSMPivot xlWB
    'Set xlWB = GeneratePivotTable_Generic(, 12)
    
    For Each ws In xlWB.worksheets
        DoEvents
        
        If ws.Name = "RSMEdit" Or ws.Name = "RSMPivot" Then
        Else
            ws.Visible = xlSheetHidden
        End If
    Next
    
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

Public Function GenerateRSMPivotCustom(ByRef xlWB As Object)
'**********************************************
'Author  :  RKP
'Date/Ver:  03-12-13/V01
'Input   :
'Output  :
'Comments:
'**********************************************
    On Error GoTo Err_Handler
    
    Dim month1      As String
    Dim month2      As String
    Dim month3      As String
    'Anjali - code added to include month 4,5 & 6
    Dim month4      As String
    Dim month5      As String
    Dim month6      As String
    
    xlWB.worksheets("RSMPivot").Activate
    xlWB.worksheets("RSMPivot").Select
    xlWB.worksheets("RSMPivot").UnProtect
    
    month1 = xlWB.worksheets("RSMSettings").Range("A2").value
    month2 = xlWB.worksheets("RSMSettings").Range("A3").value
    month3 = xlWB.worksheets("RSMSettings").Range("A4").value
    month4 = xlWB.worksheets("RSMSettings").Range("A5").value
    month5 = xlWB.worksheets("RSMSettings").Range("A6").value
    month6 = xlWB.worksheets("RSMSettings").Range("A7").value

    PivotTable_AdjustMonths xlWB.worksheets("RSMPivot"), 5, "Hist", "Sales", 4, 2
    
    xlWB.worksheets("RSMPivot").Select
    xlWB.worksheets("RSMPivot").Rows("3:4").Select
    xlWB.Application.Selection.EntireRow.Hidden = True
    
    xlWB.worksheets("RSMPivot").PivotTables(1).PivotSelect "'" & month1 & "'", xlDataAndLabel, True
    ApplyFormat_Color_Generic xlWB.Application.Selection, 10
    
    xlWB.worksheets("RSMPivot").PivotTables(1).PivotSelect "'" & month2 & "'", xlDataAndLabel, True
    ApplyFormat_Color_Generic xlWB.Application.Selection, 12
    
    xlWB.worksheets("RSMPivot").PivotTables(1).PivotSelect "'" & month3 & "'", xlDataAndLabel, True
    ApplyFormat_Color_Generic xlWB.Application.Selection, 14
    
    xlWB.worksheets("RSMPivot").PivotTables(1).PivotSelect "'" & month4 & "'", xlDataAndLabel, True
    ApplyFormat_Color_Generic xlWB.Application.Selection, 10
    
    xlWB.worksheets("RSMPivot").PivotTables(1).PivotSelect "'" & month5 & "'", xlDataAndLabel, True
    ApplyFormat_Color_Generic xlWB.Application.Selection, 12
    
    xlWB.worksheets("RSMPivot").PivotTables(1).PivotSelect "'" & month6 & "'", xlDataAndLabel, True
    ApplyFormat_Color_Generic xlWB.Application.Selection, 14
    
    
    
    
    
    xlWB.worksheets("RSMPivot").PivotTables(1).PivotSelect "'" & "Hist" & "'", xlDataAndLabel, True
    ApplyFormat_Color_Generic xlWB.Application.Selection, 16
    
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

Public Function Test()
'**********************************************
'Author  :  RKP
'Date/Ver:  03-12-13/V01
'Input   :
'Output  :
'Comments:
'**********************************************
    On Error GoTo Err_Handler
    
    Dim xlWB        As Object 'Workbook
    Dim timeStamp   As String

    'GeneratePivotTable_Generic , , , 105, xlWB
    'GeneratePivotTable_Generic , , , 105
    
    timeStamp = Application.CurrentProject.Connection.Execute("SELECT StageTimeStamp FROM tbl003PlanningCycleStages WHERE StageID = '" & ENUM_PC_STAGES.pcStage_RSMFilesGenerated & "'").Fields("StageTimeStamp").value & ""
    If VBA.IsDate(timeStamp) Then
        MsgBox "Is a date."
    Else
        MsgBox "Is not a date."
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

    Exit Function
    Resume
End Function

Public Function ApplyFormula_EditTotals(ByRef xlWB As Object, ByVal lastRowEdit As Integer)
'**********************************************
'Author  :  RKP
'Date/Ver:  04-02-13/V01
'Input   :
'Output  :
'Comments:
'**********************************************
    On Error GoTo Err_Handler

    Dim rowCtr              As Long
    Dim arrayTotals()       As TYPE_TOTALS
    Dim arrColA()           As TYPE_TOTALS
    Dim arrColB()           As TYPE_TOTALS
    Dim arrColC()           As TYPE_TOTALS
    Dim ws                  As Object 'Worksheet
    Dim grandTotalA         As String
    Dim grandTotalB         As String
    Dim grandTotalC         As String
    Dim gapCtr              As Long
    Dim startRow            As Long
    Dim colRSMRegion        As String
    Dim colBusiness         As String
    Dim colCat1             As String
    Dim colHist1            As String
    Dim colHist2            As String
    Dim colHist3            As String
    Dim colASF1             As String
    Dim colASF2             As String
    Dim colASF3             As String
    Dim colSales1           As String
    Dim colSales2           As String
    Dim colSales3           As String
    Dim colHist4            As String
    Dim colHist5            As String
    Dim colHist6            As String
    Dim colASF4             As String
    Dim colASF5             As String
    Dim colASF6             As String
    Dim colSales4           As String
    Dim colSales5           As String
    Dim colSales6           As String
    
    
    Set ws = xlWB.worksheets("RSMEdit")
    ReDim arrayTotals(0)
    ReDim arrColA(0)
    ReDim arrColB(0)
    ReDim arrColC(0)
    grandTotalA = ""
    grandTotalB = ""
    grandTotalC = ""
    gapCtr = 0
    startRow = 6
    
    colRSMRegion = "A"
    colBusiness = "B"
    colCat1 = "C"
    
'    colHist1 = "F"
'    colHist2 = "I"
'    colHist3 = "L"
'    colASF1 = "G"
'    colASF2 = "J"
'    colASF3 = "M"
'    colSales1 = "H"
'    colSales2 = "K"
'    colSales3 = "N"
    
    colHist1 = "G"
    colHist2 = "J"
    colHist3 = "M"
    colASF1 = "H"
    colASF2 = "K"
    colASF3 = "N"
    colSales1 = "I"
    colSales2 = "L"
    colSales3 = "O"
    'Anjali Code for months 4,5 & 6
    colHist4 = "P"
    colHist5 = "S"
    colHist6 = "V"
    colASF4 = "Q"
    colASF5 = "T"
    colASF6 = "W"
    colSales4 = "R"
    colSales5 = "U"
    colSales6 = "X"
    
    
    
    'Apply formulas to "Business" level
    For rowCtr = startRow To 100000
        DoEvents

        'Store the position where each "RSMRegion" starts and ends, in an array
        If ws.Range(colRSMRegion & rowCtr).value <> "" Then
            If VBA.InStr(1, ws.Range(colRSMRegion & rowCtr).value, "Total", vbTextCompare) > 0 Then
                If arrColA(0).endRow = 0 Then
                    arrColA(0).endRow = rowCtr
                Else
                    'ReDim Preserve arrayTotals(UBound(arrayTotals) + 1)
                    arrColA(UBound(arrColA)).endRow = rowCtr
                End If
            Else
                If arrColA(0).startRow = 0 Then
                    arrColA(0).startRow = rowCtr
                Else
                    ReDim Preserve arrColA(UBound(arrColA) + 1)
                    arrColA(UBound(arrColA)).startRow = rowCtr
                End If
            End If
        End If
        
        'Store the position where each "Business" starts and ends, in an array
        If ws.Range(colBusiness & rowCtr).value <> "" Then
            If VBA.InStr(1, ws.Range(colBusiness & rowCtr).value, "Total", vbTextCompare) > 0 Then
                If arrColB(0).endRow = 0 Then
                    arrColB(0).endRow = rowCtr
                Else
                    'ReDim Preserve arrayTotals(UBound(arrayTotals) + 1)
                    arrColB(UBound(arrColB)).endRow = rowCtr
                End If
            Else
                If arrColB(0).startRow = 0 Then
                    arrColB(0).startRow = rowCtr
                Else
                    ReDim Preserve arrColB(UBound(arrColB) + 1)
                    arrColB(UBound(arrColB)).startRow = rowCtr
                End If
            End If
        End If
        
        'Store the position where each "Cat1" starts and ends, in an array
        If ws.Range(colCat1 & rowCtr).value <> "" Then
            If VBA.InStr(1, ws.Range(colCat1 & rowCtr).value, "Total", vbTextCompare) > 0 Then
                If arrColC(0).endRow = 0 Then
                    arrColC(0).endRow = rowCtr
                Else
                    'ReDim Preserve arrayTotals(UBound(arrayTotals) + 1)
                    arrColC(UBound(arrColC)).endRow = rowCtr
                End If
            Else
                If arrColC(0).startRow = 0 Then
                    arrColC(0).startRow = rowCtr
                Else
                    ReDim Preserve arrColC(UBound(arrColC) + 1)
                    arrColC(UBound(arrColC)).startRow = rowCtr
                End If
            End If
        End If
        
        'Apply conditional formats, where "Sales" tons differs from "ASF" tons
        ApplyFormat_ConditionalFormat ws, rowCtr, colSales1, colASF1
        ApplyFormat_ConditionalFormat ws, rowCtr, colSales2, colASF2
        ApplyFormat_ConditionalFormat ws, rowCtr, colSales3, colASF3
        'Anjali Code
        ApplyFormat_ConditionalFormat ws, rowCtr, colSales4, colASF4
        ApplyFormat_ConditionalFormat ws, rowCtr, colSales5, colASF5
        ApplyFormat_ConditionalFormat ws, rowCtr, colSales6, colASF6
        
        
        
        
        If ws.Range(colHist1 & rowCtr).value = "" Then
            gapCtr = gapCtr + 1
        Else
            gapCtr = 0
        End If
        If gapCtr > 2 Then Exit For
    Next
    
    For rowCtr = 0 To UBound(arrColC)
        DoEvents
        
        'Add SUM to Total row
        ws.Range(colSales1 & arrColC(rowCtr).endRow).Formula = "=SUM(" & colSales1 & arrColC(rowCtr).startRow & ":" & colSales1 & arrColC(rowCtr).endRow - 1 & ")"
        ws.Range(colSales2 & arrColC(rowCtr).endRow).Formula = "=SUM(" & colSales2 & arrColC(rowCtr).startRow & ":" & colSales2 & arrColC(rowCtr).endRow - 1 & ")"
        ws.Range(colSales3 & arrColC(rowCtr).endRow).Formula = "=SUM(" & colSales3 & arrColC(rowCtr).startRow & ":" & colSales3 & arrColC(rowCtr).endRow - 1 & ")"
        ws.Range(colSales4 & arrColC(rowCtr).endRow).Formula = "=SUM(" & colSales4 & arrColC(rowCtr).startRow & ":" & colSales4 & arrColC(rowCtr).endRow - 1 & ")"
        ws.Range(colSales5 & arrColC(rowCtr).endRow).Formula = "=SUM(" & colSales5 & arrColC(rowCtr).startRow & ":" & colSales5 & arrColC(rowCtr).endRow - 1 & ")"
        ws.Range(colSales6 & arrColC(rowCtr).endRow).Formula = "=SUM(" & colSales6 & arrColC(rowCtr).startRow & ":" & colSales6 & arrColC(rowCtr).endRow - 1 & ")"
        
        
        'Color the "Sales" edit rows blue
        ApplyFormat_Color ws, arrColC(rowCtr).startRow, arrColC(rowCtr).endRow, colSales1, colSales1, 2
        ApplyFormat_Color ws, arrColC(rowCtr).startRow, arrColC(rowCtr).endRow, colSales2, colSales2, 2
        ApplyFormat_Color ws, arrColC(rowCtr).startRow, arrColC(rowCtr).endRow, colSales3, colSales3, 2
        ApplyFormat_Color ws, arrColC(rowCtr).startRow, arrColC(rowCtr).endRow, colSales4, colSales4, 2
        ApplyFormat_Color ws, arrColC(rowCtr).startRow, arrColC(rowCtr).endRow, colSales5, colSales5, 2
        ApplyFormat_Color ws, arrColC(rowCtr).startRow, arrColC(rowCtr).endRow, colSales6, colSales6, 2
        
        'Color "Hist" and "ASF" green
        ApplyFormat_Color ws, arrColC(rowCtr).startRow, arrColC(rowCtr).endRow, colHist1, colASF1, 3
        ApplyFormat_Color ws, arrColC(rowCtr).startRow, arrColC(rowCtr).endRow, colHist2, colASF2, 3
        ApplyFormat_Color ws, arrColC(rowCtr).startRow, arrColC(rowCtr).endRow, colHist3, colASF3, 3
        ApplyFormat_Color ws, arrColC(rowCtr).startRow, arrColC(rowCtr).endRow, colHist4, colASF4, 3
        ApplyFormat_Color ws, arrColC(rowCtr).startRow, arrColC(rowCtr).endRow, colHist5, colASF5, 3
        ApplyFormat_Color ws, arrColC(rowCtr).startRow, arrColC(rowCtr).endRow, colHist6, colASF6, 3
        
        'Color the "Sales" "Sum" cell dark grey
        ApplyFormat_Color ws, arrColC(rowCtr).endRow, arrColC(rowCtr).endRow, colCat1, colSales6, 1
        
        ws.Range(colSales1 & arrColC(rowCtr).startRow & ":" & colSales1 & arrColC(rowCtr).endRow - 1).Locked = False
        ws.Range(colSales2 & arrColC(rowCtr).startRow & ":" & colSales2 & arrColC(rowCtr).endRow - 1).Locked = False
        ws.Range(colSales3 & arrColC(rowCtr).startRow & ":" & colSales3 & arrColC(rowCtr).endRow - 1).Locked = False
        ws.Range(colSales4 & arrColC(rowCtr).startRow & ":" & colSales4 & arrColC(rowCtr).endRow - 1).Locked = False
        ws.Range(colSales5 & arrColC(rowCtr).startRow & ":" & colSales5 & arrColC(rowCtr).endRow - 1).Locked = False
        ws.Range(colSales6 & arrColC(rowCtr).startRow & ":" & colSales6 & arrColC(rowCtr).endRow - 1).Locked = False
        
        
    Next
    
    For rowCtr = 0 To UBound(arrColB)
        DoEvents
        
        'Add SUM to Total row
        ws.Range(colSales1 & arrColB(rowCtr).endRow).Formula = "=SUM(" & colSales1 & arrColB(rowCtr).startRow & ":" & colSales1 & arrColB(rowCtr).endRow - 1 & ")/2"
        ws.Range(colSales2 & arrColB(rowCtr).endRow).Formula = "=SUM(" & colSales2 & arrColB(rowCtr).startRow & ":" & colSales2 & arrColB(rowCtr).endRow - 1 & ")/2"
        ws.Range(colSales3 & arrColB(rowCtr).endRow).Formula = "=SUM(" & colSales3 & arrColB(rowCtr).startRow & ":" & colSales3 & arrColB(rowCtr).endRow - 1 & ")/2"
        ws.Range(colSales4 & arrColB(rowCtr).endRow).Formula = "=SUM(" & colSales4 & arrColB(rowCtr).startRow & ":" & colSales4 & arrColB(rowCtr).endRow - 1 & ")/2"
        ws.Range(colSales5 & arrColB(rowCtr).endRow).Formula = "=SUM(" & colSales5 & arrColB(rowCtr).startRow & ":" & colSales5 & arrColB(rowCtr).endRow - 1 & ")/2"
        ws.Range(colSales6 & arrColB(rowCtr).endRow).Formula = "=SUM(" & colSales6 & arrColB(rowCtr).startRow & ":" & colSales6 & arrColB(rowCtr).endRow - 1 & ")/2"
        
        
'        'Color the "Sales" edit rows blue
'        ApplyFormat_Color ws, arrColB(rowCtr).startRow, arrColB(rowCtr).endRow, colSales1, colSales1, 2
'        ApplyFormat_Color ws, arrColB(rowCtr).startRow, arrColB(rowCtr).endRow, colSales2, colSales2, 2
'        ApplyFormat_Color ws, arrColB(rowCtr).startRow, arrColB(rowCtr).endRow, colSales3, colSales3, 2
'        'Color "Hist" and "ASF" green
'        ApplyFormat_Color ws, arrColB(rowCtr).startRow, arrColB(rowCtr).endRow, colHist1, colASF1, 3
'        ApplyFormat_Color ws, arrColB(rowCtr).startRow, arrColB(rowCtr).endRow, colHist2, colASF2, 3
'        ApplyFormat_Color ws, arrColB(rowCtr).startRow, arrColB(rowCtr).endRow, colHist3, colASF3, 3
        
        'Color the "Sales" "Sum" cell dark grey
        ApplyFormat_Color ws, arrColB(rowCtr).endRow, arrColB(rowCtr).endRow, colBusiness, colSales6, 1
        
'        ws.Range(colSales1 & arrColB(rowCtr).startRow & ":" & colSales1 & arrColB(rowCtr).endRow - 1).Locked = False
'        ws.Range(colSales2 & arrColB(rowCtr).startRow & ":" & colSales2 & arrColB(rowCtr).endRow - 1).Locked = False
'        ws.Range(colSales3 & arrColB(rowCtr).startRow & ":" & colSales3 & arrColB(rowCtr).endRow - 1).Locked = False
    Next
    
    For rowCtr = 0 To UBound(arrColA)
        DoEvents
        
        'Add SUM to Total row
        ws.Range(colSales1 & arrColA(rowCtr).endRow).Formula = "=SUM(" & colSales1 & arrColA(rowCtr).startRow & ":" & colSales1 & arrColA(rowCtr).endRow - 1 & ")/3"
        ws.Range(colSales2 & arrColA(rowCtr).endRow).Formula = "=SUM(" & colSales2 & arrColA(rowCtr).startRow & ":" & colSales2 & arrColA(rowCtr).endRow - 1 & ")/3"
        ws.Range(colSales3 & arrColA(rowCtr).endRow).Formula = "=SUM(" & colSales3 & arrColA(rowCtr).startRow & ":" & colSales3 & arrColA(rowCtr).endRow - 1 & ")/3"
        ws.Range(colSales4 & arrColA(rowCtr).endRow).Formula = "=SUM(" & colSales4 & arrColA(rowCtr).startRow & ":" & colSales4 & arrColA(rowCtr).endRow - 1 & ")/3"
        ws.Range(colSales5 & arrColA(rowCtr).endRow).Formula = "=SUM(" & colSales5 & arrColA(rowCtr).startRow & ":" & colSales5 & arrColA(rowCtr).endRow - 1 & ")/3"
        ws.Range(colSales6 & arrColA(rowCtr).endRow).Formula = "=SUM(" & colSales6 & arrColA(rowCtr).startRow & ":" & colSales6 & arrColA(rowCtr).endRow - 1 & ")/3"
'        'Color the "Sales" edit rows blue
'        ApplyFormat_Color ws, arrColA(rowCtr).startRow, arrColA(rowCtr).endRow, colSales1, colSales1, 2
'        ApplyFormat_Color ws, arrColA(rowCtr).startRow, arrColA(rowCtr).endRow, colSales2, colSales2, 2
'        ApplyFormat_Color ws, arrColA(rowCtr).startRow, arrColA(rowCtr).endRow, colSales3, colSales3, 2
'        'Color "Hist" and "ASF" green
'        ApplyFormat_Color ws, arrColA(rowCtr).startRow, arrColA(rowCtr).endRow, colHist1, colASF1, 3
'        ApplyFormat_Color ws, arrColA(rowCtr).startRow, arrColA(rowCtr).endRow, colHist2, colASF2, 3
'        ApplyFormat_Color ws, arrColA(rowCtr).startRow, arrColA(rowCtr).endRow, colHist3, colASF3, 3
        
        'Color the "Sales" "Sum" cell dark grey
        ApplyFormat_Color ws, arrColA(rowCtr).endRow, arrColA(rowCtr).endRow, colRSMRegion, colSales6, 1
        
'        ws.Range(colSales1 & arrColA(rowCtr).startRow & ":" & colSales1 & arrColA(rowCtr).endRow - 1).Locked = False
'        ws.Range(colSales2 & arrColA(rowCtr).startRow & ":" & colSales2 & arrColA(rowCtr).endRow - 1).Locked = False
'        ws.Range(colSales3 & arrColA(rowCtr).startRow & ":" & colSales3 & arrColA(rowCtr).endRow - 1).Locked = False
    Next
    
''    ws.Range(colHist1 & lastRowEdit).Formula = "=SUM(" & colHist1 & startRow & ":" & colHist1 & lastRowEdit - 1 & ")/3"
''    ws.Range(colASF1 & lastRowEdit).Formula = "=SUM(" & colASF1 & startRow & ":" & colASF1 & lastRowEdit - 1 & ")/3"
''    ws.Range(colSales1 & lastRowEdit).Formula = "=SUM(" & colSales1 & startRow & ":" & colSales1 & lastRowEdit - 1 & ")/3"
''    ws.Range(colASF2 & lastRowEdit).Formula = "=SUM(" & colASF2 & startRow & ":" & colASF2 & lastRowEdit - 1 & ")/3"
''    ws.Range(colSales2 & lastRowEdit).Formula = "=SUM(" & colSales2 & startRow & ":" & colSales2 & lastRowEdit - 1 & ")/3"
''    ws.Range(colASF3 & lastRowEdit).Formula = "=SUM(" & colASF3 & startRow & ":" & colASF3 & lastRowEdit - 1 & ")/3"
''    ws.Range(colSales3 & lastRowEdit).Formula = "=SUM(" & colSales3 & startRow & ":" & colSales3 & lastRowEdit - 1 & ")/3"
''
''GoTo Skip_GrandTotal
    
    'Apply formats to Grand Total row
    ws.Range(colRSMRegion & lastRowEdit + 1).value = "GRAND TOTAL"
    ApplyFormat_Color ws, lastRowEdit + 1, lastRowEdit + 1, colRSMRegion, colSales6, 1
    ws.Range(colHist1 & lastRowEdit + 1).Formula = "=SUM(" & colHist1 & startRow & ":" & colHist1 & lastRowEdit & ")/4"
    ws.Range(colASF1 & lastRowEdit + 1).Formula = "=SUM(" & colASF1 & startRow & ":" & colASF1 & lastRowEdit & ")/4"
    ws.Range(colSales1 & lastRowEdit + 1).Formula = "=SUM(" & colSales1 & startRow & ":" & colSales1 & lastRowEdit & ")/4"
    ws.Range(colHist2 & lastRowEdit + 1).Formula = "=SUM(" & colHist2 & startRow & ":" & colHist2 & lastRowEdit & ")/4"
    ws.Range(colASF2 & lastRowEdit + 1).Formula = "=SUM(" & colASF2 & startRow & ":" & colASF2 & lastRowEdit & ")/4"
    ws.Range(colSales2 & lastRowEdit + 1).Formula = "=SUM(" & colSales2 & startRow & ":" & colSales2 & lastRowEdit & ")/4"
    ws.Range(colHist3 & lastRowEdit + 1).Formula = "=SUM(" & colHist3 & startRow & ":" & colHist3 & lastRowEdit & ")/4"
    ws.Range(colASF3 & lastRowEdit + 1).Formula = "=SUM(" & colASF3 & startRow & ":" & colASF3 & lastRowEdit & ")/4"
    ws.Range(colSales3 & lastRowEdit + 1).Formula = "=SUM(" & colSales3 & startRow & ":" & colSales3 & lastRowEdit & ")/4"
    
    ws.Range(colHist4 & lastRowEdit + 1).Formula = "=SUM(" & colHist4 & startRow & ":" & colHist4 & lastRowEdit & ")/4"
    ws.Range(colASF4 & lastRowEdit + 1).Formula = "=SUM(" & colASF4 & startRow & ":" & colASF4 & lastRowEdit & ")/4"
    ws.Range(colSales4 & lastRowEdit + 1).Formula = "=SUM(" & colSales4 & startRow & ":" & colSales4 & lastRowEdit & ")/4"
    ws.Range(colHist5 & lastRowEdit + 1).Formula = "=SUM(" & colHist5 & startRow & ":" & colHist5 & lastRowEdit & ")/4"
    ws.Range(colASF5 & lastRowEdit + 1).Formula = "=SUM(" & colASF5 & startRow & ":" & colASF5 & lastRowEdit & ")/4"
    ws.Range(colSales5 & lastRowEdit + 1).Formula = "=SUM(" & colSales5 & startRow & ":" & colSales5 & lastRowEdit & ")/4"
    ws.Range(colHist6 & lastRowEdit + 1).Formula = "=SUM(" & colHist6 & startRow & ":" & colHist6 & lastRowEdit & ")/4"
    ws.Range(colASF6 & lastRowEdit + 1).Formula = "=SUM(" & colASF6 & startRow & ":" & colASF6 & lastRowEdit & ")/4"
    ws.Range(colSales6 & lastRowEdit + 1).Formula = "=SUM(" & colSales6 & startRow & ":" & colSales6 & lastRowEdit & ")/4"
    
    
    
    'Apply formats to "GRAND TOTAL" row
    ws.Range(colRSMRegion & lastRowEdit & ":" & colSales6 & lastRowEdit).Select
    xlWB.Application.Selection.Copy
    ws.Range(colRSMRegion & lastRowEdit + 1).Select
    xlWB.Application.Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    xlWB.Application.CutCopyMode = False
    ws.Range(colSales6 & lastRowEdit + 1).Select
    
Skip_GrandTotal:

    ws.Columns("G:O").EntireColumn.AutoFit
    ws.Select
    ws.Columns("M:M").Select
    xlWB.Application.Selection.EntireColumn.Hidden = True
    ws.Columns("J:J").Select
    xlWB.Application.Selection.EntireColumn.Hidden = True
    ws.Columns("P:P").Select
    xlWB.Application.Selection.EntireColumn.Hidden = True
    ws.Columns("S:S").Select
    xlWB.Application.Selection.EntireColumn.Hidden = True
    ws.Columns("V:V").Select
    xlWB.Application.Selection.EntireColumn.Hidden = True
    
    ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True

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

Public Function ApplyFormat_Color(ByRef xlWS As Object, ByVal rowStart As Long, ByVal rowEnd As Long, ByVal colStart As String, ByVal colEnd As String, ByVal colorType As Integer)
'**********************************************
'Author  :  RKP
'Date/Ver:  04-02-13/V01
'Input   :
'Output  :
'Comments:
'**********************************************
    On Error GoTo Err_Handler
    
    If colorType = 1 Then
        With xlWS.Range(colStart & rowStart & ":" & colEnd & rowEnd).Interior
'            .Pattern = xlSolid
'            .PatternColorIndex = xlAutomatic
'            .ThemeColor = xlThemeColorDark1
'            .TintAndShade = -0.499984740745262
'            .PatternTintAndShade = 0
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = -0.249977111117893 'light
            '.TintAndShade = -0.149998474074526 'darker
            .PatternTintAndShade = 0
        End With
'        xlWS.Range(colStart & rowStart & ":" & colEnd & rowEnd).Font.Bold = True
    ElseIf colorType = 2 Then
        With xlWS.Range(colStart & rowStart & ":" & colEnd & rowEnd).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorLight2 'xlThemeColorAccent5
            .TintAndShade = 0.599993896298105 '0.399975585192419
            .PatternTintAndShade = 0
        End With
    ElseIf colorType = 3 Then
        With xlWS.Range(colStart & rowStart & ":" & colEnd & rowEnd).Interior
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent3
            .TintAndShade = 0.399975585192419
            .PatternTintAndShade = 0
        End With
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

    Exit Function
    Resume
End Function

Public Function ApplyFormat_ConditionalFormat(ByRef ws As Object, ByVal rowCtr As Long, ByVal srcColLetter As String, ByVal destColLetter As String)
'**********************************************
'Author  :  RKP
'Date/Ver:  04-02-13/V01
'Input   :
'Output  :
'Comments:
'**********************************************
    On Error GoTo Err_Handler

    'Apply conditional formats, where "Sales" tons differs from "ASF" tons
    ws.Range(srcColLetter & rowCtr).Select
    ws.Application.Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlNotEqual, Formula1:="=" & destColLetter & rowCtr
    ws.Application.Selection.FormatConditions(ws.Application.Selection.FormatConditions.count).SetFirstPriority
    With ws.Application.Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 49407
        .TintAndShade = 0
    End With
    ws.Application.Selection.FormatConditions(1).StopIfTrue = False

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

Public Function ApplyFormula_Subtotals(ByRef xlWS As Object, ByVal startRow As Long, ByVal srcColLetter As String, ByVal destColLetter As String, ByVal isContinuous As Boolean)
'**********************************************
'Author  :  RKP
'Date/Ver:  04-02-13/V01
'Input   :
'Output  :
'Comments:
'**********************************************
    On Error GoTo Err_Handler

    Dim ctr         As Long
    Dim ctrBlank    As Long
    Dim strFormula  As String
    
    strFormula = ""
    For ctr = startRow To 100000
        DoEvents
        
        If VBA.InStr(1, xlWS.Range(destColLetter & ctr).value, "Total", vbTextCompare) > 0 Then
            If strFormula <> "" Then
                xlWS.Range(srcColLetter & ctr).Formula = "=SUM(" & strFormula & ")"
                strFormula = ""
            End If
        End If
        If VBA.InStr(1, xlWS.Range(srcColLetter & ctr).value, "Total", vbTextCompare) > 0 Then
            strFormula = strFormula & "+" & srcColLetter & ctr
        End If
        
        If xlWS.Range(srcColLetter & ctr).value = "" Then
            ctrBlank = ctrBlank + 1
        End If
        If ctrBlank > 5 Then Exit For
    Next

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

Public Function Import_SalesHistory()
'**********************************************
'Author  :  RKP
'Date/Ver:  05-01-13/V01
'Input   :
'Output  :
'Comments:
'Example:
'\\sa53fs01\Secunderabad-A53\APPM S&OP\RawFiles
'\\sa53fs01\Secunderabad-A53\APPM S&OP\PlanningCycle_201304\RawFiles\APPM_SalesHistory_RAJ_PCMinus3.txt
'filePath = FormatFilePath(basePath) & "PlanningCycle_" & planningCycle & "\RawFiles\" & "APPM_SalesHistory_RAJ_PCMinus" & ctr & ".txt"
'filePath = FormatFilePath(basePath) & "PlanningCycle_" & planningCycle & "\RawFiles\" & "APPM_SalesHistory_KAD_PCMinus" & ctr & ".txt"
'**********************************************
    On Error GoTo Err_Handler

    Dim linkSpecName        As String
    Dim basePath            As String
    Dim filePath            As String
    Dim planningCycle       As String
    Dim tableName           As String
    Dim ctr                 As Integer
    Dim sql                 As String
    Dim ret                 As Long
    Dim RS                  As ADODB.Recordset
    
    linkSpecName = basUtility.GetSetting("SalesHistoryImportSpecName")
    basePath = basUtility.GetSetting("BaseFilePath")
    planningCycle = basUtility.GetScalarValue("SELECT PlanningCycle FROM tbl001PlanningCycle")
    tableName = "tbl002RawData010SalesHistory"
    
    sql = "DELETE * FROM [" & tableName & "]"
    ExecSQL sql, ret
    
    
    
   '*&*#  FIX so all saleshist comes in in one single file
   filePath = FormatFilePath(basUtility.GetSetting("SalesHistoryFolder")) & "\" & "APPM_PC_SALESHIST" & ".txt"
   If VBA.Dir(filePath) <> "" Then ImportCSVFilesUsingImportSpec linkSpecName, filePath, tableName
    
    
    
'''''''    'Anjali 02/15/2016 commented first line to include E01 February data pull in the Month of Feb, should be uncommented for March2016 PC
'''''''    For ctr = 3 To 1 Step -1
'''''''    ''''OK IN MARCH, STM, CHANGE ONLY PULL 3 MOS NOT FOUR,
'''''''    ''''For ctr = 3 To 0 Step -1 'Tobedeleted PC 201603 ****IMP**** uncomment the one above this
'''''''        DoEvents
'''''''        'Anjali 02/15/2016 added to include E01 data pull
'''''''        'filePath = FormatFilePath(basePath) & "PlanningCycle_" & planningCycle & "\RawFiles\" & "APPM_SalesHistory_RAJ_PCMinus" & ctr & ".txt"
'''''''
'''''''        '*&*#  FIX THIS
'''''''        filePath = FormatFilePath(basUtility.GetSetting("SalesHistoryFolder")) & "APPM_PC" & planningCycle & "_IN03_" & FormatDate(VBA.DateAdd("m", -1 * ctr, VBA.Now()), 1) & ".txt"
'''''''        Debug.Print filePath
'''''''
'''''''
'''''''        If VBA.Dir(filePath) = "" Then
'''''''            'error
'''''''        Else
'''''''            ImportCSVFilesUsingImportSpec linkSpecName, filePath, tableName
'''''''        End If
'''''''
'''''''
'''''''        '''' DISABLE APP INSTANCE SCRIPT MAY 2016  ''''   'filePath = FormatFilePath(basePath) & "PlanningCycle_" & planningCycle & "\RawFiles\" & "APPM_SalesHistory_RAJ_PCMinus" & ctr & ".txt"
'''''''        '''' DISABLE APP INSTANCE SCRIPT MAY 2016  ''''   filePath = FormatFilePath(basUtility.GetSetting("SalesHistoryFolder")) & "APPM_PC" & planningCycle & "_1000_" & FormatDate(VBA.DateAdd("m", -1 * ctr, VBA.Now()), 1) & ".txt"
'''''''        '''' DISABLE APP INSTANCE SCRIPT MAY 2016  ''''   Debug.Print filePath
'''''''        '''' DISABLE APP INSTANCE SCRIPT MAY 2016  ''''   If VBA.Dir(filePath) = "" Then
'''''''        '''' DISABLE APP INSTANCE SCRIPT MAY 2016  ''''       'error
'''''''        '''' DISABLE APP INSTANCE SCRIPT MAY 2016  ''''   Else
'''''''        '''' DISABLE APP INSTANCE SCRIPT MAY 2016  ''''       ImportCSVFilesUsingImportSpec linkSpecName, filePath, tableName
'''''''        '''' DISABLE APP INSTANCE SCRIPT MAY 2016  ''''   End If
'''''''
'''''''        '''' DISABLE APP INSTANCE SCRIPT MAY 2016  ''''   'filePath = FormatFilePath(basePath) & "PlanningCycle_" & planningCycle & "\RawFiles\" & "APPM_SalesHistory_KAD_PCMinus" & ctr & ".txt"
'''''''        '''' DISABLE APP INSTANCE SCRIPT MAY 2016  ''''   filePath = FormatFilePath(basUtility.GetSetting("SalesHistoryFolder")) & "APPM_PC" & planningCycle & "_2000_" & FormatDate(VBA.DateAdd("m", -1 * ctr, VBA.Now()), 1) & ".txt"
'''''''        '''' DISABLE APP INSTANCE SCRIPT MAY 2016  ''''   Debug.Print filePath
'''''''        '''' DISABLE APP INSTANCE SCRIPT MAY 2016  ''''   If VBA.Dir(filePath) = "" Then
'''''''        '''' DISABLE APP INSTANCE SCRIPT MAY 2016  ''''       'error
'''''''        '''' DISABLE APP INSTANCE SCRIPT MAY 2016  ''''   Else
'''''''        '''' DISABLE APP INSTANCE SCRIPT MAY 2016  ''''       ImportCSVFilesUsingImportSpec linkSpecName, filePath, tableName
'''''''        '''' DISABLE APP INSTANCE SCRIPT MAY 2016  ''''   End If
'''''''    Next
    
    tableName = "tbl002RawData015Customer"
    sql = "DELETE * FROM [" & tableName & "]"
    ExecSQL sql, ret
    
    '''' DISABLE APP INSTANCE SCRIPT MAY 2016  ''''   filePath = FormatFilePath(basUtility.GetSetting("SalesHistoryFolder")) & "APPM_KNA1" & ".txt"
    '''' DISABLE APP INSTANCE SCRIPT MAY 2016  ''''   Debug.Print filePath
    '''' DISABLE APP INSTANCE SCRIPT MAY 2016  ''''   linkSpecName = basUtility.GetSetting("KNA1ImportSpecName")
    '''' DISABLE APP INSTANCE SCRIPT MAY 2016  ''''   ImportCSVFilesUsingImportSpec linkSpecName, filePath, tableName
    
    'Anjali New code added to import E01 CustMaster
    filePath = FormatFilePath(basUtility.GetSetting("SalesHistoryFolder")) & "APPM_KNA1_E01" & ".txt"
    Debug.Print filePath
    linkSpecName = basUtility.GetSetting("KNA1ImportSpecName")
    ImportCSVFilesUsingImportSpec linkSpecName, filePath, tableName
    
    'Anjali 02/15/2016 Added code to import MARA E01, new spec in settings table "imexLinkSpec_MARAE01"
    tableName = "tbl002RawData015Material"
    sql = "DELETE * FROM [" & tableName & "]"
    ExecSQL sql, ret
    filePath = FormatFilePath(basUtility.GetSetting("SalesHistoryFolder")) & "APPM_MARA_E01" & ".txt"
    Debug.Print filePath
    linkSpecName = basUtility.GetSetting("MARAE01ImportSpecName")
    ImportCSVFilesUsingImportSpec linkSpecName, filePath, tableName
          
    basUtility.RunQueriesByPrefix "q010ImportData"
    
    'MsgBox "Import finished."

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
    DoCmd.SetWarnings True

    Exit Function
    Resume
End Function

'''Public Function FormatFilePath(ByVal filePath As String) As String
''''**********************************************
''''Author  :  RKP
''''Date/Ver:  05-01-13/V01
''''Input   :
''''Output  :
''''Comments:
''''**********************************************
'''    On Error GoTo Err_Handler
'''
'''    FormatFilePath = VBA.IIf(VBA.Right(VBA.Trim(filePath), 1) = "\", VBA.Trim(filePath), VBA.Trim(filePath) & "\")
'''
'''Err_Handler:
'''    mlLastErr = Err.Number
'''    msLastErr = Err.Description
'''    'Function1 = mlLastErr
'''    If Err Then
'''        If Err.Number = 49 Then 'Bad DLL calling convention
'''            mlLastErr = 0
'''            msLastErr = ""
'''            Resume Next
'''        Else
'''            'ProcessMsg Err.Number, Err.Description, "", ""
'''            MsgBox Err.Number & " - " & Err.Description
'''        End If
'''    End If
'''
'''    Exit Function
'''    Resume
'''End Function

Public Function ExecSQL(ByVal sql As String, ByRef ret As Long) As ADODB.Recordset
'**********************************************
'Author  :  RKP
'Date/Ver:  05-01-13/V01
'Input   :
'Output  :
'Comments:
'**********************************************
    On Error GoTo Err_Handler
    
    Set ExecSQL = Application.CurrentProject.Connection.Execute(sql, ret)

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

Public Function ImportCSVFilesUsingImportSpec _
( _
   ByVal importSpec As String, _
   ByVal FileName As String, _
   ByVal tableName As String _
)
'**********************************************
'Author  :  RKP
'Date/Ver:  08-23-2010/V01
'Input   :
'Output  :
'Comments:
'**********************************************
   On Error GoTo Err_Handler
   
'   Dim fileName      As String
   
'   fileName = "Z:\Work\BMOS\SAPScripts\Output\ZCOR1347_0205PM15.csv"

   'Call Application.DoCmd.TransferText(acImportDelim, "ZCOR1347_RAW_IMPORT_SPEC", "ZCOR1347_RAW", fileName, True)
   
   Call Application.DoCmd.TransferText(acImportDelim, importSpec, tableName, FileName, True)

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

'''Public Function FormatDate(ByVal inputDate As Date, ByVal id As Integer) As String
''''**********************************************
''''Author  :  RKP
''''Date/Ver:  05-01-13/V01
''''Input   :
''''Output  :
''''Comments:
''''**********************************************
'''    On Error GoTo Err_Handler
'''
'''    Select Case id
'''        Case 1
'''            FormatDate = VBA.Year(inputDate) & VBA.Format(VBA.Month(inputDate), "00")
'''    End Select
'''
'''Err_Handler:
'''    mlLastErr = Err.Number
'''    msLastErr = Err.Description
'''    'Function1 = mlLastErr
'''    If Err Then
'''        If Err.Number = 49 Then 'Bad DLL calling convention
'''            mlLastErr = 0
'''            msLastErr = ""
'''            Resume Next
'''        Else
'''            'ProcessMsg Err.Number, Err.Description, "", ""
'''            MsgBox Err.Number & " - " & Err.Description
'''        End If
'''    End If
'''
'''    Exit Function
'''    Resume
'''End Function

Public Function GetPlanningCycleStageDate(ByVal pcStageID As ENUM_PC_STAGES) As Date
'**********************************************
'Author  :  RKP
'Date/Ver:  03-12-13/V01
'Input   :
'Output  :
'Comments:
'**********************************************
    On Error GoTo Err_Handler
    
    Dim timeStampStr        As String
    
    timeStampStr = Application.CurrentProject.Connection.Execute("SELECT StageTimeStamp FROM tbl003PlanningCycleStages WHERE StageID = '" & pcStageID & "'").Fields("StageTimeStamp").value & ""
    
    GetPlanningCycleStageDate = VBA.CDate(NULL_DATE)
    If VBA.IsDate(timeStampStr) Then
        GetPlanningCycleStageDate = VBA.CDate(timeStampStr)
    End If
    
'    Select Case pcStageID
'        Case ENUM_PC_STAGES.pcStage_PlanningCycleChanged
'            timeStampStr = Application.CurrentProject.Connection.Execute("SELECT StageTimeStamp FROM tbl003PlanningCycleStages WHERE StageID = '" & pcStageID & "'").Fields("StageTimeStamp").value & ""
'        Case ENUM_PC_STAGES.pcStage_RunSAPScripts
'        Case ENUM_PC_STAGES.pcStage_ImportNewData
'        Case ENUM_PC_STAGES.pcStage_InitializeData
'        Case ENUM_PC_STAGES.pcStage_RSMFilesGenerated
'        Case ENUM_PC_STAGES.pcStage_RSMFilesImported
'        Case Else
'
'    End Select

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

Public Function SetPlanningCycleStageDate(ByVal pcStageID As ENUM_PC_STAGES, Optional value As String) As Date
'**********************************************
'Author  :  RKP
'Date/Ver:  03-12-13/V01
'Input   :
'Output  :
'Comments:
'**********************************************
    On Error GoTo Err_Handler
    
    Call Application.CurrentProject.Connection.Execute("UPDATE tbl003PlanningCycleStages SET StageTimeStamp = #" & VBA.Now() & "# WHERE StageID = '" & pcStageID & "'")
    Call Application.CurrentProject.Connection.Execute("UPDATE tbl003PlanningCycleStages SET StageCounter = StageCounter + 1 WHERE StageID = '" & pcStageID & "'")
    
    Call Application.CurrentProject.Connection.Execute("UPDATE tbl003PlanningCycleStages SET StageTimeStamp = #" & VBA.Now() & "# WHERE StageID = '" & ENUM_PC_STAGES.pcStage_Lock & "'")
    Call Application.CurrentProject.Connection.Execute("UPDATE tbl003PlanningCycleStages SET StageCounter = StageCounter + 1 WHERE StageID = '" & ENUM_PC_STAGES.pcStage_Lock & "'")
    
'    If pcStageID = pcStage_Lock Then
'        Call Application.CurrentProject.Connection.Execute("UPDATE tbl003PlanningCycleStages SET StageCounter = " & value & " WHERE StageID = '" & pcStageID & "'")
'    Else
    
    
'    End If
    
    Forms.Item("MainPage").chkPlanningCycleStagesLock.SetFocus
    Forms.Item("MainPage").chkPlanningCycleStagesLock.value = True
    
'    Select Case pcStageID
'        Case ENUM_PC_STAGES.pcStage_PlanningCycleChanged
'            timeStampStr = Application.CurrentProject.Connection.Execute("SELECT StageTimeStamp FROM tbl003PlanningCycleStages WHERE StageID = '" & pcStageID & "'").Fields("StageTimeStamp").value & ""
'        Case ENUM_PC_STAGES.pcStage_RunSAPScripts
'        Case ENUM_PC_STAGES.pcStage_ImportNewData
'        Case ENUM_PC_STAGES.pcStage_InitializeData
'        Case ENUM_PC_STAGES.pcStage_RSMFilesGenerated
'        Case ENUM_PC_STAGES.pcStage_RSMFilesImported
'        Case Else
'
'    End Select

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
            'MsgBox Err.Number & " - " & Err.Description
            Resume Next
        End If
    End If

    Exit Function
    Resume
End Function

Public Function CanRunPlanningCycleStage(ByVal pcStageID As ENUM_PC_STAGES, ByVal pcStageDesc As String) As Boolean
'**********************************************
'Author  :  RKP
'Date/Ver:  03-12-13/V01
'Input   :
'Output  :
'Comments:
'**********************************************
    On Error GoTo Err_Handler
    
    Dim timeStamp   As Date
    
    CanRunPlanningCycleStage = False
    timeStamp = modMain.GetPlanningCycleStageDate(pcStage_RSMFilesGenerated)
    If timeStamp = NULL_DATE Then
        CanRunPlanningCycleStage = True
    Else
        'If Forms.Item("MainPage").chkPlanningCycleStagesLock.value = True Then
        If modMain.GetPlanningCycleStageDate(pcStage_Lock) = NULL_DATE Then
            CanRunPlanningCycleStage = True
        Else
            MsgBox """" & pcStageDesc & """ is disallowed because RSM files have already been generated." & vbNewLine & vbNewLine & "If you still want to proceed, you must unlock the:" & vbNewLine & """Planning Cycle Stages Lock""" & vbNewLine & "on the ""Main Page"", and try the operation again.", vbCritical, Application.CurrentProject.Name
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

    Exit Function
    Resume
End Function

Public Function IsValid(ByVal pcStageID As ENUM_PC_STAGES) As Boolean
'**********************************************
'Author  :  RKP
'Date/Ver:  06-24-13/V01
'Input   :
'Output  :
'Comments:
'**********************************************
    On Error GoTo Err_Handler
    
    Dim timeStamp   As Date
    
    IsValid = False
    Select Case pcStageID
        Case ENUM_PC_STAGES.pcStage_PlanningCycleChanged
            'timeStampStr = Application.CurrentProject.Connection.Execute("SELECT StageTimeStamp FROM tbl003PlanningCycleStages WHERE StageID = '" & pcStageID & "'").Fields("StageTimeStamp").value & ""
            
        Case ENUM_PC_STAGES.pcStage_RunSAPScripts
        Case ENUM_PC_STAGES.pcStage_ImportNewData
        Case ENUM_PC_STAGES.pcStage_InitializeData
        Case ENUM_PC_STAGES.pcStage_RSMFilesGenerated
        Case ENUM_PC_STAGES.pcStage_RSMFilesImported
        Case Else

    End Select

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

Public Function SAP_RunSAPScripts()
'**********************************************
'Author  :  RKP
'Date/Ver:  08-20-13/V01
'Input   :
'Output  :
'Comments:  Anjali - on Feb 15, 2016 Added code to include E01 extract
'**********************************************
    On Error GoTo Err_Handler

    Dim session         As Object
    Dim APPsession      As Object
    Dim sapInstance     As String
    Dim APPsapInstance  As String
    '''Dim tMinus0Start    As String 'Tobedeleted PC 201603
    '''Dim tMinus0End      As String 'Tobedeleted PC 201603
    Dim tMinus1Start    As String
    Dim tMinus1End      As String
    Dim tMinus2Start    As String
    Dim tMinus2End      As String
    Dim tMinus3Start    As String
    Dim tMinus3End      As String
    
    Dim companyCode_RAJ As String
    Dim companyCode_KAD As String
    Dim companyCode_IN03 As String
    Dim layoutName      As String
    Dim userGroupName   As String
    Dim sapQueryName    As String
    Dim folderPath      As String
    Dim FileName        As String
    Dim filePath        As String
    Dim planningCycle   As String
    Dim msg             As String
    Dim startTime       As Date
    Dim RS              As DAO.Recordset
    Dim sSQL            As String
    Dim sMachMatlCode   As String
    Dim sGenOrAppend    As String
    
    startTime = VBA.Now()
    
    '''' DISABLE APP INSTANCE SCRIPT MAY 2016  ''''   APPsapInstance = "APP"
    sapInstance = VBA.CStr(basUtility.GetSetting("SAPInstance"))
    
    '''' DISABLE APP INSTANCE SCRIPT MAY 2016  ''''   Set APPsession = modMain.SAP_GetSession(APPsapInstance, False)
    Set session = modMain.SAP_GetSession(sapInstance, False)
    If session Is Nothing Then    '''' DISABLE APP INSTANCE SCRIPT MAY 2016  '''' Or APPsession Is Nothing
        If session Is Nothing Then
            MsgBox "No active SAP session found for """ & sapInstance & """." & vbNewLine & "Please log on to """ & sapInstance & """ and try the operation again.", vbExclamation, Application.CurrentProject.Name
        '''' DISABLE APP INSTANCE SCRIPT MAY 2016  ''''   ElseIf APPsession Is Nothing Then
            '''' DISABLE APP INSTANCE SCRIPT MAY 2016  ''''   MsgBox "No active SAP session found for """ & APPsapInstance & """." & vbNewLine & "Please log on to """ & sapInstance & """ and try the operation again.", vbExclamation, Application.CurrentProject.Name
        End If
    Else
        layoutName = "/SOPSLSHST"
        companyCode_RAJ = "1000"
        companyCode_KAD = "2000"
        companyCode_IN03 = "IN03"
        userGroupName = "SNOP"
        sapQueryName = "SOP_CUSTMASTER"
        sMachMatlCode = ""
        
        planningCycle = Application.CurrentProject.Connection.Execute("SELECT PlanningCycle FROM tbl001PlanningCycle").Fields("PlanningCycle").value
        
        folderPath = VBA.CStr(basUtility.GetSetting("SalesHistoryFolder"))
        
        msg = ""
            
        'Anjali 02/15/2016 added tMinus0Start & tMinus0End code to fetch feb's data from E01 - this section will be deleted in PC201603
        ''' stm deleted 3/01/16
        ''' tMinus0Start = "01" & "." & VBA.Format(VBA.Month(VBA.DateAdd("m", 0, VBA.Now())), "00") & "." & VBA.Year(VBA.DateAdd("m", 0, VBA.Now())) 'Tobedeleted PC 201603
        ''' tMinus0End = VBA.Format(VBA.Day(modMain.date_MonthLastDay( _
        '''    VBA.Year(VBA.DateAdd("m", 0, VBA.Now())) & "-" & VBA.Format(VBA.Month(VBA.DateAdd("m", 0, VBA.Now())), "00") & "-" & "01" _
        ''' )), "00") & "." & VBA.Format(VBA.Month(VBA.DateAdd("m", 0, VBA.Now())), "00") & "." & VBA.Year(VBA.DateAdd("m", 0, VBA.Now()))

        tMinus1Start = "01" & "." & VBA.Format(VBA.Month(VBA.DateAdd("m", -1, VBA.Now())), "00") & "." & VBA.Year(VBA.DateAdd("m", -1, VBA.Now()))
        tMinus1End = VBA.Format(VBA.Day(modMain.date_MonthLastDay( _
            VBA.Year(VBA.DateAdd("m", -1, VBA.Now())) & "-" & VBA.Format(VBA.Month(VBA.DateAdd("m", -1, VBA.Now())), "00") & "-" & "01" _
            )), "00") & "." & VBA.Format(VBA.Month(VBA.DateAdd("m", -1, VBA.Now())), "00") & "." & VBA.Year(VBA.DateAdd("m", -1, VBA.Now()))

        tMinus2Start = "01" & "." & VBA.Format(VBA.Month(VBA.DateAdd("m", -2, VBA.Now())), "00") & "." & VBA.Year(VBA.DateAdd("m", -2, VBA.Now()))
        tMinus2End = VBA.Format(VBA.Day(modMain.date_MonthLastDay( _
            VBA.Year(VBA.DateAdd("m", -2, VBA.Now())) & "-" & VBA.Format(VBA.Month(VBA.DateAdd("m", -2, VBA.Now())), "00") & "-" & "01" _
            )), "00") & "." & VBA.Format(VBA.Month(VBA.DateAdd("m", -2, VBA.Now())), "00") & "." & VBA.Year(VBA.DateAdd("m", -2, VBA.Now()))

        tMinus3Start = "01" & "." & VBA.Format(VBA.Month(VBA.DateAdd("m", -3, VBA.Now())), "00") & "." & VBA.Year(VBA.DateAdd("m", -3, VBA.Now()))
        tMinus3End = VBA.Format(VBA.Day(modMain.date_MonthLastDay( _
            VBA.Year(VBA.DateAdd("m", -3, VBA.Now())) & "-" & VBA.Format(VBA.Month(VBA.DateAdd("m", -3, VBA.Now())), "00") & "-" & "01" _
            )), "00") & "." & VBA.Format(VBA.Month(VBA.DateAdd("m", -3, VBA.Now())), "00") & "." & VBA.Year(VBA.DateAdd("m", -3, VBA.Now()))

        
        'Anjali 02/15/2016 added last three months of sales history from E01 (Adding the current month as well temporarily
          
        '''FileName = "APPM_PC" & planningCycle & "_" & companyCode_IN03 & "_" & VBA.Year(VBA.DateAdd("m", 0, VBA.Now())) & VBA.Format(VBA.Month(VBA.DateAdd("m", 0, VBA.Now())), "00") & ".txt" 'Tobedeleted PC 201603
        '''VBA.Kill folderPath & "\" & FileName
        '''SAP_RunSAPScript_SalesHistory session, companyCode_IN03, tMinus0Start, tMinus0End, folderPath, FileName, machineMatlCodeFilter, GenAppendFlag, layoutName
        '''If VBA.Dir(folderPath & "\" & FileName) = "" Then msg = msg & vbNewLine & "Error - Did not download: " & folderPath & "\" & FileName
        
        
        
         'new code STM  '*&*#
         FileName = "APPM_PC_SALESHIST" & ".txt"
         If FileExists(folderPath & "\" & FileName) Then VBA.Kill folderPath & "\" & FileName
         sGenOrAppend = "GEN"
         sSQL = "SELECT tbl004MasterData002MachineInfo.PMCode, tbl004MasterData002MachineInfo.MatlCode, tbl004MasterData002MachineInfo.Active " & vbCrLf & _
                "FROM tbl004MasterData002MachineInfo " & vbCrLf & _
                "WHERE (((tbl004MasterData002MachineInfo.PMCode) <> 'Missing') And ((tbl004MasterData002MachineInfo.Active) = True)) " & vbCrLf & _
                "ORDER BY tbl004MasterData002MachineInfo.MatlCode;"
         Set RS = CurrentDb.OpenRecordset(sSQL)
         RS.MoveLast
         RS.MoveFirst
    
         While RS.EOF = False
            DoEvents
            sMachMatlCode = RS!MatlCode
            If FileExists(folderPath & "\" & FileName) Then sGenOrAppend = "APPEND"
            
            SAP_RunSAPScript_SalesHistory session, companyCode_IN03, tMinus3Start, tMinus1End, folderPath, FileName, sMachMatlCode, sGenOrAppend, layoutName
            
            RS.MoveNext
         Wend

         
        
''    THIS CODE WAS ACTIVE UNTIL JUNE 2016 WHEN SAP PERFORMANCE REQUIRED SPLIT BY MACHINE INSTEAD OF MONTH
''
''        FileName = "APPM_PC" & planningCycle & "_" & companyCode_IN03 & "_" & VBA.Year(VBA.DateAdd("m", -1, VBA.Now())) & VBA.Format(VBA.Month(VBA.DateAdd("m", -1, VBA.Now())), "00") & ".txt"
''        If FileExists(folderPath & " \ " & FileName) Then VBA.Kill folderPath & "\" & FileName
''        SAP_RunSAPScript_SalesHistory session, companyCode_IN03, tMinus1Start, tMinus1End, folderPath, FileName, layoutName
''        If VBA.Dir(folderPath & "\" & FileName) = "" Then msg = msg & vbNewLine & "Error - Did not download: " & folderPath & "\" & FileName
''
''        FileName = "APPM_PC" & planningCycle & "_" & companyCode_IN03 & "_" & VBA.Year(VBA.DateAdd("m", -2, VBA.Now())) & VBA.Format(VBA.Month(VBA.DateAdd("m", -2, VBA.Now())), "00") & ".txt"
''        If FileExists(folderPath & " \ " & FileName) Then VBA.Kill folderPath & "\" & FileName
''        SAP_RunSAPScript_SalesHistory session, companyCode_IN03, tMinus2Start, tMinus2End, folderPath, FileName, layoutName
''        If VBA.Dir(folderPath & "\" & FileName) = "" Then msg = msg & vbNewLine & "Error - Did not download: " & folderPath & "\" & FileName
''
''        FileName = "APPM_PC" & planningCycle & "_" & companyCode_IN03 & "_" & VBA.Year(VBA.DateAdd("m", -3, VBA.Now())) & VBA.Format(VBA.Month(VBA.DateAdd("m", -3, VBA.Now())), "00") & ".txt"
''        If FileExists(folderPath & " \ " & FileName) Then VBA.Kill folderPath & "\" & FileName
''        SAP_RunSAPScript_SalesHistory session, companyCode_IN03, tMinus3Start, tMinus3End, folderPath, FileName, layoutName
''        If VBA.Dir(folderPath & "\" & FileName) = "" Then msg = msg & vbNewLine & "Error - Did not download: " & folderPath & "\" & FileName
        
        '''' DISABLE APP INSTANCE SCRIPT MAY 2016  ''''   'Anjali 02/15/2016 Editied the following to hardcode session & Layout
        '''' DISABLE APP INSTANCE SCRIPT MAY 2016  ''''
        '''' DISABLE APP INSTANCE SCRIPT MAY 2016  ''''   'fileName = "APPM_PC201308_1000_201307.txt"
        '''' DISABLE APP INSTANCE SCRIPT MAY 2016  ''''
        '''' DISABLE APP INSTANCE SCRIPT MAY 2016  ''''   'Rajahmundry - last three months of sales history
        '''' DISABLE APP INSTANCE SCRIPT MAY 2016  ''''   FileName = "APPM_PC" & planningCycle & "_" & companyCode_RAJ & "_" & VBA.Year(VBA.DateAdd("m", -1, VBA.Now())) & VBA.Format(VBA.Month(VBA.DateAdd("m", -1, VBA.Now())), "00") & ".txt"
        '''' DISABLE APP INSTANCE SCRIPT MAY 2016  ''''   VBA.Kill folderPath & "\" & FileName
        '''' DISABLE APP INSTANCE SCRIPT MAY 2016  ''''   SAP_RunSAPScript_SalesHistory APPsession, companyCode_RAJ, tMinus1Start, tMinus1End, folderPath, FileName, layoutName
        '''' DISABLE APP INSTANCE SCRIPT MAY 2016  ''''   If VBA.Dir(folderPath & "\" & FileName) = "" Then msg = msg & vbNewLine & "Error - Did not download: " & folderPath & "\" & FileName
        '''' DISABLE APP INSTANCE SCRIPT MAY 2016  ''''
        '''' DISABLE APP INSTANCE SCRIPT MAY 2016  ''''   FileName = "APPM_PC" & planningCycle & "_" & companyCode_RAJ & "_" & VBA.Year(VBA.DateAdd("m", -2, VBA.Now())) & VBA.Format(VBA.Month(VBA.DateAdd("m", -2, VBA.Now())), "00") & ".txt"
        '''' DISABLE APP INSTANCE SCRIPT MAY 2016  ''''   VBA.Kill folderPath & "\" & FileName
        '''' DISABLE APP INSTANCE SCRIPT MAY 2016  ''''   SAP_RunSAPScript_SalesHistory APPsession, companyCode_RAJ, tMinus2Start, tMinus2End, folderPath, FileName, layoutName
        '''' DISABLE APP INSTANCE SCRIPT MAY 2016  ''''   If VBA.Dir(folderPath & "\" & FileName) = "" Then msg = msg & vbNewLine & "Error - Did not download: " & folderPath & "\" & FileName
        '''' DISABLE APP INSTANCE SCRIPT MAY 2016  ''''
        '''' DISABLE APP INSTANCE SCRIPT MAY 2016  ''''
        '''' DISABLE APP INSTANCE SCRIPT MAY 2016  ''''   FileName = "APPM_PC" & planningCycle & "_" & companyCode_RAJ & "_" & VBA.Year(VBA.DateAdd("m", -3, VBA.Now())) & VBA.Format(VBA.Month(VBA.DateAdd("m", -3, VBA.Now())), "00") & ".txt"
        '''' DISABLE APP INSTANCE SCRIPT MAY 2016  ''''   VBA.Kill folderPath & "\" & FileName
        '''' DISABLE APP INSTANCE SCRIPT MAY 2016  ''''   SAP_RunSAPScript_SalesHistory APPsession, companyCode_RAJ, tMinus3Start, tMinus3End, folderPath, FileName, layoutName
        '''' DISABLE APP INSTANCE SCRIPT MAY 2016  ''''   If VBA.Dir(folderPath & "\" & FileName) = "" Then msg = msg & vbNewLine & "Error - Did not download: " & folderPath & "\" & FileName
        '''' DISABLE APP INSTANCE SCRIPT MAY 2016  ''''
        '''' DISABLE APP INSTANCE SCRIPT MAY 2016  ''''   'Kadiyam - last three months of sales history
        '''' DISABLE APP INSTANCE SCRIPT MAY 2016  ''''   FileName = "APPM_PC" & planningCycle & "_" & companyCode_KAD & "_" & VBA.Year(VBA.DateAdd("m", -1, VBA.Now())) & VBA.Format(VBA.Month(VBA.DateAdd("m", -1, VBA.Now())), "00") & ".txt"
        '''' DISABLE APP INSTANCE SCRIPT MAY 2016  ''''   VBA.Kill folderPath & "\" & FileName
        '''' DISABLE APP INSTANCE SCRIPT MAY 2016  ''''   SAP_RunSAPScript_SalesHistory APPsession, companyCode_KAD, tMinus1Start, tMinus1End, folderPath, FileName, layoutName
        '''' DISABLE APP INSTANCE SCRIPT MAY 2016  ''''   If VBA.Dir(folderPath & "\" & FileName) = "" Then msg = msg & vbNewLine & "Error - Did not download: " & folderPath & "\" & FileName
        '''' DISABLE APP INSTANCE SCRIPT MAY 2016  ''''
        '''' DISABLE APP INSTANCE SCRIPT MAY 2016  ''''   FileName = "APPM_PC" & planningCycle & "_" & companyCode_KAD & "_" & VBA.Year(VBA.DateAdd("m", -2, VBA.Now())) & VBA.Format(VBA.Month(VBA.DateAdd("m", -2, VBA.Now())), "00") & ".txt"
        '''' DISABLE APP INSTANCE SCRIPT MAY 2016  ''''   VBA.Kill folderPath & "\" & FileName
        '''' DISABLE APP INSTANCE SCRIPT MAY 2016  ''''   SAP_RunSAPScript_SalesHistory APPsession, companyCode_KAD, tMinus2Start, tMinus2End, folderPath, FileName, layoutName
        '''' DISABLE APP INSTANCE SCRIPT MAY 2016  ''''   If VBA.Dir(folderPath & "\" & FileName) = "" Then msg = msg & vbNewLine & "Error - Did not download: " & folderPath & "\" & FileName
        '''' DISABLE APP INSTANCE SCRIPT MAY 2016  ''''
        '''' DISABLE APP INSTANCE SCRIPT MAY 2016  ''''   FileName = "APPM_PC" & planningCycle & "_" & companyCode_KAD & "_" & VBA.Year(VBA.DateAdd("m", -3, VBA.Now())) & VBA.Format(VBA.Month(VBA.DateAdd("m", -3, VBA.Now())), "00") & ".txt"
        '''' DISABLE APP INSTANCE SCRIPT MAY 2016  ''''   VBA.Kill folderPath & "\" & FileName
        '''' DISABLE APP INSTANCE SCRIPT MAY 2016  ''''   SAP_RunSAPScript_SalesHistory APPsession, companyCode_KAD, tMinus3Start, tMinus3End, folderPath, FileName, layoutName
        '''' DISABLE APP INSTANCE SCRIPT MAY 2016  ''''   If VBA.Dir(folderPath & "\" & FileName) = "" Then msg = msg & vbNewLine & "Error - Did not download: " & folderPath & "\" & FileName

        '''' DISABLE APP INSTANCE SCRIPT MAY 2016  ''''   'Entire list of dealers from SAP (Customer Master)
        '''' DISABLE APP INSTANCE SCRIPT MAY 2016  ''''   FileName = "APPM_KNA1.txt"
        '''' DISABLE APP INSTANCE SCRIPT MAY 2016  ''''   VBA.Kill folderPath & "\" & FileName
        '''' DISABLE APP INSTANCE SCRIPT MAY 2016  ''''   SAP_RunSAPScript_CustomerMaster APPsession, userGroupName, sapQueryName, folderPath & "\" & FileName
        '''' DISABLE APP INSTANCE SCRIPT MAY 2016  ''''   If VBA.Dir(folderPath & "\" & FileName) = "" Then msg = msg & vbNewLine & "Error - Did not download: " & folderPath & "\" & FileName
         
        'Anjali 02/16/16 Edited the code to include new sub - SAP_RunSAPQueryGeneric for customer master
        userGroupName = "APPMBMOS"
        FileName = "APPM_KNA1_E01.txt"
        VBA.Kill folderPath & "\" & FileName
        SAP_RunSAPQueryGeneric session, userGroupName, sapQueryName, "STANDARD", folderPath & "\" & FileName
        If VBA.Dir(folderPath & "\" & FileName) = "" Then msg = msg & vbNewLine & "Error - Did not download: " & folderPath & "\" & FileName

        'Anjali 02/16/16 Edited the code to include new sub - SAP_RunSAPQueryGeneric for Material master
        userGroupName = "APPMBMOS"
        sapQueryName = "SOP_MATLMaster"
        FileName = "APPM_MARA_E01.txt"
        VBA.Kill folderPath & "\" & FileName
        SAP_RunSAPQueryGeneric session, userGroupName, sapQueryName, "STANDARD", folderPath & "\" & FileName
        If VBA.Dir(folderPath & "\" & FileName) = "" Then msg = msg & vbNewLine & "Error - Did not download: " & folderPath & "\" & FileName

        If msg = "" Then
            MsgBox "Successfully downloaded data from SAP." & vbNewLine & "Time taken = " & VBA.DateDiff("n", startTime, VBA.Now()) & " min.", vbInformation, Application.CurrentProject.Name
        Else
            MsgBox "Not all files were downloaded from SAP." & vbNewLine & "List of files that were NOT downloaded from SAP:" & vbNewLine & msg, vbExclamation, Application.CurrentProject.Name
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
        ElseIf Err.Number = 53 Then 'File not found
            Resume Next
        Else
            'ProcessMsg Err.Number, Err.Description, "", ""
            MsgBox Err.Number & " - " & Err.Description
        End If
    End If

    Exit Function
    Resume
End Function

Public Function SAP_RunSAPKNA1Script()
'**********************************************
'Author  :  RKP
'Date/Ver:  08-20-13/V01
'Input   :
'Output  :
'Comments:  Runs KNA1 script only
'**********************************************
    On Error GoTo Err_Handler

    Dim session         As Object
    Dim APPsession      As Object
    Dim sapInstance     As String
    Dim tMinus1Start    As String
    Dim tMinus1End      As String
    Dim tMinus2Start    As String
    Dim tMinus2End      As String
    Dim tMinus3Start    As String
    Dim tMinus3End      As String
    Dim companyCode_RAJ As String
    Dim companyCode_KAD As String
    Dim companyCode_IN03 As String
    Dim layoutName      As String
    Dim userGroupName   As String
    Dim sapQueryName    As String
    Dim folderPath      As String
    Dim FileName        As String
    Dim filePath        As String
    Dim planningCycle   As String
    Dim msg             As String
    Dim startTime       As Date
    Dim APPsapInstance  As String
    
    startTime = VBA.Now()
    
    '////////
    
    APPsapInstance = "APP"
    sapInstance = VBA.CStr(basUtility.GetSetting("SAPInstance"))
    
    Set APPsession = modMain.SAP_GetSession(APPsapInstance, False)
    Set session = modMain.SAP_GetSession(sapInstance, False)
    If session Is Nothing Or APPsession Is Nothing Then
        If session Is Nothing Then
            MsgBox "No active SAP session found for """ & sapInstance & """." & vbNewLine & "Please log on to """ & sapInstance & """ and try the operation again.", vbExclamation, Application.CurrentProject.Name
        ElseIf APPsession Is Nothing Then
            MsgBox "No active SAP session found for """ & APPsapInstance & """." & vbNewLine & "Please log on to """ & sapInstance & """ and try the operation again.", vbExclamation, Application.CurrentProject.Name
        End If
    Else
        layoutName = "/SOPSLSHST"
        companyCode_RAJ = "1000"
        companyCode_KAD = "2000"
        companyCode_IN03 = "IN03"
        userGroupName = "SNOP"
        sapQueryName = "SOP_CUSTMASTER"
        
        planningCycle = Application.CurrentProject.Connection.Execute("SELECT PlanningCycle FROM tbl001PlanningCycle").Fields("PlanningCycle").value
        
        folderPath = VBA.CStr(basUtility.GetSetting("SalesHistoryFolder"))
        
        msg = ""
    
        'Entire list of dealers from SAP (Customer Master)
        userGroupName = "SNOP"
        FileName = "APPM_KNA1.txt"
        VBA.Kill folderPath & "\" & FileName
        SAP_RunSAPScript_CustomerMaster APPsession, userGroupName, sapQueryName, folderPath & "\" & FileName
        If VBA.Dir(folderPath & "\" & FileName) = "" Then msg = msg & vbNewLine & "Error - Did not download: " & folderPath & "\" & FileName
                 
        'Anjali 02/16/16 Edited the code to include new sub - SAP_RunSAPQueryGeneric for customer master
        userGroupName = "APPMBMOS"
        FileName = "APPM_KNA1_E01.txt"
        VBA.Kill folderPath & "\" & FileName
        SAP_RunSAPQueryGeneric session, userGroupName, sapQueryName, "STANDARD", folderPath & "\" & FileName
        If VBA.Dir(folderPath & "\" & FileName) = "" Then msg = msg & vbNewLine & "Error - Did not download: " & folderPath & "\" & FileName

        'Anjali 02/16/16 Edited the code to include new sub - SAP_RunSAPQueryGeneric for Material master
        userGroupName = "APPMBMOS"
        sapQueryName = "SOP_MATLMaster"
        FileName = "APPM_MARA_E01.txt"
        VBA.Kill folderPath & "\" & FileName
        SAP_RunSAPQueryGeneric session, userGroupName, sapQueryName, "STANDARD", folderPath & "\" & FileName
        If VBA.Dir(folderPath & "\" & FileName) = "" Then msg = msg & vbNewLine & "Error - Did not download: " & folderPath & "\" & FileName
        
        
        If msg = "" Then
            MsgBox "Successfully downloaded data from SAP." & vbNewLine & "Time taken = " & VBA.DateDiff("n", startTime, VBA.Now()) & " min.", vbInformation, Application.CurrentProject.Name
        Else
            MsgBox "Not all files were downloaded from SAP." & vbNewLine & "List of files that were NOT downloaded from SAP:" & vbNewLine & msg, vbExclamation, Application.CurrentProject.Name
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
        ElseIf Err.Number = 53 Then 'File not found
            Resume Next
        Else
            'ProcessMsg Err.Number, Err.Description, "", ""
            MsgBox Err.Number & " - " & Err.Description
        End If
    End If

    Exit Function
    Resume
End Function


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


Public Function ChangeSettingsTablesTestMode(sType As String) As Boolean
'***********************************************************************
'Author  :  STM
'Date/Ver:  11-04-13/V01
'USAGE   :  i = ChangeSettingsTablesTestMode("TEST")
'           i = ChangeSettingsTablesTestMode("LIVE")
'Input   :
'Output  :
'Comments:
'***********************************************************************
Dim sTimeStamp As String
   sTimeStamp = VBA.Format(VBA.Now(), "yyyymmdd-hhmm")
   DoCmd.Rename "tbl000Settings" & "_bkup_" & sTimeStamp, acTable, "tbl000Settings"
   DoCmd.Rename "tbl000SettingsExcelPivots" & "_bkup_" & sTimeStamp, acTable, "tbl000SettingsExcelPivots"

   Select Case sType
   Case "TEST"
      DoCmd.CopyObject , "tbl000Settings", acTable, "tbl000Settings_Test"
      DoCmd.CopyObject , "tbl000SettingsExcelPivots", acTable, "tbl000SettingsExcelPivots_Test"
   Case "LIVE"
      DoCmd.CopyObject , "tbl000Settings", acTable, "tbl000Settings_Live"
      DoCmd.CopyObject , "tbl000SettingsExcelPivots", acTable, "tbl000SettingsExcelPivots_Live"
   End Select
Debug.Print sTimeStamp
ChangeSettingsTablesTestMode = True
End Function


Public Function date_GetPriorMonth(ByVal currentMonth As Integer) As String
'**********************************************
'Author  :  RKP
'Date/Ver:  08-20-13/V01
'Input   :
'Output  :
'Comments:
'**********************************************
    On Error GoTo Err_Handler

    If currentMonth = 1 Then
        date_GetPriorMonth = "12"
    Else
        date_GetPriorMonth = VBA.Format(currentMonth - 1, "00")
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

    Exit Function
    Resume
End Function

Public Function date_MonthLastDay(ByVal dCurrDate As Date)
  Dim dFirstDayNextMonth As Date
  
  On Error GoTo lbl_Error
 
  date_MonthLastDay = Empty
  dFirstDayNextMonth = DateSerial(CInt(Format(dCurrDate, "yyyy")), CInt(Format(dCurrDate, "mm")) + 1, 1)
  date_MonthLastDay = DateAdd("d", -1, dFirstDayNextMonth)
  
  Exit Function
lbl_Error:
  MsgBox Err.Description, vbOKOnly + vbExclamation
End Function

Public Function SAP_RunSAPScript_SalesHistory( _
    ByRef session As Object, _
    ByVal companyCode As String, _
    ByVal startDate As String, _
    ByVal endDate As String, _
    ByVal folderPath As String, _
    ByVal FileName As String, _
    ByVal machineMatlCodeFilter As String, _
    ByVal GenAppendFlag As String, _
    ByVal layoutName As String _
)
'**********************************************
'Author  :  RKP
'Date/Ver:  08-20-13/V01
'Input   :
'Output  :
'Comments:
'Usage:     SAP_RunSAPScript_SalesHistory session, companyCode_IN03, tMinus0Start, tMinus0End, folderPath, FileName, _
                                          machineMatlCodeFilter, GenAppendFlag, layoutName
'           SAP_RunSAPScript_SalesHistory session, "IN03", "01.08.2016", "31.10.2016", "\\sa53fs01\Secunderabad-A53\APPM S&OP\RawFiles", "APPM_PC201611_IN03_201611.txt", _
'                                         "A", "GEN", "/SOPSLSHST"
'**********************************************

On Error GoTo Err_Handler

'Anjali 02/15/2016 added the IN03 code in the If clause
Dim strCompanyCode As String
strCompanyCode = companyCode

If strCompanyCode = "IN03" Then
   session.findById("wnd[0]/tbar[0]/btn[3]").press
   session.findById("wnd[0]/tbar[0]/btn[3]").press
   session.findById("wnd[0]/tbar[0]/btn[3]").press
   session.findById("wnd[0]/tbar[0]/btn[3]").press
   
   session.findById("wnd[0]/tbar[0]/okcd").text = "ZSD04"
   session.findById("wnd[0]").sendVKey 0
   session.findById("wnd[0]/usr/ctxtSO_VTWEG-LOW").text = ""
   session.findById("wnd[0]/usr/ctxtSO_VTWEG-HIGH").text = ""
   session.findById("wnd[0]/usr/ctxtSO_WERKS-LOW").text = ""
   session.findById("wnd[0]/usr/ctxtSO_WERKS-HIGH").text = ""
   session.findById("wnd[0]/usr/ctxtSO_LGORT-LOW").text = ""
   session.findById("wnd[0]/usr/ctxtSO_LGORT-HIGH").text = ""

   session.findById("wnd[0]/usr/ctxtSO_MATNR-LOW").text = machineMatlCodeFilter & "*"
   session.findById("wnd[0]/usr/ctxtSO_MATNR-HIGH").text = ""
   
   session.findById("wnd[0]/usr/ctxtSO_VBELN-LOW").text = ""
   session.findById("wnd[0]/usr/ctxtSO_VBELN-HIGH").text = ""
   session.findById("wnd[0]/usr/ctxtSO_FKDAT-LOW").text = startDate
   session.findById("wnd[0]/usr/ctxtSO_FKDAT-HIGH").text = endDate
   session.findById("wnd[0]/usr/txtSO_SONO-LOW").text = ""
   session.findById("wnd[0]/usr/txtSO_SONO-HIGH").text = ""
   '' session.findById("wnd[0]/usr/txtSO_AUART-LOW").text = ""   ''' ERROR HERE STM
   '' session.findById("wnd[0]/usr/txtSO_AUART-HIGH").text = ""  ''' ERROR HERE STM
   session.findById("wnd[0]/usr/ctxtSO_KUNAG-LOW").text = ""
   session.findById("wnd[0]/usr/ctxtSO_KUNAG-HIGH").text = ""
   session.findById("wnd[0]/usr/txtSO_PAYER-LOW").text = ""
   session.findById("wnd[0]/usr/txtSO_PAYER-HIGH").text = ""
   session.findById("wnd[0]/usr/ctxtSP_VARI").text = layoutName
   session.findById("wnd[0]/usr/ctxtSP_VARI").SetFocus
   session.findById("wnd[0]/usr/ctxtSP_VARI").caretPosition = 9
   session.findById("wnd[0]/tbar[1]/btn[8]").press
   session.findById("wnd[0]/tbar[1]/btn[45]").press    'LOCAL FILE (CNTRL+sHIFT+f9)
   session.findById("wnd[1]/tbar[0]/btn[0]").press
      
   session.findById("wnd[1]/usr/ctxtDY_PATH").text = folderPath
   session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = FileName
   ''session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 7  'not needed, I don't think
   
   '*&*# deal with the generate append question here
   Select Case GenAppendFlag
      Case "GEN", "Gen", "Generate", "GENERATE", "generate", "gen"
         session.findById("wnd[1]/tbar[0]/btn[11]").press   '//GENERATE  [GENERATE-REPLACE-EXTEND(APPEND)-X(CANCEL)]
         Debug.Print Space(3) & machineMatlCodeFilter & "*" & "  " & GenAppendFlag
      Case "APP", "App", "Append", "APPEND", "append", "app"
         session.findById("wnd[1]/tbar[0]/btn[7]").press   '// extend (F7)
         Debug.Print Space(3) & machineMatlCodeFilter & "*" & "  " & GenAppendFlag
   End Select
   
   '' session.findById("wnd[0]").Maximize       ''' DO NOT MAXIMIZE STM
   session.findById("wnd[0]/tbar[0]/btn[3]").press
   session.findById("wnd[0]/tbar[0]/btn[3]").press
   
Else  'DEAD CODE NOW THAT NO APPM SAP INSTANCE

   session.findById("wnd[0]/tbar[0]/btn[3]").press
   session.findById("wnd[0]/tbar[0]/btn[3]").press
   session.findById("wnd[0]/tbar[0]/btn[3]").press
   session.findById("wnd[0]/tbar[0]/btn[3]").press
   
   session.findById("wnd[0]/tbar[0]/okcd").text = "ZSD04"
   session.findById("wnd[0]/tbar[0]/btn[0]").press
   session.findById("wnd[0]/usr/ctxtP_BUKRS").text = companyCode '"IN03"
   'session.findById("wnd[0]/usr/ctxtP_BUKRS").text = companyCode '"1000"
   session.findById("wnd[0]/usr/ctxtS_FKDAT-LOW").text = startDate '"01.07.2013"
   session.findById("wnd[0]/usr/ctxtS_FKDAT-HIGH").text = endDate '"31.07.2013"
   'session.findById("wnd[0]/usr/ctxtS_FKDAT-HIGH").SetFocus
   'session.findById("wnd[0]/usr/ctxtS_FKDAT-HIGH").caretPosition = 10
   session.findById("wnd[0]/usr/btn%_P_WERKS_%_APP_%-VALU_PUSH").press
   session.findById("wnd[1]/tbar[0]/btn[16]").press
   session.findById("wnd[1]/tbar[0]/btn[12]").press
   session.findById("wnd[0]/usr/btn%_S_MATNR_%_APP_%-VALU_PUSH").press
   session.findById("wnd[1]/tbar[0]/btn[16]").press
   session.findById("wnd[1]/tbar[0]/btn[12]").press
   session.findById("wnd[0]/usr/btn%_S_VBELN_%_APP_%-VALU_PUSH").press
   session.findById("wnd[1]/tbar[0]/btn[16]").press
   session.findById("wnd[1]/tbar[0]/btn[12]").press
   session.findById("wnd[0]/usr/btn%_S_SONO_%_APP_%-VALU_PUSH").press
   session.findById("wnd[1]/tbar[0]/btn[16]").press
   session.findById("wnd[1]/tbar[0]/btn[12]").press
   session.findById("wnd[0]/usr/btn%_S_SLOC_%_APP_%-VALU_PUSH").press
   session.findById("wnd[1]/tbar[0]/btn[16]").press
   session.findById("wnd[1]/tbar[0]/btn[12]").press
   session.findById("wnd[0]/usr/btn%_S_KUNAG_%_APP_%-VALU_PUSH").press
   session.findById("wnd[1]/tbar[0]/btn[16]").press
   session.findById("wnd[1]/tbar[0]/btn[12]").press
   session.findById("wnd[0]/usr/btn%_S_VTWEG_%_APP_%-VALU_PUSH").press
   session.findById("wnd[1]/tbar[0]/btn[16]").press
   session.findById("wnd[1]/tbar[0]/btn[12]").press
   session.findById("wnd[0]/usr/btn%_S_DESTI_%_APP_%-VALU_PUSH").press
   session.findById("wnd[1]/tbar[0]/btn[16]").press
   session.findById("wnd[1]/tbar[0]/btn[12]").press
   session.findById("wnd[0]/usr/btn%_S_PAYER_%_APP_%-VALU_PUSH").press
   session.findById("wnd[1]/tbar[0]/btn[16]").press
   session.findById("wnd[1]/tbar[0]/btn[12]").press
   session.findById("wnd[0]/usr/btn%_S_LZONE_%_APP_%-VALU_PUSH").press
   session.findById("wnd[1]/tbar[0]/btn[16]").press
   session.findById("wnd[1]/tbar[0]/btn[12]").press
   session.findById("wnd[0]/usr/btn%_S_TAXCD_%_APP_%-VALU_PUSH").press
   session.findById("wnd[1]/tbar[0]/btn[16]").press
   session.findById("wnd[1]/tbar[0]/btn[12]").press
   session.findById("wnd[0]/usr/btn%_S_EXKNG_%_APP_%-VALU_PUSH").press
   session.findById("wnd[1]/tbar[0]/btn[16]").press
   session.findById("wnd[1]/tbar[0]/btn[12]").press
   session.findById("wnd[0]/usr/radR1").Select
   session.findById("wnd[0]/usr/ctxtP_LAYOUT").text = layoutName
   session.findById("wnd[0]/usr/radR1").SetFocus

   session.findById("wnd[0]/tbar[1]/btn[8]").press
   session.findById("wnd[0]/tbar[1]/btn[45]").press
   session.findById("wnd[1]/usr/sub:SAPLSPO5:0101/radSPOPLI-SELFLAG[0,0]").Select
   session.findById("wnd[1]/tbar[0]/btn[0]").press
   session.findById("wnd[1]/usr/ctxtDY_PATH").text = folderPath '"C:\_Work\Downloads"
   session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = FileName '"APPM_SalesHistory.txt"
   'session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 21
   session.findById("wnd[1]/tbar[0]/btn[0]").press
   session.findById("wnd[0]/tbar[0]/btn[3]").press
   session.findById("wnd[0]/tbar[0]/btn[3]").press
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
        ElseIf Err.Number = 619 Then 'The control could not be found by ID - Anjali 02/16/2016 added to handle error 619 and get out to main menu of SAP session
            MsgBox Err.Number & " - " & Err.Description
            session.findById("wnd[0]/tbar[0]/btn[3]").press
        Else
            'ProcessMsg Err.Number, Err.Description, "", ""
            MsgBox Err.Number & " - " & Err.Description
        End If
    End If

    Exit Function
    Resume
End Function

Public Function SAP_RunSAPScript_CustomerMaster( _
    ByRef session As Object, _
    ByVal userGroupName As String, _
    ByVal sapQueryName As String, _
    ByVal filePath As String _
)
'**********************************************
'Author  :  RKP
'Date/Ver:  08-20-13/V01
'Input   :
'Output  :
'Comments:
'**********************************************
    On Error GoTo Err_Handler

session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

session.findById("wnd[0]/tbar[0]/okcd").text = "SQ00"
session.findById("wnd[0]/tbar[0]/btn[0]").press
session.findById("wnd[0]/mbar/menu[5]/menu[0]").Select
session.findById("wnd[1]/usr/lbl[0,0]").SetFocus
'session.findById("wnd[1]/usr/lbl[0,0]").caretPosition = 4
session.findById("wnd[1]/tbar[0]/btn[2]").press
'session.findById("wnd[0]/mbar/menu[5]/menu[1]").Select
'session.findById("wnd[0]/tbar[0]/btn[3]").press
'session.findById("wnd[0]/tbar[1]/btn[19]").press
'session.findById("wnd[1]/tbar[0]/btn[71]").press
'session.findById("wnd[2]/usr/txtRSYSF-STRING").text = userGroupName '"SNOP"
'session.findById("wnd[2]/usr/txtRSYSF-STRING").caretPosition = 4
'session.findById("wnd[2]/tbar[0]/btn[0]").press
'session.findById("wnd[3]/tbar[0]/btn[2]").press
'session.findById("wnd[3]/usr/lbl[0,2]").SetFocus
'session.findById("wnd[3]/usr/lbl[0,2]").caretPosition = 1
'session.findById("wnd[3]").sendVKey 2
'session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/ctxtRS38R-QNUM").text = sapQueryName '"SOP_CUSTMASTER"
'session.findById("wnd[0]/usr/ctxtRS38R-QNUM").SetFocus
'session.findById("wnd[0]/usr/ctxtRS38R-QNUM").caretPosition = 14
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/rad%ALV").SetFocus
session.findById("wnd[0]/usr/rad%ALV").Select
session.findById("wnd[0]/usr/rad%DOWN").SetFocus
session.findById("wnd[0]/usr/rad%DOWN").Select
session.findById("wnd[0]/usr/txt%PATH").text = filePath '"C:\_Work\Downloads\APPM_KNA1.txt"
'session.findById("wnd[0]/usr/txt%PATH").SetFocus
'session.findById("wnd[0]/usr/txt%PATH").caretPosition = 32
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[1]/usr/txtRS38R-DBACC").text = "99999"
'session.findById("wnd[1]/usr/txtRS38R-DBACC").caretPosition = 5
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/chkRSAQDOWN-COLUMN").Selected = True
session.findById("wnd[1]/usr/chkRSAQDOWN-COLUMN").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

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


Public Function SAP_RunSAPScript_IsSAPSessionAvailable() As Boolean
'**********************************************
'Author  :  RKP
'Date/Ver:  08-20-13/V01
'Input   :
'Output  :
'Comments:
'**********************************************
    On Error GoTo Err_Handler

    

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


Public Function SAP_RunSAPScripts_History()
'**********************************************
'Author  :  RKP
'Date/Ver:  08-20-13/V01
'Input   :
'Output  :
'Comments: Anjali this code doesn't get called
'**********************************************
    On Error GoTo Err_Handler

    Dim session         As Object
    Dim sapInstance     As String
    Dim tMinus1Start    As String
    Dim tMinus1End      As String
    Dim tMinus2Start    As String
    Dim tMinus2End      As String
    Dim tMinus3Start    As String
    Dim tMinus3End      As String
    Dim companyCode_RAJ As String
    Dim companyCode_KAD As String
    Dim companyCode_IN03 As String
    Dim layoutName      As String
    Dim userGroupName   As String
    Dim sapQueryName    As String
    Dim folderPath      As String
    Dim FileName        As String
    Dim filePath        As String
    Dim planningCycle   As String
    Dim msg             As String
    Dim startTime       As Date
    
    Dim currDate        As Date
    Dim tStart          As String
    Dim tEnd            As String
    '*********************************************************************************************
    '// THIS ROUTINE IS DEAD CODE |  NOT CALLED FROM ANYWHERE
    '*********************************************************************************************
    startTime = VBA.Now()
    sapInstance = VBA.CStr(basUtility.GetSetting("SAPInstance"))
    Set session = modMain.SAP_GetSession(sapInstance, False)
    If session Is Nothing Then
        MsgBox "No active SAP session found for """ & sapInstance & """." & vbNewLine & "Please log on to """ & sapInstance & """ and try the operation again.", vbExclamation, Application.CurrentProject.Name
    Else
        layoutName = "/SOPSLSHST"
        companyCode_RAJ = "1000"
        companyCode_KAD = "2000"
        companyCode_IN03 = "IN03"
        userGroupName = "SNOP"
        sapQueryName = "SOP_CUSTMASTER"
        
        planningCycle = Application.CurrentProject.Connection.Execute("SELECT PlanningCycle FROM tbl001PlanningCycle").Fields("PlanningCycle").value
        
        folderPath = VBA.CStr(basUtility.GetSetting("SalesHistoryFolder"))
        
        Call Application.CurrentProject.Connection.Execute("DELETE * FROM tblSalesHistory")
        
        msg = ""
        
        currDate = "04/01/2013"
        Do
            DoEvents
            
            If VBA.Dir(CreateObject("Scripting.FileSystemObject").GetFile(Application.CurrentProject.FullName).ParentFolder & "\STOP.STOP") <> "" Then GoTo Err_Handler
            
            currDate = VBA.DateAdd("m", 1, currDate)
                        
            tStart = "01" & "." & VBA.Format(VBA.Month(currDate), "00") & "." & VBA.Year(currDate)
            tEnd = VBA.Format(VBA.Day(modMain.date_MonthLastDay( _
                VBA.Year(currDate) & "-" & VBA.Format(VBA.Month(currDate), "00") & "-" & "01" _
                )), "00") & "." & VBA.Format(VBA.Month(currDate), "00") & "." & VBA.Year(currDate)
            
            Debug.Print tStart & ", " & tEnd
            'Anjali 02/15/2016 added code for E01
            FileName = "APPM_PC" & planningCycle & "_" & companyCode_IN03 & "_" & VBA.Year(currDate) & VBA.Format(VBA.Month(currDate), "00") & ".txt"
            VBA.Kill folderPath & "\" & FileName
            SAP_RunSAPScript_SalesHistory session, companyCode_IN03, tStart, tEnd, folderPath, FileName, layoutName
            If VBA.Dir(folderPath & "\" & FileName) = "" Then msg = msg & vbNewLine & "Error - Did not download: " & folderPath & "\" & FileName
            If msg = "" Then
                Import_SalesHistory_History folderPath & "\" & FileName
            End If
            
''            FileName = "APPM_PC" & planningCycle & "_" & companyCode_RAJ & "_" & VBA.Year(currDate) & VBA.Format(VBA.Month(currDate), "00") & ".txt"
''            VBA.Kill folderPath & "\" & FileName
''            SAP_RunSAPScript_SalesHistory session, companyCode_RAJ, tStart, tEnd, folderPath, FileName, layoutName
''            If VBA.Dir(folderPath & "\" & FileName) = "" Then msg = msg & vbNewLine & "Error - Did not download: " & folderPath & "\" & FileName
''            If msg = "" Then
''                Import_SalesHistory_History folderPath & "\" & FileName
''            End If
''
''            FileName = "APPM_PC" & planningCycle & "_" & companyCode_KAD & "_" & VBA.Year(currDate) & VBA.Format(VBA.Month(currDate), "00") & ".txt"
''            VBA.Kill folderPath & "\" & FileName
''            SAP_RunSAPScript_SalesHistory session, companyCode_KAD, tStart, tEnd, folderPath, FileName, layoutName
''            If VBA.Dir(folderPath & "\" & FileName) = "" Then msg = msg & vbNewLine & "Error - Did not download: " & folderPath & "\" & FileName
''            If msg = "" Then
''                Import_SalesHistory_History folderPath & "\" & FileName
''            End If
            
        Loop Until currDate = "07/01/2013"
        
        MsgBox "Sales History Import Complete!" & vbNewLine & "Time taken:" & vbNewLine & VBA.DateDiff("n", startTime, VBA.Now()) & " min.", vbInformation, Application.CurrentProject.Name
        
        GoTo Err_Handler
            
''''        tMinus1Start = "01" & "." & VBA.Format(VBA.Month(VBA.DateAdd("m", 0, VBA.Now())), "00") & "." & VBA.Year(VBA.DateAdd("m", 0, VBA.Now()))
''''        tMinus1End = VBA.Format(VBA.Day(modMain.date_MonthLastDay( _
''''            VBA.Year(VBA.DateAdd("m", 0, VBA.Now())) & "-" & VBA.Format(VBA.Month(VBA.DateAdd("m", 0, VBA.Now())), "00") & "-" & "01" _
''''        )), "00") & "." & VBA.Format(VBA.Month(VBA.DateAdd("m", 0, VBA.Now())), "00") & "." & VBA.Year(VBA.DateAdd("m", 0, VBA.Now()))
''''
''''        tMinus2Start = "01" & "." & VBA.Format(VBA.Month(VBA.DateAdd("m", -2, VBA.Now())), "00") & "." & VBA.Year(VBA.DateAdd("m", -2, VBA.Now()))
''''        tMinus2End = VBA.Format(VBA.Day(modMain.date_MonthLastDay( _
''''            VBA.Year(VBA.DateAdd("m", -2, VBA.Now())) & "-" & VBA.Format(VBA.Month(VBA.DateAdd("m", -2, VBA.Now())), "00") & "-" & "01" _
''''        )), "00") & "." & VBA.Format(VBA.Month(VBA.DateAdd("m", -2, VBA.Now())), "00") & "." & VBA.Year(VBA.DateAdd("m", -2, VBA.Now()))
''''
''''        tMinus3Start = "01" & "." & VBA.Format(VBA.Month(VBA.DateAdd("m", -3, VBA.Now())), "00") & "." & VBA.Year(VBA.DateAdd("m", -3, VBA.Now()))
''''        tMinus3End = VBA.Format(VBA.Day(modMain.date_MonthLastDay( _
''''            VBA.Year(VBA.DateAdd("m", -3, VBA.Now())) & "-" & VBA.Format(VBA.Month(VBA.DateAdd("m", -3, VBA.Now())), "00") & "-" & "01" _
''''        )), "00") & "." & VBA.Format(VBA.Month(VBA.DateAdd("m", -3, VBA.Now())), "00") & "." & VBA.Year(VBA.DateAdd("m", -3, VBA.Now()))
''''
''''        'fileName = "APPM_PC201308_1000_201307.txt"
''''
''''        '// We can change this here to pull all three months at one time, but to pull only one machine
''''        '// Also need to change so that the file created goes generate append append append etc. so
''''        '// that there is only one file at all.
''''
''''
''''        'E01 - last three months of sales history
''''        FileName = "APPM_PC" & planningCycle & "_" & companyCode_IN03 & "_" & VBA.Year(VBA.DateAdd("m", 0, VBA.Now())) & VBA.Format(VBA.Month(VBA.DateAdd("m", 0, VBA.Now())), "00") & ".txt"
''''        VBA.Kill folderPath & "\" & FileName
''''        SAP_RunSAPScript_SalesHistory session, companyCode_IN03, tMinus1Start, tMinus1End, folderPath, FileName, layoutName
''''        If VBA.Dir(folderPath & "\" & FileName) = "" Then msg = msg & vbNewLine & "Error - Did not download: " & folderPath & "\" & FileName
''''
''''        FileName = "APPM_PC" & planningCycle & "_" & companyCode_IN03 & "_" & VBA.Year(VBA.DateAdd("m", -2, VBA.Now())) & VBA.Format(VBA.Month(VBA.DateAdd("m", -2, VBA.Now())), "00") & ".txt"
''''        VBA.Kill folderPath & "\" & FileName
''''        SAP_RunSAPScript_SalesHistory session, companyCode_IN03, tMinus1Start, tMinus1End, folderPath, FileName, layoutName
''''        If VBA.Dir(folderPath & "\" & FileName) = "" Then msg = msg & vbNewLine & "Error - Did not download: " & folderPath & "\" & FileName
''''
''''        FileName = "APPM_PC" & planningCycle & "_" & companyCode_IN03 & "_" & VBA.Year(VBA.DateAdd("m", -3, VBA.Now())) & VBA.Format(VBA.Month(VBA.DateAdd("m", -3, VBA.Now())), "00") & ".txt"
''''        VBA.Kill folderPath & "\" & FileName
''''        SAP_RunSAPScript_SalesHistory session, companyCode_IN03, tMinus1Start, tMinus1End, folderPath, FileName, layoutName
''''        If VBA.Dir(folderPath & "\" & FileName) = "" Then msg = msg & vbNewLine & "Error - Did not download: " & folderPath & "\" & FileName
''''
''''
''''''        'Rajahmundry - last three months of sales history
''''''        FileName = "APPM_PC" & planningCycle & "_" & companyCode_RAJ & "_" & VBA.Year(VBA.DateAdd("m", -1, VBA.Now())) & VBA.Format(VBA.Month(VBA.DateAdd("m", -1, VBA.Now())), "00") & ".txt"
''''''        VBA.Kill folderPath & "\" & FileName
''''''        SAP_RunSAPScript_SalesHistory session, companyCode_RAJ, tMinus1Start, tMinus1End, folderPath, FileName, layoutName
''''''        If VBA.Dir(folderPath & "\" & FileName) = "" Then msg = msg & vbNewLine & "Error - Did not download: " & folderPath & "\" & FileName
''''''
''''''        FileName = "APPM_PC" & planningCycle & "_" & companyCode_RAJ & "_" & VBA.Year(VBA.DateAdd("m", -2, VBA.Now())) & VBA.Format(VBA.Month(VBA.DateAdd("m", -2, VBA.Now())), "00") & ".txt"
''''''        VBA.Kill folderPath & "\" & FileName
''''''        SAP_RunSAPScript_SalesHistory session, companyCode_RAJ, tMinus1Start, tMinus1End, folderPath, FileName, layoutName
''''''        If VBA.Dir(folderPath & "\" & FileName) = "" Then msg = msg & vbNewLine & "Error - Did not download: " & folderPath & "\" & FileName
''''''
''''''
''''''        FileName = "APPM_PC" & planningCycle & "_" & companyCode_RAJ & "_" & VBA.Year(VBA.DateAdd("m", -3, VBA.Now())) & VBA.Format(VBA.Month(VBA.DateAdd("m", -3, VBA.Now())), "00") & ".txt"
''''''        VBA.Kill folderPath & "\" & FileName
''''''        SAP_RunSAPScript_SalesHistory session, companyCode_RAJ, tMinus1Start, tMinus1End, folderPath, FileName, layoutName
''''''        If VBA.Dir(folderPath & "\" & FileName) = "" Then msg = msg & vbNewLine & "Error - Did not download: " & folderPath & "\" & FileName
''''''
''''''        'Kadiyam - last three months of sales history
''''''        FileName = "APPM_PC" & planningCycle & "_" & companyCode_KAD & "_" & VBA.Year(VBA.DateAdd("m", -1, VBA.Now())) & VBA.Format(VBA.Month(VBA.DateAdd("m", -1, VBA.Now())), "00") & ".txt"
''''''        VBA.Kill folderPath & "\" & FileName
''''''        SAP_RunSAPScript_SalesHistory session, companyCode_KAD, tMinus1Start, tMinus1End, folderPath, FileName, layoutName
''''''        If VBA.Dir(folderPath & "\" & FileName) = "" Then msg = msg & vbNewLine & "Error - Did not download: " & folderPath & "\" & FileName
''''''
''''''        FileName = "APPM_PC" & planningCycle & "_" & companyCode_KAD & "_" & VBA.Year(VBA.DateAdd("m", -2, VBA.Now())) & VBA.Format(VBA.Month(VBA.DateAdd("m", -2, VBA.Now())), "00") & ".txt"
''''''        VBA.Kill folderPath & "\" & FileName
''''''        SAP_RunSAPScript_SalesHistory session, companyCode_KAD, tMinus1Start, tMinus1End, folderPath, FileName, layoutName
''''''        If VBA.Dir(folderPath & "\" & FileName) = "" Then msg = msg & vbNewLine & "Error - Did not download: " & folderPath & "\" & FileName
''''''
''''''        FileName = "APPM_PC" & planningCycle & "_" & companyCode_KAD & "_" & VBA.Year(VBA.DateAdd("m", -3, VBA.Now())) & VBA.Format(VBA.Month(VBA.DateAdd("m", -3, VBA.Now())), "00") & ".txt"
''''''        VBA.Kill folderPath & "\" & FileName
''''''        SAP_RunSAPScript_SalesHistory session, companyCode_KAD, tMinus1Start, tMinus1End, folderPath, FileName, layoutName
''''''        If VBA.Dir(folderPath & "\" & FileName) = "" Then msg = msg & vbNewLine & "Error - Did not download: " & folderPath & "\" & FileName
''''
''''        'Entire list of dealers from SAP (Customer Master)
''''        FileName = "APPM_KNA1.txt"
''''        VBA.Kill folderPath & "\" & FileName
''''        SAP_RunSAPScript_CustomerMaster session, userGroupName, sapQueryName, folderPath & "\" & FileName
''''        If VBA.Dir(folderPath & "\" & FileName) = "" Then msg = msg & vbNewLine & "Error - Did not download: " & folderPath & "\" & FileName
''''
''''        If msg = "" Then
''''            MsgBox "Successfully downloaded data from SAP." & vbNewLine & "Time taken = " & VBA.DateDiff("n", startTime, VBA.Now()) & " min.", vbInformation, Application.CurrentProject.Name
''''        Else
''''            MsgBox "Not all files were downloaded from SAP." & vbNewLine & "List of files that were NOT downloaded from SAP:" & vbNewLine & msg, vbExclamation, Application.CurrentProject.Name
''''        End If
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
        ElseIf Err.Number = 53 Then 'File not found
            Resume Next
        Else
            'ProcessMsg Err.Number, Err.Description, "", ""
            MsgBox Err.Number & " - " & Err.Description
        End If
    End If

    Exit Function
    Resume
End Function

Public Function Import_SalesHistory_History(Optional filePath As String)
'**********************************************
'Author  :  RKP
'Date/Ver:  05-01-13/V01
'Input   :
'Output  :
'Comments:
'Example:
'\\sa53fs01\Secunderabad-A53\APPM S&OP\RawFiles
'\\sa53fs01\Secunderabad-A53\APPM S&OP\PlanningCycle_201304\RawFiles\APPM_SalesHistory_RAJ_PCMinus3.txt
'filePath = FormatFilePath(basePath) & "PlanningCycle_" & planningCycle & "\RawFiles\" & "APPM_SalesHistory_RAJ_PCMinus" & ctr & ".txt"
'filePath = FormatFilePath(basePath) & "PlanningCycle_" & planningCycle & "\RawFiles\" & "APPM_SalesHistory_KAD_PCMinus" & ctr & ".txt"
'**********************************************
    On Error GoTo Err_Handler

    Dim linkSpecName        As String
    Dim basePath            As String
'    Dim filePath            As String
    Dim planningCycle       As String
    Dim tableName           As String
    Dim ctr                 As Integer
    Dim sql                 As String
    Dim ret                 As Long
    Dim RS                  As ADODB.Recordset
    
    linkSpecName = basUtility.GetSetting("SalesHistoryImportSpecName")
    basePath = basUtility.GetSetting("BaseFilePath")
    planningCycle = basUtility.GetScalarValue("SELECT PlanningCycle FROM tbl001PlanningCycle")
    tableName = "tbl002RawData010SalesHistory"
    
    sql = "DELETE * FROM [" & tableName & "]"
    ExecSQL sql, ret
    
'    For ctr = 3 To 1 Step -1
'        DoEvents
        
        'filePath = FormatFilePath(basePath) & "PlanningCycle_" & planningCycle & "\RawFiles\" & "APPM_SalesHistory_RAJ_PCMinus" & ctr & ".txt"
'        filePath = FormatFilePath(basUtility.GetSetting("SalesHistoryFolder")) & "APPM_PC" & planningCycle & "_1000_" & FormatDate(VBA.DateAdd("m", -1 * ctr, VBA.Now()), 1) & ".txt"
    Debug.Print filePath
    If VBA.Dir(filePath) = "" Then
        'error
    Else
        ImportCSVFilesUsingImportSpec linkSpecName, filePath, tableName
    End If
        
'        'filePath = FormatFilePath(basePath) & "PlanningCycle_" & planningCycle & "\RawFiles\" & "APPM_SalesHistory_KAD_PCMinus" & ctr & ".txt"
'        filePath = FormatFilePath(basUtility.GetSetting("SalesHistoryFolder")) & "APPM_PC" & planningCycle & "_2000_" & FormatDate(VBA.DateAdd("m", -1 * ctr, VBA.Now()), 1) & ".txt"
'        Debug.Print filePath
'        If VBA.Dir(filePath) = "" Then
'            'error
'        Else
'            ImportCSVFilesUsingImportSpec linkSpecName, filePath, tableName
'        End If
'    Next
    
'    tableName = "tbl002RawData015Customer"
'    sql = "DELETE * FROM [" & tableName & "]"
'    ExecSQL sql, ret
'
'    filePath = FormatFilePath(basUtility.GetSetting("SalesHistoryFolder")) & "APPM_KNA1" & ".txt"
'    Debug.Print filePath
'    linkSpecName = basUtility.GetSetting("KNA1ImportSpecName")
'    ImportCSVFilesUsingImportSpec linkSpecName, filePath, tableName
    
    basUtility.RunQueriesByPrefix "qx010ImportData"
    
    'MsgBox "Import finished."

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
    DoCmd.SetWarnings True

    Exit Function
    Resume
End Function


Function testNSRData()

Dim filePath As String
    DoCmd.SetWarnings False
    DoCmd.RunSQL "DELETE tblRawNSRData.* FROM tblRawNSRData"
    DoCmd.SetWarnings True
    filePath = FormatFilePath(basUtility.GetSetting("SalesHistoryFolder"))
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12, "tblRawNSRData", filePath & "NSRData.xlsx", True
               
    RunQueriesByPrefix "q400NSR0"
    Debug.Print "Done Test NSRData"
 
End Function

Public Sub SAP_RunSAPQueryGeneric( _
    ByRef session As Object, _
    sUserGroup As String, _
    sQueryname As String, _
    sVariant As String, _
    sFilePath As String, _
    Optional sAction As String = "Replace" _
)
'**********************************************
'Author  :  Anjali
'Date/Ver:  02/16/16
'Input   :
'Output  :
'Comments:  02/16/16 Added bew generic function to run SAP queries
'**********************************************

On Error GoTo Err_Handler

Dim bSAPFileExistsAtBegin As Boolean
Dim sSAPAction As String

sSAPAction = sAction
If VBA.Dir(sFilePath) = "" Then 'File Does not exist
   bSAPFileExistsAtBegin = False
Else
   bSAPFileExistsAtBegin = True                    'File exist
End If


session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
   
session.findById("wnd[0]/tbar[0]/okcd").text = "SQ00"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/mbar/menu[5]/menu[0]").Select
session.findById("wnd[1]/usr/radRAD1").Select
session.findById("wnd[1]/tbar[0]/btn[2]").press
session.findById("wnd[0]/tbar[1]/btn[19]").press
session.findById("wnd[1]/tbar[0]/btn[29]").press
session.findById("wnd[2]/usr/subSUB_DYN0500:SAPLSKBH:0600/btnAPP_WL_SING").press
session.findById("wnd[2]/usr/subSUB_DYN0500:SAPLSKBH:0600/btn600_BUTTON").press
session.findById("wnd[3]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = sUserGroup 'APPMBMOS
session.findById("wnd[3]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").caretPosition = 8
session.findById("wnd[3]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell").selectedRows = "0"
session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell").doubleClickCurrentCell
session.findById("wnd[0]/usr/ctxtRS38R-QNUM").text = sQueryname '"SOP_CUSTMASTER"
session.findById("wnd[0]/usr/ctxtRS38R-QNUM").SetFocus
session.findById("wnd[0]/usr/ctxtRS38R-QNUM").caretPosition = 14
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/ctxtRS38R-VARIANT").text = sVariant 'STANDARD
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/txt%PATH").text = sFilePath
session.findById("wnd[0]/usr/txt%PATH").SetFocus
session.findById("wnd[0]/usr/txt%PATH").caretPosition = 64
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[1]/usr/chkRSAQDOWN-COLUMN").Selected = True 'Header
session.findById("wnd[1]/usr/chkRSAQDOWN-COLUMN").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press 'Transfer
If bSAPFileExistsAtBegin = True Then
   If sSAPAction = "Append" Then
      session.findById("wnd[1]/tbar[0]/btn[11]").press 'Append
   Else
      session.findById("wnd[1]/tbar[0]/btn[0]").press 'Replace
   End If
End If
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press


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

===============+
Form_MaintainMasterData
===============

Private Sub btnRatesAndCostsHighLevelRefresh_Click() 'Used to be Command111 - Refresh Rates And Costs Detail
DoCmd.SetWarnings False
'DoCmd.OpenQuery "q990Misc010UupdateRatesAndCostsExistingRecords"
'DoCmd.OpenQuery "q010Init002AppendRatesAndCosts"
DoCmd.OpenQuery "q010Init144UpdateCurrentMillPMBasedOnHighLevel"
DoCmd.SetWarnings True
End Sub

Private Sub btnHelpDistPoint_Click() 'Used to be Command112 - Help

MsgBox "New Distribution Point Assignments:" & vbNewLine & vbNewLine & "(1) If only one mill makes the product, assign that mill" & vbNewLine & "(2) If both mills make the product, default to SVET" & vbNewLine & "(Nothing defaults to DC)" & vbNewLine & vbNewLine & "MPT logic shold initially assign all distribution points." & vbNewLine & "Contact the COE if there are missing records shown."


End Sub

Private Sub btnMillToDCMappingForceDC_Click() 'Used to be Command124 - If the table above is changed, click here to update the "ForceDC" tab
DoCmd.SetWarnings False
DoCmd.OpenQuery "q010Init018ClearUncheckedForceDC"
DoCmd.OpenQuery "q010Init019AppendForceDC"
DoCmd.SetWarnings True
DoCmd.Echo False
DoCmd.Close acForm, "MaintainMasterData"
DoCmd.OpenForm "MaintainMasterData"
Form_MaintainMasterData.MillToDCMapping.SetFocus
DoCmd.Maximize
DoCmd.Echo True
MsgBox "ForceDC Table Updated"
End Sub

Private Sub btnBack_Click() 'Used to be Command32 - Back
DoCmd.OpenForm "MainPage"
DoCmd.Close acForm, "MaintainMasterData"
End Sub

Private Sub Command110_Click()
On Error GoTo Err_Command110_Click


    Screen.PreviousControl.SetFocus
    DoCmd.FindNext

Exit_Command110_Click:
    Exit Sub

Err_Command110_Click:
    MsgBox Err.Description
    Resume Exit_Command110_Click
    
End Sub


Private Sub btnUpdateCustMaster_Click()
   basUtility.RunQueriesByPrefix "q010ImportData02"
   MsgBox "Update Complete"
End Sub

Private Sub btnUpdateDefaultSourcing_Click()
   basUtility.RunQueriesByPrefix "q010Init14"
   MsgBox "Default Sourcing Updated."
End Sub

Private Sub cmdUpdateSeasonalityValues_Click()
DoCmd.SetWarnings False
   UnwindTable "tblSeasonalityEdit", "tblSeasonalityWorking", 2, 12
   DoCmd.SetWarnings True
   MsgBox "Business/Cat1 Seasonality Values Updated"
End Sub

Private Sub CmdUpdNSRFcst_Click()
DoCmd.SetWarnings False
   
   DoCmd.Close acForm, "MaintainMasterdata"
   
   UnwindTable "tblSalesHistNSRPricesEdit", "tblSalesHistNSRPricesWorking", 5, 13
   UnwindTable "tblNSRMissingRatesEdit", "tblNSRMissingRatesWorking", 5, 13
   
   DoCmd.SetWarnings False
   
   DoCmd.OpenQuery "q400NSR102UpdNSRPricesWorking"
   DoCmd.OpenQuery "q400NSR103UpdNSRPricesFromWorking"
   DoCmd.OpenQuery "q400NSR104UpdSalesHistNSRPrices"
   DoCmd.OpenQuery "q400Nsr106DelSalesHistNSRPricesEdit"
   DoCmd.OpenQuery "q400Nsr108AppTblNSRHistPricesEdit"
   DoCmd.OpenQuery "q400Nsr110DelNSRMissingRatesEdit"
   DoCmd.OpenQuery "q400NSR112AppMissingRatesEdit"
   
   DoCmd.SetWarnings True
   
   MsgBox "RSM/Business/Cat1/GSM NSR Forecast Updated This Planning Cycle Onwards"
   DoCmd.OpenForm "MaintainMasterdata"
   Forms![MaintainMasterdata]![NSRData].SetFocus
 
End Sub
===============+
Form_DataImport
===============

Private Sub btnStop_Click() 'Used to be Command59
   DoCmd.OpenForm "mainpage"
   DoCmd.Close acForm, "dataimport"
End Sub

Private Sub btnContinue_Click() 'Used to be Command67
    Dim timeStamp   As Date
    Dim proceed     As Boolean
    Dim filePath As String
    
    'MsgBox "Disabled for Demo Mode"
    DoCmd.OpenForm "cleanmessagesheet", acNormal
    DoCmd.Close acForm, "DataImport"
    
    'MsgBox "Testing - not all Master Data coming from live Majiq views."
    
    'timeStamp = Application.CurrentProject.Connection.Execute("SELECT StageTimeStamp FROM tbl003PlanningCycleStages WHERE StageID = '" & ENUM_PC_STAGES.pcStage_RSMFilesGenerated & "'").Fields("StageTimeStamp").value & ""
'    timeStamp = modMain.GetPlanningCycleStageDate(pcStage_RSMFilesGenerated)
'    If timeStamp = NULL_DATE Then
'        modMain.SetPlanningCycleStageDate pcStage_ImportNewData
'        proceed = True
'    Else
'        If Forms.Item("MainPage").chkPlanningCycleStagesLock.value = True Then
'            MsgBox """Import New Data"" is disallowed because RSM files have already been generated." & vbNewLine & vbNewLine & "If you still want to proceed, you must unlock the ""Planning Cycle Stages Lock"", on the ""Main Page"", and try the operation again.", vbExclamation, Application.CurrentProject.Name
'        Else
'            proceed = True
'        End If
'    End If

    proceed = modMain.CanRunPlanningCycleStage(pcStage_ImportNewData, "Import New Data")
    
    If proceed Then
        DoCmd.SetWarnings False
        
        ''DoCmd.RunSQL "Delete * from tbl004MasterData010ArticleID"
        'DoCmd.RunSQL "insert into tbl004MasterData010ArticleID select * from q002ImportData010PassARTICLE"
        '
        ''DoCmd.RunSQL "Delete * from tbl004MasterData012CustFinID"
        'DoCmd.RunSQL "insert into tbl004MasterData012CustFinID select * from q002ImportData012PassCUSTFIN"
        '
        ''DoCmd.RunSQL "Delete * from tbl004MasterData014CustInvID"
        'DoCmd.RunSQL "insert into tbl004MasterData014CustInvID select * from q002ImportData014PassCUSTINV"
        '
        'DoCmd.RunSQL "Delete * from tbl002RawData010APOForecast"
        'DoCmd.RunSQL "insert into tbl002RawData010APOForecast select * from q002ImportData020Consensus"
        '
        'DoCmd.RunSQL "Delete * from tbl004MasterData031HistoricalModes"
        'DoCmd.RunSQL "insert into tbl004MasterData031HistoricalModes select * from q002ImportData018TModeHistory02"
        '
        'DoCmd.OpenQuery "q002ImportData030CleanUpNullGSMOnProductMaster"
        
        
        modMain.Import_SalesHistory
        
        
        
        DoCmd.OpenForm "mainpage"
        
        DoCmd.Close acForm, "cleanmessagesheet"
        
        modMain.SetPlanningCycleStageDate pcStage_ImportNewData
        
        Forms("MainPage").UpdateMainFrmChkBoxes (5)

        
        Form_MainPage.InitializeMainFrmChkBoxesAfterImp
              
        MsgBox "Data Import Complete", vbInformation, Application.CurrentProject.Name
    Else
        DoCmd.OpenForm "mainpage"
        
        DoCmd.Close acForm, "cleanmessagesheet"
    End If

End Sub
===============+
Form_InitializeMappingWarning
===============
Option Compare Database

Private Sub btnContinue_Click() 'Used to be Command37
    Dim proceed     As Boolean
    'RKP/03-12-13
    'Moved to "InterfaceCustom"
    proceed = modMain.CanRunPlanningCycleStage(pcStage_InitializeData, "Initialize Data")
    If proceed Then
        basUtility.RunQueriesByPrefix "q010ImportData02"
        InitializeData "MainPage", "InitializeMappingWarning"
        modMain.SetPlanningCycleStageDate pcStage_InitializeData
        Forms("MainPage").UpdateMainFrmChkBoxes (6)
    Else
        
    End If

End Sub

Private Sub btnStop_Click() 'Used to be Command39
DoCmd.OpenForm "mainpage"
DoCmd.Close acForm, "initializemappingwarning"
End Sub


===============+
Form_NSRImport
===============

Private Sub btnStop_Click() 'Used to be Command59
DoCmd.OpenForm "mainpage"
DoCmd.Close acForm, "dataimport"
End Sub

Private Sub btnContinue_Click() 'Used to be Command67
    Dim timeStamp   As Date
    Dim proceed     As Boolean
    Dim filePath As String
    
    'MsgBox "Disabled for Demo Mode"
    DoCmd.OpenForm "cleanmessagesheet", acNormal
    DoCmd.Close acForm, "DataImport"
    
    'MsgBox "Testing - not all Master Data coming from live Majiq views."
    
    'timeStamp = Application.CurrentProject.Connection.Execute("SELECT StageTimeStamp FROM tbl003PlanningCycleStages WHERE StageID = '" & ENUM_PC_STAGES.pcStage_RSMFilesGenerated & "'").Fields("StageTimeStamp").value & ""
'    timeStamp = modMain.GetPlanningCycleStageDate(pcStage_RSMFilesGenerated)
'    If timeStamp = NULL_DATE Then
'        modMain.SetPlanningCycleStageDate pcStage_ImportNewData
'        proceed = True
'    Else
'        If Forms.Item("MainPage").chkPlanningCycleStagesLock.value = True Then
'            MsgBox """Import New Data"" is disallowed because RSM files have already been generated." & vbNewLine & vbNewLine & "If you still want to proceed, you must unlock the ""Planning Cycle Stages Lock"", on the ""Main Page"", and try the operation again.", vbExclamation, Application.CurrentProject.Name
'        Else
'            proceed = True
'        End If
'    End If
    
    proceed = modMain.CanRunPlanningCycleStage(pcStage_ImportNewData, "Import New Data")
    
    If proceed Then
        DoCmd.SetWarnings False
        
        ''DoCmd.RunSQL "Delete * from tbl004MasterData010ArticleID"
        'DoCmd.RunSQL "insert into tbl004MasterData010ArticleID select * from q002ImportData010PassARTICLE"
        '
        ''DoCmd.RunSQL "Delete * from tbl004MasterData012CustFinID"
        'DoCmd.RunSQL "insert into tbl004MasterData012CustFinID select * from q002ImportData012PassCUSTFIN"
        '
        ''DoCmd.RunSQL "Delete * from tbl004MasterData014CustInvID"
        'DoCmd.RunSQL "insert into tbl004MasterData014CustInvID select * from q002ImportData014PassCUSTINV"
        '
        'DoCmd.RunSQL "Delete * from tbl002RawData010APOForecast"
        'DoCmd.RunSQL "insert into tbl002RawData010APOForecast select * from q002ImportData020Consensus"
        '
        'DoCmd.RunSQL "Delete * from tbl004MasterData031HistoricalModes"
        'DoCmd.RunSQL "insert into tbl004MasterData031HistoricalModes select * from q002ImportData018TModeHistory02"
        '
        'DoCmd.OpenQuery "q002ImportData030CleanUpNullGSMOnProductMaster"
        
        
        modMain.Import_SalesHistory
        
        
        
        DoCmd.OpenForm "mainpage"
        
        DoCmd.Close acForm, "cleanmessagesheet"
        
        modMain.SetPlanningCycleStageDate pcStage_ImportNewData
        
        MsgBox "Data Import Complete", vbInformation, Application.CurrentProject.Name
    Else
        DoCmd.OpenForm "mainpage"
        
        DoCmd.Close acForm, "cleanmessagesheet"
    End If
End Sub

Private Sub BtnGetNsrData_Click_Click()
'//Anjali Importing and Processing NSR (Net Sales Revenue Price) data
   
   Dim RS As ADODB.Recordset
   Dim cTotNSRCount As Integer
   
   On Error GoTo Err_Handler
        DoCmd.SetWarnings False
        DoCmd.RunSQL "DELETE tblRawNSRData.* FROM tblRawNSRData"
        DoCmd.SetWarnings True
        filePath = FormatFilePath(basUtility.GetSetting("SalesHistoryFolder"))
        DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12, "tblRawNSRData", filePath & "NSRData.xlsx", True, "A:AA"
        Set RS = Application.CurrentProject.Connection.Execute("select * from q400NSR000RawFileDataCheck")
        cTotNSRCount = RS.Fields("TotNSR").value
        If cTotNSRCount = 0 Then
            MsgBox "Some Error in Loading NSR data"
        Else
            DoCmd.SetWarnings False
            RunQueriesByPrefix "q400NSR0"
            DoCmd.SetWarnings True
            
            Form_MainPage.UpdateMainFrmChkBoxes (7)
            MsgBox "Loaded Latest NSR data Successfully"
            
            DoCmd.Close acForm, "NSRImport"
            DoCmd.OpenForm "mainpage"
            DoCmd.Close acForm, "cleanmessagesheet"
        End If

      
Err_Handler:
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
End Sub

Private Sub btnNSRStop_Click()
DoCmd.OpenForm "mainpage"
DoCmd.Close acForm, "NSRimport"
End Sub

Private Sub btnUpdateNSRHistoryArchive_Click()
'anjali new code
 
   Dim sBaseFilePath              As String
   Dim sForecastTrendDBFolder     As String
   Dim sForecastTrendDBFile       As String




   Dim sExcelReportFile           As String
   Dim sForecastTrendReportFolder As String
   Dim sReportMasterFolder        As String
   Dim sRefreshCmd                As String


        
   sBaseFilePath = FormatFilePath(basUtility.GetSetting("BaseFilePath"))                     'usu C:\OPTMODELS\IND1\
                                                                                  '    OR \\sa53fs01\Secunderabad-A53\APPM S&OP\
   sForecastTrendDBFolder = FormatFilePath(basUtility.GetSetting("ForecastTrendDBFolder"))   'usu C:\OPTMODELS\IND1\S&OPTool\
                                                                                  '    OR \\sa53fs01\Secunderabad-A53\APPM S&OP\DataPrep\Archive\
   sForecastTrendDBFile = basUtility.GetSetting("ForecastTrendDBFileName")                   'usu IPIndiaS&OP_ForecastTrends.mdb
   sForecastTrendDBQuery = basUtility.GetSetting("ForecastTrendDBQueryName")                 'usu q110Trend910TrendReport


    
   sExcelReportFile = ""                                                          'set later, will be IPIndiaFA.xlsm
                                                                                  '    OR  IPIndiaFAnoGSM.xlsm OR IPIndiaDPTrend.xlsm
   sForecastTrendReportFolder = FormatFilePath(basUtility.GetSetting("ForecastTrendReportFolder"))    'Destination for Reports, usu C:\OPTMODELS\IND1\S&OPTool\Reports
                                                                                           '    OR \\sa53fs01\Secunderabad-A53\APPM S&OP\PlanningCycle_Current\
   sReportMasterFolder = sBaseFilePath & FormatFilePath(basUtility.GetSetting("ReportMasterFolder"))
    






    'Anjali code begins
    
    Dim q5 As String, q6 As String, q7 As String, q8 As String, q9 As String, q10 As String, q11 As String, q12 As String, q13 As String, q14 As String, q15 As String
    Dim sSQL As String
    'basUtility.RunQueriesByPrefix "q930Archive11"
    
    'Anjali code begins - replacing RunQueriesByPrefix "q930Archive11" (nine action queries)
    'file name and path is flexible in table setings for instance
    ' \\sa53fs01\Secunderabad-A53\APPM S&OP\DataPrep\Archive\IPIndiaS&OP_ForecastTrends.mdb
        
    
    'Refreshing 3 FcstAccuracy Reports
    MsgBox "If Any Master Planning Excel Report Is Open, Close It, Then Click OK To Proceed"
    DoCmd.SetWarnings False
               
   
    'query6 - q930Archive118ClearCustomerMaster

    q6 = "DELETE * " & vbCrLf & _
         "FROM tbl004MasterData010Customer IN '" & sForecastTrendDBFolder & sForecastTrendDBFile & "'; " & vbCrLf
    DoCmd.RunSQL (q6)

     
    'query7 - q930Archive118ClearMaterialMaster

    q7 = "DELETE * " & vbCrLf & _
         "FROM tbl004MasterData010Material IN '" & sForecastTrendDBFolder & sForecastTrendDBFile & "'; " & vbCrLf
    DoCmd.RunSQL (q7)

    
    'query8 - q930Archive119CustomerMaster

    q8 = "INSERT INTO tbl004MasterData010Customer ( Customer, CountryCode, CustName, City, PostalCode, Region, RegionZone, RSMRegion ) " & vbCrLf & _
         "IN '" & sForecastTrendDBFolder & sForecastTrendDBFile & "' " & vbCrLf & _
         "SELECT tbl004MasterData010Customer.Customer, tbl004MasterData010Customer.CountryCode, tbl004MasterData010Customer.CustName, " & vbCrLf & _
         "tbl004MasterData010Customer.City, tbl004MasterData010Customer.PostalCode, tbl004MasterData010Customer.Region, " & vbCrLf & _
         "tbl004MasterData010Customer.RegionZone, tbl004MasterData010Customer.RSMRegion" & vbCrLf & _
         "FROM tbl004MasterData010Customer;" & vbCrLf
    DoCmd.RunSQL (q8)
    

    'query9 - q930Archive119MaterialMaster  'ADDED PRODSTRATCODE (GOOD/BETTER/BEST)
    q9 = "INSERT INTO tbl004MasterData010Material ( Material, MaterialDesc, Cat1, Business, SSNSS, SR, Pack, ProdStratCode ) " & vbCrLf & _
        "IN '" & sForecastTrendDBFolder & sForecastTrendDBFile & "' " & vbCrLf & _
        "SELECT tbl004MasterData010Material.Material, tbl004MasterData010Material.MaterialDesc, tbl004MasterData010Material.Cat1, " & vbCrLf & _
        "tbl004MasterData010Material.Business, tbl004MasterData010Material.SSNSS, tbl004MasterData010Material.SR, tbl004MasterData010Material.Pack, " & vbCrLf & _
        "tbl004MasterData010Material.ProdStratCode " & vbCrLf & _
        "FROM tbl004MasterData010Material;" & vbCrLf
    DoCmd.RunSQL (q9)
    

    'Anjali - added query 10 - 15 to handle NSR archiving process
    
    'query10 - q930Archive130MkNSRMonths
    q10 = "SELECT tblNSRAllMonths.* INTO tblNSRAllMonths " & vbCrLf & _
         "IN '" & sForecastTrendDBFolder & sForecastTrendDBFile & "' " & vbCrLf & _
         "FROM tblNSRAllMonths;" & vbCrLf
    DoCmd.RunSQL (q10)
    
    'query12 - q930Archive114ClearSalesHist
    q12 = "DELETE DISTINCTROW tblNSRDataHist.*" & vbCrLf & _
         "FROM tblNSRDataHist INNER JOIN tblNSRAllMonths ON tblNSRDataHist.YearMonth = tblNSRAllMonths.[WorkingMonthM1]" & vbCrLf & _
         "IN '" & sForecastTrendDBFolder & sForecastTrendDBFile & "'; " & vbCrLf
    DoCmd.RunSQL (q12)
    
    'query13 - q930Archive135NSRDataHist
    q13 = "INSERT INTO tblNSRDataHist ( RSMRegion, Bus, Cat1, GSM, Material, YearMonth, Quantity, WtdAvgNSRperTon ) " & vbCrLf & _
         "IN '" & sForecastTrendDBFolder & sForecastTrendDBFile & "' " & vbCrLf & _
         "SELECT tblNSRDataMaterial.RSMRegion, tblNSRDataMaterial.Bus, tblNSRDataMaterial.Cat1, tblNSRDataMaterial.GSM, tblNSRDataMaterial.Material, tblNSRDataMaterial.YearMonth, " & vbCrLf & _
         "Sum(tblNSRDataMaterial.Quantity) AS SumOfQuantity, " & vbCrLf & _
         "Round(Sum([Quantity]*[NSRperTon])/Sum([Quantity]),0) AS WtdAvgNSRperTon" & vbCrLf & _
         "FROM tblNSRDataMaterial" & vbCrLf & _
         "GROUP BY tblNSRDataMaterial.RSMRegion, tblNSRDataMaterial.Bus, tblNSRDataMaterial.Cat1, " & vbCrLf & _
         "tblNSRDataMaterial.GSM, tblNSRDataMaterial.Material, tblNSRDataMaterial.YearMonth;" & vbCrLf
    DoCmd.RunSQL (q13)
    
    'query14 - q930Archive114ClearSalesHist
    q14 = "DELETE DISTINCTROW tblSalesHistNSRPricesArchive.*" & vbCrLf & _
         "FROM tblSalesHistNSRPricesArchive INNER JOIN tblNSRAllMonths ON tblSalesHistNSRPricesArchive.PlanningCycle = tblNSRAllMonths.[PlanningCycle]" & vbCrLf & _
         "IN '" & sForecastTrendDBFolder & sForecastTrendDBFile & "'; " & vbCrLf
    DoCmd.RunSQL (q14)

    
    'query15 - q930Archive116SalesHist
    q15 = "INSERT INTO tblSalesHistNSRPricesArchive " & vbCrLf & _
         "IN '" & sForecastTrendDBFolder & sForecastTrendDBFile & "' " & vbCrLf & _
         "SELECT tblSaleHistNSRPrices.* " & vbCrLf & _
         "FROM tblSaleHistNSRPrices INNER JOIN tbl001PlanningCycle " & vbCrLf & _
         "ON tblSaleHistNSRPrices.PlanningCycle = tbl001PlanningCycle.PlanningCycle;" & vbCrLf
    DoCmd.RunSQL (q15)

    
    'Anjali code ends


    'additions for Forecast Evolution report  'STM  2014-09-11
    
    DoCmd.SetWarnings False
    
    sSQL = "DELETE * " & vbCrLf & _
           "FROM trepForecastEvolutionReport IN '" & sForecastTrendDBFolder & sForecastTrendDBFile & "'; " & vbCrLf
    DoCmd.RunSQL (sSQL)
    
    
    sSQL = "INSERT INTO trepForecastEvolutionReport  " & vbCrLf & _
              "IN '" & sForecastTrendDBFolder & sForecastTrendDBFile & "' " & vbCrLf & _
           "SELECT q110Trend915ForecastEvolutionReport.* " & vbCrLf & _
           "FROM q110Trend915ForecastEvolutionReport " & "IN '" & sForecastTrendDBFolder & sForecastTrendDBFile & "' " & vbCrLf & _
           ";"

    DoCmd.RunSQL (sSQL)

    DoCmd.SetWarnings True

    Dim xlApp As Object
    
'       Anjali - commented as we only need DPtrend File
''''    Set xlApp = CreateObject("Excel.Application")
''''    xlApp.Visible = True
''''    sExcelReportFile = "IPIndiaFA.xlsm"
''''    xlApp.Workbooks.Open (sReportMasterFolder & sExcelReportFile)
''''    xlApp.displayalerts = False
''''    xlApp.Range("lkpCfgForecastTrendDBFolder").value = sForecastTrendDBFolder
''''    xlApp.Range("lkpCfgForecastTrendDBFile").value = sForecastTrendDBFile
''''    xlApp.Range("lkpCfgExcelReportFile").value = sExcelReportFile
''''    xlApp.Range("lkpCfgForecastTrendReportFolder").value = sForecastTrendReportFolder
''''    xlApp.Range("lkpCfgForecaseTrendDBQuery").value = sForecastTrendDBQuery
''''    xlApp.Run ("Refresh")
''''    xlApp.displayalerts = True
''''    Set xlApp = Nothing
''''
''''    Set xlApp = CreateObject("Excel.Application")
''''    xlApp.Visible = True
''''    sExcelReportFile = "IPIndiaFAnoGSM.xlsm"
''''    xlApp.Workbooks.Open (sReportMasterFolder & sExcelReportFile)
''''    xlApp.displayalerts = False
''''    xlApp.Range("lkpCfgForecastTrendDBFolder").value = sForecastTrendDBFolder
''''    xlApp.Range("lkpCfgForecastTrendDBFile").value = sForecastTrendDBFile
''''    xlApp.Range("lkpCfgExcelReportFile").value = sExcelReportFile
''''    xlApp.Range("lkpCfgForecastTrendReportFolder").value = sForecastTrendReportFolder
''''    xlApp.Range("lkpCfgForecaseTrendDBQuery").value = sForecastTrendDBQuery
''''    xlApp.Run ("Refresh")
''''    xlApp.displayalerts = True
''''    Set xlApp = Nothing

    
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = True
    sExcelReportFile = "IPIndiaDPTrend.xlsm"
    xlApp.Workbooks.Open (sReportMasterFolder & sExcelReportFile)
    xlApp.displayalerts = False
    xlApp.Range("lkpCfgForecastTrendDBFolder").value = sForecastTrendDBFolder
    xlApp.Range("lkpCfgForecastTrendDBFile").value = sForecastTrendDBFile
    xlApp.Range("lkpCfgExcelReportFile").value = sExcelReportFile
    xlApp.Range("lkpCfgForecastTrendReportFolder").value = sForecastTrendReportFolder
    xlApp.Range("lkpCfgForecaseTrendDBQuery").value = sForecastTrendDBQuery
    xlApp.Run ("Refresh")


    xlApp.displayalerts = True
    Set xlApp = Nothing


        
    MsgBox "New Reports Generated In " & sForecastTrendReportFolder & " Folder"


End Sub
===============+
Form_MainPage
===============

Private Sub btnAdjustForecastASF_Click()
'**********************************************
'Author  :  RKP
'Date/Ver:  05-03-13/V01
'Input   :
'Output  :
'Comments:
'**********************************************
    On Error GoTo Err_Handler

    DoCmd.OpenForm "fintChangesASF"

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

Private Sub btnAdjustForecastMP_Click()
'**********************************************
'Author  :  RKP
'Date/Ver:  05-03-13/V01
'Input   :
'Output  :
'Comments:
'**********************************************
    On Error GoTo Err_Handler

    DoCmd.OpenForm "fintChangesMP"

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

Private Sub btnAdjustForecastUDP_Click()
'**********************************************
'Author  :  RKP
'Date/Ver:  05-03-13/V01
'Input   :
'Output  :
'Comments:
'**********************************************
    On Error GoTo Err_Handler

    DoCmd.OpenForm "fintChangesUDP"

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

Private Sub btnApplyPriorRSMEdits_Click()
    Dim proceed     As Boolean
    
    proceed = modMain.CanRunPlanningCycleStage(pcStage_ApplyRSMEdits, "Apply RSM Edits")
    If proceed Then
        basUtility.RunQueriesByPrefix "q230DPRSM01"
        
        modMain.SetPlanningCycleStageDate pcStage_ApplyRSMEdits
        
        Forms("MainPage").UpdateMainFrmChkBoxes (13)
        Me.Refresh
        
        MsgBox "Prior RSM edits have been applied successfully!", vbInformation, Application.CurrentDb.Name
    End If
End Sub

Private Sub btnGenerateRSMFile_Click()
'**********************************************
'Author  :  RKP
'Date/Ver:  03-18-13/V01
'Input   :
'Output  :
'Comments:
'**********************************************
    On Error GoTo Err_Handler

    GenerateRSMFile
    Forms("MainPage").UpdateMainFrmChkBoxes (14)
    Me.Refresh
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



Private Sub BtnGetNsrData_Click()
DoCmd.OpenForm "NSRimport"
DoCmd.Close acForm, "mainpage"
End Sub

Private Sub btnImportRSMFile_Click()
    If MsgBox("This operation will:" & vbNewLine & vbNewLine & "Import updated sales forecasts from RSM's." & vbNewLine & "Do you want to continue?", vbOKCancel + vbQuestion, Application.CurrentProject.Name) = vbOK Then
        modMain.RSM_ImportSalesForecasts
        Forms("MainPage").UpdateMainFrmChkBoxes (15)
        Me.Refresh
    End If
End Sub

Private Sub btnInitializeData_Click()
    On Error GoTo Err_Handler
    
    Dim proceed     As Boolean
    
    proceed = modMain.CanRunPlanningCycleStage(pcStage_InitializeData, "Initialize Data")
    
    If proceed Then
        DoCmd.OpenForm "initializemappingwarning"
        DoCmd.Close acForm, "mainpage"
    Else
        
    End If
    
Err_Handler:
    If Err Then
        Resume Next
    End If
End Sub

Private Sub btnPlanningCycle_Click()
    Me.txtPlanningCycle.Enabled = True
End Sub

Private Sub btnReportDPByBusiness_Click()
    If MsgBox("This operation will:" & vbNewLine & vbNewLine & "Generate an Excel pivot table that provides a Demand Planning view by Business." & vbNewLine & vbNewLine & "Do you want to continue?", vbQuestion + vbOKCancel, Application.CurrentProject.Name) = vbOK Then
        GeneratePivotTable_Generic , , , 22
        Forms("MainPage").UpdateMainFrmChkBoxes (18)
        Me.Refresh
    End If
End Sub

Private Sub btnReportMPByMillMachine_Click()
    If MsgBox("This operation will:" & vbNewLine & vbNewLine & "Generate an Excel pivot table that provides a Master Planning view by Mill/Machine." & vbNewLine & vbNewLine & "Do you want to continue?", vbQuestion + vbOKCancel, Application.CurrentProject.Name) = vbOK Then
        GeneratePivotTable_Generic , , , 21
        Forms("MainPage").UpdateMainFrmChkBoxes (19)
        Me.Refresh
    End If
End Sub

Private Sub btnReportSalesHistory_Click()
    If MsgBox("This operation will:" & vbNewLine & vbNewLine & "Generate an Excel pivot table that provides a view of Sales History." & vbNewLine & vbNewLine & "Do you want to continue?", vbQuestion + vbOKCancel, Application.CurrentProject.Name) = vbOK Then
        GeneratePivotTable_Generic , , , 20
        Forms("MainPage").UpdateMainFrmChkBoxes (8)
        Me.Refresh
    End If
End Sub

Private Sub btnRunKNA1Script_Click()
    Dim timeStamp   As Date
    Dim proceed     As Boolean
    
'    proceed = False
'    timeStamp = modMain.GetPlanningCycleStageDate(pcStage_RSMFilesGenerated)
'    If timeStamp = NULL_DATE Then
'        proceed = True
'    Else
'        If Forms.Item("MainPage").chkPlanningCycleStagesLock.value = True Then
'            MsgBox """Run SAP Scripts"" is disallowed because RSM files have already been generated." & vbNewLine & vbNewLine & "If you still want to proceed, you must unlock the ""Planning Cycle Stages Lock"", on the ""Main Page"", and try the operation again.", vbExclamation, Application.CurrentProject.Name
'        Else
'            proceed = True
'        End If
'    End If

    proceed = modMain.CanRunPlanningCycleStage(pcStage_RunSAPScripts, "Run SAP Scripts")
    
    If proceed Then
        If MsgBox("This operation will:" & vbNewLine & vbNewLine & "Run KNA1 script to download Customer data from SAP." & vbNewLine & vbNewLine & "Do you want to continue?", vbOKCancel + vbQuestion) = vbOK Then
            'MsgBox "SAP Scripting has not yet been enabled." & vbNewLine & "This functionality will be turned on as soon as IT allows SAP Scripting." & vbNewLine & "Until then, you will have to manually download the data from SAP.", vbInformation, "IP India S&OP Tool"
            modMain.SAP_RunSAPKNA1Script
            modMain.SetPlanningCycleStageDate pcStage_RunSAPScripts
            UpdateMainFrmChkBoxes (4)
            Me.Refresh
        End If
    End If
End Sub

Private Sub btnRunSAPScripts_Click()
    Dim timeStamp   As Date
    Dim proceed     As Boolean
    
'    proceed = False
'    timeStamp = modMain.GetPlanningCycleStageDate(pcStage_RSMFilesGenerated)
'    If timeStamp = NULL_DATE Then
'        proceed = True
'    Else
'        If Forms.Item("MainPage").chkPlanningCycleStagesLock.value = True Then
'            MsgBox """Run SAP Scripts"" is disallowed because RSM files have already been generated." & vbNewLine & vbNewLine & "If you still want to proceed, you must unlock the ""Planning Cycle Stages Lock"", on the ""Main Page"", and try the operation again.", vbExclamation, Application.CurrentProject.Name
'        Else
'            proceed = True
'        End If
'    End If

    proceed = modMain.CanRunPlanningCycleStage(pcStage_RunSAPScripts, "Run SAP Scripts")
    
    If proceed Then
        If MsgBox("This operation will:" & vbNewLine & vbNewLine & "Run scripts to download Sales History data from SAP." & vbNewLine & vbNewLine & "Do you want to continue?", vbOKCancel + vbQuestion) = vbOK Then
            'MsgBox "SAP Scripting has not yet been enabled." & vbNewLine & "This functionality will be turned on as soon as IT allows SAP Scripting." & vbNewLine & "Until then, you will have to manually download the data from SAP.", vbInformation, "IP India S&OP Tool"
            modMain.SAP_RunSAPScripts
            modMain.SetPlanningCycleStageDate pcStage_RunSAPScripts
            UpdateMainFrmChkBoxes (3)
            Me.Refresh
        End If
    End If
End Sub

Private Sub btnUpdateForecastHistoryArchive_Click()
    'stm
    Dim sBaseFilePath              As String
    Dim sForecastTrendDBFolder     As String
    Dim sForecastTrendDBFile       As String
    Dim sForecastTrendDBQuery      As String
    Dim sForecastTrendDBQuery2     As String
    Dim sExcelReportFile           As String
    Dim sForecastTrendReportFolder As String
    Dim sReportMasterFolder        As String
    Dim sRefreshCmd                As String
    Dim i As Integer
        
    sBaseFilePath = FormatFilePath(basUtility.GetSetting("BaseFilePath"))                     'usu C:\OPTMODELS\IND1\
                                                                                   '    OR \\sa53fs01\Secunderabad-A53\APPM S&OP\
    sForecastTrendDBFolder = FormatFilePath(basUtility.GetSetting("ForecastTrendDBFolder"))   'usu C:\OPTMODELS\IND1\S&OPTool\
                                                                                   '    OR \\sa53fs01\Secunderabad-A53\APPM S&OP\DataPrep\Archive\
    sForecastTrendDBFile = basUtility.GetSetting("ForecastTrendDBFileName")                   'usu IPIndiaS&OP_ForecastTrends.mdb
    sForecastTrendDBQuery = basUtility.GetSetting("ForecastTrendDBQueryName")                 'usu q110Trend910TrendReport
    sForecastTrendDBQuery2 = basUtility.GetSetting("ForecastTrendDBQueryName2")                'usu q110Trend910TrendReportToDays
    
    sExcelReportFile = ""                                                          'set later, will be IPIndiaFA.xlsm
                                                                                   '    OR  IPIndiaFAnoGSM.xlsm OR IPIndiaDPTrend.xlsm
    sForecastTrendReportFolder = FormatFilePath(basUtility.GetSetting("ForecastTrendReportFolder"))    'Destination for Reports, usu C:\OPTMODELS\IND1\S&OPTool\Reports
                                                                                            '    OR \\sa53fs01\Secunderabad-A53\APPM S&OP\PlanningCycle_Current\
    sReportMasterFolder = sBaseFilePath & FormatFilePath(basUtility.GetSetting("ReportMasterFolder"))
    
    
    MsgBox "If Any Master Planning Excel Report Is Open, Close It, Then Click OK To Proceed"
    Dim xlApp As Object
    
    'Anjali code begins
    Dim q1 As String, q2 As String, q3 As String, q4 As String, q5 As String, q6 As String, q10 As String, q12 As String, q13 As String, q14 As String, q15 As String
    Dim sSQL As String
    'basUtility.RunQueriesByPrefix "q930Archive12"    'Anjali
    
    'Anjali code begins - replacing RunQueriesByPrefix "q930Archive12" (six action queries)
    'file name and path is flexible in table setings for instance
    '\\sa53fs01\Secunderabad-A53\APPM S&OP\DataPrep\Archive\IPIndiaS&OP_ForecastTrends.mdb
        
    DoCmd.SetWarnings False
        
    'query1 - q930Archive124ClearForecastTblMain
    q1 = "DELETE tbl001Main.*, tbl001Main.PlanningCycle " & vbCrLf & _
        "FROM tbl001Main IN '" & sForecastTrendDBFolder & sForecastTrendDBFile & "' " & vbCrLf & _
        "WHERE (((tbl001Main.PlanningCycle)=DLookUp('PlanningCycle','tbl001PlanningCycle'))); " & vbCrLf
    DoCmd.RunSQL (q1)
    
    'query2 - q930Archive126ForecastTblMain
    q2 = "INSERT INTO tbl001Main ( PlanningCycle, [Month], Business, Grade, GSM, Material, Pack, [Size], RSMRegion, " & vbCrLf & _
        "RegionZone, SoldTo, ShipTo, HistTons, StatTons, ASFTons, SalesTons, UDPTons, CDPTons, PlanTons, OrigConvPt, CurrConvPt, " & vbCrLf & _
        "OrigDistPt, CurrDistPt, OrigMill, CurrMill, OrigPM, CurrPM, LagValue ) " & vbCrLf & _
        "IN '" & sForecastTrendDBFolder & sForecastTrendDBFile & "' " & vbCrLf & _
        "SELECT tbl001Main.PlanningCycle, tbl001Main.Month, tbl001Main.Business, tbl001Main.Grade, tbl001Main.GSM, tbl001Main.Material, " & vbCrLf & _
        "tbl001Main.Pack, tbl001Main.Size, tbl001Main.RSMRegion, tbl001Main.RegionZone, tbl001Main.SoldTo, tbl001Main.ShipTo, " & vbCrLf & _
        "tbl001Main.HistTons, tbl001Main.StatTons, tbl001Main.ASFTons, tbl001Main.SalesTons, tbl001Main.UDPTons, tbl001Main.CDPTons, " & vbCrLf & _
        "tbl001Main.PlanTons, tbl001Main.OrigConvPt, tbl001Main.CurrConvPt, tbl001Main.OrigDistPt, tbl001Main.CurrDistPt, tbl001Main.OrigMill, " & vbCrLf & _
        "tbl001Main.CurrMill, tbl001Main.OrigPM, tbl001Main.CurrPM, IIf(Right([Month],2)-Right([PlanningCycle],2)<0,Right([Month],2)-Right([PlanningCycle],2)+12," & vbCrLf & _
        "Right([Month],2)-Right([PlanningCycle],2)) AS LagValue " & vbCrLf & _
        "FROM tbl001Main; " & vbCrLf
    DoCmd.RunSQL (q2)
    
    'query3 - q930Archive128ClearCustomerMaster
    q3 = "DELETE * " & vbCrLf & _
        "FROM tbl004MasterData010Customer IN '" & sForecastTrendDBFolder & sForecastTrendDBFile & "'; " & vbCrLf
    DoCmd.RunSQL (q3)
    
    'query4 - q930Archive128ClearMaterialMaster
    q4 = "DELETE * " & vbCrLf & _
        "FROM tbl004MasterData010Material IN '" & sForecastTrendDBFolder & sForecastTrendDBFile & "'; " & vbCrLf
    DoCmd.RunSQL (q4)
    
    'query5 - q930Archive129CustomerMaster
    q5 = "INSERT INTO tbl004MasterData010Customer ( Customer, CountryCode, CustName, City, PostalCode, Region, RegionZone, RSMRegion ) " & vbCrLf & _
        "IN '" & sForecastTrendDBFolder & sForecastTrendDBFile & "' " & vbCrLf & _
        "SELECT tbl004MasterData010Customer.Customer, tbl004MasterData010Customer.CountryCode, tbl004MasterData010Customer.CustName,  " & vbCrLf & _
        "tbl004MasterData010Customer.City, tbl004MasterData010Customer.PostalCode, tbl004MasterData010Customer.Region,  " & vbCrLf & _
        "tbl004MasterData010Customer.RegionZone, tbl004MasterData010Customer.RSMRegion " & vbCrLf & _
        "FROM tbl004MasterData010Customer; " & vbCrLf
    DoCmd.RunSQL (q5)
    
    'query6 - q930Archive129MaterialMaster  'ADDED PRODSTRAT CODE (GOOD/BETTER/BEST)
    q6 = "INSERT INTO tbl004MasterData010Material ( Material, MaterialDesc, Cat1, Business, SSNSS, SR, Pack, ProdStratCode ) " & vbCrLf & _
        "IN '" & sForecastTrendDBFolder & sForecastTrendDBFile & "' " & vbCrLf & _
        "SELECT tbl004MasterData010Material.Material, tbl004MasterData010Material.MaterialDesc, tbl004MasterData010Material.Cat1, " & vbCrLf & _
        "tbl004MasterData010Material.Business, tbl004MasterData010Material.SSNSS, tbl004MasterData010Material.SR, tbl004MasterData010Material.Pack, " & vbCrLf & _
        "tbl004MasterData010Material.ProdStratCode " & vbCrLf & _
        "FROM tbl004MasterData010Material;" & vbCrLf
    DoCmd.RunSQL (q6)
    
    
    'Anjali - added query 10 - 15 to handle NSR archiving process
    
    'query10 - q930Archive130MkNSRMonths
    q10 = "SELECT tblNSRAllMonths.* INTO tblNSRAllMonths " & vbCrLf & _
         "IN '" & sForecastTrendDBFolder & sForecastTrendDBFile & "' " & vbCrLf & _
         "FROM tblNSRAllMonths;" & vbCrLf
    DoCmd.RunSQL (q10)
    
    'query12 - q930Archive114ClearSalesHist
    q12 = "DELETE DISTINCTROW tblNSRDataHist.*" & vbCrLf & _
         "FROM tblNSRDataHist INNER JOIN tblNSRAllMonths ON tblNSRDataHist.YearMonth = tblNSRAllMonths.[WorkingMonthM1]" & vbCrLf & _
         "IN '" & sForecastTrendDBFolder & sForecastTrendDBFile & "'; " & vbCrLf
    DoCmd.RunSQL (q12)
    
    'query13 - q930Archive135NSRDataHist
    q13 = "INSERT INTO tblNSRDataHist ( RSMRegion, Bus, Cat1, GSM, Material, YearMonth, Quantity, WtdAvgNSRperTon ) " & vbCrLf & _
         "IN '" & sForecastTrendDBFolder & sForecastTrendDBFile & "' " & vbCrLf & _
         "SELECT tblNSRDataMaterial.RSMRegion, tblNSRDataMaterial.Bus, tblNSRDataMaterial.Cat1, tblNSRDataMaterial.GSM, tblNSRDataMaterial.Material, tblNSRDataMaterial.YearMonth, " & vbCrLf & _
         "Sum(tblNSRDataMaterial.Quantity) AS SumOfQuantity, " & vbCrLf & _
         "Round(Sum([Quantity]*[NSRperTon])/Sum([Quantity]),0) AS WtdAvgNSRperTon" & vbCrLf & _
         "FROM tblNSRDataMaterial" & vbCrLf & _
         "GROUP BY tblNSRDataMaterial.RSMRegion, tblNSRDataMaterial.Bus, tblNSRDataMaterial.Cat1, " & vbCrLf & _
         "tblNSRDataMaterial.GSM, tblNSRDataMaterial.Material, tblNSRDataMaterial.YearMonth;" & vbCrLf
    DoCmd.RunSQL (q13)
    
    'query14 - q930Archive114ClearSalesHist
    q14 = "DELETE DISTINCTROW tblSalesHistNSRPricesArchive.*" & vbCrLf & _
         "FROM tblSalesHistNSRPricesArchive INNER JOIN tblNSRAllMonths ON tblSalesHistNSRPricesArchive.PlanningCycle = tblNSRAllMonths.[PlanningCycle]" & vbCrLf & _
         "IN '" & sForecastTrendDBFolder & sForecastTrendDBFile & "'; " & vbCrLf
    DoCmd.RunSQL (q14)

    
    'query15 - q930Archive116SalesHist
    q15 = "INSERT INTO tblSalesHistNSRPricesArchive " & vbCrLf & _
         "IN '" & sForecastTrendDBFolder & sForecastTrendDBFile & "' " & vbCrLf & _
         "SELECT tblSaleHistNSRPrices.* " & vbCrLf & _
         "FROM tblSaleHistNSRPrices INNER JOIN tbl001PlanningCycle " & vbCrLf & _
         "ON tblSaleHistNSRPrices.PlanningCycle = tbl001PlanningCycle.PlanningCycle;" & vbCrLf
    DoCmd.RunSQL (q15)

    
    'Anjali code ends

    
    DoCmd.SetWarnings True
    
    'Anjali code ends
    
    
    'additions for Forecast Evolution report  'STM  2014-09-11
    
    DoCmd.SetWarnings False
    
    sSQL = "DELETE * " & vbCrLf & _
           "FROM trepForecastEvolutionReport IN '" & sForecastTrendDBFolder & sForecastTrendDBFile & "'; " & vbCrLf
    DoCmd.RunSQL (sSQL)
    
    
    sSQL = "INSERT INTO trepForecastEvolutionReport  " & vbCrLf & _
              "IN '" & sForecastTrendDBFolder & sForecastTrendDBFile & "' " & vbCrLf & _
           "SELECT q110Trend915ForecastEvolutionReport.* " & vbCrLf & _
           "FROM q110Trend915ForecastEvolutionReport " & "IN '" & sForecastTrendDBFolder & sForecastTrendDBFile & "' " & vbCrLf & _
           ";"

    DoCmd.RunSQL (sSQL)

    DoCmd.SetWarnings True


    
    
    
    Set xlApp = CreateObject("Excel.Application")
       
    '~~ change the connection in the master and do the saveas
    sExcelReportFile = "IPIndiaDPTrend.xlsm"
    xlApp.Visible = True
    xlApp.Workbooks.Open (sReportMasterFolder & sExcelReportFile)
    xlApp.displayalerts = False
    xlApp.Range("lkpCfgForecastTrendDBFolder").value = sForecastTrendDBFolder
    xlApp.Range("lkpCfgForecastTrendDBFile").value = sForecastTrendDBFile
    xlApp.Range("lkpCfgExcelReportFile").value = sExcelReportFile
    xlApp.Range("lkpCfgForecastTrendReportFolder").value = sForecastTrendReportFolder
    xlApp.Range("lkpCfgForecaseTrendDBQuery").value = sForecastTrendDBQuery
    xlApp.Run ("Refresh")
    'Anjali 03/04/2015 - added code for the new tab - TrendPerDay
        
''
''    xlApp.Range("lkpCfgForecaseTrendDBQuery").value = sForecastTrendDBQuery2
''    xlApp.Run ("Refresh")
    xlApp.displayalerts = True
    Set xlApp = Nothing
    '~~ done changing the connection in the master and do the saveas
    
    

''''    Set xlApp = CreateObject("Excel.Application")
''''
''''    '~~ change the connection in the master and do the saveas
''''    sExcelReportFile = "IPIndiaDPTrend.xlsm"
''''    xlApp.Visible = True
''''    xlApp.Workbooks.Open (sReportMasterFolder & sExcelReportFile)
''''    xlApp.displayalerts = False
''''    xlApp.Range("lkpCfgForecastTrendDBFolder").value = sForecastTrendDBFolder
''''    xlApp.Range("lkpCfgForecastTrendDBFile").value = sForecastTrendDBFile
''''    xlApp.Range("lkpCfgExcelReportFile").value = sExcelReportFile
''''    xlApp.Range("lkpCfgForecastTrendReportFolder").value = sForecastTrendReportFolder
''''    xlApp.Range("lkpCfgForecaseTrendDBQuery").value = sForecastTrendDBQuery
''''    xlApp.Run ("Refresh")
''''    xlApp.displayalerts = True
''''    Set xlApp = Nothing
    
    Forms("MainPage").UpdateMainFrmChkBoxes (20)
    Me.Refresh
    
  
    MsgBox "New DPTrend Report Generated In " & sForecastTrendReportFolder & " Folder"
End Sub

Private Sub btnUpdateInventory_Click() 'Used to be Command107
   MsgBox "Disabled for Demo Mode"
End Sub

Private Sub btnUpdateSalesHistoryArchive_Click()
   'stm
   Dim sBaseFilePath              As String
   Dim sForecastTrendDBFolder     As String
   Dim sForecastTrendDBFile       As String
   Dim sExcelReportFile           As String
   Dim sForecastTrendReportFolder As String
   Dim sReportMasterFolder        As String
   Dim sRefreshCmd                As String
        
   sBaseFilePath = FormatFilePath(basUtility.GetSetting("BaseFilePath"))                     'usu C:\OPTMODELS\IND1\
                                                                                  '    OR \\sa53fs01\Secunderabad-A53\APPM S&OP\
   sForecastTrendDBFolder = FormatFilePath(basUtility.GetSetting("ForecastTrendDBFolder"))   'usu C:\OPTMODELS\IND1\S&OPTool\
                                                                                  '    OR \\sa53fs01\Secunderabad-A53\APPM S&OP\DataPrep\Archive\
   sForecastTrendDBFile = basUtility.GetSetting("ForecastTrendDBFileName")                   'usu IPIndiaS&OP_ForecastTrends.mdb
   sForecastTrendDBQuery = basUtility.GetSetting("ForecastTrendDBQueryName")                 'usu q110Trend910TrendReport
    
   sExcelReportFile = ""                                                          'set later, will be IPIndiaFA.xlsm
                                                                                  '    OR  IPIndiaFAnoGSM.xlsm OR IPIndiaDPTrend.xlsm
   sForecastTrendReportFolder = FormatFilePath(basUtility.GetSetting("ForecastTrendReportFolder"))    'Destination for Reports, usu C:\OPTMODELS\IND1\S&OPTool\Reports
                                                                                           '    OR \\sa53fs01\Secunderabad-A53\APPM S&OP\PlanningCycle_Current\
   sReportMasterFolder = sBaseFilePath & FormatFilePath(basUtility.GetSetting("ReportMasterFolder"))
    
    'Anjali code begins
    Dim q1 As String
    Dim q2 As String, q3 As String, q4 As String, q5 As String, q6 As String, q7 As String, q8 As String, q9 As String, q10 As String, q11 As String, q12 As String, q13 As String, q14 As String, q15 As String
    Dim sSQL As String
    'basUtility.RunQueriesByPrefix "q930Archive11"
    
    'Anjali code begins - replacing RunQueriesByPrefix "q930Archive11" (nine action queries)
    'file name and path is flexible in table setings for instance
    ' \\sa53fs01\Secunderabad-A53\APPM S&OP\DataPrep\Archive\IPIndiaS&OP_ForecastTrends.mdb
        
    
    'Refreshing 3 FcstAccuracy Reports
    MsgBox "If Any Master Planning Excel Report Is Open, Close It, Then Click OK To Proceed"
    DoCmd.SetWarnings False
               
    'query1 - q930Archive110ClearSalesHistMonths
    q1 = "DELETE * FROM tbl002SalesHistMonths IN '" & sForecastTrendDBFolder & sForecastTrendDBFile & "'; " & vbCrLf
    DoCmd.RunSQL (q1)
    
    'query2 - q930Archive112SalesHistMonths
    q2 = "INSERT INTO tbl002SalesHistMonths ( [Month] ) IN '" & sForecastTrendDBFolder & sForecastTrendDBFile & "' " & vbCrLf & _
         "SELECT tbl002RawData010SalesHistoryClean.Month" & vbCrLf & _
         "FROM tbl002RawData010SalesHistoryClean" & vbCrLf & _
         "GROUP BY tbl002RawData010SalesHistoryClean.Month;" & vbCrLf
    DoCmd.RunSQL (q2)
    
    'query3 - q930Archive114ClearSalesHist
    q3 = "DELETE tbl002SalesHist.*" & vbCrLf & _
         "FROM tbl002SalesHist INNER JOIN tbl002SalesHistMonths ON tbl002SalesHist.Month = tbl002SalesHistMonths.Month" & vbCrLf & _
         "IN '" & sForecastTrendDBFolder & sForecastTrendDBFile & "'; " & vbCrLf
    DoCmd.RunSQL (q3)
    
    'query4 - q930Archive116SalesHist
    q4 = "INSERT INTO tbl002SalesHist ( [Month], SoldTo, ShipTo, Material, GSM, [Size], Tons, PM )" & vbCrLf & _
         "IN '" & sForecastTrendDBFolder & sForecastTrendDBFile & "' " & vbCrLf & _
         "SELECT tbl002RawData010SalesHistoryClean.Month, tbl002RawData010SalesHistoryClean.SoldTo, " & vbCrLf & _
         "tbl002RawData010SalesHistoryClean.ShipTo, tbl002RawData010SalesHistoryClean.Material, tbl002RawData010SalesHistoryClean.GSM, " & vbCrLf & _
         "tbl002RawData010SalesHistoryClean.Size, Sum(tbl002RawData010SalesHistoryClean.Tons) AS Tons, tbl004MasterData002MachineInfo.PMCode" & vbCrLf & _
         "FROM tbl002RawData010SalesHistoryClean, tbl004MasterData002MachineInfo" & vbCrLf & _
         "WHERE (((tbl004MasterData002MachineInfo.MatlCode) = Left([tbl002RawData010SalesHistoryClean].[Material], 1)))" & vbCrLf & _
         "GROUP BY tbl002RawData010SalesHistoryClean.Month, tbl002RawData010SalesHistoryClean.SoldTo, " & vbCrLf & _
         "tbl002RawData010SalesHistoryClean.ShipTo, tbl002RawData010SalesHistoryClean.Material, " & vbCrLf & _
         "tbl002RawData010SalesHistoryClean.GSM, tbl002RawData010SalesHistoryClean.Size, tbl004MasterData002MachineInfo.PMCode;" & vbCrLf
    DoCmd.RunSQL (q4)
    
    'no query5, the fifth one was a _TEST and was non-action qry, just a SELECT
    
    'query6 - q930Archive118ClearCustomerMaster
    q6 = "DELETE * " & vbCrLf & _
         "FROM tbl004MasterData010Customer IN '" & sForecastTrendDBFolder & sForecastTrendDBFile & "'; " & vbCrLf
    DoCmd.RunSQL (q6)
     
    'query7 - q930Archive118ClearMaterialMaster
    q7 = "DELETE * " & vbCrLf & _
         "FROM tbl004MasterData010Material IN '" & sForecastTrendDBFolder & sForecastTrendDBFile & "'; " & vbCrLf
    DoCmd.RunSQL (q7)
    
    'query8 - q930Archive119CustomerMaster
    q8 = "INSERT INTO tbl004MasterData010Customer ( Customer, CountryCode, CustName, City, PostalCode, Region, RegionZone, RSMRegion ) " & vbCrLf & _
         "IN '" & sForecastTrendDBFolder & sForecastTrendDBFile & "' " & vbCrLf & _
         "SELECT tbl004MasterData010Customer.Customer, tbl004MasterData010Customer.CountryCode, tbl004MasterData010Customer.CustName, " & vbCrLf & _
         "tbl004MasterData010Customer.City, tbl004MasterData010Customer.PostalCode, tbl004MasterData010Customer.Region, " & vbCrLf & _
         "tbl004MasterData010Customer.RegionZone, tbl004MasterData010Customer.RSMRegion" & vbCrLf & _
         "FROM tbl004MasterData010Customer;" & vbCrLf
    DoCmd.RunSQL (q8)
    
    'query9 - q930Archive119MaterialMaster  'ADDED PRODSTRATCODE (GOOD/BETTER/BEST)
    q9 = "INSERT INTO tbl004MasterData010Material ( Material, MaterialDesc, Cat1, Business, SSNSS, SR, Pack, ProdStratCode ) " & vbCrLf & _
        "IN '" & sForecastTrendDBFolder & sForecastTrendDBFile & "' " & vbCrLf & _
        "SELECT tbl004MasterData010Material.Material, tbl004MasterData010Material.MaterialDesc, tbl004MasterData010Material.Cat1, " & vbCrLf & _
        "tbl004MasterData010Material.Business, tbl004MasterData010Material.SSNSS, tbl004MasterData010Material.SR, tbl004MasterData010Material.Pack, " & vbCrLf & _
        "tbl004MasterData010Material.ProdStratCode " & vbCrLf & _
        "FROM tbl004MasterData010Material;" & vbCrLf
    DoCmd.RunSQL (q9)
    
    'Anjali - added query 10 - 15 to handle NSR archiving process
    
    'query10 - q930Archive130MkNSRMonths
    q10 = "SELECT tblNSRAllMonths.* INTO tblNSRAllMonths " & vbCrLf & _
         "IN '" & sForecastTrendDBFolder & sForecastTrendDBFile & "' " & vbCrLf & _
         "FROM tblNSRAllMonths;" & vbCrLf
    DoCmd.RunSQL (q10)
    
    'query12 - q930Archive114ClearSalesHist
    q12 = "DELETE DISTINCTROW tblNSRDataHist.*" & vbCrLf & _
         "FROM tblNSRDataHist INNER JOIN tblNSRAllMonths ON tblNSRDataHist.YearMonth = tblNSRAllMonths.[WorkingMonthM1]" & vbCrLf & _
         "IN '" & sForecastTrendDBFolder & sForecastTrendDBFile & "'; " & vbCrLf
    DoCmd.RunSQL (q12)
    
    'query13 - q930Archive135NSRDataHist
    q13 = "INSERT INTO tblNSRDataHist ( RSMRegion, Bus, Cat1, GSM, Material, YearMonth, Quantity, WtdAvgNSRperTon ) " & vbCrLf & _
         "IN '" & sForecastTrendDBFolder & sForecastTrendDBFile & "' " & vbCrLf & _
         "SELECT tblNSRDataMaterial.RSMRegion, tblNSRDataMaterial.Bus, tblNSRDataMaterial.Cat1, tblNSRDataMaterial.GSM, tblNSRDataMaterial.Material, tblNSRDataMaterial.YearMonth, " & vbCrLf & _
         "Sum(tblNSRDataMaterial.Quantity) AS SumOfQuantity, " & vbCrLf & _
         "Round(Sum([Quantity]*[NSRperTon])/Sum([Quantity]),0) AS WtdAvgNSRperTon" & vbCrLf & _
         "FROM tblNSRDataMaterial" & vbCrLf & _
         "GROUP BY tblNSRDataMaterial.RSMRegion, tblNSRDataMaterial.Bus, tblNSRDataMaterial.Cat1, " & vbCrLf & _
         "tblNSRDataMaterial.GSM, tblNSRDataMaterial.Material, tblNSRDataMaterial.YearMonth;" & vbCrLf
    DoCmd.RunSQL (q13)
    
    'query14 - q930Archive114ClearSalesHist
    q14 = "DELETE DISTINCTROW tblSalesHistNSRPricesArchive.*" & vbCrLf & _
         "FROM tblSalesHistNSRPricesArchive INNER JOIN tblNSRAllMonths ON tblSalesHistNSRPricesArchive.PlanningCycle = tblNSRAllMonths.[PlanningCycle]" & vbCrLf & _
         "IN '" & sForecastTrendDBFolder & sForecastTrendDBFile & "'; " & vbCrLf
    DoCmd.RunSQL (q14)
    
    'query15 - q930Archive116SalesHist
    q15 = "INSERT INTO tblSalesHistNSRPricesArchive " & vbCrLf & _
         "IN '" & sForecastTrendDBFolder & sForecastTrendDBFile & "' " & vbCrLf & _
         "SELECT tblSaleHistNSRPrices.* " & vbCrLf & _
         "FROM tblSaleHistNSRPrices INNER JOIN tbl001PlanningCycle " & vbCrLf & _
         "ON tblSaleHistNSRPrices.PlanningCycle = tbl001PlanningCycle.PlanningCycle;" & vbCrLf
    DoCmd.RunSQL (q15)
    
    'Anjali code ends

    'additions for Forecast Evolution report  'STM  2014-09-11
    
    DoCmd.SetWarnings False
    
    sSQL = "DELETE * " & vbCrLf & _
           "FROM trepForecastEvolutionReport IN '" & sForecastTrendDBFolder & sForecastTrendDBFile & "'; " & vbCrLf
    DoCmd.RunSQL (sSQL)
    
    
    sSQL = "INSERT INTO trepForecastEvolutionReport  " & vbCrLf & _
              "IN '" & sForecastTrendDBFolder & sForecastTrendDBFile & "' " & vbCrLf & _
           "SELECT q110Trend915ForecastEvolutionReport.* " & vbCrLf & _
           "FROM q110Trend915ForecastEvolutionReport " & "IN '" & sForecastTrendDBFolder & sForecastTrendDBFile & "' " & vbCrLf & _
           ";"

    DoCmd.RunSQL (sSQL)

    DoCmd.SetWarnings True

    Dim xlApp As Object
    
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = True
    sExcelReportFile = "IPIndiaFA.xlsm"
    xlApp.Workbooks.Open (sReportMasterFolder & sExcelReportFile)
    xlApp.displayalerts = False
    xlApp.Range("lkpCfgForecastTrendDBFolder").value = sForecastTrendDBFolder
    xlApp.Range("lkpCfgForecastTrendDBFile").value = sForecastTrendDBFile
    xlApp.Range("lkpCfgExcelReportFile").value = sExcelReportFile
    xlApp.Range("lkpCfgForecastTrendReportFolder").value = sForecastTrendReportFolder
    xlApp.Range("lkpCfgForecaseTrendDBQuery").value = sForecastTrendDBQuery
    xlApp.Run ("Refresh")
    xlApp.displayalerts = True
    Set xlApp = Nothing
    
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = True
    sExcelReportFile = "IPIndiaFAnoGSM.xlsm"
    xlApp.Workbooks.Open (sReportMasterFolder & sExcelReportFile)
    xlApp.displayalerts = False
    xlApp.Range("lkpCfgForecastTrendDBFolder").value = sForecastTrendDBFolder
    xlApp.Range("lkpCfgForecastTrendDBFile").value = sForecastTrendDBFile
    xlApp.Range("lkpCfgExcelReportFile").value = sExcelReportFile
    xlApp.Range("lkpCfgForecastTrendReportFolder").value = sForecastTrendReportFolder
    xlApp.Range("lkpCfgForecaseTrendDBQuery").value = sForecastTrendDBQuery
    xlApp.Run ("Refresh")
    xlApp.displayalerts = True
    Set xlApp = Nothing
    
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = True
    sExcelReportFile = "IPIndiaDPTrend.xlsm"
    xlApp.Workbooks.Open (sReportMasterFolder & sExcelReportFile)
    xlApp.displayalerts = False
    xlApp.Range("lkpCfgForecastTrendDBFolder").value = sForecastTrendDBFolder
    xlApp.Range("lkpCfgForecastTrendDBFile").value = sForecastTrendDBFile
    xlApp.Range("lkpCfgExcelReportFile").value = sExcelReportFile
    xlApp.Range("lkpCfgForecastTrendReportFolder").value = sForecastTrendReportFolder
    xlApp.Range("lkpCfgForecaseTrendDBQuery").value = sForecastTrendDBQuery
    xlApp.Run ("Refresh")
    xlApp.displayalerts = True
    Set xlApp = Nothing
    Forms("MainPage").UpdateMainFrmChkBoxes (9)
    Me.Refresh
    MsgBox "New Reports Generated In " & sForecastTrendReportFolder & " Folder"
    
End Sub

'Private Sub btnInitializeMapping_Click() 'Used to be Command37
'DoCmd.OpenForm "initializemappingwarning"
'DoCmd.Close acForm, "mainpage"
'End Sub

Private Sub btnWorkingMonth_Click() 'Used to be Command39

End Sub

Private Sub btnGoToReportsPage_Click() 'Used to be Command50 - Go To Reports Page
DoCmd.OpenForm "reportpage"
DoCmd.Close acForm, "mainpage"
End Sub

Private Sub btnMasterData_Click() 'Used to be Command51
DoCmd.OpenForm "maintainmasterdata"
DoCmd.Close acForm, "MainPage"
End Sub

Private Sub btnImportNewData_Click() 'Used to be Command54
'IMPORT DATA
   DoCmd.OpenForm "dataimport"
   DoCmd.Close acForm, "mainpage"
End Sub

Private Sub btnRestartWorkingMonth_Click() 'Used to be Command60
DoCmd.OpenForm "resetworkingmonthwarning"
DoCmd.Close acForm, "mainpage"
End Sub

Private Sub btnTransferToMP_Click() 'Used to be Command61 - Transfer to Master Planning

'If MsgBox("If you have saved this database with a new file name for the operational forecast then" & vbNewLine & "click OK to continue.", vbOKCancel) = vbOK Then

'    DoCmd.SetWarnings False
'    DoCmd.OpenQuery "q710OpFcst010ClearOpForecastMain"
'    DoCmd.OpenQuery "q710OpFcst012ImportToOpForecastMain"
'    DoCmd.OpenQuery "q710OpFcst016AppendOpForecastMain"
'    DoCmd.SetWarnings True
    
'    DoCmd.OpenForm "OpFcstMainPage"
    DoCmd.OpenForm "fintChanges"
    DoCmd.Close acForm, "MainPage"
    
'    DoCmd.ShowToolbar "Ribbon", acToolbarNo
    
'    HideNavigationPane
    
    'CurrentDb.Properties("StartupForm") = "OpFcstMainPage"
    CurrentDb.Properties("StartupForm") = "fintChanges"
'End If

'MsgBox "Operational Forecast Table Updated"

End Sub


Private Sub btnExcelReportSC_Click() 'Used to be Command64 - Create "All Data" Excel Pivot Table Report - SC
'MsgBox "Disabled for Demo Mode"
MsgBox "If Any Master Planning Excel Report Is Open, Close It, Then Click OK To Proceed"


DoCmd.SetWarnings False
DoCmd.RunSQL "SELECT q920ExcelReport010FullDataExtract.* INTO tblFullDataExtract IN '" & sBasePath & "JMAMPT\Reporting\JMAMPTReporting.mdb'FROM q920ExcelReport010FullDataExtract;"
DoCmd.RunSQL "SELECT q910Report015PMBalanceAllMonths.* INTO tblPMBalanceAllMonths IN '" & sBasePath & "JMAMPT\Reporting\JMAMPTReporting.mdb'FROM q910Report015PMBalanceAllMonths;"

DoCmd.SetWarnings True

Dim xlApp As Object
Set xlApp = CreateObject("Excel.Application")
xlApp.Visible = True
xlApp.Workbooks.Open (sBasePath & "JMAMPT\Reporting\ReportMasters\JMAMPT_AllDataMaster.xls")
xlApp.Run ("Refresh")
xlApp.displayalerts = False
xlApp.Quit
Set xlApp = Nothing

MsgBox "New Reports Generated In Folder 'Current Cycle'"
End Sub

Private Sub btnCreateReplenishmentFile_Click() 'Used to be Command66 - Create Replenishment File

End Sub

Private Sub btnHelp_Click() 'Used to be Command68 - Help (?)
   DoCmd.OpenForm "HelpMainPage"
End Sub



Private Sub btnOptModelInterface_Click() 'Used to be Command70 - OptModel Interface
   DoCmd.OpenForm "OptModelInterface"
   DoCmd.Close acForm, "MainPage"
End Sub

Private Sub btnBackupWorkingMonth_Click() 'Used to be Command71
   If MsgBox("You are about to overwrite your current backup." & vbNewLine & "Click OK to continue.", vbOKCancel) = vbOK Then
       DoCmd.SetWarnings False
       DoCmd.OpenQuery "q320Misc010ClearTblMainBackupWorkingMonth"
       DoCmd.OpenQuery "q320Misc012AppendTblMainBackupWorkingMonth"
       DoCmd.SetWarnings True
       If gBATCH_MODE Then
       Else
          MsgBox "Working Month Backed Up"
       End If
   End If

End Sub

Private Sub btnRestoreWorkingMonth_Click() 'Used to be Command72
   DoCmd.SetWarnings False
   DoCmd.OpenQuery "q320Misc014ClearTblMainWorkingMonth"
   DoCmd.OpenQuery "q320Misc016AppendTblMainWorkingMonthFromBackup"
   DoCmd.SetWarnings True
   MsgBox "Working Month Restored"
End Sub

Private Sub btnVolumeDashboard_Click() 'Used to be Command74
   DoCmd.OpenForm "fintvolumedashboard", acNormal
End Sub

Private Sub Command75_Click()
   DoCmd.SetWarnings False
   DoCmd.OpenQuery "qClear1SAPPrimaryMachine"
   DoCmd.OpenQuery "qAppend1SAPPrimaryMachine"
   DoCmd.OpenQuery "qUpdAllDaysOnConsolidatedInformation"
   DoCmd.OpenQuery "qUpdAllDaysOnbkpConsolidatedInformation"
   DoCmd.SetWarnings True
   If gBATCH_MODE Then
   Else
      MsgBox "Primary Machine Reset and Days Updated"
   End If
End Sub

Private Sub btnAdjustForecast_Click() 'Used to be Command76
   DoCmd.OpenForm "fintOpFcst"
End Sub

Private Sub btnFreightCostData_Click() 'Used to be Command85
   DoCmd.OpenForm "MaintainFreightCostData"
   DoCmd.Close acForm, "MainPage"
End Sub

Private Sub btnExcelReportPM_Click() 'Used to be CreateInvReport
   
   MsgBox "If Any MPT Excel Report Is Open, Close It, Then Click OK To Proceed"
   
   
   DoCmd.SetWarnings False
   DoCmd.RunSQL "SELECT q920ExcelReport010FullDataExtract.* INTO tblFullDataExtract IN '" & sBasePath & "PM Strategic\Reporting\JMAMPTReporting.mdb'FROM q920ExcelReport010FullDataExtract;"
   DoCmd.RunSQL "SELECT q910Report015PMBalanceAllMonths.* INTO tblPMBalanceAllMonths IN '" & sBasePath & "PM Strategic\Reporting\JMAMPTReporting.mdb'FROM q910Report015PMBalanceAllMonths;"
   
   DoCmd.SetWarnings True
   
   Dim xlApp As Object
   Set xlApp = CreateObject("Excel.Application")
   xlApp.Visible = True
   xlApp.Workbooks.Open (sBasePath & "PM Strategic\Reporting\ReportMasters\JMAMPT_AllDataMaster.xls")
   xlApp.Run ("Refresh")
   xlApp.displayalerts = False
   xlApp.Quit
   Set xlApp = Nothing
   
   MsgBox "New Reports Generated In Folder 'PM Strategic'"

End Sub

Private Sub chkPlanningCycleStagesLock_Click()
    If Me.chkPlanningCycleStagesLock.value Then
        Application.CurrentProject.Connection.Execute ("UPDATE tbl003PlanningCycleStages SET StageTimeStamp = #" & VBA.Now() & "# WHERE StageID = '" & ENUM_PC_STAGES.pcStage_Lock & "'")
    Else
        Application.CurrentProject.Connection.Execute ("UPDATE tbl003PlanningCycleStages SET StageTimeStamp = #" & NULL_DATE & "# WHERE StageID = '" & ENUM_PC_STAGES.pcStage_Lock & "'")
    End If
End Sub

Private Sub cmbPlanningCycle_BeforeUpdate(Cancel As Integer)
'    If MsgBox("Are you sure you want to change the Planning Cycle?" & vbNewLine & vbNewLine & Me.cmbPlanningCycle, vbOKCancel + vbInformation, Application.CurrentProject.Name) = vbOK Then
'    Else
'        Cancel = 1
'    End If
End Sub

Private Sub Cmd_addToMainClick_Click()
DoCmd.OpenForm "FrmAddToMain"
DoCmd.Close acForm, "mainpage"
End Sub

Private Sub cmdTblMain_Click()
   DoCmd.OpenForm "frmAddForecastRecords", acFormDS
End Sub


Private Sub CreateOldForm_Click()
   MsgBox "Disabled for Demo Mode"
   'MsgBox "If Any Master Planning Excel Report Is Open, Close It, Then Click OK To Proceed"
   '
   'DoCmd.SetWarnings False
   'DoCmd.RunSQL "SELECT qRepOldWorksheetFormExtract.* INTO tblOldFormReport IN '\\s02afs01\Lenox-275\CST\Supply Planning\Integrated Master Planning\ReportMasters\CBDMPReporting.mdb'FROM qRepOldWorksheetFormExtract;"
   'DoCmd.RunSQL "SELECT qRepMachineBalanceAllMonths.* INTO tblMachineBalance IN '\\s02afs01\Lenox-275\CST\Supply Planning\Integrated Master Planning\ReportMasters\CBDMPReporting.mdb'FROM qRepMachineBalanceAllMonths;"
   'DoCmd.RunSQL "SELECT qRepInvChange6Extract.* INTO tblInventoryReport IN '\\s02afs01\Lenox-275\CST\Supply Planning\Integrated Master Planning\ReportMasters\CBDMPReporting.mdb'FROM qRepInvChange6Extract;"
   'DoCmd.SetWarnings True
   '
   'Dim xlApp As Object
   'Set xlApp = CreateObject("Excel.Application")
   'xlApp.Visible = True
   'xlApp.Workbooks.Open ("\\s02afs01\Lenox-275\CST\Supply Planning\Integrated Master Planning\ReportMasters\OldFormBalance.xls")
   'xlApp.Run ("Refresh")
   'xlApp.displayalerts = False
   ''xlApp.Quit
   'Set xlApp = Nothing
   '
   ''MsgBox "New Report Generated In 'CurrentCycle' Folder"
End Sub

Private Sub Form_Open(Cancel As Integer)
  On Error GoTo Err_Handler
   Dim sDateParse1 As String
   Dim sDateParse2 As String
   Dim sDateParse3 As String
   Dim sDateParse4 As String
   Dim sDateParse5 As String
   Dim RS As ADODB.Recordset
   
   DoCmd.Maximize

   Set RS = Application.CurrentProject.Connection.Execute("SELECT DISTINCT [Month] FROM [tbl002ForecastMonths]")
   
   Set RS = Application.CurrentProject.Connection.Execute("select * from tbl001workingmonth")
   Me.cmbWorkingMonth.value = RS.Fields("workingmonth").value
   Set RS = Nothing
   
   Set RS = Application.CurrentProject.Connection.Execute("select * from tbl001planningcycle")
   Me.cmbPlanningCycle.value = RS.Fields("planningcycle").value
   Me.txtPlanningCycle.Enabled = True
   Me.txtPlanningCycle.SetFocus
   Me.txtPlanningCycle.text = RS.Fields("planningcycle").value
   Me.txtPlanningCycle.Enabled = False
   
   Set RS = Nothing
   
   sDateParse1 = Right(DLookup("YearMonth", "q400NSR000LatestHistMonth"), 2) & "/" & "01/" & Left(DLookup("YearMonth", "q400NSR000LatestHistMonth"), 4)
   sDateParse2 = VBA.DateAdd("m", 1, sDateParse1)
   sDateParse5 = VBA.DateAdd("m", 4, sDateParse1)
   sDateParse3 = "Last NSR updated in PC: " & Right(VBA.DateAdd("m", 1, sDateParse1), 4) & Format(Left(sDateParse2, (InStr(sDateParse2, "/") - 1)), "00")
   sDateParse4 = Right(VBA.DateAdd("m", 4, sDateParse1), 4) & Format(Left(sDateParse5, (InStr(sDateParse5, "/") - 1)), "00")
   If Me.cmbPlanningCycle.value = sDateParse4 Then
      Me.LblLastNSRMonth.Caption = "You must update last 3 Months of NSR data in this PC"
      Me.LblLastNSRMonth.FontBold = True
      Me.LblLastNSRMonth.ForeColor = RED
   Else
      Me.LblLastNSRMonth.Caption = sDateParse3
   End If
   
   
   Me.lblVersion.Caption = GetVersion
   Me.chkGenerateRSMRegionFiles.value = False
   
   If _
       modMain.GetPlanningCycleStageDate(pcStage_ImportNewData) <> NULL_DATE _
       Or _
       modMain.GetPlanningCycleStageDate(pcStage_InitializeData) <> NULL_DATE _
       Or _
       modMain.GetPlanningCycleStageDate(pcStage_RSMFilesGenerated) <> NULL_DATE _
       Or _
       modMain.GetPlanningCycleStageDate(pcStage_RSMFilesImported) <> NULL_DATE _
   Then
       Me.chkPlanningCycleStagesLock.value = True
       'modMain.SetPlanningCycleStageDate pcStage_Lock
       modMain.SetPlanningCycleStageDate pcStage_Lock, True
   Else
       Me.chkPlanningCycleStagesLock.value = False
   End If

Err_Handler:
    If Err Then
        Resume Next
    End If
End Sub

Private Sub cmbPlanningCycle_AfterUpdate()

Dim pc As String
Dim sfolder As String
Dim sTag As String
Dim filePath    As String

If VBA.IsNumeric(Me.cmbPlanningCycle.value) Then
    If VBA.Len(Me.cmbPlanningCycle.value) = 6 Then
        DoCmd.SetWarnings False
        DoCmd.RunSQL "UPDATE tbl001planningcycle SET tbl001planningcycle.planningcycle = (cmbPlanningCycle)"
        DoCmd.RunSQL "UPDATE tbl006lkpPlanningCycle SET tbl006lkpplanningcycle.planningcycle = (cmbPlanningCycle)"
        DoCmd.SetWarnings True
        
        'sTag = Left(cmbPlanningCycle, 4) & Format(Right(cmbPlanningCycle, 2), "00")
        'sFolder = sBasePath & "JMAMPT\" & sTag & " Cycle"
        '
        'If Len(Dir(sFolder, vbDirectory)) = 0 Then
        '    MkDir sFolder
        'End If
        
        filePath = FormatFilePath(basUtility.GetSetting("BaseFilePath")) & "PlanningCycle_" & Me.cmbPlanningCycle.value
        If VBA.Len(VBA.Dir(filePath, vbDirectory)) = 0 Then
            VBA.MkDir filePath
        End If
        
        
        MsgBox "Planning Cycle Changed.", vbInformation, Application.CurrentProject.Name
    End If
End If

End Sub

Private Sub cmbWorkingMonth_AfterUpdate()
   DoCmd.SetWarnings False
   DoCmd.RunSQL "UPDATE tbl001workingmonth SET tbl001workingmonth.workingmonth = (cmbWorkingMonth)"
   DoCmd.Echo False
   DoCmd.Close acForm, "MainPage"
   DoCmd.OpenForm "MainPage"
   DoCmd.Echo True
   DoCmd.SetWarnings True
   Forms("MainPage").UpdateMainFrmChkBoxes (2)
   Forms.MainPage.Refresh
   MsgBox "Working Month Changed"
End Sub

Private Sub lblShow_Click()
   'code to show the db stuff
   ShowNavigationPane
   ShowRibbon
   'Application.Forms("Form_HelpMainPage").btnApplyChanges_Click
End Sub

Private Sub lblShow_DblClick(Cancel As Integer)
   'code to HIDE the db stuff
   HideNavigationPane
   HideRibbon
   'btnApplyChanges_Click
End Sub

Private Sub txtPlanningCycle_LostFocus()
    On Error GoTo Err_Handler
    
    Dim pc As String
    Dim sfolder As String
    Dim sTag As String
    Dim filePath    As String
    Dim pcCurrent   As String
    Dim pcNew       As String
    Dim sql         As String
    
    pcCurrent = Application.CurrentProject.Connection.Execute("SELECT PlanningCycle FROM tbl001PlanningCycle").Fields("PlanningCycle").value
    
    Me.txtPlanningCycle.SetFocus
    pcNew = Me.txtPlanningCycle.text
    
    If pcCurrent <> pcNew Then
        If MsgBox("Are you sure you want to change the Planning Cycle?" & vbNewLine & vbNewLine & "This should be done ONLY during the beginning of a new Planning Cycle.", vbOKCancel + vbQuestion, Application.CurrentProject.Name) = vbOK Then
    
            If VBA.IsNumeric(pcNew) Then
                If VBA.Len(pcNew) = 6 Then
                    
                    If pcNew <> pcCurrent Then
                        DoCmd.SetWarnings False
                        
                        DoCmd.RunSQL "UPDATE tbl001planningcycle SET tbl001planningcycle.planningcycle = '" & pcNew & "'"
                        DoCmd.RunSQL "UPDATE tbl006lkpPlanningCycle SET tbl006lkpplanningcycle.planningcycle = '" & pcNew & "'"
                        
                        DoCmd.SetWarnings True
                        
                        'sTag = Left(cmbPlanningCycle, 4) & Format(Right(cmbPlanningCycle, 2), "00")
                        'sFolder = sBasePath & "JMAMPT\" & sTag & " Cycle"
                        '
                        'If Len(Dir(sFolder, vbDirectory)) = 0 Then
                        '    MkDir sFolder
                        'End If
                        
                        filePath = FormatFilePath(basUtility.GetSetting("BaseFilePath")) & "PlanningCycle_" & Me.txtPlanningCycle.value
                        If VBA.Len(VBA.Dir(filePath, vbDirectory)) = 0 Then
                            VBA.MkDir filePath
                        End If
                        
                        'tbl003PlanningCycleStages
                        sql = "UPDATE tbl003PlanningCycleStages SET PlanningCycle = '" & pcNew & "'"
                        Application.CurrentProject.Connection.Execute sql
                        
                        sql = "UPDATE tbl003PlanningCycleStages SET StageTimeStamp = NULL"
                        Application.CurrentProject.Connection.Execute sql
                        
                        sql = "UPDATE tbl003PlanningCycleStages SET StageCounter = 0"
                        Application.CurrentProject.Connection.Execute sql
                        
                        sql = "UPDATE tbl003PlanningCycleStages SET StageTimeStamp = #" & VBA.Now() & "# WHERE StageID = '" & ENUM_PC_STAGES.pcStage_PlanningCycleChanged & "'"
                        Application.CurrentProject.Connection.Execute sql
                        
                        sql = "UPDATE tbl003PlanningCycleStages SET StageCounter = StageCounter + 1 WHERE StageID = '" & ENUM_PC_STAGES.pcStage_PlanningCycleChanged & "'"
                        Application.CurrentProject.Connection.Execute sql
                        
                        InitializeMainFrmChkBoxes
                        Me.Refresh
                        
                        MsgBox "Planning Cycle Changed.", vbInformation, Application.CurrentProject.Name
                    End If
                End If
            End If
        Else
            Me.txtPlanningCycle.text = pcCurrent
        End If
    End If
    
Err_Handler:
    If Err Then
        MsgBox "An error occured while changing Planning Cycle." & vbNewLine & "Error:" & vbNewLine & Err.Number & " - " & Err.Description, vbExclamation, Application.CurrentProject.Name
    End If
    Me.txtPlanningCycle.Enabled = False
End Sub

Sub InitializeMainFrmChkBoxes()
   DoCmd.SetWarnings False
   DoCmd.OpenQuery "qryInitializemainFrmChkBoxes"
   DoCmd.SetWarnings True
End Sub

Public Function UpdateMainFrmChkBoxes(sChkBoxID As Integer)
   Dim sql As String
   sql = "UPDATE tblMainFrmChkBoxes SET tblMainFrmChkBoxes.ChkBoxValue = True WHERE (((tblMainFrmChkBoxes.ID) = " & sChkBoxID & " ))"
   Application.CurrentProject.Connection.Execute sql
End Function

Sub InitializeMainFrmChkBoxesAfterImp()
   DoCmd.SetWarnings False
   DoCmd.OpenQuery "qryInitializemainFrmChkBoxesAfterImp"
   DoCmd.SetWarnings True
End Sub
===============+
basProj
===============
Option Explicit
Option Compare Database

Public mlLastErr       As Long
Public msLastErr       As String
Public msErrSource     As String
Public msErrDesc       As String


Public Function CleanConflicts()
    DoCmd.SetWarnings False
    DoCmd.OpenQuery "q430ModelPaths030ConflictList"
    DoCmd.OpenQuery "q430ModelPaths032MarkConflicts"
    DoCmd.OpenQuery "q430ModelPaths034ClearConflicts"
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
    If i > 0 Then
       strMod = VBE.ActiveVBProject.VBComponents(mdl.Name).codemodule.Lines(1, i)
    End If
    ''and then write it to a file, first marking the start with
    ''some equal signs and the component name.
    f.writeline String(15, "=") & vbCrLf & mdl.Name _
        & vbCrLf & String(15, "=") & vbCrLf & strMod
Next

''Close eveything
f.Close
Set fs = Nothing
End Sub

