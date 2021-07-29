Attribute VB_Name = "Module1"
Sub Run_All()

Call AO_Refresh
Call SAP_FBL3N
Call Variance

MsgBox "Process Complete"

End Sub

Sub AO_Refresh()

Application.DisplayAlerts = False
Application.ScreenUpdating = False

On Error Resume Next
If Sheet2.AutoFilterMode Then Sheet2.ShowAllData
On Error GoTo 0

Sheet2.Select
Range("F4:H500000").ClearContents

Dim lResult As Long

DataSources = "DS_1"

Call Application.Run("SAPExecuteCommand", "Refresh", DataSources)

Call Application.Run("SAPSetRefreshBehaviour", "Off") '

Call Application.Run("SAPExecuteCommand", "PauseVariableSubmit", "On")

    'DS1
    lResult = Application.Run("SAPSetVariable", "0I_DAYS", Range("Beg_of_Month").Value & " - " & Range("End_of_Month").Value, "Input_String", "DS_1")

Call Application.Run("SAPExecuteCommand", "PauseVariableSubmit", "Off")

Call Application.Run("SAPSetRefreshBehaviour", "On")

Sheet2.Select
            Range("A4:A300000").Select
    Selection.TextToColumns Destination:=Range("A4"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True

            Range("D4:D300000").Select
    Selection.TextToColumns Destination:=Range("D4"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True

Range("A5").Select


End Sub

Sub SAP_FBL3N()

Application.DisplayAlerts = False
Application.ScreenUpdating = False

wb = ActiveWorkbook.Name

Application.Calculation = xlCalculationAutomatic

On Error Resume Next
If Sheet3.AutoFilterMode Then Sheet3.ShowAllData
On Error GoTo 0

Sheet3.Select
Range("M2:N500000").Clear
Range("L2").Select
Range(Selection, Selection.End(xlDown).End(xlToLeft)).ClearContents

On Error GoTo NotLoggedOnSAP:
  Set SapGuiAuto = GetObject("SAPGUI")  'Get the SAP GUI Scripting object
  Set SAPApp = SapGuiAuto.GetScriptingEngine 'Get the currently running SAP GUI
  Set SAPCon = SAPApp.Children(0) 'Get the first system that is currently connected
  Set session = SAPCon.Children(0) 'Get the first session (window) on that connection
  
      GoTo LoggedonSAP:
      
NotLoggedOnSAP:
    X = MsgBox("You are not logged on SAP.  Please log on and try again.", vbOKOnly, "Not Logged on SAP")
    Exit Sub
    
LoggedonSAP:
        On Error GoTo 0

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nfbl3n"
session.findById("wnd[0]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/txtV-LOW").Text = "GFS2_GLSALEVAL"
session.findById("wnd[1]/usr/txtENAME-LOW").Text = ""
session.findById("wnd[1]/usr/txtENAME-LOW").SetFocus
session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 0
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/ctxtSO_BUDAT-LOW").Text = Range("Beg_of_Month")
session.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").Text = Range("End_of_Month")
session.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").SetFocus
session.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").caretPosition = 10
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[2]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").Text = "C:\Users\" & Environ("Username") & "\Desktop\"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "test.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 8
session.findById("wnd[1]/tbar[0]/btn[11]").press

Workbooks.Open "C:\Users\" & Environ("Username") & "\Desktop\test.xls"

Range("C:C").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
Range("D:D").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
Range("A:B").Delete
Range("C:C").Delete

        Dim ns As Long, i As Long
 ns = Range("A1:A300000").SpecialCells(xlLastCell).Row
 For i = ns To 1 Step -1
    With Range("A1:A300000").Cells(i)
       If .Value = "Plnt" Then
             .EntireRow.Delete Shift:=xlUp
       End If
    End With
 Next i

Range("L1").Select
Range(Selection, Selection.End(xlDown).End(xlToLeft)).Copy

Windows(wb).Activate
Sheet3.Select
Range("A2").PasteSpecial xlPasteValues

Windows("test.xls").Close

Range("A2").Select

Application.Calculation = xlCalculationManual

End Sub

Sub Variance()

Application.DisplayAlerts = False
Application.ScreenUpdating = False

Application.Calculation = xlCalculationAutomatic

Sheet7.Visible = True

On Error Resume Next
If Sheet8.AutoFilterMode Then Sheet8.ShowAllData
On Error GoTo 0

Sheet8.Select
Range("D2:G2000").Clear
Range("A2:G2000").Interior.ColorIndex = xlNone

Sheet2.Select
Range("F4:H500000").Clear

Sheet3.Select
Range("M2:N500000").Clear

Sheet7.Range("A2:H2000").ClearContents

Sheet2.Select
            Range("A4:A300000").Select
    Selection.TextToColumns Destination:=Range("A4"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True

            Range("D4:D300000").Select
    Selection.TextToColumns Destination:=Range("D4"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True

Range("A4").Select
Range(Selection, Selection.End(xlDown)).Copy

Sheet7.Select
Range("A2").PasteSpecial xlPasteValues
Selection.RemoveDuplicates Columns:=Array(1)
Range("A2").Select
Range(Selection, Selection.End(xlDown)).Select

    Sheet7.Sort.SortFields.Clear
    Sheet7.Sort.SortFields.Add Key:=Range("A2"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With Sheet7.Sort
        .SetRange Selection
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

Range("B2") = "=SUMIF('CSR Data'!A:A,A2,'CSR Data'!E:E)"
Range("B2").Copy
Range("A2").Select
Range(Selection, Selection.End(xlDown)).Offset(0, 1).PasteSpecial xlPasteFormulas
Selection.Copy
Selection.PasteSpecial xlPasteValues
Range("A2").Select


Sheet3.Select
Range("A2").Select
Range(Selection, Selection.End(xlDown)).Copy

Sheet7.Select
Range("D2").PasteSpecial xlPasteValues
Selection.RemoveDuplicates Columns:=Array(1)
Range("D2").Select
Range(Selection, Selection.End(xlDown)).Select

    Sheet7.Sort.SortFields.Clear
    Sheet7.Sort.SortFields.Add Key:=Range("D2"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With Sheet7.Sort
        .SetRange Selection
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
Range("E2") = "=SUMIF('SAP Data'!A:A,D2,'SAP Data'!G:G)"
Range("E2").Copy
Range("D2").Select
Range(Selection, Selection.End(xlDown)).Offset(0, 1).PasteSpecial xlPasteFormulas
Selection.Copy
Selection.PasteSpecial xlPasteValues
Range("A2").Select

Range("B2").Select
Range(Selection, Selection.End(xlDown).End(xlToLeft)).Copy
Range("G2").PasteSpecial xlPasteValues

Range("E2").Select
Range(Selection, Selection.End(xlDown).End(xlToLeft)).Copy
Range("G2").End(xlDown).Offset(1, 0).PasteSpecial xlPasteValues

Range("G2").Select
Range(Selection, Selection.End(xlDown)).Copy

Sheet8.Select
Range("D2").PasteSpecial xlPasteValues
Selection.RemoveDuplicates Columns:=Array(1)
Range("D2").Select
Range(Selection, Selection.End(xlDown)).Select
    
    Sheet8.Sort.SortFields.Clear
    Sheet8.Sort.SortFields.Add Key:=Range("D2"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With Sheet8.Sort
        .SetRange Selection
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

Range("E2") = "=SUMIF(Holder!G:G,D2,Holder!H:H)"
Range("E2").Copy
Range("D2").Select
Range(Selection, Selection.End(xlDown)).Offset(0, 1).PasteSpecial xlPasteFormulas
Selection.Copy
Selection.PasteSpecial xlPasteValues
Range("E:E").Style = "Comma"

Range("D2").Select
Application.CutCopyMode = False

Sheet7.Visible = False

Application.Calculation = xlCalculationManual

End Sub

