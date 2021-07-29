Option Compare Database
Option Explicit
Private mlLastErr       As Long
Private msLastErr       As String
Private msErrSource     As String
Private msErrDesc       As String
'----------------------------------------------




Public Function GenerateRSMEditCustom(ByRef xlWB As Object)
'**********************************************
'Author  :  sean
'Date/Ver:  5.16.18
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
   Dim MyArray()       As Integer   ''for all the fields in range that want sum subtotals
   Dim i As Integer
   Dim j As Integer
   Dim k As Integer
   
   Const CLRchangedOrange = 49918
   Const CLRheaderBrown = 13163740
   Const CLRvaluesGreen = 8243375
   'grays
   Const CLRsubTotalGray = 12500670
   Const CLRGray1Light = 14606046      'rgb 222.222.222
   Const CLRGray2LightMed = 12632256   'rgb 192.192.192
   Const CLRGray3Medium = 10987431     'rgb 167.167.167
   Const CLRGray4MedDark = 9145227     'rgb 139.139.139
   Const CLRGray5Dark = 7303023        'rgb 111.111.111
   'blues
   Const CLRvaluesBlue = 16092160      'this 16092160 is abit too deep for edit cells
   Const CLRmediumBlue = 16738560
   Const CLRvaluesPaleBlue = 16441800
   Const CLRmildBlue = 16435300
   Const CLRcellEditBlue = 16110260
   Const CLRrichBlue = 13134085
   Const CLRmidnightBlue = 6299648
   Const CLRmachineHdrBlue = 15451035
   
   Const CLRblackoutDark = 8224125
   
   Const CLRlightCream = 13497855
   Const CLRredAsRed = 255
   Const CLRstarkWhite = -1
   
   wsName = "RSMPivot"
   
   ' =================================================== title info
   ' GCF RSM View (201805)
   ' Forecast (CDP)
   ' (ADMT)
   
   xlWB.Sheets(wsName).Activate
   xlWB.Sheets(wsName).Range("E1").select
   xlWB.Sheets(wsName).Range("E1").FormulaR1C1 = "GCF RSM View (" & xlWB.Worksheets("RSMData").Range("C2").value & ")"
   xlWB.Sheets(wsName).Range("E1").select
   xlWB.Sheets(wsName).Range("E1").Font.Bold = True
   
   ApplyFormat_Font_Generic xlWB.Sheets(wsName).Range("E1"), 108
   
   
   xlWB.Sheets(wsName).Activate
   xlWB.Sheets(wsName).Range("E2").select
   xlWB.Sheets(wsName).Range("E2").FormulaR1C1 = "Forecast (CDP)"
   xlWB.Sheets(wsName).Range("E2").select
   xlWB.Sheets(wsName).Range("E2").Font.Bold = True
   
   ApplyFormat_Font_Generic xlWB.Sheets(wsName).Range("E2"), 104
   
       
   xlWB.Sheets(wsName).Activate
   xlWB.Sheets(wsName).Range("E3").select
   xlWB.Sheets(wsName).Range("E3").FormulaR1C1 = "(ADMT)"
   xlWB.Sheets(wsName).Range("E3").select
   xlWB.Sheets(wsName).Range("E3").Font.Bold = True
   
   ApplyFormat_Font_Generic xlWB.Sheets(wsName).Range("E3"), 104
   
   
   ' =================================================== repeat item labels before copy
   xlWB.Sheets(wsName).PivotTables("RSMPivot").PivotFields("RSMBalanceGrade").RepeatLabels = True
   xlWB.Sheets(wsName).PivotTables("RSMPivot").PivotFields("RSMParentHdr").RepeatLabels = True
   xlWB.Sheets(wsName).PivotTables("RSMPivot").PivotFields("RSMParentLocDtl").RepeatLabels = True
   xlWB.Sheets(wsName).PivotTables("RSMPivot").PivotFields("RSMRegion").RepeatLabels = True
   
   
   'kill the grand total at the bottom
   xlWB.Sheets(wsName).PivotTables("RSMPivot").RowGrand = False
   
   
   ' =================================================== copy the pivot as values to 'RSMEdit' tab
   'xlWB.Worksheets("RSMPivot").Select
   xlWB.Worksheets(wsName).cells.select
   xlWB.Application.selection.Copy
   
   wsName = "RSMEdit"
   
   Worksheet_Action xlWB, actionCreate, wsName
   
   xlWB.Worksheets(wsName).select
   xlWB.Application.selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
   xlWB.Worksheets(wsName).Range("A1").select
   
   Set ws = xlWB.Worksheets(wsName)
   'ws.Name = wsName
   
   'ws.Range("A1:Z100000").value = ws.Range("A1:Z100000").value
   lastRowEdit = 600  ''GetLastColumnInRSMEdit(xlWB)  'UN-HARDWIRE
   
   'format title and header
   ApplyFormat_Font_Generic xlWB.Sheets(wsName).Range("E1"), 108
   ApplyFormat_Font_Generic xlWB.Sheets(wsName).Range("E2"), 104
   ApplyFormat_Font_Generic xlWB.Sheets(wsName).Range("E3"), 104
   
   
   ' =================================================== format the numbers
   ws.cells.select
   xlWB.Application.selection.Style = "Comma"
   xlWB.Application.selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
   
   ''ws.Columns("D:D").Select
   ''xlWB.Application.Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
   
   
   
   ''xlWB.Sheets(wsName).cells.Select
   ''xlWB.Sheets(wsName).cells.EntireColumn.AutoFit
    
    
   ' =================================================== checkerboard to indicate WIP
    xlWB.Worksheets(wsName).select
    
    xlWB.Sheets(wsName).Range("A1, B2, A3").select
    xlWB.Application.selection.Interior.Color = CLRmidnightBlue
    
    xlWB.Sheets(wsName).Range("B1, A2, B3").select
    xlWB.Application.selection.Interior.Color = CLRlightCream
    
    'copy checkerboard down about 40 rows
    xlWB.Sheets(wsName).Range("A2:B3").select
    xlWB.Application.selection.Copy
    xlWB.Sheets(wsName).Range("A2:B41").select
    xlWB.Application.selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    xlWB.Sheets(wsName).Columns("A:B").select
    xlWB.Application.selection.ColumnWidth = 1
    xlWB.Sheets(wsName).Range("A1").select
    
    
    ' =================================================== format and hide columns
    ''Sheets("RSMEdit").Select
    xlWB.Sheets(wsName).Columns("A:B").ColumnWidth = 0.8
    xlWB.Sheets(wsName).Columns("C:D").ColumnWidth = 13.57
    xlWB.Sheets(wsName).Columns("E:E").ColumnWidth = 26.57     'parent loc detail
    xlWB.Sheets(wsName).Columns("F:F").ColumnWidth = 10        'rsm region
    xlWB.Sheets(wsName).Columns("G:G").select                  'base key
    xlWB.Application.selection.EntireColumn.Hidden = True
    xlWB.Sheets(wsName).Columns("H:CN").select
    '''''''''ActiveWindow.LargeScroll ToRight:=-4
    xlWB.Application.selection.ColumnWidth = 9.43
    xlWB.Sheets(wsName).Range("A1").select
    
    
    ' =================================================== format a month's headers, copy all the way over, and hide extra header rows and freeze panes
   xlWB.Sheets(wsName).Range("H1").select
   xlWB.Sheets(wsName).Range("H1").FormulaR1C1 = "=IF(R[7]C="""","""",R[7]C)"
   xlWB.Sheets(wsName).Range("H2").select
   xlWB.Sheets(wsName).Range("H2").FormulaR1C1 = "=IF(R[8]C="""","""",R[8]C)"
   xlWB.Sheets(wsName).Range("H3").select
   xlWB.Sheets(wsName).Range("H3").FormulaR1C1 = "=IF(R[9]C="""","""",R[9]C)"
   xlWB.Sheets(wsName).Range("H1:H3").select
   xlWB.Application.selection.Copy
   xlWB.Sheets(wsName).Range("H1:T3").select
   xlWB.Application.selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
   '' activeSheet.Paste
    
    
   ''put in the total stuff
   xlWB.Sheets(wsName).Range("U3").select
   xlWB.Sheets(wsName).Range("U3").FormulaR1C1 = "TOTAL"
   xlWB.Sheets(wsName).Range("H3:U3").select
   With xlWB.Application.selection
      .HorizontalAlignment = xlCenter
      .VerticalAlignment = xlBottom
      .WrapText = False
      .Orientation = 0
      .AddIndent = False
      .IndentLevel = 0
      .ShrinkToFit = False
      .ReadingOrder = xlContext
      .MergeCells = False
    End With
    
    xlWB.Application.selection.Font.Bold = True
    xlWB.Sheets(wsName).Range("H2:U2").select
    
    With xlWB.Application.selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
   xlWB.Sheets(wsName).Range("I2:U2").select
   With xlWB.Application.selection
      .ClearContents
   End With
   
   xlWB.Sheets(wsName).Range("H2:U2").select
   ''xlWB.Sheets(wsName).Range("H2").Activate   'U2 instead??   need at all??
   With xlWB.Application.selection
      .HorizontalAlignment = xlCenter
      .VerticalAlignment = xlBottom
      .WrapText = False
      .Orientation = 0
      .AddIndent = False
      .IndentLevel = 0
      .ShrinkToFit = False
      .ReadingOrder = xlContext
      .MergeCells = False
   End With
   xlWB.Application.selection.Merge
   xlWB.Application.selection.Font.Bold = True
    
   'machine name headers
   xlWB.Sheets(wsName).Range("H3:U3").select
   xlWB.Application.selection.Interior.Color = CLRmachineHdrBlue
   '''xlWB.Sheets(wsName).Range("H3:U3").select.Interior.Color = CLRmachineHdrBlue    'nope

    
   xlWB.Sheets(wsName).Range("H2:U3").select
   xlWB.Application.selection.Borders(xlDiagonalDown).LineStyle = xlNone
   xlWB.Application.selection.Borders(xlDiagonalUp).LineStyle = xlNone
   With xlWB.Application.selection.Borders(xlEdgeLeft)
       .LineStyle = xlContinuous
       .ColorIndex = 0
       .TintAndShade = 0
       .Weight = xlThin
   End With
   With xlWB.Application.selection.Borders(xlEdgeTop)
       .LineStyle = xlContinuous
       .ColorIndex = 0
       .TintAndShade = 0
       .Weight = xlThin
   End With
   With xlWB.Application.selection.Borders(xlEdgeBottom)
       .LineStyle = xlContinuous
       .ColorIndex = 0
       .TintAndShade = 0
       .Weight = xlThin
   End With
   With xlWB.Application.selection.Borders(xlEdgeRight)
       .LineStyle = xlContinuous
       .ColorIndex = 0
       .TintAndShade = 0
       .Weight = xlThin
   End With


   xlWB.Application.selection.Borders(xlInsideVertical).LineStyle = xlNone
   xlWB.Application.selection.Borders(xlInsideHorizontal).LineStyle = xlNone

   xlWB.Application.selection.Copy
   xlWB.Application.CutCopyMode = False
   xlWB.Application.selection.Copy
   xlWB.Sheets(wsName).Range("V2:CM3").select
   xlWB.Application.selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
   '' activeSheet.Paste

   xlWB.Sheets(wsName).Rows("4:14").select
   xlWB.Application.selection.EntireRow.Hidden = True
   
   
   xlWB.Sheets(wsName).Rows("15:15").select 'center machine codes
   With xlWB.Application.selection
      .HorizontalAlignment = xlCenter
      .VerticalAlignment = xlBottom
      .WrapText = False
      .Orientation = 0
      .AddIndent = False
      .IndentLevel = 0
      .ShrinkToFit = False
      .ReadingOrder = xlContext
      .MergeCells = False
    End With

   
   xlWB.Sheets(wsName).Range("H16").select
   ''ActiveWindow.FreezePanes = True
   xlWB.Windows(1).FreezePanes = True
    
    
   xlWB.Sheets(wsName).Range("U15").select
   xlWB.Sheets(wsName).Range("U15").FormulaR1C1 = "TOTAL"
   xlWB.Sheets(wsName).Range("AI15").select
   xlWB.Sheets(wsName).Range("AI15").FormulaR1C1 = "TOTAL"
   xlWB.Sheets(wsName).Range("AW15").select
   xlWB.Sheets(wsName).Range("AW15").FormulaR1C1 = "TOTAL"
   xlWB.Sheets(wsName).Range("BK15").select
   xlWB.Sheets(wsName).Range("BK15").FormulaR1C1 = "TOTAL"
   xlWB.Sheets(wsName).Range("BY15").select
   xlWB.Sheets(wsName).Range("BY15").FormulaR1C1 = "TOTAL"
   xlWB.Sheets(wsName).Range("CM15").select
   xlWB.Sheets(wsName).Range("CM15").FormulaR1C1 = "TOTAL"
   xlWB.Sheets(wsName).Range("CN15").select
   xlWB.Sheets(wsName).Range("CN15").FormulaR1C1 = "GRAND TOTAL"


   'copy the formulas necessary into the month totals in the above columns
   xlWB.Sheets(wsName).Range("U16").FormulaR1C1 = "=SUM(RC[-13]:RC[-1])"
   ''xlWB.Sheets(wsName).Range("U16").FormulaR1C1 = "=SUBTOTAL(9,RC[-13]:RC[-1])"

   
   
   xlWB.Sheets(wsName).Range("U16").select
   xlWB.Application.selection.Copy
   xlWB.Sheets(wsName).Range("U16:U526").select
   xlWB.Application.selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

   'PASTE ALL THE REST OF THE COLUMNS
   xlWB.Sheets(wsName).Range("U16:U526").select
   xlWB.Application.selection.Copy
   xlWB.Sheets(wsName).Range("AI16, AW16, BK16, BY16, CM16").select
   xlWB.Application.selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
   
   'HIDE GRAND GRAND TOTAL
   xlWB.Sheets(wsName).Columns("CN:CN").select
   xlWB.Application.selection.EntireColumn.Hidden = True

    
  
   'color only blank cells with light gray
   xlWB.Sheets(wsName).Range("H16:CN" & lastRowEdit).SpecialCells(xlCellTypeBlanks).Interior.Color = CLRGray1Light
   'and lock/protect them
   xlWB.Sheets(wsName).Range("H16:CN" & lastRowEdit).Locked = True
   ''  doesn't work or support property    xlWB.Sheets(wsName).Range("H16:CN" & lastRowEdit).Protect Contents:=True
   
   'color formula cells with darkened gray
   xlWB.Sheets(wsName).Range("H16:CN" & lastRowEdit).SpecialCells(xlCellTypeFormulas).Interior.Color = CLRGray4MedDark  'was CLRGray5Dark
   'and lock/protect them
   xlWB.Sheets(wsName).Range("H16:CN" & lastRowEdit).Locked = True
   ''  doesn't work or support property    xlWB.Sheets(wsName).Range("H16:CN" & lastRowEdit).Protect Contents:=True
    
   'color other cells with numbers or zeros some light Values blue or something  (using xlCellTypeConstants)
   xlWB.Sheets(wsName).Range("H16:CN" & lastRowEdit).SpecialCells(xlCellTypeConstants).Interior.Color = CLRcellEditBlue  'was CLRvaluesBlue
    
   'whiteout the text in row 15
   xlWB.Sheets(wsName).Rows("15:15").select
   With xlWB.Application.selection.Font
      .ThemeColor = xlThemeColorDark1
      .TintAndShade = 0
   End With


   ReDim MyArray(6 To 90) As Integer   ''for all the fields in range that want sum subtotals
   For i = 6 To 90
      MyArray(i) = i              ''for all the fields in range that want sum subtotals
   Next
   
   'subtotal level 1
   xlWB.Sheets(wsName).Range("C15:CN" & lastRowEdit).select
   xlWB.Application.selection.Subtotal GroupBy:=1, Function:=xlSum, TotalList:=MyArray, Replace:=True, PageBreaks:=False, SummaryBelowData:=True
      
   
   'subtotal level 2
   xlWB.Sheets(wsName).Range("C15:CN" & lastRowEdit).select
   xlWB.Application.selection.Subtotal GroupBy:=2, Function:=xlSum, TotalList:=MyArray, Replace:=False, PageBreaks:=False, SummaryBelowData:=False


   'COLOR EDIT TAB BLUE AND COPY TO RSMORIG FOR COMPARISON AND TRACKING CHANGES AND COLOR THAT ONE RED
   xlWB.Sheets(wsName).select
   xlWB.Application.ActiveWorkbook.Sheets("RSMEdit").Tab.Color = CLRmediumBlue
   xlWB.Worksheets(wsName).cells.select
   xlWB.Application.selection.Copy
   wsName = "RSMOrig"
   Worksheet_Action xlWB, actionCreate, wsName
   xlWB.Worksheets(wsName).select
   xlWB.Application.selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
   xlWB.Worksheets(wsName).Range("A1").select
   Set ws = xlWB.Worksheets(wsName)
   ws.cells.select
   xlWB.Application.selection.Style = "Comma"
   xlWB.Application.selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
   
   'COLOR NEW SHEET RED
   xlWB.Application.ActiveWorkbook.Sheets(wsName).Tab.Color = CLRredAsRed
   ws.Range("C1:E15").select
   xlWB.Application.selection.Interior.Color = CLRredAsRed
   
   xlWB.Sheets("RSMOrig").Range("H16").select
   xlWB.Sheets("RSMOrig").Range("A1").select
   
   wsName = "RSMEdit"
   xlWB.Sheets(wsName).select
   
   
   'write some color samples into the settings tab
   wsName = "RSMSettings"
   xlWB.Sheets(wsName).select
   'ORANGE
   xlWB.Sheets(wsName).Range("S2").FormulaR1C1 = "CLRchangedOrange  -  49918"
   xlWB.Sheets(wsName).Range("S2").select
   xlWB.Application.selection.Interior.Color = CLRchangedOrange
   'HDR BROWN
   xlWB.Sheets(wsName).Range("S3").FormulaR1C1 = "CLRheaderBrown  -  13163740"
   xlWB.Sheets(wsName).Range("S3").select
   xlWB.Application.selection.Interior.Color = CLRheaderBrown
   'VALUES GREEN
   xlWB.Sheets(wsName).Range("S4").FormulaR1C1 = "CLRvaluesGreen  -  8243375"
   xlWB.Sheets(wsName).Range("S4").select
   xlWB.Application.selection.Interior.Color = CLRvaluesGreen
   
   
   'SUB TOTAL GRAY
   xlWB.Sheets(wsName).Range("S5").FormulaR1C1 = "CLRsubTotalGray  -  12500670"
   xlWB.Sheets(wsName).Range("S5").select
   xlWB.Application.selection.Interior.Color = CLRsubTotalGray
   'GRAY1 LIGHT
   xlWB.Sheets(wsName).Range("S6").FormulaR1C1 = "CLRGray1Light  -  14606046"
   xlWB.Sheets(wsName).Range("S6").select
   xlWB.Application.selection.Interior.Color = CLRGray1Light
   'GRAY2 LIGHT-MED
   xlWB.Sheets(wsName).Range("S7").FormulaR1C1 = "CLRGray2LightMed  -  12632256"
   xlWB.Sheets(wsName).Range("S7").select
   xlWB.Application.selection.Interior.Color = CLRGray2LightMed
   'GRAY3 MEDIUM
   xlWB.Sheets(wsName).Range("S8").FormulaR1C1 = "CLRGray3Medium  -  10987431"
   xlWB.Sheets(wsName).Range("S8").select
   xlWB.Application.selection.Interior.Color = CLRGray3Medium
   'GRAY4 MED DARK
   xlWB.Sheets(wsName).Range("S9").FormulaR1C1 = "CLRGray4MedDark  -  9145227"
   xlWB.Sheets(wsName).Range("S9").select
   xlWB.Application.selection.Interior.Color = CLRGray4MedDark
   'GRAY5 DARK
   xlWB.Sheets(wsName).Range("S10").FormulaR1C1 = "CLRGray5Dark  -  7303023"
   xlWB.Sheets(wsName).Range("S10").select
   xlWB.Application.selection.Interior.Color = CLRGray5Dark
   'VALUE BLUE
   xlWB.Sheets(wsName).Range("S11").FormulaR1C1 = "CLRvaluesBlue  -  16092160"
   xlWB.Sheets(wsName).Range("S11").select
   xlWB.Application.selection.Interior.Color = CLRvaluesBlue
   'MEDIUM BLUE
   xlWB.Sheets(wsName).Range("S12").FormulaR1C1 = "CLRmediumBlue  -  16738560"
   xlWB.Sheets(wsName).Range("S12").select
   xlWB.Application.selection.Interior.Color = CLRmediumBlue
   'VALUE PALE BLUE
   xlWB.Sheets(wsName).Range("S13").FormulaR1C1 = "CLRvaluesPaleBlue  -  16441800"
   xlWB.Sheets(wsName).Range("S13").select
   xlWB.Application.selection.Interior.Color = CLRvaluesPaleBlue
   'MILD BLUE
   xlWB.Sheets(wsName).Range("S14").FormulaR1C1 = "CLRmildBlue  -  16435300"
   xlWB.Sheets(wsName).Range("S14").select
   xlWB.Application.selection.Interior.Color = CLRmildBlue
   'CELL EDIT BLUE
   xlWB.Sheets(wsName).Range("S15").FormulaR1C1 = "CLRcellEditBlue  -  16110260"
   xlWB.Sheets(wsName).Range("S15").select
   xlWB.Application.selection.Interior.Color = CLRcellEditBlue
   'RICH BLUE
   xlWB.Sheets(wsName).Range("S16").FormulaR1C1 = "CLRrichBlue  -  13134085"
   xlWB.Sheets(wsName).Range("S16").select
   xlWB.Application.selection.Interior.Color = CLRrichBlue
   'MIDNIGHT BLUE
   xlWB.Sheets(wsName).Range("S17").FormulaR1C1 = "CLRmidnightBlue  -  6299648"
   xlWB.Sheets(wsName).Range("S17").select
   xlWB.Application.selection.Interior.Color = CLRmidnightBlue
   'MACHINE HEADER BLUE
   xlWB.Sheets(wsName).Range("S18").FormulaR1C1 = "CLRmachineHdrBlue  -  15451035"
   xlWB.Sheets(wsName).Range("S18").select
   xlWB.Application.selection.Interior.Color = CLRmachineHdrBlue
   'BLACKOUT DARK
   xlWB.Sheets(wsName).Range("S19").FormulaR1C1 = "CLRblackoutDark  -  8224125"
   xlWB.Sheets(wsName).Range("S19").select
   xlWB.Application.selection.Interior.Color = CLRblackoutDark
   'LIGHT CREAM
   xlWB.Sheets(wsName).Range("S20").FormulaR1C1 = "CLRlightCream  -  13497855"
   xlWB.Sheets(wsName).Range("S20").select
   xlWB.Application.selection.Interior.Color = CLRlightCream
   'RED RED RED
   xlWB.Sheets(wsName).Range("S21").FormulaR1C1 = "CLRredAsRed  -  255"
   xlWB.Sheets(wsName).Range("S21").select
   xlWB.Application.selection.Interior.Color = CLRredAsRed
   'STARK WHITE
   xlWB.Sheets(wsName).Range("S22").FormulaR1C1 = "CLRstarkWhite  -  -1"
   xlWB.Sheets(wsName).Range("S22").select
   xlWB.Application.selection.Interior.Color = CLRstarkWhite

   
   'ADD THE FORMULA TO A COLUMN OUT THERE TO THE RIGHT IN RSM DATA.
   'THIS HAS THE NUMBER OF THE COLUMN IN RSMEDIT THAT MATCHES THE MACHINE AND MONTH FOR THIS ROW IN RSMDATA
   wsName = "RSMData"
   xlWB.Sheets(wsName).select
   'COLOR RSMData SHEET GREEN
   xlWB.Application.ActiveWorkbook.Sheets(wsName).Tab.Color = CLRvaluesGreen
   lastRowData = 7500
   xlWB.Sheets(wsName).Range("BC1").select
   xlWB.Sheets(wsName).Range("BC1").FormulaR1C1 = "COLINRSMEDIT"
   xlWB.Sheets(wsName).Range("BC2").select
   xlWB.Sheets(wsName).Range("BC2").FormulaR1C1 = "=1+(14*(-1+VLOOKUP(RC[-49],RSMSettings!R2C26:R30C27,2,FALSE)))+ INT(VLOOKUP(RC[-14],RSMSettings!R2C33:R29C34,2,FALSE))"
   xlWB.Sheets(wsName).Range("BC2").select
   xlWB.Application.selection.Copy
   xlWB.Sheets(wsName).Range("BC2:BC" & lastRowData).select
   xlWB.Application.selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=False, Transpose:=False


   xlWB.Sheets(wsName).Range("AF2").select
   xlWB.Sheets(wsName).Range("AF2").FormulaR1C1 = "=ROUND(VLOOKUP(RC[-30],RSMEdit!C7:C92,RC[23],FALSE)/IF(VLOOKUP(RC[-30],RSMOrig!C7:C92,RC[23],FALSE)<0.0001,RC[22],VLOOKUP(RC[-30],RSMOrig!C7:C92,RC[23],FALSE))*RC[-8],2)"
                                                '"=ROUND(VLOOKUP(B2,RSMEdit!$G:$CN,BC2,FALSE)/IF(VLOOKUP(B2,RSMOrig!$G:$CN,BC2,FALSE)<0.0001,BB2,VLOOKUP(B2,RSMOrig!$G:$CN,BC2,FALSE))*X2,2)"
                              '"=ROUND(VLOOKUP(RC[-30],RSMEdit!C7:C92,RC[23],FALSE)/IF(VLOOKUP(RC[-30],RSMOrig!C7:C92,RC[23],FALSE)<0.0001,RC[22],VLOOKUP(RC[-30],RSMOrig!C7:C92,RC[23],FALSE))*RC[-8],2)"
   xlWB.Sheets(wsName).Range("AF2").select
   xlWB.Application.selection.Copy
   xlWB.Sheets(wsName).Range("AF2:AF" & lastRowData).select
   xlWB.Application.selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
   xlWB.Sheets(wsName).Range("AF2:AF" & lastRowData).select
   xlWB.Application.selection.FormatConditions.Add Type:=xlExpression, Formula1:="=AF2<>X2"
   'xlWB.Application.selection.FormatConditions(selection.FormatConditions.count).SetFirstPriority
   xlWB.Application.selection.FormatConditions(1).Interior.Color = CLRchangedOrange
   xlWB.Application.selection.FormatConditions(1).StopIfTrue = False
   xlWB.Sheets(wsName).Range("A1").select


   'GO BACK TO FORMATTING RSMEDIT
   wsName = "RSMEdit"
   xlWB.Sheets(wsName).select

   'FILTER
   xlWB.Sheets(wsName).Range("C15:F" & lastRowEdit).select
   xlWB.Application.selection.AutoFilter
   xlWB.Sheets(wsName).Range("C15:F15").select
   xlWB.Application.selection.Font.Color = -9884416
   xlWB.Application.selection.Font.Name = "Calibri"
   xlWB.Application.selection.Font.Size = 8



   'load conditional formatting in reverse order

   'MAKE ALL TOTAL ROWS DARKENED GRAY
   xlWB.Sheets(wsName).Range("$D$16:$CM$559").select
   xlWB.Application.selection.FormatConditions.Add Type:=xlExpression, Formula1:="=RIGHT($D16,6)="" Total"""
   xlWB.Application.selection.FormatConditions(xlWB.Application.selection.FormatConditions.count).SetFirstPriority
   xlWB.Application.selection.FormatConditions(1).Interior.Color = CLRGray4MedDark
   xlWB.Application.selection.FormatConditions(1).StopIfTrue = False

   'MAKE balance grade TOTAL ROWS a different DARKENED GRAY
   xlWB.Sheets(wsName).Range("$C$16:$CM$559").select
   xlWB.Application.selection.FormatConditions.Add Type:=xlExpression, Formula1:="=RIGHT($C16,6)="" Total"""
   xlWB.Application.selection.FormatConditions(xlWB.Application.selection.FormatConditions.count).SetFirstPriority
   xlWB.Application.selection.FormatConditions(1).Interior.Color = CLRGray4MedDark
   xlWB.Application.selection.FormatConditions(1).StopIfTrue = False

   'MAKE CHANGED NUMBER CELLS YELLOW-ORANGE
   xlWB.Sheets(wsName).Range("$H$16:$CM$559").select
   xlWB.Application.selection.FormatConditions.Add Type:=xlExpression, Formula1:="=H16<>RSMOrig!H16"
   xlWB.Application.selection.FormatConditions(xlWB.Application.selection.FormatConditions.count).SetFirstPriority
   xlWB.Application.selection.FormatConditions(1).Interior.Color = CLRchangedOrange
   xlWB.Application.selection.FormatConditions(1).StopIfTrue = False

   'WHITE OUT THE FIRST ZED STUFF - WHITE FONT FOR ALL ZEDS
   xlWB.Sheets(wsName).Columns("C:E").select
   xlWB.Application.selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""ZED"""
   xlWB.Application.selection.FormatConditions(xlWB.Application.selection.FormatConditions.count).SetFirstPriority
   xlWB.Application.selection.FormatConditions(1).Font.Color = CLRstarkWhite
   xlWB.Application.selection.FormatConditions(1).StopIfTrue = False
    
   'WHITE OUT THE SECOND ZED STUFF - WHITE BACKGROUND OVERRIDES FOR ZED TOTALS
   xlWB.Application.selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""ZED Total"""
   xlWB.Application.selection.FormatConditions(xlWB.Application.selection.FormatConditions.count).SetFirstPriority
   xlWB.Application.selection.FormatConditions(1).Interior.Color = CLRstarkWhite
   xlWB.Application.selection.FormatConditions(1).StopIfTrue = False
   
   'TOTAL COLUMNS GETTING OWN GRAYS, PUT HERE??




   xlWB.Sheets(wsName).Range("H16").select
   xlWB.Sheets(wsName).Range("A1").select

    
   'hide your sheet if you don't want people to see it
   For Each ws In xlWB.Worksheets
      If ws.Name = "RSMEdit" Or ws.Name = "RSMData" Or ws.Name = "RSMOrig" Then   'ok, leave your sheet out in the open
      Else
         ws.Visible = xlSheetHidden    'hide that sheet
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
