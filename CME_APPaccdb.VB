===============
Form_fUpdateTheModel
===============
Option Compare Database


'Append Kiwi history for latest month

Private Sub btnAppendKiwiData_Click()
      
 On Error GoTo Err_Handler
 
    DoCmd.SetWarnings False
    DoCmd.Hourglass True
    
    Dim strFileName As String
    
    'Clear out tblKiwiHistRaw (keep data in raw form for now - comment out)
    'DoCmd.OpenQuery "qKiwiHistRawClr"

    'Select the file to import and use Link Specification to append to tblZ_CTA_BMOS_EXTRACT
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Select the CSV file to import"
        .AllowMultiSelect = False
        .Filters.Clear
        .InitialFileName = "\\S02akiwikrs04\Out"
        .Filters.Add "CSV Files", "*.csv", 1
        'Import the .csv into the Access tblKiwiHistRaw
        If .Show = -1 Then
            strFileName = .SelectedItems(1)
            DoCmd.TransferText acImportDelim, "KPIPMPR_YYYYMM Link Specification", "tblKiwiHistRaw", strFileName, False
            Else
            Exit Sub
        End If
    End With
    
    'Update SoldToCommon and ShipToCommon fields on tblCTA_BMOS_LocationXref
    DoCmd.OpenQuery "qKiwiHistClr"
    DoCmd.OpenQuery "qKiwiHistApd"
  
  
  DoCmd.SetWarnings True
  DoCmd.Hourglass False
  Me.Requery
  Me.Refresh
  MsgBox "Kiwi data appended for latest month: " + strFileName
    
Err_Handler:
 If Err Then
      DoCmd.SetWarnings True
      DoCmd.Hourglass False
      MsgBox "Problem with appending Kiwi data " & vbNewLine & Err.Number & " - " & Err.Description
 End If
  
End Sub




Private Sub btnApdShipmentHistoryRaw_Click()
    If MsgBox("Do you want to refresh shipment history from SAP_BW_ShipmentHistoryDataExtract.xls?", vbYesNo + vbExclamation) = vbYes Then
        
    DoCmd.SetWarnings False
    DoCmd.Hourglass True
        
    DoCmd.OpenQuery "qShipmentHistoryRawDataClr"
    DoCmd.OpenQuery "qShipmentHistoryRawDataApd"
               
    DoCmd.SetWarnings True
    DoCmd.Hourglass False
    Me.Requery
    Me.Refresh
    MsgBox "New shipment history imported from spreadsheet"
  End If
End Sub

Private Sub btnAppendNewProducts_Click()
    
    'On Error GoTo Err_Handler
    
    DoCmd.SetWarnings False
    DoCmd.Hourglass True
        
    'Refresh tblZ_CTA_BMOS_Product
    DoCmd.OpenQuery "qZ_CTA_BMOS_EXTRACT_ProductClr"
    'DoCmd.OpenQuery "qZ_CTA_BMOS_EXTRACT_ProductApd"
    DoCmd.OpenQuery "qZ_CTA_BMOS_EXTRACT_ProductApdFromAnalysis" 'This is from Analysis for Office worksheet
    
    
    'TEMPORARY ***!!! remove product records for plants not in scope
    DoCmd.OpenQuery "qZ_CTA_BMOS_EXTRACT_ProductClrNotInScope"
    
   'Clear and append the Witron extract (manual WeySQL step has to be completed before this query runs)
    DoCmd.OpenQuery "qZ_CTA_BMOS_ProductInfoFromWitronClr"
    DoCmd.OpenQuery "qZ_CTA_BMOS_ProductInfoFromWitronApd"
            
    'Add standarized ProdCode, SoldTo, ShipTo
    DoCmd.OpenQuery "qZ_CTA_BMOS_EXTRACT_ProductUpd"
    
    'Clear and append the Witron extract
    DoCmd.OpenQuery "qZ_CTA_BMOS_ProductInfoFromWitronClr"
    DoCmd.OpenQuery "qZ_CTA_BMOS_ProductInfoFromWitronApd"
    
    'Update corrugator and converter wax info
    DoCmd.OpenQuery "qZ_CTA_BMOS_ProductInfoFromWitronCorrUpd"
    DoCmd.OpenQuery "qZ_CTA_BMOS_ProductInfoFromWitronConvUpd"

            
    'Update the the product table using the Witron download cross reference (tblZ_CTA_BMOS_ProductInfoFromWitron)
    DoCmd.OpenQuery "qZ_CTA_BMOS_ProductInfoWitronUpd"
    
    'Update the machine types for each routing stage on tblZ_CTA_BMOS_Product
    DoCmd.OpenQuery "qZ_CTA_BMOS_EXTRACT_MachineType1Upd"
    DoCmd.OpenQuery "qZ_CTA_BMOS_EXTRACT_MachineType2Upd"
    DoCmd.OpenQuery "qZ_CTA_BMOS_EXTRACT_MachineType3Upd"
    DoCmd.OpenQuery "qZ_CTA_BMOS_EXTRACT_MachineType4Upd"
    DoCmd.OpenQuery "qZ_CTA_BMOS_EXTRACT_MachineType5Upd"
    DoCmd.OpenQuery "qZ_CTA_BMOS_EXTRACT_MachineType6JoinUpd"
    
    'If there is a laminator, change the machine family to laminator (on MachineType5)
    DoCmd.OpenQuery "qZ_CTA_BMOS_EXTRACT_MachineFamilyLaminatorUpd"
    
    'If there is a paster (flute=Z)change the machine family to Paster
    DoCmd.OpenQuery "qZ_CTA_BMOS_EXTRACT_MachineFamilyPasterUpd"
    
    'Identify outsourced items as machine "0106" (outsourced corrugator) and don't try to include downstream machines
    DoCmd.OpenQuery "qZ_CTA_BMOS_EXTRACT_MachineTypeOutsourceUpd"
        
    'Convert stock sheets (Machine 0105) to regular sheet machine
    DoCmd.OpenQuery "qZ_CTA_BMOS_EXTRACT_MachineTypeStockSheetsUpd"
    
    'Fix any remaining CSR derived routings (they tend to have anomolies so use default machine type routing by product type)
    DoCmd.OpenQuery "qZ_CTA_BMOS_EXTRACT_MachineTypeDefaultByProdTypeUpdCSR"
        
    'Remove any record where the width is 0 (usually indicates pallets and other non-product related transactions)
    DoCmd.OpenQuery "qZ_CTA_BMOS_ProductRemoveNonProductItems"
    
    'Unless the machine type is an FFG, don't track if it has a cutting tool
    DoCmd.OpenQuery "qZ_CTA_BMOS_EXTRACT_TrackCuttingToolReqOnlyForFFG"
    
    'Unless the machine type is a RDC, don't allow it to require a bundle breaker
    DoCmd.OpenQuery "qZ_CTA_BMOS_EXTRACT_TrackCuttingToolReqOnlyForFFG"
    
    'If machine family empty ("TBD") then assign a default based on the product type and give default routing
    DoCmd.OpenQuery "qZ_CTA_BMOS_EXTRACT_MachineTypeDefaultByProdTypeUpd"
    
    'If Gross to Net Ratio is 0, make it 1
    DoCmd.OpenQuery "qZ_CTA_BMOS_EXTRACT_GrossToNetRatioFixUpd"
    
    
    '----Product fixing queries not currently used-------------------
    'If the MachineType1 is missing (e.g., 0106 Farmout), then give it the default routing by machine family
    'DoCmd.OpenQuery "qZ_CTA_BMOS_EXTRACT_MachineRoutingByMachineTypeDefaultUpd"
     
    'Use default routing based by machine family if XXXXXX (see tblZ_CTA_BMOS_EXTRACT_DefaultRoutingByMachineType)
    'DoCmd.OpenQuery "qZ_CTA_BMOS_EXTRACT_DefaultRoutingUpd"
        
    'If the second machine is empty assign a default machine type by machine family
    ' DoCmd.OpenQuery "qZ_CTA_BMOS_EXTRACT_MachineType2FixMissingUpd"
         
    'If the MachineType2 is missing then give it the default MachineType2 for the machine family
    'DoCmd.OpenQuery "qZ_CTA_BMOS_EXTRACT_MachineRoutingByMachineTypeDefault2Upd"
    
    'Refresh tblProduct which is aggregated version of tblZ_CTA_BMOS_Product(Local)
    DoCmd.OpenQuery "qProductClr"
    DoCmd.OpenQuery "qProductApd"
    
    'Fix the issue of gross width greater than gross length (when reversed - often on RDC products)
    'Go through each record of tblProduct and swap initial width & length when width>length for RDCs (may want to limit with query)
    Dim db As DAO.Database
    Dim rstSpec As DAO.Recordset
    Dim OrigIniWidth As Double
    Dim OrigIniLength As Double
    Dim OrigFinalWidth As Double
    Dim OrigFinalLength As Double
        
    Set db = CurrentDb
    Set rstSpec = db.OpenRecordset("tblProduct")
    
    rstSpec.MoveFirst
    Do Until rstSpec.EOF
      'If rstSpec!Machine_Group = "5" And rstSpec!Ini_Width > rstSpec!Ini_Length Then
       If rstSpec!GrossSheetWidth > rstSpec!GrossSheetLength Then
         OrigIniWidth = rstSpec!GrossSheetWidth
         OrigIniLength = rstSpec!GrossSheetLength
         OrigFinalWidth = rstSpec!NetSheetWidth
         OrigFinalLength = rstSpec!NetSheetLength
         
         rstSpec.Edit
         rstSpec!GrossSheetWidth = OrigIniLength
         rstSpec!GrossSheetLength = OrigIniWidth
         rstSpec!NetSheetWidth = OrigFinalLength
         rstSpec!NetSheetLength = OrigFinalWidth
         'rstSpec!WidthLengthSwap = "Y"
         rstSpec.Update
       End If
    
       rstSpec.MoveNext
    Loop
    db.Close
    Set db = Nothing
            
            
    'Refresh dbo_tblCTA_CSR_Product (SQL Server)
    DoCmd.OpenQuery "qZ_CTA_BMOS_EXTRACT_ProductClrSQLServer"
    DoCmd.OpenQuery "qZ_CTA_BMOS_EXTRACT_ProductApdSQLServer"
    
    'Refresh dbo_tblProduct (SQL Server Express)
    DoCmd.OpenQuery "qdboProductClr"
    DoCmd.OpenQuery "qdboProductApd"
           
    Me.Requery
    Me.Refresh
    
    DoCmd.SetWarnings True
    DoCmd.Hourglass False
      
   MsgBox "Refreshed product information"
   
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
    
    
    
    
End Sub



Private Sub btnClearShipmentHistoryRawData_Click()
    If MsgBox("Are you sure you want to delete records from the shipment history below?", vbYesNo + vbExclamation) = vbYes Then
        
    DoCmd.SetWarnings False
    DoCmd.Hourglass True
        
    DoCmd.RunSQL "Delete * From tblShipmentHistoryRawData"
               
    DoCmd.SetWarnings True
    DoCmd.Hourglass False
    Me.Requery
    Me.Refresh
    MsgBox "Shipment history cleared"
  End If
End Sub




Private Sub btnCheckForNewOrigDestPostalCodeLanes_Click()
    'Appends any new lanes (keys in tblLanesOrigDestPostalCodes keep out dup's)
    
    DoCmd.SetWarnings False
    DoCmd.Hourglass True
    
    'Append any new lanes to tblLanesOrigDestPostalCodes
    'DoCmd.OpenQuery "qdboLaneOrigDestPostalCodesApd"   'This was to SQL Server
    DoCmd.OpenQuery "qLaneOrigDestPostalCodesApd"       'This is to .accdb
    
    MsgBox "New Postal Code Origin / Destination Pairs Identified"
    
    DoCmd.SetWarnings True
    DoCmd.Hourglass False
    Me.Requery
    Me.Refresh
     
End Sub

Private Sub btnConvertingWax_Click()

  DoCmd.OpenQuery "qZ_CTA_BMOS_ProductInfoConvWaxXrefDisplay"

End Sub

Private Sub btnCorrugatingWax_Click()
    
  DoCmd.OpenQuery "qZ_CTA_BMOS_ProductInfoCorrWaxXrefDisplay"
    
End Sub



Private Sub btnCreateLaneSpreadsheet_Click()
   
   On Error GoTo error_handler
   
   Dim sMileageTool As String
   Dim xlApp As Object
   Dim wb As Workbook
   Dim ws As Worksheet
   Dim sFileName As String
   Dim sTableName As String
   Dim lLastRow As Long
   
   DoCmd.SetWarnings False
   DoCmd.Hourglass True
   
   'Delete spreadsheet if it exists
    sFileName = "C:\OPTMODELS\CV1\MileagesExternalLanesSQLServer.xlsx"
    If Len(Dir$(sFileName)) > 0 Then
      Kill sFileName
    End If
    
   'Determine mileage tool (either IR = IntelliRoute or PC=PC*Miler)
    sMileageTool = DLookup("MileageTool", "tblParameters")
      
   'Create Excel spreadsheet from tblFreightCost
    sTableName = "qZ_CTA_BMOS_LanesExternal"
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, sTableName, sFileName
        
   'Insert mileage formula depending on mileage tool available
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Application.ScreenUpdating = True
    xlApp.Application.Visible = True
    
    Set wb = xlApp.Workbooks.Open(sFileName)
    Set ws = wb.Worksheets(sTableName)
   
    With ws
     lLastRow = .Range("A1").SpecialCells(xlCellTypeLastCell).Row
     If sMileageTool = "IR" Then
        .Range("J2") = "=HHGNC(C2,H2)"
     Else
        .Range("J2") = "=MILES(C2,H2,""PRAC"")"
     End If
    .Range("J2").Select
    .Range("J2").Copy
    .Range("J3").Select
    .Range("J3:J" & lLastRow).Select
    .Paste
    End With
    
    xlApp.Application.ScreenUpdating = True
    xlApp.DisplayAlerts = False
    wb.Save
    wb.Close False
    xlApp.Quit
    Set xlApp = Nothing
    
    DoCmd.SetWarnings True
    DoCmd.Hourglass False
    MsgBox "MileagesExternalLanesSQLServer.xlsx created - open it as Admin and wait until all mileages are completed, then Save/Exit"
    
error_handler:
If Err Then
      
    wb.Close False
    xlApp.Quit
    Set xlApp = Nothing
      
    DoCmd.SetWarnings True
    DoCmd.Hourglass False
      
    MsgBox "Problem with Excel Mileage Process" & vbNewLine & Err.Number & " - " & Err.Description
      
 End If

End Sub


Private Sub btnCreateCustomerSpecificMSFperShipmentFactors_Click()
   If MsgBox("This will clear and recreate MSF per Shipment Factors by customer location, are you sure?", vbYesNo + vbExclamation) = vbYes Then
        
    DoCmd.SetWarnings False
    DoCmd.Hourglass True
        
    DoCmd.OpenQuery "qShipmentAnalysisRawClr"
    DoCmd.OpenQuery "qShipmentAnalysisRawApd"
    
    DoCmd.OpenQuery "qShipmentAnalysisClr"
    DoCmd.OpenQuery "qShipmentAnalysisApd"
                   
    DoCmd.SetWarnings True
    DoCmd.Hourglass False
    Me.Requery
    Me.Refresh
    MsgBox "MSF per Shipment Factors by Customer Location refreshed"
  End If


End Sub

Private Sub btnCustomerMoveDisplay_Click()

    DoCmd.OpenQuery "qSourcingMinCostSavings"

End Sub

Private Sub btnDetermineBestCustomerMove_Click()

If MsgBox("Do you want to determine best customer moves?", vbYesNo + vbQuestion) = vbYes Then
    
    DoCmd.SetWarnings False
    DoCmd.Hourglass True
    
    'Refresh dbo_tblPlant with plants in scope
    DoCmd.OpenQuery "qdboPlantClr"
    DoCmd.OpenQuery "qdboPlantApd"
    
    'Refresh tblCTA_CSR_Plant with plants in scope
    DoCmd.OpenQuery "qdbo_tblCTA_CSR_PlantClr"
    DoCmd.OpenQuery "qdbo_tblCTA_CSR_PlantApd"
    
'GoTo SkipSourcingHistRefresh
 
                
    'Clear and append dbo_tblSourcingHist (switch from SQL Server table based query to SQL Server view to speed up?)
    DoCmd.OpenQuery "qdboSourcingHistClr"
    DoCmd.OpenQuery "qdboSourcingHistApd"
    
SkipSourcingHistRefresh:

        
    'Update the freight cost and lowest cost (These were replaced by SQL Server Express Stored Procedures)
    'DoCmd.OpenQuery "qdboSourcingHistCostUpdFrtOnly"
    'DoCmd.OpenQuery "qdboSourcingLowCostUpdFrtOnly"
        
   'Update historical freight cost in dbo_tblSourcingHistt with a SQL Server Local stored procedure
    Dim qdef1 As DAO.QueryDef
    Set qdef1 = CurrentDb.CreateQueryDef("")
    qdef1.Connect = CurrentDb.TableDefs("dbo_tblSourcingHist").Connect
    qdef1.SQL = "EXEC spSourcingHistFrtUpd"
    qdef1.ReturnsRecords = False  ''avoid 3065 error
    qdef1.Execute
    qdef1.Close
    Set qdf1 = Nothing
            
   'Update low cost freight source cost in dbo_tblSourcingHist with a SQL Server Express stored procedure
    Dim qdef2 As DAO.QueryDef
    Set qdef2 = CurrentDb.CreateQueryDef("")
    qdef2.Connect = CurrentDb.TableDefs("dbo_tblSourcingHist").Connect
    qdef2.SQL = "EXEC spSourcingHistLowCostUpd"
    qdef2.ReturnsRecords = False  ''avoid 3065 error
    qdef2.Execute
    qdef2.Close
    Set qdf2 = Nothing
    
    
    '****Activate the GoTo below to skip SharePoint refresh****
    GoTo SkipSharePointSourcingLocks
    
             
    'Bring in historical Locks from SharePoint
    DoCmd.OpenQuery "qSourcingLocksUserInputClr"
    DoCmd.OpenQuery "qSourcingLocksUserInputApd"
    
    'Refresh tblSourcingLocks (will include any new lanes)
    DoCmd.OpenQuery "qSourcingLocksClr"
    DoCmd.OpenQuery "qSourcingLocksApd"
            
    'Update tblSourcingLocks with any locks from SharePoint
    DoCmd.OpenQuery "qSourcingLocksUpd"
        
        
    'Update dbo_tblSourcingHist with current locks
    DoCmd.OpenQuery "qdboSourcingLocksUpd"
    
    'Refresh SharePoint with latest moves and retained locks
    DoCmd.OpenQuery "qSourcingLocksClrSharePoint"
    DoCmd.OpenQuery "qSourcingLocksApdSharePoint"
    
SkipSharePointSourcingLocks:
    
    
    'Refresh dbo_tblSourcingMapCombined used in Power BI
    DoCmd.OpenQuery "qdboSourcingMapCombinedClr"
    'DoCmd.OpenQuery "qdboSourcingMapCombinedApd" (too slow - replaced by the three Append queries below
    
    DoCmd.OpenQuery "qdboSourcingMapCombinedHistApd"
    DoCmd.OpenQuery "qdboSourcingMapCombinedLowApd"
    DoCmd.OpenQuery "qdboSourcingMapCombinedPlantApd"
    
    'Refresh dbo_tblSourcingMapCombinedHistOnly used in Power BI
    DoCmd.OpenQuery "qdboSourcingMapCombinedHistOnlyClr"
    DoCmd.OpenQuery "qdboSourcingMapCombinedHistOnlyApd"
    
    'Refresh dbo_tblSourcingOppHeader and dbo_tblSourcingOppDetail
    DoCmd.OpenQuery "qdboSourcingOppHeaderClr"
    DoCmd.OpenQuery "qdboSourcingOppHeaderApd"
    DoCmd.OpenQuery "qdboSourcingOppDetailClr"
    DoCmd.OpenQuery "qdboSourcingOppDetailApd"
    
        
    Me.Requery
    Me.Refresh
        
    DoCmd.SetWarnings True
    DoCmd.Hourglass False
    
    MsgBox "Best Customer Move Candidates Developed"
    
  End If
End Sub

Private Sub btnFreightCostClear_Click()

'Check to make sure the user wants to regen
If MsgBox("Do you want to clear out all the freight records?", vbYesNo + vbQuestion) = vbYes Then
    
    DoCmd.SetWarnings False
    DoCmd.Hourglass True

       
    'Clear tblFreightCost
    DoCmd.OpenQuery "qFreightCostClr"
    
    
    DoCmd.SetWarnings True
    DoCmd.Hourglass False
        
    MsgBox "Records removed from tblFreightCost, Compact/Repair before proceeding"
    
End If
 
End Sub

Private Sub btnFreightCostDisplay_Click()
    DoCmd.OpenQuery "qFreightCostDisplay"
End Sub

Private Sub btnGenerateProductionAlternatives_Click()

   'On Error GoTo error_handler
    
'Check to make sure the user wants to regen
If MsgBox("Do you want to refresh all production and distribution capabilities?", vbYesNo + vbQuestion) = vbYes Then
    
    DoCmd.SetWarnings False
    DoCmd.Hourglass True

'Activate to skip long running append by plant (just for testing)
'GoTo SkipAppendByPlant

    
    'Loop through each plant and generate production options (runs qdboMakeProdApdSQLServerExpress for each in-scope plant)
    LoopThroughPlantsInScope

SkipAppendByPlant:
    
    'Update costs in dbo_tblMakeProd cost with a SQL Server Local stored procedure
    'Dim qdef As DAO.QueryDef
    'Set qdef = CurrentDb.CreateQueryDef("")
    'qdef.Connect = CurrentDb.TableDefs("dbo_tblFreightCost").Connect
    'qdef.SQL = "EXEC spMakeProdFrtUpd"
    'qdef.ReturnsRecords = False  ''avoid 3065 error
    'qdef.Execute
    
    
    'Update freight cost in dbo_tblCTA_CSR_MakeProd cost with a SQL Server stored procedure
    'Dim qdef As DAO.QueryDef
    'Set qdef = CurrentDb.CreateQueryDef("")
    'qdef.Connect = CurrentDb.TableDefs("dbo_tblCTA_CSR_MakeProd").Connect
    'qdef.SQL = "EXEC sp_tblCTA_CSR_MakeProdUpdFreight"
    'qdef.ReturnsRecords = False  ''avoid 3065 error
    'qdef.Execute
      
    'Clear and append dbo_tblMakeProdAvg
    'DoCmd.OpenQuery "qdboMakeProdAvgClr"
    'DoCmd.OpenQuery "qdboMakeProdAvgApd"
    'DoCmd.OpenQuery "qdboMakeProdAvgApdFromSQLServer"   '***This takes forever to run; look into finding an alternative (e.g., view, use "first" instead of average?)
    'DoCmd.OpenQuery "qdboMakeProdAvgApdFromSQLServerNEW"   'Uses view on SQL Server but is creating dup's -- need's work...
        
    'Clear and append dbo_tblMakeProdMin
    DoCmd.OpenQuery "qdboMakeProdMinClr"
    DoCmd.OpenQuery "qdboMakeProdMinFrtOnlyApd"
    
    'Maybe put a radio button on the form to toggle which of these two queries to run???? (ON INDEFINITE HOLD)
    'DoCmd.OpenQuery "qdboMakeProdMinMfgAndFrtApd"
    
    'Clear and append dbo_tblMakeProdLow
    DoCmd.OpenQuery "qdboMakeProdLowClr"
    DoCmd.OpenQuery "qdboMakeProdLowApd"
        
    Me.Requery
    Me.Refresh
        
    DoCmd.SetWarnings True
    DoCmd.Hourglass False
    
    MsgBox "Refresh of Production Alternatives Complete"
    
  End If
  
error_handler:
If Err Then
    
    DoCmd.SetWarnings True
    DoCmd.Hourglass False
      
    MsgBox "Problem creating MakeProd options" & vbNewLine & Err.Number & " - " & Err.Description
          
End If
  
End Sub

Private Sub btnRefreshHistAdj_Click()
  'Check to make sure user really wants to do this
  If MsgBox("Are you sure you want to refresh -- it will remove current plant selections?", vbYesNo + vbExclamation) = vbYes Then
  
    DoCmd.SetWarnings False
    DoCmd.Hourglass True
  
    'Clear and append to tblShipmentHistoryAdj
    DoCmd.OpenQuery "qShipmentHistoryAdjClr"
    DoCmd.OpenQuery "qShipmentHistoryAdjApd"
    
    Me.Requery
    Me.Refresh
        
    DoCmd.SetWarnings True
    DoCmd.Hourglass False
    
  End If
End Sub

Private Sub btnMakeProdDisplay_Click()

    DoCmd.OpenQuery "qdboMakeProdDisplaySQLServer"

End Sub


Private Sub btnRefreshFreightCosts_Click()
    DoCmd.SetWarnings False
    DoCmd.Hourglass True
         
    'Refresh Freight Costs
    'DoCmd.OpenQuery "qdboFreightCostClr"
    'DoCmd.OpenQuery "qdboFreightCostApd"
    DoCmd.OpenQuery "qFreightCostClr"
    DoCmd.OpenQuery "qFreightCostApd"
         
    'Update tblFreighCost mileages and mileage bands from tblLaneOrigDestPostalCodes
    'DoCmd.OpenQuery "qdboFreightCostMileageUpd"
    DoCmd.OpenQuery "qFreightCostMileageUpd"

    
    'Update dbo.tblFreighCost where mileages are zero in the US and set to one mile
    'DoCmd.OpenQuery "qdboFreightCostSetZeroMileageToOneBandUpd"
    DoCmd.OpenQuery "qFreightCostSetZeroMileageToOneBandUpd"
    
    'Update the tblFreightCost Volumes Per Shipment, etc.
    DoCmd.OpenQuery "qShipmentAnalysisShipToUpd"
    DoCmd.OpenQuery "qFreightCostMSFperShipmentUpd"
    
    'Update tblFreightCost using generic mileage bands then TBC rates, going from general to specific (not necessarily decreasing cost)
    DoCmd.OpenQuery "qFreightRateGenericMileageBandsUpd"
    'DoCmd.OpenQuery "qdboFreightRateGenericMileageBandsUpd"
    
    'Update the MSF per Shipment Used
    DoCmd.OpenQuery "qFreightCostMSFperShipmentUsedUpd"
    
    'Apply state default cost
    DoCmd.OpenQuery "qFreightRateStateClr"
    DoCmd.OpenQuery "qFreightRateStateApd"
    DoCmd.OpenQuery "qFreightRateStateUpd"
    'DoCmd.OpenQuery "qdboFreightRateStateUpd"
    
    'Apply split state default cost
    DoCmd.OpenQuery "qFreightRateSplitStateClr"
    DoCmd.OpenQuery "qFreightRateSplitStateApd"
    DoCmd.OpenQuery "qFreightRateSplitStateUpd"
    'DoCmd.OpenQuery "qdboFreightRateSplitStateUpd"
        
    'Apply 3 digit zip rate
    DoCmd.OpenQuery "qFreightRate3DigitZipClr"
    DoCmd.OpenQuery "qFreightRate3DigitZipApd"
    DoCmd.OpenQuery "qFreightRate3DigitZipUpd"
    'DoCmd.OpenQuery "qdboFreightRate3DigitZipUpd"
       
    'Apply additional mileage band costs
    DoCmd.OpenQuery "qFreightRateMileageBandClr"
    DoCmd.OpenQuery "qFreightRateMileageBandApd"
    DoCmd.OpenQuery "qFreightCostMileageBandAdditionalUpd"
    'DoCmd.OpenQuery "qdboFreightCostMileageBandAdditionalUpd"
   
    'Apply specific city rate
    DoCmd.OpenQuery "qFreightRateCityClr"
    DoCmd.OpenQuery "qFreightRateCityApd"
    DoCmd.OpenQuery "qFreightRateCityUpd"
    'DoCmd.OpenQuery "qdboFreightRateCityUpd"
    
    'Update fuel surcharge
    DoCmd.OpenQuery "qFreightRateFuelSurchargeUpd"
    'DoCmd.OpenQuery "qdboFreightRateFuelSurchargeUpd"
    
    'Calculate the freight cost
    DoCmd.OpenQuery "qFreightRateFuelSurchargeCostUpd"
    DoCmd.OpenQuery "qFreightRateLineHaulUpd"
    'DoCmd.OpenQuery "qdboFreightRateLineHaulUpd"
    DoCmd.OpenQuery "qFreightCostTotalCostPerMSFUpd"
    'DoCmd.OpenQuery "qdboFreightCostTotalCostPerMSFUpd"
    
    'Update the historical shipments
    DoCmd.OpenQuery "qFreightCostHistShipmentsClr"
    DoCmd.OpenQuery "qFreightCostHistShipmentsApd"
    DoCmd.OpenQuery "qFreightCostHistShipmentsUpd"
    
    'Incorpoate the Chainalytics benchmark rates for each origin/destination postal code
    DoCmd.OpenQuery "qChainalyticsBenchmarkRatesClr"
    DoCmd.OpenQuery "qChainalyticsBenchmarkRatesApd"
    DoCmd.OpenQuery "qChainalyticsBenchmarkRateUpdFreightCost"
    
    'Replace generic mileage bands with Chainalytics benchmark rates
    DoCmd.OpenQuery "qChainalyticsBenchmarkRateUpdGenericMileageBandRates"
        
    'Refresh dbo_tblFreightCost on SQL Server Express
    DoCmd.OpenQuery "qdboFreightCostClr"
    DoCmd.OpenQuery "qdboFreightCostApd"
    
    'Refresh dbo_tblFreightCost on SQL Server
    'DoCmd.OpenQuery "qdboCTA_CSR_FreightCostClrSQLServer"
    'DoCmd.OpenQuery "qdboCTA_CSR_FreightCostApdSQLServer"
          
    DoCmd.SetWarnings True
    DoCmd.Hourglass False
    
    Me.Requery
    Me.Refresh
        
    MsgBox "Estimated freight costs calculated for all lanes"
    
End Sub

'Clear and refresh tblLaneOrigDestPostalCode and get mileage from the Excel Add-in

Private Sub btnRefreshLanes_Click()
  'Check to make sure user really wants to do this
  If MsgBox("Are you sure you want to calculate mileages for new Postal Code lanes? (this can take awhile...)", vbYesNo + vbExclamation) = vbYes Then
  
    On Error GoTo error_handler
  
    DoCmd.SetWarnings False
    DoCmd.Hourglass True
  

   Dim sMileageTool As String
   Dim xlApp As Object
   Dim wb As Workbook
   Dim ws As Worksheet
   Dim sFileName As String
   Dim sTableName As String
   Dim lLastRow As Long

   'Delete spreadsheet if it exists
    sFileName = "C:\OPTMODELS\CME\MileagesExternalLanes.xlsx"
    If Len(Dir$(sFileName)) > 0 Then
      Kill sFileName
    End If
    
   'Determine mileage tool (either IR = IntelliRoute or PC=PC*Miler)
    sMileageTool = DLookup("MileageTool", "tblParameters")
      
   'Create Excel spreadsheet from tblLaneOrigDestPostalCodes
    sTableName = "qLaneOrigDestPostalCodesNew"
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, sTableName, sFileName
        
   'Insert mileage formula depending on mileage tool available
    Set xlApp = CreateObject("Excel.Application")
        
    xlApp.Workbooks.Open ("C:\Program Files\Rand McNally\IntelliRoute\ApiBin\irexcel.xla")
    xlApp.Workbooks("irexcel.xla").RunAutoMacros 1
    
    Set wb = xlApp.Workbooks.Open(sFileName)
    Set ws = wb.Worksheets(sTableName)
        
    xlApp.Application.Visible = True
    xlApp.Application.ScreenUpdating = True
    xlApp.DisplayAlerts = False
    
    With ws
     lLastRow = .Range("A1").SpecialCells(xlCellTypeLastCell).Row
     If sMileageTool = "IR" Then
        .Range("C2") = "=HHGNC(A2,B2)"
     Else
        .Range("C2") = "=MILES(A2,B2,""PRAC"")"
     End If
    .Range("C2").Select
    .Range("C2").Copy
    .Range("C3").Select
    .Range("C3:C" & lLastRow).Select
    .Paste
   
    Do Until xlApp.Application.CalculationState = xlDone
        DoEvents
    Loop
    
    
    End With
      
    wb.Save
    wb.Close True
    xlApp.Quit
     
    
    Set xlApp = Nothing
    
    'Clear and the append to tblLaneOrigDestPostalCodesTmp from Excel spreadsheet with mileages for new lanes
    DoCmd.OpenQuery "qLaneOrigDestPostalCodesTmpClr"
    DoCmd.OpenQuery "qLaneOrigDestPostalCodesTmpApdFromExcel"
    
    
    'Update tblLanesOrigDestPostalCodes
    'DoCmd.OpenQuery "qdboLaneOrigDestPostalCodesUpd"
    DoCmd.OpenQuery "qLaneOrigDestPostalCodesUpd"
    
    
    DoCmd.SetWarnings True
    DoCmd.Hourglass False

    Me.Requery
    Me.Refresh
    
    MsgBox "Mileages updated for any new Postal Code lanes"

error_handler:
If Err Then
      
    'wb.Close False
    xlApp.Quit
    Set xlApp = Nothing
      
    DoCmd.SetWarnings True
    DoCmd.Hourglass False
      
    MsgBox "Problem with Excel Mileage Process" & vbNewLine & Err.Number & " - " & Err.Description

    Me.Requery
    Me.Refresh
        
    DoCmd.SetWarnings True
    DoCmd.Hourglass False
End If
End If

End Sub
 
    



Private Sub btnRefreshMachineFromOptModel_Click()

    On Error GoTo Err_Handler
    
    DoCmd.SetWarnings False
    DoCmd.Hourglass True
    
    'Check to make sure user really wants to do this
    If MsgBox("Are you sure you want to refresh machines from Opt Model?", vbYesNo + vbExclamation) = vbYes Then
    
    'Refresh the Area and RGM from shipment history tracking plant table which gets updated monthly (may want to change this to something more permanent at some point)
    DoCmd.OpenQuery "qPlantAreaAndRGMUpd"
    
    'Refresh tblMachineLocal from the NAC Opt Model machine table
    DoCmd.OpenQuery "qMachineLocalClr"
    DoCmd.OpenQuery "qMachineLocalApd"
    
    Me.Requery
    Me.Refresh
    
    DoCmd.SetWarnings True
    DoCmd.Hourglass False

    
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
   DoCmd.SetWarnings True
   DoCmd.Hourglass False

End Sub

Private Sub btnSelectCSV_Click()
    
    On Error GoTo Err_Handler
   
    'Check to make sure user really wants to do this
     If MsgBox("Are you sure you want append a new CSV? --don't forget to Compact/Repair before appending another month", vbYesNo + vbExclamation) = vbYes Then
     'If MsgBox("Are you sure you want incorporate new data from the Analysis spreadsheet?", vbYesNo + vbExclamation) = vbYes Then
    
    DoCmd.SetWarnings False
    DoCmd.Hourglass True
    
    Dim strFileName As String
       
    'Clear the tblZ_CTA_BMOS_EXTRACT table (temporary table for latest month)
    DoCmd.RunSQL "Delete * From tblZ_CTA_BMOS_EXTRACT"
       
'GoTo SkipFileSelect
    
    'Select the file to import and use Link Specification to append to tblZ_CTA_BMOS_EXTRACT
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Select the CSV file to import"
        .AllowMultiSelect = False
        .Filters.Clear
        .InitialFileName = "\\s02aess05\Production\Sales\Transfer\BMOS"
        .Filters.Add "CSV Files", "*.csv", 1
        'Import the .csv into the Access table tblZ_CTA_BMOS_EXTRACT
        If .Show = -1 Then
            strFileName = .SelectedItems(1)
            DoCmd.TransferText acImportDelim, "Z_CTA_BMOS_EXTRACT Link Specification", "tblZ_CTA_BMOS_EXTRACT", strFileName, False
            Else
            Exit Sub
        End If
    End With
    
   'Check to make sure month hasn't aleady been loaded (?)
   'Put check steps here
      
   'Append to dbo_tblCTA_CSR_ShipmentHistory
    DoCmd.OpenQuery "qZ_CTA_BMOS_EXTRACT_ShipmentHistoryApd"
    
   'Append latest YearMonth
    DoCmd.OpenQuery "qYearMonthApd"
          
SkipFileSelect:
          

   
'GoTo TempSkip
   
    'Determine active products in history
    DoCmd.OpenQuery "qCTA_CSR_ProductsActiveClr"
    DoCmd.OpenQuery "qCTA_CSR_ProductsActiveApd"
    
    'Clear and append the Witron extract (manual WeySQL step has to be completed before this query runs)
    DoCmd.OpenQuery "qZ_CTA_BMOS_ProductInfoFromWitronClr"
    DoCmd.OpenQuery "qZ_CTA_BMOS_ProductInfoFromWitronApd"
    
    'Append any new wax code to Xref tables
    'Append any new CorrWax entries
    DoCmd.OpenQuery "qZ_CTA_BMOS_ProductInfoCorrWaxApd"
    
    'Append any new CorrWax entries
    DoCmd.OpenQuery "qZ_CTA_BMOS_ProductInfoConvWaxApd"
               
TempSkip:
               
    Me.Requery
    Me.Refresh
    
    DoCmd.SetWarnings True
    DoCmd.Hourglass False
   
   'Append any new plant/system/soldto/shipto combo's
       
      
   MsgBox "Appended latest month of BW/CSR shipment history and refreshed location xref's" & strFileName
   'MsgBox "Refreshed product info from Witron"
   'MsgBox "Refreshed shipment data from Analysis for Office and Product info from Witron"
 
   End If
   DoCmd.SetWarnings True
   DoCmd.Hourglass False
   
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
   DoCmd.Hourglass False
End Sub


Private Sub btnUpdateProdRatesAndHours_Click()
   DoCmd.SetWarnings False
   DoCmd.Hourglass True
   
   'Refresh tblKiwiHistMachineRatesAndEstAnnualizedHours with the latest YTD data
   DoCmd.OpenQuery "qKiwiHistMachineRatesAndEstAnnualizedHoursClr"
   DoCmd.OpenQuery "qKiwiHistMachineRatesAndEstAnnualizedHoursApd"
   
   'Update tblMachineLocal
    DoCmd.OpenQuery "qKiwiHistMachineLocalUpd"
    
   'Refresh Screen
   Me.Requery
   Me.Refresh
   
   DoCmd.SetWarnings True
   DoCmd.Hourglass False
   MsgBox "Production Rates and Estimated Annual Hours Updated in tblMachineLocal and dbo_tblMachineInfo"


End Sub

Private Sub btnUpdateTheModleFormClose_Click()
        DoCmd.Close acForm, Me.Name
End Sub

Private Sub btnUploadLanesToSQLServer_Click()
   DoCmd.SetWarnings False
   DoCmd.Hourglass True
   
   'Clear dbo_tblCTA_CSR_LanesExternal
   DoCmd.OpenQuery "qZ_CTA_BMOS_LanesExternalClrSQLServer"
   
   'Append from spreadsheet
   DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12Xml, "dbo_tblCTA_CSR_LanesExternal", "C:\OPTMODELS\CV1\MileagesExternalLanesSQLServer.xlsx", True
   
   'Round mileage (needed for PC*Miler)
   'DoCmd.OpenQuery "qFreightCostMileageRound_External"
   
   'Refresh screen
   'Me.Requery
   'Me.Refresh
   
   DoCmd.SetWarnings True
   DoCmd.Hourglass False
   MsgBox "Lanes and Mileages Uploaded to SQL Server from Spreadsheet"
   End Sub

Private Sub btnUploadPlantAndMachineInfo_Click()
   
   DoCmd.SetWarnings False
   DoCmd.Hourglass True
   
   'Clear and Append dbo_tblCTA_CSR_Plant
   DoCmd.OpenQuery "qdbo_tblCTA_CSR_PlantClr"
   DoCmd.OpenQuery "qdbo_tblCTA_CSR_PlantApd"
   
   'Clear and Append dbo_tblMachineInfo
   DoCmd.OpenQuery "qdboMachineInfoClr"
   DoCmd.OpenQuery "qdboMachineInfoApd"
   
   DoCmd.SetWarnings True
   DoCmd.Hourglass False
   
   MsgBox "Machine Info uploaded to SQL Server"
   
End Sub

Private Sub btnViewCustomer_Xref_Click()
  DoCmd.OpenForm "fCustomerSoldToNameCheck"
End Sub

Private Sub btnViewProductInfo_Click()
    DoCmd.OpenQuery "qZ_CTA_BMOS_EXTRACT_ProductDisplay"
End Sub

Private Sub btnViewXrefFullScreen_Click()
    DoCmd.OpenQuery "qShipmentHistoryXrefFullScreen"
End Sub

Private Sub btnViewXrefFullScreenInScope_Click()
    DoCmd.OpenQuery "qShipmentHistoryCleanXrefPlantsInScope"
End Sub


Private Sub Form_Load()
    DoCmd.Maximize
End Sub

Private Sub SelectPlantToAppend_Click()
  Me.Requery
End Sub

===============
Form_fCustomerSoldToNameCheck
===============
Option Compare Database

Private Sub btnCheckPostalCodes_Click()
    DoCmd.SetWarnings False
    DoCmd.Hourglass True

    'Clear and append tblLocationXrefCleanupExtract which is used to check Postal Codes using IntelliRoute
    DoCmd.OpenQuery "qLocationXrefCleanupExtractClr"
    DoCmd.OpenQuery "qLocationXrefCleanupExtractApd"
 
    Me.Requery
    Me.Refresh
    
    DoCmd.SetWarnings True
    DoCmd.Hourglass False
End Sub

Private Sub btnCloseForm_Click()
   DoCmd.Close acForm, Me.Name
End Sub



Private Sub btnRefreshCustomerLocationData_Click()

    DoCmd.SetWarnings False
    DoCmd.Hourglass True

    'Clear and append a customer location cross reference table with BW/CSR extract from Analysis for Excel
    DoCmd.OpenQuery "qZ_CTA_BMOS_LocationXrefClr"
    DoCmd.OpenQuery "qZ_CTA_BMOS_LocationXrefApd"
     
    'Refresh Witron location info
    DoCmd.OpenQuery "qLocationInfoClr"
    DoCmd.OpenQuery "qLocationInfoApd"
    
    'Append any new Plant/SoldTo/Ship/To records to tblLocationXrefCleanup
    DoCmd.OpenQuery "qLocationXrefCleanupApd"
    
    Me.Requery
    Me.Refresh
    
        MsgBox "Any new customer location master data appended"
    
    DoCmd.SetWarnings True
    DoCmd.Hourglass False

End Sub

Private Sub btnRefreshSoldToShipTo_Click()
'Uses location data fixes and updates tblLocationInfo, etc. with cleaner location info
    
    DoCmd.SetWarnings False
    DoCmd.Hourglass True
    
    'Refresh tblZ_CTA_BMOS_LocationXref
    DoCmd.OpenQuery "qZ_CTA_BMOS_LocationXrefClr"
    DoCmd.OpenQuery "qZ_CTA_BMOS_LocationXrefApd"
    
    'Update tblCTA_BMOS_LocationXref with changes made on tblLocationXrefCleanup
    DoCmd.OpenQuery "qZ_CTA_BMOS_LocationXrefCleanupUpd"
      
    'Append SoldTo's to tblZ_CTA_BMOS_SoldTo and update the derived SoldToCode (Common)
    DoCmd.OpenQuery "qZ_CTA_BMOS_SoldToClr"
    DoCmd.OpenQuery "qZ_CTA_BMOS_SoldToApd"
    DoCmd.OpenQuery "qZ_CTA_BMOS_SoldToUpd"
    
    'Append ShipTo's to tblZ_CTA_BMOS_ShipTo and update the derived ShipToCode (Common)
    DoCmd.OpenQuery "qZ_CTA_BMOS_ShipToClr"
    DoCmd.OpenQuery "qZ_CTA_BMOS_ShipToApd"
    DoCmd.OpenQuery "qZ_CTA_BMOS_ShipToUpd"
    
    'Update SoldToCommon and ShipToCommon fields on tblCTA_BMOS_LocationXref
    DoCmd.OpenQuery "qZ_CTA_BMOS_LocationXrefUpdSoldToCommon"
    DoCmd.OpenQuery "qZ_CTA_BMOS_LocationXrefUpdShipToCommon"
        
    'Update the Split State
    DoCmd.OpenQuery "qZ_CTA_BMOS_ShipToUpdSplitState"
        
    'Clear and append dbo_tblCTA_CSR_LocationXref (SQL Server)
    DoCmd.OpenQuery "qZ_CTA_BMOS_LocationXrefClrSQLServer"
    DoCmd.OpenQuery "qZ_CTA_BMOS_LocationXrefApdSQLServer"
    
    'Upload SoldTo data to SQL Server Express
    DoCmd.OpenQuery "qdboSoldToClr"
    DoCmd.OpenQuery "qdboSoldToApd"
    
    'Upload ShipTo data to SQL Server Express
    DoCmd.OpenQuery "qdboShipToClr"
    DoCmd.OpenQuery "qdboShipToApd"
    
    'Upload LocationXref data to SQL Server Express
    DoCmd.OpenQuery "qdboLocationXrefClr"
    DoCmd.OpenQuery "qdboLocationXrefApd"
        
    Me.Requery
    Me.Refresh
    
    DoCmd.SetWarnings True
    DoCmd.Hourglass False
    
    MsgBox "Location info master tables refreshed"
        
End Sub

Private Sub btnUpdateCodeToCity_Click()
    DoCmd.SetWarnings False
    DoCmd.Hourglass True
    
    DoCmd.OpenQuery "qLocationXrefCleanupExtractUpd"
    
    Me.Requery
    Me.Refresh
    
    DoCmd.SetWarnings True
    DoCmd.Hourglass False
    
    MsgBox "Code to City Updated"
    
End Sub
===============
GetRoadMiles
===============

Option Compare Database
'Option Explicit

#If Win64 And VBA7 Then
'Declare PtrSafe Function GetHHGDistance Lib "C:\Program Files\Rand McNally\IntelliRoute\ApiBin\miles32.dll" (ByVal Orig As String, ByVal Dest As String) As Long
Private Declare PtrSafe Function GetHHGDistance Lib "miles32" (ByVal Orig As String, ByVal Dest As String) As LongLong
#Else
'Private Declare Function GetHHGDistance Lib "miles32" (ByVal Orig As String, ByVal Dest As String) As Long
'Declare PtrSafe Function GetHHGDistance Lib "C:\Program Files\Rand McNally\IntelliRoute\ApiBin\miles32.dll" (ByVal Orig As String, ByVal Dest As String) As Long
Private Declare Function GetHHGDistance Lib "miles32" (ByVal Orig As String, ByVal Dest As String) As LongPtr
#End If

Public Sub UpdateMiles()
  On Error GoTo Err_Handler
Dim MyPath As String
Dim GetName As String
Dim Counter As Double
Dim conn As ADODB.Connection
Dim myTableRS As ADODB.Recordset
Dim targetTableName As String
Set myTableRS = New ADODB.Recordset
MyPath = CurrentProject.Path
GetName = CurrentProject.Name




Set conn = Application.CurrentProject.Connection

targetTableName = "tblFreightCost"

    myTableRS.Open targetTableName, conn, adOpenStatic, adLockOptimistic
    
    DoCmd.Hourglass True
    
    Counter = 0
    
    myTableRS.MoveFirst
    
    Do While Not myTableRS.EOF
        DoEvents
            
        Counter = Counter + 1
                    
        Orig = myTableRS.Fields("PlantZIP") & "                 "
        
        'Set mileage to zero for export and skip Milemaker
        'If myTableRS.Fields("ShipMode") = "ZC" Then
        '   Mileage = 0
        '   myTableRS.Fields("mileage") = Mileage
        '   myTableRS.Update
        '   GoTo Skip
        'End If
        
                
        If myTableRS.Fields("ShipToCountry") = "CA" Or myTableRS.Fields("ShipToCountry") = "MX" Then
            Dest = myTableRS.Fields("ShipToZIP") & "                 "
        Else
            Dest = myTableRS.Fields("ShipToZIP") & "                 "
        End If
        
        'Call Milemaker
        Mileage = GetHHGDistance(Orig, Dest)
        myTableRS.Fields("mileage") = Mileage
        myTableRS.Update

        myTableRS.MoveNext
    Loop
    
    myTableRS.Close
    Set myTableRS.ActiveConnection = Nothing
    conn.Close
    
    DoCmd.Hourglass False
    
    MsgBox "External Lane Mileages Calculated"


Err_Handler:
 If Err Then
      MsgBox "Problem with MileMaker" & vbNewLine & Err.Number & " - " & Err.Description
 End If

Exit Sub
Resume
End Sub

===============
modMain
===============
Option Compare Database


Sub AllCodeToDesktop()
''The reference for the FileSystemObject Object is Windows Script Host Object Model
''but it not necessary to add the reference for this procedure.
   Dim fs     As Object
   Dim f      As Object
   Dim strMod As String
   Dim mdl    As Object
   Dim i      As Integer
   
   Dim sfolder      As String
   Dim sHostName    As String
   Dim sUserName    As String
   Dim sUserDomain  As String
   Dim sUserProfile As String
   
   sHostName = Environ$("COMPUTERNAME")
   sUserName = Environ$("USERNAME")
   sUserDomain = Environ$("USERDOMAIN")
   sUserProfile = Environ$("USERPROFILE")
   
   Debug.Print sHostName    '// If InStr(sHostName, "SMACDER") > 0 Thenf
   Debug.Print sUserName
   Debug.Print sUserDomain
   Debug.Print sUserProfile
   Set fs = CreateObject("Scripting.FileSystemObject")
   sfolder = sUserProfile & "\Desktop"
   Set f = fs.CreateTextFile(sfolder & "\" & Replace(CurrentProject.Name, ".", "") & ".txt")
   Debug.Print sfolder
   
   ''For each component in the project ...
   For Each mdl In VBE.ActiveVBProject.VBComponents
      If InStr(1, mdl.Name, "basUtility") = 0 And InStr(1, mdl.Name, "basModule") = 0 And InStr(1, mdl.Name, "InterfaceCommon") = 0 And InStr(1, mdl.Name, "modExcelPivot") = 0 And _
         InStr(1, mdl.Name, "InterfaceSourcingLocks") = 0 And InStr(1, mdl.Name, "basUtility") = 0 And InStr(1, mdl.Name, "basUtility") = 0 Then
         ''using the count of lines ...
         i = VBE.ActiveVBProject.VBComponents(mdl.Name).CodeModule.CountOfLines
         ''put the code in a string ...
         If i > 0 Then
            strMod = VBE.ActiveVBProject.VBComponents(mdl.Name).CodeModule.Lines(1, i)
         End If
         ''and then write it to a file, first marking the start with
         ''some equal signs and the component name.
         f.WriteLine String(15, "=") & vbCrLf & mdl.Name _
             & vbCrLf & String(15, "=") & vbCrLf & strMod
      End If
   Next
   
   ''Close eveything
   f.Close
   Set fs = Nothing
End Sub


Private Sub btnRefreshCustLocAndProd_Click()
  'Check to make sure user really wants to do this
  If MsgBox("Are you sure you want to refresh SQL Server Location and Product Info?", vbYesNo + vbExclamation) = vbYes Then
  
    DoCmd.SetWarnings False
    DoCmd.Hourglass True
       
    'Clear and append to dbo.tblCTA_CSR_LocationXref
    DoCmd.OpenQuery "qZ_CTA_BMOS_LocationXrefClrSQLServer"
    DoCmd.OpenQuery "qZ_CTA_BMOS_LocationXrefApdSQLServer"
    
    'Clear and append to dbo.tblCTA_SoldTo
    DoCmd.OpenQuery "qZ_CTA_BMOS_EXTRACT_SoldToClrSQLServer"
    DoCmd.OpenQuery "qZ_CTA_BMOS_EXTRACT_SoldToApdSQLServer"
    
    'Clear and append to dbo.tblCTA_ShipTo
    DoCmd.OpenQuery "qZ_CTA_BMOS_EXTRACT_ShipToClrSQLServer"
    DoCmd.OpenQuery "qZ_CTA_BMOS_EXTRACT_ShipToApdSQLServer"
    
    'Clear and append to dbo.tblCTA_CSR_Product
    DoCmd.OpenQuery "qZ_CTA_BMOS_EXTRACT_ProductClrSQLServer"
    DoCmd.OpenQuery "qZ_CTA_BMOS_EXTRACT_ProductApdSQLServer"


'MsgBox "Manufacturing options appended!  "

  End If
  

  DoCmd.SetWarnings True
  DoCmd.Hourglass False

End Sub


Public Sub LoopThroughPlantsInScope()
'**********************************************
'Author  :  KWH
'Date/Ver:  2018-08-02 V00
'Input   :  tblPlant; tblProduct ; tblMachine
'Output  :
'Comments: Used to loop through plants in scope and run query to append to "MakeProd" table
'**********************************************

On Error GoTo Err_Handler

Dim strDatabaseName As String
Dim strDBCursorType As String
Dim strDBLockType As String
Dim strDBOptions As String
Dim rs As ADODB.Recordset
Dim cn As ADODB.Connection

Dim strUserName As String
Dim strSQL As String
Dim strPlantCode As String
Dim strPlantDesc As String

Dim Counter As Integer

Dim lRet As String
Dim strMessage As String

Dim qdf As QueryDef


strDBCursorType = adOpenDynamic
strDBLockType = adLockOptimistic
strDBOptions = adCmdText

'Remove records from tblMakeProdAppendProgressStatus
DoCmd.OpenQuery "qMakeProdAppendProgressStatusClr"

'Remove record from tblPlantCurrent
DoCmd.RunSQL "Delete * From tblPlantCurrent"

'Get plants in scope
Set rs = New ADODB.Recordset

strSQL = "SELECT PlantCode, PlantDesc FROM tblPlant WHERE InScope=-1"
rs.Open strSQL, CurrentProject.Connection, strDBCursorType, strDBLockType, strDBOptions

'Clear dbo_tblMakeProd on SQL Server (Express)
'DoCmd.OpenQuery "qdboMakeProdClr"

'Clear dbo_tblMakeProd on SQL Server
'DoCmd.OpenQuery "qdboMakeProdClrSQLServer"

'Clear the dbo_tblMakeProdAvg
DoCmd.OpenQuery "qdboMakeProdAvgClr"

'Execute qMakeProdApd by plant
Do While Not rs.EOF

    strPlantCode = rs!PlantCode
    strPlantDesc = rs!PlantDesc
        
    'Write the current plant tblPlantCurrent
    strSQL = "INSERT INTO tblPlantCurrent (PlantCode,PlantDesc)" & _
             "VALUES('" & strPlantCode & "','" & strPlantDesc & "');"
    'Debug.Print strSQL
    Call Application.CurrentProject.Connection.Execute(strSQL, lRet)
            
            
    'Run qdboMakeProdApdSQLServerExpress which has an inner join that limits it to a single plant at at time
    'DoCmd.OpenQuery "qdboMakeProdApd"
    
    'Append to SQL Server
    'DoCmd.OpenQuery "qdboMakeProdApdSQLServer"
    
    'Append directly to SQL Server Express dbo_tblMakeProdAvg (short cut that bypasses routing/mfgcost and uses keys to prevent dups...)
    DoCmd.OpenQuery "qdboMakeProdAvgClrALT_Prelim"
    DoCmd.OpenQuery "qdboMakeProdAvgApdALT_Prelim"
    
    'Update the freight cost info on dbo_tblMakeProdAvg_Prelim
    DoCmd.OpenQuery "qdboMakeProdAvgUpdPrelimFreightCost"
    
    'Append the Prelim table to dbo_tblMakeProdAvg
    DoCmd.OpenQuery "qdboMakeProdAvgApdALT_Main"
   
    'Clear record in tblCurrentPlant
    DoCmd.RunSQL "Delete * From tblPlantCurrent"
    
    'Append plant info and finish time to tblMakeProdAppendProgressStatus
    strSQL = "INSERT INTO tblMakeProdAppendProgressStatus (PlantCode,PlantDesc,FinishTime)" & _
             "VALUES('" & strPlantCode & "','" & strPlantDesc & "',Now());"
    'Debug.Print strSQL
    Call Application.CurrentProject.Connection.Execute(strSQL, lRet)
    
    Forms!fUpdateTheModel!subformMakeProdApdProgress.Form.Requery
    
    DoEvents
     
rs.MoveNext
Loop

  
If Not rs Is Nothing Then
    rs.Close
    Set rs = Nothing
    
End If

        
Err_Handler:
 If Err Then
      rs.Close
      Set rs = Nothing
      DoCmd.SetWarnings True
      DoCmd.Hourglass False
      MsgBox "Problem appending to dbo_tblMakeProd" & vbNewLine & Err.Number & " - " & Err.Description
 End If

End Sub

