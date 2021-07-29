Public Function GroupShipmentsCode(Optional sTableName = "tblFS1415_LoadSummaryAbleTest", _
                               Optional sSortFieldName = "GroupSortPriority", _
                               Optional dMaxCubes = 3130, Optional dMaxWeight = 24000, _
                               Optional iMaxStops = 3, _
                               Optional iMaxCubesToGroup = 2000) _
                               As Boolean
                               
'//=================================================================//
'/|   FUNCTION:  GroupShipments                                     |/
'/| PARAMETERS:  Requires tblName, SortField, MaxCubes, MaxWeight,  |/
'/|                                                                 |/
'/|              Table Must include fields GroupID, and GroupCode   |/
'/|              and ShipSite and Cubes and Weight                  |\
'/|                                                                 |/
'/|    RETURNS:  True on Success                                    |/
'/|    PURPOSE:  Group Shipments into truckLoads by weight          |/
'/|              limit or cubic foot limit                          |/
'/|      USAGE:  i= GroupShipments()                                |/
'/|              i= GroupShipments("TBL1","PSORT", 2800, 24000, 2)  |/
'/|         BY:  Sean                                               |/
'/|       DATE:  08/11/2015                                         |/
'/|    HISTORY:  08/11/2015                                         |/
'/|                                                                 |/
'//=================================================================//

'// i= GroupShipments("tblFS1415_LoadSummaryAble","GroupSortPriority", _
'                       0.85 * CDbl(GetSetting("FullTruckCubicFeet")), _
'                       0.85 * CDbl(GetSetting("FullTruckWeight")), _
'                       2 )
'   i= GroupShipments("t114NETPDelCostGroupShipments","GroupSortPriority", 3130, 24000, 3, 2000)

'//============  ** Enhance This:

Dim db         As Database
Dim rs         As DAO.Recordset
Dim rsClone    As DAO.Recordset
Dim rsShipFit  As DAO.Recordset
Dim sSQL       As String
Dim i          As Integer
Dim T          As Date
Dim s          As String
Dim key        As String
Dim keyTimeBucket As String
Dim weightleft As Double
Dim cubesleft  As Double
Dim sShipSite  As String
Dim lGroupCtr  As Long
Dim PctTolerance As Double
Dim lRecCtr    As Long
Dim iGrpStops  As Integer
Dim weightAvailPct As Double
Dim cubesAvailPct  As Double
Dim fNewRecordFlag As Boolean




   On Error GoTo GroupShipmentsCode_Err
   GroupShipmentsCode = False

   T = Now()

   Set db = CurrentDb
   sSQL = ""
   iGrpStops = 0
   cubesAvailPct = 0
   weightAvailPct = 0
   PctTolerance = 0.1
   fNewRecordFlag = False

   '//  G R O U P   T H E   S H I P M E N T S  //
   sSQL = "SELECT * FROM " & sTableName & " " & _
          "WHERE " & sTableName & ".Cubes <= " & iMaxCubesToGroup & " " & _
          "ORDER BY " & sSortFieldName & " ASC "
   Set rs = db.OpenRecordset(sSQL)
     
   Set rsClone = rs.Clone()
   rsClone.MoveFirst
   lRecCtr = 1

   'INIT GROUPS
   DoCmd.SetWarnings False
   DoCmd.RunSQL ("UPDATE " & sTableName & " SET " & sTableName & ".GroupID = 0, " & sTableName & ".GroupCode = '0'")
   DoCmd.SetWarnings True
   
   key = rsClone!ShipSite
   keyTimeBucket = rsClone!Year & rsClone!Month
   '''& rsClone!WK
   cubesleft = dMaxCubes
   weightleft = dMaxWeight
   lGroupCtr = 1


   While rsClone.EOF = False And lGroupCtr <= rsClone.RecordCount
      DoEvents
      '''If rsClone!id = 12512 Then Debug.Print "     ##### 12512 #####    "
     'IF IT FITS AND IT'S THE RIGHT FACILITY, TimeBucket and is still unassigned to a group, START OR ADD TO THE GROUP
      If rsClone!Cubes <= cubesleft And rsClone!Weight <= weightleft And rsClone!ShipSite = key And _
         rsClone!Year & rsClone!Month = keyTimeBucket And _
         rsClone!GroupID = 0 And _
         iGrpStops <= iMaxStops Then     '//Group this line in and fill in the Group ID and Code
         '''& rsClone!WK
         rsClone.Edit
         rsClone!GroupID = lGroupCtr
         rsClone!GroupCode = lGroupCtr
         cubesleft = cubesleft - rsClone!Cubes
         weightleft = weightleft - rsClone!Weight
         cubesAvailPct = cubesleft / dMaxCubes
         weightAvailPct = weightleft / dMaxWeight
         rsClone.Update
         iGrpStops = iGrpStops + 1
         
''''         If iGrpStops = 1 Then Debug.Print lGroupCtr & " . . . "
''''         Debug.Print "   " & "recs seen = " & lRecCtr & " . " & "currTableID " & rsClone!id & " assigned " & rsClone!Weight & " lbs. "
         key = rsClone!ShipSite
         lRecCtr = lRecCtr + 1
         rsClone.MoveNext
         fNewRecordFlag = False
      Else
         key = rsClone!ShipSite
         rsClone.MoveNext
         lRecCtr = lRecCtr + 1
      End If
      
    
       
         
     'IF IT IS ENOUGH
      If cubesAvailPct <= PctTolerance Or weightAvailPct <= PctTolerance Or iGrpStops = iMaxStops Then
         lGroupCtr = lGroupCtr + 1           'increment counter to next Group
         rsClone.MoveFirst
         rsClone.FindFirst "[GroupID] = 0 And [Cubes] <= " & dMaxCubes & " AND [Weight] <= " & dMaxWeight & " " 'go back to first unassigned record

         ''' Alternatively, reset the rsClone right here to only open candidates

         lRecCtr = lRecCtr + 1
         cubesleft = dMaxCubes
         weightleft = dMaxWeight
         cubesAvailPct = cubesleft / dMaxCubes
         weightAvailPct = weightleft / dMaxWeight
         iGrpStops = 0
         key = rsClone!ShipSite
         
''''         Debug.Print "currTableID " & rsClone!id
         '''& rsClone!WK
      Else
         If rsClone.EOF = False Then
            key = rsClone!ShipSite
            rsClone.MoveNext
            lRecCtr = lRecCtr + 1
         End If
      End If
      keyTimeBucket = rsClone!Year & rsClone!Month
   Wend
    
      
   T = Now() - T
   s = "GroupShipmentsCode....." & Format(T, "hh:nn:ss")
   Debug.Print s
   Debug.Print "--DONE--"
   
   GroupShipmentsCode = True
   
GroupShipmentsCode_Done:
  Exit Function

GroupShipmentsCode_Err:
   Select Case Err
    Case 9 'or 13 'subscript out of range
      GroupShipmentsCode = False
    Case 3021 'no current record
      GroupShipmentsCode = True
      Resume GroupShipmentsCode_Done
    Case Else
      MsgBox "Error: " & Error$ & " " & Err, , "GroupShipmentsCode"
      GroupShipmentsCode = False
   End Select
   Resume GroupShipmentsCode_Done
   Resume
   
   
End Function
