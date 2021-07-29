Function RunRollLoad2() As Boolean
'//=====================================================//
'/|   FUNCTION:  RunRollLoad2                           |/
'/| PARAMETERS:                                         |/
'/|    RETURNS:  -NONE-                                 |/
'/|    PURPOSE:  Determine if load is loadable or not,  |/
'/|              and if so, if it is regular loadable   |/
'/|              or T-Loading/RollBack loadable only    |/
'/|  Not to be used for detailed load planning like is  |/
'/|  found here: www.loadxpert.com/lxlp/en/lxlp.htm     |/
'/|      USAGE:  i= RunRollLoad2()                      |/
'/|         BY:  Sean                                   |/
'/|       DATE:  10/30/2013                             |/
'/|    HISTORY:  10/30/2013                             |/
'/|                 adapted into IPGPlannerAlertsSQL    |/
'/|                              Server_V23 and the     |/
'/|                              GSCCOERPT SQL DB       |/
'/|                                                     |/
'/|                                                     |/
'/|                                                     |/
'//=====================================================//

'//============  ** Allow for T-Loading aka RollBack loading
'//============  ** DONE Expand so that number of floor spots is "by Cartype (R1..R4)"
'//============  ** DONE Expand so that MAXIMUM STACK HT IS "by MILL AND Cartype (R1..R4)"
'//============  ** DONE Expand so that roll diameter is not constant but read in from the roll file
'//============  ** Let it go unalerted if UnloadedRolls <=8 and RBA is allowed.
'//============  ** DONE Filter input into RCRolls to only ~58"~ Rolls

   On Error GoTo RunRollLoad2_Err
   RunRollLoad2 = False
   
   Dim db As Database
   Set db = CurrentDb
   
   Dim dal As clsDAL
   Dim ret As Long
    
   Set dal = New clsDAL
   dal.Connect e_db_ACCESS, e_connType_E, Application.CurrentProject.FullName
   dal.bPersist = True
   'ret = dal.Execute(sql, GetData)
   
   Dim dAllowableFloorSpots  As Double  'Like 20
   Dim dAllowableStackSpots  As Double  'Like 2 or 3
   Dim dMaximumStackHeight   As Double  'Like 152    (inches)
   Dim dStandardRollDiameter As Double  'Like  54.25 (inches)
   
   dAllowableFloorSpots = 99.9999
   dAllowableStackSpots = 9
   dMaximumStackHeight = 999.999
   dStandardRollDiameter = 99.9999
   
'''   dAllowableFloorSpots = Application.CurrentProject.Connection.Execute( _
'''                             "SELECT dbo_tblPAS_103Settings.AllowableFloorSpots FROM dbo_tblPAS_103Settings").Fields("AllowableFloorSpots").value
'''   dAllowableStackSpots = Application.CurrentProject.Connection.Execute( _
'''                             "SELECT AllowableStackSpots FROM dbo_tblPAS_103Settings").Fields("AllowableStackSpots").value
'''   dMaximumStackHeight = Application.CurrentProject.Connection.Execute( _
'''                             "SELECT MaximumStackHeight FROM dbo_tblPAS_103Settings").Fields("MaximumStackHeight").value
'''   dStandardRollDiameter = Application.CurrentProject.Connection.Execute( _
'''                             "SELECT StandardRollDiameter FROM dbo_tblPAS_103Settings").Fields("StandardRollDiameter").value
   
   
   '// T-LOADING (RollBack Loading) IS ONE WHERE ROLLS ARE STACKED LIKE CANS BUT THE TOP ONE IS LIKE A CAN ON ITS SIDE
   
   Dim iTotalRolls         As Integer    'Like 38
   Dim iRollsOnSpot        As Integer
   Dim C, i, j, K, z       As Integer
   Dim dTLoadHeight        As Double     'Current Load of Stack if a TLoad  '>?? is this needed
   Dim dStackHeight        As Double     'Current Load of Stack
   Dim dxRollWidth         As Dictionary
   Dim dxRollStatus        As Dictionary
   Dim dxFloorSpotStatus   As Dictionary
   Dim sCarID              As String
   Dim sZstatus            As String
   Dim sMillCar            As String
         
   iTotalRolls = 0
   dTLoadHeight = 0
   dStackHeight = 0
   iRollsOnSpot = 0
      
   Dim sSQL         As String
   Dim rsRC         As DAO.Recordset   'why won't ado work???   ADODB.Recordset 'RailCar
   Dim rsRolls1     As DAO.Recordset   'why won't ado work???   ADODB.Recordset 'Rolls
   'Dim rsRolls2     As DAO.Recordset   'why won't ado work???   ADODB.Recordset 'Rolls
   'Dim rsRolls3     As DAO.Recordset   'why won't ado work???   ADODB.Recordset 'Rolls
   Dim rsDest       As ADODB.Recordset 'ADO ONE
   
   
   'clear  RCRolls  (temptblRCRolls), which  is the
   'RailCar & Rolls Table --  all mills all cars, and then
   'append RCRolls from PPR -- all 16 mills.
   C = RunQueriesByPrefix("qry103Init", True)
   
   
   sSQL = "SELECT DISTINCT " & _
          "   Trim(temptblRCRolls.CarID) as CarID, " & _
          "   Trim(First(temptblRCRolls.MillCode)) AS MillCode,  " & _
          "   Trim(First(temptblRCRolls.ShipmentStatus)) AS ShipmentStatus,  " & _
          "   Trim(First(temptblRCRolls.CarType)) AS CarType,  " & _
          "   Trim(First(temptblRCRolls.ShipToName)) AS ShipToName,  " & _
          "   Trim(First(temptblRCRolls.ShipToCity)) AS ShipToCity,  " & _
          "   Trim(First(temptblRCRolls.ShipToState)) AS ShipToState,  " & _
          "   Trim(First(temptblRCRolls.ShipToID)) AS ShipToID,  " & _
          "   Trim(First(temptblRCRolls.ID)) AS OneID,  " & _
          "   Trim(First(dbo_tblPAS_lkpMillCarType.MaximumStackHeight)) AS MaximumStackHeight " & _
          "FROM  " & _
          "   temptblRCRolls INNER JOIN dbo_tblPAS_lkpMillCarType ON  " & _
          "       (temptblRCRolls.CarType = dbo_tblPAS_lkpMillCarType.CarTypeCode) AND  " & _
          "       (temptblRCRolls.MillCode = dbo_tblPAS_lkpMillCarType.MillCode) " & _
          "GROUP BY Trim(temptblRCRolls.CarID) " & _
          "ORDER BY Trim(First(temptblRCRolls.ID)) "

   Set rsRC = Nothing
   Set rsRC = db.OpenRecordset(sSQL)
   
   'put a check right here for "if no records, message 'no rolll records' and quit
   rsRC.MoveLast
   rsRC.MoveFirst
   
   '''DoCmd.RunSQL "DELETE * FROM temptblResults"
   ''''''''''''''''''''''''''''''''''''''c = RunQueriesByPrefix("qdel_PAS_103Results", True)
   
   sSQL = "SELECT * FROM temptblResults"
   Set rsDest = New ADODB.Recordset
   rsDest.LockType = adLockBatchOptimistic
   rsDest.CursorLocation = adUseClientBatch
   rsDest.Open sSQL, Application.CurrentProject.Connection, adOpenKeyset, adLockBatchOptimistic, adCmdText
      
      
'--------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------------------------------------
   'LOOP (RAILCAR)  'RECORDSET rsRC
   While Not rsRC.EOF
      DoEvents
      Set rsRolls1 = Nothing
      sSQL = "SELECT * FROM temptblRCRolls WHERE CarID = '" & rsRC.Fields("CarID").value & "' ORDER BY CarID ASC, RollWidth DESC, ID ASC"
      Set rsRolls1 = db.OpenRecordset(sSQL)
          rsRolls1.MoveFirst
      '''Set rsRolls2 = rsRolls1.Clone()
      '''Set rsRolls3 = rsRolls1.Clone()

      'Set dxRollWidth = New Dictionary
      Set dxRollStatus = New Dictionary
          dxRollStatus.CompareMode = BinaryCompare       '// set compare mode
      Set dxFloorSpotStatus = New Dictionary
          dxFloorSpotStatus.CompareMode = BinaryCompare  '// set compare mode
      
      iTotalRolls = 0
      dTLoadHeight = 0
      dStackHeight = 0
      sCarID = rsRC!CarID
      dAllowableFloorSpots = DLookup("CarTypeFloorSpots", "dbo_tblPAS_lkpCarType", "CarTypeCode = '" & rsRC!CarType & "'")
      dMaximumStackHeight = rsRC!MaximumStackHeight
      sMillCar = rsRC("MillCode").value & rsRC("CarType").value
      
      'add not one, but TWO floorspot for NEWPORT R2'S AND R3'S
      Select Case sMillCar
         Case "NPR2", "NPR3"
            dAllowableFloorSpots = dAllowableFloorSpots + 2   'Changed Newport Additional Spots from 1 to 2  4/22/2015
         Case Else
      End Select
      
      'LOOP OVER ROLLS TO POPULATE / INITIALIZE  ROLLSTATUS VARIABLES
      While Not rsRolls1.EOF
         dxRollStatus(VBA.Trim(rsRolls1!CarID) & "." & VBA.Trim(rsRolls1!Id)) = 0
            ' .Add Key:="mike", Item:=22    '// add item using named arguments
            ' .Add "joe", 33   'add item without named arguments this would also work
         rsRolls1.MoveNext
      Wend
      'ENDLOOP OVER ROLLS
      rsRolls1.MoveFirst
      
      'LOOP OVER FLOORSPOTS TO POPULATE / INITIALIZE  FLOOR-SPOT-STATUS VARIABLES
      For z = 1 To dAllowableFloorSpots
         dxFloorSpotStatus(z) = 0
      Next z
      'ENDLOOP OVER FLOORSPOTS
      
      
      '==LOOP OVER FLOOR SPOTS (Z-LOOP)======================================================================
      For z = 1 To dAllowableFloorSpots
         'wrap this whole Z Loop in an IF statement??
         sZstatus = "ALERT"
         rsRolls1.MoveFirst
         dStackHeight = 0 'init
         iRollsOnSpot = 0 'init
                  
         '--LOOP (ROLLS)-------------------------------------------------------------------------------------
         While Not rsRolls1.EOF
            '
            If dxFloorSpotStatus(z) = 0 And sZstatus <> "DONE" And _
               dxRollStatus(VBA.Trim(rsRolls1.Fields("CarID").value) & "." & VBA.Trim(rsRolls1.Fields("ID").value)) = 0 Then
               
               'IF ROLL STACK FITS THEN LOAD BOTH ROLLS ONTO THE FLOOR SPOT
               If dStackHeight + rsRolls1.Fields("RollWidth").value <= dMaximumStackHeight Then
                  
                  'add the stack and assign roll i and roll j to floor spot z
                  dStackHeight = dStackHeight + rsRolls1.Fields("RollWidth").value
                  'dTLoadHeight = dTLoadHeight + rsRolls1.Fields("RollWidth").value
                  
                  iRollsOnSpot = iRollsOnSpot + 1
                  
                  dxRollStatus(VBA.Trim(rsRolls1!CarID) & "." & VBA.Trim(rsRolls1!Id)) = z
                  '''dxRollStatus(VBA.Trim(rsRolls2!CarID & "." & rsRolls2!Id)) = z
                  
                  If iRollsOnSpot = 1 Then
                     rsDest.AddNew
                     rsDest!CarID = VBA.Trim(rsRolls1!CarID)
                     rsDest!CarType = VBA.Trim(rsRC!CarType)
                     rsDest!RType = "Spot"
                     rsDest!FloorSpot = z
   
                  End If
                  
                  '*&*&*&*&
                  rsDest!StackHeight = dStackHeight
                  rsDest!Roll1id = rsDest!Roll1id & " " & iRollsOnSpot & "." & rsRolls1.Fields("RollWidth").value
                  '*&*&*&*&
                  
                  
                  
                  '''rsDest!Roll1width = VBA.Trim(rsRolls1!RollWidth)
                  '''rsDest!Roll2id = VBA.Trim(rsRolls2!Id)
                  '''rsDest!Roll2width = VBA.Trim(rsRolls2!RollWidth)
               
                  
                  iTotalRolls = iTotalRolls + 1
                  'RESET STACK HEIGHT, ETC.  AND MOVE ON AND JUMP OUT TO NEXT Z
                  '''dStackHeight = 0
                  'sZstatus = "DONE"
               End If
            End If
            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            rsRolls1.MoveNext
         Wend
         'ENDLOOP ROLL1
         '*******&*&*&*&*&******* *******&*&*&*&*&******* *******&*&*&*&*&******* *******&*&*&
         If iTotalRolls = rsRolls1.RecordCount Then
            sZstatus = "DONE"
         End If
      
      Next z
      'ENDLOOP FLOOR SPOTS
      '===================================================================================================
      
     
      If iTotalRolls = rsRolls1.RecordCount Then
         sZstatus = "OK"
      Else
         sZstatus = "ALERT"
      End If
      
      sZstatus = "[" & sZstatus & "]" '& Space(13 - Len("[" & sZstatus & "]")) & _
                  '"     TOTAL ROLLS IN " & rsRC!CarID & "= " & iTotalRolls & " OF " & rsRolls1.RecordCount & " TOTAL in temptblRCRolls."
      If iTotalRolls = 0 Then
         sZstatus = sZstatus & "           ***  NO ROLLS STACKED OR LOADED  ***     " & VBA.Trim(rsRC.Fields("CarID").value)
      End If
      
      
      rsDest.AddNew
      rsDest!CarID = VBA.Trim(rsRC!CarID)
      rsDest!shipmentStatus = VBA.Trim(rsRC!shipmentStatus)
      rsDest!ShipToName = VBA.Trim(rsRC!ShipToName)
      rsDest!ShipToCity = VBA.Trim(rsRC!ShipToCity)
      rsDest!ShipToState = VBA.Trim(rsRC!ShipToState)
      rsDest!ShipToID = VBA.Trim(rsRC!ShipToID)
      rsDest!CarType = VBA.Trim(rsRC!CarType)
      rsDest!StackHeight = dMaximumStackHeight
      rsDest!FloorSpot = dAllowableFloorSpots
      rsDest!RType = "TOTAL"
      rsDest!Roll1id = VBA.Left(rsRC!OneID, 2)    '// MILL CODE
      rsDest!Status = sZstatus
      rsDest!CarRollTotal = rsRolls1.RecordCount
      rsDest!CarRollsLoaded = iTotalRolls
      rsDest!CarRollsUnLoaded = rsRolls1.RecordCount - iTotalRolls
            
      Debug.Print VBA.Trim(rsRC!CarID) & "    " & VBA.Left(rsDest!Roll1id, 2) & "    " & sZstatus
            
      rsRC.MoveNext
   
   Wend
   'ENDLOOP RAILCAR
'--------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------------------------------------
   rsDest.UpdateBatch adAffectAllChapters
   
   C = RunQueriesByPrefix("qry103Push", True)
   Debug.Print "--DONE--"
   
   RunRollLoad2 = True
   
RunRollLoad2_Done:
  Exit Function

RunRollLoad2_Err:
   Select Case Err
    Case 9 'or 13 'subscript out of range
      RunRollLoad2 = False
    Case Else
      MsgBox "Error: " & Error$ & " " & Err, , "RunRollLoad"
      RunRollLoad2 = False
   End Select
   Resume RunRollLoad2_Done
   Resume
   
End Function