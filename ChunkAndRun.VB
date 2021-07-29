Option Compare Database
Option Explicit


Public Function ChunkAndRun(Optional sTableName = "TableX", Optional sSortFieldName = "GroupSortPriority", _
                        Optional dMaxCubes = 2700, Optional dMaxWeight = 24000, Optional iMaxStops = 3, _
                        Optional iMaxCubesToGroup = 2399, Optional dPctTolerance = 0.01, _
                        Optional dMaxDeltaDays = 5, Optional dMaxDeltaZip2 = 5, _
                        Optional dTgtFirstDeliverySizeInCubes = 2000 _
                       ) As Boolean
'//=====================================================================================================================//
'/|   FUNCTION:  ChunkAndRun                 RETURNS:  True on Success                                                  |/
'/| PARAMETERS:  Optional:  sTableName     sSortFieldName dMaxCubes     dMaxWeight iMaxStops iMaxCubesToGroup           |/
'/|                         dPctTolerance  dMaxDeltaDays  dMaxDeltaZip2                                                 |/
'/|              Table fields:  ID, ShipSite, GroupID, GroupCode, Zip2Numeric, Cubes, Weight, ShipDate, Year, Month, WK |/
'/|    PURPOSE:  Group Shipments into truckLoads by weight limit or cubic foot limit                                    |/
'/|      USAGE:  i= ChunkAndRun()                                                                                       |/
'/|      USAGE:  i= ChunkAndRun("TableX", "GroupSortPriority", 2750, 50000, 3, 2500, 0.10, 3, 2, 1650)                  |/
'/|         BY:  Sean                                                                                                   |/
'/|       DATE:  08/21/2015                                                                                             |/
'/|    HISTORY:  08/21/2015                                                                                             |/
'/|                                                                                                                     |/
'//=====================================================================================================================//


   '//============  Declarations and Things
   '//==================================================================================================================
   Dim db          As Database
   Dim rs          As DAO.Recordset
   Dim rs2Fit      As DAO.Recordset
   Dim rsChunk     As DAO.Recordset
   Dim sSQL        As String
   Dim i           As Integer
   Dim T           As Date
   Dim ET          As Date
   Dim weightleft  As Double
   Dim cubesleft   As Double
   Dim sShipSite   As String
   Dim bDebug      As Boolean
   Dim lGroupNum   As Long
   Dim iGrpStops   As Integer
   Dim s           As String

   On Error GoTo ChunkAndRun_Err
   ChunkAndRun = False


   '//============  Initialize Variables
   '//=====================================================================================================================
   T = Now()
   ET = Timer
   Set db = CurrentDb
   sSQL = ""
   iGrpStops = 0
   lGroupNum = 1
   cubesleft = dMaxCubes
   weightleft = dMaxWeight
   bDebug = True
   Debug.Print "START / INITIALIZE ............ ....             " & Format(Timer, "ss")


   '//============  Initialize the Groups, setting all group ID's and group Codes to "0"
   '//=====================================================================================================================
   DoCmd.SetWarnings False
   DoCmd.RunSQL ("UPDATE " & sTableName & " SET " & sTableName & ".GroupID = 0, " & sTableName & ".GroupCode = '0'")
   DoCmd.RunSQL ("DELETE * FROM TableZ")
   DoCmd.RunSQL ("DELETE * FROM TableY")
   DoCmd.SetWarnings True
   Debug.Print "INIT GROUP_IDS AND CLEAR TABLES y & z...        " & Format(Timer, "ss")


   '//============  Set rsChunk to the list of combo's of ShipSite, Yr and Month so that the 'outer' loop is over fewer recs
   '//=====================================================================================================================
   Set rsChunk = db.OpenRecordset("qryChunk")  'typically master list of PLANT (ShipSite), Year, REGN (Region)
   rsChunk.MoveFirst
   Debug.Print "open rsChunk for the master index to the outer loop         " & Format(Timer, "ss")


   '//============  rsChunk Loop -- Grab a Chunk of records to group at once (saves candidate qry work and time)
   '//=====================================================================================================================
   While Not rsChunk.EOF
      DoEvents
      '##
      If bDebug = True Then Debug.Print "Top of CHUNK Loop ....             " & Format(Timer, "ss")

      'Clear and Append The Next Chunk to TableY
      DoCmd.SetWarnings False
      DoCmd.RunSQL ("DELETE * FROM TableY")
      sSQL = "INSERT INTO TableY ( TableXID, Delivery, ShipSite, ShipDate, [Year], [Month], WK, Mode, Weight, Cubes, Dest, " & vbCrLf & _
             "                     Zip2, Zip2Numeric, Zip2State, Zip3, DestZipCode, REGN, GroupSortPriority, GroupID, GroupCode ) " & vbCrLf & _
             "SELECT ID, Delivery, ShipSite, ShipDate, Year, Month, WK,  " & vbCrLf & _
             "       Mode, Weight, Cubes, Dest, Zip2, Zip2Numeric, Zip2State,  " & vbCrLf & _
             "       Zip3, DestZipCode, REGN, GroupSortPriority, GroupID, GroupCode  " & vbCrLf & _
             "FROM  " & sTableName & " " & vbCrLf & _
             "WHERE (((ShipSite)='" & rsChunk!ShipSite & "') AND ((Year)=" & rsChunk!Year & ") AND ((REGN)='" & rsChunk!REGN & "')) " & vbCrLf
      DoCmd.RunSQL (sSQL)
      DoCmd.SetWarnings True

      '//============  Set rs to the table full of deliveries in the ChunkAndRun_Done that is used for the 'outer' loop.
      '//=====================================================================================================================
      sSQL = "SELECT * " & vbCrLf & _
             "FROM " & "TableY" & " " & vbCrLf & _
             "WHERE Cubes <= " & dMaxCubes & " AND Cubes <= " & iMaxCubesToGroup & _
             "  AND Weight <= " & dMaxWeight & " " & vbCrLf & _
             "ORDER BY " & sSortFieldName & " ASC "
      Set rs = db.OpenRecordset(sSQL)
      rs.MoveFirst
      Debug.Print "open rs for the outer loop         " & Format(Timer, "ss")
   
      
      '//============  Outer Loop on Table of Deliveries (e.g. 38,000 for FS year, or 7,400 for FS to NorthEast U.S.)
      '//=====================================================================================================================
      While Not rs.EOF
         DoEvents
         '##
         If bDebug = True And lGroupNum = 3 Then Debug.Print "Top of Outer Loop ....             " & Format(Timer, "ss")
         
         If rs!Cubes <= cubesleft And rs!Weight <= weightleft And _
            rs!GroupID = 0 And iGrpStops = 0 Then     '//Group this line in and fill in the Group ID and Code for Stop 1 (Anchor)
   
            'Assign it and update the table GroupID and GroupCode
            iGrpStops = iGrpStops + 1
           'key = rs!ShipSite
            rs.Edit
            rs!GroupID = lGroupNum
            rs!GroupCode = lGroupNum & "." & iGrpStops
            rs.Update
            cubesleft = cubesleft - rs!Cubes
            weightleft = weightleft - rs!Weight
            
            If bDebug = True And lGroupNum = 3 Then Debug.Print "Init candidates   ....             " & Format(Timer, "ss")
   
            'Subset TableY even further into a table, TableYCandidates with just the candidates
            'Clear and Append The Candidates to TableYCandidates
            DoCmd.SetWarnings False
            DoCmd.RunSQL ("DELETE * FROM TableYCandidates")
            sSQL = "INSERT INTO TableYCandidates ( TableXID, Delivery, ShipSite, ShipDate, [Year], [Month], WK, Mode, Weight, Cubes, Dest, " & vbCrLf & _
             "                     Zip2, Zip2Numeric, Zip2State, Zip3, DestZipCode, REGN, GroupSortPriority, GroupID, GroupCode ) " & vbCrLf & _
             "SELECT TableXID, Delivery, ShipSite, ShipDate, Year, Month, WK,  " & vbCrLf & _
             "       Mode, Weight, Cubes, Dest, Zip2, Zip2Numeric, Zip2State,  " & vbCrLf & _
             "       Zip3, DestZipCode, REGN, GroupSortPriority, GroupID, GroupCode  " & vbCrLf & _
             "FROM  " & " TableY " & " " & vbCrLf & _
             "WHERE " & " " & vbCrLf & _
             "   ShipSite='" & rsChunk!ShipSite & "' AND Year = " & rsChunk!Year & " AND REGN='" & rsChunk!REGN & "' " & vbCrLf & _
             "    AND Abs( " & rs!Zip2Numeric & " - [Zip2Numeric]) <= " & dMaxDeltaZip2 & "  AND Cubes <= " & cubesleft & "   " & vbCrLf & _
             "    AND Abs([ShipDate]- #" & rs!ShipDate & "# ) <= " & dMaxDeltaDays & "  "
      DoCmd.RunSQL (sSQL)
      DoCmd.SetWarnings True
   
            Set rs2Fit = Nothing
            
            'Right here, you know what you've got as the first shipment, go build the SQL to return eligible candidates to
            'join (there may be none).  If there are some return the best 5 singletons and doubletons.
            
            sSQL = GenSQLToGetGroupCandidates(rs, CStr("TableYCandidates"), CDbl(cubesleft), CDbl(dMaxCubes), CDbl(dMaxWeight), _
                                              CDbl(dMaxDeltaDays), CDbl(dMaxDeltaZip2), "BOTH")   'SINGLE, DOUBLE, OR BOTH
            
            If bDebug = True And lGroupNum = 3 Then Debug.Print "got sSQL for cand ....             " & Format(Timer, "ss")
            
            Set rs2Fit = db.OpenRecordset(sSQL)
            
            If bDebug = True And lGroupNum = 3 Then Debug.Print "ran sSQL for cand ....             " & Format(Timer, "ss")
            
            
            If rs2Fit.EOF Then  'NO CANDIDATES TO JOIN SO IT IS A SINGLETON
               Debug.Print lGroupNum & " . . . " & 1
               If bDebug = True Then Debug.Print "   " & " . " & "currTableID " & rs!TableXID & " assigned " & rs!Cubes & " Cubes. [ANCHOR] -- SINGLETON."
            Else
               rs2Fit.MoveFirst 'MOVE TO FIRST RECORD OF CANDIDATES
               Debug.Print lGroupNum & " . . . "
               If bDebug = True Then Debug.Print "   " & " . " & "currTableID " & rs!TableXID & " assigned " & rs!Cubes & " Cubes. [ANCHOR] "
            End If
            
         End If
          
         '//============ Inner Loop on candidate deliveries to group with the first (anchor) delivery
         '//===============================================================================================================
         
         If bDebug = True And lGroupNum = 3 Then Debug.Print "start inner loop  ....             " & Format(Timer, "ss")
         
         
         Do While Not rs2Fit.EOF
            DoEvents
            If rs2Fit!PkgCubes <= cubesleft And rs2Fit!PkgWeight <= weightleft And Left(rs2Fit!GroupID, 1) = 0 And iGrpStops <> 0 Then
               DoCmd.SetWarnings False
               If InStr(rs2Fit!TableXID, "...") > 0 Then 'DOUBLETON NEEDS TO BE ADDED
                  'Assign THEM and update BOTH RECORDS FOR the GroupID and GroupCode
                  iGrpStops = iGrpStops + 2
                  'UPDATE THE FIRST ONE
                  sSQL = "UPDATE " & sTableName & " SET " & sTableName & ".GroupID = " & lGroupNum & ", " & _
                                     sTableName & ".GroupCode = '" & lGroupNum & "." & iGrpStops & "d' " & _
                                "WHERE " & sTableName & ".ID=" & Trim(Left(rs2Fit!TableXID, InStr(rs2Fit!TableXID, "...") - 1)) & " "
                  DoCmd.RunSQL (sSQL)
                  If bDebug = True And lGroupNum = 3 Then Debug.Print "update first of dblton...          " & Format(Timer, "ss")
   
                  
                  'UPDATE THE SECOND ONE
                  sSQL = "UPDATE " & sTableName & " SET " & sTableName & ".GroupID = " & lGroupNum & ", " & _
                                     sTableName & ".GroupCode = '" & lGroupNum & "." & iGrpStops + 1 & "dd' " & _
                                "WHERE " & sTableName & ".ID=" & Trim(Right(rs2Fit!TableXID, Len(rs2Fit!TableXID) - (InStr(rs2Fit!TableXID, "...") + 2))) & " "
                  DoCmd.RunSQL (sSQL)
                  If bDebug = True And lGroupNum = 3 Then Debug.Print "update secnd of dblton...          " & Format(Timer, "ss")
               
               Else 'SINGLETON NEEDS TO BE ADDED
                  'Assign it and update the table GroupID and GroupCode
                  iGrpStops = iGrpStops + 1
                  DoCmd.RunSQL ("UPDATE " & sTableName & " SET " & sTableName & ".GroupID = " & lGroupNum & ", " & _
                                                                   sTableName & ".GroupCode = '" & lGroupNum & "." & iGrpStops & "s' " & _
                                "WHERE " & sTableName & ".ID=" & rs2Fit!TableXID & "")
                  If bDebug = True And lGroupNum = 3 Then Debug.Print "update singleton ........          " & Format(Now() - ET, "ss:") & " " & "." & Right(Format(Timer, "#0.000"), 3)
                                
               End If
               
               DoCmd.SetWarnings True
               cubesleft = cubesleft - rs2Fit!PkgCubes
               weightleft = weightleft - rs2Fit!PkgWeight
               If bDebug = True Then Debug.Print "   " & " . " & "currTableID " & rs2Fit!TableXID & " assigned " & rs2Fit!PkgCubes & " Cubes. "
            End If
            
            '//============ STOPPING Criteria:  If it is enough, then stop
            '//===============================================================================================================
            If cubesleft / dMaxCubes <= dPctTolerance Or weightleft / dMaxWeight <= dPctTolerance Or iGrpStops = iMaxStops Then
               'reset rs to unallocated records
               sSQL = "SELECT * FROM " & "TableY" & " " & vbCrLf & _
                      "WHERE " & "" & "Cubes <= " & dMaxCubes & " AND " & "Cubes <= " & iMaxCubesToGroup & " AND " & "Weight <= " & dMaxWeight & " " & vbCrLf & _
                      "  AND " & "GroupID = " & 0 & " " & vbCrLf & _
                      "ORDER BY " & sSortFieldName & " ASC "
               Set rs = db.OpenRecordset(sSQL)
               rs.MoveFirst
               If bDebug = True Then Debug.Print "       . . . .  . . . . " & Format(Now(), "hh:nn:ss") & "   " & rs.RecordCount - 1 & " records left working . . . . . . . . GOT FILLED UP."
               lGroupNum = lGroupNum + 1           'increment counter to next Group
               cubesleft = dMaxCubes
               weightleft = dMaxWeight
               iGrpStops = 0
               If bDebug = True And lGroupNum = 3 Then Debug.Print "done inner loop here.....          " & Format(Now() - ET, "ss:") & " " & "." & Right(Format(Timer, "#0.000"), 3)
               If bDebug = True And lGroupNum = 3 Then Debug.Print "done inner loop here.....          " & Format(Now() - T, "hh:nn:ss")
               Exit Do  'BREAK OUT OF THE WHILE LOOP
            Else
               
            End If
            
            rs2Fit.MoveNext
    
         
         Loop
         
         If cubesleft = dMaxCubes Then
            'you got filled up and advanced
         Else
            'you did not reach a stopping criteria, but you ran out of group candidates
            'reset rs
            sSQL = "SELECT * FROM " & " TableY " & " " & vbCrLf & _
                   "WHERE " & "" & " TableY" & ".Cubes <= " & dMaxCubes & " AND " & " TableY" & ".Cubes <= " & iMaxCubesToGroup & " AND " & " TableY" & ".Weight <= " & dMaxWeight & " " & vbCrLf & _
                   "  AND " & " TableY" & ".GroupID = " & 0 & " " & vbCrLf & _
                   "ORDER BY " & sSortFieldName & " ASC "
            Set rs = db.OpenRecordset(sSQL)
            If rs.EOF Then 'done
            Else
               rs.MoveFirst
               '''Debug.Print "       . . . .  . . . . " & Format(Now(), "hh:nn:ss") & "   " & rs.RecordCount - 1 & " records left working . . . . . . . . RAN OUT OF CANDIDATES."
               lGroupNum = lGroupNum + 1           'increment counter to next Group
               cubesleft = dMaxCubes
               weightleft = dMaxWeight
               iGrpStops = 0
            End If
         End If
      Wend 'RS 'outer loop'
      sSQL = "INSERT INTO TableZ ( TableXID, GroupID, GroupCode ) " & vbCrLf & _
             "SELECT TableY.TableXID, TableY.GroupID, TableY.GroupCode " & vbCrLf & _
             "FROM TableY "
      rsChunk.MoveNext
   Wend 'rsChunk
    
    
   'number the ungrouped ones
   sSQL = "SELECT * FROM " & sTableName & " " & vbCrLf & _
          "WHERE " & sTableName & ".GroupID = " & 0 & " " & vbCrLf & _
          "ORDER BY " & sSortFieldName & " ASC "
   Set rs = db.OpenRecordset(sSQL)
   While Not rs.EOF
      lGroupNum = lGroupNum + 1        'move this up to the point where the singleton is id'd
      rs.Edit       'move this up to the point where the singleton is id'd
      rs!GroupID = lGroupNum       'move this up to the point where the singleton is id'd
      rs!GroupCode = lGroupNum & ".0"  'move this up to the point where the singleton is id'd
      rs.Update
      Debug.Print lGroupNum & " . . . "
   rs.MoveNext
   Wend
   
    
      
   T = Now() - T
   s = "ChunkAndRun Grouper....." & Format(T, "hh:nn:ss")
   Debug.Print s
   Debug.Print "--DONE--"
   
   ChunkAndRun = True
   
ChunkAndRun_Done:
  Exit Function

ChunkAndRun_Err:
   Select Case Err
    Case 9 'or 13 'subscript out of range
      ChunkAndRun = False
    Case 3021 'no current record
      ChunkAndRun = True
      Resume ChunkAndRun_Done
    Case Else
      MsgBox "Error: " & Error$ & " " & Err, , " ChunkAndRun Grouper"
      ChunkAndRun = False
   End Select
   Resume ChunkAndRun_Done
   Resume
   
End Function


Public Function GenSQLToGetGroupCandidates(rs As DAO.Recordset, sTableName As String, cubesleft As Double, _
                          dMaxCubes As Double, dMaxWeight As Double, _
                          dMaxDeltaDays As Double, dMaxDeltaZip2 As Double, _
                          sMethod As String _
                          ) As String
'//=====================================================================================================================//
'/|   FUNCTION:  GenSQLToGetGroupCandidates      RETURNS:  String of the SQL                                            |/
'/|    PURPOSE:  Generate the SQL SELECT which will get candidates of singletons and doubeltons to                      |/
'/|                  group with the current anchor                                                                      |/
'/|      USAGE:  sSQL = GenSQLToGetGroupCandidates(rs, CStr(sTableName), CDbl(cubesleft), _                             |/
'/|                                                CDbl(dMaxCubes), CDbl(dMaxWeight), _                                 |/
'/|                                                CDbl(dMaxDeltaDays), CDbl(dMaxDeltaZip2), _                          |/
'/|                                                "BOTH")   'SINGLE, DOUBLE, OR BOTH                                   |/
'/|                                                                                                                     |/
'/|         BY:  Sean                                                                                                   |/
'/|       DATE:  08/21/2015                                                                                             |/
'/|    HISTORY:  08/21/2015                                                                                             |/
'/|                                                                                                                     |/
'//=====================================================================================================================//
                          
                          
'''          "   " & sTableName & ".*, " & vbCrLf & _

                          
   'sMethod is SINGLE, DOUBLE OR BOTH
   Dim sSQL As String
   GenSQLToGetGroupCandidates = ""
   sSQL = "SELECT * FROM " & vbCrLf & _
          "( " & vbCrLf & _
          "SELECT " & vbCrLf & _
          "   " & sTableName & ".TableYID  AS TableYID, " & vbCrLf & _
          "   " & sTableName & ".TableXID  AS TableXID,  " & vbCrLf & _
          "   " & sTableName & ".Delivery  AS Delivery,  " & vbCrLf & _
          "   " & sTableName & ".ShipSite  AS ShipSite,  " & vbCrLf & _
          "   " & sTableName & ".ShipDate  AS ShipDate,  " & vbCrLf & _
          "   " & sTableName & ".Year    AS [Year],   " & vbCrLf & _
          "   " & sTableName & ".Month   AS [Month],  " & vbCrLf & _
          "   " & sTableName & ".Mode    AS Mode,     " & vbCrLf & _
          "   " & sTableName & ".Weight  AS Weight,   " & vbCrLf & _
          "   " & sTableName & ".Cubes   AS Cubes,    " & vbCrLf & _
          "   " & sTableName & ".Dest    AS Dest,     " & vbCrLf & _
          "   " & sTableName & ".Zip2    AS Zip2,     "
   sSQL = sSQL & "" & vbCrLf & _
          "   " & sTableName & ".Zip2Numeric AS Zip2Numeric,   " & vbCrLf & _
          "   " & sTableName & ".Zip2State   AS Zip2State,    " & vbCrLf & _
          "   " & sTableName & ".Zip3        AS Zip3,         " & vbCrLf & _
          "   " & sTableName & ".DestZipCode AS DestZipCode,   " & vbCrLf & _
          "   " & sTableName & ".REGN        AS REGN,         " & vbCrLf & _
          "   " & sTableName & ".GroupSortPriority AS GroupSortPriority,   " & vbCrLf & _
          "   " & sTableName & ".GroupID     AS GroupID,   " & vbCrLf & _
          "   " & sTableName & ".GroupCode   AS GroupCode, "
   sSQL = sSQL & "" & vbCrLf & _
          "   " & sTableName & ".Cubes AS PkgCubes, " & vbCrLf & _
          "   " & sTableName & ".Weight AS PkgWeight, " & vbCrLf & _
          "   Round(" & dMaxCubes & " - [" & sTableName & "_1].[Cubes] - [" & sTableName & "].[Cubes],2) AS CubeLim, " & vbCrLf & _
          "   Abs([" & sTableName & "].[ShipDate]-[" & sTableName & "_1].[ShipDate]) AS DayDelta, " & vbCrLf & _
          "   Abs(CInt([" & sTableName & "].[Zip2Numeric])-CInt([" & sTableName & "_1].[Zip2Numeric])) AS Z2Delta, " & vbCrLf & _
          "   " & sTableName & "_1.GroupID AS AnchorID, " & vbCrLf & _
          "   Round( [CubeLim] + IIF([DayDelta]=0,-100,60*[DayDelta]) + IIF([Z2Delta]=0,-250,150*[Z2Delta])  ,4) AS DistFromAnchorSCORE "
   sSQL = sSQL & "" & vbCrLf & _
          "FROM " & sTableName & " INNER JOIN " & sTableName & " AS " & sTableName & "_1 ON " & sTableName & ".ShipSite = " & sTableName & "_1.ShipSite" & vbCrLf & _
          "WHERE " & vbCrLf & _
          "   " & cubesleft & "  >= [" & sTableName & "].[Cubes] AND " & vbCrLf & _
          "   Abs([" & sTableName & "].[ShipDate]-[" & sTableName & "_1].[ShipDate]) <= " & dMaxDeltaDays & " AND " & vbCrLf & _
          "   Abs(CInt([" & sTableName & "].[Zip2Numeric])-CInt([" & sTableName & "_1].[Zip2Numeric])) <= " & dMaxDeltaZip2 & " AND " & vbCrLf & _
          "   " & sTableName & "_1.GroupID = " & rs!GroupID & " AND " & vbCrLf & _
          "   " & sTableName & ".GroupID = 0" & " " & vbCrLf & _
          "ORDER BY " & vbCrLf & _
          "   " & sTableName & ".Cubes DESC " & vbCrLf & _
          " " & vbCrLf
   
   ''============================
   '' THIS CODE BELOW IS ALL ON NOW.  HERE IS THE UNION WITH THE DOUBLETONS
   ''============================
   sSQL = sSQL & "" & vbCrLf & _
          "UNION ALL" & vbCrLf & _
                 "" & vbCrLf
   
   sSQL = sSQL & "SELECT  " & vbCrLf & _
          "   " & sTableName & ".TableYID" & " & '...' & " & " " & sTableName & "Pair.TableYID AS TableYID,  " & vbCrLf & _
          "   " & sTableName & ".TableXID" & " & '...' & " & " " & sTableName & "Pair.TableXID AS TableXID,  " & vbCrLf & _
          "   " & sTableName & ".Delivery" & " & '...' & " & " " & sTableName & "Pair.Delivery AS Delivery,  " & vbCrLf & _
          "   " & sTableName & ".ShipSite" & " & '...' & " & " " & sTableName & "Pair.ShipSite AS ShipSite,  " & vbCrLf & _
          "   " & sTableName & ".ShipDate" & " & '...' & " & " " & sTableName & "Pair.ShipDate AS ShipDate,  " & vbCrLf & _
          "   " & sTableName & ".Year" & " & '...' & " & " " & sTableName & "Pair.Year AS [Year],  " & vbCrLf & _
          "   " & sTableName & ".Month" & " & '...' & " & " " & sTableName & "Pair.Month AS [Month],  " & vbCrLf & _
          "   " & sTableName & ".Mode" & " & '...' & " & " " & sTableName & "Pair.Mode AS Mode,  " & vbCrLf & _
          "   " & sTableName & ".Weight" & " & '...' & " & " " & sTableName & "Pair.Weight AS Weight,  " & vbCrLf & _
          "   " & sTableName & ".Cubes" & " & '...' & " & " " & sTableName & "Pair.Cubes AS Cubes,  " & vbCrLf & _
          "   " & sTableName & ".Dest" & " & '...' & " & " " & sTableName & "Pair.Dest AS Dest,  " & vbCrLf & _
          "   " & sTableName & ".Zip2" & " & '...' & " & " " & sTableName & "Pair.Zip2 AS Zip2,  " & vbCrLf & _
          "   " & sTableName & ".Zip2Numeric" & " & '...' & " & " " & sTableName & "Pair.Zip2Numeric AS Zip2Numeric,  " & vbCrLf & _
          "   " & sTableName & ".Zip2State" & " & '...' & " & " " & sTableName & "Pair.Zip2State AS Zip2State,  " & vbCrLf & _
          "   " & sTableName & ".Zip3" & " & '...' & " & " " & sTableName & "Pair.Zip3 AS Zip3,  " & vbCrLf & _
          "   " & sTableName & ".DestZipCode" & " & '...' & " & " " & sTableName & "Pair.DestZipCode AS DestZipCode,  " & vbCrLf & _
          "   " & sTableName & ".REGN" & " & '...' & " & " " & sTableName & "Pair.REGN AS REGN,  " & vbCrLf & _
          "   " & sTableName & ".GroupSortPriority" & " & '...' & " & " " & sTableName & "Pair.GroupSortPriority AS GroupSortPriority,  " & vbCrLf & _
          "   " & sTableName & ".GroupID" & " & '...' & " & " " & sTableName & "Pair.GroupID AS GroupID,  " & vbCrLf & _
          "   " & sTableName & ".GroupCode" & " & '...' & " & " " & sTableName & "Pair.GroupCode AS GroupCode,  " & vbCrLf

   sSQL = sSQL & _
          "   [" & sTableName & "].[Cubes]+[" & sTableName & "Pair].[Cubes] AS PkgCubes,  " & vbCrLf & _
          "   [" & sTableName & "].[Weight]+[" & sTableName & "Pair].[Weight] AS PkgWeight,  " & vbCrLf & _
          "   Round(" & cubesleft & " - ([" & sTableName & "].[Cubes]+[" & sTableName & "Pair].[Cubes]), 2) AS CubeLim,  " & vbCrLf & _
          "   Abs(([" & sTableName & "].[ShipDate]+[" & sTableName & "Pair].[ShipDate])/2-[" & sTableName & "_1].[ShipDate]) AS DayDelta,  " & vbCrLf & _
          "   Abs((CInt([" & sTableName & "].[Zip2])+CInt([" & sTableName & "Pair].[Zip2]))/2-CInt([" & sTableName & "_1].[Zip2])) AS Z2Delta,  " & vbCrLf & _
          "   " & sTableName & "_1.GroupID AS AnchorID,  " & vbCrLf & _
          "     Round( [CubeLim] + IIF([DayDelta]=0,-100,60*[DayDelta]) + IIF([Z2Delta]=0,-250,150*[Z2Delta])  ,4) AS DistFromAnchorSCORE " & vbCrLf & _
          "FROM (" & sTableName & " INNER JOIN " & sTableName & " AS " & sTableName & "_1 ON  " & vbCrLf & _
          "   " & sTableName & ".ShipSite = " & sTableName & "_1.ShipSite) INNER JOIN " & sTableName & " AS " & sTableName & "Pair ON  " & vbCrLf & _
          "   " & sTableName & ".ShipSite = " & sTableName & "Pair.ShipSite " & vbCrLf

   sSQL = sSQL & _
          "WHERE  " & vbCrLf & _
          "   (([" & sTableName & "].[ShipDate])  <= " & dMaxDeltaDays & " + ([" & sTableName & "_1].[ShipDate])) AND " & vbCrLf & _
          "   (([" & sTableName & "].[ShipDate])  >= " & -dMaxDeltaDays & " + ([" & sTableName & "_1].[ShipDate])) AND " & vbCrLf & _
          "   (([" & sTableName & "Pair].[ShipDate])  <= " & dMaxDeltaDays & " + ([" & sTableName & "_1].[ShipDate])) AND " & vbCrLf & _
          "   (([" & sTableName & "Pair].[ShipDate])  >= " & -dMaxDeltaDays & " + ([" & sTableName & "_1].[ShipDate])) AND " & vbCrLf & _
          "   (CInt([" & sTableName & "].[Zip2Numeric])      <= " & dMaxDeltaZip2 & " + CInt([" & sTableName & "_1].[Zip2Numeric])) AND " & vbCrLf & _
          "   (CInt([" & sTableName & "].[Zip2Numeric])      >= " & -dMaxDeltaZip2 & " + CInt([" & sTableName & "_1].[Zip2Numeric])) AND " & vbCrLf & _
          "   (CInt([" & sTableName & "Pair].[Zip2Numeric])  <= " & dMaxDeltaZip2 & " + CInt([" & sTableName & "_1].[Zip2Numeric])) AND " & vbCrLf & _
          "   (CInt([" & sTableName & "Pair].[Zip2Numeric])  >= " & -dMaxDeltaZip2 & " + CInt([" & sTableName & "_1].[Zip2Numeric])) AND " & vbCrLf & _
          "   ((" & sTableName & "_1.GroupID)=" & rs!GroupID & ") AND ((" & dMaxCubes & " - [" & sTableName & "_1].[Cubes]) >= [" & sTableName & "].[Cubes]+[" & sTableName & "Pair].[Cubes]) AND  " & vbCrLf & _
          "   ((" & sTableName & ".ShipDate - " & sTableName & "Pair.ShipDate) <= 2*" & dMaxDeltaDays & ") AND  " & vbCrLf & _
          "   [" & sTableName & "].[Cubes]+[" & sTableName & "Pair].[Cubes] <= " & cubesleft & " AND  " & vbCrLf & _
          "   ((" & sTableName & ".GroupID)=0) AND ((" & sTableName & "Pair.GroupID)=0) " & " AND  " & vbCrLf & _
          "   ((" & sTableName & ".TableYID) < [" & sTableName & "Pair].[TableYID]) " & vbCrLf
          
   sSQL = sSQL & _
          ") AS aliOneTwoDel " & vbCrLf & _
          "ORDER BY DistFromAnchorSCORE, CubeLim " & vbCrLf
          
   GenSQLToGetGroupCandidates = sSQL
   
   'When you run this SQL, here is a transposed version of what the SQL will yield
   
   'In this case we looked for options to fillt he remaining 1100.42 cubes left.
   'This qry result set is turned on its side.  the 1-2-3-4-5 would really be rows in the recordset.

   '                     1                                              2                       3                       4                       5
   '
   'ID                   305...306                                      305                     306
   'Delivery             KEN68983 CHUDY 14225...KEN68963 RESTA 14225    KEN68983 CHUDY 14225    KEN68963 RESTA 14225
   'ShipSite             KEN...KEN                                      KEN                     KEN
   'ShipDate             8/28/2014...8/28/2014                          8/28/2014               8/28/2014
   'Year                 2014...2014                                    2014                    2014
   'Month                8...8                                          8                       8
   'WK                   35...35                                        35                      35
   'Mode                 L...L                                          L                       L
   'Weight               4427.5...3352.2                                4427.5                  3352.2
   'Cubes                618.24...412.76                                618.24                  412.76
   'Dest                 CHUDY 14225...RESTA 14225                      CHUDY 14225             RESTA 14225
   'Zip2                 14...14                                        14                      14
   'Zip2State            14NY...14NY                                    14NY                    14NY
   'Zip3                 142...142                                      142                     142
   'DestZipCode          14225...14225                                  14225                   14225
   'GroupSortPriority    01031...01237                                  01031                   01237
   'GroupID              0...0                                          0                       0
   'GroupCode            0...0                                          0                       0
   'PkgCubes             1031                                           618.24                  412.76
   'PkgWeight            7779.7                                         4427.5                  3352.2
   'CubeLim              69.42                                          2086.3                  2291.78
   'DayDelta             3                                              3                       3
   'Z2Delta              0                                              0                       0
   'AnchorID             1                                              1                       1
   'DistFromAnchorSCORE  -0.58                                          2016.3                  2221.78
   
End Function



'//=====================================================================================================================//
'//=====================================================================================================================//
'//=====================================================================================================================//
'//=====================================================================================================================//
'//=====================================================================================================================//
'//=====================================================================================================================//



''''''Public Function Grouper(Optional sTableName = "TableX", Optional sSortFieldName = "GroupSortPriority", _
''''''                        Optional dMaxCubes = 2700, Optional dMaxWeight = 24000, Optional iMaxStops = 3, _
''''''                        Optional iMaxCubesToGroup = 2399, Optional dPctTolerance = 0.01, _
''''''                        Optional dMaxDeltaDays = 5, Optional dMaxDeltaZip2 = 5, _
''''''                        Optional dTgtFirstDeliverySizeInCubes = 2000 _
''''''                       ) As Boolean
'''''''//=====================================================================================================================//
'''''''/|   FUNCTION:  Grouper                 RETURNS:  True on Success                                                      |/
'''''''/| PARAMETERS:  Optional:  sTableName     sSortFieldName dMaxCubes     dMaxWeight iMaxStops iMaxCubesToGroup           |/
'''''''/|                         dPctTolerance  dMaxDeltaDays  dMaxDeltaZip2                                                 |/
'''''''/|              Table fields:  ID, ShipSite, GroupID, GroupCode, Zip2Numeric, Cubes, Weight, ShipDate, Year, Month, WK |/
'''''''/|    PURPOSE:  Group Shipments into truckLoads by weight limit or cubic foot limit                                    |/
'''''''/|      USAGE:  i= Grouper()                                                                                           |/
'''''''/|      USAGE:  i= Grouper("TableX","GroupSortPriority", 2750, 50000, 3, 2500, 0.10, 3, 2, 1650)                       |/
'''''''/|         BY:  Sean                                                                                                   |/
'''''''/|       DATE:  08/21/2015                                                                                             |/
'''''''/|    HISTORY:  08/21/2015                                                                                             |/
'''''''/|                                                                                                                     |/
'''''''//=====================================================================================================================//
''''''
''''''
''''''   '//============  Declarations and Things
''''''   '//==================================================================================================================
''''''   Dim db          As Database
''''''   Dim rs          As DAO.Recordset
''''''   Dim rs2Fit      As DAO.Recordset
''''''   Dim rsChunk     As DAO.Recordset
''''''   Dim sSQL        As String
''''''   Dim i           As Integer
''''''   Dim T           As Date
''''''   Dim ET          As Date
''''''   Dim weightleft  As Double
''''''   Dim cubesleft   As Double
''''''   Dim sShipSite   As String
''''''   Dim bDebug      As Boolean
''''''   Dim lGroupNum   As Long
''''''   Dim iGrpStops   As Integer
''''''   Dim s           As String
''''''
''''''   On Error GoTo Grouper_Err
''''''   Grouper = False
''''''   T = Now()
''''''   ET = Timer
''''''   Set db = CurrentDb
''''''   sSQL = ""
''''''   iGrpStops = 0
''''''   bDebug = True
''''''
''''''   Debug.Print "START............ ....             " & Format(Timer, "ss")
''''''
''''''   '//============  Initialize the Groups, setting all group ID's and group Codes to "0"
''''''   '//=====================================================================================================================
''''''   DoCmd.SetWarnings False
''''''   DoCmd.RunSQL ("UPDATE " & sTableName & " SET " & sTableName & ".GroupID = 0, " & sTableName & ".GroupCode = '0'")
''''''   DoCmd.SetWarnings True
''''''   Debug.Print "INIT GROUP_IDS AND STUFF...        " & Format(Timer, "ss")
''''''
''''''
''''''   '//============  Set rsChunk to the list of combo's of ShipSite, Yr and Month so that the 'outer' loop is over fewer recs
''''''   '//=====================================================================================================================
''''''   sSQL = "SELECT ShipSite ,Year, Month, Count(ID) AS CtID " & vbCrLf & _
''''''          "FROM " & sTableName & " " & vbCrLf & _
''''''          "GROUP BY ShipSite, Year, Month " & vbCrLf
''''''
''''''   Set rsChunk = db.OpenRecordset(sSQL)
''''''   rsChunk.MoveFirst
''''''   Debug.Print "open rsChunk for the master index to the outer loop         " & Format(Timer, "ss")
''''''
''''''
''''''   '//============  rsChunk Loop -- Grab a Chunk of records to group at once (saves candidate qry work and time)
''''''   '//=====================================================================================================================
''''''   While Not rsChunk.EOF
''''''      DoEvents
''''''      '##
''''''      If bDebug = True Then Debug.Print "Top of CHUNK Loop ....             " & Format(Timer, "ss")
''''''
''''''
''''''
''''''
''''''
''''''      '//============  Set rs to the table full of deliveries that is used for the 'outer' loop.
''''''      '//=====================================================================================================================
''''''      sSQL = "SELECT *, " & vbCrLf & _
''''''             "   IIf(IsAlpha(Left([Zip2],1)),900+Hex(Asc(Left([Zip2],1))),CInt([Zip2])) AS Zip2Numeric " & vbCrLf & _
''''''             "FROM " & sTableName & " " & vbCrLf & _
''''''             "WHERE " & "" & sTableName & ".Cubes <= " & dMaxCubes & " AND " & sTableName & ".Cubes <= " & iMaxCubesToGroup & _
''''''             " AND " & sTableName & ".Weight <= " & dMaxWeight & " " & vbCrLf & _
''''''             "ORDER BY " & sSortFieldName & " ASC "
''''''      Set rs = db.OpenRecordset(sSQL)
''''''      rs.MoveFirst
''''''      Debug.Print "open rs for the outer loop         " & Format(Timer, "ss")
''''''
''''''      '//============  Initialize Variables
''''''      '//=====================================================================================================================
''''''      'key = rs!ShipSite
''''''      'keyTime = rs!Year & rs!Month & rs!WK
''''''      cubesleft = dMaxCubes
''''''      weightleft = dMaxWeight
''''''      lGroupNum = 1
''''''      Debug.Print "initialize variables..             " & Format(Timer, "ss")
''''''
''''''
''''''      '//============  Outer Loop on Table of Deliveries (e.g. 38,000 for FS year, or 7,400 for FS to NorthEast U.S.)
''''''      '//=====================================================================================================================
''''''      While Not rs.EOF
''''''         DoEvents
''''''         '##
''''''         If bDebug = True And lGroupNum = 17 Then Debug.Print "Top of Outer Loop ....             " & Format(Timer, "ss")
''''''
''''''         If rs!Cubes <= cubesleft And rs!Weight <= weightleft And _
''''''            rs!GroupID = 0 And iGrpStops = 0 Then     '//Group this line in and fill in the Group ID and Code for Stop 1 (Anchor)
''''''
''''''            'Assign it and update the table GroupID and GroupCode
''''''            iGrpStops = iGrpStops + 1
''''''           'key = rs!ShipSite
''''''            rs.Edit
''''''            rs!GroupID = lGroupNum
''''''            rs!GroupCode = lGroupNum & "." & iGrpStops
''''''            rs.Update
''''''            cubesleft = cubesleft - rs!Cubes
''''''            weightleft = weightleft - rs!Weight
''''''
''''''            If bDebug = True And lGroupNum = 17 Then Debug.Print "Init candidates   ....             " & Format(Timer, "ss")
''''''
''''''            Set rs2Fit = Nothing
''''''
''''''            'Right here, you know what you've got as the first shipment, go build the SQL to return eligible candidates to
''''''            'join (there may be none).  If there are some return the best 5 singletons and doubletons.
''''''
''''''            sSQL = GenSQLToGetGroupCandidates(rs, CStr(sTableName), CDbl(cubesleft), CDbl(dMaxCubes), CDbl(dMaxWeight), _
''''''                                              CDbl(dMaxDeltaDays), CDbl(dMaxDeltaZip2), "BOTH")   'SINGLE, DOUBLE, OR BOTH
''''''
''''''            If bDebug = True And lGroupNum = 17 Then Debug.Print "got sSQL for cand ....             " & Format(Timer, "ss")
''''''
''''''            Set rs2Fit = db.OpenRecordset(sSQL)
''''''
''''''            If bDebug = True And lGroupNum = 17 Then Debug.Print "ran sSQL for cand ....             " & Format(Timer, "ss")
''''''
''''''
''''''            If rs2Fit.EOF Then  'NO CANDIDATES TO JOIN SO IT IS A SINGLETON
''''''               Debug.Print lGroupNum & " . . . " & 1
''''''               If bDebug = True Then Debug.Print "   " & " . " & "currTableID " & rs!id & " assigned " & rs!Cubes & " Cubes. [ANCHOR] -- SINGLETON."
''''''            Else
''''''               rs2Fit.MoveFirst 'MOVE TO FIRST RECORD OF CANDIDATES
''''''               Debug.Print lGroupNum & " . . . "
''''''               If bDebug = True Then Debug.Print "   " & " . " & "currTableID " & rs!id & " assigned " & rs!Cubes & " Cubes. [ANCHOR] "
''''''            End If
''''''
''''''         End If
''''''
''''''         '//============ Inner Loop on candidate deliveries to group with the first (anchor) delivery
''''''         '//===============================================================================================================
''''''
''''''         If bDebug = True And lGroupNum = 17 Then Debug.Print "start inner loop  ....             " & Format(Timer, "ss")
''''''
''''''
''''''         Do While Not rs2Fit.EOF
''''''            DoEvents
''''''            If rs2Fit!PkgCubes <= cubesleft And rs2Fit!PkgWeight <= weightleft And Left(rs2Fit!GroupID, 1) = 0 And iGrpStops <> 0 Then
''''''               DoCmd.SetWarnings False
''''''               If InStr(rs2Fit!id, "...") > 0 Then 'DOUBLETON NEEDS TO BE ADDED
''''''                  'Assign THEM and update BOTH RECORDS FOR the GroupID and GroupCode
''''''                  iGrpStops = iGrpStops + 2
''''''                  'UPDATE THE FIRST ONE
''''''                  sSQL = "UPDATE " & sTableName & " SET " & sTableName & ".GroupID = " & lGroupNum & ", " & _
''''''                                     sTableName & ".GroupCode = '" & lGroupNum & "." & iGrpStops & "d' " & _
''''''                                "WHERE " & sTableName & ".ID=" & Trim(Left(rs2Fit!id, InStr(rs2Fit!id, "...") - 1)) & " "
''''''                  DoCmd.RunSQL (sSQL)
''''''                  If bDebug = True And lGroupNum = 17 Then Debug.Print "update first of dblton...          " & Format(Timer, "ss")
''''''
''''''
''''''                  'UPDATE THE SECOND ONE
''''''                  sSQL = "UPDATE " & sTableName & " SET " & sTableName & ".GroupID = " & lGroupNum & ", " & _
''''''                                     sTableName & ".GroupCode = '" & lGroupNum & "." & iGrpStops + 1 & "dd' " & _
''''''                                "WHERE " & sTableName & ".ID=" & Trim(Right(rs2Fit!id, Len(rs2Fit!id) - (InStr(rs2Fit!id, "...") + 2))) & " "
''''''                  DoCmd.RunSQL (sSQL)
''''''                  If bDebug = True And lGroupNum = 17 Then Debug.Print "update secnd of dblton...          " & Format(Timer, "ss")
''''''
''''''               Else 'SINGLETON NEEDS TO BE ADDED
''''''                  'Assign it and update the table GroupID and GroupCode
''''''                  iGrpStops = iGrpStops + 1
''''''                  DoCmd.RunSQL ("UPDATE " & sTableName & " SET " & sTableName & ".GroupID = " & lGroupNum & ", " & _
''''''                                                                   sTableName & ".GroupCode = '" & lGroupNum & "." & iGrpStops & "s' " & _
''''''                                "WHERE " & sTableName & ".ID=" & rs2Fit!id & "")
''''''                  If bDebug = True And lGroupNum = 17 Then Debug.Print "update singleton ........          " & Format(Now() - ET, "ss:") & " " & "." & Right(Format(Timer, "#0.000"), 3)
''''''
''''''               End If
''''''
''''''               DoCmd.SetWarnings True
''''''               cubesleft = cubesleft - rs2Fit!PkgCubes
''''''               weightleft = weightleft - rs2Fit!PkgWeight
''''''               If bDebug = True Then Debug.Print "   " & " . " & "currTableID " & rs2Fit!id & " assigned " & rs2Fit!PkgCubes & " Cubes. "
''''''            End If
''''''
''''''            '//============ STOPPING Criteria:  If it is enough, then stop
''''''            '//===============================================================================================================
''''''            If cubesleft / dMaxCubes <= dPctTolerance Or weightleft / dMaxWeight <= dPctTolerance Or iGrpStops = iMaxStops Then
''''''               'reset rs to unallocated records
''''''               sSQL = "SELECT * FROM " & sTableName & " " & vbCrLf & _
''''''                      "WHERE " & "" & sTableName & ".Cubes <= " & dMaxCubes & " AND " & sTableName & ".Cubes <= " & iMaxCubesToGroup & " AND " & sTableName & ".Weight <= " & dMaxWeight & " " & vbCrLf & _
''''''                      "  AND " & sTableName & ".GroupID = " & 0 & " " & vbCrLf & _
''''''                      "ORDER BY " & sSortFieldName & " ASC "
''''''               Set rs = db.OpenRecordset(sSQL)
''''''               rs.MoveFirst
''''''               If bDebug = True Then Debug.Print "       . . . .  . . . . " & Format(Now(), "hh:nn:ss") & "   " & rs.RecordCount - 1 & " records left working . . . . . . . . GOT FILLED UP."
''''''               lGroupNum = lGroupNum + 1           'increment counter to next Group
''''''               cubesleft = dMaxCubes
''''''               weightleft = dMaxWeight
''''''               iGrpStops = 0
''''''               If bDebug = True And lGroupNum = 17 Then Debug.Print "done inner loop here.....          " & Format(Now() - ET, "ss:") & " " & "." & Right(Format(Timer, "#0.000"), 3)
''''''               If bDebug = True And lGroupNum = 17 Then Debug.Print "done inner loop here.....          " & Format(Now() - T, "hh:nn:ss")
''''''
''''''               Exit Do  'BREAK OUT OF THE WHILE LOOP
''''''            Else
''''''
''''''            End If
''''''
''''''            rs2Fit.MoveNext
''''''
''''''
''''''         Loop
''''''
''''''         If cubesleft = dMaxCubes Then
''''''            'you got filled up and advanced
''''''         Else
''''''            'you did not reach a stopping criteria, but you ran out of group candidates
''''''            'reset rs
''''''            sSQL = "SELECT * FROM " & sTableName & " " & vbCrLf & _
''''''                   "WHERE " & "" & sTableName & ".Cubes <= " & dMaxCubes & " AND " & sTableName & ".Cubes <= " & iMaxCubesToGroup & " AND " & sTableName & ".Weight <= " & dMaxWeight & " " & vbCrLf & _
''''''                   "  AND " & sTableName & ".GroupID = " & 0 & " " & vbCrLf & _
''''''                   "ORDER BY " & sSortFieldName & " ASC "
''''''            Set rs = db.OpenRecordset(sSQL)
''''''            If rs.EOF Then 'done
''''''            Else
''''''               rs.MoveFirst
''''''               '''Debug.Print "       . . . .  . . . . " & Format(Now(), "hh:nn:ss") & "   " & rs.RecordCount - 1 & " records left working . . . . . . . . RAN OUT OF CANDIDATES."
''''''               lGroupNum = lGroupNum + 1           'increment counter to next Group
''''''               cubesleft = dMaxCubes
''''''               weightleft = dMaxWeight
''''''               iGrpStops = 0
''''''            End If
''''''         End If
''''''      Wend 'RS 'outer loop'
''''''
''''''   Wend 'rsChunk
''''''
''''''
''''''   'number the ungrouped ones
''''''   sSQL = "SELECT * FROM " & sTableName & " " & vbCrLf & _
''''''          "WHERE " & sTableName & ".GroupID = " & 0 & " " & vbCrLf & _
''''''          "ORDER BY " & sSortFieldName & " ASC "
''''''   Set rs = db.OpenRecordset(sSQL)
''''''   While Not rs.EOF
''''''      lGroupNum = lGroupNum + 1        'move this up to the point where the singleton is id'd
''''''      rs.Edit       'move this up to the point where the singleton is id'd
''''''      rs!GroupID = lGroupNum       'move this up to the point where the singleton is id'd
''''''      rs!GroupCode = lGroupNum & ".0"  'move this up to the point where the singleton is id'd
''''''      rs.Update
''''''      Debug.Print lGroupNum & " . . . "
''''''   rs.MoveNext
''''''   Wend
''''''
''''''
''''''
''''''   T = Now() - T
''''''   s = "Grouper....." & Format(T, "hh:nn:ss")
''''''   Debug.Print s
''''''   Debug.Print "--DONE--"
''''''
''''''   Grouper = True
''''''
''''''Grouper_Done:
''''''  Exit Function
''''''
''''''Grouper_Err:
''''''   Select Case Err
''''''    Case 9 'or 13 'subscript out of range
''''''      Grouper = False
''''''    Case 3021 'no current record
''''''      Grouper = True
''''''      Resume Grouper_Done
''''''    Case Else
''''''      MsgBox "Error: " & Error$ & " " & Err, , " Grouper"
''''''      Grouper = False
''''''   End Select
''''''   Resume Grouper_Done
''''''   Resume
''''''
''''''End Function
''''''
''''''
''''''Public Function GroupShip(Optional sTableName = "tblFS1415_LoadSummaryAble", Optional sSortFieldName = "GroupSortPriority", _
''''''                          Optional dMaxCubes = 2700, Optional dMaxWeight = 24000, Optional iMaxStops = 3, _
''''''                          Optional iMaxCubesToGroup = 2399, Optional dPctTolerance = 0.01, _
''''''                          Optional dMaxDeltaDays = 5, Optional dMaxDeltaZip2 = 5, _
''''''                          Optional dTgtFirstDeliverySizeInCubes = 2000 _
''''''                          ) As Boolean
'''''''//=======================================================================================================================//
'''''''/|   FUNCTION:  GroupShip                 RETURNS:  True on Success                                                      |/
'''''''/| PARAMETERS:  Optional:  sTableName     sSortFieldName dMaxCubes     dMaxWeight iMaxStops iMaxCubesToGroup             |/
'''''''/|                         dPctTolerance  dMaxDeltaDays  dMaxDeltaZip2                                                   |/
'''''''/|              Table fields:  ID, ShipSite, GroupID, GroupCode, Cubes, Weight, ShipDate, Year, Month, WK                |/
'''''''/|    PURPOSE:  Group Shipments into truckLoads by weight limit or cubic foot limit                                      |/
'''''''/|      USAGE:  i= GroupShip()                                                                                           |/
'''''''/|      USAGE:  i= GroupShip("TableX","GroupSortPriority", 2700, 30000, 7, 4499, 0.01,  23, 3, 2000)                     |/
'''''''/|      USAGE:  i= GroupShip("TableX","GroupSortPriority", 3000, 50000, 7, 2700, 0.001,  4, 3, 1650)                     |/
'''''''/|      USAGE:  i= GroupShip("TableX","GroupSortPriority", 2750, 50000, 7, 2500, 0.001,  3, 2, 1650)                     |/
'''''''/|         BY:  Sean                                                                                                     |/
'''''''/|       DATE:  08/21/2015                                                                                               |/
'''''''/|    HISTORY:  08/21/2015                                                                                               |/
'''''''/|                                                                                                                       |/
'''''''//=======================================================================================================================//
''''''
'''''''//============  ** Enhance This:   [  ]  Add Delta Miles?
''''''
''''''   Dim db         As Database
''''''   Dim rs         As DAO.Recordset
''''''   Dim rs2Fit     As DAO.Recordset
''''''   Dim sSQL       As String
''''''   Dim i          As Integer
''''''   Dim T          As Date
''''''   Dim s          As String
''''''   Dim key        As String
''''''   Dim keyTime    As String
''''''   Dim weightleft As Double
''''''   Dim cubesleft  As Double
''''''   Dim sShipSite  As String
''''''   Dim lGroupNum  As Long
''''''   Dim iGrpStops  As Integer
''''''
''''''   On Error GoTo GroupShip_Err
''''''   GroupShip = False
''''''   T = Now()
''''''   Set db = CurrentDb
''''''   sSQL = ""
''''''   iGrpStops = 0
''''''
''''''   sSQL = "SELECT * FROM " & sTableName & " " & vbCrLf & _
''''''          "WHERE " & "" & sTableName & ".Cubes <= " & dMaxCubes & " AND " & sTableName & ".Cubes <= " & iMaxCubesToGroup & " AND " & sTableName & ".Weight <= " & dMaxWeight & " " & vbCrLf & _
''''''          "ORDER BY " & sSortFieldName & " ASC "
''''''   Set rs = db.OpenRecordset(sSQL)
''''''   rs.MoveFirst
''''''
''''''   'INIT GROUPS
''''''   DoCmd.SetWarnings False
''''''   DoCmd.RunSQL ("UPDATE " & sTableName & " SET " & sTableName & ".GroupID = 0, " & sTableName & ".GroupCode = '0'")
''''''   DoCmd.SetWarnings True
''''''
''''''   key = rs!ShipSite
''''''   keyTime = rs!Year & rs!Month & rs!WK
''''''   cubesleft = dMaxCubes
''''''   weightleft = dMaxWeight
''''''   lGroupNum = 1
''''''
''''''
''''''   While Not rs.EOF
''''''      DoEvents
''''''
''''''      If rs!Cubes <= cubesleft And rs!Weight <= weightleft And _
''''''         rs!GroupID = 0 And iGrpStops = 0 Then     '//Group this line in and fill in the Group ID and Code for Stop 1 (Anchor)
''''''
''''''         'Assign it and update the table GroupID and GroupCode
''''''         iGrpStops = iGrpStops + 1
''''''         key = rs!ShipSite
''''''         rs.Edit
''''''         rs!GroupID = lGroupNum
''''''         rs!GroupCode = lGroupNum & "." & iGrpStops
''''''         rs.Update
''''''         cubesleft = cubesleft - rs!Cubes
''''''         weightleft = weightleft - rs!Weight
''''''
''''''         Set rs2Fit = Nothing
''''''         sSQL = "SELECT * FROM " & vbCrLf & _
''''''                "( " & vbCrLf & _
''''''                "SELECT " & vbCrLf & _
''''''                "   " & sTableName & ".*, " & vbCrLf & _
''''''                "   " & sTableName & ".Cubes AS PkgCubes, " & vbCrLf & _
''''''                "   " & sTableName & ".Weight AS PkgWeight, " & vbCrLf & _
''''''                "   Round(" & dMaxCubes & " - [" & sTableName & "_1].[Cubes] - [" & sTableName & "].[Cubes],2) AS CubeLim, " & vbCrLf & _
''''''                "   Abs([" & sTableName & "].[ShipDate]-[" & sTableName & "_1].[ShipDate]) AS DayDelta, " & vbCrLf & _
''''''                "   Abs(CInt([" & sTableName & "].[Zip2])-CInt([" & sTableName & "_1].[Zip2])) AS Z2Delta, " & vbCrLf & _
''''''                "   " & sTableName & "_1.GroupID AS AnchorID, " & vbCrLf & _
''''''                "     Round( [CubeLim] + IIF([DayDelta]=0,-100,60*[DayDelta]) + IIF([Z2Delta]=0,-250,150*[Z2Delta])  ,4) AS DistFromAnchorSCORE " & vbCrLf & _
''''''                "FROM " & sTableName & " INNER JOIN " & sTableName & " AS " & sTableName & "_1 ON " & sTableName & ".ShipSite = " & sTableName & "_1.ShipSite" & vbCrLf & _
''''''                "WHERE " & vbCrLf & _
''''''                "   " & cubesleft & "  >= [" & sTableName & "].[Cubes] AND " & vbCrLf & _
''''''                "   Abs([" & sTableName & "].[ShipDate]-[" & sTableName & "_1].[ShipDate]) <= " & dMaxDeltaDays & " AND " & vbCrLf & _
''''''                "   Abs(CInt([" & sTableName & "].[Zip2])-CInt([" & sTableName & "_1].[Zip2])) <= " & dMaxDeltaZip2 & " AND " & vbCrLf & _
''''''                "   " & sTableName & "_1.GroupID = " & rs!GroupID & " AND " & vbCrLf & _
''''''                "   " & sTableName & ".GroupID = 0" & " " & vbCrLf & _
''''''                "ORDER BY " & vbCrLf & _
''''''                "   " & sTableName & ".Cubes DESC " & vbCrLf & _
''''''                " " & vbCrLf
''''''
''''''         ''============================
''''''         '' THIS CODE BELOW IS ALL ON NOW.  HERE IS THE UNION WITH THE DOUBLETONS
''''''         ''============================
''''''         sSQL = sSQL & "" & vbCrLf & _
''''''                "UNION ALL" & vbCrLf & _
''''''                       "" & vbCrLf
''''''
''''''         sSQL = sSQL & "SELECT  " & vbCrLf & _
''''''                "   " & sTableName & ".ID" & " & '...' & " & " " & sTableName & "Pair.ID AS ID,  " & vbCrLf & _
''''''                "   " & sTableName & ".Delivery" & " & '...' & " & " " & sTableName & "Pair.Delivery AS Delivery,  " & vbCrLf & _
''''''                "   " & sTableName & ".ShipSite" & " & '...' & " & " " & sTableName & "Pair.ShipSite AS ShipSite,  " & vbCrLf & _
''''''                "   " & sTableName & ".ShipDate" & " & '...' & " & " " & sTableName & "Pair.ShipDate AS ShipDate,  " & vbCrLf & _
''''''                "   " & sTableName & ".Year" & " & '...' & " & " " & sTableName & "Pair.Year AS [Year],  " & vbCrLf & _
''''''                "   " & sTableName & ".Month" & " & '...' & " & " " & sTableName & "Pair.Month AS [Month],  " & vbCrLf & _
''''''                "   " & sTableName & ".WK" & " & '...' & " & " " & sTableName & "Pair.WK AS WK,  " & vbCrLf & _
''''''                "   " & sTableName & ".Mode" & " & '...' & " & " " & sTableName & "Pair.Mode AS Mode,  " & vbCrLf & _
''''''                "   " & sTableName & ".Weight" & " & '...' & " & " " & sTableName & "Pair.Weight AS Weight,  " & vbCrLf & _
''''''                "   " & sTableName & ".Cubes" & " & '...' & " & " " & sTableName & "Pair.Cubes AS Cubes,  " & vbCrLf & _
''''''                "   " & sTableName & ".Dest" & " & '...' & " & " " & sTableName & "Pair.Dest AS Dest,  " & vbCrLf & _
''''''                "   " & sTableName & ".Zip2" & " & '...' & " & " " & sTableName & "Pair.Zip2 AS Zip2,  " & vbCrLf & _
''''''                "   " & sTableName & ".Zip2State" & " & '...' & " & " " & sTableName & "Pair.Zip2State AS Zip2State,  " & vbCrLf & _
''''''                "   " & sTableName & ".Zip3" & " & '...' & " & " " & sTableName & "Pair.Zip3 AS Zip3,  " & vbCrLf & _
''''''                "   " & sTableName & ".DestZipCode" & " & '...' & " & " " & sTableName & "Pair.DestZipCode AS DestZipCode,  " & vbCrLf & _
''''''                "   " & sTableName & ".GroupSortPriority" & " & '...' & " & " " & sTableName & "Pair.GroupSortPriority AS GroupSortPriority,  " & vbCrLf & _
''''''                "   " & sTableName & ".GroupID" & " & '...' & " & " " & sTableName & "Pair.GroupID AS GroupID,  " & vbCrLf & _
''''''                "   " & sTableName & ".GroupCode" & " & '...' & " & " " & sTableName & "Pair.GroupCode AS GroupCode,  " & vbCrLf
''''''
''''''         sSQL = sSQL & _
''''''                "   [" & sTableName & "].[Cubes]+[" & sTableName & "Pair].[Cubes] AS PkgCubes,  " & vbCrLf & _
''''''                "   [" & sTableName & "].[Weight]+[" & sTableName & "Pair].[Weight] AS PkgWeight,  " & vbCrLf & _
''''''                "   Round(" & cubesleft & " - ([" & sTableName & "].[Cubes]+[" & sTableName & "Pair].[Cubes]), 2) AS CubeLim,  " & vbCrLf & _
''''''                "   Abs(([" & sTableName & "].[ShipDate]+[" & sTableName & "Pair].[ShipDate])/2-[" & sTableName & "_1].[ShipDate]) AS DayDelta,  " & vbCrLf & _
''''''                "   Abs((CInt([" & sTableName & "].[Zip2])+CInt([" & sTableName & "Pair].[Zip2]))/2-CInt([" & sTableName & "_1].[Zip2])) AS Z2Delta,  " & vbCrLf & _
''''''                "   " & sTableName & "_1.GroupID AS AnchorID,  " & vbCrLf & _
''''''                "     Round( [CubeLim] + IIF([DayDelta]=0,-100,60*[DayDelta]) + IIF([Z2Delta]=0,-250,150*[Z2Delta])  ,4) AS DistFromAnchorSCORE " & vbCrLf & _
''''''                "FROM (" & sTableName & " INNER JOIN " & sTableName & " AS " & sTableName & "_1 ON  " & vbCrLf & _
''''''                "   " & sTableName & ".ShipSite = " & sTableName & "_1.ShipSite) INNER JOIN " & sTableName & " AS " & sTableName & "Pair ON  " & vbCrLf & _
''''''                "   " & sTableName & ".ShipSite = " & sTableName & "Pair.ShipSite " & vbCrLf
''''''
''''''         sSQL = sSQL & _
''''''                "WHERE  " & vbCrLf & _
''''''                "   (([" & sTableName & "].[ShipDate])  <= " & dMaxDeltaDays & " + ([" & sTableName & "_1].[ShipDate])) AND " & vbCrLf & _
''''''                "   (([" & sTableName & "].[ShipDate])  >= " & -dMaxDeltaDays & " + ([" & sTableName & "_1].[ShipDate])) AND " & vbCrLf & _
''''''                "   (([" & sTableName & "Pair].[ShipDate])  <= " & dMaxDeltaDays & " + ([" & sTableName & "_1].[ShipDate])) AND " & vbCrLf & _
''''''                "   (([" & sTableName & "Pair].[ShipDate])  >= " & -dMaxDeltaDays & " + ([" & sTableName & "_1].[ShipDate])) AND " & vbCrLf & _
''''''                "   (CInt([" & sTableName & "].[Zip2])  <= " & dMaxDeltaZip2 & " + CInt([" & sTableName & "_1].[Zip2])) AND " & vbCrLf & _
''''''                "   (CInt([" & sTableName & "].[Zip2])  >= " & -dMaxDeltaZip2 & " + CInt([" & sTableName & "_1].[Zip2])) AND " & vbCrLf & _
''''''                "   (CInt([" & sTableName & "Pair].[Zip2])  <= " & dMaxDeltaZip2 & " + CInt([" & sTableName & "_1].[Zip2])) AND " & vbCrLf & _
''''''                "   (CInt([" & sTableName & "Pair].[Zip2])  >= " & -dMaxDeltaZip2 & " + CInt([" & sTableName & "_1].[Zip2])) AND " & vbCrLf & _
''''''                "   ((" & sTableName & "_1.GroupID)=" & rs!GroupID & ") AND ((" & dMaxCubes & " - [" & sTableName & "_1].[Cubes]) >= [" & sTableName & "].[Cubes]+[" & sTableName & "Pair].[Cubes]) AND  " & vbCrLf & _
''''''                "   ((" & sTableName & ".ShipDate - " & sTableName & "Pair.ShipDate) <= 2*" & dMaxDeltaDays & ") AND  " & vbCrLf & _
''''''                "   [" & sTableName & "].[Cubes]+[" & sTableName & "Pair].[Cubes] <= " & cubesleft & " AND  " & vbCrLf & _
''''''                "   ((" & sTableName & ".GroupID)=0) AND ((" & sTableName & "Pair.GroupID)=0) " & vbCrLf
''''''
''''''
''''''         sSQL = sSQL & _
''''''                ") AS aliOneTwoDel " & vbCrLf & _
''''''                "ORDER BY DistFromAnchorSCORE, CubeLim " & vbCrLf
''''''
''''''         ''============================
''''''
''''''         Set rs2Fit = db.OpenRecordset(sSQL)
''''''
''''''
''''''         If rs2Fit.EOF Then  'NO CANDIDATES TO JOIN SO IT IS A SINGLETON
''''''            Debug.Print lGroupNum & " . . . "
''''''            Debug.Print "   " & " . " & "currTableID " & rs!id & " assigned " & rs!Cubes & " Cubes. [ANCHOR] -- SINGLETON."
''''''         Else
''''''            rs2Fit.MoveFirst 'MOVE TO FIRST RECORD OF CANDIDATES
''''''            Debug.Print lGroupNum & " . . . "
''''''            Debug.Print "   " & " . " & "currTableID " & rs!id & " assigned " & rs!Cubes & " Cubes. [ANCHOR] "
''''''         End If
''''''
''''''      End If
''''''
''''''
''''''      Do While Not rs2Fit.EOF
''''''         DoEvents
''''''
''''''         If rs2Fit!PkgCubes <= cubesleft And rs2Fit!PkgWeight <= weightleft And _
''''''            Left(rs2Fit!GroupID, 1) = 0 And iGrpStops <> 0 Then    '//Group this line in and fill in the Group ID and Code for next stop
''''''            '============================================================>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
''''''            DoCmd.SetWarnings False
''''''            'ok   right here check for double or single and add the one or add the two.
''''''
''''''            If InStr(rs2Fit!id, "...") > 0 Then 'DOUBLETON
''''''               'Assign THEM and update BOTH RECORDS FOR the GroupID and GroupCode
''''''               iGrpStops = iGrpStops + 2
''''''               'UPDATE THE FIRST ONE
''''''               sSQL = "UPDATE " & sTableName & " SET " & sTableName & ".GroupID = " & lGroupNum & ", " & _
''''''                                  sTableName & ".GroupCode = '" & lGroupNum & "." & iGrpStops & "d' " & _
''''''                             "WHERE " & sTableName & ".ID=" & Trim(Left(rs2Fit!id, InStr(rs2Fit!id, "...") - 1)) & " "
''''''               DoCmd.RunSQL (sSQL)
''''''
''''''               'UPDATE THE SECOND ONE
''''''               sSQL = "UPDATE " & sTableName & " SET " & sTableName & ".GroupID = " & lGroupNum & ", " & _
''''''                                  sTableName & ".GroupCode = '" & lGroupNum & "." & iGrpStops + 1 & "dd' " & _
''''''                             "WHERE " & sTableName & ".ID=" & Trim(Right(rs2Fit!id, Len(rs2Fit!id) - (InStr(rs2Fit!id, "...") + 2))) & " "
''''''               DoCmd.RunSQL (sSQL)
''''''            Else 'SINGLETON
''''''               'Assign it and update the table GroupID and GroupCode
''''''               iGrpStops = iGrpStops + 1
''''''               DoCmd.RunSQL ("UPDATE " & sTableName & " SET " & sTableName & ".GroupID = " & lGroupNum & ", " & _
''''''                                                                sTableName & ".GroupCode = '" & lGroupNum & "." & iGrpStops & "s' " & _
''''''                             "WHERE " & sTableName & ".ID=" & rs2Fit!id & "")
''''''            End If
''''''
''''''            DoCmd.SetWarnings True
''''''            cubesleft = cubesleft - rs2Fit!PkgCubes
''''''            weightleft = weightleft - rs2Fit!PkgWeight
''''''            Debug.Print "   " & " . " & "currTableID " & rs2Fit!id & " assigned " & rs2Fit!PkgCubes & " Cubes. "
''''''            '============================================================>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
''''''         End If
''''''
''''''         'IF IT IS ENOUGH
''''''         If cubesleft / dMaxCubes <= dPctTolerance Or weightleft / dMaxWeight <= dPctTolerance Or iGrpStops = iMaxStops Then
''''''            'reset rs to unallocated records
''''''            sSQL = "SELECT * FROM " & sTableName & " " & vbCrLf & _
''''''                   "WHERE " & "" & sTableName & ".Cubes <= " & dMaxCubes & " AND " & sTableName & ".Cubes <= " & iMaxCubesToGroup & " AND " & sTableName & ".Weight <= " & dMaxWeight & " " & vbCrLf & _
''''''                   "  AND " & sTableName & ".GroupID = " & 0 & " " & vbCrLf & _
''''''                   "ORDER BY " & sSortFieldName & " ASC "
''''''            Set rs = db.OpenRecordset(sSQL)
''''''            rs.MoveFirst
''''''            Debug.Print "       . . . .  . . . . " & Format(Now(), "hh:nn:ss") & "   " & rs.RecordCount - 1 & " records left working . . . . . . . . GOT FILLED UP."
''''''            lGroupNum = lGroupNum + 1           'increment counter to next Group
''''''            cubesleft = dMaxCubes
''''''            weightleft = dMaxWeight
''''''            iGrpStops = 0
''''''            Exit Do  'BREAK OUT OF THE WHILE LOOP
''''''         Else
''''''
''''''         End If
''''''
''''''         rs2Fit.MoveNext
''''''
''''''      Loop
''''''
''''''      If cubesleft = dMaxCubes Then
''''''         'you got filled up and advanced
''''''      Else
''''''         'you did not reach a stopping criteria, but you ran out of group candidates
''''''         'reset rs
''''''         sSQL = "SELECT * FROM " & sTableName & " " & vbCrLf & _
''''''                "WHERE " & "" & sTableName & ".Cubes <= " & dMaxCubes & " AND " & sTableName & ".Cubes <= " & iMaxCubesToGroup & " AND " & sTableName & ".Weight <= " & dMaxWeight & " " & vbCrLf & _
''''''                "  AND " & sTableName & ".GroupID = " & 0 & " " & vbCrLf & _
''''''                "ORDER BY " & sSortFieldName & " ASC "
''''''         Set rs = db.OpenRecordset(sSQL)
''''''         If rs.EOF Then 'done
''''''         Else
''''''            rs.MoveFirst
''''''            '''Debug.Print "       . . . .  . . . . " & Format(Now(), "hh:nn:ss") & "   " & rs.RecordCount - 1 & " records left working . . . . . . . . RAN OUT OF CANDIDATES."
''''''            lGroupNum = lGroupNum + 1           'increment counter to next Group
''''''            cubesleft = dMaxCubes
''''''            weightleft = dMaxWeight
''''''            iGrpStops = 0
''''''         End If
''''''      End If
''''''   Wend
''''''
''''''
''''''   'number the ungrouped ones
''''''   sSQL = "SELECT * FROM " & sTableName & " " & vbCrLf & _
''''''          "WHERE " & sTableName & ".GroupID = " & 0 & " " & vbCrLf & _
''''''          "ORDER BY " & sSortFieldName & " ASC "
''''''   Set rs = db.OpenRecordset(sSQL)
''''''   While Not rs.EOF
''''''      lGroupNum = lGroupNum + 1        'move this up to the point where the singleton is id'd
''''''      rs.Edit       'move this up to the point where the singleton is id'd
''''''      rs!GroupID = lGroupNum       'move this up to the point where the singleton is id'd
''''''      rs!GroupCode = lGroupNum & ".0"  'move this up to the point where the singleton is id'd
''''''      rs.Update
''''''      Debug.Print lGroupNum & " . . . "
''''''   rs.MoveNext
''''''   Wend
''''''
''''''
''''''
''''''   T = Now() - T
''''''   s = "GroupShipments....." & Format(T, "hh:nn:ss")
''''''   Debug.Print s
''''''   Debug.Print "--DONE--"
''''''
''''''   GroupShip = True
''''''
''''''GroupShip_Done:
''''''  Exit Function
''''''
''''''GroupShip_Err:
''''''   Select Case Err
''''''    Case 9 'or 13 'subscript out of range
''''''      GroupShip = False
''''''    Case 3021 'no current record
''''''      GroupShip = True
''''''      Resume GroupShip_Done
''''''    Case Else
''''''      MsgBox "Error: " & Error$ & " " & Err, , "GroupShipments"
''''''      GroupShip = False
''''''   End Select
''''''   Resume GroupShip_Done
''''''   Resume
''''''
''''''End Function
''''''
''''''
''''''Public Function GroupShipments(Optional sTableName = "tblFS1415_LoadSummaryAble", _
''''''                               Optional sSortFieldName = "GroupSortPriority", _
''''''                               Optional dMaxCubes = 3130, Optional dMaxWeight = 24000, _
''''''                               Optional iMaxStops = 3, _
''''''                               Optional iMaxCubesToGroup = 2000) _
''''''                               As Boolean
''''''
'''''''//=================================================================//
'''''''/|   FUNCTION:  GroupShipments                                     |/
'''''''/| PARAMETERS:  Requires tblName, SortField, MaxCubes, MaxWeight,  |/
'''''''/|                                                                 |/
'''''''/|              Table Must include fields GroupID, and GroupCode   |/
'''''''/|              and ShipSite and Cubes and Weight                  |\
'''''''/|                                                                 |/
'''''''/|    RETURNS:  True on Success                                    |/
'''''''/|    PURPOSE:  Group Shipments into truckLoads by weight          |/
'''''''/|              limit or cubic foot limit                          |/
'''''''/|      USAGE:  i= GroupShipments()                                |/
'''''''/|              i= GroupShipments("TBL1","PSORT", 2800, 24000, 2)  |/
'''''''/|         BY:  Sean                                               |/
'''''''/|       DATE:  08/11/2015                                         |/
'''''''/|    HISTORY:  08/11/2015                                         |/
'''''''/|                                                                 |/
'''''''//=================================================================//
''''''
'''''''// i= GroupShipments("tblFS1415_LoadSummaryAble","GroupSortPriority", _
'''''''                       0.85 * CDbl(GetSetting("FullTruckCubicFeet")), _
'''''''                       0.85 * CDbl(GetSetting("FullTruckWeight")), _
'''''''                       2 )
'''''''   i= GroupShipments("t114NETPDelCostGroupShipments","GroupSortPriority", 3130, 24000, 3, 2000)
'''''''   i= GroupShipments("t114NETPDelCostGroupShipments","GroupSortPriority", 2650, 24000, 3, 2400)
''''''
'''''''//============  ** Enhance This:
''''''
''''''Dim db         As Database
''''''Dim rs         As DAO.Recordset
''''''Dim rsClone    As DAO.Recordset
''''''Dim rsShipFit  As DAO.Recordset
''''''Dim sSQL       As String
''''''Dim i          As Integer
''''''Dim T          As Date
''''''Dim s          As String
''''''Dim key        As String
''''''Dim keyTimeBucket As String
''''''Dim weightleft As Double
''''''Dim cubesleft  As Double
''''''Dim sShipSite  As String
''''''Dim lGroupCtr  As Long
''''''Dim PctTolerance As Double
''''''Dim lRecCtr    As Long
''''''Dim iGrpStops  As Integer
''''''
''''''
''''''
''''''
''''''   On Error GoTo GroupShipments_Err
''''''   GroupShipments = False
''''''
''''''   T = Now()
''''''
''''''   Set db = CurrentDb
''''''   sSQL = ""
''''''   iGrpStops = 0
''''''   PctTolerance = 0.1
''''''
''''''   '//  G R O U P   T H E   S H I P M E N T S  //
''''''   sSQL = "SELECT * FROM " & sTableName & " " & _
''''''          "WHERE " & sTableName & ".Cubes <= " & iMaxCubesToGroup & " " & _
''''''          "ORDER BY " & sSortFieldName & " ASC "
''''''   Set rs = db.OpenRecordset(sSQL)
''''''
''''''   Set rsClone = rs.Clone()
''''''   rsClone.MoveFirst
''''''   lRecCtr = 1
''''''
''''''   'INIT GROUPS
''''''   DoCmd.SetWarnings False
''''''   DoCmd.RunSQL ("UPDATE " & sTableName & " SET " & sTableName & ".GroupID = 0, " & sTableName & ".GroupCode = '0'")
''''''   DoCmd.SetWarnings True
''''''
''''''   key = rsClone!ShipSite
''''''   keyTimeBucket = rsClone!Year & rsClone!Month
''''''   '''& rsClone!WK
''''''   cubesleft = dMaxCubes
''''''   weightleft = dMaxWeight
''''''   lGroupCtr = 1
''''''
''''''
''''''   While rsClone.EOF = False And lGroupCtr <= rsClone.RecordCount
''''''      DoEvents
''''''      If rsClone!id = 12512 Then Debug.Print "     ##### 12512 #####    "
''''''     'IF IT FITS AND IT'S THE RIGHT FACILITY, TimeBucket and is still unassigned to a group, START OR ADD TO THE GROUP
''''''      If rsClone!Cubes <= cubesleft And rsClone!Weight <= weightleft And rsClone!ShipSite = key And _
''''''         rsClone!Year & rsClone!Month = keyTimeBucket And _
''''''         rsClone!GroupID = 0 And _
''''''         iGrpStops <= iMaxStops Then     '//Group this line in and fill in the Group ID and Code
''''''         '''& rsClone!WK
''''''         rsClone.Edit
''''''         rsClone!GroupID = lGroupCtr
''''''         rsClone!GroupCode = lGroupCtr
''''''         cubesleft = cubesleft - rsClone!Cubes
''''''         weightleft = weightleft - rsClone!Weight
''''''         rsClone.Update
''''''         iGrpStops = iGrpStops + 1
''''''         If iGrpStops = 1 Then Debug.Print lGroupCtr & " . . . "
''''''         Debug.Print "   " & "recs seen = " & lRecCtr & " . " & "currTableID " & rsClone!id & " assigned " & rsClone!Weight & " lbs. "
''''''      Else
''''''         key = rsClone!ShipSite
''''''         rsClone.MoveNext
''''''         lRecCtr = lRecCtr + 1
''''''      End If
''''''
''''''
''''''     'IF IT IS ENOUGH
''''''      If cubesleft / dMaxCubes <= PctTolerance Or weightleft / dMaxWeight <= PctTolerance Or iGrpStops = iMaxStops Then
''''''         lGroupCtr = lGroupCtr + 1           'increment counter to next Group
''''''         ''' Alternatively, reset the rsClone right here to only open candidates
''''''         rsClone.MoveFirst
''''''         rsClone.FindFirst "[GroupID] = 0 And [Cubes] <= " & dMaxCubes & " AND [Weight] <= " & dMaxWeight & " "   'go back to first unassigned record
''''''         lRecCtr = lRecCtr + 1
''''''         cubesleft = dMaxCubes
''''''         weightleft = dMaxWeight
''''''         iGrpStops = 0
''''''         key = rsClone!ShipSite
''''''         '''& rsClone!WK
''''''      Else
''''''         If rsClone.EOF = False Then
''''''            key = rsClone!ShipSite
''''''            '''rsClone.MoveNext
''''''            '''lRecCtr = lRecCtr + 1
''''''         End If
''''''      End If
''''''      keyTimeBucket = rsClone!Year & rsClone!Month
''''''   Wend
''''''
''''''
''''''   T = Now() - T
''''''   s = "GroupShipments....." & Format(T, "hh:nn:ss")
''''''   Debug.Print s
''''''   Debug.Print "--DONE--"
''''''
''''''   GroupShipments = True
''''''
''''''GroupShipments_Done:
''''''  Exit Function
''''''
''''''GroupShipments_Err:
''''''   Select Case Err
''''''    Case 9 'or 13 'subscript out of range
''''''      GroupShipments = False
''''''    Case 3021 'no current record
''''''      GroupShipments = True
''''''      Resume GroupShipments_Done
''''''    Case Else
''''''      MsgBox "Error: " & Error$ & " " & Err, , "GroupShipments"
''''''      GroupShipments = False
''''''   End Select
''''''   Resume GroupShipments_Done
''''''   Resume
''''''
''''''End Function
''''''
''''''
''''''
''''''
''''''Public Function GroupShipmentsCode(Optional sTableName = "tblFS1415_LoadSummaryAbleTest", _
''''''                               Optional sSortFieldName = "GroupSortPriority", _
''''''                               Optional dMaxCubes = 3130, Optional dMaxWeight = 24000, _
''''''                               Optional iMaxStops = 3, _
''''''                               Optional iMaxCubesToGroup = 2000) _
''''''                               As Boolean
''''''
'''''''//=================================================================//
'''''''/|   FUNCTION:  GroupShipments                                     |/
'''''''/| PARAMETERS:  Requires tblName, SortField, MaxCubes, MaxWeight,  |/
'''''''/|                                                                 |/
'''''''/|              Table Must include fields GroupID, and GroupCode   |/
'''''''/|              and ShipSite and Cubes and Weight                  |\
'''''''/|                                                                 |/
'''''''/|    RETURNS:  True on Success                                    |/
'''''''/|    PURPOSE:  Group Shipments into truckLoads by weight          |/
'''''''/|              limit or cubic foot limit                          |/
'''''''/|      USAGE:  i= GroupShipments()                                |/
'''''''/|              i= GroupShipments("TBL1","PSORT", 2800, 24000, 2)  |/
'''''''/|         BY:  Sean                                               |/
'''''''/|       DATE:  08/11/2015                                         |/
'''''''/|    HISTORY:  08/11/2015                                         |/
'''''''/|                                                                 |/
'''''''//=================================================================//
''''''
'''''''// i= GroupShipmentsCode("tblFS1415_LoadSummaryAble","GroupSortPriority", _
'''''''                       0.85 * CDbl(GetSetting("FullTruckCubicFeet")), _
'''''''                       0.85 * CDbl(GetSetting("FullTruckWeight")), _
'''''''                       2 )
'''''''   i= GroupShipmentsCode("t114NETPDelCostGroupShipments","GroupSortPriority", 2650, 24000, 3, 2400)
''''''
'''''''//============  ** Enhance This:
''''''
''''''Dim db         As Database
''''''Dim rs         As DAO.Recordset
''''''Dim rsClone    As DAO.Recordset
''''''Dim rsShipFit  As DAO.Recordset
''''''Dim sSQL       As String
''''''Dim i          As Integer
''''''Dim T          As Date
''''''Dim s          As String
''''''Dim key        As String
''''''Dim keyTimeBucket As String
''''''Dim weightleft As Double
''''''Dim cubesleft  As Double
''''''Dim sShipSite  As String
''''''Dim lGroupCtr  As Long
''''''Dim PctTolerance As Double
''''''Dim lRecCtr    As Long
''''''Dim iGrpStops  As Integer
''''''Dim weightAvailPct As Double
''''''Dim cubesAvailPct  As Double
''''''Dim fNewRecordFlag As Boolean
''''''
''''''
''''''
''''''
''''''   On Error GoTo GroupShipmentsCode_Err
''''''   GroupShipmentsCode = False
''''''
''''''   T = Now()
''''''
''''''   Set db = CurrentDb
''''''   sSQL = ""
''''''   iGrpStops = 0
''''''   cubesAvailPct = 0
''''''   weightAvailPct = 0
''''''   PctTolerance = 0.1
''''''   fNewRecordFlag = False
''''''
''''''   '//  G R O U P   T H E   S H I P M E N T S  //
''''''   sSQL = "SELECT * FROM " & sTableName & " " & _
''''''          "WHERE " & sTableName & ".Cubes <= " & iMaxCubesToGroup & " " & _
''''''          "ORDER BY " & sSortFieldName & " ASC "
''''''   Set rs = db.OpenRecordset(sSQL)
''''''
''''''   Set rsClone = rs.Clone()
''''''   rsClone.MoveFirst
''''''   lRecCtr = 1
''''''
''''''   'INIT GROUPS
''''''   DoCmd.SetWarnings False
''''''   DoCmd.RunSQL ("UPDATE " & sTableName & " SET " & sTableName & ".GroupID = 0, " & sTableName & ".GroupCode = '0'")
''''''   DoCmd.SetWarnings True
''''''
''''''   key = rsClone!ShipSite
''''''   keyTimeBucket = rsClone!Year & rsClone!Month
''''''   '''& rsClone!WK
''''''   cubesleft = dMaxCubes
''''''   weightleft = dMaxWeight
''''''   lGroupCtr = 1
''''''
''''''
''''''   While rsClone.EOF = False And lGroupCtr <= rsClone.RecordCount
''''''      DoEvents
''''''      '''If rsClone!id = 12512 Then Debug.Print "     ##### 12512 #####    "
''''''     'IF IT FITS AND IT'S THE RIGHT FACILITY, TimeBucket and is still unassigned to a group, START OR ADD TO THE GROUP
''''''      If rsClone!Cubes <= cubesleft And rsClone!Weight <= weightleft And rsClone!ShipSite = key And _
''''''         rsClone!Year & rsClone!Month = keyTimeBucket And _
''''''         rsClone!GroupID = 0 And _
''''''         iGrpStops <= iMaxStops Then     '//Group this line in and fill in the Group ID and Code
''''''         '''& rsClone!WK
''''''         rsClone.Edit
''''''         rsClone!GroupID = lGroupCtr
''''''         rsClone!GroupCode = lGroupCtr
''''''         cubesleft = cubesleft - rsClone!Cubes
''''''         weightleft = weightleft - rsClone!Weight
''''''         cubesAvailPct = cubesleft / dMaxCubes
''''''         weightAvailPct = weightleft / dMaxWeight
''''''         rsClone.Update
''''''         iGrpStops = iGrpStops + 1
''''''
''''''         If iGrpStops = 1 Then Debug.Print lGroupCtr & " . . . "
''''''         Debug.Print "   " & "recs seen = " & lRecCtr & " . " & "currTableID " & rsClone!id & " assigned " & rsClone!Weight & " lbs. "
''''''         key = rsClone!ShipSite
''''''         lRecCtr = lRecCtr + 1
''''''         rsClone.MoveNext
''''''         fNewRecordFlag = False
''''''      Else
''''''         key = rsClone!ShipSite
''''''         rsClone.MoveNext
''''''         lRecCtr = lRecCtr + 1
''''''      End If
''''''
''''''
''''''
''''''
''''''     'IF IT IS ENOUGH
''''''      If cubesAvailPct <= PctTolerance Or weightAvailPct <= PctTolerance Or iGrpStops = iMaxStops Then
''''''         lGroupCtr = lGroupCtr + 1           'increment counter to next Group
''''''         rsClone.MoveFirst
''''''         rsClone.FindFirst "[GroupID] = 0 And [Cubes] <= " & dMaxCubes & " AND [Weight] <= " & dMaxWeight & " " 'go back to first unassigned record
''''''
''''''         ''' Alternatively, reset the rsClone right here to only open candidates
''''''
''''''         lRecCtr = lRecCtr + 1
''''''         cubesleft = dMaxCubes
''''''         weightleft = dMaxWeight
''''''         cubesAvailPct = cubesleft / dMaxCubes
''''''         weightAvailPct = weightleft / dMaxWeight
''''''         iGrpStops = 0
''''''         key = rsClone!ShipSite
''''''
''''''         Debug.Print "currTableID " & rsClone!id
''''''         '''& rsClone!WK
''''''      Else
''''''         If rsClone.EOF = False Then
''''''            key = rsClone!ShipSite
''''''            rsClone.MoveNext
''''''            lRecCtr = lRecCtr + 1
''''''         End If
''''''      End If
''''''      keyTimeBucket = rsClone!Year & rsClone!Month
''''''   Wend
''''''
''''''
''''''   T = Now() - T
''''''   s = "GroupShipmentsCode....." & Format(T, "hh:nn:ss")
''''''   Debug.Print s
''''''   Debug.Print "--DONE--"
''''''
''''''   GroupShipmentsCode = True
''''''
''''''GroupShipmentsCode_Done:
''''''  Exit Function
''''''
''''''GroupShipmentsCode_Err:
''''''   Select Case Err
''''''    Case 9 'or 13 'subscript out of range
''''''      GroupShipmentsCode = False
''''''    Case 3021 'no current record
''''''      GroupShipmentsCode = True
''''''      Resume GroupShipmentsCode_Done
''''''    Case Else
''''''      MsgBox "Error: " & Error$ & " " & Err, , "GroupShipmentsCode"
''''''      GroupShipmentsCode = False
''''''   End Select
''''''   Resume GroupShipmentsCode_Done
''''''   Resume
''''''
''''''
''''''End Function
''''''
''''''
