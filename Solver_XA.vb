Imports Optimizer
Imports EntLib.COPT
Namespace COPT
    <ComClass(Solver_XA.ClassId, Solver_XA.InterfaceId, Solver_XA.EventsId)> _
    Public Class Solver_XA
        Inherits Solver

        Structure XAOSL
            Dim rc As Long '  Return code            - returned
            Dim XAOSLSize As Long ' Size of this structure - user supplied
            Dim AuthCode As Long ' Activation code        - Internal usage
            Dim Reserve1 As Long ' Reserve 1
            Dim Reserve2 As Long ' Reserve 2
            Dim CumXATime As Long ' Cum time spent in XA routines
            Dim LastTime As Long ' Time spent in last function
            Dim XAComArea As Long ' Pointer to XA communication area - returned
            Dim allCount As Long ' Number of XA function calls      - returned
            Dim ProgArea As Long ' Available to programmer          - user area
            'Dim Reserved(13) As Long ' Reserved for later use
            Dim Reserved() As Long ' Reserved for later use
        End Structure
        Private Const COL_NAME_LEN As Short = 164 ' Number of characters in column names.
        Private Const ROW_NAME_LEN As Short = 164 ' Number of characters in row name.
        Private _rc As Optimizer.XA.XAOSL
        Private _solutionStatus As String
        Private _solutionRows As Integer
        Private _solutionColumns As Integer
        Private _solutionNonZeros As Integer
        Private _solutionObj As Double
        Private _solutionIterations As Integer
        Private Shadows _switches() As String
        Private Shadows _engine As COPT.Engine 'RKP/01-26-10/v2.3.127
        Private _progress As String = "" 'RKP/08-05-11/v3.0.149
#Region "COM GUIDs"
        ' These  GUIDs provide the COM identity for this class 
        ' and its COM interfaces. If you change them, existing 
        ' clients will no longer be able to access the class.
        Public Const ClassId As String = "9a4630a1-cf1f-4174-83ec-34f5f5408755"
        Public Const InterfaceId As String = "95cb6a22-12ce-4f72-8a6c-32bc86752999"
        Public Const EventsId As String = "83fb6760-a225-4ce7-9a69-e26da196c8d0"
#End Region

        ' A creatable COM class must have a Public Sub New() 
        ' with no parameters, otherwise, the class will not be 
        ' registered in the COM registry and cannot be created 
        ' via CreateObject.
        Public Sub New()
            MyBase.New(New COPT.Engine)
        End Sub

        Public Sub New(ByVal engine As COPT.Engine)
            MyBase.New(engine)
        End Sub

        Public Sub New(ByVal engine As COPT.Engine, ByVal switches() As String)
            MyBase.New(engine, switches)
            _switches = switches
            _engine = engine
        End Sub

        Public Overrides Function Solve(ByRef ds As DataSet) As Boolean
            Return Solve(ds.Tables("tsysRow"), ds.Tables("tsysCol"), ds.Tables("tsysMtx"))
        End Function

        Public Overrides Function Solve(ByRef dtRow As DataTable, ByRef dtCol As DataTable, ByRef dtMtx As DataTable) As Boolean

            Dim lpProbInfo As typLPprobInfo = MyBase.ProbInfo
            Dim xa As New Optimizer.XA.Optimize(lpProbInfo.intMaxMemory)
            Dim i As Integer, objFcnValue As Double
            Dim CName As String, Coef As Double, RName As String, Sense As String
            Dim sql As String
            Dim col As String, row As String
            Dim filePath As String = ""
            Dim startTime As Integer = My.Computer.Clock.TickCount
            Dim startSolveTime As Integer = My.Computer.Clock.TickCount
            Dim errSource As String = ""
            Dim nonZeroCount As Integer = _engine.SolutionNonZeros 'dtMtx.Rows.Count
            Dim ctyp() As Char = Nothing
            Dim proceed As Boolean = False
            'Dim linqTable As System.Data.EnumerableRowCollection(Of System.Data.DataRow)
            'Dim intGroups As System.Collections.Generic.IEnumerable(Of Integer)
            'Dim charQueryResults As System.Data.EnumerableRowCollection(Of Char)
            Dim updateSQLServer As Boolean = False 'RKP/08-05-11/v3.0.149
            Dim result As Integer 'RKP/08-05-11/v3.0.149
            Dim linkedTable As Boolean = False 'RKP/08-05-11/v3.0.149
            Dim success As Boolean = False 'RKP/04-18-12/v3.2.166
            Dim noTruncate As Boolean = GenUtils.IsSwitchAvailable(_switches, "/UseTruncateSP")

            If GenUtils.IsSwitchAvailable(_switches, "/UseMSAccessSyntax") Then
                linkedTable = True
            End If
            If GenUtils.IsSwitchAvailable(_switches, "/UseSQLServerSyntax") Then
                linkedTable = True
            End If

            Try
                xa.setActivationCodes(0, 0)
                errSource = "xa.openConnection"
                xa.openConnection()
                errSource = ""
                xa.setModelSize(lpProbInfo.intRowCt + 1, lpProbInfo.intColCt + 1, lpProbInfo.intCoefCt + lpProbInfo.intColCt + 1, lpProbInfo.intMasterLength + 1, lpProbInfo.intMasterLength + 1)
                'xa.setModelSize(lpProbInfo.intRowCt, lpProbInfo.intColCt)
                Debug.Print(lpProbInfo.intRowCt & ", " & lpProbInfo.intColCt & ", " & lpProbInfo.intCoefCt)

                'three lines from Jim B. on 12/19/07
                'xa.setCommand("Maximize Yes Presolve 0 ")
                'xa.setCommand("ListInput No Set Debug No ")
                'xa.setCommand("MpsxSolutionReport Yes Output c:\xa.log")

                'older set
                'xa.setCommand("Maximize Yes Mute No MpsxSolutionReport Yes Set Debug No Presolve 0 Output c:\xa.log")

                'the MPS generator line from Jim in early '08
                'xa.setCommand(" FileName  c:\pathInfoStuff\myMpsFileName  ToMps Yes" ) 

                'the mps file options are below near the solve
                xa.setCommand("Maximize Yes Presolve 0 ")
                xa.setCommand("ListInput No Set Debug No ")
                filePath = GenUtils.GetSwitchArgument(_switches, "/PathSolverLog", 1)
                'xa.setCommand("MpsxSolutionReport Yes Output c:\xa.log")
                If GenUtils.IsSwitchAvailable(_switches, "/LogWorkDir") Then
                    filePath = GenUtils.GetWorkDir(_switches) & "\xa.log"
                Else
                    filePath = GenUtils.GetSwitchArgument(_switches, "/PathSolverLog", 1)
                End If
                xa.setCommand("MpsxSolutionReport Yes Output " & filePath)


                'THESE WORK IN MARCH '08
                'xa.setCommand("Maximize Yes FileName C:\XA.MPS ToMps Yes Presolve 0 ")
                'xa.setCommand("ListInput No Set Debug No ")
                'xa.setCommand("MpsxSolutionReport Yes Output c:\xa.log")
                'xa.setCommand("Set CmprsName _")


                'THE ORIGINAL LINES ?aug 2007?  -- don't change them
                'xa.setCommand("Maximize Yes ToMPS Yes Set FreqLog 0:01")
                'xa.setCommand("Output c:\\xa.log MatList e ")


                startSolveTime = My.Computer.Clock.TickCount

                'MyBase.LoadSolverArrays(Nothing)
                'MyBase.LoadSolverArrays(dtRow, dtCol, dtMtx, False)

                ''---CTYP---
                '_engine.ProblemType = COPT.Engine.problemRunType.problemTypeContinuous
                'proceed = False
                'linqTable = dtCol.AsEnumerable()
                ''If usePLINQ Then linqTable.AsParallel()
                'intGroups = From r In linqTable _
                '                    Where r!BINRY = True _
                '                    Order By r("ColID") Ascending _
                '                    Group By r!ColID Into g = Group Select g.Count()
                'If intGroups.ToArray().Count > 0 Then
                '    proceed = True
                '    _engine.ProblemType = COPT.Engine.problemRunType.problemTypeBinary
                'End If
                ''Else
                'intGroups = From r In linqTable _
                '                    Where r!INTGR = True _
                '                    Order By r("ColID") Ascending _
                '                    Group By r!ColID Into g = Group Select g.Count()
                'If intGroups.ToArray().Count > 0 Then
                '    proceed = True
                '    '_engine.ProblemType = "I"
                '    If _engine.ProblemType = COPT.Engine.problemRunType.problemTypeContinuous Then
                '        _engine.ProblemType = COPT.Engine.problemRunType.problemTypeInteger
                '    Else
                '        _engine.ProblemType = _engine.ProblemType + COPT.Engine.problemRunType.problemTypeInteger
                '    End If
                'End If
                ''End If
                'If proceed Then
                '    linqTable = dtCol.AsEnumerable()
                '    'If usePLINQ Then linqTable.AsParallel()
                '    'If usePLINQ Then linqTable.AsParallel()
                '    charQueryResults = From r In linqTable _
                '                   Order By r("ColID") Ascending _
                '                   Select INTGR = _
                '                    CChar( _
                '                        IIf( _
                '                            r("BINRY") = True, _
                '                            "B", _
                '                            IIf( _
                '                                r("INTGR") = True, _
                '                                "I", _
                '                                "C" _
                '                            ) _
                '                        ) _
                '                    )
                '    array_ctyp = charQueryResults.ToArray()
                '    'ds.Tables("tsysCol")(0)("UP").ToString()
                '    For i = 0 To dtCol.Rows.Count - 1
                '        If dtCol(i)("BINRY") Then
                '            dtCol(i)("UP") = 1
                '        End If
                '    Next
                'End If
                ''---CTYP---


                '
                ' Load Column Data ****************************************
                For i = 0 To dtCol.Rows.Count - 1
                    CName = dtCol.Rows(i).Item("COL").ToString

                    Coef = CDbl(dtCol.Rows(i).Item("OBJ"))
                    xa.loadPoint("OBJ", CName, Coef)

                    If CBool(dtCol.Rows(i).Item("FREE")) Then
                        xa.setColumnFree(CName)
                    End If

                    If CBool(dtCol.Rows(i).Item("INTGR")) Then
                        xa.setColumnInteger(CName)
                    End If

                    'RKP/01-26-10/v2.3.127
                    If CBool(dtCol.Rows(i).Item("BINRY")) Then
                        xa.setColumnBinary(CName)
                    End If

                    Coef = CDbl(dtCol.Rows(i).Item("LO"))
                    xa.loadPoint("MIN", CName, Coef)

                    Coef = CDbl(dtCol.Rows(i).Item("UP"))
                    xa.loadPoint("MAX", CName, Coef)
                Next
                EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - XA - Loading Col took: ", GenUtils.FormatTime(startSolveTime, My.Computer.Clock.TickCount))

                startSolveTime = My.Computer.Clock.TickCount
                '
                ' Load Row Data *********************************************
                For i = 0 To dtRow.Rows.Count - 1
                    RName = dtRow.Rows(i).Item("ROW").ToString

                    Coef = CDbl(dtRow.Rows(i).Item("RHS"))
                    Sense = dtRow.Rows(i).Item("SENSE").ToString
                    CName = ""

                    If Sense.Equals("E") Then
                        xa.setRowFix(RName, Coef)
                        CName = "FIX"
                    ElseIf Sense.Equals("G") Then
                        xa.setRowMin(RName, Coef)
                        CName = "MIN"
                    ElseIf Sense.Equals("L") Then
                        xa.setRowMax(RName, Coef)
                        CName = "MAX"
                    Else
                        'Error
                    End If

                    xa.loadPoint(RName, CName, Coef)
                Next
                EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - XA - Loading Row took: ", GenUtils.FormatTime(startSolveTime, My.Computer.Clock.TickCount))

                startSolveTime = My.Computer.Clock.TickCount

                'RKP/06-11-10/v2.3.133
                If dtMtx Is Nothing Then
                    Try
                        'RKP/01-28-10/v2.3.128
                        'If qsysMtx is not present, resort to backup option.
                        dtMtx = _engine.CurrentDb.GetDataTable("SELECT * FROM qsysMtx") '_currentDb.GetDataTable("SELECT * FROM qsysMtx")            'SESSIONIZE
                        '_srcMtx = "qsysMtx"
                        _engine._srcMtx = "qsysMtx"
                    Catch ex As System.Exception
                        dtMtx = _engine.CurrentDb.GetDataTable("SELECT * FROM qsysMTXwithCOLS WHERE COEF <> 0 ORDER BY ColID, RowID")            'SESSIONIZE
                        '_srcMtx = "tsysMtx"
                        _engine._srcMtx = "tsysMtx"
                    End Try
                    'Me.SolutionNonZeros = dtMtx.Rows.Count
                    _engine.SolutionNonZeros = dtMtx.Rows.Count
                End If

                '
                ' Load Matrix Data ******************************************
                For i = 0 To dtMtx.Rows.Count - 1
                    RName = dtMtx.Rows(i).Item("ROW").ToString
                    CName = dtMtx.Rows(i).Item("COL").ToString
                    Coef = CDbl(dtMtx.Rows(i).Item("COEF"))

                    xa.loadPoint(RName, CName, Coef)
                Next
                EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - XA - Loading Mtx took: ", GenUtils.FormatTime(startSolveTime, My.Computer.Clock.TickCount))

                dtMtx = Nothing
                'If GenUtils.IsSwitchAvailable(_switches, "/UseMinSysRes") Then
                '    GC.Collect()
                'End If

                If GenUtils.IsSwitchAvailable(_switches, "/GenSolverMPS") Then
                    'xa.setCommand("FileName c:\XA.MPS ToMps Yes") 'Write the problem to an MPS file
                    filePath = GenUtils.GetSwitchArgument(_switches, "/GenSolverMPS", 1)
                    If filePath = "" Or Dir(filePath) = "" Then
                        filePath = GenUtils.GetWorkDir(_switches) & "\XA.MPS"
                    End If
                    Try
                        My.Computer.FileSystem.DeleteFile(filePath)
                    Catch ex As System.Exception

                    End Try
                    xa.setCommand("FileName " & filePath & " ToMps Yes") 'Write the problem to an MPS file
                    xa.setCommand("Set CmprsName _")
                End If

                'If GenUtils.IsSwitchAvailable(_switches, "/GenNativeMPS") Then
                '    MyBase.GenMPSFile(GenUtils.GetSwitchArgument(_switches, "/GenNativeMPS", 1))
                'End If



                startSolveTime = My.Computer.Clock.TickCount
                xa.solve()
                EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - XA - Solve took: ", GenUtils.FormatTime(startSolveTime, My.Computer.Clock.TickCount))

                'Try

                'Catch ex As System.Exception
                '    MsgBox(ex.Message)
                'End Try

                objFcnValue = xa.getLPObj()

                'If xa.getModelStatus() < 3 Then
                '    MsgBox(getSolutionStatusDescription(xa.getModelStatus()) & vbCrLf)
                'Else
                '    MsgBox(getSolutionStatusDescription(xa.getModelStatus()))
                'End If
                getSolutionStatusDescription(xa.getModelStatus())

                Debug.Print(" Solution Status: " & _solutionStatus)
                Debug.Print( _
                      " Number of Rows: " + Str(xa.getNumberOfRows()) + _
                      " Number of Columns: " + Str(xa.getNumberOfColumns()) + _
                      " Number of Non Zeros: " + Str(nonZeroCount) + _
                      " Optimal Obj: " + Str(xa.getLPObj()))

                _solutionRows = xa.getNumberOfRows()
                _solutionColumns = xa.getNumberOfColumns()
                _solutionNonZeros = nonZeroCount 'dtMtx.Rows.Count
                _solutionObj = xa.getLPObj()
                _solutionIterations = xa.getIterations()

                MyBase.Engine.SolutionStatus = _solutionStatus.Replace("=", "-").Trim().ToUpper() & " (" & xa.getModelStatus() & ")"

                MyBase.Engine.SolutionStatusCode = xa.getModelStatus()

                If MyBase.Engine.SolutionStatus.Contains("OPTIMAL") Then
                    MyBase.Engine.CommonSolutionStatus = "OPTIMAL SOLUTION"
                    MyBase.Engine.CommonSolutionStatusCode = Solver.commonSolutionStatus.statusOptimal
                Else
                    MyBase.Engine.CommonSolutionStatus = "INFEASIBLE SOLUTION"
                    MyBase.Engine.CommonSolutionStatusCode = Solver.commonSolutionStatus.statusInfeasible
                End If

                MyBase.Engine.SolutionRows = _solutionRows
                MyBase.Engine.SolutionColumns = _solutionColumns
                MyBase.Engine.SolutionNonZeros = _solutionNonZeros
                MyBase.Engine.SolutionObj = _solutionObj
                MyBase.Engine.SolverName = getSolverName().Replace("=", "-")
                MyBase.Engine.SolverVersion = "15.0.0.1" & " (" & " (" & IIf(GenUtils.Is64Bit, "64-bit", "32-bit") & ")"
                MyBase.Engine.SolutionIterations = _solutionIterations

                If xa.getModelStatus() < 3 Then

                    Debug.Print("C-OPT Engine - Solver - XA - Importing solution...")
                    Console.WriteLine("C-OPT Engine - Solver - XA - Importing solution...")
                    Console.WriteLine("")

                    ' Retrieve based upon index number
                    col = MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString
                    'For i = 0 To xa.getNumberOfColumns() - 1
                    '    sql = "UPDATE " & col & " SET ACTIVITY = " & xa.getColumnPrimalActivity(i) & ", DJ = " & xa.getColumnDualActivity(i) & " WHERE COL = '" & xa.getColumnName(i) & "'"
                    '    MyBase.Engine.CurrentDb.ExecuteNonQuery(sql)
                    'Next i
                    startSolveTime = My.Computer.Clock.TickCount
                    For i = 0 To dtCol.Rows.Count - 1
                        dtCol.Rows(i).Item("ACTIVITY") = xa.getColumnPrimalActivity(dtCol.Rows(i).Item("COL").ToString())
                        dtCol.Rows(i).Item("DJ") = xa.getColumnDualActivity(dtCol.Rows(i).Item("COL").ToString())
                    Next

                    'RKP/08-05-11/v3.0.149
                    If GenUtils.IsSwitchAvailable(_switches, "/UseMinSysRes") Then
                        If GenUtils.IsSwitchAvailable(_switches, "/UseMSAccessSyntax") Then
                            Try
                                MyBase.Engine.CurrentDb.UpdateDataSet(dtCol, "SELECT * FROM [" & _engine.MiscParams.Item("LINKED_COL_DB_PATH").ToString() & "].[" & MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString & "] ORDER BY ColID")
                            Catch ex As System.Exception
                                _progress = "Error using /UseMSAccessSyntax switch. Importing solution without switch."
                                Console.WriteLine(ex.Message)
                                Debug.Print(ex.Message)
                                Debug.Print("Error using /UseMSAccessSyntax switch. Importing solution without switch.")
                                Console.WriteLine("Error using /UseMSAccessSyntax switch. Importing solution without switch.")
                                EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-XA", "Error using /UseMSAccessSyntax switch. Importing solution without switch.")

                                MyBase.Engine.CurrentDb.UpdateDataSet(dtCol, "SELECT * FROM " & MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString & " ORDER BY ColID")
                            End Try

                        ElseIf GenUtils.IsSwitchAvailable(_switches, "/UseSQLServerSyntax") Then
                            updateSQLServer = True
                        Else
                            MyBase.Engine.CurrentDb.UpdateDataSet(dtCol, "SELECT * FROM " & MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString & " ORDER BY ColID")
                        End If
                    Else
                        If MyBase.Engine.CurrentDb.IsSQLExpress Then
                            updateSQLServer = True
                        Else
                            'MyBase.Engine.CurrentDb.UpdateDataSet(dtCol, "SELECT * FROM " & col & " ORDER BY ColID")
                            MyBase.Engine.CurrentDb.UpdateDataSet(dtCol, "SELECT * FROM " & MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString & " ORDER BY ColID")
                        End If
                    End If
                    If updateSQLServer Then
                        Try

                            'sql = "DELETE FROM " & MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString() & "Temp"
                            If noTruncate Then
                                'sql = "DELETE FROM " & MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString() & "Temp"
                                sql = "EXEC dbo.asp_TruncateTable '" & MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString() & "Temp" & "'"
                            Else
                                sql = "TRUNCATE TABLE " & MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString() & "Temp"
                            End If

                            Try
                                result = MyBase.Engine.CurrentDb.ExecuteNonQuery(sql, _switches, True)
                            Catch ex As System.Exception
                                _progress = "Error using /UseSQLServerSyntax switch. Importing solution without switch."
                                Console.WriteLine(ex.Message)
                                Debug.Print(ex.Message)
                                Debug.Print("Error using /UseSQLServerSyntax switch. Importing solution without switch.")
                                Console.WriteLine("Error using /UseSQLServerSyntax switch. Importing solution without switch.")
                                EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-XA", "Error using /UseSQLServerSyntax switch. Importing solution without switch.")
                            End Try

                            'GenUtils.UpdateDB(_engine.MiscParams.Item("LINKED_COL_DB_CONN_STR").ToString(), dtCol, "SELECT * FROM [" & MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString() & "] ORDER BY ColID")
                            'MyBase.Engine.CurrentDb.UpdateDataSet(dtCol, MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString() & "Temp", _switches)
                            MyBase.Engine.CurrentDb.UpdateDataSet(dtCol, MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString() & "Temp", _switches, True)

                            sql = "UPDATE " & MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString()
                            sql = sql & " SET "
                            sql = sql & MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString() & ".ACTIVITY = "
                            sql = sql & MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString() & "Temp.ACTIVITY,"
                            sql = sql & MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString() & ".DJ = "
                            sql = sql & MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString() & "Temp.DJ "
                            sql = sql & "FROM "
                            sql = sql & MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString() & " "
                            sql = sql & "INNER JOIN "
                            sql = sql & MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString() & "Temp "
                            sql = sql & "ON "
                            sql = sql & MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString() & ".ColID = "
                            sql = sql & MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString() & "Temp.ColID "
                            Try
                                result = MyBase.Engine.CurrentDb.ExecuteNonQuery(sql, _switches, True)
                            Catch ex As System.Exception
                                Debug.Print(ex.Message)
                                Debug.Print("C-OPT Engine - Solver - XA - Error importing Columns (tsysCol) - " & ex.Message & " - " & My.Computer.Clock.LocalTime.ToString)
                                Console.WriteLine("C-OPT Engine - Solver - XA - Error importing Columns (tsysCol) - " & ex.Message & " - " & My.Computer.Clock.LocalTime.ToString)
                                EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-XA", "C-OPT Engine - Solver - XA - Error importing Columns (tsysCol) - " & ex.Message & " - " & My.Computer.Clock.LocalTime.ToString)
                            End Try

                            'Empty the temp table, now that the purpose has been served.
                            'sql = "DELETE FROM " & MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString() & "Temp"
                            If noTruncate Then
                                'sql = "DELETE FROM " & MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString() & "Temp"
                                sql = "EXEC dbo.asp_TruncateTable '" & MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString() & "Temp" & "'"
                            Else
                                sql = "TRUNCATE TABLE " & MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString() & "Temp"
                            End If

                            Try
                                result = MyBase.Engine.CurrentDb.ExecuteNonQuery(sql, _switches, True)
                            Catch ex As System.Exception
                                _progress = "Error using /UseSQLServerSyntax switch. Importing solution without switch."
                                Console.WriteLine(ex.Message)
                                Debug.Print(ex.Message)
                                Debug.Print("Error using /UseSQLServerSyntax switch. Importing solution without switch.")
                                Console.WriteLine("Error using /UseSQLServerSyntax switch. Importing solution without switch.")
                                EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-XA", "Error using /UseSQLServerSyntax switch. Importing solution without switch.")
                            End Try
                            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-XA", "C-OPT Engine - Solver - XA - Successfully imported Columns. " & My.Computer.Clock.LocalTime.ToString)
                        Catch ex As System.Exception
                            _progress = "Error using /UseSQLServerSyntax switch. Importing solution without switch."
                            Console.WriteLine(ex.Message)
                            Debug.Print(ex.Message)
                            Debug.Print("Error using /UseSQLServerSyntax switch. Importing solution without switch.")
                            Console.WriteLine("Error using /UseSQLServerSyntax switch. Importing solution without switch.")
                            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-XA", "Error using /UseSQLServerSyntax switch. Importing solution without switch.")

                            MyBase.Engine.CurrentDb.UpdateDataSet(dtCol, "SELECT * FROM " & MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString & " ORDER BY ColID")
                        End Try
                    End If

                    EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - XA - Loading Solution into Col took: ", GenUtils.FormatTime(startSolveTime, My.Computer.Clock.TickCount))

                    row = MyBase.Engine.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString
                    'For i = 0 To xa.getNumberOfRows() - 1
                    '    sql = "UPDATE " & row & " SET ACTIVITY = " & xa.getRowPrimalActivity(i) & ", SHADOW = " & xa.getRowDualActivity(i) & " WHERE ROW = '" & xa.getRowName(i) & "'"
                    '    MyBase.Engine.CurrentDb.ExecuteNonQuery(sql)
                    'Next i
                    startSolveTime = My.Computer.Clock.TickCount
                    For i = 0 To dtRow.Rows.Count - 1
                        dtRow.Rows(i).Item("ACTIVITY") = xa.getRowPrimalActivity(dtRow.Rows(i).Item("ROW").ToString())
                        dtRow.Rows(i).Item("SHADOW") = xa.getRowDualActivity(dtRow.Rows(i).Item("ROW").ToString())
                    Next

                    'MyBase.Engine.CurrentDb.UpdateDataSet(dtRow, "SELECT * FROM " & row)
                    'RKP/08-05-11/v3.0.149
                    If GenUtils.IsSwitchAvailable(_switches, "/UseMinSysRes") Then
                        If GenUtils.IsSwitchAvailable(_switches, "/UseMSAccessSyntax") Then
                            Try
                                MyBase.Engine.CurrentDb.UpdateDataSet(dtRow, "SELECT * FROM [" & _engine.MiscParams.Item("LINKED_ROW_DB_PATH").ToString() & "].[" & MyBase.Engine.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString & "] ORDER BY RowID")
                            Catch ex As System.Exception
                                _progress = "Error using /UseMSAccessSyntax switch. Importing solution without switch."
                                Console.WriteLine(ex.Message)
                                Debug.Print(ex.Message)
                                Debug.Print("Error using /UseMSAccessSyntax switch. Importing solution without switch.")
                                Console.WriteLine("Error using /UseMSAccessSyntax switch. Importing solution without switch.")
                                EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-XA", "Error using /UseMSAccessSyntax switch. Importing solution without switch.")

                                MyBase.Engine.CurrentDb.UpdateDataSet(dtRow, "SELECT * FROM " & MyBase.Engine.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString & " ORDER BY RowID")
                            End Try
                        ElseIf GenUtils.IsSwitchAvailable(_switches, "/UseSQLServerSyntax") Then
                            updateSQLServer = True
                        Else
                            MyBase.Engine.CurrentDb.UpdateDataSet(dtRow, "SELECT * FROM " & MyBase.Engine.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString & " ORDER BY RowID")
                        End If
                    Else
                        If MyBase.Engine.CurrentDb.IsSQLExpress Then
                            updateSQLServer = True
                        Else
                            'MyBase.Engine.CurrentDb.UpdateDataSet(dtCol, "SELECT * FROM " & col & " ORDER BY ColID")
                            MyBase.Engine.CurrentDb.UpdateDataSet(dtRow, "SELECT * FROM " & MyBase.Engine.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString & " ORDER BY RowID")
                        End If
                    End If
                    If updateSQLServer Then
                        Try
                            If noTruncate Then
                                'sql = "DELETE FROM " & MyBase.Engine.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString() & "Temp"
                                sql = "EXEC dbo.asp_TruncateTable '" & MyBase.Engine.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString() & "Temp" & "'"
                            Else
                                sql = "TRUNCATE TABLE " & MyBase.Engine.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString() & "Temp"
                            End If

                            Try
                                result = MyBase.Engine.CurrentDb.ExecuteNonQuery(sql, _switches, linkedTable)
                            Catch ex As System.Exception
                                Debug.Print(ex.Message)
                                Debug.Print("C-OPT Engine - Solver - XA - Error importing Constraints (tsysRow) - " & ex.Message & " - " & My.Computer.Clock.LocalTime.ToString)
                                Console.WriteLine("C-OPT Engine - Solver - XA - Error importing Constraints (tsysRow) - " & ex.Message & " - " & My.Computer.Clock.LocalTime.ToString)
                                EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-XA", "C-OPT Engine - Solver - XA - Error importing Constraints (tsysRow) - " & ex.Message & " - " & My.Computer.Clock.LocalTime.ToString)
                            End Try

                            'MyBase.Engine.CurrentDb.UpdateDataSet(dtRow, "SELECT * FROM [" & _engine.MiscParams.Item("LINKED_ROW_DB_PATH").ToString() & "].[" & MyBase.Engine.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString & "] ORDER BY RowID")
                            'MyBase.Engine.CurrentDb.UpdateDataSet(dtRow, MyBase.Engine.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString() & "Temp", _switches)
                            MyBase.Engine.CurrentDb.UpdateDataSet(dtRow, MyBase.Engine.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString() & "Temp", _switches, True)

                            sql = "UPDATE " & MyBase.Engine.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString()
                            sql = sql & " SET "
                            sql = sql & MyBase.Engine.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString() & ".ACTIVITY = "
                            sql = sql & MyBase.Engine.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString() & "Temp.ACTIVITY,"
                            sql = sql & MyBase.Engine.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString() & ".SHADOW = "
                            sql = sql & MyBase.Engine.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString() & "Temp.SHADOW "
                            sql = sql & "FROM "
                            sql = sql & MyBase.Engine.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString() & " "
                            sql = sql & "INNER JOIN "
                            sql = sql & MyBase.Engine.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString() & "Temp "
                            sql = sql & "ON "
                            sql = sql & MyBase.Engine.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString() & ".RowID = "
                            sql = sql & MyBase.Engine.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString() & "Temp.RowID "
                            Try
                                result = MyBase.Engine.CurrentDb.ExecuteNonQuery(sql, _switches, True)
                            Catch ex As System.Exception
                                Debug.Print(ex.Message)
                                Debug.Print("C-OPT Engine - Solver - XA - Error importing Constraints (tsysRow) - " & ex.Message & " - " & My.Computer.Clock.LocalTime.ToString)
                                Console.WriteLine("C-OPT Engine - Solver - XA - Error importing Constraints (tsysRow) - " & ex.Message & " - " & My.Computer.Clock.LocalTime.ToString)
                                EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-XA", "C-OPT Engine - Solver - XA - Error importing Constraints (tsysRow) - " & ex.Message & " - " & My.Computer.Clock.LocalTime.ToString)
                            End Try

                            'Empty temp table now that it's purpose has been served.
                            'sql = "DELETE FROM " & MyBase.Engine.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString() & "Temp"
                            If noTruncate Then
                                'sql = "DELETE FROM " & MyBase.Engine.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString() & "Temp"
                                sql = "EXEC dbo.asp_TruncateTable '" & MyBase.Engine.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString() & "Temp" & "'"
                            Else
                                sql = "TRUNCATE TABLE " & MyBase.Engine.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString() & "Temp"
                            End If

                            Try
                                result = MyBase.Engine.CurrentDb.ExecuteNonQuery(sql, _switches, True)
                            Catch ex As System.Exception
                                Debug.Print(ex.Message)
                                Debug.Print("C-OPT Engine - Solver - XA - Error importing Constraints (tsysRow) - " & ex.Message & " - " & My.Computer.Clock.LocalTime.ToString)
                                Console.WriteLine("C-OPT Engine - Solver - XA - Error importing Constraints (tsysRow) - " & ex.Message & " - " & My.Computer.Clock.LocalTime.ToString)
                                EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-XA", "C-OPT Engine - Solver - XA - Error importing Constraints (tsysRow) - " & ex.Message & " - " & My.Computer.Clock.LocalTime.ToString)
                            End Try
                            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-XA", "C-OPT Engine - Solver - XA - Successfully imported Rows. " & My.Computer.Clock.LocalTime.ToString)
                        Catch ex As System.Exception
                            _progress = "Error using /UseMSAccessSyntax switch. Importing solution without switch."
                            Console.WriteLine(ex.Message)
                            Debug.Print(ex.Message)
                            Debug.Print("Error using /UseMSAccessSyntax switch. Importing solution without switch.")
                            Console.WriteLine("Error using /UseMSAccessSyntax switch. Importing solution without switch.")
                            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-XA", "Error using /UseMSAccessSyntax switch. Importing solution without switch.")

                            MyBase.Engine.CurrentDb.UpdateDataSet(dtRow, "SELECT * FROM " & MyBase.Engine.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString & " ORDER BY RowID")
                        End Try
                    End If

                    EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - XA - Loading Solution into Row took: ", GenUtils.FormatTime(startSolveTime, My.Computer.Clock.TickCount))
                End If

                xa.closeConnection()

                Console.WriteLine("C-OPT Engine - Solver - XA - Importing solution...done.")
                Console.WriteLine("")
                success = True
            Catch xe As Exception
                'MsgBox(xa.getXAExceptionMessage() & vbNewLine & xe.Message)
                If errSource.ToUpper() = "XA.OPENCONNECTION" Then
                    GenUtils.Message(GenUtils.MsgType.Critical, "Solver - XA", xa.getXAExceptionMessage() & vbNewLine & xe.Message & vbNewLine & vbNewLine & "The most likely cause for this error could be that you may have not attached the XA USB key to your computer." & vbNewLine & vbNewLine & "- Please attach the XA USB key to your computer and try the operation again." & vbNewLine & "- If you do not have the XA USB key, then select ""CoinMP"" as the solver and try the operation again.")
                Else
                    GenUtils.Message(GenUtils.MsgType.Critical, "Solver - XA", xa.getXAExceptionMessage() & vbNewLine & xe.Message)
                End If

                getSolutionStatusDescription(xa.getModelStatus())
                _solutionStatus = xa.getXAExceptionMessage().Replace("=", "-")
                _solutionRows = xa.getNumberOfRows()
                _solutionColumns = xa.getNumberOfColumns()
                _solutionNonZeros = nonZeroCount 'dtMtx.Rows.Count
                _solutionObj = xa.getLPObj()

                xa.closeConnection()

                MyBase.Engine.SolutionStatus = _solutionStatus
                MyBase.Engine.SolutionRows = _solutionRows
                MyBase.Engine.SolutionColumns = _solutionColumns
                MyBase.Engine.SolutionNonZeros = _solutionNonZeros
                MyBase.Engine.SolutionObj = _solutionObj
                MyBase.Engine.SolverName = getSolverName()
                MyBase.Engine.SolutionIterations = _solutionIterations
                success = False
            End Try
            EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - XA - Total solve took: ", GenUtils.FormatTime(startTime, My.Computer.Clock.TickCount))
            Return success
        End Function

        Public Overrides Function getSolutionStatusDescription(ByVal sts As Integer) As String
            Try
                Dim descrArray(12) As String
                descrArray(1) = "Optimal Solution"
                descrArray(2) = "Integer Solution (not proven the optimal integer solution)"
                descrArray(3) = "Unbounded Solution"
                descrArray(4) = "Infeasible Solution"
                descrArray(5) = "Callback function indicates Infeasible Solution"
                descrArray(6) = "Intermediate Infeasible Solution"
                descrArray(7) = "Intermediate Non-optimal Solution"
                descrArray(8) = ""
                descrArray(9) = "Intermediate Non-integer Solution"
                descrArray(10) = "Integer Infeasible"
                descrArray(11) = ""
                descrArray(12) = "Error Unknown"
                _solutionStatus = descrArray(sts)
                Return descrArray(sts)
            Catch ex As System.Exception
                Return ""
            End Try
        End Function

        Public Overrides Function Solve() As Boolean

        End Function

        Public Overrides Function getSolutionStatusDescription() As String
            Return _solutionStatus
        End Function

        Public Overrides Function getSolverName() As String
            Return "XA"
        End Function

        Public Function getSolutionRows() As Integer
            Return _solutionRows
        End Function

        Public Function getSolutionColumns() As Integer
            Return _solutionColumns
        End Function

        Public Function getSolutionObj() As Double
            Return _solutionObj
        End Function

        Public Function getSolutionIterations() As Double
            Return _solutionIterations
        End Function

        Private usbKeyAvailable As Boolean = False
        Public ReadOnly Property IsUSBKeyAvailable() As Boolean
            Get
                Return usbKeyAvailable
            End Get
        End Property
    End Class
End Namespace