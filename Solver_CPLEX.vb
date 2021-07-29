Imports System.Runtime.InteropServices
Imports System.Text
Imports System
Imports System.Linq
Imports System.Linq.Expressions
Imports System.Collections
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Data.Common
Imports System.Globalization
Imports System.Xml
Imports System.Xml.Linq
Imports System.Threading
Imports Microsoft.VisualBasic
Imports EntLib.COPT
Imports ILOG.Concert
Imports ILOG.CPLEX

Namespace COPT
    <Microsoft.VisualBasic.ComClass()>
    Public Class Solver_CPLEX
        Inherits Solver

        Private _solutionStatus As String
        Private _solutionRows As Integer
        Private _solutionColumns As Integer
        Private _solutionObj As Double
        Private Shadows _switches() As String
        Private Shadows _engine As COPT.Engine
        'Private _workDir As String = GenUtils.GetWorkDir(_switches)
        'Private _ds As New DataSet
        Private _probTypeMIP As Boolean = False
        Private _progress As String = ""

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
            'RKP/12-03-09
            Dim startTime As Integer = My.Computer.Clock.TickCount
            Dim success As Boolean = False 'RKP/04-18-12/v3.2.166

            '_ds.Tables.Add(dtCol)
            '_ds.Tables(0).TableName = "tsysCol"

            '_ds.Tables.Add(dtRow)
            '_ds.Tables(1).TableName = "tsysRow"

            '_ds.Tables.Add(dtMtx)
            '_ds.Tables(2).TableName = "tsysMtx"

            'EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "App - Status", "")
            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - Callable Library")
            success = SolveUsingCallableLibrary(dtRow, dtCol, dtMtx)
            'SolveUsingConcert()

            EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - CPLEX - Total solve took: ", GenUtils.FormatTime(startTime, My.Computer.Clock.TickCount))
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
            Return "CPLEX"
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

        ''' <summary>
        ''' CPLEX solver using API calls.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' RKP/01-19-10/v2.3.126
        ''' This is the main "solve" function.
        ''' CPLEX.Solve
        ''' </remarks>
        Private Function SolveUsingCallableLibrary(ByRef dtRow As DataTable, ByRef dtCol As DataTable, ByRef dtMtx As DataTable) As Boolean
            Dim numCols As Long = dtCol.Rows.Count  '_ds.Tables("tsysCol").Rows.Count
            Dim numRows As Long = dtRow.Rows.Count  '_ds.Tables("tsysRow").Rows.Count
            Dim numNZ As Long = _engine.SolutionNonZeros  'dtMtx.Rows.Count  '_ds.Tables("tsysMtx").Rows.Count
            Dim dobj(numCols) As Double
            Dim dclo(numCols) As Double
            Dim dcup(numCols) As Double
            Dim tickCountStart As Long = My.Computer.Clock.TickCount
            Dim status As Long = -1 'RKP/09-28-12/v4.2.180.2 'Integer = -1
            Dim statusSolution As Integer = -1 '0=Success, Non-Zero=Failure
            Dim env As IntPtr = IntPtr.Zero
            Dim lp As IntPtr = IntPtr.Zero
            Dim ret As Long
            Dim statind As Long
            Dim buffer As StringBuilder = New StringBuilder(510)
            Dim runName As String = ""
            Dim i As Integer = 0
            Dim sql As String
            Dim result As Long
            Dim outputFile As String
            Dim updateSQLServerRow As Boolean = False 'RKP/08-04-11/v3.0.149
            Dim updateSQLServerCol As Boolean = False 'RKP/08-04-11/v3.0.149
            Dim success As Boolean = False 'RKP/04-18-12/v3.2.166
            Dim noTruncate As Boolean = GenUtils.IsSwitchAvailable(_switches, "/UseTruncateSP")
            Dim switchOptionGap As Boolean = GenUtils.IsSwitchAvailable(_switches, "/SetRelativeGapTolerance")
            Dim switchOptionTime As Boolean = GenUtils.IsSwitchAvailable(_switches, "/SetTimeLimit")
            Dim mipSolutionCount As Long = 0
            Dim timeOptTotalStart As Long = My.Computer.Clock.TickCount
            Dim timeOptContinuous As Long = 0
            Dim timeOptMIPStage1 As Long = 0
            Dim timeOptMIPStage2 As Long = 0
            Dim timeOptMIPContinuous As Long = 0
            Dim timeOptMIPTotal As Long = 0
            Dim solutionStatusMIPStage1 As String = "" 'Solution Status after Stage 1 of MIP
            Dim solutionStatusMIPStage2 As String = "" 'Solution Status after Stage 2 of MIP (Continuous run, in order to get DJ and SHADOW)

            Try
                runName = _engine.MiscParams.Item("RUN_NAME").ToString()
            Catch ex As System.Exception
                runName = "RUN_NAME"
            End Try

            outputfile = GenUtils.GetWorkDir(_switches) & "\" & GenUtils.GetSwitchArgument(_switches, "/PRJ", 1) & "." & runName & ".CPLEX." & _engine.TimeStamp()

            'Initialize the CPLEX environment
            Try
                env = Wrapper_CPLEX.CPXopenCPLEX(status)
            Catch ex As System.Exception
                GenUtils.Message(GenUtils.MsgType.Critical, "CPLEX-SolveUsingCallableLibrary", "Failed at: CPXopenCPLEX" & vbNewLine & ex.Message & vbNewLine & "Make sure you are using the correct bitness of the solver (32-bit vs 64-bit)" & vbNewLine & "APM: " & GenUtils.GetAvailablePhysicalMemoryStr)
                success = False
            End Try

            If env.Equals(IntPtr.Zero) Then
                Dim errmsg As StringBuilder = New StringBuilder(1024)
                System.Console.Error.WriteLine("Could not open CPLEX environment.")
                CPXgeterrorstring(env, status, errmsg)
                System.Console.Error.WriteLine(errmsg)
                EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - CPXopenCPLEX failed. " & errmsg.ToString())
                GenUtils.Message(GenUtils.MsgType.Critical, "Solver - CPLEX", errmsg.ToString())
                success = False
                GoTo TERMINATE
            Else
                Console.WriteLine("CPLEX Solver opened successfully!")
                EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - Opened successfully.")
                success = True
            End If

            'Turn on output to the screen
            status = CPXsetintparam(env, CPX_PARAM_SCRIND, CPX_ON)
            If status <> 0 Then
                System.Console.Error.WriteLine("Failure to turn on screen indicator, error {0}.", status)
                EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - CPXsetintparam screen indicator failed.")
                success = False
                GoTo TERMINATE
            End If
            success = True

            'Turn on data checking
            status = CPXsetintparam(env, CPX_PARAM_DATACHECK, CPX_ON)
            If status <> 0 Then
                System.Console.Error.WriteLine("Failure to turn on data checking, error {0}.", status)
                EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - CPXsetintparam data check failed.")
                success = False
                GoTo TERMINATE
            End If
            success = True

            'RKP/06-04-12/v3.2.170
            'Set "Relative MIP Gap Tolerance" for MIP problems.
            If MyBase.isMIP Then

                'RKP/06-04-12/v3.2.170
                'Set "Tree Memory Limit" for MIP problems.
                If GenUtils.IsSwitchAvailable(_switches, "/SetTreeMemLimit") Then
                    status = CPXsetdblparam(env, CPX_PARAM_TRELIM, CDbl(GenUtils.GetSwitchArgument(_switches, "/SetTreeMemLimit", 1)))
                    If status <> 0 Then
                        System.Console.Error.WriteLine("Failure to /SetTreeMemLimit, error {0}.", status)
                        EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - CPXsetdblparam(CPX_PARAM_TRELIM), ""/SetTreeMemLimit"", data check failed.")
                        success = False
                        GoTo TERMINATE
                    End If
                    success = True
                End If

                'RKP/06-05-12/v3.2.170
                'Set "Solution Limit" for MIP problems.
                If GenUtils.IsSwitchAvailable(_switches, "/SetSolutionLimit") Then
                    status = CPXsetintparam(env, CPX_PARAM_INTSOLLIM, CInt(GenUtils.GetSwitchArgument(_switches, "/SetSolutionLimit", 1)))
                    If status <> 0 Then
                        System.Console.Error.WriteLine("Failure to /SetSolutionLimit, error {0}.", status)
                        EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - CPXsetintparam(CPX_PARAM_INTSOLLIM), ""/SetSolutionLimit"", data check failed.")
                        success = False
                        GoTo TERMINATE
                    End If
                    success = True
                End If
            End If

            'RKP/06-05-12/v3.2.170
            'Set "Maximum Time Limit" (in seconds) for ALL problems. and gap tolerance if you do not have time limits
            If switchOptionTime Then
                status = CPXsetdblparam(env, CPX_PARAM_TILIM, CDbl(GenUtils.GetSwitchArgument(_switches, "/SetTimeLimit", 1)))  ' this sets min time as max
                If status <> 0 Then
                    System.Console.Error.WriteLine("Failure to /SetTimeLimit, error {0}.", status)
                    EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - CPXsetdblparam(CPX_PARAM_TILIM), ""/SetTimeLimit"", data check failed.")
                    success = False
                    GoTo TERMINATE
                End If
                success = True
                ''BLOCK COMMENTED TO LEAVE MIPGAPTOLERANCE AT CPLEX DEFAULT
                'If GenUtils.IsSwitchAvailable(_switches, "/SetRelativeGapTolerance") Then
                '    status = CPXsetdblparam(env, CPX_PARAM_EPGAP, 0.000000001)  'reset gap to small number so mipopt will run to the minimum time (aas cplex maxtime)
                '    If status <> 0 Then
                '        System.Console.Error.WriteLine("Failure to /SetRelativeGapTolerance, error {0}.", status)
                '        EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - CPXsetdblparam(CPX_PARAM_EPGAP), ""/SetRelativeGapTolerance"", data check failed.")
                '        success = False
                '        GoTo TERMINATE
                '    End If
                '    success = True
                'End If
            Else  ' if there are no time limits go ahead and set gaptolerance to user switch value if it exists
                If switchOptionGap Then
                    status = CPXsetdblparam(env, CPX_PARAM_EPGAP, CDbl(GenUtils.GetSwitchArgument(_switches, "/SetRelativeGapTolerance", 1)))
                    If status <> 0 Then
                        System.Console.Error.WriteLine("Failure to /SetRelativeGapTolerance, error {0}.", status)
                        EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - CPXsetdblparam(CPX_PARAM_EPGAP), ""/SetRelativeGapTolerance"", data check failed.")
                        success = False
                        GoTo TERMINATE
                    End If
                    success = True
                End If
            End If

            'Create the problem
            Try
                lp = CPXcreateprob(env, status, GenUtils.GetSwitchArgument(_switches, "/PRJ", 1))
            Catch ex As System.Exception
                GenUtils.Message(GenUtils.MsgType.Critical, "CPLEX-SolveUsingCallableLibrary", "Failed at: CPXcreateprob" & vbNewLine & ex.Message & vbNewLine & "APM: " & GenUtils.GetAvailablePhysicalMemoryStr)
                success = False
            End Try

            ' A returned pointer of NULL may mean that not enough memory
            ' was available or there was some other problem.  In the case of 
            ' failure, an error message will have been written to the error 
            ' channel from inside CPLEX.  In this example, the setting of
            ' the parameter CPX_PARAM_SCRIND causes the error message to
            ' appear on stdout. 

            If lp.Equals(IntPtr.Zero) Then
                System.Console.Error.WriteLine("Failed to create LP.")
                EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - Failed to create LP.")
                success = False
                GoTo TERMINATE
            End If
            success = True

            'status = PopulateByRow(env, lp)
            'MyBase.LoadSolverArrays(dtRow, dtCol, dtMtx, True)
            status = PopulateByColumn(env, lp, dtRow, dtCol, dtMtx, numNZ, MyBase.isMIP)
            If status Then
                System.Console.Error.WriteLine("Failed to populate problem.")
                EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - Failed to populate problem.")
                success = False
                GoTo TERMINATE
            End If
            success = True

            GenUtils.CollectGarbage()

            'Optimize the problem and obtain solution.
            If MyBase.isMIP Then 'If _probTypeMIP Then
                Console.WriteLine("Solver - CPLEX - Problem type = MIP")

                timeOptMIPStage1 = My.Computer.Clock.TickCount
                Try
                    status = CPXmipopt(env, lp)
                Catch ex As System.Exception
                    GenUtils.Message(GenUtils.MsgType.Critical, "CPLEX-SolveUsingCallableLibrary", "Failed at: CPXmipopt" & vbNewLine & ex.Message & vbNewLine & "APM: " & GenUtils.GetAvailablePhysicalMemoryStr)
                End Try

                EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - Solver - CPLEX (MIP Stage 1) solve time took: ", Space(9) & GenUtils.FormatTime(timeOptMIPStage1, My.Computer.Clock.TickCount))
                Console.WriteLine("Solver - CPLEX - MIP Stage 1 solve time took: " & GenUtils.FormatTime(timeOptMIPStage1, My.Computer.Clock.TickCount))
                _progress = "Solver - CPLEX - MIP Stage 1 solve time took: " & GenUtils.FormatTime(timeOptMIPStage1, My.Computer.Clock.TickCount)
                Console.WriteLine("Solver - CPLEX - MIP Stage 1 solve status: " & statind.ToString())
                _progress = "Solver - CPLEX - MIP Stage 1 solve status: " & statind.ToString()


                EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - Problem type = MIP")

                If status = 0 Then
                    Try
                        mipSolutionCount = CPXgetsolnpoolnumsolns(env, lp)
                    Catch ex As System.Exception

                    End Try
                End If
            Else
                Console.WriteLine("Solver - CPLEX - Problem type = Continuous")

                timeOptContinuous = My.Computer.Clock.TickCount
                Try
                    status = CPXlpopt(env, lp)
                Catch ex As System.Exception
                    GenUtils.Message(GenUtils.MsgType.Critical, "CPLEX-SolveUsingCallableLibrary", "Failed at: CPXlpopt" & vbNewLine & ex.Message & vbNewLine & "APM: " & GenUtils.GetAvailablePhysicalMemoryStr)
                End Try

                EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - Solver - CPLEX (Continuous) solve time took: ", Space(9) & GenUtils.FormatTime(timeOptContinuous, My.Computer.Clock.TickCount))
                Console.WriteLine("Solver - CPLEX - Continuous solve time took: " & GenUtils.FormatTime(timeOptMIPStage1, My.Computer.Clock.TickCount))
                _progress = "Solver - CPLEX - Continuous solve time took: " & GenUtils.FormatTime(timeOptMIPStage1, My.Computer.Clock.TickCount)

                If status Then
                    GenUtils.Message(GenUtils.MsgType.Critical, "CPLEX-SolveUsingCallableLibrary", "Failed at: CPXlpopt" & vbNewLine & "" & vbNewLine & "APM: " & GenUtils.GetAvailablePhysicalMemoryStr)
                End If


                EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - Problem type = Continuous")
            End If

            statusSolution = status

            GenUtils.CollectGarbage()

            If (status) Then
                System.Console.Error.WriteLine("Failed to optimize LP/MIP.")
                EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - Failed to optimize LP/MIP.")
                success = False
                GoTo TERMINATE
            Else
                Console.WriteLine("Solver - CPLEX - Solved successfully!")
                EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "CPLEX - Solved successfully!" & vbNewLine & statusSolution)
                success = True
            End If

            Try
                statind = CPXgetstat(env, lp)
                success = True
            Catch ex As System.Exception
                EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - CPXgetstat - " & ex.Message)
                'MessageBox.Show(ex.Message)
                GenUtils.Message(GenUtils.MsgType.Warning, "Solver - CPLEX", ex.Message)
                success = False
            End Try

            Try
                ret = CPXgetstatstring(env, statind, buffer)
            Catch ex As System.Exception
                EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - CPXgetstatstring - " & ex.Message)
                'MessageBox.Show(ex.Message)
                GenUtils.Message(GenUtils.MsgType.Warning, "Solver - CPLEX", ex.Message)
            End Try

            If MyBase.isMIP Then

                Try
                    mipSolutionCount = CPXgetsolnpoolnumsolns(env, lp)
                Catch ex As System.Exception

                End Try

                solutionStatusMIPStage1 = buffer.ToString() & " (" & statind & ")"

                _progress = "C-OPT Engine - Solver - CPLEX (MIP Stage 1) solve status: " & buffer.ToString() & " (" & statind & ")"
                EntLib.COPT.Log.Log(_workDir, _progress)
                Console.WriteLine(_progress)
                _progress = "C-OPT Engine - Solver - CPLEX (MIP Stage 1) solve time took: " & GenUtils.FormatTime(timeOptMIPStage1, My.Computer.Clock.TickCount)
                EntLib.COPT.Log.Log(_workDir, _progress)
                Console.WriteLine(_progress)
                _progress = "C-OPT Engine - Solver - CPLEX - MIP Stage 1 Solve Status: " & statind.ToString() & " - " & buffer.ToString()
                EntLib.COPT.Log.Log(_workDir, _progress)
                Console.WriteLine(_progress)
                _progress = "C-OPT Engine - Solver - CPLEX - MIP Stage 1 Solution Count: " & mipSolutionCount.ToString()
                EntLib.COPT.Log.Log(_workDir, _progress)
                Console.WriteLine(_progress)
            Else
                _progress = "C-OPT Engine - Solver - CPLEX (Continuous) solve status: " & buffer.ToString() & " (" & statind & ")"
                EntLib.COPT.Log.Log(_workDir, _progress)
                Console.WriteLine(_progress)
                _progress = "C-OPT Engine - Solver - CPLEX (Continuous) solve time took: " & GenUtils.FormatTime(timeOptContinuous, My.Computer.Clock.TickCount)
                EntLib.COPT.Log.Log(_workDir, _progress)
                Console.WriteLine(_progress)
                _progress = "C-OPT Engine - Solver - CPLEX - Continuous Solve Status: " & statind.ToString() & " - " & buffer.ToString()
                EntLib.COPT.Log.Log(_workDir, _progress)
                Console.WriteLine(_progress)
            End If

            '*&* Stage 2 for MIP

            If MyBase.isMIP Then 'If _probTypeMIP Then
                If statind <> 101 AndAlso statind <> 102 Then
                    If switchOptionGap AndAlso switchOptionTime Then
                        'set the time to diff between timemin and timemax
                        status = CPXsetdblparam(env, CPX_PARAM_TILIM, CDbl(GenUtils.GetSwitchArgument(_switches, "/SetTimeLimit", 2)) - CDbl(GenUtils.GetSwitchArgument(_switches, "/SetTimeLimit", 1)))  ' this sets min time as max
                        If status <> 0 Then
                            System.Console.Error.WriteLine("Failure to /SetTimeLimit, error {0}.", status)
                            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - CPXsetdblparam(CPX_PARAM_TILIM), ""/SetTimeLimit"", data check failed.")
                            success = False
                            GoTo TERMINATE
                        End If
                        success = True

                        status = CPXsetdblparam(env, CPX_PARAM_EPGAP, CDbl(GenUtils.GetSwitchArgument(_switches, "/SetRelativeGapTolerance", 1)))
                        If status <> 0 Then
                            System.Console.Error.WriteLine("Failure to /SetRelativeGapTolerance, error {0}.", status)
                            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - CPXsetdblparam(CPX_PARAM_EPGAP), ""/SetRelativeGapTolerance"", data check failed.")
                            success = False
                            GoTo TERMINATE
                        End If
                        success = True

                        timeOptMIPStage2 = My.Computer.Clock.TickCount

                        'stage two optimize
                        Try
                            status = CPXmipopt(env, lp)
                        Catch ex As System.Exception
                            GenUtils.Message(GenUtils.MsgType.Critical, "CPLEX-SolveUsingCallableLibrary", "Failed at: CPXmipopt stage 2" & vbNewLine & ex.Message & vbNewLine & "APM: " & GenUtils.GetAvailablePhysicalMemoryStr)
                        End Try

                        Try
                            statind = CPXgetstat(env, lp)
                            success = True
                        Catch ex As System.Exception
                            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - CPXgetstat - " & ex.Message)
                            'MessageBox.Show(ex.Message)
                            GenUtils.Message(GenUtils.MsgType.Warning, "Solver - CPLEX", ex.Message)
                            success = False
                        End Try

                        Try
                            ret = CPXgetstatstring(env, statind, buffer)
                        Catch ex As System.Exception
                            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - CPXgetstatstring - " & ex.Message)
                            'MessageBox.Show(ex.Message)
                            GenUtils.Message(GenUtils.MsgType.Warning, "Solver - CPLEX", ex.Message)
                        End Try

                        Try
                            mipSolutionCount = CPXgetsolnpoolnumsolns(env, lp)
                        Catch ex As System.Exception

                        End Try

                        _progress = "C-OPT Engine - Solver - CPLEX - MIP Stage 1.2 (time limit + relative gap tolerance) Solve Time: " & GenUtils.FormatTime(timeOptMIPStage2, My.Computer.Clock.TickCount)
                        EntLib.COPT.Log.Log(_workDir, _progress)
                        Console.WriteLine(_progress)
                        _progress = "C-OPT Engine - Solver - CPLEX - MIP Stage 1.2 Solve Status: " & statind.ToString() & " - " & buffer.ToString()
                        EntLib.COPT.Log.Log(_workDir, _progress)
                        Console.WriteLine(_progress)
                        _progress = "C-OPT Engine - Solver - CPLEX - MIP Stage 1.2 Solution Count: " & mipSolutionCount.ToString()
                        EntLib.COPT.Log.Log(_workDir, _progress)
                        Console.WriteLine(_progress)
                    End If
                End If
            End If
            statusSolution = status

            'If _probTypeMIP Then
            '    'status = CPXpopulate(env, lp)
            'End If

            ' The size of the problem should be obtained by asking CPLEX what
            '  the actual size is, rather than using sizes from when the problem
            '  was built.  cur_numrows and cur_numcols store the current number 
            '  of rows and columns, respectively. 

            Dim cur_numrows As Integer = CPXgetnumrows(env, lp)
            Dim cur_numcols As Integer = CPXgetnumcols(env, lp)

            Dim x(cur_numcols) As Double
            Dim slack(cur_numrows) As Double
            Dim dj(cur_numcols) As Double
            Dim pi(cur_numrows) As Double
            Dim dobjval(cur_numcols) As Double

            Dim solstat As Integer = 0
            Dim objval As Double = 0.0

            Try
                status = CPXsolution(env, lp, solstat, objval, x, pi, slack, dj)
                success = True
            Catch ex As System.Exception
                EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - CPXsolution - " & ex.Message)
                'MessageBox.Show(ex.Message)
                GenUtils.Message(GenUtils.MsgType.Warning, "Solver - CPLEX", ex.Message)
                success = False
            End Try

            Try
                statind = CPXgetstat(env, lp)
                success = True
            Catch ex As System.Exception
                EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - CPXgetstat - " & ex.Message)
                'MessageBox.Show(ex.Message)
                GenUtils.Message(GenUtils.MsgType.Warning, "Solver - CPLEX", ex.Message)
                success = False
            End Try

            Try
                ret = CPXgetstatstring(env, statind, buffer)
                success = True
            Catch ex As System.Exception
                EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - CPXgetstatstring - " & ex.Message)
                'MessageBox.Show(ex.Message)
                GenUtils.Message(GenUtils.MsgType.Warning, "Solver - CPLEX", ex.Message)
                success = False
            End Try

            'Select Case solstat
            '    Case 0
            '        _engine.SolutionStatus = "Infeasible Solution"
            '    Case 1
            '        _engine.SolutionStatus = "Optimal Solution"
            '    Case Else
            '        _engine.SolutionStatus = ""
            'End Select

            If MyBase.isMIP Then 'If _probTypeMIP Then
                status = CPXgetobjval(env, lp, dobjval)
                objval = dobjval(0)
                status = CPXgetx(env, lp, x, 0, CPXgetnumcols(env, lp) - 1)
                status = CPXgetdj(env, lp, dj, 0, CPXgetnumcols(env, lp) - 1)
                status = CPXgetslack(env, lp, slack, 0, CPXgetnumrows(env, lp) - 1)
                status = CPXgetpi(env, lp, pi, 0, CPXgetnumrows(env, lp) - 1)
                Try
                    mipSolutionCount = CPXgetsolnpoolnumsolns(env, lp)
                Catch ex As System.Exception

                End Try
            End If

            '_engine.SolutionStatus = solstat & " (" & buffer.ToString() & ")" '"1=Optimal Solution"
            '_engine.SolutionStatus = statind & " (" & buffer.ToString() & ")" '"1=Optimal Solution"
            _engine.SolutionStatus = buffer.ToString().Replace("=", "-").Trim().ToUpper() & " (" & statind.ToString() & ")" '"1=Optimal Solution"

            'Continuous - OPTIMAL - statind = 1
            'MIP - INTEGER OPTIMAL - statind = 101
            _engine.SolutionStatusCode = statind

            If _engine.SolutionStatus.Contains("OPTIMAL") Then
                _engine.CommonSolutionStatus = "OPTIMAL SOLUTION"
                _engine.CommonSolutionStatusCode = Solver.commonSolutionStatus.statusOptimal
                Console.WriteLine("Solver - CPLEX - Solution Status = " & _engine.SolutionStatus)
                EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - Solution Status = " & _engine.SolutionStatus)
            Else
                _engine.CommonSolutionStatus = "INFEASIBLE SOLUTION"
                _engine.CommonSolutionStatusCode = Solver.commonSolutionStatus.statusInfeasible

                Console.WriteLine("Solver - CPLEX - Solution Status = " & _engine.SolutionStatus)
                EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - Solution Status = " & _engine.SolutionStatus)
            End If

            _engine.SolutionRows = cur_numrows
            _engine.SolutionColumns = cur_numcols
            _engine.SolutionNonZeros = numNZ 'dtMtx.Rows.Count
            _engine.SolutionObj = objval
            If MyBase.isMIP Then 'If _probTypeMIP Then
                _engine.SolutionIterations = CPXgetmipitcnt(env, lp)
            Else
                _engine.SolutionIterations = CPXgetitcnt(env, lp)
            End If
            _engine.SolverName = "CPLEX"
            'RKP/01-28-10/v2.3.128
            _engine.SolverVersion = System.Runtime.InteropServices.Marshal.PtrToStringAnsi(CPXversion(env)) '& " (" & IIf(GenUtils.Is64Bit, "64-bit", "32-bit") & ")"

            If String.IsNullOrEmpty(_engine.SolverVersion) Then
                _engine.SolverVersion = System.Diagnostics.FileVersionInfo.GetVersionInfo(My.Application.Info.DirectoryPath & "\cplex.dll").FileVersion.ToString() & " (" & IIf(GenUtils.Is64Bit, "64-bit", "32-bit") & ")"
            Else
                _engine.SolverVersion = _engine.SolverVersion & " (" & IIf(GenUtils.Is64Bit, "64-bit", "32-bit") & ")"
            End If

            'Debug.Print(System.Diagnostics.FileVersionInfo.GetVersionInfo(My.Application.Info.DirectoryPath & "\cplex.dll").FileVersion.ToString())
            ' Write the output to the screen.
            System.Console.WriteLine("\nSolution status = {0}", solstat)
            System.Console.WriteLine("Solution value  = {0,10:f6}\n", objval)

            'For i As Integer = 0 To cur_numrows - 1
            '    System.Console.WriteLine("Row {0}:  Slack = {1,10:f6}  Pi = {2,10:f6}", i, slack(i), pi(i))
            'Next i

            'For j As Integer = 0 To cur_numcols - 1
            '    System.Console.WriteLine("Column {0}:  Value = {1,10:f6}  Reduced cost = {2,10:f6}", _
            '     j, x(j), dj(j))
            'Next j

            ' Finally, write a copy of the problem to a file.
            Try
                My.Computer.FileSystem.DeleteFile(outputFile & ".sav")
            Catch ex As System.Exception

            End Try
            Try
                My.Computer.FileSystem.DeleteFile(outputFile & ".mps")
            Catch ex As System.Exception

            End Try
            Try
                My.Computer.FileSystem.DeleteFile(outputFile & ".lp")
            Catch ex As System.Exception

            End Try
            Try
                My.Computer.FileSystem.DeleteFile(outputFile & ".rew")
            Catch ex As System.Exception

            End Try
            Try
                My.Computer.FileSystem.DeleteFile(outputFile & ".rmp")
            Catch ex As System.Exception

            End Try
            Try
                My.Computer.FileSystem.DeleteFile(outputFile & ".rlp")
            Catch ex As System.Exception

            End Try
            Try
                My.Computer.FileSystem.DeleteFile(outputFile & ".sol")
            Catch ex As System.Exception

            End Try
            Try
                My.Computer.FileSystem.DeleteFile(outputFile & ".clp")
            Catch ex As System.Exception

            End Try
            'RKP/01-12-12/v3.0.157
            'Look for a file called C:\OPTMODELS\<PRJ NAME>\Output\C-OPT.CPLEX.clp and delete it.
            Try
                My.Computer.FileSystem.DeleteFile(GenUtils.GetWorkDir(_switches) & "\" & "C-OPT.CPLEX.clp")
            Catch ex As System.Exception

            End Try
            Try
                My.Computer.FileSystem.DeleteFile(outputFile & ".bas")
            Catch ex As System.Exception

            End Try
            Try
                My.Computer.FileSystem.DeleteFile(outputFile & ".mst")
            Catch ex As System.Exception

            End Try
            Try
                My.Computer.FileSystem.DeleteFile(outputFile & ".prm")
            Catch ex As System.Exception

            End Try

            If GenUtils.IsSwitchAvailable(_switches, "/GenSAVFile") Then
                Try
                    status = CPXwriteprob(env, lp, outputFile & ".sav", Nothing)
                    EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - cplex.sav written to disk.")
                Catch ex As System.Exception
                    EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - cplex.sav not written to disk.")
                    System.Console.Error.WriteLine("Failed to write SAV to disk.")
                End Try
            End If

            If GenUtils.IsSwitchAvailable(_switches, "/GenSolverMPS") Then
                Try
                    status = CPXwriteprob(env, lp, outputFile & ".mps", "MPS")
                    EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - cplex.mps written to disk.")
                Catch ex As System.Exception
                    EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - cplex.mps not written to disk.")
                    System.Console.Error.WriteLine("Failed to write MPS to disk.")
                End Try
            End If

            If GenUtils.IsSwitchAvailable(_switches, "/GenLPFile") Then
                Try
                    status = CPXwriteprob(env, lp, outputFile & ".lp", Nothing)
                    EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - cplex.lp written to disk.")
                Catch ex As System.Exception
                    EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - cplex.lp not written to disk.")
                    System.Console.Error.WriteLine("Failed to write LP to disk.")
                End Try
            End If

            If GenUtils.IsSwitchAvailable(_switches, "/GenREWFile") Then
                Try
                    status = CPXwriteprob(env, lp, outputFile & ".rew", Nothing)
                    EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - cplex.rew written to disk.")
                Catch ex As System.Exception
                    EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - cplex.rew not written to disk.")
                    System.Console.Error.WriteLine("Failed to write REW to disk.")
                End Try
            End If

            If GenUtils.IsSwitchAvailable(_switches, "/GenRMPFile") Then
                Try
                    status = CPXwriteprob(env, lp, outputFile & ".rmp", Nothing)
                    EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - cplex.rmp written to disk.")
                Catch ex As System.Exception
                    EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - cplex.rmp not written to disk.")
                    System.Console.Error.WriteLine("Failed to write RMP to disk.")
                End Try
            End If

            If GenUtils.IsSwitchAvailable(_switches, "/GenRLPFile") Then
                Try
                    status = CPXwriteprob(env, lp, outputFile & ".rlp", Nothing)
                    EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - cplex.rlp written to disk.")
                Catch ex As System.Exception
                    EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - cplex.rlp not written to disk.")
                    System.Console.Error.WriteLine("Failed to write RLP to disk.")
                End Try
            End If

            If GenUtils.IsSwitchAvailable(_switches, "/GenBASFile") Then
                Try
                    status = CPXmbasewrite(env, lp, outputFile & ".bas")
                    EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - cplex.bas written to disk.")
                Catch ex As System.Exception
                    EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - cplex.bas not written to disk.")
                    System.Console.Error.WriteLine("Failed to write BAS to disk.")
                End Try
            End If

            If GenUtils.IsSwitchAvailable(_switches, "/GenSOLXMLFile") Then
                Try
                    'Dim writer As IO.StreamWriter = New IO.StreamWriter(MyBase._workDir & "\cplex.sol")
                    'ret = CPXsolwrite(env, lp, MyBase._workDir & "\cplex.sol")
                    'ret = CPXsolwrite(env, lp, "cplex.sol")
                    'ret = CPXsolwritesolnpoolall(env, lp, MyBase._workDir & "\cplex.sol")
                    'ret = 0
                    'If ret = 0 Then
                    'MessageBox.Show("SOL write successful")
                    'Else
                    'ret = CPXsolwritesolnpoolall(env, lp, MyBase._workDir & "\cplex.sol")
                    ret = CPXsolwritesolnpool(env, lp, -1, outputFile & ".sol")
                    If ret <> 0 Then
                        EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - cplex.sol not written to disk.")
                        System.Console.Error.WriteLine("Failed to write SOL to disk.")
                    End If
                    'MessageBox.Show("SOL write not successful")
                    'End If
                    'writer.Close()
                Catch ex As System.Exception
                    EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - cplex.sol not written to disk.")
                    System.Console.Error.WriteLine("Failed to write SOL to disk.")
                End Try
            End If

            If GenUtils.IsSwitchAvailable(_switches, "/GenSOLFixedFile") Then
                Try
                    'Dim writer As IO.StreamWriter = New IO.StreamWriter(MyBase._workDir & "\cplex.sol")
                    'ret = CPXsolwrite(env, lp, MyBase._workDir & "\cplex.sol")
                    'ret = CPXsolwrite(env, lp, "cplex.sol")
                    'ret = CPXsolwritesolnpoolall(env, lp, MyBase._workDir & "\cplex.sol")
                    'ret = 0
                    'If ret = 0 Then
                    'MessageBox.Show("SOL write successful")
                    'Else
                    'ret = CPXsolwritesolnpoolall(env, lp, MyBase._workDir & "\cplex.sol")
                    ret = CPXsolwritesolnpool(env, lp, -1, outputFile & ".sol")
                    If ret <> 0 Then
                        EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - cplex.sol not written to disk.")
                        System.Console.Error.WriteLine("Failed to write SOL to disk.")
                    End If
                    'MessageBox.Show("SOL write not successful")
                    'End If
                    'writer.Close()
                Catch ex As System.Exception
                    EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - cplex.sol not written to disk.")
                    System.Console.Error.WriteLine("Failed to write SOL to disk.")
                End Try
            End If

            If GenUtils.IsSwitchAvailable(_switches, "/GenMSTFile") Then
                If MyBase.isMIP Then  'If _probTypeMIP Then
                    Try
                        ret = CPXwritemipstarts(env, lp, outputFile & ".mst", 0, CPXgetnumcols(env, lp) - 1)
                    Catch ex As System.Exception

                    End Try
                End If
            End If

            If GenUtils.IsSwitchAvailable(_switches, "/GenPRMFile") Then
                Try
                    status = CPXwriteparam(env, outputFile & ".prm")
                    EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - cplex.prm written to disk.")
                Catch ex As System.Exception
                    EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - cplex.prm not written to disk.")
                    System.Console.Error.WriteLine("Failed to write PRM to disk.")
                End Try
            End If

            'RKP/03-11-10/v2.3.132
            'Decided to apply "/GenCLPFile" as a default option.
            'If GenUtils.IsSwitchAvailable(_switches, "/GenCLPFile") Then
            If GenUtils.IsSwitchAvailable(_switches, "/NoGenCLPFile") Then
                'do nothing
            Else
                Try
                    ret = CPXrefineconflict(env, lp, Nothing, Nothing)
                    If ret = 0 Then
                        Try
                            ret = CPXclpwrite(env, lp, outputFile & ".clp")
                            If ret = 0 Then
                                EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - cplex.clp written to disk.")
                            Else
                                EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - cplex.clp not written to disk. CPXclpwrite return status = " & ret.ToString())
                            End If

                            'RKP/01-12-12/v3.0.157
                            'Create a duplicate version of the CLP file so that it is recognizable in the Output folder.
                            ret = CPXclpwrite(env, lp, GenUtils.GetWorkDir(_switches) & "\" & "C-OPT.CPLEX.clp")
                            If ret = 0 Then
                                EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - C-OPT.CPLEX.clp written to disk.")
                            Else
                                EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - C-OPT.CPLEX.clp not written to disk. CPXclpwrite return status = " & ret.ToString())
                            End If
                        Catch ex As System.Exception
                            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - cplex.clp not written to disk.")
                            System.Console.Error.WriteLine("Failed to write CLP to disk.")
                        End Try
                    End If
                Catch ex As System.Exception
                    EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - cplex.sol not written to disk.")
                    System.Console.Error.WriteLine("Failed to write SOL to disk.")
                End Try
            End If

            'Debug.Print(statusSolution)
            If (statusSolution <> 0) Then
                System.Console.Error.WriteLine("Failed to obtain solution.")
                success = False
                GoTo TERMINATE
            End If
            success = True

            'RKP/02-15-10/v2.3.130
            '
            'Don't save the resulting MIP solution yet.
            'Re-optimize the problem as Continuous, with the following changes:
            'LO(column) = activity(column)
            'UP(column) = activity(column)
            'This process should now yield DJ(column), ACTIVITY(row) and SHADOW(row).
            'Now go ahead and save the problem back to the database.
            '/NoContinuous
            If MyBase.isMIP Then
                If Not GenUtils.IsSwitchAvailable(_switches, "/NoContinuous") Then
                    If statusSolution = 0 Then 'Proceed only if optimal
                        'Free up the problem as allocated by CPXcreateprob, if necessary
                        If (Not lp.Equals(IntPtr.Zero)) Then
                            status = CPXfreeprob(env, lp)
                            If status Then
                                System.Console.Error.WriteLine("CPXfreeprob failed, error code {0}.", status)
                            End If
                        End If

                        'Free up the CPLEX environemt, if necessary
                        If Not (env.Equals(IntPtr.Zero)) Then
                            status = CPXcloseCPLEX(env)
                            If status Then
                                Dim errmsg As StringBuilder = New StringBuilder(1024)
                                System.Console.Error.WriteLine("Could not close CPLEX environment.")
                                CPXgeterrorstring(env, status, errmsg)
                                System.Console.Error.WriteLine(errmsg)
                            End If
                        End If

                        'Initialize the CPLEX environment
                        env = Wrapper_CPLEX.CPXopenCPLEX(status)
                        If env.Equals(IntPtr.Zero) Then
                            Dim errmsg As StringBuilder = New StringBuilder(1024)
                            System.Console.Error.WriteLine("Could not open CPLEX environment.")
                            CPXgeterrorstring(env, status, errmsg)
                            System.Console.Error.WriteLine(errmsg)
                            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - CPXopenCPLEX failed.")
                            GoTo TERMINATE
                        End If

                        'Turn on output to the screen
                        status = CPXsetintparam(env, CPX_PARAM_SCRIND, CPX_ON)
                        If status <> 0 Then
                            System.Console.Error.WriteLine("Failure to turn on screen indicator, error {0}.", status)
                            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - CPXsetintparam screen indicator failed.")
                            GoTo TERMINATE
                        End If

                        'Turn on data checking
                        status = CPXsetintparam(env, CPX_PARAM_DATACHECK, CPX_ON)
                        If status <> 0 Then
                            System.Console.Error.WriteLine("Failure to turn on data checking, error {0}.", status)
                            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - CPXsetintparam data check failed.")
                            GoTo TERMINATE
                        End If

                        'Create the problem
                        lp = CPXcreateprob(env, status, GenUtils.GetSwitchArgument(_switches, "/PRJ", 1))

                        ' A returned pointer of NULL may mean that not enough memory
                        ' was available or there was some other problem.  In the case of 
                        ' failure, an error message will have been written to the error 
                        ' channel from inside CPLEX.  In this example, the setting of
                        ' the parameter CPX_PARAM_SCRIND causes the error message to
                        ' appear on stdout. 

                        If lp.Equals(IntPtr.Zero) Then
                            System.Console.Error.WriteLine("Failed to create LP.")
                            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - Failed to create LP.")
                            GoTo TERMINATE
                        End If

                        For i = 0 To x.Length - 2
                            If MyBase.array_ctyp(i) <> "C" Then
                                MyBase.array_dclo(i) = x(i)
                                MyBase.array_dcup(i) = x(i)
                            End If
                        Next
                        'Pretend to run this problem as a Continuous one.
                        status = PopulateByColumn(env, lp, dtRow, dtCol, dtMtx, numNZ, False)
                        If status Then
                            System.Console.Error.WriteLine("Failed to populate problem (last stage of MIP).")
                            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - Failed to populate problem (last stage of MIP).")
                            GoTo TERMINATE
                        End If

                        timeOptMIPContinuous = My.Computer.Clock.TickCount

                        status = CPXlpopt(env, lp)
                        EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - Problem type = Continuous (last stage of MIP)")
                        statusSolution = status
                        If (status) Then
                            System.Console.Error.WriteLine("Failed to optimize LP.")
                            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - Failed to optimize LP (last stage of MIP).")
                            GoTo TERMINATE
                        End If

                        Try
                            status = CPXsolution(env, lp, solstat, objval, x, pi, slack, dj)
                        Catch ex As System.Exception
                            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - CPXsolution (last stage of MIP) - " & ex.Message)
                            'MessageBox.Show(ex.Message)
                            GenUtils.Message(GenUtils.MsgType.Warning, "Solver - CPLEX", ex.Message)
                        End Try

                        Try
                            statind = CPXgetstat(env, lp)
                        Catch ex As System.Exception
                            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - CPXgetstat - " & ex.Message)
                            'MessageBox.Show(ex.Message)
                            GenUtils.Message(GenUtils.MsgType.Warning, "Solver - CPLEX", ex.Message)
                        End Try

                        Try
                            ret = CPXgetstatstring(env, statind, buffer)
                        Catch ex As System.Exception
                            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - CPXgetstatstring - " & ex.Message)
                            'MessageBox.Show(ex.Message)
                            GenUtils.Message(GenUtils.MsgType.Warning, "Solver - CPLEX", ex.Message)
                        End Try

                        _progress = "C-OPT Engine - Solver - CPLEX - MIP Continous - Stage 2 - Solve Time: " & GenUtils.FormatTime(timeOptMIPContinuous, My.Computer.Clock.TickCount)
                        EntLib.COPT.Log.Log(_workDir, _progress)
                        Console.WriteLine(_progress)
                        _progress = "C-OPT Engine - Solver - CPLEX - MIP Continous - Stage 2 - Solve Status: " & statind.ToString() & " - " & buffer.ToString()
                        EntLib.COPT.Log.Log(_workDir, _progress)
                        Console.WriteLine(_progress)

                        '_engine.SolutionStatus = statind & " (" & buffer.ToString() & ")" '"1=Optimal Solution"
                        '_engine.SolutionRows = cur_numrows
                        '_engine.SolutionColumns = cur_numcols
                        '_engine.SolutionNonZeros = dtMtx.Rows.Count
                        '_engine.SolutionObj = objval

                        ' Write the output to the screen.
                        System.Console.WriteLine("\nSolution status (MIP+Continuous) = ", statind & " (" & buffer.ToString() & ")")
                        System.Console.WriteLine("\nSolution status (MIP+Continuous) = {0}", solstat)
                        System.Console.WriteLine("Solution value (MIP+Continuous)  = {0,10:f6}\n", objval)


                    End If 'If statusSolution = 0 Then 'Proceed only if optimal
                End If 'If Not GenUtils.IsSwitchAvailable(_switches, "/NoContinuous") Then
            End If 'If MyBase.isMIP Then

            'GenUtils.SerializeArrayDouble(Engine.GetWorkDir, "COL.ACTIVITY.txt", x)
            'GenUtils.SerializeArrayDouble(Engine.GetWorkDir, "COL.REDUCED.txt", dj)
            'GenUtils.SerializeArrayDouble(Engine.GetWorkDir, "ROW.SLACK.txt", slack)
            'GenUtils.SerializeArrayDouble(Engine.GetWorkDir, "ROW.SHADOW.txt", pi)

            _engine.Progress = "Updating solver results to in-memory column table...Start - " & My.Computer.Clock.LocalTime.ToString()
            Debug.Print(_engine.Progress)
            Console.WriteLine(_engine.Progress)
            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CPLEX", _engine.Progress)

            For i = 0 To dtCol.Rows.Count - 1
                'System.Console.WriteLine("Column {0}:  Value = {1,10:f6}  Reduced cost = {2,10:f6}", i, x(i), dj(i))
                dtCol.Rows(i).Item("ACTIVITY") = x(i) '0 'activity(i) 'xa.getColumnPrimalActivity(ds.Tables("tsysCol").Rows(i).Item("COL").ToString())
                dtCol.Rows(i).Item("DJ") = dj(i) '0 'reduced(i) 'xa.getColumnDualActivity(ds.Tables("tsysCol").Rows(i).Item("COL").ToString())
            Next

            _engine.Progress = "Updating solver results to in-memory column table...End - " & My.Computer.Clock.LocalTime.ToString()
            Debug.Print(_engine.Progress)
            Console.WriteLine(_engine.Progress)
            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CPLEX", _engine.Progress)

            'MyBase.Engine.CurrentDb.UpdateDataSet(dtCol, "SELECT * FROM " & col)

            _engine.Progress = "Updating solver results to in-memory row table...Start - " & My.Computer.Clock.LocalTime.ToString()
            Debug.Print(_engine.Progress)
            Console.WriteLine(_engine.Progress)
            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CPLEX", _engine.Progress)

            For i = 0 To dtRow.Rows.Count - 1
                'System.Console.WriteLine("Row {0}:  Slack = {1,10:f6}  Pi = {2,10:f6}", i, slack(i), pi(i))
                dtRow.Rows(i).Item("ACTIVITY") = MyBase.array_drhs(i) - slack(i) 'xa.getRowPrimalActivity(ds.Tables("tsysRow").Rows(i).Item("ROW").ToString())
                dtRow.Rows(i).Item("SHADOW") = pi(i) '0 'shadow(i) 'xa.getRowDualActivity(ds.Tables("tsysRow").Rows(i).Item("ROW").ToString())
            Next

            _engine.Progress = "Updating solver results to in-memory row table...End - " & My.Computer.Clock.LocalTime.ToString()
            Debug.Print(_engine.Progress)
            Console.WriteLine(_engine.Progress)
            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CPLEX", _engine.Progress)

            'MyBase.Engine.CurrentDb.UpdateDataSet(_ds.Tables("tsysRow"), "SELECT * FROM " & MyBase.Engine.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString)
            'MyBase.Engine.CurrentDb.UpdateDataSet(_ds.Tables("tsysCol"), "SELECT * FROM " & MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString)

            'RKP/06-16-10/v2.3.133
            'For large models, it is assumed that tsysCOL, tsysROW and tsysMTX are split into 1, 2 or 3 separate databases.
            'C-OPT will update each of the three tables directly against the source MDB rather than use the linked table, which drastically slows down performance.
            'tsysRow/tsysCol

            _engine.Progress = "Importing solution...Start - " & My.Computer.Clock.LocalTime.ToString()
            Debug.Print(_engine.Progress)
            Console.WriteLine(_engine.Progress)
            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CPLEX", _engine.Progress)

            If GenUtils.IsSwitchAvailable(_switches, "/UseMinSysRes") Then
                _engine.Progress = "Importing rows...Start - " & My.Computer.Clock.LocalTime.ToString()
                Debug.Print(_engine.Progress)
                Console.WriteLine(_engine.Progress)
                EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CPLEX", _engine.Progress)

                ''''''''''''''''ROWS''''''''''''''''
                If GenUtils.IsSwitchAvailable(_switches, "/UseMSAccessSyntax") Then
                    'GenUtils.UpdateDB(_engine.MiscParams.Item("LINKED_ROW_DB_CONN_STR").ToString(), dtRow, "SELECT * FROM [" & MyBase.Engine.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString() & "] ORDER BY RowID")
                    Try
                        'GenUtils.UpdateDB(_engine.MiscParams.Item("LINKED_ROW_DB_CONN_STR").ToString(), dtRow, "SELECT * FROM [" & MyBase.Engine.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString() & "] ORDER BY RowID")
                        MyBase.Engine.CurrentDb.UpdateDataSet(dtRow, "SELECT * FROM [" & _engine.MiscParams.Item("LINKED_ROW_DB_PATH").ToString() & "].[" & MyBase.Engine.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString & "] ORDER BY RowID")
                    Catch ex As System.Exception
                        _progress = "Error using /UseMSAccessSyntax switch. Importing solution without switch."
                        Console.WriteLine(ex.Message)
                        Debug.Print(ex.Message)
                        Debug.Print("Error using /UseMSAccessSyntax switch. Importing solution without switch.")
                        Console.WriteLine("Error using /UseMSAccessSyntax switch. Importing solution without switch.")
                        EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CPLEX", "Error using /UseMSAccessSyntax switch. Importing solution without switch.")

                        MyBase.Engine.CurrentDb.UpdateDataSet(dtRow, "SELECT * FROM " & MyBase.Engine.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString & " ORDER BY RowID")

                        'Debug.Print(ex.Message)
                        'Debug.Print("Error using /UseMSAccessSyntax switch. Importing solution without switch.")

                        'Console.WriteLine(ex.Message)
                        'Console.WriteLine("Error using /UseMSAccessSyntax switch. Importing solution without switch.")

                        'EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CPLEX", "Error using /UseMSAccessSyntax switch. Importing solution without switch.")

                        'MyBase.Engine.CurrentDb.UpdateDataSet(dtRow, "SELECT * FROM " & MyBase.Engine.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString & " ORDER BY RowID")
                    End Try
                ElseIf GenUtils.IsSwitchAvailable(_switches, "/UseSQLServerSyntax") Then
                    updateSQLServerRow = True

                Else
                    MyBase.Engine.CurrentDb.UpdateDataSet(dtRow, "SELECT * FROM " & MyBase.Engine.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString & " ORDER BY RowID")
                End If

                If updateSQLServerRow Then
                    Try
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
                            Debug.Print("C-OPT Engine - Solver - CPLEX - Error importing Constraints (tsysRow) - " & ex.Message & " - " & My.Computer.Clock.LocalTime.ToString)
                            Console.WriteLine("C-OPT Engine - Solver - CPLEX - Error importing Constraints (tsysRow) - " & ex.Message & " - " & My.Computer.Clock.LocalTime.ToString)
                            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CPLEX", "C-OPT Engine - Solver - CPLEX - Error importing Constraints (tsysRow) - " & ex.Message & " - " & My.Computer.Clock.LocalTime.ToString)
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
                            Debug.Print("C-OPT Engine - Solver - CPLEX - Error importing Constraints (tsysRow) - " & ex.Message & " - " & My.Computer.Clock.LocalTime.ToString)
                            Console.WriteLine("C-OPT Engine - Solver - CPLEX - Error importing Constraints (tsysRow) - " & ex.Message & " - " & My.Computer.Clock.LocalTime.ToString)
                            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CPLEX", "C-OPT Engine - Solver - CPLEX - Error importing Constraints (tsysRow) - " & ex.Message & " - " & My.Computer.Clock.LocalTime.ToString)
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
                            Debug.Print("C-OPT Engine - Solver - CPLEX - Error importing Constraints (tsysRow) - " & ex.Message & " - " & My.Computer.Clock.LocalTime.ToString)
                            Console.WriteLine("C-OPT Engine - Solver - CPLEX - Error importing Constraints (tsysRow) - " & ex.Message & " - " & My.Computer.Clock.LocalTime.ToString)
                            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CPLEX", "C-OPT Engine - Solver - CPLEX - Error importing Constraints (tsysRow) - " & ex.Message & " - " & My.Computer.Clock.LocalTime.ToString)
                        End Try
                        EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CPLEX", "C-OPT Engine - Solver - CPLEX - Successfully imported Rows. " & My.Computer.Clock.LocalTime.ToString)
                    Catch ex As System.Exception

                        _progress = "Error using /UseSQLServerSyntax switch. Importing solution without switch."
                        Console.WriteLine(ex.Message)
                        Debug.Print(ex.Message)
                        Debug.Print("Error using /UseSQLServerSyntax switch. Importing solution without switch.")
                        Console.WriteLine("Error using /UseSQLServerSyntax switch. Importing solution without switch.")
                        EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CPLEX", "Error using /UseSQLServerSyntax switch. Importing solution without switch.")

                        MyBase.Engine.CurrentDb.UpdateDataSet(dtRow, "SELECT * FROM " & MyBase.Engine.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString & " ORDER BY RowID")
                    End Try
                End If

                _engine.Progress = "Importing rows...End - " & My.Computer.Clock.LocalTime.ToString()
                Debug.Print(_engine.Progress)
                Console.WriteLine(_engine.Progress)
                EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CPLEX", _engine.Progress)

                _engine.Progress = "Importing columns...Start - " & My.Computer.Clock.LocalTime.ToString()
                Debug.Print(_engine.Progress)
                Console.WriteLine(_engine.Progress)
                EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CPLEX", _engine.Progress)

                ''''''''''''''''COLUMNS''''''''''''''''
                If GenUtils.IsSwitchAvailable(_switches, "/UseMSAccessSyntax") Then
                    'GenUtils.UpdateDB(_engine.MiscParams.Item("LINKED_COL_DB_CONN_STR").ToString(), dtCol, "SELECT * FROM [" & MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString() & "] ORDER BY ColID")
                    Try
                        'GenUtils.UpdateDB(_engine.MiscParams.Item("LINKED_COL_DB_CONN_STR").ToString(), dtCol, "SELECT * FROM [" & MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString() & "] ORDER BY ColID")
                        MyBase.Engine.CurrentDb.UpdateDataSet(dtCol, "SELECT * FROM [" & _engine.MiscParams.Item("LINKED_COL_DB_PATH").ToString() & "].[" & MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString & "] ORDER BY ColID")
                    Catch ex As System.Exception
                        _progress = "Error using /UseMSAccessSyntax switch. Importing solution without switch."
                        Console.WriteLine(ex.Message)
                        Debug.Print(ex.Message)
                        Debug.Print("Error using /UseMSAccessSyntax switch. Importing solution without switch.")
                        Console.WriteLine("Error using /UseMSAccessSyntax switch. Importing solution without switch.")
                        EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CPLEX", "Error using /UseMSAccessSyntax switch. Importing solution without switch.")

                        MyBase.Engine.CurrentDb.UpdateDataSet(dtCol, "SELECT * FROM " & MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString & " ORDER BY ColID")

                        'Debug.Print(ex.Message)
                        'Debug.Print("Error using /UseMSAccessSyntax switch. Importing solution without switch.")

                        'Console.WriteLine(ex.Message)
                        'Console.WriteLine("Error using /UseMSAccessSyntax switch. Importing solution without switch.")

                        'EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CPLEX", "Error using /UseMSAccessSyntax switch. Importing solution without switch.")

                        'MyBase.Engine.CurrentDb.UpdateDataSet(dtCol, "SELECT * FROM " & MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString & " ORDER BY ColID")
                    End Try
                ElseIf GenUtils.IsSwitchAvailable(_switches, "/UseSQLServerSyntax") Then
                    updateSQLServerCol = True
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
                            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CPLEX", "Error using /UseSQLServerSyntax switch. Importing solution without switch.")
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
                            Debug.Print("C-OPT Engine - Solver - CPLEX - Error importing Columns (tsysCol) - " & ex.Message & " - " & My.Computer.Clock.LocalTime.ToString)
                            Console.WriteLine("C-OPT Engine - Solver - CPLEX - Error importing Columns (tsysCol) - " & ex.Message & " - " & My.Computer.Clock.LocalTime.ToString)
                            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CPLEX", "C-OPT Engine - Solver - CPLEX - Error importing Columns (tsysCol) - " & ex.Message & " - " & My.Computer.Clock.LocalTime.ToString)
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
                            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CPLEX", "Error using /UseSQLServerSyntax switch. Importing solution without switch.")
                        End Try
                        EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CPLEX", "C-OPT Engine - Solver - CPLEX - Successfully imported Columns. " & My.Computer.Clock.LocalTime.ToString)
                    Catch ex As System.Exception
                        _progress = "Error using /UseSQLServerSyntax switch. Importing solution without switch."
                        Console.WriteLine(ex.Message)
                        Debug.Print(ex.Message)
                        Debug.Print("Error using /UseSQLServerSyntax switch. Importing solution without switch.")
                        Console.WriteLine("Error using /UseSQLServerSyntax switch. Importing solution without switch.")
                        EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CPLEX", "Error using /UseSQLServerSyntax switch. Importing solution without switch.")

                        MyBase.Engine.CurrentDb.UpdateDataSet(dtCol, "SELECT * FROM " & MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString & " ORDER BY ColID")
                    End Try
                Else
                    MyBase.Engine.CurrentDb.UpdateDataSet(dtCol, "SELECT * FROM " & MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString & " ORDER BY ColID")
                End If

                _engine.Progress = "Importing columns...End - " & My.Computer.Clock.LocalTime.ToString()
                Debug.Print(_engine.Progress)
                Console.WriteLine(_engine.Progress)
                EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CPLEX", _engine.Progress)
            Else 'If GenUtils.IsSwitchAvailable(_switches, "/UseMinSysRes") = False Then 

                If MyBase.Engine.CurrentDb.IsSQLExpress Then

                    _engine.Progress = "Importing rows...Start - " & My.Computer.Clock.LocalTime.ToString()
                    Debug.Print(_engine.Progress)
                    Console.WriteLine(_engine.Progress)
                    EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CPLEX", _engine.Progress)

                    Try
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
                            Debug.Print("C-OPT Engine - Solver - CPLEX - Error importing Constraints (tsysRow) - " & ex.Message & " - " & My.Computer.Clock.LocalTime.ToString)
                            Console.WriteLine("C-OPT Engine - Solver - CPLEX - Error importing Constraints (tsysRow) - " & ex.Message & " - " & My.Computer.Clock.LocalTime.ToString)
                            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CPLEX", "C-OPT Engine - Solver - CPLEX - Error importing Constraints (tsysRow) - " & ex.Message & " - " & My.Computer.Clock.LocalTime.ToString)
                        End Try

                        'MyBase.Engine.CurrentDb.UpdateDataSet(dtRow, "SELECT * FROM [" & _engine.MiscParams.Item("LINKED_ROW_DB_PATH").ToString() & "].[" & MyBase.Engine.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString & "] ORDER BY RowID")
                        'MyBase.Engine.CurrentDb.UpdateDataSet(dtRow, MyBase.Engine.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString() & "Temp", _switches)

                        'MyBase.Engine.CurrentDb.UpdateDataSet(dtRow, MyBase.Engine.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString() & "Temp", _switches, True)

                        If GenUtils.IsSwitchAvailable(_switches, "/SubNameWithID") Then
                            'dtRow.GetChanges()

                            'MyBase.Engine.CurrentDb.UpdateDataSet(dtRow, "SELECT RowID, NULL AS ROW, NULL AS [DESC], RHS, SENSE, ACTIVITY, SHADOW, STATUS FROM " & MyBase.Engine.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString & "Temp" & " ORDER BY RowID")
                        Else
                            'MyBase.Engine.CurrentDb.UpdateDataSet(dtRow, "SELECT * FROM " & MyBase.Engine.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString & " ORDER BY RowID")
                            'MyBase.Engine.CurrentDb.UpdateDataSet(dtRow, MyBase.Engine.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString() & "Temp", _switches, True)
                        End If
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
                            Debug.Print("C-OPT Engine - Solver - CPLEX - Error importing Constraints (tsysRow) - " & ex.Message & " - " & My.Computer.Clock.LocalTime.ToString)
                            Console.WriteLine("C-OPT Engine - Solver - CPLEX - Error importing Constraints (tsysRow) - " & ex.Message & " - " & My.Computer.Clock.LocalTime.ToString)
                            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CPLEX", "C-OPT Engine - Solver - CPLEX - Error importing Constraints (tsysRow) - " & ex.Message & " - " & My.Computer.Clock.LocalTime.ToString)
                        End Try

                        'Empty temp table now that it's purpose has been served.
                        If GenUtils.IsSwitchAvailable(_switches, "/DoNotEmptyTempTables") Then
                            'do nothing
                        Else
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
                                Debug.Print("C-OPT Engine - Solver - CPLEX - Error importing Constraints (tsysRow) - " & ex.Message & " - " & My.Computer.Clock.LocalTime.ToString)
                                Console.WriteLine("C-OPT Engine - Solver - CPLEX - Error importing Constraints (tsysRow) - " & ex.Message & " - " & My.Computer.Clock.LocalTime.ToString)
                                EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CPLEX", "C-OPT Engine - Solver - CPLEX - Error importing Constraints (tsysRow) - " & ex.Message & " - " & My.Computer.Clock.LocalTime.ToString)
                            End Try
                        End If

                        EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CPLEX", "C-OPT Engine - Solver - CPLEX - Successfully imported Rows. " & My.Computer.Clock.LocalTime.ToString)
                    Catch ex As System.Exception

                        _progress = "Error using /UseSQLServerSyntax switch. Importing solution without switch."
                        Console.WriteLine(ex.Message)
                        Debug.Print(ex.Message)
                        Debug.Print("Error using /UseSQLServerSyntax switch. Importing solution without switch.")
                        Console.WriteLine("Error using /UseSQLServerSyntax switch. Importing solution without switch.")
                        EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CPLEX", "Error using /UseSQLServerSyntax switch. Importing solution without switch.")

                        MyBase.Engine.CurrentDb.UpdateDataSet(dtRow, "SELECT * FROM " & MyBase.Engine.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString & " ORDER BY RowID")
                    End Try

                    _engine.Progress = "Importing rows...End - " & My.Computer.Clock.LocalTime.ToString()
                    Debug.Print(_engine.Progress)
                    Console.WriteLine(_engine.Progress)
                    EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CPLEX", _engine.Progress)

                    _engine.Progress = "Importing columns...Start - " & My.Computer.Clock.LocalTime.ToString()
                    Debug.Print(_engine.Progress)
                    Console.WriteLine(_engine.Progress)
                    EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CPLEX", _engine.Progress)

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
                            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CPLEX", "Error using /UseSQLServerSyntax switch. Importing solution without switch.")
                        End Try

                        'GenUtils.UpdateDB(_engine.MiscParams.Item("LINKED_COL_DB_CONN_STR").ToString(), dtCol, "SELECT * FROM [" & MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString() & "] ORDER BY ColID")
                        'MyBase.Engine.CurrentDb.UpdateDataSet(dtCol, MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString() & "Temp", _switches)

                        'MyBase.Engine.CurrentDb.UpdateDataSet(dtCol, MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString() & "Temp", _switches, True)

                        'If GenUtils.IsSwitchAvailable(_switches, "/SubNameWithID") Then
                        '    MyBase.Engine.CurrentDb.UpdateDataSet(dtCol, "SELECT ColID, NULL AS COL, NULL AS [DESC], OBJ, LO, UP, FREE, INTGR, BINRY, NULL AS SOSTYPE, NULL AS SOSMARKER, ACTIVITY, DJ, STATUS, ISVALID FROM " & MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString & "Temp" & " ORDER BY ColID")
                        'Else
                        '    MyBase.Engine.CurrentDb.UpdateDataSet(dtCol, "SELECT * FROM " & MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString & "Temp" & " ORDER BY ColID")
                        'End If

                        'RKP/09-14-12/v4.2.179
                        'Made the COL code in-sync with the ROW code, which it wasn't.
                        If GenUtils.IsSwitchAvailable(_switches, "/SubNameWithID") Then
                            'dtCol.GetChanges()
                            'MyBase.Engine.CurrentDb.UpdateDataSet(dtCol, "SELECT ColID, NULL AS COL, NULL AS [DESC], OBJ, LO, UP, FREE, INTGR, BINRY, NULL AS SOSTYPE, NULL AS SOSMARKER, ACTIVITY, DJ, STATUS, ISVALID FROM " & MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString & "Temp" & " ORDER BY ColID")
                        Else
                            'MyBase.Engine.CurrentDb.UpdateDataSet(dtRow, "SELECT * FROM " & MyBase.Engine.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString & " ORDER BY RowID")
                            'MyBase.Engine.CurrentDb.UpdateDataSet(dtCol, MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString() & "Temp", _switches, True)
                        End If
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
                            Debug.Print("C-OPT Engine - Solver - CPLEX - Error importing Columns (tsysCol) - " & ex.Message & " - " & My.Computer.Clock.LocalTime.ToString)
                            Console.WriteLine("C-OPT Engine - Solver - CPLEX - Error importing Columns (tsysCol) - " & ex.Message & " - " & My.Computer.Clock.LocalTime.ToString)
                            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CPLEX", "C-OPT Engine - Solver - CPLEX - Error importing Columns (tsysCol) - " & ex.Message & " - " & My.Computer.Clock.LocalTime.ToString)
                        End Try

                        'Empty the temp table, now that the purpose has been served.
                        If GenUtils.IsSwitchAvailable(_switches, "/DoNotEmptyTempTables") Then
                            'do nothing
                        Else
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
                                EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CPLEX", "Error using /UseSQLServerSyntax switch. Importing solution without switch.")
                            End Try
                        End If

                        EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CPLEX", "C-OPT Engine - Solver - CPLEX - Successfully imported Columns. " & My.Computer.Clock.LocalTime.ToString)
                    Catch ex As System.Exception
                        _progress = "Error using /UseSQLServerSyntax switch. Importing solution without switch."
                        Console.WriteLine(ex.Message)
                        Debug.Print(ex.Message)
                        Debug.Print("Error using /UseSQLServerSyntax switch. Importing solution without switch.")
                        Console.WriteLine("Error using /UseSQLServerSyntax switch. Importing solution without switch.")
                        EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CPLEX", "Error using /UseSQLServerSyntax switch. Importing solution without switch.")

                        MyBase.Engine.CurrentDb.UpdateDataSet(dtCol, "SELECT * FROM " & MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString & " ORDER BY ColID")
                    End Try

                    _engine.Progress = "Importing columns...End - " & My.Computer.Clock.LocalTime.ToString()
                    Debug.Print(_engine.Progress)
                    Console.WriteLine(_engine.Progress)
                    EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CPLEX", _engine.Progress)

                Else 'not SQLExpress

                    If GenUtils.IsSwitchAvailable(_switches, "/SubNameWithID") Then
                        MyBase.Engine.CurrentDb.UpdateDataSet(dtRow, "SELECT RowID, NULL AS ROW, NULL AS [DESC], RHS, SENSE, ACTIVITY, SHADOW, STATUS FROM " & MyBase.Engine.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString & " ORDER BY RowID")
                    Else
                        MyBase.Engine.CurrentDb.UpdateDataSet(dtRow, "SELECT * FROM " & MyBase.Engine.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString & " ORDER BY RowID")
                    End If

                    _engine.Progress = "Importing rows...End - " & My.Computer.Clock.LocalTime.ToString()
                    Debug.Print(_engine.Progress)
                    Console.WriteLine(_engine.Progress)
                    EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CPLEX", _engine.Progress)

                    _engine.Progress = "Importing columns...Start - " & My.Computer.Clock.LocalTime.ToString()
                    Debug.Print(_engine.Progress)
                    Console.WriteLine(_engine.Progress)
                    EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CPLEX", _engine.Progress)

                    'MyBase.Engine.CurrentDb.UpdateDataSet(dtCol, "SELECT * FROM " & MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString & " ORDER BY ColID")
                    If GenUtils.IsSwitchAvailable(_switches, "/SubNameWithID") Then
                        MyBase.Engine.CurrentDb.UpdateDataSet(dtCol, "SELECT ColID, NULL AS COL, NULL AS [DESC], OBJ, LO, UP, FREE, INTGR, BINRY, NULL AS SOSTYPE, NULL AS SOSMARKER, ACTIVITY, DJ, STATUS, ISVALID FROM " & MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString & " ORDER BY ColID")
                    Else
                        MyBase.Engine.CurrentDb.UpdateDataSet(dtCol, "SELECT * FROM " & MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString & " ORDER BY ColID")
                    End If

                    _engine.Progress = "Importing columns...End - " & My.Computer.Clock.LocalTime.ToString()
                    Debug.Print(_engine.Progress)
                    Console.WriteLine(_engine.Progress)
                    EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CPLEX", _engine.Progress)
                End If
            End If

            _engine.Progress = "Importing solution...End - " & My.Computer.Clock.LocalTime.ToString()
            Debug.Print(_engine.Progress)
            Console.WriteLine(_engine.Progress)
            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CPLEX", _engine.Progress)


            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - arrays written to tables.")
TERMINATE:
            'Free up the problem as allocated by CPXcreateprob, if necessary
            If (Not lp.Equals(IntPtr.Zero)) Then
                status = CPXfreeprob(env, lp)
                If status Then
                    System.Console.Error.WriteLine("CPXfreeprob failed, error code {0}.", status)
                End If
            End If

            'Free up the CPLEX environemt, if necessary
            If Not (env.Equals(IntPtr.Zero)) Then
                status = CPXcloseCPLEX(env)
                If status Then
                    Dim errmsg As StringBuilder = New StringBuilder(1024)
                    System.Console.Error.WriteLine("Could not close CPLEX environment.")
                    CPXgeterrorstring(env, status, errmsg)
                    System.Console.Error.WriteLine(errmsg)
                End If
            End If

            EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - Solver - CPLEX took: ", Space(9) & GenUtils.FormatTime(timeOptTotalStart, My.Computer.Clock.TickCount))

            Return success
        End Function

        Public Function SolveUsingConcert(ByRef dtRow As DataTable, ByRef dtCol As DataTable, ByRef dtMtx As DataTable) As Integer
            'Dim numCols As Integer = dtCol.Rows.Count
            'Dim numRows As Integer = dtRow.Rows.Count
            'Dim numNZ As Integer = dtMtx.Rows.Count

            'Dim concertModel As New ILOG.CPLEX.Cplex
            'Dim var(1)() As ILOG.Concert.INumVar
            'Dim rng(1)() As ILOG.Concert.IRange

            'Try
            '    If concertModel.Solve() Then
            '        Debug.Print(concertModel.GetStatus().ToString())
            '        Debug.Print(concertModel.ObjValue.ToString())

            '        Dim x As Double() = concertModel.GetValues(var(0))
            '        Dim dj As Double() = concertModel.GetReducedCosts(var(0))
            '        Dim pi As Double() = concertModel.GetDuals(rng(0))
            '        Dim slack As Double() = concertModel.GetSlacks(rng(0))

            '        concertModel.ExportModel("C:\OPTMODELS\KNOXMIX\Output\lpex1.lp")
            '    End If
            '    concertModel.End()
            'Catch ex As System.Exception

            'End Try

            'Dim concertModel As New ILOG.CPLEX.Cplex



            Return -1

        End Function

        Private Function PopulateByRow(ByVal env As IntPtr, ByVal lp As IntPtr) As Integer

            Dim status As Integer = 0

            CPXchgobjsen(env, lp, CPX_MAX)   ' Problem is maximization

            ' Now create the new columns.  First, populate the arrays.


TERMINATE:
            Return (status)

        End Function ' END populatebyrow 

        Private Function PopulateByColumn _
        ( _
            ByVal env As IntPtr, _
            ByVal lp As IntPtr, _
            ByRef dtRow As DataTable, _
            ByRef dtCol As DataTable, _
            ByRef dtMtx As DataTable, _
            ByVal nonZeroCount As Long, _
            ByVal isMIP As Boolean _
        ) As Long  'Integer
            'Dim status As Integer
            ''Dim obj(NUMCOLS) As Double
            'Dim obj(_ds.Tables("tsysCol").Rows.Count) As Double
            'Dim lb(_ds.Tables("tsysCol").Rows.Count) As Double
            'Dim ub(_ds.Tables("tsysCol").Rows.Count) As Double
            'Dim matbeg(_ds.Tables("tsysCol").Rows.Count) As Integer
            'Dim matind(_ds.Tables("tsysMtx").Rows.Count) As Integer
            'Dim matval(_ds.Tables("tsysMtx").Rows.Count) As Double
            'Dim rhs(_ds.Tables("tsysRow").Rows.Count) As Double
            'Dim sense(_ds.Tables("tsysRow").Rows.Count) As Byte

            Dim status As Integer = 0
            'Dim buffer As StringBuilder = New StringBuilder(4096)
            'Dim dobj() As Double
            'Dim dclo() As Double
            'Dim dcup() As Double
            'Dim colNames() As String
            'Dim mbeg() As Integer
            'Dim midx() As Integer
            'Dim mval() As Double
            'Dim drhs() As Double
            'Dim rtyp() As Char
            'Dim rowNames() As String
            'Dim mcnt() As Integer
            'Dim cumTotal As Integer = 0
            'Dim ctyp() As Char = Nothing  'RKP/01-22-10/v2.2.126 - to add MIP capability to CPLEX

            'Dim proceed As Boolean = True

            'Dim myArrayList As ArrayList
            'Dim linqTable As System.Data.EnumerableRowCollection(Of System.Data.DataRow)
            'Dim dblQueryResults As System.Data.EnumerableRowCollection(Of Double)
            'Dim intQueryResults As System.Data.EnumerableRowCollection(Of Integer)
            'Dim charQueryResults As System.Data.EnumerableRowCollection(Of Char)
            'Dim strQueryResults As System.Data.EnumerableRowCollection(Of String)
            'Dim intGroups As System.Collections.Generic.IEnumerable(Of Integer)

            ' Problem is maximization
            Call CPXchgobjsen(env, lp, CPX_MAX)

            ' Now create the new rows.  First, populate the arrays.



            'sense(0) = CChar(CStr(Asc("L")))
            'rhs(0) = 20.0#

            'sense(1) = CChar(CStr(Asc("L")))
            'rhs(1) = 30.0#

            'MyBase.LoadSolverArrays(_ds)
            'MyBase.LoadSolverArrays(dtRow, dtCol, dtMtx)

            Debug.Print("Rows=" & dtRow.Rows.Count)
            Debug.Print("Columns=" & dtCol.Rows.Count)
            Debug.Print("Matrix=" & nonZeroCount.ToString())

            'status = CPXnewrows(env, lp, dtRow.Rows.Count, drhs, rtyp, Nothing, rowNames)
            Try
                status = CPXnewrows(env, lp, dtRow.Rows.Count, MyBase.array_drhs, MyBase.array_rtyp, Nothing, MyBase.array_rowNames)
            Catch ex As System.Exception
                GenUtils.Message(GenUtils.MsgType.Critical, "CPLEX-PopulateByColum", "Failed at: CPXnewrows" & vbNewLine & ex.Message & vbNewLine & "APM: " & GenUtils.GetAvailablePhysicalMemoryStr)
            End Try

            If (status <> 0) Then
                'EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - CPXnewrows failed - " & status.ToString())
                'status = CPXgeterrorstring(env, status, buffer)
                EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - CPXnewrows failed - " & cplexGetErrorString(env, status))
                GoTo TERMINATE
            End If

            ' Now add the new columns.  First, populate the arrays.

            'obj(0) = 1.0# : obj(1) = 2.0# : obj(2) = 3.0#

            'matbeg(0) = 0 : matbeg(1) = 2 : matbeg(2) = 4

            'matind(0) = 0 : matind(2) = 0 : matind(4) = 0
            'matval(0) = -1.0# : matval(2) = 1.0# : matval(4) = 1.0#

            'matind(1) = 1 : matind(3) = 1 : matind(5) = 1
            'matval(1) = 1.0# : matval(3) = -3.0# : matval(5) = 1.0#

            'lb(0) = 0.0# : lb(1) = 0.0# : lb(2) = 0.0#
            'ub(0) = 40.0# : ub(1) = CPX_INFBOUND : ub(2) = CPX_INFBOUND

            'status = CPXaddcols(env, lp, _ds.Tables("tsysCol").Rows.Count, _ds.Tables("tsysMtx").Rows.Count, dobj, _
            '                          mbeg, midx, mval, dclo, dcup, colNames)
            Try
                status = CPXaddcols(env, lp, dtCol.Rows.Count, nonZeroCount, MyBase.array_dobj, _
                                          MyBase.array_mbeg, MyBase.array_midx, MyBase.array_mval, MyBase.array_dclo, MyBase.array_dcup, MyBase.array_colNames)
            Catch ex As System.Exception
                If status = 0 Then 'RKP/09-10-12/v4.2.177 - added this to avoid an arithmatic overflow error even though CPLEX was able to add columns successfully.
                    'GenUtils.Message(GenUtils.MsgType.Information, "CPLEX-PopulateByColum", "An exception occurred at CPXaddcols, even though status = 0." & vbNewLine & ex.Message & vbNewLine & "APM: " & GenUtils.GetAvailablePhysicalMemoryStr)
                    EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - An exception occurred at CPXaddcols, even though status = 0.")
                Else
                    GenUtils.Message(GenUtils.MsgType.Critical, "CPLEX-PopulateByColum", "Failed at: CPXaddcols" & vbNewLine & ex.Message & vbNewLine & "APM: " & GenUtils.GetAvailablePhysicalMemoryStr)
                End If
            End Try

            If (status <> 0) Then
                'EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - CPXaddcols failed - " & status.ToString())
                'status = CPXgeterrorstring(env, status, buffer)
                EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - CPXaddcols failed - " & cplexGetErrorString(env, status))
                GoTo TERMINATE
            End If

            If Not isMIP Then 'If MyBase.array_ctyp Is Nothing Then
                _probTypeMIP = False
            Else
                _probTypeMIP = True
                'status = CPXcopyctype(env, lp, ctyp)
                Try
                    status = CPXcopyctype(env, lp, MyBase.array_ctyp)
                Catch ex As System.Exception
                    GenUtils.Message(GenUtils.MsgType.Critical, "CPLEX-PopulateByColum", "Failed at: CPXcopyctype" & vbNewLine & ex.Message & vbNewLine & "APM: " & GenUtils.GetAvailablePhysicalMemoryStr)
                End Try

                If (status <> 0) Then
                    EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CPLEX - CPXcopyctype failed - " & status.ToString())
                    GoTo TERMINATE
                End If
            End If

TERMINATE:

            PopulateByColumn = status

        End Function

        Private Overloads Function GetSolutionStatus(ByVal solstat As Long) As String
            Select Case solstat
                Case 0
                    Return "Infeasible Solution"
                Case 1
                    Return "Optimal Solution"
                Case 2
                    Return "Unbounded Solution"
                Case 22
                    Return "AbortDualObjLim (Barrier only)"
                Case 10
                    Return "AbortItLim (Simplex or Barrier)"
                Case 12
                    Return "AbortObjLim (Simplex or Barrier)"
                Case 21
                    Return "AbortPrimObjLim (Barrier only)"
                Case 11
                    Return "AbortTimeLim (Simplex or Barrier)"
                Case 13
                    Return "AbortUser (Simplex or Barrier)"
                Case 32
                    Return "ConflictAbortContradiction (conflict refiner)"
                Case 34
                    Return "ConflictAbortItLim (conflict refiner)"
                Case Else
                    Return "Unknown Error"
            End Select
        End Function

        Friend Shared Sub concertPopulateByColumn() ' _
            '( _
            'ByVal model As ILOG.Concert.IMPModeler, _
            'ByVal var()() As ILOG.Concert.INumVar, _
            'ByVal rng()() As ILOG.Concert.IRange _
            ')

            'Dim obj As ILOG.Concert.IObjective = model.AddMaximize()

            'rng(0) = New ILOG.Concert.IRange(1) {}
            'rng(0)(0) = model.AddRange(-System.Double.MaxValue, 20.0, "c1")
            'rng(0)(1) = model.AddRange(-System.Double.MaxValue, 30.0, "c2")

            'Dim r0 As ILOG.Concert.IRange = rng(0)(0)
            'Dim r1 As ILOG.Concert.IRange = rng(0)(1)

            'var(0) = New ILOG.Concert.INumVar(2) {}
            'var(0)(0) = model.NumVar(model.Column(obj, 1.0).And(model.Column(r0, -1.0).And(model.Column(r1, 1.0))), 0.0, 40.0, "x1")
            'var(0)(1) = model.NumVar(model.Column(obj, 2.0).And(model.Column(r0, 1.0).And(model.Column(r1, -3.0))), 0.0, System.Double.MaxValue, "x2")
            'var(0)(2) = model.NumVar(model.Column(obj, 3.0).And(model.Column(r0, 1.0).And(model.Column(r1, 1.0))), 0.0, System.Double.MaxValue, "x3")
        End Sub 'PopulateByColumn
    End Class 'Solver_CPLEX

    ''' <summary>
    ''' This module serves as a repository of all CPLEX API calls from its callable library.
    ''' </summary>
    ''' <remarks>
    ''' RKP/01-28-10/v2.3.128
    ''' This module serves as a repository of all CPLEX API calls from its callable library.
    ''' 
    ''' RKP/12-09-10/v2.4.142
    ''' Changes in this module is the reason to move from v2.3 Build 141 to v2.4 Build 142.
    ''' List of changes:
    ''' IBM ILOG CPLEX V12.2 - downloaded by BMOS on 12-08-10.
    ''' This version marked the end of ILM (the ILOG License Manager).
    ''' All software-based licensing restrictions were removed starting with this version.
    ''' However, it is now upon the BMOS group to enforce restrictions (limiting use of CPLEX solver to no. of licenses purchased).
    ''' In order to provide forward-compatibility of CPLEX, all CPLEX declarations have now been changed to "cplex.dll" instead of "cplex121.dll" previously.
    ''' This allows backward and forward compatibility of C-OPT's ability to use CPLEX as a solver.
    ''' As long as the callable library function calls keep working, C-OPT should be able to keep up with CPLEX version changes.
    ''' </remarks>
    Friend Module Wrapper_CPLEX
        'http://www.ieor.berkeley.edu/Labs/ilog_docs/html/refparameterscplex/refparameterscplex6.html
        'stm
        Friend Const CPX As Integer = 1
        Friend Const CPX_PARAM_SCRIND As Integer = 1035
        Friend Const CPX_PARAM_TILIM As Integer = 1039 'RKP/06-04-12/v3.2.170
        Friend Const CPX_PARAM_DATACHECK As Integer = 1056
        Friend Const CPX_PARAM_EPGAP As Integer = 2009 'RKP/06-04-12/v3.2.170
        'MIP solution limit
        Friend Const CPX_PARAM_INTSOLLIM As Integer = 2015 'RKP/06-05-12/v3.2.170
        Friend Const CPX_PARAM_TRELIM As Integer = 2027 'RKP/06-04-12/v3.2.170
        Friend Const CPX_ON As Integer = 1
        Friend Const CPX_MAX As Integer = -1
        Friend Const CPX_INFBOUND As Double = 1.0E+20

        Friend Declare Function lstrlenA Lib "kernel32" _
          (ByVal Ptr As System.Object) As Long

        'CPXENVptr CPXopenCPLEX(int * status_p)

        Friend Declare Function CPXopenCPLEX Lib "cplex.dll" _
            (ByRef status As Integer) As IntPtr

        'int CPXcloseCPLEX(CPXENVptr * env_p)
        Friend Declare Function CPXcloseCPLEX Lib "cplex.dll" _
            (ByRef env As IntPtr) As Integer

        'CPXCCHARptr CPXgeterrorstring(CPXCENVptr env, int errcode, char * buffer_str)
        Friend Declare Function CPXgeterrorstring Lib "cplex.dll" _
            (ByVal env As IntPtr, ByVal errcode As Integer, ByVal buffer As StringBuilder) As Integer

        'CPXLPptr CPXcreateprob(CPXCENVptr env, int * status_p, const char * probname_str)
        Friend Declare Function CPXcreateprob Lib "cplex.dll" _
            (ByVal env As IntPtr, ByRef status As Integer, ByVal name As String) As IntPtr

        'int CPXfreeprob(CPXCENVptr env, CPXLPptr * lp_p)
        Friend Declare Function CPXfreeprob Lib "cplex.dll" _
            (ByVal env As IntPtr, ByRef lp As IntPtr) As Integer

        'int CPXwriteprob(CPXCENVptr env, CPXCLPptr lp, const char * filename_str, const char * filetype_str)
        Friend Declare Function CPXwriteprob Lib "cplex.dll" _
            (ByVal env As IntPtr, ByVal lp As IntPtr, ByVal filename As String, ByVal filetype As String) As Integer

        'int CPXsetintparam(CPXENVptr env, int whichparam, int newvalue)
        Friend Declare Function CPXsetintparam Lib "cplex.dll" _
            (ByVal env As IntPtr, ByVal whichparam As Integer, ByVal newvalue As Integer) As Integer

        'void CPXchgobjsen(CPXCENVptr env, CPXLPptr lp, int maxormin)
        Friend Declare Sub CPXchgobjsen Lib "cplex.dll" _
            (ByVal env As IntPtr, ByVal lp As IntPtr, ByVal maxormin As Integer)

        'int CPXnewcols(CPXCENVptr env, CPXLPptr lp, int ccnt, const double * obj, const double * lb, const double * ub, const char * xctype, char ** colname)
        Friend Declare Function CPXnewcols Lib "cplex.dll" _
            (ByVal env As IntPtr, ByVal lp As IntPtr, ByVal ccnt As Integer, ByVal obj() As Double, ByVal lb() As Double, ByVal ub() As Double, ByVal xctype() As Char, ByVal colname() As String) As Integer

        'int CPXnewrows(CPXCENVptr env, CPXLPptr lp, int rcnt, const double * rhs, const char * sense, const double * rngval, char ** rowname)
        Friend Declare Function CPXnewrows Lib "cplex.dll" _
            (ByVal env As IntPtr, ByVal lp As IntPtr, ByVal rcnt As Integer, ByVal rhs() As Double, ByVal sense() As Char, ByVal rngval() As Double, ByVal rowname() As String) As Integer

        'int CPXaddrows(CPXCENVptr env, CPXLPptr lp, int ccnt, int rcnt, int nzcnt, const double * rhs, const char * sense, const int * rmatbeg, const int * rmatind, const double * rmatval, char ** colname, char ** rowname)
        Friend Declare Function CPXaddrows Lib "cplex.dll" _
            (ByVal env As IntPtr, ByVal lp As IntPtr, ByVal ccnt As Integer, ByVal rcnt As Integer, ByVal nzcnt As Integer, ByVal rhs() As Double, ByVal sense() As Char, ByVal rmatbeg() As Integer, ByVal rmatind() As Integer, ByVal rmatval() As Double, ByVal colname() As String, ByVal rowname() As String) As Integer

        'int CPXaddcols(CPXCENVptr env, CPXLPptr lp, int ccnt, int nzcnt, const double * obj, const int * cmatbeg, const int * cmatind, const double * cmatval, const double * lb, const double * ub, char ** colname)
        'Friend Declare Function CPXaddcols Lib "cplex.dll" _
        '(ByVal env As IntPtr, ByVal lp As IntPtr, ByVal ccnt As Integer, ByVal nzcnt As Integer, ByVal obj() As Double, ByVal cmatbeg() As Integer, ByVal cmatind() As Integer, ByVal cmatval() As Double, ByVal lb() As Double, ByVal ub() As Double, ByVal colname() As String) As Integer
        Friend Declare Function CPXaddcols Lib "cplex.dll" _
            (ByVal env As IntPtr, ByVal lp As IntPtr, ByVal ccnt As Integer, ByVal nzcnt As Integer, ByVal obj() As Double, ByVal cmatbeg() As Integer, ByVal cmatind() As Integer, ByVal cmatval() As Double, ByVal lb() As Double, ByVal ub() As Double, ByVal colname() As String) As Long

        'int CPXchgcoeflist(CPXCENVptr env, CPXLPptr lp, int numcoefs, const int * rowlist, const int * collist, const double * vallist)
        Friend Declare Function CPXchgcoeflist Lib "cplex.dll" _
            (ByVal env As IntPtr, ByVal lp As IntPtr, ByVal numcoefs As Integer, ByVal rowlist() As Integer, ByVal collist() As Integer, ByVal vallist() As Double) As Integer

        'int CPXcopyctype(CPXCENVptr env, CPXLPptr lp, const char * xctype)
        Friend Declare Function CPXcopyctype Lib "cplex.dll" _
            (ByVal env As IntPtr, ByVal lp As IntPtr, ByVal xctype() As Char) As Integer

        'int CPXlpopt(CPXCENVptr env, CPXLPptr lp)
        Friend Declare Function CPXlpopt Lib "cplex.dll" _
            (ByVal env As IntPtr, ByVal lp As IntPtr) As Integer

        'int CPXmipopt(CPXCENVptr env, CPXLPptr lp)
        Friend Declare Function CPXmipopt Lib "cplex.dll" _
            (ByVal env As IntPtr, ByVal lp As IntPtr) As Integer

        'int CPXgetnumrows(CPXCENVptr env, CPXCLPptr lp)
        Friend Declare Function CPXgetnumrows Lib "cplex.dll" _
            (ByVal env As IntPtr, ByVal lp As IntPtr) As Integer

        'int CPXgetnumcols(CPXCENVptr env, CPXCLPptr lp)
        Friend Declare Function CPXgetnumcols Lib "cplex.dll" _
            (ByVal env As IntPtr, ByVal lp As IntPtr) As Integer

        'int CPXsolution(CPXCENVptr env, CPXCLPptr lp, int * lpstat_p, double * objval_p, double * x, double * pi, double * slack, double * dj)
        Friend Declare Function CPXsolution Lib "cplex.dll" _
            (ByVal env As IntPtr, ByVal lp As IntPtr, ByRef lpstat As Integer, ByRef objval As Double, <[In](), Out()> ByVal x() As Double, <[In](), Out()> ByVal pi() As Double, <[In](), Out()> ByVal slack() As Double, <[In](), Out()> ByVal dj() As Double) As Integer

        'int CPXgetx(CPXCENVptr env, CPXCLPptr lp, double * x, int begin, int end)
        Friend Declare Function CPXgetx Lib "cplex.dll" _
            (ByVal env As IntPtr, ByVal lp As IntPtr, <[In](), Out()> ByVal x() As Double, ByVal begin As Integer, ByVal end_ As Integer) As Integer

        'int CPXgetstat(CPXCENVptr env, CPXCLPptr lp)
        Friend Declare Function CPXgetstat Lib "cplex.dll" _
            (ByVal env As IntPtr, ByVal lp As IntPtr) As Integer

        'CPXCHARptr CPXgetstatstring(CPXCENVptr env, int statind, char * buffer_str)
        Friend Declare Function CPXgetstatstring Lib "cplex.dll" _
            (ByVal env As IntPtr, ByVal statind As Integer, ByVal buffer As StringBuilder) As Integer

        'int CPXgetitcnt(CPXCENVptr env, CPXCLPptr lp)
        Friend Declare Function CPXgetitcnt Lib "cplex.dll" _
            (ByVal env As IntPtr, ByVal lp As IntPtr) As Integer

        'CPXCCHARptr CPXversion(CPXCENVptr env)
        Friend Declare Function CPXversion Lib "cplex.dll" _
            (ByVal env As IntPtr) As Integer

        'CPXXversion(CPXCENVptr env)
        Friend Declare Function CPXXversion Lib "cplex.dll" _
            (ByVal env As IntPtr) As Integer

        'int CPXsolwrite(CPXCENVptr env, CPXCLPptr lp, const char * filename_str)
        Friend Declare Function CPXsolwrite Lib "cplex.dll" _
            (ByVal env As IntPtr, ByVal lp As IntPtr, ByVal filename As String) As Integer

        'int CPXsolwritesolnpool(CPXCENVptr env, CPXCLPptr lp, int soln, const char * filename_str)
        Friend Declare Function CPXsolwritesolnpool Lib "cplex.dll" _
            (ByVal env As IntPtr, ByVal lp As IntPtr, ByVal soln As Integer, ByVal filename As String) As Integer

        'int CPXsolwritesolnpoolall(CPXCENVptr env, CPXCLPptr lp, const char * filename_str)
        Friend Declare Function CPXsolwritesolnpoolall Lib "cplex.dll" _
            (ByVal env As IntPtr, ByVal lp As IntPtr, ByVal filename As String) As Integer

        'int CPXclpwrite(CPXCENVptr env, CPXCLPptr lp, const char * filename_str)
        Friend Declare Function CPXclpwrite Lib "cplex.dll" _
            (ByVal env As IntPtr, ByVal lp As IntPtr, ByVal filename As String) As Integer

        'int CPXrefineconflict(CPXCENVptr env, CPXLPptr lp, int * confnumrows_p, int * confnumcols_p)
        Friend Declare Function CPXrefineconflict Lib "cplex.dll" _
            (ByVal env As IntPtr, ByVal lp As IntPtr, ByRef confnumrows As Integer, ByRef confnumcols As Integer) As Integer

        'int CPXrefineconflictext(CPXCENVptr env, CPXLPptr lp, int grpcnt, int concnt, const double * grppref, const int * grpbeg, const int * grpind, const char * grptype)
        Friend Declare Function CPXrefineconflictext Lib "cplex.dll" _
            (ByVal env As IntPtr, ByVal lp As IntPtr, ByVal grpcnt As Integer, ByVal concnt As Integer, ByVal grppref() As Double, ByVal grpbeg() As Integer, ByVal grpind() As Integer, ByVal grptype As String) As Integer

        'int CPXpopulate(CPXCENVptr env, CPXLPptr lp)
        'The routine CPXpopulate generates multiple solutions to a mixed integer programming (MIP) problem.
        Friend Declare Function CPXpopulate Lib "cplex.dll" _
            (ByVal env As IntPtr, ByVal lp As IntPtr) As Integer

        'int CPXgetobjval(CPXCENVptr env, CPXCLPptr lp, double * objval_p)
        Friend Declare Function CPXgetobjval Lib "cplex.dll" _
            (ByVal env As IntPtr, ByVal lp As IntPtr, ByVal objval() As Double) As Integer

        'int CPXgetdj(CPXCENVptr env, CPXCLPptr lp, double * dj, int begin, int end)
        Friend Declare Function CPXgetdj Lib "cplex.dll" _
            (ByVal env As IntPtr, ByVal lp As IntPtr, ByVal dj() As Double, ByVal begin As Integer, ByVal end_ As Integer) As Integer

        'int CPXgetslack(CPXCENVptr env, CPXCLPptr lp, double * slack, int begin, int end)
        Friend Declare Function CPXgetslack Lib "cplex.dll" _
            (ByVal env As IntPtr, ByVal lp As IntPtr, ByVal slack() As Double, ByVal begin As Integer, ByVal end_ As Integer) As Integer

        'int CPXgetpi(CPXCENVptr env, CPXCLPptr lp, double * pi, int begin, int end)
        Friend Declare Function CPXgetpi Lib "cplex.dll" _
            (ByVal env As IntPtr, ByVal lp As IntPtr, ByVal slack() As Double, ByVal begin As Integer, ByVal end_ As Integer) As Integer

        'int CPXgetmipitcnt(CPXCENVptr env, CPXCLPptr lp)
        Friend Declare Function CPXgetmipitcnt Lib "cplex.dll" _
            (ByVal env As IntPtr, ByVal lp As IntPtr) As Integer

        'int CPXmbasewrite(CPXCENVptr env, CPXCLPptr lp, const char * filename_str)
        Friend Declare Function CPXmbasewrite Lib "cplex.dll" _
            (ByVal env As IntPtr, ByVal lp As IntPtr, ByVal filename As String) As Integer

        'int CPXwritemipstarts(CPXCENVptr env, CPXCLPptr lp, const char * filename_str, int begin, int end)
        Friend Declare Function CPXwritemipstarts Lib "cplex.dll" _
            (ByVal env As IntPtr, ByVal lp As IntPtr, ByVal filename As String, ByVal begin As Integer, ByVal end_ As Integer) As Integer

        'int CPXwriteparam(CPXCENVptr env, const char * filename_str)
        Friend Declare Function CPXwriteparam Lib "cplex.dll" _
            (ByVal env As IntPtr, ByVal filename As String) As Integer

        'int CPXsetdblparam(CPXENVptr env, int whichparam, double newvalue)
        'The routine CPXsetdblparam sets the value of a CPLEX parameter of type double.
        'status = CPXsetdblparam (env, CPX_PARAM_TILIM, 1000.0);
        'Friend Declare Function CPXsetdblparam Lib "cplex.dll" _
        '    (ByVal env As IntPtr, ByVal lp As IntPtr, ByVal whichparam As Integer, ByVal newvalue As Double) As Integer

        'int CPXsetdblparam(CPXENVptr env, int whichparam, int newvalue)
        Friend Declare Function CPXsetdblparam Lib "cplex.dll" _
            (ByVal env As IntPtr, ByVal whichparam As Integer, ByVal newvalue As Double) As Integer

        'RKP/01-06-12/v3.0.157
        'Allows CPLEX to read in a .MPS file.
        'http://publib.boulder.ibm.com/infocenter/cosinfoc/v12r4/index.jsp?topic=%2Filog.odms.cplex.help%2Frefcallablelibrary%2Fhtml%2Ffunctions%2FCPXreadcopyprob.html
        'int CPXreadcopyprob(CPXCENVptr env, CPXLPptr lp, const char * filename_str, const char * filetype_str)
        Friend Declare Function CPXreadcopyprob Lib "cplex.dll" _
            (ByVal env As IntPtr, ByVal lp As IntPtr, ByVal filename As String, ByVal filetype As String) As Integer

        'RKP/06-07-12/v3.2.170
        'int CPXgetsolnpoolnumsolns(CPXCENVptr env, CPXCLPptr lp)
        Friend Declare Function CPXgetsolnpoolnumsolns Lib "cplex.dll" _
            (ByVal env As IntPtr, ByVal lp As IntPtr) As Integer

        Friend Function cplexGetErrorString(ByVal env As IntPtr, ByVal errcode As Integer) As String
            Dim status As Integer
            Dim buffer As New StringBuilder(4096)

            status = CPXgeterrorstring(env, errcode, buffer)
            Return buffer.ToString()
        End Function

    End Module 'Wrapper_CPLEX
End Namespace