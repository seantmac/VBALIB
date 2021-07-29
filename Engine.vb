Imports System
Imports System.Collections.Generic
Imports System.Text
Imports System.Data.SqlClient
Imports System.Reflection
Imports System.IO
Imports System.Configuration
Imports System.Threading ' Needed for Parallel.For
Imports EntLib
Imports EntLib.COPT

Namespace COPT
    <Microsoft.VisualBasic.ComClass()> _
    Public Class Engine
        Private _startTime As Integer
        Private _solver As String
        Private _currentDb As EntLib.COPT.DAAB '= New EntLib.COPT.DAAB()
        Private _lastSql As String
        Private _dataSet_MiscParams As DataSet
        Private _miscParams As DataRow
        Private mlElmCounter As Integer
        Private _matGenOK As Boolean
        'Private _lpProbInfo As Solver.typLPprobInfo
        Private _statusLevel As Long
        Private _currentSession As EntLib.COPT.Session 'RKP/04-05-10/v2.3.132
        Private _databaseName As String 'RKP/08-02-10/v2.3.134
        Private _hasLinkedTables As Boolean = False 'RKP/08-04-11/v3.0.149

        'RKP/02-21-10/v2.3.131
        'Determines whether to use tsys* tables or qsys* queries
        'for LoadSolverArrays().
        Public _srcRow As String = "tsysRow"
        Public _srcCol As String = "tsysCol"
        Public _srcMtx As String = "tsysMtx"

        'RKP/06-11-10/v2.3.133
        'Converted arrays to Public to allow MatGen to populate arrays that rely on dtMtx.
        Public array_dobj() As Double = Nothing
        Public array_dclo() As Double = Nothing
        Public array_dcup() As Double = Nothing
        Public array_mbeg() As Integer = Nothing
        Public array_midx() As Integer = Nothing
        Public array_mval() As Double = Nothing
        Public array_drhs() As Double = Nothing
        Public array_mcnt() As Integer = Nothing
        Public array_rtyp() As Char = Nothing
        Public array_ctyp() As Char = Nothing
        Public array_initValues() As Double = Nothing 'RKP/02-15-10/v2.3.130
        Public array_colNames() As String = Nothing
        Public array_rowNames() As String = Nothing

        'RKP/07-30-10/v2.3.134
        Private _isMIP As Boolean

        'RKP/02-21-10/v2.3.131
        Private _dtRow As DataTable
        Private _dtCol As DataTable
        Private _dtMtx As DataTable

        Public Const COPT_MISC_TABLE_NAME As String = "qsysMiscParams"
        Public Const COPT_PARM_TABLE_NAME As String = "qsysMiscParams"
        Public Const COPT_RNNM_TABLE_NAME As String = "qsysMiscParams"
        Public Const COPT_RSLT_TABLE_NAME As String = "tsysModelResultFiles" 'RKP/04-01-08

        Public Const LP_NAME As String = "NAME"
        Public Const LP_OBJSENSE As String = "OBJSENSE"
        Public Const LP_SPACE5 As String = "     "
        Public Const LP_MAX As String = "MAX"
        Public Const LP_MIN As String = "MIN"
        Public Const LP_COL_HEAD As String = "COLUMNS"
        Public Const LP_ROW_HEAD As String = "ROWS"
        Public Const LP_RHS_HEAD As String = "RHS"
        Public Const LP_BND_HEAD As String = "BOUNDS"
        Public Const LP_END_HEAD As String = "ENDATA"
        Public Const LP_OBJ_TYPE As String = "N"
        Public Const LP_OBJ_SENSE_MAX As Byte = 1
        Public Const LP_OBJ_SENSE_MIN As Byte = 2

        Public Enum problemRunType
            problemTypeContinuous = 1 'C
            problemTypeBinary = 2 'B
            problemTypeInteger = 4 'I
            problemTypeSemiContinuous = 8 'S
            problemTypeSemiInteger = 16 'N
        End Enum

        Structure typCoeffType
            Dim lngID As Integer
            Dim intActive As Long
            Dim strType As String
            Dim strColType As String
            Dim lngColID As Integer
            Dim strRowType As String
            Dim lngRowID As Integer
            Dim strRecSet As String
            Dim strCoeffFld As String
        End Structure

        Structure typColType
            Dim lngID As Integer
            Dim intActive As Long
            Dim strType As String
            Dim strDesc As String
            Dim strTable As String
            Dim strRecSet As String
            Dim strPrefix As String
            Dim strColDescFld As String
            Dim strSOS As String
            Dim strSOSMarkerFld As String
            Dim intFREE As Long
            Dim intINT As Long
            Dim intBIN As Long
            Dim strOBJFld As String
            Dim strLOFld As String
            Dim strUPFld As String
            Dim intClassCount As Long
            Dim strClasses() As String
        End Structure

        Structure typDatType
            Dim lngID As Integer
            Dim intActive As Long
            Dim intMaster As Long
            Dim strType As String
            Dim strDesc As String
            Dim strTable As String
            Dim intClassCount As Long
            Dim strClasses() As String
            Dim intFieldCount As Long
        End Structure

        Structure typMPSLP
            Dim strFileName As String
            Dim strProblemName As String
            Dim intOBJSense As Long
            Dim strOBJRowName As String
            Dim strRowRS As String
            Dim strColRS As String
            Dim strCoeRS As String
        End Structure

        Structure typRowType
            Dim lngID As Integer
            Dim intActive As Long
            Dim strType As String
            Dim strDesc As String
            Dim strTable As String
            Dim strRecSet As String
            Dim strPrefix As String
            Dim strRowDescFld As String
            Dim strSense As String
            Dim strRHSFld As String
            Dim intClassCount As Long
            Dim strClasses() As String
        End Structure

        Private _solutionStatus As String = "UNKNOWN" 'RKP/03-02-10/v2.3.132 - Start off with "Unknown".
        Private _commonSolutionStatus As String = "UNKNOWN" 'RKP/04-01-10/v2.3.132
        Private _commonSolutionStatusCode As Integer = -1 'RKP/04-01-10/v2.3.132
        Private _solutionStatusCode As Integer = -1 'RKP/04-01-10/v2.3.132
        Private _solutionRows As Integer
        Private _solutionColumns As Integer
        Private _solutionNonZeros As Integer
        Private _solutionObj As Double
        Private _workDir As String = "C:\C-OPT"
        Private _solverName As String
        Private _solverVersion As String 'RKP/09-22-09
        Private _solutionTime As String
        Private _solutionIterations As Integer
        Private _switches() As String
        Private _progress As String 'RKP/09-26-09 - Used to update frmRibbon with progress.
        Private _cancel As Boolean = False  'RKP/09-26-09 - Used to abort RunAll.
        Private _solutionBadResults As String
        Private _solutionInfeasibilities As String
        'Private _problemType As String = "" 'RKP/01-26-10/v2.3.127 - Continuous, MIP, etc.
        Private _problemType As problemRunType 'RKP/01-27-10/v2.3.127 - Continuous, MIP, etc.
        Private _timeStamp As String = "" 'RKP/01-27-10/v2.3.127 - Stores a timestamp value for each run so that all output files have the same timestamp.
        Private _sessionGUID As String '= GenUtils.GenerateGUID()
        Private _runType As GenUtils.RunType  'RKP/12-14-10/v2.4.143

        Public Sub New()
            _currentDb = New EntLib.COPT.DAAB()
            GetMiscParams()
            _sessionGUID = GenUtils.GenerateGUID()
            '_lpProbInfo = GetProbInfo()
        End Sub

        Public Sub New(ByVal databaseName As String)
            _databaseName = databaseName
            _currentDb = New EntLib.COPT.DAAB(databaseName)
            GetMiscParams()
            _sessionGUID = GenUtils.GenerateGUID()
            '_lpProbInfo = GetProbInfo()
        End Sub

        Public Sub New(ByVal databaseName As String, ByVal workDir As String)
            _databaseName = databaseName
            _currentDb = New EntLib.COPT.DAAB(databaseName)
            _workDir = workDir
            GetMiscParams()
            _sessionGUID = GenUtils.GenerateGUID()
            '_lpProbInfo = GetProbInfo()
        End Sub

        Public Sub New(ByVal databaseName As String, ByVal switches() As String)
            _databaseName = databaseName
            _switches = switches
            If GenUtils.IsSwitchAvailable(switches, "/UseMinSysRes") Then
                If GenUtils.IsSwitchAvailable(switches, "/UseSQLServerSyntax") Then
                    _currentDb = New EntLib.COPT.DAAB(databaseName, switches)
                Else
                    _currentDb = New EntLib.COPT.DAAB(databaseName)
                End If
            Else
                _currentDb = New EntLib.COPT.DAAB(databaseName)
            End If
            '_currentDb = New EntLib.COPT.DAAB(databaseName)
            '_workDir = workDir

            _workDir = GenUtils.GetWorkDir(_switches) 'GenUtils.GetSwitchArgument(_switches, "/WorkDir", 1)
            GetMiscParams()
            _sessionGUID = GenUtils.GenerateGUID()
            '_lpProbInfo = GetProbInfo()
        End Sub

        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub

        Public Function GetConnectionString() As String
            'ConfigurationManager.ConnectionStrings("DB1").ConnectionString

            Return _currentDb.GetConnectionString() '"Connection String"
        End Function

        Public Function RunAll() As Integer

            Dim startTime As Integer = My.Computer.Clock.TickCount
            Dim startRunTime As Integer = My.Computer.Clock.TickCount
            Dim userAuthorized As Boolean = True

            'MsgBox("Wait...")
            'MsgBox("RunAll underway..." & COPTUtilities.Main.FormatTime(_startTime, My.Computer.Clock.TickCount))

            'EntLib.COPT.Log.Log(_workDir, "Status", "Entering RunAll...")
            'Debug.Print("RunAll...")
            'GetMiscParams() MOVE TO SUB NEW --STM

            'RunSysModelQueries("PREPROC")
            'InitModel()
            'GenElmTbls()
            'PopElmTbls()
            'MatGen()
            ''
            'SaveMatrix()
            'Solve()
            'PostSolve()

            Application.DoEvents()

            'EntLib.COPT.Log.Log(GenUtils.GetAppSettings("lastWorkDir"), "----------", "----------")
            EntLib.COPT.Log.Log("--------------------")

            _progress = "C-OPT - RunAll - Started..." & My.Computer.Clock.LocalTime.ToString()
            Debug.Print(_progress)
            Console.WriteLine(_progress)
            'Console.ForegroundColor = ConsoleColor.DarkYellow

            Console.ForegroundColor = ConsoleColor.DarkCyan
            _progress = "0% complete."
            Debug.Print(_progress)
            Console.WriteLine(_progress)
            _progress = "Model initialization is now in progress..."
            Debug.Print(_progress)
            Console.WriteLine(_progress)
            Console.ForegroundColor = ConsoleColor.DarkYellow

            Me.TimeStamp = GenUtils.GetTimeStamp()

            'Debug.Print(GenUtils.StatusLevel.Green.ToString())
            _statusLevel = GenUtils.StatusLevel.Green

            If Cancel Then Return GenUtils.ReturnStatus.Failure

            'RKP/06-21-10/v2.3.133
            _dtRow = Nothing
            _dtCol = Nothing
            _dtMtx = Nothing

            'RKP/12-09-10/v2.4.142
            'This block of code is a check for enforcement of CPLEX licensing (V12.2 and beyond).
            If GenUtils.GetSwitchArgument(_switches, "/Solver", 1).Trim().ToUpper.Equals("CPLEX") Then
                If My.Computer.FileSystem.FileExists(My.Computer.FileSystem.CombinePath(My.Application.Info.DirectoryPath, "CPLEX.DLL")) Then
                    'OK. Continue further processing.
                    '_progress = My.Computer.FileSystem.CombinePath(My.Application.Info.DirectoryPath, "CPLEX.DLL")
                    'Console.WriteLine(My.Computer.FileSystem.CombinePath(My.Application.Info.DirectoryPath, "CPLEX.DLL"))
                    _progress = "User authorized to use CPLEX solver."
                    Console.WriteLine(_progress)
                Else
                    'Can't continue further processing because CPLEX.DLL is missing.
                    'The current user is unauthorized to use CPLEX solver.
                    '_progress = My.Computer.FileSystem.CombinePath(My.Application.Info.DirectoryPath, "CPLEX.DLL")
                    'Console.WriteLine(My.Computer.FileSystem.CombinePath(My.Application.Info.DirectoryPath, "CPLEX.DLL"))
                    _progress = "Not authorized to use ""CPLEX"" solver. Use ""CoinMP"" solver instead."
                    Console.WriteLine(_progress)
                    _solutionStatus = _progress
                    userAuthorized = False
                End If
            End If

            If userAuthorized Then
                If GenUtils.IsSwitchAvailable(_switches, "/SolveOnly") Then
                    'Solve()
                    'PostSolve()
                Else
                    '(Not GenUtils.IsSwitchAvailable(_switches, "/SaveRunOnly")) _
                    If (Not GenUtils.IsSwitchAvailable(_switches, "/PostSolveOnly")) _
                       And _
                       (Not GenUtils.IsSwitchAvailable(_switches, "/GenOutputFilesOnly")) _
                    Then
                        If Not GenUtils.IsSwitchAvailable(_switches, "/SkipInit") Then
                            Try
                                If Cancel Then Return GenUtils.ReturnStatus.Failure
                                RunSysModelQueries("PREPROC")
                                If Cancel Then Return GenUtils.ReturnStatus.Failure
                            Catch ex As Exception
                                'MessageBox.Show(ex.Message)
                                GenUtils.Message(GenUtils.MsgType.Critical, "Engine - RunAll-PreProc", ex.Message)
                            End Try

                            Try
                                If Cancel Then Return GenUtils.ReturnStatus.Failure
                                InitModel()

                                Console.ForegroundColor = ConsoleColor.DarkCyan
                                _progress = "20% complete."
                                Debug.Print(_progress)
                                Console.WriteLine(_progress)
                                _progress = "Generation of element tables is now in progress..."
                                Debug.Print(_progress)
                                Console.WriteLine(_progress)
                                Console.ForegroundColor = ConsoleColor.DarkYellow

                                If Cancel Then Return GenUtils.ReturnStatus.Failure
                            Catch ex As Exception
                                'MessageBox.Show(ex.Message)
                                GenUtils.Message(GenUtils.MsgType.Critical, "Engine - RunAll-InitModel", ex.Message)
                            End Try

                            If Not GenUtils.IsSwitchAvailable(_switches, "/InitOnly") Then
                                Try
                                    If Cancel Then Return GenUtils.ReturnStatus.Failure
                                    GenElmTbls()

                                    Console.ForegroundColor = ConsoleColor.DarkCyan
                                    _progress = "50% complete."
                                    Debug.Print(_progress)
                                    Console.WriteLine(_progress)
                                    _progress = "Population of element tables is now in progress..."
                                    Debug.Print(_progress)
                                    Console.WriteLine(_progress)
                                    Console.ForegroundColor = ConsoleColor.DarkYellow

                                    If Cancel Then Return GenUtils.ReturnStatus.Failure
                                Catch ex As Exception
                                    'MessageBox.Show(ex.Message)
                                    GenUtils.Message(GenUtils.MsgType.Critical, "Engine - RunAll-GenElm", ex.Message)
                                End Try

                                Try
                                    If Cancel Then Return GenUtils.ReturnStatus.Failure
                                    PopElmTbls()

                                    Console.ForegroundColor = ConsoleColor.DarkCyan
                                    _progress = "60% complete."
                                    Debug.Print(_progress)
                                    Console.WriteLine(_progress)
                                    _progress = "Matrix generation is now in progress..."
                                    Debug.Print(_progress)
                                    Console.WriteLine(_progress)
                                    Console.ForegroundColor = ConsoleColor.DarkYellow

                                    If Cancel Then Return GenUtils.ReturnStatus.Failure
                                Catch ex As Exception
                                    'MessageBox.Show(ex.Message)
                                    GenUtils.Message(GenUtils.MsgType.Critical, "Engine - RunAll-PopElm", ex.Message)
                                End Try
                            End If
                        End If
                        If Not GenUtils.IsSwitchAvailable(_switches, "/SkipMatGen") Then
                            If Not GenUtils.IsSwitchAvailable(_switches, "/InitOnly") Then
                                Try
                                    If Cancel Then Return GenUtils.ReturnStatus.Failure

                                    'RKP/09-18-11/v3.0.150
                                    'MatGen returns True, when everything went OK
                                    'MatGen returns False, when something went wrong.
                                    'Cancel is a Property that stops further execution if True, hence "Not MatGen()".
                                    Cancel = Not MatGen()

                                    Console.ForegroundColor = ConsoleColor.DarkCyan
                                    _progress = "80% complete."
                                    Debug.Print(_progress)
                                    Console.WriteLine(_progress)
                                    '_progress = "Solving is now in progress..."
                                    'Debug.Print(_progress)
                                    'Console.WriteLine(_progress)
                                    Console.ForegroundColor = ConsoleColor.DarkYellow

                                    If Cancel Then Return GenUtils.ReturnStatus.Failure
                                Catch ex As Exception
                                    'MessageBox.Show(ex.Message)
                                    GenUtils.Message(GenUtils.MsgType.Critical, "Engine - RunAll-MatGen", ex.Message)
                                End Try
                            End If
                        End If
                    End If
                End If

                If GenUtils.IsSwitchAvailable(_switches, "/SaveMtx") Then
                    If Cancel Then Return GenUtils.ReturnStatus.Failure
                    SaveMatrix()
                    If Cancel Then Return GenUtils.ReturnStatus.Failure
                End If

                If GenUtils.IsSwitchAvailable(_switches, "/PostSolveOnly") _
                   Or _
                   GenUtils.IsSwitchAvailable(_switches, "/GenOutputFilesOnly") _
                Then
                    'Solve()
                    Try
                        If Cancel Then Return GenUtils.ReturnStatus.Failure
                        PostSolve()
                        If Cancel Then Return GenUtils.ReturnStatus.Failure
                    Catch ex As Exception
                        'MessageBox.Show(ex.Message)
                        GenUtils.Message(GenUtils.MsgType.Warning, "Engine - RunAll-PostSolve", ex.Message)

                    End Try

                Else
                    If Not GenUtils.IsSwitchAvailable(_switches, "/InitOnly") Then
                        If Not GenUtils.IsSwitchAvailable(_switches, "/SkipSolve") Then
                            Try
                                If Cancel Then Return GenUtils.ReturnStatus.Failure

                                Console.ForegroundColor = ConsoleColor.DarkCyan
                                _progress = "80% complete."
                                Debug.Print(_progress)
                                Console.WriteLine(_progress)
                                '_progress = "Solving is now in progress..."
                                'Debug.Print(_progress)
                                'Console.WriteLine(_progress)
                                Console.ForegroundColor = ConsoleColor.DarkYellow

                                Solve()

                                Console.ForegroundColor = ConsoleColor.DarkCyan
                                _progress = "90% complete."
                                Debug.Print(_progress)
                                Console.WriteLine(_progress)
                                _progress = "Post solve process is now in progress..."
                                Debug.Print(_progress)
                                Console.WriteLine(_progress)
                                Console.ForegroundColor = ConsoleColor.DarkYellow

                                If Cancel Then Return GenUtils.ReturnStatus.Failure
                            Catch ex As Exception
                                'MessageBox.Show(ex.Message)
                                GenUtils.Message(GenUtils.MsgType.Critical, "Engine - RunAll-Solve", ex.Message)
                            End Try

                            Try
                                If Cancel Then Return GenUtils.ReturnStatus.Failure
                                PostSolve()
                                If Cancel Then Return GenUtils.ReturnStatus.Failure
                            Catch ex As Exception
                                'MessageBox.Show(ex.Message)
                                GenUtils.Message(GenUtils.MsgType.Warning, "Engine - RunAll-PostSolve", ex.Message)
                            End Try
                        End If
                    End If
                End If
            End If 'userAuthorized = True


            Console.ForegroundColor = ConsoleColor.DarkCyan
            _progress = "100% complete."
            Debug.Print(_progress)
            Console.WriteLine(_progress)
            Console.ForegroundColor = ConsoleColor.DarkYellow

            _progress = "C-OPT - RunAll - Finished."
            Debug.Print(_progress)
            Console.WriteLine(_progress)

            EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - RunAll took: ", Space(9) & GenUtils.FormatTime(startTime, My.Computer.Clock.TickCount))
            'EntLib.COPT.Log.Log(_workDir, "Status", "Exiting RunAll...")

            'Ravi_Test()

            'MsgBox("RunAll completed successfully!" & vbNewLine & "The solution file is located at:" & vbNewLine & My.Application.Info.DirectoryPath & "\" & "ModelName" & "___" & "ProblemName" & ".MPS.sol")
        End Function

        Private Sub RunSysModelQueries(ByVal type As String)
            Dim startTime As Integer = My.Computer.Clock.TickCount
            Dim startTimeLoop As Integer = 0
            Dim dt As DataTable
            Dim ctr As Integer
            Dim viewName As String
            Dim viewDefinition As String

            _progress = "C-OPT - Engine - RunSysModelQueries...Started at: " & My.Computer.Clock.LocalTime.ToString()
            Debug.Print(_progress)
            Console.WriteLine(_progress)

            'EntLib.COPT.Log.Log(_workDir, "Status", "Entering RunSysModelQueries...")

            'ArchiveSysModelQueries()

            If _currentDb.IsSQLExpress Then
                _lastSql = "SELECT tsysModelQueries.* " & _
                         "FROM tsysModelQueries " & _
                         "WHERE (LEN(tsysModelQueries.Name)>=1) AND " & _
                         "   (tsysModelQueries.ACTIVE)=1 AND " & _
                         "   (tsysModelQueries.ACTIVE)=1 AND " & _
                         "   (tsysModelQueries.Type='" & type & "') " & _
                         "ORDER BY tsysModelQueries.ID;"
            Else
                _lastSql = "SELECT tsysModelQueries.* " & _
                         "FROM tsysModelQueries " & _
                         "WHERE (LEN(tsysModelQueries.Name)>=1) AND " & _
                         "   (tsysModelQueries.ACTIVE)=True AND " & _
                         "   (tsysModelQueries.ACTIVE)=True AND " & _
                         "   (tsysModelQueries.Type='" & type & "') " & _
                         "ORDER BY tsysModelQueries.ID;"
            End If


            Try
                dt = _currentDb.GetDataSet(_lastSql).Tables(0)
                For ctr = 0 To dt.Rows.Count - 1
                    'doevents()
                    Application.DoEvents()
                    If Cancel Then Exit Sub 'Return GenUtils.ReturnStatus.Failure
                    startTimeLoop = My.Computer.Clock.TickCount

                    viewName = dt.Rows(ctr).Item("Name").ToString
                    viewDefinition = _currentDb.GetViewDefinition(viewName)

                    Try
                        _currentDb.ExecuteNonQuery("UPDATE [tsysModelQueries] SET [SQLTEXT] = '" & viewDefinition & "' WHERE [Name] = '" & viewName & "' AND [Type] = '" & type & "'")
                    Catch ex As Exception
                        _progress = "Error updating database with SQLTEXT - " & viewName & " (" & type & ")"
                        Debug.Print(_progress)
                        Console.WriteLine(_progress)
                        EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - RunSysModelQueries", "#" & ctr + 1 & " - " & dt.Rows(ctr).Item("Name").ToString() & ", Time: " & GenUtils.FormatTime(startTimeLoop, My.Computer.Clock.TickCount))
                        EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - RunSysModelQueries", _progress)
                    End Try
                    Try
                        _currentDb.ExecuteNonQuery("UPDATE [tsysModelQueries] SET [SQLDATE] = '" & My.Computer.Clock.LocalTime.ToString() & "' WHERE [Name] = '" & viewName & "' AND [Type] = '" & type & "'")
                    Catch ex As Exception
                        _progress = "Error updating database with SQLTEXT - " & viewName & " (" & type & ")"
                        Debug.Print(_progress)
                        Console.WriteLine(_progress)
                        EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - RunSysModelQueries", "#" & ctr + 1 & " - " & dt.Rows(ctr).Item("Name").ToString() & ", Time: " & GenUtils.FormatTime(startTimeLoop, My.Computer.Clock.TickCount))
                        EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - RunSysModelQueries", _progress)
                    End Try

                    _lastSql = viewDefinition 'dt.Rows(ctr).Item("SQLTEXT").ToString

                    _progress = "C-OPT - Engine - RunSysModelQueries...#" & ctr + 1 & " - " & dt.Rows(ctr).Item("Name").ToString()
                    Debug.Print(_progress)
                    Console.WriteLine(_progress)

                    Try
                        If _lastSql <> "" Then
                            _currentDb.ExecuteNonQuery(_lastSql)
                            _progress = "C-OPT - Engine - RunSysModelQueries...#" & ctr + 1 & " - " & dt.Rows(ctr).Item("Name").ToString() & ", Time: " & GenUtils.FormatTime(startTimeLoop, My.Computer.Clock.TickCount)
                            Debug.Print(_progress)
                            Console.WriteLine(_progress)
                            EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - RunSysModelQueries", "#" & ctr + 1 & " - " & dt.Rows(ctr).Item("Name").ToString() & ", Time: " & GenUtils.FormatTime(startTimeLoop, My.Computer.Clock.TickCount))
                        Else
                            _progress = "C-OPT - Engine - RunSysModelQueries...#" & ctr + 1 & " - " & viewName & ", Query not found."
                            Debug.Print(_progress)
                            Console.WriteLine(_progress)
                            EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - RunSysModelQueries", "#" & ctr + 1 & " - " & dt.Rows(ctr).Item("Name").ToString() & ", Query not found.")
                        End If
                    Catch ex As Exception
                        'MsgBox(ex.Message)
                        GenUtils.Message(GenUtils.MsgType.Critical, "Engine - RunSysModelQueries", ex.Message)
                    End Try

                Next
            Catch ex As Exception
                'MessageBox.Show(ex.Message)
                GenUtils.Message(GenUtils.MsgType.Critical, "Engine - RunSysModelQueries", ex.Message)
                EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - RunSysModelQueries error: ", Space(13) & ex.Message)
                EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - RunSysModelQueries error: ", Space(13) & _lastSql)
            End Try

            _progress = "C-OPT - Engine - RunSysModelQueries...Ended at: " & My.Computer.Clock.LocalTime.ToString()
            Debug.Print(_progress)
            Console.WriteLine(_progress)

            'EntLib.COPT.Log.Log(_workDir, "Status", "Exiting RunSysModelQueries...")
            EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - RunSysModelQueries took: ", Space(13) & _
                                  GenUtils.FormatTime(startTime, My.Computer.Clock.TickCount))

        End Sub

        Private Sub InitModel()
            'stm
            Dim preclen As Integer
            Dim prerlen As Integer
            Dim i As Integer
            Dim userTables As DataTable
            Dim table As String
            Dim sBPE As String
            Dim startTime As Integer = My.Computer.Clock.TickCount
            Dim startTempTime As Integer = My.Computer.Clock.TickCount
            'startTempTime = My.Computer.Clock.TickCount
            'EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - InitModel took: ", GenUtils.FormatTime(startTime, My.Computer.Clock.TickCount))

            _lastSql = ""

            _progress = "C-OPT Engine - InitModel...started"
            Debug.Print(_progress)
            Console.WriteLine(_progress)
            'EntLib.COPT.Log.Log(_workDir, "Status", "Entering InitModel...")

            startTempTime = My.Computer.Clock.TickCount
            Try
                GetMiscParams()
            Catch ex As Exception
                'MessageBox.Show(ex.Message, "C-OPTEngine")
                GenUtils.Message(GenUtils.MsgType.Critical, "Engine - InitModel", ex.Message)
                EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - InitModel - Error: ", ex.Message)
            End Try

            EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - InitModel - GetMiscParams took: ", Space(6) & GenUtils.FormatTime(startTempTime, My.Computer.Clock.TickCount))

            'preclen = Len(Trim(moRSMiscParams!LPM_COL_ELEMENT_TABLE_PRE))
            'prerlen = Len(Trim(moRSMiscParams!LPM_ROW_ELEMENT_TABLE_PRE))
            preclen = Len(Trim(_miscParams.Item("LPM_COL_ELEMENT_TABLE_PRE").ToString))
            prerlen = Len(Trim(_miscParams.Item("LPM_ROW_ELEMENT_TABLE_PRE").ToString))

            'MsgBox("preclen = " & preclen & "; prerlen = " & prerlen)

            If GenUtils.GetSwitchArgument(_switches, "/UseGetTables", 1).Trim.ToUpper.Equals("MSYSOBJECTS") Then
                userTables = _currentDb.GetTables("")
            Else
                userTables = _currentDb.GetTables()
            End If

            'MsgBox("userTables.Rows.Count = " & userTables.Rows.Count)

            'MsgBox("User table count: " & _currentDb.GetUserTables.Rows.Count)

            For i = 0 To userTables.Rows.Count - 1
                Application.DoEvents()
                If Cancel Then Exit Sub 'Return GenUtils.ReturnStatus.Failure
                'table = userTables.Rows.Item(0).ToString
                table = userTables.Rows(i)(2).ToString

                'If GenUtils.GetSwitchArgument(_switches, "/UseGetTables", 1).Trim.ToUpper.Equals("MSYSOBJECTS") Then
                '    'userTables = _currentDb.GetTables("")
                '    table = userTables.Rows(i).Item(0).ToString()
                'Else
                '    'userTables = _currentDb.GetTables()
                '    table = userTables.Rows(i)(2).ToString
                'End If

                '// Delete all Element Tables //
                If (Left(table, preclen) = Trim(_miscParams.Item("LPM_COL_ELEMENT_TABLE_PRE").ToString)) _
                Or (Left(table, prerlen) = Trim(_miscParams.Item("LPM_ROW_ELEMENT_TABLE_PRE").ToString)) Then
                    'moDAL.DeleteTable(oADOTable.Name) 'delete table
                    'Debug.Print("deleted ... " & oADOTable.Name)
                    'Log(False, "deleted ... " & oADOTable.Name)
                    'Log(False, "deleted ... " & oADOTable.Name)

                    _lastSql = "DROP TABLE " & table
                    startTempTime = My.Computer.Clock.TickCount
                    Try
                        _currentDb.ExecuteNonQuery(_lastSql)
                    Catch ex As Exception
                        'MessageBox.Show(ex.Message, "C-OPTEngine")
                        GenUtils.Message(GenUtils.MsgType.Critical, "Engine - InitModel", ex.Message)
                        EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - InitModel - Error: ", ex.Message)
                        EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - InitModel - Error: ", _lastSql)
                    End Try

                    'EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - InitModel - DROP TABLE " & table & " took: ", GenUtils.FormatTime(startTempTime, My.Computer.Clock.TickCount))
                    'EntLib.COPT.Log.Log(_workDir, "Status", "deleted - " & table)
                End If
            Next

            EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - InitModel - Drop Element Tables: ", Space(5) & GenUtils.FormatTime(startTempTime, My.Computer.Clock.TickCount))

            If CurrentDb.IsSQLExpress Then
                _lastSql = "DELETE FROM " & Trim(_miscParams.Item("LPM_MATRIX_TABLE_NAME").ToString)
            Else
                _lastSql = "DELETE * FROM " & Trim(_miscParams.Item("LPM_MATRIX_TABLE_NAME").ToString)
            End If

            '_lastSql = "TRUNCATE TABLE " & Trim(_miscParams.Item("LPM_MATRIX_TABLE_NAME").ToString)
            startTempTime = My.Computer.Clock.TickCount
            Try
                _currentDb.ExecuteNonQuery(_lastSql)
            Catch ex As Exception
                'MessageBox.Show(ex.Message, "C-OPTEngine")
                GenUtils.Message(GenUtils.MsgType.Critical, "Engine - InitModel", ex.Message)
                EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - InitModel - Error: ", ex.Message)
                EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - InitModel - Error: ", _lastSql)
            End Try

            EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - InitModel - Clean Out NonZeroes: ", Space(5) & GenUtils.FormatTime(startTempTime, My.Computer.Clock.TickCount))

            If CurrentDb.IsSQLExpress Then
                _lastSql = "DELETE FROM " & Trim(_miscParams.Item("LPM_COLUMN_TABLE_NAME").ToString)
            Else
                _lastSql = "DELETE * FROM " & Trim(_miscParams.Item("LPM_COLUMN_TABLE_NAME").ToString)
            End If

            '_lastSql = "TRUNCATE TABLE " & Trim(_miscParams.Item("LPM_COLUMN_TABLE_NAME").ToString)
            startTempTime = My.Computer.Clock.TickCount
            Try
                _currentDb.ExecuteNonQuery(_lastSql)
            Catch ex As Exception
                'MessageBox.Show(ex.Message, "C-OPTEngine")
                GenUtils.Message(GenUtils.MsgType.Critical, "Engine - InitModel", ex.Message)
                EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - InitModel - Error: ", ex.Message)
                EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - InitModel - Error: ", _lastSql)
            End Try

            EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - InitModel - Clean Out Columns: ", Space(7) & GenUtils.FormatTime(startTempTime, My.Computer.Clock.TickCount))

            If CurrentDb.IsSQLExpress Then
                _lastSql = "DELETE FROM " & Trim(_miscParams.Item("LPM_CONSTR_TABLE_NAME").ToString)
            Else
                _lastSql = "DELETE * FROM " & Trim(_miscParams.Item("LPM_CONSTR_TABLE_NAME").ToString)
            End If

            '_lastSql = "TRUNCATE TABLE " & Trim(_miscParams.Item("LPM_CONSTR_TABLE_NAME").ToString)
            startTempTime = My.Computer.Clock.TickCount
            Try
                _currentDb.ExecuteNonQuery(_lastSql)
            Catch ex As Exception
                'MessageBox.Show(ex.Message, "C-OPTEngine")
                GenUtils.Message(GenUtils.MsgType.Critical, "Engine - InitModel", ex.Message)
                EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - InitModel - Error: ", ex.Message)
                EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - InitModel - Error: ", _lastSql)
            End Try

            EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - InitModel - Clean Out Rows: ", Space(10) & GenUtils.FormatTime(startTempTime, My.Computer.Clock.TickCount))

            'RKP/03-31-08
            'As per STM (no longer needed)
            '_lastSql = "DELETE * FROM " & Trim(_miscParams.Item("LPM_CONSTR_TABLE_NAME_IMPORT").ToString)
            '_currentDb.ExecuteNonQuery(_lastSql)

            '_lastSql = "DELETE * FROM " & Trim(_miscParams.Item("LPM_COLUMN_TABLE_NAME_IMPORT").ToString)
            '_currentDb.ExecuteNonQuery(_lastSql)

            'EntLib.COPT.Log.Log(_workDir, "Status", "Exiting InitModel...")

            '*&*STM
            sBPE = CheckDBBlueprint()
            EntLib.COPT.Log.Log(_workDir, sBPE & vbCrLf & "C-OPT Engine - InitModel took: ", Space(22) & GenUtils.FormatTime(startTime, My.Computer.Clock.TickCount))
            '*&*

            Progress = "C-OPT Engine - InitModel...finished."
            Debug.Print(Progress)
            Console.WriteLine(_progress)

            EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - InitModel took: ", Space(22) & GenUtils.FormatTime(startTime, My.Computer.Clock.TickCount))
        End Sub

        Public Sub GetMiscParams()
            'SESSIONIZE ALL SIX LINES BELOW

            'RKP/02-23-09
            _lastSql = "SELECT TOP 1 " & COPT_PARM_TABLE_NAME & ".* " _
                  & vbNewLine _
                  & "FROM " & COPT_PARM_TABLE_NAME '& " WHERE MiscID = 1"

            Try
                _dataSet_MiscParams = _currentDb.GetDataSet(_lastSql) '.Tables(0)
                '_dataSet_MiscParams.Tables(0).Rows(0).Item("LPM_CONSTR_DEF_TABLE_NAME").ToString
                _miscParams = _dataSet_MiscParams.Tables(0).Rows(0)
            Catch ex As Exception
                'MessageBox.Show(ex.Message)
                GenUtils.Message(GenUtils.MsgType.Critical, "Engine - GetMiscParams", ex.Message)
            End Try

            'RKP/08-04-11/v3.0.149
            _hasLinkedTables = False
            If _switches IsNot Nothing Then
                If GenUtils.IsSwitchAvailable(_switches, "/UseMinSysRes") Then
                    If GenUtils.IsSwitchAvailable(_switches, "/UseMSAccessSyntax") Then
                        _hasLinkedTables = True
                    ElseIf GenUtils.IsSwitchAvailable(_switches, "/UseSQLServerSyntax") Then
                        _hasLinkedTables = True
                    End If
                End If
            End If

        End Sub

        Public ReadOnly Property MiscParams() As DataRow
            Get
                Return _miscParams
            End Get
        End Property

        '****************************************************************************************


        Private Sub GenElmTbls()
            Dim startTime As Integer = My.Computer.Clock.TickCount
            Dim dt As DataTable
            Dim i As Integer
            Dim rowType As typRowType = Nothing
            Dim colType As typColType = Nothing
            Dim ret As Long

            'EntLib.COPT.Log.Log(ConfigurationManager.AppSettings.Item("lastWorkDir").ToString(), "C-OPT Message - Warning", msg)

            'EntLib.COPT.Log.Log(ConfigurationManager.AppSettings.Item("lastWorkDir").ToString(), "Debug", "GenElmTbls")

            _progress = "C-OPT Engine - GenElmTbls..."
            Debug.Print(_progress)
            Console.WriteLine(_progress)
            'EntLib.COPT.Log.Log(_workDir, "Status", "Entering GenElmTbls...")

            '// R O W S //
            _lastSql = "SELECT * FROM " & _miscParams.Item("LPM_CONSTR_DEF_TABLE_NAME").ToString  'moRSMiscParams!LPM_CONSTR_DEF_TABLE_NAME
            'Call moDAL.Execute("SELECT * FROM " & moRSMiscParams!LPM_CONSTR_DEF_TABLE_NAME, rs, adCmdText, adOpenDynamic)
            Try
                dt = _currentDb.GetDataSet(_lastSql).Tables(0)
                For i = 0 To dt.Rows.Count - 1
                    Application.DoEvents()
                    If Cancel Then Exit Sub 'Return GenUtils.ReturnStatus.Failure
                    rowType = ReadRowType(CInt(dt.Rows(i).Item("RowTypeID")))
                    If CBool(rowType.intActive) Then
                        ret = CreateRowTable(rowType)
                    End If
                Next
            Catch ex As Exception
                'MessageBox.Show(ex.Message, "C-OPTEngine")
                GenUtils.Message(GenUtils.MsgType.Critical, "Engine - GenElm", ex.Message)
                EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - GenElmTbls - Error: ", ex.Message)
                EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - GenElmTbls - Error: " & rowType.strDesc & " - ", _lastSql)
            End Try

            'rsClone = rs.Clone()
            'rsClone.MoveFirst()
            'While rsClone.EOF = False
            '    DoEvents()
            '    rowType = ReadRowType(rsClone!RowTypeID)
            '    If rowType.intActive Then
            '        i = CreateRowTable(rowType)
            '    End If
            '    rsClone.MoveNext()
            'End While


            ''// C O L U M N S //
            'msLastSQL = "SELECT * FROM " & moRSMiscParams!LPM_COLUMN_DEF_TABLE_NAME
            'Call moDAL.Execute("SELECT * FROM " & moRSMiscParams!LPM_COLUMN_DEF_TABLE_NAME, rs, adCmdText, adOpenDynamic)
            'rsClone = rs.Clone()
            'rsClone.MoveFirst()
            'While rsClone.EOF = False
            '    DoEvents()
            '    colType = ReadColType(rsClone!ColTypeID)
            '    If colType.intActive Then
            '        i = CreateColTable(colType)
            '    End If
            '    rsClone.MoveNext()
            'End While

            '// C O L U M N S //
            _lastSql = "SELECT * FROM " & _miscParams.Item("LPM_COLUMN_DEF_TABLE_NAME").ToString  'moRSMiscParams!LPM_CONSTR_DEF_TABLE_NAME
            Try
                dt = _currentDb.GetDataSet(_lastSql).Tables(0)
                For i = 0 To dt.Rows.Count - 1
                    Application.DoEvents()
                    If Cancel Then Exit Sub 'Return GenUtils.ReturnStatus.Failure
                    colType = ReadColType(CInt(dt.Rows(i).Item("ColTypeID").ToString))
                    If CBool(colType.intActive) Then
                        ret = CreateColTable(colType)
                    End If
                Next
            Catch ex As Exception
                'MessageBox.Show(ex.Message, "C-OPTEngine")
                GenUtils.Message(GenUtils.MsgType.Critical, "Engine - GenElm", ex.Message)
                EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - GenElmTbls - Error: ", ex.Message)
                EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - GenElmTbls - Error: " & colType.strDesc & " - ", _lastSql)
            End Try


            'EntLib.COPT.Log.Log(_workDir, "Status", "Exiting GenElmTbls...")
            EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - GenElmTbls took: ", Space(21) & GenUtils.FormatTime(startTime, My.Computer.Clock.TickCount))
        End Sub

        Private Function ReadRowType(ByVal id As Integer) As typRowType
            Dim dt As DataTable
            Dim i As Long
            Dim ctr As Long
            Dim rowTemp As typRowType = Nothing
            Dim strClassConcat As String = ""
            'Dim ret As Integer
            Dim classCount As Integer = 0

            _lastSql = "SELECT * FROM " & _miscParams.Item("LPM_CONSTR_DEF_TABLE_NAME").ToString & " WHERE RowTypeID = " & id
            Try
                dt = _currentDb.GetDataSet(_lastSql).Tables(0)
                'For i = 0 To dt.Rows.Count - 1
                If dt.Rows.Count > 0 Then
                    If dt.Rows(i).Item("RowTypeID").Equals(id) = True Then
                        With rowTemp
                            .lngID = CInt(dt.Rows(i).Item("RowTypeID"))
                            .intActive = CInt(dt.Rows(i).Item("RowActive"))
                            .strType = CStr(dt.Rows(i).Item("rowType"))
                            .strDesc = CStr(dt.Rows(i).Item("RowTypeDesc"))
                            .strTable = CStr(dt.Rows(i).Item("RowTypeTable"))
                            .strRecSet = CStr(dt.Rows(i).Item("RowTypeRecSet"))
                            .strPrefix = CStr(dt.Rows(i).Item("RowTypePrefix"))
                            .strRowDescFld = CStr(dt.Rows(i).Item("RowDescField"))
                            .strSense = CStr(dt.Rows(i).Item("RowTypeSNS"))
                            .strRHSFld = CStr(dt.Rows(i).Item("RHSField"))
                            .intClassCount = CountRowClasses(rowTemp.strType)    '// Count the Classes

                            ReDim rowTemp.strClasses(rowTemp.intClassCount)

                            '// Loop through the classes
                            For ctr = 1 To rowTemp.intClassCount
                                Application.DoEvents()
                                If Cancel Then Return Nothing 'GenUtils.ReturnStatus.Failure
                                'rowTemp.strClasses(ctr) = rs("R" & ctr)                      '// Load the array
                                rowTemp.strClasses(ctr) = dt.Rows(0).Item("R" & ctr).ToString
                                If Brack(dt.Rows(0).Item("R" & ctr).ToString) = "[]" Then
                                    'do nothing
                                Else
                                    classCount = classCount + 1
                                    strClassConcat = strClassConcat & Brack(dt.Rows(0).Item("R" & ctr).ToString)     '// Make the Concat string
                                End If

                                'strClassConcat = String.Concat(strClassConcat & Brack(dt.Rows(ctr).Item("R" & ctr).ToString))
                                'String.Concat(strClassConcat, Brack(dt.Rows(0).Item("R" & ctr).ToString))

                            Next

                            _lastSql = "UPDATE " & _miscParams.Item("LPM_CONSTR_DEF_TABLE_NAME").ToString & " SET ClassCount = " & classCount & ", ClassConcat = '" & strClassConcat & "'"

                            'rsElements.Tables(0).Rows(i2)(colTemp.strType & "ID") = mlElmCounter
                            dt.Rows(i)("ClassCount") = classCount
                            dt.Rows(i)("ClassConcat") = strClassConcat
                            Try
                                'If GenUtils.IsSwitchAvailable(_switches, "/UseMinSysRes") Then
                                'If GenUtils.IsSwitchAvailable(_switches, "/UseMSAccessSyntax") Then
                                ' GenUtils.UpdateDB(Me.MiscParams.Item("LINKED_ROW_DB_CONN_STR").ToString(), dt, "SELECT * FROM " & _miscParams.Item("LPM_CONSTR_DEF_TABLE_NAME").ToString())
                                'Else
                                '_currentDb.UpdateDataSet(dt, "SELECT * FROM " & _miscParams.Item("LPM_CONSTR_DEF_TABLE_NAME").ToString)
                                'End If
                                'Else
                                _currentDb.UpdateDataSet(dt, "SELECT * FROM " & _miscParams.Item("LPM_CONSTR_DEF_TABLE_NAME").ToString())
                                'End If
                            Catch ex As Exception
                                'MessageBox.Show(ex.Message, "C-OPTEngine")
                                GenUtils.Message(GenUtils.MsgType.Critical, "Engine - ReadRowType", ex.Message)
                                EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - ReadRowType - Error: ", ex.Message)
                                EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - ReadRowType - Error: ", _lastSql)
                            End Try


                            'EntLib.Log.Log("C-OPTEngine.Engine.ReadRowType", _lastSql)
                            'ret = _currentDb.ExecuteNonQuery(_lastSql)

                        End With
                    End If
                    'Next
                End If
            Catch ex As Exception
                'MessageBox.Show(ex.Message, "C-OPTEngine")
                GenUtils.Message(GenUtils.MsgType.Critical, "Engine - ReadRowType", ex.Message)
                EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - ReadRowType - Error: ", ex.Message)
                EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - ReadRowType - Error: ", _lastSql)
            End Try

            Return rowTemp
        End Function

        Private Function ReadColType(ByVal id As Integer) As typColType
            Dim dt As DataTable
            Dim i As Long = 0
            Dim ctr As Long
            Dim colTemp As typColType = Nothing
            Dim strClassConcat As String = ""
            Dim ret As Integer

            _lastSql = "SELECT * FROM " & _miscParams.Item("LPM_COLUMN_DEF_TABLE_NAME").ToString & " WHERE ColTypeID = " & id
            Try
                dt = _currentDb.GetDataSet(_lastSql).Tables(0)
                'For i = 0 To dt.Rows.Count - 1
                If dt.Rows.Count > 0 Then
                    If dt.Rows(i).Item("ColTypeID").Equals(id) Then
                        With colTemp
                            .lngID = CInt(dt.Rows(i).Item("ColTypeID"))
                            .intActive = CInt(dt.Rows(i).Item("ColActive"))
                            .strType = CStr(dt.Rows(i).Item("colType"))
                            .strDesc = CStr(dt.Rows(i).Item("ColTypeDesc"))
                            .strTable = CStr(dt.Rows(i).Item("ColTypeTable"))
                            .strRecSet = CStr(dt.Rows(i).Item("ColTypeRecSet"))
                            .strPrefix = CStr(dt.Rows(i).Item("ColTypePrefix"))
                            .strColDescFld = CStr(dt.Rows(i).Item("ColDescField"))
                            .strSOS = CStr(dt.Rows(i).Item("SOSTYPE"))
                            .strSOSMarkerFld = CStr(dt.Rows(i).Item("SOSMarkerField"))
                            .intFREE = CInt(dt.Rows(i).Item("BNDFree"))
                            .intINT = CInt(dt.Rows(i).Item("BNDInteger"))
                            .intBIN = CInt(dt.Rows(i).Item("BNDBinary"))
                            .strOBJFld = CStr(dt.Rows(i).Item("OBJField"))
                            .strLOFld = CStr(dt.Rows(i).Item("BNDLoField"))
                            .strUPFld = CStr(dt.Rows(i).Item("BNDUpField"))
                            .intClassCount = CountColClasses(colTemp.strType)    '// Count the Classes

                            ReDim colTemp.strClasses(colTemp.intClassCount)

                            '// Loop through the classes
                            For ctr = 1 To colTemp.intClassCount
                                Application.DoEvents()
                                If Cancel Then Return Nothing 'GenUtils.ReturnStatus.Failure
                                'rowTemp.strClasses(ctr) = rs("C" & ctr)                      '// Load the array
                                colTemp.strClasses(ctr) = dt.Rows(i).Item("C" & ctr).ToString
                                'strClassConcat = strClassConcat & Brack(rs("R" & ctr))     '// Make the Concat string
                                'strClassConcat = String.Concat(strClassConcat & Brack(dt.Rows(ctr).Item("R" & ctr).ToString))
                                'String.Concat(strClassConcat, Brack(dt.Rows(i).Item("C" & ctr).ToString))
                                strClassConcat = strClassConcat & Brack(dt.Rows(i).Item("C" & ctr).ToString)
                            Next

                            _lastSql = "UPDATE " & _miscParams.Item("LPM_COLUMN_DEF_TABLE_NAME").ToString & " SET ClassCount = " & colTemp.intClassCount & ", ClassConcat = '" & strClassConcat & "' WHERE ColTypeID = " & dt.Rows(i).Item("ColTypeID").ToString()
                            ret = _currentDb.ExecuteNonQuery(_lastSql)

                        End With
                    End If
                    'Next
                End If
            Catch ex As Exception
                'MessageBox.Show(ex.Message, "C-OPTEngine")
                GenUtils.Message(GenUtils.MsgType.Critical, "Engine - ReadColType", ex.Message)
                EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - ReadColType - Error: ", ex.Message)
                EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - ReadColType - Error: ", _lastSql)
            End Try

            Return colTemp
        End Function

        Private Function Brack(ByVal strIn As String) As String

            Return "[" & strIn & "]"

        End Function

        Private Function CreateRowTable(ByVal rowTemp As typRowType) As Long
            Dim i As Long
            Dim fieldLength As Long = CInt(_miscParams.Item("LPM_TEXT_LENGTH").ToString)
            Dim tableName As String = _miscParams.Item("LPM_ROW_ELEMENT_TABLE_PRE").ToString & rowTemp.strTable
            Dim fieldKey As String = rowTemp.strType & "Key"
            Dim field1 As String = rowTemp.strType & "ID"
            Dim field2 As String = rowTemp.strType & "Code"
            Dim field3 As String = rowTemp.strType & "Desc"

            Dim sql As String = ""
            Dim indexColumns As String = ""
            Dim retValue As Integer = 0

            'RKP/07-27-11/v2.5.147
            If CurrentDb.IsSQLExpress Then
                _lastSql = "CREATE TABLE [dbo].[" & tableName & "] (" & vbNewLine
                _lastSql = _lastSql & "[" & fieldKey & "] [INT] IDENTITY(1,1) NOT NULL," & vbNewLine
                _lastSql = _lastSql & "[" & field1 & "] [INT] NULL," & vbNewLine
                _lastSql = _lastSql & "[" & field2 & "] [NVARCHAR](" & fieldLength & ") NULL," & vbNewLine
                _lastSql = _lastSql & "[" & field3 & "] [NVARCHAR](" & fieldLength & ") NULL" & vbNewLine

                indexColumns = ""
                For i = 1 To rowTemp.intClassCount
                    Application.DoEvents()
                    If Cancel Then Return GenUtils.ReturnStatus.Failure
                    If rowTemp.strClasses(i).ToString <> "" Then
                        _lastSql = _lastSql & " ," & "[" & rowTemp.strClasses(i) & "]" & " [NVARCHAR](" & fieldLength & ") NULL "
                        indexColumns = indexColumns & "[" & rowTemp.strClasses(i) & "],"
                    End If

                    'RKP/07-24-08
                    'Indexes are now in place
                Next
                indexColumns = Mid(indexColumns, 1, Len(indexColumns) - 1)

                _lastSql = _lastSql & " ,[RHS] [FLOAT] NULL, [SENSE] [NVARCHAR](1) NULL, [ACTIVITY] [FLOAT] NULL, [SHADOW] [FLOAT] NULL, " & vbNewLine
                'CONSTRAINT [PK_tmtxCol01_PROD] PRIMARY KEY CLUSTERED 
                _lastSql = _lastSql & "CONSTRAINT [PK_" & tableName & "] PRIMARY KEY CLUSTERED " & vbNewLine
                _lastSql = _lastSql & "(" & vbNewLine
                _lastSql = _lastSql & " [" & fieldKey & "] ASC " & vbNewLine
                _lastSql = _lastSql & ")WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY] " & vbNewLine
                _lastSql = _lastSql & ") ON [PRIMARY] " & vbNewLine

                Try
                    _progress = _lastSql.Substring(0, _lastSql.IndexOf("(") - 1).Trim 'Left(_lastSql, 50)
                    Debug.Print(_lastSql.Substring(0, _lastSql.IndexOf("(") - 1).Trim)
                    Console.WriteLine(_lastSql.Substring(0, _lastSql.IndexOf("(") - 1).Trim)
                    'RKP/08-04-11/v3.0.149
                    retValue = _currentDb.ExecuteNonQuery(_lastSql, _switches, _hasLinkedTables)

                    sql = "CREATE UNIQUE NONCLUSTERED INDEX [idxRow_" & tableName & "] ON [dbo].[" & tableName & "] (" & indexColumns & ")"
                    Try
                        _currentDb.ExecuteNonQuery(sql)
                    Catch ex As Exception
                        'MessageBox.Show(ex.Message, "C-OPTEngine")
                        GenUtils.Message(GenUtils.MsgType.Critical, "Engine - CreateRowTable", ex.Message)
                        EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - CreateRowTable - Error: ", ex.Message)
                        EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - CreateRowTable - Error: ", _lastSql)
                    End Try
                    Return CInt(retValue)
                Catch ex As Exception
                    'MessageBox.Show(ex.Message)
                    Return Nothing
                End Try
            Else 'CurrentDb.IsSQLExpress = False
                '_lastSql = "CREATE TABLE " & tableName & " (" & rowTemp.strType & "Key, " & rowTemp.strType & "ID, INT, Code VARCHAR(" & _miscParams.Item("LPM_TEXT_LENGTH").ToString & "), [Desc] VARCHAR(" & _miscParams.Item("LPM_TEXT_LENGTH").ToString & ")"
                _lastSql = "CREATE TABLE " & tableName & " (" & fieldKey & " IDENTITY PRIMARY KEY, " & field1 & " INT, " & field2 & " VARCHAR(" & fieldLength & "), " & field3 & " VARCHAR(" & fieldLength & ")"

                indexColumns = ""
                For i = 1 To rowTemp.intClassCount
                    Application.DoEvents()
                    If Cancel Then Return GenUtils.ReturnStatus.Failure
                    If rowTemp.strClasses(i).ToString <> "" Then
                        _lastSql = _lastSql & " ," & rowTemp.strClasses(i) & " VARCHAR(" & fieldLength & ")"
                        indexColumns = indexColumns & "[" & rowTemp.strClasses(i) & "],"
                    End If

                    'RKP/07-24-08
                    'Indexes are now in place
                Next
                indexColumns = Mid(indexColumns, 1, Len(indexColumns) - 1)

                _lastSql = _lastSql & " ,RHS FLOAT, SENSE VARCHAR(1), ACTIVITY FLOAT, SHADOW FLOAT)"
                Try
                    _progress = _lastSql.Substring(0, _lastSql.IndexOf("(") - 1).Trim 'Left(_lastSql, 50)
                    Debug.Print(_lastSql.Substring(0, _lastSql.IndexOf("(") - 1).Trim)
                    Console.WriteLine(_lastSql.Substring(0, _lastSql.IndexOf("(") - 1).Trim)
                    retValue = _currentDb.ExecuteNonQuery(_lastSql)

                    sql = "CREATE INDEX [idxRow_" & tableName & "] ON [" & tableName & "] (" & indexColumns & ")"
                    Try
                        _currentDb.ExecuteNonQuery(sql)
                    Catch ex As Exception
                        'MessageBox.Show(ex.Message, "C-OPTEngine")
                        GenUtils.Message(GenUtils.MsgType.Critical, "Engine - CreateRowTable", ex.Message)
                        EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - CreateRowTable - Error: ", ex.Message)
                        EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - CreateRowTable - Error: ", _lastSql)
                    End Try


                    Return CInt(retValue)
                Catch ex As Exception
                    'MessageBox.Show(ex.Message)
                    Return Nothing
                End Try
            End If


        End Function

        Private Function CreateColTable(ByVal colTemp As typColType) As Long
            Dim i As Long
            Dim fieldLength As Long = CInt(_miscParams.Item("LPM_TEXT_LENGTH").ToString)
            Dim tableName As String = "[" & _miscParams.Item("LPM_COL_ELEMENT_TABLE_PRE").ToString & colTemp.strTable & "]"
            Dim fieldKey As String = "[" & colTemp.strType & "Key]"
            Dim field1 As String = "[" & colTemp.strType & "ID]"
            Dim field2 As String = "[" & colTemp.strType & "Code]"
            Dim field3 As String = "[" & colTemp.strType & "Desc]"
            Dim field4 As String = "[SOSTYPE]"
            Dim field5 As String = "[SOSMARKER]"
            Dim field6 As String = "[FREE]"
            Dim field7 As String = "[INTGR]"
            Dim field8 As String = "[BINRY]"
            Dim field9 As String = "[OBJ]"
            Dim field10 As String = "[LO]"
            Dim field11 As String = "[UP]"
            Dim field12 As String = "[ACTIVITY]"
            Dim field13 As String = "[DJ]"
            Dim field14 As String = "[STATUS]"

            Dim sql As String = ""
            Dim indexColumns As String = ""
            Dim retValue As Integer = 0

            'RKP/07-27-11/v2.5.147
            If CurrentDb.IsSQLExpress Then
                _lastSql = "CREATE TABLE [dbo]." & tableName & " (" & vbNewLine
                _lastSql = _lastSql & fieldKey & " [INT] IDENTITY(1,1) NOT NULL," & vbNewLine
                _lastSql = _lastSql & field1 & " [INT] NULL," & vbNewLine
                _lastSql = _lastSql & field2 & " [NVARCHAR](" & fieldLength & ") NULL," & vbNewLine
                _lastSql = _lastSql & field3 & " [NVARCHAR](" & fieldLength & ") NULL," & vbNewLine
                _lastSql = _lastSql & field4 & " [INT] NULL," & vbNewLine
                _lastSql = _lastSql & field5 & " [NVARCHAR](" & fieldLength & ") NULL," & vbNewLine
                _lastSql = _lastSql & field6 & " [BIT] NULL," & vbNewLine
                _lastSql = _lastSql & field7 & " [BIT] NULL," & vbNewLine
                _lastSql = _lastSql & field8 & " [BIT] NULL " & vbNewLine


                For i = 1 To colTemp.intClassCount
                    Application.DoEvents()
                    If Cancel Then Return GenUtils.ReturnStatus.Failure
                    If colTemp.strClasses(i) <> "" Then
                        _lastSql = _lastSql & " ,[" & colTemp.strClasses(i) & "] NVARCHAR(" & fieldLength & ") NULL "
                        'RKP/07-24-08
                        'Indexes are now in place.
                        indexColumns = indexColumns & "[" & colTemp.strClasses(i) & "],"
                    End If
                Next
                indexColumns = Mid(indexColumns, 1, Len(indexColumns) - 1)

                _lastSql = _lastSql & ", " & field9 & " [FLOAT] NULL, " & field10 & " [FLOAT] NULL, " & field11 & " [FLOAT] NULL, " & _
                    field12 & " [FLOAT] NULL, " & field13 & " [FLOAT] NULL, " & field14 & " [NVARCHAR](" & fieldLength & ") NULL, "
                'CONSTRAINT [PK_tmtxCol01_PROD] PRIMARY KEY CLUSTERED 
                _lastSql = _lastSql & "CONSTRAINT [PK_" & Left(Mid(tableName, 2), Len(Mid(tableName, 2)) - 1) & "] PRIMARY KEY CLUSTERED " & vbNewLine
                _lastSql = _lastSql & "(" & vbNewLine
                _lastSql = _lastSql & " " & fieldKey & " ASC " & vbNewLine
                _lastSql = _lastSql & ")WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY] " & vbNewLine
                _lastSql = _lastSql & ") ON [PRIMARY] " & vbNewLine

                Try
                    _progress = _lastSql.Substring(0, _lastSql.IndexOf("(") - 1).Trim 'Left(_lastSql, 50)
                    Debug.Print(_lastSql.Substring(0, _lastSql.IndexOf("(") - 1).Trim)
                    Console.WriteLine(_lastSql.Substring(0, _lastSql.IndexOf("(") - 1).Trim)
                    retValue = _currentDb.ExecuteNonQuery(_lastSql)

                    sql = "CREATE UNIQUE NONCLUSTERED INDEX [idxCol_" & Left(Mid(tableName, 2), Len(Mid(tableName, 2)) - 1) & "] ON [dbo]." & tableName & " (" & indexColumns & ")"
                    Try
                        _currentDb.ExecuteNonQuery(sql)
                    Catch ex As Exception
                        'MessageBox.Show(ex.Message, "C-OPTEngine")
                        GenUtils.Message(GenUtils.MsgType.Critical, "Engine - CreateRowTable", ex.Message)
                        EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - CreateRowTable - Error: ", ex.Message)
                        EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - CreateRowTable - Error: ", _lastSql)
                    End Try
                    Return CInt(retValue)
                Catch ex As Exception
                    'MessageBox.Show(ex.Message)
                    Return Nothing
                End Try

            Else 'IsSQLExpress = False

                '_lastSql = "CREATE TABLE " & tableName & " (ID INT, Code VARCHAR(" & _miscParams.Item("LPM_TEXT_LENGTH").ToString & "), [Desc] VARCHAR(" & _miscParams.Item("LPM_TEXT_LENGTH").ToString & ")"
                '_lastSql = _lastSql & " , SOSTYPE INT, SOSMARKER VARCHAR(" & _miscParams.Item("LPM_TEXT_LENGTH").ToString & "), FREE BIT, INTGR BIT, BINRY BIT"
                _lastSql = "CREATE TABLE " & tableName & _
                    " (" & fieldKey & " IDENTITY PRIMARY KEY, " & field1 & " INT, " & field2 & " VARCHAR(" & fieldLength & "), " & field3 & " VARCHAR(" & fieldLength & "), " & _
                    field4 & " INT, " & field5 & " VARCHAR(" & fieldLength & "), " & field6 & " BIT, " & field7 & " BIT, " & _
                    field8 & " BIT"

                For i = 1 To colTemp.intClassCount
                    Application.DoEvents()
                    If Cancel Then Return GenUtils.ReturnStatus.Failure
                    If colTemp.strClasses(i) <> "" Then
                        _lastSql = _lastSql & " ,[" & colTemp.strClasses(i) & "] VARCHAR(" & fieldLength & ")"
                        'RKP/07-24-08
                        'Indexes are now in place.
                        indexColumns = indexColumns & "[" & colTemp.strClasses(i) & "],"
                    End If
                Next
                indexColumns = Mid(indexColumns, 1, Len(indexColumns) - 1)

                _lastSql = _lastSql & ", " & field9 & " FLOAT, " & field10 & " FLOAT, " & field11 & " FLOAT, " & _
                    field12 & " FLOAT, " & field13 & " FLOAT, " & field14 & " VARCHAR(" & fieldLength & "))"
                Try
                    _progress = _lastSql.Substring(0, _lastSql.IndexOf("(") - 1).Trim 'Left(_lastSql, 50)
                    Debug.Print(_lastSql.Substring(0, _lastSql.IndexOf("(") - 1).Trim)
                    Console.WriteLine(_lastSql.Substring(0, _lastSql.IndexOf("(") - 1).Trim)
                    retValue = _currentDb.ExecuteNonQuery(_lastSql)

                    sql = "CREATE INDEX [idxCol_" & Left(Mid(tableName, 2), Len(Mid(tableName, 2)) - 1) & "] ON " & tableName & " (" & indexColumns & ")"
                    _currentDb.ExecuteNonQuery(sql)

                    Return CInt(retValue)
                Catch ex As Exception
                    'MessageBox.Show(ex.Message)
                    Return Nothing
                End Try
            End If

        End Function

        Private Function CountRowClasses(ByVal strRowType As String) As Long
            Dim intCount As Long = 0
            Dim i As Long
            Dim strRowClassField As String = ""
            Dim dt As DataTable = Nothing


            For i = 1 To CInt(_miscParams.Item("LPM_MAX_ELEMENT_CLASSES").ToString)
                Application.DoEvents()
                If Cancel Then Return GenUtils.ReturnStatus.Failure
                _lastSql = ""
                strRowClassField = "R" & CStr(i)
                _lastSql = "SELECT " & strRowClassField & " FROM " & _miscParams.Item("LPM_CONSTR_DEF_TABLE_NAME").ToString _
                         & " WHERE RowType = " & "'" & strRowType & "'"
                dt = _currentDb.GetDataSet(_lastSql).Tables(0)
                'If dt.Rows.Count = 0 Then Exit For
                If dt.Rows(0).Item(0).ToString = "" Then Exit For
                intCount = CInt(intCount + 1)
            Next

            Return intCount
        End Function

        Private Function CountColClasses(ByVal strColType As String) As Long
            Dim intCount As Long = 0
            Dim i As Long
            Dim strColClassField As String = ""
            Dim dt As DataTable = Nothing


            For i = 1 To CInt(_miscParams.Item("LPM_MAX_ELEMENT_CLASSES").ToString)
                Application.DoEvents()
                If Cancel Then Return GenUtils.ReturnStatus.Failure
                _lastSql = ""
                strColClassField = "C" & CStr(i)
                _lastSql = "SELECT " & strColClassField & " FROM " & _miscParams.Item("LPM_COLUMN_DEF_TABLE_NAME").ToString _
                         & " WHERE ColType = " & "'" & strColType & "'"
                dt = _currentDb.GetDataSet(_lastSql).Tables(0)
                If dt.Rows(0).Item(0).ToString = "" Then Exit For
                intCount = CInt(intCount + 1)
            Next

            Return intCount
        End Function

        Private Sub PopElmTbls()
            '//================================================================================//
            '/|   FUNCTION: PopElmTbls
            '/| PARAMETERS: -NONE-
            '/|    RETURNS: True on Success and False by default or Failure
            '/|    PURPOSE: Populate the tables that store individual LP vectors and rows
            '/|      USAGE: i= PopElmTbls()
            '/|         BY: Sean
            '/|       DATE: 01/07/2004
            '/|    HISTORY: 3/27/97  Added feature to generate the append queries
            '/|                      automatically, then run them.
            '/|             6/28/10  Added feature to accomodate large model processing
            '/|                      with the /UseMinSysRes switch.
            '/|                      This switch will not touch COLUMNS, but will, however, populate dtRow.
            '/|                      dtCol relies on qsysMtxCol, which will get loaded after MatGen.
            '//================================================================================//

            Dim startTime As Integer = My.Computer.Clock.TickCount
            Dim startTimeLoop As Integer = 0
            Dim rsDefs As DataTable
            'Dim rs As DataTable
            Dim rsElements As DataSet 'DataTable
            Dim colTemp As typColType
            Dim intIgNull As Boolean = True
            Dim s As String
            Dim sTemp As String
            Dim sTemp2 As String
            Dim mlElmCounter As Long
            Dim i As Long
            Dim i2 As Long
            Dim rowCount As Integer
            Dim insertCommand As System.Data.Common.DbCommand = Nothing
            Dim updateCommand As System.Data.Common.DbCommand = Nothing
            Dim deleteCommand As System.Data.Common.DbCommand = Nothing
            Dim rowTemp As typRowType
            Dim selectStmnt As String 'RKP/06-16-10/v2.3.133
            Dim dt As DataTable 'RKP/06-16-10/v2.3.133
            Dim dc As DataColumn 'RKP/06-21-10/v2.3.133
            Dim linkedDB As Boolean = False 'RKP/08-02-11/v3.0.148
            'Dim dtBulkCopy As DataTable 'RKP/08-04-11/v3.0.149

            _progress = "C-OPT Engine - PopElmTbls..."
            Debug.Print(_progress)
            Console.WriteLine(_progress)
            'EntLib.COPT.Log.Log(_workDir, "Status", "Entering PopElmTbls...")

            If GenUtils.IsSwitchAvailable(_switches, "/UseMinSysRes") Then
                If GenUtils.IsSwitchAvailable(_switches, "/UseSQLServerSyntax") Then
                    linkedDB = True
                ElseIf GenUtils.IsSwitchAvailable(_switches, "/UseMSAccessSyntax") Then
                    linkedDB = True
                End If
            End If

            '// C O L U M N S //
            '
            _lastSql = "SELECT * FROM " & _miscParams.Item("LPM_COLUMN_DEF_TABLE_NAME").ToString
            'Call moDAL.Execute("SELECT * FROM " & moRSMiscParams!LPM_COLUMN_DEF_TABLE_NAME, rsDefs, adCmdText, adOpenDynamic)
            rsDefs = _currentDb.GetDataSet(_lastSql).Tables(0)
            'rsDefs.MoveFirst()
            intIgNull = True

            s = "----- C O L U M N S -----"
            _progress = s
            Debug.Print(s)
            Console.WriteLine(s)
            'Log(False, s)
            EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - PopElmTbls - ", s)
            mlElmCounter = 0
            'If gFormRunFlag Then gtxtDest.Value = gtxtDest.Value & s & vbNewLine
            _dtCol = Nothing 'New DataTable
            'While rsDefs.EOF = False
            For i = 0 To CInt(rsDefs.Rows.Count - 1)
                Application.DoEvents()
                If Cancel Then Exit Sub 'Return GenUtils.ReturnStatus.Failure
                startTimeLoop = My.Computer.Clock.TickCount
                colTemp = ReadColType(CInt(rsDefs.Rows(i).Item("ColTypeID").ToString))
                If CBool(colTemp.intActive) Then

                    s = colTemp.strType
                    _progress = s
                    Debug.Print(_progress)
                    Console.WriteLine(_progress)

                    _lastSql = ""
                    _lastSql = MakePopColQry(colTemp)
                    s = colTemp.strType
                    sTemp2 = _lastSql
                    '_progress = s & " ... "
                    'Debug.Print(_progress)
                    'Console.WriteLine(_progress)
                    'Log(False, s & " ... ")
                    'EntLib.COPT.Log.Log(_workDir, "Status", s & " ... ")
                    'If gFormRunFlag Then gtxtDest.Value = gtxtDest.Value & s & vbNewLine

                    If GenUtils.IsSwitchAvailable(_switches, "/LogVerbose") Then
                        EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - PopElmTbls - ColType = '" & s & "' - SQL: ", vbNewLine & _lastSql)
                    End If

                    'rs = moDAL.oADOConn.Execute(strSQL, lRet)
                    'rs = _currentDb.GetDataSet(_lastSql).Tables(0)

                    'If CurrentDb.IsSQLExpress Then
                    'dtBulkCopy = _currentDb.GetDataTable(_lastSql)
                    '_currentDb.UpdateDataSet(dtBulkCopy, _miscParams.Item("LPM_COL_ELEMENT_TABLE_PRE").ToString() & colTemp.strTable, _switches, True)
                    'Else
                    Try
                        rowCount = _currentDb.ExecuteNonQuery(_lastSql)
                        'rowCount = _currentDb.ExecuteNonQuery(_lastSql, _switches, linkedDB)
                    Catch ex As Exception
                        EntLib.COPT.Log.Log("Error - PopElmTbls - " & colTemp.strType & " - " & vbNewLine & ex.Message)
                        Exit Sub
                    End Try
                    'End If

                    '*&*
                    'Call moDAL.Execute(strSQL, rs, adCmdStoredProc, adExecuteNoRecords) '// EXECUTE THE INSERT/APPEND
                    '-------------
                    'Call moDAL.Execute(strSQL, rs, adCmdText, adOpenDynamic) '// EXECUTE THE INSERT/APPEND
                    '-----------
                    '// EXECUTE THE INSERT/APPEND
                    'Call moDAL.Execute(strSQL, rs, adExecuteNoRecords + adCmdText, adOpenForwardOnly)
                    '---------
                    'moDAL.oADOConn.Execute strSQL, lRet, adExecuteNoRecords + adCmdText
                    'Call WaitSecs(6)
                    'For i = 0 To moDAL.oADOConn.Properties.Count - 1
                    'Debug.Print moDAL.oADOConn.Properties.Item(i).Name & " - " & _
                    '            moDAL.oADOConn.Properties.Item(i).Value
                    'Log False, moDAL.oADOConn.Properties.Item(i).Name & " - " & _
                    '            moDAL.oADOConn.Properties.Item(i).Value
                    'Next i
                    'MsgBox colTemp.strType & " Done with INSERT.  Continue? ", vbOKOnly

                    'update the counter field for each element
                    '
                    _lastSql = "SELECT * FROM " & _miscParams.Item("LPM_COL_ELEMENT_TABLE_PRE").ToString & colTemp.strTable
                    'Call moDAL.Execute(strSQL, rsElements, adCmdText, adOpenDynamic)
                    'rsElements.MoveFirst()

                    _progress = s & " ... " & "   "

                    _progress = s
                    Debug.Print(_progress)
                    Console.WriteLine(_progress)

                    rsElements = _currentDb.GetDataSet(_lastSql) '.Tables(0)

                    '_progress = s & " ... " & "   " & rsElements.Tables(0).Rows.Count & ", Time: " & GenUtils.FormatTime(startTimeLoop, My.Computer.Clock.TickCount)
                    'Debug.Print(_progress)
                    'Console.WriteLine(_progress)
                    ''Log(False, "   " & rsElements.Rows.Count)
                    'EntLib.COPT.Log.Log(_workDir, "   PopElm-COLS", Space(14 - Len(s)) & s & "..." & _
                    '                      Space(15 - Len(rsElements.Tables(0).Rows.Count)) & _
                    '                      rsElements.Tables(0).Rows.Count & Space(8) & _
                    '                      GenUtils.FormatTime(startTime, My.Computer.Clock.TickCount))

                    'Do While rsElements.EOF = False
                    '    DoEvents()
                    '    mlElmCounter = mlElmCounter + 1
                    '    rsElements(colTemp.strType & "ID") = mlElmCounter
                    '    rsElements.UpdateBatch()
                    '    rsElements.MoveNext()
                    'Loop


                    'mlElmCounter = 0
                    For i2 = 0 To CInt(rsElements.Tables(0).Rows.Count - 1)
                        Application.DoEvents()
                        If Cancel Then Exit Sub 'Return GenUtils.ReturnStatus.Failure

                        'mlElmCounter = 0
                        mlElmCounter = mlElmCounter + 1
                        'rsElements(colTemp.strType & "ID") = mlElmCounter
                        'rsElements.Columns(colTemp.strType & "ID") = mlElmCounter
                        rsElements.Tables(0).Rows(i2)(colTemp.strType & "ID") = mlElmCounter
                        'rsElements.AcceptChanges()
                        'rsElements.Tables(0).AcceptChanges()

                        '_lastSql = "UPDATE " & _miscParams.Item("LPM_COL_ELEMENT_TABLE_PRE").ToString & colTemp.strTable & " SET " & colTemp.strType & "ID" & " = " & mlElmCounter & " WHERE "

                    Next
                    _lastSql = "SELECT * FROM " & _miscParams.Item("LPM_COL_ELEMENT_TABLE_PRE").ToString & colTemp.strTable

                    'If GenUtils.IsSwitchAvailable(_switches, "/UseMinSysRes") Then
                    '    If GenUtils.IsSwitchAvailable(_switches, "/UseMSAccessSyntax") Then
                    '        rowCount = GenUtils.UpdateDB(Me.MiscParams.Item("LINKED_COL_DB_CONN_STR").ToString(), rsElements.Tables(0), _lastSql)
                    '    Else
                    '        rowCount = _currentDb.UpdateDataSet(rsElements, _lastSql)
                    '    End If
                    'Else
                    '    rowCount = _currentDb.UpdateDataSet(rsElements, _lastSql)
                    'End If
                    rowCount = _currentDb.UpdateDataSet(rsElements.Tables(0), _lastSql)

                    'rowCount = _currentDb.UpdateDataSet(rsElements, _miscParams.Item("LPM_COL_ELEMENT_TABLE_PRE").ToString & colTemp.strTable, insertCommand, updateCommand, deleteCommand)

                    'Now make the query to transfer the records to sysCOL and Run It
                    sTemp = "[" & _miscParams.Item("LPM_COL_ELEMENT_TABLE_PRE").ToString & colTemp.strTable & "]"

                    selectStmnt = "SELECT " & vbNewLine & _
                                sTemp & "." & s & "ID AS [ColID]," & vbNewLine & _
                                sTemp & "." & s & "Code AS [COL]," & vbNewLine & _
                                sTemp & "." & s & "Desc AS [DESC]," & vbNewLine & _
                                sTemp & ".SOSTYPE AS [SOSTYPE], " & vbNewLine & _
                                sTemp & ".SOSMARKER AS [SOSMARKER], " & vbNewLine & _
                                sTemp & ".FREE AS [FREE], " & vbNewLine & _
                                sTemp & ".INTGR AS [INTGR], " & vbNewLine & _
                                sTemp & ".BINRY AS [BINRY], " & vbNewLine & _
                                sTemp & ".OBJ AS [OBJ], " & vbNewLine & _
                                sTemp & ".LO AS [LO], " & vbNewLine & _
                                sTemp & ".UP AS [UP] " & vbNewLine & _
                             "FROM " & sTemp & " "

                    _lastSql = "INSERT INTO " & _miscParams.Item("LPM_COLUMN_TABLE_NAME").ToString & " " & vbNewLine & _
                                "   ( [ColID], [COL], [DESC], SOSTYPE, SOSMARKER, FREE, INTGR, BINRY, OBJ, LO, UP )" & vbNewLine & _
                                selectStmnt

                    '"SELECT " & vbNewLine & _
                    '   sTemp & "." & s & "ID," & vbNewLine & _
                    '   sTemp & "." & s & "Code," & vbNewLine & _
                    '   sTemp & "." & s & "Desc," & vbNewLine & _
                    '   sTemp & ".SOSTYPE, " & vbNewLine & _
                    '   sTemp & ".SOSMARKER, " & vbNewLine & _
                    '   sTemp & ".FREE, " & vbNewLine & _
                    '   sTemp & ".INTGR, " & vbNewLine & _
                    '   sTemp & ".BINRY, " & vbNewLine & _
                    '   sTemp & ".OBJ, " & vbNewLine & _
                    '   sTemp & ".LO, " & vbNewLine & _
                    '   sTemp & ".UP " & vbNewLine & _
                    '"FROM " & sTemp & " " & vbNewLine & _
                    '";"
                    'msLastSQL = strSQL
                    'Call moDAL.Execute(strSQL, rs, adCmdText, adOpenDynamic) '// EXECUTE THE INSERT/APPEND TO sysCOL
                    'Debug.Print(_lastSql)
                    Try
                        If GenUtils.IsSwitchAvailable(_switches, "/UseMinSysRes") Then
                            If GenUtils.IsSwitchAvailable(_switches, "/UseMSAccessSyntax") Then
                                _lastSql = "INSERT INTO " & _
                                                "[" & Me.MiscParams.Item("LINKED_COL_DB_PATH").ToString() & "]." & _
                                                "[" & Me.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString() & "] " & vbNewLine & _
                                                "   ( [ColID], [COL], [DESC], SOSTYPE, SOSMARKER, FREE, INTGR, BINRY, OBJ, LO, UP )" & vbNewLine & _
                                                selectStmnt
                                '"SELECT " & vbNewLine & _
                                'sTemp & "." & s & "ID," & vbNewLine & _
                                'sTemp & "." & s & "Code," & vbNewLine & _
                                'sTemp & "." & s & "Desc," & vbNewLine & _
                                'sTemp & ".SOSTYPE, " & vbNewLine & _
                                'sTemp & ".SOSMARKER, " & vbNewLine & _
                                'sTemp & ".FREE, " & vbNewLine & _
                                'sTemp & ".INTGR, " & vbNewLine & _
                                'sTemp & ".BINRY, " & vbNewLine & _
                                'sTemp & ".OBJ, " & vbNewLine & _
                                'sTemp & ".LO, " & vbNewLine & _
                                'sTemp & ".UP " & vbNewLine & _
                                '"FROM " & _
                                'sTemp & " "

                                'RKP/07-31-11/v2.5.147
                            ElseIf GenUtils.IsSwitchAvailable(_switches, "/UseSQLServerSyntax") Then
                                'do nothing
                                'because _lastSql is still correct.
                            Else
                                'do nothing
                                'because _lastSql is still correct.
                            End If

                            'Try
                            '    dt = _currentDb.GetDataTable(selectStmnt, False)
                            '    _dtCol.Merge(dt)
                            'Catch ex As Exception
                            '    EntLib.COPT.Log.Log("Error - PopElmTbls - Build dtCol - " & colTemp.strType & " - " & vbNewLine & ex.Message)
                            '    EntLib.COPT.Log.Log(selectStmnt)
                            '    Exit Sub
                            'End Try
                        Else
                            'rowCount = _currentDb.ExecuteNonQuery(_lastSql)
                        End If


                        rowCount = _currentDb.ExecuteNonQuery(_lastSql)

                        'RKP/08-02-11/v3.0.148
                        'rowCount = _currentDb.ExecuteNonQuery(_lastSql, _switches, linkedDB)
                    Catch ex As Exception
                        EntLib.COPT.Log.Log("Error - PopElmTbls - " & colTemp.strType & " - " & vbNewLine & ex.Message)
                        EntLib.COPT.Log.Log(_lastSql)
                        Exit Sub
                    Finally
                        If GenUtils.IsSwitchAvailable(_switches, "/LogVerbose") Then
                            EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - PopElmTbls - ColType = '" & s & "' - SQL: ", vbNewLine & _lastSql)
                        End If
                    End Try


                    _progress = s & " ... " & "   " & rsElements.Tables(0).Rows.Count & ", Time: " & GenUtils.FormatTime(startTimeLoop, My.Computer.Clock.TickCount)
                    Debug.Print(_progress)
                    Console.WriteLine(_progress)
                    'Log(False, "   " & rsElements.Rows.Count)
                    EntLib.COPT.Log.Log(_workDir, "   PopElm-COLS", Space(14 - Len(s)) & s & "..." & _
                                          Space(15 - Len(rsElements.Tables(0).Rows.Count)) & _
                                          rsElements.Tables(0).Rows.Count & Space(8) & _
                                          GenUtils.FormatTime(startTime, My.Computer.Clock.TickCount))
                End If
            Next

            If GenUtils.IsSwitchAvailable(_switches, "/UseMinSysRes") Then
                If Not dtCol Is Nothing Then
                    dc = New DataColumn
                    dc.ColumnName = "ACTIVITY"
                    dc.DataType = System.Type.GetType("System.Double")
                    _dtCol.Columns.Add(dc)

                    dc = New DataColumn
                    dc.ColumnName = "DJ"
                    dc.DataType = System.Type.GetType("System.Double")
                    _dtCol.Columns.Add(dc)

                    dc = New DataColumn
                    dc.ColumnName = "STATUS"
                    dc.DataType = System.Type.GetType("System.String")
                    _dtCol.Columns.Add(dc)
                End If
            End If

            _progress = "COLS:    " & mlElmCounter
            Debug.Print(_progress)
            Console.WriteLine(_progress)
            EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - PopElmTbls - ", "COLS:  " & mlElmCounter)

            '// R O W S //
            _lastSql = "SELECT * FROM " & _miscParams("LPM_CONSTR_DEF_TABLE_NAME").ToString()
            rsDefs = _currentDb.GetDataSet(_lastSql).Tables(0)
            intIgNull = True
            s = "----- R O W S -----"
            _progress = s
            Debug.Print(s)
            Console.WriteLine(s)
            EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - PopElmTbls - ", s)
            mlElmCounter = 0
            _dtRow = New DataTable
            For i = 0 To CInt(rsDefs.Rows.Count - 1)
                Application.DoEvents()
                If Cancel Then Exit Sub 'Return GenUtils.ReturnStatus.Failure
                'rowTemp = ReadRowType(rsDefs!RowTypeID)
                startTimeLoop = My.Computer.Clock.TickCount
                rowTemp = ReadRowType(CInt(rsDefs.Rows(i).Item("RowTypeID").ToString))
                If CBool(rowTemp.intActive) Then
                    _lastSql = MakePopRowQry(rowTemp)
                    s = rowTemp.strType

                    _progress = s
                    Debug.Print(_progress)
                    Console.WriteLine(_progress)

                    '_progress = s
                    'Debug.Print(s & " ... ")
                    'Console.WriteLine(s & " ... ")
                    'EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - PopElmTbls - ", s & " ... ")
                    'rs = _currentDb.GetDataSet(_lastSql).Tables(0)
                    rowCount = _currentDb.ExecuteNonQuery(_lastSql)

                    '_progress = s
                    'Debug.Print(_progress)
                    'Console.WriteLine(_progress)

                    _lastSql = "SELECT * FROM " & _miscParams.Item("LPM_ROW_ELEMENT_TABLE_PRE").ToString() & rowTemp.strTable
                    rsElements = _currentDb.GetDataSet(_lastSql)

                    '_progress = s & " ... " & "   " & rsElements.Tables(0).Rows.Count & ", Time: " & GenUtils.FormatTime(startTimeLoop, My.Computer.Clock.TickCount)
                    'Debug.Print(_progress)
                    'Console.WriteLine(_progress)
                    ''EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - PopElmTbls - ", "   " & rsElements.Tables(0).Rows.Count)
                    'EntLib.COPT.Log.Log(_workDir, "   PopElm-ROWS", Space(14 - Len(s)) & s & "..." & _
                    '                      Space(15 - Len(rsElements.Tables(0).Rows.Count)) & _
                    '                      rsElements.Tables(0).Rows.Count & Space(8) & _
                    '                      GenUtils.FormatTime(startTime, My.Computer.Clock.TickCount))

                    'mlElmCounter = 0
                    For i2 = 0 To CInt(rsElements.Tables(0).Rows.Count - 1)
                        Application.DoEvents()
                        If Cancel Then Exit Sub 'Return GenUtils.ReturnStatus.Failure
                        'mlElmCounter = 0
                        mlElmCounter = mlElmCounter + 1
                        'rsElements(colTemp.strType & "ID") = mlElmCounter
                        'rsElements.Columns(colTemp.strType & "ID") = mlElmCounter
                        rsElements.Tables(0).Rows(i2)(rowTemp.strType & "ID") = mlElmCounter
                        'rsElements.AcceptChanges()
                        'rsElements.Tables(0).AcceptChanges()

                        '_lastSql = "UPDATE " & _miscParams.Item("LPM_COL_ELEMENT_TABLE_PRE").ToString & colTemp.strTable & " SET " & colTemp.strType & "ID" & " = " & mlElmCounter & " WHERE "

                    Next
                    _lastSql = "SELECT * FROM " & _miscParams.Item("LPM_ROW_ELEMENT_TABLE_PRE").ToString() & rowTemp.strTable


                    rowCount = _currentDb.UpdateDataSet(rsElements.Tables(0), _lastSql)
                    If _currentDb.LastErrorNo <> 0 Then
                        Cancel = True
                        Exit Sub
                    End If


                    'Now make the query to transfer the records to sysCOL and Run It
                    sTemp = "[" & _miscParams.Item("LPM_ROW_ELEMENT_TABLE_PRE").ToString & rowTemp.strTable & "]"
                    selectStmnt = "SELECT " & vbNewLine & _
                                sTemp & "." & s & "ID AS [RowID]," & vbNewLine & _
                                sTemp & "." & s & "Code AS [ROW]," & vbNewLine & _
                                sTemp & "." & s & "Desc AS [DESC]," & vbNewLine & _
                                sTemp & ".RHS AS [RHS], " & vbNewLine & _
                                sTemp & ".SENSE AS [SENSE] " & vbNewLine & _
                             "FROM " & sTemp & " " & vbNewLine '& _
                    _lastSql = "INSERT INTO " & _miscParams.Item("LPM_CONSTR_TABLE_NAME").ToString & " " & vbNewLine & _
                                "   ( [RowID], [ROW], [DESC], RHS, SENSE )" & vbNewLine & _
                                selectStmnt
                    '";"
                    'Call moDAL.Execute(strSQL, rs, adCmdText, adOpenDynamic) '// EXECUTE THE INSERT/APPEND TO sysCOL
                    'rs = _currentDb.GetDataSet(_lastSql).Tables(0)

                    Try
                        If GenUtils.IsSwitchAvailable(_switches, "/UseMinSysRes") Then
                            If GenUtils.IsSwitchAvailable(_switches, "/UseMSAccessSyntax") Then
                                _lastSql = "INSERT INTO " & _
                                            "[" & Me.MiscParams.Item("LINKED_ROW_DB_PATH").ToString() & "]." & _
                                            "[" & _miscParams.Item("LPM_CONSTR_TABLE_NAME").ToString & "] " & vbNewLine & _
                                            "   ( [RowID], [ROW], [DESC], RHS, SENSE )" & vbNewLine & _
                                            selectStmnt
                                '"SELECT " & vbNewLine & _
                                'sTemp & "." & s & "ID," & vbNewLine & _
                                'sTemp & "." & s & "Code," & vbNewLine & _
                                'sTemp & "." & s & "Desc," & vbNewLine & _
                                'sTemp & ".RHS, " & vbNewLine & _
                                'sTemp & ".SENSE " & vbNewLine & _
                                '"FROM " & sTemp & " " & vbNewLine


                            Else
                                'do nothing
                            End If
                            'Try
                            '    dt = _currentDb.GetDataTable(selectStmnt, False)
                            '    If dtRow Is Nothing Then _dtRow = New DataTable
                            '    _dtRow.Merge(dt)
                            'Catch ex As Exception
                            '    EntLib.COPT.Log.Log("Error - PopElmTbls - Build dtRow - " & sTemp & " - " & vbNewLine & ex.Message)
                            '    EntLib.COPT.Log.Log(selectStmnt)
                            '    Exit Sub
                            'End Try
                        Else
                            'do nothing
                        End If

                        Try
                            dt = _currentDb.GetDataTable(selectStmnt, False)
                            If dtRow Is Nothing Then _dtRow = New DataTable
                            _dtRow.Merge(dt)
                        Catch ex As Exception
                            EntLib.COPT.Log.Log("Error - PopElmTbls - Build dtRow - " & sTemp & " - " & vbNewLine & ex.Message)
                            EntLib.COPT.Log.Log(selectStmnt)
                            Exit Sub
                        End Try

                        rowCount = _currentDb.ExecuteNonQuery(_lastSql)
                    Catch ex As Exception
                        EntLib.COPT.Log.Log("Error - PopElmTbls - " & sTemp & " - " & vbNewLine & ex.Message)
                        EntLib.COPT.Log.Log(_lastSql)
                        Exit Sub
                    Finally 'RKP/08-24-11/v3.0.150
                        If GenUtils.IsSwitchAvailable(_switches, "/LogVerbose") Then
                            EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - PopElmTbls - RowType = '" & s & "' - SQL: ", vbNewLine & _lastSql)
                        End If
                    End Try

                    _progress = s & " ... " & "   " & rsElements.Tables(0).Rows.Count & ", Time: " & GenUtils.FormatTime(startTimeLoop, My.Computer.Clock.TickCount)
                    Debug.Print(_progress)
                    Console.WriteLine(_progress)
                    'EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - PopElmTbls - ", "   " & rsElements.Tables(0).Rows.Count)
                    EntLib.COPT.Log.Log(_workDir, "   PopElm-ROWS", Space(14 - Len(s)) & s & "..." & _
                                          Space(15 - Len(rsElements.Tables(0).Rows.Count)) & _
                                          rsElements.Tables(0).Rows.Count & Space(8) & _
                                          GenUtils.FormatTime(startTime, My.Computer.Clock.TickCount))

                End If
            Next

            'If GenUtils.IsSwitchAvailable(_switches, "/UseMinSysRes") Then
            If Not dtRow Is Nothing Then
                dc = New DataColumn
                dc.ColumnName = "ACTIVITY"
                dc.DataType = System.Type.GetType("System.Double")
                _dtRow.Columns.Add(dc)

                dc = New DataColumn
                dc.ColumnName = "SHADOW"
                dc.DataType = System.Type.GetType("System.Double")
                _dtRow.Columns.Add(dc)

                dc = New DataColumn
                dc.ColumnName = "STATUS"
                dc.DataType = System.Type.GetType("System.String")
                _dtRow.Columns.Add(dc)
            End If
            'End If

            _progress = "ROWS:    " & mlElmCounter
            Debug.Print(_progress)
            Console.WriteLine(_progress)
            EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - PopElmTbls - ", "ROWS:  " & mlElmCounter)

            'EntLib.COPT.Log.Log(_workDir, "Status", "Exiting PopElmTbls...")
            EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - PopElmTbls took: ", Space(21) & GenUtils.FormatTime(startTime, My.Computer.Clock.TickCount))
        End Sub

        Private Function MakePopColQry(ByVal coltemp As typColType) As String
            Dim strINS As String, strSQL As String, strSelObj As String, strVectorName As String
            Dim strColClass As String
            Dim j As Long


            If CurrentDb.IsSQLExpress Then
                strSQL = ""
                strINS = ""

                strINS = "INSERT INTO " & _miscParams.Item("LPM_COL_ELEMENT_TABLE_PRE").ToString() & coltemp.strTable _
                         & " ( [" & coltemp.strType & "ID], "
                strSQL = strINS
                strSelObj = ""
                strVectorName = "'C" & coltemp.strPrefix & "' "

                For j = 1 To coltemp.intClassCount
                    Application.DoEvents()
                    If Cancel Then Return GenUtils.ReturnStatus.Failure
                    strSQL = strSQL & Brack(coltemp.strClasses(CInt(j))) & ", "
                    strColClass = coltemp.strClasses(CInt(j))
                    'strSelObj = strSelObj & "CStr(" & Brack(coltemp.strRecSet) & "." & Brack(strColClass) & ") AS " & Brack(strColClass) & ", " & vbNewLine
                    strSelObj = strSelObj & "CONVERT(NVARCHAR(80)," & Brack(coltemp.strRecSet) & "." & Brack(strColClass) & ") AS " & Brack(strColClass) & ", " & vbNewLine
                    strVectorName = strVectorName & "+" & " " & "'" & "_" & "'" & " " & "+" & " " & "CONVERT(NVARCHAR(255)," & Brack(strColClass) & ")" & " "
                Next j

                'If Len(colTemp.strOBJFld) > 0 Then
                strSQL = strSQL & "[OBJ], "
                'End If

                'If Len(colTemp.strLOFld) > 0 Then
                strSQL = strSQL & "[LO], "
                'End If

                'If Len(colTemp.strUPFld) > 0 Then
                strSQL = strSQL & "[UP], " & vbNewLine
                'End If

                'If Len(colTemp.strSOS) > 0 Then
                strSQL = strSQL & "[SOSTYPE], "
                'End If

                'If Len(colTemp.strSOSMarkerFld) > 0 Then
                strSQL = strSQL & "[SOSMARKER], "
                'End If

                'If colTemp.intFREE < 0 Then
                strSQL = strSQL & "[FREE], "
                'End If

                'If colTemp.intINT < 0 Then
                strSQL = strSQL & "[INTGR], "
                'End If

                'If colTemp.intBIN < 0 Then
                strSQL = strSQL & "[BINRY], "
                'End If

                strSQL = strSQL & coltemp.strType & "Code, " & coltemp.strType & "Desc )" & vbNewLine

                strSQL = strSQL & "SELECT DISTINCT" & vbNewLine & "0 AS VECTORID," & vbNewLine
                strSQL = strSQL & strSelObj

                'If Len(colTemp.strOBJFld) > 0 Then
                strSQL = strSQL & coltemp.strRecSet & "." & coltemp.strOBJFld & " AS [OBJ], " & vbNewLine
                'End If

                'If Len(colTemp.strLOFld) > 0 Then
                strSQL = strSQL & coltemp.strRecSet & "." & coltemp.strLOFld & " AS [LO], " & vbNewLine
                'End If

                'If Len(colTemp.strUPFld) > 0 Then
                strSQL = strSQL & coltemp.strRecSet & "." & coltemp.strUPFld & " AS [UP], " & vbNewLine
                'End If

                'If Len(colTemp.strSOS) > 0 Then
                strSQL = strSQL & coltemp.strSOS & " AS [SOSTYPE], " & vbNewLine
                'End If

                'If Len(colTemp.strSOSMarkerFld) > 0 Then
                strSQL = strSQL & "'" & coltemp.strSOSMarkerFld & "'" & " AS [SOSMARKER], " & vbNewLine
                'End If

                'If colTemp.intFREE < 0 Then
                strSQL = strSQL & coltemp.intFREE & " AS [FREE], " & vbNewLine
                'End If

                'If colTemp.intINT < 0 Then
                strSQL = strSQL & coltemp.intINT & " AS [INTGR], " & vbNewLine
                'End If

                'If colTemp.intBIN < 0 Then
                strSQL = strSQL & coltemp.intBIN & " AS [BINRY], " & vbNewLine
                'End If

                strVectorName = " " & strVectorName & " AS [VectorName]," & vbNewLine


                strSQL = strSQL & strVectorName & "Left(" & Brack(coltemp.strColDescFld) & "," & _miscParams.Item("LPM_TEXT_LENGTH").ToString() & ") AS [VectorDesc]" & vbNewLine

                strSQL = strSQL & "FROM " & Brack(coltemp.strRecSet) & vbNewLine
                strSQL = strSQL & ";"

            Else 'CurrentDb.IsSQLExpress = False
                strINS = "INSERT INTO " & _miscParams.Item("LPM_COL_ELEMENT_TABLE_PRE").ToString() & coltemp.strTable _
                         & " ( [" & coltemp.strType & "ID], "
                strSQL = strINS
                strSelObj = ""
                strVectorName = "'C" & coltemp.strPrefix & "' "

                For j = 1 To coltemp.intClassCount
                    Application.DoEvents()
                    If Cancel Then Return GenUtils.ReturnStatus.Failure
                    strSQL = strSQL & Brack(coltemp.strClasses(CInt(j))) & ", "
                    strColClass = coltemp.strClasses(CInt(j))
                    strSelObj = strSelObj & "CStr(" & Brack(coltemp.strRecSet) & "." & Brack(strColClass) & ") AS " & Brack(strColClass) & ", " & vbNewLine
                    strVectorName = strVectorName & "&" & " " & "'" & "_" & "'" & " " & "&" & " " & Brack(strColClass) & " "
                Next j

                'If Len(colTemp.strOBJFld) > 0 Then
                strSQL = strSQL & "[OBJ], "
                'End If

                'If Len(colTemp.strLOFld) > 0 Then
                strSQL = strSQL & "[LO], "
                'End If

                'If Len(colTemp.strUPFld) > 0 Then
                strSQL = strSQL & "[UP], " & vbNewLine
                'End If

                'If Len(colTemp.strSOS) > 0 Then
                strSQL = strSQL & "[SOSTYPE], "
                'End If

                'If Len(colTemp.strSOSMarkerFld) > 0 Then
                strSQL = strSQL & "[SOSMARKER], "
                'End If

                'If colTemp.intFREE < 0 Then
                strSQL = strSQL & "[FREE], "
                'End If

                'If colTemp.intINT < 0 Then
                strSQL = strSQL & "[INTGR], "
                'End If

                'If colTemp.intBIN < 0 Then
                strSQL = strSQL & "[BINRY], "
                'End If

                strSQL = strSQL & coltemp.strType & "Code, " & coltemp.strType & "Desc )" & vbNewLine
                strSQL = strSQL & "SELECT DISTINCT" & vbNewLine & "0 AS VECTORID," & vbNewLine
                strSQL = strSQL & strSelObj

                'If Len(colTemp.strOBJFld) > 0 Then
                strSQL = strSQL & coltemp.strRecSet & "." & coltemp.strOBJFld & " AS [OBJ], " & vbNewLine
                'End If

                'If Len(colTemp.strLOFld) > 0 Then
                strSQL = strSQL & coltemp.strRecSet & "." & coltemp.strLOFld & " AS [LO], " & vbNewLine
                'End If

                'If Len(colTemp.strUPFld) > 0 Then
                strSQL = strSQL & coltemp.strRecSet & "." & coltemp.strUPFld & " AS [UP], " & vbNewLine
                'End If

                'If Len(colTemp.strSOS) > 0 Then
                strSQL = strSQL & coltemp.strSOS & " AS [SOSTYPE], " & vbNewLine
                'End If

                'If Len(colTemp.strSOSMarkerFld) > 0 Then
                strSQL = strSQL & "'" & coltemp.strSOSMarkerFld & "'" & " AS [SOSMARKER], " & vbNewLine
                'End If

                'If colTemp.intFREE < 0 Then
                strSQL = strSQL & coltemp.intFREE & " AS [FREE], " & vbNewLine
                'End If

                'If colTemp.intINT < 0 Then
                strSQL = strSQL & coltemp.intINT & " AS [INTGR], " & vbNewLine
                'End If

                'If colTemp.intBIN < 0 Then
                strSQL = strSQL & coltemp.intBIN & " AS [BINRY], " & vbNewLine
                'End If

                strVectorName = " " & strVectorName & " AS [VectorName]," & vbNewLine


                strSQL = strSQL & strVectorName & "Left(" & Brack(coltemp.strColDescFld) & "," & _miscParams.Item("LPM_TEXT_LENGTH").ToString() & ") AS [VectorDesc]" & vbNewLine

                strSQL = strSQL & "FROM " & Brack(coltemp.strRecSet) & vbNewLine
                strSQL = strSQL & ";"
            End If 'If CurrentDb.IsSQLExpress Then


            Return strSQL

        End Function

        Private Function MakePopRowQry(ByVal rowTemp As typRowType) As String
            '//================================================================================//
            '/|   FUNCTION: MakeRowColQry
            '/| PARAMETERS: rowTemp, a UDT variable containing the RowType
            '/|    RETURNS: SQL String that is the append query used to Populate the
            '/|                 rowType 's Element Table or "FALSE" on Failure
            '/|    PURPOSE: Create an append query that will populate Elm Table
            '/|      USAGE: s = MakePopRowQry(rowTemp)
            '/|         BY: Sean
            '/|       DATE: 01/10/2004
            '/|    HISTORY: 01/10/2004 Originally Adapted to ADO
            '//================================================================================//


            Dim strINS As String, strSQL As String, strSelObj As String, strConstraintName As String
            Dim strRowClass As String
            Dim j As Long

            MakePopRowQry = "FALSE"

            If CurrentDb.IsSQLExpress Then
                'strSQL = ""

                strINS = "INSERT INTO " & _miscParams.Item("LPM_ROW_ELEMENT_TABLE_PRE").ToString() & rowTemp.strTable _
                         & " ( " & rowTemp.strType & "ID, "
                strSQL = strINS
                strSelObj = ""
                strConstraintName = "'" & rowTemp.strPrefix & "' "

                For j = 1 To rowTemp.intClassCount
                    Application.DoEvents()
                    If Cancel Then Return GenUtils.ReturnStatus.Failure
                    strRowClass = rowTemp.strClasses(CInt(j))
                    If strRowClass <> "" Then
                        strSQL = strSQL & Brack(rowTemp.strClasses(CInt(j))) & ", "
                        _progress = Brack(rowTemp.strRecSet)
                        'Debug.Print(_progress)
                        'Console.WriteLine(_progress)

                        _progress = Brack(strRowClass)
                        'Debug.Print(_progress)
                        'Debug.Print(Brack(strRowClass))
                        'Console.WriteLine(_progress)
                        strSelObj = strSelObj & Brack(rowTemp.strRecSet) & "." & Brack(strRowClass) & " AS " & Brack(strRowClass) & ", " & vbNewLine
                        strConstraintName = strConstraintName & "+" & " " & "'" & "_" & "'" & " " & "+" & " " & "CONVERT(NVARCHAR(255)," & Brack(strRowClass) & ") "
                    End If
                Next j

                If Len(rowTemp.strRHSFld) > 0 Then
                    strSQL = strSQL & "RHS, "
                End If
                If Len(rowTemp.strSense) > 0 Then
                    strSQL = strSQL & "SENSE, "
                End If
                strSQL = strSQL & rowTemp.strType & "Code, " & rowTemp.strType & "Desc )" & vbNewLine
                strSQL = strSQL & "SELECT DISTINCT" & vbNewLine & "0 AS ID," & vbNewLine
                strSQL = strSQL & strSelObj
                If Len(rowTemp.strRHSFld) > 0 Then
                    strSQL = strSQL & rowTemp.strRecSet & "." & rowTemp.strRHSFld & " AS RHS, " & vbNewLine
                End If
                If Len(rowTemp.strSense) > 0 Then
                    strSQL = strSQL & "'" & rowTemp.strSense & "'" & " AS SENSE, " & vbNewLine
                End If

                strConstraintName = strConstraintName & " AS ConstraintName," & vbNewLine

                strSQL = strSQL & strConstraintName & "Left(" & Brack(rowTemp.strRowDescFld) & "," & _miscParams.Item("LPM_TEXT_LENGTH").ToString() & ") AS ConstraintDesc" & vbNewLine

                strSQL = strSQL & "FROM " & rowTemp.strRecSet & vbNewLine
                strSQL = strSQL & ";"

            Else 'CurrentDb.IsSQLExpress = False
                strINS = "INSERT INTO " & _miscParams.Item("LPM_ROW_ELEMENT_TABLE_PRE").ToString() & rowTemp.strTable _
                         & " ( " & rowTemp.strType & "ID, "
                strSQL = strINS
                strSelObj = ""
                strConstraintName = "'" & rowTemp.strPrefix & "' "

                For j = 1 To rowTemp.intClassCount
                    Application.DoEvents()
                    If Cancel Then Return GenUtils.ReturnStatus.Failure
                    strRowClass = rowTemp.strClasses(CInt(j))
                    If strRowClass <> "" Then
                        strSQL = strSQL & Brack(rowTemp.strClasses(CInt(j))) & ", "
                        _progress = Brack(rowTemp.strRecSet)
                        'Debug.Print(_progress)
                        'Console.WriteLine(_progress)

                        _progress = Brack(strRowClass)
                        'Debug.Print(_progress)
                        'Debug.Print(Brack(strRowClass))
                        'Console.WriteLine(_progress)
                        strSelObj = strSelObj & Brack(rowTemp.strRecSet) & "." & Brack(strRowClass) & " AS " & Brack(strRowClass) & ", " & vbNewLine
                        strConstraintName = strConstraintName & Chr(38) & " " & "'" & "_" & "'" & " " & Chr(38) & " " & Brack(strRowClass)
                    End If
                Next j

                If Len(rowTemp.strRHSFld) > 0 Then
                    strSQL = strSQL & "RHS, "
                End If
                If Len(rowTemp.strSense) > 0 Then
                    strSQL = strSQL & "SENSE, "
                End If
                strSQL = strSQL & rowTemp.strType & "Code, " & rowTemp.strType & "Desc )" & vbNewLine
                strSQL = strSQL & "SELECT DISTINCT" & vbNewLine & "0 AS ID," & vbNewLine
                strSQL = strSQL & strSelObj
                If Len(rowTemp.strRHSFld) > 0 Then
                    strSQL = strSQL & rowTemp.strRecSet & "." & rowTemp.strRHSFld & " AS RHS, " & vbNewLine
                End If
                If Len(rowTemp.strSense) > 0 Then
                    strSQL = strSQL & "'" & rowTemp.strSense & "'" & " AS SENSE, " & vbNewLine
                End If

                strConstraintName = strConstraintName & " AS ConstraintName," & vbNewLine

                strSQL = strSQL & strConstraintName & "Left(" & Brack(rowTemp.strRowDescFld) & "," & _miscParams.Item("LPM_TEXT_LENGTH").ToString() & ") AS ConstraintDesc" & vbNewLine

                strSQL = strSQL & "FROM " & rowTemp.strRecSet & vbNewLine
                strSQL = strSQL & ";"
            End If 'If CurrentDb.IsSQLExpress Then


            Return strSQL

        End Function

        Private Function MatGen() As Boolean
            '//================================================================================//
            '/|   FUNCTION: MatGen
            '/| PARAMETERS: -NONE-
            '/|    RETURNS: True on Success and False by default or Failure
            '/|    PURPOSE: Generate the Matrix by Populating LPM_MATRIX_TABLE_NAME using the
            '/|             data from LPM_COLUMN_DEF_TABLE_NAME, LPM_CONSTR_DEF_TABLE_NAME,
            '/|             and LPM_COEFFS_DEF_TABLE_NAME
            '/|      USAGE: i= MatGen()
            '/|         BY: Sean
            '/|       DATE: 01/15/2004
            '/|    HISTORY: 12/15/2003 Originally Written
            '/|             01/15/2004 Originally Adapted to ADO
            '/|             02/25/2010 Extended table tsysDefCoef with two additional fields (UseAlternateQry and AlternateQryText)
            '/|                         that will now allow a more efficient custom SQL to be used instead of relying on MakeMatGenQuery2().
            '//================================================================================//

            Dim startTime As Integer = My.Computer.Clock.TickCount
            Dim rs As DataTable
            Dim i As Integer, rowCount As Integer
            Dim strINS As String, strSQL As String, s As String
            Dim T As Date
            Dim ctrRow As Integer
            'ctrCol As Integer, ctrMtx As Integer
            Dim sql As String 'RKP/02-25-10
            Dim useMatGenExt As Boolean = GenUtils.IsSwitchAvailable(_switches, "/UseMatGenExt") 'RKP/02-25-10

            Dim linqTable As System.Data.EnumerableRowCollection(Of System.Data.DataRow) = Nothing
            Dim dblQueryResults As System.Data.EnumerableRowCollection(Of Double) = Nothing
            Dim intQueryResults As System.Data.EnumerableRowCollection(Of Integer) = Nothing
            Dim charQueryResults As System.Data.EnumerableRowCollection(Of Char) = Nothing
            Dim strQueryResults As System.Data.EnumerableRowCollection(Of String) = Nothing
            Dim intGroups As System.Collections.Generic.IEnumerable(Of Integer) = Nothing
            Dim myArrayList As ArrayList = Nothing
            Dim cumTotal As Integer = 0
            'Dim dt As DataTable
            'Dim dtClone As DataTable
            'Dim dc As DataColumn
            'Dim dr As DataRow

            _progress = "C-OPT Engine - MatGen..."
            Debug.Print(_progress)
            Console.WriteLine(_progress)
            'EntLib.COPT.Log.Log(_workDir, "Status", "Entering MatGen...")
            EntLib.COPT.Log.Log(_workDir, "Status", "")
            EntLib.COPT.Log.Log(_workDir, "Status", "MatGen...Started at: " & Now())

            T = Now()

            _dtMtx = Nothing
            '_dtMtx = New DataTable
            'dc = New DataColumn("ColID")
            'dc.DataType = System.Type.GetType("System.Int32")
            '_dtMtx.Columns.Add(dc)
            'dc = New DataColumn("RowID")
            'dc.DataType = System.Type.GetType("System.Int32")
            '_dtMtx.Columns.Add(dc)
            'dc = New DataColumn("COL")
            'dc.DataType = System.Type.GetType("System.String")
            '_dtMtx.Columns.Add(dc)
            'dc = New DataColumn("ROW")
            'dc.DataType = System.Type.GetType("System.String")
            '_dtMtx.Columns.Add(dc)
            'dc = New DataColumn("COEF")
            'dc.DataType = System.Type.GetType("System.Double")
            '_dtMtx.Columns.Add(dc)

            _lastSql = "SELECT * FROM " & _miscParams.Item("LPM_COEFFS_DEF_TABLE_NAME").ToString
            rs = _currentDb.GetDataSet(_lastSql).Tables(0)

            strINS = "INSERT INTO " & _miscParams!LPM_MATRIX_TABLE_NAME.ToString() & vbNewLine & _
                     "  ( ColID, RowID, COL, ROW, COEF )" & vbNewLine

            If GenUtils.IsSwitchAvailable(_switches, "/UseMinSysRes") Then
                If GenUtils.IsSwitchAvailable(_switches, "/UseMSAccessSyntax") Then
                    Try
                        strINS = "INSERT INTO " & _
                                    "[" & Me.MiscParams.Item("LINKED_MTX_DB_PATH").ToString() & "]." & _
                                    "[" & _miscParams!LPM_MATRIX_TABLE_NAME.ToString() & "] " & vbNewLine & _
                                    "  ( ColID, RowID, COL, ROW, COEF )" & vbNewLine
                    Catch ex As Exception
                        EntLib.COPT.Log.Log(_workDir, "MatGen - Error", "/UseMSAccessSyntax switch turned on but field, LINKED_MTX_DB_PATH, was not found in qsysMiscParams - " & ex.Message)
                    End Try

                End If
            End If

            '// C O E F F  L O O P  //
            For i = 0 To rs.Rows.Count - 1
                Application.DoEvents()
                If Cancel Then
                    EntLib.COPT.Log.Log(_workDir, "Status", "MatGen was aborted by user.")
                    Return GenUtils.ReturnStatus.Failure
                End If

                'strSQL = strINS & MakeMatGenQry2(rsClone!CoeffTypeID)
                s = ""
                _matGenOK = True

                sql = MakeMatGenQry2(CLng(rs.Rows(i).Item("CoeffTypeID").ToString))

                If useMatGenExt Then
                    Try
                        If rs.Rows(i).Item("UseAlternateQry") = True Then
                            sql = rs.Rows(i).Item("AlternateQryText").ToString()
                        End If
                    Catch ex As Exception
                        EntLib.COPT.Log.Log(_workDir, "MatGen - Error", "/UseMatGenExt switch is turned ON but no valid query was found - " & ex.Message)

                    End Try
                End If

                strSQL = strINS & sql
                If _matGenOK = True Then
                    s = rs.Rows(i).Item("CoeffTypeID").ToString & " " & rs.Rows(i).Item("colType").ToString & " " & rs.Rows(i).Item("rowType").ToString
                    'Debug.Print(s)
                    'EntLib.COPT.Log.Log(_workDir, "      MatGen: " & s)
                    _lastSql = strSQL
                    'EntLib.COPT.Log.Log("Status", "MatGen: " & _lastSql)
                    If GenUtils.IsSwitchAvailable(_switches, "/LogVerbose") Then
                        'If GenUtils.GetSwitchArgument(_switches, "/LogVerbose", 1).Trim().ToUpper().Equals("TRUE") Then
                        EntLib.COPT.Log.Log(_workDir, " ", "      MatGen: " & vbNewLine & _lastSql)
                        'End If
                    End If
                    _progress = s
                    Debug.Print(s)
                    Console.WriteLine(s)
                    Try
                        If GenUtils.IsSwitchAvailable(_switches, "/UseMinSysRes") Then
                            'If GenUtils.IsSwitchAvailable(_switches, "/UseMinSysRes") Then
                            '    'dt = _currentDb.GetDataSet(sql).Tables(0)
                            '    'rowCount = _
                            '    '    GenUtils.UpdateDB( _
                            '    '        Me.MiscParams.Item("LINKED_MTX_DB_CONN_STR").ToString(), _
                            '    '        dt, _
                            '    '        "SELECT * FROM " & Me.MiscParams.Item("LPM_MATRIX_TABLE_NAME").ToString() _
                            '    '    )
                            '    '_dtMtx = dt.Copy

                            'Else
                            '    'do nothing
                            'End If

                            'RKP/06-24-10/v2.3.133
                            'Build an empty dt because for some blocks, COEF comes from qryOne, where the data is "1" (System.Single),
                            'which conflicts with the data type of COEF in _dtMtx (System.Double).
                            'dt = Nothing
                            'dt = New DataTable
                            'dc = New DataColumn("ColID")
                            'dc.DataType = System.Type.GetType("System.Int32")
                            'dt.Columns.Add(dc)
                            'dc = New DataColumn("RowID")
                            'dc.DataType = System.Type.GetType("System.Int32")
                            'dt.Columns.Add(dc)
                            'dc = New DataColumn("COL")
                            'dc.DataType = System.Type.GetType("System.String")
                            'dt.Columns.Add(dc)
                            'dc = New DataColumn("ROW")
                            'dc.DataType = System.Type.GetType("System.String")
                            'dt.Columns.Add(dc)
                            'dc = New DataColumn("COEF")
                            'dc.DataType = System.Type.GetType("System.Double")
                            'dt.Columns.Add(dc)

                            'dt = _currentDb.GetDataSet(sql).Tables(0)
                            'If dt.Columns("COEF").DataType.ToString().Equals(System.Type.GetType("System.Double").ToString()) Then
                            '    _dtMtx.Merge(dt)
                            'Else
                            '    dtClone = dt.Clone
                            '    dtClone.Columns("COEF").DataType = System.Type.GetType("System.Double")
                            '    For Each dr In dt.Rows
                            '        dtClone.ImportRow(dr)
                            '    Next
                            '    _dtMtx.Merge(dtClone)
                            'End If
                            'rowCount = _currentDb.ExecuteNonQuery(_lastSql)
                        Else
                            'rowCount = _currentDb.ExecuteNonQuery(_lastSql)
                        End If
                        rowCount = _currentDb.ExecuteNonQuery(_lastSql)
                    Catch ex As Exception
                        EntLib.COPT.Log.Log("Error - MatGen - " & s & " - " & ex.Message)
                        EntLib.COPT.Log.Log(_lastSql)

                        'RKP/09-18-11/v3.0.150
                        GenUtils.Message(GenUtils.MsgType.Critical, "C-OPT Engine - MatGen", "A serious error occured during ""MatGen"", while processing:" & vbNewLine & s & vbNewLine & "Error Message:" & vbNewLine & ex.Message & vbNewLine & vbNewLine & "Please check C-OPT.log for a more detailed explanation.")

                        Return False
                    End Try

                    s = s & "     ... " & GenUtils.FormatTime(startTime, My.Computer.Clock.TickCount) 'Format(Now() - T, "hh:nn:ss")
                    _progress = s
                    Debug.Print(s)
                    Console.WriteLine(s)
                    EntLib.COPT.Log.Log(_workDir, " ", "      MatGen: " & s)
                End If
            Next

            If GenUtils.IsSwitchAvailable(_switches, "/UseMinSysRes") Then

                'qsysMTXwithCOLS
                'SELECT tsysMTX.ColID, tsysMTX.RowID, tsysMTX.COL, tsysMTX.ROW, tsysMTX.COEF, tsysCOL.OBJ, 
                'tsysCOL.LO, tsysCOL.UP, tsysCOL.FREE, tsysCOL.INTGR, tsysCOL.BINRY, tsysCOL.SOSTYPE, 
                'tsysCOL.SOSMARKER, tsysCOL.ACTIVITY, tsysCOL.DJ, tsysCOL.STATUS
                'FROM tsysCOL INNER JOIN tsysMTX ON tsysCOL.COL=tsysMTX.COL;
                'dblQueryResults = From r In linqTable _
                '               Order By r("ColID"), r("RowID") _
                '               Select COEF = CDbl(r("COEF"))

                'Dim queryResults = From a In dtMtx.AsEnumerable() _
                '                   Join b In dtCol.AsEnumerable() _
                '                   On a("COL").ToString() = b("COL").ToString() _

                'http://msdn.microsoft.com/en-us/library/bb669071.aspx
                'Dim query = _
                '    contacts.AsEnumerable().Join(orders.AsEnumerable(), _
                '    Function(order) order.Field(Of Int32)("ContactID"), _
                '    Function(contact) contact.Field(Of Int32)("ContactID"), _
                '        Function(contact, order) New With _
                '        { _
                '            .ContactID = contact.Field(Of Int32)("ContactID"), _
                '            .SalesOrderID = order.Field(Of Int32)("SalesOrderID"), _
                '            .FirstName = contact.Field(Of String)("FirstName"), _
                '            .Lastname = contact.Field(Of String)("Lastname"), _
                '            .TotalDue = order.Field(Of Decimal)("TotalDue") _
                '        }) _
                '        .GroupBy(Function(record) record.ContactID)

                'Dim query = _
                '        From a In dtMtx _
                '        Group Join b In dtCol _
                '            On a.Field(Of String)("COL") _
                '            Equals b.Field(Of String)("COL") Into MTXwithCOLS = Group _
                '        Select New With _
                '            { _
                '                .CustomerID = a.Field(Of Integer)("SalesOrderID"), _
                '                .ords = MTXwithCOLS.Count() _
                '            }

                'The following LINQ query returns an anonymous type, which is undesirable.
                'Any time a LINQ query has more than 1 field in the "Select" statement, it 
                'returns an anonymous type.
                'An anonymous type will not translate to a DataTable (by using the CopyToDataTable method).
                'Hence, as a workaround, the LoadSolverArraysUsingJoin method will now
                'load all the solver arrays by replicating the way dtCol and dtMtx are generated
                'by taking care of empty columns and rows with COEF = 0.
                'Dim queryResults = _
                '    dtCol.AsEnumerable().Join(dtMtx.AsEnumerable(), Function(c) c.Field(Of String)("COL"), _
                '    Function(m) m.Field(Of String)("COL"), _
                '    Function(c, m) _
                '        New With {.ColID = m.Field(Of Integer)("ColID"), _
                '        .RowID = m.Field(Of Integer)("RowID"), _
                '        .COL = m.Field(Of String)("COL"), _
                '        .ROW = m.Field(Of String)("ROW"), _
                '        .COEF = m.Field(Of Double)("COEF")})


                's = "MatGen....." & GenUtils.FormatTime(startTime, My.Computer.Clock.TickCount)

                '_progress = "   Loading solver arrays started right after MatGen..." & My.Computer.Clock.LocalTime.ToString
                'Debug.Print("   Loading solver arrays started right after MatGen..." & My.Computer.Clock.LocalTime.ToString)
                'Console.WriteLine("   Loading solver arrays started after MatGen..." & My.Computer.Clock.LocalTime.ToString)
                'EntLib.COPT.Log.Log(_workDir, " ", "      MatGen-LoadSolverArrays: " & "   Loading solver arrays started after MatGen..." & My.Computer.Clock.LocalTime.ToString)

                ''LoadSolverArrays()
                ''LoadSolverArraysUsingMtxCol(_dtRow, _dtMtx)

                '_progress = "   Loading solver arrays complete after MatGen..." & My.Computer.Clock.LocalTime.ToString
                'Debug.Print("   Loading solver arrays complete after MatGen..." & My.Computer.Clock.LocalTime.ToString)
                'Console.WriteLine("   Loading solver arrays complete after MatGen..." & My.Computer.Clock.LocalTime.ToString)
                'EntLib.COPT.Log.Log(_workDir, " ", "      MatGen-LoadSolverArrays: " & "   Loading solver arrays complete after MatGen..." & My.Computer.Clock.LocalTime.ToString)
            End If

            s = "MatGen....." & GenUtils.FormatTime(startTime, My.Computer.Clock.TickCount)
            _progress = s
            Debug.Print(s)
            Console.WriteLine(s)
            EntLib.COPT.Log.Log(_workDir, "Status", "MatGen: " & s)

            'ctrCol = CInt(_currentDb.GetScalarValue("SELECT COUNT(*) FROM tsysCol"))
            'ctrRow = CInt(_currentDb.GetScalarValue("SELECT COUNT(*) FROM tsysRow"))
            'ctrMtx = CInt(_currentDb.GetScalarValue("SELECT COUNT(*) FROM tsysMtx"))
            ctrRow = dtRow.Rows.Count

            'EntLib.COPT.Log.Log(_workDir, "Status", "ColCt (tsysCol) = " & ctrCol.ToString())
            EntLib.COPT.Log.Log(_workDir, "Status", "RowCt (tsysRow) = " & ctrRow.ToString())
            'EntLib.COPT.Log.Log(_workDir, "Status", "NZCt  (tsysMtx) = " & ctrMtx.ToString())

            '_progress = "ColCt (tsysCol) = " & ctrCol.ToString()
            'Debug.Print(_progress)
            'Console.WriteLine(_progress)

            _progress = "RowCt (tsysRow) = " & ctrRow.ToString()
            Debug.Print(_progress)
            Console.WriteLine(_progress)

            '_progress = "NZCt  (tsysMtx) = " & ctrMtx.ToString()
            'Debug.Print(_progress)
            'Console.WriteLine(_progress)

            'Debug.Print("ROWS:  " & mlElmCounter)
            'EntLib.Log.Log("C-OPT Engine - PopElmTbls - ", "ROWS:  " & mlElmCounter)

            EntLib.COPT.Log.Log(_workDir, "Status", "MatGen...Ended at: " & Now())
            EntLib.COPT.Log.Log(_workDir, "Status", "Exiting MatGen...")
            EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - MatGen took: ", GenUtils.FormatTime(startTime, My.Computer.Clock.TickCount))

            Return True

        End Function

        Private Function MakeMatGenQry2(ByVal ID As Long) As String
            '//================================================================================//
            '/|   FUNCTION: MakeMatGenQry2
            '/| PARAMETERS: lngCoeffID, ID from the Table LPM_COEFFS_DEF_TABLE_NAME
            '/|    RETURNS: SQL on Success and False by default or Failure
            '/|    PURPOSE: Create SQL for getting a set of matrix coefficients for a given
            '/|             column type and row type.
            '/|      USAGE: i= MakeMatGenQry2(23)
            '/|         BY: Sean, Ravi
            '/|       DATE: 01/14/2004
            '/|    HISTORY: 3/15/97  Originally Written
            '/|             3/17/97  Handle all 8 cases (Class in Col and/or Row and/or Qry)
            '/|             3/26/97  Add Column ID and Row ID
            '/|             4/11/97  Change Modularity based on Read, Make Qry, Execute
            '/|                      also allow coeff recset to access tables as well as qry's
            '/|             7/22/99  Change the join implementation from WHERE clauses to
            '/|                      INNER JOIN syntax
            '/|             11/8/99  Complete and test the previous change.  Get all 27 unique
            '/|                      JOIN Combo Possibilites, translate to  the 11 general
            '/|                      query structures
            '/|             8/2/00
            '/|             1/14/04  adapt to ADO
            '/|             5/03/07  adapt to ADO.NET (C-OPT.net, C-OPT2)
            '//================================================================================//

            Dim startTime As Integer = My.Computer.Clock.TickCount

            Dim rowTemp As typRowType, colTemp As typColType
            Dim coeffTemp As typCoeffType
            Dim qdf As DataTable 'Object 'ADOX.Table   'QueryDef
            Dim fld As DataColumn  'Object 'ADOX.column  'Field
            'Dim tdf As Object 'ADOX.Table   'TableDef

            Dim i As Integer, j As Integer, intCols As Integer, intRows As Integer
            Dim intIgNull As Integer, intFirstWhere As Integer
            'Dim intColClassStatus As Integer
            Dim intLast As Integer, intDupe As Integer, intCountWhere As Integer
            Dim intCountC2R As Integer, intCountC2Q As Integer, intCountR2Q As Integer
            Dim intFirst3265Error As Integer, strQorT As String
            Dim strColElTable As String, strRowElTable As String
            Dim strColType As String, strRowType As String
            Dim strQryName As String, strQryField As String, strSQL As String
            Dim strColClass As String, strRowClass As String
            Dim strClassIncCode As String, strClassName As String
            Dim strC2Rcode As String, strC2Qcode As String, strR2Qcode As String
            Dim strJoinComboType As String, intMatgenQryType As Integer
            Dim sF00 As String, sF01 As String, sF02 As String, sF03 As String, sF04 As String, sF05 As String
            Dim sJ1 As String, sJ2 As String, sJ3 As String, sC01 As String, sC02 As String, sC03 As String

            Dim Classes(1, 1) As String

            'Debug.Print("MakeMatGenQry2...")
            'EntLib.COPT.Log.Log(_workDir, "Status", "Entering MakeMatGenQry2...")

            intIgNull = CInt(False)
            intFirstWhere = CInt(True)
            intCountWhere = 0
            intCountC2R = 0
            intCountC2Q = 0
            intCountR2Q = 0
            intRows = 0
            intLast = 0
            strQorT = "Q"
            strJoinComboType = ""
            intFirst3265Error = CInt(True)

            coeffTemp = ReadCoeffType(ID)
            colTemp = ReadColType(coeffTemp.lngColID)
            rowTemp = ReadRowType(coeffTemp.lngRowID)

            '// G E T  C O L   A N D   R O W   I N F O R M A T I O N  //
            If colTemp.intActive * rowTemp.intActive * coeffTemp.intActive = 0 Then
                _matGenOK = False
                'MakeMatGenQry2 = "False"
                Return "False"
                'GoTo MakeMatGenQry2_Done
            Else
                '// Get Col Type
                strColType = coeffTemp.strColType

                '// Get Row Type
                strRowType = coeffTemp.strRowType

                'For i = 0 To moDAL.oCatalog.Tables.Count - 1           'debug TABLE LOOP
                ' Debug.Print moDAL.oCatalog.Tables(i).Name
                ' Log False, moDAL.oCatalog.Tables(i).Name
                'Next i

                strQryName = Brack(coeffTemp.strRecSet)                              '// Get Query Name
                strQryField = Brack(coeffTemp.strCoeffFld)                           '// Get Query's Field Name
                strColElTable = Brack(_miscParams.Item("LPM_COL_ELEMENT_TABLE_PRE").ToString _
                                  & colTemp.strTable)                                '// Get Col Element Table Name
                strRowElTable = Brack(_miscParams.Item("LPM_ROW_ELEMENT_TABLE_PRE").ToString _
                                  & rowTemp.strTable)                                '// Get Row Element Table Name

                'strQryName = "SELECT TOP 1 * FROM " & strQryName
                'qdf = moDAL.oCatalog.Tables(UnBrack(strQryName))                 '// Open the QueryDef Object
                'qdf = _currentDb.GetDataSet(strQryName).Tables(0)
                qdf = _currentDb.GetDataSet("SELECT TOP 1 * FROM " & strQryName).Tables(0)
                intCols = colTemp.intClassCount
                intRows = rowTemp.intClassCount

                ReDim Classes(intCols + intRows, 4)    '// 4 column array:  ClassName, Col?, Row?, Qry?

                '// Initialize the Class Array with all 'N'
                For i = 1 To intCols + intRows
                    Application.DoEvents()
                    If Cancel Then Return GenUtils.ReturnStatus.Failure
                    Classes(i, 2) = "N"
                    Classes(i, 3) = "N"
                    Classes(i, 4) = "N"
                Next i

                '// Fill Array with the Column Classes
                For i = 1 To intCols
                    Application.DoEvents()
                    If Cancel Then Return GenUtils.ReturnStatus.Failure
                    strColClass = UCase(colTemp.strClasses(i))
                    Classes(i, 1) = strColClass
                    Classes(i, 2) = "Y"                 '// Put a 'Y' in the 2 column (Col?) of the array
                    intLast = i
                Next i

                '// Search Rows for Classes
                For j = 1 To intRows
                    Application.DoEvents()
                    If Cancel Then Return GenUtils.ReturnStatus.Failure
                    strRowClass = UCase(rowTemp.strClasses(j))
                    intDupe = CInt(False)
                    For i = 1 To intCols
                        If Cancel Then Return GenUtils.ReturnStatus.Failure
                        strColClass = colTemp.strClasses(i)
                        If strRowClass = strColClass Then   '// If it is also a Column Class Then
                            intDupe = CInt(True)                   '// It's a duplicate
                            Classes(i, 3) = "Y"              '// Put a 'Y' in the 3 column (Row?) of the array
                        End If
                    Next i

                    If CBool(Not intDupe) Then                          '// If this Row class is a new class Then
                        Classes(intLast + 1, 1) = strRowClass     '// Add it to the array
                        Classes(intLast + 1, 3) = "Y"             '// Put a 'Y' in the 3 column (Row?) of the array
                        intLast = intLast + 1
                    End If
                Next j


                Select Case strQorT

                    Case "Q"
                        '// Search The Query for Classes
                        For i = 1 To intCols + intRows
                            Application.DoEvents()
                            If Cancel Then Return GenUtils.ReturnStatus.Failure
                            For Each fld In qdf.Columns
                                Application.DoEvents()
                                If Classes(i, 1) = UCase(fld.Caption) Then
                                    Classes(i, 4) = "Y"                       '// Put a 'Y' in the 4 column (Qry?) of the array
                                End If
                            Next
                        Next i

                    Case "ELSE" ' ONLY FOR ACCESS ONLY, SINCE ADO THIS IS NOT NEEDED.
                        '// Search The Table for Classes
                        'For i = 1 To intCols + intRows
                        '    For Each fld In tdf.Columns
                        '        If Classes(i, 1) = UCase(fld.Caption) Then
                        '            Classes(i, 4) = "Y"                       '// Put a 'Y' in the 4 column (Qry?) of the array
                        '        End If
                        '    Next
                        'Next i

                End Select

                '// Build the SQL statement (SELECT)

                strSQL = ""
                strSQL = "SELECT DISTINCT" & vbNewLine       '// Select Distinct {Changed from DISTINCTROW 11/7/03  --STM}

                '// Field arguments to the Select Statement
                'strSQL = strSQL & strColElTable & "." & strColType & "ID AS [CID]," & vbNewLine
                'strSQL = strSQL & strRowElTable & "." & strRowType & "ID AS [RID]," & vbNewLine
                'strSQL = strSQL & strColElTable & "." & strColType & "Code AS [COLUMN]," & vbNewLine
                'strSQL = strSQL & strRowElTable & "." & strRowType & "Code AS [ROW]," & vbNewLine
                'strSQL = strSQL & strQryName & "." & strQryField & " AS [COEF]" & vbNewLine

                'RKP/06-23-10/v2.3.133
                'Changed the field names to reflect the actual fields names in tsysMTX
                strSQL = strSQL & strColElTable & "." & strColType & "ID AS [ColID]," & vbNewLine
                strSQL = strSQL & strRowElTable & "." & strRowType & "ID AS [RowID]," & vbNewLine
                strSQL = strSQL & strColElTable & "." & strColType & "Code AS [COL]," & vbNewLine
                strSQL = strSQL & strRowElTable & "." & strRowType & "Code AS [ROW]," & vbNewLine
                strSQL = strSQL & strQryName & "." & strQryField & " AS [COEF]" & vbNewLine

                intIgNull = CInt(True)                             '// Set to ignore nulls in the following loops

                For i = 1 To intCols + intRows
                    Application.DoEvents()
                    If Cancel Then Return GenUtils.ReturnStatus.Failure
                    strClassIncCode = Classes(i, 2) & Classes(i, 3) & Classes(i, 4)
                    Select Case strClassIncCode
                        Case "YYY"     '// Class is in Column, Row, and Query
                            intCountWhere = intCountWhere + 2
                            intCountC2R = intCountC2R + 1
                            intCountC2Q = intCountC2Q + 1
                        Case "YYN"     '// Class is in Column and Row but not the Query
                            intCountWhere = intCountWhere + 1
                            intCountC2R = intCountC2R + 1
                        Case "YNY"     '// Class is in Column and Query
                            intCountWhere = intCountWhere + 1
                            intCountC2Q = intCountC2Q + 1
                        Case "NYY"     '// Class is in Row and Query
                            intCountWhere = intCountWhere + 1
                            intCountR2Q = intCountR2Q + 1
                    End Select
                Next i


                '// Change all the counts that are greater than one to "2" and convert the rest to strings
                If intCountC2R >= 2 Then
                    strC2Rcode = "2"
                Else
                    strC2Rcode = CStr(intCountC2R)
                End If

                If intCountC2Q >= 2 Then
                    strC2Qcode = "2"
                Else
                    strC2Qcode = CStr(intCountC2Q)
                End If

                If intCountR2Q >= 2 Then
                    strR2Qcode = "2"
                Else
                    strR2Qcode = CStr(intCountR2Q)
                End If

                strJoinComboType = strC2Rcode & strC2Qcode & strR2Qcode


                '// convert the join combo type (27 unique possibilities) to one of the 11 (0-10) types of generic block queries
                Select Case strJoinComboType
                    Case "000"
                        intMatgenQryType = 0
                    Case "010", "020"
                        intMatgenQryType = 1
                    Case "001", "002"
                        intMatgenQryType = 2
                    Case "100", "200"
                        intMatgenQryType = 3
                    Case "110", "120"
                        intMatgenQryType = 4
                    Case "101", "102"
                        intMatgenQryType = 5
                    Case "111", "122", "121", "112", "211", "222", "221", "212"
                        intMatgenQryType = 6
                    Case "210", "220"
                        intMatgenQryType = 7
                    Case "201"
                        intMatgenQryType = 8
                    Case "011", "022", "021", "012"
                        intMatgenQryType = 9
                    Case "202"
                        intMatgenQryType = 10
                End Select




                '// The Beginning of the FROM clause
                '===================================================================
                sF00 = "FROM " & strColElTable & ", " & strRowElTable & ", " & strQryName & vbNewLine
                sF01 = "FROM " & strRowElTable & ", " & strColElTable & vbNewLine
                sF02 = "FROM " & strColElTable & ", " & strRowElTable & vbNewLine
                sF03 = "FROM " & strColElTable & vbNewLine
                sF04 = "FROM " & strQryName & ", " & strColElTable & vbNewLine
                sF05 = "FROM (" & strColElTable & vbNewLine

                '// The INNER JOIN clauses
                '===================================================================
                sJ1 = "INNER JOIN " & strQryName & " ON" & vbNewLine
                sJ2 = "INNER JOIN (" & strRowElTable & vbNewLine
                sJ3 = "INNER JOIN " & strRowElTable & " ON" & vbNewLine

                sC01 = ""
                sC02 = ""
                sC03 = ""

                '// The Field Joins
                '===================================================================

                '// Column to Row (aka C2R and C01)
                '===================================
                intFirstWhere = CInt(True)
                If intCountC2R > 0 Then

                    '// "WHERE" Statements
                    'sC01 = sC01 & "(" & vbNewLine

                    For i = 1 To intCols + intRows
                        Application.DoEvents()
                        If Cancel Then Return GenUtils.ReturnStatus.Failure
                        strClassIncCode = Classes(i, 2) & Classes(i, 3) & Classes(i, 4)
                        strClassName = Classes(i, 1)
                        If CBool(Not intFirstWhere And CInt(Len(Classes(i, 1)) > 0)) Then
                            If Not ((strClassIncCode = "YNN") Or (strClassIncCode = "NYN")) Then    '//Class is in more than one place
                                If Not ((strClassIncCode = "YNY") Or (strClassIncCode = "NYY")) Then '//Class is not in the C2R join
                                    sC01 = sC01 & "AND" & vbNewLine       '// Add 'AND' to the SQL if it's not the first Where Statement
                                End If
                            End If
                        End If
                        Select Case strClassIncCode
                            Case "YYY"     '// Class is in Column, Row, and Query
                                sC01 = sC01 & "(" & strColElTable & "." & strClassName & " = " & strRowElTable & "." & strClassName & ")" & vbNewLine
                                intFirstWhere = CInt(False)
                            Case "YYN"     '// Class is in Column and Row but not the Query
                                sC01 = sC01 & "(" & strColElTable & "." & strClassName & " = " & strRowElTable & "." & strClassName & ")" & vbNewLine
                                intFirstWhere = CInt(False)
                            Case Else
                        End Select
                    Next i

                    '// Close the WHERE Statement
                    'sC01 = sC01 & vbNewLine & ")" & vbNewLine

                End If


                '// Column to Qry (aka C2Q and C02)
                '===================================
                intFirstWhere = CInt(True)
                If intCountC2Q > 0 Then

                    '// "WHERE" Statements

                    For i = 1 To intCols + intRows
                        Application.DoEvents()
                        If Cancel Then Return GenUtils.ReturnStatus.Failure
                        strClassIncCode = Classes(i, 2) & Classes(i, 3) & Classes(i, 4)
                        strClassName = Classes(i, 1)
                        If CBool(Not intFirstWhere And CInt(Len(Classes(i, 1)) > 0)) Then
                            If Not ((strClassIncCode = "YNN") Or (strClassIncCode = "NYN")) Then    '//Class is in more than one place
                                If Not ((strClassIncCode = "YYN") Or (strClassIncCode = "NYY")) Then '//Class is not in the C2Q join
                                    sC02 = sC02 & "AND" & vbNewLine       '// Add 'AND' to the SQL if it's not the first Where Statement
                                End If
                            End If
                        End If
                        Select Case strClassIncCode
                            Case "YYY"     '// Class is in Column, Row, and Query
                                sC02 = sC02 & "(" & strColElTable & "." & strClassName & " = " & strQryName & "." & strClassName & ")" & vbNewLine
                                intFirstWhere = CInt(False)
                            Case "YNY"     '// Class is in Column and Query
                                sC02 = sC02 & "(" & strColElTable & "." & strClassName & " = " & strQryName & "." & strClassName & ")" & vbNewLine
                                intFirstWhere = CInt(False)
                            Case Else
                        End Select
                    Next i

                    '// Close the WHERE Statement

                End If


                '// Row to Qry (aka R2Q and C03)
                '===================================
                intFirstWhere = CInt(True)
                If intCountR2Q > 0 Then

                    '// "WHERE" Statements

                    For i = 1 To intCols + intRows
                        Application.DoEvents()
                        If Cancel Then Return GenUtils.ReturnStatus.Failure
                        strClassIncCode = Classes(i, 2) & Classes(i, 3) & Classes(i, 4)
                        strClassName = Classes(i, 1)
                        If CBool(Not intFirstWhere And CInt(Len(Classes(i, 1)) > 0)) Then
                            If Not ((strClassIncCode = "YNN") Or (strClassIncCode = "NYN")) Then    '//Class is in more than one place
                                If Not ((strClassIncCode = "YYY") Or (strClassIncCode = "YYN") Or (strClassIncCode = "YNY")) Then '//Class is not in the R2Q join
                                    sC03 = sC03 & "AND" & vbNewLine       '// Add 'AND' to the SQL if it's not the first Where Statement
                                End If
                            End If
                        End If
                        Select Case strClassIncCode
                            Case "NYY"     '// Class is in Row and Query
                                sC03 = sC03 & "(" & strRowElTable & "." & strClassName & " = " & strQryName & "." & strClassName & ")" & vbNewLine
                                intFirstWhere = CInt(False)
                            Case Else
                        End Select
                    Next i

                End If

                Select Case intMatgenQryType     '// note that case 7 and case 4 are the same
                    '//                8 and      5 too
                    Case 0
                        strSQL = strSQL & sF00 '& ";"
                    Case 1
                        strSQL = strSQL & sF01 & sJ1 & sC02 '& ";"
                    Case 2
                        strSQL = strSQL & sF02 & sJ1 & sC03 '& ";"
                    Case 3
                        strSQL = strSQL & sF04 & sJ3 & sC01 '& ";"
                    Case 4
                        strSQL = strSQL & sF05 & sJ3 & sC01 & ") " & vbNewLine & sJ1 & sC02 '& ";"
                    Case 5
                        strSQL = strSQL & sF05 & sJ3 & sC01 & ") " & vbNewLine & sJ1 & sC03 '& ";"
                    Case 6
                        strSQL = strSQL & sF05 & sJ3 & sC01 & ") " & vbNewLine & sJ1 & sC02 & " AND " & vbNewLine & sC03 '& ";"
                    Case 7
                        strSQL = strSQL & sF05 & sJ3 & sC01 & ") " & vbNewLine & sJ1 & sC02 '& ";"
                    Case 8
                        strSQL = strSQL & sF05 & sJ3 & sC01 & ") " & vbNewLine & sJ1 & sC03 '& ";"
                    Case 9
                        strSQL = strSQL & sF03 & sJ2 & vbNewLine & sJ1 & sC03 & ") ON" & vbNewLine & sC02 '& ";"
                    Case 10
                        strSQL = strSQL & sF03 & sJ2 & vbNewLine & sJ1 & sC03 & ") ON" & vbNewLine & sC01 '& ";"
                End Select

                Return strSQL
            End If

            EntLib.COPT.Log.Log(_workDir, "Status", "Exiting MakeMatGenQry2...")
            EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - MakeMatGenQry2 took: ", GenUtils.FormatTime(startTime, My.Computer.Clock.TickCount))
        End Function

        Private Function ReadCoeffType(ByVal ID As Long) As typCoeffType
            '//================================================================================//
            '/|   FUNCTION: ReadCoeffType
            '/| PARAMETERS: ID, the CoeffTypeID from table LPM_COEFFS_DEF_TABLE_NAME
            '/|    RETURNS: CoeffType data variable of form:
            '/|
            '/|                lngID        As Long
            '/|                intActive    As Integer
            '/|                strType      As String
            '/|                strColType   As String
            '/|                lngColID     As Long
            '/|                strRowType   As String
            '/|                lngRowID     As Long
            '/|                strRecSet    As String
            '/|                strCoeffFld  As String
            '/|
            '/|    PURPOSE: Read in all the information associated with a Coeff type
            '/|      USAGE: CoeffTemp = ReadCoeffType(7)
            '/|         BY: Sean
            '/|       DATE: 01/07/2004
            '/|    HISTORY:  3/31/97
            '//================================================================================//
            Dim startTime As Integer = My.Computer.Clock.TickCount

            Dim rs As DataTable, rsCol As DataTable, rsRow As DataTable
            Dim i As Integer, colCtr As Integer, rowCtr As Integer
            Dim coeffTemp As typCoeffType = Nothing
            'Dim strClassConcat As String

            'Debug.Print("ReadCoeffType...")
            'EntLib.COPT.Log.Log(_workDir, "Status", "Entering ReadCoeffType...")

            '// Open the Coeff Definition Table
            _lastSql = "SELECT * FROM " & _miscParams!LPM_COEFFS_DEF_TABLE_NAME.ToString()
            rs = _currentDb.GetDataSet(_lastSql).Tables(0)

            _lastSql = "SELECT * FROM " & _miscParams!LPM_COLUMN_DEF_TABLE_NAME.ToString()
            rsCol = _currentDb.GetDataSet(_lastSql).Tables(0)

            _lastSql = "SELECT * FROM " & _miscParams!LPM_CONSTR_DEF_TABLE_NAME.ToString()
            rsRow = _currentDb.GetDataSet(_lastSql).Tables(0)

            For i = 0 To rs.Rows.Count - 1
                Application.DoEvents()
                If Cancel Then Return Nothing 'GenUtils.ReturnStatus.Failure
                If CDbl(rs.Rows(i).Item("CoeffTypeID").ToString) = ID Then
                    coeffTemp.lngID = CInt(rs.Rows(i).Item("CoeffTypeID").ToString)
                    coeffTemp.intActive = CInt(rs.Rows(i).Item("CoeffActive"))
                    coeffTemp.strType = rs.Rows(i).Item("CoeffType").ToString
                    coeffTemp.strColType = rs.Rows(i).Item("colType").ToString

                    '// Get the ColTypeID
                    For colCtr = 0 To rsCol.Rows.Count - 1 '// Find the Correct Col Type
                        Application.DoEvents()
                        If Cancel Then Return Nothing 'GenUtils.ReturnStatus.Failure
                        'If rsCol!colType = coeffTemp.strColType Then    '// Read the Col ID
                        If rsCol.Rows(colCtr).Item("colType").ToString = coeffTemp.strColType Then '// Read the Col ID
                            coeffTemp.lngColID = CInt(rsCol.Rows(colCtr).Item("ColTypeID").ToString)  '!ColTypeID
                            Exit For
                        End If
                    Next

                    coeffTemp.strRowType = rs.Rows(i).Item("rowType").ToString

                    '// Get the RowTypeID
                    'Do While rsRow.EOF = False                         '// Find the Correct Row Type
                    '    If rsRow!rowType = coeffTemp.strRowType Then    '// Read the Row ID
                    '        coeffTemp.lngRowID = rsRow!RowTypeID
                    '        Exit Do
                    '    End If
                    '    rsRow.MoveNext()
                    'Loop
                    For rowCtr = 0 To rsRow.Rows.Count - 1 '// Find the Correct Row Type
                        Application.DoEvents()
                        If Cancel Then Return Nothing 'GenUtils.ReturnStatus.Failure
                        'If rsCol!colType = coeffTemp.strColType Then    '// Read the Col ID
                        If rsRow.Rows(rowCtr).Item("rowType").ToString = coeffTemp.strRowType Then '// Read the Row ID
                            coeffTemp.lngRowID = CInt(rsRow.Rows(rowCtr).Item("RowTypeID").ToString)  '!RowTypeID
                            Exit For
                        End If
                    Next

                    coeffTemp.strRecSet = CStr(rs.Rows(i).Item("CoeffRecSet")) 'rs!CoeffRecSet
                    coeffTemp.strCoeffFld = CStr(rs.Rows(i).Item("CoeffField")) 'rs!CoeffField
                    Exit For

                End If
            Next

            'MsgBox(coeffTemp.strRowType)

            Return coeffTemp

            EntLib.COPT.Log.Log(_workDir, "Status", "Exiting ReadCoeffType...")
            EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - ReadCoeffType took: ", GenUtils.FormatTime(startTime, My.Computer.Clock.TickCount))
        End Function

        Private Sub Solve()

            Dim success As Boolean
            Dim retVal As Integer
            'Dim dsRCC As New DataSet()

            'Dim dtCol As DataTable
            'Dim dtRow As DataTable
            'Dim dtMtx As DataTable
            'Dim dsRow As DataSet
            'Dim dsCol As DataSet
            'Dim dtrRow As DataTableReader

            'Dim newColName As String = "tmpID"

            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "", "")
            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solve - Started at: " & Now().ToString())

            _currentDb = New EntLib.COPT.DAAB(_databaseName)

            retVal = FilterEmptyColumnsRows()

            'EntLib.COPT.GenUtils.ImportUsingExcel(_dtRow, _dtCol, _workDir)


            LoadSolverArraysInSequence()

            'EntLib.COPT.GenUtils.SpreadsheetExport(dtCol, GetWorkDir() & "\Test.xls", "Test")

            'EntLib.COPT.GenUtils.SpreadsheetImport(My.Computer.FileSystem.GetParentPath(GetWorkDir()) & _databaseName & ".MDB", "TEST", GetWorkDir() & "\Test.xls", "Test")

            'dtCol = _currentDb.GetDataTable("SELECT * FROM tsysCOL")                    'SESSIONIZE

            '_progress = "Loading Coefficients into memory..."
            'Debug.Print(_progress)
            'Console.Write(_progress)
            'RKP/06-11-10/v2.3.133
            'If GenUtils.IsSwitchAvailable(_switches, "/UseMinSysRes") Then
            '    'This task should already have been completed in MatGen.
            '    'However, in case of /SolveOnly switch being used, load the arrays, since MatGen is skipped.
            '    'If array_mval Is Nothing Then
            '    'LoadMtx(_dtMtx)
            '    'LoadSolverArrays()

            '    'End If
            'Else
            '    If Not dtMtx Is Nothing Then
            '        If dtMtx.Rows.Count = 0 Then
            '            _dtMtx = Nothing
            '        End If
            '    End If
            '    If dtMtx Is Nothing Then
            '        If GenUtils.IsSwitchAvailable(_switches, "/UseSysQueries") Then
            '            Try
            '                'RKP/01-28-10/v2.3.128
            '                'If qsysMtx is not present, resort to backup option.
            '                _dtMtx = _currentDb.GetDataTable("SELECT * FROM qsysMtx")            'SESSIONIZE
            '                _srcMtx = "qsysMtx"
            '            Catch ex As Exception
            '                _dtMtx = _currentDb.GetDataTable("SELECT * FROM qsysMTXwithCOLS WHERE COEF <> 0 ORDER BY ColID, RowID")            'SESSIONIZE
            '                _srcMtx = "tsysMtx"
            '            End Try
            '        Else
            '            _dtMtx = _currentDb.GetDataTable("SELECT * FROM qsysMTXwithCOLS WHERE COEF <> 0 ORDER BY ColID, RowID")            'SESSIONIZE
            '            _srcMtx = "tsysMtx"
            '        End If
            '    Else
            '        _srcMtx = "tsysMtx"
            '    End If
            '    Me.SolutionNonZeros = dtMtx.Rows.Count
            'End If
            '_progress = "Done."
            'Debug.Print(_progress)
            'Console.WriteLine(_progress)

            '_progress = "Loading Constraints into memory..."
            'Debug.Print(_progress)
            'Console.Write(_progress)
            ''If GenUtils.IsSwitchAvailable(_switches, "/UseMinSysRes") Then
            ''    Try
            ''        'RKP/01-28-10/v2.3.128
            ''        'If qsysRow is not present, resort to backup option.
            ''        dtRow = _currentDb.GetDataTable("SELECT * FROM qsysRow")
            ''        _srcRow = "qsysRow"
            ''    Catch ex As Exception
            ''        dtRow = _currentDb.GetDataTable("SELECT * FROM tsysROW ORDER BY RowID")
            ''        _srcRow = "tsysRow"
            ''    End Try
            ''Else
            'Try
            '    'My.Computer.FileSystem.DeleteFile(_workDir & "\dtRow.xml")
            'Catch ex As Exception

            'End Try
            'dtRow.WriteXml(_workDir & "\dtRow.xml")
            'If Not dtRow Is Nothing Then
            '    If dtRow.Rows.Count = 0 Then
            '        _dtRow = Nothing
            '    End If
            'End If
            'If dtRow Is Nothing Then
            '    If GenUtils.IsSwitchAvailable(_switches, "/UseSysQueries") Then
            '        Try
            '            'RKP/01-28-10/v2.3.128
            '            'If qsysRow is not present, resort to backup option.
            '            _dtRow = _currentDb.GetDataTable("SELECT * FROM qsysRow")
            '            'dsRow = _currentDb.GetDataSet("SELECT * FROM qsysRow")
            '            _srcRow = "qsysRow"
            '        Catch ex As Exception
            '            _dtRow = _currentDb.GetDataTable("SELECT * FROM tsysROW ORDER BY RowID")
            '            'dsRow = _currentDb.GetDataSet("SELECT * FROM tsysROW ORDER BY RowID")
            '            'dtrRow = _currentDb.GetDataTableReader("SELECT * FROM tsysROW ORDER BY RowID")
            '            _srcRow = "tsysRow"
            '        End Try
            '    Else
            '        _dtRow = _currentDb.GetDataTable("SELECT * FROM tsysROW ORDER BY RowID")
            '        'dsRow = _currentDb.GetDataSet("SELECT * FROM tsysROW ORDER BY RowID")
            '        'dtrRow = _currentDb.GetDataTableReader("SELECT * FROM tsysROW ORDER BY RowID")
            '        _srcRow = "tsysRow"
            '    End If
            'Else
            '    _srcRow = "tsysRow"
            'End If
            'dtRow.TableName = "dtRow"
            'End If
            'Me.SolutionRows = dtRow.Rows.Count
            'Me.SolutionRows = dsRow.Tables(0).Rows.Count
            '_progress = "Done."
            'Debug.Print(_progress)
            'Console.WriteLine(_progress)

            '_progress = "Loading Decision Variables into memory..."
            'Debug.Print(_progress)
            'Console.Write(_progress)
            'If GenUtils.IsSwitchAvailable(_switches, "/UseMinSysRes") Then
            '    Try
            '        'RKP/01-28-10/v2.3.128
            '        'If qsysCol is not present, resort to backup option.
            '        dtCol = _currentDb.GetDataTable("SELECT * FROM qsysCol")
            '        _srcCol = "qsysCol"
            '    Catch ex As Exception
            '        dtCol = _currentDb.GetDataTable("SELECT * FROM tsysCol WHERE ColID IN (SELECT DISTINCT tsysMTX.ColID FROM tsysMTX INNER JOIN tsysCol ON tsysMTX.ColID = tsysCol.ColID WHERE tsysMTX.COEF <> 0) ORDER BY ColID")
            '        _srcCol = "tsysCol"
            '    End Try
            'Else
            'Try
            '    ' My.Computer.FileSystem.DeleteFile(_workDir & "\dtCol.xml")
            'Catch ex As Exception

            'End Try
            'dtCol.WriteXml(_workDir & "\dtCol.xml")

            'If Not dtCol Is Nothing Then
            '    If dtCol.Rows.Count = 0 Then
            '        _dtCol = Nothing
            '    End If
            'End If
            'If dtCol Is Nothing Then
            '    If GenUtils.IsSwitchAvailable(_switches, "/UseSysQueries") Then
            '        Try
            '            'RKP/01-28-10/v2.3.128
            '            'If qsysCol is not present, resort to backup option.
            '            _dtCol = _currentDb.GetDataTable("SELECT * FROM qsysCol")
            '            _srcCol = "qsysCol"
            '        Catch ex As Exception
            '            _dtCol = _currentDb.GetDataTable("SELECT * FROM tsysCol WHERE ColID IN (SELECT DISTINCT tsysMTX.ColID FROM tsysMTX INNER JOIN tsysCol ON tsysMTX.ColID = tsysCol.ColID WHERE tsysMTX.COEF <> 0) ORDER BY ColID")
            '            _srcCol = "tsysCol"
            '        End Try
            '    Else
            '        _dtCol = _currentDb.GetDataTable("SELECT * FROM tsysCol WHERE ColID IN (SELECT DISTINCT tsysMTX.ColID FROM tsysMTX INNER JOIN tsysCol ON tsysMTX.ColID = tsysCol.ColID WHERE tsysMTX.COEF <> 0) ORDER BY ColID")
            '        _srcCol = "tsysCol"
            '    End If
            '    _dtCol.TableName = "dtCol"
            'Else
            '    _srcCol = "tsysCol"
            'End If
            'End If
            'Me.SolutionColumns = dtCol.Rows.Count
            '_progress = "Done."
            'Debug.Print(_progress)
            'Console.WriteLine(_progress)



            'Me.CurrentDb.UpdateDataSet(dtRow, "SELECT * FROM " & Me.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString & " ORDER BY RowID")
            'Me.CurrentDb.UpdateDataSet(dtRow, Me.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString)

            'Dim tmpDS As New DataSet
            'tmpDS.DataSetName = "tmp"
            'Dim cmdInsert As System.Data.Common.DbCommand = Me.CurrentDb.GetDbCommand("SELECT * FROM " & Me.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString & " ORDER BY RowID")
            'Dim cmdUpdate As System.Data.Common.DbCommand = Me.CurrentDb.GetDbCommand("SELECT * FROM " & Me.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString & " ORDER BY RowID")
            'Dim cmdDelete As System.Data.Common.DbCommand = Me.CurrentDb.GetDbCommand("SELECT * FROM " & Me.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString & " ORDER BY RowID")
            ''tmpDS.Tables.Add(dtRow)
            ''tmpDS.Tables(0).TableName = "tmpTable1"
            'Me.CurrentDb.UpdateDataSet(dsRow, "SELECT * FROM " & Me.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString & " ORDER BY RowID", cmdInsert, cmdUpdate, cmdDelete)

            '9/17/09 - v2.2 Build 120
            'STM & RKP
            'Added ORDER BY clauses to all the three queries to prevent undesirable solution results going back into the database (for CoinMP).
            If GenUtils.IsSwitchAvailable(_switches, "/UseMinSysRes") Then
                If GenUtils.IsSwitchAvailable(_switches, "/UseSysQueries") Then
                    _srcRow = "qsysRow"
                    _srcCol = "qsysCol"
                    _srcMtx = "qsysMtx"
                ElseIf GenUtils.IsSwitchAvailable(_switches, "/UseSysTables") Then
                    _srcRow = "tsysRow"
                    _srcCol = "tsysCol"
                    _srcMtx = "tsysMtx"
                End If
            Else
                '_progress = "Loading [Matrix] into memory..."
                'Debug.Print(_progress)
                'Console.Write(_progress)
                'Try
                '    'RKP/01-28-10/v2.3.128
                '    'If qsysMtx is not present, resort to backup option.
                '    dtMtx = _currentDb.GetDataTable("SELECT * FROM qsysMtx")            'SESSIONIZE
                '    _srcMtx = "qsysMtx"
                'Catch ex As Exception
                '    dtMtx = _currentDb.GetDataTable("SELECT * FROM qsysMTXwithCOLS WHERE COEF <> 0 ORDER BY ColID, RowID")            'SESSIONIZE
                '    _srcMtx = "tsysMtx"
                'End Try
                'Me.SolutionNonZeros = dtMtx.Rows.Count
                '_progress = "Done."
                'Debug.Print(_progress)
                'Console.WriteLine(_progress)
            End If

            'SESSIONIZE
            'dtRow = _currentDb.GetDataTable("SELECT * FROM tsysRow WHERE RowID IN (SELECT DISTINCT tsysMTX.RowID FROM tsysMTX INNER JOIN tsysRow ON tsysMTX.RowID = tsysRow.RowID WHERE tsysMTX.COEF <> 0)")

            'dtMtx = _currentDb.GetDataTable("SELECT * FROM qsysMTXwithCOLS")            'SESSIONIZE

            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "After loading tsys tables - APM = " & My.Computer.Info.AvailablePhysicalMemory.ToString())
            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "After loading tsys tables - AVM = " & My.Computer.Info.AvailableVirtualMemory.ToString())
            Console.WriteLine("APM = " & My.Computer.Info.AvailablePhysicalMemory.ToString())
            Console.WriteLine("AVM = " & My.Computer.Info.AvailableVirtualMemory.ToString())

            'dsRCC.Tables.Add(dtCol)
            'dsRCC.Tables(0).TableName = "tsysCOL"
            'dsRCC.Tables.Add(dtRow)
            'dsRCC.Tables(1).TableName = "tsysROW"
            'dsRCC.Tables.Add(dtMtx)
            'dsRCC.Tables(2).TableName = "qsysMTXwithCOLS"

            'Dim mps As Solver = New Solver( _
            '    _miscParams, _
            '    My.Application.Info.DirectoryPath & "\" & "ModelName" & "___" & "ProblemName" & ".MPS", _
            '    "ModelName", "ProblemName", 1, "PROFIT", _
            '    _miscParams.Item("LPM_CONSTR_TABLE_NAME").ToString(), _
            '    _miscParams.Item("LPM_COLUMN_TABLE_NAME").ToString(), _
            '    "qsysMTXwithCOLS", 48, _
            '    dtRow.Rows.Count, _
            '    dtCol.Rows.Count, _
            '    dtMtx.Rows.Count, _
            '    0, _
            '    0, _
            '    "N/A", _
            '    "N/A", _
            '    "Sean MacDermant/Ravi Poluri [International Paper]", _
            '    "C-OPT2", _
            '    "Basic Mill:Machine Prod Planning LP Model" _
            ')

            'success = mps.mpslpOutputMPS(dsRCC)

            'mps.RunGLPKSolver()

            'EntLib.Log.Log("Status", "MPS file was successfully generated.")

            'Dim xa As Solver_XA = New Solver_XA
            'success = xa.Solve(dtRow, dtCol, dtMtx)
            Dim currentSolver As Solver = Nothing
            'If GenUtils.IsSwitchAvailable(_switches, "/Solver") Then
            If GenUtils.GetSwitchArgument(_switches, "/Solver", 1).Trim().ToUpper.Equals("XA") Then
                'Dim currentSolver As Solver_XA = New Solver_XA(Me)
                Me.SolverName = "XA"
                currentSolver = New Solver_XA(Me, _switches)
            ElseIf GenUtils.GetSwitchArgument(_switches, "/Solver", 1).Trim().ToUpper.Equals("COINMP") Then
                Me.SolverName = "CoinMP"
                currentSolver = New Solver_CoinMP(Me, _switches)
            ElseIf GenUtils.GetSwitchArgument(_switches, "/Solver", 1).Trim().ToUpper.Equals("MSF") Then
                Me.SolverName = "Microsoft Solver Foundation"
                currentSolver = New Solver_MSF(Me, _switches)
            ElseIf GenUtils.GetSwitchArgument(_switches, "/Solver", 1).Trim().ToUpper.Equals("CPLEX") Then
                Me.SolverName = "CPLEX"
                currentSolver = New Solver_CPLEX(Me, _switches)
            ElseIf GenUtils.GetSwitchArgument(_switches, "/Solver", 1).Trim().ToUpper.Equals("CPLEX WITH MPL") Then
                Me.SolverName = "CPLEX with MPL"
                currentSolver = New Solver_CPLEX(Me, _switches)
            ElseIf GenUtils.GetSwitchArgument(_switches, "/Solver", 1).Trim().ToUpper.Equals("GLPK") Then
                Me.SolverName = "GLPK"
                currentSolver = New Solver_GLPK(Me, _switches)
            ElseIf GenUtils.GetSwitchArgument(_switches, "/Solver", 1).Trim().ToUpper.Equals("LPSOLVE") Then
                Me.SolverName = "LPSolve"
                currentSolver = New Solver_LPSOLVE(Me, _switches)
            ElseIf GenUtils.GetSwitchArgument(_switches, "/Solver", 1).Trim().ToUpper.Equals("REMOTE") Then
                Me.SolverName = "Remote"
                currentSolver = New Solver_LPSOLVE(Me, _switches)
            Else
                'Dim currentSolver As Solver_XA = New Solver_XA(Me)
                'currentSolver = New Solver_XA(Me, _switches)
                Me.SolverName = "CoinMP"
                currentSolver = New Solver_CoinMP(Me, _switches)
            End If
            'End If
            'Dim xa As Solver_XA = New Solver_XA(Me)
            _solverName = currentSolver.getSolverName()
            currentSolver.SetSession(_currentSession) 'RKP/04-05-10/v2.3.132
            currentSolver.isMIP = _isMIP 'RKP/07-30-10/v2.3.134

            'RKP/05-15-09
            'The "tmpID" column allows CoinMP solver to push the results (ACTIVITY, DJ and SHADOW columns) back into a C-OPT database from in-memory arrays).
            'currentSolver.AddIdentityColumn(dtRow, newColName)
            'currentSolver.AddIdentityColumn(dtCol, newColName)
            'currentSolver.AddIdentityColumn(dtMtx, "tmpID")

            Try
                If GenUtils.IsSwitchAvailable(_switches, "/GenNativeMPS") Then
                    _progress = "Native MPS file generation started at: " & Microsoft.VisualBasic.Now().ToString()
                    Debug.Print(_progress)
                    Console.WriteLine(_progress)

                    currentSolver.GenMPSFile(GenUtils.GetSwitchArgument(_switches, "/GenNativeMPS", 1), dtRow, dtCol, dtMtx)

                    _progress = "Native MPS file generation ended at: " & Microsoft.VisualBasic.Now().ToString()
                    Debug.Print(_progress)
                    Console.WriteLine(_progress)
                End If
            Catch ex As Exception
                EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "GenMPSFile - Failed: " & ex.Message)
                _progress = "Native MPS file generation error - " & ex.Message
                Debug.Print(_progress)
                Console.WriteLine(_progress)
            End Try

            _progress = """" & Me.SolverName & """ Solver is now solving the model..." & My.Computer.Clock.LocalTime.ToString
            Debug.Print(_progress)
            Console.WriteLine(_progress)
            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", """" & Me.SolverName & """ Solver is now solving the model..." & My.Computer.Clock.LocalTime.ToString)

            Try
                With currentSolver
                    .array_colNames = array_colNames
                    .array_ctyp = array_ctyp
                    .array_dclo = array_dclo
                    .array_dcup = array_dcup
                    .array_dobj = array_dobj
                    .array_drhs = array_drhs
                    .array_initValues = array_initValues
                    .array_mbeg = array_mbeg
                    .array_mcnt = array_mcnt
                    .array_midx = array_midx
                    .array_mval = array_mval
                    .array_rowNames = array_rowNames
                    .array_rtyp = array_rtyp
                End With
                success = currentSolver.Solve(dtRow, dtCol, dtMtx)
            Catch ex As Exception
                EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solve - Failed - Ended at: " & Now().ToString() & " - " & ex.Message)
                GenUtils.Message(GenUtils.MsgType.Critical, "C-OPT Engine - RunAll - Solve", ex.Message)
            End Try

            'success = xa.Solve(dtRow, dtCol, dtMtx, lpProbInfo)
            'success = xa.Solve(dtRow, dtCol, dtMtx)

            If Not GenUtils.IsSwitchAvailable(_switches, "/NoSOLFile") Then
                Try
                    currentSolver.GenSOLFile(dtRow, dtCol)
                    _progress = "Solution file was generated successfully."
                    Debug.Print(_progress)
                    Console.WriteLine(_progress)
                Catch ex As Exception
                    EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solve - Failed to generate C-OPT.SOL file." & ex.Message)
                End Try
            End If

            'RKP/05-15-09
            'Remove the "tmpID" column since it is no longer required.
            '_ds.Tables(0).Columns.Remove(newCol)
            'dtRow.Columns.Remove(newColName)
            'dtCol.Columns.Remove(newColName)
            _dtRow = Nothing
            _dtCol = Nothing
            _dtMtx = Nothing
            currentSolver = Nothing

            'GC.Collect()
            'GC.WaitForPendingFinalizers()
            'GC.Collect()

            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solve - Ended at: " & Now().ToString())
            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "", "")

        End Sub

        Private Sub PostSolve()
            Dim startTime As Integer = My.Computer.Clock.TickCount
            Dim dt As DataTable = Nothing
            Dim ctr As Integer = 0

            _progress = "C-OPT - Engine - PostSolve...Started at: " & My.Computer.Clock.LocalTime.ToString()
            Debug.Print(_progress)
            Console.WriteLine(_progress)
            EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - PostSolve - Started at: ", My.Computer.Clock.LocalTime.ToString())

            If GenUtils.IsSwitchAvailable(_switches, "/GenOutputFilesOnly") Then
                _progress = "C-OPT - Engine - PostSolve...GenOutputFiles - Started at: " & My.Computer.Clock.LocalTime.ToString()
                Debug.Print(_progress)
                Console.WriteLine(_progress)
                EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - PostSolve - GenOutputFiles - Started at: ", My.Computer.Clock.LocalTime.ToString())

                SaveRun()

                _progress = "C-OPT - Engine - PostSolve...GenOutputFiles - Ended at: " & My.Computer.Clock.LocalTime.ToString()
                Debug.Print(_progress)
                Console.WriteLine(_progress)
                EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - PostSolve - GenOutputFiles - Ended at: ", My.Computer.Clock.LocalTime.ToString())
            Else
                'Check out of bounds in tsysCOL
                Try
                    SolutionBadResults = "0"
                    Debug.Print(_miscParams.Item("BAD_SOLVER_RESULTS_QUERY").ToString())
                    _lastSql = "SELECT COUNT(*) AS RowCount FROM [" & _miscParams.Item("BAD_SOLVER_RESULTS_QUERY").ToString() & "]"
                    dt = _currentDb.GetDataTable(_lastSql)
                    If dt IsNot Nothing Then
                        If dt.Rows(0).Item("RowCount").ToString() <> "0" Then
                            'MsgBox("More than zero.")
                            _progress = "C-OPT - Engine - PostSolve...Bad Solver Results: " & dt.Rows(0).Item("RowCount").ToString()
                            Debug.Print(_progress)
                            Console.WriteLine(_progress)

                            EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - PostSolve: ", " - Bad Solver Results - " & dt.Rows(0).Item("RowCount").ToString())
                            SolutionBadResults = dt.Rows(0).Item("RowCount").ToString()
                            'MessageBox.Show("C-OPT has encountered bad solver results (" & dt.Rows(0).Item("RowCount").ToString() & " invalid rows) coming in during this run." & vbNewLine & "Please verify and run the model again since the current results might be unreliable!", "C-OPT", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        End If
                    End If
                Catch ex As Exception
                    'Ignore this error.
                    'Debug.Print(ex.Message)
                End Try

                'qrepInfeasibilitySummary
                Try
                    SolutionInfeasibilities = "0"
                    'Debug.Print(_miscParams.Item("BAD_SOLVER_RESULTS_QUERY").ToString())
                    _lastSql = "SELECT SUM(VAL) AS Total FROM [" & "qrepInfeasibilitySummary" & "]"
                    dt = _currentDb.GetDataTable(_lastSql)
                    If dt IsNot Nothing Then
                        If dt.Rows(0).Item("Total").ToString() <> "0" Then
                            'MsgBox("More than zero.")
                            SolutionInfeasibilities = dt.Rows(0).Item("Total").ToString()
                            _lastSql = "SELECT * FROM [" & "qrepInfeasibilitySummary" & "]"
                            dt = _currentDb.GetDataTable(_lastSql)
                            For ctr = 0 To dt.Rows.Count - 1
                                _progress = "C-OPT - Engine - PostSolve...Infeasibilities: " & dt.Rows(ctr).Item(0).ToString() & " - " & dt.Rows(ctr).Item(1).ToString()
                                Debug.Print(_progress)
                                Console.WriteLine(_progress)

                                EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - PostSolve: ", " - Infeasibilities - " & dt.Rows(ctr).Item(0).ToString() & " - " & dt.Rows(ctr).Item(1).ToString())
                            Next

                            'MessageBox.Show("C-OPT has encountered bad solver results (" & dt.Rows(0).Item("RowCount").ToString() & " invalid rows) coming in during this run." & vbNewLine & "Please verify and run the model again since the current results might be unreliable!", "C-OPT", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        End If
                    End If
                Catch ex As Exception
                    'Ignore this error.
                    'Debug.Print(ex.Message)
                End Try

                _progress = "C-OPT - Engine - PostSolve...SaveResultsToElementTables - Started at: " & My.Computer.Clock.LocalTime.ToString()
                Debug.Print(_progress)
                Console.WriteLine(_progress)
                EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - PostSolve - SaveResultsToElementTables - Started at: ", My.Computer.Clock.LocalTime.ToString())

                SaveResultsToElementTables()

                _progress = "C-OPT - Engine - PostSolve...SaveResultsToElementTables - Ended at: " & My.Computer.Clock.LocalTime.ToString()
                Debug.Print(_progress)
                Console.WriteLine(_progress)
                EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - PostSolve - SaveResultsToElementTables - Ended at: ", My.Computer.Clock.LocalTime.ToString())
                EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - SaveResultsToElementTables: ", Space(9) & GenUtils.FormatTime(startTime, My.Computer.Clock.TickCount))
                'RunQueries(ByVal queryName as string)
                'GenerateRUNFile()

                'RKP/08-06-08
                'POSTPROC needs to run right before SaveRun
                RunSysModelQueries("POSTPROC")  'stm

                _progress = "C-OPT - Engine - PostSolve...GenOutputFiles - Started at: " & My.Computer.Clock.LocalTime.ToString()
                Debug.Print(_progress)
                Console.WriteLine(_progress)
                EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - PostSolve - GenOutputFiles - Started at: ", My.Computer.Clock.LocalTime.ToString())

                SaveRun()

                _progress = "C-OPT - Engine - PostSolve...GenOutputFiles - Ended at: " & My.Computer.Clock.LocalTime.ToString()
                Debug.Print(_progress)
                Console.WriteLine(_progress)
                EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - PostSolve - GenOutputFiles - Ended at: ", My.Computer.Clock.LocalTime.ToString())
            End If

            _progress = "C-OPT - Engine - PostSolve...Ended at: " & My.Computer.Clock.LocalTime.ToString()
            Debug.Print(_progress)
            Console.WriteLine(_progress)
            EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - PostSolve - Ended at: ", My.Computer.Clock.LocalTime.ToString())

            EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - PostSolve: ", Space(30) & GenUtils.FormatTime(startTime, My.Computer.Clock.TickCount))

            '_miscParams.Item("PRJ_NAME").ToString
            '_miscParams.Item("RUN_NAME").ToString

        End Sub

        Private Function UnBrack(ByVal strIn As String) As String

            If (Left(strIn, 1) = "[") Then
                strIn = Right(strIn, Len(strIn) - 1)
            End If

            If (Right(strIn, 1) = "]") Then
                strIn = Left(strIn, Len(strIn) - 1)
            End If

            Return strIn

        End Function

        Public Sub SaveMatrix()
            Dim ds As New DataSet
            Dim tableName As String
            Dim outputFile As String
            Dim startTime As Integer = My.Computer.Clock.TickCount

            tableName = "tsysRow"
            outputFile = _workDir & "\" & tableName & ".xml"
            Try
                _currentDb.GetDataSet("SELECT * FROM " & tableName).WriteXml(outputFile, XmlWriteMode.WriteSchema)
            Catch ex As Exception

            End Try

            tableName = "tsysCol"
            outputFile = _workDir & "\" & tableName & ".xml"
            Try
                _currentDb.GetDataSet("SELECT * FROM " & tableName).WriteXml(outputFile, XmlWriteMode.WriteSchema)
            Catch ex As Exception

            End Try

            tableName = "tsysMtx"
            outputFile = _workDir & "\" & tableName & ".xml"
            Try
                _currentDb.GetDataSet("SELECT * FROM " & tableName).WriteXml(outputFile, XmlWriteMode.WriteSchema)
            Catch ex As Exception

            End Try


            EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - SaveMatrix took: ", GenUtils.FormatTime(startTime, My.Computer.Clock.TickCount))
        End Sub

        Private Function GetDefColQueries() As DataTable
            Return _currentDb.GetDataTable("SELECT 'COL' AS DefType, ColTypeRecSet AS SourceQuery FROM tsysDefCol")
        End Function

        Private Function GetDefRowQueries() As DataTable
            Return _currentDb.GetDataTable("SELECT 'ROW' AS DefType, RowTypeRecSet AS SourceQuery FROM tsysDefRow")
        End Function

        Private Function GetDefCoeffQueries() As DataTable
            Return _currentDb.GetDataTable("SELECT 'COEFF' AS DefType, CoeffRecSet AS SourceQuery FROM tsysDefCoef")
        End Function

        Private Function GetDefQueries() As DataTable
            Dim defTable As DataTable

            defTable = GetDefColQueries()
            defTable.Merge(GetDefRowQueries())
            defTable.Merge(GetDefCoeffQueries())

            Return defTable
        End Function

        Public Function CurrentDb() As EntLib.COPT.DAAB
            Return _currentDb
        End Function

        'Public Property CurrentDb() As EntLib.COPT.DAAB
        '    Get
        '        Return _currentDb
        '    End Get
        '    Set(ByVal value As EntLib.COPT.DAAB)
        '        _currentDb = value
        '    End Set
        'End Property

        Public Sub SetCurrentDb(ByRef currentDb As EntLib.COPT.DAAB)
            _currentDb = currentDb
        End Sub

        Public ReadOnly Property DatabaseName() As String
            Get
                Return _databaseName
            End Get
        End Property

        ''' <summary>This is a SET property that allows to change the C-OPT database</summary>
        ''' <remarks>
        ''' RKP/01-31-08/v2.0
        ''' </remarks> 
        Public WriteOnly Property ChangeDb() As String
            Set(ByVal value As String)
                _currentDb = New EntLib.COPT.DAAB(value)
            End Set
        End Property

        ''' <summary>This method saves the results of a run to the element tables</summary>
        ''' <remarks>
        ''' Adapted from C-OPT1 codebase
        ''' </remarks> 
        ''' <author>RKP</author>
        ''' <date>02-06-08</date>
        ''' <version>2.0</version>
        Private Sub SaveResultsToElementTables()
            '//================================================================================//
            '/|   FUNCTION: SsvElmTbls
            '/| PARAMETERS: -NONE-
            '/|    RETURNS: True on Success and False by default or Failure
            '/|    PURPOSE: Store solution values for the individual LP vectors and rows
            '/|      USAGE: i= SsvElmTbls()
            '/|         BY: Sean
            '/|       DATE: 04/15/1997
            '/|    HISTORY: 04/15/1997  Added feature to generate the append queries
            '/|             02/17/2004  Adapted to ADO
            '/|             02/17/2008  Adapted to ADO.NET (RKP)
            '//================================================================================//

            'Dim sql As String
            Dim dt As DataTable
            Dim retValue As Integer
            Dim ctr As Integer
            Dim colTemp As typColType
            Dim rowTemp As typRowType
            Dim startTime As Integer = My.Computer.Clock.TickCount

            Try
                _lastSql = "SELECT * FROM " & _miscParams.Item("LPM_COLUMN_DEF_TABLE_NAME").ToString
                dt = _currentDb.GetDataTable(_lastSql)

                _progress = "----- C O L U M N S -----"
                Debug.Print(_progress)
                Console.WriteLine(_progress)
                For ctr = 0 To dt.Rows.Count - 1
                    Application.DoEvents()
                    If Cancel Then Exit Sub 'Return GenUtils.ReturnStatus.Failure
                    colTemp = ReadColType(CInt(dt.Rows(ctr)("ColTypeID").ToString()))
                    If CBool(colTemp.intActive) Then
                        _lastSql = MakeUpdColSlnQry(colTemp)

                        _progress = colTemp.strType
                        Debug.Print(_progress)
                        Console.WriteLine(_progress)
                        retValue = _currentDb.ExecuteNonQuery(_lastSql)
                    End If
                Next

                '// R O W S //
                _lastSql = "SELECT * FROM " & _miscParams.Item("LPM_CONSTR_DEF_TABLE_NAME").ToString
                dt = _currentDb.GetDataTable(_lastSql)

                _progress = "----- R O W S -----"
                Debug.Print(_progress)
                Console.WriteLine(_progress)
                For ctr = 0 To dt.Rows.Count - 1
                    Application.DoEvents()
                    If Cancel Then Exit Sub 'Return GenUtils.ReturnStatus.Failure
                    rowTemp = ReadRowType(CInt(dt.Rows(ctr)("RowTypeID").ToString())) 'RKP/20080522/v2.1.118 - Moved this line up where it should belong
                    If CBool(rowTemp.intActive) Then    'STM 20080508
                        'rowTemp = ReadRowType(CInt(dt.Rows(ctr)("RowTypeID").ToString()))
                        _lastSql = MakeUpdRowSlnQry(rowTemp)
                        _progress = rowTemp.strType
                        Debug.Print(_progress)
                        Console.WriteLine(_progress)
                        retValue = _currentDb.ExecuteNonQuery(_lastSql)
                    End If
                Next
            Catch ex As Exception
                'MsgBox(ex.Message)
                GenUtils.Message(GenUtils.MsgType.Critical, "Engine - SaveResultsToElementTables", ex.Message)
            End Try

            EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - SaveResultsToElementTables took: ", GenUtils.FormatTime(startTime, My.Computer.Clock.TickCount))

        End Sub

        Private Function MakeUpdColSlnQry(ByVal colTemp As typColType) As String
            'Dim sql As String
            Dim sb As New Text.StringBuilder

            'sb.AppendLine("UPDATE DISTINCT ")

            If CurrentDb.IsSQLExpress Then
                sb.AppendLine("UPDATE ")
                sb.Append(_miscParams.Item("LPM_COL_ELEMENT_TABLE_PRE").ToString())
                sb.AppendLine(colTemp.strTable)
                sb.AppendLine(" SET ")
                sb.Append(_miscParams.Item("LPM_COL_ELEMENT_TABLE_PRE").ToString())
                sb.Append(colTemp.strTable)
                sb.Append(".ACTIVITY = ")
                sb.Append(_miscParams.Item("LPM_COLUMN_TABLE_NAME").ToString())
                sb.AppendLine(".ACTIVITY,")
                sb.Append(_miscParams.Item("LPM_COL_ELEMENT_TABLE_PRE").ToString())
                sb.Append(colTemp.strTable)
                sb.Append(".DJ = ")
                sb.Append(_miscParams.Item("LPM_COLUMN_TABLE_NAME").ToString())
                sb.AppendLine(".DJ")
                sb.AppendLine("FROM ")
                sb.Append(_miscParams.Item("LPM_COL_ELEMENT_TABLE_PRE").ToString())
                sb.AppendLine(colTemp.strTable)
                sb.AppendLine(" INNER JOIN ")
                sb.AppendLine(_miscParams.Item("LPM_COLUMN_TABLE_NAME").ToString())
                sb.AppendLine(" ON ")
                sb.Append(_miscParams.Item("LPM_COL_ELEMENT_TABLE_PRE").ToString())
                sb.Append(colTemp.strTable)
                sb.Append(".")
                sb.Append(colTemp.strType)
                sb.Append("ID = ")
                sb.Append(_miscParams.Item("LPM_COLUMN_TABLE_NAME").ToString())
                sb.AppendLine(".ColID")

            Else
                sb.AppendLine("UPDATE ")
                sb.Append(_miscParams.Item("LPM_COL_ELEMENT_TABLE_PRE").ToString())
                sb.AppendLine(colTemp.strTable)
                sb.AppendLine(" INNER JOIN ")
                sb.AppendLine(_miscParams.Item("LPM_COLUMN_TABLE_NAME").ToString())
                sb.AppendLine(" ON ")
                sb.Append(_miscParams.Item("LPM_COL_ELEMENT_TABLE_PRE").ToString())
                sb.Append(colTemp.strTable)
                sb.Append(".")
                sb.Append(colTemp.strType)
                sb.Append("ID = ")
                sb.Append(_miscParams.Item("LPM_COLUMN_TABLE_NAME").ToString())
                sb.AppendLine(".ColID")
                sb.AppendLine(" SET ")
                sb.Append(_miscParams.Item("LPM_COL_ELEMENT_TABLE_PRE").ToString())
                sb.Append(colTemp.strTable)
                sb.Append(".ACTIVITY = ")
                sb.Append(_miscParams.Item("LPM_COLUMN_TABLE_NAME").ToString())
                sb.AppendLine(".ACTIVITY,")
                sb.Append(_miscParams.Item("LPM_COL_ELEMENT_TABLE_PRE").ToString())
                sb.Append(colTemp.strTable)
                sb.Append(".DJ = ")
                sb.Append(_miscParams.Item("LPM_COLUMN_TABLE_NAME").ToString())
                sb.Append(".DJ")
            End If

            Return sb.ToString()
        End Function

        Private Function MakeUpdRowSlnQry(ByVal rowTemp As typRowType) As String
            Dim sb As New Text.StringBuilder

            'sb.AppendLine("UPDATE DISTINCT ")


            If CurrentDb.IsSQLExpress Then
                sb.AppendLine("UPDATE ")
                sb.Append(_miscParams.Item("LPM_ROW_ELEMENT_TABLE_PRE").ToString())
                sb.AppendLine(rowTemp.strTable)
                sb.AppendLine(" SET ")
                sb.Append(_miscParams.Item("LPM_ROW_ELEMENT_TABLE_PRE").ToString())
                sb.Append(rowTemp.strTable)
                sb.Append(".ACTIVITY = ")
                sb.Append(_miscParams.Item("LPM_CONSTR_TABLE_NAME").ToString())
                sb.AppendLine(".ACTIVITY,")
                sb.Append(_miscParams.Item("LPM_ROW_ELEMENT_TABLE_PRE").ToString())
                sb.Append(rowTemp.strTable)
                sb.Append(".SHADOW = ")
                sb.Append(_miscParams.Item("LPM_CONSTR_TABLE_NAME").ToString())
                sb.AppendLine(".SHADOW")
                sb.AppendLine("FROM ")
                sb.Append(_miscParams.Item("LPM_ROW_ELEMENT_TABLE_PRE").ToString())
                sb.AppendLine(rowTemp.strTable)
                sb.AppendLine(" INNER JOIN ")
                sb.AppendLine(_miscParams.Item("LPM_CONSTR_TABLE_NAME").ToString())
                sb.AppendLine(" ON ")
                sb.Append(_miscParams.Item("LPM_ROW_ELEMENT_TABLE_PRE").ToString())
                sb.Append(rowTemp.strTable)
                sb.Append(".")
                sb.Append(rowTemp.strType)
                sb.Append("ID = ")
                sb.Append(_miscParams.Item("LPM_CONSTR_TABLE_NAME").ToString())
                sb.AppendLine(".RowID")

            Else
                sb.AppendLine("UPDATE ")
                sb.Append(_miscParams.Item("LPM_ROW_ELEMENT_TABLE_PRE").ToString())
                sb.AppendLine(rowTemp.strTable)
                sb.AppendLine(" INNER JOIN ")
                sb.AppendLine(_miscParams.Item("LPM_CONSTR_TABLE_NAME").ToString())
                sb.AppendLine(" ON ")
                sb.Append(_miscParams.Item("LPM_ROW_ELEMENT_TABLE_PRE").ToString())
                sb.Append(rowTemp.strTable)
                sb.Append(".")
                sb.Append(rowTemp.strType)
                sb.Append("ID = ")
                sb.Append(_miscParams.Item("LPM_CONSTR_TABLE_NAME").ToString())
                sb.AppendLine(".RowID")
                sb.AppendLine(" SET ")
                sb.Append(_miscParams.Item("LPM_ROW_ELEMENT_TABLE_PRE").ToString())
                sb.Append(rowTemp.strTable)
                sb.Append(".ACTIVITY = ")
                sb.Append(_miscParams.Item("LPM_CONSTR_TABLE_NAME").ToString())
                sb.AppendLine(".ACTIVITY,")
                sb.Append(_miscParams.Item("LPM_ROW_ELEMENT_TABLE_PRE").ToString())
                sb.Append(rowTemp.strTable)
                sb.Append(".SHADOW = ")
                sb.Append(_miscParams.Item("LPM_CONSTR_TABLE_NAME").ToString())
                sb.AppendLine(".SHADOW")
            End If



            Return sb.ToString()
        End Function

        Private Function ArchiveSysModelQueries() As Boolean
            Dim dt As DataTable
            Dim dtQ As DataTable
            Dim ctr As Integer
            Dim ctrQ As Integer

            _lastSql = "SELECT tsysModelQueries.* FROM tsysModelQueries WHERE LEN(tsysModelQueries.Name) >= 1"
            dt = _currentDb.GetDataTable(_lastSql)

            dtQ = _currentDb.GetTablesViews
            For ctr = 0 To dt.Rows.Count - 1
                Application.DoEvents()
                If Cancel Then Return GenUtils.ReturnStatus.Failure
                For ctrQ = 0 To dtQ.Rows.Count - 1
                    If Cancel Then Return GenUtils.ReturnStatus.Failure
                    If dtQ.Rows(ctrQ).Item("QUERY_NAME").ToString() = UnBrack(dt.Rows(ctr).Item("Name").ToString()) Then
                        dt.Rows(ctr).Item("SQLTEXT") = dtQ.Rows(ctrQ).Item("QUERY_DEFINITION").ToString()
                        dt.Rows(ctr).Item("SQLDATE") = Now()
                        Exit For
                    End If
                Next
                _lastSql = "SELECT * FROM "
                '_currentDb.UpdateDataSet(
            Next

        End Function

        Public Function getSolutionStatusDescription() As String
            Return _solutionStatus
        End Function

        Public Function getSolutionRows() As Integer
            Return _solutionRows
        End Function

        Public Function getSolutionColumns() As Integer
            Return _solutionColumns
        End Function

        Public Function getSolutionNonZeros() As Integer
            Return _solutionNonZeros
        End Function

        Public Function getSolutionObj() As Double
            Return _solutionObj
        End Function

        Public Property SolutionStatus() As String
            Get
                Return _solutionStatus
            End Get
            Set(ByVal value As String)
                _solutionStatus = value
            End Set
        End Property

        Public Property SolutionStatusCode() As Integer
            Get
                Return _solutionStatusCode
            End Get
            Set(ByVal value As Integer)
                _solutionStatusCode = value
            End Set
        End Property
        Public Function CheckDBBlueprint() As String
            Dim _lastSql As String
            Dim sSQL As String
            Dim i As Integer
            Dim z As Integer
            Dim dt As DataTable
            Dim sRet As String
            Dim sTemp As String
            CheckDBBlueprint = ""
            sRet = ""
            sTemp = ""

            sSQL = "SELECT DISTINCT " & vbCrLf & _
                   "   'PRIMARY' AS TYPE, " & vbCrLf & _
                   "   ModelOBJTYPE, " & vbCrLf & _
                   "   ModelOBJTYPEdesc, " & vbCrLf & _
                   "   ObjectActive, " & vbCrLf & _
                   "   SourceOfModelObject, " & vbCrLf & _
                   "   ModelObjectName, " & vbCrLf & _
                   "   ModelObjectDesc, " & vbCrLf & _
                   "   RecSetName " & vbCrLf & _
                   "FROM " & vbCrLf & _
                   "   qsysModelObjects " & vbCrLf
            _lastSql = sSQL

            '            C-OPT Engine - BluePrint akjsdfjajjajjjljadfj OK
            '            C-OPT Engine - BluePrint took: ,                      0m 3s
            '            EntLib.COPT.Log.Log(_workDir, "MatGen - Error", "/UseMSAccessSyntax switch turned on but field, LINKED_MTX_DB_PATH, was not found in qsysMiscParams - " & ex.Message)

            Try
                dt = _currentDb.GetDataSet(_lastSql, True).Tables(0)
                For i = 0 To dt.Rows.Count - 1
                    Application.DoEvents()
                    z = CheckDBObject(dt.Rows(i).Item("RecSetName").ToString)
                    Select Case z
                        Case -9
                            sTemp = sTemp & "*E*  BLUEPRINT_ERROR ... OBJECT " & dt.Rows(i).Item("RecSetName").ToString & " DOES NOT EXIST " & vbCrLf
                            sTemp = sTemp & "                         SEE " & dt.Rows(i).Item("SourceOfModelObject").ToString & _
                                                                      " " & dt.Rows(i).Item("ModelObjectName").ToString & _
                                                                      " " & dt.Rows(i).Item("ModelObjectDesc").ToString & _
                                                                      " " & dt.Rows(i).Item("RecSetName").ToString & _
                                                                      vbCrLf
                        Case 0
                            sTemp = sTemp & "*W*  BLUEPRINT_ERROR ... OBJECT " & dt.Rows(i).Item("RecSetName").ToString & " RETURNS 0 ROWS " & vbCrLf
                        Case Else
                            sTemp = sTemp & "                         " & dt.Rows(i).Item("RecSetName").ToString & " ... OK " & vbCrLf
                    End Select
                    sRet = sRet & sTemp
                    sTemp = ""
                Next i
            Catch ex As Exception
                'GenUtils.Message(GenUtils.MsgType.Information, "Engine - RunAll-CheckDBBlueprint", ex.Message)
                EntLib.COPT.Log.Log(_workDir, "Engine - RunAll-CheckDBBlueprint", ex.Message)
            End Try

            CheckDBBlueprint = sRet
        End Function

        Public Function CheckDBObject(ByVal strTableOrQuery As String) As Integer
            'usage:  i = CheckDBObject("qmtxC_05_FACFIX")
            'returns -9 on failure, row count otherwise.
            Dim dTable As DataTable
            _lastSql = "SELECT COUNT(*) AS ROW_COUNT FROM " & strTableOrQuery
            Try
                dTable = _currentDb.GetDataSet(_lastSql, True).Tables(0)
                Return dTable.Rows(0).Item("ROW_COUNT")
            Catch ex As Exception
                Return -9    'query chokes, most likely strTableOrQuery does not exist
            End Try
        End Function

        Public Property CommonSolutionStatus() As String
            Get
                Return _commonSolutionStatus
            End Get
            Set(ByVal value As String)
                _commonSolutionStatus = value
            End Set
        End Property

        Public Property CommonSolutionStatusCode() As Integer
            Get
                Return _commonSolutionStatusCode
            End Get
            Set(ByVal value As Integer)
                _commonSolutionStatusCode = value
            End Set
        End Property

        Public Property SolutionRows() As Integer
            Get
                Return _solutionRows
            End Get
            Set(ByVal value As Integer)
                _solutionRows = value
            End Set
        End Property

        Public Property SolutionColumns() As Integer
            Get
                Return _solutionColumns
            End Get
            Set(ByVal value As Integer)
                _solutionColumns = value
            End Set
        End Property

        Public Property SolutionNonZeros() As Integer
            Get
                Return _solutionNonZeros
            End Get
            Set(ByVal value As Integer)
                _solutionNonZeros = value
            End Set
        End Property

        Public Property SolutionObj() As Double
            Get
                Return _solutionObj
            End Get
            Set(ByVal value As Double)
                _solutionObj = value
            End Set
        End Property

        Public Property SolverName() As String
            Get
                Return _solverName
            End Get
            Set(ByVal value As String)
                _solverName = value
            End Set
        End Property

        Public Property SolverVersion() As String
            Get
                Return _solverVersion
            End Get
            Set(ByVal value As String)
                _solverVersion = value
            End Set
        End Property

        Public Property SolutionTime() As String
            Get
                Return _solutionTime
            End Get
            Set(ByVal value As String)
                _solutionTime = value
            End Set
        End Property

        Public ReadOnly Property SolutionProjectName() As String
            Get
                Try
                    Return _miscParams.Item("PRJ_NAME").ToString()
                Catch ex As Exception
                    Return "_EMPTY_"
                End Try
            End Get
        End Property

        Public ReadOnly Property SolutionRunName() As String
            Get
                Try
                    Return _miscParams.Item("RUN_NAME").ToString()
                Catch ex As Exception
                    Return "_EMPTY_"
                End Try

            End Get
        End Property

        Public Property SolutionIterations() As Integer
            Get
                Return _solutionIterations
            End Get
            Set(ByVal value As Integer)
                _solutionIterations = value
            End Set
        End Property

        Public Property ProblemType() As problemRunType
            Get
                'Select Case _problemType
                '    Case ProblemRunType.problemTypeContinuous  '"C"
                '        Return "Continuous"
                '    Case ProblemRunType.problemTypeInteger  '"I"
                '        Return "MIP (Integer)"
                '    Case ProblemRunType.problemTypeBinary  '"B"
                '        Return "MIP (Binary)"
                '    Case ProblemRunType.problemTypeBinaryAndInteger
                '        Return "MIP (Binary & Integer)"
                '    Case Else
                Return _problemType
                'End Select
                'Return _problemType
            End Get
            Set(ByVal value As problemRunType)
                _problemType = value
            End Set
        End Property

        Public Function getProblemType(ByVal problemType As problemRunType) As String
            Select Case problemType
                Case problemRunType.problemTypeContinuous  '"C"
                    Return "Continuous"
                Case problemRunType.problemTypeInteger  '"I"
                    Return "MIP (Integer)"
                Case problemRunType.problemTypeBinary  '"B"
                    Return "MIP (Binary)"
                Case problemRunType.problemTypeBinary + problemRunType.problemTypeInteger
                    Return "MIP (Binary & Integer)"
                Case Else
                    Return problemType
            End Select
        End Function

        Public Sub SaveRun()
            Dim genUtils As New GenUtils(_currentDb)

            'genUtils.SaveRun(_currentDb.GetDataTable("SELECT * FROM " & COPT_RSLT_TABLE_NAME & " WHERE ACTIVE = True"), _miscParams, GenUtils.GetSwitchArgument(_switches, "/WorkDir", 1), GenUtils.IsSwitchAvailable(_switches, "/SaveRunXML")) '_workDir)
            'genUtils.SaveRun(_currentDb.GetDataTable("SELECT * FROM " & COPT_RSLT_TABLE_NAME & " WHERE ACTIVE = True"), _miscParams, GenUtils.GetWorkDir(_switches), GenUtils.IsSwitchAvailable(_switches, "/SaveRunXML")) '_workDir)

            'genUtils.SaveRun(_currentDb.GetDataTable("SELECT * FROM " & COPT_RSLT_TABLE_NAME & " WHERE ACTIVE = True"), _miscParams, GenUtils.GetWorkDir(_switches), GenUtils.IsSwitchAvailable(_switches, "/SaveRunXML")) '_workDir)

            If CurrentDb.IsSQLExpress Then
                genUtils.SaveRun(_currentDb.GetDataTable("SELECT * FROM " & COPT_RSLT_TABLE_NAME & " WHERE ACTIVE = 1"), _miscParams, genUtils.GetWorkDir(_switches), TimeStamp, genUtils.IsSwitchAvailable(_switches, "/SaveADOXML"), genUtils.IsSwitchAvailable(_switches, "/SaveRunXML"), True) '_workDir)
            Else
                genUtils.SaveRun(_currentDb.GetDataTable("SELECT * FROM " & COPT_RSLT_TABLE_NAME & " WHERE ACTIVE = True"), _miscParams, genUtils.GetWorkDir(_switches), TimeStamp, genUtils.IsSwitchAvailable(_switches, "/SaveADOXML"), genUtils.IsSwitchAvailable(_switches, "/SaveRunXML"), True) '_workDir)
            End If



            'Dim timeStamp As String = Microsoft.VisualBasic.Format(Now(), "yyyymmdd-hhmm")
            'Dim dt As DataTable = _currentDb.GetDataTable("SELECT * FROM " & COPT_RSLT_TABLE_NAME & " WHERE ACTIVE = True")

            'For ctr As Long = 0 To dt.Rows.Count - 1
            '    Debug.Print(dt.Rows(ctr).Item("RecordsetName").ToString())
            '    'dt.WriteXml(_workDir & "\" & _miscParams.Item("PRJ_NAME").ToString() & "-" & _miscParams.Item("RUN_NAME").ToString() & "-" & dt.Rows(ctr).Item("TextAbbr").ToString() & "-" & Microsoft.VisualBasic.Format(Now(), "yyyymmdd-hhmm") & ".xml")
            '    _currentDb.GetDataTable("SELECT * FROM [" & dt.Rows(ctr).Item("RecordsetName").ToString() & "]").WriteXml(_workDir & "\" & _miscParams.Item("PRJ_NAME").ToString() & "-" & _miscParams.Item("RUN_NAME").ToString() & "-" & dt.Rows(ctr).Item("TextAbbr").ToString() & "-" & timeStamp & ".xml")
            'Next

        End Sub

        Public ReadOnly Property GetWorkDir() As String
            Get
                Return _workDir
            End Get
        End Property

        Private Sub Parallel()
            'Parallel.For(
        End Sub

        Public Property Progress() As String
            Get
                Return _progress & ""
            End Get
            Set(ByVal value As String)
                _progress = value
            End Set
        End Property

        Public Property Cancel() As Boolean
            Get
                Return _cancel
            End Get
            Set(ByVal value As Boolean)
                _cancel = value
            End Set
        End Property

        Public Property SolutionBadResults() As String
            Get
                Return _solutionBadResults
            End Get
            Set(ByVal value As String)
                _solutionBadResults = value
            End Set
        End Property

        Public Property SolutionInfeasibilities() As String
            Get
                Return _solutionInfeasibilities
            End Get
            Set(ByVal value As String)
                _solutionInfeasibilities = value
            End Set
        End Property

        Public Property TimeStamp() As String
            Get
                Return _timeStamp
            End Get
            Set(ByVal value As String)
                _timeStamp = value
            End Set
        End Property

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks>
        ''' RKP/02-21-10/v2.3.131
        ''' </remarks>
        Public ReadOnly Property dtRow() As DataTable
            Get
                Return _dtRow
            End Get
            'Set(ByVal value As DataTable)
            '    _dtRow = value
            'End Set
        End Property

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks>
        ''' RKP/02-21-10/v2.3.131
        ''' </remarks>
        Public ReadOnly Property dtCol() As DataTable
            Get
                Return _dtCol
            End Get
            'Set(ByVal value As DataTable)
            '    _dtCol = value
            'End Set
        End Property

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks>
        ''' RKP/02-21-10/v2.3.131
        ''' </remarks>
        Public ReadOnly Property dtMtx() As DataTable
            Get
                Return _dtMtx
            End Get
            'Set(ByVal value As DataTable)
            '    _dtMtx = value
            'End Set
        End Property

        Public Sub SetSession(ByRef currentSession As EntLib.COPT.Session)
            _currentSession = currentSession
        End Sub

        Public Function GetSession() As EntLib.COPT.Session
            Return _currentSession
        End Function

        Private Sub LoadSolverArrays()
            Dim linqTable As System.Data.EnumerableRowCollection(Of System.Data.DataRow) = Nothing
            Dim linqTable2 As System.Collections.Generic.IEnumerable(Of System.Data.DataRow)
            Dim linqTable3 As System.Collections.Generic.IEnumerable(Of System.Data.DataRow) = Nothing
            Dim dblQueryResults As System.Data.EnumerableRowCollection(Of Double) = Nothing
            Dim intQueryResults As System.Data.EnumerableRowCollection(Of Integer) = Nothing
            Dim charQueryResults As System.Data.EnumerableRowCollection(Of Char) = Nothing
            Dim strQueryResults As System.Data.EnumerableRowCollection(Of String) = Nothing
            Dim drQueryResults As System.Data.EnumerableRowCollection(Of System.Data.DataRow) = Nothing
            Dim intGroups As System.Collections.Generic.IEnumerable(Of Integer) = Nothing
            Dim myArrayList As ArrayList = Nothing
            Dim cumTotal As Integer = 0

            'If GenUtils.IsSwitchAvailable(_switches, "/UseMinSysRes") Then
            's = "MatGen....." & GenUtils.FormatTime(startTime, My.Computer.Clock.TickCount)

            '"SELECT * FROM qsysMtx"
            '"SELECT * FROM qsysMTXwithCOLS WHERE COEF <> 0 ORDER BY ColID, RowID"
            'strSQL = "SELECT ColID, COUNT(*) AS MatCnt FROM [" & _miscParams!LPM_MATRIX_TABLE_NAME.ToString() & "] GROUP BY ColID ORDER BY ColID "

            'GenUtils.Message(GenUtils.MsgType.Information, "Engine-MatGen-LoadSolverArrays", "APM-Before: " & My.Computer.Info.AvailablePhysicalMemory.ToString)
            'GenUtils.Message(GenUtils.MsgType.Information, "Engine - MatGen - LoadSolverArrays", "APM-Before: " & My.Computer.Info.AvailableVirtualMemory.ToString)

            _progress = "Engine-MatGen-LoadSolverArrays-APM-Before: " & My.Computer.Info.AvailablePhysicalMemory.ToString()
            Debug.Print("Engine-MatGen-LoadSolverArrays-APM-Before: " & My.Computer.Info.AvailablePhysicalMemory.ToString())
            Console.WriteLine("Engine-MatGen-LoadSolverArrays-APM-Before: " & My.Computer.Info.AvailablePhysicalMemory.ToString())
            EntLib.COPT.Log.Log(_workDir, " ", "Engine-MatGen-LoadSolverArrays-APM-Before: " & My.Computer.Info.AvailablePhysicalMemory.ToString())

            If dtMtx Is Nothing Then
                Try
                    'RKP/01-28-10/v2.3.128
                    'If qsysMtx is not present, resort to backup option.
                    _dtMtx = _currentDb.GetDataTable("SELECT * FROM qsysMtx")            'SESSIONIZE
                    _srcMtx = "qsysMtx"
                Catch ex As Exception
                    _dtMtx = _currentDb.GetDataTable("SELECT * FROM qsysMTXwithCOLS WHERE COEF <> 0 ORDER BY ColID, RowID")            'SESSIONIZE
                    _srcMtx = "tsysMtx"
                End Try
            End If

            _progress = "Engine-MatGen-LoadSolverArrays-APM-After tsysMtx loaded: " & My.Computer.Info.AvailablePhysicalMemory.ToString()
            Debug.Print("Engine-MatGen-LoadSolverArrays-APM-After tsysMtx loaded: " & My.Computer.Info.AvailablePhysicalMemory.ToString())
            Console.WriteLine("Engine-MatGen-LoadSolverArrays-APM-After tsysMtx loaded: " & My.Computer.Info.AvailablePhysicalMemory.ToString())
            EntLib.COPT.Log.Log(_workDir, " ", "Engine-MatGen-LoadSolverArrays-APM-After tsysMtx loaded: " & My.Computer.Info.AvailablePhysicalMemory.ToString())

            'linqTable = dtMtx.AsEnumerable()
            'Dim queryResults = _
            '    dtCol.AsEnumerable().Join(dtMtx.AsEnumerable(), Function(c) c.Field(Of String)("COL"), _
            '    Function(m) m.Field(Of String)("COL"), _
            '    Function(c, m) _
            '        New With {.ColID = m.Field(Of Integer)("ColID"), _
            '        .RowID = m.Field(Of Integer)("RowID"), _
            '        .COL = m.Field(Of String)("COL"), _
            '        .ROW = m.Field(Of String)("ROW"), _
            '        .COEF = m.Field(Of Double)("COEF")})

            If GenUtils.IsSwitchAvailable(_switches, "/UseMinSysRes") Then

                'linqTable = _
                '    dtCol.AsEnumerable().Join(dtMtx.AsEnumerable(), Function(c) c.Field(Of String)("COL"), _
                '    Function(m) m.Field(Of String)("COL"), _
                '    Function(c, m) _
                '        New With {.ColID = m.Field(Of Integer)("ColID"), _
                '        .RowID = m.Field(Of Integer)("RowID"), _
                '        .COL = m.Field(Of String)("COL"), _
                '        .ROW = m.Field(Of String)("ROW"), _
                '        .COEF = m.Field(Of Double)("COEF")})
                Try
                    'linqTable3 = _
                    '    From r In (From _
                    '        c In dtCol.AsEnumerable(), _
                    '        m In dtMtx.AsEnumerable() _
                    '    Where _
                    '        c!COL = m!COL _
                    '    Select _
                    '        ColID = CInt(m!ColID), _
                    '        RowID = CInt(m!RowID), _
                    '        COL = CStr(m!COL), _
                    '        ROW = CStr(m!ROW), _
                    '        COEF = CDbl(m!COEF) _
                    '    ).Cast(Of System.Data.DataRow)()

                    linqTable = _
                        From _
                            c In dtCol.AsEnumerable(), _
                            m In dtMtx.AsEnumerable() _
                        Where _
                            c!COL = m!COL _
                        Select _
                            DirectCast(m, System.Data.DataRow)

                    'linqTable = DirectCast(linqTable2, System.Data.EnumerableRowCollection(Of System.Data.DataRow))

                    'linqTable2 = linqTable3.AsEnumerable.Cast(Of System.Data.EnumerableRowCollection(Of System.Data.DataRow))
                    'linqTable2 = DirectCast(linqTable3, System.Data.EnumerableRowCollection(Of System.Data.DataRow))
                Catch ex As Exception
                    Debug.Print(ex.Message)
                End Try
                Try
                    linqTable2 = From r In linqTable3.AsEnumerable()
                    Dim a = From r In linqTable2
                Catch ex As Exception
                    Debug.Print(ex.Message)
                End Try

                'Dim dtTemp = linqTable3.CopyToDataTable()
                'linqTable = dtTemp.AsEnumerable()
            Else
                linqTable = dtMtx.AsEnumerable()
            End If

            '---MCNT---dtMtx
            If array_mcnt Is Nothing Then
                'If GenUtils.IsSwitchAvailable(_switches, "/UseMinSysRes") Then
                '    linqTable = _
                '        dtCol.AsEnumerable().Join(dtMtx.AsEnumerable(), Function(c) c.Field(Of String)("COL"), _
                '        Function(m) m.Field(Of String)("COL"), _
                '        Function(c, m) _
                '            New With {.ColID = m.Field(Of Integer)("ColID"), _
                '            .RowID = m.Field(Of Integer)("RowID"), _
                '            .COL = m.Field(Of String)("COL"), _
                '            .ROW = m.Field(Of String)("ROW"), _
                '            .COEF = m.Field(Of Double)("COEF")})
                'Else
                '    linqTable = dtMtx.AsEnumerable()
                'End If

                'If usePLINQ Then linqTable.AsParallel()
                If GenUtils.IsSwitchAvailable(_switches, "/UseMinSysRes") Then
                    intGroups = From r In linqTable.AsEnumerable() _
                                        Order By r("ColID") Ascending _
                                        Group By r!ColID Into g = Group Select CInt(g.Count())
                Else
                    intGroups = From r In linqTable _
                                        Order By r("ColID") Ascending _
                                        Group By r!ColID Into g = Group Select g.Count()
                End If
                Try
                    array_mcnt = intGroups.Cast(Of Integer())()
                Catch ex As Exception
                    Debug.Print(ex.Message)
                End Try

            End If

            '---MBEG---
            If array_mbeg Is Nothing Then
                myArrayList = New ArrayList
                myArrayList.Add(0)
                For ctr = 0 To array_mcnt.Length - 1
                    cumTotal = cumTotal + array_mcnt(ctr)
                    myArrayList.Add(cumTotal)
                Next
                array_mbeg = DirectCast(myArrayList.ToArray(GetType(Integer)), Integer())
            End If

            '---MIDX---dtMtx
            If array_midx Is Nothing Then
                'If GenUtils.IsSwitchAvailable(_switches, "/UseMinSysRes") Then
                '    linqTable = _
                '        dtCol.AsEnumerable().Join(dtMtx.AsEnumerable(), Function(c) c.Field(Of String)("COL"), _
                '        Function(m) m.Field(Of String)("COL"), _
                '        Function(c, m) _
                '            New With {.ColID = m.Field(Of Integer)("ColID"), _
                '            .RowID = m.Field(Of Integer)("RowID"), _
                '            .COL = m.Field(Of String)("COL"), _
                '            .ROW = m.Field(Of String)("ROW"), _
                '            .COEF = m.Field(Of Double)("COEF")})
                'Else
                '    linqTable = dtMtx.AsEnumerable()
                'End If
                'If usePLINQ Then linqTable.AsParallel()
                intQueryResults = From r In linqTable _
                               Order By r("ColID"), r("RowID") _
                               Select IDX = CInt(r("RowID")) - 1
                array_midx = intQueryResults.ToArray()
            End If

            '---MVAL---dtMtx
            If array_mval Is Nothing Then
                'linqTable = dtMtx.AsEnumerable()
                'If GenUtils.IsSwitchAvailable(_switches, "/UseMinSysRes") Then
                '    linqTable = _
                '        dtCol.AsEnumerable().Join(dtMtx.AsEnumerable(), Function(c) c.Field(Of String)("COL"), _
                '        Function(m) m.Field(Of String)("COL"), _
                '        Function(c, m) _
                '            New With {.ColID = m.Field(Of Integer)("ColID"), _
                '            .RowID = m.Field(Of Integer)("RowID"), _
                '            .COL = m.Field(Of String)("COL"), _
                '            .ROW = m.Field(Of String)("ROW"), _
                '            .COEF = m.Field(Of Double)("COEF")})
                'Else
                '    linqTable = dtMtx.AsEnumerable()
                'End If
                'If usePLINQ Then linqTable.AsParallel()
                dblQueryResults = From r In linqTable _
                               Order By r("ColID"), r("RowID") _
                               Select COEF = CDbl(r("COEF"))
                array_mval = dblQueryResults.ToArray()
            End If

            Me.SolutionNonZeros = dtMtx.Rows.Count

            'RKP/06-22-10/v2.3.133
            If GenUtils.IsSwitchAvailable(_switches, "/UseRemoteSolver") Then
                RemoteSolver_UploadData("MTX")
            End If

            _dtMtx = Nothing

            'http://blogs.msdn.com/b/ricom/archive/2004/11/29/271829.aspx
            'Call the Garbage Collector to clean up dtMtx, since it is no longer necessary and memory needs to be reclaimed immediately.
            'GC.Collect()
            'GC.WaitForPendingFinalizers()
            'GC.Collect()

            ''---MCNT---dtMtx
            'strSQL = "SELECT ColID, COUNT(*) AS MatCnt FROM [" & _miscParams!LPM_MATRIX_TABLE_NAME.ToString() & "] GROUP BY ColID ORDER BY ColID "
            '_lastSql = strSQL
            'rs = _currentDb.GetDataSet(_lastSql).Tables(0)
            'Try
            '    strSQL = "SELECT ColID, COUNT(*) AS MatCnt FROM [qsysMtx] GROUP BY ColID ORDER BY ColID "
            '    _lastSql = strSQL
            '    rs = _currentDb.GetDataSet(_lastSql).Tables(0)
            'Catch ex As Exception
            '    strSQL = "SELECT ColID, COUNT(*) AS MatCnt FROM (SELECT * FROM qsysMTXwithCOLS WHERE COEF <> 0 ORDER BY ColID, RowID) GROUP BY ColID ORDER BY ColID "
            '    _lastSql = strSQL
            '    rs = _currentDb.GetDataSet(_lastSql).Tables(0)
            'End Try
            'linqTable = rs.AsEnumerable()
            ''If usePLINQ Then linqTable.AsParallel()
            ''intGroups = From r In linqTable _
            ''                    Order By r("ColID") Ascending _
            ''                    Group By r!ColID Into g = Group Select g.Count()
            'intQueryResults = From r In linqTable _
            '               Order By r("ColID") _
            '               Select MatCnt = CInt(r("MatCnt"))
            'array_mcnt = intQueryResults.ToArray()

            ''---MBEG---
            'myArrayList = New ArrayList
            'myArrayList.Add(0)
            'For ctr = 0 To array_mcnt.Length - 1
            '    cumTotal = cumTotal + array_mcnt(ctr)
            '    myArrayList.Add(cumTotal)
            'Next
            'array_mbeg = DirectCast(myArrayList.ToArray(GetType(Integer)), Integer())

            ''---MIDX---dtMtx
            'strSQL = "SELECT ColID, RowID, (RowID - 1) AS MatIdx FROM [" & _miscParams!LPM_MATRIX_TABLE_NAME.ToString() & "] ORDER BY ColID, RowID "
            '_lastSql = strSQL
            'rs = _currentDb.GetDataSet(_lastSql).Tables(0)
            'linqTable = rs.AsEnumerable()
            ''If usePLINQ Then linqTable.AsParallel()
            'intQueryResults = From r In linqTable _
            '               Order By r("ColID"), r("RowID") _
            '               Select MatIdx = CInt(r("MatIdx"))
            'array_midx = intQueryResults.ToArray()

            ''---MVAL---dtMtx
            'strSQL = "SELECT ColID, RowID, COEF AS MatVal FROM [" & _miscParams!LPM_MATRIX_TABLE_NAME.ToString() & "] ORDER BY ColID, RowID "
            '_lastSql = strSQL
            'rs = _currentDb.GetDataSet(_lastSql).Tables(0)
            'linqTable = rs.AsEnumerable()
            ''If usePLINQ Then linqTable.AsParallel()
            'dblQueryResults = From r In linqTable _
            '               Order By r("ColID"), r("RowID") _
            '               Select MatVal = CDbl(r("MatVal"))
            'array_mval = dblQueryResults.ToArray()

            ''---SENSE---dtRow
            'strSQL = "SELECT ColID, RowID, COEF AS MatVal FROM [" & _miscParams!LPM_MATRIX_TABLE_NAME.ToString() & "] ORDER BY ColID, RowID "
            '_lastSql = strSQL
            'rs = _currentDb.GetDataSet(_lastSql).Tables(0)
            'linqTable = rs.AsEnumerable()
            ''If usePLINQ Then
            ''    'linqTable.AsParallel()
            ''    charQueryResults = From r In linqTable _
            ''                   Order By r("RowID") _
            ''                   Select SENSE = CChar(r("SENSE"))
            ''Else
            'charQueryResults = From r In linqTable _
            '               Order By r("RowID") _
            '               Select SENSE = CChar(r("SENSE"))
            ''End If
            'array_rtyp = charQueryResults.ToArray()
            'End If

            _progress = "Engine-MatGen-LoadSolverArrays-APM-After tsysMtx GC: " & My.Computer.Info.AvailablePhysicalMemory.ToString()
            Debug.Print("Engine-MatGen-LoadSolverArrays-APM-After tsysMtx GC: " & My.Computer.Info.AvailablePhysicalMemory.ToString())
            Console.WriteLine("Engine-MatGen-LoadSolverArrays-APM-After tsysMtx GC: " & My.Computer.Info.AvailablePhysicalMemory.ToString())
            EntLib.COPT.Log.Log(_workDir, " ", "Engine-MatGen-LoadSolverArrays-APM-After tsysMtx GC: " & My.Computer.Info.AvailablePhysicalMemory.ToString())
        End Sub

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <remarks>
        ''' RKP/06-26-10/v2.3.133
        ''' Created to allow solver arrays to be created by joining dtCol and dtMtx every time an array needs to be created.
        ''' </remarks>
        Private Sub LoadSolverArraysUsingJoin()
            Dim sql As String
            Dim linqTable As System.Data.EnumerableRowCollection(Of System.Data.DataRow) = Nothing
            Dim linqTable2 As System.Collections.Generic.IEnumerable(Of System.Data.DataRow)
            Dim linqTable3 As System.Collections.Generic.IEnumerable(Of System.Data.DataRow) = Nothing
            Dim dblQueryResults As System.Data.EnumerableRowCollection(Of Double) = Nothing
            Dim intQueryResults As System.Data.EnumerableRowCollection(Of Integer) = Nothing
            Dim charQueryResults As System.Data.EnumerableRowCollection(Of Char) = Nothing
            Dim strQueryResults As System.Data.EnumerableRowCollection(Of String) = Nothing
            Dim drQueryResults As System.Data.EnumerableRowCollection(Of System.Data.DataRow) = Nothing
            Dim intGroups As System.Collections.Generic.IEnumerable(Of Integer) = Nothing
            Dim myArrayList As ArrayList = Nothing
            Dim cumTotal As Integer = 0

            'If GenUtils.IsSwitchAvailable(_switches, "/UseMinSysRes") Then
            's = "MatGen....." & GenUtils.FormatTime(startTime, My.Computer.Clock.TickCount)

            '"SELECT * FROM qsysMtx"
            '"SELECT * FROM qsysMTXwithCOLS WHERE COEF <> 0 ORDER BY ColID, RowID"
            'strSQL = "SELECT ColID, COUNT(*) AS MatCnt FROM [" & _miscParams!LPM_MATRIX_TABLE_NAME.ToString() & "] GROUP BY ColID ORDER BY ColID "

            'GenUtils.Message(GenUtils.MsgType.Information, "Engine-MatGen-LoadSolverArrays", "APM-Before: " & My.Computer.Info.AvailablePhysicalMemory.ToString)
            'GenUtils.Message(GenUtils.MsgType.Information, "Engine - MatGen - LoadSolverArrays", "APM-Before: " & My.Computer.Info.AvailableVirtualMemory.ToString)

            _progress = "Engine-MatGen-LoadSolverArrays-APM-Before: " & My.Computer.Info.AvailablePhysicalMemory.ToString()
            Debug.Print("Engine-MatGen-LoadSolverArrays-APM-Before: " & My.Computer.Info.AvailablePhysicalMemory.ToString())
            Console.WriteLine("Engine-MatGen-LoadSolverArrays-APM-Before: " & My.Computer.Info.AvailablePhysicalMemory.ToString())
            EntLib.COPT.Log.Log(_workDir, " ", "Engine-MatGen-LoadSolverArrays-APM-Before: " & My.Computer.Info.AvailablePhysicalMemory.ToString())

            If dtMtx Is Nothing Then
                If GenUtils.IsSwitchAvailable(_switches, "/UseSysTables") Then
                    sql = "SELECT * FROM qsysMTXwithCOLS WHERE COEF <> 0 ORDER BY ColID, RowID"
                    _srcMtx = "tsysMtx"
                ElseIf GenUtils.IsSwitchAvailable(_switches, "/UseSysQueries") Then
                    sql = "SELECT * FROM qsysMtx"
                    _srcMtx = "qsysMtx"
                Else
                    sql = "SELECT * FROM qsysMTXwithCOLS WHERE COEF <> 0 ORDER BY ColID, RowID"
                    _srcMtx = "tsysMtx"
                End If
                Try
                    'RKP/01-28-10/v2.3.128
                    'If qsysMtx is not present, resort to backup option.
                    _dtMtx = _currentDb.GetDataTable(sql)            'SESSIONIZE
                Catch ex As Exception
                    '_dtMtx = _currentDb.GetDataTable("SELECT * FROM qsysMTXwithCOLS WHERE COEF <> 0 ORDER BY ColID, RowID")            'SESSIONIZE
                    '_srcMtx = "tsysMtx"
                    GenUtils.Message(GenUtils.MsgType.Critical, "C-OPT Engine-LoadSolverArraysUsingJoin", "Unable to load Coefficient arrays" & vbNewLine & ex.Message)
                End Try
            End If

            _progress = "Engine-MatGen-LoadSolverArrays-APM-After tsysMtx loaded: " & My.Computer.Info.AvailablePhysicalMemory.ToString()
            Debug.Print("Engine-MatGen-LoadSolverArrays-APM-After tsysMtx loaded: " & My.Computer.Info.AvailablePhysicalMemory.ToString())
            Console.WriteLine("Engine-MatGen-LoadSolverArrays-APM-After tsysMtx loaded: " & My.Computer.Info.AvailablePhysicalMemory.ToString())
            EntLib.COPT.Log.Log(_workDir, " ", "Engine-MatGen-LoadSolverArrays-APM-After tsysMtx loaded: " & My.Computer.Info.AvailablePhysicalMemory.ToString())

            'linqTable = dtMtx.AsEnumerable()
            'Dim queryResults = _
            '    dtCol.AsEnumerable().Join(dtMtx.AsEnumerable(), Function(c) c.Field(Of String)("COL"), _
            '    Function(m) m.Field(Of String)("COL"), _
            '    Function(c, m) _
            '        New With {.ColID = m.Field(Of Integer)("ColID"), _
            '        .RowID = m.Field(Of Integer)("RowID"), _
            '        .COL = m.Field(Of String)("COL"), _
            '        .ROW = m.Field(Of String)("ROW"), _
            '        .COEF = m.Field(Of Double)("COEF")})

            If GenUtils.IsSwitchAvailable(_switches, "/UseMinSysRes") Then

                'linqTable = _
                '    dtCol.AsEnumerable().Join(dtMtx.AsEnumerable(), Function(c) c.Field(Of String)("COL"), _
                '    Function(m) m.Field(Of String)("COL"), _
                '    Function(c, m) _
                '        New With {.ColID = m.Field(Of Integer)("ColID"), _
                '        .RowID = m.Field(Of Integer)("RowID"), _
                '        .COL = m.Field(Of String)("COL"), _
                '        .ROW = m.Field(Of String)("ROW"), _
                '        .COEF = m.Field(Of Double)("COEF")})
                Try
                    'linqTable3 = _
                    '    From r In (From _
                    '        c In dtCol.AsEnumerable(), _
                    '        m In dtMtx.AsEnumerable() _
                    '    Where _
                    '        c!COL = m!COL _
                    '    Select _
                    '        ColID = CInt(m!ColID), _
                    '        RowID = CInt(m!RowID), _
                    '        COL = CStr(m!COL), _
                    '        ROW = CStr(m!ROW), _
                    '        COEF = CDbl(m!COEF) _
                    '    ).Cast(Of System.Data.DataRow)()

                    linqTable = _
                        From _
                            c In dtCol.AsEnumerable(), _
                            m In dtMtx.AsEnumerable() _
                        Where _
                            c!COL = m!COL _
                        Select _
                            DirectCast(m, System.Data.DataRow)

                    'linqTable = DirectCast(linqTable2, System.Data.EnumerableRowCollection(Of System.Data.DataRow))

                    'linqTable2 = linqTable3.AsEnumerable.Cast(Of System.Data.EnumerableRowCollection(Of System.Data.DataRow))
                    'linqTable2 = DirectCast(linqTable3, System.Data.EnumerableRowCollection(Of System.Data.DataRow))
                Catch ex As Exception
                    Debug.Print(ex.Message)
                End Try
                Try
                    linqTable2 = From r In linqTable3.AsEnumerable()
                    Dim a = From r In linqTable2
                Catch ex As Exception
                    Debug.Print(ex.Message)
                End Try

                'Dim dtTemp = linqTable3.CopyToDataTable()
                'linqTable = dtTemp.AsEnumerable()
            Else
                linqTable = dtMtx.AsEnumerable()
            End If

            '---MCNT---dtMtx
            If array_mcnt Is Nothing Then
                'If GenUtils.IsSwitchAvailable(_switches, "/UseMinSysRes") Then
                '    linqTable = _
                '        dtCol.AsEnumerable().Join(dtMtx.AsEnumerable(), Function(c) c.Field(Of String)("COL"), _
                '        Function(m) m.Field(Of String)("COL"), _
                '        Function(c, m) _
                '            New With {.ColID = m.Field(Of Integer)("ColID"), _
                '            .RowID = m.Field(Of Integer)("RowID"), _
                '            .COL = m.Field(Of String)("COL"), _
                '            .ROW = m.Field(Of String)("ROW"), _
                '            .COEF = m.Field(Of Double)("COEF")})
                'Else
                '    linqTable = dtMtx.AsEnumerable()
                'End If

                'If usePLINQ Then linqTable.AsParallel()
                If GenUtils.IsSwitchAvailable(_switches, "/UseMinSysRes") Then
                    intGroups = From r In linqTable.AsEnumerable() _
                                        Order By r("ColID") Ascending _
                                        Group By r!ColID Into g = Group Select CInt(g.Count())
                Else
                    linqTable = dtMtx.AsEnumerable()
                    intGroups = From r In linqTable _
                                        Order By r("ColID") Ascending _
                                        Group By r!ColID Into g = Group Select g.Count()
                End If
                Try
                    array_mcnt = intGroups.ToArray()
                Catch ex As Exception
                    Debug.Print(ex.Message)
                    GenUtils.Message(GenUtils.MsgType.Critical, "C-OPT Engine-LoadSolverArraysUsingJoin", "Unable to load Coefficient arrays" & vbNewLine & ex.Message)
                End Try
            End If

            '---MBEG---
            If array_mbeg Is Nothing Then
                myArrayList = New ArrayList
                myArrayList.Add(0)
                For ctr = 0 To array_mcnt.Length - 1
                    cumTotal = cumTotal + array_mcnt(ctr)
                    myArrayList.Add(cumTotal)
                Next
                array_mbeg = DirectCast(myArrayList.ToArray(GetType(Integer)), Integer())
            End If

            '---MIDX---dtMtx
            If array_midx Is Nothing Then
                'If GenUtils.IsSwitchAvailable(_switches, "/UseMinSysRes") Then
                '    linqTable = _
                '        dtCol.AsEnumerable().Join(dtMtx.AsEnumerable(), Function(c) c.Field(Of String)("COL"), _
                '        Function(m) m.Field(Of String)("COL"), _
                '        Function(c, m) _
                '            New With {.ColID = m.Field(Of Integer)("ColID"), _
                '            .RowID = m.Field(Of Integer)("RowID"), _
                '            .COL = m.Field(Of String)("COL"), _
                '            .ROW = m.Field(Of String)("ROW"), _
                '            .COEF = m.Field(Of Double)("COEF")})
                'Else
                '    linqTable = dtMtx.AsEnumerable()
                'End If
                'If usePLINQ Then linqTable.AsParallel()
                intQueryResults = From r In linqTable _
                               Order By r("ColID"), r("RowID") _
                               Select IDX = CInt(r("RowID")) - 1
                array_midx = intQueryResults.ToArray()
            End If

            '---MVAL---dtMtx
            If array_mval Is Nothing Then
                'linqTable = dtMtx.AsEnumerable()
                'If GenUtils.IsSwitchAvailable(_switches, "/UseMinSysRes") Then
                '    linqTable = _
                '        dtCol.AsEnumerable().Join(dtMtx.AsEnumerable(), Function(c) c.Field(Of String)("COL"), _
                '        Function(m) m.Field(Of String)("COL"), _
                '        Function(c, m) _
                '            New With {.ColID = m.Field(Of Integer)("ColID"), _
                '            .RowID = m.Field(Of Integer)("RowID"), _
                '            .COL = m.Field(Of String)("COL"), _
                '            .ROW = m.Field(Of String)("ROW"), _
                '            .COEF = m.Field(Of Double)("COEF")})
                'Else
                '    linqTable = dtMtx.AsEnumerable()
                'End If
                'If usePLINQ Then linqTable.AsParallel()
                dblQueryResults = From r In linqTable _
                               Order By r("ColID"), r("RowID") _
                               Select COEF = CDbl(r("COEF"))
                array_mval = dblQueryResults.ToArray()
            End If

            Me.SolutionNonZeros = dtMtx.Rows.Count

            'RKP/06-22-10/v2.3.133
            If GenUtils.IsSwitchAvailable(_switches, "/UseRemoteSolver") Then
                RemoteSolver_UploadData("MTX")
            End If

            _dtMtx = Nothing

            'http://blogs.msdn.com/b/ricom/archive/2004/11/29/271829.aspx
            'Call the Garbage Collector to clean up dtMtx, since it is no longer necessary and memory needs to be reclaimed immediately.
            'GC.Collect()
            'GC.WaitForPendingFinalizers()
            'GC.Collect()

            ''---MCNT---dtMtx
            'strSQL = "SELECT ColID, COUNT(*) AS MatCnt FROM [" & _miscParams!LPM_MATRIX_TABLE_NAME.ToString() & "] GROUP BY ColID ORDER BY ColID "
            '_lastSql = strSQL
            'rs = _currentDb.GetDataSet(_lastSql).Tables(0)
            'Try
            '    strSQL = "SELECT ColID, COUNT(*) AS MatCnt FROM [qsysMtx] GROUP BY ColID ORDER BY ColID "
            '    _lastSql = strSQL
            '    rs = _currentDb.GetDataSet(_lastSql).Tables(0)
            'Catch ex As Exception
            '    strSQL = "SELECT ColID, COUNT(*) AS MatCnt FROM (SELECT * FROM qsysMTXwithCOLS WHERE COEF <> 0 ORDER BY ColID, RowID) GROUP BY ColID ORDER BY ColID "
            '    _lastSql = strSQL
            '    rs = _currentDb.GetDataSet(_lastSql).Tables(0)
            'End Try
            'linqTable = rs.AsEnumerable()
            ''If usePLINQ Then linqTable.AsParallel()
            ''intGroups = From r In linqTable _
            ''                    Order By r("ColID") Ascending _
            ''                    Group By r!ColID Into g = Group Select g.Count()
            'intQueryResults = From r In linqTable _
            '               Order By r("ColID") _
            '               Select MatCnt = CInt(r("MatCnt"))
            'array_mcnt = intQueryResults.ToArray()

            ''---MBEG---
            'myArrayList = New ArrayList
            'myArrayList.Add(0)
            'For ctr = 0 To array_mcnt.Length - 1
            '    cumTotal = cumTotal + array_mcnt(ctr)
            '    myArrayList.Add(cumTotal)
            'Next
            'array_mbeg = DirectCast(myArrayList.ToArray(GetType(Integer)), Integer())

            ''---MIDX---dtMtx
            'strSQL = "SELECT ColID, RowID, (RowID - 1) AS MatIdx FROM [" & _miscParams!LPM_MATRIX_TABLE_NAME.ToString() & "] ORDER BY ColID, RowID "
            '_lastSql = strSQL
            'rs = _currentDb.GetDataSet(_lastSql).Tables(0)
            'linqTable = rs.AsEnumerable()
            ''If usePLINQ Then linqTable.AsParallel()
            'intQueryResults = From r In linqTable _
            '               Order By r("ColID"), r("RowID") _
            '               Select MatIdx = CInt(r("MatIdx"))
            'array_midx = intQueryResults.ToArray()

            ''---MVAL---dtMtx
            'strSQL = "SELECT ColID, RowID, COEF AS MatVal FROM [" & _miscParams!LPM_MATRIX_TABLE_NAME.ToString() & "] ORDER BY ColID, RowID "
            '_lastSql = strSQL
            'rs = _currentDb.GetDataSet(_lastSql).Tables(0)
            'linqTable = rs.AsEnumerable()
            ''If usePLINQ Then linqTable.AsParallel()
            'dblQueryResults = From r In linqTable _
            '               Order By r("ColID"), r("RowID") _
            '               Select MatVal = CDbl(r("MatVal"))
            'array_mval = dblQueryResults.ToArray()

            ''---SENSE---dtRow
            'strSQL = "SELECT ColID, RowID, COEF AS MatVal FROM [" & _miscParams!LPM_MATRIX_TABLE_NAME.ToString() & "] ORDER BY ColID, RowID "
            '_lastSql = strSQL
            'rs = _currentDb.GetDataSet(_lastSql).Tables(0)
            'linqTable = rs.AsEnumerable()
            ''If usePLINQ Then
            ''    'linqTable.AsParallel()
            ''    charQueryResults = From r In linqTable _
            ''                   Order By r("RowID") _
            ''                   Select SENSE = CChar(r("SENSE"))
            ''Else
            'charQueryResults = From r In linqTable _
            '               Order By r("RowID") _
            '               Select SENSE = CChar(r("SENSE"))
            ''End If
            'array_rtyp = charQueryResults.ToArray()
            'End If

            _progress = "Engine-MatGen-LoadSolverArrays-APM-After tsysMtx GC: " & My.Computer.Info.AvailablePhysicalMemory.ToString()
            Debug.Print("Engine-MatGen-LoadSolverArrays-APM-After tsysMtx GC: " & My.Computer.Info.AvailablePhysicalMemory.ToString())
            Console.WriteLine("Engine-MatGen-LoadSolverArrays-APM-After tsysMtx GC: " & My.Computer.Info.AvailablePhysicalMemory.ToString())
            EntLib.COPT.Log.Log(_workDir, " ", "Engine-MatGen-LoadSolverArrays-APM-After tsysMtx GC: " & My.Computer.Info.AvailablePhysicalMemory.ToString())
        End Sub

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <remarks>
        ''' RKP/06-26-10/v2.3.133
        ''' Created to allow solver arrays to be created by joining dtCol and dtMtx every time an array needs to be created.
        ''' </remarks>
        Private Sub LoadSolverArraysUsingMtxCol(ByRef dtRow As DataTable, ByRef dtMtx As DataTable)
            Dim sql As String
            Dim linqTable As System.Data.EnumerableRowCollection(Of System.Data.DataRow) = Nothing
            Dim linqTable2 As System.Collections.Generic.IEnumerable(Of System.Data.DataRow)
            Dim linqTable3 As System.Collections.Generic.IEnumerable(Of System.Data.DataRow) = Nothing
            Dim dblQueryResults As System.Data.EnumerableRowCollection(Of Double) = Nothing
            Dim intQueryResults As System.Data.EnumerableRowCollection(Of Integer) = Nothing
            Dim charQueryResults As System.Data.EnumerableRowCollection(Of Char) = Nothing
            Dim strQueryResults As System.Data.EnumerableRowCollection(Of String) = Nothing
            Dim drQueryResults As System.Data.EnumerableRowCollection(Of System.Data.DataRow) = Nothing
            Dim intGroups As System.Collections.Generic.IEnumerable(Of Integer) = Nothing
            Dim myArrayList As ArrayList = Nothing
            Dim cumTotal As Integer = 0
            Dim tblCol As String = Me.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString() '"tsysCol"
            Dim tblMtx As String = Me.MiscParams.Item("LPM_MATRIX_TABLE_NAME").ToString() '"tsysMtx"
            Dim isValid As Boolean = False
            Dim qsysMtxCol As String = "qsysMtxCol"

            'If GenUtils.IsSwitchAvailable(_switches, "/UseMinSysRes") Then
            's = "MatGen....." & GenUtils.FormatTime(startTime, My.Computer.Clock.TickCount)

            '"SELECT * FROM qsysMtx"
            '"SELECT * FROM qsysMTXwithCOLS WHERE COEF <> 0 ORDER BY ColID, RowID"
            'strSQL = "SELECT ColID, COUNT(*) AS MatCnt FROM [" & _miscParams!LPM_MATRIX_TABLE_NAME.ToString() & "] GROUP BY ColID ORDER BY ColID "

            'GenUtils.Message(GenUtils.MsgType.Information, "Engine-MatGen-LoadSolverArrays", "APM-Before: " & My.Computer.Info.AvailablePhysicalMemory.ToString)
            'GenUtils.Message(GenUtils.MsgType.Information, "Engine - MatGen - LoadSolverArrays", "APM-Before: " & My.Computer.Info.AvailableVirtualMemory.ToString)

            _progress = "Engine-MatGen-LoadSolverArrays-APM-Before: " & My.Computer.Info.AvailablePhysicalMemory.ToString()
            Debug.Print("Engine-MatGen-LoadSolverArrays-APM-Before: " & My.Computer.Info.AvailablePhysicalMemory.ToString())
            Console.WriteLine("Engine-MatGen-LoadSolverArrays-APM-Before: " & My.Computer.Info.AvailablePhysicalMemory.ToString())
            EntLib.COPT.Log.Log(_workDir, " ", "Engine-MatGen-LoadSolverArrays-APM-Before: " & My.Computer.Info.AvailablePhysicalMemory.ToString())

            isValid = False
            If Not dtMtx Is Nothing Then
                Try
                    If dtMtx.Rows.Count > 0 Then
                        isValid = True
                    End If
                Catch ex As Exception

                End Try
            End If

            If Not isValid Then
                'If GenUtils.IsSwitchAvailable(_switches, "/UseSysTables") Then
                '    sql = "SELECT * FROM qsysMTXwithCOLS WHERE COEF <> 0 ORDER BY ColID, RowID"
                '    _srcMtx = "tsysMtx"
                'ElseIf GenUtils.IsSwitchAvailable(_switches, "/UseSysQueries") Then
                '    sql = "SELECT * FROM qsysMtx"
                '    _srcMtx = "qsysMtx"
                'Else
                '    sql = "SELECT * FROM qsysMTXwithCOLS WHERE COEF <> 0 ORDER BY ColID, RowID"
                '    _srcMtx = "tsysMtx"
                'End If

                If GenUtils.IsSwitchAvailable(_switches, "/UseMinSysRes") Then
                    If GenUtils.IsSwitchAvailable(_switches, "/UseMSAccessSyntax") Then
                        tblCol = "[" & Me.MiscParams.Item("LINKED_COL_DB_PATH").ToString() & "]." & "[" & Me.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString() & "]"
                        tblMtx = "[" & Me.MiscParams.Item("LINKED_MTX_DB_PATH").ToString() & "]." & "[" & Me.MiscParams.Item("LPM_MATRIX_TABLE_NAME").ToString() & "]"
                        _srcMtx = "tsysMtx"
                    End If
                Else
                    'If GenUtils.IsSwitchAvailable(_switches, "/UseSysTables") Then
                    '    sql = "SELECT * FROM qsysMTXwithCOLS WHERE COEF <> 0 ORDER BY ColID, RowID"
                    '    _srcMtx = "tsysMtx"
                    'ElseIf GenUtils.IsSwitchAvailable(_switches, "/UseSysQueries") Then
                    '    sql = "SELECT * FROM qsysMtx"
                    '    _srcMtx = "qsysMtx"
                    'Else
                    '    sql = "SELECT * FROM qsysMTXwithCOLS WHERE COEF <> 0 ORDER BY ColID, RowID"
                    '    _srcMtx = "tsysMtx"
                    'End If
                End If

                'qsysMtxCol
                sql = "SELECT "
                sql = sql & "tsysMTX.ColID, "
                sql = sql & "tsysMTX.RowID, "
                sql = sql & "tsysMTX.COL, "
                sql = sql & "tsysMTX.ROW, "
                sql = sql & "tsysMTX.COEF, "
                sql = sql & "tsysCOL.OBJ, "
                sql = sql & "tsysCOL.LO, "
                sql = sql & "tsysCOL.UP, "
                sql = sql & "tsysCOL.FREE, "
                sql = sql & "tsysCOL.INTGR, "
                sql = sql & "tsysCOL.BINRY, "
                sql = sql & "tsysCOL.SOSTYPE, "
                sql = sql & "tsysCOL.SOSMARKER, "
                sql = sql & "tsysCOL.ACTIVITY, "
                sql = sql & "tsysCOL.DJ, "
                sql = sql & "tsysCOL.STATUS "
                sql = sql & "FROM "
                sql = sql & tblCol & " "
                sql = sql & "INNER JOIN "
                sql = sql & tblMtx & " "
                sql = sql & "ON "
                sql = sql & "tsysCOL.ColID = tsysMTX.ColID "
                sql = sql & "WHERE "
                sql = sql & "tsysMTX.COEF <> 0 "
                sql = sql & "ORDER BY tsysMTX.ColID, tsysMTX.RowID "

                Try
                    'RKP/01-28-10/v2.3.128
                    'If qsysMtx is not present, resort to backup option.
                    _dtMtx = _currentDb.GetDataTable(sql)            'SESSIONIZE
                Catch ex As Exception
                    '_dtMtx = _currentDb.GetDataTable("SELECT * FROM qsysMTXwithCOLS WHERE COEF <> 0 ORDER BY ColID, RowID")            'SESSIONIZE
                    '_srcMtx = "tsysMtx"

                    Try
                        sql = "SELECT * FROM qsysMTXwithCOLS WHERE COEF <> 0 ORDER BY ColID, RowID"
                        _dtMtx = _currentDb.GetDataTable(sql)
                        EntLib.COPT.Log.Log(_workDir, " ", "Engine-MatGen-LoadSolverArraysUsingMtxCol- MtxCol not loaded using primary query: " & ex.Message)
                        EntLib.COPT.Log.Log(_workDir, " ", "Engine-MatGen-LoadSolverArraysUsingMtxCol- MtxCol loaded using fallback query (qsysMTXwithCOLS)")
                    Catch ex2 As Exception
                        GenUtils.Message(GenUtils.MsgType.Critical, "C-OPT Engine-LoadSolverArraysUsingMtxCol", "MtxCol not loaded using fallback query (qsysMTXwithCOLS): " & ex2.Message)
                    End Try
                End Try
            End If

            _progress = "Engine-MatGen-LoadSolverArrays-APM-After dtMtxCol loaded: " & My.Computer.Info.AvailablePhysicalMemory.ToString()
            Debug.Print("Engine-MatGen-LoadSolverArrays-APM-After dtMtxCol loaded: " & My.Computer.Info.AvailablePhysicalMemory.ToString())
            Console.WriteLine("Engine-MatGen-LoadSolverArrays-APM-After tsysMtx loaded: " & My.Computer.Info.AvailablePhysicalMemory.ToString())
            EntLib.COPT.Log.Log(_workDir, " ", "Engine-MatGen-LoadSolverArrays-APM-After dtMtxCol loaded: " & My.Computer.Info.AvailablePhysicalMemory.ToString())

            'linqTable = dtMtx.AsEnumerable()
            'Dim queryResults = _
            '    dtCol.AsEnumerable().Join(dtMtx.AsEnumerable(), Function(c) c.Field(Of String)("COL"), _
            '    Function(m) m.Field(Of String)("COL"), _
            '    Function(c, m) _
            '        New With {.ColID = m.Field(Of Integer)("ColID"), _
            '        .RowID = m.Field(Of Integer)("RowID"), _
            '        .COL = m.Field(Of String)("COL"), _
            '        .ROW = m.Field(Of String)("ROW"), _
            '        .COEF = m.Field(Of Double)("COEF")})

            If GenUtils.IsSwitchAvailable(_switches, "/UseMinSysRes") Then

                'linqTable = _
                '    dtCol.AsEnumerable().Join(dtMtx.AsEnumerable(), Function(c) c.Field(Of String)("COL"), _
                '    Function(m) m.Field(Of String)("COL"), _
                '    Function(c, m) _
                '        New With {.ColID = m.Field(Of Integer)("ColID"), _
                '        .RowID = m.Field(Of Integer)("RowID"), _
                '        .COL = m.Field(Of String)("COL"), _
                '        .ROW = m.Field(Of String)("ROW"), _
                '        .COEF = m.Field(Of Double)("COEF")})
                Try
                    'linqTable3 = _
                    '    From r In (From _
                    '        c In dtCol.AsEnumerable(), _
                    '        m In dtMtx.AsEnumerable() _
                    '    Where _
                    '        c!COL = m!COL _
                    '    Select _
                    '        ColID = CInt(m!ColID), _
                    '        RowID = CInt(m!RowID), _
                    '        COL = CStr(m!COL), _
                    '        ROW = CStr(m!ROW), _
                    '        COEF = CDbl(m!COEF) _
                    '    ).Cast(Of System.Data.DataRow)()

                    linqTable = _
                        From _
                            c In dtCol.AsEnumerable(), _
                            m In _dtMtx.AsEnumerable() _
                        Where _
                            c!COL = m!COL _
                        Select _
                            DirectCast(m, System.Data.DataRow)

                    'linqTable = DirectCast(linqTable2, System.Data.EnumerableRowCollection(Of System.Data.DataRow))

                    'linqTable2 = linqTable3.AsEnumerable.Cast(Of System.Data.EnumerableRowCollection(Of System.Data.DataRow))
                    'linqTable2 = DirectCast(linqTable3, System.Data.EnumerableRowCollection(Of System.Data.DataRow))
                Catch ex As Exception
                    Debug.Print(ex.Message)
                End Try
                Try
                    linqTable2 = From r In linqTable3.AsEnumerable()
                    Dim a = From r In linqTable2
                Catch ex As Exception
                    Debug.Print(ex.Message)
                End Try

                'Dim dtTemp = linqTable3.CopyToDataTable()
                'linqTable = dtTemp.AsEnumerable()
            Else
                linqTable = dtMtx.AsEnumerable()
            End If

            '---MCNT---dtMtx
            If array_mcnt Is Nothing Then
                'If GenUtils.IsSwitchAvailable(_switches, "/UseMinSysRes") Then
                '    linqTable = _
                '        dtCol.AsEnumerable().Join(dtMtx.AsEnumerable(), Function(c) c.Field(Of String)("COL"), _
                '        Function(m) m.Field(Of String)("COL"), _
                '        Function(c, m) _
                '            New With {.ColID = m.Field(Of Integer)("ColID"), _
                '            .RowID = m.Field(Of Integer)("RowID"), _
                '            .COL = m.Field(Of String)("COL"), _
                '            .ROW = m.Field(Of String)("ROW"), _
                '            .COEF = m.Field(Of Double)("COEF")})
                'Else
                '    linqTable = dtMtx.AsEnumerable()
                'End If

                'If usePLINQ Then linqTable.AsParallel()
                If GenUtils.IsSwitchAvailable(_switches, "/UseMinSysRes") Then
                    intGroups = From r In linqTable.AsEnumerable() _
                                        Order By r("ColID") Ascending _
                                        Group By r!ColID Into g = Group Select CInt(g.Count())
                Else
                    linqTable = dtMtx.AsEnumerable()
                    intGroups = From r In linqTable _
                                        Order By r("ColID") Ascending _
                                        Group By r!ColID Into g = Group Select g.Count()
                End If
                Try
                    array_mcnt = intGroups.ToArray()
                Catch ex As Exception
                    Debug.Print(ex.Message)
                    GenUtils.Message(GenUtils.MsgType.Critical, "C-OPT Engine-LoadSolverArraysUsingJoin", "Unable to load Coefficient arrays" & vbNewLine & ex.Message)
                End Try
            End If

            '---MBEG---
            If array_mbeg Is Nothing Then
                myArrayList = New ArrayList
                myArrayList.Add(0)
                For ctr = 0 To array_mcnt.Length - 1
                    cumTotal = cumTotal + array_mcnt(ctr)
                    myArrayList.Add(cumTotal)
                Next
                array_mbeg = DirectCast(myArrayList.ToArray(GetType(Integer)), Integer())
            End If

            '---MIDX---dtMtx
            If array_midx Is Nothing Then
                'If GenUtils.IsSwitchAvailable(_switches, "/UseMinSysRes") Then
                '    linqTable = _
                '        dtCol.AsEnumerable().Join(dtMtx.AsEnumerable(), Function(c) c.Field(Of String)("COL"), _
                '        Function(m) m.Field(Of String)("COL"), _
                '        Function(c, m) _
                '            New With {.ColID = m.Field(Of Integer)("ColID"), _
                '            .RowID = m.Field(Of Integer)("RowID"), _
                '            .COL = m.Field(Of String)("COL"), _
                '            .ROW = m.Field(Of String)("ROW"), _
                '            .COEF = m.Field(Of Double)("COEF")})
                'Else
                '    linqTable = dtMtx.AsEnumerable()
                'End If
                'If usePLINQ Then linqTable.AsParallel()
                intQueryResults = From r In linqTable _
                               Order By r("ColID"), r("RowID") _
                               Select IDX = CInt(r("RowID")) - 1
                array_midx = intQueryResults.ToArray()
            End If

            '---MVAL---dtMtx
            If array_mval Is Nothing Then
                'linqTable = dtMtx.AsEnumerable()
                'If GenUtils.IsSwitchAvailable(_switches, "/UseMinSysRes") Then
                '    linqTable = _
                '        dtCol.AsEnumerable().Join(dtMtx.AsEnumerable(), Function(c) c.Field(Of String)("COL"), _
                '        Function(m) m.Field(Of String)("COL"), _
                '        Function(c, m) _
                '            New With {.ColID = m.Field(Of Integer)("ColID"), _
                '            .RowID = m.Field(Of Integer)("RowID"), _
                '            .COL = m.Field(Of String)("COL"), _
                '            .ROW = m.Field(Of String)("ROW"), _
                '            .COEF = m.Field(Of Double)("COEF")})
                'Else
                '    linqTable = dtMtx.AsEnumerable()
                'End If
                'If usePLINQ Then linqTable.AsParallel()
                dblQueryResults = From r In linqTable _
                               Order By r("ColID"), r("RowID") _
                               Select COEF = CDbl(r("COEF"))
                array_mval = dblQueryResults.ToArray()
            End If

            Me.SolutionNonZeros = dtMtx.Rows.Count

            'RKP/06-22-10/v2.3.133
            If GenUtils.IsSwitchAvailable(_switches, "/UseRemoteSolver") Then
                RemoteSolver_UploadData("MTX")
            End If

            _dtMtx = Nothing

            'http://blogs.msdn.com/b/ricom/archive/2004/11/29/271829.aspx
            'Call the Garbage Collector to clean up dtMtx, since it is no longer necessary and memory needs to be reclaimed immediately.
            'GC.Collect()
            'GC.WaitForPendingFinalizers()
            'GC.Collect()

            ''---MCNT---dtMtx
            'strSQL = "SELECT ColID, COUNT(*) AS MatCnt FROM [" & _miscParams!LPM_MATRIX_TABLE_NAME.ToString() & "] GROUP BY ColID ORDER BY ColID "
            '_lastSql = strSQL
            'rs = _currentDb.GetDataSet(_lastSql).Tables(0)
            'Try
            '    strSQL = "SELECT ColID, COUNT(*) AS MatCnt FROM [qsysMtx] GROUP BY ColID ORDER BY ColID "
            '    _lastSql = strSQL
            '    rs = _currentDb.GetDataSet(_lastSql).Tables(0)
            'Catch ex As Exception
            '    strSQL = "SELECT ColID, COUNT(*) AS MatCnt FROM (SELECT * FROM qsysMTXwithCOLS WHERE COEF <> 0 ORDER BY ColID, RowID) GROUP BY ColID ORDER BY ColID "
            '    _lastSql = strSQL
            '    rs = _currentDb.GetDataSet(_lastSql).Tables(0)
            'End Try
            'linqTable = rs.AsEnumerable()
            ''If usePLINQ Then linqTable.AsParallel()
            ''intGroups = From r In linqTable _
            ''                    Order By r("ColID") Ascending _
            ''                    Group By r!ColID Into g = Group Select g.Count()
            'intQueryResults = From r In linqTable _
            '               Order By r("ColID") _
            '               Select MatCnt = CInt(r("MatCnt"))
            'array_mcnt = intQueryResults.ToArray()

            ''---MBEG---
            'myArrayList = New ArrayList
            'myArrayList.Add(0)
            'For ctr = 0 To array_mcnt.Length - 1
            '    cumTotal = cumTotal + array_mcnt(ctr)
            '    myArrayList.Add(cumTotal)
            'Next
            'array_mbeg = DirectCast(myArrayList.ToArray(GetType(Integer)), Integer())

            ''---MIDX---dtMtx
            'strSQL = "SELECT ColID, RowID, (RowID - 1) AS MatIdx FROM [" & _miscParams!LPM_MATRIX_TABLE_NAME.ToString() & "] ORDER BY ColID, RowID "
            '_lastSql = strSQL
            'rs = _currentDb.GetDataSet(_lastSql).Tables(0)
            'linqTable = rs.AsEnumerable()
            ''If usePLINQ Then linqTable.AsParallel()
            'intQueryResults = From r In linqTable _
            '               Order By r("ColID"), r("RowID") _
            '               Select MatIdx = CInt(r("MatIdx"))
            'array_midx = intQueryResults.ToArray()

            ''---MVAL---dtMtx
            'strSQL = "SELECT ColID, RowID, COEF AS MatVal FROM [" & _miscParams!LPM_MATRIX_TABLE_NAME.ToString() & "] ORDER BY ColID, RowID "
            '_lastSql = strSQL
            'rs = _currentDb.GetDataSet(_lastSql).Tables(0)
            'linqTable = rs.AsEnumerable()
            ''If usePLINQ Then linqTable.AsParallel()
            'dblQueryResults = From r In linqTable _
            '               Order By r("ColID"), r("RowID") _
            '               Select MatVal = CDbl(r("MatVal"))
            'array_mval = dblQueryResults.ToArray()

            ''---SENSE---dtRow
            'strSQL = "SELECT ColID, RowID, COEF AS MatVal FROM [" & _miscParams!LPM_MATRIX_TABLE_NAME.ToString() & "] ORDER BY ColID, RowID "
            '_lastSql = strSQL
            'rs = _currentDb.GetDataSet(_lastSql).Tables(0)
            'linqTable = rs.AsEnumerable()
            ''If usePLINQ Then
            ''    'linqTable.AsParallel()
            ''    charQueryResults = From r In linqTable _
            ''                   Order By r("RowID") _
            ''                   Select SENSE = CChar(r("SENSE"))
            ''Else
            'charQueryResults = From r In linqTable _
            '               Order By r("RowID") _
            '               Select SENSE = CChar(r("SENSE"))
            ''End If
            'array_rtyp = charQueryResults.ToArray()
            'End If

            _progress = "Engine-MatGen-LoadSolverArrays-APM-After tsysMtx GC: " & My.Computer.Info.AvailablePhysicalMemory.ToString()
            Debug.Print("Engine-MatGen-LoadSolverArrays-APM-After tsysMtx GC: " & My.Computer.Info.AvailablePhysicalMemory.ToString())
            Console.WriteLine("Engine-MatGen-LoadSolverArrays-APM-After tsysMtx GC: " & My.Computer.Info.AvailablePhysicalMemory.ToString())
            EntLib.COPT.Log.Log(_workDir, " ", "Engine-MatGen-LoadSolverArrays-APM-After tsysMtx GC: " & My.Computer.Info.AvailablePhysicalMemory.ToString())
        End Sub

        ''' <summary>
        ''' Used by remote solver to differentiate a dataset.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks>
        ''' RKP/06-22-10/v2.3.133
        ''' </remarks>
        Public ReadOnly Property SessionGUID() As String
            Get
                Return _sessionGUID
            End Get
        End Property



        Public Function RemoteSolver_UploadData(ByVal tableType As String) As Integer
            Dim dc As DataColumn

            Select Case tableType.Trim.ToUpper
                Case "ROW"
                    dc = New DataColumn
                    dc.ColumnName = "DATASET_GUID"
                    dc.DataType = System.Type.GetType("System.String")
                    _dtRow.Columns.Add(dc)
                    dc = New DataColumn
                    dc.ColumnName = "DATASET_FLAG"
                    dc.DataType = System.Type.GetType("System.Byte")
                    _dtRow.Columns.Add(dc)
                    For ctr = 0 To dtMtx.Rows.Count - 1
                        _dtRow.Rows(ctr).Item("DATASET_GUID") = Me.SessionGUID
                        _dtRow.Rows(ctr).Item("DATASET_FLAG") = 0
                    Next
                    GenUtils.UpdateDB(GenUtils.GetAppSettings("RemoteSolverConnectionString"), dtRow, "SELECT * FROM [" & "dbo.TSYS_ROW_REMOTE" & "]")
                Case "COL"
                    dc = New DataColumn
                    dc.ColumnName = "DATASET_GUID"
                    dc.DataType = System.Type.GetType("System.String")
                    _dtCol.Columns.Add(dc)
                    dc = New DataColumn
                    dc.ColumnName = "DATASET_FLAG"
                    dc.DataType = System.Type.GetType("System.Byte")
                    _dtCol.Columns.Add(dc)
                    For ctr = 0 To dtMtx.Rows.Count - 1
                        _dtCol.Rows(ctr).Item("DATASET_GUID") = Me.SessionGUID
                        _dtCol.Rows(ctr).Item("DATASET_FLAG") = 0
                    Next
                    GenUtils.UpdateDB(GenUtils.GetAppSettings("RemoteSolverConnectionString"), dtCol, "SELECT * FROM [" & "dbo.TSYS_COL_REMOTE" & "]")
                Case "MTX"
                    dc = New DataColumn
                    dc.ColumnName = "DATASET_GUID"
                    dc.DataType = System.Type.GetType("System.String")
                    _dtMtx.Columns.Add(dc)
                    dc = New DataColumn
                    dc.ColumnName = "DATASET_FLAG"
                    dc.DataType = System.Type.GetType("System.Byte")
                    _dtMtx.Columns.Add(dc)
                    For ctr = 0 To dtMtx.Rows.Count - 1
                        _dtMtx.Rows(ctr).Item("DATASET_GUID") = Me.SessionGUID
                        _dtMtx.Rows(ctr).Item("DATASET_FLAG") = 0
                    Next
                    GenUtils.UpdateDB(GenUtils.GetAppSettings("RemoteSolverConnectionString"), dtMtx, "SELECT * FROM " & "[dbo].[TSYS_MTX_REMOTE]")
                Case Else
            End Select

            Return 0
        End Function

        ''' <summary>
        ''' Uses dtRow (built during PopElmTbls), dtCol (build during PopElmTbls) and dtMtx (build right after MatGen).
        ''' This function replicates qsysMTXwithCOLS in LINQ.
        ''' </summary>
        ''' <remarks>
        ''' RKP/06-25-10/v2.3.133
        ''' </remarks>
        Private Sub LoadSolverArraysMtx()
            Dim linqTable As System.Data.EnumerableRowCollection(Of System.Data.DataRow) = Nothing
            Dim dblQueryResults As System.Data.EnumerableRowCollection(Of Double) = Nothing
            Dim intQueryResults As System.Data.EnumerableRowCollection(Of Integer) = Nothing
            Dim charQueryResults As System.Data.EnumerableRowCollection(Of Char) = Nothing
            Dim strQueryResults As System.Data.EnumerableRowCollection(Of String) = Nothing
            Dim intGroups As System.Collections.Generic.IEnumerable(Of Integer) = Nothing
            Dim myArrayList As ArrayList = Nothing
            Dim cumTotal As Integer = 0

            'If GenUtils.IsSwitchAvailable(_switches, "/UseMinSysRes") Then
            's = "MatGen....." & GenUtils.FormatTime(startTime, My.Computer.Clock.TickCount)

            '"SELECT * FROM qsysMtx"
            '"SELECT * FROM qsysMTXwithCOLS WHERE COEF <> 0 ORDER BY ColID, RowID"
            'strSQL = "SELECT ColID, COUNT(*) AS MatCnt FROM [" & _miscParams!LPM_MATRIX_TABLE_NAME.ToString() & "] GROUP BY ColID ORDER BY ColID "

            'GenUtils.Message(GenUtils.MsgType.Information, "Engine-MatGen-LoadSolverArrays", "APM-Before: " & My.Computer.Info.AvailablePhysicalMemory.ToString)
            'GenUtils.Message(GenUtils.MsgType.Information, "Engine - MatGen - LoadSolverArrays", "APM-Before: " & My.Computer.Info.AvailableVirtualMemory.ToString)

            _progress = "Engine-MatGen-LoadSolverArrays-APM-Before: " & My.Computer.Info.AvailablePhysicalMemory.ToString()
            Debug.Print("Engine-MatGen-LoadSolverArrays-APM-Before: " & My.Computer.Info.AvailablePhysicalMemory.ToString())
            Console.WriteLine("Engine-MatGen-LoadSolverArrays-APM-Before: " & My.Computer.Info.AvailablePhysicalMemory.ToString())
            EntLib.COPT.Log.Log(_workDir, " ", "Engine-MatGen-LoadSolverArrays-APM-Before: " & My.Computer.Info.AvailablePhysicalMemory.ToString())

            If dtMtx Is Nothing Then
                Try
                    'RKP/01-28-10/v2.3.128
                    'If qsysMtx is not present, resort to backup option.
                    _dtMtx = _currentDb.GetDataTable("SELECT * FROM qsysMtx")            'SESSIONIZE
                    _srcMtx = "qsysMtx"
                Catch ex As Exception
                    _dtMtx = _currentDb.GetDataTable("SELECT * FROM qsysMTXwithCOLS WHERE COEF <> 0 ORDER BY ColID, RowID")            'SESSIONIZE
                    _srcMtx = "tsysMtx"
                End Try
            End If

            _progress = "Engine-MatGen-LoadSolverArrays-APM-After tsysMtx loaded: " & My.Computer.Info.AvailablePhysicalMemory.ToString()
            Debug.Print("Engine-MatGen-LoadSolverArrays-APM-After tsysMtx loaded: " & My.Computer.Info.AvailablePhysicalMemory.ToString())
            Console.WriteLine("Engine-MatGen-LoadSolverArrays-APM-After tsysMtx loaded: " & My.Computer.Info.AvailablePhysicalMemory.ToString())
            EntLib.COPT.Log.Log(_workDir, " ", "Engine-MatGen-LoadSolverArrays-APM-After tsysMtx loaded: " & My.Computer.Info.AvailablePhysicalMemory.ToString())

            '---MCNT---dtMtx
            If array_mcnt Is Nothing Then
                linqTable = dtMtx.AsEnumerable()
                'If usePLINQ Then linqTable.AsParallel()
                intGroups = From r In linqTable _
                                    Order By r("ColID") Ascending _
                                    Group By r!ColID Into g = Group Select g.Count()
                array_mcnt = intGroups.ToArray()
            End If

            '---MBEG---
            If array_mbeg Is Nothing Then
                myArrayList = New ArrayList
                myArrayList.Add(0)
                For ctr = 0 To array_mcnt.Length - 1
                    cumTotal = cumTotal + array_mcnt(ctr)
                    myArrayList.Add(cumTotal)
                Next
                array_mbeg = DirectCast(myArrayList.ToArray(GetType(Integer)), Integer())
            End If

            '---MIDX---dtMtx
            If array_midx Is Nothing Then
                linqTable = dtMtx.AsEnumerable()
                'If usePLINQ Then linqTable.AsParallel()
                intQueryResults = From r In linqTable _
                               Order By r("ColID"), r("RowID") _
                               Select IDX = CInt(r("RowID")) - 1
                array_midx = intQueryResults.ToArray()
            End If

            '---MVAL---dtMtx
            If array_mval Is Nothing Then
                linqTable = dtMtx.AsEnumerable()
                'If usePLINQ Then linqTable.AsParallel()
                dblQueryResults = From r In linqTable _
                               Order By r("ColID"), r("RowID") _
                               Select COEF = CDbl(r("COEF"))
                array_mval = dblQueryResults.ToArray()
            End If

            Me.SolutionNonZeros = dtMtx.Rows.Count

            'RKP/06-22-10/v2.3.133
            If GenUtils.IsSwitchAvailable(_switches, "/UseRemoteSolver") Then
                RemoteSolver_UploadData("MTX")
            End If

            _dtMtx = Nothing

            'http://blogs.msdn.com/b/ricom/archive/2004/11/29/271829.aspx
            'Call the Garbage Collector to clean up dtMtx, since it is no longer necessary and memory needs to be reclaimed immediately.
            'GC.Collect()
            'GC.WaitForPendingFinalizers()
            'GC.Collect()

            ''---MCNT---dtMtx
            'strSQL = "SELECT ColID, COUNT(*) AS MatCnt FROM [" & _miscParams!LPM_MATRIX_TABLE_NAME.ToString() & "] GROUP BY ColID ORDER BY ColID "
            '_lastSql = strSQL
            'rs = _currentDb.GetDataSet(_lastSql).Tables(0)
            'Try
            '    strSQL = "SELECT ColID, COUNT(*) AS MatCnt FROM [qsysMtx] GROUP BY ColID ORDER BY ColID "
            '    _lastSql = strSQL
            '    rs = _currentDb.GetDataSet(_lastSql).Tables(0)
            'Catch ex As Exception
            '    strSQL = "SELECT ColID, COUNT(*) AS MatCnt FROM (SELECT * FROM qsysMTXwithCOLS WHERE COEF <> 0 ORDER BY ColID, RowID) GROUP BY ColID ORDER BY ColID "
            '    _lastSql = strSQL
            '    rs = _currentDb.GetDataSet(_lastSql).Tables(0)
            'End Try
            'linqTable = rs.AsEnumerable()
            ''If usePLINQ Then linqTable.AsParallel()
            ''intGroups = From r In linqTable _
            ''                    Order By r("ColID") Ascending _
            ''                    Group By r!ColID Into g = Group Select g.Count()
            'intQueryResults = From r In linqTable _
            '               Order By r("ColID") _
            '               Select MatCnt = CInt(r("MatCnt"))
            'array_mcnt = intQueryResults.ToArray()

            ''---MBEG---
            'myArrayList = New ArrayList
            'myArrayList.Add(0)
            'For ctr = 0 To array_mcnt.Length - 1
            '    cumTotal = cumTotal + array_mcnt(ctr)
            '    myArrayList.Add(cumTotal)
            'Next
            'array_mbeg = DirectCast(myArrayList.ToArray(GetType(Integer)), Integer())

            ''---MIDX---dtMtx
            'strSQL = "SELECT ColID, RowID, (RowID - 1) AS MatIdx FROM [" & _miscParams!LPM_MATRIX_TABLE_NAME.ToString() & "] ORDER BY ColID, RowID "
            '_lastSql = strSQL
            'rs = _currentDb.GetDataSet(_lastSql).Tables(0)
            'linqTable = rs.AsEnumerable()
            ''If usePLINQ Then linqTable.AsParallel()
            'intQueryResults = From r In linqTable _
            '               Order By r("ColID"), r("RowID") _
            '               Select MatIdx = CInt(r("MatIdx"))
            'array_midx = intQueryResults.ToArray()

            ''---MVAL---dtMtx
            'strSQL = "SELECT ColID, RowID, COEF AS MatVal FROM [" & _miscParams!LPM_MATRIX_TABLE_NAME.ToString() & "] ORDER BY ColID, RowID "
            '_lastSql = strSQL
            'rs = _currentDb.GetDataSet(_lastSql).Tables(0)
            'linqTable = rs.AsEnumerable()
            ''If usePLINQ Then linqTable.AsParallel()
            'dblQueryResults = From r In linqTable _
            '               Order By r("ColID"), r("RowID") _
            '               Select MatVal = CDbl(r("MatVal"))
            'array_mval = dblQueryResults.ToArray()

            ''---SENSE---dtRow
            'strSQL = "SELECT ColID, RowID, COEF AS MatVal FROM [" & _miscParams!LPM_MATRIX_TABLE_NAME.ToString() & "] ORDER BY ColID, RowID "
            '_lastSql = strSQL
            'rs = _currentDb.GetDataSet(_lastSql).Tables(0)
            'linqTable = rs.AsEnumerable()
            ''If usePLINQ Then
            ''    'linqTable.AsParallel()
            ''    charQueryResults = From r In linqTable _
            ''                   Order By r("RowID") _
            ''                   Select SENSE = CChar(r("SENSE"))
            ''Else
            'charQueryResults = From r In linqTable _
            '               Order By r("RowID") _
            '               Select SENSE = CChar(r("SENSE"))
            ''End If
            'array_rtyp = charQueryResults.ToArray()
            'End If

            _progress = "Engine-MatGen-LoadSolverArrays-APM-After tsysMtx GC: " & My.Computer.Info.AvailablePhysicalMemory.ToString()
            Debug.Print("Engine-MatGen-LoadSolverArrays-APM-After tsysMtx GC: " & My.Computer.Info.AvailablePhysicalMemory.ToString())
            Console.WriteLine("Engine-MatGen-LoadSolverArrays-APM-After tsysMtx GC: " & My.Computer.Info.AvailablePhysicalMemory.ToString())
            EntLib.COPT.Log.Log(_workDir, " ", "Engine-MatGen-LoadSolverArrays-APM-After tsysMtx GC: " & My.Computer.Info.AvailablePhysicalMemory.ToString())
        End Sub

        Private Function LoadCol(ByRef dtCol As DataTable)
            Dim sql As String
            Dim tblCol As String = Me.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString() '"tsysCol"
            Dim tblMtx As String = Me.MiscParams.Item("LPM_MATRIX_TABLE_NAME").ToString() '"tsysMtx"
            Dim isValid As Boolean = False
            Dim qsysMtxCol As String = "qsysMtxCol"
            Dim ret As Long
            Dim linkedDB As Boolean = False

            isValid = False
            If Not dtCol Is Nothing Then
                Try
                    If dtCol.Rows.Count > 0 Then
                        isValid = True
                    End If
                Catch ex As Exception

                End Try
            End If

            If Not isValid Then

                'UPDATE tsysCOL, tsysDefCol SET tsysCOL.UP = 1
                'WHERE ([tsysCOL].[UP]<>1) And ([tsysDefCol].[BNDBinary]=True) And ([tsysCOL].[COL] Like 'C'+[tsysDefCol].[ColTypePrefix]+'*')
                If CurrentDb.IsSQLExpress Then
                    'Sample UPDATE SQL
                    'UPDATE [tsysCOL] SET [tsysCOL].[IsValid] = 1 FROM [tsysCOL] INNER JOIN [qsysMTXwithCOLS] ON [tsysCOL].[ColID] = [qsysMTXwithCOLS].[ColID] WHERE [qsysMTXwithCOLS].[COEF] <> 0
                    sql = "UPDATE [tsysCOL] SET [tsysCOL].[UP] = 1 FROM [tsysCOL], [tsysDefCol] WHERE ([tsysCOL].[UP]<>1) And ([tsysDefCol].[BNDBinary]=1) And ([tsysCOL].[COL] Like 'C'+[tsysDefCol].[ColTypePrefix]+'%')"
                Else
                    sql = "UPDATE tsysCOL, tsysDefCol SET tsysCOL.UP = 1 WHERE ([tsysCOL].[UP]<>1) And ([tsysDefCol].[BNDBinary]=True) And ([tsysCOL].[COL] Like 'C'+[tsysDefCol].[ColTypePrefix]+'*')"
                End If

                Try
                    ret = _currentDb.ExecuteNonQuery(sql)
                Catch ex As Exception
                    GenUtils.Message(GenUtils.MsgType.Critical, "C-OPT Engine - LoadCol", ex.Message)
                End Try


                If GenUtils.IsSwitchAvailable(_switches, "/UseMinSysRes") Then
                    If GenUtils.IsSwitchAvailable(_switches, "/UseMSAccessSyntax") Then
                        tblCol = "[" & Me.MiscParams.Item("LINKED_COL_DB_PATH").ToString() & "]." & "[" & Me.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString() & "]"
                        tblMtx = "[" & Me.MiscParams.Item("LINKED_MTX_DB_PATH").ToString() & "]." & "[" & Me.MiscParams.Item("LPM_MATRIX_TABLE_NAME").ToString() & "]"
                        _srcMtx = "tsysMtx"
                    End If
                Else
                    'If GenUtils.IsSwitchAvailable(_switches, "/UseSysTables") Then
                    '    sql = "SELECT * FROM qsysMTXwithCOLS WHERE COEF <> 0 ORDER BY ColID, RowID"
                    '    _srcMtx = "tsysMtx"
                    'ElseIf GenUtils.IsSwitchAvailable(_switches, "/UseSysQueries") Then
                    '    sql = "SELECT * FROM qsysMtx"
                    '    _srcMtx = "qsysMtx"
                    'Else
                    '    sql = "SELECT * FROM qsysMTXwithCOLS WHERE COEF <> 0 ORDER BY ColID, RowID"
                    '    _srcMtx = "tsysMtx"
                    'End If
                End If

                'sql = "SELECT DISTINCT "
                'sql = sql & "tsysMTX.ColID, "
                ''sql = sql & "tsysMTX.RowID, "
                'sql = sql & "tsysMTX.COL, "
                'sql = sql & "'' AS [DESC], "
                ''sql = sql & "tsysMTX.ROW, "
                ''sql = sql & "tsysMTX.COEF, "
                'sql = sql & "tsysCOL.OBJ, "
                'sql = sql & "tsysCOL.LO, "
                'sql = sql & "tsysCOL.UP, "
                'sql = sql & "tsysCOL.FREE, "
                'sql = sql & "tsysCOL.INTGR, "
                'sql = sql & "tsysCOL.BINRY, "
                'sql = sql & "tsysCOL.SOSTYPE, "
                'sql = sql & "tsysCOL.SOSMARKER, "
                'sql = sql & "tsysCOL.ACTIVITY, "
                'sql = sql & "tsysCOL.DJ, "
                'sql = sql & "tsysCOL.STATUS "
                'sql = sql & "FROM "
                'sql = sql & tblCol & " "
                'sql = sql & "INNER JOIN "
                'sql = sql & tblMtx & " "
                'sql = sql & "ON "
                'sql = sql & "tsysCOL.ColID = tsysMTX.ColID "
                'sql = sql & "WHERE "
                'sql = sql & "tsysMTX.COEF <> 0 "
                'sql = sql & "ORDER BY tsysMTX.ColID "

                'sql = "SELECT * FROM tsysCol WHERE ColID IN (SELECT DISTINCT tsysMTX.ColID FROM tsysMTX INNER JOIN tsysCol ON tsysMTX.ColID = tsysCol.ColID WHERE tsysMTX.COEF <> 0) ORDER BY ColID"

                If CurrentDb.IsSQLExpress Then
                    sql = "SELECT * FROM [" & Me.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString() & "] WHERE ISVALID = 1 ORDER BY ColID"
                Else
                    If GenUtils.IsSwitchAvailable(_switches, "/UseSQLServerSyntax") Then
                        sql = "SELECT * FROM [" & Me.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString() & "] WHERE ISVALID = 1 ORDER BY ColID"
                    Else
                        sql = "SELECT * FROM [" & Me.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString() & "] WHERE ISVALID = TRUE ORDER BY ColID"
                    End If

                End If


                Try
                    'RKP/01-28-10/v2.3.128
                    'If qsysMtx is not present, resort to backup option.
                    'dtCol = _currentDb.GetDataTable(sql)            'SESSIONIZE

                    If GenUtils.IsSwitchAvailable(_switches, "/UseSQLServerSyntax") Then
                        linkedDB = True
                    End If
                    If GenUtils.IsSwitchAvailable(_switches, "/UseMSAccessSyntax") Then
                        linkedDB = True
                    End If
                    dtCol = _currentDb.GetDataTable(sql, linkedDB, _switches)
                Catch ex As Exception
                    '_dtMtx = _currentDb.GetDataTable("SELECT * FROM qsysMTXwithCOLS WHERE COEF <> 0 ORDER BY ColID, RowID")            'SESSIONIZE
                    '_srcMtx = "tsysMtx"

                    Try
                        'sql = "SELECT ColID, COL, '' AS DESC, OBJ, LO, UP, FREE, INTGR, BINRY, SOSTYPE, SOSMARKER, ACTIVITY, DJ, STATUS FROM qsysMTXwithCOLS WHERE COEF <> 0 ORDER BY ColID, RowID"
                        sql = "SELECT * FROM [" & Me.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString() & "] WHERE ColID IN (SELECT DISTINCT [" & Me.MiscParams.Item("LPM_MATRIX_TABLE_NAME").ToString() & "].ColID FROM [" & Me.MiscParams.Item("LPM_MATRIX_TABLE_NAME").ToString() & "] INNER JOIN [" & Me.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString() & "] ON [" & Me.MiscParams.Item("LPM_MATRIX_TABLE_NAME").ToString() & "].ColID = [" & Me.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString() & "].ColID WHERE [" & Me.MiscParams.Item("LPM_MATRIX_TABLE_NAME").ToString() & "].COEF <> 0) ORDER BY ColID"
                        dtCol = _currentDb.GetDataTable(sql)
                        EntLib.COPT.Log.Log(_workDir, " ", "Engine-LoadCol- dtCol not loaded using primary query: " & ex.Message)
                        EntLib.COPT.Log.Log(_workDir, " ", "Engine-LoadCol- dtCol loaded using fallback query (qsysMTXwithCOLS)")
                    Catch ex2 As Exception
                        GenUtils.Message(GenUtils.MsgType.Critical, "C-OPT Engine-LoadSolverArraysUsingMtxCol", "MtxCol not loaded using fallback query (qsysMTXwithCOLS): " & ex2.Message)
                    End Try
                End Try
            End If

            Return 0
        End Function

        Private Function LoadMtx(ByRef dtMtx As DataTable)
            Dim sql As String
            Dim tblCol As String = Me.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString() '"tsysCol"
            Dim tblMtx As String = Me.MiscParams.Item("LPM_MATRIX_TABLE_NAME").ToString() '"tsysMtx"
            Dim isValid As Boolean = False
            Dim qsysMtxCol As String = "qsysMtxCol"

            isValid = False
            If Not dtMtx Is Nothing Then
                Try
                    If dtMtx.Rows.Count > 0 Then
                        isValid = True
                    End If
                Catch ex As Exception

                End Try
            End If

            If Not isValid Then
                If GenUtils.IsSwitchAvailable(_switches, "/UseMinSysRes") Then
                    If GenUtils.IsSwitchAvailable(_switches, "/UseMSAccessSyntax") Then
                        tblCol = "[" & Me.MiscParams.Item("LINKED_COL_DB_PATH").ToString() & "]." & "[" & Me.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString() & "]"
                        tblMtx = "[" & Me.MiscParams.Item("LINKED_MTX_DB_PATH").ToString() & "]." & "[" & Me.MiscParams.Item("LPM_MATRIX_TABLE_NAME").ToString() & "]"
                        _srcMtx = "tsysMtx"
                    End If
                Else
                    'If GenUtils.IsSwitchAvailable(_switches, "/UseSysTables") Then
                    '    sql = "SELECT * FROM qsysMTXwithCOLS WHERE COEF <> 0 ORDER BY ColID, RowID"
                    '    _srcMtx = "tsysMtx"
                    'ElseIf GenUtils.IsSwitchAvailable(_switches, "/UseSysQueries") Then
                    '    sql = "SELECT * FROM qsysMtx"
                    '    _srcMtx = "qsysMtx"
                    'Else
                    '    sql = "SELECT * FROM qsysMTXwithCOLS WHERE COEF <> 0 ORDER BY ColID, RowID"
                    '    _srcMtx = "tsysMtx"
                    'End If
                End If

                'qsysMtxCol
                sql = "SELECT DISTINCT "
                sql = sql & "tsysMTX.ColID, "
                sql = sql & "tsysMTX.RowID, "
                sql = sql & "tsysMTX.COL, "
                sql = sql & "tsysMTX.ROW, "
                sql = sql & "tsysMTX.COEF, "
                sql = sql & "tsysCOL.OBJ, "
                sql = sql & "tsysCOL.LO, "
                sql = sql & "tsysCOL.UP, "
                sql = sql & "tsysCOL.FREE, "
                sql = sql & "tsysCOL.INTGR, "
                sql = sql & "tsysCOL.BINRY, "
                sql = sql & "tsysCOL.SOSTYPE, "
                sql = sql & "tsysCOL.SOSMARKER, "
                sql = sql & "tsysCOL.ACTIVITY, "
                sql = sql & "tsysCOL.DJ, "
                sql = sql & "tsysCOL.STATUS "
                sql = sql & "FROM "
                sql = sql & tblCol & " "
                sql = sql & "INNER JOIN "
                sql = sql & tblMtx & " "
                sql = sql & "ON "
                sql = sql & "tsysCOL.ColID = tsysMTX.ColID "
                sql = sql & "WHERE "
                sql = sql & "tsysMTX.COEF <> 0 "
                sql = sql & "ORDER BY tsysMTX.ColID, tsysMTX.RowID "

                'RKP/08-03-10/v2.3.135
                sql = "SELECT * FROM qsysMTXwithCOLS WHERE COEF <> 0 ORDER BY ColID, RowID"

                Try
                    'RKP/01-28-10/v2.3.128
                    'If qsysMtx is not present, resort to backup option.
                    dtMtx = _currentDb.GetDataTable(sql)            'SESSIONIZE
                Catch ex As Exception
                    '_dtMtx = _currentDb.GetDataTable("SELECT * FROM qsysMTXwithCOLS WHERE COEF <> 0 ORDER BY ColID, RowID")            'SESSIONIZE
                    '_srcMtx = "tsysMtx"

                    Try
                        sql = "SELECT * FROM qsysMTXwithCOLS WHERE COEF <> 0 ORDER BY ColID, RowID"
                        dtMtx = _currentDb.GetDataTable(sql)
                        EntLib.COPT.Log.Log(_workDir, " ", "Engine-LoadMtx- dtMtx not loaded using primary query: " & ex.Message)
                        EntLib.COPT.Log.Log(_workDir, " ", "Engine-LoadMtx- dtMtx loaded using fallback query (qsysMTXwithCOLS)")
                    Catch ex2 As Exception
                        GenUtils.Message(GenUtils.MsgType.Critical, "C-OPT Engine-LoadMtx", "dtMtx not loaded using fallback query (qsysMTXwithCOLS): " & ex2.Message)
                    End Try
                End Try
            End If

            Return 0
        End Function

        Private Sub LoadSolverArraysInSequence()
            Dim linqTable As System.Data.EnumerableRowCollection(Of System.Data.DataRow) = Nothing
            Dim linqTable3 As System.Collections.Generic.IEnumerable(Of System.Data.DataRow) = Nothing
            Dim dblQueryResults As System.Data.EnumerableRowCollection(Of Double) = Nothing
            Dim intQueryResults As System.Data.EnumerableRowCollection(Of Integer) = Nothing
            Dim charQueryResults As System.Data.EnumerableRowCollection(Of Char) = Nothing
            Dim strQueryResults As System.Data.EnumerableRowCollection(Of String) = Nothing

            Dim dblQueryResultsGeneric As System.Collections.Generic.IEnumerable(Of Double) = Nothing
            Dim intQueryResultsGeneric As System.Collections.Generic.IEnumerable(Of Integer) = Nothing
            Dim charQueryResultsGeneric As System.Collections.Generic.IEnumerable(Of Char) = Nothing
            Dim strQueryResultsGeneric As System.Collections.Generic.IEnumerable(Of String) = Nothing

            Dim drQueryResults As System.Data.EnumerableRowCollection(Of System.Data.DataRow) = Nothing

            Dim intGroups As System.Collections.Generic.IEnumerable(Of Integer) = Nothing
            Dim myArrayList As ArrayList = Nothing
            Dim cumTotal As Integer = 0
            Dim isValid As Boolean = False
            Dim proceed As Boolean = False
            Dim startTime As Integer = My.Computer.Clock.TickCount
            Dim startTime2 As Integer = My.Computer.Clock.TickCount
            Dim totalTime As String = ""
            'Dim isMIP As Boolean = False

            'If GenUtils.IsSwitchAvailable(_switches, "/UseMinSysRes") Then
            's = "MatGen....." & GenUtils.FormatTime(startTime, My.Computer.Clock.TickCount)

            '"SELECT * FROM qsysMtx"
            '"SELECT * FROM qsysMTXwithCOLS WHERE COEF <> 0 ORDER BY ColID, RowID"
            'strSQL = "SELECT ColID, COUNT(*) AS MatCnt FROM [" & _miscParams!LPM_MATRIX_TABLE_NAME.ToString() & "] GROUP BY ColID ORDER BY ColID "

            'GenUtils.Message(GenUtils.MsgType.Information, "Engine-MatGen-LoadSolverArrays", "APM-Before: " & My.Computer.Info.AvailablePhysicalMemory.ToString)
            'GenUtils.Message(GenUtils.MsgType.Information, "Engine - MatGen - LoadSolverArrays", "APM-Before: " & My.Computer.Info.AvailableVirtualMemory.ToString)

            '_progress = "Engine-LoadSolverArrays-APM-Before: " & My.Computer.Info.AvailablePhysicalMemory.ToString()
            'Debug.Print("Engine-LoadSolverArrays-APM-Before: " & My.Computer.Info.AvailablePhysicalMemory.ToString())
            'Console.WriteLine("Engine-MatGen-LoadSolverArrays-APM-Before: " & My.Computer.Info.AvailablePhysicalMemory.ToString())
            'EntLib.COPT.Log.Log(_workDir, " ", "Engine-LoadSolverArrays-APM-Before: " & My.Computer.Info.AvailablePhysicalMemory.ToString())

            isValid = False
            If Not dtRow Is Nothing Then
                Try
                    If dtRow.Rows.Count > 0 Then
                        isValid = True
                    End If
                Catch ex As Exception

                End Try
            End If
            If Not isValid Then
                startTime2 = My.Computer.Clock.TickCount
                'GenUtils.FormatTime(startTime, My.Computer.Clock.TickCount)
                _progress = "Loading rows from database..."

                LoadRow(_dtRow)

                totalTime = GenUtils.FormatTime(startTime2, My.Computer.Clock.TickCount)
                _progress = "Loading rows from database took: " & totalTime
                Debug.Print(_progress)
                Console.WriteLine(_progress)
                EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - " & _progress)
            End If

            linqTable = dtRow.AsEnumerable()

            Me.SolutionRows = dtRow.Rows.Count

            '---SENSE---dtRow
            If array_rtyp Is Nothing Then
                charQueryResults = From r In linqTable _
                               Order By r("RowID") _
                               Select SENSE = CChar(r("SENSE"))
                array_rtyp = charQueryResults.ToArray()
            End If

            '---RHS---dtRow
            If array_dobj Is Nothing Then
                'linqTable = dtRow.AsEnumerable()
                'If usePLINQ Then linqTableRow.AsParallel()
                dblQueryResults = From r In linqTable _
                               Order By r("RowID") _
                               Select RHS = CDbl(r("RHS"))
                array_drhs = dblQueryResults.ToArray()
            End If

            '---ROW---dtRow
            If array_rowNames Is Nothing Then
                'linqTable = dtRow.AsEnumerable()
                'If usePLINQ Then linqTableRow.AsParallel()
                '/SubNameWithID
                If GenUtils.IsSwitchAvailable(_switches, "/SubNameWithID") Then
                    strQueryResults = From r In linqTable Order By r("RowID") Select CStr(r("RowID"))
                Else
                    strQueryResults = From r In linqTable Order By r("RowID") Select CStr(r("ROW"))
                End If
                array_rowNames = strQueryResults.ToArray()
            End If

            EntLib.COPT.Log.Log(_workDir, " ", "Engine-LoadSolverArrays-APM-Before tsysMtx loaded: " & My.Computer.Info.AvailablePhysicalMemory.ToString())
            startTime2 = My.Computer.Clock.TickCount
            _progress = "Loading non-zeros from database..."

            LoadMtx(_dtMtx)

            totalTime = GenUtils.FormatTime(startTime2, My.Computer.Clock.TickCount)
            _progress = "Loading non-zeros from database took: " & totalTime
            Debug.Print(_progress)
            Console.WriteLine(_progress)
            EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - " & _progress)

            If dtMtx Is Nothing Then
                Try
                    'RKP/01-28-10/v2.3.128
                    'If qsysMtx is not present, resort to backup option.
                    _dtMtx = _currentDb.GetDataTable("SELECT * FROM qsysMtx")            'SESSIONIZE
                    _srcMtx = "qsysMtx"
                Catch ex As Exception
                    _dtMtx = _currentDb.GetDataTable("SELECT * FROM qsysMTXwithCOLS WHERE COEF <> 0 ORDER BY ColID, RowID")            'SESSIONIZE
                    _srcMtx = "tsysMtx"
                End Try
            End If

            '_progress = "Engine-LoadSolverArrays-APM-After tsysMtx loaded: " & My.Computer.Info.AvailablePhysicalMemory.ToString()
            'Debug.Print("Engine-LoadSolverArrays-APM-After tsysMtx loaded: " & My.Computer.Info.AvailablePhysicalMemory.ToString())
            'Console.WriteLine("Engine-LoadSolverArrays-APM-After tsysMtx loaded: " & My.Computer.Info.AvailablePhysicalMemory.ToString())
            EntLib.COPT.Log.Log(_workDir, " ", "Engine-LoadSolverArrays-APM-After tsysMtx loaded: " & My.Computer.Info.AvailablePhysicalMemory.ToString())

            linqTable = dtMtx.AsEnumerable()
            Me.SolutionNonZeros = dtMtx.Rows.Count

            '---MCNT---dtMtx
            If array_mcnt Is Nothing Then
                'If usePLINQ Then linqTable.AsParallel()
                intQueryResultsGeneric = From r In linqTable _
                                    Order By r("ColID") Ascending _
                                    Group By r!ColID Into g = Group Select g.Count()
                Try
                    array_mcnt = intQueryResultsGeneric.ToArray()
                Catch ex As Exception
                    Debug.Print(ex.Message)
                End Try

            End If

            '---MBEG---
            If array_mbeg Is Nothing Then
                myArrayList = New ArrayList
                myArrayList.Add(0)
                For ctr = 0 To array_mcnt.Length - 1
                    cumTotal = cumTotal + array_mcnt(ctr)
                    myArrayList.Add(cumTotal)
                Next
                array_mbeg = DirectCast(myArrayList.ToArray(GetType(Integer)), Integer())
            End If

            '---MIDX---dtMtx
            If array_midx Is Nothing Then
                'If usePLINQ Then linqTable.AsParallel()
                intQueryResults = From r In linqTable _
                               Order By r("ColID"), r("RowID") _
                               Select IDX = CInt(r("RowID")) - 1
                array_midx = intQueryResults.ToArray()
            End If

            '---MVAL---dtMtx
            If array_mval Is Nothing Then
                'If usePLINQ Then linqTable.AsParallel()
                dblQueryResults = From r In linqTable _
                               Order By r("ColID"), r("RowID") _
                               Select COEF = CDbl(r("COEF"))
                array_mval = dblQueryResults.ToArray()
            End If

            'RKP/06-22-10/v2.3.133
            If GenUtils.IsSwitchAvailable(_switches, "/UseRemoteSolver") Then
                RemoteSolver_UploadData("MTX")
            End If

            _dtMtx = Nothing

            'http://blogs.msdn.com/b/ricom/archive/2004/11/29/271829.aspx
            'Call the Garbage Collector to clean up dtMtx, since it is no longer necessary and memory needs to be reclaimed immediately.
            'GC.Collect()
            'GC.WaitForPendingFinalizers()
            'GC.Collect()

            startTime2 = My.Computer.Clock.TickCount
            _progress = "Loading columns from database..."

            LoadCol(_dtCol)

            totalTime = GenUtils.FormatTime(startTime2, My.Computer.Clock.TickCount)
            _progress = "Loading columns from database took: " & totalTime
            Debug.Print(_progress)
            Console.WriteLine(_progress)
            EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - " & _progress)

            linqTable = dtCol.AsEnumerable()
            Me.SolutionColumns = dtCol.Rows.Count
            '---OBJ---dtCol
            If array_dobj Is Nothing Then
                'linqTable = dtCol.AsEnumerable()
                'If usePLINQ Then linqTableCol.AsParallel()
                dblQueryResults = From r In linqTable _
                               Order By r("ColID") _
                               Select OBJ = CDbl(r("OBJ"))
                array_dobj = dblQueryResults.ToArray()
            End If

            '---LO---dtCol
            If array_dclo Is Nothing Then
                'If usePLINQ Then linqTableCol.AsParallel()
                dblQueryResults = From r In linqTable _
                               Order By r("ColID") _
                               Select LO = CDbl(r("LO"))
                array_dclo = dblQueryResults.ToArray()
            End If

            'ub(0) = 40.0 : ub(1) = CPX_INFBOUND : ub(2) = CPX_INFBOUND
            '---UP---dtCol
            If array_dcup Is Nothing Then
                'linqTable = dtCol.AsEnumerable()
                'If usePLINQ Then linqTableCol.AsParallel()
                dblQueryResults = From r In linqTable _
                               Order By r("ColID") _
                               Select UP = CDbl(r("UP"))
                array_dcup = dblQueryResults.ToArray()
            End If

            '---COL---dtCol
            If array_colNames Is Nothing Then
                'linqTable = dtCol.AsEnumerable()
                'If usePLINQ Then linqTableCol.AsParallel()
                '/SubNameWithID
                If GenUtils.IsSwitchAvailable(_switches, "/SubNameWithID") Then
                    strQueryResults = From r In linqTable Order By r("ColID") Select CStr(r("ColID"))
                Else
                    strQueryResults = From r In linqTable Order By r("ColID") Select CStr(r("COL"))
                End If
                array_colNames = strQueryResults.ToArray()
            End If

            startTime2 = My.Computer.Clock.TickCount
            _progress = "Loading [column type] array..."

            '---CTYP---dtCol
            If array_ctyp Is Nothing Then
                Me.ProblemType = COPT.Engine.problemRunType.problemTypeContinuous
                proceed = False
                _isMIP = False
                'linqTable = dtCol.AsEnumerable()
                'If usePLINQ Then linqTableCol.AsParallel()
                intGroups = From r In linqTable _
                                    Where r!BINRY = True _
                                    Order By r("ColID") Ascending _
                                    Group By r!ColID Into g = Group Select g.Count()
                Dim tmpArr = intGroups.ToArray()
                If intGroups.ToArray().Count > 0 Then
                    proceed = True
                    Me.ProblemType = COPT.Engine.problemRunType.problemTypeBinary
                End If
                'Else
                intGroups = From r In linqTable _
                                    Where r!INTGR = True _
                                    Order By r("ColID") Ascending _
                                    Group By r!ColID Into g = Group Select g.Count()
                If intGroups.ToArray().Count > 0 Then
                    proceed = True
                    '_engine.ProblemType = "I"
                    If Me.ProblemType = COPT.Engine.problemRunType.problemTypeContinuous Then
                        Me.ProblemType = COPT.Engine.problemRunType.problemTypeInteger
                    Else
                        Me.ProblemType = Me.ProblemType + COPT.Engine.problemRunType.problemTypeInteger
                    End If
                End If
                'End If
                If proceed Then
                    _isMIP = True
                    'linqTable = dtCol.AsEnumerable()
                    'If usePLINQ Then linqTableCol.AsParallel()
                    'If usePLINQ Then linqTable.AsParallel()
                    charQueryResults = From r In linqTable _
                                   Order By r("ColID") Ascending _
                                   Select INTGR = _
                                    CChar( _
                                        IIf( _
                                            r("BINRY") = True, _
                                            "B", _
                                            IIf( _
                                                r("INTGR") = True, _
                                                "I", _
                                                "C" _
                                            ) _
                                        ) _
                                    )
                    array_ctyp = charQueryResults.ToArray()
                    'dtCol(0)("UP").ToString()

                    'startTime2 = My.Computer.Clock.TickCount
                    '_progress = "Loading columns from database..."

                    'RKP/08-05-10/v2.3.136
                    'This block of code takes ~10 min. to finish, for HIPNET (~60,000 rows in tsysCOL).
                    'A, faster, workaround for this has been implemented in "LoadCol" via SQL UPDATE statement.
                    'For i = 0 To dtCol.Rows.Count - 1
                    '    If dtCol(i)("BINRY") Then
                    '        dtCol(i)("UP") = 1
                    '    End If
                    'Next

                    'linqTable = dtCol.AsEnumerable()
                    'If usePLINQ Then linqTableCol.AsParallel()
                    dblQueryResults = From r In linqTable _
                                   Order By r("ColID") _
                                   Select LO = CDbl(0.0)
                    array_initValues = dblQueryResults.ToArray()
                End If
            End If
            '---CTYP---

            totalTime = GenUtils.FormatTime(startTime2, My.Computer.Clock.TickCount)
            _progress = "Loading [column type] array took: " & totalTime
            Debug.Print(_progress)
            Console.WriteLine(_progress)
            EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - " & _progress)

            ''---MCNT---dtMtx
            'strSQL = "SELECT ColID, COUNT(*) AS MatCnt FROM [" & _miscParams!LPM_MATRIX_TABLE_NAME.ToString() & "] GROUP BY ColID ORDER BY ColID "
            '_lastSql = strSQL
            'rs = _currentDb.GetDataSet(_lastSql).Tables(0)
            'Try
            '    strSQL = "SELECT ColID, COUNT(*) AS MatCnt FROM [qsysMtx] GROUP BY ColID ORDER BY ColID "
            '    _lastSql = strSQL
            '    rs = _currentDb.GetDataSet(_lastSql).Tables(0)
            'Catch ex As Exception
            '    strSQL = "SELECT ColID, COUNT(*) AS MatCnt FROM (SELECT * FROM qsysMTXwithCOLS WHERE COEF <> 0 ORDER BY ColID, RowID) GROUP BY ColID ORDER BY ColID "
            '    _lastSql = strSQL
            '    rs = _currentDb.GetDataSet(_lastSql).Tables(0)
            'End Try
            'linqTable = rs.AsEnumerable()
            ''If usePLINQ Then linqTable.AsParallel()
            ''intGroups = From r In linqTable _
            ''                    Order By r("ColID") Ascending _
            ''                    Group By r!ColID Into g = Group Select g.Count()
            'intQueryResults = From r In linqTable _
            '               Order By r("ColID") _
            '               Select MatCnt = CInt(r("MatCnt"))
            'array_mcnt = intQueryResults.ToArray()

            ''---MBEG---
            'myArrayList = New ArrayList
            'myArrayList.Add(0)
            'For ctr = 0 To array_mcnt.Length - 1
            '    cumTotal = cumTotal + array_mcnt(ctr)
            '    myArrayList.Add(cumTotal)
            'Next
            'array_mbeg = DirectCast(myArrayList.ToArray(GetType(Integer)), Integer())

            ''---MIDX---dtMtx
            'strSQL = "SELECT ColID, RowID, (RowID - 1) AS MatIdx FROM [" & _miscParams!LPM_MATRIX_TABLE_NAME.ToString() & "] ORDER BY ColID, RowID "
            '_lastSql = strSQL
            'rs = _currentDb.GetDataSet(_lastSql).Tables(0)
            'linqTable = rs.AsEnumerable()
            ''If usePLINQ Then linqTable.AsParallel()
            'intQueryResults = From r In linqTable _
            '               Order By r("ColID"), r("RowID") _
            '               Select MatIdx = CInt(r("MatIdx"))
            'array_midx = intQueryResults.ToArray()

            ''---MVAL---dtMtx
            'strSQL = "SELECT ColID, RowID, COEF AS MatVal FROM [" & _miscParams!LPM_MATRIX_TABLE_NAME.ToString() & "] ORDER BY ColID, RowID "
            '_lastSql = strSQL
            'rs = _currentDb.GetDataSet(_lastSql).Tables(0)
            'linqTable = rs.AsEnumerable()
            ''If usePLINQ Then linqTable.AsParallel()
            'dblQueryResults = From r In linqTable _
            '               Order By r("ColID"), r("RowID") _
            '               Select MatVal = CDbl(r("MatVal"))
            'array_mval = dblQueryResults.ToArray()

            ''---SENSE---dtRow
            'strSQL = "SELECT ColID, RowID, COEF AS MatVal FROM [" & _miscParams!LPM_MATRIX_TABLE_NAME.ToString() & "] ORDER BY ColID, RowID "
            '_lastSql = strSQL
            'rs = _currentDb.GetDataSet(_lastSql).Tables(0)
            'linqTable = rs.AsEnumerable()
            ''If usePLINQ Then
            ''    'linqTable.AsParallel()
            ''    charQueryResults = From r In linqTable _
            ''                   Order By r("RowID") _
            ''                   Select SENSE = CChar(r("SENSE"))
            ''Else
            'charQueryResults = From r In linqTable _
            '               Order By r("RowID") _
            '               Select SENSE = CChar(r("SENSE"))
            ''End If
            'array_rtyp = charQueryResults.ToArray()
            'End If

            totalTime = GenUtils.FormatTime(startTime, My.Computer.Clock.TickCount)
            '_progress = "Engine-LoadSolverArrays-APM-After tsysMtx GC: " & My.Computer.Info.AvailablePhysicalMemory.ToString()
            'Debug.Print("Engine-LoadSolverArrays-APM-After tsysMtx GC: " & My.Computer.Info.AvailablePhysicalMemory.ToString())
            _progress = "Time to load solver arrays: " & totalTime
            Debug.Print("Time to load solver arrays: " & totalTime)
            'Console.WriteLine("Engine-LoadSolverArrays-APM-After tsysMtx GC: " & My.Computer.Info.AvailablePhysicalMemory.ToString())
            Console.WriteLine("Time to load solver arrays: " & totalTime)
            EntLib.COPT.Log.Log(_workDir, " ", "Engine-LoadSolverArrays-APM-After GC: " & My.Computer.Info.AvailablePhysicalMemory.ToString())
            EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - Time to load solver arrays: ", Space(9) & totalTime)
        End Sub

        Private Function LoadRow(ByRef dtRow As DataTable)
            Dim sql As String
            Dim tblRow As String = Me.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString() '"tsysRow"
            Dim isValid As Boolean = False
            Dim linkedDB As Boolean = False

            isValid = False
            If Not dtRow Is Nothing Then
                Try
                    If dtRow.Rows.Count > 0 Then
                        isValid = True
                    End If
                Catch ex As Exception

                End Try
            End If

            If Not isValid Then
                If GenUtils.IsSwitchAvailable(_switches, "/UseMinSysRes") Then
                    If GenUtils.IsSwitchAvailable(_switches, "/UseMSAccessSyntax") Then
                        tblRow = "[" & Me.MiscParams.Item("LINKED_ROW_DB_PATH").ToString() & "]." & "[" & Me.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString() & "]"
                        '???TODO???
                        _srcMtx = "tsysRow"
                    End If
                Else
                    'If GenUtils.IsSwitchAvailable(_switches, "/UseSysTables") Then
                    '    sql = "SELECT * FROM qsysMTXwithCOLS WHERE COEF <> 0 ORDER BY ColID, RowID"
                    '    _srcMtx = "tsysMtx"
                    'ElseIf GenUtils.IsSwitchAvailable(_switches, "/UseSysQueries") Then
                    '    sql = "SELECT * FROM qsysMtx"
                    '    _srcMtx = "qsysMtx"
                    'Else
                    '    sql = "SELECT * FROM qsysMTXwithCOLS WHERE COEF <> 0 ORDER BY ColID, RowID"
                    '    _srcMtx = "tsysMtx"
                    'End If
                End If

                sql = "SELECT * FROM " & tblRow & " ORDER BY RowID"

                Try
                    'RKP/01-28-10/v2.3.128
                    'If qsysMtx is not present, resort to backup option.
                    'dtRow = _currentDb.GetDataTable(sql)            'SESSIONIZE

                    If GenUtils.IsSwitchAvailable(_switches, "/UseSQLServerSyntax") Then
                        linkedDB = True
                    End If
                    If GenUtils.IsSwitchAvailable(_switches, "/UseMSAccessSyntax") Then
                        linkedDB = True
                    End If
                    dtRow = _currentDb.GetDataTable(sql, linkedDB, _switches)
                Catch ex As Exception
                    EntLib.COPT.Log.Log(_workDir, " ", "Engine-LoadRow- dtRow not loaded using primary query: " & ex.Message)
                End Try
            End If



            Return 0
        End Function

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' RKP/08-03-10/v2.3.135
        ''' </remarks>
        Private Function FilterEmptyColumnsRows() As Integer
            Dim sql As String

            If CurrentDb.IsSQLExpress Then
                sql = "UPDATE [tsysCOL] SET [ISVALID] = 0"
            Else
                sql = "UPDATE [tsysCOL] SET [ISVALID] = FALSE"
            End If

            Try
                _currentDb.ExecuteNonQuery(sql)

                If CurrentDb.IsSQLExpress Then
                    sql = "UPDATE [tsysCOL] SET [tsysCOL].[IsValid] = 1 FROM [tsysCOL] INNER JOIN [qsysMTXwithCOLS] ON [tsysCOL].[ColID] = [qsysMTXwithCOLS].[ColID] WHERE [qsysMTXwithCOLS].[COEF] <> 0"
                Else
                    sql = "UPDATE [tsysCOL] INNER JOIN [qsysMTXwithCOLS] ON [tsysCOL].[ColID] = [qsysMTXwithCOLS].[ColID] SET [tsysCOL].[ISVALID] = True WHERE [qsysMTXwithCOLS].[COEF] <> 0"
                End If

                Try
                    _currentDb.ExecuteNonQuery(sql)
                    Return 0
                Catch ex As Exception
                    'GenUtils.Message(GenUtils.MsgType.Critical, "Engine-FilterEmptyColumnsRows", ex.Message)
                    EntLib.COPT.Log.Log(_workDir, " ", "Engine-FilterEmptyColumns- " & ex.Message)
                    Return 100
                End Try

            Catch ex As Exception
                'GenUtils.Message(GenUtils.MsgType.Critical, "Engine-FilterEmptyColumnsRows", ex.Message)
                EntLib.COPT.Log.Log(_workDir, " ", "Engine-FilterEmptyColumns- " & ex.Message)
                Return 100
            End Try
        End Function

        ''' <summary>
        ''' Sets/Gets Run Type (Test, Development, Debug, Production)
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks>
        ''' RKP/12-14-10/v2.4.143
        ''' </remarks>
        Public Property RunType() As GenUtils.RunType
            Get
                Return _runType
            End Get
            Set(ByVal value As GenUtils.RunType)
                _runType = value
            End Set
        End Property

    End Class
End Namespace