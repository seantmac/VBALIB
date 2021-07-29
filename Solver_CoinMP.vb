Imports System
Imports System.Text
Imports System.Linq
Imports System.Linq.Expressions
Imports System.Collections
Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports System.Data
Imports System.Data.Linq
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Data.Common
Imports System.Data.SqlServerCe
Imports System.Globalization
Imports System.Xml
Imports System.Xml.Linq
Imports System.Threading
Imports Microsoft.VisualBasic
Imports EntLib.COPT

Namespace COPT
    <Microsoft.VisualBasic.ComClass()>
    Public Class Solver_CoinMP
        Inherits Solver

        Private _solutionStatus As String
        Private _solutionRows As Integer
        Private _solutionColumns As Integer
        Private _solutionObj As Double
        Private Shadows _switches() As String
        Private Shadows _engine As COPT.Engine
        Private _progress As String = ""
        Private _ds As New DataSet

        'Private _dbPath As String = "C:\OPTMODELS\KNOXMIX\C-OPT_DATA_KNOXMIX.MDB"
        'Private _dbPath As String = "C:\OPTMODELS\CBK35\CBK35.MDB"

        'Sub Main1()
        '    Dim input As String
        '    Console.WriteLine(My.Application.Info.AssemblyName)
        '    Do
        '        Console.WriteLine()
        '        Console.WriteLine("Type ""x"" to exit.")
        '        Console.Write("Command> ")
        '        input = Console.ReadLine()
        '        'Console.WriteLine(input)
        '        Console.WriteLine()
        '        Select Case input.Trim.ToUpper()
        '            Case "A"
        '                Linq()
        '            Case "B"
        '            Case "RUN"
        '                'Console.Write("Enter path to C-OPT database: ")
        '                'input = Console.ReadLine()
        '                'Console.WriteLine(input)
        '                RunCoinMP()
        '            Case Else

        '        End Select

        '    Loop Until input.Trim().ToUpper.Equals("X")
        'End Sub

        'Private Sub Linq()
        '    Console.WriteLine("LINQ")
        'End Sub

        Private Function RunCoinMP(ByRef dtRow As DataTable, ByRef dtCol As DataTable, ByRef dtMtx As DataTable) As Boolean
            ' Add any initialization after the InitializeComponent() call.
            Dim result As Integer
            Dim length As Integer
            Dim version As Double
            Dim solverName As StringBuilder = New StringBuilder(100)
            Dim prob As New SolveProblem()
            Dim bigMRows As Integer = -1
            Dim ds As New DataSet
            Dim sql As String
            Dim updateSQLServer As Boolean = False 'RKP/08-04-11/v3.0.149
            Dim linkedTable As Boolean = False 'RKP/08-04-11/v3.0.149
            Dim success As Boolean = False 'RKP/04-18-12/v3.2.166
            Dim noTruncate As Boolean = GenUtils.IsSwitchAvailable(_switches, "/UseTruncateSP")

            If GenUtils.IsSwitchAvailable(_switches, "/UseMSAccessSyntax") Then
                linkedTable = True
            End If
            If GenUtils.IsSwitchAvailable(_switches, "/UseSQLServerSyntax") Then
                linkedTable = True
            End If

            Try 'RKP/06-23-12/v4.0.170
                result = CoinMP.InitSolver("")
                success = True
            Catch ex As System.Exception
                _progress = "" & vbNewLine & "Error initializing solver (CoinMP.InitSolver)." & vbNewLine & "Make sure you are using the correct 32-bit/64-bit version of solver." & vbNewLine & ""
                Console.WriteLine(_progress)
                Debug.Print(_progress)
                EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CoinMP", _progress)
                success = False
            End Try

            If success Then '#1
                'RKP/03-26-12/v4.0.160
                'Getting runtime error, in VS2010/.NET 4
                length = CoinMP.GetSolverName(solverName, solverName.Capacity)
                version = CoinMP.GetVersion()

                Console.WriteLine("CoinMP" & " (v" & version & ")")

                success = ProblemCOPTDynamic.Solve(Me, _engine, _switches, prob, dtRow, dtCol, dtMtx)

                _engine.Progress = "Importing rows...Start - " & My.Computer.Clock.LocalTime.ToString()
                Debug.Print("Importing rows...Start - " & My.Computer.Clock.LocalTime.ToString())
                Console.WriteLine("Importing rows...Start - " & My.Computer.Clock.LocalTime.ToString())
                EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CoinMP", "Importing solution...start - " & My.Computer.Clock.LocalTime.ToString())
                EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CoinMP", "Importing rows...Start - " & My.Computer.Clock.LocalTime.ToString())

                If GenUtils.IsSwitchAvailable(_switches, "/UseMinSysRes") Then
                    '_engine.Progress = "UpdateDataSet-dtRow-Start-" & My.Computer.Clock.LocalTime.ToString()
                    'Debug.Print("UpdateDataSet-dtRow-Start-" & My.Computer.Clock.LocalTime.ToString())

                    'GenUtils.UpdateDB(_engine.MiscParams.Item("LINKED_ROW_DB_CONN_STR").ToString(), dtRow, "SELECT * FROM [" & MyBase.Engine.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString() & "]")

                    '_engine.Progress = "UpdateDataSet-dtRow-End-" & My.Computer.Clock.LocalTime.ToString()
                    'Debug.Print("UpdateDataSet-dtRow-End-" & My.Computer.Clock.LocalTime.ToString())

                    '_engine.Progress = "UpdateDataSet-dtRow-ACTIVITY-Start-" & My.Computer.Clock.LocalTime.ToString()
                    'Debug.Print("UpdateDataSet-dtRow-ACTIVITY-Start-" & My.Computer.Clock.LocalTime.ToString())


                    If GenUtils.IsSwitchAvailable(_switches, "/UseSpreadsheet") Then
                        _progress = "Importing solution (tsysROW) with /UseSpreadsheet switch."
                        Console.WriteLine(_progress)
                        Debug.Print(_progress)
                        EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CoinMP", _progress)
                        Try
                            sql = "UPDATE tsysROW INNER JOIN tsysROW_Linked ON tsysROW.RowID = tsysROW_Linked.RowID SET tsysROW.ACTIVITY = [tsysROW_Linked].[ACTIVITY], tsysROW_Linked.SHADOW = [tsysROW_Linked].[SHADOW]"
                            EntLib.COPT.GenUtils.SpreadsheetExport(dtRow, MyBase.Engine.GetWorkDir() & "\tsysROW.xlsm", "tsysROW", MyBase.Engine.CurrentDb, sql)
                            'MyBase.Engine.CurrentDb.ExecuteNonQuery("UPDATE tsysROW INNER JOIN tsysROW_Linked ON tsysROW.RowID = tsysROW_Linked.RowID SET tsysROW.ACTIVITY = [tsysROW_Linked].[ACTIVITY], tsysROW_Linked.SHADOW = [tsysROW_Linked].[SHADOW]")
                        Catch ex As System.Exception
                            _progress = "Error importing solution (tsysROW) with /UseSpreadsheet switch. Using default option."
                            Console.WriteLine(_progress)
                            Console.WriteLine(ex.Message)
                            Debug.Print(_progress)
                            Debug.Print(ex.Message)
                            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CoinMP", _progress)
                            MyBase.Engine.CurrentDb.UpdateDataSet(dtRow, "SELECT * FROM [" & _engine.MiscParams.Item("LINKED_ROW_DB_PATH").ToString() & "].[" & MyBase.Engine.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString & "] ORDER BY RowID")
                        End Try
                    ElseIf GenUtils.IsSwitchAvailable(_switches, "/UseMSAccessSyntax") Then
                        'Try
                        '    db = New EntLib.COPT.DAAB("System.Data.OleDb", _engine.MiscParams.Item("LINKED_ROW_DB_CONN_STR").ToString())

                        'Catch ex As System.Exception

                        'End Try

                        Try

                            MyBase.Engine.CurrentDb.UpdateDataSet(dtRow, "SELECT * FROM [" & _engine.MiscParams.Item("LINKED_ROW_DB_PATH").ToString() & "].[" & MyBase.Engine.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString & "] ORDER BY RowID")

                        Catch ex As System.Exception
                            _progress = "Error using /UseMSAccessSyntax switch. Importing solution without switch."
                            Console.WriteLine(ex.Message)
                            Debug.Print(ex.Message)
                            Debug.Print("Error using /UseMSAccessSyntax switch. Importing solution without switch.")
                            Console.WriteLine("Error using /UseMSAccessSyntax switch. Importing solution without switch.")
                            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CoinMP", "Error using /UseMSAccessSyntax switch. Importing solution without switch.")

                            MyBase.Engine.CurrentDb.UpdateDataSet(dtRow, "SELECT * FROM " & MyBase.Engine.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString & " ORDER BY RowID")

                            'Debug.Print(ex.Message)
                            'Debug.Print("Error using /UseMSAccessSyntax switch. Importing solution using alternate switch.")
                            'EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CoinMP", "Error using /UseMSAccessSyntax switch. Importing solution using alternate switch. " & My.Computer.Clock.LocalTime.ToString())

                            'Try
                            '    MyBase.Engine.CurrentDb.UpdateDataSet(dtRow, "SELECT * FROM [" & _engine.MiscParams.Item("LINKED_ROW_DB_PATH").ToString() & "].[" & MyBase.Engine.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString & "] ORDER BY RowID")
                            '    EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CoinMP", "Error using /UseMSAccessSyntax switch. Importing solution using alternate switch - Done. " & My.Computer.Clock.LocalTime.ToString())
                            'Catch ex2 As Exception
                            '    Console.WriteLine(ex2.Message)
                            '    Console.WriteLine("Error using /UseMSAccessSyntax switch. Importing solution without switch.")

                            '    EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CoinMP", "Error using /UseMSAccessSyntax switch. Importing solution without switch.")

                            '    MyBase.Engine.CurrentDb.UpdateDataSet(dtRow, "SELECT * FROM " & MyBase.Engine.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString & " ORDER BY RowID")
                            'End Try
                        End Try
                    ElseIf GenUtils.IsSwitchAvailable(_switches, "/UseSQLServerSyntax") Then
                        updateSQLServer = True

                    Else
                        MyBase.Engine.CurrentDb.UpdateDataSet(dtRow, "SELECT * FROM " & MyBase.Engine.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString & " ORDER BY RowID")

                        'Try
                        '    dtRow = Nothing
                        '    dtRow = New DataTable
                        '    dtRow.TableName = "dtRow"
                        '    'dtRow.ReadXmlSchema(Engine.GetWorkDir & "\dtRow.xml")
                        '    'dtRow.ReadXml(Engine.GetWorkDir & "\dtRow.xml")

                        '    ds = New DataSet
                        '    ds.ReadXml(Engine.GetWorkDir & "\dtRow.xml")
                        '    dtRow = ds.Tables(0)

                        '    MyBase.Engine.CurrentDb.UpdateDataSet(dtRow, "SELECT * FROM " & MyBase.Engine.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString & " ORDER BY RowID")
                        '    EntLib.COPT.Log.Log(Engine.GetWorkDir, "Solver-CoinMP", "C-OPT Engine - Solver - CoinMP - Unable to update dtRow using ReadXML. Switching to default update.")
                        'Catch ex As System.Exception
                        '    Try
                        '        MyBase.Engine.CurrentDb.UpdateDataSet(dtRow, "SELECT * FROM " & MyBase.Engine.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString & " ORDER BY RowID")
                        '    Catch ex2 As Exception

                        '    End Try
                        'End Try

                    End If

                    '_engine.CurrentDb = New EntLib.COPT.DAAB(_engine.DatabaseName)

                    _engine.Progress = "UpdateDataSet-dtRow-End-" & My.Computer.Clock.LocalTime.ToString()
                    Debug.Print("UpdateDataSet-dtRow-End-" & My.Computer.Clock.LocalTime.ToString())
                    Console.WriteLine(My.Computer.Clock.LocalTime.ToString())

                Else 'If GenUtils.IsSwitchAvailable(_switches, "/UseMinSysRes") Then

                    _engine.Progress = "UpdateDataSet-dtRow-Start-" & My.Computer.Clock.LocalTime.ToString()
                    Debug.Print("UpdateDataSet-dtRow-Start-" & My.Computer.Clock.LocalTime.ToString())
                    'Console.WriteLine("Importing rows...")

                    If MyBase.Engine.CurrentDb.IsSQLExpress Then

                        If GenUtils.IsSwitchAvailable(_switches, "/UseMinSQL") Then
                            'This switch was developed for avoid using SqlBulkCopy() for SQL Azure.
                            updateSQLServer = False
                            Try
                                MyBase.Engine.CurrentDb.UpdateDataSet(dtRow, "SELECT * FROM " & MyBase.Engine.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString & " ORDER BY RowID")

                                _engine.Progress = "UpdateDataSet-dtRow-End-" & My.Computer.Clock.LocalTime.ToString()
                                Debug.Print("UpdateDataSet-dtRow-End-" & My.Computer.Clock.LocalTime.ToString())
                                Console.WriteLine(My.Computer.Clock.LocalTime.ToString())
                            Catch ex As System.Exception
                                Debug.Print("C-OPT Engine - Solver - CoinMP - Error importing Constraints (tsysRow) - " & ex.Message & " - " & My.Computer.Clock.LocalTime.ToString)
                                Console.WriteLine("C-OPT Engine - Solver - CoinMP - Error importing Constraints (tsysRow) - " & ex.Message & " - " & My.Computer.Clock.LocalTime.ToString)
                                EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CoinMP", "C-OPT Engine - Solver - CoinMP - Error importing Constraints (tsysRow) - " & ex.Message & " - " & My.Computer.Clock.LocalTime.ToString)
                            End Try
                        Else
                            updateSQLServer = True
                        End If

                    Else
                        Try
                            If GenUtils.IsSwitchAvailable(_switches, "/SubNameWithID") Then
                                MyBase.Engine.CurrentDb.UpdateDataSet(dtRow, "SELECT RowID, NULL AS ROW, NULL AS [DESC], RHS, SENSE, ACTIVITY, SHADOW, STATUS FROM " & MyBase.Engine.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString & " ORDER BY RowID")
                            Else
                                MyBase.Engine.CurrentDb.UpdateDataSet(dtRow, "SELECT * FROM " & MyBase.Engine.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString & " ORDER BY RowID")
                            End If


                            _engine.Progress = "UpdateDataSet-dtRow-End-" & My.Computer.Clock.LocalTime.ToString()
                            Debug.Print("UpdateDataSet-dtRow-End-" & My.Computer.Clock.LocalTime.ToString())
                            Console.WriteLine(My.Computer.Clock.LocalTime.ToString())
                        Catch ex As System.Exception
                            Debug.Print("C-OPT Engine - Solver - CoinMP - Error importing Constraints (tsysRow) - " & ex.Message & " - " & My.Computer.Clock.LocalTime.ToString)
                            Console.WriteLine("C-OPT Engine - Solver - CoinMP - Error importing Constraints (tsysRow) - " & ex.Message & " - " & My.Computer.Clock.LocalTime.ToString)
                            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CoinMP", "C-OPT Engine - Solver - CoinMP - Error importing Constraints (tsysRow) - " & ex.Message & " - " & My.Computer.Clock.LocalTime.ToString)
                        End Try
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
                            Debug.Print("C-OPT Engine - Solver - CoinMP - Error importing Constraints (tsysRow) - " & ex.Message & " - " & My.Computer.Clock.LocalTime.ToString)
                            Console.WriteLine("C-OPT Engine - Solver - CoinMP - Error importing Constraints (tsysRow) - " & ex.Message & " - " & My.Computer.Clock.LocalTime.ToString)
                            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CoinMP", "C-OPT Engine - Solver - CoinMP - Error importing Constraints (tsysRow) - " & ex.Message & " - " & My.Computer.Clock.LocalTime.ToString)
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
                            Debug.Print("C-OPT Engine - Solver - CoinMP - Error importing Constraints (tsysRow) - " & ex.Message & " - " & My.Computer.Clock.LocalTime.ToString)
                            Console.WriteLine("C-OPT Engine - Solver - CoinMP - Error importing Constraints (tsysRow) - " & ex.Message & " - " & My.Computer.Clock.LocalTime.ToString)
                            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CoinMP", "C-OPT Engine - Solver - CoinMP - Error importing Constraints (tsysRow) - " & ex.Message & " - " & My.Computer.Clock.LocalTime.ToString)
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
                            Debug.Print("C-OPT Engine - Solver - CoinMP - Error importing Constraints (tsysRow) - " & ex.Message & " - " & My.Computer.Clock.LocalTime.ToString)
                            Console.WriteLine("C-OPT Engine - Solver - CoinMP - Error importing Constraints (tsysRow) - " & ex.Message & " - " & My.Computer.Clock.LocalTime.ToString)
                            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CoinMP", "C-OPT Engine - Solver - CoinMP - Error importing Constraints (tsysRow) - " & ex.Message & " - " & My.Computer.Clock.LocalTime.ToString)
                        End Try
                        EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CoinMP", "C-OPT Engine - Solver - CoinMP - Successfully imported Rows. " & My.Computer.Clock.LocalTime.ToString)
                    Catch ex As System.Exception
                        _progress = "Error using /UseMSAccessSyntax switch. Importing solution without switch."
                        Console.WriteLine(ex.Message)
                        Debug.Print(ex.Message)
                        Debug.Print("Error using /UseMSAccessSyntax switch. Importing solution without switch.")
                        Console.WriteLine("Error using /UseMSAccessSyntax switch. Importing solution without switch.")
                        EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CoinMP", "Error using /UseMSAccessSyntax switch. Importing solution without switch.")

                        MyBase.Engine.CurrentDb.UpdateDataSet(dtRow, "SELECT * FROM " & MyBase.Engine.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString & " ORDER BY RowID")

                        'Debug.Print(ex.Message)
                        'Debug.Print("Error using /UseMSAccessSyntax switch. Importing solution using alternate switch.")
                        'EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CoinMP", "Error using /UseMSAccessSyntax switch. Importing solution using alternate switch. " & My.Computer.Clock.LocalTime.ToString())

                        'Try
                        '    MyBase.Engine.CurrentDb.UpdateDataSet(dtRow, "SELECT * FROM [" & _engine.MiscParams.Item("LINKED_ROW_DB_PATH").ToString() & "].[" & MyBase.Engine.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString & "] ORDER BY RowID")
                        '    EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CoinMP", "Error using /UseMSAccessSyntax switch. Importing solution using alternate switch - Done. " & My.Computer.Clock.LocalTime.ToString())
                        'Catch ex2 As Exception
                        '    Console.WriteLine(ex2.Message)
                        '    Console.WriteLine("Error using /UseMSAccessSyntax switch. Importing solution without switch.")

                        '    EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CoinMP", "Error using /UseMSAccessSyntax switch. Importing solution without switch.")

                        '    MyBase.Engine.CurrentDb.UpdateDataSet(dtRow, "SELECT * FROM " & MyBase.Engine.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString & " ORDER BY RowID")
                        'End Try
                    End Try
                End If

                _engine.Progress = "Importing rows...End - " & My.Computer.Clock.LocalTime.ToString()
                Debug.Print("Importing rows...End - " & My.Computer.Clock.LocalTime.ToString())
                Console.WriteLine("Importing rows...End - " & My.Computer.Clock.LocalTime.ToString())
                EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CoinMP", "Importing rows...End - " & My.Computer.Clock.LocalTime.ToString())

                'Try
                '    MyBase.Engine.CurrentDb.UpdateDataSet(dtCol, "SELECT * FROM " & MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString & " ORDER BY ColID")
                'Catch ex As System.Exception
                '    Debug.Print("C-OPT Engine - Solver - CoinMP - Error importing Decision Variables (tsysCol) - " & ex.Message & " - " & My.Computer.Clock.LocalTime.ToString)
                '    Console.WriteLine("C-OPT Engine - Solver - CoinMP - Error importing Decision Variables (tsysCol) - " & ex.Message & " - " & My.Computer.Clock.LocalTime.ToString)
                '    EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CoinMP", "C-OPT Engine - Solver - CoinMP - Error importing Decision Variables (tsysCol) - " & ex.Message & " - " & My.Computer.Clock.LocalTime.ToString)
                'End Try

                'tsysCol
                'If GenUtils.IsSwitchAvailable(_switches, "/UseMinSysRes0") Then
                'ElseIf GenUtils.IsSwitchAvailable(_switches, "/UseMinSysRes1") Then
                '    _engine.Progress = "UpdateDataSet-dtCol-Start-" & My.Computer.Clock.LocalTime.ToString()
                '    Debug.Print("UpdateDataSet-dtCol-Start-" & My.Computer.Clock.LocalTime.ToString())

                '    'Dim provider As System.Data.Common.DbProviderFactory = System.Data.Common.DbProviderFactories.GetFactory(MyBase.Engine.CurrentDb.GetProviderName)
                '    Dim provider As System.Data.Common.DbProviderFactory = System.Data.Common.DbProviderFactories.GetFactory(GenUtils.GetAppSettings("linkedColDBProviderName"))
                '    Dim conn As System.Data.OleDb.OleDbConnection = provider.CreateConnection()
                '    'Dim sysDB As String = "C:\OPTMODELS\C-OPTSYS\System.mdw.copt" 'Environ("USERPROFILE") & "\Application Data\Microsoft\Access\System.mdw"
                '    'conn.ConnectionString = MyBase.Engine.CurrentDb.GetConnectionString
                '    conn.ConnectionString = GenUtils.GetAppSettings("linkedColDBConnectionString")
                '    conn.Open()

                '    Dim myDataAdapter As New OleDbDataAdapter()
                '    'myDataAdapter.SelectCommand = New OleDbCommand("SELECT * FROM " & MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString & " ORDER BY ColID", conn) 'MyBase.Engine.CurrentDb.GetDbConnection)
                '    myDataAdapter.SelectCommand = New OleDbCommand("SELECT * FROM " & MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString & " ORDER BY ColID", conn) 'MyBase.Engine.CurrentDb.GetDbConnection)
                '    Dim myCB As OleDbCommandBuilder = New OleDbCommandBuilder(myDataAdapter)
                '    myDataAdapter.Update(dtCol)

                '    conn.Close()

                '    _engine.Progress = "UpdateDataSet-dtCol-End-" & My.Computer.Clock.LocalTime.ToString()
                '    Debug.Print("UpdateDataSet-dtCol-End-" & My.Computer.Clock.LocalTime.ToString())
                'ElseIf GenUtils.IsSwitchAvailable(_switches, "/UseMinSysRes2") Then
                '    Try
                '        bigMRows = MyBase.Engine.MiscParams.Item("BIGMROWS")
                '    Catch ex As System.Exception
                '        bigMRows = 50000
                '    End Try
                '    'If dtRow.Rows.Count >= MyBase.Engine.MiscParams.Item("BIGMROWS").ToString Then
                '    'End If
                '    rowCtr = 0
                '    rowCnt = dtCol.Rows.Count
                '    'dtTmp = New DataTable
                '    'dtTmp = Nothing
                '    'dtTmp = dtCol.Clone
                '    'For Each dr In dtCol.Rows
                '    '    rowCtr = rowCtr + 1
                '    '    If rowCtr >= bigMRows Then
                '    '        'UpdateDataSet
                '    '        Try
                '    '            Debug.Print("dtCol-UpdateDataSet-RowCtr= " & rowCtr.ToString() & " -Start-" & My.Computer.Clock.LocalTime.ToString())
                '    '            MyBase.Engine.CurrentDb.UpdateDataSet(dtTmp, "SELECT * FROM " & MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString & " ORDER BY ColID")
                '    '            Debug.Print("dtCol-UpdateDataSet-RowCtr= " & rowCtr.ToString() & " -End-" & My.Computer.Clock.LocalTime.ToString())
                '    '        Catch ex As System.Exception
                '    '            'MessageBox.Show(ex.Message)
                '    '            Debug.Print(ex.Message)
                '    '        End Try
                '    '        dtTmp = Nothing
                '    '        dtTmp = dtCol.Clone
                '    '        dtTmp.ImportRow(dr)

                '    '    Else
                '    '        'dtTmp.Rows.Add(dr)
                '    '        dtTmp.ImportRow(dr)
                '    '    End If

                '    'Next

                '    _engine.Progress = "UpdateDataSet-dtCol-Start-" & My.Computer.Clock.LocalTime.ToString()
                '    Debug.Print("UpdateDataSet-dtCol-Start-" & My.Computer.Clock.LocalTime.ToString())
                '    For Each dr In dtCol.Rows
                '        sb = New StringBuilder
                '        sb.Append("UPDATE [")
                '        sb.Append(MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString())
                '        sb.Append("] SET [ACTIVITY] = ")
                '        sb.Append(dr.Item("ACTIVITY"))
                '        sb.Append(" ")
                '        sb.Append(", ")
                '        sb.Append("DJ = ")
                '        sb.Append(dr.Item("DJ"))
                '        sb.Append(" ")
                '        sb.Append("WHERE ")
                '        sb.Append("ColID = ")
                '        sb.Append(dr.Item("ColID"))
                '        sql = sb.ToString()
                '        rowCnt = _engine.CurrentDb.ExecuteNonQuery(sql)
                '    Next
                '    _engine.Progress = "UpdateDataSet-dtCol-End-" & My.Computer.Clock.LocalTime.ToString()
                '    Debug.Print("UpdateDataSet-dtCol-End-" & My.Computer.Clock.LocalTime.ToString())

                _engine.Progress = "Importing columns...Start - " & My.Computer.Clock.LocalTime.ToString()
                Debug.Print("Importing columns...Start - " & My.Computer.Clock.LocalTime.ToString())
                Console.WriteLine("Importing columns...Start - " & My.Computer.Clock.LocalTime.ToString())
                EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CoinMP", "Importing columns...Start - " & My.Computer.Clock.LocalTime.ToString())

                If GenUtils.IsSwitchAvailable(_switches, "/UseMinSysRes") Then
                    '_engine.Progress = "UpdateDataSet-dtCol-ACTIVITY-Start-" & My.Computer.Clock.LocalTime.ToString()
                    'Debug.Print("UpdateDataSet-dtCol-ACTIVITY-Start-" & My.Computer.Clock.LocalTime.ToString())

                    _engine.Progress = "UpdateDataSet-dtCol-Start-" & My.Computer.Clock.LocalTime.ToString()
                    Debug.Print("UpdateDataSet-dtCol-Start-" & My.Computer.Clock.LocalTime.ToString())
                    Console.WriteLine("Importing columns...")

                    If GenUtils.IsSwitchAvailable(_switches, "/UseSpreadsheet") Then
                        _progress = "Importing solution (tsysCOL) with /UseSpreadsheet switch."
                        Console.WriteLine(_progress)
                        Debug.Print(_progress)
                        EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CoinMP", _progress)

                        Try
                            'sql = "UPDATE tsysCOL INNER JOIN tsysCOL_Linked ON tsysCOL.ColID = tsysCOL_Linked.ColID SET tsysCOL.ACTIVITY = [tsysCOL_Linked].[ACTIVITY], tsysCOL.DJ = [tsysCOL_Linked].[DJ]"
                            EntLib.COPT.GenUtils.SpreadsheetExport(dtCol, MyBase.Engine.GetWorkDir() & "\tsysCOL.xlsm", "tsysCOL", MyBase.Engine.CurrentDb, "")
                            'MyBase.Engine.CurrentDb.ExecuteNonQuery("UPDATE tsysCOL INNER JOIN tsysCOL_Linked ON tsysCOL.ColID = tsysCOL_Linked.ColID SET tsysCOL.ACTIVITY = [tsysCOL_Linked].[ACTIVITY], tsysCOL.DJ = [tsysCOL_Linked].[DJ]")
                        Catch ex As System.Exception
                            _progress = "Error importing solution (tsysCOL) with /UseSpreadsheet switch. Using default option."
                            Console.WriteLine(_progress)
                            Debug.Print(_progress)
                            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CoinMP", _progress)
                            MyBase.Engine.CurrentDb.UpdateDataSet(dtCol, "SELECT * FROM " & MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString & " ORDER BY ColID")
                        End Try
                    ElseIf GenUtils.IsSwitchAvailable(_switches, "/UseMSAccessSyntax") Then
                        Try
                            'GenUtils.UpdateDB(_engine.MiscParams.Item("LINKED_COL_DB_CONN_STR").ToString(), dtCol, "SELECT * FROM [" & MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString() & "] ORDER BY ColID")
                            MyBase.Engine.CurrentDb.UpdateDataSet(dtCol, "SELECT * FROM [" & _engine.MiscParams.Item("LINKED_COL_DB_PATH").ToString() & "].[" & MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString & "] ORDER BY ColID")
                        Catch ex As System.Exception
                            _progress = "Error using /UseMSAccessSyntax switch. Importing solution without switch."
                            Console.WriteLine(ex.Message)
                            Debug.Print(ex.Message)
                            Debug.Print("Error using /UseMSAccessSyntax switch. Importing solution without switch.")
                            Console.WriteLine("Error using /UseMSAccessSyntax switch. Importing solution without switch.")
                            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CoinMP", "Error using /UseMSAccessSyntax switch. Importing solution without switch.")

                            MyBase.Engine.CurrentDb.UpdateDataSet(dtCol, "SELECT * FROM " & MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString & " ORDER BY ColID")

                            'Debug.Print(ex.Message)
                            'Debug.Print("Error using /UseMSAccessSyntax switch. Importing solution without switch.")

                            'Console.WriteLine(ex.Message)
                            'Console.WriteLine("Error using /UseMSAccessSyntax switch. Importing solution without switch.")

                            'EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CoinMP", "Error using /UseMSAccessSyntax switch. Importing solution without switch.")

                            'MyBase.Engine.CurrentDb.UpdateDataSet(dtCol, "SELECT * FROM " & MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString & " ORDER BY ColID")

                            'Debug.Print(ex.Message)
                            'Debug.Print("Error using /UseMSAccessSyntax switch. Importing solution using alternate switch.")
                            'EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CoinMP", "Error using /UseMSAccessSyntax switch. Importing solution using alternate switch. " & My.Computer.Clock.LocalTime.ToString())

                            'Debug.Print(ex.Message)
                            'Debug.Print("Error using /UseMSAccessSyntax switch. Importing solution using alternate switch.")
                            'EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CoinMP", "Error using /UseMSAccessSyntax switch. Importing solution using alternate switch. " & My.Computer.Clock.LocalTime.ToString())

                            'Try
                            '    MyBase.Engine.CurrentDb.UpdateDataSet(dtCol, "SELECT * FROM [" & _engine.MiscParams.Item("LINKED_COL_DB_PATH").ToString() & "].[" & MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString & "] ORDER BY ColID")
                            '    EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CoinMP", "Error using /UseMSAccessSyntax switch. Importing solution using alternate switch - Done. " & My.Computer.Clock.LocalTime.ToString())
                            'Catch ex2 As Exception
                            '    Console.WriteLine(ex2.Message)
                            '    Console.WriteLine("Error using /UseMSAccessSyntax switch. Importing solution without switch.")

                            '    EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CoinMP", "Error using /UseMSAccessSyntax switch. Importing solution without switch.")

                            '    MyBase.Engine.CurrentDb.UpdateDataSet(dtCol, "SELECT * FROM " & MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString & " ORDER BY ColID")
                            'End Try
                        End Try
                    ElseIf GenUtils.IsSwitchAvailable(_switches, "/UseSQLServerSyntax") Then
                        updateSQLServer = True

                    ElseIf GenUtils.IsSwitchAvailable(_switches, "/UseMSAccessSyntax2") Then

                    Else
                        MyBase.Engine.CurrentDb.UpdateDataSet(dtCol, "SELECT * FROM " & MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString & " ORDER BY ColID")
                    End If

                    _engine.Progress = "UpdateDataSet-dtCol-End-" & My.Computer.Clock.LocalTime.ToString()
                    Debug.Print("UpdateDataSet-dtCol-End-" & My.Computer.Clock.LocalTime.ToString())
                    Console.WriteLine(My.Computer.Clock.LocalTime.ToString())

                Else 'If GenUtils.IsSwitchAvailable(_switches, "/UseMinSysRes") Then

                    If MyBase.Engine.CurrentDb.IsSQLExpress Then
                        If GenUtils.IsSwitchAvailable(_switches, "/UseMinSQL") Then
                            updateSQLServer = False
                            Try
                                _engine.Progress = "UpdateDataSet-dtCol-Start-" & My.Computer.Clock.LocalTime.ToString()
                                Debug.Print("UpdateDataSet-dtCol-Start-" & My.Computer.Clock.LocalTime.ToString())
                                Console.WriteLine("Importing columns...")

                                'MyBase.Engine.CurrentDb.UpdateDataSet(dtCol, "SELECT * FROM " & MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString & " ORDER BY ColID")
                                MyBase.Engine.CurrentDb.UpdateDataSet(dtCol, "SELECT * FROM " & MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString & " ORDER BY ColID")

                                'MyBase.Engine.CurrentDb.UpdateDataSet(dtRow, "SELECT * FROM " & MyBase.Engine.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString & " ORDER BY RowID")

                                'Dim myRow As DataRow
                                'For Each myRow In dtCol.Rows
                                '    If myRow.HasErrors Then
                                '        Console.WriteLine(myRow(0).ToString() & vbCrLf & myRow.RowError.ToString())
                                '    End If
                                'Next

                                _engine.Progress = "UpdateDataSet-dtCol-End-" & My.Computer.Clock.LocalTime.ToString()
                                Debug.Print("UpdateDataSet-dtCol-End-" & My.Computer.Clock.LocalTime.ToString())
                                Console.WriteLine(My.Computer.Clock.LocalTime.ToString())
                            Catch ex As System.Exception
                                Debug.Print("C-OPT Engine - Solver - CoinMP - Error importing Decision Variables (tsysCol) - " & ex.Message & " - " & My.Computer.Clock.LocalTime.ToString)
                                Console.WriteLine("C-OPT Engine - Solver - CoinMP - Error importing Decision Variables (tsysCol) - " & ex.Message & " - " & My.Computer.Clock.LocalTime.ToString)
                                EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CoinMP", "C-OPT Engine - Solver - CoinMP - Error importing Decision Variables (tsysCol) - " & ex.Message & " - " & My.Computer.Clock.LocalTime.ToString)
                            End Try
                        Else
                            updateSQLServer = True
                        End If
                    Else
                        Try
                            _engine.Progress = "UpdateDataSet-dtCol-Start-" & My.Computer.Clock.LocalTime.ToString()
                            Debug.Print("UpdateDataSet-dtCol-Start-" & My.Computer.Clock.LocalTime.ToString())
                            Console.WriteLine("Importing columns...")

                            'MyBase.Engine.CurrentDb.UpdateDataSet(dtCol, "SELECT * FROM " & MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString & " ORDER BY ColID")

                            If GenUtils.IsSwitchAvailable(_switches, "/SubNameWithID") Then
                                MyBase.Engine.CurrentDb.UpdateDataSet(dtCol, "SELECT ColID, NULL AS COL, NULL AS [DESC], OBJ, LO, UP, FREE, INTGR, BINRY, NULL AS SOSTYPE, NULL AS SOSMARKER, ACTIVITY, DJ, STATUS, ISVALID FROM " & MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString & " ORDER BY ColID")
                            Else
                                MyBase.Engine.CurrentDb.UpdateDataSet(dtCol, "SELECT * FROM " & MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString & " ORDER BY ColID")
                            End If

                            'MyBase.Engine.CurrentDb.UpdateDataSet(dtRow, "SELECT * FROM " & MyBase.Engine.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString & " ORDER BY RowID")

                            'Dim myRow As DataRow
                            'For Each myRow In dtCol.Rows
                            '    If myRow.HasErrors Then
                            '        Console.WriteLine(myRow(0).ToString() & vbCrLf & myRow.RowError.ToString())
                            '    End If
                            'Next

                            _engine.Progress = "UpdateDataSet-dtCol-End-" & My.Computer.Clock.LocalTime.ToString()
                            Debug.Print("UpdateDataSet-dtCol-End-" & My.Computer.Clock.LocalTime.ToString())
                            Console.WriteLine(My.Computer.Clock.LocalTime.ToString())
                        Catch ex As System.Exception
                            Debug.Print("C-OPT Engine - Solver - CoinMP - Error importing Decision Variables (tsysCol) - " & ex.Message & " - " & My.Computer.Clock.LocalTime.ToString)
                            Console.WriteLine("C-OPT Engine - Solver - CoinMP - Error importing Decision Variables (tsysCol) - " & ex.Message & " - " & My.Computer.Clock.LocalTime.ToString)
                            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CoinMP", "C-OPT Engine - Solver - CoinMP - Error importing Decision Variables (tsysCol) - " & ex.Message & " - " & My.Computer.Clock.LocalTime.ToString)
                        End Try
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
                            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CoinMP", "Error using /UseSQLServerSyntax switch. Importing solution without switch.")
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
                            Debug.Print("C-OPT Engine - Solver - CoinMP - Error importing Columns (tsysCol) - " & ex.Message & " - " & My.Computer.Clock.LocalTime.ToString)
                            Console.WriteLine("C-OPT Engine - Solver - CoinMP - Error importing Columns (tsysCol) - " & ex.Message & " - " & My.Computer.Clock.LocalTime.ToString)
                            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CoinMP", "C-OPT Engine - Solver - CoinMP - Error importing Columns (tsysCol) - " & ex.Message & " - " & My.Computer.Clock.LocalTime.ToString)
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
                            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CoinMP", "Error using /UseSQLServerSyntax switch. Importing solution without switch.")
                        End Try
                        EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CoinMP", "C-OPT Engine - Solver - CoinMP - Successfully imported Columns. " & My.Computer.Clock.LocalTime.ToString)
                    Catch ex As System.Exception
                        _progress = "Error using /UseSQLServerSyntax switch. Importing solution without switch."
                        Console.WriteLine(ex.Message)
                        Debug.Print(ex.Message)
                        Debug.Print("Error using /UseSQLServerSyntax switch. Importing solution without switch.")
                        Console.WriteLine("Error using /UseSQLServerSyntax switch. Importing solution without switch.")
                        EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CoinMP", "Error using /UseSQLServerSyntax switch. Importing solution without switch.")

                        MyBase.Engine.CurrentDb.UpdateDataSet(dtCol, "SELECT * FROM " & MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString & " ORDER BY ColID")

                        'Debug.Print(ex.Message)
                        'Debug.Print("Error using /UseMSAccessSyntax switch. Importing solution without switch.")

                        'Console.WriteLine(ex.Message)
                        'Console.WriteLine("Error using /UseMSAccessSyntax switch. Importing solution without switch.")

                        'EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CoinMP", "Error using /UseMSAccessSyntax switch. Importing solution without switch.")

                        'MyBase.Engine.CurrentDb.UpdateDataSet(dtCol, "SELECT * FROM " & MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString & " ORDER BY ColID")

                        'Debug.Print(ex.Message)
                        'Debug.Print("Error using /UseMSAccessSyntax switch. Importing solution using alternate switch.")
                        'EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CoinMP", "Error using /UseMSAccessSyntax switch. Importing solution using alternate switch. " & My.Computer.Clock.LocalTime.ToString())

                        'Debug.Print(ex.Message)
                        'Debug.Print("Error using /UseMSAccessSyntax switch. Importing solution using alternate switch.")
                        'EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CoinMP", "Error using /UseMSAccessSyntax switch. Importing solution using alternate switch. " & My.Computer.Clock.LocalTime.ToString())

                        'Try
                        '    MyBase.Engine.CurrentDb.UpdateDataSet(dtCol, "SELECT * FROM [" & _engine.MiscParams.Item("LINKED_COL_DB_PATH").ToString() & "].[" & MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString & "] ORDER BY ColID")
                        '    EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CoinMP", "Error using /UseMSAccessSyntax switch. Importing solution using alternate switch - Done. " & My.Computer.Clock.LocalTime.ToString())
                        'Catch ex2 As Exception
                        '    Console.WriteLine(ex2.Message)
                        '    Console.WriteLine("Error using /UseMSAccessSyntax switch. Importing solution without switch.")

                        '    EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CoinMP", "Error using /UseMSAccessSyntax switch. Importing solution without switch.")

                        '    MyBase.Engine.CurrentDb.UpdateDataSet(dtCol, "SELECT * FROM " & MyBase.Engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString & " ORDER BY ColID")
                        'End Try
                    End Try
                End If

                _engine.Progress = "Importing columns...End - " & My.Computer.Clock.LocalTime.ToString()
                Debug.Print("Importing columns...End - " & My.Computer.Clock.LocalTime.ToString())
                Console.WriteLine("Importing columns...End - " & My.Computer.Clock.LocalTime.ToString())
                EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CoinMP", "Importing columns...End - " & My.Computer.Clock.LocalTime.ToString())

                Debug.Print("C-OPT Engine - Solver - CoinMP - Importing solution...done - " & My.Computer.Clock.LocalTime.ToString)
                Console.WriteLine("C-OPT Engine - Solver - CoinMP - Importing solution...done - " & My.Computer.Clock.LocalTime.ToString)
                EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CoinMP", "Importing solution...done - " & My.Computer.Clock.LocalTime.ToString())
            End If 'success #1
            Return success
        End Function

        ' A creatable COM class must have a Public Sub New() 
        ' with no parameters, otherwise, the class will not be 
        ' registered in the COM registry and cannot be created 
        ' via CreateObject.
        Public Sub New()
            MyBase.New(New COPT.Engine)
        End Sub

        Public Sub New(ByRef engine As COPT.Engine)
            MyBase.New(engine)
            _engine = engine
        End Sub

        Public Sub New(ByRef engine As COPT.Engine, ByVal switches() As String)
            MyBase.New(engine, switches)
            _engine = engine
            _switches = switches
        End Sub

        Public Overrides Function Solve(ByRef ds As DataSet) As Boolean
            Return Solve(ds.Tables("tsysRow"), ds.Tables("tsysCol"), ds.Tables("tsysMtx"))
        End Function

        Public Overrides Function Solve(ByRef dtRow As DataTable, ByRef dtCol As DataTable, ByRef dtMtx As DataTable) As Boolean
            Dim startTime As Integer = My.Computer.Clock.TickCount
            Dim success As Boolean = False 'RKP/04-18-12/v3.2.166

            '_ds.Tables.Add(dtCol)
            '_ds.Tables(0).TableName = "tsysCol"

            '_ds.Tables.Add(dtRow)
            '_ds.Tables(1).TableName = "tsysRow"

            '_ds.Tables.Add(dtMtx)
            '_ds.Tables(2).TableName = "tsysMtx"

            success = RunCoinMP(dtRow, dtCol, dtMtx)

            EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - CoinMP - Total solve took: ", GenUtils.FormatTime(startTime, My.Computer.Clock.TickCount))
            Return success
        End Function

        Public Overrides Function getSolutionStatusDescription(ByVal sts As Integer) As String
            Try
                'Dim descrArray(12) As String

                '0 = LpStatusOptimal
                '1 = LpStatusInfeasible
                '2 = LpStatusInfeasible
                '3 = LpStatusNotSolved
                '4 = LpStatusNotSolved
                '5 = LpStatusNotSolved
                '-1= LpStatusUndefined

                'XA solution status
                'descrArray(1) = "Optimal Solution"
                'descrArray(2) = "Integer Solution (not proven the optimal integer solution)"
                'descrArray(3) = "Unbounded Solution"
                'descrArray(4) = "Infeasible Solution"
                'descrArray(5) = "Callback function indicates Infeasible Solution"
                'descrArray(6) = "Intermediate Infeasible Solution"
                'descrArray(7) = "Intermediate Non-optimal Solution"
                'descrArray(8) = ""
                'descrArray(9) = "Intermediate Non-integer Solution"
                'descrArray(10) = "Integer Infeasible"
                'descrArray(11) = ""
                'descrArray(12) = "Error Unknown"

                'Select Case sts
                '    Case -1
                '        _solutionStatus = "Undefined Error"
                '    Case 0
                '        _solutionStatus = "Optimal Solution"
                '    Case 1
                '        _solutionStatus = "Infeasible Solution"
                '    Case 2
                '        _solutionStatus = "Infeasible Solution"
                '    Case 3
                '        _solutionStatus = "Not Solved"
                '    Case 4
                '        _solutionStatus = "Not Solved"
                '    Case 5
                '        _solutionStatus = "Not Solved"
                '    Case Else
                '        _solutionStatus = "Unknown Error"
                'End Select

                Select Case sts
                    Case 0
                        Return "Optimal solution"
                    Case 1
                        Return "Problem primal infeasible"
                    Case 2
                        Return "Problem dual infeasible"
                    Case 3
                        Return "Stopped on iterations"
                    Case 4
                        Return "Stopped due to errors"
                    Case 5
                        Return "Stopped by user"
                    Case Else
                        Return "Undefined error"
                End Select

                Return _solutionStatus 'descrArray(sts)
            Catch ex As System.Exception
                Return "Unknown Error"
            End Try
        End Function

        Public Overloads Overrides Function getSolutionStatusDescription() As String
            Return _solutionStatus
        End Function

        Public Overrides Function getSolverName() As String
            Return "CoinMP"
        End Function

        Public Overloads Overrides Function Solve() As Boolean

        End Function

        Public Sub setSolutionStatusDescription(ByVal sts As String)
            _solutionStatus = sts
        End Sub
    End Class 'Solver_CoinMP

    Module CoinMP
        'Compatible with CoinMP v1.4
        Public Enum CallResult As Integer
            Success = 0
            Failed = 1
        End Enum

        Public Enum MethodType As Integer
            Auto = 0
            Primal = &H1
            Dual = &H2
            Network = &H4
            Barrier = &H8
            Benders = &H100
            DEQ = &H200
            EV = &H400
        End Enum

        Public Enum FeatureType As Integer
            Auto = 0
            LP = &H1
            QP = &H2
            QCP = &H4
            NLP = &H8
            MIP = &H10
            MIQP = &H20
            MIQCP = &H40
            MINLP = &H80
            SP = &H10000
        End Enum

        Public Enum LoadNamesType As Integer
            Auto = 0
            List = 1
            Buffer = 2
        End Enum

        Public Enum ObjectSense As Integer
            Max = -1
            None = 0
            Min = 1
        End Enum

        Public Enum FileType As Integer
            Log = 0
            Basis = 1
            MipStart = 2
            MPS = 3
            LP = 4
            Binary = 5
            Output = 6
            BinOut = 7
            IIS = 8
        End Enum

        Public Enum CheckResult As Integer
            Passed = 0
            ColCount = 1
            RowCount = 2
            RangeCount = 3
            ObjSense = 4
            RowType = 5
            MatBegin = 6
            MatCount = 7
            MatBegCount = 8
            MatBegNonzero = 9
            MatIndex = 10
            MatIndexRow = 11
            Bounds = 12
            ColType = 13
            ColNames = 14
            ColNamesLen = 15
            RowNames = 16
            RowNamesLen = 17
        End Enum

        Public Enum OptionType As Integer
            None = 0
            OnOff = 1
            List = 2
            Int = 3
            Real = 4
            Str = 5
        End Enum

        Public Enum GroupType As Integer
            None = 0
            Simplex = 1
            PreProc = 2
            LogFile = 3
            Limits = 4
            MipStrat = 5
            MipCuts = 6
            MipTol = 7
            Barrier = 8
            Network = 9
        End Enum

        'https://projects.coin-or.org/CoinMP/browser/trunk/CoinMP/src/CoinMP.h?rev=317
        Public Enum OptionParam As Integer
            SolveMethod = 1
            PresolveType = 2
            Scaling = 3
            Perturbation = 4
            PrimalPivotAlg = 5
            DualPivotAlg = 6
            LogLevel = 7
            MaxIter = 8
            CrashInd = 9
            CrashPivot = 10
            CrashGap = 11
            PrimalObjLim = 12
            DualObjLim = 13
            PrimalObjTol = 14
            DualObjTol = 15
            MaxSeconds = 16

            MipMaxNodes = 17
            MipMaxSol = 18 '/SetSolutionLimit
            MipMaxSec = 19 '/SetTimeLimit

            MipFathomDisc = 20
            MipHotStart = 21
            MipMinimumDrop = 22
            MipMaxCutPass = 23
            MipMaxPassRoot = 24
            MipStrongBranch = 25
            MipScanGlobCuts = 26

            MipIntTol = 30 '/SetRelativeGapTolerance
            MipInfWeight = 31
            MipCutOff = 32
            MipAbsGap = 33 'MipAllowableGap
            MipFracGap = 34 'RKP/v4.0.170/06-21-12

            MipCutProbing = 110

            MipProbeFreq = 111
            MipProbeMode = 112
            MipProbeUseObj = 113
            MipProbeMaxPass = 114
            MipProbeMaxProbe = 115
            MipProbeMaxLook = 116
            MipProbeRowCuts = 117

            MipCutGomory = 120
            MipGomoryFreq = 121
            MipGomoryLimit = 122
            MipGomoryAway = 123

            MipCutKnapsack = 130
            MipKnapsackFreq = 131
            MipKnapsackMaxIn = 132

            MipCutOddHole = 140
            MipOddHoleFreq = 141
            MipOddHoleMinViol = 142
            MipOddHoleMinViolPer = 143
            MipOddHoleMaxEntries = 144

            MipCutClique = 150
            MipCliqueFreq = 151
            MipCliquePacking = 152
            MipCliqueStar = 153
            MipCliqueStarMethod = 154
            MipCliqueStarMaxLen = 155
            MipCliqueStarReport = 156
            MipCliqueRow = 157
            MipCliqueRowMaxLen = 158
            MipCliqueRowReport = 159
            MipCliqueMinViol = 160

            MipCutLiftProject = 170
            MipLiftProFreq = 171
            MipLiftProBetaOne = 172

            MipUseCBCMain = 200

        End Enum


        Public Delegate Function MsgLogDelegate(ByVal MessageStr As String) As Integer

        Public Delegate Function IterDelegate(ByVal IterCount As Integer, ByVal ObjectValue As Double, _
                                        ByVal IsFeasible As Integer, ByVal InfeasValue As Double) As Integer

        Public Delegate Function MipNodeDelegate(ByVal IterCount As Integer, ByVal MipNodeCount As Integer, _
                                        ByVal BestBound As Double, ByVal BestInteger As Double, _
                                        ByVal IsMipImproved As Integer) As Integer

#If win32 Then
        <DllImport("coinmp.dll", EntryPoint:="CoinInitSolver")> _
        Public Function InitSolver(ByVal LicenseStr As String) As CallResult
        End Function
#ElseIf win64 Then
        <DllImport("coinmp.dll", EntryPoint:="CoinInitSolver")> _
        Public Function InitSolver(ByVal LicenseStr As String) As CallResult
        End Function
#Else
        <DllImport("coinmp.dll", EntryPoint:="CoinInitSolver")> _
        Public Function InitSolver(ByVal LicenseStr As String) As CallResult
        End Function
#End If



        <DllImport("coinmp.dll", EntryPoint:="CoinFreeSolver")> _
        Public Function FreeSolver() As CallResult
        End Function

        '<DllImport("coinmp.dll", EntryPoint:="CoinGetSolverName")> _
        'Public Function GetSolverName(ByVal SolverName As StringBuilder, ByVal buflen As Integer) As Integer
        'End Function

        <DllImport("coinmp.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl, EntryPoint:="CoinGetSolverName")> _
        Public Function GetSolverName(ByVal SolverName As StringBuilder, ByVal buflen As Integer) As Integer
        End Function

        <DllImport("coinmp.dll", EntryPoint:="CoinGetVersionStr")> _
        Public Function GetVersionStr(ByVal VersionStr As StringBuilder, ByVal buflen As Integer) As Integer
        End Function

        <DllImport("coinmp.dll", EntryPoint:="CoinGetVersion")> _
        Public Function GetVersion() As Double
        End Function

        <DllImport("coinmp.dll", EntryPoint:="CoinGetFeatures")> _
        Public Function GetFeatures() As FeatureType
        End Function

        <DllImport("coinmp.dll", EntryPoint:="CoinGetMethods")> _
        Public Function GetMethods() As MethodType
        End Function

        <DllImport("coinmp.dll", EntryPoint:="CoinGetInfinity")> _
        Public Function GetInfinity() As Double
        End Function

        <DllImport("coinmp.dll", EntryPoint:="CoinCreateProblem")> _
        Public Function CreateProblem(ByVal problemName As String) As IntPtr
        End Function


        <DllImport("coinmp.dll", EntryPoint:="CoinLoadProblem")> _
        Public Function LoadProblem(ByVal hProb As IntPtr, ByVal colCount As Integer, ByVal rowCount As Integer, _
                        ByVal nonzeroCount As Integer, ByVal rangeCount As Integer, ByVal objectSense As ObjectSense, _
                        ByVal objectConst As Double, ByVal objectCoeffs() As Double, ByVal lowerBounds() As Double, _
                        ByVal upperBounds() As Double, ByVal rowType() As Char, ByVal rhsValues() As Double, _
                        ByVal rangeValues() As Double, ByVal matrixBegin() As Integer, ByVal matrixCount() As Integer, _
                        ByVal matrixIndex() As Integer, ByVal matrixValues() As Double, ByVal colNamesBuf As String, _
                        ByVal rowNamesBuf As String, _
                        ByVal objName As String) As CallResult
        End Function

        'Public Declare Function CoinLoadProblemBuf Lib "CoinMP.dll" (ByVal hProb As Long, _
        '    ByVal colCount As Long, ByVal rowCount As Long, ByVal nonZeroCount As Long, ByVal rangeCount As Long, _
        '    ByVal objectSense As Long, ByVal objectConst As Double, ByRef objectCoeffs As Double, _
        '    ByRef lowerBounds As Double, ByRef upperBounds As Double, ByRef rowType As Byte, _
        '    ByRef rhsValues As Double, ByRef rangeValues As Double, ByRef matrixBegin As Long, _
        '    ByRef matrixCount As Long, ByRef matrixIndex As Long, ByRef matrixValues As Double, _
        '    ByVal colNamesBuf As String, ByVal rowNamesBuf As String, ByVal objName As String) As Integer

        Public Function LoadProblem(ByVal hProb As IntPtr, ByVal colCount As Integer, ByVal rowCount As Integer, _
                        ByVal nonzeroCount As Integer, ByVal rangeCount As Integer, ByVal objectSense As ObjectSense, _
                        ByVal objectConst As Double, ByVal objectCoeffs() As Double, ByVal lowerBounds() As Double, _
                        ByVal upperBounds() As Double, ByVal rowType() As Char, ByVal rhsValues() As Double, _
                        ByVal rangeValues() As Double, ByVal matrixBegin() As Integer, ByVal matrixCount() As Integer, _
                        ByVal matrixIndex() As Integer, ByVal matrixValues() As Double, ByVal colNames() As String, _
                        ByVal rowNames() As String, ByVal objName As String) As CallResult
            Dim result As CallResult
            'Dim status As Integer
            result = SetLoadNamesType(hProb, LoadNamesType.Buffer)

            'Try
            result = LoadProblem(hProb, colCount, rowCount, nonzeroCount, rangeCount, objectSense, _
                        objectConst, objectCoeffs, lowerBounds, upperBounds, rowType, rhsValues, _
                        rangeValues, matrixBegin, matrixCount, matrixIndex, matrixValues, _
                        GenerateNamesBuf(colNames), GenerateNamesBuf(rowNames), objName)
            'Catch ex As System.Exception

            'End Try

            'status = LoadProblemBuf(hProb, colCount, rowCount, nonzeroCount, rangeCount, objectSense, _
            '            objectConst, objectCoeffs, lowerBounds, upperBounds, rowType, rhsValues, _
            '            rangeValues, matrixBegin, matrixCount, matrixIndex, matrixValues, _
            '            GenerateNamesBuf(colNames), GenerateNamesBuf(rowNames), objName)

            Return result
        End Function

        Public Function LoadProblem(ByVal hProb As IntPtr, ByVal colCount As Integer, ByVal rowCount As Integer, _
                        ByVal nonZeroCount As Integer, ByVal rangeCount As Integer, ByVal objectSense As ObjectSense, _
                        ByVal objectConst As Double, ByVal objectCoeffs() As Double, ByVal lowerBounds() As Double, _
                        ByVal upperBounds() As Double, ByVal rowLower() As Double, ByVal rowUpper() As Double, _
                        ByVal matrixBegin() As Integer, ByVal matrixCount() As Integer, ByVal matrixIndex() As Integer, _
                        ByVal matrixValues() As Double, ByVal colNames() As String, ByVal rowNames() As String, _
                        ByVal objName As String) As CallResult
            Dim result As CallResult
            result = SetLoadNamesType(hProb, LoadNamesType.Buffer)
            result = LoadProblem(hProb, colCount, rowCount, nonZeroCount, rangeCount, objectSense, _
                        objectConst, objectCoeffs, lowerBounds, upperBounds, Nothing, rowLower, _
                        rowUpper, matrixBegin, matrixCount, matrixIndex, matrixValues, _
                        GenerateNamesBuf(colNames), GenerateNamesBuf(rowNames), objName)
            Return result
        End Function

        'RKP/01-25-10
        'This is a new addition to CoinMP.dll v1.4+
        'The following function has been deprecated:
        'CoinSetLoadNamesType
        'https://projects.coin-or.org/CoinMP/changeset/337
        <DllImport("coinmp.dll", EntryPoint:="CoinLoadProblemBuf")> _
        Public Function LoadProblemBuf(ByVal hProb As Long, _
            ByVal colCount As Long, ByVal rowCount As Long, ByVal nonZeroCount As Long, ByVal rangeCount As Long, _
            ByVal objectSense As Long, ByVal objectConst As Double, ByRef objectCoeffs As Double, _
            ByRef lowerBounds As Double, ByRef upperBounds As Double, ByRef rowType As Byte, _
            ByRef rhsValues As Double, ByRef rangeValues As Double, ByRef matrixBegin As Long, _
            ByRef matrixCount As Long, ByRef matrixIndex As Long, ByRef matrixValues As Double, _
            ByVal colNamesBuf As String, ByVal rowNamesBuf As String, ByVal objName As String) As Integer
        End Function

        'RKP/08-29-12/v4.1.176
        'Added new entry to conform to CoinMP v1.6.

        '[DllImport("coinmp.dll")] public static extern int CoinLoadMatrix(IntPtr hProb, int colCount, int rowCount,
        '        int nzCount, int rangeCount, int objectSense, double objectConst, double[] objectCoeffs,
        '        double[] lowerBounds, double[] upperBounds, char[] rowType, double[] rhsValues,
        '        double[] rangeValues, int[] matrixBegin, int[] matrixCount, int[] matrixIndex,
        '        double[] matrixValues);

        <DllImport("coinmp.dll")> _
        Public Function CoinLoadMatrix(hProb As IntPtr, colCount As Integer, rowCount As Integer, nzCount As Integer, rangeCount As Integer, objectSense As Integer, _
 objectConst As Double, objectCoeffs As Double(), lowerBounds As Double(), upperBounds As Double(), rowType As Char(), rhsValues As Double(), _
 rangeValues As Double(), matrixBegin As Integer(), matrixCount As Integer(), matrixIndex As Integer(), matrixValues As Double()) As Integer
        End Function

        'RKP/08-29-12/v4.1.176
        'Added new entry to conform to CoinMP v1.6.

        <DllImport("coinmp.dll")> _
        Public Function CoinLoadNamesBuf(hProb As IntPtr, colNamesBuf As String, rowNamesBuf As String, objName As String) As Integer
        End Function

        Public Function CoinLoadNames(hProb As IntPtr, colNames As String(), rowNames As String(), objName As String) As Integer
            Return CoinLoadNamesBuf(hProb, GenerateNamesBuf(colNames), GenerateNamesBuf(rowNames), objName)
        End Function

        <DllImport("coinmp.dll")> _
        Public Function CoinLoadProblemBuf(hProb As IntPtr, colCount As Integer, rowCount As Integer, nzCount As Integer, rangeCount As Integer, objectSense As Integer, _
 objectConst As Double, objectCoeffs As Double(), lowerBounds As Double(), upperBounds As Double(), rowType As Char(), rhsValues As Double(), _
 rangeValues As Double(), matrixBegin As Integer(), matrixCount As Integer(), matrixIndex As Integer(), matrixValues As Double(), colNamesBuf As String, _
 rowNamesBuf As String, objName As String) As Integer
        End Function

        Public Function CoinLoadProblem(hProb As IntPtr, colCount As Integer, rowCount As Integer, nzCount As Integer, rangeCount As Integer, objectSense As Integer, _
         objectConst As Double, objectCoeffs As Double(), lowerBounds As Double(), upperBounds As Double(), rowType As Char(), rhsValues As Double(), _
         rangeValues As Double(), matrixBegin As Integer(), matrixCount As Integer(), matrixIndex As Integer(), matrixValues As Double(), colNames As String(), _
         rowNames As String(), objName As String) As Integer
            Return CoinLoadProblemBuf(hProb, colCount, rowCount, nzCount, rangeCount, objectSense, _
             objectConst, objectCoeffs, lowerBounds, upperBounds, rowType, rhsValues, _
             rangeValues, matrixBegin, matrixCount, matrixIndex, matrixValues, GenerateNamesBuf(colNames), _
             GenerateNamesBuf(rowNames), objName)

        End Function

        ' when there is no rowType argument, CoinLoadProblem switches to rowLower and rowUpper arguments
        Public Function CoinLoadProblem(hProb As IntPtr, colCount As Integer, rowCount As Integer, nzCount As Integer, rangeCount As Integer, objectSense As Integer, _
         objectConst As Double, objectCoeffs As Double(), lowerBounds As Double(), upperBounds As Double(), rowLower As Double(), rowUpper As Double(), _
         matrixBegin As Integer(), matrixCount As Integer(), matrixIndex As Integer(), matrixValues As Double(), colNames As String(), rowNames As String(), _
         objName As String) As Integer
            Return CoinLoadProblemBuf(hProb, colCount, rowCount, nzCount, rangeCount, objectSense, _
             objectConst, objectCoeffs, lowerBounds, upperBounds, Nothing, rowLower, _
             rowUpper, matrixBegin, matrixCount, matrixIndex, matrixValues, GenerateNamesBuf(colNames), _
             GenerateNamesBuf(rowNames), objName)

        End Function

        <DllImport("coinmp.dll", EntryPoint:="CoinLoadInitValues")> _
        Public Function LoadInitValues(ByVal hProb As IntPtr, ByVal InitValues() As Double) As CallResult
        End Function

        <DllImport("coinmp.dll", EntryPoint:="CoinLoadInteger")> _
        Public Function LoadInteger(ByVal hProb As IntPtr, ByVal colType() As Char) As CallResult
        End Function

        <DllImport("coinmp.dll", EntryPoint:="CoinLoadPriority")> _
        Public Function LoadPriority(ByVal hProb As IntPtr, ByVal priorCount As Integer, _
                        ByVal priorValues() As Integer, ByVal branchDir() As Integer) As CallResult
        End Function

        <DllImport("coinmp.dll", EntryPoint:="CoinLoadSos")> _
        Public Function LoadSos(ByVal hProb As IntPtr, ByVal sosCount As Integer, _
                        ByVal sosNZCount As Integer, ByVal sosType() As Integer, ByVal sosPrior() As Integer, _
                        ByVal sosBegin() As Integer, ByVal sosIndex() As Integer, ByVal sosRef() As Double) As CallResult
        End Function

        <DllImport("coinmp.dll", EntryPoint:="CoinLoadQuadratic")> _
        Public Function LoadQuadratic(ByVal hProb As IntPtr, ByVal quadBegin() As Integer, _
                        ByVal quadCount() As Integer, ByVal quadIndex() As Integer, ByVal quadValues() As Double) As CallResult
        End Function

        <DllImport("coinmp.dll", EntryPoint:="CoinLoadNonlinear")> _
        Public Function LoadNonlinear(ByVal hProb As IntPtr, ByVal nlpTreeCount As Integer, _
                        ByVal nlpLineCount As Integer, ByVal nlpBegin() As Integer, ByVal nlpOper() As Integer, _
                        ByVal nlpArg1() As Integer, ByVal nlpArg2() As Integer, ByVal nlpIndex1() As Integer, _
                        ByVal nlpIndex2() As Integer, ByVal nlpValue1() As Double, ByVal nlpValue2() As Double) As CallResult
        End Function

        <DllImport("coinmp.dll", EntryPoint:="CoinUnloadProblem")> _
        Public Function UnloadProblem(ByVal hProb As IntPtr) As CallResult
        End Function

        'RKP/01-25-10/v2.3.127
        'This function has been deprecated with CoinMP.dll v1.4+
        'The replacement function is:
        'CoinLoadProblemBuf
        <DllImport("coinmp.dll", EntryPoint:="CoinSetLoadNamesType")> _
        Public Function SetLoadNamesType(ByVal hProb As IntPtr, ByVal loadNamesType As LoadNamesType) As CallResult
        End Function

        <DllImport("coinmp.dll", EntryPoint:="CoinCheckProblem")> _
        Public Function CheckProblem(ByVal hProb As IntPtr) As CheckResult
        End Function


        <DllImport("coinmp.dll", EntryPoint:="CoinGetProblemName")> _
        Public Function GetProblemName(ByVal hProb As IntPtr, ByVal ProblemName As StringBuilder, _
                        ByVal buflen As Integer) As Integer
        End Function

        <DllImport("coinmp.dll", EntryPoint:="CoinGetColCount")> _
        Public Function GetColCount(ByVal hProb As IntPtr) As Integer
        End Function

        <DllImport("coinmp.dll", EntryPoint:="CoinGetRowCount")> _
        Public Function GetRowCount(ByVal hProb As IntPtr) As Integer
        End Function

        <DllImport("coinmp.dll", EntryPoint:="CoinGetColName")> _
        Public Function GetColName(ByVal hProb As IntPtr, ByVal col As Integer, _
                        ByVal colName As StringBuilder, ByVal buflen As Integer) As Integer
        End Function

        <DllImport("coinmp.dll", EntryPoint:="CoinGetRowName")> _
        Public Function GetRowName(ByVal hProb As IntPtr, ByVal row As Integer, _
                        ByVal rowName As StringBuilder, ByVal buflen As Integer) As Integer
        End Function


        <DllImport("coinmp.dll", EntryPoint:="CoinSetMsgLogCallback")> _
        Public Function SetMsgLogCallback(ByVal hProb As IntPtr, _
                        ByVal msgLogDelegate As MsgLogDelegate) As CallResult
        End Function

        <DllImport("coinmp.dll", EntryPoint:="CoinSetIterCallback")> _
        Public Function SetIterCallback(ByVal hProb As IntPtr, _
                        ByVal iterDelegate As IterDelegate) As CallResult
        End Function

        <DllImport("coinmp.dll", EntryPoint:="CoinSetMipNodeCallback")> _
        Public Function SetMipNodeCallback(ByVal hProb As IntPtr, _
                        ByVal mipNodeDelegate As MipNodeDelegate) As CallResult
        End Function


        <DllImport("coinmp.dll", EntryPoint:="CoinOptimizeProblem")> _
        Public Function OptimizeProblem(ByVal hProb As IntPtr, ByVal method As Integer) As CallResult
        End Function

        Public Function OptimizeProblem(ByVal hProb As IntPtr) As CallResult
            Return OptimizeProblem(hProb, MethodType.Auto)
        End Function

        <DllImport("coinmp.dll", EntryPoint:="CoinGetSolutionStatus")> _
        Public Function GetSolutionStatus(ByVal hProb As IntPtr) As Integer
        End Function

        '<DllImport("coinmp.dll", EntryPoint:="CoinGetSolutionText")> _
        'Public Function GetSolutionText(ByVal hProbs As IntPtr) As Integer
        'End Function

        '<DllImport("coinmp.dll", EntryPoint:="CoinGetSolutionText")> _
        'Public Function GetSolutionText(ByVal hProb As IntPtr, ByVal solutionStatus As Integer, _
        '                ByVal solutionText As StringBuilder, ByVal buflen As Integer) As Integer
        'End Function

        'Public Declare Function CoinGetSolutionText Lib "coinmp.dll" (ByVal hProbs As IntPtr) As Integer

        '<DllImport("coinmp.dll")> _
        'Public Function CoinGetSolutionText(ByRef hProbs As IntPtr) As Integer
        'End Function

        '<DllImport("coinmp.dll")> _
        'Public Function CoinGetSolutionText(ByRef hProbs As IntPtr, ByVal solutionStatus As Integer) As String
        'End Function

        <DllImport("coinmp.dll")> _
        Public Function CoinGetSolutionText(ByVal hProbs As IntPtr) As String
        End Function

        'length = CoinMP.GetSolutionText(CType(hProb, IntPtr), solutionStatus, solutionText, solutionText.Capacity)
        <DllImport("coinmp.dll", CallingConvention:=Runtime.InteropServices.CallingConvention.Cdecl, CharSet:=CharSet.Unicode, EntryPoint:="CoinGetSolutionTextBuf")> _
        Public Function GetSolutionText(ByVal hProb As IntPtr, ByVal solutionStatus As Integer, _
                        ByVal solutionText As StringBuilder, ByVal buflen As Integer) As Integer
        End Function

        <DllImport("coinmp.dll", EntryPoint:="CoinGetObjectValue")> _
        Public Function GetObjectValue(ByVal hProb As IntPtr) As Double
        End Function

        <DllImport("coinmp.dll", EntryPoint:="CoinGetMipBestBound")> _
        Public Function GetMipBestBound(ByVal hProb As IntPtr) As Double
        End Function

        <DllImport("coinmp.dll", EntryPoint:="CoinGetIterCount")> _
        Public Function GetIterCount(ByVal hProb As IntPtr) As Integer
        End Function

        <DllImport("coinmp.dll", EntryPoint:="CoinGetMipNodeCount")> _
        Public Function GetMipNodeCount(ByVal hProb As IntPtr) As Integer
        End Function


        <DllImport("coinmp.dll", EntryPoint:="CoinGetSolutionValues")> _
        Public Function GetSolutionValues(ByVal hProb As IntPtr, ByVal activity() As Double, _
                        ByVal reducedCost() As Double, ByVal slackValues() As Double, _
                        ByVal shadowPrice() As Double) As CallResult
        End Function

        <DllImport("coinmp.dll", EntryPoint:="CoinGetSolutionRanges")> _
        Public Function GetSolutionRanges(ByVal hProb As IntPtr, ByVal objLoRange() As Double, _
                        ByVal objUpRange() As Double, ByVal rhsLoRange() As Double, _
                        ByVal rhsUpRange() As Double) As CallResult
        End Function

        <DllImport("coinmp.dll", EntryPoint:="CoinGetSolutionBasis")> _
        Public Function GetSolutionBasis(ByVal hProb As IntPtr, ByVal colStatus() As Integer, _
                        ByVal rowStatus() As Integer) As CallResult
        End Function

        <DllImport("coinmp.dll", EntryPoint:="CoinReadFile")> _
        Public Function ReadFile(ByVal hProb As IntPtr, ByVal fileType As FileType, _
                        ByVal readFilename As String) As CallResult
        End Function

        <DllImport("coinmp.dll", EntryPoint:="CoinWriteFile")> _
        Public Function WriteFile(ByVal hProb As IntPtr, ByVal fileType As FileType, _
                        ByVal writeFilename As String) As CallResult
        End Function

        <DllImport("coinmp.dll", EntryPoint:="CoinOpenLogFile")> _
        Public Function OpenLogFile(ByVal hProb As IntPtr, ByVal logFilename As String) As CallResult
        End Function

        <DllImport("coinmp.dll", EntryPoint:="CoinCloseLogFile")> _
        Public Function CloseLogFile(ByVal hProb As IntPtr) As CallResult
        End Function


        <DllImport("coinmp.dll", EntryPoint:="CoinGetOptionCount")> _
        Public Function GetOptionCount(ByVal hProb As IntPtr) As Integer
        End Function

        <DllImport("coinmp.dll", EntryPoint:="CoinGetOptionInfo")> _
        Public Function GetOptionInfo(ByVal hProb As IntPtr, ByVal optionNr As Integer, _
                        ByRef optionID As OptionParam, ByRef groupType As GroupType, _
                        ByRef optionType As OptionType, ByVal optionName As StringBuilder, _
                        ByVal shortName As StringBuilder, ByVal buflen As Integer) As CallResult
        End Function

        <DllImport("coinmp.dll", EntryPoint:="CoinGetIntOptionMinMax")> _
        Public Function GetIntOptionMinMax(ByVal hProb As IntPtr, ByVal optionNr As Integer, _
                        ByRef minIntValue As Integer, ByRef MaxIntValue As Integer) As CallResult
        End Function

        <DllImport("coinmp.dll", EntryPoint:="CoinGetRealOptionMinMax")> _
        Public Function GetRealOptionMinMax(ByVal hProb As IntPtr, ByVal optionNr As Integer, _
                        ByRef minRealValue As Double, ByRef MaxRealValue As Double) As CallResult
        End Function

        <DllImport("coinmp.dll", EntryPoint:="CoinGetOptionChanged")> _
        Public Function GetOptionChanged(ByVal hProb As IntPtr, ByVal optionID As OptionParam) As Integer
        End Function

        <DllImport("coinmp.dll", EntryPoint:="CoinGetIntOption")> _
        Public Function GetIntOption(ByVal hProb As IntPtr, ByVal optionID As OptionParam) As Integer
        End Function

        <DllImport("coinmp.dll", EntryPoint:="CoinSetIntOption")> _
        Public Function SetIntOption(ByVal hProb As IntPtr, ByVal optionID As OptionParam, _
                        ByVal intValue As Integer) As CallResult
        End Function

        <DllImport("coinmp.dll", EntryPoint:="CoinGetRealOption")> _
        Public Function GetRealOption(ByVal hProb As IntPtr, ByVal optionID As OptionParam) As Double
        End Function

        <DllImport("coinmp.dll", EntryPoint:="CoinSetRealOption")> _
        Public Function SetRealOption(ByVal hProb As IntPtr, ByVal optionID As OptionParam, _
                        ByVal realValue As Double) As CallResult
        End Function

        <DllImport("coinmp.dll", EntryPoint:="CoinGetStringOption")> _
        Public Function GetStringOption(ByVal hProb As IntPtr, ByVal optionID As OptionParam, _
                        ByVal stringValue As StringBuilder, ByVal buflen As Integer) As CallResult
        End Function

        <DllImport("coinmp.dll", EntryPoint:="CoinSetStringOption")> _
        Public Function SetStringOption(ByVal hProb As IntPtr, ByVal optionID As OptionParam, _
                        ByVal stringValue As String) As CallResult
        End Function

        Private Function GenerateNamesBuf(ByVal NamesList() As String) As String
            Dim i As Integer
            Dim count As Integer
            Dim namesBuf As StringBuilder = New StringBuilder()

            count = NamesList.GetLength(0)
            If count > 0 Then
                namesBuf.Append(NamesList(0) + vbNullChar)
                For i = 1 To count - 1
                    namesBuf.Append(NamesList(i) + vbNullChar)
                Next i
            End If
            Return namesBuf.ToString()
        End Function

        Private Function CallingConvention() As CallingConvention
            Throw New NotImplementedException
        End Function


    End Module 'CoinMP

    Module CoinMP16
        'Compatible with CoinMP v1.6 onwards
    End Module

    Public Class SolveProblem

        'Dim logTxt As LogHandler = Nothing
        'Dim logMsg As LogHandler = Nothing
        Private _switches() As String
        Private Shadows _engine As COPT.Engine
        Private _ds As New DataSet

        Public Sub New()
            'logTxt = New LogHandler()
            'LogMsg = New LogHandler()
        End Sub

        'Public Sub New(ByRef txtLog As TextBox)
        '    logTxt = New LogHandler(txtLog)
        '    LogMsg = New LogHandler()
        'End Sub

        'Public Sub New(ByRef txtLog As TextBox, ByRef msgLog As TextBox)
        '    logTxt = New LogHandler(txtLog)
        '    LogMsg = New LogHandler(msgLog)
        'End Sub


        Private Function MsgLogCallback(ByVal msg As String) As Integer
            'logMsg.WriteLine("***" & msg)
            Return CallResult.Success
        End Function

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="solverCoinMP"></param>
        ''' <param name="engine"></param>
        ''' <param name="dtRow"></param>
        ''' <param name="dtCol"></param>
        ''' <param name="dtMtx"></param>
        ''' <param name="switches"></param>
        ''' <param name="problemName"></param>
        ''' <param name="optimalValue"></param>
        ''' <param name="colCount"></param>
        ''' <param name="rowCount"></param>
        ''' <param name="nonZeroCount"></param>
        ''' <param name="rangeCount"></param>
        ''' <param name="objectSense"></param>
        ''' <param name="objectConst"></param>
        ''' <param name="objectCoeffs"></param>
        ''' <param name="lowerBounds"></param>
        ''' <param name="upperBounds"></param>
        ''' <param name="rowType"></param>
        ''' <param name="rhsValues"></param>
        ''' <param name="rangeValues"></param>
        ''' <param name="matrixBegin"></param>
        ''' <param name="matrixCount"></param>
        ''' <param name="matrixIndex"></param>
        ''' <param name="matrixValues"></param>
        ''' <param name="colNames"></param>
        ''' <param name="rowNames"></param>
        ''' <param name="objectName"></param>
        ''' <param name="initValues"></param>
        ''' <param name="colType"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' RKP/06-21-12/v4.0.170
        ''' Created to incorporate MIP stopping conditions.
        ''' </remarks>
        'Public Function RunV3( _
        '    ByRef solverCoinMP As Solver_CoinMP, _
        '    ByRef engine As COPT.Engine, _
        '    ByRef dtRow As DataTable, _
        '    ByRef dtCol As DataTable, _
        '    ByRef dtMtx As DataTable, _
        '    ByRef switches() As String, _
        '    ByVal problemName As String, ByVal optimalValue As Double, _
        '    ByVal colCount As Integer, ByVal rowCount As Integer, ByVal nonZeroCount As Integer, _
        '    ByVal rangeCount As Integer, ByVal objectSense As Integer, ByVal objectConst As Double, _
        '    ByRef objectCoeffs() As Double, ByRef lowerBounds() As Double, ByRef upperBounds() As Double, _
        '    ByRef rowType() As Char, ByRef rhsValues() As Double, ByRef rangeValues() As Double, _
        '    ByRef matrixBegin() As Integer, ByRef matrixCount() As Integer, ByRef matrixIndex() As Integer, _
        '    ByRef matrixValues() As Double, ByRef colNames() As String, ByRef rowNames() As String, _
        '    ByVal objectName As String, ByRef initValues() As Double, ByRef colType() As Char _
        ') As Boolean

        '    'code.google.com/p/pulp-or/source/browse/trunk/pulp-or/src/solvers.py?spec=svn286&r=286

        '    Dim result As CallResult

        '    Dim hProb As Integer
        '    Dim solutionStatus As Integer
        '    Dim objectValue As Double
        '    Dim objectValueMIP As Double
        '    Dim length As Integer
        '    Dim i As Integer
        '    'Dim value As Double

        '    Dim activity(colCount) As Double
        '    Dim reduced(colCount) As Double
        '    Dim slack(rowCount) As Double
        '    Dim shadow(rowCount) As Double
        '    'Dim activity() As Double
        '    'Dim reduced() As Double
        '    'Dim slack() As Double
        '    'Dim shadow() As Double

        '    Dim solutionText As StringBuilder = New StringBuilder(100)
        '    Dim colName As StringBuilder = New StringBuilder(100)
        '    Dim rowName As StringBuilder = New StringBuilder(100)

        '    Dim MsgLogDelegate As CoinMP.MsgLogDelegate

        '    Dim tickCountStart As Integer = My.Computer.Clock.TickCount
        '    Dim success As Boolean = False 'RKP/04-18-12/v3.2.166

        '    'logTxt.NewLine()
        '    'logTxt.WriteLine("Solve Problem" & problemName)
        '    'logTxt.WriteLine("---------------------------------------------------------------")

        '    _switches = switches

        '    EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CoinMP - Loading...")


        '    dtMtx = Nothing

        '    'RKP/04-18-12/v3.2.166
        '    GenUtils.CollectGarbage()

        '    Try
        '        hProb = CInt(CoinMP.CreateProblem(problemName))
        '    Catch ex As System.Exception
        '        'EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solve - CoinMP - Unable to ""CoinMP.CreateProblem""" & vbNewLine & ex.Message)
        '        GenUtils.Message(GenUtils.MsgType.Critical, "C-OPT Engine-Solve-CoinMP-Run", "Unable to ""CoinMP.CreateProblem""" & vbNewLine & ex.Message & vbNewLine & "APM: " & GenUtils.GetAvailablePhysicalMemoryStr())
        '        success = False
        '    End Try


        '    'If self.mip Then
        '    '        CoinMP.CoinSetRealOption(hProb, self.COIN_REAL_MIPMAXSEC,
        '    '                              ctypes.c_double(self.maxSeconds))
        '    'Else
        '    '        CoinMP.CoinSetRealOption(hProb, self.COIN_REAL_MAXSECONDS,
        '    '                              ctypes.c_double(self.maxSeconds))


        '    ''result = CoinMP.LoadProblem(hProb, colCount, rowCount, nonZeroCount, rangeCount, _
        '    ''        objectSense, objectConst, objectCoeffs, lowerBounds, upperBounds, _
        '    ''        rowType, rhsValues, rangeValues, matrixBegin, matrixCount, _
        '    ''        matrixIndex, matrixValues, colNames, rowNames, objectName)

        '    Try
        '        result = CoinMP.LoadProblem(hProb, colCount, rowCount, nonZeroCount, rangeCount, _
        '            objectSense, objectConst, objectCoeffs, lowerBounds, upperBounds, _
        '            rowType, rhsValues, rangeValues, matrixBegin, matrixCount, _
        '            matrixIndex, matrixValues, colNames, rowNames, objectName)
        '    Catch ex As System.Exception
        '        GenUtils.Message(GenUtils.MsgType.Critical, "C-OPT Engine-Solve-CoinMP-Run", "Unable to ""CoinMP.LoadProblem""" & vbNewLine & ex.Message & vbNewLine & "APM: " & GenUtils.GetAvailablePhysicalMemoryStr())
        '        success = False
        '    End Try


        '    'result = CoinMP.LoadProblemBuf(hProb, colCount, rowCount, nonZeroCount, rangeCount, objectSense, objectConst, objectCoeffs, lowerBounds, upperBounds, rowType, rhsValues, rangeValues, matrixBegin, matrixCount, matrixIndex, matrixValues, colNames, rowNames, objectName)

        '    If result = CoinMP.CallResult.Failed Then
        '        'logTxt.WriteLine("CoinLoadProblem failed")
        '        Console.WriteLine("CoinLoadProblem failed")
        '        EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CoinMP - CoinLoadProblem failed.")
        '        success = False
        '    Else
        '        Console.WriteLine("CoinMP.LoadProblem finished successfully.")
        '        success = True
        '    End If
        '    If solverCoinMP.isMIP Then
        '        Console.WriteLine("Solver - CoinMP - Problem type = Integer")
        '        Try
        '            result = CoinMP.LoadInteger(CType(hProb, IntPtr), colType)
        '        Catch ex As System.Exception
        '            GenUtils.Message(GenUtils.MsgType.Critical, "C-OPT Engine-Solve-CoinMP-Run", "Unable to ""CoinMP.LoadInteger""" & vbNewLine & ex.Message & vbNewLine & "APM: " & GenUtils.GetAvailablePhysicalMemoryStr())
        '            success = False
        '        End Try

        '        If result = CoinMP.CallResult.Failed Then
        '            'logTxt.WriteLine("CoinLoadInteger failed")
        '            Console.WriteLine("CoinLoadInteger failed")
        '            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CoinMP - CoinLoadInteger failed.")
        '            success = False
        '        Else
        '            Console.WriteLine("Solver - CoinMP - CoinMP.LoadInteger finished successfully.")
        '            success = True
        '        End If
        '    Else
        '        Console.WriteLine("Solver - CoinMP - Problem type = Continuous")
        '    End If
        '    'If Not colType Is Nothing Then

        '    'End If
        '    Try
        '        result = CType(CoinMP.CheckProblem(CType(hProb, IntPtr)), CallResult)
        '    Catch ex As System.Exception
        '        GenUtils.Message(GenUtils.MsgType.Critical, "C-OPT Engine-Solve-CoinMP-Run", "Unable to ""CoinMP.CheckProblem""" & vbNewLine & ex.Message & vbNewLine & "APM: " & GenUtils.GetAvailablePhysicalMemoryStr())
        '        success = False
        '    End Try

        '    If result = CoinMP.CallResult.Failed Then
        '        'logTxt.WriteLine("Check Problem failed (result = " & result & ")")
        '        'MsgBox("Check Problem failed (result = " & result & ")")
        '        Console.WriteLine("Check Problem failed (result = " & result & ")")
        '        EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CoinMP - Check Problem failed (result = " & result & ")")
        '        success = False
        '    Else
        '        Console.WriteLine("Solver - CoinMP - CoinMP.CheckProblem finished successfully.")
        '        success = True
        '    End If

        '    'Dim MsgLogDelegate As CoinMP.MsgLogDelegate = New CoinMP.MsgLogDelegate(AddressOf MsgLogCallback)
        '    Try
        '        MsgLogDelegate = New CoinMP.MsgLogDelegate(AddressOf MsgLogCallback)
        '        success = True
        '    Catch ex As System.Exception
        '        Console.WriteLine("Solver - CoinMP - Failed to instantiate 'MsgLogDelegate'.")
        '        EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CoinMP - Failed to instantiate 'MsgLogDelegate'.")
        '        success = False
        '    End Try

        '    Try
        '        result = CoinMP.SetMsgLogCallback(CType(hProb, IntPtr), MsgLogDelegate)
        '    Catch ex As System.Exception
        '        Console.WriteLine("Solver - CoinMP - Failed to 'SetMsgLogCallback'.")
        '        EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CoinMP - Failed to 'SetMsgLogCallback'.")
        '        success = False
        '    End Try


        '    'If MsgBox("Do you want to write MPS file (before optimizing)?", CType(MsgBoxStyle.OkCancel + MsgBoxStyle.Question, MsgBoxStyle)) = MsgBoxResult.Ok Then
        '    '    result = CoinMP.WriteFile(CType(hProb, IntPtr), CoinMP.FileType.MPS, My.Application.Info.DirectoryPath & "\" & problemName + ".mps")
        '    'End If
        '    If GenUtils.IsSwitchAvailable(_switches, "/GenSolverMPS") Then
        '        result = CoinMP.WriteFile(CType(hProb, IntPtr), CoinMP.FileType.MPS, GenUtils.GetSwitchArgument(_switches, "/WorkDir", 1) & "\" & problemName + ".mps")
        '    End If

        '    GenUtils.CollectGarbage()

        '    tickCountStart = My.Computer.Clock.TickCount

        '    Try
        '        result = CoinMP.OptimizeProblem(CType(hProb, IntPtr))
        '        Console.WriteLine("Solver - CoinMP - CoinMP.OptimizeProblem finished successfully (" & GenUtils.FormatTime(tickCountStart, My.Computer.Clock.TickCount) & ")")
        '        EntLib.COPT.Log.Log(GenUtils.GetSwitchArgument(_switches, "/WorkDir", 1), "C-OPT Engine - CoinMP - Solve took: ", GenUtils.FormatTime(tickCountStart, My.Computer.Clock.TickCount))
        '        success = True
        '    Catch ex As System.Exception
        '        'MessageBox.Show(ex.Message)
        '        EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CoinMP - OptimizeProblem failed - " & ex.Message)
        '        GenUtils.Message(GenUtils.MsgType.Information, "Engine - Solver-CoinMP", ex.Message)
        '        success = False
        '    End Try

        '    GenUtils.CollectGarbage()

        '    'logMsg.WriteLine("---------------------------------------------------------------")
        '    'logTxt.WriteLine("Solve time = " & ((My.Computer.Clock.TickCount - tickCountStart) / 1000).ToString() & " seconds.")

        '    'Ravi/RKP - Commented due to memory violation error
        '    'result = CoinMP.WriteFile(CType(hProb, IntPtr), CoinMP.FileType.MPS, problemName + ".mps")



        '    solutionStatus = CoinMP.GetSolutionStatus(CType(hProb, IntPtr))
        '    Console.WriteLine("Solver - CoinMP - Solution Status = " & solutionStatus)


        '    'engine.SolutionStatus = CoinMP.GetSolutionText(hProb, solutionStatus, solutionText, solutionText.Capacity)
        '    'length = engine.SolutionStatus = CoinMP.GetSolutionText(CType(hProb, IntPtr))
        '    engine.SolutionStatus = CoinMP.CoinGetSolutionText(hProb)

        '    objectValue = CoinMP.GetObjectValue(CType(hProb, IntPtr))
        '    objectValueMIP = objectValue

        '    '0 = LpStatusOptimal
        '    '1 = LpStatusInfeasible
        '    '2 = LpStatusInfeasible
        '    '3 = LpStatusNotSolved
        '    '4 = LpStatusNotSolved
        '    '5 = LpStatusNotSolved
        '    '-1= LpStatusUndefined
        '    'Select Case solutionStatus
        '    '    Case -1
        '    '        engine.SolutionStatus = "Undefined Error"
        '    '    Case 0
        '    '        engine.SolutionStatus = "Optimal Solution"
        '    '    Case 1
        '    '        engine.SolutionStatus = "Infeasible Solution"
        '    '    Case 2
        '    '        engine.SolutionStatus = "Infeasible Solution"
        '    '    Case 3
        '    '        engine.SolutionStatus = "Solution Not Solved"
        '    '    Case 4
        '    '        engine.SolutionStatus = "Solution Not Solved"
        '    '    Case 5
        '    '        engine.SolutionStatus = "Solution Not Solved"
        '    '    Case Else
        '    '        engine.SolutionStatus = "Undefined Error"
        '    'End Select
        '    engine.SolutionStatus = getSolutionStatus(solutionStatus)
        '    engine.SolutionStatusCode = solutionStatus

        '    If engine.SolutionStatus.Contains("OPTIMAL") Then
        '        engine.CommonSolutionStatus = "OPTIMAL SOLUTION"
        '        engine.CommonSolutionStatusCode = Solver.commonSolutionStatus.statusOptimal
        '    Else
        '        engine.CommonSolutionStatus = "INFEASIBLE SOLUTION"
        '        engine.CommonSolutionStatusCode = Solver.commonSolutionStatus.statusInfeasible
        '    End If

        '    Console.ForegroundColor = ConsoleColor.DarkCyan
        '    Debug.Print("Solver has finished optimizing.")
        '    Console.WriteLine("Solver has finished optimizing.")
        '    Console.WriteLine("Solution status = " & engine.SolutionStatus)
        '    Console.ForegroundColor = ConsoleColor.DarkYellow

        '    'logTxt.WriteLine("---------------------------------------------------------------")

        '    'logTxt.WriteLine("Problem Name:    " + problemName)
        '    'logTxt.WriteLine("Solution Result: " + solutionText.ToString())
        '    'logTxt.WriteLine("Solution Status: " + solutionStatus.ToString())
        '    'logTxt.WriteLine("Optimal Value:   " + objectValue.ToString() + " (" + optimalValue.ToString() + ")")
        '    'logTxt.WriteLine("---------------------------------------------------------------")

        '    result = CoinMP.GetSolutionValues(CType(hProb, IntPtr), activity, reduced, slack, shadow)
        '    'value = CoinMP.GetMipBestBound(CType(hProb, IntPtr))
        '    'i = CoinMP.GetMipNodeCount(CType(hProb, IntPtr))

        '    'RKP/02-23-10/v2.3.131
        '    'Need to find a way to the count for MIP (Binary) problems.
        '    engine.SolutionIterations = CoinMP.GetIterCount(hProb)
        '    'engine.SolutionIterations = CoinMP.GetMipNodeCount(hProb)
        '    'RKP/02-15-10/v2.3.130
        '    '
        '    'Don't save the resulting MIP solution yet.
        '    'Re-optimize the problem as Continuous, with the following changes:
        '    'LO(column) = activity(column)
        '    'UP(column) = activity(column)
        '    'This process should now yield DJ(column), ACTIVITY(row) and SHADOW(row).
        '    'Now go ahead and save the problem back to the database.
        '    '/NoContinuous
        '    If solverCoinMP.isMIP Then
        '        If Not GenUtils.IsSwitchAvailable(switches, "/NoContinuous") Then
        '            If solutionStatus = 0 Then 'Proceed only if optimal
        '                'Unload the current problem
        '                result = CoinMP.UnloadProblem(CType(hProb, IntPtr))
        '                CoinMP.FreeSolver()
        '                'Create a new problem
        '                hProb = CInt(CoinMP.CreateProblem(problemName))
        '                'Readjust LO and UP to activity() values from the MIP run.
        '                For i = 0 To activity.Length - 2
        '                    'Only change for columns that are not continuous.
        '                    If colType(i) <> "C" Then
        '                        lowerBounds(i) = activity(i)
        '                        upperBounds(i) = activity(i)
        '                    End If
        '                Next
        '                result = CoinMP.LoadProblem(hProb, colCount, rowCount, nonZeroCount, rangeCount, _
        '                    objectSense, objectConst, objectCoeffs, lowerBounds, upperBounds, _
        '                    rowType, rhsValues, rangeValues, matrixBegin, matrixCount, _
        '                    matrixIndex, matrixValues, colNames, rowNames, objectName)
        '                If result = CoinMP.CallResult.Failed Then
        '                    Console.WriteLine("CoinLoadProblem failed")
        '                    EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CoinMP - CoinLoadProblem failed.")
        '                End If
        '                result = CType(CoinMP.CheckProblem(CType(hProb, IntPtr)), CallResult)
        '                If result = CoinMP.CallResult.Failed Then
        '                    Console.WriteLine("Check Problem failed (result = " & result & ")")
        '                    EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CoinMP - Check Problem failed (result = " & result & ")")
        '                End If
        '                MsgLogDelegate = New CoinMP.MsgLogDelegate(AddressOf MsgLogCallback)
        '                result = CoinMP.SetMsgLogCallback(CType(hProb, IntPtr), MsgLogDelegate)
        '                result = CoinMP.OptimizeProblem(CType(hProb, IntPtr))
        '                Try
        '                    EntLib.COPT.Log.Log(GenUtils.GetSwitchArgument(_switches, "/WorkDir", 1), "C-OPT Engine - CoinMP - Solve took: ", GenUtils.FormatTime(tickCountStart, My.Computer.Clock.TickCount))
        '                Catch ex As System.Exception
        '                    EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CoinMP - OptimizeProblem failed - " & ex.Message)
        '                    GenUtils.Message(GenUtils.MsgType.Information, "Engine - Solver-CoinMP", ex.Message)
        '                End Try
        '                solutionStatus = CoinMP.GetSolutionStatus(CType(hProb, IntPtr))
        '                'length = CoinMP.GetSolutionText(hProb, solutionStatus, solutionText, solutionText.Capacity)
        '                'length = CoinMP.GetSolutionText(CType(hProb, IntPtr))
        '                engine.SolutionStatus = CoinMP.CoinGetSolutionText(hProb)
        '                objectValue = CoinMP.GetObjectValue(CType(hProb, IntPtr))
        '                If engine.SolutionStatus <> getSolutionStatus(solutionStatus) Then
        '                    engine.SolutionStatus = engine.SolutionStatus & " / " & getSolutionStatus(solutionStatus)
        '                End If
        '                result = CoinMP.GetSolutionValues(CType(hProb, IntPtr), activity, reduced, slack, shadow)
        '            End If 'If solutionStatus = 0 Then
        '        End If
        '    End If

        '    If engine Is Nothing Then
        '        engine = New COPT.Engine
        '    End If

        '    'engine.SolutionStatus = solutionStatus.ToString() 'solutionText.ToString().Replace("=", "-")
        '    'If solutionStatus = 0 Then
        '    '    engine.SolutionStatus = "Optimal Solution"
        '    'Else
        '    '    engine.SolutionStatus = "Error"
        '    'End If

        '    engine.SolutionRows = CoinMP.GetRowCount(hProb)
        '    engine.SolutionColumns = CoinMP.GetColCount(hProb)
        '    engine.SolutionNonZeros = nonZeroCount
        '    If solverCoinMP.isMIP Then
        '        engine.SolutionObj = objectValueMIP
        '    Else
        '        engine.SolutionObj = objectValue
        '    End If

        '    'engine.SolverName = "CoinMP" '.Replace("=", "-")
        '    If Not solverCoinMP.isMIP Then
        '        engine.SolutionIterations = CoinMP.GetIterCount(hProb)
        '    End If

        '    length = CoinMP.GetSolverName(solutionText, solutionText.Capacity)

        '    'RKP/09-16-09
        '    'Not working
        '    'engine.SolverName = solutionText.ToString().Replace("=", "-") '& CoinMP.GetVersion
        '    engine.SolverName = "CoinMP" '(v" & CoinMP.GetVersion.ToString() & ")"
        '    engine.SolverVersion = CoinMP.GetVersion.ToString() & " (" & IIf(GenUtils.Is64Bit, "64-bit", "32-bit") & ")"

        '    'Ravi/RKP/04-20-09 - Temporarily commented
        '    'For i = 0 To colCount - 1
        '    '    'If activity(i) <> 0.0 Then
        '    '    length = CoinMP.GetColName(hProb, i, colName, colName.Capacity)
        '    '    logTxt.WriteLine(colName.ToString() & " = " & activity(i).ToString() & ", reduced = " & reduced(i).ToString())
        '    '    'End If
        '    'Next i
        '    'For i = 0 To rowCount - 1
        '    '    'If slack(i) <> 0.0 Then
        '    '    length = CoinMP.GetRowName(hProb, i, rowName, rowName.Capacity)
        '    '    logTxt.WriteLine(rowName.ToString() & " = " & slack(i).ToString() & ", shadow = " & shadow(i).ToString())
        '    '    'End If
        '    'Next i

        '    'engine.SolutionStatus = engine.getSolutionStatusDescription(solutionStatus)


        '    If solutionStatus = 0 Then '0 = Optimal solution (returned by CoinMP)

        '        success = True

        '        '0 = LpStatusOptimal
        '        '1 = LpStatusInfeasible
        '        '2 = LpStatusInfeasible
        '        '3 = LpStatusNotSolved
        '        '4 = LpStatusNotSolved
        '        '5 = LpStatusNotSolved
        '        '-1= LpStatusUndefined


        '        'Dim dtCol As New DataTable
        '        ''dtActivity.LoadDataRow(activity, True)
        '        ''dtActivity.LoadDataRow(DirectCast(activity, Object), LoadOption.OverwriteChanges)
        '        'dtCol.Columns.Add("ColID", System.Type.GetType("System.Int32"))
        '        'dtCol.Columns.Add("Activity", System.Type.GetType("System.Double"))
        '        'dtCol.Columns.Add("DJ", System.Type.GetType("System.Double"))
        '        'For i = 0 To colCount - 1
        '        '    dtCol.Rows.Add(i, activity(i), reduced(i))
        '        'Next

        '        'Dim dtRow As New DataTable
        '        'dtRow.Columns.Add("RowID", System.Type.GetType("System.Int32"))
        '        'dtRow.Columns.Add("Activity", System.Type.GetType("System.Double"))
        '        'dtRow.Columns.Add("Shadow", System.Type.GetType("System.Double"))
        '        'For i = 0 To rowCount - 1
        '        '    dtRow.Rows.Add(i, slack(i), shadow(i))
        '        'Next

        '        Debug.Print("C-OPT Engine - Solver - CoinMP - Importing solution...")
        '        Console.WriteLine("C-OPT Engine - Solver - CoinMP - Importing solution...")

        '        'GenUtils.SerializeArrayDouble(engine.GetWorkDir, "COL.ACTIVITY.txt", activity)
        '        'GenUtils.SerializeArrayDouble(engine.GetWorkDir, "COL.REDUCED.txt", reduced)
        '        'GenUtils.SerializeArrayDouble(engine.GetWorkDir, "ROW.SLACK.txt", slack)
        '        'GenUtils.SerializeArrayDouble(engine.GetWorkDir, "ROW.SHADOW.txt", shadow)

        '        'db_Update(dtCol, dtRow)
        '        'For i = 0 To ds.Tables("tsysCol").Rows.Count - 1
        '        '    ds.Tables("tsysCol").Rows(i).Item("ACTIVITY") = activity(i) 'xa.getColumnPrimalActivity(ds.Tables("tsysCol").Rows(i).Item("COL").ToString())
        '        '    ds.Tables("tsysCol").Rows(i).Item("DJ") = reduced(i) 'xa.getColumnDualActivity(ds.Tables("tsysCol").Rows(i).Item("COL").ToString())
        '        'Next

        '        engine.Progress = "Updating solver results to in-memory column table...Start - " & My.Computer.Clock.LocalTime.ToString()
        '        Debug.Print(engine.Progress)
        '        Console.WriteLine(engine.Progress)
        '        EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CoinMP", engine.Progress)

        '        For i = 0 To dtCol.Rows.Count - 1
        '            dtCol.Rows(i).Item("ACTIVITY") = activity(i) 'xa.getColumnPrimalActivity(ds.Tables("tsysCol").Rows(i).Item("COL").ToString())
        '            dtCol.Rows(i).Item("DJ") = reduced(i) 'xa.getColumnDualActivity(ds.Tables("tsysCol").Rows(i).Item("COL").ToString())
        '            'dtCol.Rows(i).Item("STATUS") = ""
        '        Next

        '        engine.Progress = "Updating solver results to in-memory column table...End - " & My.Computer.Clock.LocalTime.ToString()
        '        Debug.Print(engine.Progress)
        '        Console.WriteLine(engine.Progress)
        '        EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CoinMP", engine.Progress)

        '        'MyBase.Engine.CurrentDb.UpdateDataSet(dtCol, "SELECT * FROM " & col)
        '        'For i = 0 To ds.Tables("tsysRow").Rows.Count - 1
        '        '    ds.Tables("tsysRow").Rows(i).Item("ACTIVITY") = slack(i) 'xa.getRowPrimalActivity(ds.Tables("tsysRow").Rows(i).Item("ROW").ToString())
        '        '    ds.Tables("tsysRow").Rows(i).Item("SHADOW") = shadow(i) 'xa.getRowDualActivity(ds.Tables("tsysRow").Rows(i).Item("ROW").ToString())
        '        'Next

        '        engine.Progress = "Updating solver results to in-memory row table...Start - " & My.Computer.Clock.LocalTime.ToString()
        '        Debug.Print(engine.Progress)
        '        Console.WriteLine(engine.Progress)
        '        EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CoinMP", engine.Progress)

        '        For i = 0 To dtRow.Rows.Count - 1
        '            dtRow.Rows(i).Item("ACTIVITY") = slack(i) 'xa.getRowPrimalActivity(ds.Tables("tsysRow").Rows(i).Item("ROW").ToString())
        '            dtRow.Rows(i).Item("SHADOW") = shadow(i) 'xa.getRowDualActivity(ds.Tables("tsysRow").Rows(i).Item("ROW").ToString())
        '            'dtRow.Rows(i).Item("STATUS") = ""
        '        Next

        '        engine.Progress = "Updating solver results to in-memory row table...End - " & My.Computer.Clock.LocalTime.ToString()
        '        Debug.Print(engine.Progress)
        '        Console.WriteLine(engine.Progress)
        '        EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CoinMP", engine.Progress)

        '        'If GenUtils.IsSwitchAvailable(_switches, "/UseMinSysRes") Then
        '        '    'If GenUtils.IsSwitchAvailable(_switches, "/UseMSAccessSyntax") Then
        '        '    'End If
        '        '    dtCol.TableName = "dtCol"
        '        '    GenUtils.SerializeDataTable(engine.GetWorkDir, "dtCol.xml", dtCol)
        '        '    dtRow.TableName = "dtRow"
        '        '    GenUtils.SerializeDataTable(engine.GetWorkDir, "dtRow.xml", dtRow)
        '        'End If
        '    Else
        '        success = False
        '    End If

        '    'logTxt.WriteLine("---------------------------------------------------------------")
        '    'logTxt.NewLine()

        '    result = CoinMP.UnloadProblem(CType(hProb, IntPtr))
        '    CoinMP.FreeSolver()

        '    EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CoinMP - Finished.")

        '    'Console.WriteLine("C-OPT Engine - Solver - CoinMP - Importing solution...done.")
        '    'Console.WriteLine("")
        '    Return success
        'End Function

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="solverCoinMP"></param>
        ''' <param name="engine"></param>
        ''' <param name="dtRow"></param>
        ''' <param name="dtCol"></param>
        ''' <param name="dtMtx"></param>
        ''' <param name="switches"></param>
        ''' <param name="problemName"></param>
        ''' <param name="optimalValue"></param>
        ''' <param name="colCount"></param>
        ''' <param name="rowCount"></param>
        ''' <param name="nonZeroCount"></param>
        ''' <param name="rangeCount"></param>
        ''' <param name="objectSense"></param>
        ''' <param name="objectConst"></param>
        ''' <param name="objectCoeffs"></param>
        ''' <param name="lowerBounds"></param>
        ''' <param name="upperBounds"></param>
        ''' <param name="rowType"></param>
        ''' <param name="rhsValues"></param>
        ''' <param name="rangeValues"></param>
        ''' <param name="matrixBegin"></param>
        ''' <param name="matrixCount"></param>
        ''' <param name="matrixIndex"></param>
        ''' <param name="matrixValues"></param>
        ''' <param name="colNames"></param>
        ''' <param name="rowNames"></param>
        ''' <param name="objectName"></param>
        ''' <param name="initValues"></param>
        ''' <param name="colType"></param>
        ''' <remarks>
        ''' This is the main "solve" function.
        ''' CoinMP.Solve
        ''' </remarks>

        Public Function Run( _
            ByRef solverCoinMP As Solver_CoinMP, _
            ByRef engine As COPT.Engine, _
            ByRef dtRow As DataTable, _
            ByRef dtCol As DataTable, _
            ByRef dtMtx As DataTable, _
            ByRef switches() As String, _
            ByVal problemName As String, ByVal optimalValue As Double, _
            ByVal colCount As Long, ByVal rowCount As Long, ByVal nonZeroCount As Long, _
            ByVal rangeCount As Long, ByVal objectSense As Long, ByVal objectConst As Double, _
            ByRef objectCoeffs() As Double, ByRef lowerBounds() As Double, ByRef upperBounds() As Double, _
            ByRef rowType() As Char, ByRef rhsValues() As Double, ByRef rangeValues() As Double, _
            ByRef matrixBegin() As Integer, ByRef matrixCount() As Integer, ByRef matrixIndex() As Integer, _
            ByRef matrixValues() As Double, ByRef colNames() As String, ByRef rowNames() As String, _
            ByVal objectName As String, ByRef initValues() As Double, ByRef colType() As Char _
        ) As Boolean

            'code.google.com/p/pulp-or/source/browse/trunk/pulp-or/src/solvers.py?spec=svn286&r=286

            Dim result As CallResult

            Dim hProb As IntPtr = IntPtr.Zero 'Long
            Dim solutionStatus As Integer
            Dim objectValue As Double
            Dim objectValueMIP As Double
            Dim length As Integer
            Dim i As Integer
            'Dim value As Double

            Dim activity(colCount) As Double
            Dim reduced(colCount) As Double
            Dim slack(rowCount) As Double
            Dim shadow(rowCount) As Double
            'Dim activity() As Double
            'Dim reduced() As Double
            'Dim slack() As Double
            'Dim shadow() As Double

            Dim solutionText As StringBuilder = New StringBuilder(100)
            Dim colName As StringBuilder = New StringBuilder(100)
            Dim rowName As StringBuilder = New StringBuilder(100)

            Dim MsgLogDelegate As CoinMP.MsgLogDelegate

            Dim tickCountStart As Integer = My.Computer.Clock.TickCount
            Dim success As Boolean = False 'RKP/04-18-12/v3.2.166
            Dim switchOptionGap As Boolean = GenUtils.IsSwitchAvailable(switches, "/SetRelativeGapTolerance")
            Dim switchOptionTime As Boolean = GenUtils.IsSwitchAvailable(switches, "/SetTimeLimit")

            'logTxt.NewLine()
            'logTxt.WriteLine("Solve Problem" & problemName)
            'logTxt.WriteLine("---------------------------------------------------------------")

            _switches = switches

            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CoinMP - Loading...")

            dtMtx = Nothing

            'RKP/04-18-12/v3.2.166
            GenUtils.CollectGarbage()

            'result = CoinMP.CoinInitSolver("")
            result = CoinMP.InitSolver("")

            Try
                hProb = CoinMP.CreateProblem(problemName)
            Catch ex As System.Exception
                'EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solve - CoinMP - Unable to ""CoinMP.CreateProblem""" & vbNewLine & ex.Message)
                GenUtils.Message(GenUtils.MsgType.Critical, "C-OPT Engine-Solve-CoinMP-Run", "Unable to ""CoinMP.CreateProblem""" & vbNewLine & ex.Message & vbNewLine & "APM: " & GenUtils.GetAvailablePhysicalMemoryStr())
                success = False
            End Try


            'If self.mip Then
            '        CoinMP.CoinSetRealOption(hProb, self.COIN_REAL_MIPMAXSEC,
            '                              ctypes.c_double(self.maxSeconds))
            'Else
            '        CoinMP.CoinSetRealOption(hProb, self.COIN_REAL_MAXSECONDS,
            '                              ctypes.c_double(self.maxSeconds))


            ''result = CoinMP.LoadProblem(hProb, colCount, rowCount, nonZeroCount, rangeCount, _
            ''        objectSense, objectConst, objectCoeffs, lowerBounds, upperBounds, _
            ''        rowType, rhsValues, rangeValues, matrixBegin, matrixCount, _
            ''        matrixIndex, matrixValues, colNames, rowNames, objectName)

            'RKP/08-29-12/v4.1.176
            'Used by CoinMP v1.4
            'Try
            '    result = CoinMP.LoadProblem(hProb, colCount, rowCount, nonZeroCount, rangeCount, _
            '        objectSense, objectConst, objectCoeffs, lowerBounds, upperBounds, _
            '        rowType, rhsValues, rangeValues, matrixBegin, matrixCount, _
            '        matrixIndex, matrixValues, colNames, rowNames, objectName)
            'Catch ex As System.Exception
            '    GenUtils.Message(GenUtils.MsgType.Critical, "C-OPT Engine-Solve-CoinMP-Run", "Unable to ""CoinMP.LoadProblem""" & vbNewLine & ex.Message & vbNewLine & "APM: " & GenUtils.GetAvailablePhysicalMemoryStr())
            '    success = False
            'End Try

            'RKP/08-29-12/v4.1.176
            'Used by CoinMP v1.6
            Try

                'result = CoinMP.LoadMatrix(hProb, colCount, rowCount, nonZeroCount, rangeCount, _
                '    objectSense, objectConst, objectCoeffs, lowerBounds, upperBounds, _
                '    rowType, rhsValues, rangeValues, matrixBegin, matrixCount, _
                '    matrixIndex, matrixValues)

                result = CoinMP.CoinLoadMatrix(hProb, colCount, rowCount, nonZeroCount, rangeCount,
                                objectSense, objectConst, objectCoeffs, lowerBounds, upperBounds,
                                rowType, rhsValues, rangeValues, matrixBegin, matrixCount,
                                matrixIndex, matrixValues)

                result = CoinMP.CoinLoadNames(hProb, colNames, rowNames, objectName)

                'result = CoinMP.LoadProblem(hProb, colCount, rowCount, nonZeroCount, rangeCount, _
                '    objectSense, objectConst, objectCoeffs, lowerBounds, upperBounds, _
                '    rowType, rhsValues, rangeValues, matrixBegin, matrixCount, _
                '    matrixIndex, matrixValues, colNames, rowNames, objectName)
            Catch ex As System.Exception
                GenUtils.Message(GenUtils.MsgType.Critical, "C-OPT Engine-Solve-CoinMP-Run", "Unable to ""CoinMP.LoadProblem""" & vbNewLine & ex.Message & vbNewLine & "APM: " & GenUtils.GetAvailablePhysicalMemoryStr())
                success = False
            End Try


            'result = CoinMP.LoadProblemBuf(hProb, colCount, rowCount, nonZeroCount, rangeCount, objectSense, objectConst, objectCoeffs, lowerBounds, upperBounds, rowType, rhsValues, rangeValues, matrixBegin, matrixCount, matrixIndex, matrixValues, colNames, rowNames, objectName)

            If result = CoinMP.CallResult.Failed Then
                'logTxt.WriteLine("CoinLoadProblem failed")
                Console.WriteLine("CoinLoadProblem failed")
                EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CoinMP - CoinLoadProblem failed.")
                success = False
            Else
                Console.WriteLine("CoinMP.LoadProblem finished successfully.")
                success = True
            End If
            If solverCoinMP.isMIP Then
                Console.WriteLine("Solver - CoinMP - Problem type = Integer")
                Try
                    result = CoinMP.LoadInteger(hProb, colType)
                Catch ex As System.Exception
                    GenUtils.Message(GenUtils.MsgType.Critical, "C-OPT Engine-Solve-CoinMP-Run", "Unable to ""CoinMP.LoadInteger""" & vbNewLine & ex.Message & vbNewLine & "APM: " & GenUtils.GetAvailablePhysicalMemoryStr())
                    success = False
                End Try

                If result = CoinMP.CallResult.Failed Then
                    'logTxt.WriteLine("CoinLoadInteger failed")
                    Console.WriteLine("CoinLoadInteger failed")
                    EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CoinMP - CoinLoadInteger failed.")
                    success = False
                Else
                    Console.WriteLine("Solver - CoinMP - CoinMP.LoadInteger finished successfully.")
                    success = True
                End If
            Else
                Console.WriteLine("Solver - CoinMP - Problem type = Continuous")
            End If
            'If Not colType Is Nothing Then

            'End If
            Try
                'result = CType(CoinMP.CheckProblem(CType(hProb, IntPtr)), CallResult)
                result = CoinMP.CheckProblem(hProb)
            Catch ex As System.Exception
                GenUtils.Message(GenUtils.MsgType.Critical, "C-OPT Engine-Solve-CoinMP-Run", "Unable to ""CoinMP.CheckProblem""" & vbNewLine & ex.Message & vbNewLine & "APM: " & GenUtils.GetAvailablePhysicalMemoryStr())
                success = False
            End Try

            If result = CoinMP.CallResult.Failed Then
                'logTxt.WriteLine("Check Problem failed (result = " & result & ")")
                'MsgBox("Check Problem failed (result = " & result & ")")
                Console.WriteLine("Check Problem failed (result = " & result & ")")
                EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CoinMP - Check Problem failed (result = " & result & ")")
                success = False
            Else
                Console.WriteLine("Solver - CoinMP - CoinMP.CheckProblem finished successfully.")
                success = True
            End If

            engine.SolverVersion = CoinMP.GetVersion.ToString() & " (" & IIf(GenUtils.Is64Bit, "64-bit", "32-bit") & ")"

            'Dim MsgLogDelegate As CoinMP.MsgLogDelegate = New CoinMP.MsgLogDelegate(AddressOf MsgLogCallback)
            MsgLogDelegate = New CoinMP.MsgLogDelegate(AddressOf MsgLogCallback)
            result = CoinMP.SetMsgLogCallback(hProb, MsgLogDelegate)

            'If MsgBox("Do you want to write MPS file (before optimizing)?", CType(MsgBoxStyle.OkCancel + MsgBoxStyle.Question, MsgBoxStyle)) = MsgBoxResult.Ok Then
            '    result = CoinMP.WriteFile(CType(hProb, IntPtr), CoinMP.FileType.MPS, My.Application.Info.DirectoryPath & "\" & problemName + ".mps")
            'End If
            If GenUtils.IsSwitchAvailable(_switches, "/GenSolverMPS") Then
                result = CoinMP.WriteFile(hProb, CoinMP.FileType.MPS, GenUtils.GetSwitchArgument(_switches, "/WorkDir", 1) & "\" & problemName + ".mps")
            End If

            GenUtils.CollectGarbage()

            'RKP/09-20-12/v4.2.180
            'Adding code to apply MIP Stopping Criteria
            'MipMaxSol
            'MipMaxSec
            'MipIntTol
            If solverCoinMP.isMIP Then
                If GenUtils.IsSwitchAvailable(_switches, "/SetSolutionLimit") Then
                    Try '
                        result = CoinMP.SetIntOption(hProb, OptionParam.MipMaxSol, GenUtils.GetSwitchArgument(_switches, "/SetSolutionLimit", 1))
                        success = True
                    Catch ex As System.Exception
                        EntLib.COPT.Log.Log(GenUtils.GetWorkDir(switches), "C-OPT Engine", "Solver - CoinMP - /SetSolutionLimit failed - " & ex.Message)
                        GenUtils.Message(GenUtils.MsgType.Information, "Engine - Solver-CoinMP", ex.Message)
                        success = False
                    End Try
                    If result = 0 Then
                        success = True
                    Else
                        EntLib.COPT.Log.Log(GenUtils.GetWorkDir(switches), "C-OPT Engine", "Solver - CoinMP - /SetSolutionLimit failed - " & "CallResult = " & result.ToString())
                        GenUtils.Message(GenUtils.MsgType.Information, "Engine - Solver-CoinMP", "Solver - CoinMP - /SetSolutionLimit failed - " & "CallResult = " & result.ToString())
                        success = False
                    End If
                End If

                If success Then
                    If switchOptionTime Then
                        If GenUtils.IsSwitchAvailable(switches, "/SetTimeLimit") Then
                            Try '
                                result = CoinMP.SetRealOption(hProb, OptionParam.MipMaxSec, GenUtils.GetSwitchArgument(switches, "/SetTimeLimit", 1)) 'this sets the min time as max
                                success = True
                            Catch ex As System.Exception
                                EntLib.COPT.Log.Log(GenUtils.GetWorkDir(switches), "C-OPT Engine", "Solver - CoinMP - /SetTimeLimit failed - " & ex.Message)
                                GenUtils.Message(GenUtils.MsgType.Information, "Engine - Solver-CoinMP", ex.Message)
                                success = False
                            End Try
                            If result = 0 Then
                                success = True
                            Else
                                EntLib.COPT.Log.Log(GenUtils.GetWorkDir(switches), "C-OPT Engine", "Solver - CoinMP - /SetTimeLimit failed - " & "CallResult = " & result)
                                GenUtils.Message(GenUtils.MsgType.Information, "Engine - Solver-CoinMP", "Solver - CoinMP - /SetTimeLimit failed - " & "CallResult = " & result)
                                success = False
                            End If
                        End If
                    Else
                        If success Then
                            If GenUtils.IsSwitchAvailable(switches, "/SetRelativeGapTolerance") Then
                                Try '
                                    'result = CoinMP.SetRealOption(hProb, OptionParam.MipAbsGap, GenUtils.GetSwitchArgument(_switches, "/SetRelativeGapTolerance", 1)) 'this sets the min time as max
                                    result = CoinMP.SetRealOption(hProb, OptionParam.MipFracGap, GenUtils.GetSwitchArgument(_switches, "/SetRelativeGapTolerance", 1)) 'this sets the min time as max
                                    success = True
                                Catch ex As System.Exception
                                    EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CoinMP - /SetRelativeGapTolerance failed - " & ex.Message)
                                    GenUtils.Message(GenUtils.MsgType.Information, "Engine - Solver-CoinMP", ex.Message)
                                    success = False
                                End Try
                                If result = 0 Then
                                    success = True
                                Else
                                    EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CoinMP - /SetRelativeGapTolerance failed - " & "CallResult = " & result)
                                    GenUtils.Message(GenUtils.MsgType.Information, "Engine - Solver-CoinMP", "Solver - CoinMP - /SetRelativeGapTolerance failed - " & "CallResult = " & result)
                                    success = False
                                End If
                            End If 'If GenUtils.IsSwitchAvailable(switches, "/SetRelativeGapTolerance") Then
                        End If
                    End If
                End If
            End If

            Console.WriteLine("Solver - CoinMP is now optimizing...")

            tickCountStart = My.Computer.Clock.TickCount

            Try
                result = CoinMP.OptimizeProblem(hProb)
                Console.WriteLine("Solver - CoinMP - CoinMP.OptimizeProblem finished successfully (" & GenUtils.FormatTime(tickCountStart, My.Computer.Clock.TickCount) & ")")
                EntLib.COPT.Log.Log(GenUtils.GetSwitchArgument(_switches, "/WorkDir", 1), "C-OPT Engine - CoinMP - Solve took: ", GenUtils.FormatTime(tickCountStart, My.Computer.Clock.TickCount))
                success = True
            Catch ex As System.Exception
                'MessageBox.Show(ex.Message)
                EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CoinMP - OptimizeProblem failed - " & ex.Message)
                GenUtils.Message(GenUtils.MsgType.Information, "Engine - Solver-CoinMP", ex.Message)
                success = False
            End Try

            GenUtils.CollectGarbage()

            'logMsg.WriteLine("---------------------------------------------------------------")
            'logTxt.WriteLine("Solve time = " & ((My.Computer.Clock.TickCount - tickCountStart) / 1000).ToString() & " seconds.")

            'Ravi/RKP - Commented due to memory violation error
            'result = CoinMP.WriteFile(CType(hProb, IntPtr), CoinMP.FileType.MPS, problemName + ".mps")



            solutionStatus = CoinMP.GetSolutionStatus(hProb)
            Console.WriteLine("Solver - CoinMP - Solution Status = " & solutionStatus)

            'length = CoinMP.GetSolutionText(CType(hProb, IntPtr), solutionStatus, solutionText, solutionText.Capacity)
            Try
                i = CoinMP.GetSolutionText(hProb, solutionStatus, solutionText, 100)
            Catch ex As System.Exception
                'engine.SolutionStatus = CoinMP.CoinGetSolutionText(hProb)
                engine.SolutionStatus = Me.getSolutionStatus(solutionStatus)
            End Try

            objectValue = CoinMP.GetObjectValue(hProb)
            objectValueMIP = objectValue

            If GenUtils.IsSwitchAvailable(_switches, "/Sense") Then
                If GenUtils.GetSwitchArgument(_switches, "/Sense", 1).Trim().ToUpper() = "MIN" Then
                    objectValue = objectValue * CDbl(GenUtils.GetSwitchArgument(_switches, "/Sense", 2))
                    objectValueMIP = objectValue '* CDbl(GenUtils.GetSwitchArgument(_switches, "/Sense", 2))
                End If
            End If

            'Stage 2 Optimization - for MIP only
            If solverCoinMP.isMIP Then
                If solutionStatus <> 0 Then
                    If switchOptionGap AndAlso switchOptionTime Then
                        Try '
                            result = CoinMP.SetRealOption(hProb, OptionParam.MipMaxSec, GenUtils.GetSwitchArgument(_switches, "/SetTimeLimit", 2) - GenUtils.GetSwitchArgument(_switches, "/SetTimeLimit", 1)) 'this sets the max time as max
                            success = True
                        Catch ex As System.Exception
                            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CoinMP - /SetTimeLimit (Max-Min) failed - " & ex.Message)
                            GenUtils.Message(GenUtils.MsgType.Information, "Engine - Solver-CoinMP", ex.Message)
                            success = False
                        End Try
                        If result = 0 Then
                            success = True
                        Else
                            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CoinMP - /SetTimeLimit (Max-Min) failed - " & "CallResult = " & result)
                            GenUtils.Message(GenUtils.MsgType.Information, "Engine - Solver-CoinMP", "Solver - CoinMP - /SetTimeLimit (Max-Min) failed - " & "CallResult = " & result)
                            success = False
                        End If

                        If success Then
                            If GenUtils.IsSwitchAvailable(_switches, "/SetRelativeGapTolerance") Then
                                Try '
                                    'result = CoinMP.SetRealOption(hProb, OptionParam.MipAbsGap, GenUtils.GetSwitchArgument(_switches, "/SetRelativeGapTolerance", 1)) 'this sets the min time as max
                                    'MipFracGap
                                    result = CoinMP.SetRealOption(hProb, OptionParam.MipFracGap, GenUtils.GetSwitchArgument(_switches, "/SetRelativeGapTolerance", 1)) 'this sets the min time as max
                                    success = True
                                Catch ex As System.Exception
                                    EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CoinMP - /SetRelativeGapTolerance failed - " & ex.Message)
                                    GenUtils.Message(GenUtils.MsgType.Information, "Engine - Solver-CoinMP", ex.Message)
                                    success = False
                                End Try
                                If result = 0 Then
                                    success = True
                                Else
                                    EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CoinMP - /SetRelativeGapTolerance failed - " & "CallResult = " & result)
                                    GenUtils.Message(GenUtils.MsgType.Information, "Engine - Solver-CoinMP", "Solver - CoinMP - /SetRelativeGapTolerance failed - " & "CallResult = " & result)
                                    success = False
                                End If
                            End If
                        End If

                        If success Then
                            tickCountStart = My.Computer.Clock.TickCount

                            Try
                                result = CoinMP.OptimizeProblem(hProb)
                                Console.WriteLine("Solver - CoinMP - CoinMP.OptimizeProblem finished successfully (" & GenUtils.FormatTime(tickCountStart, My.Computer.Clock.TickCount) & ")")
                                EntLib.COPT.Log.Log(GenUtils.GetSwitchArgument(_switches, "/WorkDir", 1), "C-OPT Engine - CoinMP - Solve took: ", GenUtils.FormatTime(tickCountStart, My.Computer.Clock.TickCount))
                                success = True
                            Catch ex As System.Exception
                                'MessageBox.Show(ex.Message)
                                EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CoinMP - OptimizeProblem failed - " & ex.Message)
                                GenUtils.Message(GenUtils.MsgType.Information, "Engine - Solver-CoinMP", ex.Message)
                                success = False
                            End Try

                            GenUtils.CollectGarbage()

                            'logMsg.WriteLine("---------------------------------------------------------------")
                            'logTxt.WriteLine("Solve time = " & ((My.Computer.Clock.TickCount - tickCountStart) / 1000).ToString() & " seconds.")

                            'Ravi/RKP - Commented due to memory violation error
                            'result = CoinMP.WriteFile(CType(hProb, IntPtr), CoinMP.FileType.MPS, problemName + ".mps")



                            solutionStatus = CoinMP.GetSolutionStatus(hProb)
                            Console.WriteLine("Solver - CoinMP - Solution Status = " & solutionStatus)


                            'length = CoinMP.GetSolutionText(hProb, solutionStatus, solutionText, solutionText.Capacity)
                            'length = CoinMP.GetSolutionText(hProb)
                            Try
                                i = CoinMP.GetSolutionText(hProb, solutionStatus, solutionText, 100)
                            Catch ex As System.Exception
                                engine.SolutionStatus = Me.getSolutionStatus(solutionStatus)
                            End Try

                            objectValue = CoinMP.GetObjectValue(hProb)
                            objectValueMIP = objectValue

                            If GenUtils.IsSwitchAvailable(_switches, "/Sense") Then
                                If GenUtils.GetSwitchArgument(_switches, "/Sense", 1).Trim().ToUpper() = "MIN" Then
                                    objectValue = objectValue * CDbl(GenUtils.GetSwitchArgument(_switches, "/Sense", 2))
                                    objectValueMIP = objectValue '* CDbl(GenUtils.GetSwitchArgument(_switches, "/Sense", 2))
                                End If
                            End If
                        End If
                    End If
                End If
            End If

            '0 = LpStatusOptimal
            '1 = LpStatusInfeasible
            '2 = LpStatusInfeasible
            '3 = LpStatusNotSolved
            '4 = LpStatusNotSolved
            '5 = LpStatusNotSolved
            '-1= LpStatusUndefined
            'Select Case solutionStatus
            '    Case -1
            '        engine.SolutionStatus = "Undefined Error"
            '    Case 0
            '        engine.SolutionStatus = "Optimal Solution"
            '    Case 1
            '        engine.SolutionStatus = "Infeasible Solution"
            '    Case 2
            '        engine.SolutionStatus = "Infeasible Solution"
            '    Case 3
            '        engine.SolutionStatus = "Solution Not Solved"
            '    Case 4
            '        engine.SolutionStatus = "Solution Not Solved"
            '    Case 5
            '        engine.SolutionStatus = "Solution Not Solved"
            '    Case Else
            '        engine.SolutionStatus = "Undefined Error"
            'End Select
            engine.SolutionStatus = getSolutionStatus(solutionStatus)
            engine.SolutionStatusCode = solutionStatus

            If engine.SolutionStatus.Contains("OPTIMAL") Then
                engine.CommonSolutionStatus = "OPTIMAL SOLUTION"
                engine.CommonSolutionStatusCode = Solver.commonSolutionStatus.statusOptimal
            Else
                engine.CommonSolutionStatus = "INFEASIBLE SOLUTION"
                engine.CommonSolutionStatusCode = Solver.commonSolutionStatus.statusInfeasible
            End If

            Console.ForegroundColor = ConsoleColor.DarkCyan
            Debug.Print("Solver has finished optimizing.")
            Console.WriteLine("Solver has finished optimizing.")
            Console.WriteLine("Solution status = " & engine.SolutionStatus)
            Console.ForegroundColor = ConsoleColor.DarkYellow

            'logTxt.WriteLine("---------------------------------------------------------------")

            'logTxt.WriteLine("Problem Name:    " + problemName)
            'logTxt.WriteLine("Solution Result: " + solutionText.ToString())
            'logTxt.WriteLine("Solution Status: " + solutionStatus.ToString())
            'logTxt.WriteLine("Optimal Value:   " + objectValue.ToString() + " (" + optimalValue.ToString() + ")")
            'logTxt.WriteLine("---------------------------------------------------------------")

            result = CoinMP.GetSolutionValues(hProb, activity, reduced, slack, shadow)
            'value = CoinMP.GetMipBestBound(CType(hProb, IntPtr))
            'i = CoinMP.GetMipNodeCount(CType(hProb, IntPtr))

            'RKP/02-23-10/v2.3.131
            'Need to find a way to the count for MIP (Binary) problems.
            engine.SolutionIterations = CoinMP.GetIterCount(hProb)
            'engine.SolutionIterations = CoinMP.GetMipNodeCount(hProb)
            'RKP/02-15-10/v2.3.130

            'Don't save the resulting MIP solution yet.
            'Re-optimize the problem as Continuous, with the following changes:
            'LO(column) = activity(column)
            'UP(column) = activity(column)
            'This process should now yield DJ(column), ACTIVITY(row) and SHADOW(row).
            'Now go ahead and save the problem back to the database.
            '/NoContinuous
            If solverCoinMP.isMIP Then
                If Not GenUtils.IsSwitchAvailable(switches, "/NoContinuous") Then
                    If solutionStatus = 0 Then 'Proceed only if optimal

                        For i = 0 To activity.Length - 2
                            If activity(i) > upperBounds(i) OrElse activity(i) < lowerBounds(i) Then
                                Console.WriteLine("Index = " & i & ", Activity = " & activity(i) & ", Lower = " & lowerBounds(i) & ", Upper = " & upperBounds(i))
                            End If
                        Next

                        'Unload the current problem
                        result = CoinMP.UnloadProblem(hProb)
                        CoinMP.FreeSolver()
                        'Create a new problem
                        hProb = CoinMP.CreateProblem(problemName)

                        'RKP/08-28-12/v4.0.175
                        If GenUtils.IsSwitchAvailable(_switches, "/AddBuffer") Then

                            'For i = 0 To activity.Length - 1
                            '    If lowerBounds(i) > 0 Then
                            '        lowerBounds(i) = Math.Round(lowerBounds(i), 8) + 0.00001
                            '    Else
                            '        lowerBounds(i) = Math.Round(lowerBounds(i), 8)
                            '    End If
                            '    upperBounds(i) = Math.Round(upperBounds(i), 8) + 0.0001
                            'Next

                            'Readjust LO and UP to activity() values from the MIP run.
                            For i = 0 To activity.Length - 2
                                'Only change for columns that are not continuous.
                                If colType(i) <> "C" Then
                                    If activity(i) > 0 Then
                                        lowerBounds(i) = Math.Round(activity(i), 8) - 0.00001
                                    Else
                                        lowerBounds(i) = Math.Round(activity(i), 8)
                                    End If
                                    'lowerBounds(i) = Math.Round((activity(i)), 8)
                                    upperBounds(i) = Math.Round(activity(i), 8) + 0.0001

                                    'lowerBounds(i) = activity(i)
                                    'upperBounds(i) = activity(i)
                                Else
                                    If lowerBounds(i) > 0 Then
                                        lowerBounds(i) = Math.Round(lowerBounds(i), 8) - 0.00001
                                    Else
                                        lowerBounds(i) = Math.Round(lowerBounds(i), 8)
                                    End If
                                    upperBounds(i) = Math.Round(upperBounds(i), 8) + 0.0001
                                End If
                            Next
                        Else
                            'Readjust LO and UP to activity() values from the MIP run.
                            For i = 0 To activity.Length - 2
                                'Only change for columns that are not continuous.
                                If colType(i) <> "C" Then
                                    lowerBounds(i) = Math.Round((activity(i)), 8)
                                    upperBounds(i) = Math.Round(activity(i), 8)
                                End If
                            Next
                        End If

                        'RKP/08-29-12/v4.1.176
                        'Deprecated due to upgrade of CoinMP from v1.4 to v1.6
                        'result = CoinMP.LoadProblem(hProb, colCount, rowCount, nonZeroCount, rangeCount, _
                        '    objectSense, objectConst, objectCoeffs, lowerBounds, upperBounds, _
                        '    rowType, rhsValues, rangeValues, matrixBegin, matrixCount, _
                        '    matrixIndex, matrixValues, colNames, rowNames, objectName)

                        'RKP/08-29-12/v4.1.176
                        'New API to conform to CoinMP v1.6

                        result = CoinMP.CoinLoadMatrix(hProb, colCount, rowCount, nonZeroCount, rangeCount,
                                        objectSense, objectConst, objectCoeffs, lowerBounds, upperBounds,
                                        rowType, rhsValues, rangeValues, matrixBegin, matrixCount,
                                        matrixIndex, matrixValues)

                        result = CoinMP.CoinLoadNames(hProb, colNames, rowNames, objectName)

                        If result = CoinMP.CallResult.Failed Then
                            Console.WriteLine("CoinLoadProblem failed")
                            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CoinMP - CoinLoadProblem failed.")
                        End If
                        result = CoinMP.CheckProblem(hProb)
                        If result = CoinMP.CallResult.Failed Then
                            Console.WriteLine("Check Problem failed (result = " & result & ")")
                            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CoinMP - Check Problem failed (result = " & result & ")")
                        End If
                        MsgLogDelegate = New CoinMP.MsgLogDelegate(AddressOf MsgLogCallback)
                        result = CoinMP.SetMsgLogCallback(hProb, MsgLogDelegate)
                        result = CoinMP.OptimizeProblem(hProb)
                        Try
                            EntLib.COPT.Log.Log(GenUtils.GetSwitchArgument(_switches, "/WorkDir", 1), "C-OPT Engine - CoinMP - Solve took: ", GenUtils.FormatTime(tickCountStart, My.Computer.Clock.TickCount))
                        Catch ex As System.Exception
                            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CoinMP - OptimizeProblem failed - " & ex.Message)
                            GenUtils.Message(GenUtils.MsgType.Information, "Engine - Solver-CoinMP", ex.Message)
                        End Try
                        solutionStatus = CoinMP.GetSolutionStatus(hProb)
                        'length = CoinMP.GetSolutionText(hProb, solutionStatus, solutionText, solutionText.Capacity)
                        'length = CoinMP.GetSolutionText(hProb)
                        Try
                            i = CoinMP.GetSolutionText(hProb, solutionStatus, solutionText, 100)
                        Catch ex As System.Exception
                            engine.SolutionStatus = Me.getSolutionStatus(solutionStatus)
                        End Try
                        objectValue = CoinMP.GetObjectValue(hProb)

                        If GenUtils.IsSwitchAvailable(_switches, "/Sense") Then
                            If GenUtils.GetSwitchArgument(_switches, "/Sense", 1).Trim().ToUpper() = "MIN" Then
                                objectValue = objectValue * CDbl(GenUtils.GetSwitchArgument(_switches, "/Sense", 2))
                                objectValueMIP = objectValue '* CDbl(GenUtils.GetSwitchArgument(_switches, "/Sense", 2))
                            End If
                        End If

                        If engine.SolutionStatus <> getSolutionStatus(solutionStatus) Then
                            engine.SolutionStatus = engine.SolutionStatus & " / " & getSolutionStatus(solutionStatus)
                            If IsNumeric(engine.SolutionStatus) Then
                                engine.SolutionStatus = getSolutionStatus(solutionStatus)
                            Else
                                engine.SolutionStatus = engine.SolutionStatus & " (Stage 1-MIP)/ " & getSolutionStatus(solutionStatus) & " (Stage 2-Continuous)"
                            End If
                        End If
                        result = CoinMP.GetSolutionValues(hProb, activity, reduced, slack, shadow)


                        'If GenUtils.IsSwitchAvailable(_switches, "/Override") Then
                        '    solutionStatus = 0
                        'End If

                    End If 'If solutionStatus = 0 Then
                End If
            End If

            If engine Is Nothing Then
                engine = New COPT.Engine
            End If

            'engine.SolutionStatus = solutionStatus.ToString() 'solutionText.ToString().Replace("=", "-")
            'If solutionStatus = 0 Then
            '    engine.SolutionStatus = "Optimal Solution"
            'Else
            '    engine.SolutionStatus = "Error"
            'End If

            engine.SolutionRows = CoinMP.GetRowCount(hProb)
            engine.SolutionColumns = CoinMP.GetColCount(hProb)
            engine.SolutionNonZeros = nonZeroCount
            If solverCoinMP.isMIP Then
                engine.SolutionObj = objectValueMIP
            Else
                engine.SolutionObj = objectValue
            End If

            'engine.SolverName = "CoinMP" '.Replace("=", "-")
            If Not solverCoinMP.isMIP Then
                engine.SolutionIterations = CoinMP.GetIterCount(hProb)
            End If

            length = CoinMP.GetSolverName(solutionText, solutionText.Capacity)

            'RKP/09-16-09
            'Not working
            'engine.SolverName = solutionText.ToString().Replace("=", "-") '& CoinMP.GetVersion
            engine.SolverName = "CoinMP" '(v" & CoinMP.GetVersion.ToString() & ")"
            engine.SolverVersion = CoinMP.GetVersion.ToString() & " (" & IIf(GenUtils.Is64Bit, "64-bit", "32-bit") & ")"

            'Ravi/RKP/04-20-09 - Temporarily commented
            'For i = 0 To colCount - 1
            '    'If activity(i) <> 0.0 Then
            '    length = CoinMP.GetColName(hProb, i, colName, colName.Capacity)
            '    logTxt.WriteLine(colName.ToString() & " = " & activity(i).ToString() & ", reduced = " & reduced(i).ToString())
            '    'End If
            'Next i
            'For i = 0 To rowCount - 1
            '    'If slack(i) <> 0.0 Then
            '    length = CoinMP.GetRowName(hProb, i, rowName, rowName.Capacity)
            '    logTxt.WriteLine(rowName.ToString() & " = " & slack(i).ToString() & ", shadow = " & shadow(i).ToString())
            '    'End If
            'Next i

            'engine.SolutionStatus = engine.getSolutionStatusDescription(solutionStatus)


            If solutionStatus = 0 Then '0 = Optimal solution (returned by CoinMP)

                success = True

                '0 = LpStatusOptimal
                '1 = LpStatusInfeasible
                '2 = LpStatusInfeasible
                '3 = LpStatusNotSolved
                '4 = LpStatusNotSolved
                '5 = LpStatusNotSolved
                '-1= LpStatusUndefined


                'Dim dtCol As New DataTable
                ''dtActivity.LoadDataRow(activity, True)
                ''dtActivity.LoadDataRow(DirectCast(activity, Object), LoadOption.OverwriteChanges)
                'dtCol.Columns.Add("ColID", System.Type.GetType("System.Int32"))
                'dtCol.Columns.Add("Activity", System.Type.GetType("System.Double"))
                'dtCol.Columns.Add("DJ", System.Type.GetType("System.Double"))
                'For i = 0 To colCount - 1
                '    dtCol.Rows.Add(i, activity(i), reduced(i))
                'Next

                'Dim dtRow As New DataTable
                'dtRow.Columns.Add("RowID", System.Type.GetType("System.Int32"))
                'dtRow.Columns.Add("Activity", System.Type.GetType("System.Double"))
                'dtRow.Columns.Add("Shadow", System.Type.GetType("System.Double"))
                'For i = 0 To rowCount - 1
                '    dtRow.Rows.Add(i, slack(i), shadow(i))
                'Next

                Debug.Print("C-OPT Engine - Solver - CoinMP - Importing solution...")
                Console.WriteLine("C-OPT Engine - Solver - CoinMP - Importing solution...")

                'GenUtils.SerializeArrayDouble(engine.GetWorkDir, "COL.ACTIVITY.txt", activity)
                'GenUtils.SerializeArrayDouble(engine.GetWorkDir, "COL.REDUCED.txt", reduced)
                'GenUtils.SerializeArrayDouble(engine.GetWorkDir, "ROW.SLACK.txt", slack)
                'GenUtils.SerializeArrayDouble(engine.GetWorkDir, "ROW.SHADOW.txt", shadow)

                'db_Update(dtCol, dtRow)
                'For i = 0 To ds.Tables("tsysCol").Rows.Count - 1
                '    ds.Tables("tsysCol").Rows(i).Item("ACTIVITY") = activity(i) 'xa.getColumnPrimalActivity(ds.Tables("tsysCol").Rows(i).Item("COL").ToString())
                '    ds.Tables("tsysCol").Rows(i).Item("DJ") = reduced(i) 'xa.getColumnDualActivity(ds.Tables("tsysCol").Rows(i).Item("COL").ToString())
                'Next

                engine.Progress = "Updating solver results to in-memory column table...Start - " & My.Computer.Clock.LocalTime.ToString()
                Debug.Print(engine.Progress)
                Console.WriteLine(engine.Progress)
                EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CoinMP", engine.Progress)

                For i = 0 To dtCol.Rows.Count - 1
                    dtCol.Rows(i).Item("ACTIVITY") = activity(i) 'xa.getColumnPrimalActivity(ds.Tables("tsysCol").Rows(i).Item("COL").ToString())
                    dtCol.Rows(i).Item("DJ") = reduced(i) 'xa.getColumnDualActivity(ds.Tables("tsysCol").Rows(i).Item("COL").ToString())
                    'dtCol.Rows(i).Item("STATUS") = ""
                Next

                engine.Progress = "Updating solver results to in-memory column table...End - " & My.Computer.Clock.LocalTime.ToString()
                Debug.Print(engine.Progress)
                Console.WriteLine(engine.Progress)
                EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CoinMP", engine.Progress)

                'MyBase.Engine.CurrentDb.UpdateDataSet(dtCol, "SELECT * FROM " & col)
                'For i = 0 To ds.Tables("tsysRow").Rows.Count - 1
                '    ds.Tables("tsysRow").Rows(i).Item("ACTIVITY") = slack(i) 'xa.getRowPrimalActivity(ds.Tables("tsysRow").Rows(i).Item("ROW").ToString())
                '    ds.Tables("tsysRow").Rows(i).Item("SHADOW") = shadow(i) 'xa.getRowDualActivity(ds.Tables("tsysRow").Rows(i).Item("ROW").ToString())
                'Next

                engine.Progress = "Updating solver results to in-memory row table...Start - " & My.Computer.Clock.LocalTime.ToString()
                Debug.Print(engine.Progress)
                Console.WriteLine(engine.Progress)
                EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CoinMP", engine.Progress)

                For i = 0 To dtRow.Rows.Count - 1
                    dtRow.Rows(i).Item("ACTIVITY") = slack(i) 'xa.getRowPrimalActivity(ds.Tables("tsysRow").Rows(i).Item("ROW").ToString())
                    dtRow.Rows(i).Item("SHADOW") = shadow(i) 'xa.getRowDualActivity(ds.Tables("tsysRow").Rows(i).Item("ROW").ToString())
                    'dtRow.Rows(i).Item("STATUS") = ""
                Next

                engine.Progress = "Updating solver results to in-memory row table...End - " & My.Computer.Clock.LocalTime.ToString()
                Debug.Print(engine.Progress)
                Console.WriteLine(engine.Progress)
                EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "Solver-CoinMP", engine.Progress)

                'If GenUtils.IsSwitchAvailable(_switches, "/UseMinSysRes") Then
                '    'If GenUtils.IsSwitchAvailable(_switches, "/UseMSAccessSyntax") Then
                '    'End If
                '    dtCol.TableName = "dtCol"
                '    GenUtils.SerializeDataTable(engine.GetWorkDir, "dtCol.xml", dtCol)
                '    dtRow.TableName = "dtRow"
                '    GenUtils.SerializeDataTable(engine.GetWorkDir, "dtRow.xml", dtRow)
                'End If
            Else
                success = False
            End If

            'logTxt.WriteLine("---------------------------------------------------------------")
            'logTxt.NewLine()

            result = CoinMP.UnloadProblem(hProb)
            CoinMP.FreeSolver()

            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - CoinMP - Finished.")

            'Console.WriteLine("C-OPT Engine - Solver - CoinMP - Importing solution...done.")
            'Console.WriteLine("")
            Return success
        End Function

        Public Function Run( _
            ByRef solverCoinMP As Solver_CoinMP, _
            ByRef dtRow As DataTable, ByRef dtCol As DataTable, ByRef dtMtx As DataTable, ByVal switches() As String, ByVal problemName As String, ByVal optimalValue As Double, _
            ByVal colCount As Integer, ByVal rowCount As Integer, ByVal nonZeroCount As Integer, _
            ByVal rangeCount As Integer, ByVal objectSense As Integer, ByVal objectConst As Double, _
            ByVal objectCoeffs() As Double, ByVal lowerBounds() As Double, ByVal upperBounds() As Double, _
            ByVal rowLower() As Double, ByVal rowUpper() As Double, _
            ByVal matrixBegin() As Integer, ByVal matrixCount() As Integer, ByVal matrixIndex() As Integer, _
            ByVal matrixValues() As Double, ByVal colNames() As String, ByVal rowNames() As String, _
            ByVal objectName As String, ByVal initValues() As Double, ByVal colType() As Char _
        ) As Boolean

            Return Run(solverCoinMP, _engine, dtRow, dtCol, dtMtx, switches, problemName, optimalValue, colCount, rowCount, nonZeroCount, rangeCount, _
                objectSense, objectConst, objectCoeffs, lowerBounds, upperBounds, Nothing, _
                rowLower, rowUpper, matrixBegin, matrixCount, matrixIndex, matrixValues, _
                 colNames, rowNames, objectName, initValues, colType)
        End Function

        Private Function getSolutionStatus(ByVal solutionStatus As Integer) As String
            'Select Case solutionStatus
            '    Case -1
            '        Return "Undefined Error".ToUpper() & " (" & solutionStatus & ")"
            '    Case 0
            '        Return "Optimal Solution".ToUpper() & " (" & solutionStatus & ")"
            '    Case 1
            '        Return "Infeasible Solution".ToUpper() & " (" & solutionStatus & ")"
            '    Case 2
            '        Return "Infeasible Solution".ToUpper() & " (" & solutionStatus & ")"
            '    Case 3
            '        Return "Solution Not Solved".ToUpper() & " (" & solutionStatus & ")"
            '    Case 4
            '        Return "Solution Not Solved".ToUpper() & " (" & solutionStatus & ")"
            '    Case 5
            '        Return "Solution Not Solved".ToUpper() & " (" & solutionStatus & ")"
            '    Case Else
            '        Return "Undefined Error".ToUpper() & " (" & solutionStatus & ")"
            'End Select

            Select Case solutionStatus
                Case 0
                    Return "Optimal solution".ToUpper() & " (" & solutionStatus & ")"
                Case 1
                    Return "Problem primal infeasible".ToUpper() & " (" & solutionStatus & ")"
                Case 2
                    Return "Problem dual infeasible".ToUpper() & " (" & solutionStatus & ")"
                Case 3
                    Return "Stopped on iterations".ToUpper() & " (" & solutionStatus & ")"
                Case 4
                    Return "Stopped due to errors".ToUpper() & " (" & solutionStatus & ")"
                Case 5
                    Return "Stopped by user".ToUpper() & " (" & solutionStatus & ")"
                Case Else
                    Return "Undefined error".ToUpper() & " (" & solutionStatus & ")"
            End Select

        End Function

    End Class 'SolveProblem

    Module ProblemCOPTDynamic
        Private _engine As COPT.Engine
        Private _switches() As String
        Private _ds As New DataSet
        'Private _dbPath As String = "C:\OPTMODELS\KNOXMIX\C-OPT_DATA_KNOXMIX.MDB"
        Private _dbPath As String = "C:\OPTMODELS\CBK35\CBK35.MDB"
        'Dim logTxt As LogHandler = Nothing
        'Dim logMsg As LogHandler = Nothing

        Public Sub Main()
            'logTxt = New LogHandler()
            'logMsg = New LogHandler()
        End Sub

        'Public Sub Main(ByRef txtLog As TextBox)
        '    logTxt = New LogHandler(txtLog)
        '    logMsg = New LogHandler()
        'End Sub

        'Public Sub Main(ByRef txtLog As TextBox, ByRef msgLog As TextBox)
        '    logTxt = New LogHandler(txtLog)
        '    logMsg = New LogHandler(msgLog)
        'End Sub

        Private Function MsgLogCallback(ByVal msg As String) As Integer
            'logMsg.WriteLine("***" & msg)
            Return CallResult.Success
        End Function

        'Public Sub Solve_Old(ByRef engine As COPT.Engine, ByVal switches() As String, ByVal solveProblem As SolveProblem, ByRef ds As DataSet)
        '    'Const NUM_COLS As Integer = 32
        '    'Const NUM_ROWS As Integer = 27
        '    'Const NUM_NZ As Integer = 83
        '    Const NUM_RNG As Integer = 0
        '    'Const INF As Double = 1.0E+37

        '    Dim NUM_COLS As Integer = 0
        '    Dim NUM_ROWS As Integer = 0
        '    Dim NUM_NZ As Integer = 0

        '    _switches = switches
        '    _engine = engine

        '    'Dim probname As String = My.Computer.FileSystem.GetFileInfo(formMain.txtDBPath.Text).ToString()  '"Afiro"
        '    Dim probname As String = engine.SolutionProjectName  'GenUtils.GetSwitchArgument(_switches, "/PRJ", 1) '"CoinMP_Test"
        '    Dim ncol As Integer = NUM_COLS
        '    Dim nrow As Integer = NUM_ROWS
        '    Dim nels As Integer = NUM_NZ
        '    Dim nrng As Integer = NUM_RNG

        '    Dim objectname As String = engine.SolutionProjectName  'GenUtils.GetSwitchArgument(_switches, "/PRJ", 1) '"Cost"
        '    'Dim objsens As Integer = CoinMP.ObjectSense.Min
        '    Dim objsens As Integer = CoinMP.ObjectSense.Max
        '    Dim objconst As Double = 0.0

        '    Dim ctr As Integer
        '    Dim proceed As Boolean = True
        '    'Dim sql As String

        '    Dim dobj() As Double
        '    Dim dclo() As Double
        '    Dim dcup() As Double
        '    Dim rtyp() As Char
        '    Dim ctyp() As Char = Nothing  'RKP/01-19-10/v2.2.126 - to add MIP capability to CoinMP
        '    Dim drhs() As Double
        '    Dim mbeg() As Integer
        '    Dim mcnt() As Integer
        '    Dim midx() As Integer
        '    Dim mval() As Double
        '    'Dim colids() As Integer
        '    Dim colNames() As String
        '    'Dim rowids() As Integer
        '    Dim rowNames() As String
        '    Dim cumTotal As Integer = 0
        '    Dim optimalValue As Double = -464.753142857

        '    Dim myArrayList As ArrayList

        '    'Dim linqTable As System.Data.EnumerableRowCollection(Of System.Data.DataRow)
        '    'Dim dblQueryResults As System.Data.EnumerableRowCollection(Of Double)
        '    'Dim intQueryResults As System.Data.EnumerableRowCollection(Of Integer)
        '    'Dim charQueryResults As System.Data.EnumerableRowCollection(Of Char)
        '    'Dim strQueryResults As System.Data.EnumerableRowCollection(Of String)

        '    Dim linqTable As System.Data.EnumerableRowCollection(Of System.Data.DataRow)
        '    Dim dblQueryResults As System.Data.EnumerableRowCollection(Of Double)
        '    Dim intQueryResults As System.Data.EnumerableRowCollection(Of Integer)
        '    Dim charQueryResults As System.Data.EnumerableRowCollection(Of Char)
        '    Dim strQueryResults As System.Data.EnumerableRowCollection(Of String)

        '    Dim intGroups As System.Collections.Generic.IEnumerable(Of Integer)
        '    'Dim queryResults As System.Data.EnumerableRowCollection 'Anonymous collection
        '    Dim tickCountStart As Integer = My.Computer.Clock.TickCount

        '    Dim usePLINQ As Boolean = GenUtils.IsSwitchAvailable(_switches, "/UsePLINQ")

        '    '9/17/09 - v2.2 Build 120
        '    'STM & RKP
        '    'Added ORDER BY clauses to all LINQ queries to prevent undesirable solution results going back into the database (for CoinMP).

        '    'db_Connect()
        '    _ds = ds
        '    'logTxt.WriteLine("Time taken by: db_Connect = " & ((My.Computer.Clock.TickCount - tickCountStart) / 1000).ToString() & " seconds.")

        '    NUM_COLS = _ds.Tables("tsysCol").Rows.Count
        '    NUM_ROWS = _ds.Tables("tsysRow").Rows.Count
        '    NUM_NZ = _ds.Tables("tsysMtx").Rows.Count

        '    ncol = NUM_COLS
        '    nrow = NUM_ROWS
        '    nels = NUM_NZ
        '    nrng = NUM_RNG

        '    'ReDim dobj(NUM_COLS)

        '    'Main(formMain.txtLog)

        '    'linqTable = _ds.Tables("tsysCol").AsEnumerable()
        '    'queryResults = From r In linqTable _
        '    '               Order By r("ColID") _
        '    '               Select _
        '    '                COL = CStr(r("COL")), _
        '    '                OBJ = CDbl(r("OBJ")), _
        '    '                LO = CDbl(r("LO")), _
        '    '                UP = CDbl(r("UP"))
        '    ''strQueryResults = queryResults(0)
        '    ''colNames = strQueryResults.ToArray()
        '    'dblQueryResults = queryResults(1)
        '    'dobj = dblQueryResults.ToArray()
        '    'dclo = queryResults(2).ToArray()
        '    'dcup = queryResults(3).ToArray()

        '    'linqTable = _ds.Tables("tsysRow").AsEnumerable()
        '    'queryResults = From r In linqTable _
        '    '               Order By r("ColID") _
        '    '               Select _
        '    '                ROW = CStr(r("ROW")), _
        '    '                RHS = CDbl(r("RHS")), _
        '    '                SENSE = CChar(r("SENSE"))
        '    'rowNames = queryResults(0).ToArray()
        '    'drhs = queryResults(1).ToArray()
        '    'rtyp = queryResults(2).ToArray()

        '    'OBJ Function
        '    'Coefficient of each variable in objective function
        '    '                       x1  x2
        '    'dobj()  = {0, -0.4, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, -0.32, 0, 0, 0, -0.6, _
        '    '0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, -0.48, 0, 0, 10} '-> 32
        '    'SELECT tsysCOL.OBJ FROM tsysCOL ORDER BY ColID/ColKey

        '    'logTxt.NewLine()
        '    'logTxt.WriteLine("---dobj---")
        '    'ReDim dobj(_ds.Tables("tsysCol").Rows.Count)
        '    'For ctr = 0 To _ds.Tables("tsysCol").Rows.Count - 1
        '    '    dobj(ctr) = _ds.Tables("tsysCol").Rows(ctr).Item("OBJ")
        '    '    logTxt.WriteLine(dobj(ctr).ToString)
        '    'Next

        '    '---LINQ code that works---
        '    'linqTable = _ds.Tables("tsysCol").AsEnumerable()
        '    'dblQueryResults = From r In linqTable Select OBJ = CDbl(r("OBJ"))
        '    'dobj = dblQueryResults.ToArray()

        '    '---LINQ code that works---
        '    'linqTable = _ds.Tables("tsysCol").AsEnumerable()
        '    'queryResults = From r In _ds.Tables("tsysCol").AsEnumerable() _
        '    '               Order By r("ColID") _
        '    '               Select ColID = r("ColID"), OBJ = CDbl(r("OBJ"))

        '    tickCountStart = My.Computer.Clock.TickCount

        '    '---OBJ---
        '    linqTable = _ds.Tables("tsysCol").AsEnumerable()
        '    If usePLINQ Then linqTable.AsParallel()
        '    'Dim linqTable1 = _ds.Tables("tsysCol").AsEnumerable().AsParallel
        '    dblQueryResults = From r In linqTable _
        '                   Order By r("ColID") Ascending _
        '                   Select OBJ = CDbl(r("OBJ"))
        '    dobj = dblQueryResults.ToArray()
        '    '---OBJ---

        '    'For d = 0 To dobj.Length - 1
        '    '    logTxt.WriteLine(dobj(d))
        '    'Next


        '    'Lower limit of each variable in objective function
        '    'dclo() = {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
        '    '                        0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0} '-> 32
        '    'SELECT tsysCOL.LO FROM tsysCOL ORDER BY ColID/ColKey

        '    '---LO---
        '    linqTable = _ds.Tables("tsysCol").AsEnumerable()
        '    If usePLINQ Then linqTable.AsParallel()
        '    dblQueryResults = From r In linqTable _
        '                   Order By r("ColID") Ascending _
        '                   Select LO = CDbl(r("LO"))
        '    dclo = dblQueryResults.ToArray()
        '    '---LO---

        '    ''ReDim dclo(_ds.Tables("tsysCol").Rows.Count)
        '    ''For ctr = 0 To _ds.Tables("tsysCol").Rows.Count - 1
        '    ''    dclo(ctr) = _ds.Tables("tsysCol").Rows(ctr).Item("LO")
        '    ''Next
        '    'dtTemp = _ds.Tables("tsysCol").AsEnumerable()
        '    'doubles = From dbl In dtTemp Select dbl!OBJ
        '    'dclo = doubles.ToArray()

        '    'Upper limit of each variable in objective function
        '    'Note:
        '    'There are 42 entries in this array, which may be incorrect. 32 should be the correct number.
        '    'Dim dcup() As Double = {INF, INF, INF, INF, INF, INF, INF, INF, INF, INF, INF, INF, _
        '    '        INF, INF, INF, INF, INF, INF, INF, INF, INF, INF, INF, INF, INF, INF, INF, _
        '    '        INF, INF, INF, INF, INF, INF, INF, INF, INF, INF, INF, INF, INF, INF, INF} '-> 42
        '    'SELECT tsysCOL.UP FROM tsysCOL ORDER BY ColID/ColKey

        '    '---UP---
        '    linqTable = _ds.Tables("tsysCol").AsEnumerable()
        '    If usePLINQ Then linqTable.AsParallel()
        '    dblQueryResults = From r In linqTable _
        '                   Order By r("ColID") Ascending _
        '                   Select UP = CDbl(r("UP"))
        '    dcup = dblQueryResults.ToArray()
        '    '---UP---

        '    ''ReDim dcup(_ds.Tables("tsysCol").Rows.Count)
        '    ' ''logTxt.WriteLine("Count: " & dcup.Count.ToString)
        '    ''For ctr = 0 To _ds.Tables("tsysCol").Rows.Count - 1
        '    ''    dcup(ctr) = _ds.Tables("tsysCol").Rows(ctr).Item("UP")
        '    ''    'logTxt.WriteLine(dcup(ctr))
        '    ''Next

        '    'Equality of each constraint row
        '    'Dim rtyp() As Char = "EELLEELLLLEELLEELLLLLLLLLLL" '-> 27
        '    'Dim rtyp() As Char = "EELLEELLLLEELLEELLLLLLLLLLL" '-> 27
        '    'SELECT SENSE FROM tsysROW ORDER BY RowID/RowKey
        '    'charQueryResults = From r In _ds.Tables("tsysRow").AsEnumerable() _
        '    '               Order By r("RowID") _
        '    '               Select SENSE = CStr(r("SENSE")).Chars(1)
        '    'rtyp = strQueryResults.ToArray()

        '    '---SENSE---
        '    linqTable = _ds.Tables("tsysRow").AsEnumerable()
        '    If usePLINQ Then linqTable.AsParallel()
        '    charQueryResults = From r In linqTable _
        '                   Order By r("RowID") Ascending _
        '                   Select SENSE = CChar(r("SENSE"))
        '    rtyp = charQueryResults.ToArray()
        '    '---SENSE---

        '    ''ReDim rtyp(_ds.Tables("tsysRow").Rows.Count)
        '    ' ''logTxt.WriteLine("Count: " & rtyp.Count.ToString)
        '    ''For ctr = 0 To _ds.Tables("tsysRow").Rows.Count - 1
        '    ''    rtyp(ctr) = _ds.Tables("tsysRow").Rows(ctr).Item("SENSE")
        '    ''    'logTxt.WriteLine(rtyp(ctr))
        '    ''Next

        '    'RHS value of each constraint row
        '    'Dim drhs() As Double = {0, 0, 80, 0, 0, 0, 80, 0, 0, 0, 0, 0, 500, 0, 0, 44, 500, 0, _
        '    '                        0, 0, 0, 0, 0, 0, 0, 310, 300} '-> 27
        '    'SELECT RHS FROM tsysROW ORDER BY RowID/RowKey

        '    '---RHS---
        '    linqTable = _ds.Tables("tsysRow").AsEnumerable()
        '    If usePLINQ Then linqTable.AsParallel()
        '    dblQueryResults = From r In linqTable _
        '                   Order By r("RowID") Ascending _
        '                   Select RHS = CDbl(r("RHS"))
        '    drhs = dblQueryResults.ToArray()
        '    '---RHS---

        '    ''Dim ctyp() As Char = "CCCCBBBBCCCCCC"
        '    ''RKP/01-19-10/v2.2.126
        '    ''---CTYP---
        '    ''charQueryResults = From r In linqTable _
        '    ''               Order By r("ColID") Ascending _
        '    ''               Select INTGR = _
        '    ''                CChar( _
        '    ''                    IIf( _
        '    ''                        r("INTGR") = True, "B", "C" _
        '    ''                    ) _
        '    ''                )
        '    'linqTable = _ds.Tables("tsysCol").AsEnumerable()
        '    'If usePLINQ Then linqTable.AsParallel()
        '    'charQueryResults = From r In linqTable _
        '    '               Order By r("ColID") Ascending _
        '    '               Select INTGR = _
        '    '                CChar( _
        '    '                    IIf( _
        '    '                        r("BINRY") = True, _
        '    '                        "B", _
        '    '                        IIf( _
        '    '                            r("INTGR") = True, _
        '    '                            "I", _
        '    '                            "C" _
        '    '                        ) _
        '    '                    ) _
        '    '                )
        '    'ctyp = charQueryResults.ToArray()
        '    ''---CTYP---

        '    '---CTYP---
        '    _engine.ProblemType = "C"
        '    proceed = False
        '    linqTable = _ds.Tables("tsysCol").AsEnumerable()
        '    intGroups = From r In linqTable _
        '                        Where r!INTGR = True _
        '                        Order By r("ColID") Ascending _
        '                        Group By r!ColID Into g = Group Select g.Count()
        '    If intGroups.ToArray().Count > 0 Then
        '        proceed = True
        '        _engine.ProblemType = "I"
        '    Else
        '        intGroups = From r In linqTable _
        '                            Where r!BINRY = True _
        '                            Order By r("ColID") Ascending _
        '                            Group By r!ColID Into g = Group Select g.Count()
        '        If intGroups.ToArray().Count > 0 Then
        '            proceed = True
        '            _engine.ProblemType = "B"
        '        End If
        '    End If
        '    If proceed Then
        '        'Dim ctyp() As Char = "CCCCBBBBCCCCCC"
        '        'RKP/01-19-10/v2.2.126
        '        'charQueryResults = From r In linqTable _
        '        '               Order By r("ColID") Ascending _
        '        '               Select INTGR = _
        '        '                CChar( _
        '        '                    IIf( _
        '        '                        r("INTGR") = True, "B", "C" _
        '        '                    ) _
        '        '                )
        '        linqTable = _ds.Tables("tsysCol").AsEnumerable()
        '        'If usePLINQ Then linqTable.AsParallel()
        '        charQueryResults = From r In linqTable _
        '                       Order By r("ColID") Ascending _
        '                       Select INTGR = _
        '                        CChar( _
        '                            IIf( _
        '                                r("BINRY") = True, _
        '                                "B", _
        '                                IIf( _
        '                                    r("INTGR") = True, _
        '                                    "I", _
        '                                    "C" _
        '                                ) _
        '                            ) _
        '                        )
        '        ctyp = charQueryResults.ToArray()
        '    End If
        '    '---CTYP---

        '    ''ReDim drhs(_ds.Tables("tsysRow").Rows.Count)
        '    ' ''logTxt.WriteLine("Count: " & rtyp.Count.ToString)
        '    ''For ctr = 0 To _ds.Tables("tsysRow").Rows.Count - 1
        '    ''    drhs(ctr) = _ds.Tables("tsysRow").Rows(ctr).Item("RHS")
        '    ''    'logTxt.WriteLine(rtyp(ctr))
        '    ''Next

        '    'Cumulative total of number of occurances of each variable
        '    'Note:
        '    'This is a 0-based array
        '    'Dim mbeg() As Integer = {0, 4, 6, 8, 10, 14, 18, 22, 26, 28, 30, 32, 34, 36, 38, 40, _
        '    '                 44, 46, 48, 50, 52, 56, 60, 64, 68, 70, 72, 74, 76, 78, 80, 82, 83} '-> 33

        '    'Returns 32 rows
        '    'SELECT A.ColID, A.CT, Sum(B.CT) AS SumOfCT
        '    'FROM 
        '    '(SELECT ColId, COUNT(*) AS CT FROM tsysMTX GROUP BY ColID ORDER BY ColID)
        '    'AS A, 
        '    '(SELECT ColId, COUNT(*) AS CT FROM tsysMTX GROUP BY ColID ORDER BY ColID)
        '    'AS B
        '    'WHERE (((B.ColID)<=[A].[ColID]))
        '    'GROUP BY A.ColID, A.CT
        '    'ORDER BY A.ColID, A.CT

        '    '---MBEG---
        '    'linqTable = _ds.Tables("MBEG").AsEnumerable()
        '    'intQueryResults = From r In linqTable _
        '    '               Order By r("ColID") _
        '    '               Select MATBEG = CInt(r("SumOfCT"))
        '    'mbeg = intQueryResults.ToArray()
        '    '---MBEG---

        '    'Returns 33 rows
        '    'SELECT 0 AS ColID, 0 AS CT, 0 AS SumOfCT FROM tsysCol  
        '    'UNION  
        '    'SELECT * FROM  
        '    '(  
        '    'SELECT A.ColID, A.CT, Sum(B.CT) AS SumOfCT  
        '    'FROM  
        '    '(SELECT ColID, COUNT(*) AS CT FROM tsysMTX GROUP BY ColID ORDER BY ColID) AS A,  
        '    '(SELECT ColID, COUNT(*) AS CT FROM tsysMTX GROUP BY ColID ORDER BY ColID) AS B
        '    'WHERE 
        '    'B.ColID <= A.ColID  
        '    'GROUP BY A.ColID, A.CT  
        '    'ORDER BY A.ColID, A.CT 
        '    ') RunningTotals 



        '    'Number of coefficients for each variable
        '    'Dim mcnt() As Integer = {4, 2, 2, 2, 4, 4, 4, 4, 2, 2, 2, 2, 2, 2, 2, 4, 2, 2, 2, 2, 4, _
        '    '                         4, 4, 4, 2, 2, 2, 2, 2, 2, 2, 1} '-> 32
        '    'SELECT COUNT(*) AS NoOfCoeffs FROM tsysMTX GROUP BY ColID/ColKey ORDER BY ColID/ColKey

        '    'linqTable = _ds.Tables("tsysMtx").AsEnumerable()
        '    'intQueryResults = From r In linqTable _
        '    '                  Order By r("ColID") _
        '    '                    Group By r!ColID Into g = Group _
        '    '               Select CInt(linqTable.Distinct().Count())
        '    'mcnt = intQueryResults.ToArray()
        '    linqTable = _ds.Tables("tsysMtx").AsEnumerable()
        '    If usePLINQ Then linqTable.AsParallel()
        '    intGroups = From r In linqTable _
        '                        Order By r("ColID") Ascending, r("RowID") Ascending _
        '                        Group By r!ColID Into g = Group Select g.Count()
        '    mcnt = intGroups.ToArray()
        '    'Dim tickCountStart2 As Integer = My.Computer.Clock.TickCount

        '    'RKP/09-22-09
        '    'May want to use Enumerable.Repeat, as shown in the following link as a replacement to the For...Next loop.
        '    'https://msmvps.com/blogs/deborahk/archive/2009/09/04/enumerable-repeat.aspx
        '    myArrayList = New ArrayList
        '    myArrayList.Add(0)
        '    For ctr = 0 To mcnt.Length - 1
        '        cumTotal = cumTotal + mcnt(ctr)
        '        myArrayList.Add(cumTotal)
        '    Next
        '    'mbeg = DirectCast(myArrayList.ToArray(GetType(Integer)), Integer())
        '    If usePLINQ Then
        '        mbeg = DirectCast(myArrayList.ToArray(GetType(Integer)), Integer())
        '    Else
        '        mbeg = DirectCast(myArrayList.ToArray(GetType(Integer)), Integer())
        '    End If




        '    'logTxt.WriteLine("Time taken to load mbeg = " & ((My.Computer.Clock.TickCount - tickCountStart2) / 1000).ToString() & " seconds.")

        '    'mcnt = intQueryResults.ToArray()
        '    'queryResults = From t In linqTable Order By t("ColID") Select t!ColID, linqTable.Distinct().Count()
        '    'intQueryResults = DirectCast(queryResults(1), System.Data.EnumerableRowCollection(Of Integer))
        '    'intQueryResults = From _
        '    '    t In ( _
        '    '        From t In linqTable Order By t("ColID") Select t!ColID, CInt(linqTable.Distinct().Count()) _
        '    '        ) _
        '    '    Order By t("ColID") _
        '    '    Select linqTable.Distinct().Count()

        '    'Group r By ColID = r.Field(Of Integer)("ColID") Into Group _

        '    'This is code to populate mbeg() using mcnt() as the basis, using ArrayList
        '    'myArrayList = New ArrayList(mcnt.Length + 1)
        '    'myArrayList.Add(0)
        '    'ctr = 0
        '    'For i As Integer = 0 To mcnt.Length - 1
        '    '    ctr = ctr + mcnt.GetValue(i)
        '    '    myArrayList.Add(ctr)
        '    'Next

        '    'mbeg = DirectCast(myArrayList.ToArray(GetType(Integer)), Integer())

        '    ''ReDim mcnt(_ds.Tables("mtxColCounts").Rows.Count)
        '    ' ''logTxt.WriteLine("Count: " & rtyp.Count.ToString)
        '    ''For ctr = 0 To _ds.Tables("mtxColCounts").Rows.Count - 1
        '    ''    mcnt(ctr) = _ds.Tables("mtxColCounts").Rows(ctr).Item("RowCount")
        '    ''    'logTxt.WriteLine(rtyp(ctr))
        '    ''Next
        '    'The first four (0,1,2,23) are positions of the variable x1 in rows (c1,c2,c3,c24).
        '    'Dim midx() As Integer = {0, 1, 2, 23, 0, 3, 0, 21, 1, 25, 4, 5, 6, 24, 4, 5, 7, 24, 4, 5, _
        '    '            8, 24, 4, 5, 9, 24, 6, 20, 7, 20, 8, 20, 9, 20, 3, 4, 4, 22, 5, 26, 10, 11, _
        '    '            12, 21, 10, 13, 10, 23, 10, 20, 11, 25, 14, 15, 16, 22, 14, 15, 17, 22, 14, _
        '    '            15, 18, 22, 14, 15, 19, 22, 16, 20, 17, 20, 18, 20, 19, 20, 13, 15, 15, 24, _
        '    '            14, 26, 15} '-> 83
        '    'SELECT RowID-1 AS idx FROM tsysMTX GROUP BY ColID/ColKey, RowID/RowKey ORDER BY ColID/ColKey, RowID/RowKey
        '    linqTable = _ds.Tables("tsysMtx").AsEnumerable()
        '    If usePLINQ Then linqTable.AsParallel()
        '    intQueryResults = From r In linqTable _
        '                   Order By r("ColID") Ascending, r("RowID") Ascending _
        '                   Select IDX = CInt(r("RowID")) - 1
        '    midx = intQueryResults.ToArray()

        '    'Coefficient value of each variable as defined by colNames
        '    'Dim mval() As Double = {-1, -1.06, 1, 0.301, 1, -1, 1, -1, 1, 1, -1, -1.06, 1, 0.301, _
        '    '        -1, -1.06, 1, 0.313, -1, -0.96, 1, 0.313, -1, -0.86, 1, 0.326, -1, 2.364, -1, _
        '    '        2.386, -1, 2.408, -1, 2.429, 1.4, 1, 1, -1, 1, 1, -1, -0.43, 1, 0.109, 1, -1, _
        '    '        1, -1, 1, -1, 1, 1, -0.43, 1, 1, 0.109, -0.43, 1, 1, 0.108, -0.39, 1, 1, _
        '    '        0.108, -0.37, 1, 1, 0.107, -1, 2.191, -1, 2.219, -1, 2.249, -1, 2.279, 1.4, _
        '    '        -1, 1, -1, 1, 1, 1} '-> 83
        '    'SELECT COEF FROM tsysMTX ORDER BY ColID/ColKey, RowID/RowKey
        '    linqTable = _ds.Tables("tsysMtx").AsEnumerable()
        '    If usePLINQ Then linqTable.AsParallel()
        '    dblQueryResults = From r In linqTable _
        '                   Order By r("ColID") Ascending, r("RowID") Ascending _
        '                   Select COEF = CDbl(r("COEF"))
        '    mval = dblQueryResults.ToArray()

        '    'Column names
        '    'Dim colNames() As String = {"x01", "x02", "x03", "x04", "x06", "x07", "x08", "x09", _
        '    '        "x10", "x11", "x12", "x13", "x14", "x15", "x16", "x22", "x23", "x24", "x25", _
        '    '        "x26", "x28", "x29", "x30", "x31", "x32", "x33", "x34", "x35", "x36", "x37", _
        '    '        "x38", "x39"} '-> 32
        '    'SELECT COL FROM tsysCOL ORDER BY ColID/ColKey

        '    '---COL---
        '    linqTable = _ds.Tables("tsysCol").AsEnumerable()
        '    If usePLINQ Then linqTable.AsParallel()
        '    '/SubNameWithID
        '    If GenUtils.IsSwitchAvailable(_switches, "/SubNameWithID") Then
        '        strQueryResults = From r In linqTable Order By r("ColID") Ascending Select CStr(r("ColID"))
        '    Else
        '        strQueryResults = From r In linqTable Order By r("ColID") Ascending Select CStr(r("COL"))
        '    End If
        '    colNames = strQueryResults.ToArray()
        '    'RKP/9-21-09-Used for GenSOLFile
        '    'intQueryResults = From r In linqTable Order By r("ColID") Ascending Select CInt(r("ColID"))
        '    'colids = intQueryResults.ToArray()
        '    '---COL---

        '    'Row names
        '    'Dim rowNames() As String = {"r09", "r10", "x05", "x21", "r12", "r13", "x17", "x18", _
        '    '        "x19", "x20", "r19", "r20", "x27", "x44", "r22", "r23", "x40", "x41", "x42", _
        '    '        "x43", "x45", "x46", "x47", "x48", "x49", "x50", "x51"} '-> 27
        '    'SELECT ROW FROM tsysROW ORDER BY RowID/RowKey

        '    '---ROW---
        '    linqTable = _ds.Tables("tsysRow").AsEnumerable()
        '    If usePLINQ Then linqTable.AsParallel()
        '    '/SubNameWithID
        '    If GenUtils.IsSwitchAvailable(_switches, "/SubNameWithID") Then
        '        strQueryResults = From r In linqTable Order By r("RowID") Ascending Select CStr(r("RowID"))
        '    Else
        '        strQueryResults = From r In linqTable Order By r("RowID") Ascending Select CStr(r("ROW"))
        '    End If
        '    rowNames = strQueryResults.ToArray()
        '    'RKP/9-21-09-Used for GenSOLFile
        '    'intQueryResults = From r In linqTable Order By r("RowID") Ascending Select CInt(r("RowID"))
        '    'rowids = intQueryResults.ToArray()
        '    '---ROW---

        '    'logTxt.WriteLine("Time taken to load arrays (via LINQ) = " & ((My.Computer.Clock.TickCount - tickCountStart) / 1000).ToString() & " seconds.")

        '    'RKP/01-22-10/v2.3.126
        '    If ctyp Is Nothing Then
        '        solveProblem.Run(_engine, _ds, _switches, probname, optimalValue, ncol, nrow, nels, nrng, objsens, objconst, _
        '            dobj, dclo, dcup, rtyp, drhs, Nothing, mbeg, mcnt, midx, mval, _
        '            colNames, rowNames, objectname, Nothing, Nothing)
        '    Else
        '        'RKP/01-19-10/v2.2.126
        '        'Modified to support MIP in the form of a new parameter, ctyp.
        '        solveProblem.Run(_engine, _ds, _switches, probname, optimalValue, ncol, nrow, nels, nrng, objsens, objconst, _
        '            dobj, dclo, dcup, rtyp, drhs, Nothing, mbeg, mcnt, midx, mval, _
        '            colNames, rowNames, objectname, Nothing, ctyp)
        '    End If
        '    'solveProblem.Run(_engine, _ds, _switches, probname, optimalValue, ncol, nrow, nels, nrng, objsens, objconst, _
        '    '    dobj, dclo, dcup, rtyp, drhs, Nothing, mbeg, mcnt, midx, mval, _
        '    '    colNames, rowNames, objectname, Nothing, Nothing)


        'End Sub

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="solverCoinMP"></param>
        ''' <param name="engine"></param>
        ''' <param name="switches"></param>
        ''' <param name="solveProblem"></param>
        ''' <param name="dtRow"></param>
        ''' <param name="dtCol"></param>
        ''' <param name="dtMtx"></param>
        ''' <remarks>
        ''' RKP/01-26-10/v2.3.127
        ''' Modified to use "LoadSolverArrays" routine, which is common for every C-OPT solver (CoinMP, CPLEX, lpsolve).
        ''' </remarks>
        Public Function Solve _
        ( _
            ByRef solverCoinMP As Solver_CoinMP, _
            ByRef engine As COPT.Engine, _
            ByVal switches() As String, _
            ByVal solveProblem As SolveProblem, _
            ByRef dtRow As DataTable, _
            ByRef dtCol As DataTable, _
            ByRef dtMtx As DataTable _
        ) As Boolean

            Dim optimalValue As Double = 0
            Dim objconst As Double = 0.0
            Dim nonZeroCount As Integer = 0
            Dim success As Boolean = False 'RKP/04-18-12/v3.2.166
            Dim objSense As Long = CoinMP.ObjectSense.Max

            '_ds = ds
            _switches = switches
            _engine = engine

            'Try
            '    nonZeroCount = dtMtx.Rows.Count
            'Catch ex As System.Exception
            '    nonZeroCount = _engine.SolutionNonZeros
            'End Try
            nonZeroCount = _engine.SolutionNonZeros
            'solverCoinMP.LoadSolverArrays(ds)
            'If GenUtils.IsSwitchAvailable(_switches, "/UseMinSysRes") Then
            '    'solverCoinMP.LoadSolverArrays(dtRow, dtCol, dtMtx, Solver.loadType.fromDatabase)
            '    solverCoinMP.LoadSolverArrays(dtRow, dtCol, dtMtx, True)
            'Else
            '    solverCoinMP.LoadSolverArrays(dtRow, dtCol, dtMtx, True)
            'End If

            'nonZeroCount = _engine.SolutionNonZeros

            'RKP/10-17-12/v4.2.181
            If GenUtils.IsSwitchAvailable(_switches, "/Sense") Then
                '/Sense "MIN" "-1"
                If GenUtils.GetSwitchArgument(_switches, "/Sense", 1).Trim().ToUpper() = "MAX" Then
                    objSense = CoinMP.ObjectSense.Max
                ElseIf GenUtils.GetSwitchArgument(_switches, "/Sense", 1).Trim().ToUpper() = "MIN" Then
                    objSense = CoinMP.ObjectSense.Min
                ElseIf GenUtils.GetSwitchArgument(_switches, "/Sense", 1).Trim().ToUpper() = "NONE" Then
                    objSense = CoinMP.ObjectSense.None
                Else
                    objSense = CoinMP.ObjectSense.Max
                End If
            Else
                objSense = CoinMP.ObjectSense.Max
            End If

            'RKP/01-22-10/v2.3.126
            'If solverCoinMP.array_ctyp Is Nothing Then
            'RKP/02-15-10/v2.3.130
            If Not solverCoinMP.isMIP Then

                success = solveProblem.Run(solverCoinMP, _engine, dtRow, dtCol, dtMtx, _switches, engine.SolutionProjectName, optimalValue, _engine.SolutionColumns, _engine.SolutionRows, _engine.SolutionNonZeros, 0, objSense, objconst, _
                    solverCoinMP.array_dobj, solverCoinMP.array_dclo, solverCoinMP.array_dcup, solverCoinMP.array_rtyp, solverCoinMP.array_drhs, Nothing, solverCoinMP.array_mbeg, solverCoinMP.array_mcnt, solverCoinMP.array_midx, solverCoinMP.array_mval, _
                    solverCoinMP.array_colNames, solverCoinMP.array_rowNames, engine.SolutionProjectName, Nothing, Nothing)
            Else
                'RKP/01-19-10/v2.2.126
                'Modified to support MIP in the form of a new parameter, ctyp.
                success = solveProblem.Run(solverCoinMP, _engine, dtRow, dtCol, dtMtx, _switches, engine.SolutionProjectName, optimalValue, _engine.SolutionColumns, _engine.SolutionRows, _engine.SolutionNonZeros, 0, objSense, objconst, _
                    solverCoinMP.array_dobj, solverCoinMP.array_dclo, solverCoinMP.array_dcup, solverCoinMP.array_rtyp, solverCoinMP.array_drhs, Nothing, solverCoinMP.array_mbeg, solverCoinMP.array_mcnt, solverCoinMP.array_midx, solverCoinMP.array_mval, _
                    solverCoinMP.array_colNames, solverCoinMP.array_rowNames, engine.SolutionProjectName, Nothing, solverCoinMP.array_ctyp)
            End If

            'solveProblem.Run(_engine, _ds, _switches, probname, optimalValue, ncol, nrow, nels, nrng, objsens, objconst, _
            '    dobj, dclo, dcup, rtyp, drhs, Nothing, mbeg, mcnt, midx, mval, _
            '    colNames, rowNames, objectname, Nothing, Nothing)

            Return success
        End Function

        Public Sub db_Update(ByRef dtCol As DataTable, ByRef dtRow As DataTable)
            Dim con As System.Data.OleDb.OleDbConnection
            Dim adapter As System.Data.OleDb.OleDbDataAdapter
            Dim sql As String
            Dim conStr As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & _dbPath 'formMain.txtDBPath.Text & ";;Jet OLEDB:System Database=C:\OPTMODELS\C-OPTSYS\System.MDW.COPT;User ID=Admin;Password="
            Dim dt As New DataTable()
            Dim sb As New StringBuilder
            Dim bldr As System.Data.OleDb.OleDbCommandBuilder
            Dim cmd As System.Data.OleDb.OleDbCommand
            Dim recordsAffected As Integer
            Dim tickCountStart As Integer = My.Computer.Clock.TickCount

            'Main(formMain.txtLog)

            con = New OleDb.OleDbConnection(conStr)
            Try
                con.Open()
            Catch ex As System.Exception
                'MsgBox(ex.Message, MsgBoxStyle.Exclamation)
                GenUtils.Message(GenUtils.MsgType.Critical, "Engine - Solver-CoinMP", ex.Message)
            End Try

            sql = "DELETE * FROM tsysSolution_COL"
            cmd = New OleDb.OleDbCommand(sql, con)
            recordsAffected = cmd.ExecuteNonQuery()

            sql = "SELECT * FROM tsysSolution_COL"
            adapter = New OleDb.OleDbDataAdapter(sql, conStr)
            adapter.UpdateCommand = New OleDbCommand(sql, con)
            bldr = New OleDbCommandBuilder(adapter)
            adapter.Fill(dtCol)
            bldr.GetUpdateCommand()
            adapter.Update(dtCol.Select(Nothing, Nothing, DataViewRowState.Deleted))
            adapter.Update(dtCol.Select(Nothing, Nothing, DataViewRowState.ModifiedCurrent))
            adapter.Update(dtCol.Select(Nothing, Nothing, DataViewRowState.Added))
            'adapter.Update(dtCol.Select(Nothing, Nothing, DataViewRowState.OriginalRows))
            adapter.Dispose()

            sql = "DELETE * FROM tsysSolution_ROW"
            cmd = New OleDb.OleDbCommand(sql, con)
            recordsAffected = cmd.ExecuteNonQuery()

            sql = "SELECT * FROM tsysSolution_ROW"
            adapter = New OleDb.OleDbDataAdapter(sql, conStr)
            adapter.UpdateCommand = New OleDbCommand(sql, con)
            bldr = New OleDbCommandBuilder(adapter)
            adapter.Fill(dtRow)
            bldr.GetUpdateCommand()
            'adapter.Update(dtRow)
            adapter.Update(dtRow.Select(Nothing, Nothing, DataViewRowState.Deleted))
            adapter.Update(dtRow.Select(Nothing, Nothing, DataViewRowState.ModifiedCurrent))
            adapter.Update(dtRow.Select(Nothing, Nothing, DataViewRowState.Added))
            adapter.Dispose()

            'logTxt.WriteLine("Time taken to write solution to database = " & ((My.Computer.Clock.TickCount - tickCountStart) / 1000).ToString() & " seconds.")

            'sql = "SELECT ColID, COL, LO, UP, OBJ FROM tsysCol"
            'adapter = New OleDb.OleDbDataAdapter(sql, conStr)
            'adapter.Fill(dt)
            '_ds.Tables.Add(dt)
            'Dim workCol As DataColumn = dt.Columns.Add("ColKey", Type.GetType("System.Int32"))
            'workCol.AutoIncrement = True
            ''workCol.AllowDBNull = False
            ''workCol.Unique = True
            ''dt.PrimaryKey = workCol
            'For Each dr As DataRow In dt.Rows
            '    dr.Item("ColKey") = dt.Rows.IndexOf(dr)
            'Next
            'dt = Nothing

            'sql = "SELECT RowID, ROW, SENSE, RHS FROM tsysRow"
            'adapter = New OleDb.OleDbDataAdapter(sql, conStr)
            'dt = New DataTable
            'adapter.Fill(dt)
            '_ds.Tables.Add(dt)
            'workCol = dt.Columns.Add("RowKey", Type.GetType("System.Int32"))
            'workCol.AutoIncrement = True
            ''workCol.AllowDBNull = False
            ''workCol.Unique = True
            ''dt.PrimaryKey = workCol
            'For Each dr As DataRow In dt.Rows
            '    dr.Item("RowKey") = dt.Rows.IndexOf(dr)
            'Next
            'dt = Nothing

            'sql = "SELECT ColID, RowID, COEF FROM tsysMtx"
            'adapter = New OleDb.OleDbDataAdapter(sql, conStr)
            'dt = New DataTable
            'adapter.Fill(dt)
            '_ds.Tables.Add(dt)
            'workCol = dt.Columns.Add("MtxKey", Type.GetType("System.Int32"))
            'workCol.AutoIncrement = True
            ''workCol.AllowDBNull = False
            ''workCol.Unique = True
            ''dt.PrimaryKey = workCol
            'For Each dr As DataRow In dt.Rows
            '    dr.Item("MtxKey") = dt.Rows.IndexOf(dr)
            'Next
            'dt = Nothing


            'Returns 33 rows
            'SELECT 0 AS ColID, 0 AS CT, 0 AS SumOfCT FROM tsysCol
            'UNION
            'SELECT * FROM
            '(
            'SELECT A.ColID, A.CT, Sum(B.CT) AS SumOfCT
            'FROM 
            '(SELECT ColId, COUNT(*) AS CT FROM tsysMTX GROUP BY ColID ORDER BY ColID)
            'AS A, 
            '(SELECT ColId, COUNT(*) AS CT FROM tsysMTX GROUP BY ColID ORDER BY ColID)
            'AS B
            'WHERE (((B.ColID)<=[A].[ColID]))
            'GROUP BY A.ColID, A.CT
            'ORDER BY A.ColID, A.CT
            ') RunningTotals
            'sb.AppendLine("SELECT 0 AS ColID, 0 AS CT, 0 AS SumOfCT FROM tsysCol ")
            'sb.AppendLine("UNION ")
            'sb.AppendLine("SELECT * FROM ")
            'sb.AppendLine("( ")
            'sb.AppendLine("SELECT A.ColID, A.CT, Sum(B.CT) AS SumOfCT ")
            'sb.AppendLine("FROM ")
            'sb.AppendLine("(SELECT ColId, COUNT(*) AS CT FROM tsysMTX GROUP BY ColID ORDER BY ColID) AS A, ")
            'sb.AppendLine("(SELECT ColId, COUNT(*) AS CT FROM tsysMTX GROUP BY ColID ORDER BY ColID) AS B ")
            'sb.AppendLine("WHERE (((B.ColID)<=[A].[ColID])) ")
            'sb.AppendLine("GROUP BY A.ColID, A.CT ")
            'sb.AppendLine("ORDER BY A.ColID, A.CT ")
            'sb.AppendLine(") RunningTotals ")
            'sql = sb.ToString()
            'adapter = New OleDb.OleDbDataAdapter(sql, conStr)
            'dt = New DataTable
            'adapter.Fill(dt)
            '_ds.Tables.Add(dt)

            'workCol = dt.Columns.Add("MtxKey", Type.GetType("System.Int32"))
            'workCol.AutoIncrement = True
            ''workCol.AllowDBNull = False
            ''workCol.Unique = True
            ''dt.PrimaryKey = workCol
            'For Each dr As DataRow In dt.Rows
            '    dr.Item("MtxKey") = dt.Rows.IndexOf(dr)
            'Next
            'dt = Nothing

            'sql = "SELECT SELECT DISTINCT COL, COUNT(ColID) AS RowCount FROM tsysMTX GROUP BY COL"
            'sql = _
            '    <Query>
            '        SELECT DISTINCT a.ColID, b.RowCount 
            '        FROM 
            '            tsysMTX a 
            '        INNER JOIN
            '        (
            '            SELECT DISTINCT COL, COUNT(ColID) AS RowCount FROM tsysMTX GROUP BY COL
            '        ) b
            '        ON 
            '            a.COL = b.COL
            '        ORDER BY
            '            a.ColID
            '    </Query>.Value
            'adapter = New OleDb.OleDbDataAdapter(sql, conStr)
            'dt = New DataTable
            'adapter.Fill(dt)
            '_ds.Tables.Add(dt)
            'dt = Nothing

            ''_ds.Tables(0).TableName = "tsysCol"
            ''_ds.Tables(1).TableName = "tsysRow"
            ''_ds.Tables(2).TableName = "tsysMtx"
            '_ds.Tables(3).TableName = "MBEG"
            '_ds.Tables(3).TableName = "mtxColCounts"

            'MsgBox("tsysCol:" & _ds.Tables(0).Rows.Count)
            'MsgBox("tsysRow:" & _ds.Tables(1).Rows.Count)
            'MsgBox("tsysMtx:" & _ds.Tables(2).Rows.Count)

            'logTxt.NewLine()
            'logTxt.WriteLine("Before con.Close----")
            'logTxt.WriteLine("tsysCol:" & _ds.Tables(0).Rows.Count)
            'logTxt.WriteLine("tsysRow:" & _ds.Tables(1).Rows.Count)
            'logTxt.WriteLine("tsysMtx:" & _ds.Tables(2).Rows.Count)
            'logTxt.WriteLine("mtxColCounts:" & _ds.Tables(3).Rows.Count)

            con.Close()
            con = Nothing
            adapter = Nothing
            '_ds = Nothing
            'dt = Nothing

            'MsgBox("tsysCol:" & _ds.Tables(0).Rows.Count)
            'MsgBox("tsysRow:" & _ds.Tables(1).Rows.Count)
            'MsgBox("tsysMtx:" & _ds.Tables(2).Rows.Count)
            'logTxt.NewLine()
            'logTxt.WriteLine("After con.Close----")
            'logTxt.WriteLine("tsysCol:" & _ds.Tables(0).Rows.Count)
            'logTxt.WriteLine("tsysRow:" & _ds.Tables(1).Rows.Count)
            'logTxt.WriteLine("tsysMtx:" & _ds.Tables(2).Rows.Count)
            'logTxt.WriteLine("mtxColCounts:" & _ds.Tables(3).Rows.Count)
        End Sub

        Public Sub db_Connect()
            Dim con As System.Data.OleDb.OleDbConnection
            Dim adapter As System.Data.OleDb.OleDbDataAdapter
            Dim sql As String
            'Dim dbPath As String = _dbPath '"C:\OPTMODELS\KNOXMIX\KNOXMIX.MDB" 'formMain.txtDBPath.Text  '"C:\OPTMODELS\BR1\BR1_DATA.MDB"
            'Dim dbPath As String = "C:\OPTMODELS\KNOXMIX\C-OPT_DATA_KNOXMIX.MDB"
            Dim conStr As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & _dbPath & ";;Jet OLEDB:System Database=C:\OPTMODELS\C-OPTSYS\System.MDW.COPT;User ID=Admin;Password="
            'Dim conStr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=F:\RaviP\Afiro.MDB;;Jet OLEDB:System Database=C:\OPTMODELS\C-OPTSYS\System.MDW.COPT;User ID=Admin;Password="

            'Main(formMain.txtLog)

            Dim dt As New DataTable()
            Dim sb As New StringBuilder


            con = New OleDb.OleDbConnection(conStr)
            Try
                con.Open()
            Catch ex As System.Exception
                'MsgBox(ex.Message, MsgBoxStyle.Exclamation)
                GenUtils.Message(GenUtils.MsgType.Critical, "Engine - Solver-CoinMP", ex.Message)
            End Try

            sql = "SELECT ColID, COL, LO, UP, OBJ FROM tsysCol"
            adapter = New OleDb.OleDbDataAdapter(sql, conStr)
            adapter.Fill(dt)
            _ds.Tables.Add(dt)
            'Dim workCol As DataColumn = dt.Columns.Add("ColKey", Type.GetType("System.Int32"))
            'workCol.AutoIncrement = True
            'workCol.AllowDBNull = False
            'workCol.Unique = True
            'dt.PrimaryKey = workCol
            'For Each dr As DataRow In dt.Rows
            '    dr.Item("ColKey") = dt.Rows.IndexOf(dr)
            'Next
            dt = Nothing

            sql = "SELECT RowID, ROW, SENSE, RHS FROM tsysRow"
            adapter = New OleDb.OleDbDataAdapter(sql, conStr)
            dt = New DataTable
            adapter.Fill(dt)
            _ds.Tables.Add(dt)
            'workCol = dt.Columns.Add("RowKey", Type.GetType("System.Int32"))
            'workCol.AutoIncrement = True
            'workCol.AllowDBNull = False
            'workCol.Unique = True
            'dt.PrimaryKey = workCol
            'For Each dr As DataRow In dt.Rows
            '    dr.Item("RowKey") = dt.Rows.IndexOf(dr)
            'Next
            dt = Nothing

            sql = "SELECT ColID, RowID, COEF FROM tsysMtx"
            adapter = New OleDb.OleDbDataAdapter(sql, conStr)
            dt = New DataTable
            adapter.Fill(dt)
            _ds.Tables.Add(dt)
            'workCol = dt.Columns.Add("MtxKey", Type.GetType("System.Int32"))
            'workCol.AutoIncrement = True
            'workCol.AllowDBNull = False
            'workCol.Unique = True
            'dt.PrimaryKey = workCol
            'For Each dr As DataRow In dt.Rows
            '    dr.Item("MtxKey") = dt.Rows.IndexOf(dr)
            'Next
            dt = Nothing

            _ds.Tables(0).TableName = "tsysCol"
            _ds.Tables(1).TableName = "tsysRow"
            _ds.Tables(2).TableName = "tsysMtx"
            '_ds.Tables(3).TableName = "MBEG"
            '_ds.Tables(3).TableName = "mtxColCounts"

            'MsgBox("tsysCol:" & _ds.Tables(0).Rows.Count)
            'MsgBox("tsysRow:" & _ds.Tables(1).Rows.Count)
            'MsgBox("tsysMtx:" & _ds.Tables(2).Rows.Count)

            'logTxt.NewLine()
            'logTxt.WriteLine("Before con.Close----")
            'logTxt.WriteLine("tsysCol:" & _ds.Tables(0).Rows.Count)
            'logTxt.WriteLine("tsysRow:" & _ds.Tables(1).Rows.Count)
            'logTxt.WriteLine("tsysMtx:" & _ds.Tables(2).Rows.Count)
            'logTxt.WriteLine("mtxColCounts:" & _ds.Tables(3).Rows.Count)

            con.Close()
            con = Nothing
            adapter = Nothing
            '_ds = Nothing
            dt = Nothing

            'MsgBox("tsysCol:" & _ds.Tables(0).Rows.Count)
            'MsgBox("tsysRow:" & _ds.Tables(1).Rows.Count)
            'MsgBox("tsysMtx:" & _ds.Tables(2).Rows.Count)
            'logTxt.NewLine()
            'logTxt.WriteLine("After con.Close----")
            'logTxt.WriteLine("tsysCol:" & _ds.Tables(0).Rows.Count)
            'logTxt.WriteLine("tsysRow:" & _ds.Tables(1).Rows.Count)
            'logTxt.WriteLine("tsysMtx:" & _ds.Tables(2).Rows.Count)
            'logTxt.WriteLine("mtxColCounts:" & _ds.Tables(3).Rows.Count)

        End Sub

        Public Sub GetRowCount()
            MsgBox("tsysCol:" & _ds.Tables(0).Rows.Count)
            MsgBox("tsysRow:" & _ds.Tables(1).Rows.Count)
            MsgBox("tsysMtx:" & _ds.Tables(2).Rows.Count)
        End Sub

        Public Sub TestDynamicArray()
            'Dim ctr As Short
            'Dim sql As String

            'Dim dobj() As Double
            'Dim dclo() As Double
            'Dim dcup() As Double
            'Dim rtyp() As Char
            'Dim drhs() As Double
            'Dim mcnt() As Integer


            'logTxt.NewLine()
            'logTxt.WriteLine("---TestDynamicArray---")
            'logTxt.WriteLine("After con.Close----")
            'ReDim dobj(_ds.Tables("tsysCol").Rows.Count)
            'For ctr = 0 To _ds.Tables("tsysCol").Rows.Count - 1
            '    dobj(ctr) = _ds.Tables("tsysCol").Rows(ctr).Item("OBJ")
            '    logTxt.WriteLine(dobj(ctr))
            'Next

            'ReDim dclo(_ds.Tables("tsysCol").Rows.Count)
            'For ctr = 0 To _ds.Tables("tsysCol").Rows.Count - 1
            '    dclo(ctr) = _ds.Tables("tsysCol").Rows(ctr).Item("LO")
            '    logTxt.WriteLine(dclo(ctr))
            'Next

            'ReDim dcup(_ds.Tables("tsysCol").Rows.Count)
            'logTxt.WriteLine("Count: " & dcup.Count.ToString)
            'For ctr = 0 To _ds.Tables("tsysCol").Rows.Count - 1
            '    dcup(ctr) = _ds.Tables("tsysCol").Rows(ctr).Item("UP")
            '    logTxt.WriteLine(dcup(ctr))
            'Next

            'ReDim rtyp(_ds.Tables("tsysRow").Rows.Count)
            'logTxt.WriteLine("Count: " & rtyp.Count.ToString)
            'For ctr = 0 To _ds.Tables("tsysRow").Rows.Count - 1
            '    rtyp(ctr) = _ds.Tables("tsysRow").Rows(ctr).Item("SENSE")
            '    logTxt.WriteLine(rtyp(ctr))
            'Next

            'ReDim rtyp(_ds.Tables("tsysRow").Rows.Count)
            'logTxt.WriteLine("Count: " & rtyp.Count.ToString)
            'For ctr = 0 To _ds.Tables("tsysRow").Rows.Count - 1
            '    rtyp(ctr) = _ds.Tables("tsysRow").Rows(ctr).Item("SENSE")
            '    logTxt.WriteLine(rtyp(ctr))
            'Next

            'ReDim drhs(_ds.Tables("tsysRow").Rows.Count)
            'logTxt.WriteLine("Count: " & drhs.Count.ToString)
            'For ctr = 0 To _ds.Tables("tsysRow").Rows.Count - 1
            '    drhs(ctr) = _ds.Tables("tsysRow").Rows(ctr).Item("RHS")
            '    logTxt.WriteLine(drhs(ctr))
            'Next

            'Dim mtx = _ds.Tables("tsysMtx").AsEnumerable()
            'ReDim drhs(_ds.Tables("tsysRow").Rows.Count)
            'logTxt.WriteLine("Count: " & drhs.Count.ToString)
            'For ctr = 0 To _ds.Tables("tsysRow").Rows.Count - 1
            '    drhs(ctr) = _ds.Tables("tsysRow").Rows(ctr).Item("RHS")
            '    logTxt.WriteLine(drhs(ctr))
            'Next

            'Dim dtTemp = _ds.Tables("tsysCol").AsEnumerable()
            'Dim doubles = From dbl In dtTemp Select dbl("PrimaryKey")
            'dobj = doubles.ToArray()
            'For d = 0 To dobj.Length - 1
            '    logTxt.WriteLine(dobj(d))
            'Next
            'Dim dblArray = doubles.ToArray()
            'For d = 0 To dblArray.Length - 1
            '    logTxt.WriteLine(dblArray(d))
            'Next


        End Sub

        Public Sub TestLinq()
            db_Connect()


            Dim dtTemp = _ds.Tables("tsysCol").AsEnumerable()
            'Dim doubles = From dbl In dtTemp Select dbl = "OBJ"
            Dim doubles = From dbl In dtTemp Select dbl!OBJ
            Dim dobj() = doubles.ToArray()

            'Dim doubles = From dbl In doublesDataTable Select dbl!double

            'Dim doublesArray = doubles.ToArray()

            'Dim d As Double
            'Console.WriteLine("Every other double from highest to lowest:")
            'For d = 0 To doublesArray.Length
            'Console.WriteLine(doublesArray(d))
            'd += 1
            'Next
        End Sub
    End Module 'Module ProblemCOPTDynamic
End Namespace 'Namespace COPT