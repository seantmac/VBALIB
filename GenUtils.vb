Imports System
Imports System.IO
Imports System.Diagnostics
Imports System.Collections.Specialized
Imports System.Collections.ObjectModel
Imports System.Collections
Imports System.Text
Imports System.Configuration
Imports System.Xml
Imports System.Text.RegularExpressions
Imports System.Xml.Xsl
Imports System.Linq
Imports System.Runtime.CompilerServices
Imports System.Reflection
Imports Excel = Microsoft.Office.Interop.Excel
Imports Microsoft.Practices.EnterpriseLibrary.Data
Imports Microsoft.Practices.EnterpriseLibrary.Common.Configuration
Imports Microsoft.Office.Interop.Access

Namespace COPT

    ''' <summary>This is a general-purpose class that holds methods of generic value across the entire application</summary>
    '''
    Public Class GenUtils

        'RKP/07-12-07/v1
        '********** START OPTIONS            **********
        'Option Compare Database
        'Option Explicit
        '********** END   OPTIONS            **********
        '********** START DLL DECLARATIONS   **********

        'Private Declare Function GetShortPathNameA Lib "kernel32" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
        '<Runtime.InteropServices.DllImport("kernel32.dll", SetLastError:=True, CharSet:=Runtime.InteropServices.CharSet.Auto)> _
        Public Function GetShortPathName(ByVal longPath As String, _
        <Runtime.InteropServices.MarshalAs(Runtime.InteropServices.UnmanagedType.LPTStr)> _
        ByVal ShortPath As System.Text.StringBuilder, _
        <Runtime.InteropServices.MarshalAs(Runtime.InteropServices.UnmanagedType.U4)> _
        ByVal bufferSize As Integer) As Integer
        End Function
        Private Declare Auto Function SetForegroundWindow Lib "user32.dll" (ByVal hWnd As IntPtr) As Boolean
        Private Declare Auto Function ShowWindowAsync Lib "user32.dll" (ByVal hWnd As IntPtr, ByVal nCmdShow As Integer) As Boolean
        Private Declare Auto Function IsIconic Lib "user32.dll" (ByVal hWnd As IntPtr) As Boolean
        Private Const SW_RESTORE As Integer = 9
        Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, ByRef nSize As Integer) As Integer

        '********** END   DLL DECLARATIONS   **********
        '********** START PUBLIC CONSTANTS   **********
        Public Enum StatusLevel
            Green = 0
            Yellow = 1
            Red = 2
        End Enum

        'RKP/05-03-12/v3.2.168
        Public Enum StatusLevelColor
            Green = ConsoleColor.Green
            GreenDark = ConsoleColor.DarkGreen
            Yellow = ConsoleColor.Yellow
            YellowDark = ConsoleColor.DarkYellow
            Red = ConsoleColor.Red
            RedDark = ConsoleColor.DarkRed
        End Enum

        Public Enum ReturnStatus
            Success
            Failure
        End Enum

        Public Enum MsgType
            Information
            Warning
            Critical
            'Supress 'RKP/04-05-10/v2.3.132
        End Enum

        Public Enum RunType
            Test
            Development
            Debug
            Production
        End Enum
        '********** END   PUBLIC CONSTANTS   **********
        '********** START PUBLIC VARIABLES   **********
        '********** END   PUBLIC VARIABLES   **********
        '********** START PRIVATE CONSTANTS  **********
        '********** END   PRIVATE CONSTANTS  **********
        '********** START PRIVATE VARIABLES  **********
        Private _fswFolder As FileSystemWatcher 'RKP/07-09-07/v2
        Private Shared _viewerProcess As Process = Nothing
        Private _databaseName As String 'RKP/01-31-08
        Private _currentDb As DAAB
        Private _switches() As String
        Private _switchesDict As Dictionary(Of String, String) 'RKP/05-11-12/v3.2.168
        Private _statusLevel As Short
        Private Parameters As StringDictionary
        '********** END   PRIVATE VARIABLES  **********
        '********** START USER DEFINED TYPES **********
        '********** END   USER DEFINED TYPES **********


        Public Sub New()
            '
        End Sub

        Public Sub New(ByVal currentDb As String)
            _currentDb = New DAAB(currentDb)
        End Sub

        Public Sub New(ByVal currentDb As DAAB)
            _currentDb = currentDb
        End Sub

        Public Sub New(ByVal switches() As String)
            _switches = switches
        End Sub

        ''' <summary>
        ''' Gives a unique name to current build.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' RKP/12-13-11/v3.0.156
        ''' "Rocky Mountain Juniper" Edition - v3.1.164
        ''' "Rocky Mountain Juniper 2" Edition - v3.1.165
        ''' "California Laurel" Edition - v3.2.166
        ''' </remarks>
        Public Shared Function GetVersionName() As String
            Return VersionName()
        End Function

        Public Shared Function FormatTime(ByVal startTime As Integer, ByVal endTime As Integer) As String
            'My.Computer.Clock.TickCount
            Dim strTime As String = Format(((endTime - startTime) / 1000) / 60, "0.00")
            Dim position As String = CStr(InStr(strTime, ".", CompareMethod.Text))
            If CDbl(position) > 0 Then
                'FormatTime = VBA.Left(vsTime, nPos - 1) & "m " & 
                'VBA.Format((VBA.Mid(vsTime, nPos) * 60), "0") & "s"
                Return Left(strTime, CInt(CDbl(position) - 1)) & "m " & Format(CDbl(Mid(strTime, CInt(position))) * 60, "0") & "s"
            Else
                Return strTime
            End If
        End Function

        ''' <summary>
        ''' Returns elapsed time in Xm Ys format (eg; 23m 2s).
        ''' </summary>
        ''' <param name="startTime"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' RKP/04-30-12/v3.2.167
        ''' </remarks>
        Public Shared Function FormatTime(ByVal startTime As Integer) As String
            Dim endTime As Integer = My.Computer.Clock.TickCount
            'My.Computer.Clock.TickCount
            Dim strTime As String = Format(((endTime - startTime) / 1000) / 60, "0.00")
            Dim position As String = CStr(InStr(strTime, ".", CompareMethod.Text))
            If CDbl(position) > 0 Then
                'FormatTime = VBA.Left(vsTime, nPos - 1) & "m " & 
                'VBA.Format((VBA.Mid(vsTime, nPos) * 60), "0") & "s"
                Return Left(strTime, CInt(CDbl(position) - 1)) & "m " & Format(CDbl(Mid(strTime, CInt(position))) * 60, "0") & "s"
            Else
                Return strTime
            End If
        End Function

        Public Sub WriteAllText(ByVal file As String, ByVal text As String, ByVal append As Boolean)
            My.Computer.FileSystem.WriteAllText(file, text, append)
        End Sub

        Public Shared Function ConfigRead(ByVal dbName As String) As String

            'Dim reader As System.Configuration.IConfigurationSectionHandler
            'Dim config As Microsoft.Practices.EnterpriseLibrary.Configuration


            'Return System.Configuration.ConfigurationManager.AppSettings("author")
            Try
                'Return ConfigurationManager.ConnectionStrings(dbName).ConnectionString
                Return GetSysConfig.ConnectionStrings.ConnectionStrings(dbName).ConnectionString
            Catch ex As Exception
                Return ""
            End Try



            'Return ConfigurationSettings.AppSettings("author").ToString()
            'Return reader.GetValue("author", System.Type.GetType("String"))

            'Return ConfigurationSettings.GetConfig("dataConfiguration").ToString()

            'My.Computer.FileSystem.GetFileInfo("").Extension.tou()
        End Function

        'Public Sub ConfigAdd()

        'End Sub

        'Public Sub ConfigEdit()

        'End Sub

        'Public Sub ConfigDelete()

        'End Sub

        Public Function AppPath() As String
            Return My.Application.Info.DirectoryPath & "\"
        End Function

        Public Function ShortPathName(ByVal Path As String) As String
            Dim sb As New System.Text.StringBuilder(1024)

            Dim tempVal As Integer = GetShortPathName(Path, sb, 1024)
            If tempVal <> 0 Then
                Dim Result As String = sb.ToString()
                Return Result
            Else
                'Throw New Exception("Failed to return a short path")
                Return Path
            End If
        End Function

        Public Shared Function GetAllConnectionStrings() As Collection
            Dim connStrCollection As New ConnectionStringSettingsCollection
            Dim connStrEntry As ConnectionStringSettings
            Dim sb As New StringBuilder
            Dim list As New Collection

            'connStrCollection = ConfigurationManager.ConnectionStrings
            connStrCollection = GetSysConfig.ConnectionStrings.ConnectionStrings  'config.ConnectionStrings.ConnectionStrings

            For Each connStrEntry In connStrCollection
                'sb.Append(connStrEntry.Name)
                'sb.Append(" / ")
                'sb.Append(connStrEntry.ConnectionString)
                'sb.Append(vbNewLine)
                list.Add(connStrEntry.ConnectionString, connStrEntry.Name)
            Next

            Return list 'sb.ToString()
        End Function

        Public Shared Function GetAllDatabases() As String()
            Dim connStrCollection As New ConnectionStringSettingsCollection
            'Dim connStrEntry As ConnectionStringSettings
            'Dim sb As New StringBuilder
            Dim list As String()
            Dim ctr As Short = 0

            'connStrCollection = ConfigurationManager.ConnectionStrings
            connStrCollection = GetSysConfig.ConnectionStrings.ConnectionStrings  'config.ConnectionStrings.ConnectionStrings

            ReDim list(connStrCollection.Count - 1)
            'For Each connStrEntry In connStrCollection
            For ctr = 0 To CShort(connStrCollection.Count - 1)
                'sb.Append(connStrEntry.Name)
                'sb.Append(" / ")
                'sb.Append(connStrEntry.ConnectionString)
                'sb.Append(vbNewLine)

                'RKP/04-20-11/v2.4.143
                'Ignore the following entry, which automatically get picked up.
                'data source=.\SQLEXPRESS;Integrated Security=SSPI;AttachDBFilename=|DataDirectory|aspnetdb.mdf;User Instance=true

                'RKP/12-14-11/v3.0.156
                'Added checks to exclude projects like "LOCALSQLSERVER", "OraAspNetConString" and [PROJECTNAME].
                'Projects with [], are hidden from C-OPTApp.
                If _
                    Not connStrCollection.Item(ctr).Name.ToString().Trim().ToUpper().Contains("LOCALSQLSERVER") _
                    And _
                    Not connStrCollection.Item(ctr).Name.ToString().Trim().ToUpper().Contains("CONSTRING") _
                    And _
                    Not connStrCollection.Item(ctr).Name.ToString().Trim().ToUpper().StartsWith("[") _
                Then '.Equals("LOCALSQLSERVER") = 0 Then
                    list(ctr) = connStrCollection.Item(ctr).Name.ToString()  'connStrEntry.Name
                End If
            Next

            Return list 'sb.ToString()

        End Function

        Public Function GetCurrentDB() As String
            'Return ConfigurationManager.GetSection("defaultDatabase")
            'ConfigurationManager.RefreshSection("dataConfiguration")
            _databaseName = GenUtils.GetSysConfigValue("defaultDatabase").ToString()
            Return GenUtils.GetSysConfigValue("defaultDatabase").ToString()
            'Dim dbPath as String = ConfigurationSettings.AppSettings("DatabasePath")
        End Function

        Public Sub ChangeCurrentDB(ByVal databaseName As String)
            'ConfigurationManager.AppSettings.Item("defaultDatabase") = databaseName
            'ConfigurationManager.RefreshSection("dataConfiguration")
            GenUtils.SetSysConfigUpdateKey("defaultDatabase", databaseName)
            _databaseName = GenUtils.GetSysConfigValue("defaultDatabase").ToString()

        End Sub

        Public Property CurrentDbName() As String
            Get
                Return _databaseName
            End Get
            Set(ByVal value As String)
                _databaseName = value
            End Set
        End Property

        Public ReadOnly Property CurrentDb() As DAAB
            Get
                Return New DAAB(_databaseName)
            End Get
        End Property

        Public Shared Function GetHelpViewerExecutable() As String
            Dim common As String = Environment.GetEnvironmentVariable("CommonProgramFiles")
            Return Path.Combine(common, "Microsoft Shared\Help 8\dexplore.exe")
        End Function

        Public Shared Sub ShowHelpViewer()
            Const HelpViewerArguments As String = "/helpcol ms-help://ms.EntLib.2007Apr /LaunchFKeywordTopic ExceptionhandlingQS2"
            ' Process has never been started. Initialize and launch the viewer.
            If (_viewerProcess Is Nothing) Then
                ' Initialize the Process information for the help viewer
                _viewerProcess = New Process

                _viewerProcess.StartInfo.FileName = GetHelpViewerExecutable()
                _viewerProcess.StartInfo.Arguments = HelpViewerArguments
                _viewerProcess.Start()
            ElseIf (_viewerProcess.HasExited) Then
                ' Process previously started, then exited. Start the process again.
                _viewerProcess.Start()
            Else
                ' Process was already started - bring it to the foreground
                Dim hWnd As IntPtr = _viewerProcess.MainWindowHandle
                If (IsIconic(hWnd)) Then
                    ShowWindowAsync(hWnd, SW_RESTORE)
                End If
                SetForegroundWindow(hWnd)
            End If
        End Sub

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="file"></param>
        ''' <param name="verb"></param>
        ''' <param name="arg"></param>
        ''' <remarks>
        ''' RKP/04-27-12/v3.2.167
        ''' </remarks>
        Public Shared Function LaunchFile( _
            ByVal file As String, _
            ByVal verb As String, _
            ByVal arg As String, _
            ByVal readToEnd As Boolean, _
            ByVal waitForExit As Boolean, _
            ByVal waitTimeInMilliSeconds As Integer _
        ) As Integer

            Dim info As New System.Diagnostics.ProcessStartInfo()
            Dim p As New System.Diagnostics.Process
            Dim exitCode As Integer

            info.UseShellExecute = False
            info.CreateNoWindow = True
            info.RedirectStandardOutput = True
            info.RedirectStandardError = True
            'info.Verb = "print"
            'info.Verb = "open"
            info.Verb = verb
            info.WindowStyle = ProcessWindowStyle.Normal
            info.FileName = file
            If Not String.IsNullOrEmpty(arg) Then
                info.Arguments = arg
            End If
            Try
                'System.Diagnostics.Process.Start(info)
                p.StartInfo = info
                p.Start()
                If readToEnd Then
                    Console.WriteLine(p.StandardOutput.ReadToEnd)
                End If
                If waitForExit Then
                    If waitTimeInMilliSeconds = 0 Then
                        p.WaitForExit()
                    Else
                        p.WaitForExit(waitTimeInMilliSeconds)
                    End If
                End If
            Catch ex As Exception
                'MsgBox(ex.Message)
                GenUtils.Message(MsgType.Critical, "Launch File", ex.Message)
            Finally
                'exitCode = p.ExitCode()
                p.Dispose()
                p.Close()
            End Try

            Return exitCode

        End Function

        ''' <summary>This method runs the DiffCompare algorithm</summary>
        ''' <param name="vsDBName">Valid C-OPT Project</param>
        ''' <param name="workDir">Work Directory</param>
        ''' <param name="vsFile1">file1</param>
        ''' <param name="vsFile2">file2</param>
        ''' <param name=" vbDeleteAllTempTables">deleteAllTempTables</param>
        ''' <exception cref="System.OverflowException">
        ''' Thrown when <paramref name="denominator"/><c> = 0</c>.
        ''' </exception>
        ''' <remarks>
        ''' COMPARE /PRJ "CBK35" /RUNFile1 "C:\OPTMODELS\CBK35\Output\CBK35-CBKWY-ORANGBAS1-KEYRES-20080717-2119.xml" /RUNFile2 "C:\OPTMODELS\CBK35\Output\CBK35-CBKWY-ORANGBAS1-KEYRES-20080717-2120.xml" /DeleteTempTables "False"
        ''' COMPARE /PRJ "HIPNET" /RUNFile1 "C:\OPTMODELS\HIPNET\Output\RUNFILE1.xml" /RUNFile2 "C:\OPTMODELS\HIPNET\Output\RUNFILE2.xml" /DeleteTempTables "False"
        ''' </remarks> 
        Public Sub RunComparison(ByVal vsDBName As String, _
                                 ByVal workDir As String, _
                                 ByVal vsFile1 As String, _
                                 ByVal vsFile2 As String, _
                                 ByVal vbDeleteAllTempTables As Boolean, _
                                 ByRef rsOutput As String)
            '**********************************************
            'Author  :  Ravi Poluri/Sean MacDermant
            'Date/Ver:  01-08-08/v2.0
            'Input   :
            'Output  :
            'Comments:
            '04-09-05/v1.5.39 (adapted from C-OPT1 codebase)
            '**********************************************

            Dim startTime As Integer = My.Computer.Clock.TickCount

            Dim bContinue As Boolean
            Dim lCtr As Integer
            'Dim lCtr2 As Integer
            Dim iRet As Integer
            'Dim lRet As Long
            Dim sSQL As String
            Dim sSQL2 As String
            Dim sTemp As String
            'Dim sTmpName As String
            'Dim sTmpValue As String
            Dim nKeyFieldCtr As Integer
            'Dim lStartTime As Long
            'Dim oRS As Object 'ADODB.Recordset
            'Dim oTable As Object 'ADOX.Table
            Dim dt As DataTable
            Dim dt1 As DataTable
            'Dim dt2 As DataTable
            Dim ds As New DataSet("RunComp")
            Dim dr As DataRow
            Dim file As String = ""
            Dim runName1 As String = ""
            Dim runName2 As String = ""

            rsOutput = ""
            If vsFile1 = "" Or vsFile2 = "" Then
                bContinue = False
                MsgBox("You must enter valid ""Run Results"" file(s) in order to run a comparison.", vbInformation)
            ElseIf vsFile1 = vsFile2 Then
                bContinue = False
                MsgBox("You must enter two different ""Run Results"" files in order to run a comparison.", vbInformation)
                'ElseIf vsQueryName = "" Then
                'do nothing
            Else
                bContinue = True
            End If

            If bContinue = True Then
                Console.WriteLine("Diff Compare is now running...")
                Debug.Print("Diff Compare is now running...")


                sSQL = "DROP TABLE tRun1"
                Try
                    _currentDb.ExecuteNonQuery(sSQL, True)
                Catch ex As Exception

                End Try
                sSQL = "DROP TABLE tRun2"
                Try
                    _currentDb.ExecuteNonQuery(sSQL, True)
                Catch ex As Exception

                End Try
                sSQL = "DROP TABLE tRunDiff"
                Try
                    _currentDb.ExecuteNonQuery(sSQL, True)
                Catch ex As Exception

                End Try
                sSQL = "DROP TABLE tRunDiff1"
                Try
                    _currentDb.ExecuteNonQuery(sSQL, True)
                Catch ex As Exception

                End Try
                sSQL = "DROP TABLE tRunDiff2"
                Try
                    _currentDb.ExecuteNonQuery(sSQL, True)
                Catch ex As Exception

                End Try
                sSQL = "DROP TABLE tRunDiffDetail_Temp"
                Try
                    _currentDb.ExecuteNonQuery(sSQL, True)
                Catch ex As Exception

                End Try
                sSQL = "DROP TABLE tRunDiffDetail"
                Try
                    _currentDb.ExecuteNonQuery(sSQL, True)
                Catch ex As Exception

                End Try

                dt1 = New DataTable '("RunComp")
                'Dim ds As New DataSet("RunComp")
                'Import vsFile1 into tRun1
                Try ' 'Import vsFile1 into tRun1
                    'ds = New DataSet("RunComp")

                    dt1.ReadXmlSchema(Left(vsFile1, vsFile1.Length - 4) & ".schema.xml")
                    dt1.ReadXml(vsFile1)
                    'ds.ReadXml(vsFile1)
                Catch ex As Exception  'Import vsFile1 into tRun1
                    'MsgBox(ex.Message)
                    GenUtils.Message(MsgType.Critical, "Run Comparison", ex.Message & vbNewLine & "Unable to read schema file:" & vbNewLine & vsFile1)
                End Try  'Import vsFile1 into tRun1 'ds.ReadXml(vsFile1)

                'ds.Tables(0).WriteXml(vsFile2)
                'ds.Dispose()
                'Check if there are any records in the recordset
                If dt1.Rows.Count = 0 Then
                    'error
                Else
                    'Make sure the recordset has the ONLY "key" field: key_UniqueCode
                    nKeyFieldCtr = 0
                    For lCtr = 0 To dt1.Columns.Count - 1
                        'If ds.Tables(0).Columns(lCtr).Name = "KEY_UNIQUECODE" Then
                        If dt1.Columns(lCtr).ColumnName.ToUpper().Equals("KEY_UNIQUECODE") Then
                            nKeyFieldCtr = nKeyFieldCtr + 1
                            Exit For
                        End If
                        'End If
                    Next
                    If nKeyFieldCtr <> 1 Then
                        MsgBox("The file:" & vbNewLine & vsFile1 & vbNewLine & "is missing the key fields: key_UniqueCode" & vbNewLine & "Please make sure you have the key field before you ""Run Comparison"" again.", vbExclamation)
                    Else
                        'dt = ds.Tables.Add("tRun1")
                        'sSQL = "CREATE TABLE [tRun1] ( "
                        'sSQL = "SELECT 'SCN01' AS Field1, 'SCN02' AS Field2, 8.90 AS Field3 INTO tRun1"
                        'iRet = _currentDb.ExecuteNonQuery(sSQL)
                        sSQL = "SELECT "
                        sSQL2 = "INSERT INTO tRun1 ("
                        For lCtr = 0 To dt1.Columns.Count - 1
                            'dt.Columns.Add(ds.Tables(0).Columns(lCtr).ColumnName, ds.Tables(0).Columns(lCtr).DataType)
                            'sSQL = sSQL & ds.Tables(0).Columns(lCtr).ColumnName & ","
                            sSQL2 = sSQL2 & dt1.Columns(lCtr).ColumnName & ","
                            If dt1.Columns(lCtr).ColumnName.ToUpper().StartsWith("TXT_") Or dt1.Columns(lCtr).ColumnName.ToUpper().StartsWith("KEY_") Or dt1.Columns(lCtr).ColumnName.ToUpper().StartsWith("SUBKEY_") Then
                                sSQL = sSQL & "'TEXT' AS " & dt1.Columns(lCtr).ColumnName & ","
                            Else
                                sSQL = sSQL & "9.00000000000000001E-13 AS " & dt1.Columns(lCtr).ColumnName & ","
                            End If
                        Next
                        sSQL = Left(sSQL, Len(sSQL) - 1)
                        sSQL = sSQL & " INTO tRun1"
                        Try
                            iRet = _currentDb.ExecuteNonQuery(sSQL)
                            If iRet = -1 Then
                                EntLib.COPT.Log.Log("Error executing SQL in RunComparison." & vbNewLine & sSQL)
                            End If
                            sSQL = "DELETE * FROM tRun1"
                            iRet = _currentDb.ExecuteNonQuery(sSQL)
                            If iRet = -1 Then
                                EntLib.COPT.Log.Log("Error executing SQL in RunComparison." & vbNewLine & sSQL)
                            End If
                        Catch ex As Exception
                            'MsgBox(ex.Message)
                            GenUtils.Message(MsgType.Critical, "Run Comparison", ex.Message & vbNewLine & sSQL)
                        End Try

                        sSQL2 = Left(sSQL2, Len(sSQL2) - 1)
                        sSQL2 = sSQL2 & ") VALUES ("


                        'ds.WriteXml("C:\OPTMODELS\CBK35\Output\Test\ds.xml", XmlWriteMode.IgnoreSchema)
                        'iRet = _currentDb.UpdateDataSet(ds, "tRun1")

                        dt = _currentDb.GetDataTable("SELECT * FROM tRun1", True)
                        dt.TableName = "tRun1"
                        ds.Tables.Add(dt)

                        For Each dr In dt1.Rows  'ds.Tables(0).Rows
                            ds.Tables("tRun1").ImportRow(dr)
                        Next
                        'ds.Tables.Remove("Table")
                        _currentDb.UpdateDataSet(ds.Tables("tRun1"), "SELECT * FROM tRun1")
                        dt = _currentDb.GetDataTable("SELECT DISTINCT txt_RUN_NAME FROM tRun1")
                        runName1 = dt.Rows(0).Item(0).ToString()

                        'For lCtr = 0 To ds.Tables(0).Rows.Count - 1
                        '    sSQL = ""

                        '    'sSQL = "INSERT INTO [tRun1] ("
                        '    'sSQL = "SELECT "
                        '    'For lCtr2 = 0 To ds.Tables(0).Columns.Count - 1
                        '    '    'sSQL = sSQL & "[" & ds.Tables(0).Columns(lCtr2).ColumnName & "],"
                        '    '    If ds.Tables(0).Columns(lCtr2).ColumnName.ToUpper().StartsWith("TXT_") Or ds.Tables(0).Columns(lCtr2).ColumnName.ToUpper().StartsWith("KEY_") Or ds.Tables(0).Columns(lCtr2).ColumnName.ToUpper().StartsWith("SUBKEY_") Then
                        '    '        sSQL = sSQL & "'TEXT' AS " & ds.Tables(0).Columns(lCtr2).ColumnName & ","
                        '    '    Else
                        '    '        sSQL = sSQL & "0.0001 AS " & ds.Tables(0).Columns(lCtr2).ColumnName & ","
                        '    '    End If
                        '    'Next
                        '    'sSQL = Left(sSQL, Len(sSQL) - 1)
                        '    ''sSQL = sSQL & ") VALUES ("
                        '    'sSQL = sSQL & " INTO tRun1"
                        '    'Try
                        '    '    iRet = _currentDb.ExecuteNonQuery(sSQL)
                        '    '    sSQL = "DELETE * FROM tRun1"
                        '    '    iRet = _currentDb.ExecuteNonQuery(sSQL)
                        '    'Catch ex As Exception

                        '    'End Try

                        '    'sSQL = "INSERT INTO [tRun1] ("

                        '    For lCtr2 = 0 To ds.Tables(0).Columns.Count - 1
                        '        sTmpName = ds.Tables(0).Columns(lCtr2).Caption  'oRS.Fields(lCtr2).Name
                        '        'sTemp = sFormatSQLString(oRS.Fields(lCtr2).Value & "")
                        '        sTemp = sFormatSQLString(ds.Tables(0).Rows(lCtr).Item(lCtr2) & "")
                        '        If sTemp = "" Then
                        '            sTmpValue = 0
                        '        Else
                        '            sTmpValue = sTemp
                        '        End If
                        '        If UCase(Left(sTmpName, 4)) = "KEY_" Or UCase(Left(sTmpName, 4)) = "SUB_" Or UCase(Left(sTmpName, 4)) = "SUBK" Or UCase(Left(sTmpName, 4)) = "TXT_" Then
                        '            sSQL = sSQL & "'" & sTmpValue & "'"
                        '        Else
                        '            sSQL = sSQL & sTmpValue
                        '        End If
                        '        sSQL = sSQL & ","
                        '    Next
                        '    sSQL = Left(sSQL, Len(sSQL) - 1)
                        '    sSQL = sSQL & ")"

                        '    sSQL = sSQL2 & sSQL

                        '    iRet = _currentDb.ExecuteNonQuery(sSQL)
                        'Next

                        ds = Nothing
                        dt = Nothing
                        dt1 = Nothing

                        'Import vsFile2 into tRun2
                        ds = New DataSet("RunComp")
                        Try
                            'ds.ReadXml(vsFile2)
                            dt1 = New DataTable '("RunComp")
                            dt1.ReadXmlSchema(Left(vsFile2, vsFile2.Length - 4) & ".schema.xml")
                            dt1.ReadXml(vsFile2)
                        Catch ex As Exception
                            'MsgBox(ex.Message)
                            GenUtils.Message(MsgType.Critical, "Run Comparison", ex.Message & "Unable to read schema file:" & vbNewLine & vsFile2)
                        End Try


                        'Check if there are any records in the recordset
                        If dt1.Rows.Count = 0 Then
                        Else

                            'Make sure the recordset has the ONLY "key" field: key_UniqueCode
                            nKeyFieldCtr = 0
                            For lCtr = 0 To dt1.Columns.Count - 1
                                'If ds.Tables(0).Columns(lCtr).Name = "KEY_UNIQUECODE" Then
                                If dt1.Columns(lCtr).ColumnName.ToUpper().Equals("KEY_UNIQUECODE") Then
                                    nKeyFieldCtr = nKeyFieldCtr + 1
                                    Exit For
                                End If
                                'End If
                            Next
                            If nKeyFieldCtr <> 1 Then
                                MsgBox("The file:" & vbNewLine & vsFile2 & vbNewLine & "is missing the key fields: key_UniqueCode" & vbNewLine & "Please make sure you have the key field before you ""Run Comparison"" again.", vbExclamation)
                            Else
                                'dt = ds.Tables.Add("tRun2")
                                'For lCtr = 0 To ds.Tables(0).Columns.Count - 1
                                '    dt.Columns.Add(ds.Tables(0).Columns(lCtr).ColumnName, ds.Tables(0).Columns(lCtr).GetType())
                                'Next

                                sSQL = "SELECT "
                                sSQL2 = "INSERT INTO tRun2 ("
                                For lCtr = 0 To dt1.Columns.Count - 1
                                    'dt.Columns.Add(ds.Tables(0).Columns(lCtr).ColumnName, ds.Tables(0).Columns(lCtr).DataType)
                                    'sSQL = sSQL & ds.Tables(0).Columns(lCtr).ColumnName & ","
                                    sSQL2 = sSQL2 & dt1.Columns(lCtr).ColumnName & ","
                                    If dt1.Columns(lCtr).ColumnName.ToUpper().StartsWith("TXT_") Or dt1.Columns(lCtr).ColumnName.ToUpper().StartsWith("KEY_") Or dt1.Columns(lCtr).ColumnName.ToUpper().StartsWith("SUBKEY_") Then
                                        sSQL = sSQL & "'TEXT' AS " & dt1.Columns(lCtr).ColumnName & ","
                                    Else
                                        sSQL = sSQL & "9.00000000000000001E-13 AS " & dt1.Columns(lCtr).ColumnName & ","
                                    End If
                                Next
                                sSQL = Left(sSQL, Len(sSQL) - 1)
                                sSQL = sSQL & " INTO tRun2"
                                Try
                                    iRet = _currentDb.ExecuteNonQuery(sSQL)
                                    If iRet = -1 Then
                                        EntLib.COPT.Log.Log("Error executing SQL in RunComparison." & vbNewLine & sSQL)
                                    End If
                                    sSQL = "DELETE * FROM tRun2"
                                    iRet = _currentDb.ExecuteNonQuery(sSQL)
                                    If iRet = -1 Then
                                        EntLib.COPT.Log.Log("Error executing SQL in RunComparison." & vbNewLine & sSQL)
                                    End If
                                Catch ex As Exception
                                    'MsgBox(ex.Message)
                                    GenUtils.Message(MsgType.Critical, "Run Comparison", ex.Message & vbNewLine & sSQL)
                                End Try

                                sSQL2 = Left(sSQL2, Len(sSQL2) - 1)
                                sSQL2 = sSQL2 & ") VALUES ("


                                dt = _currentDb.GetDataTable("SELECT * FROM tRun2", True)
                                dt.TableName = "tRun2"
                                ds.Tables.Add(dt)

                                For Each dr In dt1.Rows
                                    ds.Tables("tRun2").ImportRow(dr)
                                Next
                                _currentDb.UpdateDataSet(ds.Tables("tRun2"), "SELECT * FROM tRun2")
                                dt = _currentDb.GetDataTable("SELECT DISTINCT txt_RUN_NAME FROM tRun2")
                                runName2 = dt.Rows(0).Item(0).ToString()

                                'For lCtr = 0 To ds.Tables(0).Rows.Count - 1
                                '    sSQL = ""

                                '    For lCtr2 = 0 To ds.Tables(0).Columns.Count - 1
                                '        sTmpName = ds.Tables(0).Columns(lCtr2).Caption  'oRS.Fields(lCtr2).Name
                                '        'sTemp = sFormatSQLString(oRS.Fields(lCtr2).Value & "")
                                '        sTemp = sFormatSQLString(ds.Tables(0).Rows(lCtr).Item(lCtr2) & "")
                                '        If sTemp = "" Then
                                '            sTmpValue = 0
                                '        Else
                                '            sTmpValue = sTemp
                                '        End If
                                '        If UCase(Left(sTmpName, 4)) = "KEY_" Or UCase(Left(sTmpName, 4)) = "SUB_" Or UCase(Left(sTmpName, 4)) = "SUBK" Or UCase(Left(sTmpName, 4)) = "TXT_" Then
                                '            sSQL = sSQL & "'" & sTmpValue & "'"
                                '        Else
                                '            sSQL = sSQL & sTmpValue
                                '        End If
                                '        sSQL = sSQL & ","
                                '    Next
                                '    sSQL = Left(sSQL, Len(sSQL) - 1)
                                '    sSQL = sSQL & ")"

                                '    sSQL = sSQL2 & sSQL

                                '    iRet = _currentDb.ExecuteNonQuery(sSQL)
                                'Next

                                'Run a diff between the fields of tRun1 and tRun2
                                'LEFT JOIN
                                sSQL = "SELECT TOP 1 * FROM tRun2"
                                dt = _currentDb.GetDataTable(sSQL)

                                sSQL = "SELECT "
                                For lCtr = 0 To dt.Columns.Count - 1
                                    If dt.Columns(lCtr).Caption.ToUpper.StartsWith("KEY_") Or _
                                        dt.Columns(lCtr).Caption.ToUpper.StartsWith("SUB_") Or _
                                        dt.Columns(lCtr).Caption.ToUpper.StartsWith("SUBKEY_") Then

                                        sSQL = sSQL & "tRun2." & dt.Columns(lCtr).Caption & ","
                                    ElseIf dt.Columns(lCtr).Caption.ToUpper.StartsWith("TXT_") Then
                                        sSQL = sSQL & "tRun2." & dt.Columns(lCtr).Caption & " AS " & dt.Columns(lCtr).Caption & "_2," & "tRun1." & dt.Columns(lCtr).Caption & " AS " & dt.Columns(lCtr).Caption & "_1,"
                                    Else
                                        sSQL = sSQL & "IIF(ISNULL(tRun2." & dt.Columns(lCtr).Caption & "), 0, tRun2." & dt.Columns(lCtr).Caption & ") AS " & dt.Columns(lCtr).Caption & "_2,"
                                        sSQL = sSQL & "IIf(IsNull(tRun1." & dt.Columns(lCtr).Caption & "), 0, tRun1." & dt.Columns(lCtr).Caption & ") AS " & dt.Columns(lCtr).Caption & "_1,"

                                    End If
                                Next
                                sSQL = Left(sSQL, Len(sSQL) - 1)
                                sSQL = sSQL & " INTO tRunDiff1"
                                sSQL = sSQL & " FROM "
                                sSQL = sSQL & "tRun2 LEFT JOIN tRun1 ON "
                                sSQL = sSQL & "tRun2.key_UniqueCode = tRun1.key_UniqueCode "
                                Try
                                    iRet = _currentDb.ExecuteNonQuery(sSQL)
                                    If iRet = -1 Then
                                        EntLib.COPT.Log.Log("Error executing SQL in RunComparison." & vbNewLine & sSQL)
                                    End If
                                Catch ex As Exception
                                    'MsgBox(ex.Message)
                                    GenUtils.Message(MsgType.Critical, "Run Comparison", ex.Message & vbNewLine & sSQL)
                                End Try

                                sSQL = "UPDATE tRunDiff1 SET txt_RUN_NAME_1 = '' WHERE txt_RUN_NAME_1 IS NULL"
                                Try
                                    iRet = _currentDb.ExecuteNonQuery(sSQL)
                                    If iRet = -1 Then
                                        EntLib.COPT.Log.Log("Error executing SQL in RunComparison." & vbNewLine & sSQL)
                                    End If
                                Catch ex As Exception
                                    'MsgBox(ex.Message)
                                    GenUtils.Message(MsgType.Critical, "Run Comparison", ex.Message & vbNewLine & sSQL)
                                End Try
                                sSQL = "UPDATE tRunDiff1 SET txt_TIME_STAMP_1 = '' WHERE txt_TIME_STAMP_1 IS NULL"
                                Try
                                    iRet = _currentDb.ExecuteNonQuery(sSQL)
                                    If iRet = -1 Then
                                        EntLib.COPT.Log.Log("Error executing SQL in RunComparison." & vbNewLine & sSQL)
                                    End If
                                Catch ex As Exception
                                    'MsgBox(ex.Message)
                                    GenUtils.Message(MsgType.Critical, "Run Comparison", ex.Message & vbNewLine & sSQL)
                                End Try

                                'RIGHT JOIN
                                sSQL = "SELECT TOP 1 * FROM tRun1"
                                dt = _currentDb.GetDataTable(sSQL)
                                sSQL = "SELECT "
                                For lCtr = 0 To dt.Columns.Count - 1
                                    If dt.Columns(lCtr).Caption.ToUpper.StartsWith("KEY_") Or _
                                        dt.Columns(lCtr).Caption.ToUpper.StartsWith("SUB_") Or _
                                        dt.Columns(lCtr).Caption.ToUpper.StartsWith("SUBKEY_") Then

                                        sSQL = sSQL & "tRun1." & dt.Columns(lCtr).Caption & ","
                                    ElseIf dt.Columns(lCtr).Caption.ToUpper.StartsWith("TXT_") Then
                                        sSQL = sSQL & "tRun2." & dt.Columns(lCtr).Caption & " AS " & dt.Columns(lCtr).Caption & "_2," & "tRun1." & dt.Columns(lCtr).Caption & " AS " & dt.Columns(lCtr).Caption & "_1,"
                                    Else
                                        sSQL = sSQL & "IIF(ISNULL(tRun2." & dt.Columns(lCtr).Caption & "), 0, tRun2." & dt.Columns(lCtr).Caption & ") AS " & dt.Columns(lCtr).Caption & "_2,"
                                        sSQL = sSQL & "IIf(IsNull(tRun1." & dt.Columns(lCtr).Caption & "), 0, tRun1." & dt.Columns(lCtr).Caption & ") AS " & dt.Columns(lCtr).Caption & "_1,"

                                    End If
                                Next
                                sSQL = Left(sSQL, Len(sSQL) - 1)
                                sSQL = sSQL & " INTO tRunDiff2 "
                                sSQL = sSQL & "FROM "
                                sSQL = sSQL & "tRun1 LEFT JOIN tRun2 ON "
                                sSQL = sSQL & "tRun1.key_UniqueCode = tRun2.key_UniqueCode "
                                sSQL = sSQL & "WHERE "
                                sSQL = sSQL & "tRun2.key_UniqueCode IS NULL "
                                Try
                                    iRet = _currentDb.ExecuteNonQuery(sSQL)
                                    If iRet = -1 Then
                                        EntLib.COPT.Log.Log("Error executing SQL in RunComparison." & vbNewLine & sSQL)
                                    End If
                                Catch ex As Exception
                                    'MsgBox(ex.Message)
                                    GenUtils.Message(MsgType.Critical, "Run Comparison", ex.Message & vbNewLine & sSQL)
                                End Try

                                sSQL = "UPDATE tRunDiff2 SET txt_RUN_NAME_2 = '' WHERE txt_RUN_NAME_2 IS NULL"
                                Try
                                    iRet = _currentDb.ExecuteNonQuery(sSQL)
                                    If iRet = -1 Then
                                        EntLib.COPT.Log.Log("Error executing SQL in RunComparison." & vbNewLine & sSQL)
                                    End If
                                Catch ex As Exception
                                    'MsgBox(ex.Message)
                                    GenUtils.Message(MsgType.Critical, "Run Comparison", ex.Message & vbNewLine & sSQL)
                                End Try
                                sSQL = "UPDATE tRunDiff2 SET txt_TIME_STAMP_2 = '' WHERE txt_TIME_STAMP_2 IS NULL"
                                Try
                                    iRet = _currentDb.ExecuteNonQuery(sSQL)
                                    If iRet = -1 Then
                                        EntLib.COPT.Log.Log("Error executing SQL in RunComparison." & vbNewLine & sSQL)
                                    End If
                                Catch ex As Exception
                                    'MsgBox(ex.Message)
                                    GenUtils.Message(MsgType.Critical, "Run Comparison", ex.Message & vbNewLine & sSQL)
                                End Try


                                'UNION
                                sSQL = "SELECT * INTO tRunDiff FROM (SELECT * FROM tRunDiff1 UNION SELECT * FROM tRunDiff2)"
                                Try
                                    iRet = _currentDb.ExecuteNonQuery(sSQL)
                                    If iRet = -1 Then
                                        EntLib.COPT.Log.Log("Error executing SQL in RunComparison." & vbNewLine & sSQL)
                                    End If
                                Catch ex As Exception
                                    'MsgBox(ex.Message)
                                    GenUtils.Message(MsgType.Critical, "Run Comparison", ex.Message & vbNewLine & sSQL)
                                End Try


                                'DIFF - TEMP
                                sSQL = "SELECT TOP 1 * FROM tRun2"
                                dt = _currentDb.GetDataTable(sSQL)
                                sSQL = "SELECT "
                                For lCtr = 0 To dt.Columns.Count - 1
                                    If dt.Columns(lCtr).Caption.ToUpper.StartsWith("KEY_") Or _
                                        dt.Columns(lCtr).Caption.ToUpper.StartsWith("SUB_") Or _
                                        dt.Columns(lCtr).Caption.ToUpper.StartsWith("SUBKEY_") Then

                                        sSQL = sSQL & dt.Columns(lCtr).Caption & ","
                                    ElseIf dt.Columns(lCtr).Caption.ToUpper.StartsWith("TXT_") Then
                                        sSQL = sSQL & dt.Columns(lCtr).Caption & "_2," & dt.Columns(lCtr).Caption & "_1,"
                                    Else
                                        sSQL = sSQL & dt.Columns(lCtr).Caption & "_2," & dt.Columns(lCtr).Caption & "_1," & dt.Columns(lCtr).Caption & "_2-" & dt.Columns(lCtr).Caption & "_1 AS dif_" & dt.Columns(lCtr).Caption & ","
                                    End If
                                Next
                                sSQL = Left(sSQL, Len(sSQL) - 1)
                                sSQL = sSQL & " INTO tRunDiffDetail_Temp "
                                sSQL = sSQL & "FROM "
                                sSQL = sSQL & "tRunDiff"
                                Try
                                    iRet = _currentDb.ExecuteNonQuery(sSQL)
                                    If iRet = -1 Then
                                        EntLib.COPT.Log.Log("Error executing SQL in RunComparison." & vbNewLine & sSQL)
                                    End If
                                Catch ex As Exception
                                    'MsgBox(ex.Message)
                                    GenUtils.Message(MsgType.Critical, "Run Comparison", ex.Message & vbNewLine & sSQL)
                                End Try


                                'DIFF - FINAL
                                sSQL = "SELECT TOP 1 * FROM tRunDiffDetail_Temp"
                                dt = _currentDb.GetDataTable(sSQL)
                                sSQL = "SELECT "
                                For lCtr = 0 To dt.Columns.Count - 1
                                    If dt.Columns(lCtr).Caption.ToUpper.StartsWith("KEY_") Or _
                                        dt.Columns(lCtr).Caption.ToUpper.StartsWith("SUB_") Or _
                                        dt.Columns(lCtr).Caption.ToUpper.StartsWith("SUBKEY_") Then

                                        sSQL = sSQL & dt.Columns(lCtr).Caption & ","
                                    ElseIf dt.Columns(lCtr).Caption.ToUpper.EndsWith("_2") Then
                                        sSQL = sSQL & dt.Columns(lCtr).Caption & ","
                                    End If
                                Next
                                For lCtr = 0 To dt.Columns.Count - 1
                                    If dt.Columns(lCtr).Caption.ToUpper.StartsWith("KEY_") Or _
                                        dt.Columns(lCtr).Caption.ToUpper.StartsWith("SUB_") Or _
                                        dt.Columns(lCtr).Caption.ToUpper.StartsWith("SUBKEY_") Then

                                        'do nothing
                                    ElseIf Right(dt.Columns(lCtr).Caption, 2) = "_1" Then
                                        sSQL = sSQL & dt.Columns(lCtr).Caption & ","
                                    End If
                                Next
                                For lCtr = 0 To dt.Columns.Count - 1
                                    If dt.Columns(lCtr).Caption.ToUpper.StartsWith("DIF_") Then
                                        sTemp = Mid(dt.Columns(lCtr).Caption, 5)
                                        sSQL = sSQL & dt.Columns(lCtr).Caption & " AS diff_" & Mid(sTemp, InStr(1, sTemp, "_", vbTextCompare) + 1) & ","
                                    End If
                                Next
                                sSQL = Left(sSQL, Len(sSQL) - 1)
                                sSQL = sSQL & " INTO tRunDiffDetail "
                                sSQL = sSQL & "FROM "
                                sSQL = sSQL & "tRunDiffDetail_Temp"
                                Try
                                    iRet = _currentDb.ExecuteNonQuery(sSQL)
                                    If iRet = -1 Then
                                        EntLib.COPT.Log.Log("Error executing SQL in RunComparison." & vbNewLine & sSQL)
                                    End If
                                Catch ex As Exception
                                    'MsgBox(ex.Message)
                                    GenUtils.Message(MsgType.Critical, "Run Comparison", ex.Message & vbNewLine & sSQL)
                                End Try


                                If vbDeleteAllTempTables Then
                                    sSQL = "DROP TABLE tRun1"
                                    Try
                                        _currentDb.ExecuteNonQuery(sSQL)

                                    Catch ex As Exception

                                    End Try
                                    sSQL = "DROP TABLE tRun2"
                                    Try
                                        _currentDb.ExecuteNonQuery(sSQL)
                                    Catch ex As Exception

                                    End Try
                                    sSQL = "DROP TABLE tRunDiff"
                                    Try
                                        _currentDb.ExecuteNonQuery(sSQL)
                                    Catch ex As Exception

                                    End Try
                                    sSQL = "DROP TABLE tRunDiff1"
                                    Try
                                        _currentDb.ExecuteNonQuery(sSQL)
                                    Catch ex As Exception

                                    End Try
                                    sSQL = "DROP TABLE tRunDiff2"
                                    Try
                                        _currentDb.ExecuteNonQuery(sSQL)
                                    Catch ex As Exception

                                    End Try
                                    sSQL = "DROP TABLE tRunDiffDetail_Temp"
                                    Try
                                        _currentDb.ExecuteNonQuery(sSQL)
                                    Catch ex As Exception

                                    End Try
                                End If

                                'Save tRunDiffDetail as a CSV file
                                ds = Nothing
                                dt = _currentDb.GetDataTable("SELECT * FROM tRunDiffDetail", True)
                                dt.TableName = "Table"
                                ds = New DataSet("NewDataSet")
                                ds.Tables.Add(dt)
                                'file = workDir & "\" & miscParams.Item("PRJ_NAME").ToString() & "-" & miscParams.Item("RUN_NAME").ToString() & "-" & dt.Rows(ctr).Item("TextAbbr").ToString() & "-" & timeStamp '& ".xml"

                                'file = workDir & "\" & vsDBName & "-[" & runName1 & "--" & runName2 & "]-" & vsFile1.Substring(0, vsFile1.Length - 18).Split(CChar("-"))(vsFile1.Substring(0, vsFile1.Length - 18).Split(CChar("-")).Length - 1) & "DIFF-" & My.Computer.Clock.LocalTime.Year().ToString() & My.Computer.Clock.LocalTime.Month().ToString().PadLeft(2, CChar("0")) & My.Computer.Clock.LocalTime.Day().ToString().PadLeft(2, CChar("0")) & "-" & My.Computer.Clock.LocalTime.Hour.ToString().PadLeft(2, CChar("0")) & My.Computer.Clock.LocalTime.Minute.ToString().PadLeft(2, CChar("0"))
                                file = workDir & "\" & vsDBName & "-" & runName1 & "--" & runName2 & "-" & "DIFF-" & My.Computer.Clock.LocalTime.Year().ToString() & My.Computer.Clock.LocalTime.Month().ToString().PadLeft(2, CChar("0")) & My.Computer.Clock.LocalTime.Day().ToString().PadLeft(2, CChar("0")) & "-" & My.Computer.Clock.LocalTime.Hour.ToString().PadLeft(2, CChar("0")) & My.Computer.Clock.LocalTime.Minute.ToString().PadLeft(2, CChar("0"))

                                Try
                                    My.Computer.FileSystem.DeleteFile(file & ".xml.tmp")
                                Catch ex As Exception

                                End Try
                                ds.WriteXml(file & ".xml.tmp")

                                Try
                                    My.Computer.FileSystem.DeleteFile(file & ".xml")
                                Catch ex As Exception

                                End Try
                                ds.WriteXml(file & ".xml")

                                'Create the XsltSettings object with script enabled.
                                Dim settings As New XsltSettings(False, True)

                                'Execute the transform.
                                Dim xslt As New System.Xml.Xsl.XslCompiledTransform()
                                Try
                                    xslt.Load(workDir & "\DataSetToCSV.xsl", settings, New XmlUrlResolver)
                                Catch ex As Exception
                                    Try
                                        xslt.Load(My.Application.Info.DirectoryPath & "\DataSetToCSV.xsl", settings, New XmlUrlResolver)
                                    Catch ex1 As Exception
                                        Try
                                            'C:\OPTMODELS\C-OPTSYS
                                            xslt.Load("C:\OPTMODELS\C-OPTSYS" & "\DataSetToCSV.xsl", settings, New XmlUrlResolver)
                                        Catch ex2 As Exception
                                            'MsgBox(ex1.Message)
                                            GenUtils.Message(MsgType.Critical, "Run Comparison", ex2.Message & vbNewLine & "Unable to load XSL file:" & "C:\OPTMODELS\C-OPTSYS" & "\DataSetToCSV.xsl")
                                        End Try
                                    End Try
                                End Try
                                'dsTemp.WriteXml(file & ".xml.tmp")
                                xslt.Transform(file & ".xml.tmp", file & ".csv.tmp")

                                xslt = Nothing

                                Dim sr As New IO.StreamReader(file & ".csv.tmp")
                                'Dim content As String = sr.ReadToEnd


                                Dim sw As New IO.StreamWriter(file & ".csv", False)
                                Dim delim As String = ""
                                Dim delimiter As String = ","
                                For Each col As DataColumn In dt.Columns

                                    sw.Write(delim)
                                    sw.Write(col.ColumnName)
                                    delim = delimiter
                                Next
                                sw.WriteLine()
                                'sw.Write(content)
                                sw.Write(sr.ReadToEnd)
                                sw.Close()
                                sw.Dispose()
                                sr.Close()
                                sr.Dispose()

                                Try
                                    My.Computer.FileSystem.DeleteFile(file & ".csv.tmp")
                                Catch ex As Exception

                                End Try
                                Try
                                    My.Computer.FileSystem.DeleteFile(file & ".xml.tmp")
                                Catch ex As Exception

                                End Try

                                'rsOutput = "Diff Compare file:" & vbNewLine & file & ".csv" & vbNewLine & "has been created at:" & vbNewLine & workDir
                                rsOutput = file & ".csv"

                                Console.WriteLine("Diff Compare completed successfully!")
                                Debug.Print("Diff Compare completed successfully!")

                            End If 'If nKeyFieldCtr <> 1 Then 'tRun2
                        End If
                    End If 'If nKeyFieldCtr <> 1 Then 'tRun1
                End If 'If ds.Tables(0).Rows.Count = 0 Then
            End If 'If bContinue = True Then



            COPT.GenUtils.FormatTime(startTime, My.Computer.Clock.TickCount)

        End Sub

        ''' <summary>This method runs the diff compare algorithm</summary>
        ''' <param name="vsDBName">Valid C-OPT Project</param>
        ''' <param name="workDir">Work Directory</param>
        ''' <param name="vsFile1">file1</param>
        ''' <param name="vsFile2">file2</param>
        ''' <param name=" vbDeleteAllTempTables">deleteAllTempTables</param>
        ''' <exception cref="System.OverflowException">
        ''' Thrown when <paramref name="denominator"/><c> = 0</c>.
        ''' </exception>
        ''' <remarks><c>Dim a As Integer</c></remarks> 
        Public Sub RunComparison(ByVal vsDBName As String, _
                                 ByVal workDir As String, _
                                 ByVal vsFile1 As String, _
                                 ByVal vsFile2 As String, _
                                 ByVal vbDeleteAllTempTables As Boolean, _
                                 ByRef rsOutput As String, _
                                 ByVal useADOClassic As Boolean)
            '**********************************************
            'Author  :  Ravi Poluri/Sean MacDermant
            'Date/Ver:  01-08-08/v2.0
            'Input   :
            'Output  :
            'Comments:
            '04-09-05/v1.5.39 (adapted from C-OPT1 codebase)
            '**********************************************

            Dim startTime As Integer = My.Computer.Clock.TickCount

            Dim bContinue As Boolean
            Dim lCtr As Integer
            'Dim lCtr2 As Integer
            Dim iRet As Integer
            'Dim lRet As Long
            Dim sSQL As String
            Dim sSQL2 As String
            Dim sTemp As String
            'Dim sTmpName As String
            'Dim sTmpValue As String
            Dim nKeyFieldCtr As Integer
            'Dim lStartTime As Long
            'Dim oRS As Object 'ADODB.Recordset
            'Dim oTable As Object 'ADOX.Table
            Dim dt As DataTable
            Dim ds As New DataSet("RunComp")
            Dim dr As DataRow
            Dim rs1 As ADODB.Recordset = Nothing
            Dim rs2 As ADODB.Recordset = Nothing
            Dim file As String = ""
            Dim runName1 As String = ""
            Dim runName2 As String = ""

            rsOutput = ""
            If vsFile1 = "" Or vsFile2 = "" Then
                bContinue = False
                MsgBox("You must enter valid ""Run Results"" file(s) in order to run a comparison.", vbInformation)
            ElseIf vsFile1 = vsFile2 Then
                bContinue = False
                MsgBox("You must enter two different ""Run Results"" files in order to run a comparison.", vbInformation)
                'ElseIf vsQueryName = "" Then
                'do nothing
            Else
                bContinue = True
            End If

            If bContinue = True Then
                sSQL = "DROP TABLE tRun1"
                Try
                    _currentDb.ExecuteNonQuery(sSQL, True)
                Catch ex As Exception

                End Try
                sSQL = "DROP TABLE tRun2"
                Try
                    _currentDb.ExecuteNonQuery(sSQL, True)
                Catch ex As Exception

                End Try
                sSQL = "DROP TABLE tRunDiff"
                Try
                    _currentDb.ExecuteNonQuery(sSQL, True)
                Catch ex As Exception

                End Try
                sSQL = "DROP TABLE tRunDiff1"
                Try
                    _currentDb.ExecuteNonQuery(sSQL, True)
                Catch ex As Exception

                End Try
                sSQL = "DROP TABLE tRunDiff2"
                Try
                    _currentDb.ExecuteNonQuery(sSQL, True)
                Catch ex As Exception

                End Try
                sSQL = "DROP TABLE tRunDiffDetail_Temp"
                Try
                    _currentDb.ExecuteNonQuery(sSQL, True)
                Catch ex As Exception

                End Try
                sSQL = "DROP TABLE tRunDiffDetail"
                Try
                    _currentDb.ExecuteNonQuery(sSQL, True)
                Catch ex As Exception

                End Try

                'Dim ds As New DataSet("RunComp")
                'Import vsFile1 into tRun1
                Try ' 'Import vsFile1 into tRun1
                    ds = New DataSet("RunComp")
                    'ds.ReadXml(vsFile1)
                    rs1 = New ADODB.Recordset
                    rs1.Open(vsFile1)
                Catch ex As Exception  'Import vsFile1 into tRun1
                    'MsgBox(ex.Message)
                    GenUtils.Message(MsgType.Critical, "Run Comparison", ex.Message)
                End Try  'Import vsFile1 into tRun1 'ds.ReadXml(vsFile1)

                'ds.Tables(0).WriteXml(vsFile2)
                'ds.Dispose()
                'Check if there are any records in the recordset
                'If ds.Tables(0).Rows.Count = 0 Then
                If rs1.RecordCount = 0 Then
                    'error
                Else
                    'Make sure the recordset has the ONLY "key" field: key_UniqueCode
                    nKeyFieldCtr = 0
                    For lCtr = 0 To rs1.Fields.Count - 1
                        'If ds.Tables(0).Columns(lCtr).Name = "KEY_UNIQUECODE" Then
                        'If ds.Tables(0).Columns(lCtr).ColumnName.ToUpper().Equals("KEY_UNIQUECODE") Then
                        If rs1.Fields(lCtr).Name.ToUpper().Equals("KEY_UNIQUECODE") Then
                            nKeyFieldCtr = nKeyFieldCtr + 1
                            Exit For
                        End If
                        'End If
                    Next
                    If nKeyFieldCtr <> 1 Then
                        MsgBox("The file:" & vbNewLine & vsFile1 & vbNewLine & "is missing the key fields: key_UniqueCode" & vbNewLine & "Please make sure you have the key field before you ""Run Comparison"" again.", vbExclamation)
                    Else
                        'dt = ds.Tables.Add("tRun1")
                        'sSQL = "CREATE TABLE [tRun1] ( "
                        'sSQL = "SELECT 'SCN01' AS Field1, 'SCN02' AS Field2, 8.90 AS Field3 INTO tRun1"
                        'iRet = _currentDb.ExecuteNonQuery(sSQL)
                        sSQL = "SELECT "
                        sSQL2 = "INSERT INTO tRun1 ("
                        For lCtr = 0 To rs1.Fields.Count - 1
                            'dt.Columns.Add(ds.Tables(0).Columns(lCtr).ColumnName, ds.Tables(0).Columns(lCtr).DataType)
                            'sSQL = sSQL & ds.Tables(0).Columns(lCtr).ColumnName & ","
                            'sSQL2 = sSQL2 & ds.Tables(0).Columns(lCtr).ColumnName & ","
                            sSQL2 = sSQL2 & rs1.Fields(lCtr).Name & ","
                            If rs1.Fields(lCtr).Name.ToUpper().StartsWith("TXT_") Or rs1.Fields(lCtr).Name.ToUpper().StartsWith("KEY_") Or rs1.Fields(lCtr).Name.ToUpper().StartsWith("SUBKEY_") Then
                                sSQL = sSQL & "'TEXT' AS " & rs1.Fields(lCtr).Name & ","
                            Else
                                sSQL = sSQL & "9.00000000000000001E-13 AS " & rs1.Fields(lCtr).Name & ","
                            End If
                        Next
                        sSQL = Left(sSQL, Len(sSQL) - 1)
                        sSQL = sSQL & " INTO tRun1"
                        Try
                            iRet = _currentDb.ExecuteNonQuery(sSQL)
                            If iRet = -1 Then
                                EntLib.COPT.Log.Log("Error executing SQL in RunComparison." & vbNewLine & sSQL)
                            End If
                            sSQL = "DELETE * FROM tRun1"
                            iRet = _currentDb.ExecuteNonQuery(sSQL)
                            If iRet = -1 Then
                                EntLib.COPT.Log.Log("Error executing SQL in RunComparison." & vbNewLine & sSQL)
                            End If
                        Catch ex As Exception
                            'MsgBox(ex.Message)
                            GenUtils.Message(MsgType.Critical, "Run Comparison", ex.Message)
                        End Try

                        sSQL2 = Left(sSQL2, Len(sSQL2) - 1)
                        sSQL2 = sSQL2 & ") VALUES ("


                        'ds.WriteXml("C:\OPTMODELS\CBK35\Output\Test\ds.xml", XmlWriteMode.IgnoreSchema)
                        'iRet = _currentDb.UpdateDataSet(ds, "tRun1")

                        dt = _currentDb.GetDataTable("SELECT * FROM tRun1")
                        dt.TableName = "tRun1"
                        ds.Tables.Add(dt.Copy)

                        For Each dr In ds.Tables(0).Rows
                            ds.Tables("tRun1").ImportRow(dr)
                        Next
                        'ds.Tables.Remove("Table")
                        _currentDb.UpdateDataSet(ds.Tables("tRun1"), "SELECT * FROM tRun1")
                        dt = _currentDb.GetDataTable("SELECT DISTINCT txt_RUN_NAME FROM tRun1")
                        runName1 = dt.Rows(0).Item(0).ToString()

                        'For lCtr = 0 To ds.Tables(0).Rows.Count - 1
                        '    sSQL = ""

                        '    'sSQL = "INSERT INTO [tRun1] ("
                        '    'sSQL = "SELECT "
                        '    'For lCtr2 = 0 To ds.Tables(0).Columns.Count - 1
                        '    '    'sSQL = sSQL & "[" & ds.Tables(0).Columns(lCtr2).ColumnName & "],"
                        '    '    If ds.Tables(0).Columns(lCtr2).ColumnName.ToUpper().StartsWith("TXT_") Or ds.Tables(0).Columns(lCtr2).ColumnName.ToUpper().StartsWith("KEY_") Or ds.Tables(0).Columns(lCtr2).ColumnName.ToUpper().StartsWith("SUBKEY_") Then
                        '    '        sSQL = sSQL & "'TEXT' AS " & ds.Tables(0).Columns(lCtr2).ColumnName & ","
                        '    '    Else
                        '    '        sSQL = sSQL & "0.0001 AS " & ds.Tables(0).Columns(lCtr2).ColumnName & ","
                        '    '    End If
                        '    'Next
                        '    'sSQL = Left(sSQL, Len(sSQL) - 1)
                        '    ''sSQL = sSQL & ") VALUES ("
                        '    'sSQL = sSQL & " INTO tRun1"
                        '    'Try
                        '    '    iRet = _currentDb.ExecuteNonQuery(sSQL)
                        '    '    sSQL = "DELETE * FROM tRun1"
                        '    '    iRet = _currentDb.ExecuteNonQuery(sSQL)
                        '    'Catch ex As Exception

                        '    'End Try

                        '    'sSQL = "INSERT INTO [tRun1] ("

                        '    For lCtr2 = 0 To ds.Tables(0).Columns.Count - 1
                        '        sTmpName = ds.Tables(0).Columns(lCtr2).Caption  'oRS.Fields(lCtr2).Name
                        '        'sTemp = sFormatSQLString(oRS.Fields(lCtr2).Value & "")
                        '        sTemp = sFormatSQLString(ds.Tables(0).Rows(lCtr).Item(lCtr2) & "")
                        '        If sTemp = "" Then
                        '            sTmpValue = 0
                        '        Else
                        '            sTmpValue = sTemp
                        '        End If
                        '        If UCase(Left(sTmpName, 4)) = "KEY_" Or UCase(Left(sTmpName, 4)) = "SUB_" Or UCase(Left(sTmpName, 4)) = "SUBK" Or UCase(Left(sTmpName, 4)) = "TXT_" Then
                        '            sSQL = sSQL & "'" & sTmpValue & "'"
                        '        Else
                        '            sSQL = sSQL & sTmpValue
                        '        End If
                        '        sSQL = sSQL & ","
                        '    Next
                        '    sSQL = Left(sSQL, Len(sSQL) - 1)
                        '    sSQL = sSQL & ")"

                        '    sSQL = sSQL2 & sSQL

                        '    iRet = _currentDb.ExecuteNonQuery(sSQL)
                        'Next

                        ds = Nothing

                        'Import vsFile2 into tRun2
                        ds = New DataSet("RunComp")
                        Try
                            'ds.ReadXml(vsFile2)
                            rs2 = New ADODB.Recordset
                            rs2.Open(vsFile2)
                        Catch ex As Exception
                            'MsgBox(ex.Message)
                            GenUtils.Message(MsgType.Critical, "Run Comparison", ex.Message)
                        End Try


                        'Check if there are any records in the recordset
                        If rs2.RecordCount = 0 Then
                        Else

                            'Make sure the recordset has the ONLY "key" field: key_UniqueCode
                            nKeyFieldCtr = 0
                            For lCtr = 0 To rs2.Fields.Count - 1
                                'If ds.Tables(0).Columns(lCtr).Name = "KEY_UNIQUECODE" Then
                                If rs2.Fields(lCtr).Name.ToUpper().Equals("KEY_UNIQUECODE") Then
                                    nKeyFieldCtr = nKeyFieldCtr + 1
                                    Exit For
                                End If
                                'End If
                            Next
                            If nKeyFieldCtr <> 1 Then
                                MsgBox("The file:" & vbNewLine & vsFile2 & vbNewLine & "is missing the key fields: key_UniqueCode" & vbNewLine & "Please make sure you have the key field before you ""Run Comparison"" again.", vbExclamation)
                            Else
                                'dt = ds.Tables.Add("tRun2")
                                'For lCtr = 0 To ds.Tables(0).Columns.Count - 1
                                '    dt.Columns.Add(ds.Tables(0).Columns(lCtr).ColumnName, ds.Tables(0).Columns(lCtr).GetType())
                                'Next

                                sSQL = "SELECT "
                                sSQL2 = "INSERT INTO tRun2 ("
                                For lCtr = 0 To rs2.Fields.Count - 1
                                    'dt.Columns.Add(ds.Tables(0).Columns(lCtr).ColumnName, ds.Tables(0).Columns(lCtr).DataType)
                                    'sSQL = sSQL & ds.Tables(0).Columns(lCtr).ColumnName & ","
                                    sSQL2 = sSQL2 & rs2.Fields(lCtr).Name & ","
                                    If rs2.Fields(lCtr).Name.ToUpper().StartsWith("TXT_") Or rs2.Fields(lCtr).Name.ToUpper().StartsWith("KEY_") Or rs2.Fields(lCtr).Name.ToUpper().StartsWith("SUBKEY_") Then
                                        sSQL = sSQL & "'TEXT' AS " & rs2.Fields(lCtr).Name & ","
                                    Else
                                        sSQL = sSQL & "9.00000000000000001E-13 AS " & rs2.Fields(lCtr).Name & ","
                                    End If
                                Next
                                sSQL = Left(sSQL, Len(sSQL) - 1)
                                sSQL = sSQL & " INTO tRun2"
                                Try
                                    iRet = _currentDb.ExecuteNonQuery(sSQL)
                                    If iRet = -1 Then
                                        EntLib.COPT.Log.Log("Error executing SQL in RunComparison." & vbNewLine & sSQL)
                                    End If
                                    sSQL = "DELETE * FROM tRun2"
                                    iRet = _currentDb.ExecuteNonQuery(sSQL)
                                    If iRet = -1 Then
                                        EntLib.COPT.Log.Log("Error executing SQL in RunComparison." & vbNewLine & sSQL)
                                    End If
                                Catch ex As Exception
                                    'MsgBox(ex.Message)
                                    GenUtils.Message(MsgType.Critical, "Run Comparison", ex.Message)
                                End Try

                                sSQL2 = Left(sSQL2, Len(sSQL2) - 1)
                                sSQL2 = sSQL2 & ") VALUES ("


                                dt = _currentDb.GetDataTable("SELECT * FROM tRun2")
                                dt.TableName = "tRun2"
                                ds.Tables.Add(dt)

                                For Each dr In ds.Tables(0).Rows
                                    ds.Tables("tRun2").ImportRow(dr)
                                Next
                                _currentDb.UpdateDataSet(ds.Tables("tRun2"), "SELECT * FROM tRun2")
                                dt = _currentDb.GetDataTable("SELECT DISTINCT txt_RUN_NAME FROM tRun2")
                                runName2 = dt.Rows(0).Item(0).ToString()

                                'For lCtr = 0 To ds.Tables(0).Rows.Count - 1
                                '    sSQL = ""

                                '    For lCtr2 = 0 To ds.Tables(0).Columns.Count - 1
                                '        sTmpName = ds.Tables(0).Columns(lCtr2).Caption  'oRS.Fields(lCtr2).Name
                                '        'sTemp = sFormatSQLString(oRS.Fields(lCtr2).Value & "")
                                '        sTemp = sFormatSQLString(ds.Tables(0).Rows(lCtr).Item(lCtr2) & "")
                                '        If sTemp = "" Then
                                '            sTmpValue = 0
                                '        Else
                                '            sTmpValue = sTemp
                                '        End If
                                '        If UCase(Left(sTmpName, 4)) = "KEY_" Or UCase(Left(sTmpName, 4)) = "SUB_" Or UCase(Left(sTmpName, 4)) = "SUBK" Or UCase(Left(sTmpName, 4)) = "TXT_" Then
                                '            sSQL = sSQL & "'" & sTmpValue & "'"
                                '        Else
                                '            sSQL = sSQL & sTmpValue
                                '        End If
                                '        sSQL = sSQL & ","
                                '    Next
                                '    sSQL = Left(sSQL, Len(sSQL) - 1)
                                '    sSQL = sSQL & ")"

                                '    sSQL = sSQL2 & sSQL

                                '    iRet = _currentDb.ExecuteNonQuery(sSQL)
                                'Next

                                'Run a diff between the fields of tRun1 and tRun2
                                'LEFT JOIN
                                sSQL = "SELECT TOP 1 * FROM tRun2"
                                dt = _currentDb.GetDataTable(sSQL)

                                sSQL = "SELECT "
                                For lCtr = 0 To dt.Columns.Count - 1
                                    If dt.Columns(lCtr).Caption.ToUpper.StartsWith("KEY_") Or _
                                        dt.Columns(lCtr).Caption.ToUpper.StartsWith("SUB_") Or _
                                        dt.Columns(lCtr).Caption.ToUpper.StartsWith("SUBKEY_") Then

                                        sSQL = sSQL & "tRun2." & dt.Columns(lCtr).Caption & ","
                                    ElseIf dt.Columns(lCtr).Caption.ToUpper.StartsWith("TXT_") Then
                                        sSQL = sSQL & "tRun2." & dt.Columns(lCtr).Caption & " AS " & dt.Columns(lCtr).Caption & "_2," & "tRun1." & dt.Columns(lCtr).Caption & " AS " & dt.Columns(lCtr).Caption & "_1,"
                                    Else
                                        sSQL = sSQL & "IIF(ISNULL(tRun2." & dt.Columns(lCtr).Caption & "), 0, tRun2." & dt.Columns(lCtr).Caption & ") AS " & dt.Columns(lCtr).Caption & "_2,"
                                        sSQL = sSQL & "IIf(IsNull(tRun1." & dt.Columns(lCtr).Caption & "), 0, tRun1." & dt.Columns(lCtr).Caption & ") AS " & dt.Columns(lCtr).Caption & "_1,"

                                    End If
                                Next
                                sSQL = Left(sSQL, Len(sSQL) - 1)
                                sSQL = sSQL & " INTO tRunDiff1"
                                sSQL = sSQL & " FROM "
                                sSQL = sSQL & "tRun2 LEFT JOIN tRun1 ON "
                                sSQL = sSQL & "tRun2.key_UniqueCode = tRun1.key_UniqueCode "
                                Try
                                    iRet = _currentDb.ExecuteNonQuery(sSQL)
                                    If iRet = -1 Then
                                        EntLib.COPT.Log.Log("Error executing SQL in RunComparison." & vbNewLine & sSQL)
                                    End If
                                Catch ex As Exception
                                    'MsgBox(ex.Message)
                                    GenUtils.Message(MsgType.Critical, "Run Comparison", ex.Message)
                                End Try


                                'RIGHT JOIN
                                sSQL = "SELECT TOP 1 * FROM tRun1"
                                dt = _currentDb.GetDataTable(sSQL)
                                sSQL = "SELECT "
                                For lCtr = 0 To dt.Columns.Count - 1
                                    If dt.Columns(lCtr).Caption.ToUpper.StartsWith("KEY_") Or _
                                        dt.Columns(lCtr).Caption.ToUpper.StartsWith("SUB_") Or _
                                        dt.Columns(lCtr).Caption.ToUpper.StartsWith("SUBKEY_") Then

                                        sSQL = sSQL & "tRun1." & dt.Columns(lCtr).Caption & ","
                                    ElseIf dt.Columns(lCtr).Caption.ToUpper.StartsWith("TXT_") Then
                                        sSQL = sSQL & "tRun2." & dt.Columns(lCtr).Caption & " AS " & dt.Columns(lCtr).Caption & "_2," & "tRun1." & dt.Columns(lCtr).Caption & " AS " & dt.Columns(lCtr).Caption & "_1,"
                                    Else
                                        sSQL = sSQL & "IIF(ISNULL(tRun2." & dt.Columns(lCtr).Caption & "), 0, tRun2." & dt.Columns(lCtr).Caption & ") AS " & dt.Columns(lCtr).Caption & "_2,"
                                        sSQL = sSQL & "IIf(IsNull(tRun1." & dt.Columns(lCtr).Caption & "), 0, tRun1." & dt.Columns(lCtr).Caption & ") AS " & dt.Columns(lCtr).Caption & "_1,"

                                    End If
                                Next
                                sSQL = Left(sSQL, Len(sSQL) - 1)
                                sSQL = sSQL & " INTO tRunDiff2 "
                                sSQL = sSQL & "FROM "
                                sSQL = sSQL & "tRun1 LEFT JOIN tRun2 ON "
                                sSQL = sSQL & "tRun1.key_UniqueCode = tRun2.key_UniqueCode "
                                sSQL = sSQL & "WHERE "
                                sSQL = sSQL & "tRun2.key_UniqueCode IS NULL "
                                Try
                                    iRet = _currentDb.ExecuteNonQuery(sSQL)
                                    If iRet = -1 Then
                                        EntLib.COPT.Log.Log("Error executing SQL in RunComparison." & vbNewLine & sSQL)
                                    End If
                                Catch ex As Exception
                                    'MsgBox(ex.Message)
                                    GenUtils.Message(MsgType.Critical, "Run Comparison", ex.Message)
                                End Try


                                'UNION
                                sSQL = "SELECT * INTO tRunDiff FROM (SELECT * FROM tRunDiff1 UNION SELECT * FROM tRunDiff2)"
                                Try
                                    iRet = _currentDb.ExecuteNonQuery(sSQL)
                                    If iRet = -1 Then
                                        EntLib.COPT.Log.Log("Error executing SQL in RunComparison." & vbNewLine & sSQL)
                                    End If
                                Catch ex As Exception
                                    'MsgBox(ex.Message)
                                    GenUtils.Message(MsgType.Critical, "Run Comparison", ex.Message)
                                End Try


                                'DIFF - TEMP
                                sSQL = "SELECT TOP 1 * FROM tRun2"
                                dt = _currentDb.GetDataTable(sSQL)
                                sSQL = "SELECT "
                                For lCtr = 0 To dt.Columns.Count - 1
                                    If dt.Columns(lCtr).Caption.ToUpper.StartsWith("KEY_") Or _
                                        dt.Columns(lCtr).Caption.ToUpper.StartsWith("SUB_") Or _
                                        dt.Columns(lCtr).Caption.ToUpper.StartsWith("SUBKEY_") Then

                                        sSQL = sSQL & dt.Columns(lCtr).Caption & ","
                                    ElseIf dt.Columns(lCtr).Caption.ToUpper.StartsWith("TXT_") Then
                                        sSQL = sSQL & dt.Columns(lCtr).Caption & "_2," & dt.Columns(lCtr).Caption & "_1,"
                                    Else
                                        sSQL = sSQL & dt.Columns(lCtr).Caption & "_2," & dt.Columns(lCtr).Caption & "_1," & dt.Columns(lCtr).Caption & "_2-" & dt.Columns(lCtr).Caption & "_1 AS dif_" & dt.Columns(lCtr).Caption & ","
                                    End If
                                Next
                                sSQL = Left(sSQL, Len(sSQL) - 1)
                                sSQL = sSQL & " INTO tRunDiffDetail_Temp "
                                sSQL = sSQL & "FROM "
                                sSQL = sSQL & "tRunDiff"
                                Try
                                    iRet = _currentDb.ExecuteNonQuery(sSQL)
                                    If iRet = -1 Then
                                        EntLib.COPT.Log.Log("Error executing SQL in RunComparison." & vbNewLine & sSQL)
                                    End If
                                Catch ex As Exception
                                    'MsgBox(ex.Message)
                                    GenUtils.Message(MsgType.Critical, "Run Comparison", ex.Message)
                                End Try


                                'DIFF - FINAL
                                sSQL = "SELECT TOP 1 * FROM tRunDiffDetail_Temp"
                                dt = _currentDb.GetDataTable(sSQL)
                                sSQL = "SELECT "
                                For lCtr = 0 To dt.Columns.Count - 1
                                    If dt.Columns(lCtr).Caption.ToUpper.StartsWith("KEY_") Or _
                                        dt.Columns(lCtr).Caption.ToUpper.StartsWith("SUB_") Or _
                                        dt.Columns(lCtr).Caption.ToUpper.StartsWith("SUBKEY_") Then

                                        sSQL = sSQL & dt.Columns(lCtr).Caption & ","
                                    ElseIf dt.Columns(lCtr).Caption.ToUpper.EndsWith("_2") Then
                                        sSQL = sSQL & dt.Columns(lCtr).Caption & ","
                                    End If
                                Next
                                For lCtr = 0 To dt.Columns.Count - 1
                                    If dt.Columns(lCtr).Caption.ToUpper.StartsWith("KEY_") Or _
                                        dt.Columns(lCtr).Caption.ToUpper.StartsWith("SUB_") Or _
                                        dt.Columns(lCtr).Caption.ToUpper.StartsWith("SUBKEY_") Then

                                        'do nothing
                                    ElseIf Right(dt.Columns(lCtr).Caption, 2) = "_1" Then
                                        sSQL = sSQL & dt.Columns(lCtr).Caption & ","
                                    End If
                                Next
                                For lCtr = 0 To dt.Columns.Count - 1
                                    If dt.Columns(lCtr).Caption.ToUpper.StartsWith("DIF_") Then
                                        sTemp = Mid(dt.Columns(lCtr).Caption, 5)
                                        sSQL = sSQL & dt.Columns(lCtr).Caption & " AS diff_" & Mid(sTemp, InStr(1, sTemp, "_", vbTextCompare) + 1) & ","
                                    End If
                                Next
                                sSQL = Left(sSQL, Len(sSQL) - 1)
                                sSQL = sSQL & " INTO tRunDiffDetail "
                                sSQL = sSQL & "FROM "
                                sSQL = sSQL & "tRunDiffDetail_Temp"
                                Try
                                    iRet = _currentDb.ExecuteNonQuery(sSQL)
                                    If iRet = -1 Then
                                        EntLib.COPT.Log.Log("Error executing SQL in RunComparison." & vbNewLine & sSQL)
                                    End If
                                Catch ex As Exception
                                    'MsgBox(ex.Message)
                                    GenUtils.Message(MsgType.Critical, "Run Comparison", ex.Message)
                                End Try


                                If vbDeleteAllTempTables Then
                                    sSQL = "DROP TABLE tRun1"
                                    Try
                                        _currentDb.ExecuteNonQuery(sSQL)
                                    Catch ex As Exception

                                    End Try
                                    sSQL = "DROP TABLE tRun2"
                                    Try
                                        _currentDb.ExecuteNonQuery(sSQL)
                                    Catch ex As Exception

                                    End Try
                                    sSQL = "DROP TABLE tRunDiff"
                                    Try
                                        _currentDb.ExecuteNonQuery(sSQL)
                                    Catch ex As Exception

                                    End Try
                                    sSQL = "DROP TABLE tRunDiff1"
                                    Try
                                        _currentDb.ExecuteNonQuery(sSQL)
                                    Catch ex As Exception

                                    End Try
                                    sSQL = "DROP TABLE tRunDiff2"
                                    Try
                                        _currentDb.ExecuteNonQuery(sSQL)
                                    Catch ex As Exception

                                    End Try
                                    sSQL = "DROP TABLE tRunDiffDetail_Temp"
                                    Try
                                        _currentDb.ExecuteNonQuery(sSQL)
                                    Catch ex As Exception

                                    End Try
                                End If

                                'Save tRunDiffDetail as a CSV file
                                ds = Nothing
                                dt = _currentDb.GetDataTable("SELECT * FROM tRunDiffDetail")
                                dt.TableName = "Table"
                                ds = New DataSet("NewDataSet")
                                ds.Tables.Add(dt)
                                'file = workDir & "\" & miscParams.Item("PRJ_NAME").ToString() & "-" & miscParams.Item("RUN_NAME").ToString() & "-" & dt.Rows(ctr).Item("TextAbbr").ToString() & "-" & timeStamp '& ".xml"

                                file = workDir & "\" & vsDBName & "-[" & runName1 & "--" & runName2 & "]-" & vsFile1.Substring(0, vsFile1.Length - 18).Split(CChar("-"))(vsFile1.Substring(0, vsFile1.Length - 18).Split(CChar("-")).Length - 1) & "DIFF-" & My.Computer.Clock.LocalTime.Year().ToString() & My.Computer.Clock.LocalTime.Month().ToString().PadLeft(2, CChar("0")) & My.Computer.Clock.LocalTime.Day().ToString().PadLeft(2, CChar("0")) & "-" & My.Computer.Clock.LocalTime.Hour.ToString().PadLeft(2, CChar("0")) & My.Computer.Clock.LocalTime.Minute.ToString().PadLeft(2, CChar("0"))
                                ds.WriteXml(file & ".xml.tmp")

                                'Create the XsltSettings object with script enabled.
                                Dim settings As New XsltSettings(False, True)

                                'Execute the transform.
                                Dim xslt As New System.Xml.Xsl.XslCompiledTransform()
                                Try
                                    xslt.Load(workDir & "\DataSetToCSV.xsl", settings, New XmlUrlResolver)
                                Catch ex As Exception
                                    Try
                                        xslt.Load(My.Application.Info.DirectoryPath & "\DataSetToCSV.xsl", settings, New XmlUrlResolver)
                                    Catch ex1 As Exception
                                        'MsgBox(ex1.Message)
                                        GenUtils.Message(MsgType.Critical, "Run Comparison", ex1.Message)
                                    End Try
                                End Try
                                'dsTemp.WriteXml(file & ".xml.tmp")
                                xslt.Transform(file & ".xml.tmp", file & ".csv.tmp")

                                xslt = Nothing

                                Dim sr As New IO.StreamReader(file & ".csv.tmp")
                                'Dim content As String = sr.ReadToEnd


                                Dim sw As New IO.StreamWriter(file & ".csv", False)
                                Dim delim As String = ""
                                Dim delimiter As String = ","
                                For Each col As DataColumn In dt.Columns

                                    sw.Write(delim)
                                    sw.Write(col.ColumnName)
                                    delim = delimiter
                                Next
                                sw.WriteLine()
                                'sw.Write(content)
                                sw.Write(sr.ReadToEnd)
                                sw.Close()
                                sw.Dispose()
                                sr.Close()
                                sr.Dispose()

                                Try
                                    My.Computer.FileSystem.DeleteFile(file & ".csv.tmp")
                                Catch ex As Exception

                                End Try
                                Try
                                    My.Computer.FileSystem.DeleteFile(file & ".xml.tmp")
                                Catch ex As Exception

                                End Try

                                'rsOutput = "Diff Compare file:" & vbNewLine & file & ".csv" & vbNewLine & "has been created at:" & vbNewLine & workDir
                                rsOutput = file & ".csv"

                            End If 'If nKeyFieldCtr <> 1 Then 'tRun2
                        End If
                    End If 'If nKeyFieldCtr <> 1 Then 'tRun1
                End If 'If ds.Tables(0).Rows.Count = 0 Then
            End If 'If bContinue = True Then

            COPT.GenUtils.FormatTime(startTime, My.Computer.Clock.TickCount)

        End Sub

        Public Function sFormatSQLString(ByVal vsSQLStr As String) As String
            '**********************************************
            'Author  :  Ravi Poluri
            'Date/Ver:  04-22-05/v1.5.41
            'Input   :
            'Output  :
            'Comments:
            '**********************************************

            'sFormatSQLString = VBA.Replace(vsSQLStr, "'", "", , , vbTextCompare)
            sFormatSQLString = Replace(vsSQLStr, "'", "''", , , vbTextCompare)

        End Function

        Public Function GetInstanceName() As String
            Return My.Application.Info.AssemblyName
        End Function

        '' Create a custom section.
        'Public Sub Chg()
        '    Dim customSectionName As String
        '    ' Get the application configuration file.
        '    Dim config As System.Configuration.Configuration = _
        '    ConfigurationManager.OpenExeConfiguration( _
        '    ConfigurationUserLevel.None)
        '    ' Console.WriteLine(config.FilePath);
        '    ' If the section does not exiat in the configuration
        '    ' file, create it and save it to the file.
        '    If config.Sections(customSectionName) Is Nothing Then
        '        custSection = New CustomSection()
        '        config.Sections.Add(customSectionName, custSection)
        '        custSection = config.GetSection(customSectionName)
        '        custSection.SectionInformation.ForceSave = True
        '        config.Save(ConfigurationSaveMode.Full)
        '    End If
        'End Sub 'New


        'Public Function CmdLineParse(ByVal cmdLine As String) As Boolean
        '    'Dim parser As New CommandLineParser()

        '    'CmdLineEntries(parser)

        '    'parser.CommandLine = cmdLine
        '    'If parser.Parse() Then
        '    '    Return True
        '    'Else
        '    '    Return False
        '    'End If
        'End Function


        'Sub CmdLineEntries(ByVal parser _
        '   As CommandLineParser)

        '    Dim anEntry As CommandLineEntry
        '    Dim nextEntry As CommandLineEntry

        '    parser.Errors.Clear()
        '    parser.Entries.Clear()

        '    '' RUN
        '    '' create an entry that accepts:
        '    '' RUN PRJ BB4
        '    'anEntry = parser.CreateEntry(CommandLineParse.CommandTypeEnum.Value, "RUN")
        '    'anEntry.Required = False
        '    'parser.Entries.Add(anEntry)
        '    '' store the new Entry in a local reference
        '    '' for use with the next CommandLineEntry's 
        '    '' MustFollow property.
        '    'nextEntry = anEntry
        '    '' now create am ExistingFile type entry that must
        '    '' follow the command.
        '    'anEntry = parser.CreateEntry(CommandTypeEnum.Value, "PRJ")
        '    'anEntry.MustFollowEntry = nextEntry
        '    'anEntry.Required = True
        '    'parser.Entries.Add(anEntry)
        '    'nextEntry = anEntry
        '    '' now create am ExistingFile type entry that must
        '    '' follow the previous argument.
        '    'anEntry = parser.CreateEntry(CommandTypeEnum.Value)
        '    'anEntry.MustFollowEntry = nextEntry
        '    'anEntry.Required = True
        '    'parser.Entries.Add(anEntry)

        '    ' HELP
        '    ' create an entry that accepts:
        '    ' HELP
        '    anEntry = parser.CreateEntry(CommandLineParse.CommandTypeEnum.Value, "HELP")
        '    anEntry.Required = False
        '    parser.Entries.Add(anEntry)

        '    ' RUN
        '    ' create a flag type entry that accepts:
        '    ' RUN PRJ BB4
        '    anEntry = parser.CreateEntry(CommandLineParse.CommandTypeEnum.Value, "RUN")
        '    anEntry.Required = False
        '    parser.Entries.Add(anEntry)
        '    ' store the new Entry in a local reference
        '    ' for use with the next CommandLineEntry's 
        '    ' MustFollow property.
        '    nextEntry = anEntry
        '    ' now create am ExistingFile type entry that must
        '    ' follow the command.
        '    anEntry = parser.CreateEntry(CommandTypeEnum.Value, "PRJ")
        '    anEntry.MustFollowEntry = nextEntry
        '    anEntry.Required = True
        '    parser.Entries.Add(anEntry)
        '    nextEntry = anEntry
        '    ' now create am ExistingFile type entry that must
        '    ' follow the previous argument.
        '    anEntry = parser.CreateEntry(CommandTypeEnum.Value)
        '    anEntry.MustFollowEntry = nextEntry
        '    anEntry.Required = True
        '    parser.Entries.Add(anEntry)
        'End Sub

        Public Shared Function CmdFormatArguments(ByVal cmd As String) As String()
            Dim input() As String
            Dim output() As String
            Dim len As Short
            Dim tmp As String = String.Empty
            Dim arg As Boolean = False

            input = cmd.Split(CChar(" "))
            ReDim output(0)
            len = -1
            For ctr As Short = 0 To CShort(input.Length - 1)
                If input(ctr).ToString().Trim.StartsWith("""") Then
                    tmp = String.Empty
                    If input(ctr).ToString().Trim.EndsWith("""") Then
                        arg = False
                        If input(ctr).ToString().Trim <> "" Then
                            len = CShort(len + 1)
                            ReDim Preserve output(len)
                            'output(len) = input(ctr).ToString().Trim
                            If input(ctr).ToString().Trim.StartsWith("""") Then
                                output(len) = input(ctr).ToString().Trim().Substring(1, input(ctr).ToString().Trim().Length - 2)
                            Else
                                'output(len) = input(ctr).ToString().Trim
                                If input(ctr).ToString().Trim().StartsWith("""") Then
                                    output(len) = input(ctr).ToString().Trim().Substring(1, input(ctr).ToString().Trim().Length - 2)
                                Else
                                    output(len) = input(ctr).ToString().Trim
                                End If
                            End If
                            tmp = String.Empty
                        End If
                    Else
                        arg = True
                        tmp = tmp & input(ctr).ToString().Trim & " "
                    End If
                Else
                    If input(ctr).ToString().Trim <> "" Then
                        If input(ctr).ToString().Trim.EndsWith("""") Then
                            len = CShort(len + 1)
                            ReDim Preserve output(len)
                            tmp = tmp & input(ctr).ToString().Trim & " "
                            output(len) = tmp.Trim().Substring(1, tmp.Trim().Length - 2) 'input(ctr).ToString().Trim
                            tmp = String.Empty
                            arg = False
                        Else
                            If arg = False Then
                                len = CShort(len + 1)
                                ReDim Preserve output(len)
                                If input(ctr).ToString().Trim.StartsWith("""") Then
                                    output(len) = input(ctr).ToString().Trim().Substring(1, input(ctr).ToString().Trim().Length - 2)
                                Else
                                    output(len) = input(ctr).ToString().Trim
                                End If

                                tmp = String.Empty
                            Else
                                tmp = tmp & input(ctr).ToString().Trim & " "
                            End If
                        End If
                    End If
                End If
            Next

            ''Remove the ""
            'For ctr As Short = 0 To output.Length - 1
            '    If output(ctr).StartsWith("""""") Then
            '        output(ctr) = output(ctr).Substring(2, output(ctr).Length - 2)
            '    End If
            'Next

            Return output
        End Function

        'Public Sub SaveRun_Old(ByRef dt As DataTable, ByRef miscParams As DataRow, ByVal workDir As String, ByVal saveXML As Boolean)
        '    'Dim timeStamp As String = Microsoft.VisualBasic.Format(Now(), "yyyymmdd-hhmm")
        '    Dim timeStamp As String = ""
        '    'Dim dt As DataTable = currentDB.GetDataTable("SELECT * FROM " & COPT_RSLT_TABLE_NAME & " WHERE ACTIVE = True")
        '    Dim dtTemp As DataTable = Nothing
        '    Dim dsTemp As DataSet = Nothing
        '    Dim file As String
        '    Dim delimiter As String = ","
        '    Dim delim As String = ""
        '    Dim startTime As Integer = My.Computer.Clock.TickCount
        '    Dim flagCSV As Boolean = False
        '    Dim flagXML As Boolean = False
        '    'Dim writer As TextWriter = Nothing

        '    Const FORMAT_CSV As Short = 2
        '    Const FORMAT_XML As Short = 4

        '    Try
        '        timeStamp = My.Computer.Clock.LocalTime.Year().ToString() & My.Computer.Clock.LocalTime.Month().ToString().PadLeft(2, CChar("0")) & My.Computer.Clock.LocalTime.Day().ToString().PadLeft(2, CChar("0")) & "-" & My.Computer.Clock.LocalTime.Hour.ToString().PadLeft(2, CChar("0")) & My.Computer.Clock.LocalTime.Minute.ToString().PadLeft(2, CChar("0"))
        '        For ctr As Short = 0 To CShort(dt.Rows.Count - 1)
        '            'Debug.Print(dt.Rows(ctr).Item("RecordsetName").ToString())
        '            'dt.WriteXml(_workDir & "\" & _miscParams.Item("PRJ_NAME").ToString() & "-" & _miscParams.Item("RUN_NAME").ToString() & "-" & dt.Rows(ctr).Item("TextAbbr").ToString() & "-" & Microsoft.VisualBasic.Format(Now(), "yyyymmdd-hhmm") & ".xml")
        '            'dtTemp = _currentDb.GetDataTable("SELECT * FROM [" & dt.Rows(ctr).Item("RecordsetName").ToString() & "]")
        '            dsTemp = _currentDb.GetDataSet("SELECT * FROM [" & dt.Rows(ctr).Item("RecordsetName").ToString() & "]")
        '            dtTemp = dsTemp.Tables(0)
        '            file = workDir & "\" & miscParams.Item("PRJ_NAME").ToString() & "-" & miscParams.Item("RUN_NAME").ToString() & "-" & dt.Rows(ctr).Item("TextAbbr").ToString() & "-" & timeStamp '& ".xml"

        '            If saveXML Then
        '                'dtTemp.WriteXml(file & ".datatable.xml")
        '                dsTemp.WriteXml(file & ".xml")
        '                'dsTemp.WriteXml(writer)

        '            End If

        '            If dt.Rows(ctr).Item("OutputFormat") = FORMAT_CSV + FORMAT_XML Then
        '                flagCSV = True
        '                flagXML = True
        '            End If
        '            If dt.Rows(ctr).Item("OutputFormat") = FORMAT_CSV Then
        '                flagCSV = True
        '            End If
        '            If dt.Rows(ctr).Item("OutputFormat") = FORMAT_XML Then
        '                'dtTemp.WriteXml(file & ".xml")
        '                flagXML = True
        '            End If

        '            If flagXML = True Then
        '                'dtTemp.WriteXml(file & ".xml")
        '                'dtTemp.WriteXml(file & ".datatable.xml")
        '                dsTemp.WriteXml(file & ".xml")
        '                ConvertToRecordset(dtTemp).Save(file & ".ado.xml", ADODB.PersistFormatEnum.adPersistXML)
        '            End If

        '            If flagCSV = True Then


        '                Dim sw As New IO.StreamWriter(file & ".csv")
        '                delim = ""
        '                For Each col As DataColumn In dtTemp.Columns
        '                    sw.Write(delim)
        '                    sw.Write(col.ColumnName)
        '                    delim = delimiter
        '                Next
        '                sw.WriteLine()
        '                For Each row As DataRow In dtTemp.Rows
        '                    Try
        '                        'Dim sb As New StringBuilder
        '                        'For rowCtr As Short = 0 To dtTemp.Columns.Count - 1
        '                        '    sb.Append(row(rowCtr).ToString())
        '                        '    sb.Append(",")
        '                        'Next
        '                        sw.WriteLine(Join(row.ItemArray, ","))
        '                        'sw.WriteLine(sb.ToString())
        '                    Catch ex As Exception
        '                        'MsgBox(ex.Message)
        '                        'An exception occurs when there is a NULL in one or more columns.
        '                        'Use the traditional loop to include the row with one or more NULL column values.
        '                        Dim sb As New StringBuilder
        '                        For rowCtr As Short = 0 To dtTemp.Columns.Count - 1
        '                            sb.Append(row(rowCtr).ToString())
        '                            sb.Append(",")
        '                        Next
        '                        'sw.WriteLine(Join(row.ItemArray, ","))
        '                        sw.WriteLine(sb.ToString().Substring(0, sb.ToString().Length - 1))
        '                    End Try
        '                Next
        '                sw.Close()
        '                sw.Dispose()
        '            End If
        '        Next

        '        'cleanup
        '        'dt.Dispose()
        '        dtTemp.Dispose()

        '        'RunComparison("C:\OPTMODELS\BR1\New Folder\BR1-BBB-MillProdSourcing-20081803-0118.xml", "C:\OPTMODELS\BR1\New Folder\BR1-BBB-MillProdSourcing-20081803-0118-Out.xml", "", False)
        '    Catch ex As Exception
        '        MsgBox(ex.Message)
        '    End Try
        '    EntLib.Log.Log(workDir, "EntLib - GenUtils - SaveRun took: ", EntLib.GenUtils.FormatTime(startTime, My.Computer.Clock.TickCount))
        'End Sub

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="dt"></param>
        ''' <param name="miscParams"></param>
        ''' <param name="workDir"></param>
        ''' <param name="saveADOXML"></param>
        ''' <param name="saveXML"></param>
        ''' <param name="useXSLT"></param>
        ''' <remarks>
        ''' How to use XML Literals (in VB9) to output CSV files.
        ''' http://www.emoreau.com/Entries/Articles/2009/04/Using-LINQ-and-XML-Literals-to-transform-a-DataTable-into-a-HTML-table.aspx
        ''' https://msmvps.com/blogs/deborahk/archive/2009/08/14/using-linq-with-microsoft-word-and-excel.aspx
        ''' </remarks>
        Public Sub SaveRun(ByRef dt As DataTable, _
                           ByRef miscParams As DataRow, _
                           ByVal workDir As String, _
                           ByVal timeStamp As String, _
                           ByVal saveADOXML As Boolean, _
                           ByVal saveXML As Boolean, _
                           ByVal useXSLT As Boolean)
            'Dim timeStamp As String = Microsoft.VisualBasic.Format(Now(), "yyyymmdd-hhmm")
            'Dim timeStamp As String = ""
            'Dim dt As DataTable = currentDB.GetDataTable("SELECT * FROM " & COPT_RSLT_TABLE_NAME & " WHERE ACTIVE = True")
            Dim dtTemp As DataTable = Nothing
            Dim dsTemp As DataSet = Nothing
            Dim file As String
            Dim delimiter As String = ""","""
            Dim delim As String = ""
            Dim startTime As Integer = My.Computer.Clock.TickCount
            Dim startQueryTime As Integer = My.Computer.Clock.TickCount
            Dim flagCSV As Boolean = False
            Dim flagXML As Boolean = False
            'Dim writer As TextWriter = Nothing

            Const FORMAT_CSV As Short = 2
            Const FORMAT_XML As Short = 4
            'Const FORMAT_HTML As Short = 8

            Try
                timeStamp = GenUtils.GetTimeStamp() 'My.Computer.Clock.LocalTime.Year().ToString() & My.Computer.Clock.LocalTime.Month().ToString().PadLeft(2, CChar("0")) & My.Computer.Clock.LocalTime.Day().ToString().PadLeft(2, CChar("0")) & "-" & My.Computer.Clock.LocalTime.Hour.ToString().PadLeft(2, CChar("0")) & My.Computer.Clock.LocalTime.Minute.ToString().PadLeft(2, CChar("0"))
                Console.WriteLine("Saving model results...")
                For ctr As Integer = 0 To CInt(dt.Rows.Count - 1)
                    'Debug.Print(dt.Rows(ctr).Item("RecordsetName").ToString())
                    'dt.WriteXml(_workDir & "\" & _miscParams.Item("PRJ_NAME").ToString() & "-" & _miscParams.Item("RUN_NAME").ToString() & "-" & dt.Rows(ctr).Item("TextAbbr").ToString() & "-" & Microsoft.VisualBasic.Format(Now(), "yyyymmdd-hhmm") & ".xml")
                    'dtTemp = _currentDb.GetDataTable("SELECT * FROM [" & dt.Rows(ctr).Item("RecordsetName").ToString() & "]")
                    startQueryTime = My.Computer.Clock.TickCount

                    'Console.WriteLine("SaveRun..." & dt.Rows(ctr).Item("TextAbbr").ToString())

                    dsTemp = _currentDb.GetDataSet("SELECT * FROM [" & dt.Rows(ctr).Item("RecordsetName").ToString() & "]")
                    'EntLib.Log.Log(workDir, "EntLib - GenUtils - SaveRun - " & dt.Rows(ctr).Item("RecordsetName").ToString() & " - " & dt.Rows(ctr).Item("TextAbbr").ToString() & " : ", EntLib.GenUtils.FormatTime(startQueryTime, My.Computer.Clock.TickCount))

                    'RKP/08-07-08
                    'This extra looping is done to prevent NULL columns from sliding over to the left in the CSV files.


                    'startQueryTime = My.Computer.Clock.TickCount
                    dtTemp = dsTemp.Tables(0)
                    file = workDir & "\" & miscParams.Item("PRJ_NAME").ToString() & "-" & miscParams.Item("RUN_NAME").ToString() & "-" & dt.Rows(ctr).Item("TextAbbr").ToString() & "-" & timeStamp '& ".xml"
                    Console.WriteLine(file)

                    If saveXML Then
                        'dtTemp.WriteXml(file & ".datatable.xml")
                        'dsTemp.WriteXml(file & ".xml")
                        'dsTemp.WriteXml(writer)
                        dtTemp.WriteXmlSchema(file & ".schema.xml")
                        dtTemp.WriteXml(file & ".xml")
                    End If

                    flagCSV = False
                    flagXML = False
                    If CType(dt.Rows(ctr).Item("OutputFormat"), Short) = FORMAT_CSV + FORMAT_XML Then
                        flagCSV = True
                        flagXML = True
                    End If
                    If CType(dt.Rows(ctr).Item("OutputFormat"), Short) = FORMAT_CSV Then
                        flagCSV = True
                        flagXML = False
                    End If
                    If CType(dt.Rows(ctr).Item("OutputFormat"), Short) = FORMAT_XML Then
                        'dtTemp.WriteXml(file & ".xml")
                        flagXML = True
                        flagCSV = False
                    End If

                    If flagXML = True Then
                        'dtTemp.WriteXml(file & ".xml")
                        'dtTemp.WriteXml(file & ".datatable.xml")
                        'dsTemp.WriteXml(file & ".xml", XmlWriteMode.WriteSchema)
                        dtTemp.WriteXmlSchema(file & ".schema.xml")
                        dtTemp.WriteXml(file & ".xml")
                        If saveADOXML Then
                            ConvertToRecordset(dtTemp).Save(file & ".ado.xml", ADODB.PersistFormatEnum.adPersistXML)
                        End If
                    End If

                    If flagCSV = True Then
                        If useXSLT Then
                            'Create the XsltSettings object with script enabled.
                            Dim settings As New XsltSettings(False, True)

                            'Execute the transform.
                            Dim xslt As New System.Xml.Xsl.XslCompiledTransform()
                            Try
                                xslt.Load(workDir & "\DataSetToCSV.xsl", settings, New XmlUrlResolver)
                            Catch ex As Exception
                                Try
                                    xslt.Load(My.Application.Info.DirectoryPath & "\DataSetToCSV.xsl", settings, New XmlUrlResolver)
                                Catch ex1 As Exception
                                    'MsgBox(ex1.Message)
                                    Try
                                        xslt.Load("C:\OPTMODELS\C-OPTSYS\DataSetToCSV.xsl", settings, New XmlUrlResolver)
                                    Catch ex2 As Exception
                                        'MsgBox(ex2.Message)
                                        GenUtils.Message(MsgType.Warning, "Save Run (CSV file)", ex2.Message)
                                    End Try
                                End Try
                            End Try
                            dsTemp.WriteXml(file & ".xml.tmp")
                            xslt.Transform(file & ".xml.tmp", file & ".csv.tmp")

                            xslt = Nothing

                            Dim sr As New IO.StreamReader(file & ".csv.tmp")
                            Dim content As String = sr.ReadToEnd
                            sr.Close()
                            sr.Dispose()

                            Dim sw As New IO.StreamWriter(file & ".csv", False)
                            delim = """"
                            For Each col As DataColumn In dtTemp.Columns

                                sw.Write(delim)
                                sw.Write(col.ColumnName)

                                delim = delimiter
                            Next
                            sw.Write("""")
                            sw.WriteLine()
                            sw.Write(content)
                            sw.Close()
                            sw.Dispose()

                            Try
                                My.Computer.FileSystem.DeleteFile(file & ".csv.tmp")
                            Catch ex As Exception

                            End Try
                            Try
                                My.Computer.FileSystem.DeleteFile(file & ".xml.tmp")
                            Catch ex As Exception
                            End Try

                            'Dim fs As New FileStream(file & ".csv", FileMode.Open, FileAccess.ReadWrite)
                            'Dim pos As Long = fs.Seek(0, SeekOrigin.End)

                        Else
                            Dim sw As New IO.StreamWriter(file & ".csv")
                            delim = ""
                            For Each col As DataColumn In dtTemp.Columns
                                sw.Write(delim)
                                sw.Write(col.ColumnName)
                                delim = delimiter
                            Next
                            sw.WriteLine()
                            For Each row As DataRow In dtTemp.Rows
                                Try
                                    'Dim sb As New StringBuilder
                                    'For rowCtr As Short = 0 To dtTemp.Columns.Count - 1
                                    '    sb.Append(row(rowCtr).ToString())
                                    '    sb.Append(",")
                                    'Next
                                    sw.WriteLine(Join(row.ItemArray, ","))
                                    'sw.WriteLine(sb.ToString())
                                Catch ex As Exception
                                    'MsgBox(ex.Message)
                                    'An exception occurs when there is a NULL in one or more columns.
                                    'Use the traditional loop to include the row with one or more NULL column values.
                                    Dim sb As New StringBuilder
                                    For rowCtr As Integer = 0 To CInt(dtTemp.Columns.Count - 1)
                                        sb.Append(row(rowCtr).ToString())
                                        sb.Append(",")
                                    Next
                                    'sw.WriteLine(Join(row.ItemArray, ","))
                                    sw.WriteLine(sb.ToString().Substring(0, sb.ToString().Length - 1))
                                End Try
                            Next
                            sw.Close()
                            sw.Dispose()
                        End If
                    End If
                    COPT.Log.Log(workDir, "EntLib - GenUtils - SaveRun - " & dt.Rows(ctr).Item("RecordsetName").ToString() & " - " & dt.Rows(ctr).Item("TextAbbr").ToString() & " - to CSV/XML: ", COPT.GenUtils.FormatTime(startQueryTime, My.Computer.Clock.TickCount))
                Next

                'cleanup
                'dt.Dispose()
                If Not dtTemp Is Nothing Then
                    dtTemp.Dispose()
                End If

                'RunComparison("C:\OPTMODELS\BR1\New Folder\BR1-BBB-MillProdSourcing-20081803-0118.xml", "C:\OPTMODELS\BR1\New Folder\BR1-BBB-MillProdSourcing-20081803-0118-Out.xml", "", False)
            Catch ex As Exception
                'MsgBox(ex.Message)
                'GenUtils.Message(MsgType.Warning, "Save Run (CSV file)", ex.Message)
                EntLib.COPT.Log.Log(workDir, "EntLib - GenUtils - Error saving CSV file: ", ex.Message)
                Console.WriteLine("EntLib - GenUtils - Error saving CSV file: ", ex.Message)
            End Try
            COPT.Log.Log(workDir, "EntLib - GenUtils - SaveRun took: ", COPT.GenUtils.FormatTime(startTime, My.Computer.Clock.TickCount))
        End Sub

        Public Shared Function GetAppSettings(ByVal key As String) As String
            Try
                Return GetSysConfigValue(key) 'ConfigurationManager.AppSettings(key)
            Catch ex As Exception
                Return ""
            End Try
        End Function

        Public Shared Function GetUserName() As String
            Dim Ret As Integer
            Dim UserName As String
            Dim Buffer As String
            Buffer = New String(CChar(" "), 25)
            Try
                Ret = GetUserName(Buffer, 25)
                UserName = Left(Buffer, InStr(Buffer, Chr(0)) - 1)
                Return UserName
            Catch ex As Exception
                Return ""
            End Try

        End Function

        Public Shared Function IsSwitchAvailable(ByVal switches() As String, ByVal switch As String) As Boolean
            IsSwitchAvailable = False
            Try
                For ctr As Short = 0 To CShort(UBound(switches)) '- 1
                    If switch.Trim().ToUpper().Equals(switches(ctr).Trim().ToUpper()) Then
                        IsSwitchAvailable = True
                        Exit For
                    End If
                Next
            Catch ex As Exception

            End Try

        End Function

        Public Shared Function GetSwitchArgument(ByVal switches() As String, ByVal switch As String, ByVal offset As Short) As String
            GetSwitchArgument = ""
            If switches Is Nothing Then

            Else
                'If switches.Length > 0 Then


                For ctr As Short = 0 To CShort(UBound(switches)) '- 1
                    If switch.Trim().ToUpper().Equals(switches(ctr).Trim().ToUpper()) Then
                        GetSwitchArgument = switches(ctr + offset).ToString()
                        Exit For
                        'Else
                        '    Return ""
                    End If
                Next
            End If
            If GetSwitchArgument = "" Then
                If switch.Trim.ToUpper() = "/WORKDIR" Then
                    GetSwitchArgument = GetAppSettings("defaultOptModelRootFolder") & "\" & GetSwitchArgument(switches, "/PRJ", 1) & "\" & GetAppSettings("defaultOptModelOutputFolder")
                Else
                    GetSwitchArgument = GetAppSettings(switch)
                End If

            End If
        End Function

        Public Shared Sub ConvertXMLToCSV(ByVal xmlFile As String, ByVal csvFile As String)
            Dim xslt As New Xsl.XslCompiledTransform()

        End Sub

        Public Shared Function GetWorkDir(ByVal switches() As String) As String

            If switches Is Nothing Then
                Return GenUtils.GetSysConfigValue("defaultOptModelRootFolder") & "\" & GenUtils.GetSysConfigValue("defaultProject") & "\" & GenUtils.GetSysConfigValue("defaultOptModelOutputFolder")
            Else

                Dim folderRoot As String = COPT.GenUtils.GetSwitchArgument(switches, "defaultOptModelRootFolder", 1)

                If Not folderRoot.Trim().EndsWith("\") Then
                    folderRoot = folderRoot.Trim() & "\"
                End If

                If Not switches Is Nothing Then
                    If IsSwitchAvailable(switches, "/WorkDir") Then
                        Return GetSwitchArgument(switches, "/WorkDir", 1)
                        'Return folderRoot & ConfigurationManager.AppSettings.Item("defaultProject").ToString() & "\" & EntLib.GenUtils.GetSwitchArgument(switches, "defaultOptModelOutputFolder", 1)
                    Else
                        If IsSwitchAvailable(switches, "/PRJ") Then
                            Return folderRoot & COPT.GenUtils.GetSwitchArgument(switches, "/PRJ", 1) & "\" & COPT.GenUtils.GetSwitchArgument(switches, "defaultOptModelOutputFolder", 1)
                        Else
                            'Return folderRoot & ConfigurationManager.AppSettings.Item("defaultProject").ToString() & "\" & COPT.GenUtils.GetSwitchArgument(switches, "defaultOptModelOutputFolder", 1)
                            Return folderRoot & GetSysConfigValue("defaultProject").ToString() & "\" & COPT.GenUtils.GetSwitchArgument(switches, "defaultOptModelOutputFolder", 1)
                        End If

                    End If
                Else
                    'Return folderRoot & ConfigurationManager.AppSettings.Item("defaultProject").ToString() & "\" & COPT.GenUtils.GetSwitchArgument(switches, "defaultOptModelOutputFolder", 1)
                    Return folderRoot & GetSysConfigValue("defaultProject").ToString() & "\" & COPT.GenUtils.GetSwitchArgument(switches, "defaultOptModelOutputFolder", 1)
                End If
            End If
        End Function

        Public Shared Function ConvertToRecordset(ByVal inTable As DataTable) As ADODB.Recordset
            Dim result As New ADODB.Recordset()

            result.CursorLocation = ADODB.CursorLocationEnum.adUseClient

            Dim resultFields As ADODB.Fields = result.Fields
            Dim inColumns As System.Data.DataColumnCollection = inTable.Columns

            For Each inColumn As DataColumn In inColumns
                resultFields.Append(inColumn.ColumnName, TranslateType(inColumn.DataType), inColumn.MaxLength, CType(IIf(inColumn.AllowDBNull, ADODB.FieldAttributeEnum.adFldIsNullable, ADODB.FieldAttributeEnum.adFldUnspecified), ADODB.FieldAttributeEnum), Nothing)
                'resultFields.Append(inColumn.ColumnName, TranslateType(inColumn.DataType), inColumn.MaxLength, IIf(inColumn.AllowDBNull, ADODB.FieldAttributeEnum.adFldIsNullable, ADODB.FieldAttributeEnum.adFldUnspecified), Nothing)
            Next

            result.Open(System.Reflection.Missing.Value, System.Reflection.Missing.Value, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic, 0)

            For Each dr As DataRow In inTable.Rows
                result.AddNew(System.Reflection.Missing.Value, System.Reflection.Missing.Value)
                For columnIndex As Integer = 0 To inColumns.Count - 1

                    resultFields(columnIndex).Value = dr(columnIndex)
                Next
            Next

            Return result
        End Function

        Private Shared Function TranslateType(ByVal columnType As Type) As ADODB.DataTypeEnum

            Select Case columnType.UnderlyingSystemType.ToString()
                Case "System.Boolean"
                    Return ADODB.DataTypeEnum.adBoolean
                Case "System.Byte"

                    Return ADODB.DataTypeEnum.adUnsignedTinyInt
                Case "System.Char"

                    Return ADODB.DataTypeEnum.adChar
                Case "System.DateTime"

                    Return ADODB.DataTypeEnum.adDate
                Case "System.Decimal"

                    Return ADODB.DataTypeEnum.adCurrency
                Case "System.Double"

                    Return ADODB.DataTypeEnum.adDouble
                Case "System.Int16"

                    Return ADODB.DataTypeEnum.adSmallInt
                Case "System.Int32"

                    Return ADODB.DataTypeEnum.adInteger
                Case "System.Int64"

                    Return ADODB.DataTypeEnum.adBigInt
                Case "System.SByte"

                    Return ADODB.DataTypeEnum.adTinyInt
                Case "System.Single"

                    Return ADODB.DataTypeEnum.adSingle
                Case "System.UInt16"

                    Return ADODB.DataTypeEnum.adUnsignedSmallInt
                Case "System.UInt32"

                    Return ADODB.DataTypeEnum.adUnsignedInt
                Case "System.UInt64"

                    Return ADODB.DataTypeEnum.adUnsignedBigInt
                Case "System.String"
                    'Case Else

                    Return ADODB.DataTypeEnum.adVarChar
                Case Else
                    Return ADODB.DataTypeEnum.adVarChar
            End Select
        End Function

        ''' <summary>
        ''' Generates a C-OPT blueprint file (*.c-opt.blp.xml).
        ''' </summary>
        ''' <author>RKP</author>
        ''' <date>12-01-08</date>
        ''' <version>2.0.115</version>
        ''' <remarks>
        ''' This method generates a blueprint file that allows a modeler to recreate a project.
        ''' The blueprint file has just metadata information. There is no data saved in it.
        ''' NOTE: Need to add System.Data.DataSetExtensions in order to get "AsEnumerable()" extension to a DataSet.
        ''' Link to MSDN LINQ samples:
        ''' http://msdn.microsoft.com/en-us/vbasic/bb688088.aspx
        ''' </remarks>
        Public Shared Function GenerateBlueprint(ByVal switches() As String, ByRef ds As DataSet) As Integer
            Dim prjName As String = GetSwitchArgument(switches, "/PRJ", 1)
            Dim fileName As String = GetSwitchArgument(switches, "/WorkDir", 1) & "\" & prjName + ".c-opt.blp.xml"
            'Dim dtDefCol As DataTable = Nothing
            'Dim dtDefRow As DataTable = Nothing
            'Dim dtDefCoef As DataTable = Nothing
            'Dim dtTables As DataTable = Nothing
            'Dim dtQueries As DataTable = Nothing

            Dim tbl = ds.Tables("defCol").AsEnumerable()
            Dim values = From rows In tbl Select rows(2)
            For Each value In values
                Console.WriteLine(value)
            Next

            'Dim tbl = ds.Tables("defCol").AsEnumerable()
            Dim xmlBlueprint = _
                <blueprint>
                    <defTables>
                        <defCol>
                            <%= From rows In tbl _
                                Select _
                                <row ID=<%= rows("ID") %>>
                                    <ColActive><%= rows("ColActive") %></ColActive>
                                    <ColType><%= rows("ColType") %></ColType>
                                    <ColTypeDesc><%= rows("ColTypeDesc") %></ColTypeDesc>
                                    <ColTypeTable><%= rows("ColTypeTable") %></ColTypeTable>
                                    <ColTypeRecSet><%= rows("ColTypeRecSet") %></ColTypeRecSet>
                                    <ColTypePrefix><%= rows("ColTypePrefix") %></ColTypePrefix>
                                    <ColDescField><%= rows("ColDescField") %></ColDescField>
                                    <OBJField><%= rows("OBJField") %></OBJField>
                                    <SOSType><%= rows("SOSType") %></SOSType>
                                    <SOSMarkerField><%= rows("SOSMarkerField") %></SOSMarkerField>
                                    <BNDFree><%= rows("BNDFree") %></BNDFree>
                                    <BNDInteger><%= rows("BNDInteger") %></BNDInteger>
                                    <BNDBinary><%= rows("BNDBinary") %></BNDBinary>
                                    <BNDLoField><%= rows("BNDLoField") %></BNDLoField>
                                    <BNDUpField><%= rows("BNDUpField") %></BNDUpField>
                                    <ClassCount><%= rows("ClassCount") %></ClassCount>
                                    <ClassConcat><%= rows("ClassConcat") %></ClassConcat>
                                    <C1><%= rows("C1") %></C1>
                                    <C2><%= rows("C2") %></C2>
                                    <C3><%= rows("C3") %></C3>
                                    <C4><%= rows("C4") %></C4>
                                    <C5><%= rows("C5") %></C5>
                                    <C6><%= rows("C6") %></C6>
                                    <C7><%= rows("C7") %></C7>
                                    <C8><%= rows("C8") %></C8>
                                    <C9><%= rows("C9") %></C9>
                                    <C10><%= rows("C10") %></C10>
                                    <C11><%= rows("C11") %></C11>
                                    <C12><%= rows("C12") %></C12>
                                    <C13><%= rows("C13") %></C13>
                                    <C14><%= rows("C14") %></C14>
                                    <C15><%= rows("C15") %></C15>
                                    <C16><%= rows("C16") %></C16>
                                </row> %>
                            %>
                        </defCol>
                    </defTables>
                </blueprint>

            Try
                xmlBlueprint.Save(fileName)
            Catch ex As Exception
                'MsgBox(ex.Message)
                GenUtils.Message(MsgType.Critical, "Generate Blueprint", ex.Message)
            End Try



            'dtDefCol = _currentDb.GetDataTable("SELECT * FROM tsysDefCol")

            'Dim blp = _
            '    <?xml version="1.0"?>
            '    <!--C-OPT Blueprint File Version 0.1 (Beta) 12-01-08-->
            '    <blueprint>
            '        <defTables>
            '            <defCol>
            '                <%= From column In ds.Tables("defCol").Columns _
            '                    Select _
            '                <id=<%= ds.Tables("defCol").Columns("ID").ColumnName%>>
            '                %>
            '            </defCol>
            '    </blueprint>

            'Dim xmlOrders = _
            '<orders>
            '    <%= From Customer In ds.Tables(0) _
            '        Where Customer(0) = "USA" _
            '        Select _
            '        <customer id=<%= Customer.CustomerID %>>
            '            <name><%= Customer.ContactName %></name>
            '            <address><%= Customer.Address %></address>
            '            <city><%= Customer.City %></city>
            '            <zip><%= Customer.PostalCode %></zip>
            '            <orders>
            '                <%= From Order In Customer.Orders _
            '                    Select _
            '                    <order total=<%= _
            '                                     Aggregate Detail In Order.Order_Details _
            '                                     Into Sum(Detail.UnitPrice * Detail.Quantity) %>>
            '                        <date><%= Order.OrderDate %></date>
            '                        <details>
            '                            <%= From Detail In Order.Order_Details _
            '                                Select _
            '                                <product id=<%= Detail.ProductID %>>
            '                                    <name>
            '                                        <%= Detail.Product.ProductName %>
            '                                    </name>
            '                                </product> %>
            '                        </details>
            '                    </order> %>
            '            </orders>
            '        </customer> %>
            '</orders>

        End Function

        Public Sub SaveRunUsingLINQ_Old(ByRef dt As DataTable, _
                           ByRef miscParams As DataRow, _
                           ByVal workDir As String, _
                           ByVal saveADOXML As Boolean, _
                           ByVal saveXML As Boolean, _
                           ByVal useXSLT As Boolean)
            'Dim timeStamp As String = Microsoft.VisualBasic.Format(Now(), "yyyymmdd-hhmm")
            Dim timeStamp As String = ""
            'Dim dt As DataTable = currentDB.GetDataTable("SELECT * FROM " & COPT_RSLT_TABLE_NAME & " WHERE ACTIVE = True")
            Dim dtTemp As DataTable = Nothing
            Dim dsTemp As DataSet = Nothing
            Dim file As String
            Dim delimiter As String = ""","""
            Dim delim As String = ""
            Dim startTime As Integer = My.Computer.Clock.TickCount
            Dim startQueryTime As Integer = My.Computer.Clock.TickCount
            Dim flagCSV As Boolean = False
            Dim flagXML As Boolean = False
            'Dim writer As TextWriter = Nothing

            Const FORMAT_CSV As Short = 2
            Const FORMAT_XML As Short = 4

            Try
                timeStamp = My.Computer.Clock.LocalTime.Year().ToString() & My.Computer.Clock.LocalTime.Month().ToString().PadLeft(2, CChar("0")) & My.Computer.Clock.LocalTime.Day().ToString().PadLeft(2, CChar("0")) & "-" & My.Computer.Clock.LocalTime.Hour.ToString().PadLeft(2, CChar("0")) & My.Computer.Clock.LocalTime.Minute.ToString().PadLeft(2, CChar("0"))
                For ctr As Short = 0 To CShort(dt.Rows.Count - 1)
                    'Debug.Print(dt.Rows(ctr).Item("RecordsetName").ToString())
                    'dt.WriteXml(_workDir & "\" & _miscParams.Item("PRJ_NAME").ToString() & "-" & _miscParams.Item("RUN_NAME").ToString() & "-" & dt.Rows(ctr).Item("TextAbbr").ToString() & "-" & Microsoft.VisualBasic.Format(Now(), "yyyymmdd-hhmm") & ".xml")
                    'dtTemp = _currentDb.GetDataTable("SELECT * FROM [" & dt.Rows(ctr).Item("RecordsetName").ToString() & "]")
                    startQueryTime = My.Computer.Clock.TickCount
                    dsTemp = _currentDb.GetDataSet("SELECT * FROM [" & dt.Rows(ctr).Item("RecordsetName").ToString() & "]")
                    'EntLib.Log.Log(workDir, "EntLib - GenUtils - SaveRun - " & dt.Rows(ctr).Item("RecordsetName").ToString() & " - " & dt.Rows(ctr).Item("TextAbbr").ToString() & " : ", EntLib.GenUtils.FormatTime(startQueryTime, My.Computer.Clock.TickCount))

                    'RKP/08-07-08
                    'This extra looping is done to prevent NULL columns from sliding over to the left in the CSV files.


                    startQueryTime = My.Computer.Clock.TickCount
                    dtTemp = dsTemp.Tables(0)
                    file = workDir & "\" & miscParams.Item("PRJ_NAME").ToString() & "-" & miscParams.Item("RUN_NAME").ToString() & "-" & dt.Rows(ctr).Item("TextAbbr").ToString() & "-" & timeStamp '& ".xml"

                    If saveXML Then
                        'dtTemp.WriteXml(file & ".datatable.xml")
                        dsTemp.WriteXml(file & ".xml")
                        'dsTemp.WriteXml(writer)

                    End If

                    flagCSV = False
                    flagXML = False
                    If CType(dt.Rows(ctr).Item("OutputFormat"), Short) = FORMAT_CSV + FORMAT_XML Then
                        flagCSV = True
                        flagXML = True
                    End If
                    If CType(dt.Rows(ctr).Item("OutputFormat"), Short) = FORMAT_CSV Then
                        flagCSV = True
                        flagXML = False
                    End If
                    If CType(dt.Rows(ctr).Item("OutputFormat"), Short) = FORMAT_XML Then
                        'dtTemp.WriteXml(file & ".xml")
                        flagXML = True
                        flagCSV = False
                    End If

                    If flagXML = True Then
                        'dtTemp.WriteXml(file & ".xml")
                        'dtTemp.WriteXml(file & ".datatable.xml")
                        dsTemp.WriteXml(file & ".xml", XmlWriteMode.WriteSchema)
                        If saveADOXML Then
                            ConvertToRecordset(dtTemp).Save(file & ".ado.xml", ADODB.PersistFormatEnum.adPersistXML)
                        End If
                    End If

                    If flagCSV = True Then
                        If useXSLT Then
                            'Create the XsltSettings object with script enabled.
                            Dim settings As New XsltSettings(False, True)

                            'Execute the transform.
                            Dim xslt As New System.Xml.Xsl.XslCompiledTransform()
                            Try
                                xslt.Load(workDir & "\DataSetToCSV.xsl", settings, New XmlUrlResolver)
                            Catch ex As Exception
                                Try
                                    xslt.Load(My.Application.Info.DirectoryPath & "\DataSetToCSV.xsl", settings, New XmlUrlResolver)
                                Catch ex1 As Exception
                                    'MsgBox(ex1.Message)
                                    Try
                                        xslt.Load("C:\OPTMODELS\C-OPTSYS\DataSetToCSV.xsl", settings, New XmlUrlResolver)
                                    Catch ex2 As Exception
                                        MsgBox(ex2.Message)
                                    End Try
                                End Try
                            End Try
                            dsTemp.WriteXml(file & ".xml.tmp")
                            xslt.Transform(file & ".xml.tmp", file & ".csv.tmp")

                            xslt = Nothing

                            Dim sr As New IO.StreamReader(file & ".csv.tmp")
                            Dim content As String = sr.ReadToEnd
                            sr.Close()
                            sr.Dispose()

                            Dim sw As New IO.StreamWriter(file & ".csv", False)
                            delim = """"
                            For Each col As DataColumn In dtTemp.Columns

                                sw.Write(delim)
                                sw.Write(col.ColumnName)

                                delim = delimiter
                            Next
                            sw.Write("""")
                            sw.WriteLine()
                            sw.Write(content)
                            sw.Close()
                            sw.Dispose()

                            Try
                                My.Computer.FileSystem.DeleteFile(file & ".csv.tmp")
                            Catch ex As Exception

                            End Try
                            Try
                                My.Computer.FileSystem.DeleteFile(file & ".xml.tmp")
                            Catch ex As Exception
                            End Try

                            'Dim fs As New FileStream(file & ".csv", FileMode.Open, FileAccess.ReadWrite)
                            'Dim pos As Long = fs.Seek(0, SeekOrigin.End)

                        Else
                            Dim sw As New IO.StreamWriter(file & ".csv")
                            delim = ""
                            For Each col As DataColumn In dtTemp.Columns
                                sw.Write(delim)
                                sw.Write(col.ColumnName)
                                delim = delimiter
                            Next
                            sw.WriteLine()
                            For Each row As DataRow In dtTemp.Rows
                                Try
                                    'Dim sb As New StringBuilder
                                    'For rowCtr As Short = 0 To dtTemp.Columns.Count - 1
                                    '    sb.Append(row(rowCtr).ToString())
                                    '    sb.Append(",")
                                    'Next
                                    sw.WriteLine(Join(row.ItemArray, ","))
                                    'sw.WriteLine(sb.ToString())
                                Catch ex As Exception
                                    'MsgBox(ex.Message)
                                    'An exception occurs when there is a NULL in one or more columns.
                                    'Use the traditional loop to include the row with one or more NULL column values.
                                    Dim sb As New StringBuilder
                                    For rowCtr As Short = 0 To CShort(dtTemp.Columns.Count - 1)
                                        sb.Append(row(rowCtr).ToString())
                                        sb.Append(",")
                                    Next
                                    'sw.WriteLine(Join(row.ItemArray, ","))
                                    sw.WriteLine(sb.ToString().Substring(0, sb.ToString().Length - 1))
                                End Try
                            Next
                            sw.Close()
                            sw.Dispose()
                        End If
                    End If
                    'EntLib.Log.Log(workDir, "EntLib - GenUtils - SaveRun - " & dt.Rows(ctr).Item("RecordsetName").ToString() & " - " & dt.Rows(ctr).Item("TextAbbr").ToString() & " - to CSV/XML: ", EntLib.GenUtils.FormatTime(startQueryTime, My.Computer.Clock.TickCount))
                Next

                'cleanup
                'dt.Dispose()
                dtTemp.Dispose()

                'RunComparison("C:\OPTMODELS\BR1\New Folder\BR1-BBB-MillProdSourcing-20081803-0118.xml", "C:\OPTMODELS\BR1\New Folder\BR1-BBB-MillProdSourcing-20081803-0118-Out.xml", "", False)
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            COPT.Log.Log(workDir, "EntLib - GenUtils - SaveRun took: ", COPT.GenUtils.FormatTime(startTime, My.Computer.Clock.TickCount))
        End Sub

        ''' <summary>
        ''' 
        ''' 
        ''' </summary>
        ''' <param name="dt"></param>
        ''' <param name="miscParams"></param>
        ''' <param name="workDir"></param>
        ''' <param name="saveADOXML"></param>
        ''' <param name="saveXML"></param>
        ''' <param name="useXSLT"></param>
        ''' <remarks></remarks>
        Public Sub SaveRunUsingLINQ(ByRef dt As DataTable, _
                           ByRef miscParams As DataRow, _
                           ByVal workDir As String, _
                           ByVal timeStamp As String, _
                           ByVal saveADOXML As Boolean, _
                           ByVal saveXML As Boolean, _
                           ByVal useXSLT As Boolean)
            'Dim timeStamp As String = Microsoft.VisualBasic.Format(Now(), "yyyymmdd-hhmm")
            'Dim timeStamp As String = ""
            'Dim dt As DataTable = currentDB.GetDataTable("SELECT * FROM " & COPT_RSLT_TABLE_NAME & " WHERE ACTIVE = True")
            Dim dtTemp As DataTable = Nothing
            Dim dsTemp As DataSet = Nothing
            Dim file As String
            Dim delimiter As String = ""","""
            Dim delim As String = ""
            Dim startTime As Integer = My.Computer.Clock.TickCount
            Dim startQueryTime As Integer = My.Computer.Clock.TickCount
            Dim flagCSV As Boolean = False
            Dim flagXML As Boolean = False
            Dim flagHTML As Boolean = False
            'Dim writer As TextWriter = Nothing

            Const FORMAT_CSV As Short = 2
            Const FORMAT_XML As Short = 4
            Const FORMAT_HTML As Short = 8

            Try
                timeStamp = My.Computer.Clock.LocalTime.Year().ToString() & My.Computer.Clock.LocalTime.Month().ToString().PadLeft(2, CChar("0")) & My.Computer.Clock.LocalTime.Day().ToString().PadLeft(2, CChar("0")) & "-" & My.Computer.Clock.LocalTime.Hour.ToString().PadLeft(2, CChar("0")) & My.Computer.Clock.LocalTime.Minute.ToString().PadLeft(2, CChar("0"))
                For ctr As Integer = 0 To CInt(dt.Rows.Count - 1)
                    'Debug.Print(dt.Rows(ctr).Item("RecordsetName").ToString())
                    'dt.WriteXml(_workDir & "\" & _miscParams.Item("PRJ_NAME").ToString() & "-" & _miscParams.Item("RUN_NAME").ToString() & "-" & dt.Rows(ctr).Item("TextAbbr").ToString() & "-" & Microsoft.VisualBasic.Format(Now(), "yyyymmdd-hhmm") & ".xml")
                    'dtTemp = _currentDb.GetDataTable("SELECT * FROM [" & dt.Rows(ctr).Item("RecordsetName").ToString() & "]")
                    startQueryTime = My.Computer.Clock.TickCount

                    Console.WriteLine("SaveRun...")

                    dsTemp = _currentDb.GetDataSet("SELECT * FROM [" & dt.Rows(ctr).Item("RecordsetName").ToString() & "]")
                    'EntLib.Log.Log(workDir, "EntLib - GenUtils - SaveRun - " & dt.Rows(ctr).Item("RecordsetName").ToString() & " - " & dt.Rows(ctr).Item("TextAbbr").ToString() & " : ", EntLib.GenUtils.FormatTime(startQueryTime, My.Computer.Clock.TickCount))

                    'RKP/08-07-08
                    'This extra looping is done to prevent NULL columns from sliding over to the left in the CSV files.


                    'startQueryTime = My.Computer.Clock.TickCount
                    dtTemp = dsTemp.Tables(0)
                    file = workDir & "\" & miscParams.Item("PRJ_NAME").ToString() & "-" & miscParams.Item("RUN_NAME").ToString() & "-" & dt.Rows(ctr).Item("TextAbbr").ToString() & "-" & timeStamp '& ".xml"

                    If saveXML Then
                        'dtTemp.WriteXml(file & ".datatable.xml")
                        dsTemp.WriteXml(file & ".xml")
                        'dsTemp.WriteXml(writer)

                    End If

                    flagCSV = False
                    flagXML = False
                    flagHTML = False

                    If CType(dt.Rows(ctr).Item("OutputFormat"), Short) = FORMAT_CSV Then
                        flagCSV = True
                        flagXML = False
                    End If
                    If CType(dt.Rows(ctr).Item("OutputFormat"), Short) = FORMAT_XML Then
                        flagXML = True
                        flagCSV = False
                    End If
                    If CType(dt.Rows(ctr).Item("OutputFormat"), Short) = FORMAT_HTML Then
                        flagXML = False
                        flagCSV = False
                        flagHTML = True
                    End If
                    If CType(dt.Rows(ctr).Item("OutputFormat"), Short) = FORMAT_CSV + FORMAT_XML Then
                        flagCSV = True
                        flagXML = True
                        flagHTML = False
                    End If
                    If CType(dt.Rows(ctr).Item("OutputFormat"), Short) = FORMAT_CSV + FORMAT_HTML Then
                        flagCSV = True
                        flagXML = False
                        flagHTML = True
                    End If
                    If CType(dt.Rows(ctr).Item("OutputFormat"), Short) = FORMAT_XML + FORMAT_HTML Then
                        flagCSV = False
                        flagXML = True
                        flagHTML = True
                    End If

                    If flagXML = True Then
                        'dtTemp.WriteXml(file & ".xml")
                        'dtTemp.WriteXml(file & ".datatable.xml")
                        dsTemp.WriteXml(file & ".xml", XmlWriteMode.WriteSchema)
                        If saveADOXML Then
                            ConvertToRecordset(dtTemp).Save(file & ".ado.xml", ADODB.PersistFormatEnum.adPersistXML)
                        End If
                    End If

                    If flagCSV = True Then
                        If useXSLT Then
                            'Create the XsltSettings object with script enabled.
                            Dim settings As New XsltSettings(False, True)

                            'Execute the transform.
                            Dim xslt As New System.Xml.Xsl.XslCompiledTransform()
                            Try
                                xslt.Load(workDir & "\DataSetToCSV.xsl", settings, New XmlUrlResolver)
                            Catch ex As Exception
                                Try
                                    xslt.Load(My.Application.Info.DirectoryPath & "\DataSetToCSV.xsl", settings, New XmlUrlResolver)
                                Catch ex1 As Exception
                                    'MsgBox(ex1.Message)
                                    Try
                                        xslt.Load("C:\OPTMODELS\C-OPTSYS\DataSetToCSV.xsl", settings, New XmlUrlResolver)
                                    Catch ex2 As Exception
                                        'MsgBox(ex2.Message)
                                        GenUtils.Message(MsgType.Warning, "Save Run (Using LINQ)", ex2.Message)
                                    End Try
                                End Try
                            End Try
                            dsTemp.WriteXml(file & ".xml.tmp")
                            xslt.Transform(file & ".xml.tmp", file & ".csv.tmp")

                            xslt = Nothing

                            Dim sr As New IO.StreamReader(file & ".csv.tmp")
                            Dim content As String = sr.ReadToEnd
                            sr.Close()
                            sr.Dispose()

                            Dim sw As New IO.StreamWriter(file & ".csv", False)
                            delim = """"
                            For Each col As DataColumn In dtTemp.Columns

                                sw.Write(delim)
                                sw.Write(col.ColumnName)

                                delim = delimiter
                            Next
                            sw.Write("""")
                            sw.WriteLine()
                            sw.Write(content)
                            sw.Close()
                            sw.Dispose()

                            Try
                                My.Computer.FileSystem.DeleteFile(file & ".csv.tmp")
                            Catch ex As Exception

                            End Try
                            Try
                                My.Computer.FileSystem.DeleteFile(file & ".xml.tmp")
                            Catch ex As Exception
                            End Try

                            'Dim fs As New FileStream(file & ".csv", FileMode.Open, FileAccess.ReadWrite)
                            'Dim pos As Long = fs.Seek(0, SeekOrigin.End)

                        Else
                            Dim sw As New IO.StreamWriter(file & ".csv")
                            delim = ""
                            For Each col As DataColumn In dtTemp.Columns
                                sw.Write(delim)
                                sw.Write(col.ColumnName)
                                delim = delimiter
                            Next
                            sw.WriteLine()
                            For Each row As DataRow In dtTemp.Rows
                                Try
                                    'Dim sb As New StringBuilder
                                    'For rowCtr As Short = 0 To dtTemp.Columns.Count - 1
                                    '    sb.Append(row(rowCtr).ToString())
                                    '    sb.Append(",")
                                    'Next
                                    sw.WriteLine(Join(row.ItemArray, ","))
                                    'sw.WriteLine(sb.ToString())
                                Catch ex As Exception
                                    'MsgBox(ex.Message)
                                    'An exception occurs when there is a NULL in one or more columns.
                                    'Use the traditional loop to include the row with one or more NULL column values.
                                    Dim sb As New StringBuilder
                                    For rowCtr As Integer = 0 To CInt(dtTemp.Columns.Count - 1)
                                        sb.Append(row(rowCtr).ToString())
                                        sb.Append(",")
                                    Next
                                    'sw.WriteLine(Join(row.ItemArray, ","))
                                    sw.WriteLine(sb.ToString().Substring(0, sb.ToString().Length - 1))
                                End Try
                            Next
                            sw.Close()
                            sw.Dispose()
                        End If
                    End If
                    COPT.Log.Log(workDir, "EntLib - GenUtils - SaveRun - " & dt.Rows(ctr).Item("RecordsetName").ToString() & " - " & dt.Rows(ctr).Item("TextAbbr").ToString() & " - to CSV/XML: ", COPT.GenUtils.FormatTime(startQueryTime, My.Computer.Clock.TickCount))
                Next

                'cleanup
                'dt.Dispose()
                dtTemp.Dispose()

                'RunComparison("C:\OPTMODELS\BR1\New Folder\BR1-BBB-MillProdSourcing-20081803-0118.xml", "C:\OPTMODELS\BR1\New Folder\BR1-BBB-MillProdSourcing-20081803-0118-Out.xml", "", False)
            Catch ex As Exception
                'MsgBox(ex.Message)
                GenUtils.Message(MsgType.Warning, "Save Run (Using LINQ)", ex.Message)
            End Try
            COPT.Log.Log(workDir, "EntLib - GenUtils - SaveRun took: ", COPT.GenUtils.FormatTime(startTime, My.Computer.Clock.TickCount))
        End Sub

        ''' <summary>
        ''' Formats the look and feel of any message box that is displayed by C-OPTApp or C-OPTConsole.
        ''' </summary>
        ''' <param name="type"></param>
        ''' <param name="msgSource"></param>
        ''' <param name="msg"></param>
        ''' <returns></returns>
        ''' <remarks>RKP/v2.2.125/01-13-10</remarks>
        Public Shared Function Message _
        ( _
            ByVal type As EntLib.COPT.GenUtils.MsgType, _
            ByVal msgSource As String, _
            ByVal msg As String _
        ) As String

            Dim caption As String = "C-OPT Message"
            Dim s As String = "Error source: " & msgSource
            Dim config As System.Configuration.Configuration  'Configuration.Configuration

            'Dim noMsgBox As Boolean = GenUtils.GetSwitchArgument(_switches, "/NoMsgBox", 1)

            's = s & IIf(type = MsgType.Information, "", IIf(type = MsgType.Warning, "Type: Warning (model results not affected)", "Type: Critical (call BMOS)")).ToString()
            'config = System.Configuration.ConfigurationManager.OpenExeConfiguration(System.Configuration.ConfigurationUserLevel.None)
            config = GetSysConfig()
            Select Case type
                Case MsgType.Information
                    'MessageBox.Show(msg, msgSource, MessageBoxButtons.OK, MessageBoxIcon.Information)
                    'MessageBox.Show(s & vbNewLine & vbNewLine & "Message:" & vbNewLine & msg, "C-OPT Message", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    MessageBox.Show(msg, caption, MessageBoxButtons.OK, MessageBoxIcon.Information)
                Case MsgType.Warning
                    EntLib.COPT.Log.Log(GenUtils.GetSysConfigValue("lastWorkDir").ToString(), "C-OPT Message - Warning", msg)
                    MessageBox.Show(s & vbNewLine & "Error type: Warning" & vbNewLine & "Error will affect model results: Most likely not" & vbNewLine & vbNewLine & "Message:" & vbNewLine & msg, caption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Case MsgType.Critical
                    EntLib.COPT.Log.Log(GenUtils.GetSysConfigValue("lastWorkDir").ToString(), "C-OPT Message - Critical", msg)
                    MessageBox.Show(s & vbNewLine & "Error type: Critical" & vbNewLine & "Error will affect model results: Yes" & vbNewLine & "*Contact BMOS for support*" & vbNewLine & vbNewLine & "Message:" & vbNewLine & msg, caption, MessageBoxButtons.OK, MessageBoxIcon.Stop)
                Case Else
                    EntLib.COPT.Log.Log(GenUtils.GetSysConfigValue("lastWorkDir").ToString(), "C-OPT Message - Unknown", msg)
                    MessageBox.Show(msg, msgSource, MessageBoxButtons.OK, MessageBoxIcon.Information)
            End Select
            Return Nothing
        End Function

        ''' <summary>
        ''' Dynamically Launching A WinForm Via Object Factory
        ''' </summary>
        ''' <param name="objectName"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' RKP/01-19-10/v2.2.126
        ''' http://msdn.microsoft.com/en-us/vbasic/ee794614.aspx
        ''' </remarks>
        Public Shared Function CreateAnObject(ByVal objectName As String) As Object
            'Dim asm = [Assembly].GetExecutingAssembly()
            Dim asm = System.Reflection.Assembly.GetExecutingAssembly()
            Dim asmType As Type = asm.GetType(objectName.Trim)
            Dim obj As Object = Nothing

            Try
                obj = Activator.CreateInstance(asmType)
            Catch ex As Exception
                GenUtils.Message(MsgType.Critical, "GenUtils - CreateAnObject", ex.Message)
            End Try

            Return obj
        End Function

        Public Shared Function GetTimeStamp() As String
            Return My.Computer.Clock.LocalTime.Year().ToString() & My.Computer.Clock.LocalTime.Month().ToString().PadLeft(2, CChar("0")) & My.Computer.Clock.LocalTime.Day().ToString().PadLeft(2, CChar("0")) & "-" & My.Computer.Clock.LocalTime.Hour.ToString().PadLeft(2, CChar("0")) & My.Computer.Clock.LocalTime.Minute.ToString().PadLeft(2, CChar("0")) & My.Computer.Clock.LocalTime.Second.ToString().PadLeft(2, CChar("0"))
        End Function

        Public Shared Function GenerateZMatrix(ByVal switches() As String, ByRef ds As DataSet) As Integer
            Dim outputFile As String = GenUtils.GetWorkDir(switches) & "\" & GenUtils.GetSwitchArgument(switches, "/PRJ", 1) & "-" & GenUtils.GetTimeStamp() & ".zMatrix.xlsx"
            Console.WriteLine(outputFile)
            Dim ctr As Integer = 0
            Dim rowCtr As Integer = 0
            Dim endCol As Integer = 0
            Dim endRow As Integer = 0
            Dim col As Integer = 0
            Dim row As Integer = 0
            Dim colType As String = ""
            Dim rowType As String = ""
            Dim arr() As String = Nothing

            Dim linqTable As System.Data.EnumerableRowCollection(Of System.Data.DataRow) = Nothing
            Dim strQueryResults As System.Data.EnumerableRowCollection(Of String) = Nothing

            Const xlPart = 2
            Const xlGeneral = 1
            Const xlBottom = -4107
            Const xlByColumns = 2
            Const xlByRows = 1
            Const xlCenter = -4108
            Const xlContext = -5002
            Const xlDown = -4121
            Const xlFormatFromLeftOrAbove = 0
            Const xlContinuous = 1
            Const xlEdgeBottom = 9
            Const xlEdgeLeft = 7
            Const xlEdgeRight = 10
            Const xlEdgeTop = 8
            Const xlMedium = -4138
            Const xlDash = -4115
            Const xlDashDot = 4
            Const xlDashDotDot = 5
            Const xlDot = -4118
            Const xlDouble = -4119
            Const xlLineStyleNone = -4142
            Const xlSlantDashDot = 13
            Const xlHairline = 1
            Const xlThick = 4
            Const xlThin = 2
            Const xlDiagonalDown = 5
            Const xlDiagonalUp = 6
            Const xlInsideHorizontal = 12
            Const xlInsideVertical = 11
            Const xlLandscape = 2
            Const xlPaper11x17 = 17
            Const xlPaperLegal = 5
            Const xlPaperLetter = 1
            Const xlPaperLetterSmall = 2
            Const xlAutomatic = -4105
            Const xlDownThenOver = 1

            Const zmtxColumnStart = 11  'worksheet column H -- go over 11 cols in excel to start the matrix cols
            Const zmtxXClearRows = 4    'worksheet will have 3 clear rows at the top and start in row 4
            Const zmtxRowStart = 22     'worksheet row 22   -- go down 22 rows in excel to start the matrix rows

            Try
                Dim rng As Excel.Range
                Dim xl As New Excel.Application
                xl.Visible = True
                Dim wb As Excel.Workbook = xl.Workbooks.Add
                Dim ws As Excel.Worksheet = CType(wb.Worksheets.Add, Excel.Worksheet)
                ws.Name = "zMatrix"

                '// Write the Columns/Decision Variables/Vectors
                '========================================================================
                rowCtr = zmtxColumnStart
                For ctr = 0 To ds.Tables("defCol").Rows.Count - 1
                    Application.DoEvents()
                    endCol = rowCtr
                    ws.Cells(zmtxXClearRows, rowCtr) = ds.Tables("defCol").Rows(ctr).Item("ColType")
                    ws.Cells(zmtxXClearRows + 1, rowCtr) = ds.Tables("defCol").Rows(ctr).Item("ColTypeDesc")
                    ws.Cells(zmtxXClearRows + 2, rowCtr) = ds.Tables("defCol").Rows(ctr).Item("ColTypeTable")
                    ws.Cells(zmtxXClearRows + 3, rowCtr) = ds.Tables("defCol").Rows(ctr).Item("ColTypeRecSet")
                    'ws.Cells(zmtxXClearRows + 3, rowCtr) = IIf(ds.Tables("defCol").Rows(ctr).Item("BNDInteger")
                    ws.Cells(zmtxXClearRows + 4, rowCtr) = ds.Tables("defCol").Rows(ctr).Item("ClassConcat")
                    ws.Cells(zmtxXClearRows + 5, rowCtr) = ds.Tables("defCol").Rows(ctr).Item("ColActive")
                    ws.Cells(zmtxXClearRows + 9, rowCtr) = ds.Tables("defCol").Rows(ctr).Item("OBJField")
                    rowCtr = rowCtr + 1
                Next

                '// Write the Rows/Constraints
                '========================================================================
                rowCtr = zmtxRowStart
                For ctr = 0 To ds.Tables("defRow").Rows.Count - 1
                    Application.DoEvents()
                    endRow = rowCtr
                    ws.Cells(rowCtr, 2) = ds.Tables("defRow").Rows(ctr).Item("RowType")
                    ws.Cells(rowCtr, 3) = ds.Tables("defRow").Rows(ctr).Item("RowTypeDesc")
                    ws.Cells(rowCtr, 4) = ds.Tables("defRow").Rows(ctr).Item("RowTypeTable")
                    ws.Cells(rowCtr, 5) = ds.Tables("defRow").Rows(ctr).Item("RowTypeRecSet")
                    ws.Cells(rowCtr, 6) = ds.Tables("defRow").Rows(ctr).Item("ClassConcat")
                    ws.Cells(rowCtr, 7) = ds.Tables("defRow").Rows(ctr).Item("RowActive")
                    rowCtr = rowCtr + 1
                Next

                '// Write the Blocks
                '========================================================================
                For col = zmtxColumnStart To endCol
                    Application.DoEvents()
                    For row = zmtxRowStart To endRow
                        Application.DoEvents()
                        rng = CType(ws.Cells(zmtxXClearRows, col), Excel.Range)
                        colType = CStr(rng.Text).Trim
                        rng = CType(ws.Cells(row, 2), Excel.Range)
                        rowType = CStr(rng.Text).Trim
                        'colType = CStr(ws.Cells(zmtxRowStart, col))
                        'rowType = CStr(ws.Cells(row, 2))
                        linqTable = ds.Tables("defCoef").AsEnumerable()
                        strQueryResults = _
                            From r In linqTable _
                            Where _
                                Trim(CStr(r!ColType)) = colType _
                                And _
                                Trim(CStr(r!RowType)) = rowType _
                            Select Block = CStr(r("CoeffRecSet")) & "." & CStr(r("CoeffField"))
                        'Select Block = CStr(r("CoeffRecSet")) & "." & CStr(r("CoeffField"))
                        arr = strQueryResults.ToArray()
                        If strQueryResults.ToArray.Length > 0 Then
                            ws.Cells(row, col) = strQueryResults.ToArray(0)
                        End If
                    Next
                Next

                '// write model info
                ws.Cells(zmtxXClearRows, 2) = "MODEL NAME:  " & GetSwitchArgument(switches, "/PRJ", 1)
                ws.Cells(zmtxXClearRows + 1, 2) = "OPT MODEL"
                ws.Cells(zmtxXClearRows + 2, 2) = "MATRIX SCHEMATIC"
                ws.Cells(zmtxXClearRows + 3, 2) = "'" & FormatDateTime(Now(), DateFormat.ShortDate)
                '// format model info
                ws.Range(ws.Cells(zmtxXClearRows, 2), ws.Cells(zmtxXClearRows + 3, 2)).Select()
                With xl.Selection.Font
                    .Name = "Candara"
                    .Size = 10
                    .Bold = True
                End With
                ws.Range(ws.Cells(zmtxXClearRows, 2), ws.Cells(zmtxXClearRows + 3, 2 + 2)).Select()
                With xl.Selection.Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlMedium
                End With
                With xl.Selection.Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlMedium
                End With
                With xl.Selection.Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlMedium
                End With
                With xl.Selection.Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlMedium
                End With

                '// Add the OBJ ROW
                ws.Cells(zmtxXClearRows + 9, 2) = "Objective"
                ws.Cells(zmtxXClearRows + 9, 3) = "PROFIT"
                ws.Cells(zmtxXClearRows + 9, 4) = "OBJ"
                '// FORMAT THE OBJECTIVE ROW
                ws.Range(ws.Cells(zmtxXClearRows + 9, 2), ws.Cells(zmtxXClearRows + 9, endCol + 4)).Select()
                'ws.Range(ws.Cells(13, 2), ws.Cells(13, 26)).Select()
                With xl.Selection.Font
                    .Name = "Candara"
                    .Size = 13
                End With
                With xl.Selection.Borders(xlEdgeTop)
                    .LineStyle = xlDouble
                    .Weight = xlThick
                End With
                With xl.Selection.Borders(xlEdgeBottom)
                    .LineStyle = xlDouble
                    .Weight = xlThick
                End With

                '// Add the SENSE COLUMN AND RHS
                ws.Cells(zmtxXClearRows, endCol + 2) = "'Sns"
                ws.Cells(zmtxXClearRows, endCol + 3) = "RHS"
                '// FORMAT the SENSE COLUMN AND RHS
                ws.Range(ws.Cells(zmtxXClearRows, endCol + 2), ws.Cells(zmtxXClearRows, endCol + 3)).Select()
                With xl.Selection.Font
                    .Name = "CaslonOldFace"
                    .Size = 13
                    .Bold = True
                End With

                '//Write the ROW SENSES AND THE RHS
                rowCtr = zmtxRowStart
                For ctr = 0 To ds.Tables("defRow").Rows.Count - 1
                    Application.DoEvents()
                    endRow = rowCtr
                    Select Case ds.Tables("defRow").Rows(ctr).Item("RowTypeSNS")
                        Case "L"
                            ws.Cells(rowCtr, endCol + 2) = "'<="
                        Case "G"
                            ws.Cells(rowCtr, endCol + 2) = "'>="
                        Case Else 'MOSTLY "E"
                            ws.Cells(rowCtr, endCol + 2) = "'="
                    End Select
                    ws.Cells(rowCtr, endCol + 3) = ds.Tables("defRow").Rows(ctr).Item("RHSField")
                    rowCtr = rowCtr + 1
                Next

                '// ADD THE BOUNDS
                ws.Cells(endRow + 4, 2) = "BOUNDS"
                ws.Cells(endRow + 3, 7) = "LO"
                ws.Cells(endRow + 4, 7) = "FIX"
                ws.Cells(endRow + 5, 7) = "UP"

                '//Write the Column Bounds
                rowCtr = zmtxColumnStart
                For ctr = 0 To ds.Tables("defCol").Rows.Count - 1
                    Application.DoEvents()
                    endCol = rowCtr
                    ws.Cells(endRow + 3, rowCtr) = ds.Tables("defCol").Rows(ctr).Item("BNDLoField")
                    ws.Cells(endRow + 5, rowCtr) = ds.Tables("defCol").Rows(ctr).Item("BNDUpField")
                    rowCtr = rowCtr + 1
                Next

                '//format the bounds lines
                ws.Range(ws.Cells(endRow + 3, 2), ws.Cells(endRow + 5, endCol + 4)).Select()
                With xl.Selection.Font
                    .Name = "CaslonOldFace"
                    .Size = 13
                    .Bold = True
                End With
                ws.Range(ws.Cells(endRow + 4, 2), ws.Cells(endRow + 4, endCol + 4)).Select()  'add horizontal line top of word "BOUNDS" and bottom too
                With xl.Selection.Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlMedium
                End With
                With xl.Selection.Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .Weight = xlMedium
                End With
                ws.Range(ws.Cells(zmtxXClearRows, 7), ws.Cells(endRow + 5, 7)).Select() 'add vertical line on the right of the LO/UP/BOUNDS
                With xl.Selection.Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                    .Weight = xlMedium
                End With

                '// Insert Spaces in Cell Values where needed
                ws.Cells.Select()
                xl.Selection.Replace(What:="][", Replacement:="] [", LookAt:=xlPart, _
                    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                    ReplaceFormat:=False)
                xl.Selection.Replace(What:=".", Replacement:=". ", LookAt:=xlPart, _
                    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                    ReplaceFormat:=False)
                xl.Selection.Replace(What:="_", Replacement:="_ ", LookAt:=xlPart, _
                    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                    ReplaceFormat:=False)

                '// More Formatting
                ws.Cells.Select()
                With xl.Selection
                    .VerticalAlignment = xlCenter
                    .WrapText = True
                    .MergeCells = False
                End With

                ws.Cells.Select()
                With xl.Selection
                    .VerticalAlignment = xlCenter
                    .WrapText = True
                    .MergeCells = False
                End With
                ws.Columns(2).Select()
                With xl.Selection
                    .WrapText = False
                End With
                ws.Range(ws.Columns(endCol + 1), ws.Columns(endCol + 10)).Select()
                With xl.Selection
                    .WrapText = False
                End With
                ws.Rows(2).Select()
                With xl.Selection
                    .WrapText = False
                End With


                ws.Range(ws.Columns(zmtxColumnStart), ws.Columns(endCol + 4)).Select()
                With xl.Selection
                    .HorizontalAlignment = xlCenter
                    .WrapText = True
                    .Orientation = 0
                    .MergeCells = False
                End With
                ws.Columns(1).Select()
                xl.Selection.ColumnWidth = 1
                ws.Range(ws.Columns(endCol + 2), ws.Columns(endCol + 7)).Select()
                xl.Selection.ColumnWidth = 1
                ws.Range(ws.Columns(zmtxColumnStart - 7), ws.Columns(endCol)).Select()
                xl.Selection.ColumnWidth = 16.5
                ws.Range(ws.Columns(zmtxColumnStart - 4), ws.Columns(zmtxColumnStart - 1)).Select()
                xl.Selection.ColumnWidth = 1
                ws.Range(ws.Columns(4), ws.Columns(endCol + 10)).Select()
                xl.Selection.EntireColumn.AutoFit()
                ws.Columns(endCol + 1).Select()  'blank col before sense
                xl.Selection.ColumnWidth = 1
                ws.Columns(endCol + 2).Select()  'sense
                xl.Selection.ColumnWidth = 5
                ws.Cells.Select()
                xl.Selection.EntireRow.AutoFit()  'autofit all the rows

                '// Replace "Plus" with +1, etc.
                '*&*&*&*&*&
                For Each rCell In ws.Range(ws.Cells(1, 1), ws.Cells(endRow, endCol))
                    If Right(rCell.Value, 6) = ". Plus" Then rCell.Value = "'+ 1"
                    If Right(rCell.Value, 9) = ". PlusOne" Then rCell.Value = "'+ 1"
                    If Right(rCell.Value, 7) = ". Minus" Then rCell.Value = "'- 1"
                    If Right(rCell.Value, 10) = ". MinusOne" Then rCell.Value = "'- 1"

                Next rCell


                '// Add big overall border
                ws.Range(ws.Cells(zmtxXClearRows - 2, 2), ws.Cells(endRow + 7, endCol + 6)).Select()
                With xl.Selection.Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlMedium
                End With
                With xl.Selection.Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlMedium
                End With
                With xl.Selection.Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlMedium
                End With
                With xl.Selection.Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlMedium
                End With


                '// Zoom To Selection
                ws.Range(ws.Cells(1, 1), ws.Cells(endRow + 9, endCol + 7)).Select()
                xl.Selection.Activate()
                xl.ActiveWindow.Zoom = True


                '//Page Setup for Printing
                xl.ActiveSheet.PageSetup.PrintArea = ""
                xl.Application.PrintCommunication = False
                With xl.ActiveSheet.PageSetup
                    .LeftMargin = xl.Application.InchesToPoints(0.5)
                    .RightMargin = xl.Application.InchesToPoints(0.5)
                    .TopMargin = xl.Application.InchesToPoints(0.5)
                    .BottomMargin = xl.Application.InchesToPoints(0.5)
                    .HeaderMargin = xl.Application.InchesToPoints(0.3)
                    .FooterMargin = xl.Application.InchesToPoints(0.3)
                    .PrintHeadings = False
                    .PrintGridlines = True
                    .PrintQuality = 600
                    .CenterHorizontally = True
                    .CenterVertically = True
                    .Orientation = xlLandscape
                    .Draft = False
                    .PaperSize = xlPaperLetter
                    .FirstPageNumber = xlAutomatic
                    .Order = xlDownThenOver
                    .BlackAndWhite = False
                    .Zoom = False
                    .FitToPagesWide = 1
                    .FitToPagesTall = 1
                End With
                xl.Application.PrintCommunication = True


                wb.SaveAs(outputFile)
                'xl.Quit()
                ws = Nothing
                wb = Nothing
                xl = Nothing

                Return 0
            Catch ex As Exception
                GenUtils.Message(MsgType.Warning, "Generate zMatrix", ex.Message)
                EntLib.COPT.Log.Log(GenUtils.GetWorkDir(switches), "C-OPT GenUtils", "Generate zMatrix Error - " & ex.Message)
                Return -1
            End Try

        End Function

        ''' <summary>
        ''' Generates a GUID.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' RKP/06-07-2010/v2.3.132
        ''' </remarks>
        Public Shared Function GenerateGUID() As String
            'Dim myGUID As Guid = Guid.NewGuid()
            Return Guid.NewGuid.ToString()
        End Function

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="connectionString"></param>
        ''' <param name="dt"></param>
        ''' <param name="sql"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' RKP/06-16-10/v2.3.133
        ''' </remarks>
        Public Shared Function UpdateDB(ByVal connectionString As String, ByRef dt As DataTable, ByVal sql As String) As Integer
            Dim data_adapter As OleDb.OleDbDataAdapter

            data_adapter = New OleDb.OleDbDataAdapter(sql, connectionString)

            Dim command_builder As New OleDb.OleDbCommandBuilder(data_adapter)
            dt.EndInit()
            Return data_adapter.Update(dt)
        End Function

        Public Shared Function InsertDB(ByVal connectionString As String, ByRef dt As DataTable, ByVal tableName As String, ByVal sql As String) As Integer
            'Dim myConn As New SqlConnection(connectionString)
            Dim myConn As New OleDb.OleDbConnection(connectionString)
            myConn.Open()
            'Dim myDataAdapter As New SqlDataAdapter()
            Dim myDataAdapter As New OleDb.OleDbDataAdapter(New OleDb.OleDbCommand(sql, myConn))
            'myDataAdapter.SelectCommand = New SqlCommand(sql, myconn)
            'myDataAdapter.SelectCommand = New OleDbCommand(sql, myconn)
            'Dim cb As SqlCommandBuilder = New SqlCommandBuilder(myDataAdapter)
            Dim cb As OleDb.OleDbCommandBuilder = New OleDb.OleDbCommandBuilder(myDataAdapter)
            Dim cmd As New OleDb.OleDbCommand(sql, myConn)


            myDataAdapter.InsertCommand = cmd
            myDataAdapter.InsertCommand.CommandType = CommandType.Text

            Dim ds As DataSet = New DataSet
            ds.Tables.Add(dt.Copy)
            'myDataAdapter.Fill(ds, tableName)
            'myDataAdapter.Fill(dt, tableName)

            ' Code to modify data in DataSet here 

            ' Without the SqlCommandBuilder this line would fail.
            'myDataAdapter.Update(ds, tableName)



            myDataAdapter.Update(ds, tableName)

            myConn.Close()

        End Function

        Public Shared Function GenerateSOLFile() As Integer

            Return 0
        End Function

        Public Shared Function ImportUsingExcel(ByRef dtRow As DataTable, ByRef dtCol As DataTable, ByVal workDir As String) As Integer

            Try
                My.Computer.FileSystem.DeleteFile(workDir & "\dtRow.xml")
            Catch ex As Exception

            End Try
            Try
                dtRow.TableName = "dtRow"
                dtRow.WriteXml(workDir & "\dtRow.xml")
            Catch ex As Exception
                GenUtils.Message(GenUtils.MsgType.Critical, "ImportUsingExcel", ex.Message)
            End Try

            Dim xl As New Excel.Application
            xl.Visible = True
            Dim wb As Excel.Workbook = xl.Workbooks.Add
            'Dim ws As Excel.Worksheet '= CType(wb.Worksheets.Add, Excel.Worksheet)
            'ws.Name = "DataTable"
            Try
                'ws.Range("A1").CopyFromRecordset(dtRow)
                'Workbooks.OpenXML Filename:="C:\OPTMODELS\FARMER\Output\dtRow.xml", LoadOption:=xlXmlLoadImportToList
                'xl.Workbooks.OpenXML(workDir & "\dtRow.xml", , 2)
                'xl.ActiveWorkbook.XmlImport(workDir & "\dtRow.xml", Nothing, True, xl.Range("A1"))
                'C:\OPTMODELS\HIPNET\Output

                'xl.ActiveWorkbook.XmlImport("C:\OPTMODELS\HIPNET\Output\dtRow.xml", Nothing, True, xl.Range("A1"))
                'xl.ActiveWorkbook.SaveAs("C:\OPTMODELS\HIPNET\Output\dtRow.xml.xlsx")

                wb.XmlImport("C:\OPTMODELS\HIPNET\Output\dtRow.xml", Nothing, True, xl.Range("A1"))
                'ws = CType(wb.Worksheets("Sheet1"), Excel.Worksheet)
                'ws.Select()
                'ws.Range("A1").Select()
                'xl.Range(xl.Selection, xl.Selection.end(-4161)).select()


                wb.SaveAs("C:\OPTMODELS\HIPNET\Output\dtRow.xml.xlsx")


            Catch ex As Exception
                'GenUtils.Message(GenUtils.MsgType.Critical, "ImportUsingExcel", ex.Message)
            End Try

            Return 0

        End Function

        Public Shared Function SerializeArrayDouble(ByVal workDir As String, ByVal outputFileName As String, ByRef srcArray() As Double) As Integer
            Try
                My.Computer.FileSystem.DeleteFile(workDir & "\" & outputFileName)
            Catch ex As Exception

            End Try
            Try
                System.IO.File.WriteAllLines(workDir & "\" & outputFileName, Array.ConvertAll(srcArray, New Converter(Of Double, String)(Function(t As Double) t.ToString())))
                Return 0
            Catch ex As Exception
                GenUtils.Message(MsgType.Information, "SerializeArrayDouble", ex.Message)
                Console.WriteLine("Error converting solver array to table - " & ex.Message)
                Debug.Print("Error converting solver array to table - " & ex.Message)
                EntLib.COPT.Log.Log(workDir, "C-OPT GenUtils", "SerializeArrayDouble - Error converting solver array to table - " & ex.Message)
                Return -1
            End Try
        End Function

        Public Shared Function Serialize(ByVal workDir As String, ByVal outputFileName As String, ByRef srcArray() As Integer) As Integer
            Dim msgSrc As String = "Serialize (Integer Array)"
            Dim msg As String
            'Dim dt As DataTable

            If srcArray Is Nothing Then
                Return 0
            Else
                Try
                    My.Computer.FileSystem.DeleteFile(workDir & "\" & outputFileName)
                Catch ex As Exception

                End Try
                Try
                    System.IO.File.WriteAllLines(workDir & "\" & outputFileName, Array.ConvertAll(srcArray, New Converter(Of Integer, String)(Function(t As Integer) t.ToString())))
                    Return 0
                Catch ex As Exception
                    msg = msgSrc & vbNewLine & "Error saving to disk:" & vbNewLine & workDir & "\" & outputFileName & vbNewLine & ex.Message
                    GenUtils.Message(MsgType.Critical, msgSrc, msg)
                    Console.WriteLine(msg)
                    Debug.Print(msg)
                    EntLib.COPT.Log.Log(workDir, "C-OPT GenUtils", msg)
                    Return -1
                End Try
                'dt = New DataTable
                'dt.LoadDataRow(srcArray, True)
            End If
        End Function

        Public Shared Function Serialize(ByVal workDir As String, ByVal outputFileName As String, ByRef srcArray() As Double) As Integer
            Dim msgSrc As String = "Serialize (Double Array)"
            Dim msg As String

            If srcArray Is Nothing Then
                Return 0
            Else
                Try
                    My.Computer.FileSystem.DeleteFile(workDir & "\" & outputFileName)
                Catch ex As Exception

                End Try
                Try
                    System.IO.File.WriteAllLines(workDir & "\" & outputFileName, Array.ConvertAll(srcArray, New Converter(Of Double, String)(Function(t As Double) t.ToString())))
                    Return 0
                Catch ex As Exception
                    msg = msgSrc & vbNewLine & "Error saving to disk:" & vbNewLine & workDir & "\" & outputFileName & vbNewLine & ex.Message
                    GenUtils.Message(MsgType.Critical, msgSrc, msg)
                    Console.WriteLine(msg)
                    Debug.Print(msg)
                    EntLib.COPT.Log.Log(workDir, "C-OPT GenUtils", msg)
                    Return -1
                End Try
            End If
        End Function

        Public Shared Function Serialize(ByVal workDir As String, ByVal outputFileName As String, ByRef srcArray() As Char) As Integer
            Dim msgSrc As String = "Serialize (Char Array)"
            Dim msg As String

            If srcArray Is Nothing Then
                Return 0
            Else
                Try
                    My.Computer.FileSystem.DeleteFile(workDir & "\" & outputFileName)
                Catch ex As Exception

                End Try
                Try
                    System.IO.File.WriteAllLines(workDir & "\" & outputFileName, Array.ConvertAll(srcArray, New Converter(Of Char, String)(Function(t As Char) t.ToString())))
                    Return 0
                Catch ex As Exception
                    msg = msgSrc & vbNewLine & "Error saving to disk:" & vbNewLine & workDir & "\" & outputFileName & vbNewLine & ex.Message
                    GenUtils.Message(MsgType.Critical, msgSrc, msg)
                    Console.WriteLine(msg)
                    Debug.Print(msg)
                    EntLib.COPT.Log.Log(workDir, "C-OPT GenUtils", msg)
                    Return -1
                End Try
            End If
        End Function

        Public Shared Function Serialize(ByVal workDir As String, ByVal outputFileName As String, ByRef srcArray() As String) As Integer
            Dim msgSrc As String = "Serialize (String Array)"
            Dim msg As String

            If srcArray Is Nothing Then
                Return 0
            Else
                Try
                    My.Computer.FileSystem.DeleteFile(workDir & "\" & outputFileName)
                Catch ex As Exception

                End Try
                Try
                    System.IO.File.WriteAllLines(workDir & "\" & outputFileName, Array.ConvertAll(srcArray, New Converter(Of String, String)(Function(t As String) t.ToString())))
                    Return 0
                Catch ex As Exception
                    msg = msgSrc & vbNewLine & "Error saving to disk:" & vbNewLine & workDir & "\" & outputFileName & vbNewLine & ex.Message
                    GenUtils.Message(MsgType.Critical, msgSrc, msg)
                    Console.WriteLine(msg)
                    Debug.Print(msg)
                    EntLib.COPT.Log.Log(workDir, "C-OPT GenUtils", msg)
                    Return -1
                End Try
            End If
        End Function

        Public Shared Function Serialize(ByVal workDir As String, ByVal outputFileName As String, ByRef dt As DataTable) As Integer
            Dim msgSrc As String = "Serialize (DataTable)"
            Dim msg As String

            If dt Is Nothing Then
                Return 0
            Else
                Try
                    My.Computer.FileSystem.DeleteFile(workDir & "\" & outputFileName)
                Catch ex As Exception

                End Try
                Try
                    'System.IO.File.WriteAllLines(workDir & "\" & outputFileName, Array.ConvertAll(srcArray, New Converter(Of String, String)(Function(t As String) t.ToString())))
                    dt.WriteXml(workDir & "\" & outputFileName)
                    Return 0
                Catch ex As Exception
                    msg = msgSrc & vbNewLine & "Error saving to disk:" & vbNewLine & workDir & "\" & outputFileName & vbNewLine & ex.Message
                    GenUtils.Message(MsgType.Critical, msgSrc, msg)
                    Console.WriteLine(msg)
                    Debug.Print(msg)
                    EntLib.COPT.Log.Log(workDir, "C-OPT GenUtils", msg)
                    Return -1
                End Try
            End If
        End Function

        Public Shared Function SerializeDataTable(ByVal workDir As String, ByVal outputFileName As String, ByRef dt As DataTable) As Integer
            Try
                My.Computer.FileSystem.DeleteFile(workDir & "\" & outputFileName)
            Catch ex As Exception

            End Try
            Try
                'System.IO.File.WriteAllLines(workDir & "\" & outputFileName, Array.ConvertAll(srcArray, New Converter(Of Double, String)(Function(t As Double) t.ToString())))
                dt.WriteXml(workDir & "\" & outputFileName, System.Data.XmlWriteMode.WriteSchema, False)
                Return 0
            Catch ex As Exception
                GenUtils.Message(MsgType.Critical, "SerializeDataTable", ex.Message)
                Console.WriteLine("Error in SerializeDataTable - " & ex.Message)
                Debug.Print("Error in SerializeDataTable - " & ex.Message)
                EntLib.COPT.Log.Log(workDir, "C-OPT GenUtils", "Error in SerializeDataTable - " & ex.Message)
                Return -1
            End Try

            'dt.ReadXml(workDir & "\" & outputFileName)
        End Function

        Public Shared Function ConvertDataTableToRecordset_Test() As Integer

            Dim dt As New DataTable()
            Dim rs As ADODB.Recordset
            Dim currentDb As New DAAB("HIPNET")

            Try
                dt.ReadXmlSchema("C:\OPTMODELS\HIPNET\Output\dtCol.xml")
                dt.ReadXml("C:\OPTMODELS\HIPNET\Output\dtCol.xml")
            Catch ex As Exception
                GenUtils.Message(MsgType.Critical, "ConvertDataTableToRecordset", ex.Message)
            End Try

            rs = ConvertToRecordset(dt)
            Try
                My.Computer.FileSystem.DeleteFile("C:\OPTMODELS\HIPNET\Output\Test.ado.xml")
            Catch ex As Exception

            End Try
            rs.Save("C:\OPTMODELS\HIPNET\Output\Test.ado.xml", ADODB.PersistFormatEnum.adPersistXML)

            Dim xl As New Excel.Application
            xl.Visible = True
            Dim wb As Excel.Workbook = xl.Workbooks.Open("C:\OPTMODELS\HIPNET\Output\Temp.xlsm")
            'Dim wb As Excel.Workbook = xl.Workbooks.Add
            'Dim ws As Excel.Worksheet = CType(wb.Worksheets(1), Excel.Worksheet)
            'ws.Name = "Test"
            'ws.Range("A2").CopyFromRecordset(rs)
            'Dim ctr As Integer
            'For ctr = 0 To rs.Fields.Count - 1
            '    ws.Cells(1, ctr + 1) = rs.Fields(ctr).Name
            'Next
            ''xl.WorksheetFunction.
            'ws.Range("A1:O" & rs.RecordCount + 1).Name = "Test"

            'Try
            '    My.Computer.FileSystem.DeleteFile("C:\OPTMODELS\HIPNET\Output\Test.xlsx")
            'Catch ex As Exception

            'End Try
            'wb.SaveAs("C:\OPTMODELS\HIPNET\Output\Test.xlsx")

            'wb.Close()
            'xl.Quit()

            'Try
            '    currentDb.ExecuteNonQuery("DROP TABLE TEST_LINKED")
            'Catch ex As Exception
            '    'GenUtils.Message(MsgType.Critical, "ConvertDataTableToRecordset", ex.Message)
            'End Try

            'Dim con As New ADODB.Connection
            'Dim cat As New ADOX.Catalog
            'Dim tbl As New ADOX.Table
            'Try

            '    con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & "C:\OPTMODELS\HIPNET\MTXDATA.MDB" & ";User Id=admin;Password="
            '    con.Open()
            '    Dim connStr As String = "Excel 8.0;HDR=YES;IMEX=2;ACCDB=YES;DATABASE=" & "C:\OPTMODELS\HIPNET\Output\Test.xlsx" & ";TABLE=" & "Test"

            '    cat.ActiveConnection = con

            '    With tbl
            '        .Name = "TEST_LINKED"
            '        .ParentCatalog = cat
            '        .Properties("Jet OLEDB:Link Provider String").Value = connStr
            '        .Properties("Jet OLEDB:Remote Table Name").Value = "Test" 'ws.Name.ToString()
            '        .Properties("Jet OLEDB:Create Link").Value = True
            '    End With
            '    cat.Tables.Append(tbl)
            '    cat.Tables.Refresh()
            'Catch ex As Exception
            '    GenUtils.Message(MsgType.Critical, "ConvertDataTableToRecordset", ex.Message)
            'End Try

            'Try
            '    currentDb.ExecuteNonQuery("DROP TABLE TEST")
            'Catch ex As Exception

            'End Try

            'Try
            '    currentDb.ExecuteNonQuery("SELECT * INTO TEST FROM TEST_LINKED")
            'Catch ex As Exception
            '    GenUtils.Message(MsgType.Critical, "ConvertDataTableToRecordset", ex.Message)
            'End Try


            dt = Nothing
            rs = Nothing
            'ws = Nothing
            'wb = Nothing
            'xl = Nothing
            'con = Nothing
            'cat = Nothing
            'tbl = Nothing


            Return 0
        End Function

        Public Shared Function SpreadsheetExport( _
            ByVal dt As DataTable, _
            ByVal excelFilePath As String, _
            ByVal excelRangeName As String, _
            ByRef currentDb As DAAB, _
            ByVal sql As String _
        ) As Integer

            Dim xl As New Excel.Application
            Dim ctr As Integer
            Dim rs As New ADODB.Recordset
            Dim lastCol As String

            rs = ConvertToRecordset(dt)
            lastCol = ColumnLetter(rs.Fields.Count).ToString()

            xl.DisplayAlerts = False
            xl.Visible = True
            Dim wb0 As Excel.Workbook = xl.Workbooks.Add
            Dim wb As Excel.Workbook = wb0.Application.Workbooks.Open(excelFilePath)
            'Dim wb As Excel.Worksheet = CType(xl.Workbooks.Open(excelFilePath), Excel.Worksheet)
            'Dim ws As Excel.Worksheet = CType(wb.Worksheets.Add(), Excel.Worksheet)
            Dim ws As Excel.Worksheet = CType(wb.Worksheets(excelRangeName), Excel.Worksheet)
            ws.Cells.ClearContents()
            'Dim ws As Excel.Worksheet = wb
            'ws.Name = excelRangeName
            ws.Range("A2").CopyFromRecordset(rs)
            For ctr = 0 To rs.Fields.Count - 1
                ws.Cells(1, ctr + 1) = rs.Fields(ctr).Name
            Next

            ws.Range("A1:" & lastCol & rs.RecordCount + 1).Name = excelRangeName
            'Try
            '    'My.Computer.FileSystem.DeleteFile(excelFilePath)
            'Catch ex As Exception

            'End Try
            wb.Save() '(excelFilePath)

            'xl.Run("MacroCreateLinkedTable")
            xl.Run("MacroUpdateData")

            'If currentDb IsNot Nothing Then
            '    Try
            '        currentDb.ExecuteNonQuery(sql)
            '    Catch ex As Exception
            '        GenUtils.Message(MsgType.Critical, "SpreadsheetExport", ex.Message)
            '    End Try
            'End If

            wb0.Close()
            wb.Save()
            wb.Close()
            xl.Quit()
            ws = Nothing
            wb0 = Nothing
            wb = Nothing
            xl = Nothing

            Return 0
        End Function

        'Public Shared Function SpreadsheetImport( _
        '    ByVal accessFilePath As String, _
        '    ByVal accessTableName As String, _
        '    ByVal excelFilePath As String, _
        '    ByVal excelRangeName As String _
        ') As Integer

        '    Dim accessApp As New Access.Application

        '    accessApp.Visible = True
        '    Try
        '        accessApp.OpenCurrentDatabase(accessFilePath, False)
        '        Try
        '            accessApp.CurrentProject.Connection.Execute("DROP TABLE [" & accessTableName & "]")
        '        Catch ex As Exception

        '        End Try
        '        accessApp.DoCmd.TransferSpreadsheet(Access.AcDataTransferType.acImport, Access.AcSpreadSheetType.acSpreadsheetTypeExcel9, accessTableName, excelFilePath, True, excelRangeName)
        '        accessApp.CloseCurrentDatabase()
        '    Catch ex As Exception
        '        GenUtils.Message(GenUtils.MsgType.Critical, "SpreadsheetImport", ex.Message)
        '    End Try

        '    accessApp = Nothing
        '    Return 0
        'End Function

        Private Sub Test()

        End Sub

        Public Shared Function ColumnLetter(ByVal ColumnNumber As Integer) As String
            If ColumnNumber > 26 Then

                ' 1st character:  Subtract 1 to map the characters to 0-25,
                '                 but you don't have to remap back to 1-26
                '                 after the 'Int' operation since columns
                '                 1-26 have no prefix letter

                ' 2nd character:  Subtract 1 to map the characters to 0-25,
                '                 but then must remap back to 1-26 after
                '                 the 'Mod' operation by adding 1 back in
                '                 (included in the '65')

                ColumnLetter = Chr(CInt(Int((ColumnNumber - 1) / 26) + 64)) & _
                               Chr(((ColumnNumber - 1) Mod 26) + 65)
            Else
                ' Columns A-Z
                ColumnLetter = Chr(ColumnNumber + 64)
            End If
        End Function

        Public Shared Function SpreadsheetOpen( _
            ByVal dt As DataTable, _
            ByVal excelFilePath As String, _
            ByVal excelRangeName As String _
        ) As Integer

            Dim xl As New Excel.Application
            Dim ctr As Integer
            Dim rs As New ADODB.Recordset

            rs = ConvertToRecordset(dt)

            xl.Visible = True
            Dim wb0 As Excel.Workbook = xl.Workbooks.Add
            Dim wb As Excel.Workbook = wb0.Application.Workbooks.Open(excelFilePath)
            'Dim wb As Excel.Worksheet = CType(xl.Workbooks.Open(excelFilePath), Excel.Worksheet)
            'Dim ws As Excel.Worksheet = CType(wb.Worksheets.Add(), Excel.Worksheet)
            Dim ws As Excel.Worksheet = CType(wb.Worksheets(excelRangeName), Excel.Worksheet)
            'Dim ws As Excel.Worksheet = wb
            'ws.Name = excelRangeName
            ws.Range("A2").CopyFromRecordset(rs)
            For ctr = 0 To rs.Fields.Count - 1
                ws.Cells(1, ctr + 1) = rs.Fields(ctr).Name
            Next
            ws.Range("A1:" & ColumnLetter(rs.Fields.Count).ToString() & rs.RecordCount + 1).Name = excelRangeName
            Try
                'My.Computer.FileSystem.DeleteFile(excelFilePath)
            Catch ex As Exception

            End Try
            wb0.Close()
            wb.Save() '(excelFilePath)
            wb.Close()

            xl.Quit()
            ws = Nothing
            wb0 = Nothing
            wb = Nothing
            xl = Nothing

            Return 0
        End Function

        Public Shared Function ConnectToDBUsingDAO(ByVal dbPath As String, ByRef dt As DataTable, ByVal tableName As String) As Integer
            Dim daoDbEngine As Dao.DBEngine
            Dim daoDB As Dao.Database
            Dim daoRS As Dao.Recordset
            Dim sql As String

            daoDbEngine = New Dao.DBEngine
            daoDB = daoDbEngine.OpenDatabase(dbPath)
            sql = "SELECT * FROM [" & tableName & "]"
            daoRS = daoDB.OpenRecordset(sql, Dao.RecordsetTypeEnum.dbOpenDynamic, Dao.LockTypeEnum.dbOptimisticBatch)

            Return 0
        End Function

        Public Shared Function ConnectToDBUsingADO(ByVal dbPath As String, ByRef dt As DataTable, ByVal tableName As String) As Integer
            Dim sql As String
            'Dim ret As Long
            'Dim rs As ADODB.Recordset
            'Dim conn As ADODB.Connection
            Dim connStr As String

            sql = "SELECT * FROM [" & tableName & "]"
            connStr = ""


            Return 0
        End Function

        Public Shared Function IsAuthorized(ByVal switches() As String) As Boolean



            Return False
        End Function

        Public Function IsValidUrl(ByVal Url As String) As Boolean
            Dim strRegex As String = "^(https?://)" _
                                      & "?(([0-9a-z_!~*'().&=+$%-]+: )?[0-9a-z_!~*'().&=+$%-]+@)?" _
                                      & "(([0-9]{1,3}\.){3}[0-9]{1,3}" _
                                      & "|" _
                                      & "([0-9a-z_!~*'()-]+\.)*" _
                                      & "([0-9a-z][0-9a-z-]{0,61})?[0-9a-z]\." _
                                      & "[a-z]{2,6})" _
                                      & "(:[0-9]{1,4})?" _
                                      & "((/?)|" _
                                      & "(/[0-9a-z_!~*'().;?:@&=+$,%#-]+)+/?)$"

            Dim re As RegularExpressions.Regex = New RegularExpressions.Regex(strRegex)
            'MessageBox.Show("IP: " & Net.IPAddress.TryParse(Url, Nothing))
            If re.IsMatch(Url) Then
                Return True
            Else
                Return False
            End If
        End Function

        ''' <summary>
        ''' Returns total amount of free physical memory for the computer, in GB.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' RKP/03-26-12/v3.0.159
        ''' </remarks>
        Public Shared Function GetAvailablePhysicalMemory() As Double
            Return CDbl(Microsoft.VisualBasic.FormatNumber(((My.Computer.Info.AvailablePhysicalMemory / 1024) / 1024) / 1024, 4))
        End Function

        ''' <summary>
        ''' Returns total percent of free physical memory for the computer.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' RKP/03-26-10/v3.0.159
        ''' </remarks>
        Public Shared Function GetAvailablePhysicalMemoryPercent() As Double
            'Return CDbl(Microsoft.VisualBasic.FormatNumber(((My.Computer.Info.AvailablePhysicalMemory / 1024) / 1024) / 1024, 4))
            Return CDbl(Microsoft.VisualBasic.FormatNumber((My.Computer.Info.AvailablePhysicalMemory / My.Computer.Info.TotalPhysicalMemory) * 100, 2))
        End Function

        Public Shared Function GetAvailablePhysicalMemoryStr() As String
            Return GenUtils.GetAvailablePhysicalMemory.ToString() & " GB; " & GenUtils.GetAvailablePhysicalMemoryPercent().ToString() & "%"
        End Function

        'Public Shared Function GetAvailableRAM() As Double
        '    'Dim ramCounter As New System.Diagnostics.PerformanceCounter("Memory", "Available MBytes")

        '    Return New System.Diagnostics.PerformanceCounter("Memory", "Available MBytes").NextValue() 'ramCounter.NextValue()
        'End Function

        ''' <summary>
        ''' Returns total amount of free virtual memory for the computer, in GB.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' RKP/03-26-10/v3.0.159
        ''' </remarks>
        Public Shared Function GetAvailableVirtualMemory() As Double
            Return CDbl(Microsoft.VisualBasic.FormatNumber(((My.Computer.Info.AvailableVirtualMemory / 1024) / 1024) / 1024, 4))
        End Function

        ''' <summary>
        ''' Returns the path to C-OPTSYS directory.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' RKP/04-03-12/v3.0.164
        ''' </remarks>
        Public Shared Function GetSysDirectory() As String
            Try
                Return GetNativeConfigValue("sysFolderPath")
            Catch ex As Exception
                Return "C:\OPTMODELS\C-OPTSYS"
            End Try
        End Function

        Public Shared Function GetSysConfigFilePath() As String
            Try
                Return GetSysDirectory() & "\" & GetNativeConfigValue("sysConfigFileName")
            Catch ex As Exception
                Return GetSysDirectory() & "\C-OPT.config.xml"
            End Try
        End Function

        'Public Shared Function SetSysConfigDeleteKey() As String
        '    Return ""
        'End Function

        Public Shared Function SetSysConfigUpdateKey(key As String, value As String) As String
            Dim config As System.Configuration.Configuration
            Dim fileMap As New System.Configuration.ExeConfigurationFileMap()

            fileMap.ExeConfigFilename = GenUtils.GetSysConfigFilePath()
            config = System.Configuration.ConfigurationManager.OpenMappedExeConfiguration(fileMap, ConfigurationUserLevel.None)

            config.AppSettings.Settings.Remove(key)
            'config.AppSettings.Settings.Add("lastWorkDir", GenUtils.GetSwitchArgument(switches, "/WorkDir", 1))
            config.AppSettings.Settings.Add(key, value)

            config.Save(System.Configuration.ConfigurationSaveMode.Modified)

            fileMap = Nothing
            config = Nothing

            Return "0"
        End Function

        Public Shared Function GetSysConfigValue(key As String) As String
            'config.AppSettings.Settings("logFileLength").Value

            Dim config As System.Configuration.Configuration
            Dim fileMap As New System.Configuration.ExeConfigurationFileMap()

            fileMap.ExeConfigFilename = GenUtils.GetSysConfigFilePath()
            config = System.Configuration.ConfigurationManager.OpenMappedExeConfiguration(fileMap, ConfigurationUserLevel.None)
            'Return IIf(String.IsNullOrEmpty(config.AppSettings.Settings(key).Value), "", config.AppSettings.Settings(key).Value).ToString()

            Try
                Return config.AppSettings.Settings(key).Value
            Catch ex As Exception
                Return ""
            End Try


        End Function

        ''' <summary>
        ''' Returns the value of a "key", from the native configuration file.
        ''' </summary>
        ''' <param name="key"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' RKP/04-16-12/v3.1.165
        ''' 
        ''' </remarks>
        Public Shared Function GetNativeConfigValue(key As String) As String
            'config.AppSettings.Settings("logFileLength").Value

            'Dim config As System.Configuration.Configuration
            'Dim fileMap As New System.Configuration.ExeConfigurationFileMap()

            'fileMap.ExeConfigFilename = GenUtils.GetSysConfigFilePath()
            'config = System.Configuration.ConfigurationManager.OpenMappedExeConfiguration(fileMap, ConfigurationUserLevel.None)
            'Return IIf(String.IsNullOrEmpty(config.AppSettings.Settings(key).Value), "", config.AppSettings.Settings(key).Value).ToString()


            'ConfigurationManager.AppSettings(key)
            Try
                Return System.Configuration.ConfigurationManager.AppSettings(key).ToString()
            Catch ex As Exception
                Return ""
            End Try


        End Function

        Public Shared Function GetSysConfig() As System.Configuration.Configuration
            'config.AppSettings.Settings("logFileLength").Value

            'Dim config As System.Configuration.Configuration
            Dim fileMap As New System.Configuration.ExeConfigurationFileMap()

            fileMap.ExeConfigFilename = GenUtils.GetSysConfigFilePath()
            'config = System.Configuration.ConfigurationManager.OpenMappedExeConfiguration(fileMap, ConfigurationUserLevel.None)
            'Return IIf(String.IsNullOrEmpty(config.AppSettings.Settings(key).Value), "", config.AppSettings.Settings(key).Value).ToString()

            Try
                Return System.Configuration.ConfigurationManager.OpenMappedExeConfiguration(fileMap, ConfigurationUserLevel.None)
            Catch ex As Exception
                Return Nothing
            End Try


        End Function

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function CollectGarbage() As Integer
            Try
                GC.Collect()
                GC.WaitForPendingFinalizers()
                GC.Collect()
                Return 0
            Catch ex As Exception
                Return -1
            End Try
        End Function

        ''' <summary>
        ''' Returns whether a project is available or not (as a valid connection string in C-OPT.config.xml).
        ''' </summary>
        ''' <param name="projectName"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' RKP/04-12-12/v3.1.165
        ''' </remarks>
        Public Shared Function IsProjectAvailable(ByVal projectName As String) As Boolean
            Dim list As String() = GetAllDatabases()

            If list.Contains(projectName) Then
                Return True
            Else
                Return False
            End If
        End Function

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="Args"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' RKP/04-12-12/v3.1.165
        ''' </remarks>
        Public Function CmdParse(ByVal Args As String()) As StringDictionary
            'Dim Parameters As StringDictionary

            Parameters = New StringDictionary()
            Dim Spliter As New Regex("^-{1,2}|^/|=|:", RegexOptions.IgnoreCase Or RegexOptions.Compiled)

            Dim Remover As New Regex("^['""]?(.*?)['""]?$", RegexOptions.IgnoreCase Or RegexOptions.Compiled)

            Dim Parameter As String = Nothing
            Dim Parts As String()

            ' Valid parameters forms:
            ' {-,/,--}param{ ,=,:}((",')value(",'))
            ' Examples: 
            ' -param1 value1 --param2 /param3:"Test-:-work" 
            '   /param4=happy -param5 '--=nice=--'
            For Each Txt As String In Args
                ' Look for new parameters (-,/ or --) and a
                ' possible enclosed value (=,:)
                Parts = Spliter.Split(Txt, 3)

                Select Case Parts.Length
                    ' Found a value (for the last parameter 
                    ' found (space separator))
                    Case 1
                        If Parameter IsNot Nothing Then
                            If Not Parameters.ContainsKey(Parameter) Then
                                Parts(0) = Remover.Replace(Parts(0), "$1")

                                Parameters.Add(Parameter, Parts(0))
                            End If
                            Parameter = Nothing
                        End If
                        ' else Error: no parameter waiting for a value (skipped)
                        Exit Select

                        ' Found just a parameter
                    Case 2
                        ' The last parameter is still waiting. 
                        ' With no value, set it to true.
                        If Parameter IsNot Nothing Then
                            If Not Parameters.ContainsKey(Parameter) Then
                                Parameters.Add(Parameter, "true")
                            End If
                        End If
                        Parameter = Parts(1)
                        Exit Select

                        ' Parameter with enclosed value
                    Case 3
                        ' The last parameter is still waiting. 
                        ' With no value, set it to true.
                        If Parameter IsNot Nothing Then
                            If Not Parameters.ContainsKey(Parameter) Then
                                Parameters.Add(Parameter, "true")
                            End If
                        End If

                        Parameter = Parts(1)

                        ' Remove possible enclosing characters (",')
                        If Not Parameters.ContainsKey(Parameter) Then
                            Parts(2) = Remover.Replace(Parts(2), "$1")
                            Parameters.Add(Parameter, Parts(2))
                        End If

                        Parameter = Nothing
                        Exit Select
                End Select
            Next
            ' In case a parameter is still waiting
            If Parameter IsNot Nothing Then
                If Not Parameters.ContainsKey(Parameter) Then
                    Parameters.Add(Parameter, "true")
                End If
            End If

            Return Parameters
        End Function

        ' Retrieve a parameter value if it exists 
        ' (overriding C# indexer property)
        Default Public ReadOnly Property Item(Param As String) As String
            Get
                Return (Parameters(Param))
            End Get
        End Property

        Public Shared Function GetSysDatabase() As DAAB
            Return New DAAB("[C-OPTSYS]")
        End Function

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="srcArray"></param>
        ''' <param name="filePath"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' RKP/05-18-12/v3.2.168
        ''' </remarks>
        Public Shared Function SaveToDisk(ByRef srcArray() As Double, ByVal filePath As String) As Integer
            Try
                My.Computer.FileSystem.DeleteFile(filePath)
            Catch ex As Exception

            End Try
            Try
                System.IO.File.WriteAllLines(filePath, Array.ConvertAll(srcArray, New Converter(Of Double, String)(Function(t As Double) t.ToString())))
                Return 0
            Catch ex As Exception
                GenUtils.Message(MsgType.Critical, "SerializeArrayDouble", ex.Message)
                Console.WriteLine("Error converting solver array to table - " & ex.Message)
                Debug.Print("Error converting solver array to table - " & ex.Message)
                EntLib.COPT.Log.Log("C-OPT GenUtils", "SerializeArrayDouble - Error converting solver array to table - " & ex.Message)
                Return -1
            End Try

            Return 0
        End Function

        Public Shared Function SaveToDiskChar(ByRef arrChar() As Char, ByVal filePath As String) As Integer

            Return 0
        End Function

        Public Shared Function GetDBType(ByRef dbTypeName As String) As DAAB.e_DB
            Select Case dbTypeName.Trim().ToUpper()
                Case "ACCESS"
                    Return DAAB.e_DB.e_db_ACCESS
                Case "SQLEXPRESS"
                    Return DAAB.e_DB.e_db_SQLSERVER_EXPRESS
                Case "SQLENT"
                    Return DAAB.e_DB.e_db_SQLSERVER_ENTERPRISE
                Case "SQLCOMPACT"
                    Return DAAB.e_DB.e_db_SQLSERVER_COMPACT
                Case "SQLLOCALDB"
                    Return DAAB.e_DB.e_db_SQLSERVER_LOCALDB
            End Select
        End Function

        ''' <summary>
        ''' Returns TRUE if 64-bit OS, else FALSE.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks>
        ''' RKP/06-22-12/v4.0.170
        ''' </remarks>
        Public Shared ReadOnly Property Is64Bit() As Boolean
            Get
                Return IntPtr.Size = 8
            End Get
        End Property

        ''' <summary>
        ''' Returns bitness (32-bit or 64-bit).
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks>
        ''' RKP/09-11-12/v4.2.178
        ''' </remarks>
        Public Shared ReadOnly Property GetBitnessStr() As String
            Get
                Return IIf(Is64Bit, "64", "32").ToString() & "-bit"
            End Get
        End Property

        Public Shared ReadOnly Property Version() As String
            Get
                Return My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & "." & My.Application.Info.Version.Build '& "." & My.Application.Info.Version.Revision
            End Get
        End Property

        Public Shared ReadOnly Property VersionDesc() As String
            Get
                'Return My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & " Build " & My.Application.Info.Version.Build & " Revision " & My.Application.Info.Version.Revision & " (" & My.Computer.FileSystem.GetFileInfo(My.Application.Info.DirectoryPath & "\" & My.Application.Info.AssemblyName & ".exe").LastWriteTime & ")"
                Return Version() & " (" & My.Computer.FileSystem.GetFileInfo(My.Application.Info.DirectoryPath & "\" & My.Application.Info.AssemblyName & ".exe").LastWriteTime & ")"
            End Get
        End Property

        Public Shared ReadOnly Property GetVersionRevisionStr() As String
            Get
                Return IIf(My.Application.Info.Version.Revision = 0, "", "Rev. " & My.Application.Info.Version.Revision).ToString()
            End Get
        End Property

        Public Shared ReadOnly Property VersionName() As String
            Get
                'Return """Balsam Poplar"" Edition on Win" & IIf(Is64Bit, "64", "32").ToString() '& " / RAM Available= " & GetAvailablePhysicalMemoryStr() & ""
                'Return """Balsam Poplar"" " & IIf(Is64Bit, "64", "32").ToString() & "-bit Edition"
                'Return """Balsam Poplar"" " & GetBitnessStr() & " Edition " & GetVersionRevisionStr()

                'Return """" & GetAppSettings("editionName") & """ " & GetBitnessStr() & " Edition " & GetVersionRevisionStr()

                Return """Palmetto"" " & GetBitnessStr() & " Edition " & GetVersionRevisionStr()
            End Get
        End Property

    End Class 'Public Class GenUtils
End Namespace