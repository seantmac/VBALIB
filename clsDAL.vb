Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Data.Odbc
Imports Microsoft.Office.Interop.Access

Public Class clsDAL
    Public Enum e_DB
        e_db_NONE
        e_db_ACCESS
        e_db_SQLSERVER
        e_db_ORACLE
        e_db_TEXTFILE
        e_db_EXCEL
        e_db_OTHER
    End Enum

    Public Enum e_ConnType
        e_connType_A
        e_connType_B
        e_connType_C
        e_connType_D
        e_connType_E
        e_connType_F
        e_connType_G
        e_connType_H
        e_connType_I
        e_connType_J
        e_connType_K
        e_connType_L
        e_connType_M
        e_connType_N
        e_connType_O
        e_connType_P
        e_connType_Q
        e_connType_R
        e_connType_S
        e_connType_T
        e_connType_U
        e_connType_V
        e_connType_W
        e_connType_X
        e_connType_Y
        e_connType_Z
    End Enum

    Public Enum e_DataClientType
        e_Odbc
        e_SqlServer
        e_SqlServerCe
        e_Oracle
        e_OleDb
        e_Ado
        e_Dao
    End Enum

    Private _dbType As e_DB
    Private _connType As e_ConnType
    Private _connSQLServer As System.Data.SqlClient.SqlConnection
    Private _connOleDb As System.Data.OleDb.OleDbConnection
    Private _connOdbc As System.Data.Odbc.OdbcConnection
    Private _adoConn As ADODB.Connection
    'Private _adoCatalog As ADOX.Catalog
    Private _daoDatabase As Microsoft.Office.Interop.Access.Dao.Database 'Microsoft.Office.Interop.Access.Dao.Database 'RKP/04-20-12/v3.2.166
    Private _useADO As Boolean = False
    Private _useDAO As Boolean = False 'RKP/04-20-12/v3.2.166
    Private _isActive As Boolean = False
    Private _lastErrorNo As Integer = 0
    Private _lastErrorDesc As String = ""

    Public Sub New()
        DBType = e_DB.e_db_NONE
    End Sub

    Public Sub New _
        ( _
            ByVal dataClientType As e_DataClientType, _
            ByVal connStr As String _
        )

        Select Case dataClientType
            Case e_DataClientType.e_Odbc
            Case e_DataClientType.e_SqlServer
            Case e_DataClientType.e_SqlServerCe
            Case e_DataClientType.e_Oracle
            Case e_DataClientType.e_OleDb
            Case e_DataClientType.e_Ado
            Case e_DataClientType.e_Dao
            Case Else
        End Select

    End Sub

    Public Function Connect _
        ( _
            ByVal vnDBType As e_DB, _
            ByVal vnConnType As e_ConnType, _
            ByRef ConnOdbc As System.Data.Odbc.OdbcConnection, _
            ByRef ConnOleDb As System.Data.OleDb.OleDbConnection, _
            ByRef ConnSQLServer As System.Data.SqlClient.SqlConnection, _
            ByRef ConnSQLServerCe As System.Data.SqlServerCe.SqlCeConnection, _
            ByRef ConnADO As ADODB.Connection, _
            ByVal showMsg As Boolean, _
            ByVal ParamArray pArray() As Object _
        ) As Integer

        Dim connStr As String = ""
        'Dim aceDBEngine As Microsoft.Office.Interop.Access.Dao.DBEngine
        'Dim ConnOleDb As System.Data.OleDb.OleDbConnection
        'Dim ConnSQLServer As System.Data.SqlClient.SqlConnection
        'Dim ConnADO As ADODB.Connection

        'UseADO = False
        'DBType = vnDBType
        '_isActive = False

        SetLastError(Nothing)

        DBType = vnDBType

        Select Case vnDBType
            Case e_DB.e_db_ACCESS
                Select Case vnConnType
                    Case e_ConnType.e_connType_A
                        connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & pArray(0).ToString() & ";User Id=admin;Password="
                    Case e_ConnType.e_connType_B
                        If COPT.GenUtils.Is64Bit Then
                            connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & pArray(0).ToString() & ";User Id=" & pArray(1).ToString() & ";Password=" & pArray(2).ToString()
                        Else
                            connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & pArray(0).ToString() & ";User Id=" & pArray(1).ToString() & ";Password=" & pArray(2).ToString()
                        End If
                    Case e_ConnType.e_connType_C
                        If COPT.GenUtils.Is64Bit Then
                            connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & pArray(0).ToString() & ";Jet OLEDB:System Database=" & pArray(1).ToString()
                        Else
                            connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & pArray(0).ToString() & ";Jet OLEDB:System Database=" & pArray(1).ToString()
                        End If
                    Case e_ConnType.e_connType_D
                    Case e_ConnType.e_connType_E
                    Case e_ConnType.e_connType_F
                    Case e_ConnType.e_connType_G
                    Case e_ConnType.e_connType_H
                    Case e_ConnType.e_connType_I
                    Case e_ConnType.e_connType_J
                    Case e_ConnType.e_connType_K
                    Case e_ConnType.e_connType_L
                    Case e_ConnType.e_connType_M
                    Case e_ConnType.e_connType_N
                    Case e_ConnType.e_connType_O
                    Case e_ConnType.e_connType_P
                    Case e_ConnType.e_connType_Q
                    Case e_ConnType.e_connType_Z
                        connStr = pArray(0).ToString()
                    Case Else
                        connStr = pArray(0).ToString()
                End Select

                If Not UseADO Then
                    If UseDAO Then
                        'Do nothing here. Connection opened at the end of this function.
                    Else
                        Try
                            ConnOleDb = New System.Data.OleDb.OleDbConnection(connStr)
                            ConnOleDb.Open()

                            Me.ConnOleDb = ConnOleDb
                            '_isActive = True
                        Catch ex As Exception
                            'GenUtils.Log("Error connecting to:" & vbNewLine & connStr)
                            'GenUtils.Log(ex.Message)
                            LastErrorNo = -1
                            LastErrorDesc = ex.Message
                            'GenUtils.Message(GenUtils.MsgType.Critical, "EntLib - DAAB-Connect", "Error connecting to:" & vbNewLine & connStr)
                            If showMsg Then
                                EntLib.COPT.GenUtils.Message(EntLib.COPT.GenUtils.MsgType.Critical, "EntLib - DAAB-Connect", ex.Message)
                                EntLib.COPT.GenUtils.Message(EntLib.COPT.GenUtils.MsgType.Critical, "EntLib - DAAB-Connect", ex.Message)
                            End If
                        End Try
                    End If

                End If
            Case e_DB.e_db_SQLSERVER
                'For 3.0, the provider string is: Microsoft.SQLSERVER.MOBILE.OLEDB.3.0
                'For 3.5 the provider string is: Microsoft.SQLSERVER.CE.OLEDB.3.5

                Select Case vnConnType
                    Case e_ConnType.e_connType_A  'Trusted Connection
                        'pArray(0) = "SWSQLDEV"
                        'pArray(1) = "tempdb"
                        'msConnStr = "Provider=SQLOLEDB;Data Source=" & pArray(0) & ";Initial Catalog=" & pArray(1) & ";Integrated Security=SSPI"
                        If pArray(0).ToString() = "" Then
                            pArray(0) = "SQLOLEDB"
                        End If
                        connStr = "Provider=" & pArray(0).ToString() & ";Data Source=" & pArray(1).ToString() & ";Initial Catalog=" & pArray(2).ToString() & ";Integrated Security=SSPI"
                    Case e_ConnType.e_connType_B  'Standard Security
                        'pArray(0) = "SWSQLDEV"
                        'pArray(1) = "tempdb"
                        'pArray(1) = "userid"
                        'pArray(2) = "password"
                        'msConnStr = "Provider=SQLOLEDB;Data Source=" & pArray(0) & ";Initial Catalog=" & pArray(1) & ";User Id=" & pArray(2) & ";Password=" & pArray(3)
                        If pArray(0).ToString() = "" Then
                            pArray(0) = "SQLOLEDB"
                        End If
                        connStr = "Provider=" & pArray(0).ToString() & ";Data Source=" & pArray(1).ToString() & ";Initial Catalog=" & pArray(2).ToString() & ";User Id=" & pArray(3).ToString() & ";Password=" & pArray(4).ToString()
                    Case e_ConnType.e_connType_C  'Remote Computer
                        '               oConn.Open "Provider=sqloledb;" & _
                        '               "Network Library=DBMSSOCN;" & _
                        '               "Data Source=xxx.xxx.xxx.xxx,1433;" & _
                        '               "Initial Catalog=myDatabaseName;" & _
                        '               "User ID=myUsername;" & _
                        '               "Password=myPassword"
                        'msConnStr = "Provider=SQLOLEDB;Network Library=" & pArray(0) & ";Data Source=" & pArray(1) & ";Initial Catalog=" & pArray(2) & ";User Id=" & pArray(3) & ";Password=" & pArray(4)
                        If pArray(0).ToString() = "" Then
                            pArray(0) = "SQLOLEDB"
                        End If
                        connStr = "Provider=" & pArray(0).ToString() & ";Network Library=" & pArray(1).ToString() & ";Data Source=" & pArray(2).ToString() & ";Initial Catalog=" & pArray(3).ToString() & ";User Id=" & pArray(4).ToString() & ";Password=" & pArray(5).ToString()
                    Case e_ConnType.e_connType_D
                        'msConnStr = "Provider=SQLNCLI;Server=" & pArray(0).ToString() & ";Database=" & pArray(1).ToString() & ";UID=" & pArray(2).ToString() & ";PWD=" & pArray(3).ToString() & ";"
                        connStr = "Provider=" & pArray(0).ToString() & ";Server=" & pArray(1).ToString() & ";Database=" & pArray(2).ToString() & ";UID=" & pArray(3).ToString() & ";PWD=" & pArray(4).ToString() & ";"
                    Case e_ConnType.e_connType_E
                        If pArray(2).ToString() = "" Then
                            pArray(2) = 1024
                        End If
                        connStr = "Provider=" & pArray(0).ToString() & ";Data Source=" & pArray(1).ToString() & ";Persist Security Info=False;Max Buffer Size=" & pArray(2).ToString() & ";"
                    Case e_ConnType.e_connType_F 'ODBC connection - works in 32-bit and 64-bit Office.
                        'connStr = "Data Source=" & pArray(0).ToString() & ";Initial Catalog=" & pArray(1).ToString() & ";Integrated Security=SSPI"
                        'Driver={SQL Server Native Client 10.0};Server=myServerAddress;Database=myDataBase;Trusted_Connection=yes;
                        connStr = "Driver={" & pArray(0) & "};Server=" & pArray(1) & ";Database=" & pArray(2) & ";Trusted_Connection=Yes;"
                    Case e_ConnType.e_connType_G
                        'Driver={SQL Server Native Client 10.0};Server=myServerAddress;Database=myDataBase;Uid=myUsername;Pwd=myPassword;
                        connStr = "Driver={" & pArray(0) & "};Server=" & pArray(1) & ";Database=" & pArray(2) & ";Uid=" & pArray(3) & ";Pwd=" & pArray(4) & ";"
                    Case e_ConnType.e_connType_H
                        'Use this connection to:
                        'Connect to a SQL Server Express (.mdf) database (SQL Server 2008 and above)
                        'This connection can handle an Express database that resides outside of it's default location.
                        'Default location =
                        'c:\Program Files\Microsoft SQL Server\MSSQL.1\MSSQL\DATA
                        'Usage:
                        'moDAL_SQL.Connect e_db_SQLSERVER, e_connType_H, "SQLNCLI10", VBA.Environ("COMPUTERNAME"), "DatabaseName"
                        connStr = "Provider=" & pArray(0).ToString() & ";Server=" & pArray(1).ToString() & ";Database=" & pArray(2).ToString() & ";Trusted_Connection=Yes;"
                    Case e_ConnType.e_connType_I
                        'Server=.\SQLExpress;AttachDbFilename=c:\asd\qwe\mydbfile.mdf;Database=dbname; Trusted_Connection=Yes;
                        'or
                        'Server=.\SQLExpress;AttachDbFilename=|DataDirectory|mydbfile.mdf; Database=dbname;Trusted_Connection=Yes;
                        connStr = "Server=" & pArray(0).ToString() & ";AttachDbFilename=" & pArray(1).ToString() & ";Database=" & pArray(2).ToString() & "; Trusted_Connection=Yes;"
                    Case e_ConnType.e_connType_J
                        'For SQL Server Compact 3.0
                        'msConnStr = "Provider=Microsoft.SQLSERVER.MOBILE.OLEDB.3.0;Data Source=C:\Northwind.sdf"
                        connStr = "Provider=Microsoft.SQLSERVER.MOBILE.OLEDB.3.0;Data Source=" & pArray(0).ToString()
                    Case e_ConnType.e_connType_K
                        'For SQL Server Compact 3.5
                        'msConnStr = "Provider=Microsoft.SQLSERVER.CE.OLEDB.3.5;Data Source=C:\Northwind.sdf"
                        connStr = "Provider=Microsoft.SQLSERVER.CE.OLEDB.3.5;Data Source=" & pArray(0).ToString()
                    Case e_ConnType.e_connType_L
                        'For SQL Server Compact 3.0/3.5/4.0 and beyond
                        'msConnStr = "Provider=Microsoft.SQLSERVER.CE.OLEDB.3.5;Data Source=C:\Northwind.sdf"
                        connStr = "Provider=" & pArray(0).ToString() & ";Data Source=" & pArray(1).ToString()
                    Case e_ConnType.e_connType_M
                        'Provider=SQLNCLI10;Server=.\SQLExpress;AttachDbFilename=c:\asd\qwe\mydbfile.mdf; Database=dbname; Trusted_Connection=Yes;
                        connStr = "Provider=" & pArray(0).ToString() & ";Server=" & pArray(1).ToString() & ";AttachDbFilename=" & pArray(2).ToString() & "; Database=" & pArray(3).ToString() & "; Trusted_Connection=Yes;"
                    Case e_ConnType.e_connType_N
                        'Provider=SQLNCLI10;Server=.\SQLExpress;AttachDbFilename=|DataDirectory|mydbfile.mdf; Database=dbname;Trusted_Connection=Yes;
                        connStr = "Provider=" & pArray(0).ToString() & ";Server=" & pArray(1).ToString() & ";AttachDbFilename=" & pArray(2).ToString() & "; Database=" & pArray(3).ToString() & ";Trusted_Connection=Yes;"
                    Case e_ConnType.e_connType_Z
                        connStr = pArray(0).ToString()
                    Case Else
                        'moDAL_SQL.Connect e_db_SQLSERVER, -1, "SQLNCLI10", VBA.Environ("COMPUTERNAME"), "DatabaseName"
                        connStr = pArray(0).ToString()
                End Select

                If Not UseADO Then
                    Try
                        ConnSQLServer = New System.Data.SqlClient.SqlConnection(connStr)
                        ConnSQLServer.Open()
                        Me.ConnSQLServer = ConnSQLServer
                        '_isActive = True
                    Catch ex As Exception
                        'GenUtils.Log("Error connecting to:" & vbNewLine & connStr)
                        'GenUtils.Log(ex.Message)
                        SetLastError(ex)
                        If showMsg Then
                            EntLib.COPT.GenUtils.Message(EntLib.COPT.GenUtils.MsgType.Critical, "EntLib - DAAB-Connect", "Error connecting to:" & vbNewLine & connStr)
                            EntLib.COPT.GenUtils.Message(EntLib.COPT.GenUtils.MsgType.Critical, "EntLib - DAAB-Connect", ex.Message)
                        End If
                    End Try
                End If
            Case e_DB.e_db_NONE
                Select Case vnConnType
                    Case e_ConnType.e_connType_A
                        connStr = "Data Source=" & pArray(0).ToString() & ";Initial Catalog=" & pArray(1).ToString() & ";Integrated Security=SSPI;"
                        'connStr = "Provider=SQLOLEDB;Data Source=" & pArray(0).ToString() & ";Initial Catalog=" & pArray(1).ToString() & ";Integrated Security=SSPI"
                        'connStr = "Data Source=" & pArray(0).ToString() & ";Initial Catalog=" & pArray(1).ToString() & ";User Id=" & pArray(2).ToString() & ";Password=" & pArray(3).ToString() & ";"
                    Case e_ConnType.e_connType_B
                        'connStr = "Provider=SQLOLEDB;Data Source=" & pArray(0).ToString() & ";Initial Catalog=" & pArray(1).ToString() & ";User Id=" & pArray(2).ToString() & ";Password=" & pArray(3).ToString()
                        connStr = "Data Source=" & pArray(0).ToString() & ";Initial Catalog=" & pArray(1).ToString() & ";User Id=" & pArray(2).ToString() & ";Password=" & pArray(3).ToString() & ";"
                    Case e_ConnType.e_connType_C

                        connStr = "Provider=SQLNCLI;Server=" & pArray(0).ToString() & ";Database=" & pArray(1).ToString() & ";Trusted_Connection=YES;"
                    Case e_ConnType.e_connType_D
                        connStr = "Provider=SQLNCLI;Server=" & pArray(0).ToString() & ";Database=" & pArray(1).ToString() & ";UID=" & pArray(2).ToString() & ";PWD=" & pArray(3).ToString() & ";"
                    Case e_ConnType.e_connType_E
                        connStr = "Provider=SQLNCLI;Data Source=" & pArray(0).ToString() & ";Persist Security Info=False;Max Buffer Size=1024;"
                    Case e_ConnType.e_connType_F
                    Case e_ConnType.e_connType_G
                    Case e_ConnType.e_connType_H
                        'Use this connection to:
                        'Connect to a SQL Server Express (.mdf) database (SQL Server 2008 and above)
                        'This connection can handle an Express database that resides outside of it's default location.
                        'Default location =
                        'c:\Program Files\Microsoft SQL Server\MSSQL.1\MSSQL\DATA
                        'Usage:
                        'moDAL_SQL.Connect e_db_SQLSERVER, e_connType_H, "SQLNCLI10", VBA.Environ("COMPUTERNAME"), "DatabaseName"
                        connStr = "Provider=" & pArray(0).ToString() & ";Server=" & pArray(1).ToString() & ";Database=" & pArray(2).ToString() & ";Trusted_Connection=Yes;"
                    Case e_ConnType.e_connType_I
                    Case e_ConnType.e_connType_J
                    Case e_ConnType.e_connType_K
                    Case e_ConnType.e_connType_L
                    Case e_ConnType.e_connType_M
                    Case e_ConnType.e_connType_N
                    Case e_ConnType.e_connType_O
                    Case e_ConnType.e_connType_P
                    Case e_ConnType.e_connType_Q
                    Case e_ConnType.e_connType_Z
                        connStr = pArray(0).ToString()
                    Case Else
                        connStr = pArray(0).ToString()
                End Select

                If Not UseADO Then
                    If vnConnType <> e_ConnType.e_connType_E Then
                        Try
                            ConnSQLServer = New System.Data.SqlClient.SqlConnection(connStr)
                            ConnSQLServer.Open()
                            '_isActive = True
                        Catch ex As Exception
                            'GenUtils.Log("Error connecting to:" & vbNewLine & connStr)
                            'GenUtils.Log(ex.Message)
                            SetLastError(ex)
                            If showMsg Then
                                EntLib.COPT.GenUtils.Message(EntLib.COPT.GenUtils.MsgType.Critical, "EntLib - DAAB-Connect", "Error connecting to:" & vbNewLine & connStr)
                                EntLib.COPT.GenUtils.Message(EntLib.COPT.GenUtils.MsgType.Critical, "EntLib - DAAB-Connect", ex.Message)
                            End If
                        End Try
                    Else
                        Try
                            ConnSQLServerCe = New System.Data.SqlServerCe.SqlCeConnection(connStr)
                            ConnSQLServerCe.Open()
                            '_isActive = True
                        Catch ex As Exception
                            'GenUtils.Log("Error connecting to:" & vbNewLine & connStr)
                            'GenUtils.Log(ex.Message)
                            SetLastError(ex)
                            If showMsg Then
                                EntLib.COPT.GenUtils.Message(EntLib.COPT.GenUtils.MsgType.Critical, "EntLib - DAAB-Connect", "Error connecting to:" & vbNewLine & connStr)
                                EntLib.COPT.GenUtils.Message(EntLib.COPT.GenUtils.MsgType.Critical, "EntLib - DAAB-Connect", ex.Message)
                            End If
                        End Try
                    End If
                End If

            Case e_DB.e_db_ORACLE
                Select Case vnConnType
                    Case e_ConnType.e_connType_A
                        connStr = "Provider=OraOLEDB.Oracle;Data Source=" & pArray(0).ToString() & ";User Id=" & pArray(1).ToString() & ";Password=" & pArray(2).ToString()
                    Case e_ConnType.e_connType_B
                        connStr = "Provider=OraOLEDB.Oracle;Data Source=" & pArray(0).ToString() & ";User Id=/;Password="
                    Case e_ConnType.e_connType_C
                        connStr = "Provider=OraOLEDB.Oracle;Data Source=" & pArray(0).ToString() & ";OSAuthent=1"
                    Case e_ConnType.e_connType_D
                        connStr = "Provider=MSDAORA.1;Data Source=" & pArray(0).ToString() & ";User Id=" & pArray(1).ToString() & ";Password=" & pArray(2).ToString()
                    Case e_ConnType.e_connType_E
                        connStr = "Provider=OraOLEDB.Oracle.1;Data Source=" & pArray(0).ToString() & ";User Id=" & pArray(1).ToString() & ";Password=" & pArray(2).ToString()
                    Case e_ConnType.e_connType_F
                        connStr = "Provider=OraOLEDB.Oracle;Data Source=" & pArray(0).ToString() & ";User Id=" & pArray(1).ToString() & ";Password=" & pArray(2).ToString()
                        Try
                            ConnADO = New ADODB.Connection
                            ConnADO.ConnectionString = connStr
                            ConnADO.CommandTimeout = 0
                            ConnADO.CursorLocation = ADODB.CursorLocationEnum.adUseClient
                            ConnADO.Open()
                            'UseADO = True
                            '_isActive = True
                        Catch ex As Exception
                            'GenUtils.Log("Error connecting to:" & vbNewLine & connStr & vbNewLine & "using ADO")
                            'GenUtils.Log(ex.Message)
                            'useADO = False
                            SetLastError(ex)
                            If showMsg Then
                                EntLib.COPT.GenUtils.Message(EntLib.COPT.GenUtils.MsgType.Critical, "EntLib - DAAB-Connect", "Error connecting to:" & vbNewLine & connStr)
                                EntLib.COPT.GenUtils.Message(EntLib.COPT.GenUtils.MsgType.Critical, "EntLib - DAAB-Connect", ex.Message)
                            End If
                        End Try
                    Case e_ConnType.e_connType_G
                        connStr = "Provider=OraOLEDB.Oracle;Data Source=" & pArray(0).ToString() & ";User Id=" & pArray(1).ToString() & ";Password=" & pArray(2).ToString() & ";OLEDB.NET=True;"
                    Case e_ConnType.e_connType_H
                    Case e_ConnType.e_connType_I
                    Case e_ConnType.e_connType_J
                    Case e_ConnType.e_connType_K
                    Case e_ConnType.e_connType_L
                    Case e_ConnType.e_connType_M
                    Case e_ConnType.e_connType_N
                    Case e_ConnType.e_connType_O
                    Case e_ConnType.e_connType_P
                    Case e_ConnType.e_connType_Q
                    Case e_ConnType.e_connType_Z
                        connStr = pArray(0).ToString()
                    Case Else
                        connStr = pArray(0).ToString()
                End Select

                'ConnOdbc = New System.Data.Odbc.OdbcConnection(connStr)
                'ConnOdbc.Open()

                If Not UseADO Then
                    Try
                        ConnOleDb = New System.Data.OleDb.OleDbConnection(connStr)
                        ConnOleDb.Open()
                        '_isActive = True
                    Catch ex As Exception
                        'GenUtils.Log("Error connecting to:" & vbNewLine & connStr)
                        'GenUtils.Log(ex.Message)
                        SetLastError(ex)
                        If showMsg Then
                            EntLib.COPT.GenUtils.Message(EntLib.COPT.GenUtils.MsgType.Critical, "EntLib - DAAB-Connect", "Error connecting to:" & vbNewLine & connStr)
                            EntLib.COPT.GenUtils.Message(EntLib.COPT.GenUtils.MsgType.Critical, "EntLib - DAAB-Connect", ex.Message)
                        End If
                    End Try
                End If
            Case e_DB.e_db_TEXTFILE
                Select Case vnConnType
                    Case e_ConnType.e_connType_A
                        connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & pArray(0).ToString() & ";Extended Properties=""text;HDR=Yes;FMT=Delimited"""
                    Case e_ConnType.e_connType_B
                        connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & pArray(0).ToString() & ";Extended Properties=""text;HDR=No;FMT=Delimited"""
                    Case e_ConnType.e_connType_C
                        connStr = "Driver={Microsoft Text Driver (*.txt; *.csv)};DBQ=" & pArray(0).ToString() & ";Extensions=asc,csv,tab,txt;"
                    Case e_ConnType.e_connType_D
                    Case e_ConnType.e_connType_E
                    Case e_ConnType.e_connType_F
                    Case e_ConnType.e_connType_G
                    Case e_ConnType.e_connType_H
                    Case e_ConnType.e_connType_I
                    Case e_ConnType.e_connType_J
                    Case e_ConnType.e_connType_K
                    Case e_ConnType.e_connType_L
                    Case e_ConnType.e_connType_M
                    Case e_ConnType.e_connType_N
                    Case e_ConnType.e_connType_O
                    Case e_ConnType.e_connType_P
                    Case e_ConnType.e_connType_Q
                    Case e_ConnType.e_connType_Z
                        connStr = pArray(0).ToString()
                    Case Else
                        connStr = pArray(0).ToString()
                End Select

                If Not UseADO Then
                    Try
                        ConnOleDb = New System.Data.OleDb.OleDbConnection(connStr)
                        ConnOleDb.Open()
                        '_isActive = True
                    Catch ex As Exception
                        'GenUtils.Log("Error connecting to:" & vbNewLine & connStr)
                        'GenUtils.Log(ex.Message)
                        SetLastError(ex)
                        If showMsg Then
                            EntLib.COPT.GenUtils.Message(EntLib.COPT.GenUtils.MsgType.Critical, "EntLib - DAAB-Connect", "Error connecting to:" & vbNewLine & connStr)
                            EntLib.COPT.GenUtils.Message(EntLib.COPT.GenUtils.MsgType.Critical, "EntLib - DAAB-Connect", ex.Message)
                        End If
                    End Try
                End If
            Case e_DB.e_db_EXCEL
                Select Case vnConnType
                    Case e_ConnType.e_connType_A
                        connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & pArray(0).ToString() & ";Extended Properties=""Excel 8.0;HDR=Yes"""
                    Case e_ConnType.e_connType_B
                    Case e_ConnType.e_connType_C
                    Case e_ConnType.e_connType_D
                    Case e_ConnType.e_connType_E
                    Case e_ConnType.e_connType_F
                    Case e_ConnType.e_connType_G
                    Case e_ConnType.e_connType_H
                    Case e_ConnType.e_connType_I
                    Case e_ConnType.e_connType_J
                    Case e_ConnType.e_connType_K
                    Case e_ConnType.e_connType_L
                    Case e_ConnType.e_connType_M
                    Case e_ConnType.e_connType_N
                    Case e_ConnType.e_connType_O
                    Case e_ConnType.e_connType_P
                    Case e_ConnType.e_connType_Q
                    Case e_ConnType.e_connType_Z
                        connStr = pArray(0).ToString()
                    Case Else
                        connStr = pArray(0).ToString()
                End Select

                If Not UseADO Then
                    Try
                        ConnOleDb = New System.Data.OleDb.OleDbConnection(connStr)
                        ConnOleDb.Open()
                        '_isActive = True
                    Catch ex As Exception
                        'GenUtils.Log("Error connecting to:" & vbNewLine & connStr)
                        'GenUtils.Log(ex.Message)
                        SetLastError(ex)
                        If showMsg Then
                            EntLib.COPT.GenUtils.Message(EntLib.COPT.GenUtils.MsgType.Critical, "EntLib - DAAB-Connect", "Error connecting to:" & vbNewLine & connStr)
                            EntLib.COPT.GenUtils.Message(EntLib.COPT.GenUtils.MsgType.Critical, "EntLib - DAAB-Connect", ex.Message)
                        End If
                    End Try
                End If
            Case e_DB.e_db_OTHER
                Select Case vnConnType
                    Case e_ConnType.e_connType_A
                        connStr = pArray(0).ToString()
                    Case e_ConnType.e_connType_B
                    Case e_ConnType.e_connType_C
                    Case e_ConnType.e_connType_D
                    Case e_ConnType.e_connType_E
                    Case e_ConnType.e_connType_F
                    Case e_ConnType.e_connType_G
                    Case e_ConnType.e_connType_H
                    Case e_ConnType.e_connType_I
                    Case e_ConnType.e_connType_J
                    Case e_ConnType.e_connType_K
                    Case e_ConnType.e_connType_L
                    Case e_ConnType.e_connType_M
                    Case e_ConnType.e_connType_N
                    Case e_ConnType.e_connType_O
                    Case e_ConnType.e_connType_P
                    Case e_ConnType.e_connType_Q
                    Case e_ConnType.e_connType_Z
                        connStr = pArray(0).ToString()
                    Case Else
                End Select
            Case Else
        End Select

        If UseADO Then
            Try
                ConnADO = New ADODB.Connection
                ConnADO.ConnectionString = connStr
                ConnADO.ConnectionTimeout = 0
                ConnADO.CursorLocation = ADODB.CursorLocationEnum.adUseClient
                ConnADO.Open()
                Me.ConnADO = ConnADO

                'ConnADOCatalog = New ADOX.Catalog
                'ConnADOCatalog.ActiveConnection = ConnADO
            Catch ex As Exception
                SetLastError(ex)
                If showMsg Then
                    EntLib.COPT.GenUtils.Message(EntLib.COPT.GenUtils.MsgType.Critical, "clsDAL-Connect", "Error connecting to:" & vbNewLine & connStr)
                    EntLib.COPT.GenUtils.Message(EntLib.COPT.GenUtils.MsgType.Critical, "clsDAL-Connect", ex.Message)
                End If
            End Try
        End If

        'RKP/04-20-12/v3.2.166
        If UseDAO Then
            'Dim daoDBEngine As dao.DBEngine
            'daoDBEngine = New dao.DBEngine
            '_daoDatabase = daoDBEngine.OpenDatabase(CStr(pArray(0)))
            Try
                _daoDatabase = (New Microsoft.Office.Interop.Access.Dao.DBEngine).OpenDatabase(CStr(pArray(0)))
            Catch ex As Exception
                SetLastError(ex)
                If showMsg Then
                    EntLib.COPT.GenUtils.Message(EntLib.COPT.GenUtils.MsgType.Critical, "clsDAL-Connect", "Error connecting (using DAO) to:" & vbNewLine & connStr)
                    EntLib.COPT.GenUtils.Message(EntLib.COPT.GenUtils.MsgType.Critical, "clsDAL-Connect", ex.Message)
                End If
            End Try
        End If

        Return 0
    End Function

    Private Property DBType() As e_DB
        Get
            Return _dbType
        End Get
        Set(ByVal value As e_DB)
            _dbType = value
        End Set
    End Property

    Public Property ConnSQLServer() As System.Data.SqlClient.SqlConnection
        Get
            Return _connSQLServer
        End Get
        Set(ByVal value As System.Data.SqlClient.SqlConnection)
            _connSQLServer = value
        End Set
    End Property

    Public Property ConnOleDb() As System.Data.OleDb.OleDbConnection
        Get
            Return _connOleDb
        End Get
        Set(ByVal value As System.Data.OleDb.OleDbConnection)
            _connOleDb = value
        End Set
    End Property

    Public Property ConnOdbc() As System.Data.Odbc.OdbcConnection
        Get
            Return _connOdbc
        End Get
        Set(ByVal value As System.Data.Odbc.OdbcConnection)
            _connOdbc = value
        End Set
    End Property

    Public Property ConnADO() As ADODB.Connection
        Get
            Return _adoConn
        End Get
        Set(ByVal value As ADODB.Connection)
            _adoConn = value
        End Set
    End Property

    'Public Property ConnADOCatalog() As ADOX.Catalog
    '    Get
    '        Return _adoCatalog
    '    End Get
    '    Set(ByVal value As ADOX.Catalog)
    '        _adoCatalog = value
    '    End Set
    'End Property

    Public Property UseADO() As Boolean
        Get
            Return _useADO
        End Get
        Set(ByVal value As Boolean)
            _useADO = value
            If value Then
                UseDAO = False 'RKP/04-20-12/v3.2.166
            End If
        End Set
    End Property

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>
    ''' RKP/04-20-12/v3.2.166
    ''' </remarks>
    Public Property UseDAO() As Boolean
        Get
            Return _useDAO
        End Get
        Set(ByVal value As Boolean)
            _useDAO = value
            If value Then
                UseADO = False  'RKP/04-20-12/v3.2.166
            End If

        End Set
    End Property

    Public Function Execute(ByRef sql As String, ByRef dt As DataTable, ByVal ParamArray pArray() As Object) As Integer

        Dim adapterOleDb As System.Data.OleDb.OleDbDataAdapter
        Dim adapterSql As System.Data.SqlClient.SqlDataAdapter
        'Dim adapterOdbc As System.Data.Odbc.OdbcDataAdapter
        Dim cmdSQL As System.Data.SqlClient.SqlCommand
        Dim adoRS As ADODB.Recordset
        Dim ret As Integer = 0
        Dim retStr As String = "0"
        Dim bulkCopy As SqlBulkCopy
        Dim ctr1 As Integer

        ret = -1
        LastErrorNo = -1
        LastErrorDesc = ""

        If Not UseADO Then
            Select Case DBType
                Case e_DB.e_db_ACCESS
                    Try
                        adapterOleDb = New System.Data.OleDb.OleDbDataAdapter(sql, ConnOleDb)
                        dt = New DataTable
                        adapterOleDb.Fill(dt)
                        If Not dt Is Nothing Then
                            ret = dt.Rows.Count
                        Else
                            ret = 0
                        End If
                    Catch ex As Exception
                        EntLib.COPT.Log.Log("Error executing SQL:" & vbNewLine & sql)
                        EntLib.COPT.Log.Log(ex.Message)
                        ret = -1
                        LastErrorDesc = ex.Message
                    End Try
                    LastErrorNo = ret
                Case e_DB.e_db_SQLSERVER
                    'Select  'Case sql.Trim().ToUpper().Substring(0, 6)
                    '    Case sql.Trim().ToUpper().Substring(0, 6).Equals("SELECT")
                    '        Try
                    '            adapterSql = New System.Data.SqlClient.SqlDataAdapter(sql, ConnSQLServer)
                    '            dt = New DataTable
                    '            adapterSql.Fill(dt)
                    '            If Not dt Is Nothing Then
                    '                ret = dt.Rows.Count
                    '            Else
                    '                ret = 0
                    '            End If
                    '        Catch ex As Exception
                    '            GenUtils.Log("Error executing SQL:" & vbNewLine & sql)
                    '            GenUtils.Log(ex.Message)
                    '        End Try
                    '    Case sql.Trim().ToUpper().Substring(0, 4).Equals("EXEC")

                    '    Case Else
                    '        cmdSQL = New System.Data.SqlClient.SqlCommand("SQLServer", ConnSQLServer)
                    '        cmdSQL.CommandType = System.Data.CommandType.Text
                    '        cmdSQL.CommandText = sql
                    '        Try
                    '            ret = cmdSQL.ExecuteNonQuery()
                    '        Catch ex As Exception
                    '            GenUtils.Log("Error executing SQL:" & vbNewLine & sql)
                    '            GenUtils.Log(ex.Message)
                    '        End Try
                    'End Select

                    If sql.Trim().ToUpper().Substring(0, 6).Equals("SELECT") Then
                        Try
                            adapterSql = New System.Data.SqlClient.SqlDataAdapter(sql, ConnSQLServer)
                            dt = New DataTable
                            adapterSql.Fill(dt)
                            If Not dt Is Nothing Then
                                ret = dt.Rows.Count
                            Else
                                ret = 0
                            End If
                        Catch ex As Exception
                            EntLib.COPT.Log.Log("Error executing SQL:" & vbNewLine & sql)
                            EntLib.COPT.Log.Log(ex.Message)
                            ret = -1
                            LastErrorDesc = ex.Message
                        End Try
                        LastErrorNo = ret
                    ElseIf sql.Trim().ToUpper().Substring(0, 4).Equals("EXEC") Then

                        Try
                            cmdSQL = New System.Data.SqlClient.SqlCommand("SQLServer", ConnSQLServer)
                            cmdSQL.CommandType = CommandType.StoredProcedure
                            cmdSQL.CommandTimeout = 0
                            cmdSQL.CommandText = sql.Substring(5).Split(CChar(" "))(0)
                            If pArray IsNot Nothing Then
                                If pArray.Length > 0 Then
                                    For ctr As Integer = 0 To pArray.Length - 1 Step 3
                                        cmdSQL.Parameters.AddWithValue(pArray(ctr).ToString(), pArray(ctr + 1).ToString())
                                        cmdSQL.Parameters(ctr).SqlDbType = CType(CInt(pArray(ctr + 2)), SqlDbType)
                                    Next
                                End If
                            End If
                            cmdSQL.CommandTimeout = 0
                            ret = cmdSQL.ExecuteNonQuery()
                        Catch ex As Exception
                            EntLib.COPT.Log.Log("Error executing SQL:" & vbNewLine & sql)
                            EntLib.COPT.Log.Log(ex.Message)
                            EntLib.COPT.Log.Log(ex.Source)
                            ret = -1
                            LastErrorDesc = ex.Message
                        End Try
                        LastErrorNo = ret
                    ElseIf sql.Trim().ToUpper().Contains("BULKCOPY") Then
                        bulkCopy = New SqlBulkCopy(ConnSQLServer)
                        bulkCopy.BulkCopyTimeout = 0
                        bulkCopy.DestinationTableName = pArray(0).ToString()
                        If sql.Trim().ToUpper().Contains("MAP") Then
                            bulkCopy.ColumnMappings.Clear()
                            For ctr1 = 0 To dt.Columns.Count - 1
                                bulkCopy.ColumnMappings.Add(dt.Columns(ctr1).ColumnName, dt.Columns(ctr1).ColumnName)
                            Next
                        End If
                        Try
                            bulkCopy.WriteToServer(dt)

                            ret = 0
                        Catch ex As Exception
                            EntLib.COPT.Log.Log("Error executing BULKCOPY: " & pArray(0).ToString())
                            EntLib.COPT.Log.Log("Error msg: " & ex.Message)
                            EntLib.COPT.Log.Log("Error source: " & ex.Source)
                            ret = -1
                            LastErrorDesc = ex.Message
                        End Try
                        LastErrorNo = ret
                    Else
                        cmdSQL = New System.Data.SqlClient.SqlCommand("SQLServer", ConnSQLServer)
                        cmdSQL.CommandType = System.Data.CommandType.Text
                        cmdSQL.CommandTimeout = 0
                        cmdSQL.CommandText = sql
                        Try
                            ret = cmdSQL.ExecuteNonQuery()
                            If ret = -1 Then
                                ret = 0
                            End If
                        Catch ex As Exception
                            EntLib.COPT.Log.Log("Error executing SQL:" & vbNewLine & sql)
                            EntLib.COPT.Log.Log(ex.Message)
                            ret = -1
                            LastErrorDesc = ex.Message
                        End Try
                        LastErrorNo = ret
                    End If
                Case e_DB.e_db_ORACLE
                    Try
                        adapterOleDb = New System.Data.OleDb.OleDbDataAdapter(sql, ConnOleDb)

                        dt = New DataTable
                        adapterOleDb.Fill(dt)
                        If Not dt Is Nothing Then
                            ret = dt.Rows.Count
                        Else
                            ret = 0
                        End If
                    Catch ex As Exception
                        EntLib.COPT.Log.Log("Error executing SQL:" & vbNewLine & sql)
                        EntLib.COPT.Log.Log(ex.Message)
                        ret = -1
                        LastErrorDesc = ex.Message
                    End Try
                    LastErrorNo = ret
                Case e_DB.e_db_EXCEL
                Case e_DB.e_db_TEXTFILE
                Case e_DB.e_db_NONE
                Case e_DB.e_db_OTHER
                Case Else
            End Select
        Else
            Try
                If ConnADO Is Nothing Then
                    EntLib.COPT.Log.Log("Error executing SQL:" & vbNewLine & sql)
                    EntLib.COPT.Log.Log("No ADODB.Connection found.")
                Else
                    If ConnADO.State = ADODB.ObjectStateEnum.adStateOpen Then
                        If sql.StartsWith("SELECT") Then
                            adoRS = New ADODB.Recordset
                            adoRS.ActiveConnection = ConnADO
                            adoRS.CursorType = ADODB.CursorTypeEnum.adOpenKeyset
                            adoRS.LockType = ADODB.LockTypeEnum.adLockBatchOptimistic
                            adoRS.Open(sql, ConnADO, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockBatchOptimistic, ADODB.CommandTypeEnum.adCmdText)
                            'sql = ""
                            If Not adoRS Is Nothing Then
                                If adoRS.State = ADODB.ObjectStateEnum.adStateOpen Then
                                    If Not (adoRS.BOF And adoRS.EOF) Then
                                        ret = adoRS.RecordCount
                                        sql = adoRS.Fields(0).Value.ToString() & ""

                                        dt = ConvertADORecordsetToDataTable(adoRS, "MyTable")
                                        'sql = adoRS.Fields(0).Value.ToString() & ""
                                    End If
                                End If
                            End If
                        Else
                            adoRS = ConnADO.Execute(sql, retStr.ToString())
                            ret = CInt(retStr)
                        End If
                    Else
                        EntLib.COPT.Log.Log("Error executing SQL:" & vbNewLine & sql)
                        EntLib.COPT.Log.Log("ADODB.Connection.State is not OPEN.")
                        ret = -1
                        LastErrorDesc = "Error executing SQL:" & vbNewLine & sql & vbNewLine & "ADODB.Connection.State is not OPEN."
                    End If
                    LastErrorNo = ret
                End If
            Catch ex As Exception
                EntLib.COPT.Log.Log("Error executing SQL:" & vbNewLine & sql & vbNewLine & "using ADO.")
                EntLib.COPT.Log.Log(ex.Message)
                ret = -1
                LastErrorDesc = ex.Message
            End Try
        End If

        Return ret
    End Function

    Public Function ConvertADORecordsetToDataTable(ByRef adoRS As ADODB.Recordset, ByVal tableName As String) As DataTable
        Dim ds As New DataSet(tableName)
        'Dim dt As DataTable
        Dim da As New System.Data.OleDb.OleDbDataAdapter

        da.Fill(ds, adoRS, tableName)

        Return ds.Tables(0)
    End Function

    Public Shared Function PrepareSQL(ByVal dr As DataRow, ByVal colCtr As Integer) As String
        Dim sql As String = ""

        Select Case dr.Item(colCtr).GetType().ToString()
            Case "System.String"
                sql = sql & "'" & dr.Item(colCtr).ToString().Replace("'", "''").Trim() & "',"
            Case "System.DateTime"
                sql = sql & "'" & dr.Item(colCtr).ToString() & "',"
            Case "System.DBNull"
                sql = sql & "NULL,"
            Case Else
                sql = sql & dr.Item(colCtr).ToString() & ","
        End Select

        Return sql
    End Function


    Public ReadOnly Property IsActive() As Boolean
        Get
            Return _isActive
        End Get
    End Property

    'RKP/07-28-11/v2.5.147
    Public Property LastErrorNo() As Integer
        Get
            Return _lastErrorNo
        End Get
        Set(ByVal value As Integer)
            _lastErrorNo = value
        End Set
    End Property

    'RKP/07-28-11/v2.5.147
    Public Property LastErrorDesc() As String
        Get
            Return _lastErrorDesc
        End Get
        Set(ByVal value As String)
            _lastErrorDesc = value
        End Set
    End Property

    ''RKP/07-28-11/v2.5.147
    'Public Sub SetLastError(ByVal errorNo As Integer)
    '    LastErrorNo = errorNo
    '    'If errorNo = 0 Then
    '    '    LastErrorDesc = ""
    '    'End If
    'End Sub

    'RKP/07-28-11/v2.5.147
    Public Sub SetLastError(ByRef ex As Exception)
        If ex Is Nothing Then
            LastErrorNo = 0
            LastErrorDesc = ""
        Else
            LastErrorNo = -1
            LastErrorDesc = ex.Message
        End If
    End Sub

    'RKP/07-28-11/v2.5.147
    Public Sub SetLastError(ByVal errorNo As Integer, ByRef ex As Exception)
        LastErrorNo = errorNo
        LastErrorDesc = ex.Message
    End Sub

    Public Function GetConnectionString( _
            ByVal vnDBType As e_DB, _
            ByVal vnConnType As e_ConnType, _
            ByVal ParamArray pArray() As Object _
    ) As String

        Dim connStr As String = ""

        Select Case vnDBType
            Case e_DB.e_db_ACCESS
                Select Case vnConnType
                    Case e_ConnType.e_connType_A
                        connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & pArray(0).ToString() & ";User Id=admin;Password="
                    Case e_ConnType.e_connType_B
                        connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & pArray(0).ToString() & ";User Id=" & pArray(1).ToString() & ";Password=" & pArray(2).ToString()
                    Case e_ConnType.e_connType_C
                        connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & pArray(0).ToString() & ";Jet OLEDB:System Database=" & pArray(1).ToString()
                    Case e_ConnType.e_connType_D
                    Case e_ConnType.e_connType_E
                    Case e_ConnType.e_connType_F
                    Case e_ConnType.e_connType_G
                    Case e_ConnType.e_connType_H
                    Case e_ConnType.e_connType_I
                    Case e_ConnType.e_connType_J
                    Case e_ConnType.e_connType_K
                    Case e_ConnType.e_connType_L
                    Case e_ConnType.e_connType_M
                    Case e_ConnType.e_connType_N
                    Case e_ConnType.e_connType_O
                    Case e_ConnType.e_connType_P
                    Case e_ConnType.e_connType_Q
                    Case e_ConnType.e_connType_R
                    Case e_ConnType.e_connType_S
                    Case e_ConnType.e_connType_T
                    Case e_ConnType.e_connType_U
                    Case e_ConnType.e_connType_V
                    Case e_ConnType.e_connType_W
                    Case e_ConnType.e_connType_X
                    Case e_ConnType.e_connType_Y
                    Case e_ConnType.e_connType_Z
                    Case Else
                        connStr = pArray(0).ToString()
                End Select
            Case e_DB.e_db_SQLSERVER
                Select Case vnConnType
                    Case e_ConnType.e_connType_A
                    Case e_ConnType.e_connType_B
                    Case e_ConnType.e_connType_C
                    Case e_ConnType.e_connType_D
                    Case e_ConnType.e_connType_E
                    Case e_ConnType.e_connType_F
                    Case e_ConnType.e_connType_G
                    Case e_ConnType.e_connType_H
                    Case e_ConnType.e_connType_I
                    Case e_ConnType.e_connType_J
                    Case e_ConnType.e_connType_K
                    Case e_ConnType.e_connType_L
                    Case e_ConnType.e_connType_M
                    Case e_ConnType.e_connType_N
                    Case e_ConnType.e_connType_O
                    Case e_ConnType.e_connType_P
                    Case e_ConnType.e_connType_Q
                    Case e_ConnType.e_connType_R
                    Case e_ConnType.e_connType_S
                    Case e_ConnType.e_connType_T
                    Case e_ConnType.e_connType_U
                    Case e_ConnType.e_connType_V
                    Case e_ConnType.e_connType_W
                    Case e_ConnType.e_connType_X
                    Case e_ConnType.e_connType_Y
                    Case e_ConnType.e_connType_Z
                    Case Else
                End Select
            Case e_DB.e_db_ORACLE
                Select Case vnConnType
                    Case e_ConnType.e_connType_A
                    Case e_ConnType.e_connType_B
                    Case e_ConnType.e_connType_C
                    Case e_ConnType.e_connType_D
                    Case e_ConnType.e_connType_E
                    Case e_ConnType.e_connType_F
                    Case e_ConnType.e_connType_G
                    Case e_ConnType.e_connType_H
                    Case e_ConnType.e_connType_I
                    Case e_ConnType.e_connType_J
                    Case e_ConnType.e_connType_K
                    Case e_ConnType.e_connType_L
                    Case e_ConnType.e_connType_M
                    Case e_ConnType.e_connType_N
                    Case e_ConnType.e_connType_O
                    Case e_ConnType.e_connType_P
                    Case e_ConnType.e_connType_Q
                    Case e_ConnType.e_connType_R
                    Case e_ConnType.e_connType_S
                    Case e_ConnType.e_connType_T
                    Case e_ConnType.e_connType_U
                    Case e_ConnType.e_connType_V
                    Case e_ConnType.e_connType_W
                    Case e_ConnType.e_connType_X
                    Case e_ConnType.e_connType_Y
                    Case e_ConnType.e_connType_Z
                    Case Else
                End Select
            Case e_DB.e_db_EXCEL
                Select Case vnConnType
                    Case e_ConnType.e_connType_A
                    Case e_ConnType.e_connType_B
                    Case e_ConnType.e_connType_C
                    Case e_ConnType.e_connType_D
                    Case e_ConnType.e_connType_E
                    Case e_ConnType.e_connType_F
                    Case e_ConnType.e_connType_G
                    Case e_ConnType.e_connType_H
                    Case e_ConnType.e_connType_I
                    Case e_ConnType.e_connType_J
                    Case e_ConnType.e_connType_K
                    Case e_ConnType.e_connType_L
                    Case e_ConnType.e_connType_M
                    Case e_ConnType.e_connType_N
                    Case e_ConnType.e_connType_O
                    Case e_ConnType.e_connType_P
                    Case e_ConnType.e_connType_Q
                    Case e_ConnType.e_connType_R
                    Case e_ConnType.e_connType_S
                    Case e_ConnType.e_connType_T
                    Case e_ConnType.e_connType_U
                    Case e_ConnType.e_connType_V
                    Case e_ConnType.e_connType_W
                    Case e_ConnType.e_connType_X
                    Case e_ConnType.e_connType_Y
                    Case e_ConnType.e_connType_Z
                    Case Else
                End Select
            Case e_DB.e_db_TEXTFILE
                Select Case vnConnType
                    Case e_ConnType.e_connType_A
                    Case e_ConnType.e_connType_B
                    Case e_ConnType.e_connType_C
                    Case e_ConnType.e_connType_D
                    Case e_ConnType.e_connType_E
                    Case e_ConnType.e_connType_F
                    Case e_ConnType.e_connType_G
                    Case e_ConnType.e_connType_H
                    Case e_ConnType.e_connType_I
                    Case e_ConnType.e_connType_J
                    Case e_ConnType.e_connType_K
                    Case e_ConnType.e_connType_L
                    Case e_ConnType.e_connType_M
                    Case e_ConnType.e_connType_N
                    Case e_ConnType.e_connType_O
                    Case e_ConnType.e_connType_P
                    Case e_ConnType.e_connType_Q
                    Case e_ConnType.e_connType_R
                    Case e_ConnType.e_connType_S
                    Case e_ConnType.e_connType_T
                    Case e_ConnType.e_connType_U
                    Case e_ConnType.e_connType_V
                    Case e_ConnType.e_connType_W
                    Case e_ConnType.e_connType_X
                    Case e_ConnType.e_connType_Y
                    Case e_ConnType.e_connType_Z
                    Case Else
                End Select
            Case e_DB.e_db_OTHER
                Select Case vnConnType
                    Case e_ConnType.e_connType_A
                    Case e_ConnType.e_connType_B
                    Case e_ConnType.e_connType_C
                    Case e_ConnType.e_connType_D
                    Case e_ConnType.e_connType_E
                    Case e_ConnType.e_connType_F
                    Case e_ConnType.e_connType_G
                    Case e_ConnType.e_connType_H
                    Case e_ConnType.e_connType_I
                    Case e_ConnType.e_connType_J
                    Case e_ConnType.e_connType_K
                    Case e_ConnType.e_connType_L
                    Case e_ConnType.e_connType_M
                    Case e_ConnType.e_connType_N
                    Case e_ConnType.e_connType_O
                    Case e_ConnType.e_connType_P
                    Case e_ConnType.e_connType_Q
                    Case e_ConnType.e_connType_R
                    Case e_ConnType.e_connType_S
                    Case e_ConnType.e_connType_T
                    Case e_ConnType.e_connType_U
                    Case e_ConnType.e_connType_V
                    Case e_ConnType.e_connType_W
                    Case e_ConnType.e_connType_X
                    Case e_ConnType.e_connType_Y
                    Case e_ConnType.e_connType_Z
                    Case Else
                End Select
            Case e_DB.e_db_NONE
            Case Else

        End Select

        Return connStr
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="adoDataType"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' RKP/08-22-11
    ''' </remarks>
    Public Shared Function GetADOToSQLDataType(ByVal adoDataType As String) As String
        Select Case adoDataType.ToUpper()
            Case "ADINTEGER"
                Return "INT"
            Case "ADDOUBLE"
                Return "FLOAT"
            Case "ADVARWCHAR"
                Return "NVARCHAR"
            Case "ADCURRENCY"
                Return "MONEY"
            Case "ADBOOLEAN"
                Return "BIT"
            Case "ADBIGINT"
                Return "BIGINT"
            Case "ADBINARY"
                Return "BINARY"
            Case "ADCHAR"
                Return "CHAR"
            Case "ADDATE"
                Return "DATETIME"
            Case "ADDBTIMESTAMP"
                Return "DATETIME"
            Case "ADGUID"
                Return "UNIQUEIDENTIFIER"
            Case "ADLONGVARBINARY"
                Return "IMAGE"
            Case "ADLONGVARCHAR"
                Return "TEXT"
            Case "ADLONGVARWCHAR"
                Return "TEXT"
            Case "ADNUMERIC"
                Return "DECIMAL"
            Case "ADSINGLE"
                Return "REAL"
            Case "ADSMALLINT"
                Return "SMALLINT"
            Case "ADUNSIGNEDTINYINT"
                Return "TINYINT"
            Case "ADVARCHAR"
                Return "VARCHAR"
            Case "ADVARIANT"
                Return "SQL_VARIANT"
            Case "ADWCHAR"
                Return "NCHAR"
            Case Else
                Return "N/A"
        End Select
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="adoDataType"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' RKP/04-20-12/v3.2.166
    ''' </remarks>
    Public Shared Function GetDAOToSQLDataType(ByVal adoDataType As String) As String
        Select Case CInt(adoDataType) '.ToUpper()
            Case dao.DataTypeEnum.dbBigInt '"ADINTEGER"
                Return "INT"
            Case dao.DataTypeEnum.dbBinary  '"ADDOUBLE"
                Return "BINARY"
                'Return "FLOAT"
            Case dao.DataTypeEnum.dbBoolean
                Return "BIT"
            Case dao.DataTypeEnum.dbByte
                Return "TINYINT"
            Case dao.DataTypeEnum.dbChar
                Return "CHAR"
            Case dao.DataTypeEnum.dbCurrency
                Return "MONEY"
            Case dao.DataTypeEnum.dbDate
                Return "DATETIME"
            Case dao.DataTypeEnum.dbDecimal
                Return "DECIMAL"
            Case dao.DataTypeEnum.dbDouble
                Return "FLOAT"
            Case dao.DataTypeEnum.dbFloat
                Return "FLOAT"
            Case dao.DataTypeEnum.dbGUID
                Return "UNIQUEIDENTIFIER"
            Case dao.DataTypeEnum.dbInteger
                Return "INT"
            Case dao.DataTypeEnum.dbLong
                Return "INT"
            Case dao.DataTypeEnum.dbLongBinary
                Return "IMAGE"
            Case dao.DataTypeEnum.dbMemo
                Return "TEXT"
            Case dao.DataTypeEnum.dbNumeric
                Return "DECIMAL"
            Case dao.DataTypeEnum.dbSingle
                Return "REAL"
            Case dao.DataTypeEnum.dbText
                Return "NVARCHAR"
            Case dao.DataTypeEnum.dbTime
                Return "DATETIME"
            Case dao.DataTypeEnum.dbTimeStamp
                Return "DATETIME"
            Case dao.DataTypeEnum.dbVarBinary
                Return "IMAGE"
            Case Else
                Return "N/A"
                'Case "ADVARWCHAR"
                '    Return "NVARCHAR"
                'Case "ADCURRENCY"
                '    Return "MONEY"
                'Case "ADBOOLEAN"
                '    Return "BIT"
                'Case "ADBIGINT"
                '    Return "BIGINT"
                'Case "ADBINARY"
                '    Return "BINARY"
                'Case "ADCHAR"
                '    Return "CHAR"
                'Case "ADDATE"
                '    Return "DATETIME"
                'Case "ADDBTIMESTAMP"
                '    Return "DATETIME"
                'Case "ADGUID"
                '    Return "UNIQUEIDENTIFIER"
                'Case "ADLONGVARBINARY"
                '    Return "IMAGE"
                'Case "ADLONGVARCHAR"
                '    Return "TEXT"
                'Case "ADLONGVARWCHAR"
                '    Return "TEXT"
                'Case "ADNUMERIC"
                '    Return "DECIMAL"
                'Case "ADSINGLE"
                '    Return "REAL"
                'Case "ADSMALLINT"
                '    Return "SMALLINT"
                'Case "ADUNSIGNEDTINYINT"
                '    Return "TINYINT"
                'Case "ADVARCHAR"
                '    Return "VARCHAR"
                'Case "ADVARIANT"
                '    Return "SQL_VARIANT"
                'Case "ADWCHAR"
                '    Return "NCHAR"
                'Case Else

        End Select
    End Function

    Public ReadOnly Property DAODatabase() As Microsoft.Office.Interop.Access.Dao.Database
        Get
            Return _daoDatabase
        End Get
    End Property

End Class

