Imports Microsoft.Practices.EnterpriseLibrary.Common
Imports Microsoft.Practices.EnterpriseLibrary.Common.Configuration
Imports Microsoft.Practices.EnterpriseLibrary.Data
Imports Microsoft.Practices.EnterpriseLibrary.Data.Configuration
Imports Microsoft.Practices.EnterpriseLibrary.Data.Sql
Imports Microsoft.Practices.EnterpriseLibrary.Data.Oracle
Imports System.Data
Imports System.Data.SqlClient
Imports System.Xml
Imports System.Xml.XPath
Imports System.Xml.Xsl
Imports System.IO
Imports System.Runtime.CompilerServices
Imports Microsoft.Office.Interop.Access

Namespace COPT
    <Microsoft.VisualBasic.ComClass()>
    Public Class DAAB
        Private _db As Database
        Private _dbConn As DbConnection 'RKP/03-23-12/v3.0.158
        Private _databaseName As String
        Private _provider As System.Data.Common.DbProviderFactory
        Private _conn As System.Data.OleDb.OleDbConnection
        Private _lastErrorNo As Integer = 0
        Private _lastErrorDesc As String = ""

        'RKP/07-26-11/v2.5.147
        'Private _connSQLExpress As System.Data.SqlClient.SqlConnection = Nothing
        Private _connSQLExpress As SqlConnection = Nothing

        Private _switches() As String

        Private _dal As clsDAL
        Private _dalOleDb As System.Data.OleDb.OleDbConnection
        Private _dalSql As System.Data.SqlClient.SqlConnection
        Private _dalSqlCe As System.Data.SqlServerCe.SqlCeConnection
        Private _dalOdbc As System.Data.Odbc.OdbcConnection
        Private _dalAdoDb As ADODB.Connection
        Private _dalDao As dao.Connection

        Public Enum e_DB
            e_db_NONE = 0
            e_db_ACCESS = 100
            e_db_ACCESS_MDB = 110
            e_db_ACCESS_ACCDB = 120
            e_db_SQLSERVER = 200
            e_db_SQLSERVER_COMPACT = 210
            e_db_SQLSERVER_LOCALDB = 220
            e_db_SQLSERVER_EXPRESS = 230
            e_db_SQLSERVER_ENTERPRISE = 240
            e_db_ORACLE = 300
            e_db_ORACLE_EXPRESS = 310
            e_db_MYSQL
            e_db_DB2_EXPRESS
            e_db_DB2
            e_db_TEXTFILE = 400
            e_db_EXCEL = 500
            e_db_OTHER = 10000
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

        'RKP/11-29-11/v3.0.153
        Public Enum e_TransferDataType
            e_MSAccessToSQLExpress
            e_SQLExpressToMSAccess
        End Enum

        Private Structure STRUCT_TABLEDEF
            Dim fieldName As String
            Dim fieldType As String
            Dim fieldSize As Integer
            Dim fieldIsIdentity As Boolean
            Dim fieldIsNullable As Boolean
            Dim fieldTypeNew As String
        End Structure
        'RKP/11-23-11/v3.0.153
        Private Structure STRUCT_IDX
            Dim idxName As String
            Dim idxCols() As String
            Dim isUnique As Boolean
        End Structure
        'RKP/11-23-11/v3.0.153
        Private Structure STRUCT_KEYVALUE
            Dim key As String
            Dim value As String
            Dim valueInt As Integer 'RKP/04-04-12/v3.0.164
            Dim flag As Integer 'RKP/04-04-12/v3.0.164
        End Structure

        Public Sub DAAB()
        End Sub

        Public Sub New()
            ConnectToDB()
        End Sub

        Public Sub New(ByVal databaseName As String)
            _databaseName = databaseName
            ConnectToDB(databaseName)
        End Sub

        Public Sub New(ByVal databaseName As String, ByVal switches() As String)
            _switches = switches
            _databaseName = databaseName

            If GenUtils.IsSwitchAvailable(switches, "/UseMinSysRes") Then
                If GenUtils.IsSwitchAvailable(switches, "/UseSQLServerSyntax") Then
                    ConnectToDB(databaseName, switches)
                Else
                    ConnectToDB(databaseName)
                End If
            Else
                ConnectToDB(databaseName)
            End If

        End Sub

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="providerName"></param>
        ''' <param name="connectionString"></param>
        ''' <remarks>
        ''' RKP/06-16-10/v2.3.133
        ''' To allow ad-hoc connections to linked databases (tsysCOL, tsysROW, tsysMTX)
        ''' </remarks>
        Public Sub New(ByVal providerName As String, ByVal connectionString As String)
            '_databaseName = databaseName
            'ConnectToDB(databaseName)
        End Sub

        Public Sub ConnectToDB()
            '_db = DatabaseFactory.CreateDatabase()
            ConnectToDB("")
        End Sub

        Public Sub ConnectToDB(ByVal databaseName As String)
            Dim connStr As String
            Dim providerName As String
            Dim dbFactory As System.Data.Common.DbProviderFactory
            Dim dbGeneric As Microsoft.Practices.EnterpriseLibrary.Data.GenericDatabase   'GenericDatabase

            DAL = New clsDAL

            If String.IsNullOrEmpty(databaseName) Then
                databaseName = GenUtils.GetSysConfigValue("defaultProject")
            End If

            _databaseName = databaseName
            SetLastError(Nothing)

            connStr = GenUtils.GetSysConfig.ConnectionStrings.ConnectionStrings(databaseName).ConnectionString
            providerName = GenUtils.GetSysConfig.ConnectionStrings.ConnectionStrings(databaseName).ProviderName
            If providerName.Trim().ToUpper().Contains("System.Data.SqlClient".ToUpper()) Then
                dbFactory = SqlClientFactory.Instance

                DAL.UseADO = False
                DAL.Connect(clsDAL.e_DB.e_db_SQLSERVER, clsDAL.e_ConnType.e_connType_Z, Nothing, Nothing, _dalSql, Nothing, Nothing, True, connStr)
            ElseIf providerName.Trim().ToUpper().Contains("System.Data.OleDb".ToUpper()) Then
                dbFactory = OleDbFactory.Instance

                DAL.UseADO = False
                DAL.Connect(clsDAL.e_DB.e_db_ACCESS, clsDAL.e_ConnType.e_connType_Z, Nothing, _dalOleDb, Nothing, Nothing, Nothing, True, connStr)
            ElseIf providerName.Trim().ToUpper().Contains("System.Data.Odbc".ToUpper()) Then
                dbFactory = OdbcFactory.Instance

                DAL.UseADO = False
                DAL.Connect(clsDAL.e_DB.e_db_ACCESS, clsDAL.e_ConnType.e_connType_Z, Nothing, _dalOleDb, Nothing, Nothing, Nothing, True, connStr)
            Else
                dbFactory = OleDbFactory.Instance
            End If

            Try
                dbGeneric = New Microsoft.Practices.EnterpriseLibrary.Data.GenericDatabase(connStr, dbFactory)
                _db = dbGeneric

                '_db = DatabaseFactory.CreateDatabase(databaseName)

                _dbConn = _db.CreateConnection()

                'If _db.ConnectionString.Contains("SQLEXPRESS") Then
                'If _db.DbProviderFactory.ToString.Contains("System.Data.SqlClient") Then
                If IsSQLExpress() Then
                    '_connSQLExpress = New System.Data.SqlClient.SqlConnection(_db.ConnectionString)
                    _connSQLExpress = New SqlConnection(_db.ConnectionString)
                    '_connSQLExpress.ConnectionTimeout = 0 'Read only
                    'Debug.Print(_connSQLExpress.ConnectionTimeout.ToString())
                    _connSQLExpress.Open()
                End If
            Catch ex As Exception
                'MsgBox(ex.Message)
                SetLastError(ex)
                GenUtils.Message(GenUtils.MsgType.Critical, "EntLib - DAAB-ConnectToDB", ex.Message)
            End Try



        End Sub

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="databaseName"></param>
        ''' <param name="switches"></param>
        ''' <remarks>
        ''' RKP/07-31-11/v2.5.147
        ''' Created to accomodate connecting to linked SQL Server databases.
        ''' </remarks>
        Public Sub ConnectToDB(ByVal databaseName As String, ByVal switches() As String)

            SetLastError(Nothing)
            If (Not String.IsNullOrEmpty(databaseName)) AndAlso (switches IsNot Nothing) Then

                _databaseName = databaseName
                _switches = switches

                Try
                    '_db = DatabaseFactory.CreateDatabase(databaseName)
                    ConnectToDB(databaseName)

                    'If _db.ConnectionString.Contains("SQLEXPRESS") Then
                    'If _db.DbProviderFactory.ToString.Contains("System.Data.SqlClient") Then
                    If IsSQLExpress() Then
                        'do nothing - ConnectToDB(databaseName) already does the work.

                        ''_connSQLExpress = New System.Data.SqlClient.SqlConnection(_db.ConnectionString)
                        '_connSQLExpress = New SqlConnection(_db.ConnectionString)
                        ''Debug.Print(_connSQLExpress.ConnectionTimeout.ToString())
                        '_connSQLExpress.Open()
                    ElseIf EntLib.COPT.GenUtils.IsSwitchAvailable(switches, "/UseSQLServerSyntax") Then
                        'If EntLib.COPT.GenUtils.IsSwitchAvailable(switches, "/UseMinSysRes") Then
                        'If EntLib.COPT.GenUtils.IsSwitchAvailable(switches, "/UseSQLServerSyntax") Then
                        _connSQLExpress = New SqlConnection(Me.GetDataTable("SELECT LINKED_SQL_SERVER_CONN_STR FROM qsysMiscParams").Rows(0).Item("LINKED_SQL_SERVER_CONN_STR").ToString())
                        'Debug.Print(_connSQLExpress.ConnectionTimeout.ToString())
                        _connSQLExpress.Open()
                        'End If
                        'End If
                    End If
                Catch ex As Exception
                    'MsgBox(ex.Message)
                    SetLastError(ex)
                    GenUtils.Message(GenUtils.MsgType.Critical, "EntLib - DAAB-ConnectToDB: " & databaseName, ex.Message)
                End Try
            Else
                '_db = DatabaseFactory.CreateDatabase()
                ConnectToDB("")
            End If

        End Sub

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="providerName"></param>
        ''' <param name="connectionString"></param>
        ''' <remarks>
        ''' RKP/06-16-10/v2.3.133
        ''' Created to allow ad-hoc connections to linked databases (tsysCOL, tsysROW, tsysMTX).
        ''' </remarks>
        Public Sub ConnectToDB(ByVal providerName As String, ByVal connectionString As String)
            'The Microsoft.Practices.EnterpriseLibrary.Data.Database class leverages 
            'the provider factory model from ADO.NET. 
            'A database instance holds a reference to a concrete 
            'Microsoft.Practices.EnterpriseLibrary.Data.Database.DbProviderFactory object 
            'to which it forwards the creation of ADO.NET objects.
            '_db = DatabaseFactory.CreateDatabase()
            '_db = System.Data.Common.DbProviderFactories.GetFactory(GenUtils.GetAppSettings("linkedColDBProviderName"))

            'Dim provider As System.Data.Common.DbProviderFactory = System.Data.Common.DbProviderFactories.GetFactory(GenUtils.GetAppSettings("linkedColDBProviderName"))
            _provider = System.Data.Common.DbProviderFactories.GetFactory(providerName)
            _conn = CType(_provider.CreateConnection(), OleDbConnection)
            '_conn.ConnectionString = GenUtils.GetAppSettings("linkedColDBConnectionString")
            _conn.ConnectionString = connectionString
            'Debug.Print(_conn.ConnectionTimeout.ToString())
            _conn.Open()
        End Sub

        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub

        'Public Function GetDatabase() As Database
        '    ConnectToDB()
        '    Return _db
        'End Function

        ' Retrieves a list of customers from the database.
        ' Returns: List of customers in a string.
        ' Remarks: Demonstrates retrieving multiple rows of data using
        ' a DataReader
        'Public Function GetData() As String
        '    ' DataReader that will hold the returned results		
        '    ' Create the Database object, using the default database service. The
        '    ' default database service is determined through configuration.

        '    'Dim db As Database = DatabaseFactory.CreateDatabase("C-OPTSysDbDefault")
        '    Dim db As Database = DatabaseFactory.CreateDatabase()

        '    Dim sqlCommand As String = "SELECT * FROM dbo.tblMills"

        '    'db.

        '    Dim dbCommand As DbCommand = db.GetSqlStringCommand(sqlCommand)

        '    Dim readerData As StringBuilder = New StringBuilder

        '    ' The ExecuteReader call will request the connection to be closed upon
        '    ' the closing of the DataReader. The DataReader will be closed 
        '    ' automatically when it is disposed.
        '    Using dataReader As IDataReader = db.ExecuteReader(dbCommand)

        '        ' Iterate through DataReader and put results to the text box.
        '        ' DataReaders cannot be bound to Windows Form controls (e.g. the
        '        ' resultsDataGrid), but may be bound to Web Form controls.
        '        While (dataReader.Read())
        '            ' Get the value of the 'Name' column in the DataReader
        '            readerData.Append(dataReader("MILL_ID").ToString() & " - " & dataReader("MILL_DESC").ToString())
        '            readerData.Append(Environment.NewLine)
        '        End While
        '    End Using
        '    Return readerData.ToString()

        'End Function

        'Public Function GetDataStoredProc() As DataSet

        '    Dim Category As Integer

        '    ' Create the Database object, using the default database service. The
        '    ' default database service is determined through configuration.
        '    Dim db As Database = DatabaseFactory.CreateDatabase()

        '    Dim sqlCommand As String = "dbo.asp_GetProductsByCategory"
        '    Dim dbCommand As DbCommand = db.GetStoredProcCommand(sqlCommand)

        '    Category = 1
        '    ' Retrieve products from the specified category.
        '    db.AddInParameter(dbCommand, "CategoryID", DbType.Int32, Category)

        '    ' DataSet that will hold the returned results		
        '    Dim productsDataSet As DataSet = Nothing

        '    productsDataSet = db.ExecuteDataSet(dbCommand)

        '    ' Note: connection was closed by ExecuteDataSet method call 

        '    Return productsDataSet

        'End Function

        ' Retreives all products in the specified category.
        ' Category: The category containing the products.
        ' Returns: DataSet containing the products.
        ' Remarks: Demonstrates retrieving multiple rows using a DataSet.
        'Public Function GetProductsInCategory(ByRef Category As Integer) As DataSet

        '    ' Create the Database object, using the default database service. The
        '    ' default database service is determined through configuration.
        '    Dim db As Microsoft.Practices.EnterpriseLibrary.Data.Database = DatabaseFactory.CreateDatabase()

        '    Dim sqlCommand As String = "GetProductsByCategory"
        '    Dim dbCommand As Common.DbCommand = db.GetStoredProcCommand(sqlCommand)

        '    ' Retrieve products from the specified category.
        '    db.AddInParameter(dbCommand, "CategoryID", DbType.Int32, Category)

        '    ' DataSet that will hold the returned results		
        '    Dim productsDataSet As DataSet = Nothing

        '    productsDataSet = db.ExecuteDataSet(dbCommand)

        '    ' Note: connection was closed by ExecuteDataSet method call 

        '    Return productsDataSet
        'End Function

        'Public Function GetTestData() As DataSet
        '    'C-OPTSystem
        '    'Return GetDataStoredProc()

        '    Return GetDataSet("SELECT * FROM tsysMapPoint")

        'End Function

        Public Function GetDataSet(ByVal sql As String) As DataSet

            ' Create the Database object, using the default database service. The
            ' default database service is determined through configuration.
            'Dim db As Database = DatabaseFactory.CreateDatabase()
            '_db = DatabaseFactory.CreateDatabase()

            Dim returnDataSet As DataSet = Nothing

            Dim adapterSql As System.Data.SqlClient.SqlDataAdapter
            'Dim cmdSQL As System.Data.SqlClient.SqlCommand
            Dim dt As DataTable
            'Dim ret As Integer

            SetLastError(Nothing)

            If _db Is Nothing Then
                ConnectToDB()
            End If


            'Dim sqlCommand As String = "SELECT * FROM " & nameTableOrQueryOrView
            'Dim dbCommand As DbCommand = db.GetStoredProcCommand(sqlCommand)
            Try
                If IsSQLExpress() Then '_db.ConnectionString.Contains("SQLEXPRESS") Then
                    'returnDataSet = Nothing
                    If sql.Trim().ToUpper().Substring(0, 6).Equals("SELECT") Then
                        Try
                            adapterSql = New System.Data.SqlClient.SqlDataAdapter(sql, _connSQLExpress)
                            dt = New DataTable
                            adapterSql.Fill(dt)
                            If Not dt Is Nothing Then
                                returnDataSet = New DataSet()
                                returnDataSet.Tables.Add(dt)
                                'ret = dt.Rows.Count
                            Else
                                'ret = 0
                            End If
                        Catch ex As Exception
                            'GenUtils.Log("Error executing SQL:" & vbNewLine & sql)
                            'GenUtils.Log(ex.Message)
                            SetLastError(ex)
                            returnDataSet = Nothing
                            GenUtils.Message(GenUtils.MsgType.Critical, "EntLib - DAAB-GetDataSet", ex.Message)
                        End Try
                        'ElseIf sql.Trim().ToUpper().Substring(0, 4).Equals("EXEC") Then

                        '    Try
                        '        cmdSQL = New System.Data.SqlClient.SqlCommand("SQLServer", ConnSQLServer)
                        '        cmdSQL.CommandType = CommandType.StoredProcedure
                        '        cmdSQL.CommandText = sql.Substring(5).Split(CChar(" "))(0)
                        '        If pArray.Length > 0 Then
                        '            For ctr As Integer = 0 To pArray.Length Step 3
                        '                cmdSQL.Parameters.AddWithValue(pArray(ctr).ToString(), pArray(ctr + 1).ToString())
                        '                cmdSQL.Parameters(ctr).SqlDbType = CType(CInt(pArray(ctr + 2)), SqlDbType)
                        '            Next
                        '        End If
                        '        cmdSQL.CommandTimeout = 0
                        '        ret = cmdSQL.ExecuteNonQuery()
                        '    Catch ex As Exception
                        '        GenUtils.Log("Error executing SQL:" & vbNewLine & sql)
                        '        GenUtils.Log(ex.Message)
                        '        GenUtils.Log(ex.Source)
                        '    End Try
                        'ElseIf sql.Trim().ToUpper().Substring(0, 8).Equals("BULKCOPY") Then
                        '    bulkCopy = New SqlBulkCopy(ConnSQLServer)
                        '    bulkCopy.DestinationTableName = pArray(0).ToString()
                        '    Try
                        '        bulkCopy.WriteToServer(dt)
                        '    Catch ex As Exception
                        '        GenUtils.Log("Error executing BULKCOPY: " & pArray(0).ToString())
                        '        GenUtils.Log("Error msg: " & ex.Message)
                        '        GenUtils.Log("Error source: " & ex.Source)
                        '    End Try
                        'Else
                        '    cmdSQL = New System.Data.SqlClient.SqlCommand("SQLServer", ConnSQLServer)
                        '    cmdSQL.CommandType = System.Data.CommandType.Text
                        '    cmdSQL.CommandText = sql
                        '    Try
                        '        ret = cmdSQL.ExecuteNonQuery()
                        '    Catch ex As Exception
                        '        GenUtils.Log("Error executing SQL:" & vbNewLine & sql)
                        '        GenUtils.Log(ex.Message)
                        '    End Try
                    End If
                Else
                    returnDataSet = _db.ExecuteDataSet(System.Data.CommandType.Text, sql)
                End If

            Catch ex As Exception
                SetLastError(ex)
                returnDataSet = Nothing
                GenUtils.Message(GenUtils.MsgType.Critical, "EntLib - DAAB-GetDataSet", ex.Message)
            End Try

            ' Note: connection was closed by ExecuteDataSet method call 

            Return returnDataSet

        End Function

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="sql"></param>
        ''' <param name="suppressMsg"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' RKP/08-02-11/v2.5.147
        ''' Created to allow calls for Blueprint checks.
        ''' </remarks>
        Public Function GetDataSet(ByVal sql As String, ByVal suppressMsg As Boolean) As DataSet

            ' Create the Database object, using the default database service. The
            ' default database service is determined through configuration.
            'Dim db As Database = DatabaseFactory.CreateDatabase()
            '_db = DatabaseFactory.CreateDatabase()

            Dim returnDataSet As DataSet = Nothing

            Dim adapterSql As System.Data.SqlClient.SqlDataAdapter
            'Dim cmdSQL As System.Data.SqlClient.SqlCommand
            Dim dt As DataTable
            'Dim ret As Integer

            SetLastError(Nothing)

            If _db Is Nothing Then
                ConnectToDB()
            End If


            'Dim sqlCommand As String = "SELECT * FROM " & nameTableOrQueryOrView
            'Dim dbCommand As DbCommand = db.GetStoredProcCommand(sqlCommand)
            Try
                If IsSQLExpress() Then '_db.ConnectionString.Contains("SQLEXPRESS") Then
                    'returnDataSet = Nothing
                    If sql.Trim().ToUpper().Substring(0, 6).Equals("SELECT") Then
                        Try
                            adapterSql = New System.Data.SqlClient.SqlDataAdapter(sql, _connSQLExpress)
                            dt = New DataTable
                            adapterSql.Fill(dt)
                            If Not dt Is Nothing Then
                                returnDataSet = New DataSet()
                                returnDataSet.Tables.Add(dt)
                                'ret = dt.Rows.Count
                            Else
                                'ret = 0
                            End If
                        Catch ex As Exception
                            'GenUtils.Log("Error executing SQL:" & vbNewLine & sql)
                            'GenUtils.Log(ex.Message)
                            SetLastError(ex)
                            returnDataSet = Nothing
                            If suppressMsg Then
                                EntLib.COPT.Log.Log(GenUtils.GetSysConfigValue("lastWorkDir").ToString(), "C-OPT Message - Critical - DAAB.GetDataSet", ex.Message)
                            Else
                                GenUtils.Message(GenUtils.MsgType.Critical, "EntLib - DAAB-GetDataSet", ex.Message)
                            End If

                        End Try
                        'ElseIf sql.Trim().ToUpper().Substring(0, 4).Equals("EXEC") Then

                        '    Try
                        '        cmdSQL = New System.Data.SqlClient.SqlCommand("SQLServer", ConnSQLServer)
                        '        cmdSQL.CommandType = CommandType.StoredProcedure
                        '        cmdSQL.CommandText = sql.Substring(5).Split(CChar(" "))(0)
                        '        If pArray.Length > 0 Then
                        '            For ctr As Integer = 0 To pArray.Length Step 3
                        '                cmdSQL.Parameters.AddWithValue(pArray(ctr).ToString(), pArray(ctr + 1).ToString())
                        '                cmdSQL.Parameters(ctr).SqlDbType = CType(CInt(pArray(ctr + 2)), SqlDbType)
                        '            Next
                        '        End If
                        '        cmdSQL.CommandTimeout = 0
                        '        ret = cmdSQL.ExecuteNonQuery()
                        '    Catch ex As Exception
                        '        GenUtils.Log("Error executing SQL:" & vbNewLine & sql)
                        '        GenUtils.Log(ex.Message)
                        '        GenUtils.Log(ex.Source)
                        '    End Try
                        'ElseIf sql.Trim().ToUpper().Substring(0, 8).Equals("BULKCOPY") Then
                        '    bulkCopy = New SqlBulkCopy(ConnSQLServer)
                        '    bulkCopy.DestinationTableName = pArray(0).ToString()
                        '    Try
                        '        bulkCopy.WriteToServer(dt)
                        '    Catch ex As Exception
                        '        GenUtils.Log("Error executing BULKCOPY: " & pArray(0).ToString())
                        '        GenUtils.Log("Error msg: " & ex.Message)
                        '        GenUtils.Log("Error source: " & ex.Source)
                        '    End Try
                        'Else
                        '    cmdSQL = New System.Data.SqlClient.SqlCommand("SQLServer", ConnSQLServer)
                        '    cmdSQL.CommandType = System.Data.CommandType.Text
                        '    cmdSQL.CommandText = sql
                        '    Try
                        '        ret = cmdSQL.ExecuteNonQuery()
                        '    Catch ex As Exception
                        '        GenUtils.Log("Error executing SQL:" & vbNewLine & sql)
                        '        GenUtils.Log(ex.Message)
                        '    End Try
                    End If
                Else
                    returnDataSet = _db.ExecuteDataSet(System.Data.CommandType.Text, sql)
                End If

            Catch ex As Exception
                SetLastError(ex)
                returnDataSet = Nothing
                If suppressMsg Then
                    EntLib.COPT.Log.Log(GenUtils.GetSysConfigValue("lastWorkDir").ToString(), "C-OPT Message - Critical - DAAB.GetDataSet", ex.Message)
                Else
                    GenUtils.Message(GenUtils.MsgType.Critical, "EntLib - DAAB-GetDataSet", ex.Message)
                End If
                'GenUtils.Message(GenUtils.MsgType.Critical, "EntLib - DAAB-GetDataSet", ex.Message)
            End Try

            ' Note: connection was closed by ExecuteDataSet method call 

            Return returnDataSet

        End Function

        Public Function GetDataTable(ByVal sql As String, ByVal returnCopy As Boolean) As DataTable

            ' Create the Database object, using the default database service. The
            ' default database service is determined through configuration.
            'Dim db As Database = DatabaseFactory.CreateDatabase()
            '_db = DatabaseFactory.CreateDatabase()
            If _db Is Nothing Then
                ConnectToDB()
            End If

            'Dim sqlCommand As String = "SELECT * FROM " & nameTableOrQueryOrView
            'Dim dbCommand As DbCommand = db.GetStoredProcCommand(sqlCommand)
            Dim returnDataTable As DataTable = _db.ExecuteDataSet(System.Data.CommandType.Text, sql).Tables(0)


            ' Note: connection was closed by ExecuteDataSet method call 
            If returnCopy Then
                Return returnDataTable.Copy
            Else
                Return returnDataTable
            End If

        End Function
        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="sql"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' RKP/02-18-10/v2.3.131
        ''' Removed requirement to return copy of DataTable for performance efficiencies, especially with PC31.
        ''' </remarks>
        Public Function GetDataTable(ByVal sql As String) As DataTable

            ' Create the Database object, using the default database service. The
            ' default database service is determined through configuration.
            'Dim db As Database = DatabaseFactory.CreateDatabase()
            '_db = DatabaseFactory.CreateDatabase()

            If _db Is Nothing Then
                ConnectToDB()
            End If

            'Debug.Print(_db.CreateConnection.ConnectionTimeout.ToString())

            'Dim sqlCommand As String = "SELECT * FROM " & nameTableOrQueryOrView
            'Dim dbCommand As DbCommand = db.GetStoredProcCommand(sqlCommand)


            '*&*
            ''Debug.Print "here is the sql for qryBadSolverResults ".ToString
            Dim returnDataTable As DataTable = _db.ExecuteDataSet(System.Data.CommandType.Text, sql).Tables(0)


            ' Note: connection was closed by ExecuteDataSet method call 

            Return returnDataTable '.Copy

        End Function

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="sql"></param>
        ''' <param name="linkedDB"></param>
        ''' <param name="switches"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' RKP/08-01-11/v2.5.147
        ''' Created to accomodate SQL Express processing
        ''' </remarks>
        Public Function GetDataTable(ByVal sql As String, ByVal linkedDB As Boolean, ByVal switches() As String) As DataTable

            ' Create the Database object, using the default database service. The
            ' default database service is determined through configuration.
            'Dim db As Database = DatabaseFactory.CreateDatabase()
            '_db = DatabaseFactory.CreateDatabase()

            Dim returnDataTable As DataTable
            Dim sqlDA As SqlDataAdapter
            Dim proceed As Boolean = False

            _switches = switches

            If _db Is Nothing Then
                ConnectToDB(_databaseName, _switches)
            End If

            'Dim sqlCommand As String = "SELECT * FROM " & nameTableOrQueryOrView
            'Dim dbCommand As DbCommand = db.GetStoredProcCommand(sqlCommand)

            If IsSQLExpress() Then
                If _connSQLExpress Is Nothing Then
                    ConnectToDB(_databaseName, _switches)
                End If
                sqlDA = New SqlDataAdapter(sql, _connSQLExpress)
                returnDataTable = New DataTable
                sqlDA.Fill(returnDataTable)
            Else
                If linkedDB Then
                    If GenUtils.IsSwitchAvailable(_switches, "/UseSQLServerSyntax") Then
                        If _connSQLExpress Is Nothing Then
                            ConnectToDB(_databaseName, _switches)
                        End If
                        sqlDA = New SqlDataAdapter(sql, _connSQLExpress)
                        returnDataTable = New DataTable
                        sqlDA.Fill(returnDataTable)
                    Else
                        returnDataTable = New DataTable
                        proceed = True
                    End If
                Else
                    If GenUtils.IsSwitchAvailable(_switches, "/UseMinSysRes") Then
                        If GenUtils.IsSwitchAvailable(_switches, "/UseSQLServerSyntax") Then
                            If _connSQLExpress Is Nothing Then
                                ConnectToDB(_databaseName, _switches)
                            End If
                            sqlDA = New SqlDataAdapter(sql, _connSQLExpress)
                            returnDataTable = New DataTable
                            sqlDA.Fill(returnDataTable)
                        Else
                            returnDataTable = New DataTable
                            proceed = True
                        End If
                    Else
                        returnDataTable = New DataTable
                        proceed = True
                    End If
                End If
            End If

            If proceed Then
                returnDataTable = New DataTable
                returnDataTable = _db.ExecuteDataSet(System.Data.CommandType.Text, sql).Tables(0)
            End If

            ' Note: connection was closed by ExecuteDataSet method call 

            Return returnDataTable '.Copy

        End Function

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="sql"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' RKP/06-13-10/v2.3.133
        ''' Created this function in order to provide a fast efficient in-memory DataTable.
        ''' </remarks>
        Public Function GetDataTableReader(ByVal sql As String) As DataTableReader

            ' Create the Database object, using the default database service. The
            ' default database service is determined through configuration.
            'Dim db As Database = DatabaseFactory.CreateDatabase()
            '_db = DatabaseFactory.CreateDatabase()
            If _db Is Nothing Then
                ConnectToDB()
            End If

            'Dim sqlCommand As String = "SELECT * FROM " & nameTableOrQueryOrView
            'Dim dbCommand As DbCommand = db.GetStoredProcCommand(sqlCommand)
            Dim returnDataTable As DataTable = _db.ExecuteDataSet(System.Data.CommandType.Text, sql).Tables(0)


            ' Note: connection was closed by ExecuteDataSet method call 
            Return returnDataTable.CreateDataReader   '.Copy

        End Function

        Public Function GetDbDataAdapter() As System.Data.Common.DbDataAdapter
            Return _db.DbProviderFactory.CreateDataAdapter()
        End Function

        Public Function GetCommandBuilder() As System.Data.Common.DbCommandBuilder
            Return _db.DbProviderFactory.CreateCommandBuilder()
        End Function

        Public Function GetDbCommand() As System.Data.Common.DbCommand
            Return _db.DbProviderFactory.CreateCommand()
        End Function

        Public Function GetDbCommand(ByVal sql As String) As System.Data.Common.DbCommand
            Return _db.GetSqlStringCommand(sql)
        End Function

        Public Function GetDbConnection() As System.Data.Common.DbConnection
            'Return _db.DbProviderFactory.CreateConnection()
            Return _db.CreateConnection()

        End Function

        Public Function ExecuteNonQuery(ByVal sql As String) As Integer
            Dim cmdSQL As System.Data.SqlClient.SqlCommand
            Dim ret As Integer = 0
            Dim rowCount As Integer = -1

            SetLastError(Nothing)

            '_db = DatabaseFactory.CreateDatabase()
            If _db Is Nothing Then
                If (Not String.IsNullOrEmpty(_databaseName)) AndAlso (_switches IsNot Nothing) Then
                    ConnectToDB(_databaseName, _switches)
                Else
                    ConnectToDB()
                End If
            End If
            If IsSQLExpress() Then
                cmdSQL = New System.Data.SqlClient.SqlCommand("SQLServer", _connSQLExpress)
                cmdSQL.CommandTimeout = 0

                'RKP/03-28-12/v3.0.161
                'If there is a space, then "sql" is SQL statement.
                'If there is/are no space(s), then "sql" is the name of a Stored Procedure.
                If sql.Contains(" ") Then
                    cmdSQL.CommandType = System.Data.CommandType.Text
                Else
                    cmdSQL.CommandType = System.Data.CommandType.StoredProcedure
                End If
                cmdSQL.CommandText = sql

                Try
                    rowCount = cmdSQL.ExecuteNonQuery()
                    'Return ret
                Catch ex As Exception
                    'GenUtils.Log("Error executing SQL:" & vbNewLine & sql)
                    'GenUtils.Log(ex.Message)
                    rowCount = -1
                    SetLastError(ex)
                    GenUtils.Message(GenUtils.MsgType.Critical, "EntLib - DAAB-ExecuteNonQuery", ex.Message & vbNewLine & sql)
                End Try
            Else
                Try
                    rowCount = _db.ExecuteNonQuery(System.Data.CommandType.Text, sql)
                Catch ex As Exception
                    rowCount = -1
                    SetLastError(ex)
                    GenUtils.Message(GenUtils.MsgType.Critical, "EntLib - DAAB-ExecuteNonQuery", ex.Message & vbNewLine & sql)
                End Try
            End If

            Return rowCount
        End Function

        ''' <summary>
        ''' RKP/12-10-11/v3.0.155
        ''' </summary>
        ''' <param name="sql"></param>
        ''' <param name="supressMsg"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' Same as "ExecuteNonQuery(ByVal sql As String)
        ''' except with the option to suppress all messages.
        ''' Useful for scenarios where C-OPT doesn't care if the sql is executed successfully or unsuccessfully.
        ''' eg; ShrinkDatabase()
        ''' </remarks>
        Public Function ExecuteNonQuery(ByVal sql As String, ByVal supressMsg As Boolean) As Integer
            Dim cmdSQL As System.Data.SqlClient.SqlCommand
            Dim ret As Integer = 0
            Dim rowCount As Integer = -1

            SetLastError(Nothing)

            '_db = DatabaseFactory.CreateDatabase()
            If _db Is Nothing Then
                If (Not String.IsNullOrEmpty(_databaseName)) AndAlso (_switches IsNot Nothing) Then
                    ConnectToDB(_databaseName, _switches)
                Else
                    ConnectToDB()
                End If
            End If
            If IsSQLExpress() Then
                cmdSQL = New System.Data.SqlClient.SqlCommand("SQLServer", _connSQLExpress)
                cmdSQL.CommandTimeout = 0
                cmdSQL.CommandType = System.Data.CommandType.Text
                cmdSQL.CommandText = sql
                Try
                    rowCount = cmdSQL.ExecuteNonQuery()
                    'Return ret
                Catch ex As Exception
                    'GenUtils.Log("Error executing SQL:" & vbNewLine & sql)
                    'GenUtils.Log(ex.Message)
                    rowCount = -1
                    SetLastError(ex)
                    If supressMsg Then
                    Else
                        GenUtils.Message(GenUtils.MsgType.Critical, "EntLib - DAAB-ExecuteNonQuery", ex.Message & vbNewLine & sql)
                    End If
                End Try
            Else
                Try
                    rowCount = _db.ExecuteNonQuery(System.Data.CommandType.Text, sql)
                Catch ex As Exception
                    rowCount = -1
                    SetLastError(ex)
                    If supressMsg Then
                    Else
                        GenUtils.Message(GenUtils.MsgType.Critical, "EntLib - DAAB-ExecuteNonQuery", ex.Message & vbNewLine & sql)
                    End If
                End Try

                'Return rowCount
            End If

            Return rowCount
        End Function

        Public Function ExecuteNonQuery(ByVal sql As String, ByVal switches() As String, ByVal linkedDB As Boolean) As Integer
            Dim cmdSQL As System.Data.SqlClient.SqlCommand
            Dim ret As Integer = 0
            Dim rowCount As Integer = -1

            _switches = switches

            SetLastError(Nothing)

            '_db = DatabaseFactory.CreateDatabase()
            If _db Is Nothing Then
                ConnectToDB(_databaseName, switches)
            End If
            If _db Is Nothing Then
                ConnectToDB()
            End If

            If linkedDB Then
                If GenUtils.IsSwitchAvailable(_switches, "/UseMinSysRes") Then
                    If GenUtils.IsSwitchAvailable(_switches, "/UseSQLServerSyntax") Then

                        If _connSQLExpress Is Nothing Then
                            ConnectToDB(_databaseName, _switches)
                        End If

                        cmdSQL = New System.Data.SqlClient.SqlCommand("SQLServer", _connSQLExpress)
                        cmdSQL.CommandTimeout = 0
                        cmdSQL.CommandType = System.Data.CommandType.Text
                        cmdSQL.CommandText = sql
                        Try
                            rowCount = cmdSQL.ExecuteNonQuery()
                            'Return ret
                        Catch ex As Exception
                            'GenUtils.Log("Error executing SQL:" & vbNewLine & sql)
                            'GenUtils.Log(ex.Message)
                            rowCount = -1
                            SetLastError(ex)
                            GenUtils.Message(GenUtils.MsgType.Critical, "EntLib - DAAB-ExecuteNonQuery", ex.Message)
                        End Try
                    Else
                        Try
                            rowCount = _db.ExecuteNonQuery(System.Data.CommandType.Text, sql)
                        Catch ex As Exception
                            rowCount = -1
                            SetLastError(ex)
                            GenUtils.Message(GenUtils.MsgType.Critical, "EntLib - DAAB-ExecuteNonQuery", ex.Message)
                        End Try

                        'Return errCount
                    End If
                Else
                    Try
                        rowCount = _db.ExecuteNonQuery(System.Data.CommandType.Text, sql)
                    Catch ex As Exception
                        rowCount = -1
                        SetLastError(ex)
                        GenUtils.Message(GenUtils.MsgType.Critical, "EntLib - DAAB-ExecuteNonQuery", ex.Message)
                    End Try

                    'Return errCount
                End If
            Else
                If IsSQLExpress() Then
                    cmdSQL = New System.Data.SqlClient.SqlCommand("SQLServer", _connSQLExpress)
                    cmdSQL.CommandTimeout = 0
                    cmdSQL.CommandType = System.Data.CommandType.Text
                    cmdSQL.CommandText = sql
                    Try
                        rowCount = cmdSQL.ExecuteNonQuery()
                        'Return ret
                    Catch ex As Exception
                        'GenUtils.Log("Error executing SQL:" & vbNewLine & sql)
                        'GenUtils.Log(ex.Message)
                        rowCount = -1
                        SetLastError(ex)
                        GenUtils.Message(GenUtils.MsgType.Critical, "EntLib - DAAB-ExecuteNonQuery", ex.Message)
                    End Try
                Else
                    Try
                        rowCount = _db.ExecuteNonQuery(System.Data.CommandType.Text, sql)
                    Catch ex As Exception
                        rowCount = -1
                        SetLastError(ex)
                        GenUtils.Message(GenUtils.MsgType.Critical, "EntLib - DAAB-ExecuteNonQuery", ex.Message)
                    End Try

                    'Return errCount
                End If
            End If

            Return rowCount

        End Function

        Public Function GetTables() As DataTable
            Dim startTempTime As Integer = My.Computer.Clock.TickCount
            Dim userTables As DataTable = Nothing
            ' We only want user tables, not system tables
            Dim restrictions() As String = New String(3) {}
            restrictions(3) = "Table"

            SetLastError(Nothing)

            If _db Is Nothing Then
                ConnectToDB()
            End If

            Try
                Dim conn As DbConnection = _db.CreateConnection()
                conn.Open()
                userTables = conn.GetSchema("Tables", restrictions)

                'EntLib.Log.Log("C:\OPTMODELS\", "EntLib - DAAB.GetTables - GetSchema: ", Space(7) & EntLib.GenUtils.FormatTime(startTempTime, My.Computer.Clock.TickCount))

                Return userTables
            Catch ex As Exception
                'MsgBox(ex.Message)
                SetLastError(ex)
                GenUtils.Message(GenUtils.MsgType.Critical, "EntLib - DAAB-GetTables", ex.Message)
                Return Nothing
            End Try
        End Function

        Public Function GetTables(ByVal workDir As String) As DataTable
            Dim startTempTime As Integer = My.Computer.Clock.TickCount
            Dim userTables As DataTable = Nothing
            ' We only want user tables, not system tables
            Dim restrictions() As String = New String(3) {}
            Dim sql As String = ""

            restrictions(3) = "Table"

            SetLastError(Nothing)

            sql = "SELECT * FROM qsysMSysObjects WHERE OBJ_TYPE = '" & restrictions(3).ToUpper & "'"
            Try
                Return GetDataSet(sql).Tables(0)
            Catch ex As Exception
                'MsgBox(ex.Message)
                SetLastError(ex)
                GenUtils.Message(GenUtils.MsgType.Critical, "EntLib - DAAB-GetTables", ex.Message)
                Return Nothing
            End Try

            'If _db Is Nothing Then
            '    ConnectToDB()
            'End If

            'Try
            '    Dim conn As DbConnection = _db.CreateConnection()
            '    conn.Open()
            '    userTables = conn.GetSchema("Tables", restrictions)

            '    EntLib.Log.Log(workDir, "EntLib - DAAB.GetTables - GetSchema: ", Space(7) & EntLib.GenUtils.FormatTime(startTempTime, My.Computer.Clock.TickCount))

            '    Return userTables
            'Catch ex As Exception
            '    MsgBox(ex.Message)
            '    Return Nothing
            'End Try
        End Function

        Public Function GetViews() As DataTable
            Dim userTables As DataTable = Nothing
            ' We only want user tables, not system tables
            Dim restrictions() As String = New String(3) {}
            restrictions(3) = "View"

            SetLastError(Nothing)

            If _db Is Nothing Then
                ConnectToDB()
            End If

            Try
                Dim conn As DbConnection = _db.CreateConnection()
                conn.Open()
                userTables = conn.GetSchema("Tables", restrictions)

                Return userTables
            Catch ex As Exception
                'MsgBox(ex.Message)
                SetLastError(ex)
                GenUtils.Message(GenUtils.MsgType.Critical, "EntLib - DAAB-GetViews", ex.Message)
                Return Nothing
            End Try
        End Function

        Public Function GetProcedures() As DataTable
            Dim userTables As DataTable = Nothing
            ' We only want user tables, not system tables
            Dim restrictions() As String = New String(3) {}
            restrictions(3) = "Procedures"

            SetLastError(Nothing)

            If _db Is Nothing Then
                ConnectToDB()
            End If

            Try
                Dim conn As DbConnection = _db.CreateConnection()
                conn.Open()
                userTables = conn.GetSchema("Tables", restrictions)

                Return userTables
            Catch ex As Exception
                'MsgBox(ex.Message)
                SetLastError(ex)
                GenUtils.Message(GenUtils.MsgType.Critical, "EntLib - DAAB-GetProcedures", ex.Message)
                Return Nothing
            End Try
        End Function

        Public Function GetColumns() As DataTable
            Dim userTables As DataTable = Nothing
            ' We only want user tables, not system tables
            Dim restrictions() As String = New String(3) {}
            restrictions(3) = "Columns"

            LastErrorNo = 0
            LastErrorDesc = ""

            SetLastError(Nothing)

            If _db Is Nothing Then
                ConnectToDB()
            End If

            Try
                Dim conn As DbConnection = _db.CreateConnection()
                conn.Open()
                userTables = conn.GetSchema("Tables", restrictions)

                Return userTables
            Catch ex As Exception
                'MsgBox(ex.Message)
                SetLastError(ex)
                GenUtils.Message(GenUtils.MsgType.Critical, "EntLib - DAAB-GetColumns", ex.Message)
                Return Nothing
            End Try
        End Function

        Public Function GetIndexes() As DataTable
            Dim userTables As DataTable = Nothing
            ' We only want user tables, not system tables
            Dim restrictions() As String = New String(3) {}
            restrictions(3) = "Indexes"

            SetLastError(Nothing)

            If _db Is Nothing Then
                ConnectToDB()
            End If

            Try
                Dim conn As DbConnection = _db.CreateConnection()
                conn.Open()
                userTables = conn.GetSchema("Tables", restrictions)

                Return userTables
            Catch ex As Exception
                'MsgBox(ex.Message)
                SetLastError(ex)
                GenUtils.Message(GenUtils.MsgType.Critical, "EntLib - DAAB-GetIndexes", ex.Message)
                Return Nothing
            End Try
        End Function

        Public Function GetCatalog() As DataTable
            Dim userTables As DataTable = Nothing
            ' We only want user tables, not system tables
            Dim restrictions() As String = New String(3) {}
            restrictions(3) = "Catalog"

            SetLastError(Nothing)

            If _db Is Nothing Then
                ConnectToDB()
            End If

            Try
                Dim conn As DbConnection = _db.CreateConnection()
                conn.Open()
                userTables = conn.GetSchema("Tables", restrictions)

                Return userTables
            Catch ex As Exception
                'MsgBox(ex.Message)
                SetLastError(ex)
                GenUtils.Message(GenUtils.MsgType.Critical, "EntLib - DAAB-GetCatalog", ex.Message)
                Return Nothing
            End Try
        End Function

        Public Function GetProcedureParameters() As DataTable
            Dim userTables As DataTable = Nothing
            ' We only want user tables, not system tables
            Dim restrictions() As String = New String(3) {}
            restrictions(3) = "ProcedureParameters"

            SetLastError(Nothing)

            If _db Is Nothing Then
                ConnectToDB()
            End If

            Try
                Dim conn As DbConnection = _db.CreateConnection()
                conn.Open()
                userTables = conn.GetSchema("Tables", restrictions)

                Return userTables
            Catch ex As Exception
                'MsgBox(ex.Message)
                SetLastError(ex)
                GenUtils.Message(GenUtils.MsgType.Critical, "EntLib - DAAB-GetProcedureParameters", ex.Message)
                Return Nothing
            End Try
        End Function

        Public Function GetTablesViews() As DataTable
            'http://msdn2.microsoft.com/en-us/library/ms254969(VS.80).aspx
            Dim dtTables As DataTable
            Dim dtViews As DataTable

            SetLastError(Nothing)

            Try
                dtTables = GetTables()
                dtViews = GetViews()

                dtTables.Merge(dtViews)

                dtTables.Merge(GetProcedures())

                dtTables.Merge(GetColumns())

                dtTables.Merge(GetIndexes())

                dtTables.Merge(GetCatalog())

                dtTables.Merge(GetProcedureParameters())

                Return dtTables
            Catch ex As Exception
                'MsgBox(ex.Message)
                SetLastError(ex)
                GenUtils.Message(GenUtils.MsgType.Critical, "EntLib - DAAB-GetTableViews", ex.Message)
                Return Nothing
            Finally

            End Try
        End Function

        Public Function GetSchema() As DataTable
            Dim userTables As DataTable = Nothing
            ' We only want user tables, not system tables
            Dim restrictions() As String = New String(3) {}
            restrictions(3) = "Table"

            SetLastError(Nothing)

            If _db Is Nothing Then
                ConnectToDB()
            End If

            Try
                Dim conn As DbConnection = _db.CreateConnection()
                conn.Open()
                userTables = conn.GetSchema("Tables", restrictions)

                Return userTables
            Catch ex As Exception
                'MsgBox(ex.Message)
                SetLastError(ex)
                GenUtils.Message(GenUtils.MsgType.Critical, "EntLib - DAAB-GetSchema", ex.Message)
                Return Nothing
            End Try
        End Function

        Public Function GetConnectionString() As String
            Dim dbConfigView As New DatabaseConfigurationView(New SystemConfigurationSource())
            Dim defaultConnStr As String = dbConfigView.DefaultName

            'Return _db.ConnectionStringWithoutCredentials
            Return _db.ConnectionString() 'returns password, if any.
        End Function

        Public Function GetProviderName() As String
            'Dim dbConfigView As New DatabaseConfigurationView(New SystemConfigurationSource())
            'dbconfigview.GetConnectionStringSettings(

            Return GetConfig("").ConnectionStrings.ConnectionStrings(_databaseName).ProviderName

            'Return Nothing
        End Function

        Public Function GetProviderName(ByVal configFilePath As String) As String
            'RKP/06-15-10
            'Dim dbConfigView As New DatabaseConfigurationView(New SystemConfigurationSource())
            'dbconfigview.GetConnectionStringSettings(

            If configFilePath = "" Then
                configFilePath = System.Configuration.ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None).FilePath
            End If

            Return GetConfig(configFilePath).ConnectionStrings.ConnectionStrings(_databaseName).ProviderName

            'Return Nothing
        End Function

        ' Updates the database.
        ' Returns: The number of rows affected by the update.
        ' Remarks: Demonstrates updating a database using a DataSet.
        Public Function UpdateDataSet(ByRef ds As DataSet, ByRef tableName As String, ByVal insertCommand As DbCommand, ByVal updateCommand As DbCommand, ByVal deleteCommand As DbCommand) As Integer
            ' Create the Database object, using the default database service. The
            ' default database service is determined through configuration.
            'Dim db As Database = DatabaseFactory.CreateDatabase()

            Dim rowsAffected As Integer

            SetLastError(Nothing)

            If _db Is Nothing Then
                ConnectToDB()
            End If

            Try
                rowsAffected = _db.UpdateDataSet(ds, tableName, insertCommand, updateCommand, deleteCommand, UpdateBehavior.Standard)
            Catch ex As Exception
                'MsgBox(ex.Message)
                SetLastError(ex)
                GenUtils.Message(GenUtils.MsgType.Critical, "EntLib - DAAB-UpdateDataSet", ex.Message)
            End Try

            Return rowsAffected

            'Dim productsDataSet As DataSet = New DataSet

            'Dim sqlCommand As String = "Select ProductID, ProductName, CategoryID, UnitPrice, LastUpdate " & _
            '"From Products"
            'Dim dbCommand As DbCommand = _db.GetSqlStringCommand(sqlCommand)

            'Dim productsTable As String = "Products"

            ' Retrieve the initial data
            '_db.LoadDataSet(dbCommand, productsDataSet, productsTable)

            ' Get the table that will be modified
            'Dim table As DataTable = productsDataSet.Tables(productsTable)

            ' Add a new product to existing DataSet
            'Dim addedRow As DataRow = table.Rows.Add(New Object() {DBNull.Value, "New product", 11, 25})

            ' Modify an existing product
            'table.Rows(0)("ProductName") = "Modified product"

            ' Establish our Insert, Delete, and Update commands
            'Dim insertCommand As DbCommand = _db.GetStoredProcCommand("AddProduct")
            '_db.AddInParameter(insertCommand, "ProductName", DbType.String, "ProductName", DataRowVersion.Current)
            '_db.AddInParameter(insertCommand, "CategoryID", DbType.Int32, "CategoryID", DataRowVersion.Current)
            '_db.AddInParameter(insertCommand, "UnitPrice", DbType.Currency, "UnitPrice", DataRowVersion.Current)

            'Dim deleteCommand As DbCommand = _db.GetStoredProcCommand("DeleteProduct")
            '_db.AddInParameter(deleteCommand, "ProductID", DbType.Int32, "ProductID", DataRowVersion.Current)

            'Dim updateCommand As DbCommand = _db.GetStoredProcCommand("UpdateProduct")
            '_db.AddInParameter(updateCommand, "ProductID", DbType.Int32, "ProductID", DataRowVersion.Current)
            '_db.AddInParameter(updateCommand, "ProductName", DbType.String, "ProductName", DataRowVersion.Current)
            '_db.AddInParameter(updateCommand, "LastUpdate", DbType.DateTime, "LastUpdate", DataRowVersion.Current)
            '_db.UpdateDataSet(

            ' Submit the DataSet, capturing the number of rows that were affected


        End Function

        Public Function UpdateDataSet(ByRef ds As DataSet, ByRef sql As String) As Integer
            'This function works ONLY when the table has a Primary Key.

            'Dim sql As String = "SELECT * FROM tsysMapPoint"
            'Dim connStr As String = _currentDb.GetConnectionString()
            'Dim data_adapter As New OleDb.OleDbDataAdapter(sql, connStr)
            'Dim command_builder As New OleDb.OleDbCommandBuilder(data_adapter)
            'Dim rowsAffected As Integer = data_adapter.Update(_dataSet)

            Dim data_adapter As OleDb.OleDbDataAdapter
            Dim command_builder As OleDb.OleDbCommandBuilder 'New OleDb.OleDbCommandBuilder(data_adapter)
            'Dim data_adapter As New OleDb.OleDbDataAdapter(sql, GetConnectionString())
            'Dim cmdSQL As System.Data.SqlClient.SqlCommand
            Dim ret As Integer = 0
            'Dim command_builder As New OleDb.OleDbCommandBuilder(data_adapter)

            Dim data_adapter_sql As System.Data.SqlClient.SqlDataAdapter
            Dim command_builder_sql As System.Data.SqlClient.SqlCommandBuilder

            SetLastError(Nothing)

            If IsSQLExpress() Then
                'cmdSQL = New System.Data.SqlClient.SqlCommand("SQLServer", _connSQLExpress)
                'cmdSQL.CommandType = System.Data.CommandType.Text
                'cmdSQL.CommandText = sql
                'Try
                '    ret = cmdSQL.ExecuteNonQuery()
                'Catch ex As Exception
                '    'GenUtils.Log("Error executing SQL:" & vbNewLine & sql)
                '    'GenUtils.Log(ex.Message)
                '    GenUtils.Message(GenUtils.MsgType.Critical, "EntLib - DAAB-UpdateDataSet", ex.Message)
                'End Try

                data_adapter_sql = New System.Data.SqlClient.SqlDataAdapter(sql, GetConnectionString())
                command_builder_sql = New System.Data.SqlClient.SqlCommandBuilder(data_adapter_sql)

                Try
                    Return data_adapter_sql.Update(ds)
                Catch ex As Exception
                    SetLastError(ex)
                    GenUtils.Message(GenUtils.MsgType.Critical, "EntLib - DAAB-UpdateDataSet", ex.Message)
                    Return Nothing
                End Try
            Else
                Try
                    data_adapter = New OleDb.OleDbDataAdapter(sql, GetConnectionString())
                    command_builder = New OleDb.OleDbCommandBuilder(data_adapter)
                    Return data_adapter.Update(ds)
                Catch ex As Exception
                    'MsgBox(ex.Message)
                    SetLastError(ex)
                    GenUtils.Message(GenUtils.MsgType.Critical, "EntLib - DAAB-UpdateDataSet", ex.Message)
                    Return Nothing
                End Try
            End If




        End Function

        ''' <summary>
        ''' Updates underlying table with data from the DataSet.
        ''' </summary>
        ''' <param name="dt"></param>
        ''' <param name="sql"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' This is the version that is called most often.
        ''' </remarks>
        Public Function UpdateDataSet(ByRef dt As DataTable, ByRef sql As String) As Integer
            'This function works ONLY when the table has a Primary Key.

            'Dim sql As String = "SELECT * FROM tsysMapPoint"
            'Dim connStr As String = _currentDb.GetConnectionString()
            'Dim data_adapter As New OleDb.OleDbDataAdapter(sql, connStr)
            'Dim command_builder As New OleDb.OleDbCommandBuilder(data_adapter)
            'Dim rowsAffected As Integer = data_adapter.Update(_dataSet)

            'Dim cmdSQL As System.Data.SqlClient.SqlCommand

            Dim ret As Integer = 0
            Dim data_adapter As OleDb.OleDbDataAdapter 'New OleDb.OleDbDataAdapter(sql, GetConnectionString())
            Dim command_builder As OleDb.OleDbCommandBuilder 'New OleDb.OleDbCommandBuilder(data_adapter)
            Dim data_adapter_sql As System.Data.SqlClient.SqlDataAdapter
            Dim command_builder_sql As System.Data.SqlClient.SqlCommandBuilder
            Dim proceed As Boolean = False
            Dim bulkCopy As SqlBulkCopy
            Dim ctr As Integer

            SetLastError(Nothing)

            'http://msdn.microsoft.com/en-us/library/kbbwt18a%28VS.80%29.aspx
            'data_adapter.UpdateBatchSize = 0

            If IsSQLExpress() Then
                'cmdSQL = New System.Data.SqlClient.SqlCommand("SQLServer", _connSQLExpress)
                'cmdSQL.CommandType = System.Data.CommandType.Text
                'cmdSQL.CommandText = sql
                'Try
                '    ret = cmdSQL.ExecuteNonQuery()
                'Catch ex As Exception
                '    'GenUtils.Log("Error executing SQL:" & vbNewLine & sql)
                '    'GenUtils.Log(ex.Message)
                '    GenUtils.Message(GenUtils.MsgType.Critical, "EntLib - DAAB-UpdateDataSet", ex.Message)
                'End Try

                data_adapter_sql = New System.Data.SqlClient.SqlDataAdapter(sql, GetConnectionString())
                command_builder_sql = New System.Data.SqlClient.SqlCommandBuilder(data_adapter_sql)
                Try
                    Return data_adapter_sql.Update(dt)
                Catch ex As Exception
                    SetLastError(ex)
                    GenUtils.Message(GenUtils.MsgType.Critical, "EntLib - DAAB-UpdateDataSet", ex.Message & vbNewLine & "In-memory DataTable row count = " & dt.Rows.Count & vbNewLine & "Please check this row count with the database table in the query:" & vbNewLine & sql & vbNewLine & "A mismatch is a sign that there is a model error.")
                    Return Nothing
                End Try
            Else
                If GenUtils.IsSwitchAvailable(_switches, "/UseMinSysRes") Then
                    If GenUtils.IsSwitchAvailable(_switches, "/UseSQLServerSyntax") Then
                        If sql.StartsWith("SELECT") Then
                            proceed = True
                        Else
                            'Use BulkCopy
                            bulkCopy = New SqlBulkCopy(_connSQLExpress)
                            bulkCopy.BulkCopyTimeout = 0
                            bulkCopy.BatchSize = 5000
                            bulkCopy.DestinationTableName = sql
                            bulkCopy.ColumnMappings.Clear()
                            For ctr = 0 To dt.Columns.Count - 1
                                bulkCopy.ColumnMappings.Add(dt.Columns(ctr).ColumnName, dt.Columns(ctr).ColumnName)
                            Next
                            Try
                                bulkCopy.WriteToServer(dt)
                                bulkCopy.Close()
                                bulkCopy = Nothing
                            Catch ex As Exception
                                SetLastError(ex)
                                'GenUtils.Message(GenUtils.MsgType.Critical, "EntLib - DAAB-UpdateDataSet", ex.Message)
                                GenUtils.Message(GenUtils.MsgType.Critical, "EntLib - DAAB-UpdateDataSet", ex.Message & vbNewLine & "In-memory DataTable row count = " & dt.Rows.Count & vbNewLine & "Please check this row count with the database table in the query:" & vbNewLine & sql & vbNewLine & "A mismatch is a sign that there is a model error.")
                            End Try
                        End If
                    Else
                        proceed = True
                    End If
                Else
                    proceed = True
                End If

                If proceed Then
                    data_adapter = New OleDb.OleDbDataAdapter(sql, GetConnectionString())
                    command_builder = New OleDb.OleDbCommandBuilder(data_adapter)
                    Try
                        Return data_adapter.Update(dt)
                    Catch ex As Exception

                        SetLastError(ex)
                        'GenUtils.Message(GenUtils.MsgType.Critical, "EntLib - DAAB-UpdateDataSet", ex.Message)
                        GenUtils.Message(GenUtils.MsgType.Critical, "EntLib - DAAB-UpdateDataSet", ex.Message & vbNewLine & "In-memory DataTable row count = " & dt.Rows.Count & vbNewLine & "Please check this row count with the database table in the query:" & vbNewLine & sql & vbNewLine & "A mismatch is a sign that there is a model error.")
                    End Try
                End If
            End If

        End Function

        Public Function UpdateDataSet(ByRef dt As DataTable, ByRef sql As String, ByVal switches() As String) As Integer
            'This function works ONLY when the table has a Primary Key.

            'Dim sql As String = "SELECT * FROM tsysMapPoint"
            'Dim connStr As String = _currentDb.GetConnectionString()
            'Dim data_adapter As New OleDb.OleDbDataAdapter(sql, connStr)
            'Dim command_builder As New OleDb.OleDbCommandBuilder(data_adapter)
            'Dim rowsAffected As Integer = data_adapter.Update(_dataSet)

            'Dim cmdSQL As System.Data.SqlClient.SqlCommand
            Dim ret As Integer = 0

            Dim data_adapter As OleDb.OleDbDataAdapter 'New OleDb.OleDbDataAdapter(sql, GetConnectionString())
            Dim command_builder As OleDb.OleDbCommandBuilder 'New OleDb.OleDbCommandBuilder(data_adapter)

            Dim data_adapter_sql As System.Data.SqlClient.SqlDataAdapter
            Dim command_builder_sql As System.Data.SqlClient.SqlCommandBuilder

            Dim proceed As Boolean = False

            Dim bulkCopy As SqlBulkCopy
            'Dim cmdSQL As System.Data.SqlClient.SqlCommand
            Dim ctr As Integer

            _switches = switches

            SetLastError(Nothing)

            'http://msdn.microsoft.com/en-us/library/kbbwt18a%28VS.80%29.aspx
            'data_adapter.UpdateBatchSize = 0

            If IsSQLExpress() Then
                data_adapter_sql = New System.Data.SqlClient.SqlDataAdapter(sql, GetConnectionString())
                command_builder_sql = New System.Data.SqlClient.SqlCommandBuilder(data_adapter_sql)
                Try
                    Return data_adapter_sql.Update(dt)
                Catch ex As Exception
                    SetLastError(ex)
                    GenUtils.Message(GenUtils.MsgType.Critical, "EntLib - DAAB-UpdateDataSet", ex.Message)
                    Return Nothing
                End Try
            Else
                If GenUtils.IsSwitchAvailable(_switches, "/UseMinSysRes") Then
                    If GenUtils.IsSwitchAvailable(_switches, "/UseSQLServerSyntax") Then
                        If sql.StartsWith("SELECT") Then
                            proceed = True
                        Else
                            'Use BulkCopy
                            If _connSQLExpress Is Nothing Then
                                ConnectToDB(_databaseName, _switches)
                            End If

                            'cmdSQL = New System.Data.SqlClient.SqlCommand("SQLServer", _connSQLExpress)
                            'cmdSQL.CommandType = System.Data.CommandType.Text
                            'cmdSQL.CommandText = "DELETE FROM " & sql
                            'Try
                            '    ret = cmdSQL.ExecuteNonQuery()
                            '    'Return ret
                            'Catch ex As Exception
                            '    'GenUtils.Log("Error executing SQL:" & vbNewLine & sql)
                            '    'GenUtils.Log(ex.Message)
                            '    SetLastError(ex)
                            '    'GenUtils.Message(GenUtils.MsgType.Critical, "EntLib - DAAB-ExecuteNonQuery", ex.Message)
                            'End Try

                            bulkCopy = New SqlBulkCopy(_connSQLExpress)
                            bulkCopy.BulkCopyTimeout = 0
                            bulkCopy.BatchSize = 5000
                            bulkCopy.DestinationTableName = sql
                            bulkCopy.ColumnMappings.Clear()
                            For ctr = 0 To dt.Columns.Count - 1
                                bulkCopy.ColumnMappings.Add(dt.Columns(ctr).ColumnName, dt.Columns(ctr).ColumnName)
                            Next
                            Try
                                bulkCopy.WriteToServer(dt)
                                bulkCopy.Close()
                                bulkCopy = Nothing
                            Catch ex As Exception
                                SetLastError(ex)
                                GenUtils.Message(GenUtils.MsgType.Critical, "EntLib - DAAB-UpdateDataSet", ex.Message)
                            End Try
                        End If
                    Else
                        proceed = True
                    End If
                Else
                    proceed = True
                End If

                If proceed Then
                    data_adapter = New OleDb.OleDbDataAdapter(sql, GetConnectionString())
                    command_builder = New OleDb.OleDbCommandBuilder(data_adapter)
                    Try
                        Return data_adapter.Update(dt)
                    Catch ex As Exception
                        SetLastError(ex)
                        GenUtils.Message(GenUtils.MsgType.Critical, "EntLib - DAAB-UpdateDataSet", ex.Message)
                    End Try
                End If
            End If

        End Function

        Public Function UpdateDataSet(ByRef dt As DataTable, ByRef sql As String, ByVal switches() As String, ByVal useBulkCopy As Boolean) As Integer
            'This function works ONLY when the table has a Primary Key.

            'Dim sql As String = "SELECT * FROM tsysMapPoint"
            'Dim connStr As String = _currentDb.GetConnectionString()
            'Dim data_adapter As New OleDb.OleDbDataAdapter(sql, connStr)
            'Dim command_builder As New OleDb.OleDbCommandBuilder(data_adapter)
            'Dim rowsAffected As Integer = data_adapter.Update(_dataSet)

            'Dim cmdSQL As System.Data.SqlClient.SqlCommand
            Dim ret As Integer = 0

            Dim data_adapter As OleDb.OleDbDataAdapter 'New OleDb.OleDbDataAdapter(sql, GetConnectionString())
            Dim command_builder As OleDb.OleDbCommandBuilder 'New OleDb.OleDbCommandBuilder(data_adapter)

            Dim data_adapter_sql As System.Data.SqlClient.SqlDataAdapter
            Dim command_builder_sql As System.Data.SqlClient.SqlCommandBuilder

            Dim proceed As Boolean = False

            Dim bulkCopy As SqlBulkCopy
            'Dim cmdSQL As System.Data.SqlClient.SqlCommand
            Dim ctr As Integer

            _switches = switches

            SetLastError(Nothing)

            'http://msdn.microsoft.com/en-us/library/kbbwt18a%28VS.80%29.aspx
            'data_adapter.UpdateBatchSize = 0

            If useBulkCopy Then
                If sql.StartsWith("SELECT") Then
                    proceed = True
                Else
                    'Use BulkCopy
                    If _connSQLExpress Is Nothing Then
                        ConnectToDB(_databaseName, _switches)
                    End If

                    'cmdSQL = New System.Data.SqlClient.SqlCommand("SQLServer", _connSQLExpress)
                    'cmdSQL.CommandType = System.Data.CommandType.Text
                    'cmdSQL.CommandText = "DELETE FROM " & sql
                    'Try
                    '    ret = cmdSQL.ExecuteNonQuery()
                    '    'Return ret
                    'Catch ex As Exception
                    '    'GenUtils.Log("Error executing SQL:" & vbNewLine & sql)
                    '    'GenUtils.Log(ex.Message)
                    '    SetLastError(ex)
                    '    'GenUtils.Message(GenUtils.MsgType.Critical, "EntLib - DAAB-ExecuteNonQuery", ex.Message)
                    'End Try

                    bulkCopy = New SqlBulkCopy(_connSQLExpress)
                    bulkCopy.BulkCopyTimeout = 0
                    bulkCopy.BatchSize = 5000
                    bulkCopy.DestinationTableName = sql
                    bulkCopy.ColumnMappings.Clear()

                    'NOTE: Column names are case-sensitive.
                    For ctr = 0 To dt.Columns.Count - 1
                        bulkCopy.ColumnMappings.Add(dt.Columns(ctr).ColumnName.Trim(), dt.Columns(ctr).ColumnName.Trim())
                    Next
                    Try
                        bulkCopy.WriteToServer(dt)
                        bulkCopy.Close()
                        bulkCopy = Nothing
                    Catch ex As Exception
                        SetLastError(ex)
                        GenUtils.Message(GenUtils.MsgType.Critical, "EntLib - DAAB-UpdateDataSet", ex.Message)
                    End Try
                End If
            Else
                proceed = True
            End If

            If proceed Then
                If IsSQLExpress() Then
                    data_adapter_sql = New System.Data.SqlClient.SqlDataAdapter(sql, GetConnectionString())
                    command_builder_sql = New System.Data.SqlClient.SqlCommandBuilder(data_adapter_sql)
                    Try
                        Return data_adapter_sql.Update(dt)
                    Catch ex As Exception
                        SetLastError(ex)
                        GenUtils.Message(GenUtils.MsgType.Critical, "EntLib - DAAB-UpdateDataSet", ex.Message)
                        Return Nothing
                    End Try
                Else
                    data_adapter = New OleDb.OleDbDataAdapter(sql, GetConnectionString())
                    command_builder = New OleDb.OleDbCommandBuilder(data_adapter)
                    Try
                        Return data_adapter.Update(dt)
                    Catch ex As Exception
                        SetLastError(ex)
                        GenUtils.Message(GenUtils.MsgType.Critical, "EntLib - DAAB-UpdateDataSet", ex.Message)
                    End Try
                End If
            End If
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
        Public Function UpdateDB(ByVal connectionString As String, ByRef dt As DataTable, ByVal sql As String) As Integer
            Dim data_adapter As OleDb.OleDbDataAdapter

            data_adapter = New OleDb.OleDbDataAdapter(sql, connectionString)

            Dim command_builder As New OleDb.OleDbCommandBuilder(data_adapter)

            Return data_adapter.Update(dt)
        End Function

        Public Sub SaveADORS(ByRef adoRS As ADODB.Recordset, ByVal filePath As String)
            adoRS.Save(filePath, ADODB.PersistFormatEnum.adPersistXML)
        End Sub

        '**************************************************************************
        '   Method Name : GetADORS
        '   Description : Takes a DataSet and converts into a Recordset. The converted 
        '                 ADODB recordset is saved as an XML file. The data is saved 
        '                 to the file path passed as parameter.
        '   Output      : The output of this method is long. Returns 1 if successfull. 
        '                 If not throws an exception. 
        '   Input parameters:
        '               1. DataSet object
        '               2. Database Name
        '               3. Output file - where the converted should be written.
        '**************************************************************************
        Public Function GetADORS(ByVal ds As DataSet, ByVal dbname As String, ByVal tableName As String, ByVal xslfile As String, _
          ByVal outputfile As String) As Long

            'Create an xmlwriter object, to write the ADO Recordset Format XML
            Try
                Dim xwriter As New XmlTextWriter(outputfile, System.Text.Encoding.Default)

                'call this Sub to write the ADONamespaces to the XMLTextWriter
                WriteADONamespaces(xwriter)
                'call this Sub to write the ADO Recordset Schema
                WriteSchemaElement(ds, dbname, xwriter)

                Dim TransformedDatastrm As New MemoryStream
                'Call this Function to transform the Dataset xml to ADO Recordset XML
                TransformedDatastrm = TransformData(ds, xslfile, tableName)
                'Pass the Transformed ADO REcordset XML to this Sub
                'to write in correct format.
                HackADOXML(xwriter, TransformedDatastrm)

                xwriter.Flush()
                xwriter.Close()
                'returns 1 if success
                Return 1

            Catch ex As Exception
                'Returns error message to the calling function.
                'Err.Raise(100, ex.Source, ex.ToString)
                GenUtils.Message(GenUtils.MsgType.Critical, "EntLib - DAAB-GetADORS", ex.Message)
            End Try

        End Function


        Private Sub WriteADONamespaces(ByRef writer As XmlTextWriter)
            'The following is to specify the encoding of the xml file
            'writer.WriteProcessingInstruction("xml", "version='1.0' encoding='ISO-8859-1'")

            'The following is the ado recordset format
            '<xml xmlns:s='uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882' 
            '        xmlns:dt='uuid:C2F41010-65B3-11d1-A29F-00AA00C14882'
            '        xmlns:rs='urn:schemas-microsoft-com:rowset' 
            '        xmlns:z='#RowsetSchema'>
            '    </xml>

            'Write the root element
            writer.WriteStartElement("", "xml", "")

            'Append the ADO Recordset namespaces
            writer.WriteAttributeString("xmlns", "s", Nothing, "uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882")
            writer.WriteAttributeString("xmlns", "dt", Nothing, "uuid:C2F41010-65B3-11d1-A29F-00AA00C14882")
            writer.WriteAttributeString("xmlns", "rs", Nothing, "urn:schemas-microsoft-com:rowset")
            writer.WriteAttributeString("xmlns", "z", Nothing, "#RowsetSchema")
            writer.Flush()

        End Sub


        Private Sub WriteSchemaElement(ByVal ds As DataSet, ByVal dbname As String, ByRef writer As XmlTextWriter)
            'ADO Recordset format for defining the schema
            ' <s:Schema id='RowsetSchema'>
            '            <s:ElementType name='row' content='eltOnly' rs:updatable='true'>
            '            </s:ElementType>
            '        </s:Schema>

            'write element schema
            writer.WriteStartElement("s", "Schema", "uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882")
            writer.WriteAttributeString("id", "RowsetSchema")

            'write element ElementTyoe
            writer.WriteStartElement("s", "ElementType", "uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882")

            'write the attributes for ElementType
            writer.WriteAttributeString("name", "", "row")
            writer.WriteAttributeString("content", "", "eltOnly")
            writer.WriteAttributeString("rs", "updatable", "urn:schemas-microsoft-com:rowset", "true")

            WriteSchema(ds, dbname, writer)
            'write the end element for ElementType
            writer.WriteFullEndElement()

            'write the end element for Schema 
            writer.WriteFullEndElement()
            writer.Flush()
        End Sub


        Private Sub WriteSchema(ByVal ds As DataSet, ByVal dbname As String, ByRef writer As XmlTextWriter)

            Dim i As Int32 = 1
            Dim dc As DataColumn

            For Each dc In ds.Tables(0).Columns

                dc.ColumnMapping = MappingType.Attribute

                writer.WriteStartElement("s", "AttributeType", "uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882")
                'write all the attributes 
                writer.WriteAttributeString("name", "", dc.ToString)
                writer.WriteAttributeString("rs", "number", "urn:schemas-microsoft-com:rowset", i.ToString)
                writer.WriteAttributeString("rs", "baseCatalog", "urn:schemas-microsoft-com:rowset", dbname)
                writer.WriteAttributeString("rs", "baseTable", "urn:schemas-microsoft-com:rowset", _
                      dc.Table.TableName.ToString)
                writer.WriteAttributeString("rs", "keycolumn", "urn:schemas-microsoft-com:rowset", _
                      dc.Unique.ToString)
                writer.WriteAttributeString("rs", "autoincrement", "urn:schemas-microsoft-com:rowset", _
                      dc.AutoIncrement.ToString)
                'write child element
                writer.WriteStartElement("s", "datatype", "uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882")
                'write attributes
                writer.WriteAttributeString("dt", "type", "uuid:C2F41010-65B3-11d1-A29F-00AA00C14882", _
                      GetDatatype(dc.DataType.ToString))
                writer.WriteAttributeString("dt", "maxlength", "uuid:C2F41010-65B3-11d1-A29F-00AA00C14882", _
                      dc.MaxLength.ToString)
                writer.WriteAttributeString("rs", "maybenull", "urn:schemas-microsoft-com:rowset", _
                      dc.AllowDBNull.ToString)
                'write end element for datatype
                writer.WriteEndElement()
                'end element for AttributeType
                writer.WriteEndElement()
                writer.Flush()
                i = i + 1
            Next
            dc = Nothing

        End Sub


        'Function to get the ADO compatible datatype
        Private Function GetDatatype(ByVal dtype As String) As String
            Select Case (dtype)
                Case "System.Int32"
                    Return "int"
                Case "System.DateTime"
                    Return "dateTime"
                Case Else
                    Return ""
            End Select
        End Function


        'Transform the data set format to ADO Recordset format
        'This only transforms the data
        Private Function TransformData(ByVal ds As DataSet, ByVal xslfile As String, ByVal tableName As String) As MemoryStream

            Dim instream As New MemoryStream
            Dim outstream As New MemoryStream

            'write the xml into a memorystream
            ds.WriteXml(instream, XmlWriteMode.IgnoreSchema)
            instream.Position = 0

            'load the xsl document
            Dim xslt As New XslCompiledTransform
            xslt.Load(xslfile)

            'create the xmltextreader using the memory stream
            Dim xmltr As New XmlTextReader(instream)
            'create the xpathdoc
            Dim xpathdoc As XPathDocument = New XPathDocument(xmltr)

            'create XpathNavigator
            Dim nav As XPathNavigator
            nav = xpathdoc.CreateNavigator

            'Create the XsltArgumentList.
            Dim xslArg As XsltArgumentList = New XsltArgumentList

            'Create a parameter that represents the current date and time.
            'Dim tablename As String
            xslArg.AddParam(tableName, "", ds.Tables(0).TableName)

            'transform the xml to a memory stream
            xslt.Transform(nav, xslArg, outstream)

            instream = Nothing
            xslt = Nothing
            '        xmltr = Nothing
            xpathdoc = Nothing
            nav = Nothing

            Return outstream

        End Function


        '**************************************************************************
        '   Method Name : ConvertToRs
        '   Description : The XSLT does not tranform with fullendelements. For example, 
        '               <root attr=""/> intead of <root attr=""><root/>. ADO Recordset 
        '               cannot read this. This method is used to convert the 
        '               elements to have fullendelements.
        '**************************************************************************
        Private Sub HackADOXML(ByRef wrt As XmlTextWriter, ByVal ADOXmlStream As System.IO.MemoryStream)

            ADOXmlStream.Position = 0
            Dim rdr As New XmlTextReader(ADOXmlStream)
            Dim outStream As New MemoryStream
            'Dim wrt As New XmlTextWriter(outStream, System.Text.Encoding.Default)

            rdr.MoveToContent()
            'if the ReadState is not EndofFile, read the XmlTextReader for nodes.
            Do While rdr.ReadState <> ReadState.EndOfFile
                If rdr.Name = "s:Schema" Then
                    wrt.WriteNode(rdr, False)
                    wrt.Flush()
                ElseIf rdr.Name = "z:row" And rdr.NodeType = XmlNodeType.Element Then
                    wrt.WriteStartElement("z", "row", "#RowsetSchema")
                    rdr.MoveToFirstAttribute()
                    wrt.WriteAttributes(rdr, False)
                    wrt.Flush()
                ElseIf rdr.Name = "z:row" And rdr.NodeType = XmlNodeType.EndElement Then
                    'The following is the key statement that closes the z:row 
                    'element without generating a full end element
                    wrt.WriteEndElement()
                    wrt.Flush()
                ElseIf rdr.Name = "rs:data" And rdr.NodeType = XmlNodeType.Element Then
                    wrt.WriteStartElement("rs", "data", "urn:schemas-microsoft-com:rowset")
                ElseIf rdr.Name = "rs:data" And rdr.NodeType = XmlNodeType.EndElement Then
                    wrt.WriteEndElement()
                    wrt.Flush()
                End If
                rdr.Read()
            Loop

            wrt.WriteEndElement()
            wrt.Flush()
        End Sub

        Public Function GetScalarValue(ByVal sql As String) As String
            Return _db.ExecuteScalar(System.Data.CommandType.Text, sql).ToString()
        End Function

        Public Function GetDBConnectionString() As String
            Try
                'RKP/12-13-11/v3.0.156
                'Return _db.ConnectionStringWithoutCredentials.ToString()
                Return _db.ConnectionString().ToString()
            Catch ex As Exception
                Return Nothing
            End Try

        End Function

        Public Sub DoCumulativeTotals(ByVal tableName As String, ByVal srcField As String, ByVal destField As String)

        End Sub

        Private Sub ReadCSV()


            Dim InputFileFullPath As String = "c:\argus\Ftp"

            Dim InputFile As String = "File.csv"

            Dim myConnString As String = "Driver={Microsoft Text Driver (*.txt; *.csv)};DBQ=" & InputFileFullPath & ";DSN=" & InputFile & ";Extensions=csv;"

            Dim myConnection As New OdbcConnection(myConnString)

            Dim query As String

            Dim dTable As New DataTable("Prices")

            myConnection.Open()

            Dim adapter As New OdbcDataAdapter

            query = "Select * from " & InputFile

            adapter.SelectCommand = New OdbcCommand(query, myConnection)

            adapter.Fill(dTable)



            myConnection.Close()
        End Sub

        Public Function GetConfig(ByVal path As String) As System.Configuration.Configuration
            'RKP/06-15-10
            'http://www.renevo.com/blogs/vbdotnet/archive/2008/01/31/loading-config-files-from-non-default-locations.aspx
            'Dim retVal As Configuration = Nothing
            Dim configFileMap As New ExeConfigurationFileMap()

            If path = "" Then
                path = System.Configuration.ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None).FilePath
            End If

            configFileMap.ExeConfigFilename = path
            Return ConfigurationManager.OpenMappedExeConfiguration(configFileMap, ConfigurationUserLevel.None)
        End Function

        Public Function GetDatabaseName() As String
            'RKP/06-15-10
            Return _databaseName
        End Function

        Public Function GetViewDefinition(ByVal viewName As String) As String
            Dim conStr As String = _db.ConnectionString
            'Dim daoWorkspace As DAO.Workspace
            Dim daoDBEngine As DAO.DBEngine
            Dim daoDB As DAO.Database
            Dim daoQueryDef As DAO.QueryDef
            'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\OPTMODELS\FARMER\FARMER.MDB;Jet OLEDB:System Database=C:\OPTMODELS\C-OPTSYS\System.MDW.COPT
            Dim dbPath As String
            Dim pos1 As Integer
            Dim pos2 As Integer
            Dim sql As String
            Dim dt As DataTable = Nothing
            Dim ctr As Integer = 0
            'Dim dtRow As DataRow
            Dim sb As StringBuilder = Nothing
            Dim keyCtr As Integer = 0

            SetLastError(Nothing)


            If conStr.Contains("Microsoft.Jet") Or conStr.Contains("Microsoft.ACE") Then
                pos1 = InStr(1, conStr, "Data Source=", CompareMethod.Text)
                If pos1 > 0 Then
                    dbPath = Mid(conStr, pos1 + 12)
                    pos2 = InStr(1, dbPath, ";", vbTextCompare)
                    dbPath = Mid(dbPath, 1, pos2 - 1)
                Else
                    pos1 = InStr(conStr, "Data Source =", CompareMethod.Text)
                    If pos1 > 0 Then
                        dbPath = Mid(conStr, pos1 + 12)
                        pos2 = InStr(1, dbPath, ";", vbTextCompare)
                        dbPath = Mid(dbPath, 1, pos2 - 1)
                    Else
                        Return ""
                    End If
                End If

                daoDBEngine = New Dao.DBEngine
                daoDB = daoDBEngine.OpenDatabase(dbPath)

                For Each daoQueryDef In daoDB.QueryDefs
                    If daoQueryDef.Name.Equals(viewName) Then
                        Return daoQueryDef.SQL.Trim
                        Exit For
                    End If
                Next
            ElseIf IsSQLExpress() Then 'RKP/11-22-11/v3.0.153
                'EXEC dbo.sp_helptext 'dbo.qMTXappCBBDSALES'
                sql = "EXEC dbo.sp_helptext 'dbo." & viewName & "'"
                EntLib.COPT.Log.Log("GetViewDefinition: " & sql)
                dt = GetDataTable(sql)
                If dt Is Nothing Then
                    Return ""
                Else

                    ctr = 0
                    sb = New StringBuilder()
                    For ctr = 0 To dt.Rows.Count - 1
                        'Debug.Print(dt.Rows(ctr).Item(0).ToString())
                        'EntLib.COPT.Log.Log(dt.Rows(ctr).Item(0).ToString())
                        If dt.Rows(ctr).Item(0).ToString().Trim().ToUpper().StartsWith("BEGIN") Then
                            keyCtr += 1
                        Else
                            If keyCtr = 1 Then
                                If dt.Rows(ctr).Item(0).ToString().Trim().ToUpper().StartsWith("--SPEND") Then
                                    'keyCtr += 1
                                    Exit For
                                End If
                                'If _
                                '    dt.Rows(ctr).Item(0).ToString().Trim().ToUpper().StartsWith("SELECT") _
                                '    Or _
                                '    dt.Rows(ctr).Item(0).ToString().Trim().ToUpper().StartsWith("INSERT") _
                                '    Or _
                                '    dt.Rows(ctr).Item(0).ToString().Trim().ToUpper().StartsWith("UPDATE") _
                                '    Or _
                                '    dt.Rows(ctr).Item(0).ToString().Trim().ToUpper().StartsWith("DELETE") _
                                'Then

                                'End If

                                If dt.Rows(ctr).Item(0).ToString().Trim().ToUpper().StartsWith("--") Then
                                    'ignore
                                ElseIf dt.Rows(ctr).Item(0).ToString().Trim().ToUpper().StartsWith("SET") Then
                                    'ignore
                                ElseIf String.IsNullOrEmpty(dt.Rows(ctr).Item(0).ToString().Trim()) Then
                                    'ignore
                                Else
                                    'sb.AppendLine(dt.Rows(ctr).Item(0).ToString())
                                    sb.Append(dt.Rows(ctr).Item(0).ToString())
                                End If


                            End If
                        End If
                    Next
                    Return sb.ToString()
                    'For Each dtRow In dt.Rows
                    '    'Debug.Print(dtRow.Item(0).ToString())
                    'Next
                End If
            Else
                Return ""
            End If
            Return ""
        End Function

        Private Function DB_Connect() As String
            Dim settings As DatabaseSettings = New DatabaseSettings()
            'Dim db As Microsoft.Practices.EnterpriseLibrary.Data.Database
            Dim db As Database

            SetLastError(Nothing)

            '_databaseName = databaseName
            Try
                '_db = DatabaseFactory.CreateDatabase(databaseName)
                'db = New OleDbDataAdapter()
                db = New OracleDatabase("")
                db = New SqlDatabase("")
                'db = New GenericDatabase("",New )
                db = New GenericDatabase("", OleDbFactory.Instance)
            Catch ex As Exception
                'MsgBox(ex.Message)
                SetLastError(ex)
                GenUtils.Message(GenUtils.MsgType.Critical, "EntLib - DAAB-ConnectToDB", ex.Message)
            End Try

            Return ""

        End Function

        Public Function Connect _
            ( _
                ByVal vnDBType As e_DB, _
                ByVal vnConnType As e_ConnType, _
                ByRef ConnOdbc As System.Data.Odbc.OdbcConnection, _
                ByRef ConnOleDb As System.Data.OleDb.OleDbConnection, _
                ByRef ConnSQLServer As System.Data.SqlClient.SqlConnection, _
                ByRef ConnSQLServerCe As System.Data.SqlServerCe.SqlCeConnection, _
                ByRef ConnADO As ADODB.Connection, _
                ByVal ParamArray pArray() As Object _
            ) As Integer

            Dim connStr As String = ""
            'Dim ConnOleDb As System.Data.OleDb.OleDbConnection
            'Dim ConnSQLServer As System.Data.SqlClient.SqlConnection
            'Dim ConnADO As ADODB.Connection

            'UseADO = False
            'DBType = vnDBType
            '_isActive = False

            SetLastError(Nothing)

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
                        Case Else
                            connStr = pArray(0).ToString()
                    End Select

                    'If Not useADO Then
                    Try
                        ConnOleDb = New System.Data.OleDb.OleDbConnection(connStr)
                        ConnOleDb.Open()
                        '_isActive = True
                    Catch ex As Exception
                        'GenUtils.Log("Error connecting to:" & vbNewLine & connStr)
                        'GenUtils.Log(ex.Message)
                        LastErrorNo = -1
                        LastErrorDesc = ex.Message
                        GenUtils.Message(GenUtils.MsgType.Critical, "EntLib - DAAB-Connect", "Error connecting to:" & vbNewLine & connStr)
                        GenUtils.Message(GenUtils.MsgType.Critical, "EntLib - DAAB-Connect", ex.Message)
                    End Try
                    'End If
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
                            'Server=.\SQLExpress;AttachDbFilename=c:\asd\qwe\mydbfile.mdf;Database=dbname; Trusted_Connection=Yes;
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
                        Case Else
                            'moDAL_SQL.Connect e_db_SQLSERVER, -1, "SQLNCLI10", VBA.Environ("COMPUTERNAME"), "DatabaseName"
                            connStr = pArray(0).ToString()
                    End Select
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
                        Case Else
                            connStr = pArray(0).ToString()
                    End Select

                    'If Not useADO Then
                    If vnConnType <> e_ConnType.e_connType_E Then
                        Try
                            ConnSQLServer = New System.Data.SqlClient.SqlConnection(connStr)
                            ConnSQLServer.Open()
                            '_isActive = True
                        Catch ex As Exception
                            'GenUtils.Log("Error connecting to:" & vbNewLine & connStr)
                            'GenUtils.Log(ex.Message)
                            SetLastError(ex)
                            GenUtils.Message(GenUtils.MsgType.Critical, "EntLib - DAAB-Connect", "Error connecting to:" & vbNewLine & connStr)
                            GenUtils.Message(GenUtils.MsgType.Critical, "EntLib - DAAB-Connect", ex.Message)
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
                            GenUtils.Message(GenUtils.MsgType.Critical, "EntLib - DAAB-Connect", "Error connecting to:" & vbNewLine & connStr)
                            GenUtils.Message(GenUtils.MsgType.Critical, "EntLib - DAAB-Connect", ex.Message)
                        End Try
                    End If
                    'End If

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
                                GenUtils.Message(GenUtils.MsgType.Critical, "EntLib - DAAB-Connect", "Error connecting to:" & vbNewLine & connStr)
                                GenUtils.Message(GenUtils.MsgType.Critical, "EntLib - DAAB-Connect", ex.Message)
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
                        Case Else
                            connStr = pArray(0).ToString()
                    End Select

                    'ConnOdbc = New System.Data.Odbc.OdbcConnection(connStr)
                    'ConnOdbc.Open()

                    'If Not useADO Then
                    Try
                        ConnOleDb = New System.Data.OleDb.OleDbConnection(connStr)
                        ConnOleDb.Open()
                        '_isActive = True
                    Catch ex As Exception
                        'GenUtils.Log("Error connecting to:" & vbNewLine & connStr)
                        'GenUtils.Log(ex.Message)
                        SetLastError(ex)
                        GenUtils.Message(GenUtils.MsgType.Critical, "EntLib - DAAB-Connect", "Error connecting to:" & vbNewLine & connStr)
                        GenUtils.Message(GenUtils.MsgType.Critical, "EntLib - DAAB-Connect", ex.Message)
                    End Try
                    'End If
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
                        Case Else
                            connStr = pArray(0).ToString()
                    End Select

                    'If Not useADO Then
                    Try
                        ConnOleDb = New System.Data.OleDb.OleDbConnection(connStr)
                        ConnOleDb.Open()
                        '_isActive = True
                    Catch ex As Exception
                        'GenUtils.Log("Error connecting to:" & vbNewLine & connStr)
                        'GenUtils.Log(ex.Message)
                        SetLastError(ex)
                        GenUtils.Message(GenUtils.MsgType.Critical, "EntLib - DAAB-Connect", "Error connecting to:" & vbNewLine & connStr)
                        GenUtils.Message(GenUtils.MsgType.Critical, "EntLib - DAAB-Connect", ex.Message)
                    End Try
                    'End If
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
                        Case Else
                            connStr = pArray(0).ToString()
                    End Select

                    'If Not useADO Then
                    Try
                        ConnOleDb = New System.Data.OleDb.OleDbConnection(connStr)
                        ConnOleDb.Open()
                        '_isActive = True
                    Catch ex As Exception
                        'GenUtils.Log("Error connecting to:" & vbNewLine & connStr)
                        'GenUtils.Log(ex.Message)
                        SetLastError(ex)
                        GenUtils.Message(GenUtils.MsgType.Critical, "EntLib - DAAB-Connect", "Error connecting to:" & vbNewLine & connStr)
                        GenUtils.Message(GenUtils.MsgType.Critical, "EntLib - DAAB-Connect", ex.Message)
                    End Try
                    'End If
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
                        Case Else
                    End Select
                Case Else
            End Select

            Return 0
        End Function

        ''' <summary>
        ''' RKP/07-11-11/v2.5.147
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function ConnectToSysDB() As Integer

            Return 0
        End Function

        'RKP/07-25-11/v2.5.147
        Public Function IsSQLExpress() As Boolean

            'If _connSQLExpress IsNot Nothing Then
            '    Return True
            'Else
            '    Return False
            'End If

            If _db.DbProviderFactory.ToString.Contains("System.Data.SqlClient") Then
                Return True
            Else
                Return False
            End If

        End Function

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

        Public Function DB_CreateDatabase _
            ( _
                ByVal targetDBType As e_DB, _
                ByRef targetDBName As String, _
                ByRef targetDBFolderPath As String _
            ) As Integer
            'RKP/08-16-11/v3.0.150
            'http://www.withvb.net/tutorials/sql-databases-create-update-and-query

            Select Case targetDBType
                Case e_DB.e_db_ACCESS
                Case e_DB.e_db_SQLSERVER

                Case Else
            End Select

        End Function

        'Public Shared Function DB_TransferObjects( _
        '    ByRef switches() As String, _
        '    ByVal srcDBName As String, _
        '    ByVal srcDBFolder As String, _
        '    ByVal destDBName As String, _
        '    ByVal destDBFolder As String _
        ') As Integer

        '    Dim dalOleDb As New clsDAL
        '    Dim dalSql As New clsDAL
        '    Dim connOleDb As New OleDb.OleDbConnection
        '    Dim connSql As New SqlClient.SqlConnection
        '    Dim adoCatalog As New ADOX.Catalog
        '    Dim adoTables As ADOX.Tables
        '    Dim adoTable As ADOX.Table
        '    Dim adoIdxs As ADOX.Indexes
        '    Dim adoIdx As ADOX.Index
        '    'Dim adoIdxCols As ADOX.Columns
        '    'Dim adoIdxCol As ADOX.Column
        '    Dim adoCols As ADOX.Columns
        '    Dim adoCol As ADOX.Column
        '    Dim adoProperties As ADOX.Properties
        '    'Dim adoProperty As ADOX.Property
        '    Dim sql As String = ""
        '    Dim dt As New DataTable
        '    Dim dt1 As New DataTable
        '    Dim filePath As String = ""
        '    Dim listPK() As String
        '    Dim listIdx() As String
        '    Dim ctr As Integer = 0
        '    Dim ctr1 As Integer = 0
        '    Dim ctr2 As Integer = 0
        '    Dim tableDefs() As STRUCT_TABLEDEF
        '    'Dim idxDefs() As STRUCT_IDX
        '    Dim createObject As Boolean = False
        '    Dim proceed As Boolean = False
        '    Dim input As String
        '    Dim srcFilePath As String = ""
        '    Dim destFilePath As String = ""
        '    Dim msg As String = ""
        '    Dim noTruncate As Boolean = GenUtils.IsSwitchAvailable(switches, "/UseTruncateSP")

        '    'Dim srcDBName As String = "IPG50"
        '    'Dim srcDBFolder As String = "C:\OPTMODELS\IPG50"
        '    'Dim destDBName As String = "IPG50SQL"
        '    'Dim destDBFolder As String = "C:\OPTMODELS\IPG50SQL"

        '    If String.IsNullOrEmpty(srcDBFolder) Then
        '        srcDBFolder = "C:\OPTMODELS\" & srcDBName
        '    End If

        '    If String.IsNullOrEmpty(destDBFolder) Then
        '        destDBFolder = "C:\OPTMODELS\" & destDBName
        '    End If

        '    filePath = srcDBFolder & "\" & srcDBName & ".mdb"  '"C:\OPTMODELS\SRC1\SRC1.mdb"
        '    srcFilePath = filePath
        '    If My.Computer.FileSystem.FileExists(filePath) Then
        '        proceed = True
        '    Else
        '        proceed = False
        '        msg = msg & vbNewLine & "Source database not found."
        '    End If
        '    If proceed Then
        '        filePath = destDBFolder & "\" & destDBName & ".mdf"
        '        destFilePath = filePath
        '        If My.Computer.FileSystem.FileExists(filePath) Then
        '            proceed = True
        '        Else
        '            proceed = False
        '            msg = msg & vbNewLine & "Destination database not found."
        '        End If
        '    End If

        '    If proceed Then 'proceed #1

        '        If EntLib.COPT.GenUtils.IsSwitchAvailable(switches, "/NoPrompt") Then
        '            proceed = True
        '        Else
        '            Console.WriteLine()
        '            Console.WriteLine("Source File Path: " & srcFilePath)
        '            Console.WriteLine("Destination File Path: " & destFilePath)
        '            Console.Write("Do you want to continue? (Y/N): ")
        '            input = Console.ReadLine()
        '            If input.Trim.ToUpper.Equals("Y") Then
        '                proceed = True
        '            Else
        '                proceed = False
        '            End If
        '        End If

        '        If proceed Then 'proceed #2

        '            EntLib.COPT.Log.Log("********************")
        '            EntLib.COPT.Log.Log("TRANSFER - Started at: " & Now())

        '            EntLib.COPT.Log.Log("Source File Path: " & srcFilePath)
        '            EntLib.COPT.Log.Log("Destination File Path: " & destFilePath)

        '            filePath = srcDBFolder & "\" & srcDBName & ".mdb"  '"C:\OPTMODELS\SRC1\SRC1.mdb"
        '            dalOleDb.UseADO = True
        '            Try
        '                dalOleDb.Connect(clsDAL.e_DB.e_db_ACCESS, clsDAL.e_ConnType.e_connType_A, Nothing, connOleDb, Nothing, Nothing, Nothing, True, filePath)
        '                EntLib.COPT.Log.Log("Connected to source database successfully.")
        '            Catch ex As Exception
        '                EntLib.COPT.Log.Log("Error connecting to source database.")
        '                EntLib.COPT.Log.Log(ex.Message)
        '            End Try

        '            adoCatalog = dalOleDb.ConnADOCatalog
        '            adoTables = adoCatalog.Tables

        '            'dalSql.Connect(clsDAL.e_DB.e_db_SQLSERVER, clsDAL.e_ConnType.e_connType_F, Nothing, Nothing, connSql, Nothing, Nothing, "S02ASQLNPD01", "BMOS")
        '            'Works:
        '            'dalSql.Connect(clsDAL.e_DB.e_db_SQLSERVER, clsDAL.e_ConnType.e_connType_I, Nothing, Nothing, connSql, Nothing, Nothing, ".\SQLEXPRESS", "C:\OPTMODELS\SRC1SQL\SRC1SQL.mdf", "SRC1SQL")
        '            EntLib.COPT.Log.Log("")
        '            Try
        '                dalSql.Connect(clsDAL.e_DB.e_db_SQLSERVER, clsDAL.e_ConnType.e_connType_I, Nothing, Nothing, connSql, Nothing, Nothing, True, ".\SQLEXPRESS", destDBFolder & "\" & destDBName & ".mdf", destDBName)
        '                EntLib.COPT.Log.Log("Connected to destination database successfully.")
        '            Catch ex As Exception
        '                EntLib.COPT.Log.Log("Error connecting to destination database.")
        '                EntLib.COPT.Log.Log(ex.Message)
        '            End Try


        '            'sql = "SELECT * FROM tsysCOL"
        '            'dalSql.Execute(sql, dt)

        '            sql = "SELECT * FROM tsysCoreObjects WHERE OBJ_TYPE_ID = 1 AND [Active] = True"
        '            EntLib.COPT.Log.Log(sql)
        '            Try
        '                dalOleDb.Execute(sql, dt)
        '            Catch ex As Exception
        '                EntLib.COPT.Log.Log("Error:")
        '                EntLib.COPT.Log.Log(ex.Message)
        '            End Try

        '            'sql = "TRUNCATE TABLE dbo.ZCOR1347_ENT"
        '            'dalSql.Execute(sql, dt)

        '            'sql = "BULKCOPY"
        '            'dalSql.Execute(sql, dt, "dbo.ZCOR1347_ENT")

        '            For ctr = 0 To dt.Rows.Count - 1
        '                For Each adoTable In adoTables
        '                    If dt.Rows(ctr).Item("OBJ_NAME").ToString().Equals(adoTable.Name.ToString()) Then
        '                        Console.Write("Reading source table: " & adoTable.Name.ToString())
        '                        EntLib.COPT.Log.Log("Reading source table: " & adoTable.Name.ToString())
        '                        ReDim tableDefs(0)
        '                        adoCols = adoTable.Columns
        '                        For Each adoCol In adoCols

        '                            adoProperties = adoCol.Properties

        '                            'For Each adoProperty In adoProperties
        '                            '    Try
        '                            '        If String.IsNullOrEmpty(adoProperty.Value.ToString()) Then
        '                            '            Debug.Print(adoCol.Name.ToString() & " - " & adoProperty.Name & " - " & "")
        '                            '        Else
        '                            '            Debug.Print(adoCol.Name.ToString() & " - " & adoProperty.Name & " - " & adoProperty.Value.ToString())
        '                            '        End If
        '                            '    Catch ex As Exception
        '                            '        Debug.Print(adoCol.Name.ToString() & " - " & adoProperty.Name & " - " & "")
        '                            '    End Try
        '                            'Next
        '                            'Debug.Print(adoCol.DefinedSize.ToString())

        '                            If String.IsNullOrEmpty(tableDefs(0).fieldName) Then
        '                                tableDefs(0).fieldName = adoCol.Name
        '                                tableDefs(0).fieldType = adoCol.Type.ToString()
        '                                tableDefs(0).fieldSize = 0
        '                            Else
        '                                ReDim Preserve tableDefs(UBound(tableDefs) + 1)
        '                                tableDefs(UBound(tableDefs)).fieldName = adoCol.Name
        '                                tableDefs(UBound(tableDefs)).fieldType = adoCol.Type.ToString()
        '                                tableDefs(UBound(tableDefs)).fieldSize = 0
        '                            End If
        '                            If tableDefs(UBound(tableDefs)).fieldType.Contains("Char") Then
        '                                tableDefs(UBound(tableDefs)).fieldSize = CInt(adoCol.DefinedSize.ToString())
        '                            End If
        '                            tableDefs(UBound(tableDefs)).fieldTypeNew = clsDAL.GetADOToSQLDataType(tableDefs(UBound(tableDefs)).fieldType)
        '                            tableDefs(UBound(tableDefs)).fieldIsIdentity = CBool(adoProperties("Autoincrement").Value.ToString())
        '                            If tableDefs(UBound(tableDefs)).fieldIsIdentity Then
        '                                tableDefs(UBound(tableDefs)).fieldIsNullable = False
        '                            Else
        '                                tableDefs(UBound(tableDefs)).fieldIsNullable = True
        '                            End If
        '                            'tableDefs(UBound(tableDefs)).fieldIsNullable = CBool(adoProperties("Nullable").Value.ToString())
        '                            'If tableDefs(UBound(tableDefs)).fieldIsIdentity Then
        '                            ' tableDefs(UBound(tableDefs)).fieldIsNullable = False
        '                            'Else
        '                            'tableDefs(UBound(tableDefs)).fieldIsNullable = True
        '                            'End If
        '                        Next

        '                        adoIdxs = adoTable.Indexes
        '                        ReDim listPK(0)
        '                        ReDim listIdx(0)
        '                        For Each adoIdx In adoIdxs
        '                            'Debug.Print(adoTable.Name.ToString() & " - " & adoIdx.Name.ToString())
        '                            'adoIdx.Columns.c
        '                            'adoProperties = adoIdx.Properties
        '                            'For Each adoProperty In adoProperties
        '                            '    Debug.Print(adoProperty.Name & " - " & adoProperty.Value.ToString())
        '                            'Next
        '                            adoCols = adoIdx.Columns
        '                            For Each adoCol In adoCols
        '                                'Debug.Print("TableName=" & adoTable.Name.ToString() & ", PK=" & adoIdx.PrimaryKey.ToString() & ", IdxName=" & adoIdx.Name.ToString() & ", ColName=" & adoCol.Name.ToString())
        '                                If adoIdx.PrimaryKey Then
        '                                    If String.IsNullOrEmpty(listPK(0)) Then
        '                                        listPK(0) = adoCol.Name.ToString()
        '                                    Else
        '                                        ReDim Preserve listPK(UBound(listPK) + 1)
        '                                        listPK(UBound(listPK)) = adoCol.Name.ToString()
        '                                    End If
        '                                Else
        '                                    If String.IsNullOrEmpty(listIdx(0)) Then
        '                                        listIdx(0) = adoCol.Name.ToString()
        '                                    Else
        '                                        ReDim Preserve listIdx(UBound(listIdx) + 1)
        '                                        listIdx(UBound(listIdx)) = adoCol.Name.ToString()
        '                                    End If
        '                                End If
        '                            Next
        '                        Next

        '                        For ctr1 = 0 To listPK.Length - 1
        '                            For ctr2 = 0 To tableDefs.Length - 1
        '                                If tableDefs(ctr2).fieldName.Equals(listPK(ctr1)) Then
        '                                    tableDefs(ctr2).fieldIsNullable = False
        '                                    Exit For
        '                                End If
        '                            Next
        '                        Next

        '                        Console.WriteLine("...Complete.")
        '                        EntLib.COPT.Log.Log("...Complete.")

        '                        Console.Write("Creating destination table: " & adoTable.Name.ToString())
        '                        EntLib.COPT.Log.Log("Creating destination table: " & adoTable.Name.ToString())

        '                        'Create table + indexes in the SQL Express database - START
        '                        If CBool(dt.Rows(ctr).Item("Recreate").ToString()) Then
        '                            sql = "IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[" & adoTable.Name.ToString() & "]') AND type in (N'U')) DROP TABLE [dbo].[" & adoTable.Name.ToString() & "]"
        '                            Try
        '                                dalSql.Execute(sql, dt1)
        '                                createObject = True
        '                            Catch ex As Exception
        '                                Console.WriteLine("Error creating table+indexes in SQLEXPRESS database.")
        '                                Console.WriteLine(ex.Message)
        '                                EntLib.COPT.Log.Log("Error creating table+indexes in SQLEXPRESS database.")
        '                                EntLib.COPT.Log.Log(ex.Message)
        '                                createObject = False
        '                            End Try

        '                        Else
        '                            sql = "SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[" & adoTable.Name.ToString() & "]') AND type in (N'U')"
        '                            'EntLib.COPT.Log.Log(ex.Message)
        '                            Try
        '                                dalSql.Execute(sql, dt1)
        '                            Catch ex As Exception
        '                                Console.WriteLine("Error executing SQL.")
        '                                Console.WriteLine(sql)
        '                                Console.WriteLine(ex.Message)
        '                                EntLib.COPT.Log.Log("Error executing SQL.")
        '                                EntLib.COPT.Log.Log(sql)
        '                                EntLib.COPT.Log.Log(ex.Message)
        '                            End Try

        '                            If dt1 Is Nothing Then
        '                                createObject = True
        '                            Else
        '                                If dt1.Rows.Count > 0 Then
        '                                    createObject = False
        '                                Else
        '                                    createObject = True
        '                                End If
        '                            End If
        '                        End If
        '                        If createObject Then
        '                            'Step 1: Drop table in destination database, if it already exists
        '                            sql = "IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[" & adoTable.Name.ToString() & "]') AND type in (N'U')) DROP TABLE [dbo].[" & adoTable.Name.ToString() & "]"
        '                            Try
        '                                dalSql.Execute(sql, dt1)
        '                            Catch ex As Exception
        '                                Console.WriteLine("Error executing SQL.")
        '                                Console.WriteLine(sql)
        '                                Console.WriteLine(ex.Message)
        '                                EntLib.COPT.Log.Log("Error executing SQL.")
        '                                EntLib.COPT.Log.Log(sql)
        '                                EntLib.COPT.Log.Log(ex.Message)
        '                            End Try

        '                            'Step 2: Create table in destination database + create primary key
        '                            sql = "SET ANSI_NULLS ON"
        '                            Try
        '                                dalSql.Execute(sql, dt1)
        '                            Catch ex As Exception
        '                                Console.WriteLine("Error executing SQL.")
        '                                Console.WriteLine(sql)
        '                                Console.WriteLine(ex.Message)
        '                                EntLib.COPT.Log.Log("Error executing SQL.")
        '                                EntLib.COPT.Log.Log(sql)
        '                                EntLib.COPT.Log.Log(ex.Message)
        '                            End Try

        '                            sql = "SET QUOTED_IDENTIFIER ON"
        '                            Try
        '                                dalSql.Execute(sql, dt1)
        '                            Catch ex As Exception
        '                                Console.WriteLine("Error executing SQL.")
        '                                Console.WriteLine(sql)
        '                                Console.WriteLine(ex.Message)
        '                                EntLib.COPT.Log.Log("Error executing SQL.")
        '                                EntLib.COPT.Log.Log(sql)
        '                                EntLib.COPT.Log.Log(ex.Message)
        '                            End Try

        '                            sql = "CREATE TABLE [dbo].[" & adoTable.Name.ToString() & "] ("
        '                            For ctr1 = 0 To tableDefs.Length - 1
        '                                sql = sql & "[" & tableDefs(ctr1).fieldName & "] [" & tableDefs(ctr1).fieldTypeNew & "]" & IIf(tableDefs(ctr1).fieldSize > 0, "(" & tableDefs(ctr1).fieldSize & ") ", "").ToString() & IIf(tableDefs(ctr1).fieldIsIdentity, " IDENTITY(1,1)", "").ToString() & IIf(tableDefs(ctr1).fieldIsNullable, " NULL", " NOT NULL").ToString() & ","
        '                            Next
        '                            If String.IsNullOrEmpty(listPK(0)) Then
        '                                sql = Left(sql, Len(sql) - 1) & vbNewLine
        '                            Else
        '                                sql = sql & " CONSTRAINT [PK_" & adoTable.Name.ToString() & "] PRIMARY KEY CLUSTERED " & vbNewLine
        '                                sql = sql & "(" & vbNewLine
        '                                For ctr1 = 0 To listPK.Length - 1
        '                                    sql = sql & "[" & listPK(ctr1) & "] ASC,"
        '                                Next
        '                                sql = Left(sql, Len(sql) - 1) & vbNewLine
        '                                sql = sql & ")WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]" & vbNewLine
        '                            End If
        '                            sql = sql & ") ON [PRIMARY]" & vbNewLine
        '                            Try
        '                                dalSql.Execute(sql, dt1)
        '                            Catch ex As Exception
        '                                Debug.Print("Error creating TABLE - " & adoTable.Name.ToString())
        '                                Debug.Print(ex.Message)
        '                                Debug.Print(sql)
        '                                Debug.Print("")
        '                                Console.WriteLine("Error executing SQL.")
        '                                Console.WriteLine(sql)
        '                                Console.WriteLine(ex.Message)
        '                                EntLib.COPT.Log.Log("Error executing SQL.")
        '                                EntLib.COPT.Log.Log(sql)
        '                                EntLib.COPT.Log.Log(ex.Message)
        '                            End Try
        '                            'Step 3: Create nonclustered index on the table
        '                            If Not String.IsNullOrEmpty(listIdx(0)) Then
        '                                sql = "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[" & adoTable.Name.ToString() & "]') AND name = N'" & "IDX_" & adoTable.Name.ToString() & "')"
        '                                sql = sql & "CREATE UNIQUE NONCLUSTERED INDEX [" & "IDX_" & adoTable.Name.ToString() & "] ON [dbo].[" & adoTable.Name.ToString() & "] "
        '                                sql = sql & "( "
        '                                For ctr1 = 0 To listIdx.Length - 1
        '                                    sql = sql & "[" & listIdx(ctr1) & "] ASC,"
        '                                Next
        '                                sql = Left(sql, Len(sql) - 1) & vbNewLine
        '                                sql = sql & ")WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY] "
        '                                Try
        '                                    dalSql.Execute(sql, dt1)
        '                                Catch ex As Exception
        '                                    Debug.Print("Error creating NONCLUSTERED INDEX on TABLE - " & adoTable.Name.ToString())
        '                                    Debug.Print(ex.Message)
        '                                    Debug.Print(sql)
        '                                    Debug.Print("")
        '                                    Console.WriteLine("Error executing SQL.")
        '                                    Console.WriteLine(sql)
        '                                    Console.WriteLine(ex.Message)
        '                                    EntLib.COPT.Log.Log("Error executing SQL.")
        '                                    EntLib.COPT.Log.Log(sql)
        '                                    EntLib.COPT.Log.Log(ex.Message)
        '                                End Try
        '                            End If
        '                        End If
        '                        Console.Write("...Created")
        '                        'Create table + indexes in the SQL Express database - END

        '                        'Copy data - START
        '                        If CBool(dt.Rows(ctr).Item("TransferData").ToString()) Then
        '                            'Empty destination table
        '                            If noTruncate Then
        '                                'sql = "DELETE FROM [" & adoTable.Name.ToString() & "]"
        '                                sql = "EXEC dbo.asp_TruncateTable '" & adoTable.Name.ToString() & "'"
        '                            Else
        '                                sql = "TRUNCATE TABLE [" & adoTable.Name.ToString() & "]"
        '                            End If

        '                            Try
        '                                dalSql.Execute(sql, dt1)
        '                            Catch ex As Exception
        '                                Console.WriteLine("Error executing SQL.")
        '                                Console.WriteLine(sql)
        '                                Console.WriteLine(ex.Message)
        '                                EntLib.COPT.Log.Log("Error executing SQL.")
        '                                EntLib.COPT.Log.Log(sql)
        '                                EntLib.COPT.Log.Log(ex.Message)
        '                            End Try

        '                            sql = "SELECT * FROM [" & adoTable.Name.ToString() & "]"
        '                            Try
        '                                dalOleDb.Execute(sql, dt1)
        '                            Catch ex As Exception
        '                                Console.WriteLine("Error executing SQL.")
        '                                Console.WriteLine(sql)
        '                                Console.WriteLine(ex.Message)
        '                                EntLib.COPT.Log.Log("Error executing SQL.")
        '                                EntLib.COPT.Log.Log(sql)
        '                                EntLib.COPT.Log.Log(ex.Message)
        '                            End Try

        '                            sql = "dbo.[" & adoTable.Name.ToString() & "]"
        '                            Try
        '                                dalSql.Execute("BULKCOPY-MAP", dt1, sql)
        '                            Catch ex As Exception
        '                                Console.WriteLine("Error executing SQL.")
        '                                Console.WriteLine(sql)
        '                                Console.WriteLine(ex.Message)
        '                                EntLib.COPT.Log.Log("Error executing SQL.")
        '                                EntLib.COPT.Log.Log(sql)
        '                                EntLib.COPT.Log.Log(ex.Message)
        '                            End Try

        '                            Console.Write("...Populated with data.")
        '                        End If
        '                        'Copy data - END
        '                        Console.WriteLine("")

        '                        Exit For
        '                    End If
        '                Next 'For Each adoTable In adoTables
        '            Next 'For ctr = 0 To dt.Rows.Count - 1

        '            Console.WriteLine("All TRANSFER tasks finished successfully.")
        '            EntLib.COPT.Log.Log("All TRANSFER tasks finished successfully.")
        '        End If 'If proceed Then 'proceed #2
        '    End If 'If proceed Then 'proceed #1

        '    dalOleDb = Nothing
        '    dalSql = Nothing
        '    dt = Nothing

        '    EntLib.COPT.Log.Log("TRANSFER - Ended at: " & Now())
        '    EntLib.COPT.Log.Log("********************")

        '    Return 0
        'End Function

        ' TRANSFER (TRF) - Transfers objects and data from source (MS Access) database to destination (SQL Server Express) database, as defined in the system table, tsysCoreObjects (in the source database).

        'Public Shared Function DB_TransferADO( _
        '    ByRef switches() As String, _
        '    ByVal srcDBName As String, _
        '    ByVal srcDBFolder As String, _
        '    ByVal destDBName As String, _
        '    ByVal destDBFolder As String _
        ') As Integer

        '    Dim dalOleDbADO As New clsDAL 'ADO version 'RKP/04-20-12/v3.2.166
        '    Dim dalOleDb As New clsDAL 'ADO.NET version 'RKP/04-20-12/v3.2.166
        '    Dim dalSql As New clsDAL
        '    Dim connOleDb As New OleDb.OleDbConnection
        '    Dim connSql As New SqlClient.SqlConnection
        '    Dim adoCatalog As New ADOX.Catalog
        '    Dim adoTables As ADOX.Tables
        '    Dim adoTable As ADOX.Table
        '    Dim adoIdxs As ADOX.Indexes
        '    Dim adoIdx As ADOX.Index
        '    Dim adoCols As ADOX.Columns
        '    Dim adoCol As ADOX.Column
        '    Dim adoProperties As ADOX.Properties
        '    Dim adoRSSchema As ADODB.Recordset
        '    Dim sql As String = ""
        '    Dim dt As New DataTable
        '    Dim dt1 As New DataTable
        '    Dim filePath As String = ""
        '    Dim listPK() As String
        '    Dim listIdx() As String
        '    Dim ctr As Integer = 0
        '    Dim ctr1 As Integer = 0
        '    Dim ctr2 As Integer = 0
        '    Dim tableDefs() As STRUCT_TABLEDEF
        '    Dim idxDefs() As STRUCT_IDX
        '    Dim createObject As Boolean = False
        '    Dim proceed As Boolean = False
        '    'Dim input As String
        '    Dim srcFilePath As String = ""
        '    Dim destFilePath As String = ""
        '    Dim msg As String = ""
        '    Dim colCtr As Integer = 0
        '    Dim idxComposite As String = ""
        '    Dim adoTableCtr As Integer = 0
        '    'Dim dtADOTables As DataTable
        '    Dim ret As Integer
        '    Dim startTime As Long = My.Computer.Clock.TickCount
        '    Dim errCount As Integer = 0
        '    Dim noTruncate As Boolean = GenUtils.IsSwitchAvailable(switches, "/UseTruncateSP")

        '    'Dim input As String
        '    'Dim arrayGetRows()() As String
        '    'Dim ret As Integer
        '    'Dim arrayRestrictions() As String
        '    'Dim list As New List(Of String)
        '    'Dim srcDBName As String = "IPG50"
        '    'Dim srcDBFolder As String = "C:\OPTMODELS\IPG50"
        '    'Dim destDBName As String = "IPG50SQL"
        '    'Dim destDBFolder As String = "C:\OPTMODELS\IPG50SQL"

        '    'MessageBox.Show("Entered DB_TransferObjects2.", "C-OPT", MessageBoxButtons.OK, MessageBoxIcon.Information)

        '    'Console.Write("Do you want to continue? (Y/N): ")
        '    'input = Console.ReadLine()
        '    'If input.Trim.ToUpper.Equals("Y") Then
        '    '    proceed = True
        '    'Else
        '    '    proceed = False
        '    'End If

        '    'Console.ReadKey()

        '    'If String.IsNullOrEmpty(srcDBFolder) Then
        '    '    srcDBFolder = "C:\OPTMODELS\" & srcDBName
        '    'End If

        '    'If String.IsNullOrEmpty(destDBFolder) Then
        '    '    destDBFolder = "C:\OPTMODELS\" & destDBName
        '    'End If

        '    'filePath = srcDBFolder & "\" & srcDBName & IIf(srcDBName.Trim().ToUpper().EndsWith(".MDB"), "", ".mdb").ToString() '& ".mdb"  '"C:\OPTMODELS\SRC1\SRC1.mdb"
        '    'srcFilePath = filePath
        '    'If My.Computer.FileSystem.FileExists(filePath) Then
        '    '    proceed = True
        '    'Else
        '    '    proceed = False
        '    '    msg = msg & vbNewLine & "Source database not found."
        '    'End If
        '    'If proceed Then
        '    '    filePath = destDBFolder & "\" & destDBName & IIf(destDBName.Trim().ToUpper().EndsWith(".MDF"), "", ".mdf").ToString() '& ".mdf"
        '    '    destFilePath = filePath
        '    '    If My.Computer.FileSystem.FileExists(filePath) Then
        '    '        proceed = True
        '    '    Else
        '    '        proceed = False
        '    '        msg = msg & vbNewLine & "Destination database not found."
        '    '    End If
        '    'End If


        '    startTime = My.Computer.Clock.TickCount

        '    proceed = True
        '    If proceed Then 'proceed #1

        '        'If EntLib.COPT.GenUtils.IsSwitchAvailable(switches, "/NoPrompt") Then
        '        '    proceed = True
        '        'Else
        '        '    Console.WriteLine()
        '        '    Console.WriteLine("Source File Path: " & srcFilePath)
        '        '    Console.WriteLine("Destination File Path: " & destFilePath)
        '        '    Console.Write("Do you want to continue? (Y/N): ")
        '        '    input = Console.ReadLine()
        '        '    If input.Trim.ToUpper.Equals("Y") Then
        '        '        proceed = True
        '        '    Else
        '        '        proceed = False
        '        '    End If
        '        'End If

        '        If proceed Then 'proceed #2

        '            'MessageBox.Show("About to create log entry.", "C-OPT", MessageBoxButtons.OK, MessageBoxIcon.Information)

        '            'Console.Write("Do you want to continue? (Y/N): ")
        '            'Input = Console.ReadLine()
        '            'If Input.Trim.ToUpper.Equals("Y") Then
        '            '    proceed = True
        '            'Else
        '            '    proceed = False
        '            'End If

        '            EntLib.COPT.Log.Log("")
        '            EntLib.COPT.Log.Log("********************")
        '            EntLib.COPT.Log.Log("TRANSFER (TRF) - Started at: " & Now())

        '            EntLib.COPT.Log.Log("Source File Path: " & srcFilePath)
        '            EntLib.COPT.Log.Log("Destination File Path: " & destFilePath)

        '            filePath = srcDBFolder & "\" & srcDBName & IIf(srcDBName.Trim().ToUpper().EndsWith(".MDB"), "", ".mdb").ToString() '& ".mdb"  '"C:\OPTMODELS\SRC1\SRC1.mdb"
        '            dalOleDbADO.UseADO = True
        '            Try
        '                dalOleDbADO.Connect(clsDAL.e_DB.e_db_ACCESS, clsDAL.e_ConnType.e_connType_C, Nothing, connOleDb, Nothing, Nothing, Nothing, True, filePath, "C:\OPTMODELS\C-OPTSYS\System.mdw.copt")
        '                EntLib.COPT.Log.Log("Connected to source database (using ADO) successfully.")
        '                Console.WriteLine("Connected to source database (using ADO) successfully.")
        '            Catch ex As Exception
        '                Console.WriteLine("Error connecting to source database (using ADO).")
        '                EntLib.COPT.Log.Log("Error connecting to source database (using ADO).")
        '                EntLib.COPT.Log.Log(ex.Message)
        '                errCount += 1
        '            End Try

        '            'RKP/04-20-12/v3.2.166
        '            'This is the ADO.NET version of the source connection object.
        '            'The purpose of this object is to avoid the use of "ConvertADORecordsetToDataTable", which avoids duplicate in-memory objects (adoRS and dt).
        '            'If not for this object, memory utilization would be adversely affected, especially for large tables (~ million rows).
        '            dalOleDb.UseADO = False
        '            Try
        '                dalOleDb.Connect(clsDAL.e_DB.e_db_ACCESS, clsDAL.e_ConnType.e_connType_C, Nothing, connOleDb, Nothing, Nothing, Nothing, True, filePath, "C:\OPTMODELS\C-OPTSYS\System.mdw.copt")
        '                EntLib.COPT.Log.Log("Connected to source database successfully (using ADO.NET).")
        '                Console.WriteLine("Connected to source database successfully (using ADO.NET).")
        '            Catch ex As Exception
        '                Console.WriteLine("Error connecting to source database (using ADO.NET).")
        '                EntLib.COPT.Log.Log("Error connecting to source database (using ADO.NET).")
        '                EntLib.COPT.Log.Log(ex.Message)
        '                errCount += 1
        '            End Try


        '            Console.WriteLine("Reading catalog from source database...")
        '            Console.Write("Started at: " & TimeSpan.FromMilliseconds(startTime).TotalMinutes & " min. ")

        '            'ret = dalOleDb.Execute("SELECT * FROM qsysMSysObjects WHERE OBJ_TYPE_ID = 1", dtADOTables)

        '            'RKP/02-12-12/v3.0.157
        '            'Added additional error-checking to prevent C-OPT from crashing at this point.
        '            Try
        '                adoCatalog = dalOleDbADO.ConnADOCatalog
        '                adoTables = adoCatalog.Tables
        '                ret = adoTables.Count
        '            Catch ex As Exception
        '                Console.WriteLine("Error.")
        '                Console.WriteLine("adoTables.Count = 0.")
        '                Console.WriteLine("Please reach out to BMOS for resolution.")
        '                Console.WriteLine("The source database has objects in it (like, linked tables, etc) that are incompatible with C-OPT.")
        '                Console.WriteLine("You must remove all incompatible objects in the source database (like linked tables to external databases, invalid queries, etc) and try the operation again.")
        '                EntLib.COPT.Log.Log("Error. adoTables.Count = 0.")
        '                EntLib.COPT.Log.Log("The source database has objects in it (like, linked tables, etc) that are incompatible with C-OPT.")
        '                EntLib.COPT.Log.Log("If, for example, the database has linked tables to, say, a MySQL database, C-OPT will have difficulty connecting to it.")
        '                EntLib.COPT.Log.Log("You must remove all incompatible objects in the source database (like linked tables to external databases, invalid queries, etc) and try the operation again.")
        '                EntLib.COPT.Log.Log(ex.Message)
        '                errCount += 1
        '                Return -1
        '            End Try

        '            If ret = 0 Then
        '                Console.WriteLine("Error.")
        '                Console.WriteLine("adoTables.Count = 0.")
        '                Console.WriteLine("Please reach out to BMOS for resolution.")
        '                EntLib.COPT.Log.Log("Error. adoTables.Count = 0.")
        '                errCount += 1
        '                Return -1
        '            Else
        '                Console.WriteLine("Ended at: " & Now().ToString() & " (" & GenUtils.FormatTime(startTime) & ".)")
        '                EntLib.COPT.Log.Log("Reading Catalog took: " & GenUtils.FormatTime(startTime) & ".")
        '            End If



        '            'Console.WriteLine(adoTables.Count.ToString() & " tables & views found in database.")

        '            'http://support.microsoft.com/kb/299484
        '            'adoRSSchema = dalOleDb.ConnADO.OpenSchema(ADODB.SchemaEnum.adSchemaColumns, Array(Nothing, Nothing, "Products"))

        '            'dalSql.Connect(clsDAL.e_DB.e_db_SQLSERVER, clsDAL.e_ConnType.e_connType_F, Nothing, Nothing, connSql, Nothing, Nothing, "S02ASQLNPD01", "BMOS")
        '            'Works:
        '            'dalSql.Connect(clsDAL.e_DB.e_db_SQLSERVER, clsDAL.e_ConnType.e_connType_I, Nothing, Nothing, connSql, Nothing, Nothing, ".\SQLEXPRESS", "C:\OPTMODELS\SRC1SQL\SRC1SQL.mdf", "SRC1SQL")
        '            'EntLib.COPT.Log.Log("")
        '            Try
        '                dalSql.Connect( _
        '                        clsDAL.e_DB.e_db_SQLSERVER, _
        '                        clsDAL.e_ConnType.e_connType_I, _
        '                        Nothing, _
        '                        Nothing, _
        '                        connSql, _
        '                        Nothing, _
        '                        Nothing, _
        '                        True, _
        '                        ".\SQLEXPRESS", _
        '                        destDBFolder & "\" & destDBName & _
        '                        IIf(destDBName.Trim().ToUpper().EndsWith(".MDF"), "", ".mdf").ToString(), _
        '                        IIf(destDBName.Trim().ToUpper().EndsWith(".MDF"), _
        '                        Left(destDBName.Trim(), Len(destDBName.Trim()) - 4), destDBName.Trim()).ToString() _
        '                )
        '                EntLib.COPT.Log.Log("Connected to destination database successfully.")
        '                Console.WriteLine("Connected to destination database successfully.")
        '            Catch ex As Exception
        '                Console.WriteLine("Error connecting to destination database.")
        '                EntLib.COPT.Log.Log("Error connecting to destination database.")
        '                EntLib.COPT.Log.Log(ex.Message)
        '                errCount += 1
        '            End Try

        '            'sql = "SELECT * FROM tsysCOL"
        '            'dalSql.Execute(sql, dt)

        '            sql = "SELECT * FROM tsysCoreObjects WHERE OBJ_TYPE_ID = 1 AND [Active] = True"
        '            Console.Write("Reading list of core objects...")
        '            EntLib.COPT.Log.Log(sql)
        '            Try
        '                dalOleDbADO.Execute(sql, dt)
        '                Console.WriteLine("complete.")
        '            Catch ex As Exception
        '                Console.WriteLine("error.")
        '                Console.WriteLine("Error:")
        '                Console.WriteLine(ex.Message)
        '                EntLib.COPT.Log.Log("Error:")
        '                EntLib.COPT.Log.Log(ex.Message)
        '                errCount += 1
        '            End Try

        '            'sql = "TRUNCATE TABLE dbo.ZCOR1347_ENT"
        '            'dalSql.Execute(sql, dt)

        '            'sql = "BULKCOPY"
        '            'dalSql.Execute(sql, dt, "dbo.ZCOR1347_ENT")

        '            'RKP/01-06-12/v3.0.157
        '            'Sometimes, if the MDB size gets large, C-OPT might get hung up here.
        '            Console.WriteLine("Starting to loop through list of core objects for processing...")
        '            Console.WriteLine("Started at: " & Now())
        '            Console.WriteLine("Note: Please check size of source database (to compact & repair) if you don't notice any progress from here on, after more than a few minutes.")
        '            Console.WriteLine("")
        '            For ctr = 0 To dt.Rows.Count - 1
        '                'For Each adoTable In adoTables
        '                For adoTableCtr = 0 To adoTables.Count - 1
        '                    adoTable = adoTables(adoTableCtr)
        '                    If dt.Rows(ctr).Item("OBJ_NAME").ToString().Equals(adoTable.Name.ToString()) Then

        '                        Console.WriteLine("Processing table: " & adoTable.Name.ToString() & ". APM: " & GenUtils.GetAvailablePhysicalMemoryStr())
        '                        EntLib.COPT.Log.Log("Processing table: " & adoTable.Name.ToString() & ". APM: " & GenUtils.GetAvailablePhysicalMemoryStr())
        '                        msg = adoTable.Name.ToString() & "..."

        '                        'Console.Write("Reading source table: " & adoTable.Name.ToString())
        '                        'EntLib.COPT.Log.Log("Reading source table: " & adoTable.Name.ToString())
        '                        'ReDim arrayRestrictions(2)
        '                        'arrayRestrictions(0) = Nothing
        '                        'arrayRestrictions(1) = Nothing
        '                        'arrayRestrictions(2) = adoTable.Name

        '                        'list = New List(Of String)

        '                        'list.Add(String.Empty)
        '                        'list.Add(String.Empty)
        '                        'list.Add(adoTable.Name)

        '                        'Try
        '                        '    adoRSSchema = dalOleDb.ConnADO.OpenSchema(ADODB.SchemaEnum.adSchemaColumns, list.ToArray())
        '                        '    adoRSSchema.Sort = "ORDINAL_POSITION"
        '                        'Catch ex As Exception
        '                        '    EntLib.COPT.Log.Log("Error returning Schema.")
        '                        '    EntLib.COPT.Log.Log(ex.Message)
        '                        'End Try

        '                        adoRSSchema = New ADODB.Recordset

        '                        sql = "SELECT * FROM [" & adoTable.Name & "] WHERE 0 = 1"
        '                        Try
        '                            adoRSSchema = dalOleDbADO.ConnADO.Execute(sql)
        '                            'arrayGetRows = CType(adoRSSchema.GetRows(), String()())
        '                        Catch ex As Exception
        '                            EntLib.COPT.Log.Log("Error returning Schema - GetRows()")
        '                            EntLib.COPT.Log.Log(ex.Message)
        '                        End Try

        '                        ReDim tableDefs(0)
        '                        ReDim idxDefs(0)
        '                        adoCols = adoTable.Columns
        '                        For ctr1 = 0 To adoCols.Count - 1
        '                            adoCol = adoCols(adoRSSchema.Fields(ctr1).Name.ToString())

        '                            adoProperties = adoCol.Properties

        '                            'For Each adoProperty In adoProperties
        '                            '    Try
        '                            '        If String.IsNullOrEmpty(adoProperty.Value.ToString()) Then
        '                            '            Debug.Print(adoCol.Name.ToString() & " - " & adoProperty.Name & " - " & "")
        '                            '        Else
        '                            '            Debug.Print(adoCol.Name.ToString() & " - " & adoProperty.Name & " - " & adoProperty.Value.ToString())
        '                            '        End If
        '                            '    Catch ex As Exception
        '                            '        Debug.Print(adoCol.Name.ToString() & " - " & adoProperty.Name & " - " & "")
        '                            '    End Try
        '                            'Next
        '                            'Debug.Print(adoCol.DefinedSize.ToString())

        '                            If String.IsNullOrEmpty(tableDefs(0).fieldName) Then
        '                                tableDefs(0).fieldName = adoCol.Name
        '                                tableDefs(0).fieldType = adoCol.Type.ToString()
        '                                tableDefs(0).fieldSize = 0
        '                            Else
        '                                ReDim Preserve tableDefs(UBound(tableDefs) + 1)
        '                                tableDefs(UBound(tableDefs)).fieldName = adoCol.Name
        '                                tableDefs(UBound(tableDefs)).fieldType = adoCol.Type.ToString()
        '                                tableDefs(UBound(tableDefs)).fieldSize = 0
        '                            End If
        '                            If tableDefs(UBound(tableDefs)).fieldType.Contains("Char") Then
        '                                tableDefs(UBound(tableDefs)).fieldSize = CInt(adoCol.DefinedSize.ToString())
        '                            End If
        '                            tableDefs(UBound(tableDefs)).fieldTypeNew = clsDAL.GetADOToSQLDataType(tableDefs(UBound(tableDefs)).fieldType)
        '                            tableDefs(UBound(tableDefs)).fieldIsIdentity = CBool(adoProperties("Autoincrement").Value.ToString())
        '                            If tableDefs(UBound(tableDefs)).fieldIsIdentity Then
        '                                tableDefs(UBound(tableDefs)).fieldIsNullable = False
        '                            Else
        '                                tableDefs(UBound(tableDefs)).fieldIsNullable = True
        '                            End If
        '                            'tableDefs(UBound(tableDefs)).fieldIsNullable = CBool(adoProperties("Nullable").Value.ToString())
        '                            'If tableDefs(UBound(tableDefs)).fieldIsIdentity Then
        '                            ' tableDefs(UBound(tableDefs)).fieldIsNullable = False
        '                            'Else
        '                            'tableDefs(UBound(tableDefs)).fieldIsNullable = True
        '                            'End If
        '                        Next 'For ctr1 = 0 To adoCols.Count - 1

        '                        'EntLib.COPT.Log.Log(adoTable.Name & "..." & adoTable.Indexes.Count)

        '                        adoIdxs = adoTable.Indexes
        '                        ReDim listPK(0)
        '                        ReDim listIdx(0)
        '                        ReDim idxDefs(0)
        '                        For Each adoIdx In adoIdxs
        '                            'Debug.Print(adoTable.Name.ToString() & " - " & adoIdx.Name.ToString())
        '                            'adoIdx.Columns.c
        '                            'adoProperties = adoIdx.Properties
        '                            'For Each adoProperty In adoProperties
        '                            '    Debug.Print(adoProperty.Name & " - " & adoProperty.Value.ToString())
        '                            'Next

        '                            If String.IsNullOrEmpty(idxDefs(0).idxName) Then
        '                                'idxDefs(UBound(idxDefs)).idxName = adoIdx.Name
        '                            Else
        '                                ReDim Preserve idxDefs(UBound(idxDefs) + 1)
        '                            End If
        '                            idxDefs(UBound(idxDefs)).idxName = adoIdx.Name

        '                            adoCols = adoIdx.Columns

        '                            ReDim Preserve idxDefs(UBound(idxDefs)).idxCols(adoCols.Count - 1)

        '                            'EntLib.COPT.Log.Log("   Index name: " & adoIdx.Name & " - No. of index columns: " & adoIdx.Columns.Count)
        '                            colCtr = -1
        '                            For Each adoCol In adoCols
        '                                colCtr += 1
        '                                'Debug.Print("TableName=" & adoTable.Name.ToString() & ", PK=" & adoIdx.PrimaryKey.ToString() & ", IdxName=" & adoIdx.Name.ToString() & ", ColName=" & adoCol.Name.ToString())
        '                                If adoIdx.PrimaryKey Then
        '                                    If String.IsNullOrEmpty(listPK(0)) Then
        '                                        listPK(0) = adoCol.Name.ToString()
        '                                    Else
        '                                        ReDim Preserve listPK(UBound(listPK) + 1)
        '                                        listPK(UBound(listPK)) = adoCol.Name.ToString()
        '                                    End If
        '                                Else
        '                                    If String.IsNullOrEmpty(listIdx(0)) Then
        '                                        listIdx(0) = adoCol.Name.ToString()
        '                                    Else
        '                                        ReDim Preserve listIdx(UBound(listIdx) + 1)
        '                                        listIdx(UBound(listIdx)) = adoCol.Name.ToString()
        '                                    End If
        '                                End If

        '                                idxDefs(UBound(idxDefs)).idxCols( _
        '                                    colCtr _
        '                                    ) = adoCol.Name.ToString()

        '                                idxDefs(UBound(idxDefs)).isUnique = adoIdx.Unique

        '                                'Console.WriteLine(adoTable.Name & "..." & adoTable.Indexes.Count & "...Index name: " & idxDefs(UBound(idxDefs)).idxName & ", Column name: " & adoCol.Name.ToString() & ", Is unique: " & idxDefs(UBound(idxDefs)).isUnique)
        '                                'EntLib.COPT.Log.Log(adoTable.Name & "..." & adoTable.Indexes.Count & "...Index name: " & idxDefs(UBound(idxDefs)).idxName & ", Column name: " & adoCol.Name.ToString() & ", Is unique: " & idxDefs(UBound(idxDefs)).isUnique)

        '                            Next 'For Each adoCol In adoCols
        '                        Next 'For Each adoIdx In adoIdxs

        '                        For ctr1 = 0 To listPK.Length - 1
        '                            For ctr2 = 0 To tableDefs.Length - 1
        '                                If tableDefs(ctr2).fieldName.Equals(listPK(ctr1)) Then
        '                                    tableDefs(ctr2).fieldIsNullable = False
        '                                    Exit For
        '                                End If
        '                            Next
        '                        Next

        '                        'Console.WriteLine("...Complete.")
        '                        'EntLib.COPT.Log.Log("...Complete.")

        '                        'Console.Write("Creating destination table: " & adoTable.Name.ToString())
        '                        'EntLib.COPT.Log.Log("Creating destination table: " & adoTable.Name.ToString())

        '                        'Create table + indexes in the SQL Express database - START
        '                        If CBool(dt.Rows(ctr).Item("Recreate").ToString()) Then
        '                            sql = "IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[" & adoTable.Name.ToString() & "]') AND type in (N'U')) DROP TABLE [dbo].[" & adoTable.Name.ToString() & "]"
        '                            Try
        '                                dalSql.Execute(sql, dt1)
        '                                createObject = True
        '                                'msg = msg & "...DROPPED..."
        '                                'EntLib.COPT.Log.Log(adoTable.Name.ToString() & "...DROP TABLE..." & CBool(dt.Rows(ctr).Item("Recreate").ToString()) & "...OK")
        '                            Catch ex As Exception
        '                                msg = msg & "Error in DROP TABLE..."
        '                                Console.WriteLine("Error creating table+indexes in SQLEXPRESS database.")
        '                                Console.WriteLine(ex.Message)
        '                                EntLib.COPT.Log.Log("Error creating table+indexes in SQLEXPRESS database.")
        '                                EntLib.COPT.Log.Log(ex.Message)
        '                                createObject = False
        '                                errCount += 1
        '                            End Try
        '                        Else
        '                            sql = "SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[" & adoTable.Name.ToString() & "]') AND type in (N'U')"
        '                            'EntLib.COPT.Log.Log(ex.Message)
        '                            Try
        '                                dalSql.Execute(sql, dt1)
        '                                'EntLib.COPT.Log.Log(adoTable.Name.ToString() & "...DROP TABLE..." & CBool(dt.Rows(ctr).Item("Recreate").ToString()) & "...OK")
        '                            Catch ex As Exception
        '                                Console.WriteLine("Error executing SQL.")
        '                                Console.WriteLine(sql)
        '                                Console.WriteLine(ex.Message)
        '                                EntLib.COPT.Log.Log("Error executing SQL.")
        '                                EntLib.COPT.Log.Log(sql)
        '                                EntLib.COPT.Log.Log(ex.Message)
        '                                errCount += 1
        '                            End Try

        '                            If dt1 Is Nothing Then
        '                                createObject = True
        '                            Else
        '                                If dt1.Rows.Count > 0 Then
        '                                    createObject = False
        '                                Else
        '                                    createObject = True
        '                                End If
        '                            End If
        '                        End If
        '                        If createObject Then
        '                            'Step 1: Drop table in destination database, if it already exists
        '                            sql = "IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[" & adoTable.Name.ToString() & "]') AND type in (N'U')) DROP TABLE [dbo].[" & adoTable.Name.ToString() & "]"
        '                            Try
        '                                dalSql.Execute(sql, dt1)
        '                                EntLib.COPT.Log.Log(adoTable.Name.ToString() & "...DROP TABLE..." & CBool(dt.Rows(ctr).Item("Recreate").ToString()) & "...OK")
        '                                msg = msg & "DROPTABLE-OK..."
        '                                Console.WriteLine("   DROP   TABLE...OK")
        '                            Catch ex As Exception
        '                                msg = msg & "DROPTABLE-Error..."
        '                                Console.WriteLine("   DROP TABLE...Error.")
        '                                Console.WriteLine("Error executing SQL.")
        '                                Console.WriteLine(sql)
        '                                Console.WriteLine(ex.Message)
        '                                EntLib.COPT.Log.Log("Error executing SQL.")
        '                                EntLib.COPT.Log.Log(sql)
        '                                EntLib.COPT.Log.Log(ex.Message)
        '                                errCount += 1
        '                            End Try

        '                            'Step 2: Create table in destination database + create primary key
        '                            sql = "SET ANSI_NULLS ON"
        '                            Try
        '                                dalSql.Execute(sql, dt1)
        '                                EntLib.COPT.Log.Log(adoTable.Name.ToString() & "...SET ANSI_NULLS ON..." & "...OK")
        '                            Catch ex As Exception
        '                                msg = msg & "Error in SET ANSI_NULLS ON..."
        '                                Console.WriteLine("Error executing SQL.")
        '                                Console.WriteLine(sql)
        '                                Console.WriteLine(ex.Message)
        '                                EntLib.COPT.Log.Log("Error executing SQL.")
        '                                EntLib.COPT.Log.Log(sql)
        '                                EntLib.COPT.Log.Log(ex.Message)
        '                                errCount += 1
        '                            End Try

        '                            sql = "SET QUOTED_IDENTIFIER ON"
        '                            Try
        '                                dalSql.Execute(sql, dt1)
        '                                EntLib.COPT.Log.Log(adoTable.Name.ToString() & "...SET QUOTED_IDENTIFIER ON..." & "...OK")
        '                            Catch ex As Exception
        '                                msg = msg & "Error in SET QUOTED_IDENTIFIER ON..."
        '                                Console.WriteLine("Error executing SQL.")
        '                                Console.WriteLine(sql)
        '                                Console.WriteLine(ex.Message)
        '                                EntLib.COPT.Log.Log("Error executing SQL.")
        '                                EntLib.COPT.Log.Log(sql)
        '                                EntLib.COPT.Log.Log(ex.Message)
        '                                errCount += 1
        '                            End Try

        '                            sql = "CREATE TABLE [dbo].[" & adoTable.Name.ToString() & "] ("
        '                            For ctr1 = 0 To tableDefs.Length - 1
        '                                sql = sql & "[" & tableDefs(ctr1).fieldName & "] [" & tableDefs(ctr1).fieldTypeNew & "]" & IIf(tableDefs(ctr1).fieldSize > 0, "(" & tableDefs(ctr1).fieldSize & ") ", "").ToString() & IIf(tableDefs(ctr1).fieldIsIdentity, " IDENTITY(1,1)", "").ToString() & IIf(tableDefs(ctr1).fieldIsNullable, " NULL", " NOT NULL").ToString() & ","
        '                            Next
        '                            If String.IsNullOrEmpty(listPK(0)) Then
        '                                sql = Left(sql, Len(sql) - 1) & vbNewLine
        '                            Else
        '                                sql = sql & " CONSTRAINT [PK_" & adoTable.Name.ToString() & "] PRIMARY KEY CLUSTERED " & vbNewLine
        '                                sql = sql & "(" & vbNewLine
        '                                For ctr1 = 0 To listPK.Length - 1
        '                                    sql = sql & "[" & listPK(ctr1) & "] ASC,"
        '                                Next
        '                                sql = Left(sql, Len(sql) - 1) & vbNewLine
        '                                sql = sql & ")WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]" & vbNewLine
        '                            End If
        '                            sql = sql & ") ON [PRIMARY]" & vbNewLine
        '                            Try
        '                                dalSql.Execute(sql, dt1)
        '                                EntLib.COPT.Log.Log(adoTable.Name.ToString() & "...CREATE TABLE..." & "...OK")
        '                                msg = msg & "CREATETABLE-OK..."
        '                                Console.WriteLine("   CREATE TABLE...OK")
        '                            Catch ex As Exception
        '                                msg = msg & "Error in CREATETABLE..."
        '                                Debug.Print("Error creating TABLE - " & adoTable.Name.ToString())
        '                                Debug.Print(ex.Message)
        '                                Debug.Print(sql)
        '                                Debug.Print("")
        '                                Console.WriteLine("Error executing SQL.")
        '                                Console.WriteLine(sql)
        '                                Console.WriteLine(ex.Message)
        '                                EntLib.COPT.Log.Log("Error executing SQL.")
        '                                EntLib.COPT.Log.Log(sql)
        '                                EntLib.COPT.Log.Log(ex.Message)
        '                                errCount += 1
        '                            End Try
        '                            'Step 3: Create nonclustered index on the table

        '                            'If Not String.IsNullOrEmpty(listIdx(0)) Then
        '                            '    sql = "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[" & adoTable.Name.ToString() & "]') AND name = N'" & "IDX_" & adoTable.Name.ToString() & "')"
        '                            '    sql = sql & "CREATE UNIQUE NONCLUSTERED INDEX [" & "IDX_" & adoTable.Name.ToString() & "] ON [dbo].[" & adoTable.Name.ToString() & "] "
        '                            '    sql = sql & "( "
        '                            '    For ctr1 = 0 To listIdx.Length - 1
        '                            '        sql = sql & "[" & listIdx(ctr1) & "] ASC,"
        '                            '    Next
        '                            '    sql = Left(sql, Len(sql) - 1) & vbNewLine
        '                            '    sql = sql & ")WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY] "
        '                            '    Try
        '                            '        dalSql.Execute(sql, dt1)
        '                            '    Catch ex As Exception
        '                            '        Debug.Print("Error creating NONCLUSTERED INDEX on TABLE - " & adoTable.Name.ToString())
        '                            '        Debug.Print(ex.Message)
        '                            '        Debug.Print(sql)
        '                            '        Debug.Print("")
        '                            '        Console.WriteLine("Error executing SQL.")
        '                            '        Console.WriteLine(sql)
        '                            '        Console.WriteLine(ex.Message)
        '                            '        EntLib.COPT.Log.Log("Error executing SQL.")
        '                            '        EntLib.COPT.Log.Log(sql)
        '                            '        EntLib.COPT.Log.Log(ex.Message)
        '                            '    End Try
        '                            'End If

        '                            If idxDefs IsNot Nothing Then
        '                                If Not String.IsNullOrEmpty(idxDefs(0).idxName) Then
        '                                    idxComposite = ""
        '                                    For ctr1 = 0 To idxDefs.Length - 1
        '                                        idxComposite = ""
        '                                        For ctr2 = 0 To idxDefs(ctr1).idxCols.Length - 1
        '                                            If String.IsNullOrEmpty(idxComposite) Then
        '                                                idxComposite = idxDefs(ctr1).idxCols(ctr2)
        '                                            Else
        '                                                idxComposite = idxComposite & "," & idxDefs(ctr1).idxCols(ctr2)
        '                                            End If
        '                                        Next

        '                                        sql = "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[" & adoTable.Name.ToString() & "]') AND name = N'" & "IDX_" & idxDefs(ctr1).idxName & "')"
        '                                        sql = sql & "CREATE " & IIf(idxDefs(ctr1).isUnique, "UNIQUE", "").ToString() & " NONCLUSTERED INDEX [" & "IDX_" & idxDefs(ctr1).idxName & "] ON [dbo].[" & adoTable.Name.ToString() & "] "
        '                                        sql = sql & "( "
        '                                        'For ctr1 = 0 To listIdx.Length - 1
        '                                        '    sql = sql & "[" & listIdx(ctr1) & "] ASC,"
        '                                        'Next
        '                                        'sql = Left(sql, Len(sql) - 1) & vbNewLine
        '                                        sql = sql & idxComposite
        '                                        sql = sql & ") WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY] "
        '                                        Try
        '                                            dalSql.Execute(sql, dt1)
        '                                            EntLib.COPT.Log.Log(adoTable.Name.ToString() & "..." & "CREATE " & IIf(idxDefs(ctr1).isUnique, "UNIQUE", "").ToString() & " NONCLUSTERED INDEX " & idxDefs(ctr1).idxName & "..." & "...OK")
        '                                        Catch ex As Exception
        '                                            Debug.Print("Error creating NONCLUSTERED INDEX on TABLE - " & adoTable.Name.ToString())
        '                                            Debug.Print(ex.Message)
        '                                            Debug.Print(sql)
        '                                            Debug.Print("")
        '                                            Console.WriteLine("Error executing SQL.")
        '                                            Console.WriteLine(sql)
        '                                            Console.WriteLine(ex.Message)
        '                                            EntLib.COPT.Log.Log("Error executing SQL.")
        '                                            EntLib.COPT.Log.Log(sql)
        '                                            EntLib.COPT.Log.Log(ex.Message)
        '                                            errCount += 1
        '                                        End Try

        '                                    Next
        '                                    msg = msg & "CREATEINDEX-OK..."
        '                                    Console.WriteLine("   CREATE INDEX...OK")
        '                                End If
        '                            End If

        '                        End If
        '                        'Console.Write("...Created")
        '                        'Create table + indexes in the SQL Express database - END

        '                        'Copy data - START
        '                        If _
        '                            CInt(dt.Rows(ctr).Item("TransferData").ToString()) = 1 _
        '                            Or _
        '                            CInt(dt.Rows(ctr).Item("TransferData").ToString()) = 3 _
        '                        Then
        '                            'Empty destination table
        '                            If noTruncate Then
        '                                'sql = "DELETE FROM [" & adoTable.Name.ToString() & "]"
        '                                sql = "EXEC dbo.asp_TruncateTable '" & adoTable.Name.ToString() & "'"
        '                            Else
        '                                sql = "TRUNCATE TABLE [" & adoTable.Name.ToString() & "]"
        '                            End If

        '                            Try
        '                                dalSql.Execute(sql, dt1)
        '                                EntLib.COPT.Log.Log(adoTable.Name.ToString() & "...TRUNCATE TABLE...OK")
        '                            Catch ex As Exception
        '                                Console.WriteLine("Error executing SQL.")
        '                                Console.WriteLine(sql)
        '                                Console.WriteLine(ex.Message)
        '                                EntLib.COPT.Log.Log("Error executing SQL.")
        '                                EntLib.COPT.Log.Log(sql)
        '                                EntLib.COPT.Log.Log(ex.Message)
        '                            End Try

        '                            Console.Write("   BULK    COPY...")
        '                            EntLib.COPT.Log.Log("Create ADO Recordset...")
        '                            sql = "SELECT * FROM [" & adoTable.Name.ToString() & "]"
        '                            Try
        '                                'dalOleDbADO.Execute(sql, dt1)
        '                                dalOleDb.Execute(sql, dt1)
        '                                If dalOleDbADO.LastErrorNo = -1 Then
        '                                    Console.WriteLine("ERROR. APM: " & GenUtils.GetAvailablePhysicalMemoryStr())
        '                                    Console.WriteLine("Error executing SQL.")
        '                                    Console.WriteLine("SELECT * FROM [" & adoTable.Name.ToString() & "]")
        '                                    Console.WriteLine(dalOleDbADO.LastErrorDesc)
        '                                    'Console.WriteLine(ex.Message)
        '                                    EntLib.COPT.Log.Log("Error executing SQL.")
        '                                    EntLib.COPT.Log.Log("SELECT * FROM [" & adoTable.Name.ToString() & "]")
        '                                    EntLib.COPT.Log.Log(dalOleDbADO.LastErrorDesc)
        '                                    EntLib.COPT.Log.Log("APM: " & GenUtils.GetAvailablePhysicalMemoryStr())
        '                                Else
        '                                    EntLib.COPT.Log.Log(adoTable.Name.ToString() & "...SELECT * FROM...OK")
        '                                    EntLib.COPT.Log.Log("Create ADO.NET Recordset...SUCCESS.")
        '                                End If
        '                            Catch ex As Exception
        '                                Console.WriteLine("Error executing SQL (using ADO.NET).")
        '                                Console.WriteLine(sql)
        '                                Console.WriteLine(ex.Message)
        '                                EntLib.COPT.Log.Log("Error executing SQL (using ADO.NET).")
        '                                EntLib.COPT.Log.Log(sql)
        '                                EntLib.COPT.Log.Log(ex.Message)
        '                                errCount += 1
        '                            End Try
        '                            If dalOleDbADO.LastErrorNo <> -1 Then
        '                                sql = "dbo.[" & adoTable.Name.ToString() & "]"
        '                                Try
        '                                    'Console.Write("   BULK    COPY...")
        '                                    EntLib.COPT.Log.Log("BulkCopy...")
        '                                    ret = dalSql.Execute("BULKCOPY-MAP", dt1, sql)
        '                                    If ret = -1 Then
        '                                        EntLib.COPT.Log.Log(adoTable.Name.ToString() & "...BULKCOPY-MAP...ERROR")
        '                                        msg = msg & "...BULKCOPY-ERROR..."
        '                                        'Console.WriteLine("   BULK    COPY...OK")
        '                                        Console.WriteLine("ERROR. APM: " & GenUtils.GetAvailablePhysicalMemoryStr())
        '                                        Console.WriteLine("No. of rows: " & dt1.Rows.Count.ToString())
        '                                        Console.WriteLine("ERROR Desc:")
        '                                        Console.WriteLine(dalSql.LastErrorDesc)
        '                                        Console.WriteLine("Please check C-OPT.log for details.")
        '                                        EntLib.COPT.Log.Log("No. of rows: " & dt1.Rows.Count.ToString())
        '                                        errCount += 1
        '                                    Else
        '                                        EntLib.COPT.Log.Log(adoTable.Name.ToString() & "...BULKCOPY-MAP...OK (" & dt1.Rows.Count.ToString() & " rows)")
        '                                        msg = msg & "...BULKCOPY-OK..."
        '                                        EntLib.COPT.Log.Log("BulkCopy...SUCCESS.")
        '                                        'Console.WriteLine("   BULK    COPY...OK")
        '                                        Console.WriteLine("OK (" & dt1.Rows.Count.ToString() & " rows)")
        '                                    End If
        '                                Catch ex As Exception
        '                                    Console.WriteLine("Error executing SQL.")
        '                                    Console.WriteLine(sql)
        '                                    Console.WriteLine(ex.Message)
        '                                    Console.WriteLine("No. of rows: " & dt1.Rows.Count.ToString())
        '                                    EntLib.COPT.Log.Log("Error executing SQL.")
        '                                    EntLib.COPT.Log.Log(sql)
        '                                    EntLib.COPT.Log.Log(ex.Message)
        '                                    EntLib.COPT.Log.Log("No. of rows: " & dt1.Rows.Count.ToString())
        '                                    errCount += 1
        '                                End Try
        '                            End If
        '                            'Console.Write("...Populated with data.")
        '                        End If
        '                        'Copy data - END

        '                        'Console.WriteLine(msg)

        '                        Exit For
        '                    End If 'CInt(dt.Rows(ctr).Item("TransferData").ToString()) = 1 Or CInt(dt.Rows(ctr).Item("TransferData").ToString()) = 3
        '                Next 'For Each adoTable In adoTables
        '                'If GenUtils.IsSwitchAvailable(switches, "/UseGC") Then
        '                GenUtils.CollectGarbage()
        '                'End If
        '                'dt1 = Nothing
        '            Next 'For ctr = 0 To dt.Rows.Count - 1

        '            Console.WriteLine("All TRANSFER tasks finished successfully.")
        '            EntLib.COPT.Log.Log("All TRANSFER tasks finished successfully.")

        '        End If 'If proceed Then 'proceed #2
        '    End If 'If proceed Then 'proceed #1

        '    dalOleDbADO = Nothing
        '    dalOleDb = Nothing
        '    dalSql = Nothing
        '    dt = Nothing
        '    dt1 = Nothing

        '    GenUtils.CollectGarbage()

        '    Try
        '        GenUtils.GetSysDatabase().ExecuteNonQuery( _
        '            "INSERT INTO tblRunDetails " & _
        '            "SELECT " & _
        '            "'" & My.Computer.Clock.LocalTime.ToString() & "' AS LCT, " & _
        '            "'" & My.Computer.Clock.GmtTime.ToString() & "' AS GMT, " & _
        '            "'" & My.Computer.Name.Replace("=", ">") & "' AS CNM, " & _
        '            "'" & GenUtils.GetUserName().ToString() & "' AS UNM, " & _
        '            "'" & My.Application.Info.AssemblyName & "' AS EXE, " & _
        '            "'" & GenUtils.Version & "' AS VER, " & _
        '            "'" & GenUtils.VersionName & "' AS VEN, " & _
        '            "'" & "TRANSFERBACK" & "' AS PRJ, " & _
        '            "'" & "" & "' AS RUN, " & _
        '            "'" & "' AS SLV, " & _
        '            "'" & "' AS SVE, " & _
        '            "'" & "' AS PTY, " & _
        '            "'" & "" & "' AS SST, " & _
        '            "'" & "" & "' AS ROW, " & _
        '            "'" & "" & "' AS COL, " & _
        '            "'" & "" & "' AS NZE, " & _
        '            "'" & "" & "' AS OBJ, " & _
        '            "'" & "" & "' AS ITE, " & _
        '            "'" & GenUtils.FormatTime(startTime) & "' AS STM, " & _
        '            "'" & errCount & "' AS BAD, " & _
        '            "'" & "" & "' AS INF, " & _
        '            "'" & GenUtils.GetAvailablePhysicalMemoryStr() & "" & "' AS APM, " & _
        '            "'" & "" & "' AS CST, " & _
        '            "'" & "TRANSFERBACK /SrcDBName """ & srcDBName & """ /SrcDBFolder """ & srcDBFolder & """ /DestDBName """ & destDBName & """ /DestDBFolder """ & destDBFolder & """" & "" & "' AS SWT " & _
        '            "" _
        '        )
        '    Catch ex1 As Exception
        '        EntLib.COPT.Log.Log(GenUtils.GetWorkDir(switches), "", "Error logging run details to C-OPTSYS database." & vbNewLine & "Error:" & vbNewLine & ex1.Message)
        '    End Try

        '    Console.WriteLine("TRANSFER took: " & GenUtils.FormatTime(startTime) & ".")
        '    EntLib.COPT.Log.Log("TRANSFER took: " & GenUtils.FormatTime(startTime) & ".")
        '    EntLib.COPT.Log.Log("********************")

        '    Return 0
        'End Function


        ' TRANSFER (TRF) - Transfers objects and data from source (MS Access) database to destination (SQL Server Express) database, as defined in the system table, tsysCoreObjects (in the source database).

        '' RKP/04-20-12/v3.2.166
        '' This is a faster version to DB_TransferADO.
        '' The "ADO" version takes ~ 10 minutes to load ADOX, especially the following line:
        '' ret = adoTables.Count

        'Public Shared Function DB_TransferDAO1( _
        '    ByRef switches() As String, _
        '    ByVal srcDBName As String, _
        '    ByVal srcDBFolder As String, _
        '    ByVal destDBName As String, _
        '    ByVal destDBFolder As String _
        ') As Integer

        '    'Console.WriteLine("Entered TRANSFER.")

        '    Dim dalOleDbADO As New clsDAL 'ADO version 'RKP/04-20-12/v3.2.166
        '    Dim dalOleDb As New clsDAL 'ADO.NET version 'RKP/04-20-12/v3.2.166
        '    Dim dalDAO As New clsDAL 'DAO version 'RKP/04-20-12/v3.2.166
        '    Dim dalSql As New clsDAL
        '    Dim connOleDb As New OleDb.OleDbConnection
        '    Dim connSql As New SqlClient.SqlConnection
        '    Dim adoCatalog As New ADOX.Catalog
        '    'Dim adoTables As ADOX.Tables
        '    'Dim adoTable As ADOX.Table
        '    'Dim adoIdxs As ADOX.Indexes
        '    'Dim adoIdx As ADOX.Index
        '    'Dim adoCols As ADOX.Columns
        '    'Dim adoCol As ADOX.Column
        '    'Dim adoProperties As ADOX.Properties
        '    'Dim adoRSSchema As ADODB.Recordset
        '    Dim sql As String = ""
        '    Dim dt As New DataTable
        '    Dim dt1 As New DataTable
        '    Dim filePath As String = ""
        '    Dim listPK() As String
        '    Dim listIdx() As String
        '    Dim ctr As Integer = 0
        '    Dim ctr1 As Integer = 0
        '    Dim ctr2 As Integer = 0
        '    Dim tableDefs() As STRUCT_TABLEDEF
        '    Dim idxDefs() As STRUCT_IDX
        '    Dim createObject As Boolean = False
        '    Dim proceed As Boolean = False
        '    Dim srcFilePath As String = ""
        '    Dim destFilePath As String = ""
        '    Dim msg As String = ""
        '    Dim colCtr As Integer = 0
        '    Dim idxComposite As String = ""
        '    Dim adoTableCtr As Integer = 0
        '    Dim ret As Integer
        '    Dim startTime As Date
        '    Dim daoTableDefs As dao.TableDefs
        '    Dim daoTableDef As dao.TableDef
        '    Dim daoFields As dao.Fields
        '    Dim daoField As dao.Field
        '    Dim daoProperties As dao.Properties
        '    Dim daoIndexes As dao.Indexes
        '    Dim daoIndex As dao.Index
        '    Dim daoCols As dao.Fields
        '    Dim daoCol As dao.Field
        '    Dim daoRSSchema As dao.Recordset
        '    Dim noTruncate As Boolean = GenUtils.IsSwitchAvailable(switches, "/UseTruncateSP")

        '    'MessageBox.Show("Entered DB_TransferObjects2.", "C-OPT", MessageBoxButtons.OK, MessageBoxIcon.Information)

        '    'Console.Write("Do you want to continue? (Y/N): ")
        '    'input = Console.ReadLine()
        '    'If input.Trim.ToUpper.Equals("Y") Then
        '    '    proceed = True
        '    'Else
        '    '    proceed = False
        '    'End If

        '    'Console.ReadKey()

        '    'If String.IsNullOrEmpty(srcDBFolder) Then
        '    '    srcDBFolder = "C:\OPTMODELS\" & srcDBName
        '    'End If

        '    'If String.IsNullOrEmpty(destDBFolder) Then
        '    '    destDBFolder = "C:\OPTMODELS\" & destDBName
        '    'End If

        '    'filePath = srcDBFolder & "\" & srcDBName & IIf(srcDBName.Trim().ToUpper().EndsWith(".MDB"), "", ".mdb").ToString() '& ".mdb"  '"C:\OPTMODELS\SRC1\SRC1.mdb"
        '    'srcFilePath = filePath
        '    'If My.Computer.FileSystem.FileExists(filePath) Then
        '    '    proceed = True
        '    'Else
        '    '    proceed = False
        '    '    msg = msg & vbNewLine & "Source database not found."
        '    'End If
        '    'If proceed Then
        '    '    filePath = destDBFolder & "\" & destDBName & IIf(destDBName.Trim().ToUpper().EndsWith(".MDF"), "", ".mdf").ToString() '& ".mdf"
        '    '    destFilePath = filePath
        '    '    If My.Computer.FileSystem.FileExists(filePath) Then
        '    '        proceed = True
        '    '    Else
        '    '        proceed = False
        '    '        msg = msg & vbNewLine & "Destination database not found."
        '    '    End If
        '    'End If
        '    proceed = True
        '    If proceed Then 'proceed #1

        '        'If EntLib.COPT.GenUtils.IsSwitchAvailable(switches, "/NoPrompt") Then
        '        '    proceed = True
        '        'Else
        '        '    Console.WriteLine()
        '        '    Console.WriteLine("Source File Path: " & srcFilePath)
        '        '    Console.WriteLine("Destination File Path: " & destFilePath)
        '        '    Console.Write("Do you want to continue? (Y/N): ")
        '        '    input = Console.ReadLine()
        '        '    If input.Trim.ToUpper.Equals("Y") Then
        '        '        proceed = True
        '        '    Else
        '        '        proceed = False
        '        '    End If
        '        'End If

        '        If proceed Then 'proceed #2

        '            'MessageBox.Show("About to create log entry.", "C-OPT", MessageBoxButtons.OK, MessageBoxIcon.Information)

        '            'Console.Write("Do you want to continue? (Y/N): ")
        '            'Input = Console.ReadLine()
        '            'If Input.Trim.ToUpper.Equals("Y") Then
        '            '    proceed = True
        '            'Else
        '            '    proceed = False
        '            'End If

        '            EntLib.COPT.Log.Log("")
        '            EntLib.COPT.Log.Log("********************")
        '            EntLib.COPT.Log.Log("TRANSFER (TRF) - Started at: " & Now())

        '            EntLib.COPT.Log.Log("Source File Path: " & srcFilePath)
        '            EntLib.COPT.Log.Log("Destination File Path: " & destFilePath)

        '            filePath = srcDBFolder & "\" & srcDBName & IIf(srcDBName.Trim().ToUpper().EndsWith(".MDB"), "", ".mdb").ToString() '& ".mdb"  '"C:\OPTMODELS\SRC1\SRC1.mdb"
        '            dalOleDbADO.UseADO = True
        '            Try
        '                dalOleDbADO.Connect(clsDAL.e_DB.e_db_ACCESS, clsDAL.e_ConnType.e_connType_C, Nothing, connOleDb, Nothing, Nothing, Nothing, True, filePath, "C:\OPTMODELS\C-OPTSYS\System.mdw.copt")
        '                EntLib.COPT.Log.Log("Connected to source database (using ADO) successfully.")
        '                Console.WriteLine("Connected to source database (using ADO) successfully.")
        '            Catch ex As Exception
        '                Console.WriteLine("Error connecting to source database (using ADO).")
        '                EntLib.COPT.Log.Log("Error connecting to source database (using ADO).")
        '                EntLib.COPT.Log.Log(ex.Message)
        '            End Try

        '            'RKP/04-20-12/v3.2.166
        '            'This is the ADO.NET version of the source connection object.
        '            'The purpose of this object is to avoid the use of "ConvertADORecordsetToDataTable", which avoids duplicate in-memory objects (adoRS and dt).
        '            'This adversely affects performance, especially for large tables (~ million rows).
        '            dalOleDb.UseADO = False
        '            Try
        '                dalOleDb.Connect(clsDAL.e_DB.e_db_ACCESS, clsDAL.e_ConnType.e_connType_C, Nothing, connOleDb, Nothing, Nothing, Nothing, True, filePath, "C:\OPTMODELS\C-OPTSYS\System.mdw.copt")
        '                EntLib.COPT.Log.Log("Connected to source database successfully (using ADO.NET).")
        '                Console.WriteLine("Connected to source database successfully (using ADO.NET).")
        '            Catch ex As Exception
        '                Console.WriteLine("Error connecting to source database (using ADO.NET).")
        '                EntLib.COPT.Log.Log("Error connecting to source database (using ADO.NET).")
        '                EntLib.COPT.Log.Log(ex.Message)
        '            End Try

        '            dalDAO.UseDAO = True
        '            Try
        '                dalDAO.Connect(clsDAL.e_DB.e_db_ACCESS, clsDAL.e_ConnType.e_connType_C, Nothing, Nothing, Nothing, Nothing, Nothing, True, filePath, "C:\OPTMODELS\C-OPTSYS\System.mdw.copt")
        '                EntLib.COPT.Log.Log("Connected to source database successfully (using DAO).")
        '                Console.WriteLine("Connected to source database successfully (using DAO).")
        '            Catch ex As Exception
        '                Console.WriteLine("Error connecting to source database (using DAO).")
        '                EntLib.COPT.Log.Log("Error connecting to source database (using DAO).")
        '                EntLib.COPT.Log.Log(ex.Message)
        '            End Try

        '            startTime = Now()
        '            Console.WriteLine("Reading catalog from source database...")
        '            Console.Write("Started at: " & startTime.ToString() & ". ")

        '            'ret = dalOleDb.Execute("SELECT * FROM qsysMSysObjects WHERE OBJ_TYPE_ID = 1", dtADOTables)

        '            'RKP/02-12-12/v3.0.157
        '            'Added additional error-checking to prevent C-OPT from crashing at this point.
        '            Try
        '                'adoCatalog = dalOleDbADO.ConnADOCatalog
        '                'adoTables = adoCatalog.Tables
        '                'ret = adoTables.Count
        '                daoTableDefs = dalDAO.DAODatabase.TableDefs
        '                ret = daoTableDefs.Count
        '            Catch ex As Exception
        '                Console.WriteLine("Error.")
        '                Console.WriteLine("adoTables.Count = 0.")
        '                Console.WriteLine("Please reach out to BMOS for resolution.")
        '                Console.WriteLine("The source database has objects in it (like, linked tables, etc) that are incompatible with C-OPT.")
        '                Console.WriteLine("You must remove all incompatible objects in the source database (like linked tables to external databases, invalid queries, etc) and try the operation again.")
        '                EntLib.COPT.Log.Log("Error. adoTables.Count = 0.")
        '                EntLib.COPT.Log.Log("The source database has objects in it (like, linked tables, etc) that are incompatible with C-OPT.")
        '                EntLib.COPT.Log.Log("If, for example, the database has linked tables to, say, a MySQL database, C-OPT will have difficulty connecting to it.")
        '                EntLib.COPT.Log.Log("You must remove all incompatible objects in the source database (like linked tables to external databases, invalid queries, etc) and try the operation again.")
        '                EntLib.COPT.Log.Log(ex.Message)
        '                Return -1
        '            End Try

        '            If ret = 0 Then
        '                Console.WriteLine("Error.")
        '                Console.WriteLine("Tables.Count = 0.")
        '                Console.WriteLine("Please reach out to BMOS for resolution.")
        '                EntLib.COPT.Log.Log("Error - Tables.Count = 0.")
        '                Return -1
        '            Else
        '                Console.WriteLine("Ended at: " & Now().ToString() & " (" & DateDiff(DateInterval.Minute, startTime, Now()).ToString() & " min.)")
        '                EntLib.COPT.Log.Log("Reading Catalog took: " & DateDiff(DateInterval.Minute, startTime, Now()).ToString() & " min.")
        '            End If



        '            'Console.WriteLine(adoTables.Count.ToString() & " tables & views found in database.")

        '            'http://support.microsoft.com/kb/299484
        '            'adoRSSchema = dalOleDb.ConnADO.OpenSchema(ADODB.SchemaEnum.adSchemaColumns, Array(Nothing, Nothing, "Products"))

        '            'dalSql.Connect(clsDAL.e_DB.e_db_SQLSERVER, clsDAL.e_ConnType.e_connType_F, Nothing, Nothing, connSql, Nothing, Nothing, "S02ASQLNPD01", "BMOS")
        '            'Works:
        '            'dalSql.Connect(clsDAL.e_DB.e_db_SQLSERVER, clsDAL.e_ConnType.e_connType_I, Nothing, Nothing, connSql, Nothing, Nothing, ".\SQLEXPRESS", "C:\OPTMODELS\SRC1SQL\SRC1SQL.mdf", "SRC1SQL")
        '            'EntLib.COPT.Log.Log("")
        '            Try
        '                dalSql.Connect( _
        '                        clsDAL.e_DB.e_db_SQLSERVER, _
        '                        clsDAL.e_ConnType.e_connType_I, _
        '                        Nothing, _
        '                        Nothing, _
        '                        connSql, _
        '                        Nothing, _
        '                        Nothing, _
        '                        True, _
        '                        ".\SQLEXPRESS", _
        '                        destDBFolder & "\" & destDBName & _
        '                        IIf(destDBName.Trim().ToUpper().EndsWith(".MDF"), "", ".mdf").ToString(), _
        '                        IIf(destDBName.Trim().ToUpper().EndsWith(".MDF"), _
        '                        Left(destDBName.Trim(), Len(destDBName.Trim()) - 4), destDBName.Trim()).ToString() _
        '                )
        '                EntLib.COPT.Log.Log("Connected to destination database successfully.")
        '                Console.WriteLine("Connected to destination database successfully.")
        '            Catch ex As Exception
        '                Console.WriteLine("Error connecting to destination database.")
        '                EntLib.COPT.Log.Log("Error connecting to destination database.")
        '                EntLib.COPT.Log.Log(ex.Message)
        '            End Try

        '            'sql = "SELECT * FROM tsysCOL"
        '            'dalSql.Execute(sql, dt)

        '            sql = "SELECT * FROM tsysCoreObjects WHERE OBJ_TYPE_ID = 1 AND [Active] = True"
        '            Console.Write("Reading list of core objects...")
        '            EntLib.COPT.Log.Log(sql)
        '            Try
        '                dalOleDbADO.Execute(sql, dt)
        '                Console.WriteLine("complete.")
        '            Catch ex As Exception
        '                Console.WriteLine("error.")
        '                Console.WriteLine("Error:")
        '                Console.WriteLine(ex.Message)
        '                EntLib.COPT.Log.Log("Error:")
        '                EntLib.COPT.Log.Log(ex.Message)
        '            End Try

        '            'sql = "TRUNCATE TABLE dbo.ZCOR1347_ENT"
        '            'dalSql.Execute(sql, dt)

        '            'sql = "BULKCOPY"
        '            'dalSql.Execute(sql, dt, "dbo.ZCOR1347_ENT")

        '            'RKP/01-06-12/v3.0.157
        '            'Sometimes, if the MDB size gets large, C-OPT might get hung up here.
        '            Console.WriteLine("Starting to loop through list of core objects for processing...")
        '            Console.WriteLine("Started at: " & Now())
        '            Console.WriteLine("Note: Please check size of source database (to compact & repair) if you don't notice any progress from here on, after more than a few minutes.")
        '            Console.WriteLine("")
        '            For ctr = 0 To dt.Rows.Count - 1
        '                'For Each adoTable In adoTables
        '                For adoTableCtr = 0 To daoTableDefs.Count - 1 'adoTables.Count - 1
        '                    'adoTable = adoTables(adoTableCtr)
        '                    daoTableDef = daoTableDefs(adoTableCtr)
        '                    'If dt.Rows(ctr).Item("OBJ_NAME").ToString().Equals(adoTable.Name.ToString()) Then
        '                    If dt.Rows(ctr).Item("OBJ_NAME").ToString().Equals(daoTableDef.Name.ToString()) Then

        '                        'End If
        '                        Console.WriteLine("Processing table: " & daoTableDef.Name.ToString() & ". APM: " & GenUtils.GetAvailablePhysicalMemoryStr())
        '                        EntLib.COPT.Log.Log("Processing table: " & daoTableDef.Name.ToString() & ". APM: " & GenUtils.GetAvailablePhysicalMemoryStr())
        '                        msg = daoTableDef.Name.ToString() & "..."

        '                        'Console.Write("Reading source table: " & adoTable.Name.ToString())
        '                        'EntLib.COPT.Log.Log("Reading source table: " & adoTable.Name.ToString())
        '                        'ReDim arrayRestrictions(2)
        '                        'arrayRestrictions(0) = Nothing
        '                        'arrayRestrictions(1) = Nothing
        '                        'arrayRestrictions(2) = adoTable.Name

        '                        'list = New List(Of String)

        '                        'list.Add(String.Empty)
        '                        'list.Add(String.Empty)
        '                        'list.Add(adoTable.Name)

        '                        'Try
        '                        '    adoRSSchema = dalOleDb.ConnADO.OpenSchema(ADODB.SchemaEnum.adSchemaColumns, list.ToArray())
        '                        '    adoRSSchema.Sort = "ORDINAL_POSITION"
        '                        'Catch ex As Exception
        '                        '    EntLib.COPT.Log.Log("Error returning Schema.")
        '                        '    EntLib.COPT.Log.Log(ex.Message)
        '                        'End Try

        '                        'adoRSSchema = New ADODB.Recordset


        '                        sql = "SELECT * FROM [" & daoTableDef.Name & "] WHERE 0 = 1"
        '                        Try
        '                            'adoRSSchema = dalOleDbADO.ConnADO.Execute(sql)
        '                            daoRSSchema = dalDAO.DAODatabase.OpenRecordset(sql)
        '                            'arrayGetRows = CType(adoRSSchema.GetRows(), String()())
        '                        Catch ex As Exception
        '                            EntLib.COPT.Log.Log("Error returning Schema - GetRows()")
        '                            EntLib.COPT.Log.Log(ex.Message)
        '                        End Try

        '                        ReDim tableDefs(0)
        '                        ReDim idxDefs(0)
        '                        'adoCols = adoTable.Columns
        '                        daoFields = daoTableDef.Fields

        '                        For ctr1 = 0 To daoFields.Count - 1
        '                            'TODO:
        '                            'adoCol = daoFields(adoRSSchema.Fields(ctr1).Name.ToString())

        '                            'For Each daoField In daoFields
        '                            '    daoProperties = daoField.Properties
        '                            'Next

        '                            daoField = daoFields(ctr1)
        '                            'adoProperties = adoCol.Properties
        '                            daoProperties = daoField.Properties

        '                            'For Each adoProperty In adoProperties
        '                            '    Try
        '                            '        If String.IsNullOrEmpty(adoProperty.Value.ToString()) Then
        '                            '            Debug.Print(adoCol.Name.ToString() & " - " & adoProperty.Name & " - " & "")
        '                            '        Else
        '                            '            Debug.Print(adoCol.Name.ToString() & " - " & adoProperty.Name & " - " & adoProperty.Value.ToString())
        '                            '        End If
        '                            '    Catch ex As Exception
        '                            '        Debug.Print(adoCol.Name.ToString() & " - " & adoProperty.Name & " - " & "")
        '                            '    End Try
        '                            'Next
        '                            'Debug.Print(adoCol.DefinedSize.ToString())

        '                            If String.IsNullOrEmpty(tableDefs(0).fieldName) Then
        '                                tableDefs(0).fieldName = daoField.Name  'adoCol.Name
        '                                tableDefs(0).fieldType = daoField.Type.ToString() 'adoCol.Type.ToString()
        '                                tableDefs(0).fieldSize = 0
        '                            Else
        '                                ReDim Preserve tableDefs(UBound(tableDefs) + 1)
        '                                tableDefs(UBound(tableDefs)).fieldName = daoField.Name  'adoCol.Name
        '                                tableDefs(UBound(tableDefs)).fieldType = daoField.Type.ToString() 'adoCol.Type.ToString()
        '                                tableDefs(UBound(tableDefs)).fieldSize = 0
        '                            End If
        '                            'If tableDefs(UBound(tableDefs)).fieldType.Contains("Char") Then
        '                            '    'tableDefs(UBound(tableDefs)).fieldSize = CInt(adoCol.DefinedSize.ToString())
        '                            '    tableDefs(UBound(tableDefs)).fieldSize = CInt(daoField.Size.ToString()) 'CInt(adoCol.DefinedSize.ToString())
        '                            'End If
        '                            If _
        '                                tableDefs(UBound(tableDefs)).fieldType = dao.DataTypeEnum.dbChar.ToString() _
        '                                OrElse _
        '                                tableDefs(UBound(tableDefs)).fieldType = dao.DataTypeEnum.dbMemo.ToString() _
        '                                OrElse _
        '                                tableDefs(UBound(tableDefs)).fieldType = dao.DataTypeEnum.dbText.ToString() _
        '                            Then
        '                                tableDefs(UBound(tableDefs)).fieldSize = CInt(daoField.Size.ToString()) 'CInt(adoCol.DefinedSize.ToString())
        '                            End If

        '                            tableDefs(UBound(tableDefs)).fieldTypeNew = clsDAL.GetDAOToSQLDataType(tableDefs(UBound(tableDefs)).fieldType)

        '                            'tableDefs(UBound(tableDefs)).fieldTypeNew = clsDAL.GetADOToSQLDataType(tableDefs(UBound(tableDefs)).fieldType)
        '                            'tableDefs(UBound(tableDefs)).fieldIsIdentity = CBool(daoProperties("Autoincrement").Value.ToString()) 'CBool(adoProperties("Autoincrement").Value.ToString())
        '                            If dao.FieldAttributeEnum.dbAutoIncrField = (CInt(daoProperties("Attributes").Value) And dao.FieldAttributeEnum.dbAutoIncrField) Then
        '                                tableDefs(UBound(tableDefs)).fieldIsIdentity = True
        '                            Else
        '                                tableDefs(UBound(tableDefs)).fieldIsIdentity = False
        '                            End If
        '                            If tableDefs(UBound(tableDefs)).fieldIsIdentity Then
        '                                tableDefs(UBound(tableDefs)).fieldIsNullable = False
        '                            Else
        '                                tableDefs(UBound(tableDefs)).fieldIsNullable = True
        '                            End If
        '                            'tableDefs(UBound(tableDefs)).fieldIsNullable = CBool(adoProperties("Nullable").Value.ToString())
        '                            'If tableDefs(UBound(tableDefs)).fieldIsIdentity Then
        '                            ' tableDefs(UBound(tableDefs)).fieldIsNullable = False
        '                            'Else
        '                            'tableDefs(UBound(tableDefs)).fieldIsNullable = True
        '                            'End If
        '                        Next 'For ctr1 = 0 To adoCols.Count - 1

        '                        'EntLib.COPT.Log.Log(adoTable.Name & "..." & adoTable.Indexes.Count)

        '                        'adoIdxs = adoTable.Indexes
        '                        daoIndexes = daoTableDef.Indexes
        '                        ReDim listPK(0)
        '                        ReDim listIdx(0)
        '                        ReDim idxDefs(0)
        '                        'For Each adoIdx In adoIdxs
        '                        For Each daoIndex In daoIndexes

        '                            'Next
        '                            'Debug.Print(adoTable.Name.ToString() & " - " & adoIdx.Name.ToString())
        '                            'adoIdx.Columns.c
        '                            'adoProperties = adoIdx.Properties
        '                            'For Each adoProperty In adoProperties
        '                            '    Debug.Print(adoProperty.Name & " - " & adoProperty.Value.ToString())
        '                            'Next

        '                            If String.IsNullOrEmpty(idxDefs(0).idxName) Then
        '                                'idxDefs(UBound(idxDefs)).idxName = adoIdx.Name
        '                            Else
        '                                ReDim Preserve idxDefs(UBound(idxDefs) + 1)
        '                            End If
        '                            idxDefs(UBound(idxDefs)).idxName = daoIndex.Name  'adoIdx.Name

        '                            Debug.Print(CStr(daoIndex.Fields))
        '                            'adoCols = adoIdx.Columns
        '                            'daoCols = CType(daoIndex.Fields, dao.Fields)
        '                            daoCols = DirectCast(daoIndex.Fields, dao.Fields)

        '                            'ReDim Preserve idxDefs(UBound(idxDefs)).idxCols(adoCols.Count - 1)
        '                            ReDim Preserve idxDefs(UBound(idxDefs)).idxCols(daoCols.Count - 1)

        '                            'EntLib.COPT.Log.Log("   Index name: " & adoIdx.Name & " - No. of index columns: " & adoIdx.Columns.Count)
        '                            colCtr = -1
        '                            'For Each adoCol In adoCols
        '                            For Each daoCol In daoCols
        '                                colCtr += 1
        '                                'Debug.Print("TableName=" & adoTable.Name.ToString() & ", PK=" & adoIdx.PrimaryKey.ToString() & ", IdxName=" & adoIdx.Name.ToString() & ", ColName=" & adoCol.Name.ToString())
        '                                'If adoIdx.PrimaryKey Then
        '                                If daoIndex.Primary Then
        '                                    If String.IsNullOrEmpty(listPK(0)) Then
        '                                        'listPK(0) = adoCol.Name.ToString()
        '                                        listPK(0) = daoCol.Name.ToString()
        '                                    Else
        '                                        ReDim Preserve listPK(UBound(listPK) + 1)
        '                                        'listPK(UBound(listPK)) = adoCol.Name.ToString()
        '                                        listPK(UBound(listPK)) = daoCol.Name.ToString()
        '                                    End If
        '                                Else
        '                                    If String.IsNullOrEmpty(listIdx(0)) Then
        '                                        'listIdx(0) = adoCol.Name.ToString()
        '                                        listIdx(0) = daoCol.Name.ToString()
        '                                    Else
        '                                        ReDim Preserve listIdx(UBound(listIdx) + 1)
        '                                        'listIdx(UBound(listIdx)) = adoCol.Name.ToString()
        '                                        listIdx(UBound(listIdx)) = daoCol.Name.ToString()
        '                                    End If
        '                                End If

        '                                idxDefs(UBound(idxDefs)).idxCols( _
        '                                    colCtr _
        '                                    ) = daoCol.Name.ToString() 'adoCol.Name.ToString()

        '                                idxDefs(UBound(idxDefs)).isUnique = daoIndex.Unique  'adoIdx.Unique

        '                                'Console.WriteLine(adoTable.Name & "..." & adoTable.Indexes.Count & "...Index name: " & idxDefs(UBound(idxDefs)).idxName & ", Column name: " & adoCol.Name.ToString() & ", Is unique: " & idxDefs(UBound(idxDefs)).isUnique)
        '                                'EntLib.COPT.Log.Log(adoTable.Name & "..." & adoTable.Indexes.Count & "...Index name: " & idxDefs(UBound(idxDefs)).idxName & ", Column name: " & adoCol.Name.ToString() & ", Is unique: " & idxDefs(UBound(idxDefs)).isUnique)

        '                            Next 'For Each adoCol In adoCols
        '                        Next 'For Each adoIdx In adoIdxs

        '                        For ctr1 = 0 To listPK.Length - 1
        '                            For ctr2 = 0 To tableDefs.Length - 1
        '                                If tableDefs(ctr2).fieldName.Equals(listPK(ctr1)) Then
        '                                    tableDefs(ctr2).fieldIsNullable = False
        '                                    Exit For
        '                                End If
        '                            Next
        '                        Next

        '                        'Console.WriteLine("...Complete.")
        '                        'EntLib.COPT.Log.Log("...Complete.")

        '                        'Console.Write("Creating destination table: " & adoTable.Name.ToString())
        '                        'EntLib.COPT.Log.Log("Creating destination table: " & adoTable.Name.ToString())

        '                        'Create table + indexes in the SQL Express database - START
        '                        If CBool(dt.Rows(ctr).Item("Recreate").ToString()) Then
        '                            sql = "IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[" & daoTableDef.Name.ToString() & "]') AND type in (N'U')) DROP TABLE [dbo].[" & daoTableDef.Name.ToString() & "]"
        '                            Try
        '                                dalSql.Execute(sql, dt1)
        '                                createObject = True
        '                                'msg = msg & "...DROPPED..."
        '                                'EntLib.COPT.Log.Log(adoTable.Name.ToString() & "...DROP TABLE..." & CBool(dt.Rows(ctr).Item("Recreate").ToString()) & "...OK")
        '                            Catch ex As Exception
        '                                msg = msg & "Error in DROP TABLE..."
        '                                Console.WriteLine("Error creating table+indexes in SQLEXPRESS database.")
        '                                Console.WriteLine(ex.Message)
        '                                EntLib.COPT.Log.Log("Error creating table+indexes in SQLEXPRESS database.")
        '                                EntLib.COPT.Log.Log(ex.Message)
        '                                createObject = False
        '                            End Try
        '                        Else
        '                            sql = "SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[" & daoTableDef.Name.ToString() & "]') AND type in (N'U')"
        '                            'EntLib.COPT.Log.Log(ex.Message)
        '                            Try
        '                                dalSql.Execute(sql, dt1)
        '                                'EntLib.COPT.Log.Log(adoTable.Name.ToString() & "...DROP TABLE..." & CBool(dt.Rows(ctr).Item("Recreate").ToString()) & "...OK")
        '                            Catch ex As Exception
        '                                Console.WriteLine("Error executing SQL.")
        '                                Console.WriteLine(sql)
        '                                Console.WriteLine(ex.Message)
        '                                EntLib.COPT.Log.Log("Error executing SQL.")
        '                                EntLib.COPT.Log.Log(sql)
        '                                EntLib.COPT.Log.Log(ex.Message)
        '                            End Try

        '                            If dt1 Is Nothing Then
        '                                createObject = True
        '                            Else
        '                                If dt1.Rows.Count > 0 Then
        '                                    createObject = False
        '                                Else
        '                                    createObject = True
        '                                End If
        '                            End If
        '                        End If
        '                        If createObject Then
        '                            'Step 1: Drop table in destination database, if it already exists
        '                            sql = "IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[" & daoTableDef.Name.ToString() & "]') AND type in (N'U')) DROP TABLE [dbo].[" & daoTableDef.Name.ToString() & "]"
        '                            Try
        '                                dalSql.Execute(sql, dt1)
        '                                EntLib.COPT.Log.Log(daoTableDef.Name.ToString() & "...DROP TABLE..." & CBool(dt.Rows(ctr).Item("Recreate").ToString()) & "...OK")
        '                                msg = msg & "DROPTABLE-OK..."
        '                                Console.WriteLine("   DROP   TABLE...OK")
        '                            Catch ex As Exception
        '                                msg = msg & "DROPTABLE-Error..."
        '                                Console.WriteLine("   DROP TABLE...Error.")
        '                                Console.WriteLine("Error executing SQL.")
        '                                Console.WriteLine(sql)
        '                                Console.WriteLine(ex.Message)
        '                                EntLib.COPT.Log.Log("Error executing SQL.")
        '                                EntLib.COPT.Log.Log(sql)
        '                                EntLib.COPT.Log.Log(ex.Message)
        '                            End Try

        '                            'Step 2: Create table in destination database + create primary key
        '                            sql = "SET ANSI_NULLS ON"
        '                            Try
        '                                dalSql.Execute(sql, dt1)
        '                                EntLib.COPT.Log.Log(daoTableDef.Name.ToString() & "...SET ANSI_NULLS ON..." & "...OK")
        '                            Catch ex As Exception
        '                                msg = msg & "Error in SET ANSI_NULLS ON..."
        '                                Console.WriteLine("Error executing SQL.")
        '                                Console.WriteLine(sql)
        '                                Console.WriteLine(ex.Message)
        '                                EntLib.COPT.Log.Log("Error executing SQL.")
        '                                EntLib.COPT.Log.Log(sql)
        '                                EntLib.COPT.Log.Log(ex.Message)
        '                            End Try

        '                            sql = "SET QUOTED_IDENTIFIER ON"
        '                            Try
        '                                dalSql.Execute(sql, dt1)
        '                                EntLib.COPT.Log.Log(daoTableDef.Name.ToString() & "...SET QUOTED_IDENTIFIER ON..." & "...OK")
        '                            Catch ex As Exception
        '                                msg = msg & "Error in SET QUOTED_IDENTIFIER ON..."
        '                                Console.WriteLine("Error executing SQL.")
        '                                Console.WriteLine(sql)
        '                                Console.WriteLine(ex.Message)
        '                                EntLib.COPT.Log.Log("Error executing SQL.")
        '                                EntLib.COPT.Log.Log(sql)
        '                                EntLib.COPT.Log.Log(ex.Message)
        '                            End Try

        '                            sql = "CREATE TABLE [dbo].[" & daoTableDef.Name.ToString() & "] ("
        '                            For ctr1 = 0 To tableDefs.Length - 1
        '                                sql = sql & "[" & tableDefs(ctr1).fieldName & "] [" & tableDefs(ctr1).fieldTypeNew & "]" & IIf(tableDefs(ctr1).fieldSize > 0, "(" & tableDefs(ctr1).fieldSize & ") ", "").ToString() & IIf(tableDefs(ctr1).fieldIsIdentity, " IDENTITY(1,1)", "").ToString() & IIf(tableDefs(ctr1).fieldIsNullable, " NULL", " NOT NULL").ToString() & ","
        '                            Next
        '                            If String.IsNullOrEmpty(listPK(0)) Then
        '                                sql = Left(sql, Len(sql) - 1) & vbNewLine
        '                            Else
        '                                sql = sql & " CONSTRAINT [PK_" & daoTableDef.Name.ToString() & "] PRIMARY KEY CLUSTERED " & vbNewLine
        '                                sql = sql & "(" & vbNewLine
        '                                For ctr1 = 0 To listPK.Length - 1
        '                                    sql = sql & "[" & listPK(ctr1) & "] ASC,"
        '                                Next
        '                                sql = Left(sql, Len(sql) - 1) & vbNewLine
        '                                sql = sql & ")WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]" & vbNewLine
        '                            End If
        '                            sql = sql & ") ON [PRIMARY]" & vbNewLine
        '                            Try
        '                                dalSql.Execute(sql, dt1)
        '                                EntLib.COPT.Log.Log(daoTableDef.Name.ToString() & "...CREATE TABLE..." & "...OK")
        '                                msg = msg & "CREATETABLE-OK..."
        '                                Console.WriteLine("   CREATE TABLE...OK")
        '                            Catch ex As Exception
        '                                msg = msg & "Error in CREATETABLE..."
        '                                Debug.Print("Error creating TABLE - " & daoTableDef.Name.ToString())
        '                                Debug.Print(ex.Message)
        '                                Debug.Print(sql)
        '                                Debug.Print("")
        '                                Console.WriteLine("Error executing SQL.")
        '                                Console.WriteLine(sql)
        '                                Console.WriteLine(ex.Message)
        '                                EntLib.COPT.Log.Log("Error executing SQL.")
        '                                EntLib.COPT.Log.Log(sql)
        '                                EntLib.COPT.Log.Log(ex.Message)
        '                            End Try
        '                            'Step 3: Create nonclustered index on the table

        '                            'If Not String.IsNullOrEmpty(listIdx(0)) Then
        '                            '    sql = "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[" & adoTable.Name.ToString() & "]') AND name = N'" & "IDX_" & adoTable.Name.ToString() & "')"
        '                            '    sql = sql & "CREATE UNIQUE NONCLUSTERED INDEX [" & "IDX_" & adoTable.Name.ToString() & "] ON [dbo].[" & adoTable.Name.ToString() & "] "
        '                            '    sql = sql & "( "
        '                            '    For ctr1 = 0 To listIdx.Length - 1
        '                            '        sql = sql & "[" & listIdx(ctr1) & "] ASC,"
        '                            '    Next
        '                            '    sql = Left(sql, Len(sql) - 1) & vbNewLine
        '                            '    sql = sql & ")WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY] "
        '                            '    Try
        '                            '        dalSql.Execute(sql, dt1)
        '                            '    Catch ex As Exception
        '                            '        Debug.Print("Error creating NONCLUSTERED INDEX on TABLE - " & adoTable.Name.ToString())
        '                            '        Debug.Print(ex.Message)
        '                            '        Debug.Print(sql)
        '                            '        Debug.Print("")
        '                            '        Console.WriteLine("Error executing SQL.")
        '                            '        Console.WriteLine(sql)
        '                            '        Console.WriteLine(ex.Message)
        '                            '        EntLib.COPT.Log.Log("Error executing SQL.")
        '                            '        EntLib.COPT.Log.Log(sql)
        '                            '        EntLib.COPT.Log.Log(ex.Message)
        '                            '    End Try
        '                            'End If

        '                            If idxDefs IsNot Nothing Then
        '                                If Not String.IsNullOrEmpty(idxDefs(0).idxName) Then
        '                                    idxComposite = ""
        '                                    For ctr1 = 0 To idxDefs.Length - 1
        '                                        idxComposite = ""
        '                                        For ctr2 = 0 To idxDefs(ctr1).idxCols.Length - 1
        '                                            If String.IsNullOrEmpty(idxComposite) Then
        '                                                idxComposite = idxDefs(ctr1).idxCols(ctr2)
        '                                            Else
        '                                                idxComposite = idxComposite & "," & idxDefs(ctr1).idxCols(ctr2)
        '                                            End If
        '                                        Next

        '                                        sql = "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[" & daoTableDef.Name.ToString() & "]') AND name = N'" & "IDX_" & idxDefs(ctr1).idxName & "')"
        '                                        sql = sql & "CREATE " & IIf(idxDefs(ctr1).isUnique, "UNIQUE", "").ToString() & " NONCLUSTERED INDEX [" & "IDX_" & idxDefs(ctr1).idxName & "] ON [dbo].[" & daoTableDef.Name.ToString() & "] "
        '                                        sql = sql & "( "
        '                                        'For ctr1 = 0 To listIdx.Length - 1
        '                                        '    sql = sql & "[" & listIdx(ctr1) & "] ASC,"
        '                                        'Next
        '                                        'sql = Left(sql, Len(sql) - 1) & vbNewLine
        '                                        sql = sql & idxComposite
        '                                        sql = sql & ") WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY] "
        '                                        Try
        '                                            dalSql.Execute(sql, dt1)
        '                                            EntLib.COPT.Log.Log(daoTableDef.Name.ToString() & "..." & "CREATE " & IIf(idxDefs(ctr1).isUnique, "UNIQUE", "").ToString() & " NONCLUSTERED INDEX " & idxDefs(ctr1).idxName & "..." & "...OK")
        '                                        Catch ex As Exception
        '                                            Debug.Print("Error creating NONCLUSTERED INDEX on TABLE - " & daoTableDef.Name.ToString())
        '                                            Debug.Print(ex.Message)
        '                                            Debug.Print(sql)
        '                                            Debug.Print("")
        '                                            Console.WriteLine("Error executing SQL.")
        '                                            Console.WriteLine(sql)
        '                                            Console.WriteLine(ex.Message)
        '                                            EntLib.COPT.Log.Log("Error executing SQL.")
        '                                            EntLib.COPT.Log.Log(sql)
        '                                            EntLib.COPT.Log.Log(ex.Message)
        '                                        End Try

        '                                    Next
        '                                    msg = msg & "CREATEINDEX-OK..."
        '                                    Console.WriteLine("   CREATE INDEX...OK")
        '                                End If
        '                            End If

        '                        End If
        '                        'Console.Write("...Created")
        '                        'Create table + indexes in the SQL Express database - END

        '                        'Copy data - START
        '                        If _
        '                            CInt(dt.Rows(ctr).Item("TransferData").ToString()) = 1 _
        '                            Or _
        '                            CInt(dt.Rows(ctr).Item("TransferData").ToString()) = 3 _
        '                        Then
        '                            'Empty destination table
        '                            If noTruncate Then
        '                                'sql = "DELETE FROM [" & daoTableDef.Name.ToString() & "]"
        '                                sql = "EXEC dbo.asp_TruncateTable '" & daoTableDef.Name.ToString() & "'"
        '                            Else
        '                                sql = "TRUNCATE TABLE [" & daoTableDef.Name.ToString() & "]"
        '                            End If

        '                            Try
        '                                dalSql.Execute(sql, dt1)
        '                                EntLib.COPT.Log.Log(daoTableDef.Name.ToString() & "...TRUNCATE TABLE...OK")
        '                            Catch ex As Exception
        '                                Console.WriteLine("Error executing SQL.")
        '                                Console.WriteLine(sql)
        '                                Console.WriteLine(ex.Message)
        '                                EntLib.COPT.Log.Log("Error executing SQL.")
        '                                EntLib.COPT.Log.Log(sql)
        '                                EntLib.COPT.Log.Log(ex.Message)
        '                            End Try

        '                            Console.Write("   BULK    COPY...")
        '                            EntLib.COPT.Log.Log("Create ADO Recordset...")
        '                            sql = "SELECT * FROM [" & daoTableDef.Name.ToString() & "]"
        '                            Try
        '                                'dalOleDbADO.Execute(sql, dt1)
        '                                dalOleDb.Execute(sql, dt1)
        '                                If dalOleDbADO.LastErrorNo = -1 Then
        '                                    Console.WriteLine("ERROR. APM: " & GenUtils.GetAvailablePhysicalMemoryStr())
        '                                    Console.WriteLine("Error executing SQL.")
        '                                    Console.WriteLine("SELECT * FROM [" & daoTableDef.Name.ToString() & "]")
        '                                    Console.WriteLine(dalOleDbADO.LastErrorDesc)
        '                                    'Console.WriteLine(ex.Message)
        '                                    EntLib.COPT.Log.Log("Error executing SQL.")
        '                                    EntLib.COPT.Log.Log("SELECT * FROM [" & daoTableDef.Name.ToString() & "]")
        '                                    EntLib.COPT.Log.Log(dalOleDbADO.LastErrorDesc)
        '                                    EntLib.COPT.Log.Log("APM: " & GenUtils.GetAvailablePhysicalMemoryStr())
        '                                Else
        '                                    EntLib.COPT.Log.Log(daoTableDef.Name.ToString() & "...SELECT * FROM...OK")
        '                                    EntLib.COPT.Log.Log("Create ADO.NET Recordset...SUCCESS.")
        '                                End If
        '                            Catch ex As Exception
        '                                Console.WriteLine("Error executing SQL (using ADO.NET).")
        '                                Console.WriteLine(sql)
        '                                Console.WriteLine(ex.Message)
        '                                EntLib.COPT.Log.Log("Error executing SQL (using ADO.NET).")
        '                                EntLib.COPT.Log.Log(sql)
        '                                EntLib.COPT.Log.Log(ex.Message)
        '                            End Try
        '                            If dalOleDbADO.LastErrorNo <> -1 Then
        '                                sql = "dbo.[" & daoTableDef.Name.ToString() & "]"
        '                                Try
        '                                    'Console.Write("   BULK    COPY...")
        '                                    EntLib.COPT.Log.Log("BulkCopy...")
        '                                    ret = dalSql.Execute("BULKCOPY-MAP", dt1, sql)
        '                                    If ret = -1 Then
        '                                        EntLib.COPT.Log.Log(daoTableDef.Name.ToString() & "...BULKCOPY-MAP...ERROR")
        '                                        msg = msg & "...BULKCOPY-ERROR..."
        '                                        'Console.WriteLine("   BULK    COPY...OK")
        '                                        Console.WriteLine("ERROR. APM: " & GenUtils.GetAvailablePhysicalMemoryStr())
        '                                        Console.WriteLine("No. of rows: " & dt1.Rows.Count.ToString())
        '                                        Console.WriteLine("ERROR Desc:")
        '                                        Console.WriteLine(dalSql.LastErrorDesc)
        '                                        Console.WriteLine("Please check C-OPT.log for details.")
        '                                        EntLib.COPT.Log.Log("No. of rows: " & dt1.Rows.Count.ToString())
        '                                    Else
        '                                        EntLib.COPT.Log.Log(daoTableDef.Name.ToString() & "...BULKCOPY-MAP...OK (" & dt1.Rows.Count.ToString() & " rows)")
        '                                        msg = msg & "...BULKCOPY-OK..."
        '                                        EntLib.COPT.Log.Log("BulkCopy...SUCCESS.")
        '                                        'Console.WriteLine("   BULK    COPY...OK")
        '                                        Console.WriteLine("OK (" & dt1.Rows.Count.ToString() & " rows)")
        '                                    End If
        '                                Catch ex As Exception
        '                                    Console.WriteLine("Error executing SQL.")
        '                                    Console.WriteLine(sql)
        '                                    Console.WriteLine(ex.Message)
        '                                    Console.WriteLine("No. of rows: " & dt1.Rows.Count.ToString())
        '                                    EntLib.COPT.Log.Log("Error executing SQL.")
        '                                    EntLib.COPT.Log.Log(sql)
        '                                    EntLib.COPT.Log.Log(ex.Message)
        '                                    EntLib.COPT.Log.Log("No. of rows: " & dt1.Rows.Count.ToString())
        '                                End Try
        '                            End If
        '                            'Console.Write("...Populated with data.")
        '                        End If
        '                        'Copy data - END

        '                        'Console.WriteLine(msg)

        '                        Exit For
        '                    End If 'CInt(dt.Rows(ctr).Item("TransferData").ToString()) = 1 Or CInt(dt.Rows(ctr).Item("TransferData").ToString()) = 3
        '                Next 'For Each adoTable In adoTables
        '                If GenUtils.IsSwitchAvailable(switches, "/UseGC") Then
        '                    GenUtils.CollectGarbage()
        '                End If
        '                'dt1 = Nothing
        '            Next 'For ctr = 0 To dt.Rows.Count - 1


        '            Console.WriteLine("All TRANSFER tasks finished successfully.")
        '            EntLib.COPT.Log.Log("All TRANSFER tasks finished successfully.")

        '        End If 'If proceed Then 'proceed #2
        '    End If 'If proceed Then 'proceed #1

        '    dalOleDbADO = Nothing
        '    dalOleDb = Nothing
        '    dalSql = Nothing
        '    dt = Nothing
        '    dt1 = Nothing

        '    GenUtils.CollectGarbage()

        '    Console.WriteLine("TRANSFER took: " & DateDiff(DateInterval.Minute, startTime, Now()).ToString() & " min.")
        '    EntLib.COPT.Log.Log("TRANSFER took: " & DateDiff(DateInterval.Minute, startTime, Now()).ToString() & " min.")
        '    EntLib.COPT.Log.Log("********************")

        '    Return 0
        'End Function

        ''' <summary>
        ''' TRANSFER (TRF) - Transfers objects and data from source (MS Access) database to destination (SQL Server Express) database, as defined in the system table, tsysCoreObjects (in the source database).
        ''' </summary>
        ''' <param name="switches"></param>
        ''' <param name="srcDBName"></param>
        ''' <param name="srcDBFolder"></param>
        ''' <param name="destDBName"></param>
        ''' <param name="destDBFolder"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' RKP/04-24-12/v3.2.166
        ''' </remarks>
        Public Shared Function DB_TransferDAO( _
            ByRef switches() As String, _
            ByVal srcDBName As String, _
            ByVal srcDBFolder As String, _
            ByVal destDBName As String, _
            ByVal destDBFolder As String, _
            Optional srcDBType As e_DB = e_DB.e_db_ACCESS, _
            Optional destDBType As e_DB = e_DB.e_db_SQLSERVER_EXPRESS, _
            Optional pArray() As String = Nothing _
        ) As Integer

            Dim dalOleDbADO As New clsDAL 'ADO version 'RKP/04-20-12/v3.2.166
            Dim dalOleDb As New clsDAL 'ADO.NET version 'RKP/04-20-12/v3.2.166
            Dim dalDAO As New clsDAL
            Dim dalSql As New clsDAL
            Dim connOleDb As New OleDb.OleDbConnection
            Dim connSql As New SqlClient.SqlConnection
            'Dim adoCatalog As New ADOX.Catalog
            'Dim adoTables As ADOX.Tables
            'Dim adoTable As ADOX.Table
            'Dim adoIdxs As ADOX.Indexes
            'Dim adoIdx As ADOX.Index
            'Dim adoCols As ADOX.Columns
            'Dim adoCol As ADOX.Column
            'Dim adoProperties As ADOX.Properties
            Dim adoRSSchema As ADODB.Recordset
            Dim sql As String = ""
            Dim dt As New DataTable
            Dim dt1 As New DataTable
            Dim filePath As String = ""
            Dim listPK() As String
            Dim listIdx() As String
            Dim ctr As Integer = 0
            Dim ctr1 As Integer = 0
            Dim ctr2 As Integer = 0
            Dim tableDefs() As STRUCT_TABLEDEF
            Dim idxDefs() As STRUCT_IDX
            Dim createObject As Boolean = False
            Dim proceed As Boolean = False
            Dim srcFilePath As String = ""
            Dim destFilePath As String = ""
            Dim msg As String = ""
            Dim colCtr As Integer = 0
            Dim idxComposite As String = ""
            Dim adoTableCtr As Integer = 0
            Dim ret As Integer
            'Dim startTime As Date
            Dim startTime As Long = My.Computer.Clock.TickCount
            Dim startTimeD As Date = Now() 'RKP/04-30-12/v3.2.167
            Dim tableName As String
            Dim fieldName As String
            Dim daoTableDefs As Dao.TableDefs
            Dim daoTableDef As Dao.TableDef
            Dim daoField As Dao.Field
            Dim daoProperties As Dao.Properties
            'Dim daoProperty As dao.Property
            Dim daoIndexes As Dao.Indexes
            Dim daoIndex As Dao.Index
            'Dim daoIdxFields As dao.Fields
            Dim daoIdxField As Dao.Field
            Dim errCount As Integer = 0
            Dim noTruncate As Boolean = GenUtils.IsSwitchAvailable(switches, "/UseTruncateSP")

            'startTime = My.Computer.Clock.TickCount
            proceed = True
            If proceed Then 'proceed #1

                'If EntLib.COPT.GenUtils.IsSwitchAvailable(switches, "/NoPrompt") Then
                '    proceed = True
                'Else
                '    Console.WriteLine()
                '    Console.WriteLine("Source File Path: " & srcFilePath)
                '    Console.WriteLine("Destination File Path: " & destFilePath)
                '    Console.Write("Do you want to continue? (Y/N): ")
                '    input = Console.ReadLine()
                '    If input.Trim.ToUpper.Equals("Y") Then
                '        proceed = True
                '    Else
                '        proceed = False
                '    End If
                'End If

                If proceed Then 'proceed #2

                    'MessageBox.Show("About to create log entry.", "C-OPT", MessageBoxButtons.OK, MessageBoxIcon.Information)

                    'Console.Write("Do you want to continue? (Y/N): ")
                    'Input = Console.ReadLine()
                    'If Input.Trim.ToUpper.Equals("Y") Then
                    '    proceed = True
                    'Else
                    '    proceed = False
                    'End If

                    EntLib.COPT.Log.Log("")
                    EntLib.COPT.Log.Log("********************")
                    EntLib.COPT.Log.Log("TRANSFER (TRF) - Started at: " & Now())

                    EntLib.COPT.Log.Log("Source File Path: " & srcFilePath)
                    EntLib.COPT.Log.Log("Destination File Path: " & destFilePath)

                    filePath = srcDBFolder & "\" & srcDBName & IIf(srcDBName.Trim().ToUpper().EndsWith(".MDB"), "", ".mdb").ToString() '& ".mdb"  '"C:\OPTMODELS\SRC1\SRC1.mdb"
                    dalOleDbADO.UseADO = True
                    Try
                        dalOleDbADO.Connect(clsDAL.e_DB.e_db_ACCESS, clsDAL.e_ConnType.e_connType_C, Nothing, connOleDb, Nothing, Nothing, Nothing, True, filePath, "C:\OPTMODELS\C-OPTSYS\System.mdw.copt")
                        EntLib.COPT.Log.Log("Connected to source database successfully (using ADO).")
                        Console.WriteLine("Connected to source database successfully (using ADO).")
                    Catch ex As Exception
                        Console.WriteLine("Error connecting to source database (using ADO).")
                        EntLib.COPT.Log.Log("Error connecting to source database (using ADO).")
                        EntLib.COPT.Log.Log(ex.Message)
                        errCount += 1
                    End Try

                    'RKP/04-20-12/v3.2.166
                    'This is the ADO.NET version of the source connection object.
                    'The purpose of this object is to avoid the use of "ConvertADORecordsetToDataTable", which avoids duplicate in-memory objects (adoRS and dt).
                    'If not for this object, memory utilization would be adversely affected, especially for large tables (~ million rows).
                    dalOleDb.UseADO = False
                    Try
                        dalOleDb.Connect(clsDAL.e_DB.e_db_ACCESS, clsDAL.e_ConnType.e_connType_C, Nothing, connOleDb, Nothing, Nothing, Nothing, True, filePath, "C:\OPTMODELS\C-OPTSYS\System.mdw.copt")
                        EntLib.COPT.Log.Log("Connected to source database successfully (using ADO.NET).")
                        Console.WriteLine("Connected to source database successfully (using ADO.NET).")
                    Catch ex As Exception
                        Console.WriteLine("Error connecting to source database (using ADO.NET).")
                        EntLib.COPT.Log.Log("Error connecting to source database (using ADO.NET).")
                        EntLib.COPT.Log.Log(ex.Message)
                        errCount += 1
                    End Try

                    dalDAO.UseDAO = True
                    Try
                        dalDAO.Connect(clsDAL.e_DB.e_db_ACCESS, clsDAL.e_ConnType.e_connType_C, Nothing, Nothing, Nothing, Nothing, Nothing, True, filePath, "C:\OPTMODELS\C-OPTSYS\System.mdw.copt")
                        EntLib.COPT.Log.Log("Connected to source database successfully (using DAO).")
                        Console.WriteLine("Connected to source database successfully (using DAO).")
                    Catch ex As Exception
                        Console.WriteLine("Error connecting to source database (using DAO).")
                        EntLib.COPT.Log.Log("Error connecting to source database (using DAO).")
                        EntLib.COPT.Log.Log(ex.Message)
                        errCount += 1
                    End Try


                    Console.WriteLine("Reading catalog from source database...")
                    Console.WriteLine("Started at: " & Now() & ". ")

                    'ret = dalOleDb.Execute("SELECT * FROM qsysMSysObjects WHERE OBJ_TYPE_ID = 1", dtADOTables)

                    'RKP/02-12-12/v3.0.157
                    'Added additional error-checking to prevent C-OPT from crashing at this point.
                    Try
                        'adoCatalog = dalOleDbADO.ConnADOCatalog
                        'adoTables = adoCatalog.Tables
                        'ret = adoTables.Count

                        daoTableDefs = dalDAO.DAODatabase.TableDefs
                        ret = daoTableDefs.Count
                    Catch ex As Exception
                        Console.WriteLine("Error.")
                        Console.WriteLine("adoTables.Count = 0.")
                        Console.WriteLine("Please reach out to BMOS for resolution.")
                        Console.WriteLine("The source database has objects in it (like, linked tables, etc) that are incompatible with C-OPT.")
                        Console.WriteLine("You must remove all incompatible objects in the source database (like linked tables to external databases, invalid queries, etc) and try the operation again.")
                        EntLib.COPT.Log.Log("Error. adoTables.Count = 0.")
                        EntLib.COPT.Log.Log("The source database has objects in it (like, linked tables, etc) that are incompatible with C-OPT.")
                        EntLib.COPT.Log.Log("If, for example, the database has linked tables to, say, a MySQL database, C-OPT will have difficulty connecting to it.")
                        EntLib.COPT.Log.Log("You must remove all incompatible objects in the source database (like linked tables to external databases, invalid queries, etc) and try the operation again.")
                        EntLib.COPT.Log.Log(ex.Message)
                        errCount += 1
                        Return -1
                    End Try

                    If ret = 0 Then
                        Console.WriteLine("Error.")
                        Console.WriteLine("adoTables.Count = 0.")
                        Console.WriteLine("Please reach out to BMOS for resolution.")
                        EntLib.COPT.Log.Log("Error. adoTables.Count = 0.")
                        errCount += 1
                        Return -1
                    Else
                        Console.WriteLine("Ended at: " & Now().ToString() & " (" & GenUtils.FormatTime(startTime) & ".)")
                        EntLib.COPT.Log.Log("Reading Catalog took: " & GenUtils.FormatTime(startTime) & ".")
                    End If

                    'Console.WriteLine(adoTables.Count.ToString() & " tables & views found in database.")

                    'http://support.microsoft.com/kb/299484
                    'adoRSSchema = dalOleDb.ConnADO.OpenSchema(ADODB.SchemaEnum.adSchemaColumns, Array(Nothing, Nothing, "Products"))

                    'dalSql.Connect(clsDAL.e_DB.e_db_SQLSERVER, clsDAL.e_ConnType.e_connType_F, Nothing, Nothing, connSql, Nothing, Nothing, "S02ASQLNPD01", "BMOS")
                    'Works:
                    'dalSql.Connect(clsDAL.e_DB.e_db_SQLSERVER, clsDAL.e_ConnType.e_connType_I, Nothing, Nothing, connSql, Nothing, Nothing, ".\SQLEXPRESS", "C:\OPTMODELS\SRC1SQL\SRC1SQL.mdf", "SRC1SQL")
                    'EntLib.COPT.Log.Log("")

                    Select Case destDBType
                        Case e_DB.e_db_SQLSERVER_EXPRESS
                            Try
                                dalSql.Connect( _
                                        clsDAL.e_DB.e_db_SQLSERVER, _
                                        clsDAL.e_ConnType.e_connType_I, _
                                        Nothing, _
                                        Nothing, _
                                        connSql, _
                                        Nothing, _
                                        Nothing, _
                                        True, _
                                        ".\SQLEXPRESS", _
                                        destDBFolder & "\" & destDBName & _
                                        IIf(destDBName.Trim().ToUpper().EndsWith(".MDF"), "", ".mdf").ToString(), _
                                        IIf(destDBName.Trim().ToUpper().EndsWith(".MDF"), _
                                        Left(destDBName.Trim(), Len(destDBName.Trim()) - 4), destDBName.Trim()).ToString() _
                                )
                                EntLib.COPT.Log.Log("Connected to destination database successfully.")
                                Console.WriteLine("Connected to destination database successfully.")
                            Catch ex As Exception
                                Console.WriteLine("Error connecting to destination database.")
                                EntLib.COPT.Log.Log("Error connecting to destination database.")
                                EntLib.COPT.Log.Log(ex.Message)
                                errCount += 1
                            End Try
                        Case e_DB.e_db_SQLSERVER_ENTERPRISE
                            Try
                                'Driver={SQL Server Native Client 10.0};Server=myServerAddress;Database=myDataBase;Trusted_Connection=yes;
                                dalSql.Connect( _
                                        clsDAL.e_DB.e_db_SQLSERVER, _
                                        clsDAL.e_ConnType.e_connType_Z, _
                                        Nothing, _
                                        Nothing, _
                                        connSql, _
                                        Nothing, _
                                        Nothing, _
                                        True, _
                                        GenUtils.ConfigRead(destDBName) _
                                )
                                EntLib.COPT.Log.Log("Connected to destination database successfully.")
                                Console.WriteLine("Connected to destination database successfully.")
                            Catch ex As Exception
                                Console.WriteLine("Error connecting to destination database.")
                                EntLib.COPT.Log.Log("Error connecting to destination database.")
                                EntLib.COPT.Log.Log(ex.Message)
                                errCount += 1
                            End Try
                        Case Else
                            Try
                                dalSql.Connect( _
                                        clsDAL.e_DB.e_db_SQLSERVER, _
                                        clsDAL.e_ConnType.e_connType_I, _
                                        Nothing, _
                                        Nothing, _
                                        connSql, _
                                        Nothing, _
                                        Nothing, _
                                        True, _
                                        ".\SQLEXPRESS", _
                                        destDBFolder & "\" & destDBName & _
                                        IIf(destDBName.Trim().ToUpper().EndsWith(".MDF"), "", ".mdf").ToString(), _
                                        IIf(destDBName.Trim().ToUpper().EndsWith(".MDF"), _
                                        Left(destDBName.Trim(), Len(destDBName.Trim()) - 4), destDBName.Trim()).ToString() _
                                )
                                EntLib.COPT.Log.Log("Connected to destination database successfully.")
                                Console.WriteLine("Connected to destination database successfully.")
                            Catch ex As Exception
                                Console.WriteLine("Error connecting to destination database.")
                                EntLib.COPT.Log.Log("Error connecting to destination database.")
                                EntLib.COPT.Log.Log(ex.Message)
                                errCount += 1
                            End Try
                    End Select



                    'sql = "SELECT * FROM tsysCOL"
                    'dalSql.Execute(sql, dt)

                    'sql = "SELECT * FROM tsysCoreObjects WHERE OBJ_TYPE_ID = 1 AND [Active] = True" 'AND OBJ_NAME = '" & "tMTXtbli00_C_BBDSALES" & "'"
                    sql = "SELECT * FROM tsysCoreObjects WHERE OBJ_TYPE_ID = 1 AND [Active] <> 0 " '& "AND OBJ_NAME = '" & "tsysCOL" & "'"
                    Console.Write("Reading list of core objects...")
                    EntLib.COPT.Log.Log(sql)
                    Try
                        dalOleDbADO.Execute(sql, dt)
                        Console.WriteLine("complete.")
                    Catch ex As Exception
                        Console.WriteLine("error.")
                        Console.WriteLine("Error:")
                        Console.WriteLine(ex.Message)
                        Console.WriteLine("SQL:" & vbNewLine & sql)
                        EntLib.COPT.Log.Log("Error:")
                        EntLib.COPT.Log.Log(ex.Message)
                        EntLib.COPT.Log.Log("SQL:" & vbNewLine & sql)
                        errCount += 1
                    End Try

                    'sql = "TRUNCATE TABLE dbo.ZCOR1347_ENT"
                    'dalSql.Execute(sql, dt)

                    'sql = "BULKCOPY"
                    'dalSql.Execute(sql, dt, "dbo.ZCOR1347_ENT")

                    'RKP/01-06-12/v3.0.157
                    'Sometimes, if the MDB size gets large, C-OPT might get hung up here.
                    Console.WriteLine("Starting to loop through list of core objects for processing...")
                    Console.WriteLine("Started at: " & Now())
                    Console.WriteLine("Note: Please check size of source database (to compact & repair) if you don't notice any progress from here on, after more than a few minutes.")
                    Console.WriteLine("")
                    For ctr = 0 To dt.Rows.Count - 1
                        'For Each adoTable In adoTables
                        'For adoTableCtr = 0 To daoTableDefs.Count - 1 'adoTables.Count - 1
                        adoTableCtr = 0
                        For Each daoTableDef In daoTableDefs
                            adoTableCtr += 1
                            'adoTable = adoTables(adoTableCtr)
                            'adoTable = adoTables(tableName)
                            tableName = daoTableDef.Name

                            'Console.WriteLine(tableName & ": Reading from Catalog took: " & Now().ToString() & " (" & DateDiff(DateInterval.Minute, startTime, Now()).ToString() & " min.)")
                            'EntLib.COPT.Log.Log(tableName & ": Reading from Catalog took: " & DateDiff(DateInterval.Minute, startTime, Now()).ToString() & " min.")

                            If dt.Rows(ctr).Item("OBJ_NAME").ToString().Trim().ToUpper().Equals(tableName.Trim().ToUpper()) Then

                                Console.WriteLine("Processing table: " & tableName & ". APM: " & GenUtils.GetAvailablePhysicalMemoryStr())
                                EntLib.COPT.Log.Log("")
                                EntLib.COPT.Log.Log("Processing table: " & tableName & ". APM: " & GenUtils.GetAvailablePhysicalMemoryStr())
                                msg = tableName & "..."

                                'Console.Write("Reading source table: " & tableName)
                                'EntLib.COPT.Log.Log("Reading source table: " & tableName)
                                'ReDim arrayRestrictions(2)
                                'arrayRestrictions(0) = Nothing
                                'arrayRestrictions(1) = Nothing
                                'arrayRestrictions(2) = adoTable.Name

                                'list = New List(Of String)

                                'list.Add(String.Empty)
                                'list.Add(String.Empty)
                                'list.Add(adoTable.Name)

                                'Try
                                '    adoRSSchema = dalOleDb.ConnADO.OpenSchema(ADODB.SchemaEnum.adSchemaColumns, list.ToArray())
                                '    adoRSSchema.Sort = "ORDINAL_POSITION"
                                'Catch ex As Exception
                                '    EntLib.COPT.Log.Log("Error returning Schema.")
                                '    EntLib.COPT.Log.Log(ex.Message)
                                'End Try

                                adoRSSchema = New ADODB.Recordset

                                sql = "SELECT * FROM [" & tableName & "] WHERE 0 = 1"
                                Try
                                    adoRSSchema = dalOleDbADO.ConnADO.Execute(sql)
                                    'arrayGetRows = CType(adoRSSchema.GetRows(), String()())
                                Catch ex As Exception
                                    EntLib.COPT.Log.Log("Error returning Schema - " & "SELECT * FROM [" & tableName & "] WHERE 0 = 1")
                                    EntLib.COPT.Log.Log(ex.Message)
                                End Try

                                ReDim tableDefs(0)
                                ReDim idxDefs(0)
                                'adoCols = adoTable.Columns
                                'adoCols = adoRSSchema.Fields
                                'ctr1 = adoRSSchema.Fields.Count
                                'For ctr1 = 0 To adoCols.Count - 1
                                'For ctr1 = 0 To daoTableDef.Fields.Count - 1
                                For Each daoField In daoTableDef.Fields
                                    'adoCol = adoCols(adoRSSchema.Fields(ctr1).Name.ToString())
                                    fieldName = daoField.Name
                                    daoProperties = daoField.Properties

                                    'adoProperties = adoCol.Properties

                                    'For Each adoProperty In adoProperties
                                    '    Try
                                    '        If String.IsNullOrEmpty(adoProperty.Value.ToString()) Then
                                    '            Debug.Print(adoCol.Name.ToString() & " - " & adoProperty.Name & " - " & "")
                                    '        Else
                                    '            Debug.Print(adoCol.Name.ToString() & " - " & adoProperty.Name & " - " & adoProperty.Value.ToString())
                                    '        End If
                                    '    Catch ex As Exception
                                    '        Debug.Print(adoCol.Name.ToString() & " - " & adoProperty.Name & " - " & "")
                                    '    End Try
                                    'Next
                                    'If tableName = "tMTXtbli00_C_BBDSALES" Then
                                    'EntLib.COPT.Log.Log("")
                                    'EntLib.COPT.Log.Log("")
                                    'msg = ""
                                    'For Each daoProperty In daoProperties
                                    '    Try
                                    '        msg = tableName & " - " & daoField.Name.ToString() & " - " & daoProperty.Name.ToString()
                                    '        'Debug.Print(tableName & " - " & daoField.Name.ToString() & " - " & daoProperty.Name.ToString())
                                    '        'Debug.Print("      - " & daoProperty.Value.ToString())
                                    '        msg = msg & " - " & daoProperty.Value.ToString()
                                    '        'If String.IsNullOrEmpty(daoProperty.Value.ToString()) Then
                                    '        '    Debug.Print(daoField.Name.ToString() & " - " & daoProperty.Name & " - " & "")
                                    '        'Else
                                    '        '    Debug.Print(daoField.Name.ToString() & " - " & daoProperty.Name & " - " & daoProperty.Value.ToString())
                                    '        'End If
                                    '    Catch ex As Exception
                                    '        'Debug.Print("      - N/A")
                                    '        msg = msg & " - N/A"
                                    '    Finally
                                    '        EntLib.COPT.Log.Log(msg)
                                    '    End Try
                                    'Next
                                    'End If
                                    If String.IsNullOrEmpty(tableDefs(0).fieldName) Then
                                        tableDefs(0).fieldName = fieldName 'adoCol.Name
                                        tableDefs(0).fieldType = daoField.Type.ToString() 'adoCol.Type.ToString()
                                        tableDefs(0).fieldSize = 0
                                    Else
                                        ReDim Preserve tableDefs(UBound(tableDefs) + 1)
                                        tableDefs(UBound(tableDefs)).fieldName = fieldName 'adoCol.Name
                                        tableDefs(UBound(tableDefs)).fieldType = daoField.Type.ToString() 'adoCol.Type.ToString()
                                        tableDefs(UBound(tableDefs)).fieldSize = 0
                                    End If
                                    If _
                                        CInt(tableDefs(UBound(tableDefs)).fieldType) = 18 _
                                        OrElse _
                                        CInt(tableDefs(UBound(tableDefs)).fieldType) = 12 _
                                        OrElse _
                                        CInt(tableDefs(UBound(tableDefs)).fieldType) = 10 _
                                    Then
                                        'dbChar = 18; dbMemo = 12; dbText = 10
                                        tableDefs(UBound(tableDefs)).fieldSize = daoField.Size  'CInt(adoCol.DefinedSize.ToString())
                                    Else
                                        tableDefs(UBound(tableDefs)).fieldSize = 0 'daoField.Size
                                    End If
                                    'tableDefs(UBound(tableDefs)).fieldSize = daoField.Size
                                    'If tableDefs(UBound(tableDefs)).fieldType.Contains("Char") Then
                                    '    tableDefs(UBound(tableDefs)).fieldSize = CInt(adoCol.DefinedSize.ToString())
                                    'End If
                                    tableDefs(UBound(tableDefs)).fieldTypeNew = clsDAL.GetDAOToSQLDataType(tableDefs(UBound(tableDefs)).fieldType)
                                    'tableDefs(UBound(tableDefs)).fieldIsIdentity = CBool(adoProperties("Autoincrement").Value.ToString())
                                    'If CInt(daoField.Attributes.ToString()) And CInt(dao.FieldAttributeEnum.dbAutoIncrField.ToString()) Then

                                    'End If
                                    If (CType(daoField.Attributes, Int32) And CType(Dao.FieldAttributeEnum.dbAutoIncrField, Int32)) = CType(Dao.FieldAttributeEnum.dbAutoIncrField, Int32) Then
                                        tableDefs(UBound(tableDefs)).fieldIsIdentity = True
                                        tableDefs(UBound(tableDefs)).fieldIsNullable = False
                                    Else
                                        tableDefs(UBound(tableDefs)).fieldIsIdentity = False
                                        tableDefs(UBound(tableDefs)).fieldIsNullable = True
                                    End If
                                    'If tableDefs(UBound(tableDefs)).fieldIsIdentity Then
                                    '    tableDefs(UBound(tableDefs)).fieldIsNullable = False
                                    'Else
                                    '    tableDefs(UBound(tableDefs)).fieldIsNullable = True
                                    'End If
                                    'tableDefs(UBound(tableDefs)).fieldIsNullable = CBool(adoProperties("Nullable").Value.ToString())
                                    'If tableDefs(UBound(tableDefs)).fieldIsIdentity Then
                                    ' tableDefs(UBound(tableDefs)).fieldIsNullable = False
                                    'Else
                                    'tableDefs(UBound(tableDefs)).fieldIsNullable = True
                                    'End If
                                Next 'For ctr1 = 0 To adoCols.Count - 1
                                'EntLib.COPT.Log.Log("")
                                'EntLib.COPT.Log.Log("")
                                'EntLib.COPT.Log.Log(adoTable.Name & "..." & adoTable.Indexes.Count)

                                daoIndexes = daoTableDef.Indexes

                                'adoIdxs = adoTable.Indexes
                                ReDim listPK(0)
                                ReDim listIdx(0)
                                ReDim idxDefs(0)
                                'For Each adoIdx In adoIdxs
                                For Each daoIndex In daoIndexes
                                    'Debug.Print(tableName & " - " & adoIdx.Name.ToString())
                                    'adoIdx.Columns.c
                                    'adoProperties = adoIdx.Properties
                                    'For Each adoProperty In adoProperties
                                    '    Debug.Print(adoProperty.Name & " - " & adoProperty.Value.ToString())
                                    'Next

                                    If String.IsNullOrEmpty(idxDefs(0).idxName) Then
                                        'idxDefs(UBound(idxDefs)).idxName = adoIdx.Name
                                    Else
                                        ReDim Preserve idxDefs(UBound(idxDefs) + 1)
                                    End If
                                    idxDefs(UBound(idxDefs)).idxName = daoIndex.Name  'adoIdx.Name

                                    'adoCols = adoIdx.Columns

                                    'daoIdxFields = CType(daoIndex.Fields, dao.Fields)
                                    'For Each daoIdxField In daoIndex.Fields
                                    '    Debug.Print(daoIdxField.Name)

                                    'Next

                                    'ReDim Preserve idxDefs(UBound(idxDefs)).idxCols(adoCols.Count - 1)
                                    'CInt(daoIndex.Fields.Count.ToString())

                                    'RKP/04-24-12/v3.2.166
                                    'The following line of code is the reason why "Option Strict" is turned OFF for this project.
                                    ReDim Preserve idxDefs(UBound(idxDefs)).idxCols(CInt(daoIndex.Fields.Count.ToString()) - 1)
                                    'EntLib.COPT.Log.Log("   Index name: " & adoIdx.Name & " - No. of index columns: " & adoIdx.Columns.Count)
                                    colCtr = -1

                                    'For Each adoCol In adoCols

                                    'RKP/04-24-12/v3.2.166
                                    'The following line of code is the reason why "Option Strict" is turned OFF for this project.
                                    For Each daoIdxField In daoIndex.Fields
                                        'For Each daoIdxField In daoIdxFields
                                        '    Debug.Print(daoIdxField.Name)
                                        'Next
                                        colCtr += 1
                                        'Debug.Print("TableName=" & tableName & ", PK=" & adoIdx.PrimaryKey.ToString() & ", IdxName=" & adoIdx.Name.ToString() & ", ColName=" & adoCol.Name.ToString())
                                        'If (CType(daoIdxField.Attributes, Int32) And CType(dao.FieldAttributeEnum.dbAutoIncrField, Int32)) = CType(dao.FieldAttributeEnum.dbAutoIncrField, Int32) Then
                                        If daoIndex.Primary Then
                                            If String.IsNullOrEmpty(listPK(0)) Then
                                                listPK(0) = daoIdxField.Name.ToString() 'adoCol.Name.ToString()
                                            Else
                                                ReDim Preserve listPK(UBound(listPK) + 1)
                                                listPK(UBound(listPK)) = daoIdxField.Name.ToString() 'adoCol.Name.ToString()
                                            End If
                                        Else
                                            If String.IsNullOrEmpty(listIdx(0)) Then
                                                listIdx(0) = daoIdxField.Name.ToString() 'adoCol.Name.ToString()
                                            Else
                                                ReDim Preserve listIdx(UBound(listIdx) + 1)
                                                listIdx(UBound(listIdx)) = daoIdxField.Name.ToString() 'adoCol.Name.ToString()
                                            End If
                                        End If
                                        'If adoIdx.PrimaryKey Then
                                        '    If String.IsNullOrEmpty(listPK(0)) Then
                                        '        listPK(0) = adoCol.Name.ToString()
                                        '    Else
                                        '        ReDim Preserve listPK(UBound(listPK) + 1)
                                        '        listPK(UBound(listPK)) = adoCol.Name.ToString()
                                        '    End If
                                        'Else
                                        '    If String.IsNullOrEmpty(listIdx(0)) Then
                                        '        listIdx(0) = adoCol.Name.ToString()
                                        '    Else
                                        '        ReDim Preserve listIdx(UBound(listIdx) + 1)
                                        '        listIdx(UBound(listIdx)) = adoCol.Name.ToString()
                                        '    End If
                                        'End If

                                        idxDefs(UBound(idxDefs)).idxCols( _
                                            colCtr _
                                            ) = daoIdxField.Name.ToString() 'adoCol.Name.ToString()

                                        idxDefs(UBound(idxDefs)).isUnique = daoIndex.Unique  'adoIdx.Unique

                                        'Console.WriteLine(adoTable.Name & "..." & adoTable.Indexes.Count & "...Index name: " & idxDefs(UBound(idxDefs)).idxName & ", Column name: " & adoCol.Name.ToString() & ", Is unique: " & idxDefs(UBound(idxDefs)).isUnique)
                                        'EntLib.COPT.Log.Log(adoTable.Name & "..." & adoTable.Indexes.Count & "...Index name: " & idxDefs(UBound(idxDefs)).idxName & ", Column name: " & adoCol.Name.ToString() & ", Is unique: " & idxDefs(UBound(idxDefs)).isUnique)

                                    Next 'For Each adoCol In adoCols
                                Next 'For Each adoIdx In adoIdxs

                                For ctr1 = 0 To listPK.Length - 1
                                    For ctr2 = 0 To tableDefs.Length - 1
                                        If tableDefs(ctr2).fieldName.Equals(listPK(ctr1)) Then
                                            tableDefs(ctr2).fieldIsNullable = False
                                            Exit For
                                        End If
                                    Next
                                Next

                                'Console.WriteLine("...Complete.")
                                'EntLib.COPT.Log.Log("...Complete.")

                                'Console.Write("Creating destination table: " & tableName)
                                'EntLib.COPT.Log.Log("Creating destination table: " & tableName)

                                'Create table + indexes in the SQL Express database - START
                                If CBool(dt.Rows(ctr).Item("Recreate").ToString()) Then
                                    sql = "IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[" & tableName & "]') AND type in (N'U')) DROP TABLE [dbo].[" & tableName & "]"
                                    Try
                                        dalSql.Execute(sql, dt1)
                                        If dalSql.LastErrorNo >= 0 Then
                                            createObject = True
                                        Else

                                            msg = msg & "Error in DROP TABLE..."
                                            Console.WriteLine("Error creating table+indexes in SQLEXPRESS database.")

                                            EntLib.COPT.Log.Log("Error creating table+indexes in SQLEXPRESS database.")
                                            EntLib.COPT.Log.Log(sql)
                                            createObject = False
                                            errCount += 1
                                        End If
                                        'createObject = True
                                        'msg = msg & "...DROPPED..."
                                        'EntLib.COPT.Log.Log(tableName & "...DROP TABLE..." & CBool(dt.Rows(ctr).Item("Recreate").ToString()) & "...OK")
                                    Catch ex As Exception
                                        msg = msg & "Error in DROP TABLE..."
                                        Console.WriteLine("Error creating table+indexes in SQLEXPRESS database.")
                                        Console.WriteLine(ex.Message)
                                        EntLib.COPT.Log.Log("Error creating table+indexes in SQLEXPRESS database.")
                                        EntLib.COPT.Log.Log(ex.Message)
                                        createObject = False
                                        errCount += 1
                                    End Try
                                Else
                                    'sql = "SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[" & tableName & "]') AND type in (N'U')"
                                    sql = "SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE UPPER(LTRIM(RTRIM(TABLE_NAME))) = '" & tableName.Trim().ToUpper() & "' "
                                    'EntLib.COPT.Log.Log(ex.Message)
                                    Try
                                        dalSql.Execute(sql, dt1)
                                        If dalSql.LastErrorNo = -1 Then
                                            Console.WriteLine("Error executing SQL.")
                                            Console.WriteLine(sql)

                                            EntLib.COPT.Log.Log("Error executing SQL.")
                                            EntLib.COPT.Log.Log(sql)
                                        Else

                                        End If
                                        'EntLib.COPT.Log.Log(tableName & "...DROP TABLE..." & CBool(dt.Rows(ctr).Item("Recreate").ToString()) & "...OK")
                                    Catch ex As Exception
                                        Console.WriteLine("Error executing SQL.")
                                        Console.WriteLine(sql)
                                        Console.WriteLine(ex.Message)
                                        EntLib.COPT.Log.Log("Error executing SQL.")
                                        EntLib.COPT.Log.Log(sql)
                                        EntLib.COPT.Log.Log(ex.Message)
                                        errCount += 1
                                    End Try

                                    If dt1 Is Nothing Then
                                        createObject = True
                                    Else
                                        If dt1.Rows.Count > 0 Then
                                            createObject = False
                                        Else
                                            createObject = True
                                        End If
                                    End If
                                End If
                                If createObject Then
                                    'Step 1: Drop table in destination database, if it already exists
                                    sql = "IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[" & tableName & "]') AND type in (N'U')) DROP TABLE [dbo].[" & tableName & "]"
                                    Try
                                        dalSql.Execute(sql, dt1)
                                        If dalSql.LastErrorNo = 0 Then
                                            EntLib.COPT.Log.Log(tableName & "...DROP TABLE..." & CBool(dt.Rows(ctr).Item("Recreate").ToString()) & "...OK")
                                            msg = msg & "DROPTABLE-OK..."
                                            Console.WriteLine("   DROP   TABLE...OK")
                                        Else
                                            msg = msg & "DROPTABLE-Error..."
                                            Console.WriteLine("   DROP TABLE...Error.")
                                            Console.WriteLine("Error executing SQL.")
                                            Console.WriteLine(sql)
                                            EntLib.COPT.Log.Log("Error executing SQL.")
                                            EntLib.COPT.Log.Log(sql)
                                            errCount += 1
                                        End If
                                    Catch ex As Exception
                                        msg = msg & "DROPTABLE-Error..."
                                        Console.WriteLine("   DROP TABLE...Error.")
                                        Console.WriteLine("Error executing SQL.")
                                        Console.WriteLine(sql)
                                        Console.WriteLine(ex.Message)
                                        EntLib.COPT.Log.Log("Error executing SQL.")
                                        EntLib.COPT.Log.Log(sql)
                                        EntLib.COPT.Log.Log(ex.Message)
                                        errCount += 1
                                    End Try

                                    'Step 2: Create table in destination database + create primary key
                                    sql = "SET ANSI_NULLS ON"
                                    Try
                                        dalSql.Execute(sql, dt1)
                                        If dalSql.LastErrorNo = 0 Then
                                            EntLib.COPT.Log.Log(tableName & "...SET ANSI_NULLS ON..." & "...OK")
                                        Else
                                            msg = msg & "Error in SET ANSI_NULLS ON..."
                                            Console.WriteLine("Error executing SQL.")
                                            Console.WriteLine(sql)
                                            EntLib.COPT.Log.Log("Error executing SQL.")
                                            EntLib.COPT.Log.Log(sql)
                                            errCount += 1
                                        End If
                                    Catch ex As Exception
                                        msg = msg & "Error in SET ANSI_NULLS ON..."
                                        Console.WriteLine("Error executing SQL.")
                                        Console.WriteLine(sql)
                                        Console.WriteLine(ex.Message)
                                        EntLib.COPT.Log.Log("Error executing SQL.")
                                        EntLib.COPT.Log.Log(sql)
                                        EntLib.COPT.Log.Log(ex.Message)
                                        errCount += 1
                                    End Try

                                    sql = "SET QUOTED_IDENTIFIER ON"
                                    Try
                                        dalSql.Execute(sql, dt1)
                                        If dalSql.LastErrorNo = 0 Then
                                            EntLib.COPT.Log.Log(tableName & "...SET QUOTED_IDENTIFIER ON..." & "...OK")
                                        Else
                                            msg = msg & "Error in SET QUOTED_IDENTIFIER ON..."
                                            Console.WriteLine("Error executing SQL.")
                                            Console.WriteLine(sql)
                                            EntLib.COPT.Log.Log("Error executing SQL.")
                                            EntLib.COPT.Log.Log(sql)
                                            errCount += 1
                                        End If
                                    Catch ex As Exception
                                        msg = msg & "Error in SET QUOTED_IDENTIFIER ON..."
                                        Console.WriteLine("Error executing SQL.")
                                        Console.WriteLine(sql)
                                        Console.WriteLine(ex.Message)
                                        EntLib.COPT.Log.Log("Error executing SQL.")
                                        EntLib.COPT.Log.Log(sql)
                                        EntLib.COPT.Log.Log(ex.Message)
                                        errCount += 1
                                    End Try

                                    sql = "CREATE TABLE [dbo].[" & tableName & "] ("
                                    For ctr1 = 0 To tableDefs.Length - 1
                                        sql = sql & "[" & tableDefs(ctr1).fieldName & "] [" & tableDefs(ctr1).fieldTypeNew & "]" & IIf(tableDefs(ctr1).fieldSize > 0, "(" & tableDefs(ctr1).fieldSize & ") ", "").ToString() & IIf(tableDefs(ctr1).fieldIsIdentity, " IDENTITY(1,1)", "").ToString() & IIf(tableDefs(ctr1).fieldIsNullable, " NULL", " NOT NULL").ToString() & ","
                                    Next
                                    If String.IsNullOrEmpty(listPK(0)) Then
                                        sql = Left(sql, Len(sql) - 1) & vbNewLine
                                    Else
                                        sql = sql & " CONSTRAINT [PK_" & tableName & "] PRIMARY KEY CLUSTERED " & vbNewLine
                                        sql = sql & "(" & vbNewLine
                                        For ctr1 = 0 To listPK.Length - 1
                                            sql = sql & "[" & listPK(ctr1) & "] ASC,"
                                        Next
                                        sql = Left(sql, Len(sql) - 1) & vbNewLine
                                        sql = sql & ")WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]" & vbNewLine
                                    End If
                                    sql = sql & ") ON [PRIMARY]" & vbNewLine
                                    Try
                                        dalSql.Execute(sql, dt1)
                                        If dalSql.LastErrorNo = 0 Then
                                            EntLib.COPT.Log.Log(tableName & "...CREATE TABLE..." & "...OK")
                                            msg = msg & "CREATETABLE-OK..."
                                            Console.WriteLine("   CREATE TABLE...OK")
                                        Else
                                            msg = msg & "Error in CREATETABLE..."
                                            Debug.Print("Error creating TABLE - " & tableName)
                                            Debug.Print(sql)
                                            Debug.Print("")
                                            Console.WriteLine("Error executing SQL.")
                                            Console.WriteLine(sql)
                                            errCount += 1
                                        End If
                                    Catch ex As Exception
                                        msg = msg & "Error in CREATETABLE..."
                                        Debug.Print("Error creating TABLE - " & tableName)
                                        Debug.Print(ex.Message)
                                        Debug.Print(sql)
                                        Debug.Print("")
                                        Console.WriteLine("Error executing SQL.")
                                        Console.WriteLine(sql)
                                        Console.WriteLine(ex.Message)
                                        errCount += 1
                                    End Try
                                    'Step 3: Create nonclustered index on the table

                                    'If Not String.IsNullOrEmpty(listIdx(0)) Then
                                    '    sql = "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[" & tableName & "]') AND name = N'" & "IDX_" & tableName & "')"
                                    '    sql = sql & "CREATE UNIQUE NONCLUSTERED INDEX [" & "IDX_" & tableName & "] ON [dbo].[" & tableName & "] "
                                    '    sql = sql & "( "
                                    '    For ctr1 = 0 To listIdx.Length - 1
                                    '        sql = sql & "[" & listIdx(ctr1) & "] ASC,"
                                    '    Next
                                    '    sql = Left(sql, Len(sql) - 1) & vbNewLine
                                    '    sql = sql & ")WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY] "
                                    '    Try
                                    '        dalSql.Execute(sql, dt1)
                                    '    Catch ex As Exception
                                    '        Debug.Print("Error creating NONCLUSTERED INDEX on TABLE - " & tableName)
                                    '        Debug.Print(ex.Message)
                                    '        Debug.Print(sql)
                                    '        Debug.Print("")
                                    '        Console.WriteLine("Error executing SQL.")
                                    '        Console.WriteLine(sql)
                                    '        Console.WriteLine(ex.Message)
                                    '        EntLib.COPT.Log.Log("Error executing SQL.")
                                    '        EntLib.COPT.Log.Log(sql)
                                    '        EntLib.COPT.Log.Log(ex.Message)
                                    '    End Try
                                    'End If

                                    If idxDefs IsNot Nothing Then
                                        If Not String.IsNullOrEmpty(idxDefs(0).idxName) Then
                                            idxComposite = ""
                                            For ctr1 = 0 To idxDefs.Length - 1
                                                idxComposite = ""
                                                For ctr2 = 0 To idxDefs(ctr1).idxCols.Length - 1
                                                    If String.IsNullOrEmpty(idxComposite) Then
                                                        idxComposite = idxDefs(ctr1).idxCols(ctr2)
                                                    Else
                                                        idxComposite = idxComposite & "," & idxDefs(ctr1).idxCols(ctr2)
                                                    End If
                                                Next

                                                sql = "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[" & tableName & "]') AND name = N'" & "IDX_" & idxDefs(ctr1).idxName & "')"
                                                sql = sql & "CREATE " & IIf(idxDefs(ctr1).isUnique, "UNIQUE", "").ToString() & " NONCLUSTERED INDEX [" & "IDX_" & idxDefs(ctr1).idxName & "] ON [dbo].[" & tableName & "] "
                                                sql = sql & "( "
                                                'For ctr1 = 0 To listIdx.Length - 1
                                                '    sql = sql & "[" & listIdx(ctr1) & "] ASC,"
                                                'Next
                                                'sql = Left(sql, Len(sql) - 1) & vbNewLine
                                                sql = sql & idxComposite
                                                sql = sql & ") WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY] "
                                                Try
                                                    dalSql.Execute(sql, dt1)
                                                    If dalSql.LastErrorNo = 0 Then
                                                        EntLib.COPT.Log.Log(tableName & "..." & "CREATE " & IIf(idxDefs(ctr1).isUnique, "UNIQUE", "").ToString() & " NONCLUSTERED INDEX " & idxDefs(ctr1).idxName & "..." & "...OK")
                                                    Else
                                                        Debug.Print("Error creating NONCLUSTERED INDEX on TABLE - " & tableName)
                                                        Debug.Print(sql)
                                                        Debug.Print("")
                                                        Console.WriteLine("Error creating NONCLUSTERED INDEX on TABLE - " & tableName)
                                                        Console.WriteLine(sql)
                                                        Console.WriteLine(dalSql.LastErrorDesc)
                                                        EntLib.COPT.Log.Log("Error creating NONCLUSTERED INDEX on TABLE - " & tableName)
                                                        EntLib.COPT.Log.Log(dalSql.LastErrorDesc)
                                                        EntLib.COPT.Log.Log(sql)
                                                        errCount += 1
                                                    End If

                                                Catch ex As Exception
                                                    Debug.Print("Error creating NONCLUSTERED INDEX on TABLE - " & tableName)
                                                    Debug.Print(ex.Message)
                                                    Debug.Print(sql)
                                                    Debug.Print("")
                                                    Console.WriteLine("Error creating NONCLUSTERED INDEX on TABLE - " & tableName)
                                                    Console.WriteLine(sql)
                                                    Console.WriteLine(ex.Message)
                                                    EntLib.COPT.Log.Log("Error creating NONCLUSTERED INDEX on TABLE - " & tableName)
                                                    EntLib.COPT.Log.Log(sql)
                                                    EntLib.COPT.Log.Log(ex.Message)
                                                    errCount += 1
                                                End Try
                                            Next
                                            msg = msg & "CREATEINDEX-OK..."
                                            Console.WriteLine("   CREATE INDEX...OK")
                                        End If
                                    End If

                                End If
                                'Console.Write("...Created")
                                'Create table + indexes in the SQL Express database - END

                                'Copy data - START
                                If _
                                    CInt(dt.Rows(ctr).Item("TransferData").ToString()) = 1 _
                                    Or _
                                    CInt(dt.Rows(ctr).Item("TransferData").ToString()) = 3 _
                                Then
                                    'Empty destination table
                                    If noTruncate Then
                                        'sql = "DELETE FROM [" & tableName & "]"
                                        sql = "EXEC dbo.asp_TruncateTable '" & tableName & "'"
                                    Else
                                        sql = "TRUNCATE TABLE [" & tableName & "]"
                                    End If

                                    Try
                                        dalSql.Execute(sql, dt1)
                                        If dalSql.LastErrorNo = 0 Then
                                            EntLib.COPT.Log.Log(tableName & "...TRUNCATE TABLE...OK")
                                        Else
                                            Console.WriteLine("Error executing SQL.")
                                            Console.WriteLine(dalSql.LastErrorDesc)
                                            Console.WriteLine(sql)
                                            EntLib.COPT.Log.Log("Error executing SQL.")
                                            EntLib.COPT.Log.Log(dalSql.LastErrorDesc)
                                            EntLib.COPT.Log.Log(sql)
                                            errCount += 1
                                        End If
                                    Catch ex As Exception
                                        Console.WriteLine("Error executing SQL.")
                                        Console.WriteLine(sql)
                                        Console.WriteLine(ex.Message)
                                        EntLib.COPT.Log.Log("Error executing SQL.")
                                        EntLib.COPT.Log.Log(sql)
                                        EntLib.COPT.Log.Log(ex.Message)
                                        errCount += 1
                                    End Try

                                    Console.Write("   BULK    COPY...")
                                    'EntLib.COPT.Log.Log("Create ADO Recordset...")
                                    sql = "SELECT * FROM [" & tableName & "]"
                                    Try
                                        'dalOleDbADO.Execute(sql, dt1)
                                        dalOleDb.Execute(sql, dt1)
                                        If dalOleDb.LastErrorNo = -1 Then
                                            Console.WriteLine("ERROR. APM: " & GenUtils.GetAvailablePhysicalMemoryStr())
                                            Console.WriteLine("Error executing SQL.")
                                            Console.WriteLine("SELECT * FROM [" & tableName & "]")
                                            Console.WriteLine(dalOleDb.LastErrorDesc)
                                            'Console.WriteLine(ex.Message)
                                            EntLib.COPT.Log.Log("Error executing SQL.")
                                            EntLib.COPT.Log.Log("SELECT * FROM [" & tableName & "]")
                                            EntLib.COPT.Log.Log(dalOleDb.LastErrorDesc)
                                            EntLib.COPT.Log.Log("APM: " & GenUtils.GetAvailablePhysicalMemoryStr())
                                            errCount += 1
                                        Else
                                            'EntLib.COPT.Log.Log(tableName & "...SELECT * FROM...OK")
                                            'EntLib.COPT.Log.Log("Create ADO.NET Recordset...SUCCESS.")
                                        End If
                                    Catch ex As Exception
                                        Console.WriteLine("Error executing SQL (using ADO.NET).")
                                        Console.WriteLine(sql)
                                        Console.WriteLine(ex.Message)
                                        EntLib.COPT.Log.Log("Error executing SQL (using ADO.NET).")
                                        EntLib.COPT.Log.Log(sql)
                                        EntLib.COPT.Log.Log(ex.Message)
                                        errCount += 1
                                    End Try
                                    If dalOleDb.LastErrorNo <> -1 Then
                                        sql = "dbo.[" & tableName & "]"
                                        Try
                                            'Console.Write("   BULK    COPY...")
                                            'EntLib.COPT.Log.Log("BulkCopy...")
                                            ret = dalSql.Execute("BULKCOPY-MAP", dt1, sql)
                                            If ret = -1 Then
                                                EntLib.COPT.Log.Log(tableName & "...BULKCOPY-MAP...ERROR")
                                                msg = msg & "...BULKCOPY-ERROR..."
                                                'Console.WriteLine("   BULK    COPY...OK")
                                                Console.WriteLine("ERROR. APM: " & GenUtils.GetAvailablePhysicalMemoryStr())
                                                Console.WriteLine("No. of rows: " & dt1.Rows.Count.ToString())
                                                Console.WriteLine("ERROR Desc:")
                                                Console.WriteLine(dalSql.LastErrorDesc)
                                                Console.WriteLine("Please check C-OPT.log for details.")
                                                EntLib.COPT.Log.Log("No. of rows: " & dt1.Rows.Count.ToString())
                                                errCount += 1
                                            Else
                                                EntLib.COPT.Log.Log(tableName & "...BULKCOPY-MAP...OK (" & dt1.Rows.Count.ToString() & " rows)")
                                                msg = msg & "...BULKCOPY-OK..."
                                                'EntLib.COPT.Log.Log("BulkCopy...SUCCESS.")
                                                'Console.WriteLine("   BULK    COPY...OK")
                                                Console.WriteLine("OK (" & dt1.Rows.Count.ToString() & " rows)")
                                            End If
                                        Catch ex As Exception
                                            Console.WriteLine("Error executing SQL.")
                                            Console.WriteLine(sql)
                                            Console.WriteLine(ex.Message)
                                            Console.WriteLine("No. of rows: " & dt1.Rows.Count.ToString())
                                            EntLib.COPT.Log.Log("Error executing SQL.")
                                            EntLib.COPT.Log.Log(sql)
                                            EntLib.COPT.Log.Log(ex.Message)
                                            EntLib.COPT.Log.Log("No. of rows: " & dt1.Rows.Count.ToString())
                                            errCount += 1
                                        End Try
                                    End If
                                    'Console.Write("...Populated with data.")
                                End If
                                'Copy data - END

                                'Console.WriteLine(msg)

                                Exit For
                            End If 'CInt(dt.Rows(ctr).Item("TransferData").ToString()) = 1 Or CInt(dt.Rows(ctr).Item("TransferData").ToString()) = 3
                        Next 'For Each adoTable In adoTables
                        'If GenUtils.IsSwitchAvailable(switches, "/UseGC") Then
                        GenUtils.CollectGarbage()
                        'End If
                        'dt1 = Nothing
                    Next 'For ctr = 0 To dt.Rows.Count - 1

                    Console.WriteLine("")
                    EntLib.COPT.Log.Log("")
                    If errCount = 0 Then
                        Console.WriteLine("All TRANSFER tasks finished successfully.")
                        EntLib.COPT.Log.Log("All TRANSFER tasks finished successfully.")
                    Else
                        Console.WriteLine("TRANSFER finished with: " & errCount & " error(s).")
                        EntLib.COPT.Log.Log("TRANSFER finished with: " & errCount & " error(s).")
                    End If

                End If 'If proceed Then 'proceed #2
            End If 'If proceed Then 'proceed #1

            dalOleDbADO = Nothing
            dalOleDb = Nothing
            dalSql = Nothing
            dt = Nothing
            dt1 = Nothing

            GenUtils.CollectGarbage()

            Try
                GenUtils.GetSysDatabase().ExecuteNonQuery( _
                    "INSERT INTO tblRunDetails " & _
                    "SELECT " & _
                    "'" & My.Computer.Clock.LocalTime.ToString() & "' AS LCT, " & _
                    "'" & My.Computer.Clock.GmtTime.ToString() & "' AS GMT, " & _
                    "'" & My.Computer.Name.Replace("=", ">") & "' AS CNM, " & _
                    "'" & GenUtils.GetUserName().ToString() & "' AS UNM, " & _
                    "'" & My.Application.Info.AssemblyName & "' AS EXE, " & _
                    "'" & GenUtils.Version & "' AS VER, " & _
                    "'" & GenUtils.VersionName & "' AS VEN, " & _
                    "'" & "TRANSFER" & "' AS PRJ, " & _
                    "'" & "" & "' AS RUN, " & _
                    "'" & "' AS SLV, " & _
                    "'" & "' AS SVE, " & _
                    "'" & "' AS PTY, " & _
                    "'" & "" & "' AS SST, " & _
                    "'" & "" & "' AS ROW, " & _
                    "'" & "" & "' AS COL, " & _
                    "'" & "" & "' AS NZE, " & _
                    "'" & "" & "' AS OBJ, " & _
                    "'" & "" & "' AS ITE, " & _
                    "'" & GenUtils.FormatTime(startTime) & "' AS STM, " & _
                    "'" & errCount & "' AS BAD, " & _
                    "'" & "" & "' AS INF, " & _
                    "'" & GenUtils.GetAvailablePhysicalMemoryStr() & "" & "' AS APM, " & _
                    "'" & "" & "' AS CST, " & _
                    "'" & "TRANSFER /SrcDBName """ & srcDBName & """ /SrcDBFolder """ & srcDBFolder & """ /DestDBName """ & destDBName & """ /DestDBFolder """ & destDBFolder & """" & "" & "' AS SWT " & _
                    "" _
                )
            Catch ex1 As Exception
                EntLib.COPT.Log.Log(GenUtils.GetWorkDir(switches), "", "Error logging run details to C-OPTSYS database." & vbNewLine & "Error:" & vbNewLine & ex1.Message)
            End Try

            'GenUtils.FormatTime(startTime, My.Computer.Clock.TickCount)
            'Console.WriteLine("TRANSFER took: " & DateDiff(DateInterval.Minute, startTime, Now()).ToString() & " min. Start time: " & startTime.ToString())
            Console.WriteLine("TRANSFER took: " & GenUtils.FormatTime(startTime) & ". Start time: " & startTimeD)
            EntLib.COPT.Log.Log("TRANSFER took: " & GenUtils.FormatTime(startTime) & ". Start time: " & startTimeD)
            EntLib.COPT.Log.Log("********************")

            Return 0
        End Function

        ''' <summary>
        ''' TRANSFERBACK or TRB - Transfers data from source (SQL Server Express) (.MDF) database to destination (MS Access) (.MDB) database.
        ''' </summary>
        ''' <param name="srcDBType"></param>
        ''' <param name="destDBType"></param>
        ''' <param name="switches"></param>
        ''' <param name="srcDBName"></param>
        ''' <param name="srcDBFolder"></param>
        ''' <param name="destDBName"></param>
        ''' <param name="destDBFolder"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' RKP/03-28-12/v3.0.161
        ''' Made changes to Console.WriteLine() output.
        ''' </remarks>
        Public Shared Function DB_TransferBack( _
            ByVal srcDBType As e_DB, _
            ByVal destDBType As e_DB, _
            ByRef switches() As String, _
            ByVal srcDBName As String, _
            ByVal srcDBFolder As String, _
            ByVal destDBName As String, _
            ByVal destDBFolder As String _
        ) As Integer

            Dim srcDBExt As String = ""
            Dim destDBExt As String = ""
            Dim filePath As String = ""
            Dim srcFilePath As String = ""
            Dim destFilePath As String = ""
            Dim proceed As Boolean = False
            Dim msg As String = ""
            Dim input As String = ""
            Dim dalOleDb As New clsDAL
            Dim dalSql As New clsDAL
            Dim connOleDb As New OleDb.OleDbConnection
            Dim connSql As New SqlClient.SqlConnection
            'Dim adoCatalog As New ADOX.Catalog
            'Dim adoTables As ADOX.Tables
            'Dim adoTable As ADOX.Table
            'Dim adoCols As ADOX.Columns
            'Dim adoCol As ADOX.Column
            Dim sql As String = ""
            Dim sqlSelect As String = ""
            Dim dt As DataTable
            Dim dt1 As New DataTable
            Dim ctr As Integer
            Dim ret As Integer
            Dim tableCtrDelete As Integer 'RKP/03-29-12/v3.0.163
            Dim tableCtrInsert As Integer 'RKP/04-02-12/v3.0.163
            'Dim tableCtr As Integer 'RKP/04-02-12/v3.0.163
            Dim arrKeyValue() As STRUCT_KEYVALUE
            Dim rowCount As Integer 'RKP/04-11-12/v3.1.165
            Dim startTime As Long = My.Computer.Clock.TickCount 'RKP/04-30-12/v3.2.167
            Dim startTimeD As Date = Now()

            EntLib.COPT.Log.Log("********************")
            EntLib.COPT.Log.Log("TRANSFERBACK (TRB)")
            EntLib.COPT.Log.Log("Start: " & Now())
            'startTime = Now()

            If srcDBType = e_DB.e_db_SQLSERVER_EXPRESS Then
                srcDBExt = ".mdf"
                If destDBType = e_DB.e_db_ACCESS_MDB Then
                    destDBExt = ".mdb"
                Else

                End If
            ElseIf srcDBType = e_DB.e_db_ACCESS_MDB Then
                srcDBExt = ".mdb"
                If destDBType = e_DB.e_db_SQLSERVER_EXPRESS Then
                    destDBExt = ".mdf"
                End If
            Else


            End If

            If (Not String.IsNullOrEmpty(srcDBExt)) AndAlso (Not String.IsNullOrEmpty(destDBExt)) Then
                If srcDBType = e_DB.e_db_SQLSERVER_EXPRESS And destDBType = e_DB.e_db_ACCESS_MDB Then
                    If String.IsNullOrEmpty(srcDBFolder) Then
                        srcDBFolder = "C:\OPTMODELS\" & srcDBName
                    End If

                    If String.IsNullOrEmpty(destDBFolder) Then
                        destDBFolder = "C:\OPTMODELS\" & destDBName
                    End If

                    filePath = srcDBFolder & "\" & srcDBName & IIf(srcDBName.Trim().ToUpper().EndsWith(".MDF"), "", ".mdf").ToString() '& ".mdf"  '"C:\OPTMODELS\IPG50SQL\IPG50SQL.mdf"
                    srcFilePath = filePath
                    If My.Computer.FileSystem.FileExists(filePath) Then
                        proceed = True
                    Else
                        proceed = False
                        msg = msg & vbNewLine & "Source database not found."
                    End If
                    If proceed Then 'proceed #1
                        filePath = destDBFolder & "\" & destDBName & IIf(destDBName.Trim().ToUpper().EndsWith(".MDB"), "", ".mdb").ToString()
                        destFilePath = filePath
                        If My.Computer.FileSystem.FileExists(filePath) Then
                            proceed = True
                        Else
                            proceed = False
                            msg = msg & vbNewLine & "Destination database not found."
                        End If

                        If proceed Then 'proceed #2
                            input = "N"
                            If EntLib.COPT.GenUtils.IsSwitchAvailable(switches, "/NoPrompt") Then
                                proceed = True
                                input = "Y"
                            Else 'If EntLib.COPT.GenUtils.IsSwitchAvailable(switches, "/NoPrompt") Then
                                Console.WriteLine()
                                Console.WriteLine("Source File Path: " & srcFilePath)
                                Console.WriteLine("Destination File Path: " & destFilePath)
                                Console.Write("Do you want to continue? (Y/N): ")
                                input = Console.ReadLine()
                            End If

                            If String.IsNullOrEmpty(input) Then
                                input = "Y"
                            End If
                            If input.Trim.ToUpper.Equals("Y") Then
                                proceed = True
                            Else
                                proceed = False
                            End If

                            If proceed Then 'proceed #3
                                filePath = srcDBFolder & "\" & srcDBName & IIf(srcDBName.Trim().ToUpper().EndsWith(".MDF"), "", ".mdf").ToString()  '"C:\OPTMODELS\IPG50SQL\IPG50SQL.mdf"
                                Try
                                    dalSql.Connect(clsDAL.e_DB.e_db_SQLSERVER, clsDAL.e_ConnType.e_connType_I, Nothing, Nothing, connSql, Nothing, Nothing, False, ".\SQLEXPRESS", srcDBFolder & "\" & srcDBName & ".mdf", srcDBName)
                                    EntLib.COPT.Log.Log("Connected to source database successfully - attempt #1 (" & srcDBName & ") (" & srcFilePath & ").")
                                Catch ex As Exception
                                    EntLib.COPT.Log.Log("Error connecting to source database - attempt #1 (" & srcDBName & ") (" & srcFilePath & ").")
                                    EntLib.COPT.Log.Log(ex.Message)

                                    'RKP/03-28-12/v3.0.161
                                    'If there is an error connecting for the first time, try one more time.
                                    Try
                                        dalSql.Connect(clsDAL.e_DB.e_db_SQLSERVER, clsDAL.e_ConnType.e_connType_I, Nothing, Nothing, connSql, Nothing, Nothing, True, ".\SQLEXPRESS", srcDBFolder & "\" & srcDBName & ".mdf", srcDBName)
                                        EntLib.COPT.Log.Log("Connected to source database successfully - attempt #2 (" & srcDBName & ") (" & srcFilePath & ").")
                                    Catch ex2 As Exception
                                        EntLib.COPT.Log.Log("Error connecting to source database - attempt #2 (" & srcDBName & ") (" & srcFilePath & ").")
                                        EntLib.COPT.Log.Log(ex2.Message)
                                        proceed = False
                                    End Try
                                End Try

                                If proceed Then 'proceed #4
                                    filePath = destDBFolder & "\" & destDBName & IIf(destDBName.Trim().ToUpper().EndsWith(".MDB"), "", ".mdb").ToString() '& ".mdb"  '"C:\OPTMODELS\IPG50\IPG50.mdb"
                                    dalOleDb.UseADO = True
                                    Try
                                        dalOleDb.Connect(clsDAL.e_DB.e_db_ACCESS, clsDAL.e_ConnType.e_connType_A, Nothing, connOleDb, Nothing, Nothing, Nothing, True, filePath)
                                        EntLib.COPT.Log.Log("Connected to destination database successfully - attempt #1 (" & destDBName & ") (" & destFilePath & ").")
                                    Catch ex As Exception
                                        EntLib.COPT.Log.Log("Connected to destination database - attempt #1 (" & destDBName & ") (" & destFilePath & ").")
                                        EntLib.COPT.Log.Log(ex.Message)
                                        proceed = False
                                    End Try

                                    If proceed Then 'proceed #5
                                        'adoCatalog = dalOleDb.ConnADOCatalog
                                        'adoTables = adoCatalog.Tables
                                        'adoCatalog.ActiveConnection = Nothing

                                        sql = "SELECT * FROM tsysCoreObjects WHERE (OBJ_TYPE_ID = 1) AND ([Active] <> 0) AND (TransferData = 2 OR TransferData = 3)"

                                        dt = New DataTable
                                        Try
                                            'dalOleDb.Execute(sql, dt)

                                            'RKP/03-24-12/v3.0.159
                                            'Always look for tsysCoreObjects in the source database.
                                            dalSql.Execute(sql, dt)

                                            If dalSql.LastErrorNo >= 0 Then
                                                'do nothing
                                            Else
                                                EntLib.COPT.Log.Log("Error reading from tsysCoreObjects table.")
                                                EntLib.COPT.Log.Log(dalSql.LastErrorDesc)
                                                EntLib.COPT.Log.Log(sql)
                                                proceed = False
                                            End If
                                        Catch ex As Exception
                                            EntLib.COPT.Log.Log("Error reading from tsysCoreObjects table.")
                                            EntLib.COPT.Log.Log(ex.Message)
                                            EntLib.COPT.Log.Log(sql)
                                            proceed = False
                                        End Try

                                        If proceed Then 'proceed #6
                                            Console.WriteLine("")

                                            tableCtrDelete = 0
                                            tableCtrInsert = 0

                                            'Console.WriteLine("Deleting all rows in destination tables.")
                                            'EntLib.COPT.Log.Log("")
                                            'EntLib.COPT.Log.Log("Deleting all rows in destination tables.")
                                            'tableCtr = 0
                                            'For ctr = 0 To dt.Rows.Count - 1
                                            '    Try
                                            '        ret = dalOleDb.Execute("DELETE * FROM [" & dt.Rows(ctr).Item("OBJ_NAME").ToString() & "]", dt1)
                                            '        Console.WriteLine("SUCCESS - DestDB: " & destDBName & " - DELETE * FROM [" & dt.Rows(ctr).Item("OBJ_NAME").ToString() & "]")
                                            '        EntLib.COPT.Log.Log("SUCCESS - DestDB: " & destDBName & " - DELETE * FROM [" & dt.Rows(ctr).Item("OBJ_NAME").ToString() & "]")
                                            '        tableCtr += 1
                                            '    Catch ex As Exception
                                            '        Console.WriteLine("FAILURE - DestDB: " & destDBName & " - DELETE * FROM [" & dt.Rows(ctr).Item("OBJ_NAME").ToString() & "]")
                                            '        EntLib.COPT.Log.Log("FAILURE - DestDB: " & destDBName & " - DELETE * FROM [" & dt.Rows(ctr).Item("OBJ_NAME").ToString() & "]")
                                            '        EntLib.COPT.Log.Log(ex.Message)
                                            '        EntLib.COPT.Log.Log(sql)
                                            '    End Try
                                            'Next
                                            'Console.WriteLine("Cleared all rows in " & tableCtr & " destination tables.")
                                            'EntLib.COPT.Log.Log("Cleared all rows in " & tableCtr & " destination tables.")

                                            'RKP/04-04-12/v3.0.164
                                            'Fill up an array with table name and row count.
                                            'This will be used to make sure the destination table is cleared&appended, if and only if the source table has row count > 0.
                                            ReDim arrKeyValue(dt.Rows.Count - 1)
                                            For ctr = 0 To dt.Rows.Count - 1
                                                sql = "SELECT COUNT(*) AS ROW_COUNT "
                                                sql = sql & "FROM "
                                                sql = sql & srcDBName & ".dbo.[" & dt.Rows(ctr).Item("OBJ_NAME").ToString() & "] a "
                                                Try
                                                    ret = dalSql.Execute(sql, dt1)
                                                    arrKeyValue(ctr).key = dt.Rows(ctr).Item("OBJ_NAME").ToString()
                                                    arrKeyValue(ctr).valueInt = CInt(dt1.Rows(0).Item("ROW_COUNT"))
                                                Catch ex As Exception
                                                    Console.WriteLine("FAILURE-TRB-SrcDB:" & srcDBName & "DestDB:" & destDBName & "-[" & dt.Rows(ctr).Item("OBJ_NAME").ToString() & "]")
                                                    Console.WriteLine("Error getting row count from source table: " & srcDBName & ".dbo.[" & dt.Rows(ctr).Item("OBJ_NAME").ToString() & "]")
                                                    Console.WriteLine(ex.Message)
                                                    Console.WriteLine(sql)
                                                    EntLib.COPT.Log.Log("FAILURE-TRB-SrcDB:" & srcDBName & "DestDB:" & destDBName & "-[" & dt.Rows(ctr).Item("OBJ_NAME").ToString() & "]")
                                                    EntLib.COPT.Log.Log("Error getting row count from source table: " & srcDBName & ".dbo.[" & dt.Rows(ctr).Item("OBJ_NAME").ToString() & "]")
                                                    EntLib.COPT.Log.Log(ex.Message)
                                                    EntLib.COPT.Log.Log(sql)
                                                End Try
                                            Next

                                            For ctr = 0 To arrKeyValue.Length - 1
                                                If arrKeyValue(ctr).valueInt > 0 Then
                                                    Try
                                                        sql = "DELETE * FROM [" & dt.Rows(ctr).Item("OBJ_NAME").ToString() & "]"
                                                        ret = dalOleDb.Execute(sql, dt1)
                                                        tableCtrDelete += 1
                                                        'tableCtr = 1
                                                        arrKeyValue(ctr).flag = 1
                                                    Catch ex As Exception
                                                        Console.WriteLine("FAILURE-TRB-DestDB:" & destDBName & "-DELETE * FROM [" & dt.Rows(ctr).Item("OBJ_NAME").ToString() & "]")
                                                        EntLib.COPT.Log.Log("FAILURE-TRB-DestDB:" & destDBName & "-DELETE * FROM [" & dt.Rows(ctr).Item("OBJ_NAME").ToString() & "]")
                                                        EntLib.COPT.Log.Log(ex.Message)
                                                        EntLib.COPT.Log.Log(sql)
                                                    End Try
                                                Else
                                                    Console.WriteLine("TRB-'DELETE * FROM '[" & srcDBName & "].[dbo].[" & dt.Rows(ctr).Item("OBJ_NAME").ToString() & "] skipped due to 0 rows.")
                                                    EntLib.COPT.Log.Log("TRB-'DELETE * FROM '[" & srcDBName & "].[dbo].[" & dt.Rows(ctr).Item("OBJ_NAME").ToString() & "] skipped due to 0 rows.")
                                                    EntLib.COPT.Log.Log(sql)
                                                End If
                                            Next

                                            For ctr = 0 To dt.Rows.Count - 1
                                                'adoTable = adoTables(dt.Rows(ctr).Item("OBJ_NAME").ToString())
                                                'adoCols = adoTable.Columns
                                                'tableCtr = 0

                                                'sql = "SELECT COUNT(*) AS ROW_COUNT "
                                                'sql = sql & "FROM "
                                                'sql = sql & srcDBName & ".dbo.[" & dt.Rows(ctr).Item("OBJ_NAME").ToString() & "] a "
                                                'Try 'Get row count
                                                'ret = dalSql.Execute(sql, dt1)
                                                'If CInt(dt1.Rows(0).Item("ROW_COUNT")) > 0 Then
                                                If arrKeyValue(ctr).valueInt > 0 Then
                                                    'Try 'Clear all rows
                                                    '    'ret = dalSql.Execute("DELETE * FROM [" & destDBName & "MDB]...[" & adoTable.Name & "]", dt1)

                                                    '    ret = dalOleDb.Execute("DELETE * FROM [" & dt.Rows(ctr).Item("OBJ_NAME").ToString() & "]", dt1)

                                                    '    'Console.WriteLine(dt.Rows(ctr).Item("OBJ_NAME").ToString() & " - SUCCESS - DELETE * FROM [DestTable]")
                                                    '    'EntLib.COPT.Log.Log(dt.Rows(ctr).Item("OBJ_NAME").ToString() & " - SUCCESS - DELETE * FROM [DestTable]")

                                                    '    'Console.WriteLine("SUCCESS - DestDB: " & destDBName & " - DELETE * FROM [" & dt.Rows(ctr).Item("OBJ_NAME").ToString() & "]")
                                                    '    'EntLib.COPT.Log.Log("SUCCESS - DestDB: " & destDBName & " - DELETE * FROM [" & dt.Rows(ctr).Item("OBJ_NAME").ToString() & "]")
                                                    '    tableCtrDelete += 1
                                                    '    'tableCtr = 1
                                                    'Catch ex1 As Exception 'Clear all rows
                                                    '    'Console.WriteLine(dt.Rows(ctr).Item("OBJ_NAME").ToString() & " - FAILURE - DELETE * FROM [DestTable]")
                                                    '    'EntLib.COPT.Log.Log(dt.Rows(ctr).Item("OBJ_NAME").ToString() & " - FAILURE - DELETE * FROM [DestTable]")
                                                    '    Console.WriteLine("FAILURE - TRB - DestDB: " & destDBName & " - DELETE * FROM [" & dt.Rows(ctr).Item("OBJ_NAME").ToString() & "]")
                                                    '    EntLib.COPT.Log.Log("FAILURE - TRB - DestDB: " & destDBName & " - DELETE * FROM [" & dt.Rows(ctr).Item("OBJ_NAME").ToString() & "]")
                                                    '    EntLib.COPT.Log.Log(ex1.Message)
                                                    '    EntLib.COPT.Log.Log(sql)
                                                    'End Try 'Clear all rows

                                                    sql = "SELECT COUNT(*) AS ROW_COUNT FROM " & srcDBName & ".dbo.[" & dt.Rows(ctr).Item("OBJ_NAME").ToString() & "] a "
                                                    Try
                                                        ret = dalSql.Execute(sql, dt1)
                                                        rowCount = CInt(dt1.Rows(0).Item("ROW_COUNT").ToString())
                                                    Catch ex As Exception
                                                        rowCount = 0
                                                    End Try

                                                    'sql = "INSERT INTO [" & destDBName & "MDB]...[" & adoTable.Name & "] "
                                                    sql = "INSERT INTO [" & destDBName & "MDB]...[" & dt.Rows(ctr).Item("OBJ_NAME").ToString() & "] "
                                                    sql = sql & "SELECT a.* "
                                                    sql = sql & "FROM "
                                                    'dt.Rows(ctr).Item("OBJ_NAME").ToString()
                                                    'sql = sql & srcDBName & ".dbo.[" & adoTable.Name & "] a "
                                                    sql = sql & srcDBName & ".dbo.[" & dt.Rows(ctr).Item("OBJ_NAME").ToString() & "] a "
                                                    Try 'Insert rows
                                                        ret = dalSql.Execute(sql, dt1)
                                                        If ret = -1 Then
                                                            'Console.WriteLine(adoTable.Name & " - FAILURE - Src to Dest.")
                                                            'dt.Rows(ctr).Item("OBJ_NAME").ToString()
                                                            'Console.WriteLine("FAILURE - TRB - SrcDB: " & srcDBName & " DestDB: " & destDBName & " - [" & adoTable.Name().ToString() & "]")
                                                            Console.WriteLine("FAILURE-TRB-SrcDB:" & srcDBName & ",DestDB:" & destDBName & "-[" & dt.Rows(ctr).Item("OBJ_NAME").ToString() & "] (" & rowCount & " rows/" & Format(((My.Computer.FileSystem.GetFileInfo(destFilePath).Length / 1024) / 1024), "0.00") & " MB)")
                                                            Console.WriteLine("  Possible Reasons: (a) Database size exceeded size limit (b) Compact & Repair required. (c) Need to ""Run as Administrator."" (d) Need to set up Linked Server.")
                                                            'EntLib.COPT.Log.Log(adoTable.Name & " - FAILURE - Src to Dest.")
                                                            EntLib.COPT.Log.Log("FAILURE-TRB-SrcDB:" & srcDBName & ",DestDB:" & destDBName & " - [" & dt.Rows(ctr).Item("OBJ_NAME").ToString() & "] (" & rowCount & " rows/" & Format(((My.Computer.FileSystem.GetFileInfo(destFilePath).Length / 1024) / 1024), "0.00") & " MB)")
                                                            EntLib.COPT.Log.Log("  Possible Reasons: (a) Database size exceeded size limit (b) Compact & Repair required. (c) Need to ""Run as Administrator."" (d) Need to set up Linked Server.")
                                                        Else 'If ret = -1 Then
                                                            'Console.WriteLine(adoTable.Name & " - SUCCESS - Src to Dest.")
                                                            'EntLib.COPT.Log.Log(adoTable.Name & " - SUCCESS - Src to Dest.")

                                                            Console.WriteLine("OK-TRB-SrcDB:" & srcDBName & ",DestDB:" & destDBName & "-" & "[" & dt.Rows(ctr).Item("OBJ_NAME").ToString() & "] (" & rowCount & " rows/" & Format(((My.Computer.FileSystem.GetFileInfo(destFilePath).Length / 1024) / 1024), "0.00") & " MB)")
                                                            EntLib.COPT.Log.Log("OK-TRB-SrcDB:" & srcDBName & ",DestDB:" & destDBName & "-" & "[" & dt.Rows(ctr).Item("OBJ_NAME").ToString() & "] (" & rowCount & " rows/" & Format(((My.Computer.FileSystem.GetFileInfo(destFilePath).Length / 1024) / 1024), "0.00") & " MB)")

                                                            tableCtrInsert += 1
                                                            arrKeyValue(ctr).flag = arrKeyValue(ctr).flag + 2
                                                            'tableCtr = tableCtr + 2

                                                            'Console.WriteLine("SUCCESS - TRB - SrcDB: " & srcDBName & " DestDB: " & destDBName & " - [" & dt.Rows(ctr).Item("OBJ_NAME").ToString() & "]")
                                                            'EntLib.COPT.Log.Log("SUCCESS - TRB - SrcDB: " & srcDBName & " DestDB: " & destDBName & " - [" & dt.Rows(ctr).Item("OBJ_NAME").ToString() & "]")
                                                        End If 'If ret = -1 Then
                                                    Catch ex As Exception 'Insert rows
                                                        'Try
                                                        '    ret = dalSql.Execute(sql, dt1)
                                                        '    Console.WriteLine(adoTable.Name & " - SUCCESS - Src to Dest.")
                                                        '    EntLib.COPT.Log.Log(adoTable.Name & " - SUCCESS - Src to Dest.")
                                                        'Catch ex1 As Exception
                                                        '    Console.WriteLine(adoTable.Name & " - FAILURE - Src to Dest.")
                                                        '    EntLib.COPT.Log.Log(adoTable.Name & " - FAILURE - Src to Dest.")
                                                        '    EntLib.COPT.Log.Log(ex.Message)
                                                        '    EntLib.COPT.Log.Log(ex1.Message)
                                                        '    EntLib.COPT.Log.Log(sql)
                                                        'End Try
                                                        'EntLib.COPT.Log.Log(adoTable.Name & " - FAILURE - Src to Dest.")
                                                        EntLib.COPT.Log.Log("FAILURE-TRB-SrcDB:" & srcDBName & ",DestDB:" & destDBName & "-[" & dt.Rows(ctr).Item("OBJ_NAME").ToString() & "] (" & rowCount & " rows/" & Format(((My.Computer.FileSystem.GetFileInfo(destFilePath).Length / 1024) / 1024), "0.00") & " MB)")
                                                        EntLib.COPT.Log.Log(ex.Message)
                                                        EntLib.COPT.Log.Log(sql)
                                                    End Try 'Insert rows
                                                Else 'If CInt(dt1.Rows(0).Item("ROW_COUNT")) > 0 Then
                                                    Console.WriteLine("FAILURE-TRB-SrcDB:" & srcDBName & ",DestDB: " & destDBName & "-[" & dt.Rows(ctr).Item("OBJ_NAME").ToString() & "] (" & rowCount & " rows/" & Format(((My.Computer.FileSystem.GetFileInfo(destFilePath).Length / 1024) / 1024), "0.00") & " MB)")
                                                    Console.WriteLine("  Possible Reasons: (a) Database size exceeded size limit (b) Compact & Repair required. (c) Need to ""Run as Administrator."" (d) Need to set up Linked Server.")
                                                    'EntLib.COPT.Log.Log("FAILURE - TRB - SrcDB: " & srcDBName & " DestDB: " & destDBName & " - [" & adoTable.Name().ToString() & "]")
                                                    EntLib.COPT.Log.Log("TRB skipped for table: " & srcDBName & ".dbo.[" & dt.Rows(ctr).Item("OBJ_NAME").ToString() & "] due to 0 rows. (" & rowCount & " src rows/" & Format(((My.Computer.FileSystem.GetFileInfo(destFilePath).Length / 1024) / 1024), "0.00") & " MB)")
                                                    'EntLib.COPT.Log.Log(ex.Message)
                                                    EntLib.COPT.Log.Log(sql)
                                                End If 'If CInt(dt1.Rows(0).Item("ROW_COUNT")) > 0 Then
                                                'Catch ex As Exception 'Get row count
                                                '    Console.WriteLine("FAILURE - TRB - SrcDB: " & srcDBName & " DestDB: " & destDBName & " - [" & dt.Rows(ctr).Item("OBJ_NAME").ToString() & "]")
                                                '    Console.WriteLine("Error getting row count from source table: " & srcDBName & ".dbo.[" & dt.Rows(ctr).Item("OBJ_NAME").ToString() & "]")
                                                '    Console.WriteLine(ex.Message)
                                                '    Console.WriteLine(sql)
                                                '    EntLib.COPT.Log.Log("FAILURE - TRB - SrcDB: " & srcDBName & " DestDB: " & destDBName & " - [" & dt.Rows(ctr).Item("OBJ_NAME").ToString() & "]")
                                                '    EntLib.COPT.Log.Log("Error getting row count from source table: " & srcDBName & ".dbo.[" & dt.Rows(ctr).Item("OBJ_NAME").ToString() & "]")
                                                '    EntLib.COPT.Log.Log(ex.Message)
                                                '    EntLib.COPT.Log.Log(sql)
                                                'End Try 'Get row count

                                                'Console.WriteLine("SUCCESS - TRB - SrcDB: " & srcDBName & ", DestDB: " & destDBName & " - Clear:" & IIf(tableCtr = 1 OrElse tableCtr = 3, " OK", "ERR").ToString() & ", Append:" & IIf(tableCtr = 2 OrElse tableCtr = 3, " OK", "ERR").ToString() & " - " & "[" & dt.Rows(ctr).Item("OBJ_NAME").ToString() & "]")
                                                'EntLib.COPT.Log.Log("SUCCESS - TRB - SrcDB: " & srcDBName & ", DestDB: " & destDBName & " - Clear:" & IIf(tableCtr = 1 OrElse tableCtr = 3, " OK", "ERR").ToString() & ", Append:" & IIf(tableCtr = 2 OrElse tableCtr = 3, " OK", "ERR").ToString() & " - " & "[" & dt.Rows(ctr).Item("OBJ_NAME").ToString() & "]")

                                                'Console.WriteLine("SUCCESS - TRB - SrcDB: " & srcDBName & ", DestDB: " & destDBName & " - " & "[" & dt.Rows(ctr).Item("OBJ_NAME").ToString() & "]")
                                                'EntLib.COPT.Log.Log("SUCCESS - TRB - SrcDB: " & srcDBName & ", DestDB: " & destDBName & " - " & "[" & dt.Rows(ctr).Item("OBJ_NAME").ToString() & "]")
                                            Next 'For ctr = 0 To dt.Rows.Count - 1
                                            Console.WriteLine("TRB - Cleared rows in: " & tableCtrDelete & " source tables/Appended rows in: " & tableCtrInsert & " destination tables. (" & Format(((My.Computer.FileSystem.GetFileInfo(destFilePath).Length / 1024) / 1024), "0.00") & " MB)")
                                            EntLib.COPT.Log.Log("TRB - Cleared rows in: " & tableCtrDelete & " source tables/Appended rows in: " & tableCtrInsert & " destination tables. (" & Format(((My.Computer.FileSystem.GetFileInfo(destFilePath).Length / 1024) / 1024), "0.00") & " MB)")
                                            If tableCtrDelete <> tableCtrInsert Then
                                                Console.WriteLine("Please refer to C-OPT.log for more details about error(s).")
                                            End If
                                        End If 'proceed #6
                                    End If 'proceed #5
                                End If 'proceed #4
                            End If 'proceed #3
                            'End If 'If EntLib.COPT.GenUtils.IsSwitchAvailable(switches, "/NoPrompt") Then
                        End If 'proceed #2
                    End If 'proceed #1
                Else 'If srcDBType = e_DB.e_db_SQLSERVER_EXPRESS And destDBType = e_DB.e_db_ACCESS_MDB Then

                End If 'If srcDBType = e_DB.e_db_SQLSERVER_EXPRESS And destDBType = e_DB.e_db_ACCESS_MDB Then
            Else 'If (Not String.IsNullOrEmpty(srcDBExt)) AndAlso (Not String.IsNullOrEmpty(destDBExt)) Then
                'quit
            End If 'If (Not String.IsNullOrEmpty(srcDBExt)) AndAlso (Not String.IsNullOrEmpty(destDBExt)) Then

            Console.WriteLine("TRANSFERBACK took: " & GenUtils.FormatTime(startTime) & ". Start time: " & startTimeD)
            EntLib.COPT.Log.Log("TRANSFERBACK took: " & GenUtils.FormatTime(startTime) & ". Start time: " & startTimeD)
            EntLib.COPT.Log.Log("********************")

            Return 0
        End Function

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="switches"></param>
        ''' <param name="dbName"></param>
        ''' <param name="dbFolderPath"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' DBCOMMAND /SQLExpress /Detach /DBName "IPG60" /DBFolder "C:\OPTMODELS\IPG60"
        ''' DBC /SQLExpress /Detach /DBName "IPG60" /DBFolder "C:\OPTMODELS\IPG60"
        ''' </remarks>
        Public Shared Function DB_SQLExpress_Detach( _
            ByRef switches() As String, _
            ByRef dbName As String, _
            ByVal dbFolderPath As String _
        ) As Integer

            Dim dalSql As New clsDAL
            Dim proceed As Boolean
            Dim connSql As New SqlClient.SqlConnection
            Dim sql As String
            Dim ret As Integer
            Dim dt As DataTable
            'Dim input As String

            EntLib.COPT.Log.Log("")
            EntLib.COPT.Log.Log("********************")
            EntLib.COPT.Log.Log("SQLEXPRESS - DETACH - Started at: " & Now())

            proceed = True

            If String.IsNullOrEmpty(dbName) Then
                proceed = False
            End If
            If proceed Then 'proceed #1
                If String.IsNullOrEmpty(dbFolderPath) Then
                    dbFolderPath = "C:\OPTMODELS\" & dbName
                End If
                Try
                    'C:\Program Files\Microsoft SQL Server\MSSQL10_50.SQLEXPRESS\MSSQL\DATA\master.mdf
                    'dalSql.Connect(clsDAL.e_DB.e_db_SQLSERVER, clsDAL.e_ConnType.e_connType_I, Nothing, Nothing, connSql, Nothing, Nothing, True, ".\SQLEXPRESS", "|DataDirectory|" & "master.mdf", "MASTER")
                    dalSql.Connect( _
                        clsDAL.e_DB.e_db_SQLSERVER, _
                        clsDAL.e_ConnType.e_connType_I, _
                        Nothing, _
                        Nothing, _
                        connSql, _
                        Nothing, _
                        Nothing, _
                        True, _
                        ".\SQLEXPRESS", _
                        dbFolderPath & "\" & IIf(dbName.Trim().ToUpper().EndsWith(".MDF"), dbName, dbName & ".mdf").ToString(), _
                        IIf(dbName.Trim().ToUpper().EndsWith(".MDF"), Left(dbName.Trim().ToUpper(), dbName.Trim().ToUpper().Length - 4), dbName.Trim().ToUpper()).ToString() _
                    )
                    EntLib.COPT.Log.Log("Connected to SQLEXPRESS database successfully - " & dbName)
                Catch ex As Exception
                    EntLib.COPT.Log.Log("Error connecting to SQLEXPRESS database - " & dbName)
                    EntLib.COPT.Log.Log(ex.Message)
                    proceed = False
                End Try
                If proceed Then 'proceed #2
                    'sql = "USE MASTER GO " & vbNewLine
                    'sql = sql & "ALTER DATABASE [" & dbName & "] SET  SINGLE_USER WITH ROLLBACK IMMEDIATE GO " & vbNewLine
                    'sql = sql & "EXEC master.dbo.sp_detach_db @dbname = N'" & dbName & "', @skipchecks = 'false' GO " & vbNewLine

                    sql = "EXEC master.dbo.asp_DetachDatabase"
                    'EXEC master.dbo.sp_detach_db @dbname = N'CV1SQL', @skipchecks = 'false'
                    'sql = "EXEC master.dbo.sp_detach_db @dbname = N'" & IIf(dbName.Trim().ToUpper().EndsWith(".MDF"), Left(dbName.Trim().ToUpper(), dbName.Trim().ToUpper().Length - 4), dbName.Trim().ToUpper()).ToString() & "', @skipchecks = 'false'"
                    'sql = "EXEC master.dbo.sp_detach_db" '@dbname = N'" & IIf(dbName.Trim().ToUpper().EndsWith(".MDF"), Left(dbName.Trim().ToUpper(), dbName.Trim().ToUpper().Length - 4), dbName.Trim().ToUpper()).ToString() & "', @skipchecks = 'false'"
                    EntLib.COPT.Log.Log(sql)
                    Try
                        dt = New DataTable
                        ret = dalSql.Execute(sql, dt, "@dbname", dbName, System.Data.SqlDbType.VarChar)
                        'ret = dalSql.Execute(sql, dt, Nothing)
                        'ret = dalSql.Execute(sql, dt, "@dbname", dbName, System.Data.SqlDbType.VarChar, "@skipchecks", 0, System.Data.SqlDbType.Bit)
                        If ret = -1 Then
                            EntLib.COPT.Log.Log("Error executing SQLEXPRESS DETACH command on: " & dbName)
                        Else
                            EntLib.COPT.Log.Log("Successfully executed SQLEXPRESS DETACH command on: " & dbName)
                        End If
                    Catch ex As Exception
                        EntLib.COPT.Log.Log("Error executing SQLEXPRESS DETACH command on: " & dbName)
                        EntLib.COPT.Log.Log(ex.Message)
                        proceed = False
                    End Try
                    'If proceed Then 'proceed #3
                    '    sql = "ALTER DATABASE [" & dbName & "] SET  SINGLE_USER WITH ROLLBACK IMMEDIATE"
                    '    EntLib.COPT.Log.Log(sql)
                    '    Try
                    '        dt = New DataTable
                    '        ret = dalSql.Execute(sql, dt)
                    '        If ret = -1 Then
                    '            EntLib.COPT.Log.Log("Error executing SQLEXPRESS DETACH command on: " & dbName)
                    '        Else
                    '            EntLib.COPT.Log.Log("Successfully executed SQLEXPRESS DETACH command on: " & dbName)
                    '        End If
                    '    Catch ex As Exception
                    '        EntLib.COPT.Log.Log("Error executing SQLEXPRESS DETACH command on: " & dbName)
                    '        EntLib.COPT.Log.Log(ex.Message)
                    '        proceed = False
                    '    End Try
                    '    If proceed Then 'proceed #4
                    '        'sql = "EXEC master.dbo.sp_detach_db @dbname = N'" & dbName & "', @skipchecks = 'false'"
                    '        sql = "EXEC sys.sp_detach_db '" & dbName & "', 'false'"
                    '        EntLib.COPT.Log.Log(sql)
                    '        Try
                    '            dt = New DataTable
                    '            ret = dalSql.Execute(sql, dt)
                    '            If ret = -1 Then
                    '                EntLib.COPT.Log.Log("Error executing SQLEXPRESS DETACH command on: " & dbName)
                    '            Else
                    '                EntLib.COPT.Log.Log("Successfully executed SQLEXPRESS DETACH command on: " & dbName)
                    '            End If
                    '        Catch ex As Exception
                    '            EntLib.COPT.Log.Log("Error executing SQLEXPRESS DETACH command on: " & dbName)
                    '            EntLib.COPT.Log.Log(ex.Message)
                    '            proceed = False
                    '        End Try
                    '    End If 'proceed #4
                    'End If 'proceed #3
                End If 'proceed #2
            End If 'proceed #1

            EntLib.COPT.Log.Log("SQLEXPRESS - DETACH - Ended at: " & Now())
            EntLib.COPT.Log.Log("********************")

            Return 0
        End Function

        Public Shared Function DB_SQLExpress_RunSQLScript( _
            ByVal dbName As String, _
            ByVal filePathSQLScript As String _
        ) As Integer

            GenUtils.LaunchFile("SQLCMD.EXE", "open", " -S .\SQLEXPRESS -i " & filePathSQLScript, True, True, 0)
            Debug.Print(System.IO.Path.GetFullPath(filePathSQLScript))
            Return 0
        End Function

        Public Property DAL() As clsDAL
            Get
                Return _dal
            End Get
            Set(ByVal value As clsDAL)
                _dal = value
            End Set
        End Property

    End Class
End Namespace