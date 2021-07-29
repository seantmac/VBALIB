Namespace COPT
    ''' <summary>
    ''' An instance of the session object holds C-OPT instance data.
    ''' </summary>
    ''' <remarks></remarks>
    Public Class Session
        Private _switches() As String
        Private _databaseName As String
        Private _lastSql As String
        Private _miscParams As DataRow

        Structure typBlueColType
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
            Dim strQuerySQL As String
        End Structure

        Structure typBlueRowType
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
            Dim strQuerySQL As String
        End Structure
        Dim _currentDb As EntLib.COPT.DAAB
        Dim _workDir As String

        Public Sub New()

        End Sub

        Public Sub New(ByRef switches() As String)
            _switches = switches
        End Sub

        Public Sub New(ByVal databaseName As String, ByVal switches() As String)
            _databaseName = databaseName
            _currentDb = New EntLib.COPT.DAAB(databaseName)
            '_workDir = workDir
            _switches = switches
            _workDir = GenUtils.GetWorkDir(_switches) 'GenUtils.GetSwitchArgument(_switches, "/WorkDir", 1)
            FindModelErrors()
            '_sessionGUID = GenUtils.GenerateGUID()
            '_lpProbInfo = GetProbInfo()
        End Sub

        Public Sub FindModelErrors()
            Dim lastSql As String
            Dim sSQL As String
            Dim dt As DataTable
            Dim ds As New DataSet
            Dim dr As DataRow
            Dim i As Integer
            Dim writer As StreamWriter = New StreamWriter("C:\OPTMODELS\BLUEPRINT_TEST.LOG")
            'Dim ret As Long


            writer.WriteLine("*** File created using StreamWriter class and C-OPT Model BluePrint Module. ***")
            writer.WriteLine("*")
            writer.WriteLine("*")
            writer.Close()


            '===============================================
            ' FIND MISSING PRIMARY MODEL OBJECTS
            '===============================================
            sSQL = ""
            sSQL = sSQL & _
                   "SELECT " & vbCrLf & _
                   "  'MISSING MODEL OBJECT:  ' & [TYPE] & ' ' & [ObjectName] AS BLUEPRINT_ERROR, " & vbCrLf & _
                   "  Q9.TYPE, " & vbCrLf & _
                   "  Q9.ObjectActive, " & vbCrLf & _
                   "  Q9.ModelOBJTYPE, " & vbCrLf & _
                   "  Q9.ModelOBJTYPEdesc, " & vbCrLf & _
                   "  Q9.SourceOfModelObject, " & vbCrLf & _
                   "  Q9.ModelObjectName, " & vbCrLf & _
                   "  Q9.ModelObjectDesc, " & vbCrLf & _
                   "  Q9.ObjectName, " & vbCrLf & _
                   "  qsysMSysObjects.OBJ_NAME, " & vbCrLf & _
                   "  'OBJECTEXISTS' AS ConditionCodeForNormalOper, " & vbCrLf & _
                   "  255 AS CoditionCodeColorINT, " & vbCrLf & _
                   "  'RED' AS CoditionCodeColorDesc " & vbCrLf & _
                   "FROM " & vbCrLf & _
                   "  " & vbCrLf & _
                   "( " & vbCrLf & _
                   "SELECT DISTINCT " & vbCrLf & _
                   "  'PRIMARY' AS TYPE, " & vbCrLf & _
                   "  qsysModelObjects.ModelOBJTYPE, " & vbCrLf & _
                   "  qsysModelObjects.ModelOBJTYPEdesc, " & vbCrLf & _
                   "  qsysModelObjects.ObjectActive, " & vbCrLf & _
                   "  qsysModelObjects.SourceOfModelObject,  " & vbCrLf & _
                   "  qsysModelObjects.ModelObjectName, " & vbCrLf & _
                   "  qsysModelObjects.ModelObjectDesc, " & vbCrLf & _
                   "  qsysModelObjects.RecSetName AS ObjectName " & vbCrLf & _
                   "FROM " & vbCrLf & _
                   "  qsysModelObjects " & vbCrLf & _
                   " " & vbCrLf & _
                   "UNION SELECT DISTINCT " & vbCrLf & _
                   "  'SECONDARY' AS TYPE, " & vbCrLf & _
                   "  qsysModelObjects.ModelOBJTYPE, " & vbCrLf & _
                   "  qsysModelObjects.ModelOBJTYPEdesc, " & vbCrLf & _
                   "  qsysModelObjects.ObjectActive, " & vbCrLf & _
                   "  qsysModelObjects.SourceOfModelObject, " & vbCrLf & _
                   "  qsysModelObjects.ModelObjectName, " & vbCrLf & _
                   "  qsysModelObjects.ModelObjectDesc, " & vbCrLf & _
                   "  ztblReferencedQueries.RefName AS ObjectName " & vbCrLf & _
                   "FROM " & vbCrLf & _
                   "  qsysModelObjects INNER JOIN ztblReferencedQueries ON " & vbCrLf & _
                   "     qsysModelObjects.RecSetName = ztblReferencedQueries.ObjectName " & vbCrLf & _
                   ") Q9 " & vbCrLf & _
                   " " & vbCrLf & _
                   "  LEFT JOIN qsysMSysObjects ON " & vbCrLf & _
                   "  Q9.ObjectName = qsysMSysObjects.OBJ_NAME " & vbCrLf & _
                   "WHERE (((qsysMSysObjects.OBJ_NAME) Is Null)) "

            lastSql = sSQL
            '
            '
            '

            Try
                ds = _currentDb.GetDataSet(lastSql)
                If ds.Tables(0).Rows.Count > 0 Then
                    For i = 0 To ds.Tables(0).Rows.Count - 1
                        dr = ds.Tables(0).Rows(i)
                        Console.WriteLine("*E*  " & dr.Item("BLUEPRINT_ERROR").ToString())
                        writer.WriteLine("*E*  " & dr.Item("BLUEPRINT_ERROR").ToString())
                    Next i
                End If
            Catch ex As Exception
                'MessageBox.Show(ex.Message)
                GenUtils.Message(GenUtils.MsgType.Critical, "Session - GetMiscParams", ex.Message)
            End Try





            '==================================================
            ' LOAD ALL THE BLUEPRINT INFO FOR   R O W S
            '==================================================
            _lastSql = "SELECT * FROM " & _miscParams.Item("LPM_CONSTR_DEF_TABLE_NAME").ToString  'moRSMiscParams!LPM_CONSTR_DEF_TABLE_NAME
            'Call moDAL.Execute("SELECT * FROM " & moRSMiscParams!LPM_CONSTR_DEF_TABLE_NAME, rs, adCmdText, adOpenDynamic)
            Try
                dt = _currentDb.GetDataSet(_lastSql).Tables(0)
                For i = 0 To dt.Rows.Count - 1
                    Application.DoEvents()
                    'rowType(i) = ReadRowType(CInt(dt.Rows(i).Item("RowTypeID")))
                Next
            Catch ex As Exception
                'MessageBox.Show(ex.Message, "C-OPTEngine")
                GenUtils.Message(GenUtils.MsgType.Critical, "Session - FindModelErrors|ROW", ex.Message)
                EntLib.COPT.Log.Log(_workDir, "Session - FindModelErrors|ROW - Error: ", ex.Message)
                'EntLib.COPT.Log.Log(_workDir, "Session - FindModelErrors|ROW - Error: " & rowType(i).strDesc & " - ", _lastSql)
            End Try

            '==================================================
            ' LOAD ALL THE BLUEPRINT INFO FOR   C O L U M N S
            '==================================================
            _lastSql = "SELECT * FROM " & _miscParams.Item("LPM_COLUMN_DEF_TABLE_NAME").ToString  'moRSMiscParams!LPM_CONSTR_DEF_TABLE_NAME
            Try
                dt = _currentDb.GetDataSet(_lastSql).Tables(0)
                For i = 0 To dt.Rows.Count - 1
                    Application.DoEvents()
                    'colType(i) = ReadColType(CInt(dt.Rows(i).Item("ColTypeID").ToString))
                    'colType(i) = Engine.x
                Next
            Catch ex As Exception
                'MessageBox.Show(ex.Message, "C-OPTEngine")
                GenUtils.Message(GenUtils.MsgType.Critical, "Session - FindModelErrors|COL", ex.Message)
                EntLib.COPT.Log.Log(_workDir, "Session - FindModelErrors|COL - Error: ", ex.Message)
                'EntLib.COPT.Log.Log(_workDir, "Session - FindModelErrors|COL - Error: " & rowType(i).strDesc & " - ", _lastSql)
            End Try








        End Sub









    End Class
End Namespace
