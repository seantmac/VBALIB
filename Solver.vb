Imports System.Runtime.InteropServices
Imports System.Text
Imports System
Imports System.Linq
Imports System.Linq.Expressions
Imports System.Collections
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.Linq
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Data.Common
Imports System.Globalization
Imports System.Xml
Imports System.Xml.Linq
Imports System.Threading
Imports Microsoft.VisualBasic
Imports EntLib.COPT

Namespace COPT
    Public MustInherit Class Solver
        'This is an abstract class.
        'The "MustInherit" keyword specifies that the abstract class "Solver" cannot be
        'directly instantiated. The class can be used only if inherited by a subclass.
        'The following is not valid:
        'Dim slvr As New Solver '<--not allowed
        'The purpose of this class is to provide the base properties and methods for subclasses.

        'Ravi Poluri/05-04-07/v1
        '********** START OPTIONS            **********
        '********** END   OPTIONS            **********
        '********** START DLL DECLARATIONS   **********
        '********** END   DLL DECLARATIONS   **********
        '********** START PUBLIC CONSTANTS   **********
        'RKP/02-21-10/v2.3.131
        'Used by LoadSolverArrays
        Public Enum loadType
            fromMemory
            fromDatabase
        End Enum
        Public Enum commonSolutionStatus
            statusOptimal
            statusInfeasible
        End Enum
        '********** END   PUBLIC CONSTANTS   **********
        '********** START PUBLIC VARIABLES   **********
        '********** END   PUBLIC VARIABLES   **********
        '********** START PRIVATE CONSTANTS  **********
        '********** END   PRIVATE CONSTANTS  **********
        '********** START PRIVATE VARIABLES  **********
        Private _lpProbInfo As Solver.typLPprobInfo
        Private _engine As COPT.Engine
        Private _currentSession As EntLib.COPT.Session

        'Protected Friend array_dobj() As Double = Nothing
        'Protected Friend array_dclo() As Double = Nothing
        'Protected Friend array_dcup() As Double = Nothing
        'Protected Friend array_mbeg() As Integer = Nothing
        'Protected Friend array_midx() As Integer = Nothing
        'Protected Friend array_mval() As Double = Nothing
        'Protected Friend array_drhs() As Double = Nothing
        'Protected Friend array_mcnt() As Integer = Nothing
        'Protected Friend array_rtyp() As Char = Nothing
        'Protected Friend array_ctyp() As Char = Nothing
        'Protected Friend array_initValues() As Double = Nothing 'RKP/02-15-10/v2.3.130
        'Protected Friend array_colNames() As String = Nothing
        'Protected Friend array_rowNames() As String = Nothing

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

        Protected Friend _dtRow As DataTable
        Protected Friend _dtCol As DataTable
        Protected Friend _dtMtx As DataTable
        Protected Friend _isMIP As Boolean = False
        '********** END   PRIVATE VARIABLES  **********
        '********** START USER DEFINED TYPES **********
        Protected Friend Structure typLPprobInfo
            Dim strProjectName As String    'Project
            Dim strScenarioName As String   'Scenario/RunName
            Dim intColCt As Integer
            Dim intRowCt As Integer
            Dim intCoefCt As Integer
            Dim intIntegerCt As Integer     'Integer Variable Count
            Dim intColNameLength As Integer 'Stdized Col Name Length
            Dim intRowNameLength As Integer 'Stdized Row Name Length
            Dim intMasterLength As Integer  'Master Element Name Length
            Dim intMaxMemory As Integer
            Dim intMaxItems As Integer      'for XA
            Dim intOBJSense As Short
            Dim strOBJRowName As String
            Dim strRHSName As String
            Dim strBOUNDSName As String
        End Structure

        ''' <summary>
        ''' STM - 9/21/09
        ''' </summary>
        ''' <remarks>
        ''' This structure is used to hold solution that comes back from the solver.
        ''' This standard representation is used by GenSOLFile, regardless of the solver used.
        ''' </remarks>
        Public Structure typLPsolution
            Dim strProjectName As String    'Project
            Dim strScenarioName As String   'Scenario/RunName
            Dim strSolverStatus As String
            Dim dblObjectiveValue As Double
            Dim intIterations As Long
            Dim strObjective As String
            '
            Dim rowID() As Integer
            Dim rowNames() As String
            Dim rowStatus() As String
            Dim rowActivity() As Double
            Dim rowSlack() As Double
            Dim rowLower() As Double
            Dim rowUpper() As Double
            Dim rowRHS() As Double
            Dim rowDual() As Double
            '
            Dim colID() As Integer
            Dim colStatus() As String
            Dim colActivity() As Double
            Dim colObjcoef() As Double
            Dim colLower() As Double
            Dim colUpper() As Double
            Dim colDJ() As Double
        End Structure
        Protected _switches() As String
        Protected _workDir As String
        '********** END   USER DEFINED TYPES **********


        Public Sub New()
            _lpProbInfo = GetProbInfo()
        End Sub

        Public Sub New(ByRef engine As COPT.Engine)
            _engine = engine
        End Sub

        Public Sub New(ByRef engine As COPT.Engine, ByVal switches() As String)
            _engine = engine
            _switches = switches
            _workDir = GenUtils.GetWorkDir(switches)

            'RKP/06-11-10/v2.3.133
            'Make the array's in Solver.vb to reference the array's that are already built in C-OPT Engine (during MatGen).
            If Not engine.array_mcnt Is Nothing Then
                array_mcnt = engine.array_mcnt
                'engine.array_mcnt.CopyTo(array_mcnt, 0)
            End If
            If Not engine.array_mbeg Is Nothing Then
                array_mbeg = engine.array_mbeg
            End If
            If Not engine.array_midx Is Nothing Then
                array_midx = engine.array_midx
            End If
            If Not engine.array_mval Is Nothing Then
                array_mval = engine.array_mval
            End If
        End Sub

        Public MustOverride Function Solve() As Boolean
        Public MustOverride Function Solve(ByRef dtRow As DataTable, ByRef dtCol As DataTable, ByRef dtMtx As DataTable) As Boolean
        Public MustOverride Function Solve(ByRef ds As DataSet) As Boolean
        Public MustOverride Function getSolutionStatusDescription() As String
        Public MustOverride Function getSolutionStatusDescription(ByVal sts As Integer) As String
        Public MustOverride Function getSolverName() As String


        '****************************************************************************************

        Public Function GenMPSFile(ByVal strMPSOPTION As String, ByRef dtRow As DataTable, ByRef dtCol As DataTable, ByRef dtMtx As DataTable) As Boolean
            '//================================================================================//
            '/|   FUNCTION: GenMPSFile()                                                       |/
            '/| PARAMETERS: strMPSOPTION; MPSID, MPS08 ... MPS99 or MPSFREE                    |/
            '/|             ID'S, 8-char names, or other number or 16-char names,              |/
            '/|                99-char names, free length names                                |/
            '/|    RETURNS: True on Success and False by default or Failure                    |/
            '/|    PURPOSE: Create MPS file                                                    |/
            '/|      USAGE: i = GenMPSFile("MPS16")                                            |/
            '/|         BY: Sean                                                               |/
            '/|       DATE: 7/19/07                                                            |/
            '/|    HISTORY:                                                                    |/
            '//================================================================================//
            '* Then set it up so there is another option to either start with C1115_  
            '*       or just use the name in COL
            '* Set up timer to debug performance
            '* See if it's necessary to check the PROJECT_SCENARIO name is not too long

            Dim lpMPS As typLPprobInfo
            Dim sb As New System.Text.StringBuilder()
            Dim ctr As Integer
            Dim intlen As Integer
            Dim strlen As String
            Dim colName As String
            Dim strdbg As String
            Dim sizeNameRowCol As String
            Dim intLengthToUse As Integer
            'Dim dsRCC As New DataSet()
            'Dim dtCol As New DataTable()
            'Dim dtRow As New DataTable()
            'Dim dtMtx As New DataTable()
            Dim strDebug As String
            Dim strRHSstring As String
            Dim strBNDstring As String
            Dim strTempRow As String
            Dim startTime As Integer = My.Computer.Clock.TickCount
            Dim engine As COPT.Engine = _engine

            '// INIT
            If strMPSOPTION = "" Then
                strMPSOPTION = GenUtils.GetAppSettings("strMPSOPTION")
            End If
            strRHSstring = ""
            strBNDstring = ""
            strdbg = ""
            If dtCol Is Nothing Then
                dtCol = engine.CurrentDb.GetDataTable("SELECT * FROM tsysCOL")              'SESSIONIZE
            End If
            If dtRow Is Nothing Then
                dtRow = engine.CurrentDb.GetDataTable("SELECT * FROM tsysROW")              'SESSIONIZE
            End If
            If dtMtx Is Nothing Then
                dtMtx = engine.CurrentDb.GetDataTable("SELECT * FROM qsysMTXwithCOLS")      'SESSIONIZE
            End If
            dtCol.TableName = "tsysCOL"
            dtRow.TableName = "tsysROW"
            dtMtx.TableName = "qsysMTXwithCOLS"

            'dsRCC.Tables.Add(dtCol)
            'dsRCC.Tables(0).TableName = "tsysCOL"               'SESSIONIZE
            'dsRCC.Tables.Add(dtRow)
            'dsRCC.Tables(1).TableName = "tsysROW"               'SESSIONIZE
            'dsRCC.Tables.Add(dtMtx)
            'dsRCC.Tables(2).TableName = "qsysMTXwithCOLS"       'SESSIONIZE

            lpMPS = GetProbInfo()
            strlen = Right(strMPSOPTION, Len(strMPSOPTION) - 3)  'Take all to the right of the "MPS"

            If strlen.ToUpper = "FREE" Then
                intlen = lpMPS.intMasterLength
            ElseIf strlen.ToUpper = "ID" Then
                intlen = 16
            Else
                intlen = CInt(strlen)
            End If

            If intlen >= 100 Then intlen = 9999

            'If dsRCC.Tables.Count.Equals(3) Then
            Try
                My.Computer.FileSystem.DeleteFile(GenUtils.GetWorkDir(_switches) & "\" & lpMPS.strProjectName & "_" & lpMPS.strScenarioName & ".MPS")      'SESSIONIZE?
            Catch ex As Exception       'no need to trap this error.  if the file exists, delete it
            End Try

            intLengthToUse = intlen

            'stm add path provision...

            '//  Print the MPS HEADER INFO (all just comments in the MPS file)
            sb.Append("*NAME:" & Space(9) & lpMPS.strProjectName & "_") : sb.AppendLine(lpMPS.strScenarioName)      'SESSIONIZE?
            sb.Append("*ROWS:" & Space(9) & "(" & lpMPS.intRowCt & ")") : sb.AppendLine("")
            sb.Append("*COLUMNS:" & Space(6) & "(" & lpMPS.intColCt & ")") : sb.AppendLine("")
            sb.Append("*INTEGER:" & Space(6) & "(" & lpMPS.intIntegerCt & ")") : sb.AppendLine("")
            sb.Append("*NONZERO:" & Space(6) & "(" & lpMPS.intCoefCt & ")") : sb.AppendLine("")

            'stm add some typical results functionality here in the next two lines
            sb.Append("*BEST SOLN:" & Space(4)) : sb.AppendLine("")         'SESSIONIZE
            sb.Append("*LP SOLN:" & Space(6)) : sb.AppendLine("")           'SESSIONIZE

            sb.Append("*MPS OPTION:" & Space(3)) : sb.AppendLine(strMPSOPTION & " (MasterLength=" & intLengthToUse & ") ")

            'stm grab this junk from somewhere in the database
            sb.Append("*SOURCE:" & Space(7)) : sb.AppendLine("Sean MacDermant (INTERNATIONAL PAPER)")       'SESSIONIZE?
            sb.Append("*       " & Space(7)) : sb.AppendLine("Ravi Poluri     (INTERNATIONAL PAPER)")       'SESSIONIZE?
            sb.Append("*APPLICATION:" & Space(2)) : sb.AppendLine("C-OPT2")                                 'SESSIONIZE?
            sb.Append("*COMMENTS:" & Space(5)) : sb.AppendLine("Capacity Optimization Model")               'SESSIONIZE
            sb.Append("*         " & Space(5)) : sb.AppendLine("")
            sb.Append("*         " & Space(5)) : sb.AppendLine("")

            '//  Print the PROBLEM NAME (Scenario Instance) Name Line
            sb.Append(engine.LP_NAME & Space(11)) : sb.AppendLine(lpMPS.strProjectName & "_" & lpMPS.strScenarioName)  'SESSIONIZE

            '//  Print the ROWS SECTION
            sb.AppendLine(engine.LP_ROW_HEAD)

            '//  Print the Objective row
            'stm fix this. it should not be hardwired
            sb.Append(Space(1) & engine.LP_OBJ_TYPE & Space(2)) : sb.AppendLine(lpMPS.strOBJRowName)

            '//  Print the rows, and build the strRHSstring, which is the RHS section but is appended later
            For ctr = 0 To dtRow.Rows.Count - 1
                Application.DoEvents()
                sb.Append(Space(1) & dtRow.Rows(ctr)("SENSE").ToString() & Space(2))
                strRHSstring = strRHSstring & Space(4) & lpMPS.strRHSName & Space(intLengthToUse - Len(lpMPS.strRHSName)) & engine.LP_SPACE5
                Select Case strMPSOPTION.ToUpper
                    Case "MPS" & "08" To "MPS" & "99"
                        strdbg = "case " & strMPSOPTION & " worked"
                        sb.AppendLine("R" & dtRow.Rows(ctr)("RowID").ToString() & "_" & _
                            Left(dtRow.Rows(ctr)("ROW").ToString(), intlen - Len("R" & dtRow.Rows(ctr)("RowID").ToString() & "_")))
                        strRHSstring = strRHSstring & SizeName("R" & dtRow.Rows(ctr)("RowID").ToString() & "_" & _
                            Left(dtRow.Rows(ctr)("ROW").ToString(), intlen - Len("R" & dtRow.Rows(ctr)("RowID").ToString() & "_")), intLengthToUse)
                    Case "MPSFREE"
                        sb.AppendLine(dtRow.Rows(ctr)("ROW").ToString())
                        strRHSstring = strRHSstring & SizeName(dtRow.Rows(ctr)("ROW").ToString(), intLengthToUse)
                    Case Else 'MPSID and any other attempted strMPSOPTION
                        sb.AppendLine("R" & dtRow.Rows(ctr)("RowID").ToString())
                        strRHSstring = strRHSstring & SizeName("R" & dtRow.Rows(ctr)("RowID").ToString(), intLengthToUse)
                End Select
                strRHSstring = strRHSstring & engine.LP_SPACE5 & dtRow.Rows(ctr)("RHS").ToString() & vbCrLf
            Next

            strDebug = "MPS GEN:  ROWS " & GenUtils.FormatTime(startTime, My.Computer.Clock.TickCount)
            Debug.Print(strDebug)

            '//  Print the Columns Section
            sb.AppendLine(engine.LP_COL_HEAD)
            colName = String.Empty
            For ctr = 0 To dtMtx.Rows.Count - 1
                Application.DoEvents()
                sizeNameRowCol = SizeName(dtMtx.Rows(ctr)("COL").ToString(), intLengthToUse)
                Select Case strMPSOPTION.ToUpper
                    Case "MPS" & "08" To "MPS" & "99"
                        sizeNameRowCol = SizeName("C" & dtMtx.Rows(ctr)("ColID").ToString() & "_" & _
                            Left(dtMtx.Rows(ctr)("COL").ToString(), intlen - Len("C" & dtMtx.Rows(ctr)("ColID").ToString() & "_")), intLengthToUse)
                    Case "MPSFREE"
                        sizeNameRowCol = SizeName(dtMtx.Rows(ctr)("COL").ToString(), intLengthToUse)
                    Case Else 'MPSID and any other attempted strMPSOPTION
                        sizeNameRowCol = SizeName("C" & dtMtx.Rows(ctr)("ColID").ToString(), intLengthToUse)
                End Select

                '// FIRST INSTANCE OF THIS COLUMN, SO ADD OBJ COEFFICIENT AND BUILD BOUNDS SECTION TEXT
                If Not colName.Equals(sizeNameRowCol) Then
                    '//  Print the ColName and the OBJ row name and the OBJ coefficient value
                    sb.Append(Space(4) & sizeNameRowCol & engine.LP_SPACE5 & SizeName(lpMPS.strOBJRowName, intLengthToUse) & engine.LP_SPACE5)
                    sb.AppendLine(dtMtx.Rows(ctr)("OBJ").ToString())
                    colName = String.Copy(sizeNameRowCol)
                    '//  And now build the strBNDstring, which is the BOUNDS section but is appended later

                    If dtMtx.Rows(ctr)("FREE").Equals(True) Then
                        strBNDstring = strBNDstring & "" & Space(1) & "FR" & Space(1) & lpMPS.strBOUNDSName & _
                           Space(intLengthToUse - Len(lpMPS.strBOUNDSName)) & engine.LP_SPACE5
                        strBNDstring = strBNDstring & "" & sizeNameRowCol & engine.LP_SPACE5
                        strBNDstring = strBNDstring & "" & dtMtx.Rows(ctr)("UP").ToString() & vbCrLf
                    Else
                        'print LOWER bound
                        strBNDstring = strBNDstring & "" & Space(1) & "LO" & Space(1) & lpMPS.strBOUNDSName & _
                           Space(intLengthToUse - Len(lpMPS.strBOUNDSName)) & engine.LP_SPACE5
                        strBNDstring = strBNDstring & "" & sizeNameRowCol & engine.LP_SPACE5
                        strBNDstring = strBNDstring & "" & dtMtx.Rows(ctr)("LO").ToString() & vbCrLf
                        'print UPPER bound
                        strBNDstring = strBNDstring & "" & Space(1) & "UP" & Space(1) & lpMPS.strBOUNDSName & _
                           Space(intLengthToUse - Len(lpMPS.strBOUNDSName)) & engine.LP_SPACE5
                        strBNDstring = strBNDstring & "" & sizeNameRowCol & engine.LP_SPACE5
                        strBNDstring = strBNDstring & "" & dtMtx.Rows(ctr)("UP").ToString() & vbCrLf
                    End If
                End If

                Select Case strMPSOPTION.ToUpper
                    Case "MPS" & "08" To "MPS" & "99"
                        strTempRow = SizeName("R" & dtMtx.Rows(ctr)("RowID").ToString() & "_" & _
                            Left(dtMtx.Rows(ctr)("ROW").ToString(), intlen - Len("R" & dtMtx.Rows(ctr)("RowID").ToString() & "_")), intLengthToUse)
                    Case "MPSFREE"
                        strTempRow = SizeName(dtMtx.Rows(ctr)("ROW").ToString(), intLengthToUse)
                    Case Else 'MPSID and any other attempted strMPSOPTION
                        strTempRow = SizeName("R" & dtMtx.Rows(ctr)("RowID").ToString(), intLengthToUse)
                End Select
                sb.Append(Space(4) & sizeNameRowCol & engine.LP_SPACE5 & strTempRow)
                sb.Append(engine.LP_SPACE5)
                sb.AppendLine(dtMtx.Rows(ctr)("COEF").ToString())
            Next

            strDebug = "MPS GEN:  COLS " & GenUtils.FormatTime(startTime, My.Computer.Clock.TickCount)
            Debug.Print(strDebug)

            '//  Print the Rows RHS Section Header
            sb.AppendLine(engine.LP_RHS_HEAD)
            sb.Append(strRHSstring)

            '//  Print the BOUNDS Section
            sb.AppendLine(engine.LP_BND_HEAD)
            sb.Append(strBNDstring)

            '//  Print the ENDATA Card
            sb.AppendLine(engine.LP_END_HEAD)

            strDebug = "MPS GEN: DONE " & GenUtils.FormatTime(startTime, My.Computer.Clock.TickCount)
            Debug.Print(strDebug)

            'Write the contents of sb to output file
            My.Computer.FileSystem.WriteAllText(GenUtils.GetWorkDir(_switches) & "\" & lpMPS.strProjectName & "_" & lpMPS.strScenarioName & ".MPS", _
                sb.ToString(), False, System.Text.Encoding.ASCII)
            Debug.Print("MPS file was successfully created:  " & lpMPS.strProjectName & "_" & lpMPS.strScenarioName & ".MPS")
            EntLib.COPT.Log.Log(_workDir, "Status", "MPS file was successfully created:  " & lpMPS.strProjectName & "_" & lpMPS.strScenarioName & ".MPS")
            Debug.Print(strdbg)
            Return True
            'Else
            '    Return False
            'End If 'If dsRCC.Tables.Count = 3 Then

            strDebug = "MPS GEN: SAVE " & GenUtils.FormatTime(startTime, My.Computer.Clock.TickCount)
            Debug.Print(strDebug)

        End Function

        Private Function GetProbInfo() As typLPprobInfo
            Dim lp As typLPprobInfo
            'Dim engine As COPT.Engine = New COPT.Engine("BR1")     'SESSIONIZE don't new engine here, use the current project.

            Try
                lp.strProjectName = _engine.MiscParams.Item("PRJ_NAME").ToString() '"PROJECTZERO"  '"BR1"      _engine.MiscParams.Item("PRJ_NAME").ToString                                             'SESSIONIZE
            Catch ex As Exception
                lp.strProjectName = "_EMPTY_"
            End Try

            Try
                lp.strScenarioName = _engine.MiscParams.Item("RUN_NAME").ToString() '_engine.CurrentDb.GetScalarValue("SELECT RUN_NAME FROM qsysMiscParams")            'SESSIONIZE
            Catch ex As Exception
                lp.strScenarioName = "_EMPTY_"
            End Try


            Try
                lp.intColCt = CInt(_engine.CurrentDb.GetScalarValue("SELECT COUNT(*) FROM " & Trim(_engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString)))
                lp.intRowCt = CInt(_engine.CurrentDb.GetScalarValue("SELECT COUNT(*) FROM " & Trim(_engine.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString)))
                lp.intCoefCt = CInt(_engine.CurrentDb.GetScalarValue("SELECT COUNT(*) FROM " & Trim(_engine.MiscParams.Item("LPM_MATRIX_TABLE_NAME").ToString)))
                lp.intIntegerCt = CInt(_engine.CurrentDb.GetScalarValue("SELECT COUNT(*) FROM " & Trim(_engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString) & _
                                                                          " WHERE INTGR = TRUE"))
                lp.intColNameLength = CInt(_engine.CurrentDb.GetScalarValue("SELECT MAX(LEN(COL)) FROM " & Trim(_engine.MiscParams.Item("LPM_COLUMN_TABLE_NAME").ToString)))
                lp.intRowNameLength = CInt(_engine.CurrentDb.GetScalarValue("SELECT Max(LEN(ROW)) FROM " & Trim(_engine.MiscParams.Item("LPM_CONSTR_TABLE_NAME").ToString)))

            Catch ex As Exception

            End Try

            ''lp.strProjectName = "PROJECTZERO"                                                               'SESSIONIZE
            ''lp.strScenarioName = CurrentDb.GetScalarValue("SELECT RUN_NAME FROM qsysMiscParams")            'SESSIONIZE
            ''lp.intColCt = CurrentDb.GetScalarValue("SELECT COUNT(*) FROM tsysCol")                          'SESSIONIZE     32  (3)
            ''lp.intRowCt = CurrentDb.GetScalarValue("SELECT COUNT(*) FROM tsysRow")                          'SESSIONIZE     30  (4)
            ''lp.intCoefCt = CurrentDb.GetScalarValue("SELECT COUNT(*) FROM tsysMtx")                         'SESSIONIZE
            ''lp.intIntegerCt = CurrentDb.GetScalarValue("SELECT COUNT(*) FROM tsysCol WHERE INTGR = TRUE")   'SESSIONIZE
            ''lp.intColNameLength = CurrentDb.GetScalarValue("SELECT MAX(LEN(COL)) FROM [tsysCOL]")           'SESSIONIZE
            ''lp.intRowNameLength = CurrentDb.GetScalarValue("SELECT Max(LEN(ROW)) FROM [tsysROW] ")          'SESSIONIZE

            If lp.intRowNameLength <= lp.intColNameLength Then
                lp.intMasterLength = lp.intColNameLength
            Else
                lp.intMasterLength = lp.intRowNameLength
            End If

            lp.intMaxMemory = lp.intColCt + lp.intCoefCt + lp.intRowCt + 1
            lp.intMaxItems = 0
            lp.intOBJSense = -1 ' -1 = MAX, 0 = MIN
            lp.strOBJRowName = "PROFIT"
            lp.strRHSName = "RHS1"
            lp.strBOUNDSName = "BOUND1"
            Return lp
        End Function


        '****************************************************************************************
        Private Function SizeName(ByVal name As String, ByVal nameLength As Integer) As String
            Dim sb As New System.Text.StringBuilder

            If name.Length >= nameLength Then
                sb.Append(name)
                Return sb.ToString()
            Else
                sb.Append(name)
                sb.Append(Space(nameLength - name.Length))
                Return sb.ToString()
            End If
        End Function

        Protected Friend ReadOnly Property ProbInfo() As typLPprobInfo
            Get
                Return GetProbInfo()
            End Get
        End Property

        'Protected Friend ReadOnly Property Engine() As COPT.Engine
        '    Get
        '        Return _engine
        '    End Get
        'End Property

        Public Property Engine() As COPT.Engine
            Get
                Return _engine
            End Get
            Set(ByVal engine As COPT.Engine)
                _engine = engine
            End Set
        End Property

        Public Function getSolutionStatus(ByVal statusCode As Short) As String
            Select Case statusCode
                Case 1
                    Return "Optimal solution"
                Case 2
                    Return "Integer solution"
                Case 3
                    Return "Unbound solution"
                Case 4
                    Return "Infeasible solution"
                Case 5
                    Return "Callback/presolve indicates infeasible solution"
                Case 10
                    Return "Integer infeasible solution"
                Case 12
                    Return "Error unknown"
                Case Else
                    Return "Error unknown"
            End Select
        End Function

        Public Sub AddIdentityColumn(ByRef dt As DataTable, ByVal newColName As String)
            'Dim newCol As DataColumn = dt.Columns.Add("tmpID", Type.GetType("System.Int32"))
            Dim newCol As DataColumn = dt.Columns.Add(newColName, Type.GetType("System.Int32"))
            newCol.AutoIncrement = True
            'newCol.AllowDBNull = False
            'newCol.Unique = True
            For Each dr As DataRow In dt.Rows
                dr.Item(newColName) = dt.Rows.IndexOf(dr)
            Next
            '_ds.Tables(0).Columns.Remove(newCol)
        End Sub

        Public Function GenSOLFile(ByRef dtRow As DataTable, ByRef dtCol As DataTable) As Boolean
            'This method generates a solution (SOL) file, regardless of the solver used.
            'STM/RKP - 9/21/09

            Dim probInfo As typLPprobInfo = GetProbInfo()
            Dim outputFile As String = GenUtils.GetSwitchArgument(_switches, "/WorkDir", 1) & "\C-OPT.sol"
            Dim sHdr As String
            Dim a As Long  'Integer
            Dim b As Long  'Integer
            'Dim sb As StringBuilder

            'C-OPT version
            Dim ver As String = My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & " Build " & My.Application.Info.Version.Build  '& " (" & fileTimeStamp.ToString() & ")" '(US Central Time)"
            Dim sdb As String

            Dim linqTable As System.Data.EnumerableRowCollection(Of System.Data.DataRow)
            Dim strQueryResults As System.Data.EnumerableRowCollection(Of String)

            sdb = _engine.CurrentDb.GetDBConnectionString
            a = InStr(sdb, "source=") + 7
            sdb = Right(sdb, Len(sdb) - a + 1)
            b = InStr(sdb, ";")
            sdb = Left(sdb, b - 1)
            sdb = sdb.ToUpper
            
            Debug.Print(_engine.SolverVersion)  '            'Use _engine for anything else
            Debug.Print("- - - - -")

            sHdr = Nothing
            sHdr = "   " & vbCrLf

            'sHdr = sHdr & "                                                                                                   1         1         1         1         1         1         1         1" & vbCrLf
            'sHdr = sHdr & "         1         2         3         4         5         6         7         8         9         0         1         2         3         4         5         6         7" & vbCrLf
            'sHdr = sHdr & "12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890" & vbCrLf
            sHdr = sHdr & vbCrLf
            sHdr = sHdr & vbCrLf
            sHdr = sHdr & "C-OPT.  " & "Version = " & ver & vbCrLf
            sHdr = sHdr & vbCrLf
            sHdr = sHdr & vbCrLf
            '
            sHdr = sHdr & "PROJECT NAME       " & probInfo.strProjectName & vbCrLf
            sHdr = sHdr & "PROBLEM NAME       " & probInfo.strScenarioName & vbCrLf
            sHdr = sHdr & "DATABASE NAME      " & sdb & vbCrLf
            sHdr = sHdr & vbCrLf
            sHdr = sHdr & "DATE               " & Microsoft.VisualBasic.Format(Now(), "MMM dd, yyyy") & vbCrLf
            sHdr = sHdr & "TIME               " & Microsoft.VisualBasic.Format(Now(), "SHORT TIME") & vbCrLf
            sHdr = sHdr & "SOLVER NAME        " & _engine.SolverName & "  Version  " & _engine.SolverVersion & vbCrLf
            sHdr = sHdr & "SOLUTION TIME      " & _engine.SolutionTime & vbCrLf
            sHdr = sHdr & "RESULT             " & _engine.SolutionStatus & "   " & " RESULT CODE = " & "XX" & vbCrLf
            sHdr = sHdr & vbCrLf
            sHdr = sHdr & "OBJECTIVE VALUE    " & _engine.SolutionObj & vbCrLf
            sHdr = sHdr & "STATUS             " & _engine.SolutionStatus & vbCrLf
            sHdr = sHdr & "ITERATIONS         " & _engine.SolutionIterations & vbCrLf
            sHdr = sHdr & vbCrLf
            sHdr = sHdr & "PROBLEM SIZE       " & _engine.SolutionColumns & " COLS x " & _engine.SolutionRows & " ROWS (" & _engine.SolutionNonZeros & ")" & vbCrLf
            sHdr = sHdr & "DENSITY            " & Microsoft.VisualBasic.Format((_engine.SolutionNonZeros / (_engine.SolutionColumns * _engine.SolutionRows) * 100), "0.00") & " %" & vbCrLf
            sHdr = sHdr & vbCrLf
            sHdr = sHdr & "OBJECTIVE          " & probInfo.strOBJRowName & "(MAX)" & vbCrLf
            sHdr = sHdr & vbCrLf
            sHdr = sHdr & vbCrLf
            sHdr = sHdr & vbCrLf
            sHdr = sHdr & "SECTION 1 - COLUMNS" & vbCrLf
            sHdr = sHdr & vbCrLf
            sHdr = sHdr & "  NUMBER  " & "....COLUMN.....".ToString().PadRight(probInfo.intMasterLength, CChar(".")) & "  AT  " & _
                                       "....ACTIVITY....  ...OBJ COEF....  " & _
                                       ".....LOWER.....  .....UPPER.....  ....REDUCED COST...." & vbCrLf
            sHdr = sHdr & vbCrLf

            Try
                My.Computer.FileSystem.WriteAllText(outputFile, sHdr, False)
            Catch ex As Exception
                EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - GenSOLFile Error - " & ex.Message)
            End Try

            Try
                linqTable = dtCol.AsEnumerable()

                strQueryResults = _
                   From r _
                   In linqTable _
                   Order By r("ColID") Ascending _
                   Select CStr( _
                      r("ColID").ToString().PadLeft(8) & "  " + _
                      r("COL").ToString().PadRight(probInfo.intMasterLength) & "  " + _
                      ("XX").ToString().PadLeft(2) & "  " + _
                      r("ACTIVITY").ToString().PadLeft(16) & "  " + _
                      r("OBJ").ToString().PadLeft(15) & "  " + _
                      r("LO").ToString().PadLeft(15) & "  " + _
                      r("UP").ToString().PadLeft(15) & "  " + _
                      r("DJ").ToString().PadLeft(15) & "  " _
                   )

                My.Computer.FileSystem.WriteAllText(outputFile, String.Join(Environment.NewLine, _
                                                                            strQueryResults.ToArray), True)
            Catch ex As Exception
                'MessageBox.Show(ex.Message)
                'GenUtils.Message(GenUtils.MsgType.Warning, "Solver - GenSOLFile", ex.Message)
                EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - GenSOLFile - " & ex.Message)
            End Try


            sHdr = "" & vbCrLf
            sHdr = sHdr & "" & vbCrLf
            sHdr = sHdr & "" & vbCrLf
            sHdr = sHdr & "SECTION 2 - ROWS" & vbCrLf
            sHdr = sHdr & vbCrLf
            sHdr = sHdr & "  NUMBER  " & "......ROW......".ToString().PadRight(probInfo.intMasterLength, CChar(".")) & "  AT  " & _
                                       "....ACTIVITY....  ......SLACK......  " & _
                                       ".....LOWER.....  .....UPPER.....  ......DUAL......." & vbCrLf
            sHdr = sHdr & vbCrLf

            Try
                My.Computer.FileSystem.WriteAllText(outputFile, sHdr, True)
            Catch ex As Exception
                EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - GenSOLFile Error - " & ex.Message)
            End Try


            Try
                linqTable = dtRow.AsEnumerable()

                strQueryResults = _
                   From r _
                   In linqTable _
                   Order By r("RowID") Ascending _
                   Select CStr( _
                      r("RowID").ToString().PadLeft(8) & "  " + _
                      r("ROW").ToString().PadRight(probInfo.intMasterLength) & "  " + _
                      ("XX").ToString().PadLeft(2) & "  " + _
                      r("ACTIVITY").ToString().PadLeft(16) & "  " + _
                      (r("RHS") - r("ACTIVITY")).ToString().PadLeft(17) & "  " + _
                      r("RHS").ToString().PadLeft(15) & "  " + _
                      r("RHS").ToString().PadLeft(15) & "  " + _
                      r("SHADOW").ToString().PadLeft(17) & "  " _
                   )

                My.Computer.FileSystem.WriteAllText(outputFile, String.Join(Environment.NewLine, _
                                                                            strQueryResults.ToArray), True)
            Catch ex As Exception
                'MessageBox.Show(ex.Message)
                'GenUtils.Message(GenUtils.MsgType.Warning, "Solver - GenSOLFile", ex.Message)
                EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - GenSOLFile Error - " & ex.Message)
                'GenUtils.Message(GenUtils.MsgType.Information, "Solver - GenSOLFile", ex.Message)
                'GenUtils.Message(GenUtils.MsgType.Critical, "Solver - GenSOLFile", ex.Message)
            End Try

            sHdr = "" & vbCrLf
            sHdr = sHdr & "" & vbCrLf
            sHdr = sHdr & "" & vbCrLf
            sHdr = sHdr & "Copyright International Paper 2005 - 2010.  All rights reserved." & vbCrLf
            sHdr = sHdr & "Global Supply Chain  |  Center Of Excellence  |  Planning and Scheduling  |  Business Modeling and Optimization" & vbCrLf
            sHdr = sHdr & vbCrLf
            'sHdr = sHdr & "---------------------------------------------------------" & vbCrLf
            'sHdr = sHdr & "  .g8~~~bgd         .g8~~8q.   `7MM~~~Mq. MMP~~MM~~YMM   " & vbCrLf
            'sHdr = sHdr & ".dP'     `M       .dP'    `YM.   MM   `MM.P'   MM   `7   " & vbCrLf
            'sHdr = sHdr & "dM'       `       dM'      `MM   MM   ,M9      MM        " & vbCrLf
            'sHdr = sHdr & "MM                MM        MM   MMmmdM9       MM        " & vbCrLf
            'sHdr = sHdr & "MM.         mmmmm MM.      ,MP   MM            MM        " & vbCrLf
            'sHdr = sHdr & "`Mb.     ,'       `Mb.    ,dP'   MM            MM      ,," & vbCrLf
            'sHdr = sHdr & "  `~bmmmd'          `~bmmd~'   .JMML.        .JMML.    db" & vbCrLf
            'sHdr = sHdr & "---------------------------------------------------------" & vbCrLf
            sHdr = sHdr & vbCrLf

            Try
                My.Computer.FileSystem.WriteAllText(outputFile, sHdr, True)
            Catch ex As Exception
                EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "Solver - GenSOLFile Error - " & ex.Message)
            End Try


        End Function

        ''' <summary>
        ''' This function loads all the arrays required by a solver's callable library.
        ''' </summary>
        ''' <param name="ds"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' RKP/01-26-10/v2.3.127
        ''' http://www.manning-source.com/books/marguerie/bonusch14.pdf
        ''' </remarks>
        Public Function LoadSolverArrays(ByRef ds As DataSet, ByVal disposeMatrix As Boolean) As Integer
            Return LoadSolverArrays(ds.Tables("tsysRow"), ds.Tables("tsysCol"), ds.Tables("tsysMtx"), disposeMatrix)
        End Function

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="dtRow"></param>
        ''' <param name="dtCol"></param>
        ''' <param name="dtMtx"></param>
        ''' <param name="disposeMatrix"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' This is the LoadSolverArrays function that is currently being called.
        ''' </remarks>
        Public Function LoadSolverArrays _
        ( _
            ByRef dtRow As DataTable, _
            ByRef dtCol As DataTable, _
            ByRef dtMtx As DataTable, _
            ByVal disposeMatrix As Boolean _
        ) As Integer

            Dim cumTotal As Integer = 0
            Dim myArrayList As ArrayList = Nothing
            Dim tblMtx = dtMtx.AsEnumerable()
            Dim tblRow = dtRow.AsEnumerable()
            Dim tblCol = dtCol.AsEnumerable()
            Dim linqTable As System.Data.EnumerableRowCollection(Of System.Data.DataRow) = Nothing
            Dim linqTableMtx As System.Data.EnumerableRowCollection(Of System.Data.DataRow) = Nothing
            Dim linqTableRow As System.Data.EnumerableRowCollection(Of System.Data.DataRow) = Nothing
            Dim linqTableCol As System.Data.EnumerableRowCollection(Of System.Data.DataRow) = Nothing
            Dim dblQueryResults As System.Data.EnumerableRowCollection(Of Double) = Nothing
            Dim intQueryResults As System.Data.EnumerableRowCollection(Of Integer) = Nothing
            Dim charQueryResults As System.Data.EnumerableRowCollection(Of Char) = Nothing
            Dim strQueryResults As System.Data.EnumerableRowCollection(Of String) = Nothing
            Dim intGroups As System.Collections.Generic.IEnumerable(Of Integer) = Nothing

            Dim plinq As ParallelEnumerable = Nothing

            Dim proceed As Boolean = False

            '_dtRow = dtRow
            '_dtCol = dtCol
            '_dtMtx = dtMtx

            Dim usePLINQ As Boolean = GenUtils.IsSwitchAvailable(_switches, "/UsePLINQ")

            Debug.Print("Loading solver arrays into memory...")
            Console.Write("Loading solver arrays into memory...")

            linqTableMtx = tblMtx
            linqTableRow = tblRow
            If GenUtils.IsSwitchAvailable(_switches, "/UseMinSysRes1") Then
                'SELECT * FROM tsysCol WHERE ColID IN (SELECT DISTINCT tsysMTX.ColID FROM tsysMTX INNER JOIN tsysCol ON tsysMTX.ColID = tsysCol.ColID WHERE tsysMTX.COEF <> 0) ORDER BY ColID

                'SELECT * FROM tsysCol WHERE 
                'ColID IN 
                '(SELECT DISTINCT tsysMTX.ColID 
                'FROM tsysMTX 
                'INNER JOIN tsysCol 
                'ON tsysMTX.ColID = tsysCol.ColID 
                'WHERE tsysMTX.COEF <> 0) 
                'ORDER BY ColID

                'SELECT DISTINCT tsysCOL.[ColID], tsysCOL.[COL], tsysCOL.[DESC], tsysCOL.[OBJ], tsysCOL.[LO], tsysCOL.[UP], 
                'tsysCOL.[FREE], tsysCOL.[INTGR], tsysCOL.[BINRY], tsysCOL.[SOSTYPE], tsysCOL.[SOSMARKER],
                'tsysCOL.[ACTIVITY], tsysCOL.[DJ], tsysCOL.[STATUS]
                'FROM tsysMTX INNER JOIN tsysCol ON tsysMTX.ColID = tsysCol.ColID WHERE tsysMTX.COEF <> 0

                'linqTableCol = _
                '    dtMtx.AsEnumerable().Join(dtCol.AsEnumerable(), Function(m) m.Field(Of Integer)("ColID"), _
                '    Function(c) c.Field(Of Integer)("ColID"), _
                '    Function(m, c) _
                '        New With {.ColID = c.Field(Of Integer)("ColID"), _
                '        .COL = c.Field(Of String)("COL"), _
                '        .DESC = c.Field(Of String)("DESC"), _
                '        .OBJ = c.Field(Of Double)("OBJ"), _
                '        .LO = c.Field(Of Double)("LO"), _
                '        .UP = c.Field(Of Double)("UP"), _
                '        .FREE = c.Field(Of Double)("FREE"), _
                '        .INTGR = c.Field(Of Double)("INTGR"), _
                '        .BINRY = c.Field(Of Double)("BINRY"), _
                '        .SOSTYPE = c.Field(Of Double)("SOSTYPE"), _
                '        .SOSMARKER = c.Field(Of Double)("SOSMARKER"), _
                '        .ACTIVITY = c.Field(Of Double)("ACTIVITY"), _
                '        .DJ = c.Field(Of Double)("DJ"), _
                '        .STATUS = c.Field(Of Double)("STATUS") _
                '    })

                linqTableCol = _
                    From _
                        m In tblMtx, _
                        c In tblCol _
                    Where _
                        m!ColID = c!ColID _
                        And _
                        m!COEF <> 0 _
                    Select _
                        c!ColID, _
                        c!COL, _
                        c!DESC, _
                        c!OBJ, _
                        c!LO, _
                        c!UP, _
                        c!FREE, _
                        c!INTGR, _
                        c!BINRY, _
                        c!SOSTYPE, _
                        c!SOSMARKER, _
                        c!ACTIVITY, _
                        c!DJ, _
                        c!STATUS
            Else
                linqTableCol = tblCol
            End If

            '---SENSE---dtRow
            If array_rtyp Is Nothing Then
                'linqTable = dtRow.AsEnumerable()
                If usePLINQ Then
                    'linqTable.AsParallel()
                    charQueryResults = From r In linqTableRow _
                                   Order By r("RowID") _
                                   Select SENSE = CChar(r("SENSE"))
                Else
                    charQueryResults = From r In linqTableRow _
                                   Order By r("RowID") _
                                   Select SENSE = CChar(r("SENSE"))
                End If
                array_rtyp = charQueryResults.ToArray()
            End If

            '---RHS---dtRow
            If array_dobj Is Nothing Then
                'linqTable = dtRow.AsEnumerable()
                If usePLINQ Then linqTableRow.AsParallel()
                dblQueryResults = From r In linqTableRow _
                               Order By r("RowID") _
                               Select RHS = CDbl(r("RHS"))
                array_drhs = dblQueryResults.ToArray()
            End If

            '---OBJ---dtCol
            If array_dobj Is Nothing Then
                'linqTable = dtCol.AsEnumerable()
                If usePLINQ Then linqTableCol.AsParallel()
                dblQueryResults = From r In linqTableCol _
                               Order By r("ColID") _
                               Select OBJ = CDbl(r("OBJ"))
                array_dobj = dblQueryResults.ToArray()
            End If

            '---MCNT---dtMtx
            If array_mcnt Is Nothing Then
                'linqTable = dtMtx.AsEnumerable()
                If usePLINQ Then linqTableMtx.AsParallel()
                intGroups = From r In linqTableMtx _
                                    Order By r("ColID") Ascending _
                                    Group By r!ColID Into g = Group Select g.Count()
                array_mcnt = intGroups.ToArray()
            End If

            '---MBEG---dtMtx
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
                'linqTable = dtMtx.AsEnumerable()
                If usePLINQ Then linqTableMtx.AsParallel()
                intQueryResults = From r In linqTableMtx _
                               Order By r("ColID"), r("RowID") _
                               Select IDX = CInt(r("RowID")) - 1
                array_midx = intQueryResults.ToArray()
            End If

            '---MVAL---dtMtx
            If array_mval Is Nothing Then
                'linqTable = dtMtx.AsEnumerable()
                If usePLINQ Then linqTableMtx.AsParallel()
                dblQueryResults = From r In linqTableMtx _
                               Order By r("ColID"), r("RowID") _
                               Select COEF = CDbl(r("COEF"))
                array_mval = dblQueryResults.ToArray()
            End If

            '---LO---dtCol
            If array_dclo Is Nothing Then
                'linqTable = dtCol.AsEnumerable()
                If usePLINQ Then linqTableCol.AsParallel()
                dblQueryResults = From r In linqTableCol _
                               Order By r("ColID") _
                               Select LO = CDbl(r("LO"))
                array_dclo = dblQueryResults.ToArray()
            End If

            '---COL---dtCol
            If array_colNames Is Nothing Then
                'linqTable = dtCol.AsEnumerable()
                If usePLINQ Then linqTableCol.AsParallel()
                '/SubNameWithID
                If GenUtils.IsSwitchAvailable(_switches, "/SubNameWithID") Then
                    strQueryResults = From r In linqTableCol Order By r("ColID") Select CStr(r("ColID"))
                Else
                    strQueryResults = From r In linqTableCol Order By r("ColID") Select CStr(r("COL"))
                End If
                array_colNames = strQueryResults.ToArray()
            End If

            '---CTYP---dtCol
            If array_ctyp Is Nothing Then
                _engine.ProblemType = COPT.Engine.problemRunType.problemTypeContinuous
                proceed = False
                isMIP = False
                'linqTable = dtCol.AsEnumerable()
                If usePLINQ Then linqTableCol.AsParallel()
                intGroups = From r In linqTableCol _
                                    Where r!BINRY = True _
                                    Order By r("ColID") Ascending _
                                    Group By r!ColID Into g = Group Select g.Count()
                If intGroups.ToArray().Count > 0 Then
                    proceed = True
                    _engine.ProblemType = COPT.Engine.problemRunType.problemTypeBinary
                End If
                'Else
                intGroups = From r In linqTableCol _
                                    Where r!INTGR = True _
                                    Order By r("ColID") Ascending _
                                    Group By r!ColID Into g = Group Select g.Count()
                If intGroups.ToArray().Count > 0 Then
                    proceed = True
                    '_engine.ProblemType = "I"
                    If _engine.ProblemType = COPT.Engine.problemRunType.problemTypeContinuous Then
                        _engine.ProblemType = COPT.Engine.problemRunType.problemTypeInteger
                    Else
                        _engine.ProblemType = _engine.ProblemType + COPT.Engine.problemRunType.problemTypeInteger
                    End If
                End If
                'End If
                If proceed Then
                    isMIP = True
                    'linqTable = dtCol.AsEnumerable()
                    If usePLINQ Then linqTableCol.AsParallel()
                    'If usePLINQ Then linqTable.AsParallel()
                    charQueryResults = From r In linqTableCol _
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
                    For i = 0 To dtCol.Rows.Count - 1
                        If dtCol(i)("BINRY") Then
                            dtCol(i)("UP") = 1
                        End If
                    Next
                    'linqTable = dtCol.AsEnumerable()
                    If usePLINQ Then linqTableCol.AsParallel()
                    dblQueryResults = From r In linqTableCol _
                                   Order By r("ColID") _
                                   Select LO = CDbl(0.0)
                    array_initValues = dblQueryResults.ToArray()
                End If
            End If
            '---CTYP---

            'ub(0) = 40.0 : ub(1) = CPX_INFBOUND : ub(2) = CPX_INFBOUND
            '---UP---dtCol
            If array_dcup Is Nothing Then
                'linqTable = dtCol.AsEnumerable()
                If usePLINQ Then linqTableCol.AsParallel()
                dblQueryResults = From r In linqTableCol _
                               Order By r("ColID") _
                               Select UP = CDbl(r("UP"))
                array_dcup = dblQueryResults.ToArray()
            End If

            '---ROW---dtRow
            If array_rowNames Is Nothing Then
                'linqTable = dtRow.AsEnumerable()
                If usePLINQ Then linqTableRow.AsParallel()
                '/SubNameWithID
                If GenUtils.IsSwitchAvailable(_switches, "/SubNameWithID") Then
                    strQueryResults = From r In linqTableRow Order By r("RowID") Select CStr(r("RowID"))
                Else
                    strQueryResults = From r In linqTableRow Order By r("RowID") Select CStr(r("ROW"))
                End If
                array_rowNames = strQueryResults.ToArray()
            End If

            If disposeMatrix Then
                'dtMtx.Dispose()
                dtMtx = Nothing

                'GC.Collect()
                'GC.WaitForPendingFinalizers()
                'GC.Collect()
                'If GenUtils.IsSwitchAvailable(_switches, "/UseMinSysRes") Then
                '    GC.Collect()
                'End If
            End If

            '_progress = "Loaded tsysCol into memory."
            Debug.Print("Loaded solver arrays into memory.")
            Console.WriteLine("Done.")

            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "After loading solver arrays - APM = " & My.Computer.Info.AvailablePhysicalMemory.ToString())
            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "After loading solver arrays - AVM = " & My.Computer.Info.AvailableVirtualMemory.ToString())
            Console.WriteLine("APM = " & My.Computer.Info.AvailablePhysicalMemory.ToString())
            Console.WriteLine("AVM = " & My.Computer.Info.AvailableVirtualMemory.ToString())

            Return 0
        End Function

        Public Function LoadSolverArrays _
        ( _
            ByRef dtRow As DataTable, _
            ByRef dtCol As DataTable, _
            ByRef dtMtx As DataTable, _
            ByVal solverArraysLoadType As loadType _
        ) As Integer

            Dim cumTotal As Integer = 0
            Dim proceed As Boolean = False
            Dim disposeMatrix As Boolean = True
            Dim usePLINQ As Boolean = GenUtils.IsSwitchAvailable(_switches, "/UsePLINQ")
            Dim myArrayList As ArrayList = Nothing

            'Dim linqTable As System.Data.EnumerableRowCollection(Of System.Data.DataRow) = Nothing
            'Dim dblQueryResults As System.Data.EnumerableRowCollection(Of Double) = Nothing
            'Dim intQueryResults As System.Data.EnumerableRowCollection(Of Integer) = Nothing
            'Dim charQueryResults As System.Data.EnumerableRowCollection(Of Char) = Nothing
            'Dim strQueryResults As System.Data.EnumerableRowCollection(Of String) = Nothing
            'Dim intGroups As System.Collections.Generic.IEnumerable(Of Integer) = Nothing

            'Dim linqTable As System.Collections.Generic.IEnumerable(Of System.Data.DataRow) = Nothing
            Dim dblQueryResults As System.Collections.Generic.IEnumerable(Of Double) = Nothing
            Dim intQueryResults As System.Collections.Generic.IEnumerable(Of Integer) = Nothing
            Dim charQueryResults As System.Collections.Generic.IEnumerable(Of Char) = Nothing
            Dim strQueryResults As System.Collections.Generic.IEnumerable(Of String) = Nothing
            Dim intGroups As System.Collections.Generic.IEnumerable(Of Integer) = Nothing

            Debug.Print("Loading solver arrays into memory...")
            Console.Write("Loading solver arrays into memory...")

            '---Row - Begin---
            '---SENSE---
            charQueryResults = From r In dtRow.AsEnumerable() _
                           Order By r("RowID") _
                           Select SENSE = CChar(r("SENSE"))
            array_rtyp = charQueryResults.ToArray()

            '---RHS---
            dblQueryResults = From r In dtRow.AsEnumerable() _
                           Order By r("RowID") _
                           Select RHS = CDbl(r("RHS"))
            array_drhs = dblQueryResults.ToArray()

            '---ROW---
            '/SubNameWithID
            If GenUtils.IsSwitchAvailable(_switches, "/SubNameWithID") Then
                strQueryResults = From r In dtRow.AsEnumerable() Order By r("RowID") Select CStr(r("RowID"))
            Else
                strQueryResults = From r In dtRow.AsEnumerable() Order By r("RowID") Select CStr(r("ROW"))
            End If
            array_rowNames = strQueryResults.ToArray()
            '---Row - End---

            '---Col - Begin---
            '---COL---
            '/SubNameWithID
            If GenUtils.IsSwitchAvailable(_switches, "/SubNameWithID") Then
                strQueryResults = From r In dtCol.AsEnumerable() Order By r("ColID") Select CStr(r("ColID"))
            Else
                strQueryResults = From r In dtCol.AsEnumerable() Order By r("ColID") Select CStr(r("COL"))
            End If
            array_colNames = strQueryResults.ToArray()

            '---OBJ---
            dblQueryResults = From r In dtCol.AsEnumerable() _
                           Order By r("ColID") _
                           Select OBJ = CDbl(r("OBJ"))
            array_dobj = dblQueryResults.ToArray()

            '---LO---
            dblQueryResults = From r In dtCol.AsEnumerable() _
                           Order By r("ColID") _
                           Select LO = CDbl(r("LO"))
            array_dclo = dblQueryResults.ToArray()

            'ub(0) = 40.0 : ub(1) = CPX_INFBOUND : ub(2) = CPX_INFBOUND
            '---UP---
            dblQueryResults = From r In dtCol.AsEnumerable() _
                           Order By r("ColID") _
                           Select UP = CDbl(r("UP"))
            array_dcup = dblQueryResults.ToArray()

            '---CTYP---
            _engine.ProblemType = COPT.Engine.problemRunType.problemTypeContinuous
            proceed = False
            isMIP = False
            intGroups = From r In dtCol.AsEnumerable() _
                                Where r!BINRY = True _
                                Order By r("ColID") Ascending _
                                Group By r!ColID Into g = Group Select g.Count()
            If intGroups.ToArray().Count > 0 Then
                proceed = True
                _engine.ProblemType = COPT.Engine.problemRunType.problemTypeBinary
            End If
            intGroups = From r In dtCol.AsEnumerable() _
                                Where r!INTGR = True _
                                Order By r("ColID") Ascending _
                                Group By r!ColID Into g = Group Select g.Count()
            If intGroups.ToArray().Count > 0 Then
                proceed = True
                '_engine.ProblemType = "I"
                If _engine.ProblemType = COPT.Engine.problemRunType.problemTypeContinuous Then
                    _engine.ProblemType = COPT.Engine.problemRunType.problemTypeInteger
                Else
                    _engine.ProblemType = _engine.ProblemType + COPT.Engine.problemRunType.problemTypeInteger
                End If
            End If
            If proceed Then
                isMIP = True
                'If usePLINQ Then linqTable.AsParallel()
                charQueryResults = From r In dtCol.AsEnumerable() _
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
                For i = 0 To dtCol.Rows.Count - 1
                    If dtCol(i)("BINRY") Then
                        dtCol(i)("UP") = 1
                    End If
                Next
                dblQueryResults = From r In dtCol.AsEnumerable() _
                               Order By r("ColID") _
                               Select LO = CDbl(0.0)
                array_initValues = dblQueryResults.ToArray()
            End If
            '---CTYP---
            '---Col - End---

            '---Mtx - Begin---
            '---MCNT---
            intGroups = From r In dtMtx.AsEnumerable() _
                                Order By r("ColID") Ascending _
                                Group By r!ColID Into g = Group Select g.Count()
            array_mcnt = intGroups.ToArray()
            '---MBEG---
            myArrayList = New ArrayList
            myArrayList.Add(0)
            For ctr = 0 To array_mcnt.Length - 1
                cumTotal = cumTotal + array_mcnt(ctr)
                myArrayList.Add(cumTotal)
            Next
            array_mbeg = DirectCast(myArrayList.ToArray(GetType(Integer)), Integer())
            '---Mtx - End---

            '---MIDX---
            intQueryResults = From r In dtMtx.AsEnumerable() _
                           Order By r("ColID"), r("RowID") _
                           Select IDX = CInt(r("RowID")) - 1
            array_midx = intQueryResults.ToArray()

            '---MVAL---
            dblQueryResults = From r In dtMtx.AsEnumerable() _
                           Order By r("ColID"), r("RowID") _
                           Select COEF = CDbl(r("COEF"))
            array_mval = dblQueryResults.ToArray()

            If disposeMatrix Then
                'dtMtx.Dispose()
                dtMtx = Nothing
                'If GenUtils.IsSwitchAvailable(_switches, "/UseMinSysRes") Then
                '    GC.Collect()
                'End If
            End If

            '_progress = "Loaded tsysCol into memory."
            Debug.Print("Loaded solver arrays into memory.")
            Console.WriteLine("Done.")

            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "After loading solver arrays - APM = " & My.Computer.Info.AvailablePhysicalMemory.ToString())
            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "After loading solver arrays - AVM = " & My.Computer.Info.AvailableVirtualMemory.ToString())
            Console.WriteLine("APM = " & My.Computer.Info.AvailablePhysicalMemory.ToString())
            Console.WriteLine("AVM = " & My.Computer.Info.AvailableVirtualMemory.ToString())

            Return 0
        End Function

        Public Function LoadSolverArrays _
        ( _
            ByRef dtRow As DataTable, _
            ByRef dtCol As DataTable, _
            ByRef dtMtx As DataTable, _
            ByVal disposeMatrix As Boolean, _
            ByVal solverArraysLoadType As loadType _
        ) As Integer

            Dim cumTotal As Integer = 0
            Dim myArrayList As ArrayList = Nothing

            Dim linqTable As System.Data.EnumerableRowCollection(Of System.Data.DataRow) = Nothing
            Dim dblQueryResults As System.Data.EnumerableRowCollection(Of Double) = Nothing
            Dim intQueryResults As System.Data.EnumerableRowCollection(Of Integer) = Nothing
            Dim charQueryResults As System.Data.EnumerableRowCollection(Of Char) = Nothing
            Dim strQueryResults As System.Data.EnumerableRowCollection(Of String) = Nothing
            Dim intGroups As System.Collections.Generic.IEnumerable(Of Integer) = Nothing

            Dim plinq As ParallelEnumerable = Nothing

            Dim proceed As Boolean = False

            '_dtRow = dtRow
            '_dtCol = dtCol
            '_dtMtx = dtMtx

            Dim usePLINQ As Boolean = GenUtils.IsSwitchAvailable(_switches, "/UsePLINQ")

            Dim sql As String = ""
            Dim dt As DataTable = Nothing

            Debug.Print("Loading solver arrays into memory...")
            Console.Write("Loading solver arrays into memory...")

            'Load tsysRow - Start
            'If GenUtils.IsSwitchAvailable(_switches, "/SubNameWithID") Then
            '    sql = "SELECT [RowID], [SENSE], [RHS] FROM [" & _engine._srcRow & "] ORDER BY [RowID]"
            'Else
            '    sql = "SELECT [RowID], [ROW], [SENSE], [RHS] FROM [" & _engine._srcRow & "] ORDER BY [RowID]"
            'End If
            'dt = _engine.CurrentDb.GetDataTable(sql)
            '_engine.SolutionRows = dt.Rows.Count

            '---SENSE---
            linqTable = dtRow.AsEnumerable()
            If usePLINQ Then
                'linqTable.AsParallel()
                charQueryResults = From r In linqTable _
                               Order By r("RowID") _
                               Select SENSE = CChar(r("SENSE"))
            Else
                charQueryResults = From r In linqTable _
                               Order By r("RowID") _
                               Select SENSE = CChar(r("SENSE"))
            End If
            array_rtyp = charQueryResults.ToArray()

            '---RHS---
            linqTable = dtRow.AsEnumerable()
            If usePLINQ Then linqTable.AsParallel()
            dblQueryResults = From r In linqTable _
                           Order By r("RowID") _
                           Select RHS = CDbl(r("RHS"))
            array_drhs = dblQueryResults.ToArray()

            '---ROW---
            linqTable = dtRow.AsEnumerable()
            If usePLINQ Then linqTable.AsParallel()
            '/SubNameWithID
            If GenUtils.IsSwitchAvailable(_switches, "/SubNameWithID") Then
                strQueryResults = From r In linqTable Order By r("RowID") Select CStr(r("RowID"))
            Else
                strQueryResults = From r In linqTable Order By r("RowID") Select CStr(r("ROW"))
            End If
            array_rowNames = strQueryResults.ToArray()
            'Load tsysRow - End

            'Load tsysCol - Start
            dt = Nothing
            'If GenUtils.IsSwitchAvailable(_switches, "/SubNameWithID") Then
            '    sql = "SELECT [ColID], [OBJ], [LO], [UP], [INTGR], [BINRY] FROM [" & _engine._srcCol & "] ORDER BY [ColID]"
            'Else
            '    sql = "SELECT [ColID], [COL], [OBJ], [LO], [UP], [INTGR], [BINRY] FROM [" & _engine._srcCol & "] ORDER BY [ColID]"
            'End If
            'dt = _engine.CurrentDb.GetDataTable(sql)
            '_engine.SolutionColumns = dt.Rows.Count

            '---OBJ---
            linqTable = dtCol.AsEnumerable()
            If usePLINQ Then linqTable.AsParallel()
            dblQueryResults = From r In linqTable _
                           Order By r("ColID") _
                           Select OBJ = CDbl(r("OBJ"))
            array_dobj = dblQueryResults.ToArray()

            '---LO---
            linqTable = dtCol.AsEnumerable()
            If usePLINQ Then linqTable.AsParallel()
            dblQueryResults = From r In linqTable _
                           Order By r("ColID") _
                           Select LO = CDbl(r("LO"))
            array_dclo = dblQueryResults.ToArray()

            'ub(0) = 40.0 : ub(1) = CPX_INFBOUND : ub(2) = CPX_INFBOUND
            '---UP---
            linqTable = dtCol.AsEnumerable()
            If usePLINQ Then linqTable.AsParallel()
            dblQueryResults = From r In linqTable _
                           Order By r("ColID") _
                           Select UP = CDbl(r("UP"))
            array_dcup = dblQueryResults.ToArray()

            '---COL---
            linqTable = dtCol.AsEnumerable()
            If usePLINQ Then linqTable.AsParallel()
            '/SubNameWithID
            If GenUtils.IsSwitchAvailable(_switches, "/SubNameWithID") Then
                strQueryResults = From r In linqTable Order By r("ColID") Select CStr(r("ColID"))
            Else
                strQueryResults = From r In linqTable Order By r("ColID") Select CStr(r("COL"))
            End If
            array_colNames = strQueryResults.ToArray()

            '---CTYP---
            _engine.ProblemType = COPT.Engine.problemRunType.problemTypeContinuous
            proceed = False
            isMIP = False
            linqTable = dtCol.AsEnumerable()
            If usePLINQ Then linqTable.AsParallel()
            intGroups = From r In linqTable _
                                Where r!BINRY = True _
                                Order By r("ColID") Ascending _
                                Group By r!ColID Into g = Group Select g.Count()
            If intGroups.ToArray().Count > 0 Then
                proceed = True
                _engine.ProblemType = COPT.Engine.problemRunType.problemTypeBinary
            End If
            'Else
            intGroups = From r In linqTable _
                                Where r!INTGR = True _
                                Order By r("ColID") Ascending _
                                Group By r!ColID Into g = Group Select g.Count()
            If intGroups.ToArray().Count > 0 Then
                proceed = True
                '_engine.ProblemType = "I"
                If _engine.ProblemType = COPT.Engine.problemRunType.problemTypeContinuous Then
                    _engine.ProblemType = COPT.Engine.problemRunType.problemTypeInteger
                Else
                    _engine.ProblemType = _engine.ProblemType + COPT.Engine.problemRunType.problemTypeInteger
                End If
            End If
            'End If
            If proceed Then
                isMIP = True
                linqTable = dtCol.AsEnumerable()
                If usePLINQ Then linqTable.AsParallel()
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
                For i = 0 To dtCol.Rows.Count - 1
                    If dtCol(i)("BINRY") Then
                        dtCol(i)("UP") = 1
                    End If
                Next
                linqTable = dtCol.AsEnumerable()
                If usePLINQ Then linqTable.AsParallel()
                dblQueryResults = From r In linqTable _
                               Order By r("ColID") _
                               Select LO = CDbl(0.0)
                array_initValues = dblQueryResults.ToArray()
            End If
            '---CTYP---
            'Load tsysCol - End

            'Load tsysMtx - Start
            dt = Nothing
            'If GenUtils.IsSwitchAvailable(_switches, "/SubNameWithID") Then
            sql = "SELECT [ColID], [RowID], [COEF] FROM [" & _engine._srcMtx & "] ORDER BY [ColID], [RowID]"
            'Else
            'sql = "SELECT [ColID], [RowID], [COL], [ROW], [COEF] FROM [" & _engine._srcMtx & "] ORDER BY [ColID], [RowID]"
            'End If
            dt = _engine.CurrentDb.GetDataTable(sql)
            _engine.SolutionNonZeros = dt.Rows.Count

            '---MCNT---
            linqTable = dt.AsEnumerable()
            If usePLINQ Then linqTable.AsParallel()
            intGroups = From r In linqTable _
                                Order By r("ColID") Ascending _
                                Group By r!ColID Into g = Group Select g.Count()
            array_mcnt = intGroups.ToArray()

            '---MBEG---
            myArrayList = New ArrayList
            myArrayList.Add(0)
            For ctr = 0 To array_mcnt.Length - 1
                cumTotal = cumTotal + array_mcnt(ctr)
                myArrayList.Add(cumTotal)
            Next
            array_mbeg = DirectCast(myArrayList.ToArray(GetType(Integer)), Integer())

            '---MIDX---
            linqTable = dt.AsEnumerable()
            If usePLINQ Then linqTable.AsParallel()
            intQueryResults = From r In linqTable _
                           Order By r("ColID"), r("RowID") _
                           Select IDX = CInt(r("RowID")) - 1
            array_midx = intQueryResults.ToArray()

            '---MVAL---
            linqTable = dt.AsEnumerable()
            If usePLINQ Then linqTable.AsParallel()
            dblQueryResults = From r In linqTable _
                           Order By r("ColID"), r("RowID") _
                           Select COEF = CDbl(r("COEF"))
            array_mval = dblQueryResults.ToArray()
            'Load tsysMtx - End

            'If disposeMatrix Then
            '    'dtMtx.Dispose()
            dtMtx = Nothing
            dt = Nothing
            '    'If GenUtils.IsSwitchAvailable(_switches, "/UseMinSysRes") Then
            '    '    GC.Collect()
            '    'End If
            'End If

            '_progress = "Loaded tsysCol into memory."
            Debug.Print("Loaded solver arrays into memory.")
            Console.WriteLine("Done.")

            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "After loading solver arrays - APM = " & My.Computer.Info.AvailablePhysicalMemory.ToString())
            EntLib.COPT.Log.Log(GenUtils.GetWorkDir(_switches), "C-OPT Engine", "After loading solver arrays - AVM = " & My.Computer.Info.AvailableVirtualMemory.ToString())
            Console.WriteLine("APM = " & My.Computer.Info.AvailablePhysicalMemory.ToString())
            Console.WriteLine("AVM = " & My.Computer.Info.AvailableVirtualMemory.ToString())

            Return 0
        End Function

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="dtRow"></param>
        ''' <param name="dtCol"></param>
        ''' <param name="dtMtx"></param>
        ''' <param name="disposeMatrix"></param>
        ''' <param name="usePLINQ"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function LoadSolverArrays _
        ( _
            ByRef dtRow As DataTable, _
            ByRef dtCol As DataTable, _
            ByRef dtMtx As DataTable, _
            ByVal disposeMatrix As Boolean, _
            ByVal usePLINQ As Boolean _
        ) As Integer

            Dim cumTotal As Integer = 0
            Dim myArrayList As ArrayList = Nothing

            'Dim linqTable As System.Data.EnumerableRowCollection(Of System.Data.DataRow) = Nothing
            'Dim dblQueryResults As System.Data.EnumerableRowCollection(Of Double) = Nothing
            'Dim intQueryResults As System.Data.EnumerableRowCollection(Of Integer) = Nothing
            'Dim charQueryResults As System.Data.EnumerableRowCollection(Of Char) = Nothing
            'Dim strQueryResults As System.Data.EnumerableRowCollection(Of String) = Nothing
            'Dim intGroups As System.Collections.Generic.IEnumerable(Of Integer) = Nothing

            Dim linqTable As System.Collections.Generic.IEnumerable(Of System.Data.DataRow) = Nothing
            Dim dblQueryResults As System.Collections.Generic.IEnumerable(Of Double) = Nothing
            Dim intQueryResults As System.Collections.Generic.IEnumerable(Of Integer) = Nothing
            Dim charQueryResults As System.Collections.Generic.IEnumerable(Of Char) = Nothing
            Dim strQueryResults As System.Collections.Generic.IEnumerable(Of String) = Nothing
            Dim intGroups As System.Collections.Generic.IEnumerable(Of Integer) = Nothing
            Dim array_test() As Char = Nothing
            Dim proceed As Boolean = False

            '_dtRow = dtRow
            '_dtCol = dtCol
            '_dtMtx = dtMtx

            'Dim usePLINQ As Boolean = GenUtils.IsSwitchAvailable(_switches, "/UsePLINQ")

            '---TEST---
            If usePLINQ Then
                charQueryResults = From r In linqTable.AsParallel().AsOrdered() _
                               Order By r.Field(Of Integer)("RowID") Ascending _
                               Select r.Field(Of Char)("SENSE")
            Else
                charQueryResults = From r In linqTable.AsEnumerable() _
                               Order By r.Field(Of Integer)("RowID") Ascending _
                               Select r.Field(Of Char)("SENSE")
            End If
            array_test = charQueryResults.ToArray()
            '---TEST---

            '---SENSE---
            If usePLINQ Then
                charQueryResults = From r In dtRow.AsEnumerable().AsParallel().AsOrdered() _
                               Order By r.Field(Of Integer)("RowID") Ascending _
                               Select r.Field(Of Char)("SENSE")
            Else
                charQueryResults = From r In dtRow.AsEnumerable() _
                               Order By r.Field(Of Integer)("RowID") Ascending _
                               Select r.Field(Of Char)("SENSE")
            End If
            array_rtyp = charQueryResults.ToArray()
            '---SENSE---

            '---RHS---
            If usePLINQ Then
                charQueryResults = From r In linqTable.AsParallel().AsOrdered() _
                               Order By r.Field(Of Integer)("RowID") Ascending _
                               Select r.Field(Of Char)("SENSE")
            Else
                charQueryResults = From r In linqTable.AsEnumerable() _
                               Order By r.Field(Of Integer)("RowID") Ascending _
                               Select r.Field(Of Char)("SENSE")
            End If
            If usePLINQ Then linqTable.AsParallel()
            dblQueryResults = From r In linqTable _
                           Order By r("RowID") _
                           Select RHS = CDbl(r("RHS"))
            array_drhs = dblQueryResults.ToArray()

            '---OBJ---
            linqTable = dtCol.AsEnumerable()
            If usePLINQ Then linqTable.AsParallel()
            dblQueryResults = From r In linqTable _
                           Order By r("ColID") _
                           Select OBJ = CDbl(r("OBJ"))
            array_dobj = dblQueryResults.ToArray()

            '---MCNT---
            linqTable = dtMtx.AsEnumerable()
            If usePLINQ Then linqTable.AsParallel()
            intGroups = From r In linqTable _
                                Order By r("ColID") Ascending _
                                Group By r!ColID Into g = Group Select g.Count()
            array_mcnt = intGroups.ToArray()

            '---MBEG---
            myArrayList = New ArrayList
            myArrayList.Add(0)
            For ctr = 0 To array_mcnt.Length - 1
                cumTotal = cumTotal + array_mcnt(ctr)
                myArrayList.Add(cumTotal)
            Next
            array_mbeg = DirectCast(myArrayList.ToArray(GetType(Integer)), Integer())

            '---MIDX---
            linqTable = dtMtx.AsEnumerable()
            If usePLINQ Then linqTable.AsParallel()
            intQueryResults = From r In linqTable _
                           Order By r("ColID"), r("RowID") _
                           Select IDX = CInt(r("RowID")) - 1
            array_midx = intQueryResults.ToArray()

            '---MVAL---
            linqTable = dtMtx.AsEnumerable()
            If usePLINQ Then linqTable.AsParallel()
            dblQueryResults = From r In linqTable _
                           Order By r("ColID"), r("RowID") _
                           Select COEF = CDbl(r("COEF"))
            array_mval = dblQueryResults.ToArray()

            '---LO---
            linqTable = dtCol.AsEnumerable()
            If usePLINQ Then linqTable.AsParallel()
            dblQueryResults = From r In linqTable _
                           Order By r("ColID") _
                           Select LO = CDbl(r("LO"))
            array_dclo = dblQueryResults.ToArray()

            '---COL---
            linqTable = dtCol.AsEnumerable()
            If usePLINQ Then linqTable.AsParallel()
            '/SubNameWithID
            If GenUtils.IsSwitchAvailable(_switches, "/SubNameWithID") Then
                strQueryResults = From r In linqTable Order By r("ColID") Select CStr(r("ColID"))
            Else
                strQueryResults = From r In linqTable Order By r("ColID") Select CStr(r("COL"))
            End If
            array_colNames = strQueryResults.ToArray()

            '---CTYP---
            _engine.ProblemType = COPT.Engine.problemRunType.problemTypeContinuous
            proceed = False
            isMIP = False
            linqTable = dtCol.AsEnumerable()
            If usePLINQ Then linqTable.AsParallel()
            intGroups = From r In linqTable _
                                Where r!BINRY = True _
                                Order By r("ColID") Ascending _
                                Group By r!ColID Into g = Group Select g.Count()
            If intGroups.ToArray().Count > 0 Then
                proceed = True
                _engine.ProblemType = COPT.Engine.problemRunType.problemTypeBinary
            End If
            'Else
            intGroups = From r In linqTable _
                                Where r!INTGR = True _
                                Order By r("ColID") Ascending _
                                Group By r!ColID Into g = Group Select g.Count()
            If intGroups.ToArray().Count > 0 Then
                proceed = True
                '_engine.ProblemType = "I"
                If _engine.ProblemType = COPT.Engine.problemRunType.problemTypeContinuous Then
                    _engine.ProblemType = COPT.Engine.problemRunType.problemTypeInteger
                Else
                    _engine.ProblemType = _engine.ProblemType + COPT.Engine.problemRunType.problemTypeInteger
                End If
            End If
            'End If
            If proceed Then
                isMIP = True
                linqTable = dtCol.AsEnumerable()
                If usePLINQ Then linqTable.AsParallel()
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
                For i = 0 To dtCol.Rows.Count - 1
                    If dtCol(i)("BINRY") Then
                        dtCol(i)("UP") = 1
                    End If
                Next
                linqTable = dtCol.AsEnumerable()
                If usePLINQ Then linqTable.AsParallel()
                dblQueryResults = From r In linqTable _
                               Order By r("ColID") _
                               Select LO = CDbl(0.0)
                array_initValues = dblQueryResults.ToArray()
            End If
            '---CTYP---

            'ub(0) = 40.0 : ub(1) = CPX_INFBOUND : ub(2) = CPX_INFBOUND
            '---UP---
            linqTable = dtCol.AsEnumerable()
            If usePLINQ Then linqTable.AsParallel()
            dblQueryResults = From r In linqTable _
                           Order By r("ColID") _
                           Select UP = CDbl(r("UP"))
            array_dcup = dblQueryResults.ToArray()

            '---ROW---
            linqTable = dtRow.AsEnumerable()
            If usePLINQ Then linqTable.AsParallel()
            '/SubNameWithID
            If GenUtils.IsSwitchAvailable(_switches, "/SubNameWithID") Then
                strQueryResults = From r In linqTable Order By r("RowID") Select CStr(r("RowID"))
            Else
                strQueryResults = From r In linqTable Order By r("RowID") Select CStr(r("ROW"))
            End If
            array_rowNames = strQueryResults.ToArray()

            Return 0
        End Function

        Public Sub SetSession(ByRef currentSession As EntLib.COPT.Session)
            _currentSession = currentSession
        End Sub

        Public Function GetSession() As EntLib.COPT.Session
            Return _currentSession
        End Function

        Public Property isMIP() As Boolean
            Get
                Return _isMIP
            End Get
            Set(ByVal value As Boolean)
                _isMIP = value
            End Set
        End Property

    End Class
End Namespace