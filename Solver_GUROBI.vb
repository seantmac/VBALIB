Imports System
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Linq
Imports System.Linq.Expressions
Imports System.Collections
Imports System.Collections.Generic
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
    Public Class Solver_GUROBI
        Inherits Solver

        Private _solutionStatus As String
        Private _solutionRows As Integer
        Private _solutionColumns As Integer
        Private _solutionObj As Double
        Private Shadows _switches() As String

        Public Sub New()
            MyBase.New(New COPT.Engine)
        End Sub

        Public Sub New(ByVal engine As COPT.Engine)
            MyBase.New(engine)
        End Sub

        Public Sub New(ByVal engine As COPT.Engine, ByVal switches() As String)
            MyBase.New(engine, switches)
            _switches = switches
        End Sub

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

        Public Overrides Function Solve(ByRef dtRow As DataTable, ByRef dtCol As DataTable, ByRef dtMtx As DataTable) As Boolean
            Dim startTime As Integer = My.Computer.Clock.TickCount



            EntLib.COPT.Log.Log(_workDir, "C-OPT Engine - CoinMP - Total solve took: ", GenUtils.FormatTime(startTime, My.Computer.Clock.TickCount))
        End Function

        Public Overrides Function Solve(ByRef ds As DataSet) As Boolean

        End Function

        Public Overrides Function getSolutionStatusDescription() As String
            Return _solutionStatus
        End Function

        Public Overrides Function getSolverName() As String
            Return "Remote"
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

        Private Function UploadToServer(ByRef dtRow As DataTable, ByRef dtCol As DataTable, ByRef dtMtx As DataTable) As Integer

            Return 0
        End Function
    End Class
End Namespace
