Namespace COPT
    <ComClass(Solver_GLPK.ClassId, Solver_GLPK.InterfaceId, Solver_GLPK.EventsId)> _
    Public Class Solver_GLPK
        Inherits Solver
#Region "COM GUIDs"
        ' These  GUIDs provide the COM identity for this class 
        ' and its COM interfaces. If you change them, existing 
        ' clients will no longer be able to access the class.
        Public Const ClassId As String = "00ded50b-e005-4d0e-8212-0799c3033789"
        Public Const InterfaceId As String = "35304705-44a4-4b54-b1cd-9ad334b8ce23"
        Public Const EventsId As String = "294cdcb5-aae6-4d66-834c-099b490f1fed"
#End Region

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

        Public Overrides Function Solve(ByRef dtRow As DataTable, ByRef dtCol As DataTable, ByRef dtMtx As DataTable) As Boolean

        End Function

        Public Overrides Function Solve(ByRef ds As DataSet) As Boolean

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
            Return "GLPK"
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

    End Class
End Namespace