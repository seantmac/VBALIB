Imports System.Text
Imports CoinMPTestVB8.Coin.CoinMP

Namespace CoinMPTest

    Module ProblemCOPTDynamic
        Private _ds As New DataSet

        Public Sub Solve(ByVal solveProblem As SolveProblem)
            Const NUM_COLS As Integer = 32
            Const NUM_ROWS As Integer = 27
            Const NUM_NZ As Integer = 83
            Const NUM_RNG As Integer = 0
            Const INF As Double = 1.0E+37

            Dim probname As String = "Afiro"
            Dim ncol As Integer = NUM_COLS
            Dim nrow As Integer = NUM_ROWS
            Dim nels As Integer = NUM_NZ
            Dim nrng As Integer = NUM_RNG

            Dim objectname As String = "Cost"
            Dim objsens As Integer = CoinMP.ObjectSense.Min
            Dim objconst As Double = 0.0

            'OBJ Function
            'Coefficient of each variable in objective function
            '                       x1  x2
            Dim dobj() As Double = {0, -0.4, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, -0.32, 0, 0, 0, -0.6, _
                                    0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, -0.48, 0, 0, 10} '-> 32

            'Lower limit of each variable in objective function
            Dim dclo() As Double = {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                                    0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0} '-> 32

            'Upper limit of each variable in objective function
            'Note:
            'There are 42 entries in this array, which may be incorrect. 32 should be the correct number.
            Dim dcup() As Double = {INF, INF, INF, INF, INF, INF, INF, INF, INF, INF, INF, INF, _
                    INF, INF, INF, INF, INF, INF, INF, INF, INF, INF, INF, INF, INF, INF, INF, _
                    INF, INF, INF, INF, INF, INF, INF, INF, INF, INF, INF, INF, INF, INF, INF} '-> 42

            'Equality of each constraint row
            Dim rtyp() As Char = "EELLEELLLLEELLEELLLLLLLLLLL" '-> 27

            'RHS value of each constraint row
            Dim drhs() As Double = {0, 0, 80, 0, 0, 0, 80, 0, 0, 0, 0, 0, 500, 0, 0, 44, 500, 0, _
                                    0, 0, 0, 0, 0, 0, 0, 310, 300} '-> 27

            'Cumulative total of number of occurances of each variable
            'Note:
            'This is a 0-based array
            Dim mbeg() As Integer = {0, 4, 6, 8, 10, 14, 18, 22, 26, 28, 30, 32, 34, 36, 38, 40, _
                             44, 46, 48, 50, 52, 56, 60, 64, 68, 70, 72, 74, 76, 78, 80, 82, 83} '-> 33

            'Number of coefficients for each variable
            Dim mcnt() As Integer = {4, 2, 2, 2, 4, 4, 4, 4, 2, 2, 2, 2, 2, 2, 2, 4, 2, 2, 2, 2, 4, _
                                     4, 4, 4, 2, 2, 2, 2, 2, 2, 2, 1} '-> 32

            'Ordinal position of each variable in the rows/constraints
            'Note:
            'This is a 0-based array
            'The first four (0,1,2,23) are positions of the variable x1 in rows (c1,c2,c3,c24).
            Dim midx() As Integer = {0, 1, 2, 23, 0, 3, 0, 21, 1, 25, 4, 5, 6, 24, 4, 5, 7, 24, 4, 5, _
                        8, 24, 4, 5, 9, 24, 6, 20, 7, 20, 8, 20, 9, 20, 3, 4, 4, 22, 5, 26, 10, 11, _
                        12, 21, 10, 13, 10, 23, 10, 20, 11, 25, 14, 15, 16, 22, 14, 15, 17, 22, 14, _
                        15, 18, 22, 14, 15, 19, 22, 16, 20, 17, 20, 18, 20, 19, 20, 13, 15, 15, 24, _
                        14, 26, 15} '-> 83

            'Coefficient value of each variable as defined by colNames
            Dim mval() As Double = {-1, -1.06, 1, 0.301, 1, -1, 1, -1, 1, 1, -1, -1.06, 1, 0.301, _
                    -1, -1.06, 1, 0.313, -1, -0.96, 1, 0.313, -1, -0.86, 1, 0.326, -1, 2.364, -1, _
                    2.386, -1, 2.408, -1, 2.429, 1.4, 1, 1, -1, 1, 1, -1, -0.43, 1, 0.109, 1, -1, _
                    1, -1, 1, -1, 1, 1, -0.43, 1, 1, 0.109, -0.43, 1, 1, 0.108, -0.39, 1, 1, _
                    0.108, -0.37, 1, 1, 0.107, -1, 2.191, -1, 2.219, -1, 2.249, -1, 2.279, 1.4, _
                    -1, 1, -1, 1, 1, 1} '-> 83

            'Column names
            Dim colNames() As String = {"x01", "x02", "x03", "x04", "x06", "x07", "x08", "x09", _
                    "x10", "x11", "x12", "x13", "x14", "x15", "x16", "x22", "x23", "x24", "x25", _
                    "x26", "x28", "x29", "x30", "x31", "x32", "x33", "x34", "x35", "x36", "x37", _
                    "x38", "x39"} '-> 32

            'Row names
            Dim rowNames() As String = {"r09", "r10", "x05", "x21", "r12", "r13", "x17", "x18", _
                    "x19", "x20", "r19", "r20", "x27", "x44", "r22", "r23", "x40", "x41", "x42", _
                    "x43", "x45", "x46", "x47", "x48", "x49", "x50", "x51"} '-> 27

            Dim optimalValue As Double = -464.753142857

            solveProblem.Run(probname, optimalValue, ncol, nrow, nels, nrng, objsens, objconst, _
                dobj, dclo, dcup, rtyp, drhs, Nothing, mbeg, mcnt, midx, mval, _
                colNames, rowNames, objectname, Nothing, Nothing)
        End Sub

        Public Sub db_Connect()
            Dim con As System.Data.OleDb.OleDbConnection
            Dim adapter As System.Data.OleDb.OleDbDataAdapter
            Dim sql As String
            Dim conStr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\OPTMODELS\Afiro\Afiro.MDB;;Jet OLEDB:System Database=C:\OPTMODELS\C-OPTSYS\System.MDW.COPT;User ID=Admin;Password="

            Dim dt As New DataTable()

            con = New OleDb.OleDbConnection(conStr)
            Try
                con.Open()
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation)
            End Try

            sql = "SELECT * FROM tsysCol"
            adapter = New OleDb.OleDbDataAdapter(sql, conStr)
            adapter.Fill(dt)
            _ds.Tables.Add(dt)
            dt = Nothing

            sql = "SELECT * FROM tsysRow"
            adapter = New OleDb.OleDbDataAdapter(sql, conStr)
            dt = New DataTable
            adapter.Fill(dt)
            _ds.Tables.Add(dt)
            dt = Nothing

            sql = "SELECT * FROM tsysMtx"
            adapter = New OleDb.OleDbDataAdapter(sql, conStr)
            dt = New DataTable
            adapter.Fill(dt)
            _ds.Tables.Add(dt)
            dt = Nothing

            _ds.Tables(0).TableName = "tsysCol"
            _ds.Tables(1).TableName = "tsysRow"
            _ds.Tables(2).TableName = "tsysMtx"

            'MsgBox("tsysCol:" & _ds.Tables(0).Rows.Count)
            'MsgBox("tsysRow:" & _ds.Tables(1).Rows.Count)
            'MsgBox("tsysMtx:" & _ds.Tables(2).Rows.Count)

            con.Close()
            con = Nothing
            adapter = Nothing
            '_ds = Nothing
            dt = Nothing

            'MsgBox("tsysCol:" & _ds.Tables(0).Rows.Count)
            'MsgBox("tsysRow:" & _ds.Tables(1).Rows.Count)
            'MsgBox("tsysMtx:" & _ds.Tables(2).Rows.Count)

        End Sub

        Public Sub GetRowCount()
            MsgBox("tsysCol:" & _ds.Tables(0).Rows.Count)
            MsgBox("tsysRow:" & _ds.Tables(1).Rows.Count)
            MsgBox("tsysMtx:" & _ds.Tables(2).Rows.Count)
        End Sub
    End Module

End Namespace
