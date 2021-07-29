Imports System
Imports System.Collections.Generic
Imports System.Text
Imports System.Data.SqlClient
Imports System.Reflection
Imports System.IO
Imports System.Configuration
Imports System.Data.OleDb

<Microsoft.VisualBasic.ComClass()> _
Public Class DAL

    Public Sub DAL()
        '
    End Sub

    Public Sub New()
        '
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Public Function GetData() As String
        Return "GetData"
    End Function

    Public Function GetDataSet(ByVal name As String) As DataSet
        Dim db As OleDb.OleDbConnection = CurrentAccessDb()

        Return Nothing
    End Function

    Public Function GetDataSingleValue() As String
        Return Nothing
    End Function

    Public Function CurrentAccessDb() As OleDb.OleDbConnection
        'Dim connectionString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\myPath\myJet.mdb;User ID=Admin;Password="
        Dim connectionString As String = ""
        Dim db As OleDb.OleDbConnection = New OleDb.OleDbConnection(connectionString)
        db.Open()

        Return db
    End Function
End Class
