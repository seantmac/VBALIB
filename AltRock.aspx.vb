Imports System.Data.OleDb

Public Class AltRock
    Inherits System.Web.UI.Page

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents dgrAltRock As System.Web.UI.WebControls.DataGrid

    'NOTE: The following placeholder declaration is required by the Web Form Designer.
    'Do not delete or move it.
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: This method call is required by the Web Form Designer
        'Do not modify it using the code editor.
        InitializeComponent()
    End Sub

#End Region

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' You will need to edit the Data Source value to correspond
        ' to the location of the 17-03.mdb database on your system.
        Dim cnx As OleDbConnection = _
         New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;" & _
         "Data Source=D:\Books\AccessCookBook2003\ch17\17-03.mdb")

        cnx.Open()

        ' Constuct a OleDbCommand to execute the query
        Dim cmdAltRock As OleDbCommand = _
         New OleDbCommand("qryAlternativeAlbums", cnx)

        ' Odd as it may seem, you need to set the CommandType
        ' to CommandType.StoredProcedure.
        cmdAltRock.CommandType = CommandType.StoredProcedure

        ' Run the query and place the rows in an OleDbDataReader.
        Dim drAltRock As OleDbDataReader
        drAltRock = cmdAltRock.ExecuteReader

        ' Bind the OleDbDataReader to the DataGrid
        dgrAltRock.DataSource = drAltRock
        dgrAltRock.DataBind()
    End Sub

End Class
