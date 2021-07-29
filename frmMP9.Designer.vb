<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMP9
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponentx()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMP9))
        Me.MappointControl1 = New AxMapPoint.AxMappointControl
        CType(Me.MappointControl1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'MappointControl1
        '
        Me.MappointControl1.Enabled = True
        Me.MappointControl1.Location = New System.Drawing.Point(12, 12)
        Me.MappointControl1.Name = "MappointControl1"
        Me.MappointControl1.OcxState = CType(resources.GetObject("MappointControl1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.MappointControl1.Size = New System.Drawing.Size(762, 561)
        Me.MappointControl1.TabIndex = 0
        '
        'frmMP9
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(786, 585)
        Me.Controls.Add(Me.MappointControl1)
        Me.Name = "frmMP9"
        Me.Text = "frmMP9"
        CType(Me.MappointControl1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents MappointControl1 As AxMapPoint.AxMappointControl
    Friend WithEvents txtMPVer As System.Windows.Forms.TextBox
    Friend WithEvents cmdSave As System.Windows.Forms.Button
    Friend WithEvents cmdClear As System.Windows.Forms.Button
    Friend WithEvents cmdMapShipments As System.Windows.Forms.Button
    Friend WithEvents cmdShadedDemand As System.Windows.Forms.Button
    Friend WithEvents cmdCircleDemand As System.Windows.Forms.Button
    Friend WithEvents cmdSizedPie As System.Windows.Forms.Button
    Friend WithEvents cmdMapSource As System.Windows.Forms.Button
    Friend WithEvents cmdSourceProdLinePie As System.Windows.Forms.Button
    Friend WithEvents cmdMakeStateMaps As System.Windows.Forms.Button
    Friend WithEvents lstMileage As System.Windows.Forms.ListBox
    Friend WithEvents cmdCalcMiles As System.Windows.Forms.Button
    Friend WithEvents cmdDemandSourcingPies As System.Windows.Forms.Button

End Class
