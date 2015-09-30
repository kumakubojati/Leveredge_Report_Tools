<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmFCS_RepViewer
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
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim ReportDataSource1 As Microsoft.Reporting.WinForms.ReportDataSource = New Microsoft.Reporting.WinForms.ReportDataSource()
        Dim ReportDataSource2 As Microsoft.Reporting.WinForms.ReportDataSource = New Microsoft.Reporting.WinForms.ReportDataSource()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmFCS_RepViewer))
        Me.dsFCSRepBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.dsHeadRepBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.rvFCS_New = New Microsoft.Reporting.WinForms.ReportViewer()
        CType(Me.dsFCSRepBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dsHeadRepBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dsFCSRepBindingSource
        '
        Me.dsFCSRepBindingSource.DataMember = "dtFCS"
        Me.dsFCSRepBindingSource.DataSource = GetType(Leveredge_Report_Tools.dsFCSRep)
        '
        'dsHeadRepBindingSource
        '
        Me.dsHeadRepBindingSource.DataMember = "dtHeadRep"
        Me.dsHeadRepBindingSource.DataSource = GetType(Leveredge_Report_Tools.dsFCSRep)
        '
        'rvFCS_New
        '
        ReportDataSource1.Name = "dsFCS"
        ReportDataSource1.Value = Me.dsFCSRepBindingSource
        ReportDataSource2.Name = "dsHeadRep"
        ReportDataSource2.Value = Me.dsHeadRepBindingSource
        Me.rvFCS_New.LocalReport.DataSources.Add(ReportDataSource1)
        Me.rvFCS_New.LocalReport.DataSources.Add(ReportDataSource2)
        Me.rvFCS_New.LocalReport.ReportEmbeddedResource = "Leveredge_Report_Tools.FCSRep_New.rdlc"
        Me.rvFCS_New.Location = New System.Drawing.Point(1, -1)
        Me.rvFCS_New.Name = "rvFCS_New"
        Me.rvFCS_New.Size = New System.Drawing.Size(1254, 557)
        Me.rvFCS_New.TabIndex = 0
        '
        'frmFCS_RepViewer
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(1256, 557)
        Me.Controls.Add(Me.rvFCS_New)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmFCS_RepViewer"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "FCS Report"
        CType(Me.dsFCSRepBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dsHeadRepBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents rvFCS_New As Microsoft.Reporting.WinForms.ReportViewer
    Friend WithEvents dsFCSRepBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents dsHeadRepBindingSource As System.Windows.Forms.BindingSource
End Class
