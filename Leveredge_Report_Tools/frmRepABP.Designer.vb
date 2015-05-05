<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmRepABP
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
        Dim ReportDataSource4 As Microsoft.Reporting.WinForms.ReportDataSource = New Microsoft.Reporting.WinForms.ReportDataSource()
        Dim ReportDataSource5 As Microsoft.Reporting.WinForms.ReportDataSource = New Microsoft.Reporting.WinForms.ReportDataSource()
        Dim ReportDataSource6 As Microsoft.Reporting.WinForms.ReportDataSource = New Microsoft.Reporting.WinForms.ReportDataSource()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmRepABP))
        Me.rvABP = New Microsoft.Reporting.WinForms.ReportViewer()
        Me.dsRep = New Leveredge_Report_Tools.dsRep()
        Me.dtRepBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.dtHeadRepBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.dtHeadRep2BindingSource = New System.Windows.Forms.BindingSource(Me.components)
        CType(Me.dsRep, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dtRepBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dtHeadRepBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dtHeadRep2BindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'rvABP
        '
        Me.rvABP.AutoSize = True
        Me.rvABP.Dock = System.Windows.Forms.DockStyle.Fill
        ReportDataSource4.Name = "dsAchByProd"
        ReportDataSource4.Value = Me.dtRepBindingSource
        ReportDataSource5.Name = "dsHeadRep"
        ReportDataSource5.Value = Me.dtHeadRepBindingSource
        ReportDataSource6.Name = "dsHeadRep2"
        ReportDataSource6.Value = Me.dtHeadRep2BindingSource
        Me.rvABP.LocalReport.DataSources.Add(ReportDataSource4)
        Me.rvABP.LocalReport.DataSources.Add(ReportDataSource5)
        Me.rvABP.LocalReport.DataSources.Add(ReportDataSource6)
        Me.rvABP.LocalReport.ReportEmbeddedResource = "Leveredge_Report_Tools.AchievementByProduct.rdlc"
        Me.rvABP.Location = New System.Drawing.Point(0, 0)
        Me.rvABP.Name = "rvABP"
        Me.rvABP.Size = New System.Drawing.Size(766, 430)
        Me.rvABP.TabIndex = 0
        '
        'dsRep
        '
        Me.dsRep.DataSetName = "dsRep"
        Me.dsRep.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'dtRepBindingSource
        '
        Me.dtRepBindingSource.DataMember = "dtRep"
        Me.dtRepBindingSource.DataSource = Me.dsRep
        '
        'dtHeadRepBindingSource
        '
        Me.dtHeadRepBindingSource.DataMember = "dtHeadRep"
        Me.dtHeadRepBindingSource.DataSource = Me.dsRep
        '
        'dtHeadRep2BindingSource
        '
        Me.dtHeadRep2BindingSource.DataMember = "dtHeadRep2"
        Me.dtHeadRep2BindingSource.DataSource = Me.dsRep
        '
        'frmRepABP
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSize = True
        Me.ClientSize = New System.Drawing.Size(766, 430)
        Me.Controls.Add(Me.rvABP)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmRepABP"
        Me.Text = "Achievement By Product"
        CType(Me.dsRep, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dtRepBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dtHeadRepBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dtHeadRep2BindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents rvABP As Microsoft.Reporting.WinForms.ReportViewer
    Friend WithEvents dtRepBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents dsRep As Leveredge_Report_Tools.dsRep
    Friend WithEvents dtHeadRepBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents dtHeadRep2BindingSource As System.Windows.Forms.BindingSource
End Class
