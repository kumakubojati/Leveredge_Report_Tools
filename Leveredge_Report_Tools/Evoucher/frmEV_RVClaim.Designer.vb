<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmEV_RVClaim
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmEV_RVClaim))
        Me.EV_REPBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.DS_EV = New Leveredge_Report_Tools.DS_EV()
        Me.dtHeader_EVBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.RV_EV_CLAIM = New Microsoft.Reporting.WinForms.ReportViewer()
        CType(Me.EV_REPBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DS_EV, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dtHeader_EVBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'EV_REPBindingSource
        '
        Me.EV_REPBindingSource.DataMember = "EV_REP"
        Me.EV_REPBindingSource.DataSource = Me.DS_EV
        '
        'DS_EV
        '
        Me.DS_EV.DataSetName = "DS_EV"
        Me.DS_EV.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'dtHeader_EVBindingSource
        '
        Me.dtHeader_EVBindingSource.DataMember = "dtHeader_EV"
        Me.dtHeader_EVBindingSource.DataSource = Me.DS_EV
        '
        'RV_EV_CLAIM
        '
        ReportDataSource1.Name = "EV_DATASET"
        ReportDataSource1.Value = Me.EV_REPBindingSource
        ReportDataSource2.Name = "HDR_DATASET"
        ReportDataSource2.Value = Me.dtHeader_EVBindingSource
        Me.RV_EV_CLAIM.LocalReport.DataSources.Add(ReportDataSource1)
        Me.RV_EV_CLAIM.LocalReport.DataSources.Add(ReportDataSource2)
        Me.RV_EV_CLAIM.LocalReport.ReportEmbeddedResource = "Leveredge_Report_Tools.EV_Claim.rdlc"
        Me.RV_EV_CLAIM.Location = New System.Drawing.Point(1, 1)
        Me.RV_EV_CLAIM.Name = "RV_EV_CLAIM"
        Me.RV_EV_CLAIM.Size = New System.Drawing.Size(1024, 586)
        Me.RV_EV_CLAIM.TabIndex = 0
        '
        'frmEV_RVClaim
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1027, 588)
        Me.Controls.Add(Me.RV_EV_CLAIM)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmEV_RVClaim"
        Me.Text = "E-Voucher Claim Report"
        CType(Me.EV_REPBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DS_EV, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dtHeader_EVBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents RV_EV_CLAIM As Microsoft.Reporting.WinForms.ReportViewer
    Friend WithEvents EV_REPBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents DS_EV As Leveredge_Report_Tools.DS_EV
    Friend WithEvents dtHeader_EVBindingSource As System.Windows.Forms.BindingSource
End Class
