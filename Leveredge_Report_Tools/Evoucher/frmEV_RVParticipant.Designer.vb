<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmEV_RVParticipant
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmEV_RVParticipant))
        Me.RV_EVParticipant = New Microsoft.Reporting.WinForms.ReportViewer()
        Me.DS_EV = New Leveredge_Report_Tools.DS_EV()
        Me.EV_PARTIBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.dtHeader_EVBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        CType(Me.DS_EV, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.EV_PARTIBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dtHeader_EVBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'RV_EVParticipant
        '
        Me.RV_EVParticipant.Dock = System.Windows.Forms.DockStyle.Fill
        ReportDataSource1.Name = "dsEV_PARTI"
        ReportDataSource1.Value = Me.EV_PARTIBindingSource
        ReportDataSource2.Name = "HDR_DS"
        ReportDataSource2.Value = Me.dtHeader_EVBindingSource
        Me.RV_EVParticipant.LocalReport.DataSources.Add(ReportDataSource1)
        Me.RV_EVParticipant.LocalReport.DataSources.Add(ReportDataSource2)
        Me.RV_EVParticipant.LocalReport.ReportEmbeddedResource = "Leveredge_Report_Tools.EV_Participant.rdlc"
        Me.RV_EVParticipant.Location = New System.Drawing.Point(0, 0)
        Me.RV_EVParticipant.Name = "RV_EVParticipant"
        Me.RV_EVParticipant.Size = New System.Drawing.Size(937, 505)
        Me.RV_EVParticipant.TabIndex = 0
        '
        'DS_EV
        '
        Me.DS_EV.DataSetName = "DS_EV"
        Me.DS_EV.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'EV_PARTIBindingSource
        '
        Me.EV_PARTIBindingSource.DataMember = "EV_PARTI"
        Me.EV_PARTIBindingSource.DataSource = Me.DS_EV
        '
        'dtHeader_EVBindingSource
        '
        Me.dtHeader_EVBindingSource.DataMember = "dtHeader_EV"
        Me.dtHeader_EVBindingSource.DataSource = Me.DS_EV
        '
        'frmEV_RVParticipant
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(937, 505)
        Me.Controls.Add(Me.RV_EVParticipant)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmEV_RVParticipant"
        Me.Text = "E-Voucher Participant Report"
        CType(Me.DS_EV, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.EV_PARTIBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dtHeader_EVBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents RV_EVParticipant As Microsoft.Reporting.WinForms.ReportViewer
    Friend WithEvents EV_PARTIBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents DS_EV As Leveredge_Report_Tools.DS_EV
    Friend WithEvents dtHeader_EVBindingSource As System.Windows.Forms.BindingSource
End Class
