<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmONI_IC
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmONI_IC))
        Me.gbRepType_DD = New System.Windows.Forms.GroupBox()
        Me.RBONI_Detail = New System.Windows.Forms.RadioButton()
        Me.RBONI_Summary = New System.Windows.Forms.RadioButton()
        Me.PicBar_ONI_IC = New System.Windows.Forms.PictureBox()
        Me.btnBrow_ONI_IC_dest = New System.Windows.Forms.Button()
        Me.txtONI_IC_dest = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.btnNeu_ONI_IC = New System.Windows.Forms.Button()
        Me.btnBrow_ONI_IC_src = New System.Windows.Forms.Button()
        Me.txtONI_IC_src = New System.Windows.Forms.TextBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.OFD_ONI_IC = New System.Windows.Forms.OpenFileDialog()
        Me.SFD_ONI_IC = New System.Windows.Forms.SaveFileDialog()
        Me.BWONI_IC = New System.ComponentModel.BackgroundWorker()
        Me.gbRepType_DD.SuspendLayout()
        CType(Me.PicBar_ONI_IC, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'gbRepType_DD
        '
        Me.gbRepType_DD.Controls.Add(Me.RBONI_Detail)
        Me.gbRepType_DD.Controls.Add(Me.RBONI_Summary)
        Me.gbRepType_DD.Location = New System.Drawing.Point(8, 10)
        Me.gbRepType_DD.Name = "gbRepType_DD"
        Me.gbRepType_DD.Size = New System.Drawing.Size(167, 53)
        Me.gbRepType_DD.TabIndex = 85
        Me.gbRepType_DD.TabStop = False
        Me.gbRepType_DD.Text = "Report Type"
        '
        'RBONI_Detail
        '
        Me.RBONI_Detail.AutoSize = True
        Me.RBONI_Detail.Checked = True
        Me.RBONI_Detail.Location = New System.Drawing.Point(7, 22)
        Me.RBONI_Detail.Name = "RBONI_Detail"
        Me.RBONI_Detail.Size = New System.Drawing.Size(52, 17)
        Me.RBONI_Detail.TabIndex = 1
        Me.RBONI_Detail.TabStop = True
        Me.RBONI_Detail.Text = "Detail"
        Me.RBONI_Detail.UseVisualStyleBackColor = True
        '
        'RBONI_Summary
        '
        Me.RBONI_Summary.AutoSize = True
        Me.RBONI_Summary.Location = New System.Drawing.Point(85, 22)
        Me.RBONI_Summary.Name = "RBONI_Summary"
        Me.RBONI_Summary.Size = New System.Drawing.Size(68, 17)
        Me.RBONI_Summary.TabIndex = 0
        Me.RBONI_Summary.Text = "Summary"
        Me.RBONI_Summary.UseVisualStyleBackColor = True
        '
        'PicBar_ONI_IC
        '
        Me.PicBar_ONI_IC.Image = CType(resources.GetObject("PicBar_ONI_IC.Image"), System.Drawing.Image)
        Me.PicBar_ONI_IC.Location = New System.Drawing.Point(9, 129)
        Me.PicBar_ONI_IC.Name = "PicBar_ONI_IC"
        Me.PicBar_ONI_IC.Size = New System.Drawing.Size(80, 80)
        Me.PicBar_ONI_IC.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PicBar_ONI_IC.TabIndex = 81
        Me.PicBar_ONI_IC.TabStop = False
        Me.PicBar_ONI_IC.Visible = False
        '
        'btnBrow_ONI_IC_dest
        '
        Me.btnBrow_ONI_IC_dest.Image = CType(resources.GetObject("btnBrow_ONI_IC_dest.Image"), System.Drawing.Image)
        Me.btnBrow_ONI_IC_dest.Location = New System.Drawing.Point(314, 101)
        Me.btnBrow_ONI_IC_dest.Name = "btnBrow_ONI_IC_dest"
        Me.btnBrow_ONI_IC_dest.Size = New System.Drawing.Size(26, 23)
        Me.btnBrow_ONI_IC_dest.TabIndex = 84
        Me.btnBrow_ONI_IC_dest.UseVisualStyleBackColor = True
        '
        'txtONI_IC_dest
        '
        Me.txtONI_IC_dest.Location = New System.Drawing.Point(80, 102)
        Me.txtONI_IC_dest.Name = "txtONI_IC_dest"
        Me.txtONI_IC_dest.Size = New System.Drawing.Size(229, 20)
        Me.txtONI_IC_dest.TabIndex = 83
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(9, 107)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(60, 13)
        Me.Label7.TabIndex = 82
        Me.Label7.Text = "Destination"
        '
        'btnNeu_ONI_IC
        '
        Me.btnNeu_ONI_IC.Enabled = False
        Me.btnNeu_ONI_IC.Image = CType(resources.GetObject("btnNeu_ONI_IC.Image"), System.Drawing.Image)
        Me.btnNeu_ONI_IC.Location = New System.Drawing.Point(347, 65)
        Me.btnNeu_ONI_IC.Name = "btnNeu_ONI_IC"
        Me.btnNeu_ONI_IC.Size = New System.Drawing.Size(54, 59)
        Me.btnNeu_ONI_IC.TabIndex = 80
        Me.btnNeu_ONI_IC.UseVisualStyleBackColor = True
        '
        'btnBrow_ONI_IC_src
        '
        Me.btnBrow_ONI_IC_src.Image = CType(resources.GetObject("btnBrow_ONI_IC_src.Image"), System.Drawing.Image)
        Me.btnBrow_ONI_IC_src.Location = New System.Drawing.Point(314, 67)
        Me.btnBrow_ONI_IC_src.Name = "btnBrow_ONI_IC_src"
        Me.btnBrow_ONI_IC_src.Size = New System.Drawing.Size(26, 23)
        Me.btnBrow_ONI_IC_src.TabIndex = 79
        Me.btnBrow_ONI_IC_src.UseVisualStyleBackColor = True
        '
        'txtONI_IC_src
        '
        Me.txtONI_IC_src.Location = New System.Drawing.Point(80, 68)
        Me.txtONI_IC_src.Name = "txtONI_IC_src"
        Me.txtONI_IC_src.Size = New System.Drawing.Size(229, 20)
        Me.txtONI_IC_src.TabIndex = 78
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(10, 73)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(41, 13)
        Me.Label13.TabIndex = 77
        Me.Label13.Text = "Source"
        '
        'OFD_ONI_IC
        '
        Me.OFD_ONI_IC.FileName = "Source File"
        Me.OFD_ONI_IC.Filter = "Excel files (*.xls,*xlsx)|*.xls;*.xlsx"
        '
        'SFD_ONI_IC
        '
        Me.SFD_ONI_IC.Filter = "Excel files (*.xls,*.xlsx)|*.xls;*xlsx"
        '
        'frmONI_IC
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.ClientSize = New System.Drawing.Size(410, 214)
        Me.Controls.Add(Me.gbRepType_DD)
        Me.Controls.Add(Me.PicBar_ONI_IC)
        Me.Controls.Add(Me.btnBrow_ONI_IC_dest)
        Me.Controls.Add(Me.txtONI_IC_dest)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.btnNeu_ONI_IC)
        Me.Controls.Add(Me.btnBrow_ONI_IC_src)
        Me.Controls.Add(Me.txtONI_IC_src)
        Me.Controls.Add(Me.Label13)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmONI_IC"
        Me.Text = "Outlet Net Increase Report"
        Me.gbRepType_DD.ResumeLayout(False)
        Me.gbRepType_DD.PerformLayout()
        CType(Me.PicBar_ONI_IC, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents gbRepType_DD As System.Windows.Forms.GroupBox
    Friend WithEvents RBONI_Detail As System.Windows.Forms.RadioButton
    Friend WithEvents RBONI_Summary As System.Windows.Forms.RadioButton
    Friend WithEvents PicBar_ONI_IC As System.Windows.Forms.PictureBox
    Friend WithEvents btnBrow_ONI_IC_dest As System.Windows.Forms.Button
    Friend WithEvents txtONI_IC_dest As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents btnNeu_ONI_IC As System.Windows.Forms.Button
    Friend WithEvents btnBrow_ONI_IC_src As System.Windows.Forms.Button
    Friend WithEvents txtONI_IC_src As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents OFD_ONI_IC As System.Windows.Forms.OpenFileDialog
    Friend WithEvents SFD_ONI_IC As System.Windows.Forms.SaveFileDialog
    Friend WithEvents BWONI_IC As System.ComponentModel.BackgroundWorker
End Class
