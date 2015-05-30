<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSL
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmSL))
        Me.gbRepType_SPR = New System.Windows.Forms.GroupBox()
        Me.RBSL_Rekap = New System.Windows.Forms.RadioButton()
        Me.RBSL_Detail = New System.Windows.Forms.RadioButton()
        Me.PicBar_SL = New System.Windows.Forms.PictureBox()
        Me.btnBrow_SL_dest = New System.Windows.Forms.Button()
        Me.txtSL_dest = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.btnNeu_SL = New System.Windows.Forms.Button()
        Me.btnBrow_SL_src = New System.Windows.Forms.Button()
        Me.txtSL_src = New System.Windows.Forms.TextBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.OFD_SL = New System.Windows.Forms.OpenFileDialog()
        Me.SFD_SL = New System.Windows.Forms.SaveFileDialog()
        Me.BW_SL = New System.ComponentModel.BackgroundWorker()
        Me.gbRepType_SPR.SuspendLayout()
        CType(Me.PicBar_SL, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'gbRepType_SPR
        '
        Me.gbRepType_SPR.Controls.Add(Me.RBSL_Rekap)
        Me.gbRepType_SPR.Controls.Add(Me.RBSL_Detail)
        Me.gbRepType_SPR.Location = New System.Drawing.Point(7, 4)
        Me.gbRepType_SPR.Name = "gbRepType_SPR"
        Me.gbRepType_SPR.Size = New System.Drawing.Size(191, 53)
        Me.gbRepType_SPR.TabIndex = 58
        Me.gbRepType_SPR.TabStop = False
        Me.gbRepType_SPR.Text = "Report Type"
        '
        'RBSL_Rekap
        '
        Me.RBSL_Rekap.AutoSize = True
        Me.RBSL_Rekap.Checked = True
        Me.RBSL_Rekap.Location = New System.Drawing.Point(7, 22)
        Me.RBSL_Rekap.Name = "RBSL_Rekap"
        Me.RBSL_Rekap.Size = New System.Drawing.Size(106, 17)
        Me.RBSL_Rekap.TabIndex = 1
        Me.RBSL_Rekap.TabStop = True
        Me.RBSL_Rekap.Text = "Rekap(Summary)"
        Me.RBSL_Rekap.UseVisualStyleBackColor = True
        '
        'RBSL_Detail
        '
        Me.RBSL_Detail.AutoSize = True
        Me.RBSL_Detail.Location = New System.Drawing.Point(119, 22)
        Me.RBSL_Detail.Name = "RBSL_Detail"
        Me.RBSL_Detail.Size = New System.Drawing.Size(52, 17)
        Me.RBSL_Detail.TabIndex = 0
        Me.RBSL_Detail.Text = "Detail"
        Me.RBSL_Detail.UseVisualStyleBackColor = True
        '
        'PicBar_SL
        '
        Me.PicBar_SL.Image = CType(resources.GetObject("PicBar_SL.Image"), System.Drawing.Image)
        Me.PicBar_SL.Location = New System.Drawing.Point(8, 123)
        Me.PicBar_SL.Name = "PicBar_SL"
        Me.PicBar_SL.Size = New System.Drawing.Size(80, 80)
        Me.PicBar_SL.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PicBar_SL.TabIndex = 54
        Me.PicBar_SL.TabStop = False
        Me.PicBar_SL.Visible = False
        '
        'btnBrow_SL_dest
        '
        Me.btnBrow_SL_dest.Image = CType(resources.GetObject("btnBrow_SL_dest.Image"), System.Drawing.Image)
        Me.btnBrow_SL_dest.Location = New System.Drawing.Point(313, 95)
        Me.btnBrow_SL_dest.Name = "btnBrow_SL_dest"
        Me.btnBrow_SL_dest.Size = New System.Drawing.Size(26, 23)
        Me.btnBrow_SL_dest.TabIndex = 57
        Me.btnBrow_SL_dest.UseVisualStyleBackColor = True
        '
        'txtSL_dest
        '
        Me.txtSL_dest.Location = New System.Drawing.Point(79, 96)
        Me.txtSL_dest.Name = "txtSL_dest"
        Me.txtSL_dest.Size = New System.Drawing.Size(229, 20)
        Me.txtSL_dest.TabIndex = 56
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(8, 101)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(60, 13)
        Me.Label7.TabIndex = 55
        Me.Label7.Text = "Destination"
        '
        'btnNeu_SL
        '
        Me.btnNeu_SL.Enabled = False
        Me.btnNeu_SL.Image = CType(resources.GetObject("btnNeu_SL.Image"), System.Drawing.Image)
        Me.btnNeu_SL.Location = New System.Drawing.Point(346, 59)
        Me.btnNeu_SL.Name = "btnNeu_SL"
        Me.btnNeu_SL.Size = New System.Drawing.Size(54, 59)
        Me.btnNeu_SL.TabIndex = 53
        Me.btnNeu_SL.UseVisualStyleBackColor = True
        '
        'btnBrow_SL_src
        '
        Me.btnBrow_SL_src.Image = CType(resources.GetObject("btnBrow_SL_src.Image"), System.Drawing.Image)
        Me.btnBrow_SL_src.Location = New System.Drawing.Point(313, 61)
        Me.btnBrow_SL_src.Name = "btnBrow_SL_src"
        Me.btnBrow_SL_src.Size = New System.Drawing.Size(26, 23)
        Me.btnBrow_SL_src.TabIndex = 52
        Me.btnBrow_SL_src.UseVisualStyleBackColor = True
        '
        'txtSL_src
        '
        Me.txtSL_src.Location = New System.Drawing.Point(79, 62)
        Me.txtSL_src.Name = "txtSL_src"
        Me.txtSL_src.Size = New System.Drawing.Size(229, 20)
        Me.txtSL_src.TabIndex = 51
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(9, 67)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(41, 13)
        Me.Label13.TabIndex = 50
        Me.Label13.Text = "Source"
        '
        'OFD_SL
        '
        Me.OFD_SL.FileName = "Source File"
        Me.OFD_SL.Filter = "Excel files (*.xls,*xlsx)|*.xls;*.xlsx"
        '
        'SFD_SL
        '
        Me.SFD_SL.Filter = "Excel files (*.xls,*.xlsx)|*.xls;*xlsx"
        '
        'BW_SL
        '
        '
        'frmSL
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.ClientSize = New System.Drawing.Size(406, 207)
        Me.Controls.Add(Me.gbRepType_SPR)
        Me.Controls.Add(Me.PicBar_SL)
        Me.Controls.Add(Me.btnBrow_SL_dest)
        Me.Controls.Add(Me.txtSL_dest)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.btnNeu_SL)
        Me.Controls.Add(Me.btnBrow_SL_src)
        Me.Controls.Add(Me.txtSL_src)
        Me.Controls.Add(Me.Label13)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmSL"
        Me.Text = "Service Level Report"
        Me.gbRepType_SPR.ResumeLayout(False)
        Me.gbRepType_SPR.PerformLayout()
        CType(Me.PicBar_SL, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents gbRepType_SPR As System.Windows.Forms.GroupBox
    Friend WithEvents RBSL_Rekap As System.Windows.Forms.RadioButton
    Friend WithEvents RBSL_Detail As System.Windows.Forms.RadioButton
    Friend WithEvents PicBar_SL As System.Windows.Forms.PictureBox
    Friend WithEvents btnBrow_SL_dest As System.Windows.Forms.Button
    Friend WithEvents txtSL_dest As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents btnNeu_SL As System.Windows.Forms.Button
    Friend WithEvents btnBrow_SL_src As System.Windows.Forms.Button
    Friend WithEvents txtSL_src As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents OFD_SL As System.Windows.Forms.OpenFileDialog
    Friend WithEvents SFD_SL As System.Windows.Forms.SaveFileDialog
    Friend WithEvents BW_SL As System.ComponentModel.BackgroundWorker
End Class
