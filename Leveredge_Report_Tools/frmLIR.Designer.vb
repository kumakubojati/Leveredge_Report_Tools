<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmLIR
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmLIR))
        Me.gbRepType_SPR = New System.Windows.Forms.GroupBox()
        Me.RBLIR_Detail = New System.Windows.Forms.RadioButton()
        Me.RBLIR_Recap = New System.Windows.Forms.RadioButton()
        Me.PicBar_LIR = New System.Windows.Forms.PictureBox()
        Me.btnBrow_LIR_dest = New System.Windows.Forms.Button()
        Me.txtLIR_dest = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.btnNeu_LIR = New System.Windows.Forms.Button()
        Me.btnBrow_LIR_src = New System.Windows.Forms.Button()
        Me.txtLIR_src = New System.Windows.Forms.TextBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.OFD_LIR = New System.Windows.Forms.OpenFileDialog()
        Me.SFD_LIR = New System.Windows.Forms.SaveFileDialog()
        Me.BWLIR = New System.ComponentModel.BackgroundWorker()
        Me.gbRepType_SPR.SuspendLayout()
        CType(Me.PicBar_LIR, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'gbRepType_SPR
        '
        Me.gbRepType_SPR.Controls.Add(Me.RBLIR_Detail)
        Me.gbRepType_SPR.Controls.Add(Me.RBLIR_Recap)
        Me.gbRepType_SPR.Location = New System.Drawing.Point(4, 6)
        Me.gbRepType_SPR.Name = "gbRepType_SPR"
        Me.gbRepType_SPR.Size = New System.Drawing.Size(156, 53)
        Me.gbRepType_SPR.TabIndex = 49
        Me.gbRepType_SPR.TabStop = False
        Me.gbRepType_SPR.Text = "Report Type"
        '
        'RBLIR_Detail
        '
        Me.RBLIR_Detail.AutoSize = True
        Me.RBLIR_Detail.Checked = True
        Me.RBLIR_Detail.Location = New System.Drawing.Point(7, 22)
        Me.RBLIR_Detail.Name = "RBLIR_Detail"
        Me.RBLIR_Detail.Size = New System.Drawing.Size(52, 17)
        Me.RBLIR_Detail.TabIndex = 1
        Me.RBLIR_Detail.TabStop = True
        Me.RBLIR_Detail.Text = "Detail"
        Me.RBLIR_Detail.UseVisualStyleBackColor = True
        '
        'RBLIR_Recap
        '
        Me.RBLIR_Recap.AutoSize = True
        Me.RBLIR_Recap.Location = New System.Drawing.Point(81, 22)
        Me.RBLIR_Recap.Name = "RBLIR_Recap"
        Me.RBLIR_Recap.Size = New System.Drawing.Size(57, 17)
        Me.RBLIR_Recap.TabIndex = 0
        Me.RBLIR_Recap.Text = "Rekap"
        Me.RBLIR_Recap.UseVisualStyleBackColor = True
        '
        'PicBar_LIR
        '
        Me.PicBar_LIR.Image = CType(resources.GetObject("PicBar_LIR.Image"), System.Drawing.Image)
        Me.PicBar_LIR.Location = New System.Drawing.Point(5, 125)
        Me.PicBar_LIR.Name = "PicBar_LIR"
        Me.PicBar_LIR.Size = New System.Drawing.Size(80, 80)
        Me.PicBar_LIR.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PicBar_LIR.TabIndex = 45
        Me.PicBar_LIR.TabStop = False
        Me.PicBar_LIR.Visible = False
        '
        'btnBrow_LIR_dest
        '
        Me.btnBrow_LIR_dest.Image = CType(resources.GetObject("btnBrow_LIR_dest.Image"), System.Drawing.Image)
        Me.btnBrow_LIR_dest.Location = New System.Drawing.Point(310, 97)
        Me.btnBrow_LIR_dest.Name = "btnBrow_LIR_dest"
        Me.btnBrow_LIR_dest.Size = New System.Drawing.Size(26, 23)
        Me.btnBrow_LIR_dest.TabIndex = 48
        Me.btnBrow_LIR_dest.UseVisualStyleBackColor = True
        '
        'txtLIR_dest
        '
        Me.txtLIR_dest.Location = New System.Drawing.Point(76, 98)
        Me.txtLIR_dest.Name = "txtLIR_dest"
        Me.txtLIR_dest.Size = New System.Drawing.Size(229, 20)
        Me.txtLIR_dest.TabIndex = 47
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(5, 103)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(60, 13)
        Me.Label7.TabIndex = 46
        Me.Label7.Text = "Destination"
        '
        'btnNeu_LIR
        '
        Me.btnNeu_LIR.Enabled = False
        Me.btnNeu_LIR.Image = CType(resources.GetObject("btnNeu_LIR.Image"), System.Drawing.Image)
        Me.btnNeu_LIR.Location = New System.Drawing.Point(343, 61)
        Me.btnNeu_LIR.Name = "btnNeu_LIR"
        Me.btnNeu_LIR.Size = New System.Drawing.Size(54, 59)
        Me.btnNeu_LIR.TabIndex = 44
        Me.btnNeu_LIR.UseVisualStyleBackColor = True
        '
        'btnBrow_LIR_src
        '
        Me.btnBrow_LIR_src.Image = CType(resources.GetObject("btnBrow_LIR_src.Image"), System.Drawing.Image)
        Me.btnBrow_LIR_src.Location = New System.Drawing.Point(310, 63)
        Me.btnBrow_LIR_src.Name = "btnBrow_LIR_src"
        Me.btnBrow_LIR_src.Size = New System.Drawing.Size(26, 23)
        Me.btnBrow_LIR_src.TabIndex = 43
        Me.btnBrow_LIR_src.UseVisualStyleBackColor = True
        '
        'txtLIR_src
        '
        Me.txtLIR_src.Location = New System.Drawing.Point(76, 64)
        Me.txtLIR_src.Name = "txtLIR_src"
        Me.txtLIR_src.Size = New System.Drawing.Size(229, 20)
        Me.txtLIR_src.TabIndex = 42
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(6, 69)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(41, 13)
        Me.Label13.TabIndex = 41
        Me.Label13.Text = "Source"
        '
        'OFD_LIR
        '
        Me.OFD_LIR.FileName = "Source File"
        Me.OFD_LIR.Filter = "Excel files (*.xls,*xlsx)|*.xls;*.xlsx"
        '
        'SFD_LIR
        '
        Me.SFD_LIR.Filter = "Excel files (*.xls,*.xlsx)|*.xls;*xlsx"
        '
        'BWLIR
        '
        '
        'frmLIR
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.ClientSize = New System.Drawing.Size(402, 210)
        Me.Controls.Add(Me.gbRepType_SPR)
        Me.Controls.Add(Me.PicBar_LIR)
        Me.Controls.Add(Me.btnBrow_LIR_dest)
        Me.Controls.Add(Me.txtLIR_dest)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.btnNeu_LIR)
        Me.Controls.Add(Me.btnBrow_LIR_src)
        Me.Controls.Add(Me.txtLIR_src)
        Me.Controls.Add(Me.Label13)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmLIR"
        Me.Text = "List Of Invoice Report"
        Me.gbRepType_SPR.ResumeLayout(False)
        Me.gbRepType_SPR.PerformLayout()
        CType(Me.PicBar_LIR, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents gbRepType_SPR As System.Windows.Forms.GroupBox
    Friend WithEvents RBLIR_Detail As System.Windows.Forms.RadioButton
    Friend WithEvents RBLIR_Recap As System.Windows.Forms.RadioButton
    Friend WithEvents PicBar_LIR As System.Windows.Forms.PictureBox
    Friend WithEvents btnBrow_LIR_dest As System.Windows.Forms.Button
    Friend WithEvents txtLIR_dest As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents btnNeu_LIR As System.Windows.Forms.Button
    Friend WithEvents btnBrow_LIR_src As System.Windows.Forms.Button
    Friend WithEvents txtLIR_src As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents OFD_LIR As System.Windows.Forms.OpenFileDialog
    Friend WithEvents SFD_LIR As System.Windows.Forms.SaveFileDialog
    Friend WithEvents BWLIR As System.ComponentModel.BackgroundWorker
End Class
