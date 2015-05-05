<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSPR
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmSPR))
        Me.gbRepType_SPR = New System.Windows.Forms.GroupBox()
        Me.RBSPR_Avg = New System.Windows.Forms.RadioButton()
        Me.RBSPR_Val = New System.Windows.Forms.RadioButton()
        Me.RBSPR_Qty = New System.Windows.Forms.RadioButton()
        Me.PicBar_SPR = New System.Windows.Forms.PictureBox()
        Me.btnBrow_SPR_dest = New System.Windows.Forms.Button()
        Me.txtSPR_dest = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.btnNeu_SPR = New System.Windows.Forms.Button()
        Me.btnBrow_SPR_src = New System.Windows.Forms.Button()
        Me.txtSPR_src = New System.Windows.Forms.TextBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.OFD_SPR = New System.Windows.Forms.OpenFileDialog()
        Me.SFD_SPR = New System.Windows.Forms.SaveFileDialog()
        Me.BWSPR = New System.ComponentModel.BackgroundWorker()
        Me.gbRepType_SPR.SuspendLayout()
        CType(Me.PicBar_SPR, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'gbRepType_SPR
        '
        Me.gbRepType_SPR.Controls.Add(Me.RBSPR_Avg)
        Me.gbRepType_SPR.Controls.Add(Me.RBSPR_Val)
        Me.gbRepType_SPR.Controls.Add(Me.RBSPR_Qty)
        Me.gbRepType_SPR.Location = New System.Drawing.Point(5, 7)
        Me.gbRepType_SPR.Name = "gbRepType_SPR"
        Me.gbRepType_SPR.Size = New System.Drawing.Size(228, 53)
        Me.gbRepType_SPR.TabIndex = 31
        Me.gbRepType_SPR.TabStop = False
        Me.gbRepType_SPR.Text = "Report Type"
        '
        'RBSPR_Avg
        '
        Me.RBSPR_Avg.AutoSize = True
        Me.RBSPR_Avg.Enabled = False
        Me.RBSPR_Avg.Location = New System.Drawing.Point(148, 22)
        Me.RBSPR_Avg.Name = "RBSPR_Avg"
        Me.RBSPR_Avg.Size = New System.Drawing.Size(65, 17)
        Me.RBSPR_Avg.TabIndex = 2
        Me.RBSPR_Avg.TabStop = True
        Me.RBSPR_Avg.Text = "Average"
        Me.RBSPR_Avg.UseVisualStyleBackColor = True
        '
        'RBSPR_Val
        '
        Me.RBSPR_Val.AutoSize = True
        Me.RBSPR_Val.Location = New System.Drawing.Point(85, 22)
        Me.RBSPR_Val.Name = "RBSPR_Val"
        Me.RBSPR_Val.Size = New System.Drawing.Size(52, 17)
        Me.RBSPR_Val.TabIndex = 1
        Me.RBSPR_Val.TabStop = True
        Me.RBSPR_Val.Text = "Value"
        Me.RBSPR_Val.UseVisualStyleBackColor = True
        '
        'RBSPR_Qty
        '
        Me.RBSPR_Qty.AutoSize = True
        Me.RBSPR_Qty.Enabled = False
        Me.RBSPR_Qty.Location = New System.Drawing.Point(7, 22)
        Me.RBSPR_Qty.Name = "RBSPR_Qty"
        Me.RBSPR_Qty.Size = New System.Drawing.Size(64, 17)
        Me.RBSPR_Qty.TabIndex = 0
        Me.RBSPR_Qty.TabStop = True
        Me.RBSPR_Qty.Text = "Quantity"
        Me.RBSPR_Qty.UseVisualStyleBackColor = True
        '
        'PicBar_SPR
        '
        Me.PicBar_SPR.Image = CType(resources.GetObject("PicBar_SPR.Image"), System.Drawing.Image)
        Me.PicBar_SPR.Location = New System.Drawing.Point(6, 126)
        Me.PicBar_SPR.Name = "PicBar_SPR"
        Me.PicBar_SPR.Size = New System.Drawing.Size(80, 80)
        Me.PicBar_SPR.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PicBar_SPR.TabIndex = 27
        Me.PicBar_SPR.TabStop = False
        Me.PicBar_SPR.Visible = False
        '
        'btnBrow_SPR_dest
        '
        Me.btnBrow_SPR_dest.Image = CType(resources.GetObject("btnBrow_SPR_dest.Image"), System.Drawing.Image)
        Me.btnBrow_SPR_dest.Location = New System.Drawing.Point(311, 98)
        Me.btnBrow_SPR_dest.Name = "btnBrow_SPR_dest"
        Me.btnBrow_SPR_dest.Size = New System.Drawing.Size(26, 23)
        Me.btnBrow_SPR_dest.TabIndex = 30
        Me.btnBrow_SPR_dest.UseVisualStyleBackColor = True
        '
        'txtSPR_dest
        '
        Me.txtSPR_dest.Location = New System.Drawing.Point(77, 99)
        Me.txtSPR_dest.Name = "txtSPR_dest"
        Me.txtSPR_dest.Size = New System.Drawing.Size(229, 20)
        Me.txtSPR_dest.TabIndex = 29
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(6, 104)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(60, 13)
        Me.Label7.TabIndex = 28
        Me.Label7.Text = "Destination"
        '
        'btnNeu_SPR
        '
        Me.btnNeu_SPR.Enabled = False
        Me.btnNeu_SPR.Image = CType(resources.GetObject("btnNeu_SPR.Image"), System.Drawing.Image)
        Me.btnNeu_SPR.Location = New System.Drawing.Point(344, 62)
        Me.btnNeu_SPR.Name = "btnNeu_SPR"
        Me.btnNeu_SPR.Size = New System.Drawing.Size(54, 59)
        Me.btnNeu_SPR.TabIndex = 26
        Me.btnNeu_SPR.UseVisualStyleBackColor = True
        '
        'btnBrow_SPR_src
        '
        Me.btnBrow_SPR_src.Image = CType(resources.GetObject("btnBrow_SPR_src.Image"), System.Drawing.Image)
        Me.btnBrow_SPR_src.Location = New System.Drawing.Point(311, 64)
        Me.btnBrow_SPR_src.Name = "btnBrow_SPR_src"
        Me.btnBrow_SPR_src.Size = New System.Drawing.Size(26, 23)
        Me.btnBrow_SPR_src.TabIndex = 25
        Me.btnBrow_SPR_src.UseVisualStyleBackColor = True
        '
        'txtSPR_src
        '
        Me.txtSPR_src.Location = New System.Drawing.Point(77, 65)
        Me.txtSPR_src.Name = "txtSPR_src"
        Me.txtSPR_src.Size = New System.Drawing.Size(229, 20)
        Me.txtSPR_src.TabIndex = 24
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(7, 70)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(41, 13)
        Me.Label13.TabIndex = 23
        Me.Label13.Text = "Source"
        '
        'OFD_SPR
        '
        Me.OFD_SPR.FileName = "Source File"
        Me.OFD_SPR.Filter = "Excel files (*.xls,*xlsx)|*.xls;*.xlsx"
        '
        'SFD_SPR
        '
        Me.SFD_SPR.Filter = "Excel files (*.xls,*.xlsx)|*.xls;*xlsx"
        '
        'BWSPR
        '
        '
        'frmSPR
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.ClientSize = New System.Drawing.Size(406, 216)
        Me.Controls.Add(Me.gbRepType_SPR)
        Me.Controls.Add(Me.PicBar_SPR)
        Me.Controls.Add(Me.btnBrow_SPR_dest)
        Me.Controls.Add(Me.txtSPR_dest)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.btnNeu_SPR)
        Me.Controls.Add(Me.btnBrow_SPR_src)
        Me.Controls.Add(Me.txtSPR_src)
        Me.Controls.Add(Me.Label13)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmSPR"
        Me.Text = "Sales Performance Report"
        Me.gbRepType_SPR.ResumeLayout(False)
        Me.gbRepType_SPR.PerformLayout()
        CType(Me.PicBar_SPR, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents gbRepType_SPR As System.Windows.Forms.GroupBox
    Friend WithEvents RBSPR_Avg As System.Windows.Forms.RadioButton
    Friend WithEvents RBSPR_Val As System.Windows.Forms.RadioButton
    Friend WithEvents RBSPR_Qty As System.Windows.Forms.RadioButton
    Friend WithEvents PicBar_SPR As System.Windows.Forms.PictureBox
    Friend WithEvents btnBrow_SPR_dest As System.Windows.Forms.Button
    Friend WithEvents txtSPR_dest As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents btnNeu_SPR As System.Windows.Forms.Button
    Friend WithEvents btnBrow_SPR_src As System.Windows.Forms.Button
    Friend WithEvents txtSPR_src As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents OFD_SPR As System.Windows.Forms.OpenFileDialog
    Friend WithEvents SFD_SPR As System.Windows.Forms.SaveFileDialog
    Friend WithEvents BWSPR As System.ComponentModel.BackgroundWorker
End Class
