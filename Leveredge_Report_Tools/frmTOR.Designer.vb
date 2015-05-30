<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmTOR
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmTOR))
        Me.gbRepType_SPR = New System.Windows.Forms.GroupBox()
        Me.RBTOR_RUP = New System.Windows.Forms.RadioButton()
        Me.RBTOR_Weight = New System.Windows.Forms.RadioButton()
        Me.PicBar_TOR = New System.Windows.Forms.PictureBox()
        Me.btnBrow_TOR_dest = New System.Windows.Forms.Button()
        Me.txtTOR_dest = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.btnNeu_TOR = New System.Windows.Forms.Button()
        Me.btnBrow_TOR_src = New System.Windows.Forms.Button()
        Me.txtTOR_src = New System.Windows.Forms.TextBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.OFD_TOR = New System.Windows.Forms.OpenFileDialog()
        Me.SFD_TOR = New System.Windows.Forms.SaveFileDialog()
        Me.BWTOR = New System.ComponentModel.BackgroundWorker()
        Me.gbRepType_SPR.SuspendLayout()
        CType(Me.PicBar_TOR, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'gbRepType_SPR
        '
        Me.gbRepType_SPR.Controls.Add(Me.RBTOR_RUP)
        Me.gbRepType_SPR.Controls.Add(Me.RBTOR_Weight)
        Me.gbRepType_SPR.Location = New System.Drawing.Point(5, 7)
        Me.gbRepType_SPR.Name = "gbRepType_SPR"
        Me.gbRepType_SPR.Size = New System.Drawing.Size(156, 53)
        Me.gbRepType_SPR.TabIndex = 40
        Me.gbRepType_SPR.TabStop = False
        Me.gbRepType_SPR.Text = "Report Type"
        '
        'RBTOR_RUP
        '
        Me.RBTOR_RUP.AutoSize = True
        Me.RBTOR_RUP.Checked = True
        Me.RBTOR_RUP.Location = New System.Drawing.Point(7, 22)
        Me.RBTOR_RUP.Name = "RBTOR_RUP"
        Me.RBTOR_RUP.Size = New System.Drawing.Size(59, 17)
        Me.RBTOR_RUP.TabIndex = 1
        Me.RBTOR_RUP.TabStop = True
        Me.RBTOR_RUP.Text = "Rupiah"
        Me.RBTOR_RUP.UseVisualStyleBackColor = True
        '
        'RBTOR_Weight
        '
        Me.RBTOR_Weight.AutoSize = True
        Me.RBTOR_Weight.Location = New System.Drawing.Point(81, 22)
        Me.RBTOR_Weight.Name = "RBTOR_Weight"
        Me.RBTOR_Weight.Size = New System.Drawing.Size(59, 17)
        Me.RBTOR_Weight.TabIndex = 0
        Me.RBTOR_Weight.TabStop = True
        Me.RBTOR_Weight.Text = "Weight"
        Me.RBTOR_Weight.UseVisualStyleBackColor = True
        '
        'PicBar_TOR
        '
        Me.PicBar_TOR.Image = CType(resources.GetObject("PicBar_TOR.Image"), System.Drawing.Image)
        Me.PicBar_TOR.Location = New System.Drawing.Point(6, 126)
        Me.PicBar_TOR.Name = "PicBar_TOR"
        Me.PicBar_TOR.Size = New System.Drawing.Size(80, 80)
        Me.PicBar_TOR.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PicBar_TOR.TabIndex = 36
        Me.PicBar_TOR.TabStop = False
        Me.PicBar_TOR.Visible = False
        '
        'btnBrow_TOR_dest
        '
        Me.btnBrow_TOR_dest.Image = CType(resources.GetObject("btnBrow_TOR_dest.Image"), System.Drawing.Image)
        Me.btnBrow_TOR_dest.Location = New System.Drawing.Point(311, 98)
        Me.btnBrow_TOR_dest.Name = "btnBrow_TOR_dest"
        Me.btnBrow_TOR_dest.Size = New System.Drawing.Size(26, 23)
        Me.btnBrow_TOR_dest.TabIndex = 39
        Me.btnBrow_TOR_dest.UseVisualStyleBackColor = True
        '
        'txtTOR_dest
        '
        Me.txtTOR_dest.Location = New System.Drawing.Point(77, 99)
        Me.txtTOR_dest.Name = "txtTOR_dest"
        Me.txtTOR_dest.Size = New System.Drawing.Size(229, 20)
        Me.txtTOR_dest.TabIndex = 38
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(6, 104)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(60, 13)
        Me.Label7.TabIndex = 37
        Me.Label7.Text = "Destination"
        '
        'btnNeu_TOR
        '
        Me.btnNeu_TOR.Enabled = False
        Me.btnNeu_TOR.Image = CType(resources.GetObject("btnNeu_TOR.Image"), System.Drawing.Image)
        Me.btnNeu_TOR.Location = New System.Drawing.Point(344, 62)
        Me.btnNeu_TOR.Name = "btnNeu_TOR"
        Me.btnNeu_TOR.Size = New System.Drawing.Size(54, 59)
        Me.btnNeu_TOR.TabIndex = 35
        Me.btnNeu_TOR.UseVisualStyleBackColor = True
        '
        'btnBrow_TOR_src
        '
        Me.btnBrow_TOR_src.Image = CType(resources.GetObject("btnBrow_TOR_src.Image"), System.Drawing.Image)
        Me.btnBrow_TOR_src.Location = New System.Drawing.Point(311, 64)
        Me.btnBrow_TOR_src.Name = "btnBrow_TOR_src"
        Me.btnBrow_TOR_src.Size = New System.Drawing.Size(26, 23)
        Me.btnBrow_TOR_src.TabIndex = 34
        Me.btnBrow_TOR_src.UseVisualStyleBackColor = True
        '
        'txtTOR_src
        '
        Me.txtTOR_src.Location = New System.Drawing.Point(77, 65)
        Me.txtTOR_src.Name = "txtTOR_src"
        Me.txtTOR_src.Size = New System.Drawing.Size(229, 20)
        Me.txtTOR_src.TabIndex = 33
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(7, 70)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(41, 13)
        Me.Label13.TabIndex = 32
        Me.Label13.Text = "Source"
        '
        'OFD_TOR
        '
        Me.OFD_TOR.FileName = "File Source"
        Me.OFD_TOR.Filter = "Excel files (*.xls,*xlsx)|*.xls;*.xlsx"
        '
        'SFD_TOR
        '
        Me.SFD_TOR.Filter = "Excel files (*.xls,*.xlsx)|*.xls;*xlsx"
        '
        'BWTOR
        '
        '
        'frmTOR
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.ClientSize = New System.Drawing.Size(404, 209)
        Me.Controls.Add(Me.gbRepType_SPR)
        Me.Controls.Add(Me.PicBar_TOR)
        Me.Controls.Add(Me.btnBrow_TOR_dest)
        Me.Controls.Add(Me.txtTOR_dest)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.btnNeu_TOR)
        Me.Controls.Add(Me.btnBrow_TOR_src)
        Me.Controls.Add(Me.txtTOR_src)
        Me.Controls.Add(Me.Label13)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmTOR"
        Me.Text = "Turn Over Report"
        Me.gbRepType_SPR.ResumeLayout(False)
        Me.gbRepType_SPR.PerformLayout()
        CType(Me.PicBar_TOR, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents gbRepType_SPR As System.Windows.Forms.GroupBox
    Friend WithEvents RBTOR_RUP As System.Windows.Forms.RadioButton
    Friend WithEvents RBTOR_Weight As System.Windows.Forms.RadioButton
    Friend WithEvents PicBar_TOR As System.Windows.Forms.PictureBox
    Friend WithEvents btnBrow_TOR_dest As System.Windows.Forms.Button
    Friend WithEvents txtTOR_dest As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents btnNeu_TOR As System.Windows.Forms.Button
    Friend WithEvents btnBrow_TOR_src As System.Windows.Forms.Button
    Friend WithEvents txtTOR_src As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents OFD_TOR As System.Windows.Forms.OpenFileDialog
    Friend WithEvents SFD_TOR As System.Windows.Forms.SaveFileDialog
    Friend WithEvents BWTOR As System.ComponentModel.BackgroundWorker
End Class
