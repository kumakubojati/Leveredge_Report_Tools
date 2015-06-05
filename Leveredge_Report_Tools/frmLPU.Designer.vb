<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmLPU
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmLPU))
        Me.PicBar_LPU = New System.Windows.Forms.PictureBox()
        Me.btnBrow_LPU_dest = New System.Windows.Forms.Button()
        Me.txtLPU_dest = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.btnNeu_LPU = New System.Windows.Forms.Button()
        Me.btnBrow_LPU_src = New System.Windows.Forms.Button()
        Me.txtLPU_src = New System.Windows.Forms.TextBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.OFD_LPU = New System.Windows.Forms.OpenFileDialog()
        Me.SFD_LPU = New System.Windows.Forms.SaveFileDialog()
        Me.BWLPU = New System.ComponentModel.BackgroundWorker()
        CType(Me.PicBar_LPU, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PicBar_LPU
        '
        Me.PicBar_LPU.Image = CType(resources.GetObject("PicBar_LPU.Image"), System.Drawing.Image)
        Me.PicBar_LPU.Location = New System.Drawing.Point(5, 79)
        Me.PicBar_LPU.Name = "PicBar_LPU"
        Me.PicBar_LPU.Size = New System.Drawing.Size(80, 80)
        Me.PicBar_LPU.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PicBar_LPU.TabIndex = 62
        Me.PicBar_LPU.TabStop = False
        Me.PicBar_LPU.Visible = False
        '
        'btnBrow_LPU_dest
        '
        Me.btnBrow_LPU_dest.Image = CType(resources.GetObject("btnBrow_LPU_dest.Image"), System.Drawing.Image)
        Me.btnBrow_LPU_dest.Location = New System.Drawing.Point(310, 51)
        Me.btnBrow_LPU_dest.Name = "btnBrow_LPU_dest"
        Me.btnBrow_LPU_dest.Size = New System.Drawing.Size(26, 23)
        Me.btnBrow_LPU_dest.TabIndex = 65
        Me.btnBrow_LPU_dest.UseVisualStyleBackColor = True
        '
        'txtLPU_dest
        '
        Me.txtLPU_dest.Location = New System.Drawing.Point(76, 52)
        Me.txtLPU_dest.Name = "txtLPU_dest"
        Me.txtLPU_dest.Size = New System.Drawing.Size(229, 20)
        Me.txtLPU_dest.TabIndex = 64
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(5, 57)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(60, 13)
        Me.Label7.TabIndex = 63
        Me.Label7.Text = "Destination"
        '
        'btnNeu_LPU
        '
        Me.btnNeu_LPU.Enabled = False
        Me.btnNeu_LPU.Image = CType(resources.GetObject("btnNeu_LPU.Image"), System.Drawing.Image)
        Me.btnNeu_LPU.Location = New System.Drawing.Point(343, 15)
        Me.btnNeu_LPU.Name = "btnNeu_LPU"
        Me.btnNeu_LPU.Size = New System.Drawing.Size(54, 59)
        Me.btnNeu_LPU.TabIndex = 61
        Me.btnNeu_LPU.UseVisualStyleBackColor = True
        '
        'btnBrow_LPU_src
        '
        Me.btnBrow_LPU_src.Image = CType(resources.GetObject("btnBrow_LPU_src.Image"), System.Drawing.Image)
        Me.btnBrow_LPU_src.Location = New System.Drawing.Point(310, 17)
        Me.btnBrow_LPU_src.Name = "btnBrow_LPU_src"
        Me.btnBrow_LPU_src.Size = New System.Drawing.Size(26, 23)
        Me.btnBrow_LPU_src.TabIndex = 60
        Me.btnBrow_LPU_src.UseVisualStyleBackColor = True
        '
        'txtLPU_src
        '
        Me.txtLPU_src.Location = New System.Drawing.Point(76, 18)
        Me.txtLPU_src.Name = "txtLPU_src"
        Me.txtLPU_src.Size = New System.Drawing.Size(229, 20)
        Me.txtLPU_src.TabIndex = 59
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(6, 23)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(41, 13)
        Me.Label13.TabIndex = 58
        Me.Label13.Text = "Source"
        '
        'OFD_LPU
        '
        Me.OFD_LPU.FileName = "Source File"
        Me.OFD_LPU.Filter = "Excel files (*.xls,*xlsx)|*.xls;*.xlsx"
        '
        'SFD_LPU
        '
        Me.SFD_LPU.Filter = "Excel files (*.xls,*.xlsx)|*.xls;*xlsx"
        '
        'BWLPU
        '
        '
        'frmLPU
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.ClientSize = New System.Drawing.Size(403, 169)
        Me.Controls.Add(Me.PicBar_LPU)
        Me.Controls.Add(Me.btnBrow_LPU_dest)
        Me.Controls.Add(Me.txtLPU_dest)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.btnNeu_LPU)
        Me.Controls.Add(Me.btnBrow_LPU_src)
        Me.Controls.Add(Me.txtLPU_src)
        Me.Controls.Add(Me.Label13)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmLPU"
        Me.Text = "List Of Prormotion Utilization"
        CType(Me.PicBar_LPU, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents PicBar_LPU As System.Windows.Forms.PictureBox
    Friend WithEvents btnBrow_LPU_dest As System.Windows.Forms.Button
    Friend WithEvents txtLPU_dest As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents btnNeu_LPU As System.Windows.Forms.Button
    Friend WithEvents btnBrow_LPU_src As System.Windows.Forms.Button
    Friend WithEvents txtLPU_src As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents OFD_LPU As System.Windows.Forms.OpenFileDialog
    Friend WithEvents SFD_LPU As System.Windows.Forms.SaveFileDialog
    Friend WithEvents BWLPU As System.ComponentModel.BackgroundWorker
End Class
