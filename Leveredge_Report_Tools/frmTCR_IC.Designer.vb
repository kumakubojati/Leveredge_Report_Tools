<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmTCR_IC
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmTCR_IC))
        Me.PicBar_TCR_IC = New System.Windows.Forms.PictureBox()
        Me.btnBrow_TCR_IC_dest = New System.Windows.Forms.Button()
        Me.txtTCR_IC_dest = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.btnNeu_TCR_IC = New System.Windows.Forms.Button()
        Me.btnBrow_TCR_IC_src = New System.Windows.Forms.Button()
        Me.txtTCR_IC_src = New System.Windows.Forms.TextBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.OFD_TCR_IC = New System.Windows.Forms.OpenFileDialog()
        Me.SFD_TCR_IC = New System.Windows.Forms.SaveFileDialog()
        Me.BWTCR_IC = New System.ComponentModel.BackgroundWorker()
        CType(Me.PicBar_TCR_IC, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PicBar_TCR_IC
        '
        Me.PicBar_TCR_IC.Image = CType(resources.GetObject("PicBar_TCR_IC.Image"), System.Drawing.Image)
        Me.PicBar_TCR_IC.Location = New System.Drawing.Point(7, 71)
        Me.PicBar_TCR_IC.Name = "PicBar_TCR_IC"
        Me.PicBar_TCR_IC.Size = New System.Drawing.Size(80, 80)
        Me.PicBar_TCR_IC.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PicBar_TCR_IC.TabIndex = 71
        Me.PicBar_TCR_IC.TabStop = False
        Me.PicBar_TCR_IC.Visible = False
        '
        'btnBrow_TCR_IC_dest
        '
        Me.btnBrow_TCR_IC_dest.Image = CType(resources.GetObject("btnBrow_TCR_IC_dest.Image"), System.Drawing.Image)
        Me.btnBrow_TCR_IC_dest.Location = New System.Drawing.Point(312, 43)
        Me.btnBrow_TCR_IC_dest.Name = "btnBrow_TCR_IC_dest"
        Me.btnBrow_TCR_IC_dest.Size = New System.Drawing.Size(26, 23)
        Me.btnBrow_TCR_IC_dest.TabIndex = 74
        Me.btnBrow_TCR_IC_dest.UseVisualStyleBackColor = True
        '
        'txtTCR_IC_dest
        '
        Me.txtTCR_IC_dest.Location = New System.Drawing.Point(78, 44)
        Me.txtTCR_IC_dest.Name = "txtTCR_IC_dest"
        Me.txtTCR_IC_dest.Size = New System.Drawing.Size(229, 20)
        Me.txtTCR_IC_dest.TabIndex = 73
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(7, 49)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(60, 13)
        Me.Label7.TabIndex = 72
        Me.Label7.Text = "Destination"
        '
        'btnNeu_TCR_IC
        '
        Me.btnNeu_TCR_IC.Enabled = False
        Me.btnNeu_TCR_IC.Image = CType(resources.GetObject("btnNeu_TCR_IC.Image"), System.Drawing.Image)
        Me.btnNeu_TCR_IC.Location = New System.Drawing.Point(345, 7)
        Me.btnNeu_TCR_IC.Name = "btnNeu_TCR_IC"
        Me.btnNeu_TCR_IC.Size = New System.Drawing.Size(54, 59)
        Me.btnNeu_TCR_IC.TabIndex = 70
        Me.btnNeu_TCR_IC.UseVisualStyleBackColor = True
        '
        'btnBrow_TCR_IC_src
        '
        Me.btnBrow_TCR_IC_src.Image = CType(resources.GetObject("btnBrow_TCR_IC_src.Image"), System.Drawing.Image)
        Me.btnBrow_TCR_IC_src.Location = New System.Drawing.Point(312, 9)
        Me.btnBrow_TCR_IC_src.Name = "btnBrow_TCR_IC_src"
        Me.btnBrow_TCR_IC_src.Size = New System.Drawing.Size(26, 23)
        Me.btnBrow_TCR_IC_src.TabIndex = 69
        Me.btnBrow_TCR_IC_src.UseVisualStyleBackColor = True
        '
        'txtTCR_IC_src
        '
        Me.txtTCR_IC_src.Location = New System.Drawing.Point(78, 10)
        Me.txtTCR_IC_src.Name = "txtTCR_IC_src"
        Me.txtTCR_IC_src.Size = New System.Drawing.Size(229, 20)
        Me.txtTCR_IC_src.TabIndex = 68
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(8, 15)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(41, 13)
        Me.Label13.TabIndex = 67
        Me.Label13.Text = "Source"
        '
        'OFD_TCR_IC
        '
        Me.OFD_TCR_IC.FileName = "Source File"
        Me.OFD_TCR_IC.Filter = "Excel files (*.xls,*xlsx)|*.xls;*.xlsx"
        '
        'SFD_TCR_IC
        '
        Me.SFD_TCR_IC.Filter = "Excel files (*.xls,*xlsx)|*.xls;*.xlsx"
        '
        'BWTCR_IC
        '
        '
        'frmTCR_IC
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.ClientSize = New System.Drawing.Size(406, 158)
        Me.Controls.Add(Me.PicBar_TCR_IC)
        Me.Controls.Add(Me.btnBrow_TCR_IC_dest)
        Me.Controls.Add(Me.txtTCR_IC_dest)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.btnNeu_TCR_IC)
        Me.Controls.Add(Me.btnBrow_TCR_IC_src)
        Me.Controls.Add(Me.txtTCR_IC_src)
        Me.Controls.Add(Me.Label13)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmTCR_IC"
        Me.Text = "ThroughPut Cabinet Report"
        CType(Me.PicBar_TCR_IC, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents PicBar_TCR_IC As System.Windows.Forms.PictureBox
    Friend WithEvents btnBrow_TCR_IC_dest As System.Windows.Forms.Button
    Friend WithEvents txtTCR_IC_dest As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents btnNeu_TCR_IC As System.Windows.Forms.Button
    Friend WithEvents btnBrow_TCR_IC_src As System.Windows.Forms.Button
    Friend WithEvents txtTCR_IC_src As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents OFD_TCR_IC As System.Windows.Forms.OpenFileDialog
    Friend WithEvents SFD_TCR_IC As System.Windows.Forms.SaveFileDialog
    Friend WithEvents BWTCR_IC As System.ComponentModel.BackgroundWorker
End Class
