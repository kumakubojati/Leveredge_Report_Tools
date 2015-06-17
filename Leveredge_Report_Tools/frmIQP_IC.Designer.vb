<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmIQP_IC
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmIQP_IC))
        Me.PicBar_IQP_IC = New System.Windows.Forms.PictureBox()
        Me.btnBrow_IQP_IC_dest = New System.Windows.Forms.Button()
        Me.txtIQP_IC_dest = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.btnNeu_IQP_IC = New System.Windows.Forms.Button()
        Me.btnBrow_IQP_IC_src = New System.Windows.Forms.Button()
        Me.txtIQP_IC_src = New System.Windows.Forms.TextBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.OFD_IQP_IC = New System.Windows.Forms.OpenFileDialog()
        Me.SFD_IQP_IC = New System.Windows.Forms.SaveFileDialog()
        Me.BWIQP_IC = New System.ComponentModel.BackgroundWorker()
        CType(Me.PicBar_IQP_IC, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PicBar_IQP_IC
        '
        Me.PicBar_IQP_IC.Image = CType(resources.GetObject("PicBar_IQP_IC.Image"), System.Drawing.Image)
        Me.PicBar_IQP_IC.Location = New System.Drawing.Point(9, 74)
        Me.PicBar_IQP_IC.Name = "PicBar_IQP_IC"
        Me.PicBar_IQP_IC.Size = New System.Drawing.Size(80, 80)
        Me.PicBar_IQP_IC.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PicBar_IQP_IC.TabIndex = 78
        Me.PicBar_IQP_IC.TabStop = False
        Me.PicBar_IQP_IC.Visible = False
        '
        'btnBrow_IQP_IC_dest
        '
        Me.btnBrow_IQP_IC_dest.Image = CType(resources.GetObject("btnBrow_IQP_IC_dest.Image"), System.Drawing.Image)
        Me.btnBrow_IQP_IC_dest.Location = New System.Drawing.Point(314, 46)
        Me.btnBrow_IQP_IC_dest.Name = "btnBrow_IQP_IC_dest"
        Me.btnBrow_IQP_IC_dest.Size = New System.Drawing.Size(26, 23)
        Me.btnBrow_IQP_IC_dest.TabIndex = 81
        Me.btnBrow_IQP_IC_dest.UseVisualStyleBackColor = True
        '
        'txtIQP_IC_dest
        '
        Me.txtIQP_IC_dest.Location = New System.Drawing.Point(80, 47)
        Me.txtIQP_IC_dest.Name = "txtIQP_IC_dest"
        Me.txtIQP_IC_dest.Size = New System.Drawing.Size(229, 20)
        Me.txtIQP_IC_dest.TabIndex = 80
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(9, 52)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(60, 13)
        Me.Label7.TabIndex = 79
        Me.Label7.Text = "Destination"
        '
        'btnNeu_IQP_IC
        '
        Me.btnNeu_IQP_IC.Enabled = False
        Me.btnNeu_IQP_IC.Image = CType(resources.GetObject("btnNeu_IQP_IC.Image"), System.Drawing.Image)
        Me.btnNeu_IQP_IC.Location = New System.Drawing.Point(347, 10)
        Me.btnNeu_IQP_IC.Name = "btnNeu_IQP_IC"
        Me.btnNeu_IQP_IC.Size = New System.Drawing.Size(54, 59)
        Me.btnNeu_IQP_IC.TabIndex = 77
        Me.btnNeu_IQP_IC.UseVisualStyleBackColor = True
        '
        'btnBrow_IQP_IC_src
        '
        Me.btnBrow_IQP_IC_src.Image = CType(resources.GetObject("btnBrow_IQP_IC_src.Image"), System.Drawing.Image)
        Me.btnBrow_IQP_IC_src.Location = New System.Drawing.Point(314, 12)
        Me.btnBrow_IQP_IC_src.Name = "btnBrow_IQP_IC_src"
        Me.btnBrow_IQP_IC_src.Size = New System.Drawing.Size(26, 23)
        Me.btnBrow_IQP_IC_src.TabIndex = 76
        Me.btnBrow_IQP_IC_src.UseVisualStyleBackColor = True
        '
        'txtIQP_IC_src
        '
        Me.txtIQP_IC_src.Location = New System.Drawing.Point(80, 13)
        Me.txtIQP_IC_src.Name = "txtIQP_IC_src"
        Me.txtIQP_IC_src.Size = New System.Drawing.Size(229, 20)
        Me.txtIQP_IC_src.TabIndex = 75
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(10, 18)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(41, 13)
        Me.Label13.TabIndex = 74
        Me.Label13.Text = "Source"
        '
        'OFD_IQP_IC
        '
        Me.OFD_IQP_IC.FileName = "Source File"
        Me.OFD_IQP_IC.Filter = "Excel files (*.xls,*xlsx)|*.xls;*.xlsx"
        '
        'SFD_IQP_IC
        '
        Me.SFD_IQP_IC.Filter = "Excel files (*.xls,*.xlsx)|*.xls;*xlsx"
        '
        'BWIQP_IC
        '
        '
        'frmIQP_IC
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.ClientSize = New System.Drawing.Size(411, 159)
        Me.Controls.Add(Me.PicBar_IQP_IC)
        Me.Controls.Add(Me.btnBrow_IQP_IC_dest)
        Me.Controls.Add(Me.txtIQP_IC_dest)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.btnNeu_IQP_IC)
        Me.Controls.Add(Me.btnBrow_IQP_IC_src)
        Me.Controls.Add(Me.txtIQP_IC_src)
        Me.Controls.Add(Me.Label13)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmIQP_IC"
        Me.Text = "Ice Cream IQ Performance Report"
        CType(Me.PicBar_IQP_IC, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents PicBar_IQP_IC As System.Windows.Forms.PictureBox
    Friend WithEvents btnBrow_IQP_IC_dest As System.Windows.Forms.Button
    Friend WithEvents txtIQP_IC_dest As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents btnNeu_IQP_IC As System.Windows.Forms.Button
    Friend WithEvents btnBrow_IQP_IC_src As System.Windows.Forms.Button
    Friend WithEvents txtIQP_IC_src As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents OFD_IQP_IC As System.Windows.Forms.OpenFileDialog
    Friend WithEvents SFD_IQP_IC As System.Windows.Forms.SaveFileDialog
    Friend WithEvents BWIQP_IC As System.ComponentModel.BackgroundWorker
End Class
