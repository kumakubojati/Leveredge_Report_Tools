<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmIQS_IC
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmIQS_IC))
        Me.PicBar_IQS_IC = New System.Windows.Forms.PictureBox()
        Me.btnBrow_IQS_IC_dest = New System.Windows.Forms.Button()
        Me.txtIQS_IC_dest = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.btnNeu_IQS_IC = New System.Windows.Forms.Button()
        Me.btnBrow_IQS_IC_src = New System.Windows.Forms.Button()
        Me.txtIQS_IC_src = New System.Windows.Forms.TextBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.OFD_IQS_IC = New System.Windows.Forms.OpenFileDialog()
        Me.SFD_IQS_IC = New System.Windows.Forms.SaveFileDialog()
        Me.BWIQS_IC = New System.ComponentModel.BackgroundWorker()
        CType(Me.PicBar_IQS_IC, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PicBar_IQS_IC
        '
        Me.PicBar_IQS_IC.Image = CType(resources.GetObject("PicBar_IQS_IC.Image"), System.Drawing.Image)
        Me.PicBar_IQS_IC.Location = New System.Drawing.Point(7, 73)
        Me.PicBar_IQS_IC.Name = "PicBar_IQS_IC"
        Me.PicBar_IQS_IC.Size = New System.Drawing.Size(80, 80)
        Me.PicBar_IQS_IC.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PicBar_IQS_IC.TabIndex = 86
        Me.PicBar_IQS_IC.TabStop = False
        Me.PicBar_IQS_IC.Visible = False
        '
        'btnBrow_IQS_IC_dest
        '
        Me.btnBrow_IQS_IC_dest.Image = CType(resources.GetObject("btnBrow_IQS_IC_dest.Image"), System.Drawing.Image)
        Me.btnBrow_IQS_IC_dest.Location = New System.Drawing.Point(312, 45)
        Me.btnBrow_IQS_IC_dest.Name = "btnBrow_IQS_IC_dest"
        Me.btnBrow_IQS_IC_dest.Size = New System.Drawing.Size(26, 23)
        Me.btnBrow_IQS_IC_dest.TabIndex = 89
        Me.btnBrow_IQS_IC_dest.UseVisualStyleBackColor = True
        '
        'txtIQS_IC_dest
        '
        Me.txtIQS_IC_dest.Location = New System.Drawing.Point(78, 46)
        Me.txtIQS_IC_dest.Name = "txtIQS_IC_dest"
        Me.txtIQS_IC_dest.Size = New System.Drawing.Size(229, 20)
        Me.txtIQS_IC_dest.TabIndex = 88
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(7, 51)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(60, 13)
        Me.Label7.TabIndex = 87
        Me.Label7.Text = "Destination"
        '
        'btnNeu_IQS_IC
        '
        Me.btnNeu_IQS_IC.Enabled = False
        Me.btnNeu_IQS_IC.Image = CType(resources.GetObject("btnNeu_IQS_IC.Image"), System.Drawing.Image)
        Me.btnNeu_IQS_IC.Location = New System.Drawing.Point(345, 9)
        Me.btnNeu_IQS_IC.Name = "btnNeu_IQS_IC"
        Me.btnNeu_IQS_IC.Size = New System.Drawing.Size(54, 59)
        Me.btnNeu_IQS_IC.TabIndex = 85
        Me.btnNeu_IQS_IC.UseVisualStyleBackColor = True
        '
        'btnBrow_IQS_IC_src
        '
        Me.btnBrow_IQS_IC_src.Image = CType(resources.GetObject("btnBrow_IQS_IC_src.Image"), System.Drawing.Image)
        Me.btnBrow_IQS_IC_src.Location = New System.Drawing.Point(312, 11)
        Me.btnBrow_IQS_IC_src.Name = "btnBrow_IQS_IC_src"
        Me.btnBrow_IQS_IC_src.Size = New System.Drawing.Size(26, 23)
        Me.btnBrow_IQS_IC_src.TabIndex = 84
        Me.btnBrow_IQS_IC_src.UseVisualStyleBackColor = True
        '
        'txtIQS_IC_src
        '
        Me.txtIQS_IC_src.Location = New System.Drawing.Point(78, 12)
        Me.txtIQS_IC_src.Name = "txtIQS_IC_src"
        Me.txtIQS_IC_src.Size = New System.Drawing.Size(229, 20)
        Me.txtIQS_IC_src.TabIndex = 83
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(8, 17)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(41, 13)
        Me.Label13.TabIndex = 82
        Me.Label13.Text = "Source"
        '
        'OFD_IQS_IC
        '
        Me.OFD_IQS_IC.FileName = "Source File"
        Me.OFD_IQS_IC.Filter = "Excel files (*.xls,*xlsx)|*.xls;*.xlsx"
        '
        'SFD_IQS_IC
        '
        Me.SFD_IQS_IC.Filter = "Excel files (*.xls,*.xlsx)|*.xls;*xlsx"
        '
        'BWIQS_IC
        '
        '
        'frmIQS_IC
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.ClientSize = New System.Drawing.Size(404, 156)
        Me.Controls.Add(Me.PicBar_IQS_IC)
        Me.Controls.Add(Me.btnBrow_IQS_IC_dest)
        Me.Controls.Add(Me.txtIQS_IC_dest)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.btnNeu_IQS_IC)
        Me.Controls.Add(Me.btnBrow_IQS_IC_src)
        Me.Controls.Add(Me.txtIQS_IC_src)
        Me.Controls.Add(Me.Label13)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmIQS_IC"
        Me.Text = "Ice Cream IQ Summary Report"
        CType(Me.PicBar_IQS_IC, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents PicBar_IQS_IC As System.Windows.Forms.PictureBox
    Friend WithEvents btnBrow_IQS_IC_dest As System.Windows.Forms.Button
    Friend WithEvents txtIQS_IC_dest As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents btnNeu_IQS_IC As System.Windows.Forms.Button
    Friend WithEvents btnBrow_IQS_IC_src As System.Windows.Forms.Button
    Friend WithEvents txtIQS_IC_src As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents OFD_IQS_IC As System.Windows.Forms.OpenFileDialog
    Friend WithEvents SFD_IQS_IC As System.Windows.Forms.SaveFileDialog
    Friend WithEvents BWIQS_IC As System.ComponentModel.BackgroundWorker
End Class
