<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSRDR
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmSRDR))
        Me.PicBar_SRDR = New System.Windows.Forms.PictureBox()
        Me.btnBrow_SRDR_dest = New System.Windows.Forms.Button()
        Me.txtSRDR_dest = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.btnNeu_SRDR = New System.Windows.Forms.Button()
        Me.btnBrow_SRDR_src = New System.Windows.Forms.Button()
        Me.txtSRDR_src = New System.Windows.Forms.TextBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.OFD_SRDR = New System.Windows.Forms.OpenFileDialog()
        Me.SFD_SRDR = New System.Windows.Forms.SaveFileDialog()
        Me.BWSRDR = New System.ComponentModel.BackgroundWorker()
        CType(Me.PicBar_SRDR, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PicBar_SRDR
        '
        Me.PicBar_SRDR.Image = CType(resources.GetObject("PicBar_SRDR.Image"), System.Drawing.Image)
        Me.PicBar_SRDR.Location = New System.Drawing.Point(8, 70)
        Me.PicBar_SRDR.Name = "PicBar_SRDR"
        Me.PicBar_SRDR.Size = New System.Drawing.Size(80, 80)
        Me.PicBar_SRDR.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PicBar_SRDR.TabIndex = 87
        Me.PicBar_SRDR.TabStop = False
        Me.PicBar_SRDR.Visible = False
        '
        'btnBrow_SRDR_dest
        '
        Me.btnBrow_SRDR_dest.Image = CType(resources.GetObject("btnBrow_SRDR_dest.Image"), System.Drawing.Image)
        Me.btnBrow_SRDR_dest.Location = New System.Drawing.Point(313, 42)
        Me.btnBrow_SRDR_dest.Name = "btnBrow_SRDR_dest"
        Me.btnBrow_SRDR_dest.Size = New System.Drawing.Size(26, 23)
        Me.btnBrow_SRDR_dest.TabIndex = 90
        Me.btnBrow_SRDR_dest.UseVisualStyleBackColor = True
        '
        'txtSRDR_dest
        '
        Me.txtSRDR_dest.Location = New System.Drawing.Point(79, 43)
        Me.txtSRDR_dest.Name = "txtSRDR_dest"
        Me.txtSRDR_dest.Size = New System.Drawing.Size(229, 20)
        Me.txtSRDR_dest.TabIndex = 89
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(8, 48)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(60, 13)
        Me.Label7.TabIndex = 88
        Me.Label7.Text = "Destination"
        '
        'btnNeu_SRDR
        '
        Me.btnNeu_SRDR.Enabled = False
        Me.btnNeu_SRDR.Image = CType(resources.GetObject("btnNeu_SRDR.Image"), System.Drawing.Image)
        Me.btnNeu_SRDR.Location = New System.Drawing.Point(346, 6)
        Me.btnNeu_SRDR.Name = "btnNeu_SRDR"
        Me.btnNeu_SRDR.Size = New System.Drawing.Size(54, 59)
        Me.btnNeu_SRDR.TabIndex = 86
        Me.btnNeu_SRDR.UseVisualStyleBackColor = True
        '
        'btnBrow_SRDR_src
        '
        Me.btnBrow_SRDR_src.Image = CType(resources.GetObject("btnBrow_SRDR_src.Image"), System.Drawing.Image)
        Me.btnBrow_SRDR_src.Location = New System.Drawing.Point(313, 8)
        Me.btnBrow_SRDR_src.Name = "btnBrow_SRDR_src"
        Me.btnBrow_SRDR_src.Size = New System.Drawing.Size(26, 23)
        Me.btnBrow_SRDR_src.TabIndex = 85
        Me.btnBrow_SRDR_src.UseVisualStyleBackColor = True
        '
        'txtSRDR_src
        '
        Me.txtSRDR_src.Location = New System.Drawing.Point(79, 9)
        Me.txtSRDR_src.Name = "txtSRDR_src"
        Me.txtSRDR_src.Size = New System.Drawing.Size(229, 20)
        Me.txtSRDR_src.TabIndex = 84
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(9, 14)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(41, 13)
        Me.Label13.TabIndex = 83
        Me.Label13.Text = "Source"
        '
        'OFD_SRDR
        '
        Me.OFD_SRDR.FileName = "Source File"
        Me.OFD_SRDR.Filter = "Excel files (*.xls,*xlsx)|*.xls;*.xlsx"
        '
        'SFD_SRDR
        '
        Me.SFD_SRDR.Filter = "Excel files (*.xls,*xlsx)|*.xls;*.xlsx"
        '
        'BWSRDR
        '
        '
        'frmSRDR
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.ClientSize = New System.Drawing.Size(409, 157)
        Me.Controls.Add(Me.PicBar_SRDR)
        Me.Controls.Add(Me.btnBrow_SRDR_dest)
        Me.Controls.Add(Me.txtSRDR_dest)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.btnNeu_SRDR)
        Me.Controls.Add(Me.btnBrow_SRDR_src)
        Me.Controls.Add(Me.txtSRDR_src)
        Me.Controls.Add(Me.Label13)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmSRDR"
        Me.Text = "Sales Return Detail Report"
        CType(Me.PicBar_SRDR, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents PicBar_SRDR As System.Windows.Forms.PictureBox
    Friend WithEvents btnBrow_SRDR_dest As System.Windows.Forms.Button
    Friend WithEvents txtSRDR_dest As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents btnNeu_SRDR As System.Windows.Forms.Button
    Friend WithEvents btnBrow_SRDR_src As System.Windows.Forms.Button
    Friend WithEvents txtSRDR_src As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents OFD_SRDR As System.Windows.Forms.OpenFileDialog
    Friend WithEvents SFD_SRDR As System.Windows.Forms.SaveFileDialog
    Friend WithEvents BWSRDR As System.ComponentModel.BackgroundWorker
End Class
