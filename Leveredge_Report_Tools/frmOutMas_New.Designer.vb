<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmOutMas_New
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmOutMas_New))
        Me.PicBar_OutMas = New System.Windows.Forms.PictureBox()
        Me.btnBrow_OutMas_dest = New System.Windows.Forms.Button()
        Me.txtOutMas_dest = New System.Windows.Forms.TextBox()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.btnNeu_OutMas = New System.Windows.Forms.Button()
        Me.btnBrow_OutMas_src = New System.Windows.Forms.Button()
        Me.txtOutMas_src = New System.Windows.Forms.TextBox()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.BWOUTMAS = New System.ComponentModel.BackgroundWorker()
        Me.OFD_OUTMAS = New System.Windows.Forms.OpenFileDialog()
        Me.SFD_OUTMAS = New System.Windows.Forms.SaveFileDialog()
        CType(Me.PicBar_OutMas, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PicBar_OutMas
        '
        Me.PicBar_OutMas.Image = CType(resources.GetObject("PicBar_OutMas.Image"), System.Drawing.Image)
        Me.PicBar_OutMas.Location = New System.Drawing.Point(9, 74)
        Me.PicBar_OutMas.Name = "PicBar_OutMas"
        Me.PicBar_OutMas.Size = New System.Drawing.Size(80, 80)
        Me.PicBar_OutMas.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PicBar_OutMas.TabIndex = 33
        Me.PicBar_OutMas.TabStop = False
        Me.PicBar_OutMas.Visible = False
        '
        'btnBrow_OutMas_dest
        '
        Me.btnBrow_OutMas_dest.BackColor = System.Drawing.Color.Transparent
        Me.btnBrow_OutMas_dest.Image = CType(resources.GetObject("btnBrow_OutMas_dest.Image"), System.Drawing.Image)
        Me.btnBrow_OutMas_dest.Location = New System.Drawing.Point(316, 44)
        Me.btnBrow_OutMas_dest.Name = "btnBrow_OutMas_dest"
        Me.btnBrow_OutMas_dest.Size = New System.Drawing.Size(25, 23)
        Me.btnBrow_OutMas_dest.TabIndex = 32
        Me.btnBrow_OutMas_dest.UseVisualStyleBackColor = False
        '
        'txtOutMas_dest
        '
        Me.txtOutMas_dest.Location = New System.Drawing.Point(82, 45)
        Me.txtOutMas_dest.Name = "txtOutMas_dest"
        Me.txtOutMas_dest.Size = New System.Drawing.Size(229, 20)
        Me.txtOutMas_dest.TabIndex = 31
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(9, 50)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(60, 13)
        Me.Label14.TabIndex = 30
        Me.Label14.Text = "Destination"
        '
        'btnNeu_OutMas
        '
        Me.btnNeu_OutMas.Enabled = False
        Me.btnNeu_OutMas.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnNeu_OutMas.Image = CType(resources.GetObject("btnNeu_OutMas.Image"), System.Drawing.Image)
        Me.btnNeu_OutMas.Location = New System.Drawing.Point(347, 11)
        Me.btnNeu_OutMas.Name = "btnNeu_OutMas"
        Me.btnNeu_OutMas.Size = New System.Drawing.Size(55, 56)
        Me.btnNeu_OutMas.TabIndex = 29
        Me.btnNeu_OutMas.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnNeu_OutMas.UseVisualStyleBackColor = True
        '
        'btnBrow_OutMas_src
        '
        Me.btnBrow_OutMas_src.BackColor = System.Drawing.Color.Transparent
        Me.btnBrow_OutMas_src.Image = CType(resources.GetObject("btnBrow_OutMas_src.Image"), System.Drawing.Image)
        Me.btnBrow_OutMas_src.Location = New System.Drawing.Point(316, 10)
        Me.btnBrow_OutMas_src.Name = "btnBrow_OutMas_src"
        Me.btnBrow_OutMas_src.Size = New System.Drawing.Size(25, 23)
        Me.btnBrow_OutMas_src.TabIndex = 28
        Me.btnBrow_OutMas_src.UseVisualStyleBackColor = False
        '
        'txtOutMas_src
        '
        Me.txtOutMas_src.Location = New System.Drawing.Point(82, 11)
        Me.txtOutMas_src.Name = "txtOutMas_src"
        Me.txtOutMas_src.Size = New System.Drawing.Size(229, 20)
        Me.txtOutMas_src.TabIndex = 27
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(9, 14)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(41, 13)
        Me.Label15.TabIndex = 26
        Me.Label15.Text = "Source"
        '
        'BWOUTMAS
        '
        '
        'OFD_OUTMAS
        '
        Me.OFD_OUTMAS.FileName = "Source File"
        Me.OFD_OUTMAS.Filter = "Excel files (*.xls,*xlsx)|*.xls;*.xlsx"
        '
        'SFD_OUTMAS
        '
        Me.SFD_OUTMAS.Filter = "Excel files (*.xls,*.xlsx)|*.xls;*xlsx"
        '
        'frmOutMas_New
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(410, 158)
        Me.Controls.Add(Me.PicBar_OutMas)
        Me.Controls.Add(Me.btnBrow_OutMas_dest)
        Me.Controls.Add(Me.txtOutMas_dest)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.btnNeu_OutMas)
        Me.Controls.Add(Me.btnBrow_OutMas_src)
        Me.Controls.Add(Me.txtOutMas_src)
        Me.Controls.Add(Me.Label15)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmOutMas_New"
        Me.Text = "Outlet Master"
        CType(Me.PicBar_OutMas, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents PicBar_OutMas As System.Windows.Forms.PictureBox
    Friend WithEvents btnBrow_OutMas_dest As System.Windows.Forms.Button
    Friend WithEvents txtOutMas_dest As System.Windows.Forms.TextBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents btnNeu_OutMas As System.Windows.Forms.Button
    Friend WithEvents btnBrow_OutMas_src As System.Windows.Forms.Button
    Friend WithEvents txtOutMas_src As System.Windows.Forms.TextBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents BWOUTMAS As System.ComponentModel.BackgroundWorker
    Friend WithEvents OFD_OUTMAS As System.Windows.Forms.OpenFileDialog
    Friend WithEvents SFD_OUTMAS As System.Windows.Forms.SaveFileDialog
End Class
