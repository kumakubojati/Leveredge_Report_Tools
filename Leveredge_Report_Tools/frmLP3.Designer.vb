<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmLP3
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmLP3))
        Me.GBLP3 = New System.Windows.Forms.GroupBox()
        Me.RBQty = New System.Windows.Forms.RadioButton()
        Me.RBRupiah = New System.Windows.Forms.RadioButton()
        Me.RBAll = New System.Windows.Forms.RadioButton()
        Me.PicBarLP3 = New System.Windows.Forms.PictureBox()
        Me.btnBrow_LP3_Dest = New System.Windows.Forms.Button()
        Me.txtLP3_dest = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.btnNeuLP3 = New System.Windows.Forms.Button()
        Me.btnBrow_LP3_src = New System.Windows.Forms.Button()
        Me.txtLP3_src = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.OFD_LP3 = New System.Windows.Forms.OpenFileDialog()
        Me.SFD_LP3 = New System.Windows.Forms.SaveFileDialog()
        Me.BWLP3 = New System.ComponentModel.BackgroundWorker()
        Me.GBLP3.SuspendLayout()
        CType(Me.PicBarLP3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GBLP3
        '
        Me.GBLP3.Controls.Add(Me.RBQty)
        Me.GBLP3.Controls.Add(Me.RBRupiah)
        Me.GBLP3.Controls.Add(Me.RBAll)
        Me.GBLP3.Location = New System.Drawing.Point(4, 3)
        Me.GBLP3.Name = "GBLP3"
        Me.GBLP3.Size = New System.Drawing.Size(251, 53)
        Me.GBLP3.TabIndex = 22
        Me.GBLP3.TabStop = False
        Me.GBLP3.Text = "Report Type"
        '
        'RBQty
        '
        Me.RBQty.AutoSize = True
        Me.RBQty.Location = New System.Drawing.Point(146, 22)
        Me.RBQty.Name = "RBQty"
        Me.RBQty.Size = New System.Drawing.Size(96, 17)
        Me.RBQty.TabIndex = 2
        Me.RBQty.TabStop = True
        Me.RBQty.Text = "Detail Per SKU"
        Me.RBQty.UseVisualStyleBackColor = True
        '
        'RBRupiah
        '
        Me.RBRupiah.AutoSize = True
        Me.RBRupiah.Location = New System.Drawing.Point(67, 22)
        Me.RBRupiah.Name = "RBRupiah"
        Me.RBRupiah.Size = New System.Drawing.Size(59, 17)
        Me.RBRupiah.TabIndex = 1
        Me.RBRupiah.TabStop = True
        Me.RBRupiah.Text = "Rupiah"
        Me.RBRupiah.UseVisualStyleBackColor = True
        '
        'RBAll
        '
        Me.RBAll.AutoSize = True
        Me.RBAll.Location = New System.Drawing.Point(7, 22)
        Me.RBAll.Name = "RBAll"
        Me.RBAll.Size = New System.Drawing.Size(36, 17)
        Me.RBAll.TabIndex = 0
        Me.RBAll.TabStop = True
        Me.RBAll.Text = "All"
        Me.RBAll.UseVisualStyleBackColor = True
        '
        'PicBarLP3
        '
        Me.PicBarLP3.Image = CType(resources.GetObject("PicBarLP3.Image"), System.Drawing.Image)
        Me.PicBarLP3.Location = New System.Drawing.Point(5, 122)
        Me.PicBarLP3.Name = "PicBarLP3"
        Me.PicBarLP3.Size = New System.Drawing.Size(80, 80)
        Me.PicBarLP3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PicBarLP3.TabIndex = 18
        Me.PicBarLP3.TabStop = False
        Me.PicBarLP3.Visible = False
        '
        'btnBrow_LP3_Dest
        '
        Me.btnBrow_LP3_Dest.Image = CType(resources.GetObject("btnBrow_LP3_Dest.Image"), System.Drawing.Image)
        Me.btnBrow_LP3_Dest.Location = New System.Drawing.Point(310, 94)
        Me.btnBrow_LP3_Dest.Name = "btnBrow_LP3_Dest"
        Me.btnBrow_LP3_Dest.Size = New System.Drawing.Size(26, 23)
        Me.btnBrow_LP3_Dest.TabIndex = 21
        Me.btnBrow_LP3_Dest.UseVisualStyleBackColor = True
        '
        'txtLP3_dest
        '
        Me.txtLP3_dest.Location = New System.Drawing.Point(76, 95)
        Me.txtLP3_dest.Name = "txtLP3_dest"
        Me.txtLP3_dest.Size = New System.Drawing.Size(229, 20)
        Me.txtLP3_dest.TabIndex = 20
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(3, 100)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(60, 13)
        Me.Label6.TabIndex = 19
        Me.Label6.Text = "Destination"
        '
        'btnNeuLP3
        '
        Me.btnNeuLP3.Enabled = False
        Me.btnNeuLP3.Image = CType(resources.GetObject("btnNeuLP3.Image"), System.Drawing.Image)
        Me.btnNeuLP3.Location = New System.Drawing.Point(343, 58)
        Me.btnNeuLP3.Name = "btnNeuLP3"
        Me.btnNeuLP3.Size = New System.Drawing.Size(54, 59)
        Me.btnNeuLP3.TabIndex = 17
        Me.btnNeuLP3.UseVisualStyleBackColor = True
        '
        'btnBrow_LP3_src
        '
        Me.btnBrow_LP3_src.Image = CType(resources.GetObject("btnBrow_LP3_src.Image"), System.Drawing.Image)
        Me.btnBrow_LP3_src.Location = New System.Drawing.Point(310, 60)
        Me.btnBrow_LP3_src.Name = "btnBrow_LP3_src"
        Me.btnBrow_LP3_src.Size = New System.Drawing.Size(26, 23)
        Me.btnBrow_LP3_src.TabIndex = 16
        Me.btnBrow_LP3_src.UseVisualStyleBackColor = True
        '
        'txtLP3_src
        '
        Me.txtLP3_src.Location = New System.Drawing.Point(76, 61)
        Me.txtLP3_src.Name = "txtLP3_src"
        Me.txtLP3_src.Size = New System.Drawing.Size(229, 20)
        Me.txtLP3_src.TabIndex = 15
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(6, 66)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(41, 13)
        Me.Label5.TabIndex = 14
        Me.Label5.Text = "Source"
        '
        'OFD_LP3
        '
        Me.OFD_LP3.FileName = "Source File"
        Me.OFD_LP3.Filter = "Excel files (*.xls,*xlsx)|*.xls;*.xlsx"
        '
        'SFD_LP3
        '
        Me.SFD_LP3.Filter = "Excel files (*.xls,*.xlsx)|*.xls;*xlsx"
        '
        'BWLP3
        '
        '
        'frmLP3
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.ClientSize = New System.Drawing.Size(397, 206)
        Me.Controls.Add(Me.GBLP3)
        Me.Controls.Add(Me.PicBarLP3)
        Me.Controls.Add(Me.btnBrow_LP3_Dest)
        Me.Controls.Add(Me.txtLP3_dest)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.btnNeuLP3)
        Me.Controls.Add(Me.btnBrow_LP3_src)
        Me.Controls.Add(Me.txtLP3_src)
        Me.Controls.Add(Me.Label5)
        Me.Name = "frmLP3"
        Me.Text = "Weekly Stock And Sales (LP3)"
        Me.GBLP3.ResumeLayout(False)
        Me.GBLP3.PerformLayout()
        CType(Me.PicBarLP3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GBLP3 As System.Windows.Forms.GroupBox
    Friend WithEvents RBQty As System.Windows.Forms.RadioButton
    Friend WithEvents RBRupiah As System.Windows.Forms.RadioButton
    Friend WithEvents RBAll As System.Windows.Forms.RadioButton
    Friend WithEvents PicBarLP3 As System.Windows.Forms.PictureBox
    Friend WithEvents btnBrow_LP3_Dest As System.Windows.Forms.Button
    Friend WithEvents txtLP3_dest As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents btnNeuLP3 As System.Windows.Forms.Button
    Friend WithEvents btnBrow_LP3_src As System.Windows.Forms.Button
    Friend WithEvents txtLP3_src As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents OFD_LP3 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents SFD_LP3 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents BWLP3 As System.ComponentModel.BackgroundWorker
End Class
