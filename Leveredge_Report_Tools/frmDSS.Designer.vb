<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmDSS
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmDSS))
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.rbDSSprod = New System.Windows.Forms.RadioButton()
        Me.rbDSSrupiah = New System.Windows.Forms.RadioButton()
        Me.rbDSSall = New System.Windows.Forms.RadioButton()
        Me.PicBar_DSS = New System.Windows.Forms.PictureBox()
        Me.btnBrowDSS_dest = New System.Windows.Forms.Button()
        Me.txtDSS_dest = New System.Windows.Forms.TextBox()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.btnNeu_DSS = New System.Windows.Forms.Button()
        Me.bntBrowDSS_src = New System.Windows.Forms.Button()
        Me.txtDSS_src = New System.Windows.Forms.TextBox()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.BWDSS = New System.ComponentModel.BackgroundWorker()
        Me.OFD_DSS = New System.Windows.Forms.OpenFileDialog()
        Me.SFD_DSS = New System.Windows.Forms.SaveFileDialog()
        Me.GroupBox2.SuspendLayout()
        CType(Me.PicBar_DSS, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.rbDSSprod)
        Me.GroupBox2.Controls.Add(Me.rbDSSrupiah)
        Me.GroupBox2.Controls.Add(Me.rbDSSall)
        Me.GroupBox2.Location = New System.Drawing.Point(3, 3)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(251, 53)
        Me.GroupBox2.TabIndex = 31
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Report Type"
        '
        'rbDSSprod
        '
        Me.rbDSSprod.AutoSize = True
        Me.rbDSSprod.Location = New System.Drawing.Point(146, 22)
        Me.rbDSSprod.Name = "rbDSSprod"
        Me.rbDSSprod.Size = New System.Drawing.Size(59, 17)
        Me.rbDSSprod.TabIndex = 2
        Me.rbDSSprod.TabStop = True
        Me.rbDSSprod.Text = "Produk"
        Me.rbDSSprod.UseVisualStyleBackColor = True
        '
        'rbDSSrupiah
        '
        Me.rbDSSrupiah.AutoSize = True
        Me.rbDSSrupiah.Location = New System.Drawing.Point(67, 22)
        Me.rbDSSrupiah.Name = "rbDSSrupiah"
        Me.rbDSSrupiah.Size = New System.Drawing.Size(59, 17)
        Me.rbDSSrupiah.TabIndex = 1
        Me.rbDSSrupiah.TabStop = True
        Me.rbDSSrupiah.Text = "Rupiah"
        Me.rbDSSrupiah.UseVisualStyleBackColor = True
        '
        'rbDSSall
        '
        Me.rbDSSall.AutoSize = True
        Me.rbDSSall.Checked = True
        Me.rbDSSall.Location = New System.Drawing.Point(7, 22)
        Me.rbDSSall.Name = "rbDSSall"
        Me.rbDSSall.Size = New System.Drawing.Size(36, 17)
        Me.rbDSSall.TabIndex = 0
        Me.rbDSSall.TabStop = True
        Me.rbDSSall.Text = "All"
        Me.rbDSSall.UseVisualStyleBackColor = True
        '
        'PicBar_DSS
        '
        Me.PicBar_DSS.Image = CType(resources.GetObject("PicBar_DSS.Image"), System.Drawing.Image)
        Me.PicBar_DSS.Location = New System.Drawing.Point(4, 122)
        Me.PicBar_DSS.Name = "PicBar_DSS"
        Me.PicBar_DSS.Size = New System.Drawing.Size(80, 80)
        Me.PicBar_DSS.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PicBar_DSS.TabIndex = 27
        Me.PicBar_DSS.TabStop = False
        Me.PicBar_DSS.Visible = False
        '
        'btnBrowDSS_dest
        '
        Me.btnBrowDSS_dest.Image = CType(resources.GetObject("btnBrowDSS_dest.Image"), System.Drawing.Image)
        Me.btnBrowDSS_dest.Location = New System.Drawing.Point(309, 94)
        Me.btnBrowDSS_dest.Name = "btnBrowDSS_dest"
        Me.btnBrowDSS_dest.Size = New System.Drawing.Size(26, 23)
        Me.btnBrowDSS_dest.TabIndex = 30
        Me.btnBrowDSS_dest.UseVisualStyleBackColor = True
        '
        'txtDSS_dest
        '
        Me.txtDSS_dest.Location = New System.Drawing.Point(75, 95)
        Me.txtDSS_dest.Name = "txtDSS_dest"
        Me.txtDSS_dest.Size = New System.Drawing.Size(229, 20)
        Me.txtDSS_dest.TabIndex = 29
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Location = New System.Drawing.Point(4, 100)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(60, 13)
        Me.Label19.TabIndex = 28
        Me.Label19.Text = "Destination"
        '
        'btnNeu_DSS
        '
        Me.btnNeu_DSS.Enabled = False
        Me.btnNeu_DSS.Image = CType(resources.GetObject("btnNeu_DSS.Image"), System.Drawing.Image)
        Me.btnNeu_DSS.Location = New System.Drawing.Point(342, 58)
        Me.btnNeu_DSS.Name = "btnNeu_DSS"
        Me.btnNeu_DSS.Size = New System.Drawing.Size(54, 59)
        Me.btnNeu_DSS.TabIndex = 26
        Me.btnNeu_DSS.UseVisualStyleBackColor = True
        '
        'bntBrowDSS_src
        '
        Me.bntBrowDSS_src.Image = CType(resources.GetObject("bntBrowDSS_src.Image"), System.Drawing.Image)
        Me.bntBrowDSS_src.Location = New System.Drawing.Point(309, 60)
        Me.bntBrowDSS_src.Name = "bntBrowDSS_src"
        Me.bntBrowDSS_src.Size = New System.Drawing.Size(26, 23)
        Me.bntBrowDSS_src.TabIndex = 25
        Me.bntBrowDSS_src.UseVisualStyleBackColor = True
        '
        'txtDSS_src
        '
        Me.txtDSS_src.Location = New System.Drawing.Point(75, 61)
        Me.txtDSS_src.Name = "txtDSS_src"
        Me.txtDSS_src.Size = New System.Drawing.Size(229, 20)
        Me.txtDSS_src.TabIndex = 24
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.Location = New System.Drawing.Point(5, 66)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(41, 13)
        Me.Label20.TabIndex = 23
        Me.Label20.Text = "Source"
        '
        'BWDSS
        '
        '
        'OFD_DSS
        '
        Me.OFD_DSS.FileName = "Source File"
        Me.OFD_DSS.Filter = "Excel files (*.xls,*xlsx)|*.xls;*.xlsx"
        '
        'SFD_DSS
        '
        Me.SFD_DSS.Filter = "Excel files (*.xls,*.xlsx)|*.xls;*xlsx"
        '
        'frmDSS
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.ClientSize = New System.Drawing.Size(406, 206)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.PicBar_DSS)
        Me.Controls.Add(Me.btnBrowDSS_dest)
        Me.Controls.Add(Me.txtDSS_dest)
        Me.Controls.Add(Me.Label19)
        Me.Controls.Add(Me.btnNeu_DSS)
        Me.Controls.Add(Me.bntBrowDSS_src)
        Me.Controls.Add(Me.txtDSS_src)
        Me.Controls.Add(Me.Label20)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmDSS"
        Me.Text = "Daily Sales Summary"
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.PicBar_DSS, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents rbDSSprod As System.Windows.Forms.RadioButton
    Friend WithEvents rbDSSrupiah As System.Windows.Forms.RadioButton
    Friend WithEvents rbDSSall As System.Windows.Forms.RadioButton
    Friend WithEvents PicBar_DSS As System.Windows.Forms.PictureBox
    Friend WithEvents btnBrowDSS_dest As System.Windows.Forms.Button
    Friend WithEvents txtDSS_dest As System.Windows.Forms.TextBox
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents btnNeu_DSS As System.Windows.Forms.Button
    Friend WithEvents bntBrowDSS_src As System.Windows.Forms.Button
    Friend WithEvents txtDSS_src As System.Windows.Forms.TextBox
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents BWDSS As System.ComponentModel.BackgroundWorker
    Friend WithEvents OFD_DSS As System.Windows.Forms.OpenFileDialog
    Friend WithEvents SFD_DSS As System.Windows.Forms.SaveFileDialog
End Class
