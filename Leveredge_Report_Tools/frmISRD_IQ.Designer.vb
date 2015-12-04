<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmISRD_IQ
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmISRD_IQ))
        Me.PicBar_ISRD_IQ = New System.Windows.Forms.PictureBox()
        Me.btnBrow_ISRD_IQ_dest = New System.Windows.Forms.Button()
        Me.txtISRD_IQ_dest = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.btnNeu_ISRD_IQ = New System.Windows.Forms.Button()
        Me.btnBrow_ISRD_IQ_src = New System.Windows.Forms.Button()
        Me.txtISRD_IQ_src = New System.Windows.Forms.TextBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.OFD_ISRD_IQ = New System.Windows.Forms.OpenFileDialog()
        Me.SFD_ISRD_IQ = New System.Windows.Forms.SaveFileDialog()
        Me.BWISRD_IQ = New System.ComponentModel.BackgroundWorker()
        CType(Me.PicBar_ISRD_IQ, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PicBar_ISRD_IQ
        '
        Me.PicBar_ISRD_IQ.Image = CType(resources.GetObject("PicBar_ISRD_IQ.Image"), System.Drawing.Image)
        Me.PicBar_ISRD_IQ.Location = New System.Drawing.Point(11, 75)
        Me.PicBar_ISRD_IQ.Name = "PicBar_ISRD_IQ"
        Me.PicBar_ISRD_IQ.Size = New System.Drawing.Size(80, 80)
        Me.PicBar_ISRD_IQ.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PicBar_ISRD_IQ.TabIndex = 79
        Me.PicBar_ISRD_IQ.TabStop = False
        Me.PicBar_ISRD_IQ.Visible = False
        '
        'btnBrow_ISRD_IQ_dest
        '
        Me.btnBrow_ISRD_IQ_dest.Image = CType(resources.GetObject("btnBrow_ISRD_IQ_dest.Image"), System.Drawing.Image)
        Me.btnBrow_ISRD_IQ_dest.Location = New System.Drawing.Point(316, 47)
        Me.btnBrow_ISRD_IQ_dest.Name = "btnBrow_ISRD_IQ_dest"
        Me.btnBrow_ISRD_IQ_dest.Size = New System.Drawing.Size(26, 23)
        Me.btnBrow_ISRD_IQ_dest.TabIndex = 82
        Me.btnBrow_ISRD_IQ_dest.UseVisualStyleBackColor = True
        '
        'txtISRD_IQ_dest
        '
        Me.txtISRD_IQ_dest.Location = New System.Drawing.Point(82, 48)
        Me.txtISRD_IQ_dest.Name = "txtISRD_IQ_dest"
        Me.txtISRD_IQ_dest.Size = New System.Drawing.Size(229, 20)
        Me.txtISRD_IQ_dest.TabIndex = 81
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(11, 53)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(60, 13)
        Me.Label7.TabIndex = 80
        Me.Label7.Text = "Destination"
        '
        'btnNeu_ISRD_IQ
        '
        Me.btnNeu_ISRD_IQ.Enabled = False
        Me.btnNeu_ISRD_IQ.Image = CType(resources.GetObject("btnNeu_ISRD_IQ.Image"), System.Drawing.Image)
        Me.btnNeu_ISRD_IQ.Location = New System.Drawing.Point(349, 11)
        Me.btnNeu_ISRD_IQ.Name = "btnNeu_ISRD_IQ"
        Me.btnNeu_ISRD_IQ.Size = New System.Drawing.Size(54, 59)
        Me.btnNeu_ISRD_IQ.TabIndex = 78
        Me.btnNeu_ISRD_IQ.UseVisualStyleBackColor = True
        '
        'btnBrow_ISRD_IQ_src
        '
        Me.btnBrow_ISRD_IQ_src.Image = CType(resources.GetObject("btnBrow_ISRD_IQ_src.Image"), System.Drawing.Image)
        Me.btnBrow_ISRD_IQ_src.Location = New System.Drawing.Point(316, 13)
        Me.btnBrow_ISRD_IQ_src.Name = "btnBrow_ISRD_IQ_src"
        Me.btnBrow_ISRD_IQ_src.Size = New System.Drawing.Size(26, 23)
        Me.btnBrow_ISRD_IQ_src.TabIndex = 77
        Me.btnBrow_ISRD_IQ_src.UseVisualStyleBackColor = True
        '
        'txtISRD_IQ_src
        '
        Me.txtISRD_IQ_src.Location = New System.Drawing.Point(82, 14)
        Me.txtISRD_IQ_src.Name = "txtISRD_IQ_src"
        Me.txtISRD_IQ_src.Size = New System.Drawing.Size(229, 20)
        Me.txtISRD_IQ_src.TabIndex = 76
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(12, 19)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(41, 13)
        Me.Label13.TabIndex = 75
        Me.Label13.Text = "Source"
        '
        'OFD_ISRD_IQ
        '
        Me.OFD_ISRD_IQ.FileName = "Source File"
        Me.OFD_ISRD_IQ.Filter = "Excel files (*.xls,*xlsx)|*.xls;*.xlsx"
        '
        'SFD_ISRD_IQ
        '
        Me.SFD_ISRD_IQ.Filter = "Excel files (*.xls,*xlsx)|*.xls;*.xlsx"
        '
        'BWISRD_IQ
        '
        '
        'frmISRD_IQ
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(409, 164)
        Me.Controls.Add(Me.PicBar_ISRD_IQ)
        Me.Controls.Add(Me.btnBrow_ISRD_IQ_dest)
        Me.Controls.Add(Me.txtISRD_IQ_dest)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.btnNeu_ISRD_IQ)
        Me.Controls.Add(Me.btnBrow_ISRD_IQ_src)
        Me.Controls.Add(Me.txtISRD_IQ_src)
        Me.Controls.Add(Me.Label13)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmISRD_IQ"
        Me.Text = "IQ Summary Report For DSR"
        CType(Me.PicBar_ISRD_IQ, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents PicBar_ISRD_IQ As System.Windows.Forms.PictureBox
    Friend WithEvents btnBrow_ISRD_IQ_dest As System.Windows.Forms.Button
    Friend WithEvents txtISRD_IQ_dest As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents btnNeu_ISRD_IQ As System.Windows.Forms.Button
    Friend WithEvents btnBrow_ISRD_IQ_src As System.Windows.Forms.Button
    Friend WithEvents txtISRD_IQ_src As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents OFD_ISRD_IQ As System.Windows.Forms.OpenFileDialog
    Friend WithEvents SFD_ISRD_IQ As System.Windows.Forms.SaveFileDialog
    Friend WithEvents BWISRD_IQ As System.ComponentModel.BackgroundWorker
End Class
