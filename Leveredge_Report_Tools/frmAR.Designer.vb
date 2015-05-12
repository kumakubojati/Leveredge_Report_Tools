<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmAR
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmAR))
        Me.PicBar_AR = New System.Windows.Forms.PictureBox()
        Me.btnBrow_AR_dest = New System.Windows.Forms.Button()
        Me.txtAR_dest = New System.Windows.Forms.TextBox()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.btnNeu_AR = New System.Windows.Forms.Button()
        Me.btnBrow_AR_src = New System.Windows.Forms.Button()
        Me.txtAR_src = New System.Windows.Forms.TextBox()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.OFD_AR = New System.Windows.Forms.OpenFileDialog()
        Me.SFD_AR = New System.Windows.Forms.SaveFileDialog()
        Me.BWAR = New System.ComponentModel.BackgroundWorker()
        CType(Me.PicBar_AR, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PicBar_AR
        '
        Me.PicBar_AR.Image = CType(resources.GetObject("PicBar_AR.Image"), System.Drawing.Image)
        Me.PicBar_AR.Location = New System.Drawing.Point(9, 70)
        Me.PicBar_AR.Name = "PicBar_AR"
        Me.PicBar_AR.Size = New System.Drawing.Size(80, 80)
        Me.PicBar_AR.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PicBar_AR.TabIndex = 33
        Me.PicBar_AR.TabStop = False
        Me.PicBar_AR.Visible = False
        '
        'btnBrow_AR_dest
        '
        Me.btnBrow_AR_dest.BackColor = System.Drawing.Color.Transparent
        Me.btnBrow_AR_dest.Image = CType(resources.GetObject("btnBrow_AR_dest.Image"), System.Drawing.Image)
        Me.btnBrow_AR_dest.Location = New System.Drawing.Point(316, 40)
        Me.btnBrow_AR_dest.Name = "btnBrow_AR_dest"
        Me.btnBrow_AR_dest.Size = New System.Drawing.Size(25, 23)
        Me.btnBrow_AR_dest.TabIndex = 32
        Me.btnBrow_AR_dest.UseVisualStyleBackColor = False
        '
        'txtAR_dest
        '
        Me.txtAR_dest.Location = New System.Drawing.Point(82, 41)
        Me.txtAR_dest.Name = "txtAR_dest"
        Me.txtAR_dest.Size = New System.Drawing.Size(229, 20)
        Me.txtAR_dest.TabIndex = 31
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(9, 46)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(60, 13)
        Me.Label14.TabIndex = 30
        Me.Label14.Text = "Destination"
        '
        'btnNeu_AR
        '
        Me.btnNeu_AR.Enabled = False
        Me.btnNeu_AR.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnNeu_AR.Image = CType(resources.GetObject("btnNeu_AR.Image"), System.Drawing.Image)
        Me.btnNeu_AR.Location = New System.Drawing.Point(347, 7)
        Me.btnNeu_AR.Name = "btnNeu_AR"
        Me.btnNeu_AR.Size = New System.Drawing.Size(55, 56)
        Me.btnNeu_AR.TabIndex = 29
        Me.btnNeu_AR.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnNeu_AR.UseVisualStyleBackColor = True
        '
        'btnBrow_AR_src
        '
        Me.btnBrow_AR_src.BackColor = System.Drawing.Color.Transparent
        Me.btnBrow_AR_src.Image = CType(resources.GetObject("btnBrow_AR_src.Image"), System.Drawing.Image)
        Me.btnBrow_AR_src.Location = New System.Drawing.Point(316, 6)
        Me.btnBrow_AR_src.Name = "btnBrow_AR_src"
        Me.btnBrow_AR_src.Size = New System.Drawing.Size(25, 23)
        Me.btnBrow_AR_src.TabIndex = 28
        Me.btnBrow_AR_src.UseVisualStyleBackColor = False
        '
        'txtAR_src
        '
        Me.txtAR_src.Location = New System.Drawing.Point(82, 7)
        Me.txtAR_src.Name = "txtAR_src"
        Me.txtAR_src.Size = New System.Drawing.Size(229, 20)
        Me.txtAR_src.TabIndex = 27
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(9, 10)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(41, 13)
        Me.Label15.TabIndex = 26
        Me.Label15.Text = "Source"
        '
        'OFD_AR
        '
        Me.OFD_AR.FileName = "Source File"
        Me.OFD_AR.Filter = "Excel files (*.xls,*xlsx)|*.xls;*.xlsx"
        '
        'SFD_AR
        '
        Me.SFD_AR.Filter = "Excel files (*.xls,*.xlsx)|*.xls;*xlsx"
        '
        'BWAR
        '
        '
        'frmAR
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.ClientSize = New System.Drawing.Size(406, 159)
        Me.Controls.Add(Me.PicBar_AR)
        Me.Controls.Add(Me.btnBrow_AR_dest)
        Me.Controls.Add(Me.txtAR_dest)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.btnNeu_AR)
        Me.Controls.Add(Me.btnBrow_AR_src)
        Me.Controls.Add(Me.txtAR_src)
        Me.Controls.Add(Me.Label15)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmAR"
        Me.Text = "AR Report"
        CType(Me.PicBar_AR, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents PicBar_AR As System.Windows.Forms.PictureBox
    Friend WithEvents btnBrow_AR_dest As System.Windows.Forms.Button
    Friend WithEvents txtAR_dest As System.Windows.Forms.TextBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents btnNeu_AR As System.Windows.Forms.Button
    Friend WithEvents btnBrow_AR_src As System.Windows.Forms.Button
    Friend WithEvents txtAR_src As System.Windows.Forms.TextBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents OFD_AR As System.Windows.Forms.OpenFileDialog
    Friend WithEvents SFD_AR As System.Windows.Forms.SaveFileDialog
    Friend WithEvents BWAR As System.ComponentModel.BackgroundWorker
End Class
