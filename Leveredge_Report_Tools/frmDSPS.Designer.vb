<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmDSPS
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmDSPS))
        Me.PicBar_DSPS = New System.Windows.Forms.PictureBox()
        Me.btnBrow_DSPS_dest = New System.Windows.Forms.Button()
        Me.txtDSPS_dest = New System.Windows.Forms.TextBox()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.btnNeu_DSPS = New System.Windows.Forms.Button()
        Me.btnBrow_DSPS_src = New System.Windows.Forms.Button()
        Me.txtDSPS_src = New System.Windows.Forms.TextBox()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.OFD_DSPS = New System.Windows.Forms.OpenFileDialog()
        Me.SFD_DSPS = New System.Windows.Forms.SaveFileDialog()
        Me.BWDSPS = New System.ComponentModel.BackgroundWorker()
        CType(Me.PicBar_DSPS, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PicBar_DSPS
        '
        Me.PicBar_DSPS.Image = CType(resources.GetObject("PicBar_DSPS.Image"), System.Drawing.Image)
        Me.PicBar_DSPS.Location = New System.Drawing.Point(7, 73)
        Me.PicBar_DSPS.Name = "PicBar_DSPS"
        Me.PicBar_DSPS.Size = New System.Drawing.Size(80, 80)
        Me.PicBar_DSPS.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PicBar_DSPS.TabIndex = 41
        Me.PicBar_DSPS.TabStop = False
        Me.PicBar_DSPS.Visible = False
        '
        'btnBrow_DSPS_dest
        '
        Me.btnBrow_DSPS_dest.BackColor = System.Drawing.Color.Transparent
        Me.btnBrow_DSPS_dest.Image = CType(resources.GetObject("btnBrow_DSPS_dest.Image"), System.Drawing.Image)
        Me.btnBrow_DSPS_dest.Location = New System.Drawing.Point(314, 43)
        Me.btnBrow_DSPS_dest.Name = "btnBrow_DSPS_dest"
        Me.btnBrow_DSPS_dest.Size = New System.Drawing.Size(25, 23)
        Me.btnBrow_DSPS_dest.TabIndex = 40
        Me.btnBrow_DSPS_dest.UseVisualStyleBackColor = False
        '
        'txtDSPS_dest
        '
        Me.txtDSPS_dest.Location = New System.Drawing.Point(80, 44)
        Me.txtDSPS_dest.Name = "txtDSPS_dest"
        Me.txtDSPS_dest.Size = New System.Drawing.Size(229, 20)
        Me.txtDSPS_dest.TabIndex = 39
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(7, 49)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(60, 13)
        Me.Label14.TabIndex = 38
        Me.Label14.Text = "Destination"
        '
        'btnNeu_DSPS
        '
        Me.btnNeu_DSPS.Enabled = False
        Me.btnNeu_DSPS.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnNeu_DSPS.Image = CType(resources.GetObject("btnNeu_DSPS.Image"), System.Drawing.Image)
        Me.btnNeu_DSPS.Location = New System.Drawing.Point(345, 10)
        Me.btnNeu_DSPS.Name = "btnNeu_DSPS"
        Me.btnNeu_DSPS.Size = New System.Drawing.Size(55, 56)
        Me.btnNeu_DSPS.TabIndex = 37
        Me.btnNeu_DSPS.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnNeu_DSPS.UseVisualStyleBackColor = True
        '
        'btnBrow_DSPS_src
        '
        Me.btnBrow_DSPS_src.BackColor = System.Drawing.Color.Transparent
        Me.btnBrow_DSPS_src.Image = CType(resources.GetObject("btnBrow_DSPS_src.Image"), System.Drawing.Image)
        Me.btnBrow_DSPS_src.Location = New System.Drawing.Point(314, 9)
        Me.btnBrow_DSPS_src.Name = "btnBrow_DSPS_src"
        Me.btnBrow_DSPS_src.Size = New System.Drawing.Size(25, 23)
        Me.btnBrow_DSPS_src.TabIndex = 36
        Me.btnBrow_DSPS_src.UseVisualStyleBackColor = False
        '
        'txtDSPS_src
        '
        Me.txtDSPS_src.Location = New System.Drawing.Point(80, 10)
        Me.txtDSPS_src.Name = "txtDSPS_src"
        Me.txtDSPS_src.Size = New System.Drawing.Size(229, 20)
        Me.txtDSPS_src.TabIndex = 35
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(7, 13)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(41, 13)
        Me.Label15.TabIndex = 34
        Me.Label15.Text = "Source"
        '
        'OFD_DSPS
        '
        Me.OFD_DSPS.FileName = "Source File"
        Me.OFD_DSPS.Filter = "Excel files (*.xls,*xlsx)|*.xls;*.xlsx"
        '
        'SFD_DSPS
        '
        Me.SFD_DSPS.Filter = "Excel files (*.xls,*.xlsx)|*.xls;*xlsx"
        '
        'BWDSPS
        '
        '
        'frmDSPS
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.ClientSize = New System.Drawing.Size(407, 157)
        Me.Controls.Add(Me.PicBar_DSPS)
        Me.Controls.Add(Me.btnBrow_DSPS_dest)
        Me.Controls.Add(Me.txtDSPS_dest)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.btnNeu_DSPS)
        Me.Controls.Add(Me.btnBrow_DSPS_src)
        Me.Controls.Add(Me.txtDSPS_src)
        Me.Controls.Add(Me.Label15)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmDSPS"
        Me.Text = "Daily Sales And Payment Summary Report"
        CType(Me.PicBar_DSPS, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents PicBar_DSPS As System.Windows.Forms.PictureBox
    Friend WithEvents btnBrow_DSPS_dest As System.Windows.Forms.Button
    Friend WithEvents txtDSPS_dest As System.Windows.Forms.TextBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents btnNeu_DSPS As System.Windows.Forms.Button
    Friend WithEvents btnBrow_DSPS_src As System.Windows.Forms.Button
    Friend WithEvents txtDSPS_src As System.Windows.Forms.TextBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents OFD_DSPS As System.Windows.Forms.OpenFileDialog
    Friend WithEvents SFD_DSPS As System.Windows.Forms.SaveFileDialog
    Friend WithEvents BWDSPS As System.ComponentModel.BackgroundWorker
End Class
