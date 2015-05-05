<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmDSM
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmDSM))
        Me.PicBar_DSM = New System.Windows.Forms.PictureBox()
        Me.btnBrow_DSM_dest = New System.Windows.Forms.Button()
        Me.txtDSM_dest = New System.Windows.Forms.TextBox()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.btnNeu_DSM = New System.Windows.Forms.Button()
        Me.btnBrow_DSM_src = New System.Windows.Forms.Button()
        Me.txtDSM_src = New System.Windows.Forms.TextBox()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.BWDSM = New System.ComponentModel.BackgroundWorker()
        Me.OFD_DSM = New System.Windows.Forms.OpenFileDialog()
        Me.SFD_DSM = New System.Windows.Forms.SaveFileDialog()
        CType(Me.PicBar_DSM, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PicBar_DSM
        '
        Me.PicBar_DSM.Image = CType(resources.GetObject("PicBar_DSM.Image"), System.Drawing.Image)
        Me.PicBar_DSM.Location = New System.Drawing.Point(6, 75)
        Me.PicBar_DSM.Name = "PicBar_DSM"
        Me.PicBar_DSM.Size = New System.Drawing.Size(80, 80)
        Me.PicBar_DSM.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PicBar_DSM.TabIndex = 25
        Me.PicBar_DSM.TabStop = False
        Me.PicBar_DSM.Visible = False
        '
        'btnBrow_DSM_dest
        '
        Me.btnBrow_DSM_dest.BackColor = System.Drawing.Color.Transparent
        Me.btnBrow_DSM_dest.Image = CType(resources.GetObject("btnBrow_DSM_dest.Image"), System.Drawing.Image)
        Me.btnBrow_DSM_dest.Location = New System.Drawing.Point(313, 45)
        Me.btnBrow_DSM_dest.Name = "btnBrow_DSM_dest"
        Me.btnBrow_DSM_dest.Size = New System.Drawing.Size(25, 23)
        Me.btnBrow_DSM_dest.TabIndex = 24
        Me.btnBrow_DSM_dest.UseVisualStyleBackColor = False
        '
        'txtDSM_dest
        '
        Me.txtDSM_dest.Location = New System.Drawing.Point(79, 46)
        Me.txtDSM_dest.Name = "txtDSM_dest"
        Me.txtDSM_dest.Size = New System.Drawing.Size(229, 20)
        Me.txtDSM_dest.TabIndex = 23
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(6, 51)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(60, 13)
        Me.Label14.TabIndex = 22
        Me.Label14.Text = "Destination"
        '
        'btnNeu_DSM
        '
        Me.btnNeu_DSM.Enabled = False
        Me.btnNeu_DSM.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnNeu_DSM.Image = CType(resources.GetObject("btnNeu_DSM.Image"), System.Drawing.Image)
        Me.btnNeu_DSM.Location = New System.Drawing.Point(344, 12)
        Me.btnNeu_DSM.Name = "btnNeu_DSM"
        Me.btnNeu_DSM.Size = New System.Drawing.Size(55, 56)
        Me.btnNeu_DSM.TabIndex = 21
        Me.btnNeu_DSM.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnNeu_DSM.UseVisualStyleBackColor = True
        '
        'btnBrow_DSM_src
        '
        Me.btnBrow_DSM_src.BackColor = System.Drawing.Color.Transparent
        Me.btnBrow_DSM_src.Image = CType(resources.GetObject("btnBrow_DSM_src.Image"), System.Drawing.Image)
        Me.btnBrow_DSM_src.Location = New System.Drawing.Point(313, 11)
        Me.btnBrow_DSM_src.Name = "btnBrow_DSM_src"
        Me.btnBrow_DSM_src.Size = New System.Drawing.Size(25, 23)
        Me.btnBrow_DSM_src.TabIndex = 20
        Me.btnBrow_DSM_src.UseVisualStyleBackColor = False
        '
        'txtDSM_src
        '
        Me.txtDSM_src.Location = New System.Drawing.Point(79, 12)
        Me.txtDSM_src.Name = "txtDSM_src"
        Me.txtDSM_src.Size = New System.Drawing.Size(229, 20)
        Me.txtDSM_src.TabIndex = 19
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(6, 15)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(41, 13)
        Me.Label15.TabIndex = 18
        Me.Label15.Text = "Source"
        '
        'BWDSM
        '
        '
        'OFD_DSM
        '
        Me.OFD_DSM.FileName = "Source File"
        Me.OFD_DSM.Filter = "Excel files (*.xls,*xlsx)|*.xls;*.xlsx"
        '
        'SFD_DSM
        '
        Me.SFD_DSM.Filter = "Excel files (*.xls,*.xlsx)|*.xls;*xlsx"
        '
        'frmDSM
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.ClientSize = New System.Drawing.Size(404, 159)
        Me.Controls.Add(Me.PicBar_DSM)
        Me.Controls.Add(Me.btnBrow_DSM_dest)
        Me.Controls.Add(Me.txtDSM_dest)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.btnNeu_DSM)
        Me.Controls.Add(Me.btnBrow_DSM_src)
        Me.Controls.Add(Me.txtDSM_src)
        Me.Controls.Add(Me.Label15)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmDSM"
        Me.Text = "Daily Stock Mutation"
        CType(Me.PicBar_DSM, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents PicBar_DSM As System.Windows.Forms.PictureBox
    Friend WithEvents btnBrow_DSM_dest As System.Windows.Forms.Button
    Friend WithEvents txtDSM_dest As System.Windows.Forms.TextBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents btnNeu_DSM As System.Windows.Forms.Button
    Friend WithEvents btnBrow_DSM_src As System.Windows.Forms.Button
    Friend WithEvents txtDSM_src As System.Windows.Forms.TextBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents BWDSM As System.ComponentModel.BackgroundWorker
    Friend WithEvents OFD_DSM As System.Windows.Forms.OpenFileDialog
    Friend WithEvents SFD_DSM As System.Windows.Forms.SaveFileDialog
End Class
