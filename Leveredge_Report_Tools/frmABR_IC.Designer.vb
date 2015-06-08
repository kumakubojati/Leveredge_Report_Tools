<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmABR_IC
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmABR_IC))
        Me.PicBar_ABR_IC = New System.Windows.Forms.PictureBox()
        Me.btnBrow_ABR_IC_dest = New System.Windows.Forms.Button()
        Me.txtABR_IC_dest = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.btnNeu_ABR_IC = New System.Windows.Forms.Button()
        Me.btnBrow_ABR_IC_src = New System.Windows.Forms.Button()
        Me.txtABR_IC_src = New System.Windows.Forms.TextBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.OFD_ABR_IC = New System.Windows.Forms.OpenFileDialog()
        Me.SFD_ABR_IC = New System.Windows.Forms.SaveFileDialog()
        Me.BWABR_IC = New System.ComponentModel.BackgroundWorker()
        CType(Me.PicBar_ABR_IC, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PicBar_ABR_IC
        '
        Me.PicBar_ABR_IC.Image = CType(resources.GetObject("PicBar_ABR_IC.Image"), System.Drawing.Image)
        Me.PicBar_ABR_IC.Location = New System.Drawing.Point(8, 74)
        Me.PicBar_ABR_IC.Name = "PicBar_ABR_IC"
        Me.PicBar_ABR_IC.Size = New System.Drawing.Size(80, 80)
        Me.PicBar_ABR_IC.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PicBar_ABR_IC.TabIndex = 70
        Me.PicBar_ABR_IC.TabStop = False
        Me.PicBar_ABR_IC.Visible = False
        '
        'btnBrow_ABR_IC_dest
        '
        Me.btnBrow_ABR_IC_dest.Image = CType(resources.GetObject("btnBrow_ABR_IC_dest.Image"), System.Drawing.Image)
        Me.btnBrow_ABR_IC_dest.Location = New System.Drawing.Point(313, 46)
        Me.btnBrow_ABR_IC_dest.Name = "btnBrow_ABR_IC_dest"
        Me.btnBrow_ABR_IC_dest.Size = New System.Drawing.Size(26, 23)
        Me.btnBrow_ABR_IC_dest.TabIndex = 73
        Me.btnBrow_ABR_IC_dest.UseVisualStyleBackColor = True
        '
        'txtABR_IC_dest
        '
        Me.txtABR_IC_dest.Location = New System.Drawing.Point(79, 47)
        Me.txtABR_IC_dest.Name = "txtABR_IC_dest"
        Me.txtABR_IC_dest.Size = New System.Drawing.Size(229, 20)
        Me.txtABR_IC_dest.TabIndex = 72
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(8, 52)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(60, 13)
        Me.Label7.TabIndex = 71
        Me.Label7.Text = "Destination"
        '
        'btnNeu_ABR_IC
        '
        Me.btnNeu_ABR_IC.Enabled = False
        Me.btnNeu_ABR_IC.Image = CType(resources.GetObject("btnNeu_ABR_IC.Image"), System.Drawing.Image)
        Me.btnNeu_ABR_IC.Location = New System.Drawing.Point(346, 10)
        Me.btnNeu_ABR_IC.Name = "btnNeu_ABR_IC"
        Me.btnNeu_ABR_IC.Size = New System.Drawing.Size(54, 59)
        Me.btnNeu_ABR_IC.TabIndex = 69
        Me.btnNeu_ABR_IC.UseVisualStyleBackColor = True
        '
        'btnBrow_ABR_IC_src
        '
        Me.btnBrow_ABR_IC_src.Image = CType(resources.GetObject("btnBrow_ABR_IC_src.Image"), System.Drawing.Image)
        Me.btnBrow_ABR_IC_src.Location = New System.Drawing.Point(313, 12)
        Me.btnBrow_ABR_IC_src.Name = "btnBrow_ABR_IC_src"
        Me.btnBrow_ABR_IC_src.Size = New System.Drawing.Size(26, 23)
        Me.btnBrow_ABR_IC_src.TabIndex = 68
        Me.btnBrow_ABR_IC_src.UseVisualStyleBackColor = True
        '
        'txtABR_IC_src
        '
        Me.txtABR_IC_src.Location = New System.Drawing.Point(79, 13)
        Me.txtABR_IC_src.Name = "txtABR_IC_src"
        Me.txtABR_IC_src.Size = New System.Drawing.Size(229, 20)
        Me.txtABR_IC_src.TabIndex = 67
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(9, 18)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(41, 13)
        Me.Label13.TabIndex = 66
        Me.Label13.Text = "Source"
        '
        'OFD_ABR_IC
        '
        Me.OFD_ABR_IC.FileName = "Source File"
        Me.OFD_ABR_IC.Filter = "Excel files (*.xls,*xlsx)|*.xls;*.xlsx"
        '
        'SFD_ABR_IC
        '
        Me.SFD_ABR_IC.Filter = "Excel files (*.xls,*.xlsx)|*.xls;*xlsx"
        '
        'BWABR_IC
        '
        '
        'frmABR_IC
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.ClientSize = New System.Drawing.Size(409, 160)
        Me.Controls.Add(Me.PicBar_ABR_IC)
        Me.Controls.Add(Me.btnBrow_ABR_IC_dest)
        Me.Controls.Add(Me.txtABR_IC_dest)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.btnNeu_ABR_IC)
        Me.Controls.Add(Me.btnBrow_ABR_IC_src)
        Me.Controls.Add(Me.txtABR_IC_src)
        Me.Controls.Add(Me.Label13)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmABR_IC"
        Me.Text = "Analysis Backup Report"
        CType(Me.PicBar_ABR_IC, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents PicBar_ABR_IC As System.Windows.Forms.PictureBox
    Friend WithEvents btnBrow_ABR_IC_dest As System.Windows.Forms.Button
    Friend WithEvents txtABR_IC_dest As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents btnNeu_ABR_IC As System.Windows.Forms.Button
    Friend WithEvents btnBrow_ABR_IC_src As System.Windows.Forms.Button
    Friend WithEvents txtABR_IC_src As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents OFD_ABR_IC As System.Windows.Forms.OpenFileDialog
    Friend WithEvents SFD_ABR_IC As System.Windows.Forms.SaveFileDialog
    Friend WithEvents BWABR_IC As System.ComponentModel.BackgroundWorker
End Class
