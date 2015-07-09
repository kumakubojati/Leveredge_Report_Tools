<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmPSCW_IC
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmPSCW_IC))
        Me.PicBar_PSCW_IC = New System.Windows.Forms.PictureBox()
        Me.btnBrow_PSCW_IC_dest = New System.Windows.Forms.Button()
        Me.txtPSCW_IC_dest = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.btnNeu_PSCW_IC = New System.Windows.Forms.Button()
        Me.btnBrow_PSCW_IC_src = New System.Windows.Forms.Button()
        Me.txtPSCW_IC_src = New System.Windows.Forms.TextBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.OFD_PSCW_IC = New System.Windows.Forms.OpenFileDialog()
        Me.SFD_PSCW_IC = New System.Windows.Forms.SaveFileDialog()
        Me.BWPSCW_IC = New System.ComponentModel.BackgroundWorker()
        CType(Me.PicBar_PSCW_IC, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PicBar_PSCW_IC
        '
        Me.PicBar_PSCW_IC.Image = CType(resources.GetObject("PicBar_PSCW_IC.Image"), System.Drawing.Image)
        Me.PicBar_PSCW_IC.Location = New System.Drawing.Point(11, 74)
        Me.PicBar_PSCW_IC.Name = "PicBar_PSCW_IC"
        Me.PicBar_PSCW_IC.Size = New System.Drawing.Size(80, 80)
        Me.PicBar_PSCW_IC.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PicBar_PSCW_IC.TabIndex = 79
        Me.PicBar_PSCW_IC.TabStop = False
        Me.PicBar_PSCW_IC.Visible = False
        '
        'btnBrow_PSCW_IC_dest
        '
        Me.btnBrow_PSCW_IC_dest.Image = CType(resources.GetObject("btnBrow_PSCW_IC_dest.Image"), System.Drawing.Image)
        Me.btnBrow_PSCW_IC_dest.Location = New System.Drawing.Point(316, 46)
        Me.btnBrow_PSCW_IC_dest.Name = "btnBrow_PSCW_IC_dest"
        Me.btnBrow_PSCW_IC_dest.Size = New System.Drawing.Size(26, 23)
        Me.btnBrow_PSCW_IC_dest.TabIndex = 82
        Me.btnBrow_PSCW_IC_dest.UseVisualStyleBackColor = True
        '
        'txtPSCW_IC_dest
        '
        Me.txtPSCW_IC_dest.Location = New System.Drawing.Point(82, 47)
        Me.txtPSCW_IC_dest.Name = "txtPSCW_IC_dest"
        Me.txtPSCW_IC_dest.Size = New System.Drawing.Size(229, 20)
        Me.txtPSCW_IC_dest.TabIndex = 81
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(11, 52)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(60, 13)
        Me.Label7.TabIndex = 80
        Me.Label7.Text = "Destination"
        '
        'btnNeu_PSCW_IC
        '
        Me.btnNeu_PSCW_IC.Enabled = False
        Me.btnNeu_PSCW_IC.Image = CType(resources.GetObject("btnNeu_PSCW_IC.Image"), System.Drawing.Image)
        Me.btnNeu_PSCW_IC.Location = New System.Drawing.Point(349, 10)
        Me.btnNeu_PSCW_IC.Name = "btnNeu_PSCW_IC"
        Me.btnNeu_PSCW_IC.Size = New System.Drawing.Size(54, 59)
        Me.btnNeu_PSCW_IC.TabIndex = 78
        Me.btnNeu_PSCW_IC.UseVisualStyleBackColor = True
        '
        'btnBrow_PSCW_IC_src
        '
        Me.btnBrow_PSCW_IC_src.Image = CType(resources.GetObject("btnBrow_PSCW_IC_src.Image"), System.Drawing.Image)
        Me.btnBrow_PSCW_IC_src.Location = New System.Drawing.Point(316, 12)
        Me.btnBrow_PSCW_IC_src.Name = "btnBrow_PSCW_IC_src"
        Me.btnBrow_PSCW_IC_src.Size = New System.Drawing.Size(26, 23)
        Me.btnBrow_PSCW_IC_src.TabIndex = 77
        Me.btnBrow_PSCW_IC_src.UseVisualStyleBackColor = True
        '
        'txtPSCW_IC_src
        '
        Me.txtPSCW_IC_src.Location = New System.Drawing.Point(82, 13)
        Me.txtPSCW_IC_src.Name = "txtPSCW_IC_src"
        Me.txtPSCW_IC_src.Size = New System.Drawing.Size(229, 20)
        Me.txtPSCW_IC_src.TabIndex = 76
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(12, 18)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(41, 13)
        Me.Label13.TabIndex = 75
        Me.Label13.Text = "Source"
        '
        'OFD_PSCW_IC
        '
        Me.OFD_PSCW_IC.FileName = "Source File"
        Me.OFD_PSCW_IC.Filter = "Excel files (*.xls,*xlsx)|*.xls;*.xlsx"
        '
        'SFD_PSCW_IC
        '
        Me.SFD_PSCW_IC.Filter = "Excel files (*.xls,*xlsx)|*.xls;*.xlsx"
        '
        'BWPSCW_IC
        '
        '
        'frmPSCW
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.ClientSize = New System.Drawing.Size(411, 161)
        Me.Controls.Add(Me.PicBar_PSCW_IC)
        Me.Controls.Add(Me.btnBrow_PSCW_IC_dest)
        Me.Controls.Add(Me.txtPSCW_IC_dest)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.btnNeu_PSCW_IC)
        Me.Controls.Add(Me.btnBrow_PSCW_IC_src)
        Me.Controls.Add(Me.txtPSCW_IC_src)
        Me.Controls.Add(Me.Label13)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmPSCW"
        Me.Text = "Product Sales By Case Weekly Report"
        CType(Me.PicBar_PSCW_IC, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents PicBar_PSCW_IC As System.Windows.Forms.PictureBox
    Friend WithEvents btnBrow_PSCW_IC_dest As System.Windows.Forms.Button
    Friend WithEvents txtPSCW_IC_dest As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents btnNeu_PSCW_IC As System.Windows.Forms.Button
    Friend WithEvents btnBrow_PSCW_IC_src As System.Windows.Forms.Button
    Friend WithEvents txtPSCW_IC_src As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents OFD_PSCW_IC As System.Windows.Forms.OpenFileDialog
    Friend WithEvents SFD_PSCW_IC As System.Windows.Forms.SaveFileDialog
    Friend WithEvents BWPSCW_IC As System.ComponentModel.BackgroundWorker
End Class
