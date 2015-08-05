<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmPSVW_IC
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmPSVW_IC))
        Me.PicBar_PSVW_IC = New System.Windows.Forms.PictureBox()
        Me.btnBrow_PSVW_IC_dest = New System.Windows.Forms.Button()
        Me.txtPSVW_IC_dest = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.btnNeu_PSVW_IC = New System.Windows.Forms.Button()
        Me.btnBrow_PSVW_IC_src = New System.Windows.Forms.Button()
        Me.txtPSVW_IC_src = New System.Windows.Forms.TextBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.OFD_PSVW_IC = New System.Windows.Forms.OpenFileDialog()
        Me.SFD_PSVW_IC = New System.Windows.Forms.SaveFileDialog()
        Me.BWPSVW_IC = New System.ComponentModel.BackgroundWorker()
        CType(Me.PicBar_PSVW_IC, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PicBar_PSVW_IC
        '
        Me.PicBar_PSVW_IC.Image = CType(resources.GetObject("PicBar_PSVW_IC.Image"), System.Drawing.Image)
        Me.PicBar_PSVW_IC.Location = New System.Drawing.Point(8, 72)
        Me.PicBar_PSVW_IC.Name = "PicBar_PSVW_IC"
        Me.PicBar_PSVW_IC.Size = New System.Drawing.Size(80, 80)
        Me.PicBar_PSVW_IC.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PicBar_PSVW_IC.TabIndex = 87
        Me.PicBar_PSVW_IC.TabStop = False
        Me.PicBar_PSVW_IC.Visible = False
        '
        'btnBrow_PSVW_IC_dest
        '
        Me.btnBrow_PSVW_IC_dest.Image = CType(resources.GetObject("btnBrow_PSVW_IC_dest.Image"), System.Drawing.Image)
        Me.btnBrow_PSVW_IC_dest.Location = New System.Drawing.Point(313, 44)
        Me.btnBrow_PSVW_IC_dest.Name = "btnBrow_PSVW_IC_dest"
        Me.btnBrow_PSVW_IC_dest.Size = New System.Drawing.Size(26, 23)
        Me.btnBrow_PSVW_IC_dest.TabIndex = 90
        Me.btnBrow_PSVW_IC_dest.UseVisualStyleBackColor = True
        '
        'txtPSVW_IC_dest
        '
        Me.txtPSVW_IC_dest.Location = New System.Drawing.Point(79, 45)
        Me.txtPSVW_IC_dest.Name = "txtPSVW_IC_dest"
        Me.txtPSVW_IC_dest.Size = New System.Drawing.Size(229, 20)
        Me.txtPSVW_IC_dest.TabIndex = 89
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(8, 50)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(60, 13)
        Me.Label7.TabIndex = 88
        Me.Label7.Text = "Destination"
        '
        'btnNeu_PSVW_IC
        '
        Me.btnNeu_PSVW_IC.Enabled = False
        Me.btnNeu_PSVW_IC.Image = CType(resources.GetObject("btnNeu_PSVW_IC.Image"), System.Drawing.Image)
        Me.btnNeu_PSVW_IC.Location = New System.Drawing.Point(346, 8)
        Me.btnNeu_PSVW_IC.Name = "btnNeu_PSVW_IC"
        Me.btnNeu_PSVW_IC.Size = New System.Drawing.Size(54, 59)
        Me.btnNeu_PSVW_IC.TabIndex = 86
        Me.btnNeu_PSVW_IC.UseVisualStyleBackColor = True
        '
        'btnBrow_PSVW_IC_src
        '
        Me.btnBrow_PSVW_IC_src.Image = CType(resources.GetObject("btnBrow_PSVW_IC_src.Image"), System.Drawing.Image)
        Me.btnBrow_PSVW_IC_src.Location = New System.Drawing.Point(313, 10)
        Me.btnBrow_PSVW_IC_src.Name = "btnBrow_PSVW_IC_src"
        Me.btnBrow_PSVW_IC_src.Size = New System.Drawing.Size(26, 23)
        Me.btnBrow_PSVW_IC_src.TabIndex = 85
        Me.btnBrow_PSVW_IC_src.UseVisualStyleBackColor = True
        '
        'txtPSVW_IC_src
        '
        Me.txtPSVW_IC_src.Location = New System.Drawing.Point(79, 11)
        Me.txtPSVW_IC_src.Name = "txtPSVW_IC_src"
        Me.txtPSVW_IC_src.Size = New System.Drawing.Size(229, 20)
        Me.txtPSVW_IC_src.TabIndex = 84
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(9, 16)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(41, 13)
        Me.Label13.TabIndex = 83
        Me.Label13.Text = "Source"
        '
        'OFD_PSVW_IC
        '
        Me.OFD_PSVW_IC.FileName = "Source File"
        Me.OFD_PSVW_IC.Filter = "Excel files (*.xls,*xlsx)|*.xls;*.xlsx"
        '
        'SFD_PSVW_IC
        '
        Me.SFD_PSVW_IC.Filter = "Excel files (*.xls,*xlsx)|*.xls;*.xlsx"
        '
        'BWPSVW_IC
        '
        '
        'frmPSVW_IC
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.ClientSize = New System.Drawing.Size(412, 157)
        Me.Controls.Add(Me.PicBar_PSVW_IC)
        Me.Controls.Add(Me.btnBrow_PSVW_IC_dest)
        Me.Controls.Add(Me.txtPSVW_IC_dest)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.btnNeu_PSVW_IC)
        Me.Controls.Add(Me.btnBrow_PSVW_IC_src)
        Me.Controls.Add(Me.txtPSVW_IC_src)
        Me.Controls.Add(Me.Label13)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmPSVW_IC"
        Me.Text = "Product Sales By Volume Weekly Report"
        CType(Me.PicBar_PSVW_IC, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents PicBar_PSVW_IC As System.Windows.Forms.PictureBox
    Friend WithEvents btnBrow_PSVW_IC_dest As System.Windows.Forms.Button
    Friend WithEvents txtPSVW_IC_dest As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents btnNeu_PSVW_IC As System.Windows.Forms.Button
    Friend WithEvents btnBrow_PSVW_IC_src As System.Windows.Forms.Button
    Friend WithEvents txtPSVW_IC_src As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents OFD_PSVW_IC As System.Windows.Forms.OpenFileDialog
    Friend WithEvents SFD_PSVW_IC As System.Windows.Forms.SaveFileDialog
    Friend WithEvents BWPSVW_IC As System.ComponentModel.BackgroundWorker
End Class
