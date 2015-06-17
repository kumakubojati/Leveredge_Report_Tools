<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmOSC_IC
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmOSC_IC))
        Me.PicBar_OSC_IC = New System.Windows.Forms.PictureBox()
        Me.btnBrow_OSC_IC_dest = New System.Windows.Forms.Button()
        Me.txtOSC_IC_dest = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.btnNeu_OSC_IC = New System.Windows.Forms.Button()
        Me.btnBrow_OSC_IC_src = New System.Windows.Forms.Button()
        Me.txtOSC_IC_src = New System.Windows.Forms.TextBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.OFD_OSC_IC = New System.Windows.Forms.OpenFileDialog()
        Me.SFD_OSC_IC = New System.Windows.Forms.SaveFileDialog()
        Me.BWOSC_IC = New System.ComponentModel.BackgroundWorker()
        CType(Me.PicBar_OSC_IC, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PicBar_OSC_IC
        '
        Me.PicBar_OSC_IC.Image = CType(resources.GetObject("PicBar_OSC_IC.Image"), System.Drawing.Image)
        Me.PicBar_OSC_IC.Location = New System.Drawing.Point(8, 71)
        Me.PicBar_OSC_IC.Name = "PicBar_OSC_IC"
        Me.PicBar_OSC_IC.Size = New System.Drawing.Size(80, 80)
        Me.PicBar_OSC_IC.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PicBar_OSC_IC.TabIndex = 89
        Me.PicBar_OSC_IC.TabStop = False
        Me.PicBar_OSC_IC.Visible = False
        '
        'btnBrow_OSC_IC_dest
        '
        Me.btnBrow_OSC_IC_dest.Image = CType(resources.GetObject("btnBrow_OSC_IC_dest.Image"), System.Drawing.Image)
        Me.btnBrow_OSC_IC_dest.Location = New System.Drawing.Point(313, 43)
        Me.btnBrow_OSC_IC_dest.Name = "btnBrow_OSC_IC_dest"
        Me.btnBrow_OSC_IC_dest.Size = New System.Drawing.Size(26, 23)
        Me.btnBrow_OSC_IC_dest.TabIndex = 92
        Me.btnBrow_OSC_IC_dest.UseVisualStyleBackColor = True
        '
        'txtOSC_IC_dest
        '
        Me.txtOSC_IC_dest.Location = New System.Drawing.Point(79, 44)
        Me.txtOSC_IC_dest.Name = "txtOSC_IC_dest"
        Me.txtOSC_IC_dest.Size = New System.Drawing.Size(229, 20)
        Me.txtOSC_IC_dest.TabIndex = 91
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(8, 49)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(60, 13)
        Me.Label7.TabIndex = 90
        Me.Label7.Text = "Destination"
        '
        'btnNeu_OSC_IC
        '
        Me.btnNeu_OSC_IC.Enabled = False
        Me.btnNeu_OSC_IC.Image = CType(resources.GetObject("btnNeu_OSC_IC.Image"), System.Drawing.Image)
        Me.btnNeu_OSC_IC.Location = New System.Drawing.Point(346, 7)
        Me.btnNeu_OSC_IC.Name = "btnNeu_OSC_IC"
        Me.btnNeu_OSC_IC.Size = New System.Drawing.Size(54, 59)
        Me.btnNeu_OSC_IC.TabIndex = 88
        Me.btnNeu_OSC_IC.UseVisualStyleBackColor = True
        '
        'btnBrow_OSC_IC_src
        '
        Me.btnBrow_OSC_IC_src.Image = CType(resources.GetObject("btnBrow_OSC_IC_src.Image"), System.Drawing.Image)
        Me.btnBrow_OSC_IC_src.Location = New System.Drawing.Point(313, 9)
        Me.btnBrow_OSC_IC_src.Name = "btnBrow_OSC_IC_src"
        Me.btnBrow_OSC_IC_src.Size = New System.Drawing.Size(26, 23)
        Me.btnBrow_OSC_IC_src.TabIndex = 87
        Me.btnBrow_OSC_IC_src.UseVisualStyleBackColor = True
        '
        'txtOSC_IC_src
        '
        Me.txtOSC_IC_src.Location = New System.Drawing.Point(79, 10)
        Me.txtOSC_IC_src.Name = "txtOSC_IC_src"
        Me.txtOSC_IC_src.Size = New System.Drawing.Size(229, 20)
        Me.txtOSC_IC_src.TabIndex = 86
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(9, 15)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(41, 13)
        Me.Label13.TabIndex = 85
        Me.Label13.Text = "Source"
        '
        'OFD_OSC_IC
        '
        Me.OFD_OSC_IC.FileName = "Source File"
        Me.OFD_OSC_IC.Filter = "Excel files (*.xls,*xlsx)|*.xls;*.xlsx"
        '
        'SFD_OSC_IC
        '
        Me.SFD_OSC_IC.Filter = "Excel files (*.xls,*.xlsx)|*.xls;*xlsx"
        '
        'BWOSC_IC
        '
        '
        'frmOSC_IC
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.ClientSize = New System.Drawing.Size(407, 159)
        Me.Controls.Add(Me.PicBar_OSC_IC)
        Me.Controls.Add(Me.btnBrow_OSC_IC_dest)
        Me.Controls.Add(Me.txtOSC_IC_dest)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.btnNeu_OSC_IC)
        Me.Controls.Add(Me.btnBrow_OSC_IC_src)
        Me.Controls.Add(Me.txtOSC_IC_src)
        Me.Controls.Add(Me.Label13)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmOSC_IC"
        Me.Text = "Outlet Store Class Report"
        CType(Me.PicBar_OSC_IC, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents PicBar_OSC_IC As System.Windows.Forms.PictureBox
    Friend WithEvents btnBrow_OSC_IC_dest As System.Windows.Forms.Button
    Friend WithEvents txtOSC_IC_dest As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents btnNeu_OSC_IC As System.Windows.Forms.Button
    Friend WithEvents btnBrow_OSC_IC_src As System.Windows.Forms.Button
    Friend WithEvents txtOSC_IC_src As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents OFD_OSC_IC As System.Windows.Forms.OpenFileDialog
    Friend WithEvents SFD_OSC_IC As System.Windows.Forms.SaveFileDialog
    Friend WithEvents BWOSC_IC As System.ComponentModel.BackgroundWorker
End Class
