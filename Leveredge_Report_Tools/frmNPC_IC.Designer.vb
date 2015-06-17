<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmNPC_IC
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmNPC_IC))
        Me.PicBar_NPC_IC = New System.Windows.Forms.PictureBox()
        Me.btnBrow_NPC_IC_dest = New System.Windows.Forms.Button()
        Me.txtNPC_IC_dest = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.btnNeu_NPC_IC = New System.Windows.Forms.Button()
        Me.btnBrow_NPC_IC_src = New System.Windows.Forms.Button()
        Me.txtNPC_IC_src = New System.Windows.Forms.TextBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.OFD_NPC_IC = New System.Windows.Forms.OpenFileDialog()
        Me.SFD_NPC_IC = New System.Windows.Forms.SaveFileDialog()
        Me.BWNPC_IC = New System.ComponentModel.BackgroundWorker()
        CType(Me.PicBar_NPC_IC, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PicBar_NPC_IC
        '
        Me.PicBar_NPC_IC.Image = CType(resources.GetObject("PicBar_NPC_IC.Image"), System.Drawing.Image)
        Me.PicBar_NPC_IC.Location = New System.Drawing.Point(8, 73)
        Me.PicBar_NPC_IC.Name = "PicBar_NPC_IC"
        Me.PicBar_NPC_IC.Size = New System.Drawing.Size(80, 80)
        Me.PicBar_NPC_IC.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PicBar_NPC_IC.TabIndex = 86
        Me.PicBar_NPC_IC.TabStop = False
        Me.PicBar_NPC_IC.Visible = False
        '
        'btnBrow_NPC_IC_dest
        '
        Me.btnBrow_NPC_IC_dest.Image = CType(resources.GetObject("btnBrow_NPC_IC_dest.Image"), System.Drawing.Image)
        Me.btnBrow_NPC_IC_dest.Location = New System.Drawing.Point(313, 45)
        Me.btnBrow_NPC_IC_dest.Name = "btnBrow_NPC_IC_dest"
        Me.btnBrow_NPC_IC_dest.Size = New System.Drawing.Size(26, 23)
        Me.btnBrow_NPC_IC_dest.TabIndex = 89
        Me.btnBrow_NPC_IC_dest.UseVisualStyleBackColor = True
        '
        'txtNPC_IC_dest
        '
        Me.txtNPC_IC_dest.Location = New System.Drawing.Point(79, 46)
        Me.txtNPC_IC_dest.Name = "txtNPC_IC_dest"
        Me.txtNPC_IC_dest.Size = New System.Drawing.Size(229, 20)
        Me.txtNPC_IC_dest.TabIndex = 88
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(8, 51)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(60, 13)
        Me.Label7.TabIndex = 87
        Me.Label7.Text = "Destination"
        '
        'btnNeu_NPC_IC
        '
        Me.btnNeu_NPC_IC.Enabled = False
        Me.btnNeu_NPC_IC.Image = CType(resources.GetObject("btnNeu_NPC_IC.Image"), System.Drawing.Image)
        Me.btnNeu_NPC_IC.Location = New System.Drawing.Point(346, 9)
        Me.btnNeu_NPC_IC.Name = "btnNeu_NPC_IC"
        Me.btnNeu_NPC_IC.Size = New System.Drawing.Size(54, 59)
        Me.btnNeu_NPC_IC.TabIndex = 85
        Me.btnNeu_NPC_IC.UseVisualStyleBackColor = True
        '
        'btnBrow_NPC_IC_src
        '
        Me.btnBrow_NPC_IC_src.Image = CType(resources.GetObject("btnBrow_NPC_IC_src.Image"), System.Drawing.Image)
        Me.btnBrow_NPC_IC_src.Location = New System.Drawing.Point(313, 11)
        Me.btnBrow_NPC_IC_src.Name = "btnBrow_NPC_IC_src"
        Me.btnBrow_NPC_IC_src.Size = New System.Drawing.Size(26, 23)
        Me.btnBrow_NPC_IC_src.TabIndex = 84
        Me.btnBrow_NPC_IC_src.UseVisualStyleBackColor = True
        '
        'txtNPC_IC_src
        '
        Me.txtNPC_IC_src.Location = New System.Drawing.Point(79, 12)
        Me.txtNPC_IC_src.Name = "txtNPC_IC_src"
        Me.txtNPC_IC_src.Size = New System.Drawing.Size(229, 20)
        Me.txtNPC_IC_src.TabIndex = 83
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(9, 17)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(41, 13)
        Me.Label13.TabIndex = 82
        Me.Label13.Text = "Source"
        '
        'OFD_NPC_IC
        '
        Me.OFD_NPC_IC.FileName = "Source File"
        Me.OFD_NPC_IC.Filter = "Excel files (*.xls,*xlsx)|*.xls;*.xlsx"
        '
        'SFD_NPC_IC
        '
        Me.SFD_NPC_IC.Filter = "Excel files (*.xls,*.xlsx)|*.xls;*xlsx"
        '
        'BWNPC_IC
        '
        '
        'frmNPC_IC
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.ClientSize = New System.Drawing.Size(407, 159)
        Me.Controls.Add(Me.PicBar_NPC_IC)
        Me.Controls.Add(Me.btnBrow_NPC_IC_dest)
        Me.Controls.Add(Me.txtNPC_IC_dest)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.btnNeu_NPC_IC)
        Me.Controls.Add(Me.btnBrow_NPC_IC_src)
        Me.Controls.Add(Me.txtNPC_IC_src)
        Me.Controls.Add(Me.Label13)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmNPC_IC"
        Me.Text = "Non Performance Cabinet Report"
        CType(Me.PicBar_NPC_IC, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents PicBar_NPC_IC As System.Windows.Forms.PictureBox
    Friend WithEvents btnBrow_NPC_IC_dest As System.Windows.Forms.Button
    Friend WithEvents txtNPC_IC_dest As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents btnNeu_NPC_IC As System.Windows.Forms.Button
    Friend WithEvents btnBrow_NPC_IC_src As System.Windows.Forms.Button
    Friend WithEvents txtNPC_IC_src As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents OFD_NPC_IC As System.Windows.Forms.OpenFileDialog
    Friend WithEvents SFD_NPC_IC As System.Windows.Forms.SaveFileDialog
    Friend WithEvents BWNPC_IC As System.ComponentModel.BackgroundWorker
End Class
