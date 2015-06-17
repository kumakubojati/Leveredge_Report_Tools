<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmDD_IC
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmDD_IC))
        Me.gbRepType_DD = New System.Windows.Forms.GroupBox()
        Me.RBDD_Dist = New System.Windows.Forms.RadioButton()
        Me.RBDD_NotDist = New System.Windows.Forms.RadioButton()
        Me.PicBar_DD_IC = New System.Windows.Forms.PictureBox()
        Me.btnBrow_DD_IC_dest = New System.Windows.Forms.Button()
        Me.txtDD_IC_dest = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.btnNeu_DD_IC = New System.Windows.Forms.Button()
        Me.btnBrow_DD_IC_src = New System.Windows.Forms.Button()
        Me.txtDD_IC_src = New System.Windows.Forms.TextBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.OFD_DD_IC = New System.Windows.Forms.OpenFileDialog()
        Me.SFD_DD_IC = New System.Windows.Forms.SaveFileDialog()
        Me.BWDD_IC = New System.ComponentModel.BackgroundWorker()
        Me.gbRepType_DD.SuspendLayout()
        CType(Me.PicBar_DD_IC, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'gbRepType_DD
        '
        Me.gbRepType_DD.Controls.Add(Me.RBDD_Dist)
        Me.gbRepType_DD.Controls.Add(Me.RBDD_NotDist)
        Me.gbRepType_DD.Location = New System.Drawing.Point(8, 10)
        Me.gbRepType_DD.Name = "gbRepType_DD"
        Me.gbRepType_DD.Size = New System.Drawing.Size(227, 53)
        Me.gbRepType_DD.TabIndex = 76
        Me.gbRepType_DD.TabStop = False
        Me.gbRepType_DD.Text = "Report Type"
        '
        'RBDD_Dist
        '
        Me.RBDD_Dist.AutoSize = True
        Me.RBDD_Dist.Checked = True
        Me.RBDD_Dist.Location = New System.Drawing.Point(7, 22)
        Me.RBDD_Dist.Name = "RBDD_Dist"
        Me.RBDD_Dist.Size = New System.Drawing.Size(75, 17)
        Me.RBDD_Dist.TabIndex = 1
        Me.RBDD_Dist.TabStop = True
        Me.RBDD_Dist.Text = "Distributed"
        Me.RBDD_Dist.UseVisualStyleBackColor = True
        '
        'RBDD_NotDist
        '
        Me.RBDD_NotDist.AutoSize = True
        Me.RBDD_NotDist.Location = New System.Drawing.Point(96, 22)
        Me.RBDD_NotDist.Name = "RBDD_NotDist"
        Me.RBDD_NotDist.Size = New System.Drawing.Size(95, 17)
        Me.RBDD_NotDist.TabIndex = 0
        Me.RBDD_NotDist.Text = "Not Distributed"
        Me.RBDD_NotDist.UseVisualStyleBackColor = True
        '
        'PicBar_DD_IC
        '
        Me.PicBar_DD_IC.Image = CType(resources.GetObject("PicBar_DD_IC.Image"), System.Drawing.Image)
        Me.PicBar_DD_IC.Location = New System.Drawing.Point(9, 129)
        Me.PicBar_DD_IC.Name = "PicBar_DD_IC"
        Me.PicBar_DD_IC.Size = New System.Drawing.Size(80, 80)
        Me.PicBar_DD_IC.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PicBar_DD_IC.TabIndex = 72
        Me.PicBar_DD_IC.TabStop = False
        Me.PicBar_DD_IC.Visible = False
        '
        'btnBrow_DD_IC_dest
        '
        Me.btnBrow_DD_IC_dest.Image = CType(resources.GetObject("btnBrow_DD_IC_dest.Image"), System.Drawing.Image)
        Me.btnBrow_DD_IC_dest.Location = New System.Drawing.Point(314, 101)
        Me.btnBrow_DD_IC_dest.Name = "btnBrow_DD_IC_dest"
        Me.btnBrow_DD_IC_dest.Size = New System.Drawing.Size(26, 23)
        Me.btnBrow_DD_IC_dest.TabIndex = 75
        Me.btnBrow_DD_IC_dest.UseVisualStyleBackColor = True
        '
        'txtDD_IC_dest
        '
        Me.txtDD_IC_dest.Location = New System.Drawing.Point(80, 102)
        Me.txtDD_IC_dest.Name = "txtDD_IC_dest"
        Me.txtDD_IC_dest.Size = New System.Drawing.Size(229, 20)
        Me.txtDD_IC_dest.TabIndex = 74
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(9, 107)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(60, 13)
        Me.Label7.TabIndex = 73
        Me.Label7.Text = "Destination"
        '
        'btnNeu_DD_IC
        '
        Me.btnNeu_DD_IC.Enabled = False
        Me.btnNeu_DD_IC.Image = CType(resources.GetObject("btnNeu_DD_IC.Image"), System.Drawing.Image)
        Me.btnNeu_DD_IC.Location = New System.Drawing.Point(347, 65)
        Me.btnNeu_DD_IC.Name = "btnNeu_DD_IC"
        Me.btnNeu_DD_IC.Size = New System.Drawing.Size(54, 59)
        Me.btnNeu_DD_IC.TabIndex = 71
        Me.btnNeu_DD_IC.UseVisualStyleBackColor = True
        '
        'btnBrow_DD_IC_src
        '
        Me.btnBrow_DD_IC_src.Image = CType(resources.GetObject("btnBrow_DD_IC_src.Image"), System.Drawing.Image)
        Me.btnBrow_DD_IC_src.Location = New System.Drawing.Point(314, 67)
        Me.btnBrow_DD_IC_src.Name = "btnBrow_DD_IC_src"
        Me.btnBrow_DD_IC_src.Size = New System.Drawing.Size(26, 23)
        Me.btnBrow_DD_IC_src.TabIndex = 70
        Me.btnBrow_DD_IC_src.UseVisualStyleBackColor = True
        '
        'txtDD_IC_src
        '
        Me.txtDD_IC_src.Location = New System.Drawing.Point(80, 68)
        Me.txtDD_IC_src.Name = "txtDD_IC_src"
        Me.txtDD_IC_src.Size = New System.Drawing.Size(229, 20)
        Me.txtDD_IC_src.TabIndex = 69
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(10, 73)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(41, 13)
        Me.Label13.TabIndex = 68
        Me.Label13.Text = "Source"
        '
        'OFD_DD_IC
        '
        Me.OFD_DD_IC.FileName = "Source File"
        Me.OFD_DD_IC.Filter = "Excel files (*.xls,*xlsx)|*.xls;*.xlsx"
        '
        'SFD_DD_IC
        '
        Me.SFD_DD_IC.Filter = "Excel files (*.xls,*xlsx)|*.xls;*.xlsx"
        '
        'BWDD_IC
        '
        '
        'frmDD_IC
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.ClientSize = New System.Drawing.Size(408, 215)
        Me.Controls.Add(Me.gbRepType_DD)
        Me.Controls.Add(Me.PicBar_DD_IC)
        Me.Controls.Add(Me.btnBrow_DD_IC_dest)
        Me.Controls.Add(Me.txtDD_IC_dest)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.btnNeu_DD_IC)
        Me.Controls.Add(Me.btnBrow_DD_IC_src)
        Me.Controls.Add(Me.txtDD_IC_src)
        Me.Controls.Add(Me.Label13)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmDD_IC"
        Me.Text = "Distributor Drive Report"
        Me.gbRepType_DD.ResumeLayout(False)
        Me.gbRepType_DD.PerformLayout()
        CType(Me.PicBar_DD_IC, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents gbRepType_DD As System.Windows.Forms.GroupBox
    Friend WithEvents RBDD_Dist As System.Windows.Forms.RadioButton
    Friend WithEvents RBDD_NotDist As System.Windows.Forms.RadioButton
    Friend WithEvents PicBar_DD_IC As System.Windows.Forms.PictureBox
    Friend WithEvents btnBrow_DD_IC_dest As System.Windows.Forms.Button
    Friend WithEvents txtDD_IC_dest As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents btnNeu_DD_IC As System.Windows.Forms.Button
    Friend WithEvents btnBrow_DD_IC_src As System.Windows.Forms.Button
    Friend WithEvents txtDD_IC_src As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents OFD_DD_IC As System.Windows.Forms.OpenFileDialog
    Friend WithEvents SFD_DD_IC As System.Windows.Forms.SaveFileDialog
    Friend WithEvents BWDD_IC As System.ComponentModel.BackgroundWorker
End Class
