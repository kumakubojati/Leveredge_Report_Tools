<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmCNI_IC
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmCNI_IC))
        Me.gbRepType_SPR = New System.Windows.Forms.GroupBox()
        Me.RBCNI_Detail = New System.Windows.Forms.RadioButton()
        Me.RBCNI_Rekap = New System.Windows.Forms.RadioButton()
        Me.PicBar_CNI_IC = New System.Windows.Forms.PictureBox()
        Me.btnBrow_CNI_IC_dest = New System.Windows.Forms.Button()
        Me.txtCNI_IC_dest = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.btnNeu_CNI_IC = New System.Windows.Forms.Button()
        Me.btnBrow_CNI_IC_src = New System.Windows.Forms.Button()
        Me.txtCNI_IC_src = New System.Windows.Forms.TextBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.OFD_CNI_IC = New System.Windows.Forms.OpenFileDialog()
        Me.SFD_CNI_IC = New System.Windows.Forms.SaveFileDialog()
        Me.BWCNI_IC = New System.ComponentModel.BackgroundWorker()
        Me.gbRepType_SPR.SuspendLayout()
        CType(Me.PicBar_CNI_IC, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'gbRepType_SPR
        '
        Me.gbRepType_SPR.Controls.Add(Me.RBCNI_Detail)
        Me.gbRepType_SPR.Controls.Add(Me.RBCNI_Rekap)
        Me.gbRepType_SPR.Location = New System.Drawing.Point(8, 10)
        Me.gbRepType_SPR.Name = "gbRepType_SPR"
        Me.gbRepType_SPR.Size = New System.Drawing.Size(191, 53)
        Me.gbRepType_SPR.TabIndex = 67
        Me.gbRepType_SPR.TabStop = False
        Me.gbRepType_SPR.Text = "Report Type"
        '
        'RBCNI_Detail
        '
        Me.RBCNI_Detail.AutoSize = True
        Me.RBCNI_Detail.Checked = True
        Me.RBCNI_Detail.Location = New System.Drawing.Point(7, 22)
        Me.RBCNI_Detail.Name = "RBCNI_Detail"
        Me.RBCNI_Detail.Size = New System.Drawing.Size(52, 17)
        Me.RBCNI_Detail.TabIndex = 1
        Me.RBCNI_Detail.TabStop = True
        Me.RBCNI_Detail.Text = "Detail"
        Me.RBCNI_Detail.UseVisualStyleBackColor = True
        '
        'RBCNI_Rekap
        '
        Me.RBCNI_Rekap.AutoSize = True
        Me.RBCNI_Rekap.Location = New System.Drawing.Point(82, 22)
        Me.RBCNI_Rekap.Name = "RBCNI_Rekap"
        Me.RBCNI_Rekap.Size = New System.Drawing.Size(68, 17)
        Me.RBCNI_Rekap.TabIndex = 0
        Me.RBCNI_Rekap.Text = "Summary"
        Me.RBCNI_Rekap.UseVisualStyleBackColor = True
        '
        'PicBar_CNI_IC
        '
        Me.PicBar_CNI_IC.Image = CType(resources.GetObject("PicBar_CNI_IC.Image"), System.Drawing.Image)
        Me.PicBar_CNI_IC.Location = New System.Drawing.Point(9, 129)
        Me.PicBar_CNI_IC.Name = "PicBar_CNI_IC"
        Me.PicBar_CNI_IC.Size = New System.Drawing.Size(80, 80)
        Me.PicBar_CNI_IC.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PicBar_CNI_IC.TabIndex = 63
        Me.PicBar_CNI_IC.TabStop = False
        Me.PicBar_CNI_IC.Visible = False
        '
        'btnBrow_CNI_IC_dest
        '
        Me.btnBrow_CNI_IC_dest.Image = CType(resources.GetObject("btnBrow_CNI_IC_dest.Image"), System.Drawing.Image)
        Me.btnBrow_CNI_IC_dest.Location = New System.Drawing.Point(314, 101)
        Me.btnBrow_CNI_IC_dest.Name = "btnBrow_CNI_IC_dest"
        Me.btnBrow_CNI_IC_dest.Size = New System.Drawing.Size(26, 23)
        Me.btnBrow_CNI_IC_dest.TabIndex = 66
        Me.btnBrow_CNI_IC_dest.UseVisualStyleBackColor = True
        '
        'txtCNI_IC_dest
        '
        Me.txtCNI_IC_dest.Location = New System.Drawing.Point(80, 102)
        Me.txtCNI_IC_dest.Name = "txtCNI_IC_dest"
        Me.txtCNI_IC_dest.Size = New System.Drawing.Size(229, 20)
        Me.txtCNI_IC_dest.TabIndex = 65
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(9, 107)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(60, 13)
        Me.Label7.TabIndex = 64
        Me.Label7.Text = "Destination"
        '
        'btnNeu_CNI_IC
        '
        Me.btnNeu_CNI_IC.Enabled = False
        Me.btnNeu_CNI_IC.Image = CType(resources.GetObject("btnNeu_CNI_IC.Image"), System.Drawing.Image)
        Me.btnNeu_CNI_IC.Location = New System.Drawing.Point(347, 65)
        Me.btnNeu_CNI_IC.Name = "btnNeu_CNI_IC"
        Me.btnNeu_CNI_IC.Size = New System.Drawing.Size(54, 59)
        Me.btnNeu_CNI_IC.TabIndex = 62
        Me.btnNeu_CNI_IC.UseVisualStyleBackColor = True
        '
        'btnBrow_CNI_IC_src
        '
        Me.btnBrow_CNI_IC_src.Image = CType(resources.GetObject("btnBrow_CNI_IC_src.Image"), System.Drawing.Image)
        Me.btnBrow_CNI_IC_src.Location = New System.Drawing.Point(314, 67)
        Me.btnBrow_CNI_IC_src.Name = "btnBrow_CNI_IC_src"
        Me.btnBrow_CNI_IC_src.Size = New System.Drawing.Size(26, 23)
        Me.btnBrow_CNI_IC_src.TabIndex = 61
        Me.btnBrow_CNI_IC_src.UseVisualStyleBackColor = True
        '
        'txtCNI_IC_src
        '
        Me.txtCNI_IC_src.Location = New System.Drawing.Point(80, 68)
        Me.txtCNI_IC_src.Name = "txtCNI_IC_src"
        Me.txtCNI_IC_src.Size = New System.Drawing.Size(229, 20)
        Me.txtCNI_IC_src.TabIndex = 60
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(10, 73)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(41, 13)
        Me.Label13.TabIndex = 59
        Me.Label13.Text = "Source"
        '
        'OFD_CNI_IC
        '
        Me.OFD_CNI_IC.FileName = "Source File"
        Me.OFD_CNI_IC.Filter = "Excel files (*.xls,*xlsx)|*.xls;*.xlsx"
        '
        'SFD_CNI_IC
        '
        Me.SFD_CNI_IC.Filter = "Excel files (*.xls,*xlsx)|*.xls;*.xlsx"
        '
        'BWCNI_IC
        '
        '
        'frmCNI_IC
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.ClientSize = New System.Drawing.Size(410, 215)
        Me.Controls.Add(Me.gbRepType_SPR)
        Me.Controls.Add(Me.PicBar_CNI_IC)
        Me.Controls.Add(Me.btnBrow_CNI_IC_dest)
        Me.Controls.Add(Me.txtCNI_IC_dest)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.btnNeu_CNI_IC)
        Me.Controls.Add(Me.btnBrow_CNI_IC_src)
        Me.Controls.Add(Me.txtCNI_IC_src)
        Me.Controls.Add(Me.Label13)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmCNI_IC"
        Me.Text = "Cabinet Net Increase Report"
        Me.gbRepType_SPR.ResumeLayout(False)
        Me.gbRepType_SPR.PerformLayout()
        CType(Me.PicBar_CNI_IC, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents gbRepType_SPR As System.Windows.Forms.GroupBox
    Friend WithEvents RBCNI_Detail As System.Windows.Forms.RadioButton
    Friend WithEvents RBCNI_Rekap As System.Windows.Forms.RadioButton
    Friend WithEvents PicBar_CNI_IC As System.Windows.Forms.PictureBox
    Friend WithEvents btnBrow_CNI_IC_dest As System.Windows.Forms.Button
    Friend WithEvents txtCNI_IC_dest As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents btnNeu_CNI_IC As System.Windows.Forms.Button
    Friend WithEvents btnBrow_CNI_IC_src As System.Windows.Forms.Button
    Friend WithEvents txtCNI_IC_src As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents OFD_CNI_IC As System.Windows.Forms.OpenFileDialog
    Friend WithEvents SFD_CNI_IC As System.Windows.Forms.SaveFileDialog
    Friend WithEvents BWCNI_IC As System.ComponentModel.BackgroundWorker
End Class
