<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSISR
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmSISR))
        Me.gbSISR = New System.Windows.Forms.GroupBox()
        Me.rbNoTax = New System.Windows.Forms.RadioButton()
        Me.rbWithTax = New System.Windows.Forms.RadioButton()
        Me.PicBar_SISR = New System.Windows.Forms.PictureBox()
        Me.btnBrow_SISR_dest = New System.Windows.Forms.Button()
        Me.txtSISR_dest = New System.Windows.Forms.TextBox()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.btnNeu_SISR = New System.Windows.Forms.Button()
        Me.btnBrow_SISR_src = New System.Windows.Forms.Button()
        Me.txtSISR_src = New System.Windows.Forms.TextBox()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.OFD_SISR = New System.Windows.Forms.OpenFileDialog()
        Me.SFD_SISR = New System.Windows.Forms.SaveFileDialog()
        Me.BWSISR = New System.ComponentModel.BackgroundWorker()
        Me.gbSISR.SuspendLayout()
        CType(Me.PicBar_SISR, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'gbSISR
        '
        Me.gbSISR.Controls.Add(Me.rbNoTax)
        Me.gbSISR.Controls.Add(Me.rbWithTax)
        Me.gbSISR.Location = New System.Drawing.Point(12, 12)
        Me.gbSISR.Name = "gbSISR"
        Me.gbSISR.Size = New System.Drawing.Size(267, 48)
        Me.gbSISR.TabIndex = 0
        Me.gbSISR.TabStop = False
        Me.gbSISR.Text = "Report Type"
        '
        'rbNoTax
        '
        Me.rbNoTax.AutoSize = True
        Me.rbNoTax.Location = New System.Drawing.Point(153, 19)
        Me.rbNoTax.Name = "rbNoTax"
        Me.rbNoTax.Size = New System.Drawing.Size(106, 17)
        Me.rbNoTax.TabIndex = 1
        Me.rbNoTax.TabStop = True
        Me.rbNoTax.Text = "Tanpa No. Pajak"
        Me.rbNoTax.UseVisualStyleBackColor = True
        '
        'rbWithTax
        '
        Me.rbWithTax.AutoSize = True
        Me.rbWithTax.Checked = True
        Me.rbWithTax.Location = New System.Drawing.Point(6, 19)
        Me.rbWithTax.Name = "rbWithTax"
        Me.rbWithTax.Size = New System.Drawing.Size(113, 17)
        Me.rbWithTax.TabIndex = 0
        Me.rbWithTax.TabStop = True
        Me.rbWithTax.Text = "Dengan No. Pajak"
        Me.rbWithTax.UseVisualStyleBackColor = True
        '
        'PicBar_SISR
        '
        Me.PicBar_SISR.Image = CType(resources.GetObject("PicBar_SISR.Image"), System.Drawing.Image)
        Me.PicBar_SISR.Location = New System.Drawing.Point(12, 131)
        Me.PicBar_SISR.Name = "PicBar_SISR"
        Me.PicBar_SISR.Size = New System.Drawing.Size(80, 80)
        Me.PicBar_SISR.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PicBar_SISR.TabIndex = 41
        Me.PicBar_SISR.TabStop = False
        Me.PicBar_SISR.Visible = False
        '
        'btnBrow_SISR_dest
        '
        Me.btnBrow_SISR_dest.BackColor = System.Drawing.Color.Transparent
        Me.btnBrow_SISR_dest.Image = CType(resources.GetObject("btnBrow_SISR_dest.Image"), System.Drawing.Image)
        Me.btnBrow_SISR_dest.Location = New System.Drawing.Point(319, 101)
        Me.btnBrow_SISR_dest.Name = "btnBrow_SISR_dest"
        Me.btnBrow_SISR_dest.Size = New System.Drawing.Size(25, 23)
        Me.btnBrow_SISR_dest.TabIndex = 40
        Me.btnBrow_SISR_dest.UseVisualStyleBackColor = False
        '
        'txtSISR_dest
        '
        Me.txtSISR_dest.Location = New System.Drawing.Point(85, 102)
        Me.txtSISR_dest.Name = "txtSISR_dest"
        Me.txtSISR_dest.Size = New System.Drawing.Size(229, 20)
        Me.txtSISR_dest.TabIndex = 39
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(12, 107)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(60, 13)
        Me.Label14.TabIndex = 38
        Me.Label14.Text = "Destination"
        '
        'btnNeu_SISR
        '
        Me.btnNeu_SISR.Enabled = False
        Me.btnNeu_SISR.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnNeu_SISR.Image = CType(resources.GetObject("btnNeu_SISR.Image"), System.Drawing.Image)
        Me.btnNeu_SISR.Location = New System.Drawing.Point(350, 68)
        Me.btnNeu_SISR.Name = "btnNeu_SISR"
        Me.btnNeu_SISR.Size = New System.Drawing.Size(55, 56)
        Me.btnNeu_SISR.TabIndex = 37
        Me.btnNeu_SISR.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnNeu_SISR.UseVisualStyleBackColor = True
        '
        'btnBrow_SISR_src
        '
        Me.btnBrow_SISR_src.BackColor = System.Drawing.Color.Transparent
        Me.btnBrow_SISR_src.Image = CType(resources.GetObject("btnBrow_SISR_src.Image"), System.Drawing.Image)
        Me.btnBrow_SISR_src.Location = New System.Drawing.Point(319, 67)
        Me.btnBrow_SISR_src.Name = "btnBrow_SISR_src"
        Me.btnBrow_SISR_src.Size = New System.Drawing.Size(25, 23)
        Me.btnBrow_SISR_src.TabIndex = 36
        Me.btnBrow_SISR_src.UseVisualStyleBackColor = False
        '
        'txtSISR_src
        '
        Me.txtSISR_src.Location = New System.Drawing.Point(85, 68)
        Me.txtSISR_src.Name = "txtSISR_src"
        Me.txtSISR_src.Size = New System.Drawing.Size(229, 20)
        Me.txtSISR_src.TabIndex = 35
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(12, 71)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(41, 13)
        Me.Label15.TabIndex = 34
        Me.Label15.Text = "Source"
        '
        'OFD_SISR
        '
        Me.OFD_SISR.FileName = "Source File"
        Me.OFD_SISR.Filter = "Excel files (*.xls,*xlsx)|*.xls;*.xlsx"
        '
        'SFD_SISR
        '
        Me.SFD_SISR.Filter = "Excel files (*.xls,*.xlsx)|*.xls;*xlsx"
        '
        'BWSISR
        '
        '
        'frmSISR
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.ClientSize = New System.Drawing.Size(415, 212)
        Me.Controls.Add(Me.PicBar_SISR)
        Me.Controls.Add(Me.btnBrow_SISR_dest)
        Me.Controls.Add(Me.txtSISR_dest)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.btnNeu_SISR)
        Me.Controls.Add(Me.btnBrow_SISR_src)
        Me.Controls.Add(Me.txtSISR_src)
        Me.Controls.Add(Me.Label15)
        Me.Controls.Add(Me.gbSISR)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmSISR"
        Me.Text = "Summary Invoice And Sales Return Report"
        Me.gbSISR.ResumeLayout(False)
        Me.gbSISR.PerformLayout()
        CType(Me.PicBar_SISR, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents gbSISR As System.Windows.Forms.GroupBox
    Friend WithEvents rbNoTax As System.Windows.Forms.RadioButton
    Friend WithEvents rbWithTax As System.Windows.Forms.RadioButton
    Friend WithEvents PicBar_SISR As System.Windows.Forms.PictureBox
    Friend WithEvents btnBrow_SISR_dest As System.Windows.Forms.Button
    Friend WithEvents txtSISR_dest As System.Windows.Forms.TextBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents btnNeu_SISR As System.Windows.Forms.Button
    Friend WithEvents btnBrow_SISR_src As System.Windows.Forms.Button
    Friend WithEvents txtSISR_src As System.Windows.Forms.TextBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents OFD_SISR As System.Windows.Forms.OpenFileDialog
    Friend WithEvents SFD_SISR As System.Windows.Forms.SaveFileDialog
    Friend WithEvents BWSISR As System.ComponentModel.BackgroundWorker
End Class
