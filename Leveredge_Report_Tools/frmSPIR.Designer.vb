<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSPIR
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmSPIR))
        Me.PicBar_SPIR = New System.Windows.Forms.PictureBox()
        Me.btnBrow_SPIR_dest = New System.Windows.Forms.Button()
        Me.txtSPIR_dest = New System.Windows.Forms.TextBox()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.btnNeu_SPIR = New System.Windows.Forms.Button()
        Me.btnBrow_SPIR_src = New System.Windows.Forms.Button()
        Me.txtSPIR_src = New System.Windows.Forms.TextBox()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.OFD_SPIR = New System.Windows.Forms.OpenFileDialog()
        Me.SFD_SPIR = New System.Windows.Forms.SaveFileDialog()
        Me.BWSPIR = New System.ComponentModel.BackgroundWorker()
        CType(Me.PicBar_SPIR, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PicBar_SPIR
        '
        Me.PicBar_SPIR.Image = CType(resources.GetObject("PicBar_SPIR.Image"), System.Drawing.Image)
        Me.PicBar_SPIR.Location = New System.Drawing.Point(6, 72)
        Me.PicBar_SPIR.Name = "PicBar_SPIR"
        Me.PicBar_SPIR.Size = New System.Drawing.Size(80, 80)
        Me.PicBar_SPIR.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PicBar_SPIR.TabIndex = 49
        Me.PicBar_SPIR.TabStop = False
        Me.PicBar_SPIR.Visible = False
        '
        'btnBrow_SPIR_dest
        '
        Me.btnBrow_SPIR_dest.BackColor = System.Drawing.Color.Transparent
        Me.btnBrow_SPIR_dest.Image = CType(resources.GetObject("btnBrow_SPIR_dest.Image"), System.Drawing.Image)
        Me.btnBrow_SPIR_dest.Location = New System.Drawing.Point(313, 42)
        Me.btnBrow_SPIR_dest.Name = "btnBrow_SPIR_dest"
        Me.btnBrow_SPIR_dest.Size = New System.Drawing.Size(25, 23)
        Me.btnBrow_SPIR_dest.TabIndex = 48
        Me.btnBrow_SPIR_dest.UseVisualStyleBackColor = False
        '
        'txtSPIR_dest
        '
        Me.txtSPIR_dest.Location = New System.Drawing.Point(79, 43)
        Me.txtSPIR_dest.Name = "txtSPIR_dest"
        Me.txtSPIR_dest.Size = New System.Drawing.Size(229, 20)
        Me.txtSPIR_dest.TabIndex = 47
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(6, 48)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(60, 13)
        Me.Label14.TabIndex = 46
        Me.Label14.Text = "Destination"
        '
        'btnNeu_SPIR
        '
        Me.btnNeu_SPIR.Enabled = False
        Me.btnNeu_SPIR.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnNeu_SPIR.Image = CType(resources.GetObject("btnNeu_SPIR.Image"), System.Drawing.Image)
        Me.btnNeu_SPIR.Location = New System.Drawing.Point(344, 9)
        Me.btnNeu_SPIR.Name = "btnNeu_SPIR"
        Me.btnNeu_SPIR.Size = New System.Drawing.Size(55, 56)
        Me.btnNeu_SPIR.TabIndex = 45
        Me.btnNeu_SPIR.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnNeu_SPIR.UseVisualStyleBackColor = True
        '
        'btnBrow_SPIR_src
        '
        Me.btnBrow_SPIR_src.BackColor = System.Drawing.Color.Transparent
        Me.btnBrow_SPIR_src.Image = CType(resources.GetObject("btnBrow_SPIR_src.Image"), System.Drawing.Image)
        Me.btnBrow_SPIR_src.Location = New System.Drawing.Point(313, 8)
        Me.btnBrow_SPIR_src.Name = "btnBrow_SPIR_src"
        Me.btnBrow_SPIR_src.Size = New System.Drawing.Size(25, 23)
        Me.btnBrow_SPIR_src.TabIndex = 44
        Me.btnBrow_SPIR_src.UseVisualStyleBackColor = False
        '
        'txtSPIR_src
        '
        Me.txtSPIR_src.Location = New System.Drawing.Point(79, 9)
        Me.txtSPIR_src.Name = "txtSPIR_src"
        Me.txtSPIR_src.Size = New System.Drawing.Size(229, 20)
        Me.txtSPIR_src.TabIndex = 43
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(6, 12)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(41, 13)
        Me.Label15.TabIndex = 42
        Me.Label15.Text = "Source"
        '
        'OFD_SPIR
        '
        Me.OFD_SPIR.FileName = "Source File"
        Me.OFD_SPIR.Filter = "Excel files (*.xls,*xlsx)|*.xls;*.xlsx"
        '
        'SFD_SPIR
        '
        Me.SFD_SPIR.Filter = "Excel files (*.xls,*.xlsx)|*.xls;*xlsx"
        '
        'BWSPIR
        '
        '
        'frmSPIR
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.ClientSize = New System.Drawing.Size(406, 155)
        Me.Controls.Add(Me.PicBar_SPIR)
        Me.Controls.Add(Me.btnBrow_SPIR_dest)
        Me.Controls.Add(Me.txtSPIR_dest)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.btnNeu_SPIR)
        Me.Controls.Add(Me.btnBrow_SPIR_src)
        Me.Controls.Add(Me.txtSPIR_src)
        Me.Controls.Add(Me.Label15)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmSPIR"
        Me.Text = "Salesman Performance Incentive Report"
        CType(Me.PicBar_SPIR, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents PicBar_SPIR As System.Windows.Forms.PictureBox
    Friend WithEvents btnBrow_SPIR_dest As System.Windows.Forms.Button
    Friend WithEvents txtSPIR_dest As System.Windows.Forms.TextBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents btnNeu_SPIR As System.Windows.Forms.Button
    Friend WithEvents btnBrow_SPIR_src As System.Windows.Forms.Button
    Friend WithEvents txtSPIR_src As System.Windows.Forms.TextBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents OFD_SPIR As System.Windows.Forms.OpenFileDialog
    Friend WithEvents SFD_SPIR As System.Windows.Forms.SaveFileDialog
    Friend WithEvents BWSPIR As System.ComponentModel.BackgroundWorker
End Class
