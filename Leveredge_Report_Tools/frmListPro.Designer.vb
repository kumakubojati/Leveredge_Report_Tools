<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmListPro
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmListPro))
        Me.PicBar_Promo = New System.Windows.Forms.PictureBox()
        Me.btnBrowPromo_dest = New System.Windows.Forms.Button()
        Me.txtProm_dest = New System.Windows.Forms.TextBox()
        Me.txtPromo_src = New System.Windows.Forms.TextBox()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.btnNeu_Promo = New System.Windows.Forms.Button()
        Me.btnBrowProm_src = New System.Windows.Forms.Button()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.BWPro = New System.ComponentModel.BackgroundWorker()
        Me.OFD_Pro = New System.Windows.Forms.OpenFileDialog()
        Me.SFD_Pro = New System.Windows.Forms.SaveFileDialog()
        CType(Me.PicBar_Promo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PicBar_Promo
        '
        Me.PicBar_Promo.Image = CType(resources.GetObject("PicBar_Promo.Image"), System.Drawing.Image)
        Me.PicBar_Promo.Location = New System.Drawing.Point(6, 69)
        Me.PicBar_Promo.Name = "PicBar_Promo"
        Me.PicBar_Promo.Size = New System.Drawing.Size(80, 80)
        Me.PicBar_Promo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PicBar_Promo.TabIndex = 17
        Me.PicBar_Promo.TabStop = False
        Me.PicBar_Promo.Visible = False
        '
        'btnBrowPromo_dest
        '
        Me.btnBrowPromo_dest.Image = CType(resources.GetObject("btnBrowPromo_dest.Image"), System.Drawing.Image)
        Me.btnBrowPromo_dest.Location = New System.Drawing.Point(311, 41)
        Me.btnBrowPromo_dest.Name = "btnBrowPromo_dest"
        Me.btnBrowPromo_dest.Size = New System.Drawing.Size(26, 23)
        Me.btnBrowPromo_dest.TabIndex = 20
        Me.btnBrowPromo_dest.UseVisualStyleBackColor = True
        '
        'txtProm_dest
        '
        Me.txtProm_dest.Location = New System.Drawing.Point(77, 42)
        Me.txtProm_dest.Name = "txtProm_dest"
        Me.txtProm_dest.Size = New System.Drawing.Size(229, 20)
        Me.txtProm_dest.TabIndex = 19
        '
        'txtPromo_src
        '
        Me.txtPromo_src.Location = New System.Drawing.Point(77, 8)
        Me.txtPromo_src.Name = "txtPromo_src"
        Me.txtPromo_src.Size = New System.Drawing.Size(229, 20)
        Me.txtPromo_src.TabIndex = 14
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(7, 47)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(60, 13)
        Me.Label16.TabIndex = 18
        Me.Label16.Text = "Destination"
        '
        'btnNeu_Promo
        '
        Me.btnNeu_Promo.Enabled = False
        Me.btnNeu_Promo.Image = CType(resources.GetObject("btnNeu_Promo.Image"), System.Drawing.Image)
        Me.btnNeu_Promo.Location = New System.Drawing.Point(344, 5)
        Me.btnNeu_Promo.Name = "btnNeu_Promo"
        Me.btnNeu_Promo.Size = New System.Drawing.Size(54, 59)
        Me.btnNeu_Promo.TabIndex = 16
        Me.btnNeu_Promo.UseVisualStyleBackColor = True
        '
        'btnBrowProm_src
        '
        Me.btnBrowProm_src.Image = CType(resources.GetObject("btnBrowProm_src.Image"), System.Drawing.Image)
        Me.btnBrowProm_src.Location = New System.Drawing.Point(311, 7)
        Me.btnBrowProm_src.Name = "btnBrowProm_src"
        Me.btnBrowProm_src.Size = New System.Drawing.Size(26, 23)
        Me.btnBrowProm_src.TabIndex = 15
        Me.btnBrowProm_src.UseVisualStyleBackColor = True
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(7, 13)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(41, 13)
        Me.Label17.TabIndex = 13
        Me.Label17.Text = "Source"
        '
        'BWPro
        '
        '
        'OFD_Pro
        '
        Me.OFD_Pro.FileName = "Source File"
        Me.OFD_Pro.Filter = "Excel files (*.xls,*xlsx)|*.xls;*.xlsx"
        '
        'SFD_Pro
        '
        Me.SFD_Pro.Filter = "Excel files (*.xls,*.xlsx)|*.xls;*xlsx"
        '
        'frmListPro
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.ClientSize = New System.Drawing.Size(403, 154)
        Me.Controls.Add(Me.PicBar_Promo)
        Me.Controls.Add(Me.btnBrowPromo_dest)
        Me.Controls.Add(Me.txtProm_dest)
        Me.Controls.Add(Me.txtPromo_src)
        Me.Controls.Add(Me.Label16)
        Me.Controls.Add(Me.btnNeu_Promo)
        Me.Controls.Add(Me.btnBrowProm_src)
        Me.Controls.Add(Me.Label17)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmListPro"
        Me.Text = "List Of Promotion"
        CType(Me.PicBar_Promo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents PicBar_Promo As System.Windows.Forms.PictureBox
    Friend WithEvents btnBrowPromo_dest As System.Windows.Forms.Button
    Friend WithEvents txtProm_dest As System.Windows.Forms.TextBox
    Friend WithEvents txtPromo_src As System.Windows.Forms.TextBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents btnNeu_Promo As System.Windows.Forms.Button
    Friend WithEvents btnBrowProm_src As System.Windows.Forms.Button
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents BWPro As System.ComponentModel.BackgroundWorker
    Friend WithEvents OFD_Pro As System.Windows.Forms.OpenFileDialog
    Friend WithEvents SFD_Pro As System.Windows.Forms.SaveFileDialog
End Class
