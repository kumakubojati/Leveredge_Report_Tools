<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmProdMas
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmProdMas))
        Me.PicBar_ProdMas = New System.Windows.Forms.PictureBox()
        Me.btnBrowProdMas_dest = New System.Windows.Forms.Button()
        Me.txtProdMas_dest = New System.Windows.Forms.TextBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.btnNeu_ProdMas = New System.Windows.Forms.Button()
        Me.btnBrowProdMas_src = New System.Windows.Forms.Button()
        Me.txtProdMas_src = New System.Windows.Forms.TextBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.BWProdMas = New System.ComponentModel.BackgroundWorker()
        Me.OFD_ProdMas = New System.Windows.Forms.OpenFileDialog()
        Me.SFD_ProdMas = New System.Windows.Forms.SaveFileDialog()
        CType(Me.PicBar_ProdMas, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PicBar_ProdMas
        '
        Me.PicBar_ProdMas.Image = CType(resources.GetObject("PicBar_ProdMas.Image"), System.Drawing.Image)
        Me.PicBar_ProdMas.Location = New System.Drawing.Point(6, 66)
        Me.PicBar_ProdMas.Name = "PicBar_ProdMas"
        Me.PicBar_ProdMas.Size = New System.Drawing.Size(80, 80)
        Me.PicBar_ProdMas.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PicBar_ProdMas.TabIndex = 26
        Me.PicBar_ProdMas.TabStop = False
        Me.PicBar_ProdMas.Visible = False
        '
        'btnBrowProdMas_dest
        '
        Me.btnBrowProdMas_dest.BackColor = System.Drawing.Color.Transparent
        Me.btnBrowProdMas_dest.Image = CType(resources.GetObject("btnBrowProdMas_dest.Image"), System.Drawing.Image)
        Me.btnBrowProdMas_dest.Location = New System.Drawing.Point(314, 39)
        Me.btnBrowProdMas_dest.Name = "btnBrowProdMas_dest"
        Me.btnBrowProdMas_dest.Size = New System.Drawing.Size(25, 23)
        Me.btnBrowProdMas_dest.TabIndex = 25
        Me.btnBrowProdMas_dest.UseVisualStyleBackColor = False
        '
        'txtProdMas_dest
        '
        Me.txtProdMas_dest.Location = New System.Drawing.Point(80, 40)
        Me.txtProdMas_dest.Name = "txtProdMas_dest"
        Me.txtProdMas_dest.Size = New System.Drawing.Size(229, 20)
        Me.txtProdMas_dest.TabIndex = 24
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(7, 45)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(60, 13)
        Me.Label11.TabIndex = 23
        Me.Label11.Text = "Destination"
        '
        'btnNeu_ProdMas
        '
        Me.btnNeu_ProdMas.Enabled = False
        Me.btnNeu_ProdMas.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnNeu_ProdMas.Image = CType(resources.GetObject("btnNeu_ProdMas.Image"), System.Drawing.Image)
        Me.btnNeu_ProdMas.Location = New System.Drawing.Point(345, 6)
        Me.btnNeu_ProdMas.Name = "btnNeu_ProdMas"
        Me.btnNeu_ProdMas.Size = New System.Drawing.Size(55, 56)
        Me.btnNeu_ProdMas.TabIndex = 22
        Me.btnNeu_ProdMas.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnNeu_ProdMas.UseVisualStyleBackColor = True
        '
        'btnBrowProdMas_src
        '
        Me.btnBrowProdMas_src.BackColor = System.Drawing.Color.Transparent
        Me.btnBrowProdMas_src.Image = CType(resources.GetObject("btnBrowProdMas_src.Image"), System.Drawing.Image)
        Me.btnBrowProdMas_src.Location = New System.Drawing.Point(314, 5)
        Me.btnBrowProdMas_src.Name = "btnBrowProdMas_src"
        Me.btnBrowProdMas_src.Size = New System.Drawing.Size(25, 23)
        Me.btnBrowProdMas_src.TabIndex = 21
        Me.btnBrowProdMas_src.UseVisualStyleBackColor = False
        '
        'txtProdMas_src
        '
        Me.txtProdMas_src.Location = New System.Drawing.Point(80, 6)
        Me.txtProdMas_src.Name = "txtProdMas_src"
        Me.txtProdMas_src.Size = New System.Drawing.Size(229, 20)
        Me.txtProdMas_src.TabIndex = 20
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(7, 9)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(41, 13)
        Me.Label12.TabIndex = 19
        Me.Label12.Text = "Source"
        '
        'BWProdMas
        '
        '
        'OFD_ProdMas
        '
        Me.OFD_ProdMas.FileName = "Source File"
        Me.OFD_ProdMas.Filter = "Excel files (*.xls,*xlsx)|*.xls;*.xlsx"
        '
        'SFD_ProdMas
        '
        Me.SFD_ProdMas.Filter = "Excel files (*.xls,*.xlsx)|*.xls;*xlsx"
        '
        'frmProdMas
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.ClientSize = New System.Drawing.Size(410, 150)
        Me.Controls.Add(Me.PicBar_ProdMas)
        Me.Controls.Add(Me.btnBrowProdMas_dest)
        Me.Controls.Add(Me.txtProdMas_dest)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.btnNeu_ProdMas)
        Me.Controls.Add(Me.btnBrowProdMas_src)
        Me.Controls.Add(Me.txtProdMas_src)
        Me.Controls.Add(Me.Label12)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmProdMas"
        Me.Text = "Product Master"
        CType(Me.PicBar_ProdMas, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents PicBar_ProdMas As System.Windows.Forms.PictureBox
    Friend WithEvents btnBrowProdMas_dest As System.Windows.Forms.Button
    Friend WithEvents txtProdMas_dest As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents btnNeu_ProdMas As System.Windows.Forms.Button
    Friend WithEvents btnBrowProdMas_src As System.Windows.Forms.Button
    Friend WithEvents txtProdMas_src As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents BWProdMas As System.ComponentModel.BackgroundWorker
    Friend WithEvents OFD_ProdMas As System.Windows.Forms.OpenFileDialog
    Friend WithEvents SFD_ProdMas As System.Windows.Forms.SaveFileDialog
End Class
