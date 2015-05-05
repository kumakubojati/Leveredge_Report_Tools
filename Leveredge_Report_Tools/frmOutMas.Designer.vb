<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmOutMas
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmOutMas))
        Me.PicBar = New System.Windows.Forms.PictureBox()
        Me.btnBrow_OutMasDes = New System.Windows.Forms.Button()
        Me.txtDest_OutMas = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.btnNeuOutMas = New System.Windows.Forms.Button()
        Me.btnbrow_outmassrc = New System.Windows.Forms.Button()
        Me.txtoutmas_src = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.OFD_OutMas = New System.Windows.Forms.OpenFileDialog()
        Me.SFD_OutMas = New System.Windows.Forms.SaveFileDialog()
        Me.BWOutMas = New System.ComponentModel.BackgroundWorker()
        CType(Me.PicBar, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PicBar
        '
        Me.PicBar.Image = CType(resources.GetObject("PicBar.Image"), System.Drawing.Image)
        Me.PicBar.Location = New System.Drawing.Point(7, 65)
        Me.PicBar.Name = "PicBar"
        Me.PicBar.Size = New System.Drawing.Size(80, 80)
        Me.PicBar.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PicBar.TabIndex = 17
        Me.PicBar.TabStop = False
        Me.PicBar.Visible = False
        '
        'btnBrow_OutMasDes
        '
        Me.btnBrow_OutMasDes.BackColor = System.Drawing.Color.Transparent
        Me.btnBrow_OutMasDes.Image = CType(resources.GetObject("btnBrow_OutMasDes.Image"), System.Drawing.Image)
        Me.btnBrow_OutMasDes.Location = New System.Drawing.Point(315, 38)
        Me.btnBrow_OutMasDes.Name = "btnBrow_OutMasDes"
        Me.btnBrow_OutMasDes.Size = New System.Drawing.Size(25, 23)
        Me.btnBrow_OutMasDes.TabIndex = 16
        Me.btnBrow_OutMasDes.UseVisualStyleBackColor = False
        '
        'txtDest_OutMas
        '
        Me.txtDest_OutMas.Location = New System.Drawing.Point(81, 39)
        Me.txtDest_OutMas.Name = "txtDest_OutMas"
        Me.txtDest_OutMas.Size = New System.Drawing.Size(229, 20)
        Me.txtDest_OutMas.TabIndex = 15
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(8, 44)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(60, 13)
        Me.Label2.TabIndex = 14
        Me.Label2.Text = "Destination"
        '
        'btnNeuOutMas
        '
        Me.btnNeuOutMas.Enabled = False
        Me.btnNeuOutMas.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnNeuOutMas.Image = CType(resources.GetObject("btnNeuOutMas.Image"), System.Drawing.Image)
        Me.btnNeuOutMas.Location = New System.Drawing.Point(346, 5)
        Me.btnNeuOutMas.Name = "btnNeuOutMas"
        Me.btnNeuOutMas.Size = New System.Drawing.Size(55, 56)
        Me.btnNeuOutMas.TabIndex = 13
        Me.btnNeuOutMas.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnNeuOutMas.UseVisualStyleBackColor = True
        '
        'btnbrow_outmassrc
        '
        Me.btnbrow_outmassrc.BackColor = System.Drawing.Color.Transparent
        Me.btnbrow_outmassrc.Image = CType(resources.GetObject("btnbrow_outmassrc.Image"), System.Drawing.Image)
        Me.btnbrow_outmassrc.Location = New System.Drawing.Point(315, 4)
        Me.btnbrow_outmassrc.Name = "btnbrow_outmassrc"
        Me.btnbrow_outmassrc.Size = New System.Drawing.Size(25, 23)
        Me.btnbrow_outmassrc.TabIndex = 12
        Me.btnbrow_outmassrc.UseVisualStyleBackColor = False
        '
        'txtoutmas_src
        '
        Me.txtoutmas_src.Location = New System.Drawing.Point(81, 5)
        Me.txtoutmas_src.Name = "txtoutmas_src"
        Me.txtoutmas_src.Size = New System.Drawing.Size(229, 20)
        Me.txtoutmas_src.TabIndex = 11
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(8, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(41, 13)
        Me.Label1.TabIndex = 10
        Me.Label1.Text = "Source"
        '
        'OFD_OutMas
        '
        Me.OFD_OutMas.FileName = "Source File"
        Me.OFD_OutMas.Filter = "Excel files (*.xls,*xlsx)|*.xls;*.xlsx"
        '
        'SFD_OutMas
        '
        Me.SFD_OutMas.Filter = "Excel files (*.xls,*.xlsx)|*.xls;*xlsx"
        '
        'BWOutMas
        '
        '
        'frmOutMas
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.ClientSize = New System.Drawing.Size(408, 154)
        Me.Controls.Add(Me.PicBar)
        Me.Controls.Add(Me.btnBrow_OutMasDes)
        Me.Controls.Add(Me.txtDest_OutMas)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.btnNeuOutMas)
        Me.Controls.Add(Me.btnbrow_outmassrc)
        Me.Controls.Add(Me.txtoutmas_src)
        Me.Controls.Add(Me.Label1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmOutMas"
        Me.Text = "Outlet Master"
        CType(Me.PicBar, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents PicBar As System.Windows.Forms.PictureBox
    Friend WithEvents btnBrow_OutMasDes As System.Windows.Forms.Button
    Friend WithEvents txtDest_OutMas As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnNeuOutMas As System.Windows.Forms.Button
    Friend WithEvents btnbrow_outmassrc As System.Windows.Forms.Button
    Friend WithEvents txtoutmas_src As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents OFD_OutMas As System.Windows.Forms.OpenFileDialog
    Friend WithEvents SFD_OutMas As System.Windows.Forms.SaveFileDialog
    Friend WithEvents BWOutMas As System.ComponentModel.BackgroundWorker
End Class
