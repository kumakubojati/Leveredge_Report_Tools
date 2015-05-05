<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmDSR
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmDSR))
        Me.PicBar_DistStock = New System.Windows.Forms.PictureBox()
        Me.btnBrowDistStock_Dest = New System.Windows.Forms.Button()
        Me.txtDistStock_Dest = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.btnNeu_DistStock = New System.Windows.Forms.Button()
        Me.btnBrowDistStock_src = New System.Windows.Forms.Button()
        Me.txtDistStock_src = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.OFD_DSR = New System.Windows.Forms.OpenFileDialog()
        Me.SFD_DSR = New System.Windows.Forms.SaveFileDialog()
        Me.BWDSR = New System.ComponentModel.BackgroundWorker()
        CType(Me.PicBar_DistStock, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PicBar_DistStock
        '
        Me.PicBar_DistStock.Image = CType(resources.GetObject("PicBar_DistStock.Image"), System.Drawing.Image)
        Me.PicBar_DistStock.Location = New System.Drawing.Point(8, 76)
        Me.PicBar_DistStock.Name = "PicBar_DistStock"
        Me.PicBar_DistStock.Size = New System.Drawing.Size(80, 80)
        Me.PicBar_DistStock.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PicBar_DistStock.TabIndex = 17
        Me.PicBar_DistStock.TabStop = False
        Me.PicBar_DistStock.Visible = False
        '
        'btnBrowDistStock_Dest
        '
        Me.btnBrowDistStock_Dest.BackColor = System.Drawing.Color.Transparent
        Me.btnBrowDistStock_Dest.Image = CType(resources.GetObject("btnBrowDistStock_Dest.Image"), System.Drawing.Image)
        Me.btnBrowDistStock_Dest.Location = New System.Drawing.Point(315, 46)
        Me.btnBrowDistStock_Dest.Name = "btnBrowDistStock_Dest"
        Me.btnBrowDistStock_Dest.Size = New System.Drawing.Size(25, 23)
        Me.btnBrowDistStock_Dest.TabIndex = 16
        Me.btnBrowDistStock_Dest.UseVisualStyleBackColor = False
        '
        'txtDistStock_Dest
        '
        Me.txtDistStock_Dest.Location = New System.Drawing.Point(81, 47)
        Me.txtDistStock_Dest.Name = "txtDistStock_Dest"
        Me.txtDistStock_Dest.Size = New System.Drawing.Size(229, 20)
        Me.txtDistStock_Dest.TabIndex = 15
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(8, 52)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(60, 13)
        Me.Label8.TabIndex = 14
        Me.Label8.Text = "Destination"
        '
        'btnNeu_DistStock
        '
        Me.btnNeu_DistStock.Enabled = False
        Me.btnNeu_DistStock.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnNeu_DistStock.Image = CType(resources.GetObject("btnNeu_DistStock.Image"), System.Drawing.Image)
        Me.btnNeu_DistStock.Location = New System.Drawing.Point(346, 13)
        Me.btnNeu_DistStock.Name = "btnNeu_DistStock"
        Me.btnNeu_DistStock.Size = New System.Drawing.Size(55, 56)
        Me.btnNeu_DistStock.TabIndex = 13
        Me.btnNeu_DistStock.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnNeu_DistStock.UseVisualStyleBackColor = True
        '
        'btnBrowDistStock_src
        '
        Me.btnBrowDistStock_src.BackColor = System.Drawing.Color.Transparent
        Me.btnBrowDistStock_src.Image = CType(resources.GetObject("btnBrowDistStock_src.Image"), System.Drawing.Image)
        Me.btnBrowDistStock_src.Location = New System.Drawing.Point(315, 12)
        Me.btnBrowDistStock_src.Name = "btnBrowDistStock_src"
        Me.btnBrowDistStock_src.Size = New System.Drawing.Size(25, 23)
        Me.btnBrowDistStock_src.TabIndex = 12
        Me.btnBrowDistStock_src.UseVisualStyleBackColor = False
        '
        'txtDistStock_src
        '
        Me.txtDistStock_src.Location = New System.Drawing.Point(81, 13)
        Me.txtDistStock_src.Name = "txtDistStock_src"
        Me.txtDistStock_src.Size = New System.Drawing.Size(229, 20)
        Me.txtDistStock_src.TabIndex = 11
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(8, 16)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(41, 13)
        Me.Label9.TabIndex = 10
        Me.Label9.Text = "Source"
        '
        'OFD_DSR
        '
        Me.OFD_DSR.FileName = "Source File"
        Me.OFD_DSR.Filter = "Excel files (*.xls,*xlsx)|*.xls;*.xlsx"
        '
        'SFD_DSR
        '
        Me.SFD_DSR.Filter = "Excel files (*.xls,*.xlsx)|*.xls;*xlsx"
        '
        'BWDSR
        '
        '
        'frmDSR
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.ClientSize = New System.Drawing.Size(406, 160)
        Me.Controls.Add(Me.PicBar_DistStock)
        Me.Controls.Add(Me.btnBrowDistStock_Dest)
        Me.Controls.Add(Me.txtDistStock_Dest)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.btnNeu_DistStock)
        Me.Controls.Add(Me.btnBrowDistStock_src)
        Me.Controls.Add(Me.txtDistStock_src)
        Me.Controls.Add(Me.Label9)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmDSR"
        Me.Text = "Distributor Stock Report"
        CType(Me.PicBar_DistStock, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents PicBar_DistStock As System.Windows.Forms.PictureBox
    Friend WithEvents btnBrowDistStock_Dest As System.Windows.Forms.Button
    Friend WithEvents txtDistStock_Dest As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents btnNeu_DistStock As System.Windows.Forms.Button
    Friend WithEvents btnBrowDistStock_src As System.Windows.Forms.Button
    Friend WithEvents txtDistStock_src As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents OFD_DSR As System.Windows.Forms.OpenFileDialog
    Friend WithEvents SFD_DSR As System.Windows.Forms.SaveFileDialog
    Friend WithEvents BWDSR As System.ComponentModel.BackgroundWorker
End Class
