<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmUpdater
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmUpdater))
        Me.lblNotif_Updater = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.WebBrow_Updater = New System.Windows.Forms.WebBrowser()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.btnSkipVersion = New System.Windows.Forms.Button()
        Me.btnUpdateNow = New System.Windows.Forms.Button()
        Me.BWDownloadFile = New System.ComponentModel.BackgroundWorker()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lblNotif_Updater
        '
        Me.lblNotif_Updater.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblNotif_Updater.Location = New System.Drawing.Point(90, 36)
        Me.lblNotif_Updater.Name = "lblNotif_Updater"
        Me.lblNotif_Updater.Size = New System.Drawing.Size(523, 23)
        Me.lblNotif_Updater.TabIndex = 0
        Me.lblNotif_Updater.Text = "Label1"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 95)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(72, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Release Note"
        '
        'WebBrow_Updater
        '
        Me.WebBrow_Updater.Location = New System.Drawing.Point(12, 115)
        Me.WebBrow_Updater.MinimumSize = New System.Drawing.Size(20, 20)
        Me.WebBrow_Updater.Name = "WebBrow_Updater"
        Me.WebBrow_Updater.Size = New System.Drawing.Size(793, 363)
        Me.WebBrow_Updater.TabIndex = 2
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(12, 12)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(72, 65)
        Me.PictureBox1.TabIndex = 3
        Me.PictureBox1.TabStop = False
        '
        'btnSkipVersion
        '
        Me.btnSkipVersion.Location = New System.Drawing.Point(340, 484)
        Me.btnSkipVersion.Name = "btnSkipVersion"
        Me.btnSkipVersion.Size = New System.Drawing.Size(214, 32)
        Me.btnSkipVersion.TabIndex = 4
        Me.btnSkipVersion.Text = "Skip This Version"
        Me.btnSkipVersion.UseVisualStyleBackColor = True
        '
        'btnUpdateNow
        '
        Me.btnUpdateNow.Location = New System.Drawing.Point(576, 484)
        Me.btnUpdateNow.Name = "btnUpdateNow"
        Me.btnUpdateNow.Size = New System.Drawing.Size(224, 32)
        Me.btnUpdateNow.TabIndex = 5
        Me.btnUpdateNow.Text = "Update Now"
        Me.btnUpdateNow.UseVisualStyleBackColor = True
        '
        'BWDownloadFile
        '
        '
        'frmUpdater
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(812, 528)
        Me.Controls.Add(Me.btnUpdateNow)
        Me.Controls.Add(Me.btnSkipVersion)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.WebBrow_Updater)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.lblNotif_Updater)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmUpdater"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Leveredge Report Tools Updater"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblNotif_Updater As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents WebBrow_Updater As System.Windows.Forms.WebBrowser
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents btnSkipVersion As System.Windows.Forms.Button
    Friend WithEvents btnUpdateNow As System.Windows.Forms.Button
    Friend WithEvents BWDownloadFile As System.ComponentModel.BackgroundWorker

End Class
