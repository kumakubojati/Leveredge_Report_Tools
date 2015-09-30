<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmInitDB
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmInitDB))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtCalPath = New System.Windows.Forms.TextBox()
        Me.txtRCFPath = New System.Windows.Forms.TextBox()
        Me.btnBrowCal = New System.Windows.Forms.Button()
        Me.btnBrowRCF = New System.Windows.Forms.Button()
        Me.btnCleanDB = New System.Windows.Forms.Button()
        Me.btnInitDB = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.pbProcess = New System.Windows.Forms.PictureBox()
        Me.cbYear = New System.Windows.Forms.ComboBox()
        Me.OFDCal = New System.Windows.Forms.OpenFileDialog()
        Me.OFDRCF = New System.Windows.Forms.OpenFileDialog()
        Me.BW1 = New System.ComponentModel.BackgroundWorker()
        Me.BW2 = New System.ComponentModel.BackgroundWorker()
        Me.lblProgress = New System.Windows.Forms.Label()
        CType(Me.pbProcess, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(13, 53)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(68, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Calendar File"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(13, 82)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(91, 13)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Route Control File"
        Me.Label2.Visible = False
        '
        'txtCalPath
        '
        Me.txtCalPath.Location = New System.Drawing.Point(122, 50)
        Me.txtCalPath.Name = "txtCalPath"
        Me.txtCalPath.Size = New System.Drawing.Size(253, 20)
        Me.txtCalPath.TabIndex = 2
        '
        'txtRCFPath
        '
        Me.txtRCFPath.Location = New System.Drawing.Point(122, 79)
        Me.txtRCFPath.Name = "txtRCFPath"
        Me.txtRCFPath.Size = New System.Drawing.Size(253, 20)
        Me.txtRCFPath.TabIndex = 3
        Me.txtRCFPath.Visible = False
        '
        'btnBrowCal
        '
        Me.btnBrowCal.Image = CType(resources.GetObject("btnBrowCal.Image"), System.Drawing.Image)
        Me.btnBrowCal.Location = New System.Drawing.Point(382, 48)
        Me.btnBrowCal.Name = "btnBrowCal"
        Me.btnBrowCal.Size = New System.Drawing.Size(30, 23)
        Me.btnBrowCal.TabIndex = 4
        Me.btnBrowCal.UseVisualStyleBackColor = True
        '
        'btnBrowRCF
        '
        Me.btnBrowRCF.Image = CType(resources.GetObject("btnBrowRCF.Image"), System.Drawing.Image)
        Me.btnBrowRCF.Location = New System.Drawing.Point(382, 76)
        Me.btnBrowRCF.Name = "btnBrowRCF"
        Me.btnBrowRCF.Size = New System.Drawing.Size(30, 23)
        Me.btnBrowRCF.TabIndex = 5
        Me.btnBrowRCF.UseVisualStyleBackColor = True
        Me.btnBrowRCF.Visible = False
        '
        'btnCleanDB
        '
        Me.btnCleanDB.Enabled = False
        Me.btnCleanDB.Location = New System.Drawing.Point(13, 149)
        Me.btnCleanDB.Name = "btnCleanDB"
        Me.btnCleanDB.Size = New System.Drawing.Size(75, 23)
        Me.btnCleanDB.TabIndex = 6
        Me.btnCleanDB.Text = "Clean DB"
        Me.btnCleanDB.UseVisualStyleBackColor = True
        '
        'btnInitDB
        '
        Me.btnInitDB.Enabled = False
        Me.btnInitDB.Location = New System.Drawing.Point(337, 149)
        Me.btnInitDB.Name = "btnInitDB"
        Me.btnInitDB.Size = New System.Drawing.Size(75, 23)
        Me.btnInitDB.TabIndex = 7
        Me.btnInitDB.Text = "Initialize DB"
        Me.btnInitDB.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(13, 9)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(55, 13)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "Year Data"
        '
        'pbProcess
        '
        Me.pbProcess.Image = CType(resources.GetObject("pbProcess.Image"), System.Drawing.Image)
        Me.pbProcess.Location = New System.Drawing.Point(174, 117)
        Me.pbProcess.Name = "pbProcess"
        Me.pbProcess.Size = New System.Drawing.Size(85, 83)
        Me.pbProcess.TabIndex = 10
        Me.pbProcess.TabStop = False
        Me.pbProcess.Visible = False
        '
        'cbYear
        '
        Me.cbYear.Enabled = False
        Me.cbYear.FormattingEnabled = True
        Me.cbYear.Items.AddRange(New Object() {"2014", "2015", "2016", "2017", "2018", "2019", "2020"})
        Me.cbYear.Location = New System.Drawing.Point(122, 9)
        Me.cbYear.Name = "cbYear"
        Me.cbYear.Size = New System.Drawing.Size(67, 21)
        Me.cbYear.TabIndex = 11
        '
        'OFDCal
        '
        Me.OFDCal.FileName = "Calendar File"
        Me.OFDCal.Filter = "SQL File|*.sql"
        '
        'OFDRCF
        '
        Me.OFDRCF.FileName = "File Route Control"
        Me.OFDRCF.Filter = "Excel 97-2003|*.xls|Excel 2007-2013|*.xlsx"
        '
        'BW1
        '
        Me.BW1.WorkerReportsProgress = True
        '
        'BW2
        '
        '
        'lblProgress
        '
        Me.lblProgress.Location = New System.Drawing.Point(46, 203)
        Me.lblProgress.Name = "lblProgress"
        Me.lblProgress.Size = New System.Drawing.Size(329, 21)
        Me.lblProgress.TabIndex = 12
        Me.lblProgress.Text = "Label4"
        Me.lblProgress.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.lblProgress.Visible = False
        '
        'frmInitDB
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(421, 227)
        Me.Controls.Add(Me.lblProgress)
        Me.Controls.Add(Me.cbYear)
        Me.Controls.Add(Me.pbProcess)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.btnInitDB)
        Me.Controls.Add(Me.btnCleanDB)
        Me.Controls.Add(Me.btnBrowRCF)
        Me.Controls.Add(Me.btnBrowCal)
        Me.Controls.Add(Me.txtRCFPath)
        Me.Controls.Add(Me.txtCalPath)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmInitDB"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Initialize DB"
        CType(Me.pbProcess, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtCalPath As System.Windows.Forms.TextBox
    Friend WithEvents txtRCFPath As System.Windows.Forms.TextBox
    Friend WithEvents btnBrowCal As System.Windows.Forms.Button
    Friend WithEvents btnBrowRCF As System.Windows.Forms.Button
    Friend WithEvents btnCleanDB As System.Windows.Forms.Button
    Friend WithEvents btnInitDB As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents pbProcess As System.Windows.Forms.PictureBox
    Friend WithEvents cbYear As System.Windows.Forms.ComboBox
    Friend WithEvents OFDCal As System.Windows.Forms.OpenFileDialog
    Friend WithEvents OFDRCF As System.Windows.Forms.OpenFileDialog
    Friend WithEvents BW1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents BW2 As System.ComponentModel.BackgroundWorker
    Friend WithEvents lblProgress As System.Windows.Forms.Label
End Class
