<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmTarget
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmTarget))
        Me.cbYear = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnFetchData = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.dgDetailTarget = New System.Windows.Forms.DataGridView()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.cbPeriod = New System.Windows.Forms.ComboBox()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.btnUpdate = New System.Windows.Forms.Button()
        Me.txtECO = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtBP = New System.Windows.Forms.TextBox()
        Me.txtLPPC = New System.Windows.Forms.TextBox()
        Me.btnForAll = New System.Windows.Forms.Button()
        CType(Me.dgDetailTarget, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cbYear
        '
        Me.cbYear.FormattingEnabled = True
        Me.cbYear.Location = New System.Drawing.Point(71, 18)
        Me.cbYear.Name = "cbYear"
        Me.cbYear.Size = New System.Drawing.Size(62, 21)
        Me.cbYear.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 21)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(29, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Year"
        '
        'btnFetchData
        '
        Me.btnFetchData.Location = New System.Drawing.Point(15, 75)
        Me.btnFetchData.Name = "btnFetchData"
        Me.btnFetchData.Size = New System.Drawing.Size(81, 23)
        Me.btnFetchData.TabIndex = 2
        Me.btnFetchData.Text = "Fetch Data"
        Me.btnFetchData.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 139)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(68, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Target Detail"
        '
        'dgDetailTarget
        '
        Me.dgDetailTarget.AllowUserToAddRows = False
        Me.dgDetailTarget.AllowUserToDeleteRows = False
        Me.dgDetailTarget.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgDetailTarget.Location = New System.Drawing.Point(12, 156)
        Me.dgDetailTarget.Name = "dgDetailTarget"
        Me.dgDetailTarget.Size = New System.Drawing.Size(386, 177)
        Me.dgDetailTarget.TabIndex = 4
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 51)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(37, 13)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Period"
        '
        'cbPeriod
        '
        Me.cbPeriod.FormattingEnabled = True
        Me.cbPeriod.Location = New System.Drawing.Point(71, 48)
        Me.cbPeriod.Name = "cbPeriod"
        Me.cbPeriod.Size = New System.Drawing.Size(44, 21)
        Me.cbPeriod.TabIndex = 6
        '
        'btnSave
        '
        Me.btnSave.Enabled = False
        Me.btnSave.Location = New System.Drawing.Point(12, 344)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(75, 23)
        Me.btnSave.TabIndex = 7
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Enabled = False
        Me.btnDelete.Location = New System.Drawing.Point(94, 344)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(75, 23)
        Me.btnDelete.TabIndex = 8
        Me.btnDelete.Text = "Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnUpdate
        '
        Me.btnUpdate.Enabled = False
        Me.btnUpdate.Location = New System.Drawing.Point(175, 344)
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.Size = New System.Drawing.Size(75, 23)
        Me.btnUpdate.TabIndex = 9
        Me.btnUpdate.Text = "Update"
        Me.btnUpdate.UseVisualStyleBackColor = True
        '
        'txtECO
        '
        Me.txtECO.Location = New System.Drawing.Point(253, 85)
        Me.txtECO.Name = "txtECO"
        Me.txtECO.Size = New System.Drawing.Size(36, 20)
        Me.txtECO.TabIndex = 10
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(218, 88)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(29, 13)
        Me.Label4.TabIndex = 11
        Me.Label4.Text = "ECO"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(218, 112)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(21, 13)
        Me.Label5.TabIndex = 12
        Me.Label5.Text = "BP"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(218, 137)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(34, 13)
        Me.Label6.TabIndex = 13
        Me.Label6.Text = "LPPC"
        '
        'txtBP
        '
        Me.txtBP.Location = New System.Drawing.Point(253, 109)
        Me.txtBP.Name = "txtBP"
        Me.txtBP.Size = New System.Drawing.Size(36, 20)
        Me.txtBP.TabIndex = 14
        '
        'txtLPPC
        '
        Me.txtLPPC.Location = New System.Drawing.Point(253, 133)
        Me.txtLPPC.Name = "txtLPPC"
        Me.txtLPPC.Size = New System.Drawing.Size(36, 20)
        Me.txtLPPC.TabIndex = 15
        '
        'btnForAll
        '
        Me.btnForAll.BackColor = System.Drawing.SystemColors.Control
        Me.btnForAll.Image = CType(resources.GetObject("btnForAll.Image"), System.Drawing.Image)
        Me.btnForAll.Location = New System.Drawing.Point(295, 84)
        Me.btnForAll.Name = "btnForAll"
        Me.btnForAll.Size = New System.Drawing.Size(33, 69)
        Me.btnForAll.TabIndex = 16
        Me.btnForAll.UseVisualStyleBackColor = False
        '
        'frmTarget
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(410, 379)
        Me.Controls.Add(Me.btnForAll)
        Me.Controls.Add(Me.txtLPPC)
        Me.Controls.Add(Me.txtBP)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.txtECO)
        Me.Controls.Add(Me.btnUpdate)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.cbPeriod)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.dgDetailTarget)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.btnFetchData)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cbYear)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmTarget"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Target"
        CType(Me.dgDetailTarget, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cbYear As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnFetchData As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents dgDetailTarget As System.Windows.Forms.DataGridView
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cbPeriod As System.Windows.Forms.ComboBox
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnUpdate As System.Windows.Forms.Button
    Friend WithEvents txtECO As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtBP As System.Windows.Forms.TextBox
    Friend WithEvents txtLPPC As System.Windows.Forms.TextBox
    Friend WithEvents btnForAll As System.Windows.Forms.Button
End Class
