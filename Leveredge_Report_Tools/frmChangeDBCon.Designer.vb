<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmChangeDBCon
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmChangeDBCon))
        Me.TCChangeDB = New System.Windows.Forms.TabControl()
        Me.tbMaster = New System.Windows.Forms.TabPage()
        Me.txtPass_Master = New System.Windows.Forms.TextBox()
        Me.txtUserId_Master = New System.Windows.Forms.TextBox()
        Me.txtDB_Master = New System.Windows.Forms.TextBox()
        Me.txtServer_Master = New System.Windows.Forms.TextBox()
        Me.lblSqlExpress_Master = New System.Windows.Forms.Label()
        Me.btnSubmit_Master = New System.Windows.Forms.Button()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.tbFCS = New System.Windows.Forms.TabPage()
        Me.btnSubmit_FCS = New System.Windows.Forms.Button()
        Me.lblSQLEXPRESS_FCS = New System.Windows.Forms.Label()
        Me.txtPass_FCS = New System.Windows.Forms.TextBox()
        Me.txtUserid_FCS = New System.Windows.Forms.TextBox()
        Me.txtDB_FCS = New System.Windows.Forms.TextBox()
        Me.txtServer_FCS = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TCChangeDB.SuspendLayout()
        Me.tbMaster.SuspendLayout()
        Me.tbFCS.SuspendLayout()
        Me.SuspendLayout()
        '
        'TCChangeDB
        '
        Me.TCChangeDB.Controls.Add(Me.tbMaster)
        Me.TCChangeDB.Controls.Add(Me.tbFCS)
        Me.TCChangeDB.Location = New System.Drawing.Point(2, 2)
        Me.TCChangeDB.Name = "TCChangeDB"
        Me.TCChangeDB.SelectedIndex = 0
        Me.TCChangeDB.Size = New System.Drawing.Size(341, 166)
        Me.TCChangeDB.TabIndex = 0
        '
        'tbMaster
        '
        Me.tbMaster.Controls.Add(Me.txtPass_Master)
        Me.tbMaster.Controls.Add(Me.txtUserId_Master)
        Me.tbMaster.Controls.Add(Me.txtDB_Master)
        Me.tbMaster.Controls.Add(Me.txtServer_Master)
        Me.tbMaster.Controls.Add(Me.lblSqlExpress_Master)
        Me.tbMaster.Controls.Add(Me.btnSubmit_Master)
        Me.tbMaster.Controls.Add(Me.Label8)
        Me.tbMaster.Controls.Add(Me.Label7)
        Me.tbMaster.Controls.Add(Me.Label6)
        Me.tbMaster.Controls.Add(Me.Label5)
        Me.tbMaster.Location = New System.Drawing.Point(4, 22)
        Me.tbMaster.Name = "tbMaster"
        Me.tbMaster.Padding = New System.Windows.Forms.Padding(3)
        Me.tbMaster.Size = New System.Drawing.Size(333, 140)
        Me.tbMaster.TabIndex = 0
        Me.tbMaster.Text = "Master DB"
        Me.tbMaster.UseVisualStyleBackColor = True
        '
        'txtPass_Master
        '
        Me.txtPass_Master.Location = New System.Drawing.Point(74, 88)
        Me.txtPass_Master.Name = "txtPass_Master"
        Me.txtPass_Master.Size = New System.Drawing.Size(145, 20)
        Me.txtPass_Master.TabIndex = 9
        Me.txtPass_Master.UseSystemPasswordChar = True
        '
        'txtUserId_Master
        '
        Me.txtUserId_Master.Location = New System.Drawing.Point(74, 64)
        Me.txtUserId_Master.Name = "txtUserId_Master"
        Me.txtUserId_Master.Size = New System.Drawing.Size(145, 20)
        Me.txtUserId_Master.TabIndex = 8
        '
        'txtDB_Master
        '
        Me.txtDB_Master.Location = New System.Drawing.Point(74, 39)
        Me.txtDB_Master.Name = "txtDB_Master"
        Me.txtDB_Master.Size = New System.Drawing.Size(145, 20)
        Me.txtDB_Master.TabIndex = 7
        '
        'txtServer_Master
        '
        Me.txtServer_Master.Location = New System.Drawing.Point(74, 15)
        Me.txtServer_Master.Name = "txtServer_Master"
        Me.txtServer_Master.Size = New System.Drawing.Size(135, 20)
        Me.txtServer_Master.TabIndex = 6
        '
        'lblSqlExpress_Master
        '
        Me.lblSqlExpress_Master.AutoSize = True
        Me.lblSqlExpress_Master.Location = New System.Drawing.Point(215, 18)
        Me.lblSqlExpress_Master.Name = "lblSqlExpress_Master"
        Me.lblSqlExpress_Master.Size = New System.Drawing.Size(83, 13)
        Me.lblSqlExpress_Master.TabIndex = 5
        Me.lblSqlExpress_Master.Text = "\SQLEXPRESS"
        '
        'btnSubmit_Master
        '
        Me.btnSubmit_Master.Location = New System.Drawing.Point(241, 111)
        Me.btnSubmit_Master.Name = "btnSubmit_Master"
        Me.btnSubmit_Master.Size = New System.Drawing.Size(75, 23)
        Me.btnSubmit_Master.TabIndex = 4
        Me.btnSubmit_Master.Text = "Submit"
        Me.btnSubmit_Master.UseVisualStyleBackColor = True
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(7, 91)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(53, 13)
        Me.Label8.TabIndex = 3
        Me.Label8.Text = "Password"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(7, 67)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(41, 13)
        Me.Label7.TabIndex = 2
        Me.Label7.Text = "User Id"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(7, 42)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(53, 13)
        Me.Label6.TabIndex = 1
        Me.Label6.Text = "Database"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(7, 18)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(38, 13)
        Me.Label5.TabIndex = 0
        Me.Label5.Text = "Server"
        '
        'tbFCS
        '
        Me.tbFCS.Controls.Add(Me.btnSubmit_FCS)
        Me.tbFCS.Controls.Add(Me.lblSQLEXPRESS_FCS)
        Me.tbFCS.Controls.Add(Me.txtPass_FCS)
        Me.tbFCS.Controls.Add(Me.txtUserid_FCS)
        Me.tbFCS.Controls.Add(Me.txtDB_FCS)
        Me.tbFCS.Controls.Add(Me.txtServer_FCS)
        Me.tbFCS.Controls.Add(Me.Label4)
        Me.tbFCS.Controls.Add(Me.Label3)
        Me.tbFCS.Controls.Add(Me.Label2)
        Me.tbFCS.Controls.Add(Me.Label1)
        Me.tbFCS.Location = New System.Drawing.Point(4, 22)
        Me.tbFCS.Name = "tbFCS"
        Me.tbFCS.Padding = New System.Windows.Forms.Padding(3)
        Me.tbFCS.Size = New System.Drawing.Size(333, 140)
        Me.tbFCS.TabIndex = 1
        Me.tbFCS.Text = "FCS DB"
        Me.tbFCS.UseVisualStyleBackColor = True
        '
        'btnSubmit_FCS
        '
        Me.btnSubmit_FCS.Location = New System.Drawing.Point(226, 100)
        Me.btnSubmit_FCS.Name = "btnSubmit_FCS"
        Me.btnSubmit_FCS.Size = New System.Drawing.Size(75, 23)
        Me.btnSubmit_FCS.TabIndex = 9
        Me.btnSubmit_FCS.Text = "Submit"
        Me.btnSubmit_FCS.UseVisualStyleBackColor = True
        '
        'lblSQLEXPRESS_FCS
        '
        Me.lblSQLEXPRESS_FCS.AutoSize = True
        Me.lblSQLEXPRESS_FCS.Location = New System.Drawing.Point(218, 17)
        Me.lblSQLEXPRESS_FCS.Name = "lblSQLEXPRESS_FCS"
        Me.lblSQLEXPRESS_FCS.Size = New System.Drawing.Size(83, 13)
        Me.lblSQLEXPRESS_FCS.TabIndex = 8
        Me.lblSQLEXPRESS_FCS.Text = "\SQLEXPRESS"
        '
        'txtPass_FCS
        '
        Me.txtPass_FCS.Location = New System.Drawing.Point(85, 85)
        Me.txtPass_FCS.Name = "txtPass_FCS"
        Me.txtPass_FCS.Size = New System.Drawing.Size(127, 20)
        Me.txtPass_FCS.TabIndex = 7
        Me.txtPass_FCS.UseSystemPasswordChar = True
        '
        'txtUserid_FCS
        '
        Me.txtUserid_FCS.Location = New System.Drawing.Point(85, 61)
        Me.txtUserid_FCS.Name = "txtUserid_FCS"
        Me.txtUserid_FCS.Size = New System.Drawing.Size(127, 20)
        Me.txtUserid_FCS.TabIndex = 6
        '
        'txtDB_FCS
        '
        Me.txtDB_FCS.Location = New System.Drawing.Point(85, 37)
        Me.txtDB_FCS.Name = "txtDB_FCS"
        Me.txtDB_FCS.Size = New System.Drawing.Size(142, 20)
        Me.txtDB_FCS.TabIndex = 5
        '
        'txtServer_FCS
        '
        Me.txtServer_FCS.Location = New System.Drawing.Point(85, 14)
        Me.txtServer_FCS.Name = "txtServer_FCS"
        Me.txtServer_FCS.Size = New System.Drawing.Size(127, 20)
        Me.txtServer_FCS.TabIndex = 4
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(7, 88)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(53, 13)
        Me.Label4.TabIndex = 3
        Me.Label4.Text = "Password"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(7, 64)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(41, 13)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "User Id"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(7, 40)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(53, 13)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Database"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(7, 17)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(38, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Server"
        '
        'frmChangeDBCon
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(345, 170)
        Me.Controls.Add(Me.TCChangeDB)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmChangeDBCon"
        Me.Text = "Change Database Connection"
        Me.TCChangeDB.ResumeLayout(False)
        Me.tbMaster.ResumeLayout(False)
        Me.tbMaster.PerformLayout()
        Me.tbFCS.ResumeLayout(False)
        Me.tbFCS.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents TCChangeDB As System.Windows.Forms.TabControl
    Friend WithEvents tbMaster As System.Windows.Forms.TabPage
    Friend WithEvents tbFCS As System.Windows.Forms.TabPage
    Friend WithEvents lblSQLEXPRESS_FCS As System.Windows.Forms.Label
    Friend WithEvents txtPass_FCS As System.Windows.Forms.TextBox
    Friend WithEvents txtUserid_FCS As System.Windows.Forms.TextBox
    Friend WithEvents txtDB_FCS As System.Windows.Forms.TextBox
    Friend WithEvents txtServer_FCS As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtPass_Master As System.Windows.Forms.TextBox
    Friend WithEvents txtUserId_Master As System.Windows.Forms.TextBox
    Friend WithEvents txtDB_Master As System.Windows.Forms.TextBox
    Friend WithEvents txtServer_Master As System.Windows.Forms.TextBox
    Friend WithEvents lblSqlExpress_Master As System.Windows.Forms.Label
    Friend WithEvents btnSubmit_Master As System.Windows.Forms.Button
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents btnSubmit_FCS As System.Windows.Forms.Button
End Class
