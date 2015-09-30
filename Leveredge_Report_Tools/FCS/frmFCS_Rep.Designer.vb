<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmFCS_Rep
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmFCS_Rep))
        Me.clbDSR = New System.Windows.Forms.CheckedListBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cbYear = New System.Windows.Forms.ComboBox()
        Me.cbPeriod = New System.Windows.Forms.ComboBox()
        Me.btnViewRep = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'clbDSR
        '
        Me.clbDSR.CheckOnClick = True
        Me.clbDSR.FormattingEnabled = True
        Me.clbDSR.Items.AddRange(New Object() {"Select All"})
        Me.clbDSR.Location = New System.Drawing.Point(236, 10)
        Me.clbDSR.Name = "clbDSR"
        Me.clbDSR.Size = New System.Drawing.Size(211, 349)
        Me.clbDSR.TabIndex = 7
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(13, 13)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(29, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Year"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(13, 41)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(37, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Period"
        '
        'cbYear
        '
        Me.cbYear.Enabled = False
        Me.cbYear.FormattingEnabled = True
        Me.cbYear.Location = New System.Drawing.Point(66, 10)
        Me.cbYear.Name = "cbYear"
        Me.cbYear.Size = New System.Drawing.Size(84, 21)
        Me.cbYear.TabIndex = 3
        '
        'cbPeriod
        '
        Me.cbPeriod.FormattingEnabled = True
        Me.cbPeriod.Location = New System.Drawing.Point(66, 38)
        Me.cbPeriod.Name = "cbPeriod"
        Me.cbPeriod.Size = New System.Drawing.Size(49, 21)
        Me.cbPeriod.TabIndex = 4
        '
        'btnViewRep
        '
        Me.btnViewRep.Location = New System.Drawing.Point(474, 13)
        Me.btnViewRep.Name = "btnViewRep"
        Me.btnViewRep.Size = New System.Drawing.Size(75, 23)
        Me.btnViewRep.TabIndex = 5
        Me.btnViewRep.Text = "View Report"
        Me.btnViewRep.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(204, 10)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(30, 13)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "DSR"
        '
        'frmFCS_Rep
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(557, 369)
        Me.Controls.Add(Me.clbDSR)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.btnViewRep)
        Me.Controls.Add(Me.cbPeriod)
        Me.Controls.Add(Me.cbYear)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmFCS_Rep"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "FCS Report"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cbYear As System.Windows.Forms.ComboBox
    Friend WithEvents cbPeriod As System.Windows.Forms.ComboBox
    Friend WithEvents btnViewRep As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents clbDSR As System.Windows.Forms.CheckedListBox
End Class
