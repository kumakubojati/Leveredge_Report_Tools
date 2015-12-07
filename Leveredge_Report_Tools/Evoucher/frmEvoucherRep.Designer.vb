<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmEvoucherRep
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmEvoucherRep))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.DTPEV = New System.Windows.Forms.DateTimePicker()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TVevoucherMaster = New System.Windows.Forms.TreeView()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.cbRepType = New System.Windows.Forms.ComboBox()
        Me.btnViewRep = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(30, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Date"
        '
        'DTPEV
        '
        Me.DTPEV.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DTPEV.Location = New System.Drawing.Point(133, 7)
        Me.DTPEV.Name = "DTPEV"
        Me.DTPEV.Size = New System.Drawing.Size(97, 20)
        Me.DTPEV.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 37)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(88, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Evoucher Master"
        '
        'TVevoucherMaster
        '
        Me.TVevoucherMaster.HideSelection = False
        Me.TVevoucherMaster.Location = New System.Drawing.Point(133, 34)
        Me.TVevoucherMaster.Name = "TVevoucherMaster"
        Me.TVevoucherMaster.Size = New System.Drawing.Size(316, 215)
        Me.TVevoucherMaster.TabIndex = 3
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 260)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(66, 13)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Report Type"
        '
        'cbRepType
        '
        Me.cbRepType.FormattingEnabled = True
        Me.cbRepType.Items.AddRange(New Object() {"Evoucher Claim", "Evoucher Participant"})
        Me.cbRepType.Location = New System.Drawing.Point(133, 257)
        Me.cbRepType.Name = "cbRepType"
        Me.cbRepType.Size = New System.Drawing.Size(162, 21)
        Me.cbRepType.TabIndex = 5
        '
        'btnViewRep
        '
        Me.btnViewRep.Location = New System.Drawing.Point(342, 255)
        Me.btnViewRep.Name = "btnViewRep"
        Me.btnViewRep.Size = New System.Drawing.Size(107, 23)
        Me.btnViewRep.TabIndex = 6
        Me.btnViewRep.Text = "View Report"
        Me.btnViewRep.UseVisualStyleBackColor = True
        '
        'frmEvoucherRep
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(461, 283)
        Me.Controls.Add(Me.btnViewRep)
        Me.Controls.Add(Me.cbRepType)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.TVevoucherMaster)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.DTPEV)
        Me.Controls.Add(Me.Label1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmEvoucherRep"
        Me.Text = "Evoucher Report Generator"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents DTPEV As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TVevoucherMaster As System.Windows.Forms.TreeView
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cbRepType As System.Windows.Forms.ComboBox
    Friend WithEvents btnViewRep As System.Windows.Forms.Button
End Class
