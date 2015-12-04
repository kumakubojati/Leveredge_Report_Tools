Public Class frmDBConPass

    
    Private Sub frmDBConPass_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        lbldate.Text = Now.Date.ToString("ddMMyyyy")
    End Sub

    Private Sub txtPass_KeyDown(sender As Object, e As KeyEventArgs) Handles txtPass.KeyDown
        Dim password As String = "uidleveredge" & lbldate.Text
        If e.KeyCode = Keys.Enter Then
            Dim inp_pass As String = txtPass.Text.Trim
            If inp_pass = password Then
                frmChangeDBCon.Show()
                Me.Hide()
            Else
                lblnotif.Text = "Access Denied"
            End If
        End If
    End Sub
End Class