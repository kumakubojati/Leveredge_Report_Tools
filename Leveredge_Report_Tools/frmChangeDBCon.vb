Imports System.Windows.Forms
Public Class frmChangeDBCon

    Private Sub Load_MySetting_DBConnection()
        Dim sqlSetting As String = My.Settings("MasterDBCon").ToString
        Dim fcssetting As String = My.Settings("FCSDBCon").ToString

        Dim sqlsplit = sqlSetting.Split(";")
        Dim sqlsplit_server1 = sqlsplit(0).Split("=")
        Dim sqlsplit_server2 = sqlsplit_server1(1).Split("\")
        Dim sqlsplit_db = sqlsplit(1).Split("=")
        Dim sqlsplit_userid = sqlsplit(2).Split("=")
        Dim sqlsplit_pass = sqlsplit(3).Split("=")

        Dim fcssplit = fcssetting.Split(";")
        Dim fcssplit_server1 = fcssplit(0).Split("=")
        Dim fcssplit_server2 = fcssplit_server1(1).Split("\")
        Dim fcssplit_db = fcssplit(1).Split("=")
        Dim fcssplit_userid = fcssplit(2).Split("=")
        Dim fcssplit_pass = fcssplit(3).Split("=")

        txtServer_Master.Text = sqlsplit_server2(0).ToString
        txtDB_Master.Text = sqlsplit_db(1).ToString
        txtUserId_Master.Text = sqlsplit_userid(1).ToString
        txtPass_Master.Text = sqlsplit_pass(1).ToString

        txtServer_FCS.Text = fcssplit_server2(0).ToString
        txtDB_FCS.Text = fcssplit_db(1).ToString
        txtuserid_FCS.Text = fcssplit_userid(1).ToString
        txtPass_FCS.Text = fcssplit_pass(1).ToString
    End Sub

    Private Sub frmChangeDBCon_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Load_MySetting_DBConnection()
    End Sub

    Private Sub btnSubmit_Master_Click(sender As Object, e As EventArgs) Handles btnSubmit_Master.Click
        Try
            If txtServer_Master.Text Is Nothing Then
                MessageBox.Show("Should define Server", "Caution", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                txtServer_Master.SelectAll()
                txtServer_Master.Focus()
                Exit Sub
            ElseIf txtDB_Master.Text Is Nothing Then
                MessageBox.Show("Should define Database", "Caution", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                txtDB_Master.SelectAll()
                txtDB_Master.Focus()
                Exit Sub
            ElseIf txtUserId_Master.Text Is Nothing Then
                MessageBox.Show("Should define User Id", "Caution", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                txtUserId_Master.SelectAll()
                txtUserId_Master.Focus()
            ElseIf txtPass_Master.Text Is Nothing Then
                MessageBox.Show("Should define Password", "Caution", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                txtDB_Master.SelectAll()
                txtDB_Master.Focus()
                Exit Sub
            Else
                Dim sqlsetting_string As String
                sqlsetting_string = "Server=" & txtServer_Master.Text & lblSqlExpress_Master.Text & "; Database=" _
                    & txtDB_Master.Text & "; User Id=" & txtUserId_Master.Text & "; Password=" & txtPass_Master.Text & ";"
                My.Settings.Item("MasterDBCon") = sqlsetting_string
                My.Settings.Save()
                MessageBox.Show("Submit Succes")
                Load_MySetting_DBConnection()
            End If
        Catch ex As Exception
            MessageBox.Show("Error On Submit Db Connection Master: " & ex.Message.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnSubmit_FCS_Click(sender As Object, e As EventArgs) Handles btnSubmit_FCS.Click
        Try
            If txtServer_Master.Text Is Nothing Then
                MessageBox.Show("Should define Server", "Caution", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                txtServer_Master.SelectAll()
                txtServer_Master.Focus()
                Exit Sub
            ElseIf txtDB_Master.Text Is Nothing Then
                MessageBox.Show("Should define Database", "Caution", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                txtDB_Master.SelectAll()
                txtDB_Master.Focus()
                Exit Sub
            ElseIf txtUserId_Master.Text Is Nothing Then
                MessageBox.Show("Should define User Id", "Caution", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                txtUserId_Master.SelectAll()
                txtUserId_Master.Focus()
            ElseIf txtPass_Master.Text Is Nothing Then
                MessageBox.Show("Should define Password", "Caution", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                txtDB_Master.SelectAll()
                txtDB_Master.Focus()
                Exit Sub
            Else
                Dim FCSsetting_string As String
                FCSsetting_string = "Server=" & txtServer_FCS.Text & lblSQLEXPRESS_FCS.Text & "; Database=" _
                    & txtDB_FCS.Text & "; User Id=" & txtUserid_FCS.Text & "; Password=" & txtPass_FCS.Text & ";"
                My.Settings.Item("FCSDBCon") = FCSsetting_string
                My.Settings.Save()
                MessageBox.Show("Submit Succes")
                Application.Restart()
            End If
        Catch ex As Exception
            MessageBox.Show("Error Submit Db Connection FCS: " & ex.Message.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class