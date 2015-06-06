Imports System.Net
Public NotInheritable Class SSdownload

    'TODO: This form can easily be set as the splash screen for the application by going to the "Application" tab
    '  of the Project Designer ("Properties" under the "Project" menu).


    Private Sub SSdownload_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        frmMain.Hide()
        Dim newfolder As New FolderBrowserDialog
        If newfolder.ShowDialog = Windows.Forms.DialogResult.OK Then
            savepath = newfolder.SelectedPath
        End If
        Control.CheckForIllegalCrossThreadCalls = False
        bwdown.RunWorkerAsync()
    End Sub
    Dim FTPurl As String = "ftp://182.253.21.82/Leveredge Report Tools ver." & frmMain.newversion & ".rar"
    Public savepath As String
    Private Sub bwdown_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles bwdown.DoWork
        Dim buffer(1023) As Byte
        Dim bytesIn As Integer
        Dim totalbytes As Integer
        Dim output As IO.Stream
        Dim filength As Integer

        Try
            Dim ftprequest As FtpWebRequest = DirectCast(WebRequest.Create(FTPurl), FtpWebRequest)
            ftprequest.Credentials = New NetworkCredential("ltrupdater", "ltrupdater123")
            ftprequest.Method = Net.WebRequestMethods.Ftp.GetFileSize
            filength = CInt(ftprequest.GetResponse.ContentLength)
            lblfilesize.Text = filength & " bytes"
        Catch ex As Exception

        End Try

        Try
            Dim ftprequest As FtpWebRequest = DirectCast(WebRequest.Create(FTPurl), FtpWebRequest)
            ftprequest.Credentials = New NetworkCredential("ltrupdater", "ltrupdater123")
            ftprequest.Method = WebRequestMethods.Ftp.DownloadFile
            Dim stream As System.IO.Stream = ftprequest.GetResponse.GetResponseStream
            Dim outfilepath As String = savepath & "\" & IO.Path.GetFileName(FTPurl)
            output = System.IO.File.Create(outfilepath)
            bytesIn = 1
            Do Until bytesIn < 1
                bytesIn = stream.Read(buffer, 0, 1024)
                If bytesIn > 0 Then
                    output.Write(buffer, 0, bytesIn)
                    totalbytes += bytesIn
                    lbldownbyte.Text = totalbytes.ToString & " bytes"
                    If filength > 0 Then
                        Dim perc As Integer = (totalbytes / filength) * 100
                        bwdown.ReportProgress(perc)
                    End If
                End If
            Loop
            output.Close()
            stream.Close()
        Catch ex As Exception
            MessageBox.Show("Error on : " & ex.Message.ToString, "Error")
        End Try
    End Sub

    Private Sub bwdown_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles bwdown.ProgressChanged
        pbdown.Value = e.ProgressPercentage
        lblpercentage.Text = e.ProgressPercentage.ToString & " %"
    End Sub

    Private Sub bwdown_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bwdown.RunWorkerCompleted
        MsgBox("Download Complete")
        Application.Exit()
    End Sub

End Class
