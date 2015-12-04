Imports System
Imports System.Net
Imports System.IO
Imports System.Linq
Imports System.Globalization
Imports System.Reflection
Imports System.Diagnostics
Imports System.Xml
Imports NUnrar.Archive
Imports System.Collections
Public Class frmUpdater
    Public Property TitleVal As String
    Public Property CurrentVersion As String
    Public Property NewVersion As String
    Public Property xmlAddress As String = "http://www.uid-leveredge.cf/LRT/AppCast.xml"

    Private Sub ReadTitle()
        ' Read XML new version title
        Dim xmlReader As New XmlDocument
        xmlReader.Load(xmlAddress)
        Dim titles = xmlReader.GetElementsByTagName("title")
        For Each title As XmlElement In titles
            TitleVal = title.InnerText
        Next
        lblNotif_Updater.Text = TitleVal
    End Sub

    Private Sub LoadReleaseNote()
        WebBrow_Updater.AllowWebBrowserDrop = False
        'WebBrow_Updater.Url = New Uri("http://www.uid-leveredge.cf/LRT/releasenotes.html")
        WebBrow_Updater.Navigate("http://www.uid-leveredge.cf/LRT/releasenotes.html")
    End Sub

    Private Sub frmUpdater_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ReadTitle()
        LoadReleaseNote()
    End Sub

    Private Sub BWDownloadFile_DoWork(sender As Object, e As ComponentModel.DoWorkEventArgs) Handles BWDownloadFile.DoWork
        Dim filename As String
        Dim urlfile As New ArrayList

        ' Read Local File XML in folder Downloaded
        Dim NewxmlReader As New XmlDocument
        NewxmlReader.Load(xmlAddress)

        'Get File to be Download
        Dim urls = From file In NewxmlReader.GetElementsByTagName("url")
        For Each url As XmlElement In urls
            urlfile.Add(url.InnerText)
        Next

        For Each urladdress As String In urlfile
            Try
                Dim wbClient As New WebClient()
                filename = urladdress.Substring(urladdress.LastIndexOf("/") + 1)
                If Not (File.Exists(My.Application.Info.DirectoryPath + "\Downloaded\" + filename)) Then
                    wbClient.DownloadFile(urladdress, My.Application.Info.DirectoryPath + "\Downloaded\" + filename)
                    Dim path As String = My.Application.Info.DirectoryPath + "\Downloaded\"
                    RarArchive.WriteToDirectory(path + filename, path, NUnrar.Common.ExtractOptions.Overwrite)
                Else
                    Dim path As String = My.Application.Info.DirectoryPath + "\Downloaded\" + filename
                    File.Delete(path)
                    wbClient.DownloadFile(urladdress, path)
                    Dim pathextract As String = My.Application.Info.DirectoryPath + "\Downloaded\"
                    RarArchive.WriteToDirectory(pathextract + filename, path, NUnrar.Common.ExtractOptions.Overwrite)
                End If
            Catch ex As Exception
                MessageBox.Show("Error on background process: " & ex.Message.ToString,"Error",MessageBoxButtons.OK)
            End Try
        Next
    End Sub

    Private Sub BWDownloadFile_RunWorkerCompleted(sender As Object, e As ComponentModel.RunWorkerCompletedEventArgs) Handles BWDownloadFile.RunWorkerCompleted
        Dim Filename As String
        Dim downloadedfile = Directory.GetFiles(My.Application.Info.DirectoryPath + "\Downloaded\")

        For Each Listfile As String In downloadedfile
            Try
                Filename = Path.GetFileName(Listfile)
                If Not (File.Exists(My.Application.Info.DirectoryPath + "\" + Filename)) Then
                    File.Copy(My.Application.Info.DirectoryPath + "\Downloaded\" + Filename, _
                              My.Application.Info.DirectoryPath + "\" + Filename)
                Else
                    Dim path As String = My.Application.Info.DirectoryPath + "\" + Filename
                    File.Delete(path)
                    File.Copy(My.Application.Info.DirectoryPath + "\Downloaded\" + Filename, _
                             My.Application.Info.DirectoryPath + "\" + Filename)
                End If
            Catch ex As Exception
                MessageBox.Show("Error replacing file, with message: " & ex.Message.ToString, "Error", MessageBoxButtons.OK)
            End Try
        Next



        frmDownload.Close()

        ' Start Leveredge_Report_Tools.exe again
        Process.Start(My.Application.Info.DirectoryPath + "\Leveredge_Report_Tools.exe")
        Me.Close()
    End Sub

    Private Sub btnUpdateNow_Click(sender As Object, e As EventArgs) Handles btnUpdateNow.Click
        For Each proc As Process In Process.GetProcesses()
            If proc.ProcessName = "Leveredge_Report_Tools" Then
                proc.Kill()
            End If
        Next

        'Show Processing Update Form
        frmDownload.Show()

        ' Run Background Worker for Processing Update
        BWDownloadFile.RunWorkerAsync()
    End Sub

    Private Sub btnSkipVersion_Click(sender As Object, e As EventArgs) Handles btnSkipVersion.Click
        Me.Close()
    End Sub

End Class
