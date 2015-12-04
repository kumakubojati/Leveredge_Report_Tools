Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO.File
Imports System.Globalization
Imports System.Windows.Forms
Imports System.Xml
Public Class frmMain

    Private Sub ProductToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ProductToolStripMenuItem.Click
        frmOutMas_New.Show()
    End Sub

    Private Sub ProductMasterReportToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ProductMasterReportToolStripMenuItem.Click
        frmProdMas.Show()
    End Sub

    Private Sub WeeklySalesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles WeeklySalesToolStripMenuItem.Click
        frmLP3.Show()
    End Sub

    Private Sub SalesPerformanceReportToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SalesPerformanceReportToolStripMenuItem.Click
        frmSPR.Show()
    End Sub

    Private Sub DailySalesSummaryToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DailySalesSummaryToolStripMenuItem.Click
        frmDSS.Show()
    End Sub

    Private Sub DistributorStockReportToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DistributorStockReportToolStripMenuItem.Click
        frmDSR.Show()
    End Sub

    Private Sub DailyStockMutationToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DailyStockMutationToolStripMenuItem.Click
        frmDSM.Show()
    End Sub

    Private Sub ListOfPromotionToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ListOfPromotionToolStripMenuItem.Click
        frmListPro.Show()
    End Sub

    Private Sub ARReportToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ARReportToolStripMenuItem.Click
        frmAR.Show()
    End Sub

    Private Sub SummaryInvoiceAndSalesReturnReportToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SummaryInvoiceAndSalesReturnReportToolStripMenuItem.Click
        frmSISR.Show()
    End Sub

    Private Sub DailySalesAndPaymentSummaryReportToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DailySalesAndPaymentSummaryReportToolStripMenuItem.Click
        frmDSPS.Show()
    End Sub

    Private Sub SalesPerformanceIncentiveReportToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SalesPerformanceIncentiveReportToolStripMenuItem.Click
        frmSPIR.Show()
    End Sub

    Private Sub TurnOverReportToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TurnOverReportToolStripMenuItem.Click
        frmTOR.Show()
    End Sub

    Private Sub ListOfInvoiceToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ListOfInvoiceToolStripMenuItem.Click
        frmLIR.Show()
    End Sub

    Private Sub ServiceLevelReportToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ServiceLevelReportToolStripMenuItem.Click
        frmSL.Show()
    End Sub

    Private Sub ListOfPromotionUtilizationToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ListOfPromotionUtilizationToolStripMenuItem.Click
        frmLPU.Show()
    End Sub

    'Public Function IsInternetAvailable() As Boolean
    '    Dim objUrl As New System.Uri("http://www.Google.com")
    '    Dim objWebReq As WebRequest
    '    objWebReq = WebRequest.Create(objUrl)
    '    Dim objresp As WebResponse

    '    Try
    '        objresp = objWebReq.GetResponse
    '        objresp.Close()
    '        objresp = Nothing
    '        Return True

    '    Catch ex As Exception
    '        objresp = Nothing
    '        objWebReq = Nothing
    '        Return False
    '    End Try
    'End Function
    Public newversion As String
    Dim sqlconSA As New SqlConnection
    Dim constrSA As String
    Public cekDBResult As String
    Private Sub saDBCon()
        Dim strfile As String = My.Settings("MasterDBCon").ToString
        'constrSA = ReadAllLines(strfile)(5)
        sqlconSA.ConnectionString = strfile
    End Sub
    Public Sub CheckDB()
        Dim comstr As String
        comstr = "SELECT CASE WHEN EXISTS(SELECT 1 FROM sys.databases WHERE Name='FCS_REPORT') THEN 1 ELSE 0 END AS DbExisting"
        Dim cmdcek As New SqlCommand(comstr, sqlconSA)
        Try
            saDBCon()
            sqlconSA.Open()
            cekDBResult = cmdcek.ExecuteScalar()
            sqlconSA.Close()
        Catch ex As Exception
            MessageBox.Show("Error on Checking DB, with messag : " & ex.Message.ToString, "ERROR!!", MessageBoxButtons.OK, MessageBoxIcon.Error)
            sqlconSA.Close()
        End Try
    End Sub

    Private Sub RestoreNewDB()
        Dim comstr As String
        comstr = "USE MASTER RESTORE DATABASE [FCS_REPORT] FROM DISK = N'" & My.Application.Info.DirectoryPath & "\BlankDB.bak' " _
            & "WITH REPLACE"
        Dim cmdNewrestore As New SqlCommand(comstr, sqlconSA)
        Try
            saDBCon()
            sqlconSA.Open()
            cmdNewrestore.CommandTimeout = 600
            cmdNewrestore.ExecuteNonQuery()
            sqlconSA.Close()
        Catch ex As Exception
            MessageBox.Show("Error when restoring DB with message: " & ex.Message.ToString, "Error!!", MessageBoxButtons.OK, MessageBoxIcon.Error)
            sqlconSA.Close()
        End Try
    End Sub

    Public Sub FCSDB()
        Dim comstr As String = "SELECT (SELECT CASE WHEN (SELECT COUNT(*) FROM CASHMEMO) > 0 THEN 'TRUE' ELSE 'FALSE' END) AS DB_EXISITING"
        Dim strfile As String = My.Settings("FCSDBCon").ToString
        'Dim constr As String
        'constr = ReadAllLines(strfile)(3)
        Dim sqlcon As New SqlConnection
        sqlcon.ConnectionString = strfile
        Try
            Dim cmdcek As New SqlCommand(comstr, sqlcon)
            sqlcon.Open()
            Dim cekresult As String
            cekresult = cmdcek.ExecuteScalar
            If cekresult = "TRUE" Then
                FCSRepMenuItem.Enabled = True
                TargetMenuItem.Enabled = True
            Else
                Exit Try
            End If
        Catch ex As Exception
            MessageBox.Show("Error on Checking Data FCS_REPORT, with message: " & ex.Message.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            sqlcon.Close()
        End Try
        sqlcon.Close()
    End Sub
    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckVersion()
        CheckDB()
        If cekDBResult = "1" Then
            GoTo LineMain1
        ElseIf cekDBResult = "0" Then
            RestoreNewDB()
            Dim comstr_CreateUser, comstr_SetSA As String
            comstr_CreateUser = "IF NOT EXISTS (SELECT * FROM sys.server_principals WHERE name=N'FCSREP') " _
                & "CREATE LOGIN [FCSREP] WITH PASSWORD=N'unilever123', DEFAULT_LANGUAGE=[British],DEFAULT_DATABASE=[FCS_REPORT], " _
                & "CHECK_EXPIRATION=OFF, CHECK_POLICY=OFF"
            comstr_SetSA = "EXEC sp_addsrvrolemember @loginame = N'FCSREP', @rolename = N'sysadmin'"
            Dim cmd_CreateUser As New SqlCommand(comstr_CreateUser, sqlconSA)
            Dim cmd_SetSA As New SqlCommand(comstr_SetSA, sqlconSA)
            Try
                saDBCon()
                sqlconSA.Open()
                cmd_CreateUser.ExecuteNonQuery()
                cmd_SetSA.ExecuteNonQuery()
                sqlconSA.Close()
            Catch ex As Exception
                MessageBox.Show("Error on creating DB, with message: " & ex.Message.ToString, "Error!!", MessageBoxButtons.OK, MessageBoxIcon.Error)
                sqlconSA.Close()
            End Try
        End If
LineMain1: FCSDB()
        'If IsInternetAvailable() = True Then
        '    Try
        '        Dim request As HttpWebRequest = HttpWebRequest.Create("Http://182.253.21.82:8020/LastVersion.txt")
        '        Dim response As HttpWebResponse = request.GetResponse()
        '        Dim sr As StreamReader = New StreamReader(response.GetResponseStream())
        '        newversion = sr.ReadToEnd()
        '    Catch ex As Exception
        '        MessageBox.Show("No Connection Update Server, with  message: " & ex.Message.ToString, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
        '        Exit Sub
        '    End Try
        '    Dim currentverison As String = Application.ProductVersion
        '    Dim splitcurver = currentverison.Split(".")
        '    Dim splitnewver = newversion.Split(".")
        '    Dim joincurver As String = splitcurver(0) & splitcurver(1) & splitcurver(2) & splitcurver(3)
        '    Dim joinnewver As String = splitnewver(0) & splitnewver(1) & splitnewver(2) & splitnewver(3)
        '    Dim rescurver As Integer = Convert.ToInt32(joincurver)
        '    Dim resnewver As Integer = Convert.ToInt32(joinnewver)
        '    'If newversion.Contains(currentverison) Then
        '    '    Exit Sub
        '    'Else
        '    '    Dim update As Integer = MessageBox.Show("New version (ver " & newversion & ") are available, proceed to download? ", "New Version", MessageBoxButtons.YesNo)
        '    '    If update = DialogResult.Yes Then
        '    '        SSdownload.Show()
        '    '    ElseIf update = DialogResult.No Then
        '    '        Exit Sub
        '    '    End If
        '    'End If

        '    If resnewver <= rescurver Then
        '        Exit Sub
        '    ElseIf resnewver > rescurver Then
        '        Dim update As Integer = MessageBox.Show("New version (ver " & newversion & ") are available, proceed to download? ", "New Version", MessageBoxButtons.YesNo)
        '        If update = DialogResult.Yes Then
        '            SSdownload.Show()
        '        ElseIf update = DialogResult.No Then
        '            Exit Sub
        '        End If
        '    End If
        'Else
        '    MessageBox.Show("No Internet Connection")
        'End If

        
    End Sub

    Private Sub AnalysisBackupToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AnalysisBackupToolStripMenuItem.Click
        frmABR_IC.Show()
    End Sub

    Private Sub BEPThroughputToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles BEPThroughputToolStripMenuItem.Click
        frmBEPT.Show()
    End Sub

    Private Sub CabinetNetIncreaseReportToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CabinetNetIncreaseReportToolStripMenuItem.Click
        frmCNI_IC.Show()
    End Sub

    Private Sub CustomerVisitReportToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CustomerVisitReportToolStripMenuItem.Click
        frmCV_IC.Show()
    End Sub

    Private Sub DistributorDriveReportToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DistributorDriveReportToolStripMenuItem.Click
        frmDD_IC.Show()
    End Sub

    Private Sub IQPerformanceReportToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles IQPerformanceReportToolStripMenuItem.Click
        frmIQP_IC.Show()
    End Sub

    Private Sub IQSummaryReportToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles IQSummaryReportToolStripMenuItem.Click
        frmIQS_IC.Show()
    End Sub

    Private Sub NonPerformanceToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NonPerformanceToolStripMenuItem.Click
        frmNPC_IC.Show()
    End Sub

    Private Sub NetIncreaseReportToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NetIncreaseReportToolStripMenuItem.Click
        frmONI_IC.Show()
    End Sub

    Private Sub OutletStoreClassReportToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OutletStoreClassReportToolStripMenuItem.Click
        frmOSC_IC.Show()
    End Sub

    Private Sub ThroughPutCabinetReportToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ThroughPutCabinetReportToolStripMenuItem.Click
        frmTCR_IC.Show()
    End Sub

    Private Sub TroughutOutletReportToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TroughutOutletReportToolStripMenuItem.Click
        frmTPO_IC.Show()
    End Sub

    Private Sub WaiveStoreReportToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles WaiveStoreReportToolStripMenuItem.Click
        frmWSR_IC.Show()
    End Sub

    Private Sub ProductSalesByCaseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ProductSalesByCaseToolStripMenuItem.Click
        frmPSCW_IC.Show()
    End Sub

    Private Sub SalesReturnToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SalesReturnToolStripMenuItem.Click
        frmSRDR.Show()
    End Sub

    Private Sub ProductSalesByVolumeWeeklyReportToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ProductSalesByVolumeWeeklyReportToolStripMenuItem.Click
        frmPSVW_IC.Show()
    End Sub

    Private Sub AchievementByProductToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles AchievementByProductToolStripMenuItem1.Click
        frmABP.Show()
    End Sub

    Private Sub TargetToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TargetMenuItem.Click
        frmTarget.Show()
    End Sub

    Private Sub InitDBMenuItem_Click(sender As Object, e As EventArgs) Handles InitDBMenuItem.Click
        frmInitDB.Show()
    End Sub

    Private Sub FCSRepMenuItem_Click(sender As Object, e As EventArgs) Handles FCSRepMenuItem.Click
        SSUploadDB.Show()
    End Sub

    Private Sub EvoucherToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EvoucherToolStripMenuItem.Click
        frmEvoucherRep.Show()
    End Sub

    Private Sub DBConnectionToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DBConnectionToolStripMenuItem.Click
        frmDBConPass.Show()
    End Sub
    Public NewVers, CurrVers As String
    Private Sub CheckVersion()
        ' Delete all file in Download folder
        For Each f As String In Directory.GetFiles("D:\Unilever Indonesia\Leveredge Report Tools\Downloaded")
            File.Delete(f)
        Next f
        'Read xml current and new version
        Const xmlAddress As String = "http://www.uid-leveredge.cf/LRT/AppCast.xml"
        Dim xmlReader As New XmlDocument
        Dim current_xmlAddress As String = My.Application.Info.DirectoryPath & "\AppCast.xml"
        Dim xmlReader_version As New XmlDocument
        xmlReader_version.Load(current_xmlAddress)
        xmlReader.Load(xmlAddress)
        Dim Curr_Vers = xmlReader_version.GetElementsByTagName("version")
        For Each Curr_Ver As XmlElement In Curr_Vers
            CurrVers = Curr_Ver.InnerText
            CurrVers = CurrVers.Replace(".", "")
        Next
        Dim New_Vers = xmlReader.GetElementsByTagName("version")
        For Each NewVer As XmlElement In New_Vers
            NewVers = NewVer.InnerText
            NewVers = NewVers.Replace(".", "")
        Next
        If NewVers > CurrVers Then
            Process.Start(My.Application.Info.DirectoryPath + "\Updater.exe")
        Else
            Exit Sub
        End If
    End Sub

    Private Sub SummaryReportForDSRToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SummaryReportForDSRToolStripMenuItem.Click
        frmISRD_IQ.Show()
    End Sub
End Class
