Imports System.Net
Imports System.IO
Public Class frmMain

    Private Sub ProductToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ProductToolStripMenuItem.Click
        frmOutMas.Show()
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

    Private Sub AchievementByOutletToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AchievementByOutletToolStripMenuItem.Click
        frmABP.Show()
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

    Public Function IsInternetAvailable() As Boolean
        Dim objUrl As New System.Uri("http://www.Google.com")
        Dim objWebReq As WebRequest
        objWebReq = WebRequest.Create(objUrl)
        Dim objresp As WebResponse

        Try
            objresp = objWebReq.GetResponse
            objresp.Close()
            objresp = Nothing
            Return True

        Catch ex As Exception
            objresp = Nothing
            objWebReq = Nothing
            Return False
        End Try
    End Function
    Public newversion As String
    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If IsInternetAvailable() = True Then
            Dim request As HttpWebRequest = HttpWebRequest.Create("Http://182.253.21.82:8020/LastVersion.txt")
            Dim response As HttpWebResponse = request.GetResponse()
            Dim sr As StreamReader = New StreamReader(response.GetResponseStream())
            newversion = sr.ReadToEnd()
            Dim currentverison As String = Application.ProductVersion
            If newversion.Contains(currentverison) Then
                Exit Sub
            Else
                Dim update As Integer = MessageBox.Show("New version (ver " & newversion & ") are available, proceed to download? ", "New Version", MessageBoxButtons.YesNo)
                If update = DialogResult.Yes Then
                    SSdownload.Show()
                ElseIf update = DialogResult.No Then
                    Exit Sub
                End If
            End If
        Else
            MessageBox.Show("No Internet Connection")
        End If
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
End Class
