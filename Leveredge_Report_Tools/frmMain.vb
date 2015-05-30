﻿
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
End Class
