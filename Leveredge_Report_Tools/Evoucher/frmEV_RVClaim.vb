Public Class frmEV_RVClaim

    Private Sub frmEV_RepView_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        With Me.RV_EV_CLAIM.LocalReport
            .ReportPath = My.Application.Info.DirectoryPath & "\EV_Claim.rdlc"
        End With
        frmEvoucherRep.Get_EV_HEAD()
        frmEvoucherRep.Get_EV_CLAIM()
        Me.RV_EV_CLAIM.RefreshReport()
    End Sub
End Class