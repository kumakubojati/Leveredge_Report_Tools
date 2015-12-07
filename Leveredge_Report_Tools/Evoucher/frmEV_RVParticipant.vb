Public Class frmEV_RVParticipant

    Private Sub frmEV_RVParticipant_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        With Me.RV_EVParticipant.LocalReport
            .ReportPath = My.Application.Info.DirectoryPath & "\EV_Participant.rdlc"
        End With
        frmEvoucherRep.Get_EV_HEAD()
        frmEvoucherRep.Get_EV_Participant()
        Me.RV_EVParticipant.RefreshReport()
    End Sub
End Class