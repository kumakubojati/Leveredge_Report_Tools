Public Class frmFCS_RepViewer

    Private Sub frmFCS_RepViewer_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        With Me.rvFCS_New.LocalReport
            .ReportPath = My.Application.Info.DirectoryPath & "\FCSRep_New.rdlc"
        End With
        frmFCS_Rep.GetSelectedCheckBox()
        frmFCS_Rep.GetDataHead()
        'frmFCS_Rep.GetECO()
        'frmFCS_Rep.GetBPSchPRod()
        'frmFCS_Rep.GetLPPC()
        frmFCS_Rep.GetFCSData()
        SSLoading.Close()
        Me.rvFCS_New.RefreshReport()
    End Sub
End Class