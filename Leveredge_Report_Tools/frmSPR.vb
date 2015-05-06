Imports Microsoft.Office.Interop
Imports System.Windows.Forms
Imports System.Text
Imports System.Drawing
Imports Microsoft.Win32

Public Class frmSPR
    Dim AppsOffice As String
    Private Sub frmSPR_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim TargetKey As RegistryKey
        TargetKey = Registry.ClassesRoot.OpenSubKey("Excel.Application")
        If TargetKey Is Nothing Then
            AppsOffice = "XL_NotInstalled"
        Else
            AppsOffice = "XL_Installed"
            TargetKey.Close()
        End If
    End Sub

    Private Sub btnBrow_SPR_src_Click(sender As Object, e As EventArgs) Handles btnBrow_SPR_src.Click
        Dim SPRpathSrc As String
        If OFD_SPR.ShowDialog = DialogResult.OK Then
            SPRpathSrc = OFD_SPR.FileName()
            txtSPR_src.Text = SPRpathSrc
        End If
        If txtSPR_dest.Text <> "" Then
            btnNeu_SPR.Enabled = True
        Else
            btnNeu_SPR.Enabled = False
        End If
    End Sub

    Private Sub btnBrow_SPR_dest_Click(sender As Object, e As EventArgs) Handles btnBrow_SPR_dest.Click
        Dim datenow As DateTime = DateTime.Now
        Dim filename As String = "Neutralize_SalesPerf" & SPRreptype & "_" & datenow.ToString("ddMMyyyy_HHmm")
        SFD_SPR.FileName = filename
        Dim SPRpath_Dest As String
        If SFD_SPR.ShowDialog = DialogResult.OK Then
            SPRpath_Dest = SFD_SPR.FileName
            txtSPR_dest.Text = SPRpath_Dest
        End If
        If txtSPR_src.Text <> "" Then
            btnNeu_SPR.Enabled = True
        Else
            btnNeu_SPR.Enabled = False
        End If
        If txtSPR_dest.Text <> "" Then
            btnNeu_SPR.Enabled = True
        Else
            btnNeu_SPR.Enabled = False
        End If
    End Sub

    Private Sub btnNeu_SPR_Click(sender As Object, e As EventArgs) Handles btnNeu_SPR.Click
        PicBar_SPR.Visible = True
        BWSPR.RunWorkerAsync()
    End Sub

    Private Sub BWSPR_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BWSPR.RunWorkerCompleted
        PicBar_SPR.Visible = False
        btnNeu_SPR.Enabled = False
        txtSPR_dest.Text = ""
        txtSPR_src.Text = ""
    End Sub
    Dim SPRreptype As String

    Private Sub RBSPR_Qty_CheckedChanged(sender As Object, e As EventArgs) Handles RBSPR_Qty.CheckedChanged
        SPRreptype = "QTY"
    End Sub

    Private Sub RBSPR_Val_CheckedChanged(sender As Object, e As EventArgs) Handles RBSPR_Val.CheckedChanged
        SPRreptype = "VAL"
    End Sub

    Private Sub RBSPR_Avg_CheckedChanged(sender As Object, e As EventArgs) Handles RBSPR_Avg.CheckedChanged
        SPRreptype = "AVG"
    End Sub

    Private Sub BWSPR_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BWSPR.DoWork
        Select Case AppsOffice
            Case "XL_NotInstalled"
                Dim xlAppSPR As Object
                Dim xlWbookSPR As Object
                Dim xlWsheetSPR As Object

                Select Case SPRreptype
                    Case "QTY"
                        'Need Confirmation to Pak Arif for Dynamic Column

                    Case "VAL"
                        Try
                            xlAppSPR = CreateObject("Ket.Application")
                            xlWbookSPR = xlAppSPR.Workbooks.Open(txtSPR_src.Text)
                            xlWsheetSPR = xlWbookSPR.Worksheets("UID Sales Perfomance Report")

                            xlWsheetSPR.UsedRange.UnMerge()
                            xlWsheetSPR.UsedRange.WrapText = False
                            xlWsheetSPR.UsedRange.ColumnWidth = 15
                            xlWsheetSPR.UsedRange.RowHeight = 15

                            Dim rg_head_cut1 As Object = xlWsheetSPR.Range("B2")
                            Dim rg_head_paste1 As Object = xlWsheetSPR.Range("A2")
                            rg_head_cut1.Select()
                            rg_head_cut1.Cut(rg_head_paste1)

                            Dim rg_head_cut2 As Object = xlWsheetSPR.Range("B4")
                            Dim rg_head_paste2 As Object = xlWsheetSPR.Range("A3")
                            rg_head_cut2.Select()
                            rg_head_cut2.Cut(rg_head_paste2)

                            xlWsheetSPR.Range("A2").RowHeight = 27

                            Dim paramhead1, paramhead2 As String
                            paramhead1 = xlWsheetSPR.Range("C6").Value & " " & xlWsheetSPR.Range("F6").Value & "; "
                            paramhead1 = paramhead1 & xlWsheetSPR.Range("H6").Value & " " & xlWsheetSPR.Range("K6").Value & "; "

                            paramhead2 = xlWsheetSPR.Range("P6").Value & " " & xlWsheetSPR.Range("S6").Value & "; "
                            paramhead2 = paramhead2 & xlWsheetSPR.Range("W6").Value & " " & xlWsheetSPR.Range("Z6").Value & "; "

                            xlWsheetSPR.Range("A6").Value = paramhead1
                            xlWsheetSPR.Range("A6").EntireRow.Font.Name = "Calibri"
                            xlWsheetSPR.Range("A7").Value = paramhead2
                            xlWsheetSPR.Range("A7").EntireRow.Font.Name = "Calibri"

                            Dim rg1, rg2, rg3, rg4, rg5, rg6 As Object
                            rg1 = xlWsheetSPR.Range("B:C")
                            rg1.Select()
                            rg1.Delete()

                            rg2 = xlWsheetSPR.Range("C:F")
                            rg2.Select()
                            rg2.Delete()

                            rg3 = xlWsheetSPR.Range("D:E")
                            rg3.Select()
                            rg3.Delete()

                            xlWsheetSPR.Range("E:E").EntireColumn.Delete()

                            rg4 = xlWsheetSPR.Range("F:G")
                            rg4.Select()
                            rg4.Delete()

                            rg5 = xlWsheetSPR.Range("G:H")
                            rg5.Select()
                            rg5.Delete()

                            rg6 = xlWsheetSPR.Range("H:J")
                            rg6.Select()
                            rg6.Delete()

                            xlWsheetSPR.Range("J:J").EntireColumn.Delete()
                            xlWsheetSPR.Range("A4").EntireRow.Delete()

                            xlWbookSPR.SaveAs(txtSPR_dest.Text)
                            xlWbookSPR.Close()
                            xlAppSPR.Quit()

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetSPR)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookSPR)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppSPR)

                            xlWsheetSPR = Nothing
                            xlWbookSPR = Nothing
                            xlAppSPR = Nothing

                            MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                        Catch ex As Exception
                            MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try

                    Case "AVG"
                        'Need Confirmation to Pak Arif for Dynamic Column
                End Select

            Case "XL_Installed"
                Dim xlAppSPR As Object
                Dim xlWbookSPR As Object
                Dim xlWsheetSPR As Object

                Select Case SPRreptype
                    Case "QTY"
                        'Need Confirmation to Pak Arif for Dynamic Column

                    Case "VAL"
                        Try
                            xlAppSPR = CreateObject("Excel.Application")
                            xlWbookSPR = xlAppSPR.Workbooks.Open(txtSPR_src.Text)
                            xlWsheetSPR = xlWbookSPR.Worksheets("UID Sales Perfomance Report")

                            xlWsheetSPR.UsedRange.UnMerge()
                            xlWsheetSPR.UsedRange.WrapText = False
                            xlWsheetSPR.UsedRange.ColumnWidth = 15
                            xlWsheetSPR.UsedRange.RowHeight = 15

                            Dim rg_head_cut1 As Excel.Range = xlWsheetSPR.Range("B2")
                            Dim rg_head_paste1 As Excel.Range = xlWsheetSPR.Range("A2")
                            rg_head_cut1.Select()
                            rg_head_cut1.Cut(rg_head_paste1)

                            Dim rg_head_cut2 As Excel.Range = xlWsheetSPR.Range("B4")
                            Dim rg_head_paste2 As Excel.Range = xlWsheetSPR.Range("A3")
                            rg_head_cut2.Select()
                            rg_head_cut2.Cut(rg_head_paste2)

                            xlWsheetSPR.Range("A2").RowHeight = 27

                            Dim paramhead1, paramhead2 As String
                            paramhead1 = xlWsheetSPR.Range("C6").Value & " " & xlWsheetSPR.Range("F6").Value & "; "
                            paramhead1 = paramhead1 & xlWsheetSPR.Range("H6").Value & " " & xlWsheetSPR.Range("K6").Value & "; "

                            paramhead2 = xlWsheetSPR.Range("P6").Value & " " & xlWsheetSPR.Range("S6").Value & "; "
                            paramhead2 = paramhead2 & xlWsheetSPR.Range("W6").Value & " " & xlWsheetSPR.Range("Z6").Value & "; "

                            xlWsheetSPR.Range("A6").Value = paramhead1
                            xlWsheetSPR.Range("A6").EntireRow.Font.Name = "Calibri"
                            xlWsheetSPR.Range("A7").Value = paramhead2
                            xlWsheetSPR.Range("A7").EntireRow.Font.Name = "Calibri"

                            Dim rg1, rg2, rg3, rg4, rg5, rg6 As Excel.Range
                            rg1 = xlWsheetSPR.Range("B:C")
                            rg1.Select()
                            rg1.Delete()

                            rg2 = xlWsheetSPR.Range("C:F")
                            rg2.Select()
                            rg2.Delete()

                            rg3 = xlWsheetSPR.Range("D:E")
                            rg3.Select()
                            rg3.Delete()

                            xlWsheetSPR.Range("E:E").EntireColumn.Delete()

                            rg4 = xlWsheetSPR.Range("F:G")
                            rg4.Select()
                            rg4.Delete()

                            rg5 = xlWsheetSPR.Range("G:H")
                            rg5.Select()
                            rg5.Delete()

                            rg6 = xlWsheetSPR.Range("H:J")
                            rg6.Select()
                            rg6.Delete()

                            xlWsheetSPR.Range("J:J").EntireColumn.Delete()
                            xlWsheetSPR.Range("A4").EntireRow.Delete()

                            xlWbookSPR.SaveAs(txtSPR_dest.Text)
                            xlWbookSPR.Close()
                            xlAppSPR.Quit()

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetSPR)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookSPR)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppSPR)

                            xlWsheetSPR = Nothing
                            xlWbookSPR = Nothing
                            xlAppSPR = Nothing

                            MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                        Catch ex As Exception
                            MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try

                    Case "AVG"
                        'Need Confirmation to Pak Arif for Dynamic Column
                End Select

        End Select
        
    End Sub
End Class