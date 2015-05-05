Imports Microsoft.Office.Interop
Imports System.Windows.Forms
Imports System.Text
Imports System.Drawing
Imports Microsoft.Win32
Public Class frmDSS
    Dim AppsOffice As String
    Private Sub frmDSS_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim TargetKey As RegistryKey
        TargetKey = Registry.ClassesRoot.OpenSubKey("Excel.Application")
        If TargetKey Is Nothing Then
            AppsOffice = "XL_NotInstalled"
        Else
            AppsOffice = "XL_Installed"
            TargetKey.Close()
        End If
    End Sub
    Dim DSStype As String

    Private Sub rbDSSall_CheckedChanged(sender As Object, e As EventArgs) Handles rbDSSall.CheckedChanged
        DSStype = "ALL"
    End Sub

    Private Sub rbDSSrupiah_CheckedChanged(sender As Object, e As EventArgs) Handles rbDSSrupiah.CheckedChanged
        DSStype = "RUPIAH"
    End Sub

    Private Sub rbDSSprod_CheckedChanged(sender As Object, e As EventArgs) Handles rbDSSprod.CheckedChanged
        DSStype = "PRODUK"
    End Sub

    Private Sub bntBrowDSS_src_Click(sender As Object, e As EventArgs) Handles bntBrowDSS_src.Click
        Dim DSSpathSrc As String
        If OFD_DSS.ShowDialog = DialogResult.OK Then
            DSSpathSrc = OFD_DSS.FileName
            txtDSS_src.Text = DSSpathSrc
        End If
        If txtDSS_dest.Text <> "" Then
            btnNeu_DSS.Enabled = True
        Else
            btnNeu_DSS.Enabled = False
        End If
    End Sub

    Private Sub btnBrowDSS_dest_Click(sender As Object, e As EventArgs) Handles btnBrowDSS_dest.Click
        Dim datenow As DateTime = DateTime.Now
        Dim filename As String = "Neutralize_DailySalesSumm_" & datenow.ToString("ddMMyyyy_HHmm")
        SFD_DSS.FileName = filename
        Dim DSSpath_Dest As String
        If SFD_DSS.ShowDialog = DialogResult.OK Then
            DSSpath_Dest = SFD_DSS.FileName
            txtDSS_dest.Text = DSSpath_Dest
        End If
        If txtDSS_src.Text <> "" Then
            btnNeu_DSS.Enabled = True
        Else
            btnNeu_DSS.Enabled = False
        End If
        If txtDSS_dest.Text <> "" Then
            btnNeu_DSS.Enabled = True
        Else
            btnNeu_DSS.Enabled = False
        End If
    End Sub

    Private Sub btnNeu_DSS_Click(sender As Object, e As EventArgs) Handles btnNeu_DSS.Click
        PicBar_DSS.Visible = True
        BWDSS.RunWorkerAsync()
    End Sub

    Private Sub BWDSS_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BWDSS.RunWorkerCompleted
        PicBar_DSS.Visible = False
        btnNeu_DSS.Enabled = False
        txtDSS_dest.Text = ""
        txtDSS_src.Text = ""
    End Sub

    Private Sub BWDSS_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BWDSS.DoWork
        Dim xlAppDSS As Excel.Application
        Dim xlWbookDSS As Excel.Workbook
        Dim xlWsheetDSS As Excel.Worksheet

        Select Case DSStype
            Case "ALL"
                Try
                    xlAppDSS = New Excel.Application
                    xlWbookDSS = xlAppDSS.Workbooks.Open(txtDSS_src.Text)
                    xlWsheetDSS = xlWbookDSS.Worksheets("UID Daily Sales Summary Report")

                    xlWsheetDSS.UsedRange.UnMerge()
                    xlWsheetDSS.UsedRange.WrapText = False
                    xlWsheetDSS.UsedRange.ColumnWidth = 15
                    xlWsheetDSS.UsedRange.RowHeight = 15

                    Dim rg_head_cut1 As Excel.Range = xlWsheetDSS.Range("B2")
                    Dim rg_head_paste1 As Excel.Range = xlWsheetDSS.Range("A2")
                    rg_head_cut1.Select()
                    rg_head_cut1.Cut(rg_head_paste1)

                    Dim rg_head_cut2 As Excel.Range = xlWsheetDSS.Range("B4")
                    Dim rg_head_paste2 As Excel.Range = xlWsheetDSS.Range("A3")
                    rg_head_cut2.Select()
                    rg_head_cut2.Cut(rg_head_paste2)

                    xlWsheetDSS.Range("A2").RowHeight = 27

                    Dim paramhead1, paramhead2, paramhead3 As String
                    paramhead1 = xlWsheetDSS.Range("C6").Value & " " & xlWsheetDSS.Range("F6").Value & "; "
                    paramhead1 = paramhead1 & xlWsheetDSS.Range("I6").Value & " " & xlWsheetDSS.Range("M6").Value & "; "

                    paramhead2 = xlWsheetDSS.Range("R6").Value & " " & xlWsheetDSS.Range("V6").Value & "; "
                    paramhead2 = paramhead2 & xlWsheetDSS.Range("AE6").Value & " " & xlWsheetDSS.Range("AI6").Value & "; "

                    paramhead3 = xlWsheetDSS.Range("AM6").Value & " " & xlWsheetDSS.Range("AP6").Value

                    xlWsheetDSS.Range("A5").Value = paramhead1
                    xlWsheetDSS.Range("A5").EntireRow.Font.Name = "Calibri"
                    xlWsheetDSS.Range("A6").Value = paramhead2
                    xlWsheetDSS.Range("A6").EntireRow.Font.Name = "Calibri"
                    xlWsheetDSS.Range("A7").Value = paramhead3
                    xlWsheetDSS.Range("A7").EntireColumn.Font.Name = "Calibri"

                    Dim rg1, rg2, rg3, rg4, rg5, rg6, rg7, rg8, rg9 As Excel.Range
                    rg1 = xlWsheetDSS.Range("B:C")
                    rg1.Select()
                    rg1.Delete()
                    rg2 = xlWsheetDSS.Range("C:D")
                    rg2.Select()
                    rg2.Delete()
                    rg3 = xlWsheetDSS.Range("D:E")
                    rg3.Select()
                    rg3.Delete()
                    xlWsheetDSS.Range("E:E").EntireColumn.Delete()
                    xlWsheetDSS.Range("F:F").EntireColumn.Delete()
                    xlWsheetDSS.Range("G:G").EntireColumn.Delete()
                    xlWsheetDSS.Range("I:I").EntireColumn.Delete()
                    rg4 = xlWsheetDSS.Range("K:L")
                    rg4.Select()
                    rg4.Delete()
                    xlWsheetDSS.Range("M:M").EntireColumn.Delete()
                    xlWsheetDSS.Range("N:N").EntireColumn.Delete()
                    xlWsheetDSS.Range("O:O").EntireColumn.Delete()
                    rg5 = xlWsheetDSS.Range("P:Q")
                    rg5.Select()
                    rg5.Delete()
                    rg6 = xlWsheetDSS.Range("Q:R")
                    rg6.Select()
                    rg6.Delete()
                    rg7 = xlWsheetDSS.Range("S:T")
                    rg7.Select()
                    rg7.Delete()
                    rg8 = xlWsheetDSS.Range("T:U")
                    rg8.Select()
                    rg8.Delete()
                    xlWsheetDSS.Range("W:W").EntireColumn.Delete()

                    Dim lrow8 As Long = xlWsheetDSS.Range("A8").End(Excel.XlDirection.xlDown).Row
                    Dim r8a As Integer = lrow8 + 1
                    Dim r8b As Integer = lrow8 + 2
                    Dim cr8 As String = "A" & r8a & ":A" & r8b
                    xlWsheetDSS.Range(cr8).EntireRow.Delete()

                    Dim lrow1 As Long
                    lrow1 = xlWsheetDSS.Range("B65536").End(Excel.XlDirection.xlUp).Row
                    Dim r1 As Integer = lrow1 + 6
                    Dim cr1 As String = "B" & r1 & ":B65536"
                    Dim cellrange1 As Excel.Range = xlWsheetDSS.Range(cr1)
                    cellrange1.Select()
                    cellrange1.Delete(Shift:=Excel.XlDirection.xlToLeft)

                    xlWsheetDSS.Range("C:C").EntireColumn.Delete()
                    xlWsheetDSS.Range("D:D").EntireColumn.Delete()

                    Dim lrow2 As Long
                    lrow2 = xlWsheetDSS.Range("D65536").End(Excel.XlDirection.xlUp).Row
                    Dim r2 As Integer = lrow2 + 6
                    Dim cr2 As String = "D" & r2 & ":D65536"
                    Dim cellrange2 As Excel.Range = xlWsheetDSS.Range(cr2)
                    cellrange2.Select()
                    cellrange2.Delete(Shift:=Excel.XlDirection.xlToLeft)

                    xlWsheetDSS.Range("F:F").EntireColumn.Delete()

                    Dim lrow3 As Long
                    lrow3 = xlWsheetDSS.Range("E65536").End(Excel.XlDirection.xlUp).Row
                    Dim r3 As Integer = lrow3 + 6
                    Dim cr3 As String = "E" & r3 & ":E65536"
                    Dim cellrange3 As Excel.Range = xlWsheetDSS.Range(cr3)
                    cellrange3.Select()
                    cellrange3.Delete(Shift:=Excel.XlDirection.xlToLeft)

                    Dim lrow4 As Long = xlWsheetDSS.Range("E65536").End(Excel.XlDirection.xlUp).Row
                    Dim r4 As Integer = lrow4 + 6
                    Dim cr4 As String = "E" & r4 & ":E65536"
                    Dim cellrange4 As Excel.Range = xlWsheetDSS.Range(cr4)
                    cellrange4.Select()
                    cellrange4.Delete(Shift:=Excel.XlDirection.xlToLeft)

                    Dim lrow9 As Long = xlWsheetDSS.Range("F1").End(Excel.XlDirection.xlDown).Row
                    Dim r9 As Integer = lrow9 + 3
                    Dim cr9 As String = "F" & lrow9 & ":F" & r9

                    Dim lrow10 As Long = xlWsheetDSS.Range("A8").End(Excel.XlDirection.xlDown).Row
                    Dim r10 As Integer = lrow10 - 3
                    Dim cr10 As String = "B" & lrow10 & ":B" & r10

                    xlWsheetDSS.Range(cr9).Copy(xlWsheetDSS.Range(cr10))
                    xlWsheetDSS.Range("F:F").EntireColumn.Delete()

                    Dim lrow5 As Long = xlWsheetDSS.Range("F65536").End(Excel.XlDirection.xlUp).Row
                    Dim r5 As Integer = lrow5 + 6
                    Dim cr5 As String = "E" & r5 & ":F65536"
                    Dim cellrange5 As Excel.Range = xlWsheetDSS.Range(cr5)
                    cellrange5.Select()
                    cellrange5.Delete(Shift:=Excel.XlDirection.xlToLeft)

                    Dim lrow6 As Long = xlWsheetDSS.Range("H1").End(Excel.XlDirection.xlDown).Row
                    Dim r6 As Integer = lrow6 - 1
                    Dim cr6 As String = "H1:H" & r6
                    Dim cellrange6 As Excel.Range = xlWsheetDSS.Range(cr6)
                    cellrange6.Select()
                    cellrange6.Delete(Shift:=Excel.XlDirection.xlToLeft)

                    Dim lrow7 As Long = xlWsheetDSS.Range("J1").End(Excel.XlDirection.xlDown).Row
                    Dim r7 As Integer = lrow7 - 1
                    Dim cr7 As String = "H1:H" & r7
                    Dim cellrange7 As Excel.Range = xlWsheetDSS.Range(cr7)
                    cellrange7.Select()
                    cellrange7.Delete(Shift:=Excel.XlDirection.xlToLeft)

                    rg9 = xlWsheetDSS.Range("Q:T")
                    rg9.Select()
                    rg9.Delete()

                    xlWbookDSS.SaveAs(txtDSS_dest.Text)
                    xlWbookDSS.Close()
                    xlAppDSS.Quit()

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetDSS)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookDSS)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppDSS)

                    xlWsheetDSS = Nothing
                    xlWbookDSS = Nothing
                    xlAppDSS = Nothing

                    MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Catch ex As Exception
                    MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try

            Case "RUPIAH"
                Try
                    xlAppDSS = New Excel.Application
                    xlWbookDSS = xlAppDSS.Workbooks.Open(txtDSS_src.Text)
                    xlWsheetDSS = xlWbookDSS.Worksheets("UID Daily Sales Summary Report")

                    xlWsheetDSS.UsedRange.UnMerge()
                    xlWsheetDSS.UsedRange.WrapText = False
                    xlWsheetDSS.UsedRange.ColumnWidth = 15
                    xlWsheetDSS.UsedRange.RowHeight = 15

                    Dim rg_head_cut1 As Excel.Range = xlWsheetDSS.Range("B2")
                    Dim rg_head_paste1 As Excel.Range = xlWsheetDSS.Range("A2")
                    rg_head_cut1.Select()
                    rg_head_cut1.Cut(rg_head_paste1)

                    Dim rg_head_cut2 As Excel.Range = xlWsheetDSS.Range("B4")
                    Dim rg_head_paste2 As Excel.Range = xlWsheetDSS.Range("A3")
                    rg_head_cut2.Select()
                    rg_head_cut2.Cut(rg_head_paste2)

                    xlWsheetDSS.Range("A2").RowHeight = 27

                    Dim paramhead1, paramhead2, paramhead3 As String
                    paramhead1 = xlWsheetDSS.Range("C6").Value & " " & xlWsheetDSS.Range("E6").Value & "; "
                    paramhead1 = paramhead1 & xlWsheetDSS.Range("H6").Value & " " & xlWsheetDSS.Range("K6").Value & "; "

                    paramhead2 = xlWsheetDSS.Range("N6").Value & " " & xlWsheetDSS.Range("P6").Value & "; "
                    paramhead2 = paramhead2 & xlWsheetDSS.Range("U6").Value & " " & xlWsheetDSS.Range("X6").Value & "; "

                    paramhead3 = xlWsheetDSS.Range("AB6").Value & " " & xlWsheetDSS.Range("AE6").Value

                    xlWsheetDSS.Range("A5").Value = paramhead1
                    xlWsheetDSS.Range("A5").EntireRow.Font.Name = "Calibri"
                    xlWsheetDSS.Range("A6").Value = paramhead2
                    xlWsheetDSS.Range("A6").EntireRow.Font.Name = "Calibri"
                    xlWsheetDSS.Range("A7").Value = paramhead3
                    xlWsheetDSS.Range("A7").EntireColumn.Font.Name = "Calibri"

                    Dim rg1, rg2, rg3, rg4, rg5, rg6, rg7, rg8 As Excel.Range
                    rg1 = xlWsheetDSS.Range("B:E")
                    rg1.Select()
                    rg1.Delete()
                    rg2 = xlWsheetDSS.Range("C:E")
                    rg2.Select()
                    rg2.Delete()
                    rg3 = xlWsheetDSS.Range("D:E")
                    rg3.Select()
                    rg3.Delete()
                    rg4 = xlWsheetDSS.Range("E:G")
                    rg4.Select()
                    rg4.Delete()
                    rg5 = xlWsheetDSS.Range("F:G")
                    rg5.Select()
                    rg5.Delete()
                    xlWsheetDSS.Range("G:G").EntireColumn.Delete()
                    rg6 = xlWsheetDSS.Range("H:I")
                    rg6.Select()
                    rg6.Delete()
                    rg7 = xlWsheetDSS.Range("J:K")
                    rg7.Select()
                    rg7.Delete()
                    rg8 = xlWsheetDSS.Range("K:L")
                    rg8.Select()
                    rg8.Delete()
                    xlWsheetDSS.Range("N:N").EntireColumn.Delete()
                    xlWsheetDSS.Range("Q:Q").EntireColumn.Delete()

                    xlWbookDSS.SaveAs(txtDSS_dest.Text)
                    xlWbookDSS.Close()
                    xlAppDSS.Quit()

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetDSS)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookDSS)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppDSS)

                    xlWsheetDSS = Nothing
                    xlWbookDSS = Nothing
                    xlAppDSS = Nothing

                    MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Catch ex As Exception
                    MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            Case "PRODUK"
                Try
                    xlAppDSS = New Excel.Application
                    xlWbookDSS = xlAppDSS.Workbooks.Open(txtDSS_src.Text)
                    xlWsheetDSS = xlWbookDSS.Worksheets("UID Daily Sales Summary Report")

                    xlWsheetDSS.UsedRange.UnMerge()
                    xlWsheetDSS.UsedRange.WrapText = False
                    xlWsheetDSS.UsedRange.ColumnWidth = 15
                    xlWsheetDSS.UsedRange.RowHeight = 15

                    Dim rg_head_cut1 As Excel.Range = xlWsheetDSS.Range("B2")
                    Dim rg_head_paste1 As Excel.Range = xlWsheetDSS.Range("A2")
                    rg_head_cut1.Select()
                    rg_head_cut1.Cut(rg_head_paste1)

                    Dim rg_head_cut2 As Excel.Range = xlWsheetDSS.Range("B4")
                    Dim rg_head_paste2 As Excel.Range = xlWsheetDSS.Range("A3")
                    rg_head_cut2.Select()
                    rg_head_cut2.Cut(rg_head_paste2)

                    xlWsheetDSS.Range("A2").RowHeight = 27

                    Dim paramhead1, paramhead2, paramhead3 As String
                    paramhead1 = xlWsheetDSS.Range("C6").Value & " " & xlWsheetDSS.Range("F6").Value & "; "
                    paramhead1 = paramhead1 & xlWsheetDSS.Range("H6").Value & " " & xlWsheetDSS.Range("K6").Value & "; "

                    paramhead2 = xlWsheetDSS.Range("O6").Value & " " & xlWsheetDSS.Range("S6").Value & "; "
                    paramhead2 = paramhead2 & xlWsheetDSS.Range("Z6").Value & " " & xlWsheetDSS.Range("AC6").Value & "; "

                    paramhead3 = xlWsheetDSS.Range("AE6").Value & " " & xlWsheetDSS.Range("AG6").Value

                    xlWsheetDSS.Range("A5").Value = paramhead1
                    xlWsheetDSS.Range("A5").EntireRow.Font.Name = "Calibri"
                    xlWsheetDSS.Range("A6").Value = paramhead2
                    xlWsheetDSS.Range("A6").EntireRow.Font.Name = "Calibri"
                    xlWsheetDSS.Range("A7").Value = paramhead3
                    xlWsheetDSS.Range("A7").EntireColumn.Font.Name = "Calibri"

                    Dim rg1, rg2, rg3, rg4, rg5, rg6 As Excel.Range
                    rg1 = xlWsheetDSS.Range("B:C")
                    rg1.Select()
                    rg1.Delete()
                    rg2 = xlWsheetDSS.Range("C:F")
                    rg2.Select()
                    rg2.Delete()
                    rg3 = xlWsheetDSS.Range("D:E")
                    rg3.Select()
                    rg3.Delete()
                    xlWsheetDSS.Range("E:E").EntireColumn.Delete()

                    Dim lrow1 As Long = xlWsheetDSS.Range("G1").End(Excel.XlDirection.xlDown).Row
                    Dim r1 As Integer = lrow1 + 3
                    Dim cr1 As String = "G" & lrow1 & ":G" & r1
                    Dim cr2 As String = "B" & lrow1 & ":B" & r1

                    Dim cellrange As Excel.Range = xlWsheetDSS.Range(cr1)
                    cellrange.Select()
                    cellrange.Copy(xlWsheetDSS.Range(cr2))

                    rg4 = xlWsheetDSS.Range("F:G")
                    rg4.Select()
                    rg4.Delete()
                    rg5 = xlWsheetDSS.Range("G:H")
                    rg5.Select()
                    rg5.Delete()
                    xlWsheetDSS.Range("H:H").EntireColumn.Delete()
                    xlWsheetDSS.Range("I:I").EntireColumn.Delete()
                    rg6 = xlWsheetDSS.Range("J:S")
                    rg6.Select()
                    rg6.Delete()

                    xlWbookDSS.SaveAs(txtDSS_dest.Text)
                    xlWbookDSS.Close()
                    xlAppDSS.Quit()

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetDSS)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookDSS)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppDSS)

                    xlWsheetDSS = Nothing
                    xlWbookDSS = Nothing
                    xlAppDSS = Nothing

                    MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Catch ex As Exception
                    MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
        End Select
    End Sub
End Class