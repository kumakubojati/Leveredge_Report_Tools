Imports Microsoft.Office.Interop
Imports System.Windows.Forms
Imports System.Text
Imports System.Drawing
Imports Microsoft.Win32
Public Class frmDSR
    Dim AppsOffice As String
    Private Sub frmDSR_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim TargetKey As RegistryKey
        TargetKey = Registry.ClassesRoot.OpenSubKey("Excel.Application")
        If TargetKey Is Nothing Then
            AppsOffice = "XL_NotInstalled"
        Else
            AppsOffice = "XL_Installed"
            TargetKey.Close()
        End If
    End Sub

    Private Sub btnBrowDistStock_src_Click(sender As Object, e As EventArgs) Handles btnBrowDistStock_src.Click
        Dim diststockpathSrc As String
        If OFD_DSR.ShowDialog = DialogResult.OK Then
            diststockpathSrc = OFD_DSR.FileName
            txtDistStock_src.Text = diststockpathSrc
        End If
        If txtDistStock_Dest.Text <> "" Then
            btnNeu_DistStock.Enabled = True
        Else
            btnNeu_DistStock.Enabled = False
        End If
    End Sub

    Private Sub btnBrowDistStock_Dest_Click(sender As Object, e As EventArgs) Handles btnBrowDistStock_Dest.Click
        Dim datenow As DateTime = DateTime.Now
        Dim filename As String = "Neutralize_DistStock_" & datenow.ToString("ddMMyyyy_HHmm")
        SFD_DSR.FileName = filename
        Dim DistStockPath_Dest As String
        If SFD_DSR.ShowDialog = DialogResult.OK Then
            DistStockPath_Dest = SFD_DSR.FileName
            txtDistStock_Dest.Text = DistStockPath_Dest
        End If
        If txtDistStock_src.Text <> "" Then
            btnNeu_DistStock.Enabled = True
        Else
            btnNeu_DistStock.Enabled = False
        End If
        If txtDistStock_Dest.Text <> "" Then
            btnNeu_DistStock.Enabled = True
        Else
            btnNeu_DistStock.Enabled = False
        End If
    End Sub

    Private Sub btnNeu_DistStock_Click(sender As Object, e As EventArgs) Handles btnNeu_DistStock.Click
        PicBar_DistStock.Visible = True
        BWDSR.RunWorkerAsync()
    End Sub

    Private Sub BWDSR_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BWDSR.RunWorkerCompleted
        PicBar_DistStock.Visible = False
        btnNeu_DistStock.Enabled = False
        txtDistStock_Dest.Text = ""
        txtDistStock_src.Text = ""
    End Sub

    Private Sub BWDSR_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BWDSR.DoWork
        Dim xlAppDistStock As Excel.Application
        Dim xlWbookDistStock As Excel.Workbook
        Dim xlWsheetDistStock As Excel.Worksheet

        Try
            xlAppDistStock = New Excel.Application
            xlWbookDistStock = xlAppDistStock.Workbooks.Open(txtDistStock_src.Text)
            xlWsheetDistStock = xlWbookDistStock.Worksheets("UID Distributor Stock Report")

            xlWsheetDistStock.UsedRange.UnMerge()
            xlWsheetDistStock.UsedRange.WrapText = False
            xlWsheetDistStock.UsedRange.ColumnWidth = 15
            xlWsheetDistStock.UsedRange.RowHeight = 15

            Dim rg_head_cut1 As Excel.Range = xlWsheetDistStock.Range("B2")
            Dim rg_head_paste1 As Excel.Range = xlWsheetDistStock.Range("A2")
            rg_head_cut1.Select()
            rg_head_cut1.Cut(rg_head_paste1)

            Dim rg_head_cut2 As Excel.Range = xlWsheetDistStock.Range("B3")
            Dim rg_head_paste2 As Excel.Range = xlWsheetDistStock.Range("A3")
            rg_head_cut2.Select()
            rg_head_cut2.Cut(rg_head_paste2)

            xlWsheetDistStock.Range("A2").RowHeight = 27

            Dim paramhead1, paramhead2, paramhead3 As String
            paramhead1 = xlWsheetDistStock.Range("C5").Value & " " & xlWsheetDistStock.Range("G5").Value & "; "
            paramhead1 = paramhead1 & xlWsheetDistStock.Range("J5").Value & " " & xlWsheetDistStock.Range("M5").Value & "; "
            paramhead1 = paramhead1 & xlWsheetDistStock.Range("V5").Value & " " & xlWsheetDistStock.Range("Z5").Value

            paramhead2 = xlWsheetDistStock.Range("AC5").Value & " " & xlWsheetDistStock.Range("AG5").Value

            paramhead3 = xlWsheetDistStock.Range("C7").Value & " " & xlWsheetDistStock.Range("G7").Value & "; "
            paramhead3 = paramhead3 & xlWsheetDistStock.Range("J7").Value & " " & xlWsheetDistStock.Range("P7").Value & "; "
            paramhead3 = paramhead3 & xlWsheetDistStock.Range("Z7").Value & " " & xlWsheetDistStock.Range("AG7").Value

            xlWsheetDistStock.Range("A5").Value = paramhead1
            xlWsheetDistStock.Range("A5").EntireRow.Font.Name = "Calibri"
            xlWsheetDistStock.Range("A6").Value = paramhead2
            xlWsheetDistStock.Range("A6").EntireRow.Font.Name = "Calibri"
            xlWsheetDistStock.Range("A7").Value = paramhead3
            xlWsheetDistStock.Range("A7").EntireRow.Font.Name = "Calibri"

            Dim rg1, rg2, rg3, rg4, rg5, rg6, rg7, rg8, rg9, rg10 As Excel.Range
            rg1 = xlWsheetDistStock.Range("B:C")
            rg1.Select()
            rg1.Delete()
            rg2 = xlWsheetDistStock.Range("C:E")
            rg2.Select()
            rg2.Delete()
            rg3 = xlWsheetDistStock.Range("D:E")
            rg3.Select()
            rg3.Delete()
            rg4 = xlWsheetDistStock.Range("E:F")
            rg4.Select()
            rg4.Delete()
            rg5 = xlWsheetDistStock.Range("F:G")
            rg5.Select()
            rg5.Delete()

            xlWsheetDistStock.Range("G:G").EntireColumn.Delete()
            xlWsheetDistStock.Range("H:H").EntireColumn.Delete()

            rg6 = xlWsheetDistStock.Range("I:K")
            rg6.Select()
            rg6.Delete()

            xlWsheetDistStock.Range("J:J").EntireColumn.Delete()

            rg7 = xlWsheetDistStock.Range("K:L")
            rg7.Select()
            rg7.Delete()
            rg8 = xlWsheetDistStock.Range("L:O")
            rg8.Select()
            rg8.Delete()
            rg9 = xlWsheetDistStock.Range("M:N")
            rg9.Select()
            rg9.Delete()

            rg10 = xlWsheetDistStock.Range("A9:A10")
            rg10.Select()
            xlWsheetDistStock.Range("B9:B10").Copy(rg10)
            xlWsheetDistStock.Range("A9").Value = "Kode Produk"
            xlWsheetDistStock.Range("A9").Font.Color = System.Drawing.ColorTranslator.ToOle(Color.White)
            xlWsheetDistStock.Range("A9:A10").Merge()
            xlWsheetDistStock.Range("A9:A10").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetDistStock.Range("B9:B10").Merge()
            xlWsheetDistStock.Range("B9:B10").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetDistStock.Range("C9:C10").Merge()
            xlWsheetDistStock.Range("C9:C10").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetDistStock.Range("D9:D10").Merge()
            xlWsheetDistStock.Range("D9:D10").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetDistStock.Range("E9:G9").Merge()
            xlWsheetDistStock.Range("E9:G9").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetDistStock.Range("H9:H10").Merge()
            xlWsheetDistStock.Range("H9:H10").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetDistStock.Range("H9:H10").WrapText = True
            xlWsheetDistStock.Range("I9:I10").Merge()
            xlWsheetDistStock.Range("I9:I10").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetDistStock.Range("I9:I10").WrapText = True
            xlWsheetDistStock.Range("J9:J10").Merge()
            xlWsheetDistStock.Range("J9:J10").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetDistStock.Range("K9:K10").Merge()
            xlWsheetDistStock.Range("K9:K10").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetDistStock.Range("L9:L10").Merge()
            xlWsheetDistStock.Range("L9:L10").HorizontalAlignment = Excel.Constants.xlCenter

            xlWbookDistStock.SaveAs(txtDistStock_Dest.Text)
            xlWbookDistStock.Close()
            xlAppDistStock.Quit()

            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetDistStock)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookDistStock)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppDistStock)

            xlWsheetDistStock = Nothing
            xlWbookDistStock = Nothing
            xlAppDistStock = Nothing

            MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class