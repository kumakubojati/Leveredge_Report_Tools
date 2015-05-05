Imports Microsoft.Office.Interop
Imports System.Windows.Forms
Imports System.Text
Imports System.Drawing
Imports Microsoft.Win32
Public Class frmLP3
    Dim AppsOffice As String
    Private Sub frmLP3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim TargetKey As RegistryKey
        TargetKey = Registry.ClassesRoot.OpenSubKey("Excel.Application")
        If TargetKey Is Nothing Then
            AppsOffice = "XL_NotInstalled"
        Else
            AppsOffice = "XL_Installed"
            TargetKey.Close()
        End If
    End Sub

    Private Sub btnBrow_LP3_src_Click(sender As Object, e As EventArgs) Handles btnBrow_LP3_src.Click
        Dim lp3srcpath As String
        If OFD_LP3.ShowDialog = DialogResult.OK Then
            lp3srcpath = OFD_LP3.FileName
            txtLP3_src.Text = lp3srcpath
        End If
        If txtLP3_dest.Text <> "" Then
            btnNeuLP3.Enabled = True
        Else
            btnNeuLP3.Enabled = False
        End If
    End Sub

    Private Sub btnBrow_LP3_Dest_Click(sender As Object, e As EventArgs) Handles btnBrow_LP3_Dest.Click
        Dim datenow As DateTime = DateTime.Now
        Dim filename As String = "Neutralize_LP3_" & datenow.ToString("ddMMyyyy_HHmm")
        SFD_LP3.FileName = filename
        Dim lp3destpath As String
        If SFD_LP3.ShowDialog = DialogResult.OK Then
            lp3destpath = SFD_LP3.FileName
            txtLP3_dest.Text = lp3destpath
        End If
        If txtLP3_src.Text <> "" Then
            btnNeuLP3.Enabled = True
        Else
            btnNeuLP3.Enabled = False
        End If
        If txtLP3_dest.Text <> "" Then
            btnNeuLP3.Enabled = True
        Else
            btnNeuLP3.Enabled = False
        End If
    End Sub

    Private Sub btnNeuLP3_Click(sender As Object, e As EventArgs) Handles btnNeuLP3.Click
        If RBAll.Checked = True Then
            GoTo Gothrough
        ElseIf RBRupiah.Checked = True Then
            GoTo Gothrough
        ElseIf RBQty.Checked = True Then
            GoTo Gothrough
        Else
            Exit Sub
        End If
Gothrough: PicBarLP3.Visible = True
        BWLP3.RunWorkerAsync()
    End Sub

    Private Sub BWLP3_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BWLP3.RunWorkerCompleted
        PicBarLP3.Visible = False
        btnNeuLP3.Enabled = False
        txtLP3_dest.Text = ""
        txtLP3_src.Text = ""
    End Sub

    Private Sub BWLP3_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BWLP3.DoWork
        Dim xlAppLP3 As Excel.Application
        Dim xlWbookLP3 As Excel.Workbook
        Dim xlWsheetLP3 As Excel.Worksheet

        Select Case LP3RptType
            Case "ALL"
                Try
                    xlAppLP3 = New Excel.Application
                    xlWbookLP3 = xlAppLP3.Workbooks.Open(txtLP3_src.Text)
                    xlWsheetLP3 = xlWbookLP3.Worksheets("UID Weekly Stock and Sales Repo")

                    xlWsheetLP3.UsedRange.UnMerge()
                    xlWsheetLP3.UsedRange.WrapText = False
                    xlWsheetLP3.UsedRange.ColumnWidth = 15
                    xlWsheetLP3.UsedRange.RowHeight = 15

                    Dim rg_head_cut1 As Excel.Range = xlWsheetLP3.Range("B2")
                    Dim rg_head_paste1 As Excel.Range = xlWsheetLP3.Range("A2")
                    rg_head_cut1.Select()
                    rg_head_cut1.Cut(rg_head_paste1)

                    Dim rg_head_cut2 As Excel.Range = xlWsheetLP3.Range("B4")
                    Dim rg_head_paste2 As Excel.Range = xlWsheetLP3.Range("A4")
                    rg_head_cut2.Select()
                    rg_head_cut2.Cut(rg_head_paste2)

                    'xlWsheetLP3.Range("A2").Value = xlWsheetLP3.Range("B2").Value
                    xlWsheetLP3.Range("A2").RowHeight = 27
                    'xlWsheetLP3.Range("A2").Font.Bold = True
                    'xlWsheetLP3.Range("A2").Font.Name = "Calibri"
                    'xlWsheetLP3.Range("A2").Font.Size = "20"

                    Dim paramhead1 As String
                    paramhead1 = xlWsheetLP3.Range("C6").Value & " " & xlWsheetLP3.Range("F6").Value & "; "
                    paramhead1 = paramhead1 & xlWsheetLP3.Range("H6").Value & " " & xlWsheetLP3.Range("K6").Value
                    Dim paramhead2 As String
                    paramhead2 = xlWsheetLP3.Range("N6").Value & " " & xlWsheetLP3.Range("P6").Value & "; "
                    paramhead2 = paramhead2 & xlWsheetLP3.Range("U6").Value & " " & xlWsheetLP3.Range("Y6").Value

                    xlWsheetLP3.Range("A6").Value = paramhead1
                    xlWsheetLP3.Range("A6").EntireRow.Font.Name = "Calibri"
                    xlWsheetLP3.Range("A7").Value = paramhead2
                    xlWsheetLP3.Range("A7").EntireRow.Font.Name = "Calibri"

                    Dim rg1 As Excel.Range = xlWsheetLP3.Range("B:C")
                    rg1.EntireColumn.Select()
                    rg1.EntireColumn.Delete()
                    Dim rg2 As Excel.Range = xlWsheetLP3.Range("C:F")
                    rg2.EntireColumn.Select()
                    rg2.EntireColumn.Delete()
                    Dim rg3 As Excel.Range = xlWsheetLP3.Range("D:E")
                    rg3.EntireColumn.Select()
                    rg3.EntireColumn.Delete()
                    Dim rg4 As Excel.Range = xlWsheetLP3.Range("E:I")
                    rg4.EntireColumn.Select()
                    rg4.EntireColumn.Delete()
                    Dim rg5 As Excel.Range = xlWsheetLP3.Range("G:I")
                    rg5.EntireColumn.Select()
                    rg5.EntireColumn.Delete()
                    Dim rg6 As Excel.Range = xlWsheetLP3.Range("H:I")
                    rg6.EntireColumn.Select()
                    rg6.EntireColumn.Delete()
                    xlWsheetLP3.Range("I:I").EntireColumn.Delete()

                    Dim lrow1 As Long
                    lrow1 = xlWsheetLP3.Range("X65536").End(Excel.XlDirection.xlUp).Row

                    Dim r As Integer = lrow1 + 3
                    Dim cr As String = "X" & r & ":X65536"

                    Dim cellrange As Excel.Range = xlWsheetLP3.Range(cr)
                    cellrange.Select()
                    cellrange.Delete(Shift:=Excel.XlDirection.xlToLeft)

                    xlWsheetLP3.Range("Y:Y").EntireColumn.Delete()

                    Dim lrow2 As Long
                    lrow2 = xlWsheetLP3.Range("Z1").End(Excel.XlDirection.xlDown).Row

                    Dim r2 As Integer = lrow2 - 3
                    Dim cr2 As String = "Z1:Z" & r2

                    Dim cellrange2 As Excel.Range = xlWsheetLP3.Range(cr2)
                    cellrange2.Select()
                    cellrange2.Delete(Shift:=Excel.XlDirection.xlToLeft)

                    Dim rg7 As Excel.Range = xlWsheetLP3.Range("AA:AB")
                    rg7.EntireColumn.Select()
                    rg7.EntireColumn.Delete()

                    xlWsheetLP3.Range("A3").EntireRow.Delete()

                    Dim lrow3 As Long
                    lrow3 = xlWsheetLP3.Range("B65536").End(Excel.XlDirection.xlUp).Row

                    Dim r3 As Integer = lrow3
                    Dim cr3 As String = "B" & r3 & ":B65536"

                    Dim cellrange3 As Excel.Range = xlWsheetLP3.Range(cr3)
                    cellrange3.Select()
                    cellrange3.Delete(Shift:=Excel.XlDirection.xlToLeft)

                    Dim lrow4, lrow5 As Long
                    lrow4 = xlWsheetLP3.Range("A65536").End(Excel.XlDirection.xlUp).Row

                    Dim r4 As Integer = lrow4 - 1

                    Dim cr4 As String = "A" & r4

                    lrow5 = xlWsheetLP3.Range(cr4).End(Excel.XlDirection.xlUp).Row
                    Dim r5 As Integer = lrow5 - 1
                    Dim r6 As Integer = lrow5 - 2
                    Dim cr5 As String = "A" & r5 & ":A" & r6

                    xlWsheetLP3.Range(cr5).EntireRow.Delete()

                    xlWsheetLP3.Range("AB:AB").EntireColumn.Delete()

                    xlWbookLP3.SaveAs(txtLP3_dest.Text)
                    xlWbookLP3.Close()
                    xlAppLP3.Quit()

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetLP3)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookLP3)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppLP3)

                    xlWsheetLP3 = Nothing
                    xlWbookLP3 = Nothing
                    xlAppLP3 = Nothing

                    MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Catch ex As Exception
                    MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try

            Case "RPH"
                Try
                    xlAppLP3 = New Excel.Application
                    xlWbookLP3 = xlAppLP3.Workbooks.Open(txtLP3_src.Text)
                    xlWsheetLP3 = xlWbookLP3.Worksheets("UID Weekly Stock and Sales Repo")

                    xlWsheetLP3.UsedRange.UnMerge()
                    xlWsheetLP3.UsedRange.WrapText = False
                    xlWsheetLP3.UsedRange.ColumnWidth = 15
                    xlWsheetLP3.UsedRange.RowHeight = 15

                    Dim rg_head_cut1 As Excel.Range = xlWsheetLP3.Range("B2")
                    Dim rg_head_paste1 As Excel.Range = xlWsheetLP3.Range("A2")
                    rg_head_cut1.Select()
                    rg_head_cut1.Cut(rg_head_paste1)

                    Dim rg_head_cut2 As Excel.Range = xlWsheetLP3.Range("B4")
                    Dim rg_head_paste2 As Excel.Range = xlWsheetLP3.Range("A4")
                    rg_head_cut2.Select()
                    rg_head_cut2.Cut(rg_head_paste2)

                    xlWsheetLP3.Range("A2").RowHeight = 27

                    Dim paramhead1 As String
                    paramhead1 = xlWsheetLP3.Range("C6").Value & " " & xlWsheetLP3.Range("F6").Value & "; "
                    paramhead1 = paramhead1 & xlWsheetLP3.Range("H6").Value & " " & xlWsheetLP3.Range("K6").Value
                    Dim paramhead2 As String
                    paramhead2 = xlWsheetLP3.Range("N6").Value & " " & xlWsheetLP3.Range("P6").Value & "; "
                    paramhead2 = paramhead2 & xlWsheetLP3.Range("U6").Value & " " & xlWsheetLP3.Range("Y6").Value

                    xlWsheetLP3.Range("A6").Value = paramhead1
                    xlWsheetLP3.Range("A6").EntireRow.Font.Name = "Calibri"
                    xlWsheetLP3.Range("A7").Value = paramhead2
                    xlWsheetLP3.Range("A7").EntireRow.Font.Name = "Calibri"

                    Dim rg1, rg2, rg3, rg4, rg5 As Excel.Range
                    rg1 = xlWsheetLP3.Range("B:H")
                    rg1.Select()
                    rg1.Delete()
                    rg2 = xlWsheetLP3.Range("C:D")
                    rg2.Select()
                    rg2.Delete()
                    rg3 = xlWsheetLP3.Range("D:H")
                    rg3.Select()
                    rg3.Delete()
                    rg4 = xlWsheetLP3.Range("F:H")
                    rg4.Select()
                    rg4.Delete()
                    rg5 = xlWsheetLP3.Range("G:H")
                    rg5.Select()
                    rg5.Delete()

                    xlWsheetLP3.Range("H:H").EntireColumn.Delete()
                    xlWsheetLP3.Range("A3").EntireRow.Delete()

                    xlWbookLP3.SaveAs(txtLP3_dest.Text)
                    xlWbookLP3.Close()
                    xlAppLP3.Quit()

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetLP3)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookLP3)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppLP3)

                    xlWsheetLP3 = Nothing
                    xlWbookLP3 = Nothing
                    xlAppLP3 = Nothing

                    MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                Catch ex As Exception
                    MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try

            Case "QTY"
                Try
                    xlAppLP3 = New Excel.Application
                    xlWbookLP3 = xlAppLP3.Workbooks.Open(txtLP3_src.Text)
                    xlWsheetLP3 = xlWbookLP3.Worksheets("UID Weekly Stock and Sales Repo")

                    xlWsheetLP3.UsedRange.UnMerge()
                    xlWsheetLP3.UsedRange.WrapText = False
                    xlWsheetLP3.UsedRange.ColumnWidth = 15
                    xlWsheetLP3.UsedRange.RowHeight = 15

                    Dim rg_head_cut1 As Excel.Range = xlWsheetLP3.Range("B2")
                    Dim rg_head_paste1 As Excel.Range = xlWsheetLP3.Range("A2")
                    rg_head_cut1.Select()
                    rg_head_cut1.Cut(rg_head_paste1)

                    Dim rg_head_cut2 As Excel.Range = xlWsheetLP3.Range("B4")
                    Dim rg_head_paste2 As Excel.Range = xlWsheetLP3.Range("A4")
                    rg_head_cut2.Select()
                    rg_head_cut2.Cut(rg_head_paste2)

                    xlWsheetLP3.Range("A2").RowHeight = 27

                    Dim paramhead1 As String
                    paramhead1 = xlWsheetLP3.Range("C6").Value & " " & xlWsheetLP3.Range("F6").Value & "; "
                    paramhead1 = paramhead1 & xlWsheetLP3.Range("H6").Value & " " & xlWsheetLP3.Range("K6").Value
                    Dim paramhead2 As String
                    paramhead2 = xlWsheetLP3.Range("N6").Value & " " & xlWsheetLP3.Range("P6").Value & "; "
                    paramhead2 = paramhead2 & xlWsheetLP3.Range("U6").Value & " " & xlWsheetLP3.Range("Y6").Value

                    xlWsheetLP3.Range("A6").Value = paramhead1
                    xlWsheetLP3.Range("A6").EntireRow.Font.Name = "Calibri"
                    xlWsheetLP3.Range("A7").Value = paramhead2
                    xlWsheetLP3.Range("A7").EntireRow.Font.Name = "Calibri"

                    Dim rg1, rg2, rg3, rg4, rg5, rg6, rg7 As Excel.Range
                    rg1 = xlWsheetLP3.Range("B:C")
                    rg1.Select()
                    rg1.Delete()
                    rg2 = xlWsheetLP3.Range("C:F")
                    rg2.Select()
                    rg2.Delete()
                    rg3 = xlWsheetLP3.Range("D:E")
                    rg3.Select()
                    rg3.Delete()
                    rg4 = xlWsheetLP3.Range("E:I")
                    rg4.Select()
                    rg4.Delete()
                    rg5 = xlWsheetLP3.Range("G:I")
                    rg5.Select()
                    rg5.Delete()
                    rg6 = xlWsheetLP3.Range("H:I")
                    rg6.Select()
                    rg6.Delete()

                    xlWsheetLP3.Range("I:I").EntireColumn.Delete()

                    rg7 = xlWsheetLP3.Range("AA:AB")
                    rg7.Select()
                    rg7.Delete()

                    xlWsheetLP3.Range("A3").EntireRow.Delete()

                    xlWbookLP3.SaveAs(txtLP3_dest.Text)
                    xlWbookLP3.Close()
                    xlAppLP3.Quit()

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetLP3)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookLP3)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppLP3)

                    xlWsheetLP3 = Nothing
                    xlWbookLP3 = Nothing
                    xlAppLP3 = Nothing

                    MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                Catch ex As Exception
                    MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try

        End Select
    End Sub
    Dim LP3RptType As String

    Private Sub RBAll_CheckedChanged(sender As Object, e As EventArgs) Handles RBAll.CheckedChanged
        LP3RptType = "ALL"
    End Sub

    Private Sub RBRupiah_CheckedChanged(sender As Object, e As EventArgs) Handles RBRupiah.CheckedChanged
        LP3RptType = "RPH"
    End Sub

    Private Sub RBQty_CheckedChanged(sender As Object, e As EventArgs) Handles RBQty.CheckedChanged
        LP3RptType = "QTY"
    End Sub
End Class