Imports Microsoft.Office.Interop
Imports System.Windows.Forms
Imports System.Text
Imports System.Drawing
Imports Microsoft.Win32

Public Class frmISRD_IQ
    Dim AppsOffice As String
    Private Sub frmISRD_IQ_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim TargetKey As RegistryKey
        TargetKey = Registry.ClassesRoot.OpenSubKey("Excel.Application")
        If TargetKey Is Nothing Then
            AppsOffice = "XL_NotInstalled"
        Else
            AppsOffice = "XL_Installed"
            TargetKey.Close()
        End If
    End Sub

    Private Sub btnBrow_ISRD_IQ_src_Click(sender As Object, e As EventArgs) Handles btnBrow_ISRD_IQ_src.Click
        Dim TCRpathSrc As String
        If OFD_ISRD_IQ.ShowDialog = DialogResult.OK Then
            TCRpathSrc = OFD_ISRD_IQ.FileName()
            txtISRD_IQ_src.Text = TCRpathSrc
        End If
        If txtISRD_IQ_dest.Text <> "" Then
            btnNeu_ISRD_IQ.Enabled = True
        Else
            btnNeu_ISRD_IQ.Enabled = False
        End If
    End Sub

    Private Sub btnBrow_ISRD_IQ_dest_Click(sender As Object, e As EventArgs) Handles btnBrow_ISRD_IQ_dest.Click
        Dim datenow As DateTime = DateTime.Now
        Dim filename As String = "Neutralize_IQ_SummaryReportDSR_" & datenow.ToString("ddMMyyyy_HHmm")
        SFD_ISRD_IQ.FileName = filename
        Dim TCRpath_Dest As String
        If SFD_ISRD_IQ.ShowDialog = DialogResult.OK Then
            TCRpath_Dest = SFD_ISRD_IQ.FileName
            txtISRD_IQ_dest.Text = TCRpath_Dest
        End If
        If txtISRD_IQ_src.Text <> "" Then
            btnNeu_ISRD_IQ.Enabled = True
        Else
            btnNeu_ISRD_IQ.Enabled = False
        End If
        If txtISRD_IQ_dest.Text <> "" Then
            btnNeu_ISRD_IQ.Enabled = True
        Else
            btnNeu_ISRD_IQ.Enabled = False
        End If
    End Sub

    Private Sub btnNeu_ISRD_IQ_Click(sender As Object, e As EventArgs) Handles btnNeu_ISRD_IQ.Click
        PicBar_ISRD_IQ.Visible = True
        BWISRD_IQ.RunWorkerAsync()
    End Sub

    Private Sub BWISRD_IQ_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BWISRD_IQ.DoWork
        Select Case AppsOffice
            Case "XL_NotInstalled"

            Case "XL_Installed"
                Dim xlAppISRD_IQ As Object
                Dim xlWbookISRD_IQ As Object
                Dim xlWsheetISRD_IQ As Object

                Try
                    xlAppISRD_IQ = CreateObject("Excel.Application")
                    xlWbookISRD_IQ = xlAppISRD_IQ.Workbooks.Open(txtISRD_IQ_src.Text)
                    xlWsheetISRD_IQ = xlWbookISRD_IQ.Worksheets("NGDMS IQ Summary Report for DSR")

                    xlWsheetISRD_IQ.UsedRange.UnMerge()
                    xlWsheetISRD_IQ.UsedRange.WrapText = False
                    xlWsheetISRD_IQ.UsedRange.ColumnWidth = 15
                    xlWsheetISRD_IQ.UsedRange.RowHeight = 15

                    Dim rg_head_cut1 As Excel.Range = xlWsheetISRD_IQ.Range("B2")
                    Dim rg_head_paste1 As Excel.Range = xlWsheetISRD_IQ.Range("A2")
                    rg_head_cut1.Select()
                    rg_head_cut1.Cut(rg_head_paste1)

                    Dim rg_head_cut2 As Excel.Range = xlWsheetISRD_IQ.Range("B4")
                    Dim rg_head_paste2 As Excel.Range = xlWsheetISRD_IQ.Range("A3")
                    rg_head_cut2.Select()
                    rg_head_cut2.Cut(rg_head_paste2)

                    xlWsheetISRD_IQ.Range("A2").RowHeight = 27

                    Dim paramhead1, paramhead2, paramhead3 As String
                    paramhead1 = xlWsheetISRD_IQ.Range("C6").Value & " " & xlWsheetISRD_IQ.Range("H6").Value & "; "
                    paramhead1 = paramhead1 & xlWsheetISRD_IQ.Range("L6").Value & " " & xlWsheetISRD_IQ.Range("O6").Value

                    paramhead2 = xlWsheetISRD_IQ.Range("S6").Value & " " & xlWsheetISRD_IQ.Range("W6").Value & ";"
                    paramhead2 = paramhead2 & xlWsheetISRD_IQ.Range("AA6").Value & " " & xlWsheetISRD_IQ.Range("AC6").Value & "; "

                    paramhead3 = xlWsheetISRD_IQ.Range("C8").Value & " " & xlWsheetISRD_IQ.Range("F8").Value

                    xlWsheetISRD_IQ.Range("A5").Value = paramhead1 + "; " + paramhead2
                    xlWsheetISRD_IQ.Range("A5").EntireRow.Font.Name = "Calibri"
                    xlWsheetISRD_IQ.Range("A6").Value = paramhead3
                    xlWsheetISRD_IQ.Range("A6").EntireRow.Font.Name = "Calibri"

                    xlWsheetISRD_IQ.Range("B6:AC8").value = ""
                    xlWsheetISRD_IQ.Range("A8:A9").EntireRow.Delete()

                    Dim xlfunc As Excel.WorksheetFunction
                    xlfunc = xlAppISRD_IQ.WorksheetFunction()
                    Dim lnCol As Long
                    Dim i As Long
                    Dim rnarea As Excel.Range = xlWsheetISRD_IQ.UsedRange

                    lnCol = rnarea.Columns.Count
                    For i = lnCol To 1 Step -1
                        If xlfunc.CountA(rnarea.Columns(i)) = 0 Then
                            rnarea.Columns(i).Delete()
                        End If
                    Next

                    xlWsheetISRD_IQ.Range("C9:E9").Merge()
                    xlWsheetISRD_IQ.Range("C9:E9").HorizontalAlignment = Excel.Constants.xlCenter

                    xlWsheetISRD_IQ.Range("F9:H9").Merge()
                    xlWsheetISRD_IQ.Range("F9:H9").HorizontalAlignment = Excel.Constants.xlCenter

                    xlWsheetISRD_IQ.Range("I9:K9").Merge()
                    xlWsheetISRD_IQ.Range("I9:K9").HorizontalAlignment = Excel.Constants.xlCenter

                    xlWsheetISRD_IQ.Range("L9:N9").Merge()
                    xlWsheetISRD_IQ.Range("L9:N9").HorizontalAlignment = Excel.Constants.xlCenter

                    xlWsheetISRD_IQ.Range("O9:Q9").Merge()
                    xlWsheetISRD_IQ.Range("O9:Q9").HorizontalAlignment = Excel.Constants.xlCenter

                    xlWsheetISRD_IQ.Range("R9:T9").Merge()
                    xlWsheetISRD_IQ.Range("R9:T9").HorizontalAlignment = Excel.Constants.xlCenter

                    xlWsheetISRD_IQ.Range("U9:W9").Merge()
                    xlWsheetISRD_IQ.Range("U9:W9").HorizontalAlignment = Excel.Constants.xlCenter

                    xlWsheetISRD_IQ.Range("X9:Z9").Merge()
                    xlWsheetISRD_IQ.Range("X9:Z9").HorizontalAlignment = Excel.Constants.xlCenter

                    xlWsheetISRD_IQ.Range("AA9:AC9").Merge()
                    xlWsheetISRD_IQ.Range("AA9:AC9").HorizontalAlignment = Excel.Constants.xlCenter

                    xlWsheetISRD_IQ.Range("A9:A10").Merge()
                    xlWsheetISRD_IQ.Range("A9:A10").HorizontalAlignment = Excel.Constants.xlCenter

                    xlWsheetISRD_IQ.Range("B9:B10").Merge()
                    xlWsheetISRD_IQ.Range("B9:B10").HorizontalAlignment = Excel.Constants.xlCenter

                    Dim lrow As Long
                    lrow = xlWsheetISRD_IQ.Range("A65536").End(Excel.XlDirection.xlUp).Row

                    Dim r As Integer = lrow - 2

                    xlWsheetISRD_IQ.Range("C" & r & ":E" & r).Merge()
                    xlWsheetISRD_IQ.Range("C" & r & ":E" & r).HorizontalAlignment = Excel.Constants.xlCenter

                    xlWsheetISRD_IQ.Range("F" & r & ":H" & r).Merge()
                    xlWsheetISRD_IQ.Range("F" & r & ":H" & r).HorizontalAlignment = Excel.Constants.xlCenter

                    xlWsheetISRD_IQ.Range("I" & r & ":K" & r).Merge()
                    xlWsheetISRD_IQ.Range("I" & r & ":K" & r).HorizontalAlignment = Excel.Constants.xlCenter

                    xlWsheetISRD_IQ.Range("L" & r & ":N" & r).Merge()
                    xlWsheetISRD_IQ.Range("L" & r & ":N" & r).HorizontalAlignment = Excel.Constants.xlCenter

                    xlWsheetISRD_IQ.Range("O" & r & ":Q" & r).Merge()
                    xlWsheetISRD_IQ.Range("O" & r & ":Q" & r).HorizontalAlignment = Excel.Constants.xlCenter

                    xlWsheetISRD_IQ.Range("R" & r & ":T" & r).Merge()
                    xlWsheetISRD_IQ.Range("R" & r & ":T" & r).HorizontalAlignment = Excel.Constants.xlCenter

                    xlWsheetISRD_IQ.Range("U" & r & ":W" & r).Merge()
                    xlWsheetISRD_IQ.Range("U" & r & ":W" & r).HorizontalAlignment = Excel.Constants.xlCenter

                    xlWsheetISRD_IQ.Range("X" & r & ":Z" & r).Merge()
                    xlWsheetISRD_IQ.Range("X" & r & ":Z" & r).HorizontalAlignment = Excel.Constants.xlCenter

                    xlWsheetISRD_IQ.Range("AA" & r & ":AC" & r).Merge()
                    xlWsheetISRD_IQ.Range("AA" & r & ":AC" & r).HorizontalAlignment = Excel.Constants.xlCenter

                    Dim r2 As Integer = r + 1
                    xlWsheetISRD_IQ.Range("A" & r & ":A" & r2).Merge()
                    xlWsheetISRD_IQ.Range("A" & r & ":A" & r2).HorizontalAlignment = Excel.Constants.xlCenter

                    xlWsheetISRD_IQ.Range("B" & r & ":B" & r2).Merge()
                    xlWsheetISRD_IQ.Range("B" & r & ":B" & r2).HorizontalAlignment = Excel.Constants.xlCenter

                    xlWbookISRD_IQ.SaveAs(txtISRD_IQ_dest.Text)
                    xlWbookISRD_IQ.Close()
                    xlAppISRD_IQ.Quit()

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetISRD_IQ)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookISRD_IQ)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppISRD_IQ)

                    xlWsheetISRD_IQ = Nothing
                    xlWbookISRD_IQ = Nothing
                    xlAppISRD_IQ = Nothing

                    MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                Catch ex As Exception
                    MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
        End Select
    End Sub

    Private Sub BWISRD_IQ_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BWISRD_IQ.RunWorkerCompleted
        PicBar_ISRD_IQ.Visible = False
        btnNeu_ISRD_IQ.Enabled = False
        txtISRD_IQ_dest.Text = ""
        txtISRD_IQ_src.Text = ""
    End Sub
End Class