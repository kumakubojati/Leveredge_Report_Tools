Imports Microsoft.Office.Interop
Imports System.Windows.Forms
Imports System.Text
Imports System.Drawing
Imports Microsoft.Win32

Public Class frmSPIR
    Dim AppsOffice As String
    Private Sub frmSPIR_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim TargetKey As RegistryKey
        TargetKey = Registry.ClassesRoot.OpenSubKey("Excel.Application")
        If TargetKey Is Nothing Then
            AppsOffice = "XL_NotInstalled"
        Else
            AppsOffice = "XL_Installed"
            TargetKey.Close()
        End If
    End Sub

    Private Sub btnBrow_SPIR_src_Click(sender As Object, e As EventArgs) Handles btnBrow_SPIR_src.Click
        Dim DSMpathSrc As String
        If OFD_SPIR.ShowDialog = DialogResult.OK Then
            DSMpathSrc = OFD_SPIR.FileName
            txtSPIR_src.Text = DSMpathSrc
        End If
        If txtSPIR_dest.Text <> "" Then
            btnNeu_SPIR.Enabled = True
        Else
            btnNeu_SPIR.Enabled = False
        End If
    End Sub

    Private Sub btnBrow_SPIR_dest_Click(sender As Object, e As EventArgs) Handles btnBrow_SPIR_dest.Click
        Dim datenow As DateTime = DateTime.Now
        Dim filename As String = "Neutralize_SalesPerfIncentive_" & datenow.ToString("ddMMyyyy_HHmm")
        SFD_SPIR.FileName = filename
        Dim DSMpath_Dest As String
        If SFD_SPIR.ShowDialog = DialogResult.OK Then
            DSMpath_Dest = SFD_SPIR.FileName
            txtSPIR_dest.Text = DSMpath_Dest
        End If
        If txtSPIR_src.Text <> "" Then
            btnNeu_SPIR.Enabled = True
        Else
            btnNeu_SPIR.Enabled = False
        End If
        If txtSPIR_dest.Text <> "" Then
            btnNeu_SPIR.Enabled = True
        Else
            btnNeu_SPIR.Enabled = False
        End If
    End Sub

    Private Sub btnNeu_SPIR_Click(sender As Object, e As EventArgs) Handles btnNeu_SPIR.Click
        PicBar_SPIR.Visible = True
        BWSPIR.RunWorkerAsync()
    End Sub

    Private Sub BWSPIR_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BWSPIR.DoWork
        Select Case AppsOffice
            Case "XL_NotInstalled"
                Dim xlAppSPIR As Object
                Dim xlWbookSPIR As Object
                Dim xlWsheetSPIR As Object

                Try
                    xlAppSPIR = CreateObject("Ket.Application")
                    xlWbookSPIR = xlAppSPIR.Workbooks.Open(txtSPIR_src.Text)
                    xlWsheetSPIR = xlWbookSPIR.Worksheets("UID Salesman Performance Incent")

                    xlWsheetSPIR.UsedRange.UnMerge()
                    xlWsheetSPIR.UsedRange.WrapText = False
                    xlWsheetSPIR.UsedRange.ColumnWidth = 15
                    xlWsheetSPIR.UsedRange.RowHeight = 15

                    xlWsheetSPIR.Range("A:A").EntireColumn.Delete()

                    Dim rg_head_cut1 As Object = xlWsheetSPIR.Range("B2")
                    Dim rg_head_paste1 As Object = xlWsheetSPIR.Range("A2")
                    rg_head_cut1.Select()
                    rg_head_cut1.Cut(rg_head_paste1)

                    Dim rg_head_cut2 As Object = xlWsheetSPIR.Range("B4")
                    Dim rg_head_paste2 As Object = xlWsheetSPIR.Range("A3")
                    rg_head_cut2.Select()
                    rg_head_cut2.Cut(rg_head_paste2)

                    xlWsheetSPIR.Range("A2").RowHeight = 27

                    Dim paramhead1, paramhead2 As String
                    paramhead1 = xlWsheetSPIR.Range("C6").Value & " " & xlWsheetSPIR.Range("F6").Value & "; "
                    paramhead1 = paramhead1 & xlWsheetSPIR.Range("H6").Value & " " & xlWsheetSPIR.Range("K6").Value

                    paramhead2 = xlWsheetSPIR.Range("O6").Value & " " & xlWsheetSPIR.Range("R6").Value & "; "

                    xlWsheetSPIR.Range("A5").Value = paramhead1
                    xlWsheetSPIR.Range("A5").EntireRow.Font.Name = "Calibri"
                    xlWsheetSPIR.Range("A6").Value = paramhead2
                    xlWsheetSPIR.Range("A6").EntireRow.Font.Name = "Calibri"

                    xlWsheetSPIR.Range("C6:R6").Value = ""

                    Dim xlfunc As Excel.WorksheetFunction
                    xlfunc = xlAppSPIR.WorksheetFunction
                    Dim lnCol As Long
                    Dim i As Long
                    Dim rnarea As Object = xlWsheetSPIR.UsedRange

                    lnCol = rnarea.Columns.Count
                    For i = lnCol To 1 Step -1
                        If xlfunc.CountA(rnarea.Columns(i)) = 0 Then
                            rnarea.Columns(i).Delete()
                        End If
                    Next

                    xlWsheetSPIR.Range("A8:A9").Merge()
                    xlWsheetSPIR.Range("A8:A9").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetSPIR.Range("B8:B9").Merge()
                    xlWsheetSPIR.Range("B8:B9").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetSPIR.Range("C8:C9").Merge()
                    xlWsheetSPIR.Range("C8:C9").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetSPIR.Range("D8:D9").Merge()
                    xlWsheetSPIR.Range("D8:D9").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetSPIR.Range("E8:G8").Merge()
                    xlWsheetSPIR.Range("E8:G8").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetSPIR.Range("H8:H9").Merge()
                    xlWsheetSPIR.Range("H8:H9").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetSPIR.Range("I8:I9").Merge()
                    xlWsheetSPIR.Range("I8:I9").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetSPIR.Range("J8:L8").Merge()
                    xlWsheetSPIR.Range("J8:L8").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetSPIR.Range("M8:M9").Merge()
                    xlWsheetSPIR.Range("M8:M9").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetSPIR.Range("N8:N9").Merge()
                    xlWsheetSPIR.Range("N8:N9").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetSPIR.Range("N8:N9").WrapText = True
                    xlWsheetSPIR.Range("O8:O9").Merge()
                    xlWsheetSPIR.Range("O8:O9").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetSPIR.Range("O8:O9").WrapText = True
                    xlWsheetSPIR.Range("P8:R8").Merge()
                    xlWsheetSPIR.Range("P8:R8").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetSPIR.Range("S8:S9").Merge()
                    xlWsheetSPIR.Range("S8:S9").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetSPIR.Range("T8:T9").Merge()
                    xlWsheetSPIR.Range("T8:T9").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetSPIR.Range("U8:U9").Merge()
                    xlWsheetSPIR.Range("U8:U9").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetSPIR.Range("V8:V9").Merge()
                    xlWsheetSPIR.Range("V8:V9").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetSPIR.Range("W8:W9").Merge()
                    xlWsheetSPIR.Range("W8:W9").HorizontalAlignment = Excel.Constants.xlCenter

                    xlWbookSPIR.SaveAs(txtSPIR_dest.Text)
                    xlWbookSPIR.Close()
                    xlAppSPIR.Quit()

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetSPIR)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookSPIR)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppSPIR)

                    xlWsheetSPIR = Nothing
                    xlWbookSPIR = Nothing
                    xlAppSPIR = Nothing

                    MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                Catch ex As Exception
                    MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try

            Case "XL_Installed"
                Dim xlAppSPIR As Object
                Dim xlWbookSPIR As Object
                Dim xlWsheetSPIR As Object

                Try
                    xlAppSPIR = CreateObject("Excel.Application")
                    xlWbookSPIR = xlAppSPIR.Workbooks.Open(txtSPIR_src.Text)
                    xlWsheetSPIR = xlWbookSPIR.Worksheets("UID Salesman Performance Incent")

                    xlWsheetSPIR.UsedRange.UnMerge()
                    xlWsheetSPIR.UsedRange.WrapText = False
                    xlWsheetSPIR.UsedRange.ColumnWidth = 15
                    xlWsheetSPIR.UsedRange.RowHeight = 15

                    xlWsheetSPIR.Range("A:A").EntireColumn.Delete()

                    Dim rg_head_cut1 As Excel.Range = xlWsheetSPIR.Range("B2")
                    Dim rg_head_paste1 As Excel.Range = xlWsheetSPIR.Range("A2")
                    rg_head_cut1.Select()
                    rg_head_cut1.Cut(rg_head_paste1)

                    Dim rg_head_cut2 As Excel.Range = xlWsheetSPIR.Range("B4")
                    Dim rg_head_paste2 As Excel.Range = xlWsheetSPIR.Range("A3")
                    rg_head_cut2.Select()
                    rg_head_cut2.Cut(rg_head_paste2)

                    xlWsheetSPIR.Range("A2").RowHeight = 27

                    Dim paramhead1, paramhead2 As String
                    paramhead1 = xlWsheetSPIR.Range("C6").Value & " " & xlWsheetSPIR.Range("F6").Value & "; "
                    paramhead1 = paramhead1 & xlWsheetSPIR.Range("H6").Value & " " & xlWsheetSPIR.Range("K6").Value

                    paramhead2 = xlWsheetSPIR.Range("O6").Value & " " & xlWsheetSPIR.Range("R6").Value & "; "

                    xlWsheetSPIR.Range("A5").Value = paramhead1
                    xlWsheetSPIR.Range("A5").EntireRow.Font.Name = "Calibri"
                    xlWsheetSPIR.Range("A6").Value = paramhead2
                    xlWsheetSPIR.Range("A6").EntireRow.Font.Name = "Calibri"

                    xlWsheetSPIR.Range("C6:R6").Value = ""

                    Dim xlfunc As Excel.WorksheetFunction
                    xlfunc = xlAppSPIR.WorksheetFunction
                    Dim lnCol As Long
                    Dim i As Long
                    Dim rnarea As Excel.Range = xlWsheetSPIR.UsedRange

                    lnCol = rnarea.Columns.Count
                    For i = lnCol To 1 Step -1
                        If xlfunc.CountA(rnarea.Columns(i)) = 0 Then
                            rnarea.Columns(i).Delete()
                        End If
                    Next

                    xlWsheetSPIR.Range("A8:A9").Merge()
                    xlWsheetSPIR.Range("A8:A9").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetSPIR.Range("B8:B9").Merge()
                    xlWsheetSPIR.Range("B8:B9").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetSPIR.Range("C8:C9").Merge()
                    xlWsheetSPIR.Range("C8:C9").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetSPIR.Range("D8:D9").Merge()
                    xlWsheetSPIR.Range("D8:D9").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetSPIR.Range("E8:G8").Merge()
                    xlWsheetSPIR.Range("E8:G8").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetSPIR.Range("H8:H9").Merge()
                    xlWsheetSPIR.Range("H8:H9").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetSPIR.Range("I8:I9").Merge()
                    xlWsheetSPIR.Range("I8:I9").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetSPIR.Range("J8:L8").Merge()
                    xlWsheetSPIR.Range("J8:L8").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetSPIR.Range("M8:M9").Merge()
                    xlWsheetSPIR.Range("M8:M9").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetSPIR.Range("N8:N9").Merge()
                    xlWsheetSPIR.Range("N8:N9").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetSPIR.Range("N8:N9").WrapText = True
                    xlWsheetSPIR.Range("O8:O9").Merge()
                    xlWsheetSPIR.Range("O8:O9").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetSPIR.Range("O8:O9").WrapText = True
                    xlWsheetSPIR.Range("P8:R8").Merge()
                    xlWsheetSPIR.Range("P8:R8").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetSPIR.Range("S8:S9").Merge()
                    xlWsheetSPIR.Range("S8:S9").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetSPIR.Range("T8:T9").Merge()
                    xlWsheetSPIR.Range("T8:T9").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetSPIR.Range("U8:U9").Merge()
                    xlWsheetSPIR.Range("U8:U9").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetSPIR.Range("V8:V9").Merge()
                    xlWsheetSPIR.Range("V8:V9").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetSPIR.Range("W8:W9").Merge()
                    xlWsheetSPIR.Range("W8:W9").HorizontalAlignment = Excel.Constants.xlCenter

                    xlWbookSPIR.SaveAs(txtSPIR_dest.Text)
                    xlWbookSPIR.Close()
                    xlAppSPIR.Quit()

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetSPIR)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookSPIR)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppSPIR)

                    xlWsheetSPIR = Nothing
                    xlWbookSPIR = Nothing
                    xlAppSPIR = Nothing

                    MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                Catch ex As Exception
                    MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try

        End Select
    End Sub

    Private Sub BWSPIR_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BWSPIR.RunWorkerCompleted
        PicBar_SPIR.Visible = False
        btnNeu_SPIR.Enabled = False
        txtSPIR_dest.Text = ""
        txtSPIR_src.Text = ""
    End Sub
End Class