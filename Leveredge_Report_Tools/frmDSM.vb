Imports Microsoft.Office.Interop
Imports System.Windows.Forms
Imports System.Text
Imports System.Drawing
Imports Microsoft.Win32

Public Class frmDSM
    Dim AppsOffice As String
    Private Sub frmDSM_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim TargetKey As RegistryKey
        TargetKey = Registry.ClassesRoot.OpenSubKey("Excel.Application")
        If TargetKey Is Nothing Then
            AppsOffice = "XL_NotInstalled"
        Else
            AppsOffice = "XL_Installed"
            TargetKey.Close()
        End If
    End Sub

    Private Sub btnBrow_DSM_src_Click(sender As Object, e As EventArgs) Handles btnBrow_DSM_src.Click
        Dim DSMpathSrc As String
        If OFD_DSM.ShowDialog = DialogResult.OK Then
            DSMpathSrc = OFD_DSM.FileName
            txtDSM_src.Text = DSMpathSrc
        End If
        If txtDSM_dest.Text <> "" Then
            btnNeu_DSM.Enabled = True
        Else
            btnNeu_DSM.Enabled = False
        End If
    End Sub

    Private Sub btnBrow_DSM_dest_Click(sender As Object, e As EventArgs) Handles btnBrow_DSM_dest.Click
        Dim datenow As DateTime = DateTime.Now
        Dim filename As String = "Neutralize_DailyStockMut_" & DSMreptype & "_" & datenow.ToString("ddMMyyyy_HHmm")
        SFD_DSM.FileName = filename
        Dim DSMpath_Dest As String
        If SFD_DSM.ShowDialog = DialogResult.OK Then
            DSMpath_Dest = SFD_DSM.FileName
            txtDSM_dest.Text = DSMpath_Dest
        End If
        If txtDSM_src.Text <> "" Then
            btnNeu_DSM.Enabled = True
        Else
            btnNeu_DSM.Enabled = False
        End If
        If txtDSM_dest.Text <> "" Then
            btnNeu_DSM.Enabled = True
        Else
            btnNeu_DSM.Enabled = False
        End If
    End Sub

    Private Sub btnNeu_DSM_Click(sender As Object, e As EventArgs) Handles btnNeu_DSM.Click
        PicBar_DSM.Visible = True
        BWDSM.RunWorkerAsync()
    End Sub

    Private Sub BWDSM_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BWDSM.RunWorkerCompleted
        PicBar_DSM.Visible = False
        btnNeu_DSM.Enabled = False
        txtDSM_dest.Text = ""
        txtDSM_src.Text = ""
    End Sub

    Dim DSMreptype As String
    Private Sub BWDSM_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BWDSM.DoWork
        Select Case AppsOffice
            Case "XL_NotInstalled"
                Dim xlAppDSM As Object
                Dim xlWbookDSM As Object
                Dim xlWsheetDSM As Object

                Select Case DSMreptype
                    Case "QTY"
                        Try
                            xlAppDSM = CreateObject("Ket.Application")
                            xlWbookDSM = xlAppDSM.Workbooks.Open(txtDSM_src.Text)
                            xlWsheetDSM = xlWbookDSM.Worksheets("UID Daily Stock Mutation Report")

                            xlWsheetDSM.UsedRange.UnMerge()
                            xlWsheetDSM.UsedRange.WrapText = False
                            xlWsheetDSM.UsedRange.ColumnWidth = 15
                            xlWsheetDSM.UsedRange.RowHeight = 15

                            Dim rg_head_cut1 As Object = xlWsheetDSM.Range("B2")
                            Dim rg_head_paste1 As Object = xlWsheetDSM.Range("A2")
                            rg_head_cut1.Select()
                            rg_head_cut1.Cut(rg_head_paste1)

                            Dim rg_head_cut2 As Object = xlWsheetDSM.Range("B4")
                            Dim rg_head_paste2 As Object = xlWsheetDSM.Range("A3")
                            rg_head_cut2.Select()
                            rg_head_cut2.Cut(rg_head_paste2)

                            xlWsheetDSM.Range("A2").RowHeight = 27

                            Dim paramhead1, paramhead2, paramhead3 As String
                            paramhead1 = xlWsheetDSM.Range("C6").Value & " " & xlWsheetDSM.Range("F6").Value & "; "
                            paramhead1 = paramhead1 & xlWsheetDSM.Range("J6").Value & " " & xlWsheetDSM.Range("M6").Value

                            paramhead2 = xlWsheetDSM.Range("R6").Value & " " & xlWsheetDSM.Range("U6").Value & "; "
                            paramhead2 = paramhead2 & xlWsheetDSM.Range("Y6").Value & " " & xlWsheetDSM.Range("AC6").Value

                            paramhead3 = xlWsheetDSM.Range("AG6").Value & " " & xlWsheetDSM.Range("AK6").Value & "; "
                            paramhead3 = paramhead3 & xlWsheetDSM.Range("AP6").Value & " " & xlWsheetDSM.Range("AS6").Value

                            xlWsheetDSM.Range("A5").Value = paramhead1
                            xlWsheetDSM.Range("A5").EntireRow.Font.Name = "Calibri"
                            xlWsheetDSM.Range("A6").Value = paramhead2
                            xlWsheetDSM.Range("A6").EntireRow.Font.Name = "Calibri"
                            xlWsheetDSM.Range("A7").Value = paramhead3
                            xlWsheetDSM.Range("A7").EntireRow.Font.Name = "Calibri"

                            xlWsheetDSM.Range("G8:G9").Copy(xlWsheetDSM.Range("A8:A9"))
                            xlWsheetDSM.Range("A8").Value = "PCODE"
                            xlWsheetDSM.Range("A8").Font.Color = System.Drawing.ColorTranslator.ToOle(Color.White)
                            xlWsheetDSM.Range("A8:A9").Merge()
                            xlWsheetDSM.Range("A8:A9").HorizontalAlignment = 3

                            Dim rg1, rg2, rg3, rg4, rg5, rg6, rg7, rg8, rg9, rg10, rg11 As Object
                            rg1 = xlWsheetDSM.Range("B:C")
                            rg1.Select()
                            rg1.Delete()

                            xlWsheetDSM.Range("E8:E9").Copy(xlWsheetDSM.Range("B8:B9"))
                            xlWsheetDSM.Range("B8").Value = "Nama Barang"
                            xlWsheetDSM.Range("B8").Font.Color = System.Drawing.ColorTranslator.ToOle(Color.White)
                            xlWsheetDSM.Range("B8:B9").Merge()
                            xlWsheetDSM.Range("B8:B9").HorizontalAlignment = 3

                            rg2 = xlWsheetDSM.Range("C:D")
                            rg2.Select()
                            rg2.Delete()

                            xlWsheetDSM.Range("C8:C9").Merge()
                            xlWsheetDSM.Range("C8:C9").HorizontalAlignment = 3
                            xlWsheetDSM.Range("C8:C9").WrapText = True
                            xlWsheetDSM.Range("D8:D9").Merge()
                            xlWsheetDSM.Range("D8:D9").HorizontalAlignment = 3
                            xlWsheetDSM.Range("D8:D9").WrapText = True

                            rg3 = xlWsheetDSM.Range("E:F")
                            rg3.Select()
                            rg3.Delete()

                            xlWsheetDSM.Range("E8:E9").Merge()
                            xlWsheetDSM.Range("E8:E9").HorizontalAlignment = 3
                            xlWsheetDSM.Range("E8:E9").WrapText = True

                            rg4 = xlWsheetDSM.Range("F:G")
                            rg4.Select()
                            rg4.Delete()

                            xlWsheetDSM.Range("F8:F9").Merge()
                            xlWsheetDSM.Range("F8:F9").HorizontalAlignment = 3
                            xlWsheetDSM.Range("F8:F9").WrapText = True

                            xlWsheetDSM.Range("G:G").EntireColumn.Delete()

                            rg5 = xlWsheetDSM.Range("H:I")
                            rg5.Select()
                            rg5.Delete()

                            xlWsheetDSM.Range("G8:H8").Merge()
                            xlWsheetDSM.Range("G8:H8").HorizontalAlignment = 3

                            rg6 = xlWsheetDSM.Range("I:J")
                            rg6.Select()
                            rg6.Delete()

                            xlWsheetDSM.Range("I8:I9").Merge()
                            xlWsheetDSM.Range("I8:I9").HorizontalAlignment = 3
                            xlWsheetDSM.Range("I8:I9").WrapText = True

                            xlWsheetDSM.Range("J8:J9").Merge()
                            xlWsheetDSM.Range("J8:J9").HorizontalAlignment = 3
                            xlWsheetDSM.Range("J8:J9").WrapText = True

                            rg7 = xlWsheetDSM.Range("K:L")
                            rg7.Select()
                            rg7.Delete()

                            xlWsheetDSM.Range("K8:K9").Merge()
                            xlWsheetDSM.Range("K8:K9").HorizontalAlignment = 3
                            xlWsheetDSM.Range("K8:K9").WrapText = True

                            rg8 = xlWsheetDSM.Range("L:N")
                            rg8.Select()
                            rg8.Delete()

                            xlWsheetDSM.Range("L8:L9").Merge()
                            xlWsheetDSM.Range("L8:L9").HorizontalAlignment = 3
                            xlWsheetDSM.Range("L8:L9").WrapText = True

                            xlWsheetDSM.Range("M8:M9").Merge()
                            xlWsheetDSM.Range("M8:M9").HorizontalAlignment = 3
                            xlWsheetDSM.Range("M8:M9").WrapText = True

                            rg9 = xlWsheetDSM.Range("N:O")
                            rg9.Select()
                            rg9.Delete()

                            xlWsheetDSM.Range("N8:N9").Merge()
                            xlWsheetDSM.Range("N8:N9").HorizontalAlignment = 3
                            xlWsheetDSM.Range("N8:N9").WrapText = True

                            xlWsheetDSM.Range("O:O").EntireColumn.Delete()

                            xlWsheetDSM.Range("O8:O9").Merge()
                            xlWsheetDSM.Range("O8:O9").HorizontalAlignment = 3
                            xlWsheetDSM.Range("O8:O9").WrapText = True

                            xlWsheetDSM.Range("P:P").EntireColumn.Delete()

                            xlWsheetDSM.Range("P8:P9").Merge()
                            xlWsheetDSM.Range("P8:P9").HorizontalAlignment = 3
                            xlWsheetDSM.Range("P8:P9").WrapText = True

                            xlWsheetDSM.Range("Q8:Q9").Merge()
                            xlWsheetDSM.Range("Q8:Q9").HorizontalAlignment = 3
                            xlWsheetDSM.Range("Q8:Q9").WrapText = True

                            xlWsheetDSM.Range("R8:R9").Merge()
                            xlWsheetDSM.Range("R8:R9").HorizontalAlignment = 3
                            xlWsheetDSM.Range("R8:R9").WrapText = True

                            rg10 = xlWsheetDSM.Range("S:T")
                            rg10.Select()
                            rg10.Delete()

                            xlWsheetDSM.Range("S8:S9").Merge()
                            xlWsheetDSM.Range("S8:S9").HorizontalAlignment = 3
                            xlWsheetDSM.Range("S8:S9").WrapText = True

                            rg11 = xlWsheetDSM.Range("T:U")
                            rg11.Select()
                            rg11.Delete()

                            xlWsheetDSM.Range("T8:T9").Merge()
                            xlWsheetDSM.Range("T8:T9").HorizontalAlignment = 3
                            xlWsheetDSM.Range("T8:T9").WrapText = True

                            xlWsheetDSM.Range("U8:U9").Merge()
                            xlWsheetDSM.Range("U8:U9").HorizontalAlignment = 3
                            xlWsheetDSM.Range("U8:U9").WrapText = True

                            xlWsheetDSM.Range("V8:V9").Merge()
                            xlWsheetDSM.Range("V8:V9").HorizontalAlignment = 3
                            xlWsheetDSM.Range("V8:V9").WrapText = True

                            xlWsheetDSM.Range("W8:W9").Merge()
                            xlWsheetDSM.Range("W8:W9").HorizontalAlignment = 3
                            xlWsheetDSM.Range("W8:W9").WrapText = True

                            xlWsheetDSM.Range("X:X").EntireColumn.Delete()

                            xlWsheetDSM.Range("X8:X9").Merge()
                            xlWsheetDSM.Range("X8:X9").HorizontalAlignment = 3
                            xlWsheetDSM.Range("X8:X9").WrapText = True

                            Dim lrow1 As Long
                            lrow1 = xlWsheetDSM.Range("A65536").End(Excel.XlDirection.xlUp).Row

                            Dim cr As String = "A" & lrow1 & ":B" & lrow1

                            xlWsheetDSM.Range(cr).Merge()
                            xlWsheetDSM.Range(cr).HorizontalAlignment = 4

                            xlWbookDSM.SaveAs(txtDSM_dest.Text)
                            xlWbookDSM.Close()
                            xlAppDSM.Quit()

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetDSM)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookDSM)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppDSM)

                            xlWsheetDSM = Nothing
                            xlWbookDSM = Nothing
                            xlAppDSM = Nothing

                            MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Catch ex As Exception
                            MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try

                    Case "RPH"
                        Try
                            xlAppDSM = CreateObject("Ket.Application")
                            xlWbookDSM = xlAppDSM.Workbooks.Open(txtDSM_src.Text)
                            xlWsheetDSM = xlWbookDSM.Worksheets("UID Daily Stock Mutation Report")

                            xlWsheetDSM.UsedRange.UnMerge()
                            xlWsheetDSM.UsedRange.WrapText = False
                            xlWsheetDSM.UsedRange.ColumnWidth = 15
                            xlWsheetDSM.UsedRange.RowHeight = 15

                            Dim rg_head_cut1 As Object = xlWsheetDSM.Range("B2")
                            Dim rg_head_paste1 As Object = xlWsheetDSM.Range("A2")
                            rg_head_cut1.Select()
                            rg_head_cut1.Cut(rg_head_paste1)

                            Dim rg_head_cut2 As Object = xlWsheetDSM.Range("B4")
                            Dim rg_head_paste2 As Object = xlWsheetDSM.Range("A3")
                            rg_head_cut2.Select()
                            rg_head_cut2.Cut(rg_head_paste2)

                            xlWsheetDSM.Range("A2").RowHeight = 27

                            Dim paramhead1, paramhead2, paramhead3 As String
                            paramhead1 = xlWsheetDSM.Range("C6").Value & " " & xlWsheetDSM.Range("F6").Value & "; "
                            paramhead1 = paramhead1 & xlWsheetDSM.Range("J6").Value & " " & xlWsheetDSM.Range("M6").Value

                            paramhead2 = xlWsheetDSM.Range("R6").Value & " " & xlWsheetDSM.Range("U6").Value & "; "
                            paramhead2 = paramhead2 & xlWsheetDSM.Range("Y6").Value & " " & xlWsheetDSM.Range("AC6").Value

                            paramhead3 = xlWsheetDSM.Range("AG6").Value & " " & xlWsheetDSM.Range("AK6").Value & "; "
                            paramhead3 = paramhead3 & xlWsheetDSM.Range("AP6").Value & " " & xlWsheetDSM.Range("AS6").Value

                            xlWsheetDSM.Range("A5").Value = paramhead1
                            xlWsheetDSM.Range("A5").EntireRow.Font.Name = "Calibri"
                            xlWsheetDSM.Range("A6").Value = paramhead2
                            xlWsheetDSM.Range("A6").EntireRow.Font.Name = "Calibri"
                            xlWsheetDSM.Range("A7").Value = paramhead3
                            xlWsheetDSM.Range("A7").EntireRow.Font.Name = "Calibri"

                            xlWsheetDSM.Range("G8:G9").Copy(xlWsheetDSM.Range("A8:A9"))
                            xlWsheetDSM.Range("A8").Value = "PCODE"
                            xlWsheetDSM.Range("A8").Font.Color = System.Drawing.ColorTranslator.ToOle(Color.White)
                            xlWsheetDSM.Range("A8:A9").Merge()
                            xlWsheetDSM.Range("A8:A9").HorizontalAlignment = 3

                            Dim rg1, rg2, rg3, rg4, rg5, rg6, rg7, rg8, rg9, rg10, rg11 As Object
                            rg1 = xlWsheetDSM.Range("B:C")
                            rg1.Select()
                            rg1.Delete()

                            xlWsheetDSM.Range("E8:E9").Copy(xlWsheetDSM.Range("B8:B9"))
                            xlWsheetDSM.Range("B8").Value = "Nama Barang"
                            xlWsheetDSM.Range("B8").Font.Color = System.Drawing.ColorTranslator.ToOle(Color.White)
                            xlWsheetDSM.Range("B8:B9").Merge()
                            xlWsheetDSM.Range("B8:B9").HorizontalAlignment = Excel.Constants.xlCenter

                            rg2 = xlWsheetDSM.Range("C:D")
                            rg2.Select()
                            rg2.Delete()
                            rg3 = xlWsheetDSM.Range("E:F")
                            rg3.Select()
                            rg3.Delete()
                            rg4 = xlWsheetDSM.Range("F:G")
                            rg4.Select()
                            rg4.Delete()
                            xlWsheetDSM.Range("G:G").EntireColumn.Delete()
                            rg5 = xlWsheetDSM.Range("H:I")
                            rg5.Select()
                            rg5.Delete()
                            rg6 = xlWsheetDSM.Range("I:J")
                            rg6.Select()
                            rg6.Delete()
                            rg7 = xlWsheetDSM.Range("K:M")
                            rg7.Select()
                            rg7.Delete()
                            rg8 = xlWsheetDSM.Range("L:N")
                            rg8.Select()
                            rg8.Delete()
                            rg9 = xlWsheetDSM.Range("N:O")
                            rg9.Select()
                            rg9.Delete()
                            xlWsheetDSM.Range("O:O").EntireColumn.Delete()
                            xlWsheetDSM.Range("P:P").EntireColumn.Delete()
                            rg10 = xlWsheetDSM.Range("S:T")
                            rg10.Select()
                            rg10.Delete()
                            rg11 = xlWsheetDSM.Range("T:U")
                            rg11.Select()
                            rg11.Delete()
                            xlWsheetDSM.Range("X:X").EntireColumn.Delete()

                            Dim lrow As Long
                            lrow = xlWsheetDSM.Range("A65536").End(Excel.XlDirection.xlUp).Row
                            Dim cr As String
                            cr = "A" & lrow & ":B" & lrow
                            xlWsheetDSM.Range(cr).Merge()
                            xlWsheetDSM.Range(cr).HorizontalAlignment = 4
                            Dim cr2 As String
                            cr2 = "A" & lrow
                            xlWsheetDSM.Range(cr2).EntireRow.Font.Bold = True

                            xlWsheetDSM.Range("C8:C9").Merge()
                            xlWsheetDSM.Range("C8:C9").HorizontalAlignment = 3
                            xlWsheetDSM.Range("C8:C9").WrapText = True
                            xlWsheetDSM.Range("D8:D9").Merge()
                            xlWsheetDSM.Range("D8:D9").HorizontalAlignment = 3
                            xlWsheetDSM.Range("D8:D9").WrapText = True
                            xlWsheetDSM.Range("E8:E9").Merge()
                            xlWsheetDSM.Range("E8:E9").HorizontalAlignment = 3
                            xlWsheetDSM.Range("E8:E9").WrapText = True
                            xlWsheetDSM.Range("F8:F9").Merge()
                            xlWsheetDSM.Range("F8:F9").HorizontalAlignment = 3
                            xlWsheetDSM.Range("F8:F9").WrapText = True
                            xlWsheetDSM.Range("G8:H8").Merge()
                            xlWsheetDSM.Range("G8:H8").HorizontalAlignment = 3
                            xlWsheetDSM.Range("G8:H8").WrapText = True
                            xlWsheetDSM.Range("G9:H9").HorizontalAlignment = 3
                            xlWsheetDSM.Range("I8:I9").Merge()
                            xlWsheetDSM.Range("I8:I9").HorizontalAlignment = 3
                            xlWsheetDSM.Range("I8:I9").WrapText = True
                            xlWsheetDSM.Range("J8:J9").Merge()
                            xlWsheetDSM.Range("J8:J9").HorizontalAlignment = 3
                            xlWsheetDSM.Range("J8:J9").WrapText = True
                            xlWsheetDSM.Range("K8:K9").Merge()
                            xlWsheetDSM.Range("K8:K9").HorizontalAlignment = 3
                            xlWsheetDSM.Range("K8:K9").WrapText = True
                            xlWsheetDSM.Range("L8:L9").Merge()
                            xlWsheetDSM.Range("L8:L9").HorizontalAlignment = 3
                            xlWsheetDSM.Range("L8:L9").WrapText = True
                            xlWsheetDSM.Range("M8:M9").Merge()
                            xlWsheetDSM.Range("M8:M9").HorizontalAlignment = 3
                            xlWsheetDSM.Range("M8:M9").WrapText = True
                            xlWsheetDSM.Range("N8:N9").Merge()
                            xlWsheetDSM.Range("N8:N9").HorizontalAlignment = 3
                            xlWsheetDSM.Range("N8:N9").WrapText = True
                            xlWsheetDSM.Range("N8:N9").Merge()
                            xlWsheetDSM.Range("N8:N9").HorizontalAlignment = 3
                            xlWsheetDSM.Range("N8:N9").WrapText = True
                            xlWsheetDSM.Range("O8:O9").Merge()
                            xlWsheetDSM.Range("O8:O9").HorizontalAlignment = 3
                            xlWsheetDSM.Range("O8:O9").WrapText = True
                            xlWsheetDSM.Range("P8:P9").Merge()
                            xlWsheetDSM.Range("P8:P9").HorizontalAlignment = 3
                            xlWsheetDSM.Range("P8:P9").WrapText = True
                            xlWsheetDSM.Range("Q8:Q9").Merge()
                            xlWsheetDSM.Range("Q8:Q9").HorizontalAlignment = 3
                            xlWsheetDSM.Range("Q8:Q9").WrapText = True
                            xlWsheetDSM.Range("R8:R9").Merge()
                            xlWsheetDSM.Range("R8:R9").HorizontalAlignment = 3
                            xlWsheetDSM.Range("R8:R9").WrapText = True
                            xlWsheetDSM.Range("S8:S9").Merge()
                            xlWsheetDSM.Range("S8:S9").HorizontalAlignment = 3
                            xlWsheetDSM.Range("S8:S9").WrapText = True
                            xlWsheetDSM.Range("T8:T9").Merge()
                            xlWsheetDSM.Range("T8:T9").HorizontalAlignment = 3
                            xlWsheetDSM.Range("T8:T9").WrapText = True
                            xlWsheetDSM.Range("U8:U9").Merge()
                            xlWsheetDSM.Range("U8:U9").HorizontalAlignment = 3
                            xlWsheetDSM.Range("U8:U9").WrapText = True
                            xlWsheetDSM.Range("V8:V9").Merge()
                            xlWsheetDSM.Range("V8:V9").HorizontalAlignment = 3
                            xlWsheetDSM.Range("V8:V9").WrapText = True
                            xlWsheetDSM.Range("W8:W9").Merge()
                            xlWsheetDSM.Range("W8:W9").HorizontalAlignment = 3
                            xlWsheetDSM.Range("W8:W9").WrapText = True
                            xlWsheetDSM.Range("X8:X9").Merge()
                            xlWsheetDSM.Range("X8:X9").HorizontalAlignment = 3
                            xlWsheetDSM.Range("X8:X9").WrapText = True

                            xlWbookDSM.SaveAs(txtDSM_dest.Text)
                            xlWbookDSM.Close()
                            xlAppDSM.Quit()

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetDSM)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookDSM)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppDSM)

                            xlWsheetDSM = Nothing
                            xlWbookDSM = Nothing
                            xlAppDSM = Nothing

                            MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Catch ex As Exception
                            MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try
                End Select

            Case "XL_Installed"
                Dim xlAppDSM As Object
                Dim xlWbookDSM As Object
                Dim xlWsheetDSM As Object

                Select Case DSMreptype
                    Case "QTY"
                        Try
                            xlAppDSM = CreateObject("Excel.Application")
                            xlWbookDSM = xlAppDSM.Workbooks.Open(txtDSM_src.Text)
                            xlWsheetDSM = xlWbookDSM.Worksheets("UID Daily Stock Mutation Report")

                            xlWsheetDSM.UsedRange.UnMerge()
                            xlWsheetDSM.UsedRange.WrapText = False
                            xlWsheetDSM.UsedRange.ColumnWidth = 15
                            xlWsheetDSM.UsedRange.RowHeight = 15

                            Dim rg_head_cut1 As Excel.Range = xlWsheetDSM.Range("B2")
                            Dim rg_head_paste1 As Excel.Range = xlWsheetDSM.Range("A2")
                            rg_head_cut1.Select()
                            rg_head_cut1.Cut(rg_head_paste1)

                            Dim rg_head_cut2 As Excel.Range = xlWsheetDSM.Range("B4")
                            Dim rg_head_paste2 As Excel.Range = xlWsheetDSM.Range("A3")
                            rg_head_cut2.Select()
                            rg_head_cut2.Cut(rg_head_paste2)

                            xlWsheetDSM.Range("A2").RowHeight = 27

                            Dim paramhead1, paramhead2, paramhead3 As String
                            paramhead1 = xlWsheetDSM.Range("C6").Value & " " & xlWsheetDSM.Range("F6").Value & "; "
                            paramhead1 = paramhead1 & xlWsheetDSM.Range("J6").Value & " " & xlWsheetDSM.Range("M6").Value

                            paramhead2 = xlWsheetDSM.Range("R6").Value & " " & xlWsheetDSM.Range("U6").Value & "; "
                            paramhead2 = paramhead2 & xlWsheetDSM.Range("Y6").Value & " " & xlWsheetDSM.Range("AC6").Value

                            paramhead3 = xlWsheetDSM.Range("AG6").Value & " " & xlWsheetDSM.Range("AK6").Value & "; "
                            paramhead3 = paramhead3 & xlWsheetDSM.Range("AP6").Value & " " & xlWsheetDSM.Range("AS6").Value

                            xlWsheetDSM.Range("A5").Value = paramhead1
                            xlWsheetDSM.Range("A5").EntireRow.Font.Name = "Calibri"
                            xlWsheetDSM.Range("A6").Value = paramhead2
                            xlWsheetDSM.Range("A6").EntireRow.Font.Name = "Calibri"
                            xlWsheetDSM.Range("A7").Value = paramhead3
                            xlWsheetDSM.Range("A7").EntireRow.Font.Name = "Calibri"

                            xlWsheetDSM.Range("G8:G9").Copy(xlWsheetDSM.Range("A8:A9"))
                            xlWsheetDSM.Range("A8").Value = "PCODE"
                            xlWsheetDSM.Range("A8").Font.Color = System.Drawing.ColorTranslator.ToOle(Color.White)
                            xlWsheetDSM.Range("A8:A9").Merge()
                            xlWsheetDSM.Range("A8:A9").HorizontalAlignment = Excel.Constants.xlCenter

                            Dim rg1, rg2, rg3, rg4, rg5, rg6, rg7, rg8, rg9, rg10, rg11 As Excel.Range
                            rg1 = xlWsheetDSM.Range("B:C")
                            rg1.Select()
                            rg1.Delete()

                            xlWsheetDSM.Range("E8:E9").Copy(xlWsheetDSM.Range("B8:B9"))
                            xlWsheetDSM.Range("B8").Value = "Nama Barang"
                            xlWsheetDSM.Range("B8").Font.Color = System.Drawing.ColorTranslator.ToOle(Color.White)
                            xlWsheetDSM.Range("B8:B9").Merge()
                            xlWsheetDSM.Range("B8:B9").HorizontalAlignment = Excel.Constants.xlCenter

                            rg2 = xlWsheetDSM.Range("C:D")
                            rg2.Select()
                            rg2.Delete()

                            xlWsheetDSM.Range("C8:C9").Merge()
                            xlWsheetDSM.Range("C8:C9").HorizontalAlignment = Excel.Constants.xlCenter
                            xlWsheetDSM.Range("C8:C9").WrapText = True
                            xlWsheetDSM.Range("D8:D9").Merge()
                            xlWsheetDSM.Range("D8:D9").HorizontalAlignment = Excel.Constants.xlCenter
                            xlWsheetDSM.Range("D8:D9").WrapText = True

                            rg3 = xlWsheetDSM.Range("E:F")
                            rg3.Select()
                            rg3.Delete()

                            xlWsheetDSM.Range("E8:E9").Merge()
                            xlWsheetDSM.Range("E8:E9").HorizontalAlignment = Excel.Constants.xlCenter
                            xlWsheetDSM.Range("E8:E9").WrapText = True

                            rg4 = xlWsheetDSM.Range("F:G")
                            rg4.Select()
                            rg4.Delete()

                            xlWsheetDSM.Range("F8:F9").Merge()
                            xlWsheetDSM.Range("F8:F9").HorizontalAlignment = Excel.Constants.xlCenter
                            xlWsheetDSM.Range("F8:F9").WrapText = True

                            xlWsheetDSM.Range("G:G").EntireColumn.Delete()

                            rg5 = xlWsheetDSM.Range("H:I")
                            rg5.Select()
                            rg5.Delete()

                            xlWsheetDSM.Range("G8:H8").Merge()
                            xlWsheetDSM.Range("G8:H8").HorizontalAlignment = Excel.Constants.xlCenter

                            rg6 = xlWsheetDSM.Range("I:J")
                            rg6.Select()
                            rg6.Delete()

                            xlWsheetDSM.Range("I8:I9").Merge()
                            xlWsheetDSM.Range("I8:I9").HorizontalAlignment = Excel.Constants.xlCenter
                            xlWsheetDSM.Range("I8:I9").WrapText = True

                            xlWsheetDSM.Range("J8:J9").Merge()
                            xlWsheetDSM.Range("J8:J9").HorizontalAlignment = Excel.Constants.xlCenter
                            xlWsheetDSM.Range("J8:J9").WrapText = True

                            rg7 = xlWsheetDSM.Range("K:L")
                            rg7.Select()
                            rg7.Delete()

                            xlWsheetDSM.Range("K8:K9").Merge()
                            xlWsheetDSM.Range("K8:K9").HorizontalAlignment = Excel.Constants.xlCenter
                            xlWsheetDSM.Range("K8:K9").WrapText = True

                            rg8 = xlWsheetDSM.Range("L:N")
                            rg8.Select()
                            rg8.Delete()

                            xlWsheetDSM.Range("L8:L9").Merge()
                            xlWsheetDSM.Range("L8:L9").HorizontalAlignment = Excel.Constants.xlCenter
                            xlWsheetDSM.Range("L8:L9").WrapText = True

                            xlWsheetDSM.Range("M8:M9").Merge()
                            xlWsheetDSM.Range("M8:M9").HorizontalAlignment = Excel.Constants.xlCenter
                            xlWsheetDSM.Range("M8:M9").WrapText = True

                            rg9 = xlWsheetDSM.Range("N:O")
                            rg9.Select()
                            rg9.Delete()

                            xlWsheetDSM.Range("N8:N9").Merge()
                            xlWsheetDSM.Range("N8:N9").HorizontalAlignment = Excel.Constants.xlCenter
                            xlWsheetDSM.Range("N8:N9").WrapText = True

                            xlWsheetDSM.Range("O:O").EntireColumn.Delete()

                            xlWsheetDSM.Range("O8:O9").Merge()
                            xlWsheetDSM.Range("O8:O9").HorizontalAlignment = Excel.Constants.xlCenter
                            xlWsheetDSM.Range("O8:O9").WrapText = True

                            xlWsheetDSM.Range("P:P").EntireColumn.Delete()

                            xlWsheetDSM.Range("P8:P9").Merge()
                            xlWsheetDSM.Range("P8:P9").HorizontalAlignment = Excel.Constants.xlCenter
                            xlWsheetDSM.Range("P8:P9").WrapText = True

                            xlWsheetDSM.Range("Q8:Q9").Merge()
                            xlWsheetDSM.Range("Q8:Q9").HorizontalAlignment = Excel.Constants.xlCenter
                            xlWsheetDSM.Range("Q8:Q9").WrapText = True

                            xlWsheetDSM.Range("R8:R9").Merge()
                            xlWsheetDSM.Range("R8:R9").HorizontalAlignment = Excel.Constants.xlCenter
                            xlWsheetDSM.Range("R8:R9").WrapText = True

                            rg10 = xlWsheetDSM.Range("S:T")
                            rg10.Select()
                            rg10.Delete()

                            xlWsheetDSM.Range("S8:S9").Merge()
                            xlWsheetDSM.Range("S8:S9").HorizontalAlignment = Excel.Constants.xlCenter
                            xlWsheetDSM.Range("S8:S9").WrapText = True

                            rg11 = xlWsheetDSM.Range("T:U")
                            rg11.Select()
                            rg11.Delete()

                            xlWsheetDSM.Range("T8:T9").Merge()
                            xlWsheetDSM.Range("T8:T9").HorizontalAlignment = Excel.Constants.xlCenter
                            xlWsheetDSM.Range("T8:T9").WrapText = True

                            xlWsheetDSM.Range("U8:U9").Merge()
                            xlWsheetDSM.Range("U8:U9").HorizontalAlignment = Excel.Constants.xlCenter
                            xlWsheetDSM.Range("U8:U9").WrapText = True

                            xlWsheetDSM.Range("V8:V9").Merge()
                            xlWsheetDSM.Range("V8:V9").HorizontalAlignment = Excel.Constants.xlCenter
                            xlWsheetDSM.Range("V8:V9").WrapText = True

                            xlWsheetDSM.Range("W8:W9").Merge()
                            xlWsheetDSM.Range("W8:W9").HorizontalAlignment = Excel.Constants.xlCenter
                            xlWsheetDSM.Range("W8:W9").WrapText = True

                            xlWsheetDSM.Range("X:X").EntireColumn.Delete()

                            xlWsheetDSM.Range("X8:X9").Merge()
                            xlWsheetDSM.Range("X8:X9").HorizontalAlignment = Excel.Constants.xlCenter
                            xlWsheetDSM.Range("X8:X9").WrapText = True

                            Dim lrow1 As Long
                            lrow1 = xlWsheetDSM.Range("A65536").End(Excel.XlDirection.xlUp).Row

                            Dim cr As String = "A" & lrow1 & ":B" & lrow1

                            xlWsheetDSM.Range(cr).Merge()
                            xlWsheetDSM.Range(cr).HorizontalAlignment = Excel.Constants.xlRight

                            xlWbookDSM.SaveAs(txtDSM_dest.Text)
                            xlWbookDSM.Close()
                            xlAppDSM.Quit()

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetDSM)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookDSM)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppDSM)

                            xlWsheetDSM = Nothing
                            xlWbookDSM = Nothing
                            xlAppDSM = Nothing

                            MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Catch ex As Exception
                            MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try

                    Case "RPH"
                        Try
                            xlAppDSM = CreateObject("Excel.Application")
                            xlWbookDSM = xlAppDSM.Workbooks.Open(txtDSM_src.Text)
                            xlWsheetDSM = xlWbookDSM.Worksheets("UID Daily Stock Mutation Report")

                            xlWsheetDSM.UsedRange.UnMerge()
                            xlWsheetDSM.UsedRange.WrapText = False
                            xlWsheetDSM.UsedRange.ColumnWidth = 15
                            xlWsheetDSM.UsedRange.RowHeight = 15

                            Dim rg_head_cut1 As Excel.Range = xlWsheetDSM.Range("B2")
                            Dim rg_head_paste1 As Excel.Range = xlWsheetDSM.Range("A2")
                            rg_head_cut1.Select()
                            rg_head_cut1.Cut(rg_head_paste1)

                            Dim rg_head_cut2 As Excel.Range = xlWsheetDSM.Range("B4")
                            Dim rg_head_paste2 As Excel.Range = xlWsheetDSM.Range("A3")
                            rg_head_cut2.Select()
                            rg_head_cut2.Cut(rg_head_paste2)

                            xlWsheetDSM.Range("A2").RowHeight = 27

                            Dim paramhead1, paramhead2, paramhead3 As String
                            paramhead1 = xlWsheetDSM.Range("C6").Value & " " & xlWsheetDSM.Range("F6").Value & "; "
                            paramhead1 = paramhead1 & xlWsheetDSM.Range("J6").Value & " " & xlWsheetDSM.Range("M6").Value

                            paramhead2 = xlWsheetDSM.Range("R6").Value & " " & xlWsheetDSM.Range("U6").Value & "; "
                            paramhead2 = paramhead2 & xlWsheetDSM.Range("Y6").Value & " " & xlWsheetDSM.Range("AC6").Value

                            paramhead3 = xlWsheetDSM.Range("AG6").Value & " " & xlWsheetDSM.Range("AK6").Value & "; "
                            paramhead3 = paramhead3 & xlWsheetDSM.Range("AP6").Value & " " & xlWsheetDSM.Range("AS6").Value

                            xlWsheetDSM.Range("A5").Value = paramhead1
                            xlWsheetDSM.Range("A5").EntireRow.Font.Name = "Calibri"
                            xlWsheetDSM.Range("A6").Value = paramhead2
                            xlWsheetDSM.Range("A6").EntireRow.Font.Name = "Calibri"
                            xlWsheetDSM.Range("A7").Value = paramhead3
                            xlWsheetDSM.Range("A7").EntireRow.Font.Name = "Calibri"

                            xlWsheetDSM.Range("G8:G9").Copy(xlWsheetDSM.Range("A8:A9"))
                            xlWsheetDSM.Range("A8").Value = "PCODE"
                            xlWsheetDSM.Range("A8").Font.Color = System.Drawing.ColorTranslator.ToOle(Color.White)
                            xlWsheetDSM.Range("A8:A9").Merge()
                            xlWsheetDSM.Range("A8:A9").HorizontalAlignment = Excel.Constants.xlCenter

                            Dim rg1, rg2, rg3, rg4, rg5, rg6, rg7, rg8, rg9, rg10, rg11 As Excel.Range
                            rg1 = xlWsheetDSM.Range("B:C")
                            rg1.Select()
                            rg1.Delete()

                            xlWsheetDSM.Range("E8:E9").Copy(xlWsheetDSM.Range("B8:B9"))
                            xlWsheetDSM.Range("B8").Value = "Nama Barang"
                            xlWsheetDSM.Range("B8").Font.Color = System.Drawing.ColorTranslator.ToOle(Color.White)
                            xlWsheetDSM.Range("B8:B9").Merge()
                            xlWsheetDSM.Range("B8:B9").HorizontalAlignment = Excel.Constants.xlCenter

                            rg2 = xlWsheetDSM.Range("C:D")
                            rg2.Select()
                            rg2.Delete()
                            rg3 = xlWsheetDSM.Range("E:F")
                            rg3.Select()
                            rg3.Delete()
                            rg4 = xlWsheetDSM.Range("F:G")
                            rg4.Select()
                            rg4.Delete()
                            xlWsheetDSM.Range("G:G").EntireColumn.Delete()
                            rg5 = xlWsheetDSM.Range("H:I")
                            rg5.Select()
                            rg5.Delete()
                            rg6 = xlWsheetDSM.Range("I:J")
                            rg6.Select()
                            rg6.Delete()
                            rg7 = xlWsheetDSM.Range("K:M")
                            rg7.Select()
                            rg7.Delete()
                            rg8 = xlWsheetDSM.Range("L:N")
                            rg8.Select()
                            rg8.Delete()
                            rg9 = xlWsheetDSM.Range("N:O")
                            rg9.Select()
                            rg9.Delete()
                            xlWsheetDSM.Range("O:O").EntireColumn.Delete()
                            xlWsheetDSM.Range("P:P").EntireColumn.Delete()
                            rg10 = xlWsheetDSM.Range("S:T")
                            rg10.Select()
                            rg10.Delete()
                            rg11 = xlWsheetDSM.Range("T:U")
                            rg11.Select()
                            rg11.Delete()
                            xlWsheetDSM.Range("X:X").EntireColumn.Delete()

                            Dim lrow As Long
                            lrow = xlWsheetDSM.Range("A65536").End(Excel.XlDirection.xlUp).Row
                            Dim cr As String
                            cr = "A" & lrow & ":B" & lrow
                            xlWsheetDSM.Range(cr).Merge()
                            xlWsheetDSM.Range(cr).HorizontalAlignment = Excel.Constants.xlRight
                            Dim cr2 As String
                            cr2 = "A" & lrow
                            xlWsheetDSM.Range(cr2).EntireRow.Font.Bold = True

                            xlWsheetDSM.Range("C8:C9").Merge()
                            xlWsheetDSM.Range("C8:C9").HorizontalAlignment = Excel.Constants.xlCenter
                            xlWsheetDSM.Range("C8:C9").WrapText = True
                            xlWsheetDSM.Range("D8:D9").Merge()
                            xlWsheetDSM.Range("D8:D9").HorizontalAlignment = Excel.Constants.xlCenter
                            xlWsheetDSM.Range("D8:D9").WrapText = True
                            xlWsheetDSM.Range("E8:E9").Merge()
                            xlWsheetDSM.Range("E8:E9").HorizontalAlignment = Excel.Constants.xlCenter
                            xlWsheetDSM.Range("E8:E9").WrapText = True
                            xlWsheetDSM.Range("F8:F9").Merge()
                            xlWsheetDSM.Range("F8:F9").HorizontalAlignment = Excel.Constants.xlCenter
                            xlWsheetDSM.Range("F8:F9").WrapText = True
                            xlWsheetDSM.Range("G8:H8").Merge()
                            xlWsheetDSM.Range("G8:H8").HorizontalAlignment = Excel.Constants.xlCenter
                            xlWsheetDSM.Range("G8:H8").WrapText = True
                            xlWsheetDSM.Range("G9:H9").HorizontalAlignment = Excel.Constants.xlCenter
                            xlWsheetDSM.Range("I8:I9").Merge()
                            xlWsheetDSM.Range("I8:I9").HorizontalAlignment = Excel.Constants.xlCenter
                            xlWsheetDSM.Range("I8:I9").WrapText = True
                            xlWsheetDSM.Range("J8:J9").Merge()
                            xlWsheetDSM.Range("J8:J9").HorizontalAlignment = Excel.Constants.xlCenter
                            xlWsheetDSM.Range("J8:J9").WrapText = True
                            xlWsheetDSM.Range("K8:K9").Merge()
                            xlWsheetDSM.Range("K8:K9").HorizontalAlignment = Excel.Constants.xlCenter
                            xlWsheetDSM.Range("K8:K9").WrapText = True
                            xlWsheetDSM.Range("L8:L9").Merge()
                            xlWsheetDSM.Range("L8:L9").HorizontalAlignment = Excel.Constants.xlCenter
                            xlWsheetDSM.Range("L8:L9").WrapText = True
                            xlWsheetDSM.Range("M8:M9").Merge()
                            xlWsheetDSM.Range("M8:M9").HorizontalAlignment = Excel.Constants.xlCenter
                            xlWsheetDSM.Range("M8:M9").WrapText = True
                            xlWsheetDSM.Range("N8:N9").Merge()
                            xlWsheetDSM.Range("N8:N9").HorizontalAlignment = Excel.Constants.xlCenter
                            xlWsheetDSM.Range("N8:N9").WrapText = True
                            xlWsheetDSM.Range("N8:N9").Merge()
                            xlWsheetDSM.Range("N8:N9").HorizontalAlignment = Excel.Constants.xlCenter
                            xlWsheetDSM.Range("N8:N9").WrapText = True
                            xlWsheetDSM.Range("O8:O9").Merge()
                            xlWsheetDSM.Range("O8:O9").HorizontalAlignment = Excel.Constants.xlCenter
                            xlWsheetDSM.Range("O8:O9").WrapText = True
                            xlWsheetDSM.Range("P8:P9").Merge()
                            xlWsheetDSM.Range("P8:P9").HorizontalAlignment = Excel.Constants.xlCenter
                            xlWsheetDSM.Range("P8:P9").WrapText = True
                            xlWsheetDSM.Range("Q8:Q9").Merge()
                            xlWsheetDSM.Range("Q8:Q9").HorizontalAlignment = Excel.Constants.xlCenter
                            xlWsheetDSM.Range("Q8:Q9").WrapText = True
                            xlWsheetDSM.Range("R8:R9").Merge()
                            xlWsheetDSM.Range("R8:R9").HorizontalAlignment = Excel.Constants.xlCenter
                            xlWsheetDSM.Range("R8:R9").WrapText = True
                            xlWsheetDSM.Range("S8:S9").Merge()
                            xlWsheetDSM.Range("S8:S9").HorizontalAlignment = Excel.Constants.xlCenter
                            xlWsheetDSM.Range("S8:S9").WrapText = True
                            xlWsheetDSM.Range("T8:T9").Merge()
                            xlWsheetDSM.Range("T8:T9").HorizontalAlignment = Excel.Constants.xlCenter
                            xlWsheetDSM.Range("T8:T9").WrapText = True
                            xlWsheetDSM.Range("U8:U9").Merge()
                            xlWsheetDSM.Range("U8:U9").HorizontalAlignment = Excel.Constants.xlCenter
                            xlWsheetDSM.Range("U8:U9").WrapText = True
                            xlWsheetDSM.Range("V8:V9").Merge()
                            xlWsheetDSM.Range("V8:V9").HorizontalAlignment = Excel.Constants.xlCenter
                            xlWsheetDSM.Range("V8:V9").WrapText = True
                            xlWsheetDSM.Range("W8:W9").Merge()
                            xlWsheetDSM.Range("W8:W9").HorizontalAlignment = Excel.Constants.xlCenter
                            xlWsheetDSM.Range("W8:W9").WrapText = True
                            xlWsheetDSM.Range("X8:X9").Merge()
                            xlWsheetDSM.Range("X8:X9").HorizontalAlignment = Excel.Constants.xlCenter
                            xlWsheetDSM.Range("X8:X9").WrapText = True

                            xlWbookDSM.SaveAs(txtDSM_dest.Text)
                            xlWbookDSM.Close()
                            xlAppDSM.Quit()

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetDSM)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookDSM)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppDSM)

                            xlWsheetDSM = Nothing
                            xlWbookDSM = Nothing
                            xlAppDSM = Nothing

                            MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Catch ex As Exception
                            MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try
                End Select
        End Select

    End Sub

    Private Sub rbDSM_qty_CheckedChanged(sender As Object, e As EventArgs) Handles rbDSM_qty.CheckedChanged
        DSMreptype = "QTY"
    End Sub

    Private Sub rbDSM_rph_CheckedChanged(sender As Object, e As EventArgs) Handles rbDSM_rph.CheckedChanged
        DSMreptype = "RPH"
    End Sub
End Class