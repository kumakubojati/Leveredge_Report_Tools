Imports Microsoft.Office.Interop
Imports System.Windows.Forms
Imports System.Text
Imports System.Drawing
Imports Microsoft.Win32
Public Class frmSISR
    Dim AppsOffice As String
    Private Sub frmSISR_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim TargetKey As RegistryKey
        TargetKey = Registry.ClassesRoot.OpenSubKey("Excel.Application")
        If TargetKey Is Nothing Then
            AppsOffice = "XL_NotInstalled"
        Else
            AppsOffice = "XL_Installed"
            TargetKey.Close()
        End If
    End Sub
    Dim SISRRepType As String

    Private Sub rbWithTax_CheckedChanged(sender As Object, e As EventArgs) Handles rbWithTax.CheckedChanged
        'WT = With Tax
        SISRRepType = "WT"
    End Sub

    Private Sub rbNoTax_CheckedChanged(sender As Object, e As EventArgs) Handles rbNoTax.CheckedChanged
        'NT = No Tax
        SISRRepType = "NT"
    End Sub

    Private Sub btnBrow_SISR_src_Click(sender As Object, e As EventArgs) Handles btnBrow_SISR_src.Click
        Dim DSMpathSrc As String
        If OFD_SISR.ShowDialog = DialogResult.OK Then
            DSMpathSrc = OFD_SISR.FileName
            txtSISR_src.Text = DSMpathSrc
        End If
        If txtSISR_dest.Text <> "" Then
            btnNeu_SISR.Enabled = True
        Else
            btnNeu_SISR.Enabled = False
        End If
    End Sub

    Private Sub btnBrow_SISR_dest_Click(sender As Object, e As EventArgs) Handles btnBrow_SISR_dest.Click
        Dim datenow As DateTime = DateTime.Now
        Dim filename As String = "Neutralize_SISR_" & SISRRepType & " " & datenow.ToString("ddMMyyyy_HHmm")
        SFD_SISR.FileName = filename
        Dim DSMpath_Dest As String
        If SFD_SISR.ShowDialog = DialogResult.OK Then
            DSMpath_Dest = SFD_SISR.FileName
            txtSISR_dest.Text = DSMpath_Dest
        End If
        If txtSISR_src.Text <> "" Then
            btnNeu_SISR.Enabled = True
        Else
            btnNeu_SISR.Enabled = False
        End If
        If txtSISR_dest.Text <> "" Then
            btnNeu_SISR.Enabled = True
        Else
            btnNeu_SISR.Enabled = False
        End If
    End Sub

    Private Sub btnNeu_SISR_Click(sender As Object, e As EventArgs) Handles btnNeu_SISR.Click
        PicBar_SISR.Visible = True
        BWSISR.RunWorkerAsync()
    End Sub

    Private Sub BWSISR_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BWSISR.DoWork
        Select Case AppsOffice
            Case "XL_NotInstalled"
                Dim xlAppSISR As Object
                Dim xlWbookSISR As Object
                Dim xlWsheetSISR As Object

                Select Case SISRRepType
                    Case "WT"
                        Try
                            xlAppSISR = CreateObject("Ket.Application")
                            xlWbookSISR = xlAppSISR.Workbooks.Open(txtSISR_src.Text)
                            xlWsheetSISR = xlWbookSISR.Worksheets("UID Summary Invoice And Sales R")

                            xlWsheetSISR.UsedRange.UnMerge()
                            xlWsheetSISR.UsedRange.WrapText = False
                            xlWsheetSISR.UsedRange.ColumnWidth = 15
                            xlWsheetSISR.UsedRange.RowHeight = 15

                            Dim rg_head_cut1 As Object = xlWsheetSISR.Range("B2")
                            Dim rg_head_paste1 As Object = xlWsheetSISR.Range("A2")
                            rg_head_cut1.Select()
                            rg_head_cut1.Cut(rg_head_paste1)

                            Dim rg_head_cut2 As Object = xlWsheetSISR.Range("B4")
                            Dim rg_head_paste2 As Object = xlWsheetSISR.Range("A3")
                            rg_head_cut2.Select()
                            rg_head_cut2.Cut(rg_head_paste2)

                            xlWsheetSISR.Range("A2").RowHeight = 27

                            Dim paramhead1, paramhead2, paramhead3 As String
                            paramhead1 = xlWsheetSISR.Range("C6").Value & " " & xlWsheetSISR.Range("H6").Value & "; "
                            paramhead1 = paramhead1 & xlWsheetSISR.Range("L6").Value & " " & xlWsheetSISR.Range("P6").Value

                            paramhead2 = xlWsheetSISR.Range("T6").Value & " " & xlWsheetSISR.Range("Y6").Value & "; "

                            paramhead3 = xlWsheetSISR.Range("C8").Value & " " & xlWsheetSISR.Range("G8").Value & "; "
                            paramhead3 = paramhead3 & xlWsheetSISR.Range("L8").Value & " " & xlWsheetSISR.Range("P8").Value

                            xlWsheetSISR.Range("A5").Value = paramhead1
                            xlWsheetSISR.Range("A5").EntireRow.Font.Name = "Calibri"
                            xlWsheetSISR.Range("A6").Value = paramhead2
                            xlWsheetSISR.Range("A6").EntireRow.Font.Name = "Calibri"
                            xlWsheetSISR.Range("A7").Value = paramhead3
                            xlWsheetSISR.Range("A7").EntireRow.Font.Name = "Calibri"

                            xlWsheetSISR.Range("C6:Z6").Value = ""
                            xlWsheetSISR.Range("C8:Z8").Value = ""

                            xlWsheetSISR.Range("A9").EntireRow.Delete()

                            Dim xlfunc As Object
                            xlfunc = xlAppSISR.WorksheetFunction
                            Dim lnCol As Long
                            Dim i, j As Long
                            Dim rnarea As Object = xlWsheetSISR.UsedRange

                            lnCol = rnarea.Columns.Count
                            For i = lnCol To 1 Step -1
                                If xlfunc.CountA(rnarea.Columns(i)) = 0 Then
                                    rnarea.Columns(i).Delete()
                                    j = j + 1
                                End If
                            Next

                            xlWbookSISR.SaveAs(txtSISR_dest.Text)
                            xlWbookSISR.Close()
                            xlAppSISR.Quit()

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetSISR)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookSISR)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppSISR)

                            xlWsheetSISR = Nothing
                            xlWbookSISR = Nothing
                            xlAppSISR = Nothing

                            MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                        Catch ex As Exception
                            MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try

                    Case "NT"
                        Try
                            xlAppSISR = CreateObject("Ket.Application")
                            xlWbookSISR = xlAppSISR.Workbooks.Open(txtSISR_src.Text)
                            xlWsheetSISR = xlWbookSISR.Worksheets("UID Summary Invoice And Sales R")

                            xlWsheetSISR.UsedRange.UnMerge()
                            xlWsheetSISR.UsedRange.WrapText = False
                            xlWsheetSISR.UsedRange.ColumnWidth = 15
                            xlWsheetSISR.UsedRange.RowHeight = 15

                            Dim rg_head_cut1 As Object = xlWsheetSISR.Range("B2")
                            Dim rg_head_paste1 As Object = xlWsheetSISR.Range("A2")
                            rg_head_cut1.Select()
                            rg_head_cut1.Cut(rg_head_paste1)

                            Dim rg_head_cut2 As Object = xlWsheetSISR.Range("B4")
                            Dim rg_head_paste2 As Object = xlWsheetSISR.Range("A3")
                            rg_head_cut2.Select()
                            rg_head_cut2.Cut(rg_head_paste2)

                            xlWsheetSISR.Range("A2").RowHeight = 27

                            Dim paramhead1, paramhead2, paramhead3 As String
                            paramhead1 = xlWsheetSISR.Range("C6").Value & " " & xlWsheetSISR.Range("H6").Value & "; "
                            paramhead1 = paramhead1 & xlWsheetSISR.Range("L6").Value & " " & xlWsheetSISR.Range("P6").Value

                            paramhead2 = xlWsheetSISR.Range("T6").Value & " " & xlWsheetSISR.Range("Y6").Value & "; "

                            paramhead3 = xlWsheetSISR.Range("C8").Value & " " & xlWsheetSISR.Range("G8").Value & "; "
                            paramhead3 = paramhead3 & xlWsheetSISR.Range("L8").Value & " " & xlWsheetSISR.Range("P8").Value

                            xlWsheetSISR.Range("A5").Value = paramhead1
                            xlWsheetSISR.Range("A5").EntireRow.Font.Name = "Calibri"
                            xlWsheetSISR.Range("A6").Value = paramhead2
                            xlWsheetSISR.Range("A6").EntireRow.Font.Name = "Calibri"
                            xlWsheetSISR.Range("A7").Value = paramhead3
                            xlWsheetSISR.Range("A7").EntireRow.Font.Name = "Calibri"

                            xlWsheetSISR.Range("C6:Z6").Value = ""
                            xlWsheetSISR.Range("C8:Z8").Value = ""

                            xlWsheetSISR.Range("A9").EntireRow.Delete()

                            Dim xlfunc As Object
                            xlfunc = xlAppSISR.WorksheetFunction
                            Dim lnCol As Long
                            Dim i, j As Long
                            Dim rnarea As Object = xlWsheetSISR.UsedRange

                            lnCol = rnarea.Columns.Count
                            For i = lnCol To 1 Step -1
                                If xlfunc.CountA(rnarea.Columns(i)) = 0 Then
                                    rnarea.Columns(i).Delete()
                                    j = j + 1
                                End If
                            Next

                            xlWbookSISR.SaveAs(txtSISR_dest.Text)
                            xlWbookSISR.Close()
                            xlAppSISR.Quit()

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetSISR)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookSISR)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppSISR)

                            xlWsheetSISR = Nothing
                            xlWbookSISR = Nothing
                            xlAppSISR = Nothing

                            MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                        Catch ex As Exception
                            MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try

                End Select

            Case "XL_Installed"
                Dim xlAppSISR As Object
                Dim xlWbookSISR As Object
                Dim xlWsheetSISR As Object

                Select Case SISRRepType
                    Case "WT"
                        Try
                            xlAppSISR = CreateObject("Excel.Application")
                            xlWbookSISR = xlAppSISR.Workbooks.Open(txtSISR_src.Text)
                            xlWsheetSISR = xlWbookSISR.Worksheets("UID Summary Invoice And Sales R")

                            xlWsheetSISR.UsedRange.UnMerge()
                            xlWsheetSISR.UsedRange.WrapText = False
                            xlWsheetSISR.UsedRange.ColumnWidth = 15
                            xlWsheetSISR.UsedRange.RowHeight = 15

                            Dim rg_head_cut1 As Excel.Range = xlWsheetSISR.Range("B2")
                            Dim rg_head_paste1 As Excel.Range = xlWsheetSISR.Range("A2")
                            rg_head_cut1.Select()
                            rg_head_cut1.Cut(rg_head_paste1)

                            Dim rg_head_cut2 As Excel.Range = xlWsheetSISR.Range("B4")
                            Dim rg_head_paste2 As Excel.Range = xlWsheetSISR.Range("A3")
                            rg_head_cut2.Select()
                            rg_head_cut2.Cut(rg_head_paste2)

                            xlWsheetSISR.Range("A2").RowHeight = 27

                            Dim paramhead1, paramhead2, paramhead3 As String
                            paramhead1 = xlWsheetSISR.Range("C6").Value & " " & xlWsheetSISR.Range("H6").Value & "; "
                            paramhead1 = paramhead1 & xlWsheetSISR.Range("L6").Value & " " & xlWsheetSISR.Range("P6").Value

                            paramhead2 = xlWsheetSISR.Range("T6").Value & " " & xlWsheetSISR.Range("Y6").Value & "; "

                            paramhead3 = xlWsheetSISR.Range("C8").Value & " " & xlWsheetSISR.Range("G8").Value & "; "
                            paramhead3 = paramhead3 & xlWsheetSISR.Range("L8").Value & " " & xlWsheetSISR.Range("P8").Value

                            xlWsheetSISR.Range("A5").Value = paramhead1
                            xlWsheetSISR.Range("A5").EntireRow.Font.Name = "Calibri"
                            xlWsheetSISR.Range("A6").Value = paramhead2
                            xlWsheetSISR.Range("A6").EntireRow.Font.Name = "Calibri"
                            xlWsheetSISR.Range("A7").Value = paramhead3
                            xlWsheetSISR.Range("A7").EntireRow.Font.Name = "Calibri"

                            xlWsheetSISR.Range("C6:Z6").Value = ""
                            xlWsheetSISR.Range("C8:Z8").Value = ""

                            xlWsheetSISR.Range("A9").EntireRow.Delete()

                            Dim xlfunc As Excel.WorksheetFunction
                            xlfunc = xlAppSISR.WorksheetFunction
                            Dim lnCol As Long
                            Dim i, j As Long
                            Dim rnarea As Excel.Range = xlWsheetSISR.UsedRange

                            lnCol = rnarea.Columns.Count
                            For i = lnCol To 1 Step -1
                                If xlfunc.CountA(rnarea.Columns(i)) = 0 Then
                                    rnarea.Columns(i).Delete()
                                    j = j + 1
                                End If
                            Next

                            xlWbookSISR.SaveAs(txtSISR_dest.Text)
                            xlWbookSISR.Close()
                            xlAppSISR.Quit()

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetSISR)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookSISR)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppSISR)

                            xlWsheetSISR = Nothing
                            xlWbookSISR = Nothing
                            xlAppSISR = Nothing

                            MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                        Catch ex As Exception
                            MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try

                    Case "NT"
                        Try
                            xlAppSISR = CreateObject("Excel.Application")
                            xlWbookSISR = xlAppSISR.Workbooks.Open(txtSISR_src.Text)
                            xlWsheetSISR = xlWbookSISR.Worksheets("UID Summary Invoice And Sales R")

                            xlWsheetSISR.UsedRange.UnMerge()
                            xlWsheetSISR.UsedRange.WrapText = False
                            xlWsheetSISR.UsedRange.ColumnWidth = 15
                            xlWsheetSISR.UsedRange.RowHeight = 15

                            Dim rg_head_cut1 As Excel.Range = xlWsheetSISR.Range("B2")
                            Dim rg_head_paste1 As Excel.Range = xlWsheetSISR.Range("A2")
                            rg_head_cut1.Select()
                            rg_head_cut1.Cut(rg_head_paste1)

                            Dim rg_head_cut2 As Excel.Range = xlWsheetSISR.Range("B4")
                            Dim rg_head_paste2 As Excel.Range = xlWsheetSISR.Range("A3")
                            rg_head_cut2.Select()
                            rg_head_cut2.Cut(rg_head_paste2)

                            xlWsheetSISR.Range("A2").RowHeight = 27

                            Dim paramhead1, paramhead2, paramhead3 As String
                            paramhead1 = xlWsheetSISR.Range("C6").Value & " " & xlWsheetSISR.Range("H6").Value & "; "
                            paramhead1 = paramhead1 & xlWsheetSISR.Range("L6").Value & " " & xlWsheetSISR.Range("P6").Value

                            paramhead2 = xlWsheetSISR.Range("T6").Value & " " & xlWsheetSISR.Range("Y6").Value & "; "

                            paramhead3 = xlWsheetSISR.Range("C8").Value & " " & xlWsheetSISR.Range("G8").Value & "; "
                            paramhead3 = paramhead3 & xlWsheetSISR.Range("L8").Value & " " & xlWsheetSISR.Range("P8").Value

                            xlWsheetSISR.Range("A5").Value = paramhead1
                            xlWsheetSISR.Range("A5").EntireRow.Font.Name = "Calibri"
                            xlWsheetSISR.Range("A6").Value = paramhead2
                            xlWsheetSISR.Range("A6").EntireRow.Font.Name = "Calibri"
                            xlWsheetSISR.Range("A7").Value = paramhead3
                            xlWsheetSISR.Range("A7").EntireRow.Font.Name = "Calibri"

                            xlWsheetSISR.Range("C6:Z6").Value = ""
                            xlWsheetSISR.Range("C8:Z8").Value = ""

                            xlWsheetSISR.Range("A9").EntireRow.Delete()

                            Dim xlfunc As Excel.WorksheetFunction
                            xlfunc = xlAppSISR.WorksheetFunction
                            Dim lnCol As Long
                            Dim i, j As Long
                            Dim rnarea As Excel.Range = xlWsheetSISR.UsedRange

                            lnCol = rnarea.Columns.Count
                            For i = lnCol To 1 Step -1
                                If xlfunc.CountA(rnarea.Columns(i)) = 0 Then
                                    rnarea.Columns(i).Delete()
                                    j = j + 1
                                End If
                            Next

                            xlWbookSISR.SaveAs(txtSISR_dest.Text)
                            xlWbookSISR.Close()
                            xlAppSISR.Quit()

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetSISR)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookSISR)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppSISR)

                            xlWsheetSISR = Nothing
                            xlWbookSISR = Nothing
                            xlAppSISR = Nothing

                            MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                        Catch ex As Exception
                            MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try

                End Select

        End Select
    End Sub

    Private Sub BWSISR_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BWSISR.RunWorkerCompleted
        PicBar_SISR.Visible = False
        btnNeu_SISR.Enabled = False
        txtSISR_dest.Text = ""
        txtSISR_src.Text = ""
    End Sub
End Class