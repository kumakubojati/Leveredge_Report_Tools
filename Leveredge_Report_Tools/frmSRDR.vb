Imports Microsoft.Office.Interop
Imports System.Windows.Forms
Imports System.Text
Imports System.Drawing
Imports Microsoft.Win32

Public Class frmSRDR
    Dim AppsOffice As String
    Private Sub frmSRDR_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim TargetKey As RegistryKey
        TargetKey = Registry.ClassesRoot.OpenSubKey("Excel.Application")
        If TargetKey Is Nothing Then
            AppsOffice = "XL_NotInstalled"
        Else
            AppsOffice = "XL_Installed"
            TargetKey.Close()
        End If
    End Sub

    Private Sub btnBrow_SRDR_src_Click(sender As Object, e As EventArgs) Handles btnBrow_SRDR_src.Click
        Dim SRDRpathSrc As String
        If OFD_SRDR.ShowDialog = DialogResult.OK Then
            SRDRpathSrc = OFD_SRDR.FileName()
            txtSRDR_src.Text = SRDRpathSrc
        End If
        If txtSRDR_dest.Text <> "" Then
            btnNeu_SRDR.Enabled = True
        Else
            btnNeu_SRDR.Enabled = False
        End If
    End Sub

    Private Sub btnBrow_SRDR_dest_Click(sender As Object, e As EventArgs) Handles btnBrow_SRDR_dest.Click
        Dim datenow As DateTime = DateTime.Now
        Dim filename As String = "Neutralize_SalesReturnDet_" & datenow.ToString("ddMMyyyy_HHmm")
        SFD_SRDR.FileName = filename
        Dim SRDRpath_Dest As String
        If SFD_SRDR.ShowDialog = DialogResult.OK Then
            SRDRpath_Dest = SFD_SRDR.FileName
            txtSRDR_dest.Text = SRDRpath_Dest
        End If
        If txtSRDR_src.Text <> "" Then
            btnNeu_SRDR.Enabled = True
        Else
            btnNeu_SRDR.Enabled = False
        End If
        If txtSRDR_dest.Text <> "" Then
            btnNeu_SRDR.Enabled = True
        Else
            btnNeu_SRDR.Enabled = False
        End If
    End Sub

    Private Sub btnNeu_SRDR_Click(sender As Object, e As EventArgs) Handles btnNeu_SRDR.Click
        PicBar_SRDR.Visible = True
        BWSRDR.RunWorkerAsync()
    End Sub

    Private Sub BWSRDR_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BWSRDR.DoWork
        Select Case AppsOffice
            Case "XL_NotInstalled"
                Dim xlAppSRDR As Object
                Dim xlWbookSRDR As Object
                Dim xlWsheetSRDR As Object

                Try
                    xlAppSRDR = CreateObject("Ket.Application")
                    xlWbookSRDR = xlAppSRDR.Workbooks.Open(txtSRDR_src.Text)
                    xlWsheetSRDR = xlWbookSRDR.Worksheets("UID Sales Return Detail Report")

                    xlWsheetSRDR.UsedRange.UnMerge()
                    xlWsheetSRDR.UsedRange.WrapText = False
                    xlWsheetSRDR.UsedRange.ColumnWidth = 15
                    xlWsheetSRDR.UsedRange.RowHeight = 15

                    Dim rg_head_cut1 As Object = xlWsheetSRDR.Range("C2")
                    Dim rg_head_paste1 As Object = xlWsheetSRDR.Range("B2")
                    rg_head_cut1.Select()
                    rg_head_cut1.Cut(rg_head_paste1)

                    Dim rg_head_cut2 As Object = xlWsheetSRDR.Range("C4")
                    Dim rg_head_paste2 As Object = xlWsheetSRDR.Range("B3")
                    rg_head_cut2.Select()
                    rg_head_cut2.Cut(rg_head_paste2)

                    xlWsheetSRDR.Range("A2").RowHeight = 27

                    Dim paramhead1, paramhead2, paramhead3 As String
                    paramhead1 = xlWsheetSRDR.Range("D6").Value & " " & xlWsheetSRDR.Range("G6").Value & "; "
                    paramhead1 = paramhead1 & xlWsheetSRDR.Range("K6").Value & " " & xlWsheetSRDR.Range("P6").Value

                    paramhead2 = xlWsheetSRDR.Range("P6").Value & " " & xlWsheetSRDR.Range("X6").Value & ";"
                    paramhead2 = paramhead2 & xlWsheetSRDR.Range("AB6").Value & " " & xlWsheetSRDR.Range("AF6").Value & "; "

                    paramhead3 = xlWsheetSRDR.Range("D8").Value & " " & xlWsheetSRDR.Range("G8").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetSRDR.Range("J8").Value & " " & xlWsheetSRDR.Range("O8").Value & ";"
                    paramhead3 = paramhead3 & xlWsheetSRDR.Range("U9").Value & " " & xlWsheetSRDR.Range("X9").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetSRDR.Range("AB9").Value & " " & xlWsheetSRDR.Range("AF9").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetSRDR.Range("AI9").Value & " " & xlWsheetSRDR.Range("AM9").Value

                    xlWsheetSRDR.Range("B5").Value = paramhead1
                    xlWsheetSRDR.Range("B5").EntireRow.Font.Name = "Calibri"
                    xlWsheetSRDR.Range("B6").Value = paramhead2
                    xlWsheetSRDR.Range("B6").EntireRow.Font.Name = "Calibri"
                    xlWsheetSRDR.Range("B7").Value = paramhead3
                    xlWsheetSRDR.Range("B7").EntireRow.Font.Name = "Calibri"

                    xlWsheetSRDR.Range("C6:AM9").value = ""
                    xlWsheetSRDR.Range("A10:A11").EntireRow.Delete()
                    xlWsheetSRDR.Range("A:A").EntireColumn.Delete()

                    Dim xlfunc As Object
                    xlfunc = xlAppSRDR.WorksheetFunction
                    Dim lnCol As Long
                    Dim i, j As Long
                    Dim rnarea As Object = xlWsheetSRDR.UsedRange

                    lnCol = rnarea.Columns.Count
                    For i = lnCol To 1 Step -1
                        If xlfunc.CountA(rnarea.Columns(i)) = 0 Then
                            rnarea.Columns(i).Delete()
                            j = j + 1
                        End If
                    Next

                    xlWbookSRDR.SaveAs(txtSRDR_dest.Text)
                    xlWbookSRDR.Close()
                    xlAppSRDR.Quit()

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetSRDR)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookSRDR)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppSRDR)

                    xlWsheetSRDR = Nothing
                    xlWbookSRDR = Nothing
                    xlAppSRDR = Nothing

                    MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                Catch ex As Exception
                    MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try

            Case "XL_Installed"
                Dim xlAppSRDR As Object
                Dim xlWbookSRDR As Object
                Dim xlWsheetSRDR As Object

                Try
                    xlAppSRDR = CreateObject("Excel.Application")
                    xlWbookSRDR = xlAppSRDR.Workbooks.Open(txtSRDR_src.Text)
                    xlWsheetSRDR = xlWbookSRDR.Worksheets("UID Sales Return Detail Report")

                    xlWsheetSRDR.UsedRange.UnMerge()
                    xlWsheetSRDR.UsedRange.WrapText = False
                    xlWsheetSRDR.UsedRange.ColumnWidth = 15
                    xlWsheetSRDR.UsedRange.RowHeight = 15

                    Dim rg_head_cut1 As Excel.Range = xlWsheetSRDR.Range("C2")
                    Dim rg_head_paste1 As Excel.Range = xlWsheetSRDR.Range("B2")
                    rg_head_cut1.Select()
                    rg_head_cut1.Cut(rg_head_paste1)

                    Dim rg_head_cut2 As Excel.Range = xlWsheetSRDR.Range("C4")
                    Dim rg_head_paste2 As Excel.Range = xlWsheetSRDR.Range("B3")
                    rg_head_cut2.Select()
                    rg_head_cut2.Cut(rg_head_paste2)

                    xlWsheetSRDR.Range("A2").RowHeight = 27

                    Dim paramhead1, paramhead2, paramhead3 As String
                    paramhead1 = xlWsheetSRDR.Range("D6").Value & " " & xlWsheetSRDR.Range("G6").Value & "; "
                    paramhead1 = paramhead1 & xlWsheetSRDR.Range("K6").Value & " " & xlWsheetSRDR.Range("P6").Value

                    paramhead2 = xlWsheetSRDR.Range("P6").Value & " " & xlWsheetSRDR.Range("X6").Value & ";"
                    paramhead2 = paramhead2 & xlWsheetSRDR.Range("AB6").Value & " " & xlWsheetSRDR.Range("AF6").Value & "; "

                    paramhead3 = xlWsheetSRDR.Range("D8").Value & " " & xlWsheetSRDR.Range("G8").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetSRDR.Range("J8").Value & " " & xlWsheetSRDR.Range("O8").Value & ";"
                    paramhead3 = paramhead3 & xlWsheetSRDR.Range("U9").Value & " " & xlWsheetSRDR.Range("X9").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetSRDR.Range("AB9").Value & " " & xlWsheetSRDR.Range("AF9").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetSRDR.Range("AI9").Value & " " & xlWsheetSRDR.Range("AM9").Value

                    xlWsheetSRDR.Range("B5").Value = paramhead1
                    xlWsheetSRDR.Range("B5").EntireRow.Font.Name = "Calibri"
                    xlWsheetSRDR.Range("B6").Value = paramhead2
                    xlWsheetSRDR.Range("B6").EntireRow.Font.Name = "Calibri"
                    xlWsheetSRDR.Range("B7").Value = paramhead3
                    xlWsheetSRDR.Range("B7").EntireRow.Font.Name = "Calibri"

                    xlWsheetSRDR.Range("C6:AM9").value = ""
                    xlWsheetSRDR.Range("A10:A11").EntireRow.Delete()
                    xlWsheetSRDR.Range("A:A").EntireColumn.Delete()

                    Dim xlfunc As Object
                    xlfunc = xlAppSRDR.WorksheetFunction
                    Dim lnCol As Long
                    Dim i, j As Long
                    Dim rnarea As Object = xlWsheetSRDR.UsedRange

                    lnCol = rnarea.Columns.Count
                    For i = lnCol To 1 Step -1
                        If xlfunc.CountA(rnarea.Columns(i)) = 0 Then
                            rnarea.Columns(i).Delete()
                            j = j + 1
                        End If
                    Next

                    xlWbookSRDR.SaveAs(txtSRDR_dest.Text)
                    xlWbookSRDR.Close()
                    xlAppSRDR.Quit()

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetSRDR)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookSRDR)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppSRDR)

                    xlWsheetSRDR = Nothing
                    xlWbookSRDR = Nothing
                    xlAppSRDR = Nothing

                    MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                Catch ex As Exception
                    MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try

        End Select
    End Sub

    Private Sub BWSRDR_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BWSRDR.RunWorkerCompleted
        PicBar_SRDR.Visible = False
        btnNeu_SRDR.Enabled = False
        txtSRDR_dest.Text = ""
        txtSRDR_src.Text = ""
    End Sub
End Class