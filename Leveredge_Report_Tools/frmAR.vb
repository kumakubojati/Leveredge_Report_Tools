Imports Microsoft.Office.Interop
Imports System.Windows.Forms
Imports System.Text
Imports System.Drawing
Imports Microsoft.Win32
Public Class frmAR
    Dim AppsOffice As String

    Private Sub frmAR_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim TargetKey As RegistryKey
        TargetKey = Registry.ClassesRoot.OpenSubKey("Excel.Application")
        If TargetKey Is Nothing Then
            AppsOffice = "XL_NotInstalled"
        Else
            AppsOffice = "XL_Installed"
            TargetKey.Close()
        End If
    End Sub

    Private Sub btnBrow_AR_src_Click(sender As Object, e As EventArgs) Handles btnBrow_AR_src.Click
        Dim DSMpathSrc As String
        If OFD_AR.ShowDialog = DialogResult.OK Then
            DSMpathSrc = OFD_AR.FileName
            txtAR_src.Text = DSMpathSrc
        End If
        If txtAR_dest.Text <> "" Then
            btnNeu_AR.Enabled = True
        Else
            btnNeu_AR.Enabled = False
        End If
    End Sub

    Private Sub btnBrow_AR_dest_Click(sender As Object, e As EventArgs) Handles btnBrow_AR_dest.Click
        Dim datenow As DateTime = DateTime.Now
        Dim filename As String = "Neutralize_ARReport_" & datenow.ToString("ddMMyyyy_HHmm")
        SFD_AR.FileName = filename
        Dim DSMpath_Dest As String
        If SFD_AR.ShowDialog = DialogResult.OK Then
            DSMpath_Dest = SFD_AR.FileName
            txtAR_dest.Text = DSMpath_Dest
        End If
        If txtAR_src.Text <> "" Then
            btnNeu_AR.Enabled = True
        Else
            btnNeu_AR.Enabled = False
        End If
        If txtAR_dest.Text <> "" Then
            btnNeu_AR.Enabled = True
        Else
            btnNeu_AR.Enabled = False
        End If
    End Sub

    Private Sub btnNeu_AR_Click(sender As Object, e As EventArgs) Handles btnNeu_AR.Click
        PicBar_AR.Visible = True
        BWAR.RunWorkerAsync()
    End Sub

    Private Sub BWAR_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BWAR.DoWork
        Select Case AppsOffice
            Case "XL_NotInstalled"
                Dim xlAppAR As Object
                Dim xlWbookAR As Object
                Dim xlWsheetAR As Object

                Try
                    xlAppAR = CreateObject("Ket.Application")
                    xlWbookAR = xlAppAR.Workbooks.Open(txtAR_src.Text)
                    xlWsheetAR = xlWbookAR.Worksheets("UID AR Report")

                    xlWsheetAR.UsedRange.UnMerge()
                    xlWsheetAR.UsedRange.WrapText = False
                    xlWsheetAR.UsedRange.ColumnWidth = 15
                    xlWsheetAR.UsedRange.RowHeight = 15

                    xlWsheetAR.Range("A:A").EntireColumn.Delete()

                    Dim rg_head_cut1 As Object = xlWsheetAR.Range("C2")
                    Dim rg_head_paste1 As Object = xlWsheetAR.Range("A2")
                    rg_head_cut1.Select()
                    rg_head_cut1.Cut(rg_head_paste1)

                    Dim rg_head_cut2 As Object = xlWsheetAR.Range("C4")
                    Dim rg_head_paste2 As Object = xlWsheetAR.Range("A3")
                    rg_head_cut2.Select()
                    rg_head_cut2.Cut(rg_head_paste2)

                    xlWsheetAR.Range("A2").RowHeight = 27

                    Dim paramhead1, paramhead2, paramhead3 As String
                    paramhead1 = xlWsheetAR.Range("B6").Value & " " & xlWsheetAR.Range("F6").Value & "; "
                    paramhead1 = paramhead1 & xlWsheetAR.Range("K6").Value & " " & xlWsheetAR.Range("O6").Value

                    paramhead2 = xlWsheetAR.Range("B7").Value & " " & xlWsheetAR.Range("F7").Value & "; "
                    paramhead2 = paramhead2 & xlWsheetAR.Range("K7").Value & " " & xlWsheetAR.Range("O7").Value & "; "
                    paramhead2 = paramhead2 & xlWsheetAR.Range("R7").Value & " " & xlWsheetAR.Range("W7").Value

                    paramhead3 = xlWsheetAR.Range("C9").Value & " " & xlWsheetAR.Range("F9").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetAR.Range("J9").Value & " " & xlWsheetAR.Range("O9").Value

                    xlWsheetAR.Range("A4").Value = paramhead1
                    xlWsheetAR.Range("A4").EntireRow.Font.Name = "Calibri"
                    xlWsheetAR.Range("A5").Value = paramhead2
                    xlWsheetAR.Range("A5").EntireRow.Font.Name = "Calibri"
                    xlWsheetAR.Range("A6").Value = paramhead3
                    xlWsheetAR.Range("A6").EntireRow.Font.Name = "Calibri"

                    xlWsheetAR.Range("B6:W7").Value = ""
                    xlWsheetAR.Range("A8:A9").EntireRow.Delete()

                    Dim xlfunc As Excel.WorksheetFunction
                    xlfunc = xlAppAR.WorksheetFunction
                    Dim lnCol As Long
                    Dim i, j As Long
                    Dim rnarea As Object = xlWsheetAR.UsedRange

                    lnCol = rnarea.Columns.Count
                    For i = lnCol To 1 Step -1
                        If xlfunc.CountA(rnarea.Columns(i)) = 0 Then
                            rnarea.Columns(i).Delete()
                            j = j + 1
                        End If
                    Next

                    xlWbookAR.SaveAs(txtAR_dest.Text)
                    xlWbookAR.Close()
                    xlAppAR.Quit()

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetAR)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookAR)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppAR)

                    xlWsheetAR = Nothing
                    xlWbookAR = Nothing
                    xlAppAR = Nothing

                    MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                Catch ex As Exception
                    MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try

            Case "XL_Installed"
                Dim xlAppAR As Object
                Dim xlWbookAR As Object
                Dim xlWsheetAR As Object

                Try
                    xlAppAR = CreateObject("Excel.Application")
                    xlWbookAR = xlAppAR.Workbooks.Open(txtAR_src.Text)
                    xlWsheetAR = xlWbookAR.Worksheets("UID AR Report")

                    xlWsheetAR.UsedRange.UnMerge()
                    xlWsheetAR.UsedRange.WrapText = False
                    xlWsheetAR.UsedRange.ColumnWidth = 15
                    xlWsheetAR.UsedRange.RowHeight = 15

                    xlWsheetAR.Range("A:A").EntireColumn.Delete()

                    Dim rg_head_cut1 As Excel.Range = xlWsheetAR.Range("C2")
                    Dim rg_head_paste1 As Excel.Range = xlWsheetAR.Range("A2")
                    rg_head_cut1.Select()
                    rg_head_cut1.Cut(rg_head_paste1)

                    Dim rg_head_cut2 As Excel.Range = xlWsheetAR.Range("C4")
                    Dim rg_head_paste2 As Excel.Range = xlWsheetAR.Range("A3")
                    rg_head_cut2.Select()
                    rg_head_cut2.Cut(rg_head_paste2)

                    xlWsheetAR.Range("A2").RowHeight = 27

                    Dim paramhead1, paramhead2, paramhead3 As String
                    paramhead1 = xlWsheetAR.Range("B6").Value & " " & xlWsheetAR.Range("F6").Value & "; "
                    paramhead1 = paramhead1 & xlWsheetAR.Range("K6").Value & " " & xlWsheetAR.Range("O6").Value

                    paramhead2 = xlWsheetAR.Range("B7").Value & " " & xlWsheetAR.Range("F7").Value & "; "
                    paramhead2 = paramhead2 & xlWsheetAR.Range("K7").Value & " " & xlWsheetAR.Range("O7").Value & "; "
                    paramhead2 = paramhead2 & xlWsheetAR.Range("R7").Value & " " & xlWsheetAR.Range("W7").Value

                    paramhead3 = xlWsheetAR.Range("C9").Value & " " & xlWsheetAR.Range("F9").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetAR.Range("J9").Value & " " & xlWsheetAR.Range("O9").Value

                    xlWsheetAR.Range("A4").Value = paramhead1
                    xlWsheetAR.Range("A4").EntireRow.Font.Name = "Calibri"
                    xlWsheetAR.Range("A5").Value = paramhead2
                    xlWsheetAR.Range("A5").EntireRow.Font.Name = "Calibri"
                    xlWsheetAR.Range("A6").Value = paramhead3
                    xlWsheetAR.Range("A6").EntireRow.Font.Name = "Calibri"

                    xlWsheetAR.Range("B6:W7").Value = ""
                    xlWsheetAR.Range("A8:A9").EntireRow.Delete()

                    Dim xlfunc As Excel.WorksheetFunction
                    xlfunc = xlAppAR.WorksheetFunction
                    Dim lnCol As Long
                    Dim i, j As Long
                    Dim rnarea As Excel.Range = xlWsheetAR.UsedRange

                    lnCol = rnarea.Columns.Count
                    For i = lnCol To 1 Step -1
                        If xlfunc.CountA(rnarea.Columns(i)) = 0 Then
                            rnarea.Columns(i).Delete()
                            j = j + 1
                        End If
                    Next

                    xlWbookAR.SaveAs(txtAR_dest.Text)
                    xlWbookAR.Close()
                    xlAppAR.Quit()

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetAR)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookAR)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppAR)

                    xlWsheetAR = Nothing
                    xlWbookAR = Nothing
                    xlAppAR = Nothing

                    MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                Catch ex As Exception
                    MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
        End Select
    End Sub

    Private Sub BWAR_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BWAR.RunWorkerCompleted
        PicBar_AR.Visible = False
        btnNeu_AR.Enabled = False
        txtAR_dest.Text = ""
        txtAR_src.Text = ""
    End Sub
End Class