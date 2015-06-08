Imports Microsoft.Office.Interop
Imports System.Windows.Forms
Imports System.Text
Imports System.Drawing
Imports Microsoft.Win32
Public Class frmABR_IC
    Dim AppsOffice As String
    Private Sub frmABR_IC_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim TargetKey As RegistryKey
        TargetKey = Registry.ClassesRoot.OpenSubKey("Excel.Application")
        If TargetKey Is Nothing Then
            AppsOffice = "XL_NotInstalled"
        Else
            AppsOffice = "XL_Installed"
            TargetKey.Close()
        End If
    End Sub

    Private Sub btnBrow_ABR_IC_src_Click(sender As Object, e As EventArgs) Handles btnBrow_ABR_IC_src.Click
        Dim ABRpathSrc As String
        If OFD_ABR_IC.ShowDialog = DialogResult.OK Then
            ABRpathSrc = OFD_ABR_IC.FileName()
            txtABR_IC_src.Text = ABRpathSrc
        End If
        If txtABR_IC_dest.Text <> "" Then
            btnNeu_ABR_IC.Enabled = True
        Else
            btnNeu_ABR_IC.Enabled = False
        End If
    End Sub

    Private Sub btnBrow_ABR_IC_dest_Click(sender As Object, e As EventArgs) Handles btnBrow_ABR_IC_dest.Click
        Dim datenow As DateTime = DateTime.Now
        Dim filename As String = "Neutralize_IC_AnalysisBackup_" & datenow.ToString("ddMMyyyy_HHmm")
        SFD_ABR_IC.FileName = filename
        Dim ABRpath_Dest As String
        If SFD_ABR_IC.ShowDialog = DialogResult.OK Then
            ABRpath_Dest = SFD_ABR_IC.FileName
            txtABR_IC_dest.Text = ABRpath_Dest
        End If
        If txtABR_IC_src.Text <> "" Then
            btnNeu_ABR_IC.Enabled = True
        Else
            btnNeu_ABR_IC.Enabled = False
        End If
        If txtABR_IC_dest.Text <> "" Then
            btnNeu_ABR_IC.Enabled = True
        Else
            btnNeu_ABR_IC.Enabled = False
        End If
    End Sub

    Private Sub btnNeu_ABR_IC_Click(sender As Object, e As EventArgs) Handles btnNeu_ABR_IC.Click
        PicBar_ABR_IC.Visible = True
        BWABR_IC.RunWorkerAsync()
    End Sub

    Private Sub BWABR_IC_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BWABR_IC.DoWork
        Select Case AppsOffice
            Case "XL_NotInstalled"
                Dim xlAppABR_IC As Object
                Dim xlWbookABR_IC As Object
                Dim xlWsheetABR_IC As Object
                Try
                    xlAppABR_IC = CreateObject("Ket.Application")
                    xlWbookABR_IC = xlAppABR_IC.Workbooks.Open(txtABR_IC_src.Text)
                    xlWsheetABR_IC = xlWbookABR_IC.Worksheets("UID IC Analysis Backup Report")

                    xlWsheetABR_IC.UsedRange.UnMerge()
                    xlWsheetABR_IC.UsedRange.WrapText = False
                    xlWsheetABR_IC.UsedRange.ColumnWidth = 15
                    xlWsheetABR_IC.UsedRange.RowHeight = 15

                    Dim rg_head_cut1 As Object = xlWsheetABR_IC.Range("B2")
                    Dim rg_head_paste1 As Object = xlWsheetABR_IC.Range("A2")
                    rg_head_cut1.Select()
                    rg_head_cut1.Cut(rg_head_paste1)

                    Dim rg_head_cut2 As Object = xlWsheetABR_IC.Range("B4")
                    Dim rg_head_paste2 As Object = xlWsheetABR_IC.Range("A3")
                    rg_head_cut2.Select()
                    rg_head_cut2.Cut(rg_head_paste2)

                    xlWsheetABR_IC.Range("A2").RowHeight = 27

                    Dim paramhead1, paramhead2, paramhead3 As String
                    paramhead1 = xlWsheetABR_IC.Range("B6").Value & " " & xlWsheetABR_IC.Range("D6").Value & "; "
                    paramhead1 = paramhead1 & xlWsheetABR_IC.Range("F6").Value & " " & xlWsheetABR_IC.Range("I6").Value

                    paramhead2 = xlWsheetABR_IC.Range("M6").Value & " " & xlWsheetABR_IC.Range("P6").Value & ";"
                    paramhead2 = paramhead2 & xlWsheetABR_IC.Range("B9").Value & " " & xlWsheetABR_IC.Range("D9").Value

                    paramhead3 = xlWsheetABR_IC.Range("F9").Value & " " & xlWsheetABR_IC.Range("I9").Value & "; "

                    xlWsheetABR_IC.Range("A5").Value = paramhead1
                    xlWsheetABR_IC.Range("A5").EntireRow.Font.Name = "Calibri"
                    xlWsheetABR_IC.Range("A6").Value = paramhead2
                    xlWsheetABR_IC.Range("A6").EntireRow.Font.Name = "Calibri"
                    xlWsheetABR_IC.Range("A7").Value = paramhead3
                    xlWsheetABR_IC.Range("A7").EntireRow.Font.Name = "Calibri"

                    xlWsheetABR_IC.Range("B6:AA9").value = ""
                    xlWsheetABR_IC.Range("A9:A10").EntireRow.Delete()

                    Dim xlfunc As Excel.WorksheetFunction
                    xlfunc = xlAppABR_IC.WorksheetFunction
                    Dim lnCol As Long
                    Dim i As Long
                    Dim rnarea As Object = xlWsheetABR_IC.UsedRange

                    lnCol = rnarea.Columns.Count
                    For i = lnCol To 1 Step -1
                        If xlfunc.CountA(rnarea.Columns(i)) = 0 Then
                            rnarea.Columns(i).Delete()
                        End If
                    Next

                    xlWbookABR_IC.SaveAs(txtABR_IC_dest.Text)
                    xlWbookABR_IC.Close()
                    xlAppABR_IC.Quit()

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetABR_IC)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookABR_IC)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppABR_IC)

                    xlWsheetABR_IC = Nothing
                    xlWbookABR_IC = Nothing
                    xlAppABR_IC = Nothing

                    MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                Catch ex As Exception
                    MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try

            Case "XL_Installed"
                Dim xlAppABR_IC As Object
                Dim xlWbookABR_IC As Object
                Dim xlWsheetABR_IC As Object
                Try
                    xlAppABR_IC = CreateObject("Excel.Application")
                    xlWbookABR_IC = xlAppABR_IC.Workbooks.Open(txtABR_IC_src.Text)
                    xlWsheetABR_IC = xlWbookABR_IC.Worksheets("UID IC Analysis Backup Report")

                    xlWsheetABR_IC.UsedRange.UnMerge()
                    xlWsheetABR_IC.UsedRange.WrapText = False
                    xlWsheetABR_IC.UsedRange.ColumnWidth = 15
                    xlWsheetABR_IC.UsedRange.RowHeight = 15

                    Dim rg_head_cut1 As Excel.Range = xlWsheetABR_IC.Range("B2")
                    Dim rg_head_paste1 As Excel.Range = xlWsheetABR_IC.Range("A2")
                    rg_head_cut1.Select()
                    rg_head_cut1.Cut(rg_head_paste1)

                    Dim rg_head_cut2 As Excel.Range = xlWsheetABR_IC.Range("B4")
                    Dim rg_head_paste2 As Excel.Range = xlWsheetABR_IC.Range("A3")
                    rg_head_cut2.Select()
                    rg_head_cut2.Cut(rg_head_paste2)

                    xlWsheetABR_IC.Range("A2").RowHeight = 27

                    Dim paramhead1, paramhead2, paramhead3 As String
                    paramhead1 = xlWsheetABR_IC.Range("B6").Value & " " & xlWsheetABR_IC.Range("D6").Value & "; "
                    paramhead1 = paramhead1 & xlWsheetABR_IC.Range("F6").Value & " " & xlWsheetABR_IC.Range("I6").Value

                    paramhead2 = xlWsheetABR_IC.Range("M6").Value & " " & xlWsheetABR_IC.Range("P6").Value & ";"
                    paramhead2 = paramhead2 & xlWsheetABR_IC.Range("B9").Value & " " & xlWsheetABR_IC.Range("D9").Value

                    paramhead3 = xlWsheetABR_IC.Range("F9").Value & " " & xlWsheetABR_IC.Range("I9").Value & "; "

                    xlWsheetABR_IC.Range("A5").Value = paramhead1
                    xlWsheetABR_IC.Range("A5").EntireRow.Font.Name = "Calibri"
                    xlWsheetABR_IC.Range("A6").Value = paramhead2
                    xlWsheetABR_IC.Range("A6").EntireRow.Font.Name = "Calibri"
                    xlWsheetABR_IC.Range("A7").Value = paramhead3
                    xlWsheetABR_IC.Range("A7").EntireRow.Font.Name = "Calibri"

                    xlWsheetABR_IC.Range("B6:AA9").value = ""
                    xlWsheetABR_IC.Range("A9:A10").EntireRow.Delete()

                    Dim xlfunc As Excel.WorksheetFunction
                    xlfunc = xlAppABR_IC.WorksheetFunction
                    Dim lnCol As Long
                    Dim i As Long
                    Dim rnarea As Excel.Range = xlWsheetABR_IC.UsedRange

                    lnCol = rnarea.Columns.Count
                    For i = lnCol To 1 Step -1
                        If xlfunc.CountA(rnarea.Columns(i)) = 0 Then
                            rnarea.Columns(i).Delete()
                        End If
                    Next

                    xlWbookABR_IC.SaveAs(txtABR_IC_dest.Text)
                    xlWbookABR_IC.Close()
                    xlAppABR_IC.Quit()

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetABR_IC)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookABR_IC)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppABR_IC)

                    xlWsheetABR_IC = Nothing
                    xlWbookABR_IC = Nothing
                    xlAppABR_IC = Nothing

                    MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                Catch ex As Exception
                    MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
        End Select
    End Sub

    Private Sub BWABR_IC_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BWABR_IC.RunWorkerCompleted
        PicBar_ABR_IC.Visible = False
        btnNeu_ABR_IC.Enabled = False
        txtABR_IC_dest.Text = ""
        txtABR_IC_src.Text = ""
    End Sub
End Class