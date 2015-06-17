Imports Microsoft.Office.Interop
Imports System.Windows.Forms
Imports System.Text
Imports System.Drawing
Imports Microsoft.Win32
Public Class frmCV_IC
    Dim AppsOffice As String
    Private Sub frmCV_IC_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim TargetKey As RegistryKey
        TargetKey = Registry.ClassesRoot.OpenSubKey("Excel.Application")
        If TargetKey Is Nothing Then
            AppsOffice = "XL_NotInstalled"
        Else
            AppsOffice = "XL_Installed"
            TargetKey.Close()
        End If
    End Sub

    Private Sub btnBrow_CV_IC_src_Click(sender As Object, e As EventArgs) Handles btnBrow_CV_IC_src.Click
        Dim CVpathSrc As String
        If OFD_CV_IC.ShowDialog = DialogResult.OK Then
            CVpathSrc = OFD_CV_IC.FileName()
            txtCV_IC_src.Text = CVpathSrc
        End If
        If txtCV_IC_dest.Text <> "" Then
            btnNeu_CV_IC.Enabled = True
        Else
            btnNeu_CV_IC.Enabled = False
        End If
    End Sub

    Private Sub btnBrow_CV_IC_dest_Click(sender As Object, e As EventArgs) Handles btnBrow_CV_IC_dest.Click
        Dim datenow As DateTime = DateTime.Now
        Dim filename As String = "Neutralize_IC_CustomerVisit_" & datenow.ToString("ddMMyyyy_HHmm")
        SFD_CV_IC.FileName = filename
        Dim CVpath_Dest As String
        If SFD_CV_IC.ShowDialog = DialogResult.OK Then
            CVpath_Dest = SFD_CV_IC.FileName
            txtCV_IC_dest.Text = CVpath_Dest
        End If
        If txtCV_IC_src.Text <> "" Then
            btnNeu_CV_IC.Enabled = True
        Else
            btnNeu_CV_IC.Enabled = False
        End If
        If txtCV_IC_dest.Text <> "" Then
            btnNeu_CV_IC.Enabled = True
        Else
            btnNeu_CV_IC.Enabled = False
        End If
    End Sub

    Private Sub btnNeu_CV_IC_Click(sender As Object, e As EventArgs) Handles btnNeu_CV_IC.Click
        PicBar_CV_IC.Visible = True
        BWCV_IC.RunWorkerAsync()
    End Sub

    Private Sub BWCV_IC_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BWCV_IC.DoWork
        Select Case AppsOffice
            Case "XL_NotInstalled"
                Dim xlAppCV_IC As Object
                Dim xlWbookCV_IC As Object
                Dim xlWsheetCV_IC As Object

                Try
                    xlAppCV_IC = CreateObject("Ket.Application")
                    xlWbookCV_IC = xlAppCV_IC.Workbooks.Open(txtCV_IC_src.Text)
                    xlWsheetCV_IC = xlWbookCV_IC.Worksheets("UID IC Customer Visit Report")

                    xlWsheetCV_IC.UsedRange.UnMerge()
                    xlWsheetCV_IC.UsedRange.WrapText = False
                    xlWsheetCV_IC.UsedRange.ColumnWidth = 15
                    xlWsheetCV_IC.UsedRange.RowHeight = 15

                    Dim rg_head_cut1 As Object = xlWsheetCV_IC.Range("B2")
                    Dim rg_head_paste1 As Object = xlWsheetCV_IC.Range("A2")
                    rg_head_cut1.Select()
                    rg_head_cut1.Cut(rg_head_paste1)

                    Dim rg_head_cut2 As Object = xlWsheetCV_IC.Range("B4")
                    Dim rg_head_paste2 As Object = xlWsheetCV_IC.Range("A3")
                    rg_head_cut2.Select()
                    rg_head_cut2.Cut(rg_head_paste2)

                    xlWsheetCV_IC.Range("A2").RowHeight = 27

                    Dim paramhead1, paramhead2, paramhead3 As String
                    paramhead1 = xlWsheetCV_IC.Range("B6").Value & " " & xlWsheetCV_IC.Range("E6").Value & "; "
                    paramhead1 = paramhead1 & xlWsheetCV_IC.Range("H6").Value & " " & xlWsheetCV_IC.Range("J6").Value

                    paramhead2 = xlWsheetCV_IC.Range("P6").Value & " " & xlWsheetCV_IC.Range("S6").Value & ";"
                    paramhead2 = paramhead2 & xlWsheetCV_IC.Range("W6").Value & " " & xlWsheetCV_IC.Range("Z6").Value

                    paramhead3 = xlWsheetCV_IC.Range("AC6").Value & " " & xlWsheetCV_IC.Range("AE6").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetCV_IC.Range("AJ6").Value & " " & xlWsheetCV_IC.Range("AL6").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetCV_IC.Range("B8").Value & " " & xlWsheetCV_IC.Range("E8").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetCV_IC.Range("H8").Value & " " & xlWsheetCV_IC.Range("M8").Value & "; "

                    xlWsheetCV_IC.Range("A5").Value = paramhead1
                    xlWsheetCV_IC.Range("A5").EntireRow.Font.Name = "Calibri"
                    xlWsheetCV_IC.Range("A6").Value = paramhead2
                    xlWsheetCV_IC.Range("A6").EntireRow.Font.Name = "Calibri"
                    xlWsheetCV_IC.Range("A7").Value = paramhead3
                    xlWsheetCV_IC.Range("A7").EntireRow.Font.Name = "Calibri"

                    xlWsheetCV_IC.Range("B6:AL8").value = ""
                    xlWsheetCV_IC.Range("A9").EntireRow.Delete()

                    Dim xlfunc As Excel.WorksheetFunction
                    xlfunc = xlAppCV_IC.WorksheetFunction
                    Dim lnCol As Long
                    Dim i As Long
                    Dim rnarea As Object = xlWsheetCV_IC.UsedRange

                    lnCol = rnarea.Columns.Count
                    For i = lnCol To 1 Step -1
                        If xlfunc.CountA(rnarea.Columns(i)) = 0 Then
                            rnarea.Columns(i).Delete()
                        End If
                    Next

                    xlWbookCV_IC.SaveAs(txtCV_IC_dest.Text)
                    xlWbookCV_IC.Close()
                    xlAppCV_IC.Quit()

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetCV_IC)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookCV_IC)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppCV_IC)

                    xlWsheetCV_IC = Nothing
                    xlWbookCV_IC = Nothing
                    xlAppCV_IC = Nothing

                    MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                Catch ex As Exception
                    MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try

            Case "XL_Installed"
                Dim xlAppCV_IC As Object
                Dim xlWbookCV_IC As Object
                Dim xlWsheetCV_IC As Object

                Try
                    xlAppCV_IC = CreateObject("Excel.Application")
                    xlWbookCV_IC = xlAppCV_IC.Workbooks.Open(txtCV_IC_src.Text)
                    xlWsheetCV_IC = xlWbookCV_IC.Worksheets("UID IC Customer Visit Report")

                    xlWsheetCV_IC.UsedRange.UnMerge()
                    xlWsheetCV_IC.UsedRange.WrapText = False
                    xlWsheetCV_IC.UsedRange.ColumnWidth = 15
                    xlWsheetCV_IC.UsedRange.RowHeight = 15

                    Dim rg_head_cut1 As Excel.Range = xlWsheetCV_IC.Range("B2")
                    Dim rg_head_paste1 As Excel.Range = xlWsheetCV_IC.Range("A2")
                    rg_head_cut1.Select()
                    rg_head_cut1.Cut(rg_head_paste1)

                    Dim rg_head_cut2 As Excel.Range = xlWsheetCV_IC.Range("B4")
                    Dim rg_head_paste2 As Excel.Range = xlWsheetCV_IC.Range("A3")
                    rg_head_cut2.Select()
                    rg_head_cut2.Cut(rg_head_paste2)

                    xlWsheetCV_IC.Range("A2").RowHeight = 27

                    Dim paramhead1, paramhead2, paramhead3 As String
                    paramhead1 = xlWsheetCV_IC.Range("B6").Value & " " & xlWsheetCV_IC.Range("E6").Value & "; "
                    paramhead1 = paramhead1 & xlWsheetCV_IC.Range("H6").Value & " " & xlWsheetCV_IC.Range("J6").Value

                    paramhead2 = xlWsheetCV_IC.Range("P6").Value & " " & xlWsheetCV_IC.Range("S6").Value & ";"
                    paramhead2 = paramhead2 & xlWsheetCV_IC.Range("W6").Value & " " & xlWsheetCV_IC.Range("Z6").Value

                    paramhead3 = xlWsheetCV_IC.Range("AC6").Value & " " & xlWsheetCV_IC.Range("AE6").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetCV_IC.Range("AJ6").Value & " " & xlWsheetCV_IC.Range("AL6").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetCV_IC.Range("B8").Value & " " & xlWsheetCV_IC.Range("E8").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetCV_IC.Range("H8").Value & " " & xlWsheetCV_IC.Range("M8").Value & "; "

                    xlWsheetCV_IC.Range("A5").Value = paramhead1
                    xlWsheetCV_IC.Range("A5").EntireRow.Font.Name = "Calibri"
                    xlWsheetCV_IC.Range("A6").Value = paramhead2
                    xlWsheetCV_IC.Range("A6").EntireRow.Font.Name = "Calibri"
                    xlWsheetCV_IC.Range("A7").Value = paramhead3
                    xlWsheetCV_IC.Range("A7").EntireRow.Font.Name = "Calibri"

                    xlWsheetCV_IC.Range("B6:AL8").value = ""
                    xlWsheetCV_IC.Range("A9").EntireRow.Delete()

                    Dim xlfunc As Excel.WorksheetFunction
                    xlfunc = xlAppCV_IC.WorksheetFunction
                    Dim lnCol As Long
                    Dim i As Long
                    Dim rnarea As Excel.Range = xlWsheetCV_IC.UsedRange

                    lnCol = rnarea.Columns.Count
                    For i = lnCol To 1 Step -1
                        If xlfunc.CountA(rnarea.Columns(i)) = 0 Then
                            rnarea.Columns(i).Delete()
                        End If
                    Next

                    xlWbookCV_IC.SaveAs(txtCV_IC_dest.Text)
                    xlWbookCV_IC.Close()
                    xlAppCV_IC.Quit()

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetCV_IC)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookCV_IC)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppCV_IC)

                    xlWsheetCV_IC = Nothing
                    xlWbookCV_IC = Nothing
                    xlAppCV_IC = Nothing

                    MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                Catch ex As Exception
                    MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try

        End Select
    End Sub

    Private Sub BWCV_IC_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BWCV_IC.RunWorkerCompleted
        PicBar_CV_IC.Visible = False
        btnNeu_CV_IC.Enabled = False
        txtCV_IC_dest.Text = ""
        txtCV_IC_src.Text = ""
    End Sub
End Class