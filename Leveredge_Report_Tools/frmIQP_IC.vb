Imports Microsoft.Office.Interop
Imports System.Windows.Forms
Imports System.Text
Imports System.Drawing
Imports Microsoft.Win32
Public Class frmIQP_IC
    Dim AppsOffice As String
    Private Sub frmIQP_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim TargetKey As RegistryKey
        TargetKey = Registry.ClassesRoot.OpenSubKey("Excel.Application")
        If TargetKey Is Nothing Then
            AppsOffice = "XL_NotInstalled"
        Else
            AppsOffice = "XL_Installed"
            TargetKey.Close()
        End If
    End Sub

    Private Sub btnBrow_IQP_IC_src_Click(sender As Object, e As EventArgs) Handles btnBrow_IQP_IC_src.Click
        Dim IQPpathSrc As String
        If OFD_IQP_IC.ShowDialog = DialogResult.OK Then
            IQPpathSrc = OFD_IQP_IC.FileName()
            txtIQP_IC_src.Text = IQPpathSrc
        End If
        If txtIQP_IC_dest.Text <> "" Then
            btnNeu_IQP_IC.Enabled = True
        Else
            btnNeu_IQP_IC.Enabled = False
        End If
    End Sub

    Private Sub btnBrow_IQP_IC_dest_Click(sender As Object, e As EventArgs) Handles btnBrow_IQP_IC_dest.Click
        Dim datenow As DateTime = DateTime.Now
        Dim filename As String = "Neutralize_IC_IQPerformance_" & datenow.ToString("ddMMyyyy_HHmm")
        SFD_IQP_IC.FileName = filename
        Dim IQPpath_Dest As String
        If SFD_IQP_IC.ShowDialog = DialogResult.OK Then
            IQPpath_Dest = SFD_IQP_IC.FileName
            txtIQP_IC_dest.Text = IQPpath_Dest
        End If
        If txtIQP_IC_src.Text <> "" Then
            btnNeu_IQP_IC.Enabled = True
        Else
            btnNeu_IQP_IC.Enabled = False
        End If
        If txtIQP_IC_dest.Text <> "" Then
            btnNeu_IQP_IC.Enabled = True
        Else
            btnNeu_IQP_IC.Enabled = False
        End If
    End Sub

    Private Sub btnNeu_IQP_IC_Click(sender As Object, e As EventArgs) Handles btnNeu_IQP_IC.Click
        PicBar_IQP_IC.Visible = True
        BWIQP_IC.RunWorkerAsync()
    End Sub

    Private Sub BWIQP_IC_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BWIQP_IC.DoWork
        Select Case AppsOffice
            Case "XL_NotInstalled"
                Dim xlAppIQP_IC As Object
                Dim xlWbookIQP_IC As Object
                Dim xlWsheetIQP_IC As Object
                Try
                    xlAppIQP_IC = CreateObject("Ket.Application")
                    xlWbookIQP_IC = xlAppIQP_IC.Workbooks.Open(txtIQP_IC_src.Text)
                    xlWsheetIQP_IC = xlWbookIQP_IC.Worksheets("UID IC IQ Perfomance Report")

                    xlWsheetIQP_IC.UsedRange.UnMerge()
                    xlWsheetIQP_IC.UsedRange.WrapText = False
                    xlWsheetIQP_IC.UsedRange.ColumnWidth = 15
                    xlWsheetIQP_IC.UsedRange.RowHeight = 15

                    Dim rg_head_cut1 As Object = xlWsheetIQP_IC.Range("B2")
                    Dim rg_head_paste1 As Object = xlWsheetIQP_IC.Range("A2")
                    rg_head_cut1.Select()
                    rg_head_cut1.Cut(rg_head_paste1)

                    Dim rg_head_cut2 As Object = xlWsheetIQP_IC.Range("B4")
                    Dim rg_head_paste2 As Object = xlWsheetIQP_IC.Range("A3")
                    rg_head_cut2.Select()
                    rg_head_cut2.Cut(rg_head_paste2)

                    xlWsheetIQP_IC.Range("A2").RowHeight = 27

                    Dim paramhead1, paramhead2 As String
                    paramhead1 = xlWsheetIQP_IC.Range("C6").Value & " " & xlWsheetIQP_IC.Range("E6").Value & "; "
                    paramhead1 = paramhead1 & xlWsheetIQP_IC.Range("G6").Value & " " & xlWsheetIQP_IC.Range("K6").Value

                    paramhead2 = xlWsheetIQP_IC.Range("N6").Value & " " & xlWsheetIQP_IC.Range("Q6").Value & ";"
                    paramhead2 = paramhead2 & xlWsheetIQP_IC.Range("C9").Value & " " & xlWsheetIQP_IC.Range("E9").Value

                    xlWsheetIQP_IC.Range("A5").Value = paramhead1
                    xlWsheetIQP_IC.Range("A5").EntireRow.Font.Name = "Calibri"
                    xlWsheetIQP_IC.Range("A6").Value = paramhead2
                    xlWsheetIQP_IC.Range("A6").EntireRow.Font.Name = "Calibri"

                    xlWsheetIQP_IC.Range("B6:AA9").value = ""
                    xlWsheetIQP_IC.Range("A8:A12").EntireRow.Delete()

                    Dim xlfunc As Excel.WorksheetFunction
                    xlfunc = xlAppIQP_IC.WorksheetFunction
                    Dim lnCol As Long
                    Dim i As Long
                    Dim rnarea As Object = xlWsheetIQP_IC.UsedRange

                    lnCol = rnarea.Columns.Count
                    For i = lnCol To 1 Step -1
                        If xlfunc.CountA(rnarea.Columns(i)) = 0 Then
                            rnarea.Columns(i).Delete()
                        End If
                    Next

                    xlWbookIQP_IC.SaveAs(txtIQP_IC_dest.Text)
                    xlWbookIQP_IC.Close()
                    xlAppIQP_IC.Quit()

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetIQP_IC)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookIQP_IC)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppIQP_IC)

                    xlWsheetIQP_IC = Nothing
                    xlWbookIQP_IC = Nothing
                    xlAppIQP_IC = Nothing

                    MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                Catch ex As Exception
                    MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try

            Case "XL_Installed"
                Dim xlAppIQP_IC As Object
                Dim xlWbookIQP_IC As Object
                Dim xlWsheetIQP_IC As Object
                Try
                    xlAppIQP_IC = CreateObject("Excel.Application")
                    xlWbookIQP_IC = xlAppIQP_IC.Workbooks.Open(txtIQP_IC_src.Text)
                    xlWsheetIQP_IC = xlWbookIQP_IC.Worksheets("UID IC IQ Perfomance Report")

                    xlWsheetIQP_IC.UsedRange.UnMerge()
                    xlWsheetIQP_IC.UsedRange.WrapText = False
                    xlWsheetIQP_IC.UsedRange.ColumnWidth = 15
                    xlWsheetIQP_IC.UsedRange.RowHeight = 15

                    Dim rg_head_cut1 As Excel.Range = xlWsheetIQP_IC.Range("B2")
                    Dim rg_head_paste1 As Excel.Range = xlWsheetIQP_IC.Range("A2")
                    rg_head_cut1.Select()
                    rg_head_cut1.Cut(rg_head_paste1)

                    Dim rg_head_cut2 As Excel.Range = xlWsheetIQP_IC.Range("B4")
                    Dim rg_head_paste2 As Excel.Range = xlWsheetIQP_IC.Range("A3")
                    rg_head_cut2.Select()
                    rg_head_cut2.Cut(rg_head_paste2)

                    xlWsheetIQP_IC.Range("A2").RowHeight = 27

                    Dim paramhead1, paramhead2 As String
                    paramhead1 = xlWsheetIQP_IC.Range("C6").Value & " " & xlWsheetIQP_IC.Range("E6").Value & "; "
                    paramhead1 = paramhead1 & xlWsheetIQP_IC.Range("G6").Value & " " & xlWsheetIQP_IC.Range("K6").Value

                    paramhead2 = xlWsheetIQP_IC.Range("N6").Value & " " & xlWsheetIQP_IC.Range("Q6").Value & ";"
                    paramhead2 = paramhead2 & xlWsheetIQP_IC.Range("C9").Value & " " & xlWsheetIQP_IC.Range("E9").Value

                    xlWsheetIQP_IC.Range("A5").Value = paramhead1
                    xlWsheetIQP_IC.Range("A5").EntireRow.Font.Name = "Calibri"
                    xlWsheetIQP_IC.Range("A6").Value = paramhead2
                    xlWsheetIQP_IC.Range("A6").EntireRow.Font.Name = "Calibri"

                    xlWsheetIQP_IC.Range("B6:AA9").value = ""
                    xlWsheetIQP_IC.Range("A8:A12").EntireRow.Delete()

                    Dim xlfunc As Excel.WorksheetFunction
                    xlfunc = xlAppIQP_IC.WorksheetFunction
                    Dim lnCol As Long
                    Dim i As Long
                    Dim rnarea As Excel.Range = xlWsheetIQP_IC.UsedRange

                    lnCol = rnarea.Columns.Count
                    For i = lnCol To 1 Step -1
                        If xlfunc.CountA(rnarea.Columns(i)) = 0 Then
                            rnarea.Columns(i).Delete()
                        End If
                    Next

                    xlWbookIQP_IC.SaveAs(txtIQP_IC_dest.Text)
                    xlWbookIQP_IC.Close()
                    xlAppIQP_IC.Quit()

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetIQP_IC)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookIQP_IC)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppIQP_IC)

                    xlWsheetIQP_IC = Nothing
                    xlWbookIQP_IC = Nothing
                    xlAppIQP_IC = Nothing

                    MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                Catch ex As Exception
                    MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
        End Select
    End Sub

    Private Sub BWIQP_IC_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BWIQP_IC.RunWorkerCompleted
        PicBar_IQP_IC.Visible = False
        btnNeu_IQP_IC.Enabled = False
        txtIQP_IC_dest.Text = ""
        txtIQP_IC_src.Text = ""
    End Sub
End Class