Imports Microsoft.Office.Interop
Imports System.Windows.Forms
Imports System.Text
Imports System.Drawing
Imports Microsoft.Win32
Public Class frmIQS_IC
    Dim AppsOffice As String
    Private Sub frmIQS_IC_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim TargetKey As RegistryKey
        TargetKey = Registry.ClassesRoot.OpenSubKey("Excel.Application")
        If TargetKey Is Nothing Then
            AppsOffice = "XL_NotInstalled"
        Else
            AppsOffice = "XL_Installed"
            TargetKey.Close()
        End If
    End Sub

    Private Sub btnBrow_IQS_IC_src_Click(sender As Object, e As EventArgs) Handles btnBrow_IQS_IC_src.Click
        Dim IQSpathSrc As String
        If OFD_IQS_IC.ShowDialog = DialogResult.OK Then
            IQSpathSrc = OFD_IQS_IC.FileName()
            txtIQS_IC_src.Text = IQSpathSrc
        End If
        If txtIQS_IC_dest.Text <> "" Then
            btnNeu_IQS_IC.Enabled = True
        Else
            btnNeu_IQS_IC.Enabled = False
        End If
    End Sub

    Private Sub btnBrow_IQS_IC_dest_Click(sender As Object, e As EventArgs) Handles btnBrow_IQS_IC_dest.Click
        Dim datenow As DateTime = DateTime.Now
        Dim filename As String = "Neutralize_IC_IQSummary_" & datenow.ToString("ddMMyyyy_HHmm")
        SFD_IQS_IC.FileName = filename
        Dim IQSpath_Dest As String
        If SFD_IQS_IC.ShowDialog = DialogResult.OK Then
            IQSpath_Dest = SFD_IQS_IC.FileName
            txtIQS_IC_dest.Text = IQSpath_Dest
        End If
        If txtIQS_IC_src.Text <> "" Then
            btnNeu_IQS_IC.Enabled = True
        Else
            btnNeu_IQS_IC.Enabled = False
        End If
        If txtIQS_IC_dest.Text <> "" Then
            btnNeu_IQS_IC.Enabled = True
        Else
            btnNeu_IQS_IC.Enabled = False
        End If
    End Sub

    Private Sub btnNeu_IQS_IC_Click(sender As Object, e As EventArgs) Handles btnNeu_IQS_IC.Click
        PicBar_IQS_IC.Visible = True
        BWIQS_IC.RunWorkerAsync()
    End Sub

    Private Sub BWIQS_IC_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BWIQS_IC.DoWork
        Select Case AppsOffice
            Case "XL_NotInstalled"
                Dim xlAppIQS_IC As Object
                Dim xlWbookIQS_IC As Object
                Dim xlWsheetIQS_IC As Object
                Try
                    xlAppIQS_IC = CreateObject("Ket.Application")
                    xlWbookIQS_IC = xlAppIQS_IC.Workbooks.Open(txtIQS_IC_src.Text)
                    xlWsheetIQS_IC = xlWbookIQS_IC.Worksheets("UID IC IQ Summary Report")

                    xlWsheetIQS_IC.UsedRange.UnMerge()
                    xlWsheetIQS_IC.UsedRange.WrapText = False
                    xlWsheetIQS_IC.UsedRange.ColumnWidth = 15
                    xlWsheetIQS_IC.UsedRange.RowHeight = 15

                    Dim rg_head_cut1 As Object = xlWsheetIQS_IC.Range("B2")
                    Dim rg_head_paste1 As Object = xlWsheetIQS_IC.Range("A2")
                    rg_head_cut1.Select()
                    rg_head_cut1.Cut(rg_head_paste1)

                    Dim rg_head_cut2 As Object = xlWsheetIQS_IC.Range("B4")
                    Dim rg_head_paste2 As Object = xlWsheetIQS_IC.Range("A3")
                    rg_head_cut2.Select()
                    rg_head_cut2.Cut(rg_head_paste2)

                    xlWsheetIQS_IC.Range("A2").RowHeight = 27

                    Dim paramhead1, paramhead2 As String
                    paramhead1 = xlWsheetIQS_IC.Range("C6").Value & " " & xlWsheetIQS_IC.Range("F6").Value & "; "
                    paramhead1 = paramhead1 & xlWsheetIQS_IC.Range("H6").Value & " " & xlWsheetIQS_IC.Range("K6").Value

                    paramhead2 = xlWsheetIQS_IC.Range("C8").Value & " " & xlWsheetIQS_IC.Range("F8").Value

                    xlWsheetIQS_IC.Range("A5").Value = paramhead1
                    xlWsheetIQS_IC.Range("A5").EntireRow.Font.Name = "Calibri"
                    xlWsheetIQS_IC.Range("A6").Value = paramhead2
                    xlWsheetIQS_IC.Range("A6").EntireRow.Font.Name = "Calibri"

                    xlWsheetIQS_IC.Range("B6:AA9").value = ""
                    xlWsheetIQS_IC.Range("A8:A10").EntireRow.Delete()

                    Dim xlfunc As Excel.WorksheetFunction
                    xlfunc = xlAppIQS_IC.WorksheetFunction
                    Dim lnCol As Long
                    Dim i As Long
                    Dim rnarea As Object = xlWsheetIQS_IC.UsedRange

                    lnCol = rnarea.Columns.Count
                    For i = lnCol To 1 Step -1
                        If xlfunc.CountA(rnarea.Columns(i)) = 0 Then
                            rnarea.Columns(i).Delete()
                        End If
                    Next

                    xlWbookIQS_IC.SaveAs(txtIQS_IC_dest.Text)
                    xlWbookIQS_IC.Close()
                    xlAppIQS_IC.Quit()

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetIQS_IC)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookIQS_IC)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppIQS_IC)

                    xlWsheetIQS_IC = Nothing
                    xlWbookIQS_IC = Nothing
                    xlAppIQS_IC = Nothing

                    MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                Catch ex As Exception
                    MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try

            Case "XL_Installed"
                Dim xlAppIQS_IC As Object
                Dim xlWbookIQS_IC As Object
                Dim xlWsheetIQS_IC As Object
                Try
                    xlAppIQS_IC = CreateObject("Excel.Application")
                    xlWbookIQS_IC = xlAppIQS_IC.Workbooks.Open(txtIQS_IC_src.Text)
                    xlWsheetIQS_IC = xlWbookIQS_IC.Worksheets("UID IC IQ Summary Report")

                    xlWsheetIQS_IC.UsedRange.UnMerge()
                    xlWsheetIQS_IC.UsedRange.WrapText = False
                    xlWsheetIQS_IC.UsedRange.ColumnWidth = 15
                    xlWsheetIQS_IC.UsedRange.RowHeight = 15

                    Dim rg_head_cut1 As Excel.Range = xlWsheetIQS_IC.Range("B2")
                    Dim rg_head_paste1 As Excel.Range = xlWsheetIQS_IC.Range("A2")
                    rg_head_cut1.Select()
                    rg_head_cut1.Cut(rg_head_paste1)

                    Dim rg_head_cut2 As Excel.Range = xlWsheetIQS_IC.Range("B4")
                    Dim rg_head_paste2 As Excel.Range = xlWsheetIQS_IC.Range("A3")
                    rg_head_cut2.Select()
                    rg_head_cut2.Cut(rg_head_paste2)

                    xlWsheetIQS_IC.Range("A2").RowHeight = 27

                    Dim paramhead1, paramhead2 As String
                    paramhead1 = xlWsheetIQS_IC.Range("C6").Value & " " & xlWsheetIQS_IC.Range("F6").Value & "; "
                    paramhead1 = paramhead1 & xlWsheetIQS_IC.Range("H6").Value & " " & xlWsheetIQS_IC.Range("K6").Value

                    paramhead2 = xlWsheetIQS_IC.Range("C8").Value & " " & xlWsheetIQS_IC.Range("F8").Value

                    xlWsheetIQS_IC.Range("A5").Value = paramhead1
                    xlWsheetIQS_IC.Range("A5").EntireRow.Font.Name = "Calibri"
                    xlWsheetIQS_IC.Range("A6").Value = paramhead2
                    xlWsheetIQS_IC.Range("A6").EntireRow.Font.Name = "Calibri"

                    xlWsheetIQS_IC.Range("B6:AA9").value = ""
                    xlWsheetIQS_IC.Range("A8:A10").EntireRow.Delete()

                    Dim xlfunc As Excel.WorksheetFunction
                    xlfunc = xlAppIQS_IC.WorksheetFunction
                    Dim lnCol As Long
                    Dim i As Long
                    Dim rnarea As Excel.Range = xlWsheetIQS_IC.UsedRange

                    lnCol = rnarea.Columns.Count
                    For i = lnCol To 1 Step -1
                        If xlfunc.CountA(rnarea.Columns(i)) = 0 Then
                            rnarea.Columns(i).Delete()
                        End If
                    Next

                    xlWbookIQS_IC.SaveAs(txtIQS_IC_dest.Text)
                    xlWbookIQS_IC.Close()
                    xlAppIQS_IC.Quit()

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetIQS_IC)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookIQS_IC)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppIQS_IC)

                    xlWsheetIQS_IC = Nothing
                    xlWbookIQS_IC = Nothing
                    xlAppIQS_IC = Nothing

                    MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                Catch ex As Exception
                    MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
        End Select
    End Sub

    Private Sub BWIQS_IC_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BWIQS_IC.RunWorkerCompleted
        PicBar_IQS_IC.Visible = False
        btnNeu_IQS_IC.Enabled = False
        txtIQS_IC_dest.Text = ""
        txtIQS_IC_src.Text = ""
    End Sub
End Class