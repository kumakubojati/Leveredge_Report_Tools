Imports Microsoft.Office.Interop
Imports System.Windows.Forms
Imports System.Text
Imports System.Drawing
Imports Microsoft.Win32
Public Class frmTPO_IC
    Dim AppsOffice As String
    Private Sub frmTPO_IC_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim TargetKey As RegistryKey
        TargetKey = Registry.ClassesRoot.OpenSubKey("Excel.Application")
        If TargetKey Is Nothing Then
            AppsOffice = "XL_NotInstalled"
        Else
            AppsOffice = "XL_Installed"
            TargetKey.Close()
        End If
    End Sub

    Private Sub btnBrow_TPO_IC_src_Click(sender As Object, e As EventArgs) Handles btnBrow_TPO_IC_src.Click
        Dim TPOpathSrc As String
        If OFD_TPO_IC.ShowDialog = DialogResult.OK Then
            TPOpathSrc = OFD_TPO_IC.FileName()
            txtTPO_IC_src.Text = TPOpathSrc
        End If
        If txtTPO_IC_dest.Text <> "" Then
            btnNeu_TPO_IC.Enabled = True
        Else
            btnNeu_TPO_IC.Enabled = False
        End If
    End Sub

    Private Sub btnBrow_TPO_IC_dest_Click(sender As Object, e As EventArgs) Handles btnBrow_TPO_IC_dest.Click
        Dim datenow As DateTime = DateTime.Now
        Dim filename As String = "Neutralize_IC_ThroughPutOutlet_" & datenow.ToString("ddMMyyyy_HHmm")
        SFD_TPO_IC.FileName = filename
        Dim TPOpath_Dest As String
        If SFD_TPO_IC.ShowDialog = DialogResult.OK Then
            TPOpath_Dest = SFD_TPO_IC.FileName
            txtTPO_IC_dest.Text = TPOpath_Dest
        End If
        If txtTPO_IC_src.Text <> "" Then
            btnNeu_TPO_IC.Enabled = True
        Else
            btnNeu_TPO_IC.Enabled = False
        End If
        If txtTPO_IC_dest.Text <> "" Then
            btnNeu_TPO_IC.Enabled = True
        Else
            btnNeu_TPO_IC.Enabled = False
        End If
    End Sub

    Private Sub btnNeu_TPO_IC_Click(sender As Object, e As EventArgs) Handles btnNeu_TPO_IC.Click
        PicBar_TPO_IC.Visible = True
        BWTPO_IC.RunWorkerAsync()
    End Sub

    Private Sub BWTPO_IC_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BWTPO_IC.DoWork
        Select Case AppsOffice
            Case "XL_NotInstalled"
                Dim xlAppTPO_IC As Object
                Dim xlWbookTPO_IC As Object
                Dim xlWsheetTPO_IC As Object

                Try
                    xlAppTPO_IC = CreateObject("Ket.Application")
                    xlWbookTPO_IC = xlAppTPO_IC.Workbooks.Open(txtTPO_IC_src.Text)
                    xlWsheetTPO_IC = xlWbookTPO_IC.Worksheets("UID IC ThroughPut Outlet Report")

                    xlWsheetTPO_IC.UsedRange.UnMerge()
                    xlWsheetTPO_IC.UsedRange.WrapText = False
                    xlWsheetTPO_IC.UsedRange.ColumnWidth = 15
                    xlWsheetTPO_IC.UsedRange.RowHeight = 15

                    Dim rg_head_cut1 As Object = xlWsheetTPO_IC.Range("B2")
                    Dim rg_head_paste1 As Object = xlWsheetTPO_IC.Range("A2")
                    rg_head_cut1.Select()
                    rg_head_cut1.Cut(rg_head_paste1)

                    Dim rg_head_cut2 As Object = xlWsheetTPO_IC.Range("B4")
                    Dim rg_head_paste2 As Object = xlWsheetTPO_IC.Range("A3")
                    rg_head_cut2.Select()
                    rg_head_cut2.Cut(rg_head_paste2)

                    xlWsheetTPO_IC.Range("A2").RowHeight = 27

                    Dim paramhead1, paramhead2, paramhead3 As String
                    paramhead1 = xlWsheetTPO_IC.Range("B6").Value & " " & xlWsheetTPO_IC.Range("D6").Value & "; "
                    paramhead1 = paramhead1 & xlWsheetTPO_IC.Range("G6").Value & " " & xlWsheetTPO_IC.Range("I6").Value

                    paramhead2 = xlWsheetTPO_IC.Range("N6").Value & " " & xlWsheetTPO_IC.Range("Q6").Value & ";"
                    paramhead2 = paramhead2 & xlWsheetTPO_IC.Range("U6").Value & " " & xlWsheetTPO_IC.Range("Z6").Value & "; "
                    paramhead2 = paramhead2 & xlWsheetTPO_IC.Range("AE6").Value & " " & xlWsheetTPO_IC.Range("AG6").Value


                    paramhead3 = xlWsheetTPO_IC.Range("B8").Value & " " & xlWsheetTPO_IC.Range("D8").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetTPO_IC.Range("G8").Value & " " & xlWsheetTPO_IC.Range("I8").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetTPO_IC.Range("N8").Value & " " & xlWsheetTPO_IC.Range("Q8").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetTPO_IC.Range("V8").Value & " " & xlWsheetTPO_IC.Range("AA8").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetTPO_IC.Range("AE8").Value & " " & xlWsheetTPO_IC.Range("AG8").Value

                    xlWsheetTPO_IC.Range("A5").Value = paramhead1
                    xlWsheetTPO_IC.Range("A5").EntireRow.Font.Name = "Calibri"
                    xlWsheetTPO_IC.Range("A6").Value = paramhead2
                    xlWsheetTPO_IC.Range("A6").EntireRow.Font.Name = "Calibri"
                    xlWsheetTPO_IC.Range("A7").Value = paramhead3
                    xlWsheetTPO_IC.Range("A7").EntireRow.Font.Name = "Calibri"

                    xlWsheetTPO_IC.Range("B6:AH8").value = ""
                    xlWsheetTPO_IC.Range("A9").EntireRow.Delete()

                    Dim xlfunc As Excel.WorksheetFunction
                    xlfunc = xlAppTPO_IC.WorksheetFunction()
                    Dim lnCol As Long
                    Dim i As Long
                    Dim rnarea As Object = xlWsheetTPO_IC.UsedRange

                    lnCol = rnarea.Columns.Count
                    For i = lnCol To 1 Step -1
                        If xlfunc.CountA(rnarea.Columns(i)) = 0 Then
                            rnarea.Columns(i).Delete()
                        End If
                    Next

                    xlWbookTPO_IC.SaveAs(txtTPO_IC_dest.Text)
                    xlWbookTPO_IC.Close()
                    xlAppTPO_IC.Quit()

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetTPO_IC)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookTPO_IC)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppTPO_IC)

                    xlWsheetTPO_IC = Nothing
                    xlWbookTPO_IC = Nothing
                    xlAppTPO_IC = Nothing

                    MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                Catch ex As Exception
                    MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try

            Case "XL_Installed"
                Dim xlAppTPO_IC As Object
                Dim xlWbookTPO_IC As Object
                Dim xlWsheetTPO_IC As Object

                Try
                    xlAppTPO_IC = CreateObject("Excel.Application")
                    xlWbookTPO_IC = xlAppTPO_IC.Workbooks.Open(txtTPO_IC_src.Text)
                    xlWsheetTPO_IC = xlWbookTPO_IC.Worksheets("UID IC ThroughPut Outlet Report")

                    xlWsheetTPO_IC.UsedRange.UnMerge()
                    xlWsheetTPO_IC.UsedRange.WrapText = False
                    xlWsheetTPO_IC.UsedRange.ColumnWidth = 15
                    xlWsheetTPO_IC.UsedRange.RowHeight = 15

                    Dim rg_head_cut1 As Excel.Range = xlWsheetTPO_IC.Range("B2")
                    Dim rg_head_paste1 As Excel.Range = xlWsheetTPO_IC.Range("A2")
                    rg_head_cut1.Select()
                    rg_head_cut1.Cut(rg_head_paste1)

                    Dim rg_head_cut2 As Excel.Range = xlWsheetTPO_IC.Range("B4")
                    Dim rg_head_paste2 As Excel.Range = xlWsheetTPO_IC.Range("A3")
                    rg_head_cut2.Select()
                    rg_head_cut2.Cut(rg_head_paste2)

                    xlWsheetTPO_IC.Range("A2").RowHeight = 27

                    Dim paramhead1, paramhead2, paramhead3 As String
                    paramhead1 = xlWsheetTPO_IC.Range("B6").Value & " " & xlWsheetTPO_IC.Range("D6").Value & "; "
                    paramhead1 = paramhead1 & xlWsheetTPO_IC.Range("G6").Value & " " & xlWsheetTPO_IC.Range("I6").Value

                    paramhead2 = xlWsheetTPO_IC.Range("N6").Value & " " & xlWsheetTPO_IC.Range("Q6").Value & ";"
                    paramhead2 = paramhead2 & xlWsheetTPO_IC.Range("U6").Value & " " & xlWsheetTPO_IC.Range("Z6").Value & "; "
                    paramhead2 = paramhead2 & xlWsheetTPO_IC.Range("AE6").Value & " " & xlWsheetTPO_IC.Range("AG6").Value


                    paramhead3 = xlWsheetTPO_IC.Range("B8").Value & " " & xlWsheetTPO_IC.Range("D8").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetTPO_IC.Range("G8").Value & " " & xlWsheetTPO_IC.Range("I8").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetTPO_IC.Range("N8").Value & " " & xlWsheetTPO_IC.Range("Q8").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetTPO_IC.Range("V8").Value & " " & xlWsheetTPO_IC.Range("AA8").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetTPO_IC.Range("AE8").Value & " " & xlWsheetTPO_IC.Range("AG8").Value

                    xlWsheetTPO_IC.Range("A5").Value = paramhead1
                    xlWsheetTPO_IC.Range("A5").EntireRow.Font.Name = "Calibri"
                    xlWsheetTPO_IC.Range("A6").Value = paramhead2
                    xlWsheetTPO_IC.Range("A6").EntireRow.Font.Name = "Calibri"
                    xlWsheetTPO_IC.Range("A7").Value = paramhead3
                    xlWsheetTPO_IC.Range("A7").EntireRow.Font.Name = "Calibri"

                    xlWsheetTPO_IC.Range("B6:AH8").value = ""
                    xlWsheetTPO_IC.Range("A9").EntireRow.Delete()

                    Dim xlfunc As Excel.WorksheetFunction
                    xlfunc = xlAppTPO_IC.WorksheetFunction()
                    Dim lnCol As Long
                    Dim i As Long
                    Dim rnarea As Excel.Range = xlWsheetTPO_IC.UsedRange

                    lnCol = rnarea.Columns.Count
                    For i = lnCol To 1 Step -1
                        If xlfunc.CountA(rnarea.Columns(i)) = 0 Then
                            rnarea.Columns(i).Delete()
                        End If
                    Next

                    xlWbookTPO_IC.SaveAs(txtTPO_IC_dest.Text)
                    xlWbookTPO_IC.Close()
                    xlAppTPO_IC.Quit()

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetTPO_IC)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookTPO_IC)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppTPO_IC)

                    xlWsheetTPO_IC = Nothing
                    xlWbookTPO_IC = Nothing
                    xlAppTPO_IC = Nothing

                    MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                Catch ex As Exception
                    MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
        End Select
    End Sub

    Private Sub BWTPO_IC_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BWTPO_IC.RunWorkerCompleted
        PicBar_TPO_IC.Visible = False
        btnNeu_TPO_IC.Enabled = False
        txtTPO_IC_dest.Text = ""
        txtTPO_IC_src.Text = ""
    End Sub
End Class