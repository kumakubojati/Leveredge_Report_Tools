Imports Microsoft.Office.Interop
Imports System.Windows.Forms
Imports System.Text
Imports System.Drawing
Imports Microsoft.Win32
Public Class frmLPU
    Dim AppsOffice As String
    Private Sub frmLPU_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim TargetKey As RegistryKey
        TargetKey = Registry.ClassesRoot.OpenSubKey("Excel.Application")
        If TargetKey Is Nothing Then
            AppsOffice = "XL_NotInstalled"
        Else
            AppsOffice = "XL_Installed"
            TargetKey.Close()
        End If
    End Sub

    Private Sub btnBrow_LPU_src_Click(sender As Object, e As EventArgs) Handles btnBrow_LPU_src.Click
        Dim LPUpathSrc As String
        If OFD_LPU.ShowDialog = DialogResult.OK Then
            LPUpathSrc = OFD_LPU.FileName()
            txtLPU_src.Text = LPUpathSrc
        End If
        If txtLPU_dest.Text <> "" Then
            btnNeu_LPU.Enabled = True
        Else
            btnNeu_LPU.Enabled = False
        End If
    End Sub

    Private Sub btnBrow_LPU_dest_Click(sender As Object, e As EventArgs) Handles btnBrow_LPU_dest.Click
        Dim datenow As DateTime = DateTime.Now
        Dim filename As String = "Neutralize_ListPromoUtilize_" & datenow.ToString("ddMMyyyy_HHmm")
        SFD_LPU.FileName = filename
        Dim LPUpath_Dest As String
        If SFD_LPU.ShowDialog = DialogResult.OK Then
            LPUpath_Dest = SFD_LPU.FileName
            txtLPU_dest.Text = LPUpath_Dest
        End If
        If txtLPU_src.Text <> "" Then
            btnNeu_LPU.Enabled = True
        Else
            btnNeu_LPU.Enabled = False
        End If
        If txtLPU_dest.Text <> "" Then
            btnNeu_LPU.Enabled = True
        Else
            btnNeu_LPU.Enabled = False
        End If
    End Sub

    Private Sub btnNeu_LPU_Click(sender As Object, e As EventArgs) Handles btnNeu_LPU.Click
        PicBar_LPU.Visible = True
        BWLPU.RunWorkerAsync()
    End Sub

    Private Sub BWLPU_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BWLPU.DoWork
        Select Case AppsOffice
            Case "XL_NotInstalled"
                Dim xlAppLPU As Object
                Dim xlWbookLPU As Object
                Dim xlWsheetLPU As Object

                Try
                    xlAppLPU = CreateObject("Ket.Application")
                    xlWbookLPU = xlAppLPU.Workbooks.Open(txtLPU_src.Text)
                    xlWsheetLPU = xlWbookLPU.Worksheets("UID List Of Promotion Utilizati")

                    xlWsheetLPU.UsedRange.UnMerge()
                    xlWsheetLPU.UsedRange.WrapText = False
                    xlWsheetLPU.UsedRange.ColumnWidth = 15
                    xlWsheetLPU.UsedRange.RowHeight = 15

                    xlWsheetLPU.Range("A8").EntireRow.Delete()

                    Dim rg_head_cut1 As Object = xlWsheetLPU.Range("B2")
                    Dim rg_head_paste1 As Object = xlWsheetLPU.Range("A2")
                    rg_head_cut1.Select()
                    rg_head_cut1.Cut(rg_head_paste1)

                    Dim rg_head_cut2 As Object = xlWsheetLPU.Range("B4")
                    Dim rg_head_paste2 As Object = xlWsheetLPU.Range("A3")
                    rg_head_cut2.Select()
                    rg_head_cut2.Cut(rg_head_paste2)

                    xlWsheetLPU.Range("A2").RowHeight = 27

                    Dim paramhead1, paramhead2, paramhead3 As String
                    paramhead1 = xlWsheetLPU.Range("C6").Value & " " & xlWsheetLPU.Range("F6").Value & "; "
                    paramhead1 = paramhead1 & xlWsheetLPU.Range("J6").Value & " " & xlWsheetLPU.Range("M6").Value

                    paramhead2 = xlWsheetLPU.Range("Q6  ").Value & " " & xlWsheetLPU.Range("T6").Value & "; "
                    paramhead2 = paramhead2 & xlWsheetLPU.Range("W6").Value & " " & xlWsheetLPU.Range("Z6").Value

                    paramhead3 = xlWsheetLPU.Range("AD6").Value & " " & xlWsheetLPU.Range("AF6").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetLPU.Range("AL6").Value & " " & xlWsheetLPU.Range("AN6").Value

                    xlWsheetLPU.Range("A5").Value = paramhead1
                    xlWsheetLPU.Range("A5").EntireRow.Font.Name = "Calibri"
                    xlWsheetLPU.Range("A6").Value = paramhead2
                    xlWsheetLPU.Range("A6").EntireRow.Font.Name = "Calibri"
                    xlWsheetLPU.Range("A7").Value = paramhead3
                    xlWsheetLPU.Range("A7").EntireRow.Font.Name = "Calibri"

                    xlWsheetLPU.range("C6:AN6").value = ""

                    Dim xlfunc As Excel.WorksheetFunction
                    xlfunc = xlAppLPU.WorksheetFunction
                    Dim lnCol As Long
                    Dim i As Long
                    Dim rnarea As Object = xlWsheetLPU.UsedRange

                    lnCol = rnarea.Columns.Count
                    For i = lnCol To 1 Step -1
                        If xlfunc.CountA(rnarea.Columns(i)) = 0 Then
                            rnarea.Columns(i).Delete()
                        End If
                    Next

                    xlWsheetLPU.Range("A8:A9").Merge()
                    xlWsheetLPU.Range("A8:A9").HorizontalAlignment = 3
                    xlWsheetLPU.Range("B8:B9").Merge()
                    xlWsheetLPU.Range("B8:B9").HorizontalAlignment = 3
                    xlWsheetLPU.Range("C8:C9").Merge()
                    xlWsheetLPU.Range("C8:C9").HorizontalAlignment = 3
                    xlWsheetLPU.Range("D8:D9").Merge()
                    xlWsheetLPU.Range("D8:D9").HorizontalAlignment = 3
                    xlWsheetLPU.Range("B8:B9").Merge()
                    xlWsheetLPU.Range("B8:B9").HorizontalAlignment = 3
                    xlWsheetLPU.Range("E8:F8").Merge()
                    xlWsheetLPU.Range("E8:F8").HorizontalAlignment = 3
                    xlWsheetLPU.Range("E9:F9").HorizontalAlignment = 3
                    xlWsheetLPU.Range("G8:G9").Merge()
                    xlWsheetLPU.Range("G8:G9").HorizontalAlignment = 3
                    xlWsheetLPU.Range("H8:H9").Merge()
                    xlWsheetLPU.Range("H8:H9").HorizontalAlignment = 3
                    xlWsheetLPU.Range("I8:I9").Merge()
                    xlWsheetLPU.Range("I8:I9").HorizontalAlignment = 3
                    xlWsheetLPU.Range("B8:B9").Merge()
                    xlWsheetLPU.Range("B8:B9").HorizontalAlignment = 3
                    xlWsheetLPU.Range("J8:J9").Merge()
                    xlWsheetLPU.Range("J8:J9").HorizontalAlignment = 3
                    xlWsheetLPU.Range("K8:K9").Merge()
                    xlWsheetLPU.Range("K8:K9").HorizontalAlignment = 3
                    xlWsheetLPU.Range("L8:L9").Merge()
                    xlWsheetLPU.Range("L8:L9").HorizontalAlignment = 3
                    xlWsheetLPU.Range("M8:M9").Merge()
                    xlWsheetLPU.Range("M8:M9").HorizontalAlignment = 3
                    xlWsheetLPU.Range("N8:N9").Merge()
                    xlWsheetLPU.Range("N8:N9").HorizontalAlignment = 3
                    xlWsheetLPU.Range("O8:Q8").Merge()
                    xlWsheetLPU.Range("O8:Q8").HorizontalAlignment = 3
                    xlWsheetLPU.Range("O8:Q8").Merge()
                    xlWsheetLPU.Range("O8:Q8").HorizontalAlignment = 3
                    xlWsheetLPU.Range("R8:R9").Merge()
                    xlWsheetLPU.Range("R8:R9").HorizontalAlignment = 3
                    xlWsheetLPU.Range("S8:S9").Merge()
                    xlWsheetLPU.Range("S8:S9").HorizontalAlignment = 3

                    xlWbookLPU.SaveAs(txtLPU_dest.Text)
                    xlWbookLPU.Close()
                    xlAppLPU.Quit()

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetLPU)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookLPU)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppLPU)

                    xlWsheetLPU = Nothing
                    xlWbookLPU = Nothing
                    xlAppLPU = Nothing

                    MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                Catch ex As Exception
                    MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try

            Case "XL_Installed"
                Dim xlAppLPU As Object
                Dim xlWbookLPU As Object
                Dim xlWsheetLPU As Object

                Try
                    xlAppLPU = CreateObject("Excel.Application")
                    xlWbookLPU = xlAppLPU.Workbooks.Open(txtLPU_src.Text)
                    xlWsheetLPU = xlWbookLPU.Worksheets("UID List Of Promotion Utilizati")

                    xlWsheetLPU.UsedRange.UnMerge()
                    xlWsheetLPU.UsedRange.WrapText = False
                    xlWsheetLPU.UsedRange.ColumnWidth = 15
                    xlWsheetLPU.UsedRange.RowHeight = 15

                    xlWsheetLPU.Range("A8").EntireRow.Delete()

                    Dim rg_head_cut1 As Excel.Range = xlWsheetLPU.Range("B2")
                    Dim rg_head_paste1 As Excel.Range = xlWsheetLPU.Range("A2")
                    rg_head_cut1.Select()
                    rg_head_cut1.Cut(rg_head_paste1)

                    Dim rg_head_cut2 As Excel.Range = xlWsheetLPU.Range("B4")
                    Dim rg_head_paste2 As Excel.Range = xlWsheetLPU.Range("A3")
                    rg_head_cut2.Select()
                    rg_head_cut2.Cut(rg_head_paste2)

                    xlWsheetLPU.Range("A2").RowHeight = 27

                    Dim paramhead1, paramhead2, paramhead3 As String
                    paramhead1 = xlWsheetLPU.Range("C6").Value & " " & xlWsheetLPU.Range("F6").Value & "; "
                    paramhead1 = paramhead1 & xlWsheetLPU.Range("J6").Value & " " & xlWsheetLPU.Range("M6").Value

                    paramhead2 = xlWsheetLPU.Range("Q6  ").Value & " " & xlWsheetLPU.Range("T6").Value & "; "
                    paramhead2 = paramhead2 & xlWsheetLPU.Range("W6").Value & " " & xlWsheetLPU.Range("Z6").Value

                    paramhead3 = xlWsheetLPU.Range("AD6").Value & " " & xlWsheetLPU.Range("AF6").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetLPU.Range("AL6").Value & " " & xlWsheetLPU.Range("AN6").Value

                    xlWsheetLPU.Range("A5").Value = paramhead1
                    xlWsheetLPU.Range("A5").EntireRow.Font.Name = "Calibri"
                    xlWsheetLPU.Range("A6").Value = paramhead2
                    xlWsheetLPU.Range("A6").EntireRow.Font.Name = "Calibri"
                    xlWsheetLPU.Range("A7").Value = paramhead3
                    xlWsheetLPU.Range("A7").EntireRow.Font.Name = "Calibri"

                    xlWsheetLPU.range("C6:AN6").value = ""

                    Dim xlfunc As Excel.WorksheetFunction
                    xlfunc = xlAppLPU.WorksheetFunction
                    Dim lnCol As Long
                    Dim i As Long
                    Dim rnarea As Excel.Range = xlWsheetLPU.UsedRange

                    lnCol = rnarea.Columns.Count
                    For i = lnCol To 1 Step -1
                        If xlfunc.CountA(rnarea.Columns(i)) = 0 Then
                            rnarea.Columns(i).Delete()
                        End If
                    Next

                    xlWsheetLPU.Range("A8:A9").Merge()
                    xlWsheetLPU.Range("A8:A9").HorizontalAlignment = 3
                    xlWsheetLPU.Range("B8:B9").Merge()
                    xlWsheetLPU.Range("B8:B9").HorizontalAlignment = 3
                    xlWsheetLPU.Range("C8:C9").Merge()
                    xlWsheetLPU.Range("C8:C9").HorizontalAlignment = 3
                    xlWsheetLPU.Range("D8:D9").Merge()
                    xlWsheetLPU.Range("D8:D9").HorizontalAlignment = 3
                    xlWsheetLPU.Range("B8:B9").Merge()
                    xlWsheetLPU.Range("B8:B9").HorizontalAlignment = 3
                    xlWsheetLPU.Range("E8:F8").Merge()
                    xlWsheetLPU.Range("E8:F8").HorizontalAlignment = 3
                    xlWsheetLPU.Range("E9:F9").HorizontalAlignment = 3
                    xlWsheetLPU.Range("G8:G9").Merge()
                    xlWsheetLPU.Range("G8:G9").HorizontalAlignment = 3
                    xlWsheetLPU.Range("H8:H9").Merge()
                    xlWsheetLPU.Range("H8:H9").HorizontalAlignment = 3
                    xlWsheetLPU.Range("I8:I9").Merge()
                    xlWsheetLPU.Range("I8:I9").HorizontalAlignment = 3
                    xlWsheetLPU.Range("B8:B9").Merge()
                    xlWsheetLPU.Range("B8:B9").HorizontalAlignment = 3
                    xlWsheetLPU.Range("J8:J9").Merge()
                    xlWsheetLPU.Range("J8:J9").HorizontalAlignment = 3
                    xlWsheetLPU.Range("K8:K9").Merge()
                    xlWsheetLPU.Range("K8:K9").HorizontalAlignment = 3
                    xlWsheetLPU.Range("L8:L9").Merge()
                    xlWsheetLPU.Range("L8:L9").HorizontalAlignment = 3
                    xlWsheetLPU.Range("M8:M9").Merge()
                    xlWsheetLPU.Range("M8:M9").HorizontalAlignment = 3
                    xlWsheetLPU.Range("N8:N9").Merge()
                    xlWsheetLPU.Range("N8:N9").HorizontalAlignment = 3
                    xlWsheetLPU.Range("O8:Q8").Merge()
                    xlWsheetLPU.Range("O8:Q8").HorizontalAlignment = 3
                    xlWsheetLPU.Range("O8:Q8").Merge()
                    xlWsheetLPU.Range("O8:Q8").HorizontalAlignment = 3
                    xlWsheetLPU.Range("R8:R9").Merge()
                    xlWsheetLPU.Range("R8:R9").HorizontalAlignment = 3
                    xlWsheetLPU.Range("S8:S9").Merge()
                    xlWsheetLPU.Range("S8:S9").HorizontalAlignment = 3

                    xlWbookLPU.SaveAs(txtLPU_dest.Text)
                    xlWbookLPU.Close()
                    xlAppLPU.Quit()

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetLPU)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookLPU)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppLPU)

                    xlWsheetLPU = Nothing
                    xlWbookLPU = Nothing
                    xlAppLPU = Nothing

                    MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                Catch ex As Exception
                    MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try

        End Select
    End Sub

    Private Sub BWLPU_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BWLPU.RunWorkerCompleted
        PicBar_LPU.Visible = False
        btnNeu_LPU.Enabled = False
        txtLPU_dest.Text = ""
        txtLPU_src.Text = ""
    End Sub
End Class