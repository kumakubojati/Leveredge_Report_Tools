Imports Microsoft.Office.Interop
Imports System.Windows.Forms
Imports System.Text
Imports System.Drawing
Imports Microsoft.Win32

Public Class frmTCR_IC
    Dim AppsOffice As String
    Private Sub frmTCR_IC_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim TargetKey As RegistryKey
        TargetKey = Registry.ClassesRoot.OpenSubKey("Excel.Application")
        If TargetKey Is Nothing Then
            AppsOffice = "XL_NotInstalled"
        Else
            AppsOffice = "XL_Installed"
            TargetKey.Close()
        End If
    End Sub

    Private Sub btnBrow_TCR_IC_src_Click(sender As Object, e As EventArgs) Handles btnBrow_TCR_IC_src.Click
        Dim TCRpathSrc As String
        If OFD_TCR_IC.ShowDialog = DialogResult.OK Then
            TCRpathSrc = OFD_TCR_IC.FileName()
            txtTCR_IC_src.Text = TCRpathSrc
        End If
        If txtTCR_IC_dest.Text <> "" Then
            btnNeu_TCR_IC.Enabled = True
        Else
            btnNeu_TCR_IC.Enabled = False
        End If
    End Sub

    Private Sub btnBrow_TCR_IC_dest_Click(sender As Object, e As EventArgs) Handles btnBrow_TCR_IC_dest.Click
        Dim datenow As DateTime = DateTime.Now
        Dim filename As String = "Neutralize_IC_ThroughPutCab_" & datenow.ToString("ddMMyyyy_HHmm")
        SFD_TCR_IC.FileName = filename
        Dim TCRpath_Dest As String
        If SFD_TCR_IC.ShowDialog = DialogResult.OK Then
            TCRpath_Dest = SFD_TCR_IC.FileName
            txtTCR_IC_dest.Text = TCRpath_Dest
        End If
        If txtTCR_IC_src.Text <> "" Then
            btnNeu_TCR_IC.Enabled = True
        Else
            btnNeu_TCR_IC.Enabled = False
        End If
        If txtTCR_IC_dest.Text <> "" Then
            btnNeu_TCR_IC.Enabled = True
        Else
            btnNeu_TCR_IC.Enabled = False
        End If
    End Sub

    Private Sub btnNeu_TCR_IC_Click(sender As Object, e As EventArgs) Handles btnNeu_TCR_IC.Click
        PicBar_TCR_IC.Visible = True
        BWTCR_IC.RunWorkerAsync()
    End Sub

    Private Sub BWTCR_IC_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BWTCR_IC.DoWork
        Select Case AppsOffice
            Case "XL_NotInstalled"
                Dim xlAppTCR_IC As Object
                Dim xlWbookTCR_IC As Object
                Dim xlWsheetTCR_IC As Object

                Try
                    xlAppTCR_IC = CreateObject("Ket.Application")
                    xlWbookTCR_IC = xlAppTCR_IC.Workbooks.Open(txtTCR_IC_src.Text)
                    xlWsheetTCR_IC = xlWbookTCR_IC.Worksheets("UID IC ThroughPut Cabinet Repor")

                    xlWsheetTCR_IC.UsedRange.UnMerge()
                    xlWsheetTCR_IC.UsedRange.WrapText = False
                    xlWsheetTCR_IC.UsedRange.ColumnWidth = 15
                    xlWsheetTCR_IC.UsedRange.RowHeight = 15

                    Dim rg_head_cut1 As Object = xlWsheetTCR_IC.Range("B2")
                    Dim rg_head_paste1 As Object = xlWsheetTCR_IC.Range("A2")
                    rg_head_cut1.Select()
                    rg_head_cut1.Cut(rg_head_paste1)

                    Dim rg_head_cut2 As Object = xlWsheetTCR_IC.Range("B4")
                    Dim rg_head_paste2 As Object = xlWsheetTCR_IC.Range("A3")
                    rg_head_cut2.Select()
                    rg_head_cut2.Cut(rg_head_paste2)

                    xlWsheetTCR_IC.Range("A2").RowHeight = 27

                    Dim paramhead1, paramhead2, paramhead3 As String
                    paramhead1 = xlWsheetTCR_IC.Range("B6").Value & " " & xlWsheetTCR_IC.Range("E6").Value & "; "
                    paramhead1 = paramhead1 & xlWsheetTCR_IC.Range("H6").Value & " " & xlWsheetTCR_IC.Range("J6").Value

                    paramhead2 = xlWsheetTCR_IC.Range("O6").Value & " " & xlWsheetTCR_IC.Range("R6").Value & ";"
                    paramhead2 = paramhead2 & xlWsheetTCR_IC.Range("V6").Value & " " & xlWsheetTCR_IC.Range("Y6").Value & "; "
                    paramhead2 = paramhead2 & xlWsheetTCR_IC.Range("AB6").Value & " " & xlWsheetTCR_IC.Range("AE6").Value


                    paramhead3 = xlWsheetTCR_IC.Range("B8").Value & " " & xlWsheetTCR_IC.Range("E8").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetTCR_IC.Range("H8").Value & " " & xlWsheetTCR_IC.Range("J8").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetTCR_IC.Range("O8").Value & " " & xlWsheetTCR_IC.Range("R8").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetTCR_IC.Range("V8").Value & " " & xlWsheetTCR_IC.Range("Y8").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetTCR_IC.Range("AB8").Value & " " & xlWsheetTCR_IC.Range("AE8").Value

                    xlWsheetTCR_IC.Range("A5").Value = paramhead1
                    xlWsheetTCR_IC.Range("A5").EntireRow.Font.Name = "Calibri"
                    xlWsheetTCR_IC.Range("A6").Value = paramhead2
                    xlWsheetTCR_IC.Range("A6").EntireRow.Font.Name = "Calibri"
                    xlWsheetTCR_IC.Range("A7").Value = paramhead3
                    xlWsheetTCR_IC.Range("A7").EntireRow.Font.Name = "Calibri"

                    xlWsheetTCR_IC.Range("B6:AF8").value = ""
                    xlWsheetTCR_IC.Range("A9").EntireRow.Delete()

                    Dim xlfunc As Excel.WorksheetFunction
                    xlfunc = xlAppTCR_IC.WorksheetFunction()
                    Dim lnCol As Long
                    Dim i As Long
                    Dim rnarea As Object = xlWsheetTCR_IC.UsedRange

                    lnCol = rnarea.Columns.Count
                    For i = lnCol To 1 Step -1
                        If xlfunc.CountA(rnarea.Columns(i)) = 0 Then
                            rnarea.Columns(i).Delete()
                        End If
                    Next

                    xlWbookTCR_IC.SaveAs(txtTCR_IC_dest.Text)
                    xlWbookTCR_IC.Close()
                    xlAppTCR_IC.Quit()

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetTCR_IC)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookTCR_IC)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppTCR_IC)

                    xlWsheetTCR_IC = Nothing
                    xlWbookTCR_IC = Nothing
                    xlAppTCR_IC = Nothing

                    MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                Catch ex As Exception
                    MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try

            Case "XL_Installed"
                Dim xlAppTCR_IC As Object
                Dim xlWbookTCR_IC As Object
                Dim xlWsheetTCR_IC As Object

                Try
                    xlAppTCR_IC = CreateObject("Excel.Application")
                    xlWbookTCR_IC = xlAppTCR_IC.Workbooks.Open(txtTCR_IC_src.Text)
                    xlWsheetTCR_IC = xlWbookTCR_IC.Worksheets("UID IC ThroughPut Cabinet Repor")

                    xlWsheetTCR_IC.UsedRange.UnMerge()
                    xlWsheetTCR_IC.UsedRange.WrapText = False
                    xlWsheetTCR_IC.UsedRange.ColumnWidth = 15
                    xlWsheetTCR_IC.UsedRange.RowHeight = 15

                    Dim rg_head_cut1 As Excel.Range = xlWsheetTCR_IC.Range("B2")
                    Dim rg_head_paste1 As Excel.Range = xlWsheetTCR_IC.Range("A2")
                    rg_head_cut1.Select()
                    rg_head_cut1.Cut(rg_head_paste1)

                    Dim rg_head_cut2 As Excel.Range = xlWsheetTCR_IC.Range("B4")
                    Dim rg_head_paste2 As Excel.Range = xlWsheetTCR_IC.Range("A3")
                    rg_head_cut2.Select()
                    rg_head_cut2.Cut(rg_head_paste2)

                    xlWsheetTCR_IC.Range("A2").RowHeight = 27

                    Dim paramhead1, paramhead2, paramhead3 As String
                    paramhead1 = xlWsheetTCR_IC.Range("B6").Value & " " & xlWsheetTCR_IC.Range("E6").Value & "; "
                    paramhead1 = paramhead1 & xlWsheetTCR_IC.Range("H6").Value & " " & xlWsheetTCR_IC.Range("J6").Value

                    paramhead2 = xlWsheetTCR_IC.Range("O6").Value & " " & xlWsheetTCR_IC.Range("R6").Value & ";"
                    paramhead2 = paramhead2 & xlWsheetTCR_IC.Range("V6").Value & " " & xlWsheetTCR_IC.Range("Y6").Value & "; "
                    paramhead2 = paramhead2 & xlWsheetTCR_IC.Range("AB6").Value & " " & xlWsheetTCR_IC.Range("AE6").Value


                    paramhead3 = xlWsheetTCR_IC.Range("B8").Value & " " & xlWsheetTCR_IC.Range("E8").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetTCR_IC.Range("H8").Value & " " & xlWsheetTCR_IC.Range("J8").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetTCR_IC.Range("O8").Value & " " & xlWsheetTCR_IC.Range("R8").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetTCR_IC.Range("V8").Value & " " & xlWsheetTCR_IC.Range("Y8").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetTCR_IC.Range("AB8").Value & " " & xlWsheetTCR_IC.Range("AE8").Value

                    xlWsheetTCR_IC.Range("A5").Value = paramhead1
                    xlWsheetTCR_IC.Range("A5").EntireRow.Font.Name = "Calibri"
                    xlWsheetTCR_IC.Range("A6").Value = paramhead2
                    xlWsheetTCR_IC.Range("A6").EntireRow.Font.Name = "Calibri"
                    xlWsheetTCR_IC.Range("A7").Value = paramhead3
                    xlWsheetTCR_IC.Range("A7").EntireRow.Font.Name = "Calibri"

                    xlWsheetTCR_IC.Range("B6:AF8").value = ""
                    xlWsheetTCR_IC.Range("A9").EntireRow.Delete()

                    Dim xlfunc As Excel.WorksheetFunction
                    xlfunc = xlAppTCR_IC.WorksheetFunction()
                    Dim lnCol As Long
                    Dim i As Long
                    Dim rnarea As Excel.Range = xlWsheetTCR_IC.UsedRange

                    lnCol = rnarea.Columns.Count
                    For i = lnCol To 1 Step -1
                        If xlfunc.CountA(rnarea.Columns(i)) = 0 Then
                            rnarea.Columns(i).Delete()
                        End If
                    Next

                    xlWbookTCR_IC.SaveAs(txtTCR_IC_dest.Text)
                    xlWbookTCR_IC.Close()
                    xlAppTCR_IC.Quit()

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetTCR_IC)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookTCR_IC)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppTCR_IC)

                    xlWsheetTCR_IC = Nothing
                    xlWbookTCR_IC = Nothing
                    xlAppTCR_IC = Nothing

                    MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                Catch ex As Exception
                    MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
        End Select
    End Sub

    Private Sub BWTCR_IC_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BWTCR_IC.RunWorkerCompleted
        PicBar_TCR_IC.Visible = False
        btnNeu_TCR_IC.Enabled = False
        txtTCR_IC_dest.Text = ""
        txtTCR_IC_src.Text = ""
    End Sub
End Class