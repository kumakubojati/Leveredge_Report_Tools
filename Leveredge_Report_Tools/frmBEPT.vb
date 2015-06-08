Imports Microsoft.Office.Interop
Imports System.Windows.Forms
Imports System.Text
Imports System.Drawing
Imports Microsoft.Win32
Public Class frmBEPT
    Dim AppsOffice As String
    Private Sub frmBEPT_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim TargetKey As RegistryKey
        TargetKey = Registry.ClassesRoot.OpenSubKey("Excel.Application")
        If TargetKey Is Nothing Then
            AppsOffice = "XL_NotInstalled"
        Else
            AppsOffice = "XL_Installed"
            TargetKey.Close()
        End If
    End Sub

    Private Sub btnBrow_BEPT_IC_src_Click(sender As Object, e As EventArgs) Handles btnBrow_BEPT_IC_src.Click
        Dim BEPTpathSrc As String
        If OFD_BEPT_IC.ShowDialog = DialogResult.OK Then
            BEPTpathSrc = OFD_BEPT_IC.FileName()
            txtBEPT_IC_src.Text = BEPTpathSrc
        End If
        If txtBEPT_IC_dest.Text <> "" Then
            btnNeu_BEPT_IC.Enabled = True
        Else
            btnNeu_BEPT_IC.Enabled = False
        End If
    End Sub

    Private Sub btnBrow_BEPT_IC_dest_Click(sender As Object, e As EventArgs) Handles btnBrow_BEPT_IC_dest.Click
        Dim datenow As DateTime = DateTime.Now
        Dim filename As String = "Neutralize_IC_BEPThroughput_" & datenow.ToString("ddMMyyyy_HHmm")
        SFD_BEPT_IC.FileName = filename
        Dim BEPTpath_Dest As String
        If SFD_BEPT_IC.ShowDialog = DialogResult.OK Then
            BEPTpath_Dest = SFD_BEPT_IC.FileName
            txtBEPT_IC_dest.Text = BEPTpath_Dest
        End If
        If txtBEPT_IC_src.Text <> "" Then
            btnNeu_BEPT_IC.Enabled = True
        Else
            btnNeu_BEPT_IC.Enabled = False
        End If
        If txtBEPT_IC_src.Text <> "" Then
            btnNeu_BEPT_IC.Enabled = True
        Else
            btnNeu_BEPT_IC.Enabled = False
        End If
    End Sub

    Private Sub btnNeu_BEPT_IC_Click(sender As Object, e As EventArgs) Handles btnNeu_BEPT_IC.Click
        PicBar_BEPT_IC.Visible = True
        BWBEPT_IC.RunWorkerAsync()
    End Sub

    Private Sub BWBEPT_IC_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BWBEPT_IC.DoWork
        Select Case AppsOffice
            Case "XL_NotInstalled"
                Dim xlAppBEPT_IC As Object
                Dim xlWbookBEPT_IC As Object
                Dim xlWsheetBEPT_IC As Object
                Try
                    xlAppBEPT_IC = CreateObject("Ket.Application")
                    xlWbookBEPT_IC = xlAppBEPT_IC.Workbooks.Open(txtBEPT_IC_src.Text)
                    xlWsheetBEPT_IC = xlWbookBEPT_IC.Worksheets("UID IC BEP Throughput Report")

                    xlWsheetBEPT_IC.UsedRange.UnMerge()
                    xlWsheetBEPT_IC.UsedRange.WrapText = False
                    xlWsheetBEPT_IC.UsedRange.ColumnWidth = 15
                    xlWsheetBEPT_IC.UsedRange.RowHeight = 15

                    Dim rg_head_cut1 As Object = xlWsheetBEPT_IC.Range("C2")
                    Dim rg_head_paste1 As Object = xlWsheetBEPT_IC.Range("A2")
                    rg_head_cut1.Select()
                    rg_head_cut1.Cut(rg_head_paste1)

                    Dim rg_head_cut2 As Object = xlWsheetBEPT_IC.Range("C4")
                    Dim rg_head_paste2 As Object = xlWsheetBEPT_IC.Range("A3")
                    rg_head_cut2.Select()
                    rg_head_cut2.Cut(rg_head_paste2)

                    xlWsheetBEPT_IC.Range("A2").RowHeight = 27

                    Dim paramhead1, paramhead2, paramhead3 As String
                    paramhead1 = xlWsheetBEPT_IC.Range("B6").Value & " " & xlWsheetBEPT_IC.Range("G6").Value & "; "
                    paramhead1 = paramhead1 & xlWsheetBEPT_IC.Range("J6").Value & " " & xlWsheetBEPT_IC.Range("L6").Value

                    paramhead2 = xlWsheetBEPT_IC.Range("R6").Value & " " & xlWsheetBEPT_IC.Range("U6").Value & "; "
                    paramhead2 = paramhead2 & xlWsheetBEPT_IC.Range("W6").Value & " " & xlWsheetBEPT_IC.Range("Z6").Value

                    paramhead3 = xlWsheetBEPT_IC.Range("B8").Value & " " & xlWsheetBEPT_IC.Range("G8").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetBEPT_IC.Range("J8").Value & " " & xlWsheetBEPT_IC.Range("M8").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetBEPT_IC.Range("U8").Value & " " & xlWsheetBEPT_IC.Range("Z8").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetBEPT_IC.Range("AD8").Value & " " & xlWsheetBEPT_IC.Range("AG8").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetBEPT_IC.Range("AK8").Value & " " & xlWsheetBEPT_IC.Range("AO8").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetBEPT_IC.Range("AQ8").Value & " " & xlWsheetBEPT_IC.Range("AT8").Value

                    xlWsheetBEPT_IC.Range("A5").Value = paramhead1
                    xlWsheetBEPT_IC.Range("A5").EntireRow.Font.Name = "Calibri"
                    xlWsheetBEPT_IC.Range("A6").Value = paramhead2
                    xlWsheetBEPT_IC.Range("A6").EntireRow.Font.Name = "Calibri"
                    xlWsheetBEPT_IC.Range("A7").Value = paramhead3
                    xlWsheetBEPT_IC.Range("A7").EntireRow.Font.Name = "Calibri"

                    xlWsheetBEPT_IC.Range("B6:AT8").value = ""

                    Dim xlfunc As Excel.WorksheetFunction
                    xlfunc = xlAppBEPT_IC.WorksheetFunction
                    Dim lnCol As Long
                    Dim i As Long
                    Dim rnarea As Object = xlWsheetBEPT_IC.UsedRange

                    lnCol = rnarea.Columns.Count
                    For i = lnCol To 1 Step -1
                        If xlfunc.CountA(rnarea.Columns(i)) = 0 Then
                            rnarea.Columns(i).Delete()
                        End If
                    Next

                    xlWbookBEPT_IC.SaveAs(txtBEPT_IC_dest.Text)
                    xlWbookBEPT_IC.Close()
                    xlAppBEPT_IC.Quit()

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetBEPT_IC)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookBEPT_IC)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppBEPT_IC)

                    xlWsheetBEPT_IC = Nothing
                    xlWbookBEPT_IC = Nothing
                    xlAppBEPT_IC = Nothing

                    MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                Catch ex As Exception
                    MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try

            Case "XL_Installed"
                Dim xlAppBEPT_IC As Object
                Dim xlWbookBEPT_IC As Object
                Dim xlWsheetBEPT_IC As Object
                Try
                    xlAppBEPT_IC = CreateObject("Excel.Application")
                    xlWbookBEPT_IC = xlAppBEPT_IC.Workbooks.Open(txtBEPT_IC_src.Text)
                    xlWsheetBEPT_IC = xlWbookBEPT_IC.Worksheets("UID IC BEP Throughput Report")

                    xlWsheetBEPT_IC.UsedRange.UnMerge()
                    xlWsheetBEPT_IC.UsedRange.WrapText = False
                    xlWsheetBEPT_IC.UsedRange.ColumnWidth = 15
                    xlWsheetBEPT_IC.UsedRange.RowHeight = 15

                    Dim rg_head_cut1 As Excel.Range = xlWsheetBEPT_IC.Range("C2")
                    Dim rg_head_paste1 As Excel.Range = xlWsheetBEPT_IC.Range("A2")
                    rg_head_cut1.Select()
                    rg_head_cut1.Cut(rg_head_paste1)

                    Dim rg_head_cut2 As Excel.Range = xlWsheetBEPT_IC.Range("C4")
                    Dim rg_head_paste2 As Excel.Range = xlWsheetBEPT_IC.Range("A3")
                    rg_head_cut2.Select()
                    rg_head_cut2.Cut(rg_head_paste2)

                    xlWsheetBEPT_IC.Range("A2").RowHeight = 27

                    Dim paramhead1, paramhead2, paramhead3 As String
                    paramhead1 = xlWsheetBEPT_IC.Range("B6").Value & " " & xlWsheetBEPT_IC.Range("G6").Value & "; "
                    paramhead1 = paramhead1 & xlWsheetBEPT_IC.Range("J6").Value & " " & xlWsheetBEPT_IC.Range("L6").Value

                    paramhead2 = xlWsheetBEPT_IC.Range("R6").Value & " " & xlWsheetBEPT_IC.Range("U6").Value & "; "
                    paramhead2 = paramhead2 & xlWsheetBEPT_IC.Range("W6").Value & " " & xlWsheetBEPT_IC.Range("Z6").Value

                    paramhead3 = xlWsheetBEPT_IC.Range("B8").Value & " " & xlWsheetBEPT_IC.Range("G8").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetBEPT_IC.Range("J8").Value & " " & xlWsheetBEPT_IC.Range("M8").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetBEPT_IC.Range("U8").Value & " " & xlWsheetBEPT_IC.Range("Z8").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetBEPT_IC.Range("AD8").Value & " " & xlWsheetBEPT_IC.Range("AG8").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetBEPT_IC.Range("AK8").Value & " " & xlWsheetBEPT_IC.Range("AO8").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetBEPT_IC.Range("AQ8").Value & " " & xlWsheetBEPT_IC.Range("AT8").Value

                    xlWsheetBEPT_IC.Range("A5").Value = paramhead1
                    xlWsheetBEPT_IC.Range("A5").EntireRow.Font.Name = "Calibri"
                    xlWsheetBEPT_IC.Range("A6").Value = paramhead2
                    xlWsheetBEPT_IC.Range("A6").EntireRow.Font.Name = "Calibri"
                    xlWsheetBEPT_IC.Range("A7").Value = paramhead3
                    xlWsheetBEPT_IC.Range("A7").EntireRow.Font.Name = "Calibri"

                    xlWsheetBEPT_IC.Range("B6:AT8").value = ""

                    Dim xlfunc As Excel.WorksheetFunction
                    xlfunc = xlAppBEPT_IC.WorksheetFunction
                    Dim lnCol As Long
                    Dim i As Long
                    Dim rnarea As Excel.Range = xlWsheetBEPT_IC.UsedRange

                    lnCol = rnarea.Columns.Count
                    For i = lnCol To 1 Step -1
                        If xlfunc.CountA(rnarea.Columns(i)) = 0 Then
                            rnarea.Columns(i).Delete()
                        End If
                    Next

                    xlWbookBEPT_IC.SaveAs(txtBEPT_IC_dest.Text)
                    xlWbookBEPT_IC.Close()
                    xlAppBEPT_IC.Quit()

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetBEPT_IC)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookBEPT_IC)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppBEPT_IC)

                    xlWsheetBEPT_IC = Nothing
                    xlWbookBEPT_IC = Nothing
                    xlAppBEPT_IC = Nothing

                    MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                Catch ex As Exception
                    MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try

        End Select
    End Sub

    Private Sub BWBEPT_IC_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BWBEPT_IC.RunWorkerCompleted
        PicBar_BEPT_IC.Visible = False
        btnNeu_BEPT_IC.Enabled = False
        txtBEPT_IC_dest.Text = ""
        txtBEPT_IC_src.Text = ""
    End Sub
End Class