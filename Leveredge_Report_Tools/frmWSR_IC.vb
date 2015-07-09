Imports Microsoft.Office.Interop
Imports System.Windows.Forms
Imports System.Text
Imports System.Drawing
Imports Microsoft.Win32
Public Class frmWSR_IC
    Dim AppsOffice As String
    Private Sub frmWSR_IC_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim TargetKey As RegistryKey
        TargetKey = Registry.ClassesRoot.OpenSubKey("Excel.Application")
        If TargetKey Is Nothing Then
            AppsOffice = "XL_NotInstalled"
        Else
            AppsOffice = "XL_Installed"
            TargetKey.Close()
        End If
    End Sub

    Private Sub btnBrow_WSR_IC_src_Click(sender As Object, e As EventArgs) Handles btnBrow_WSR_IC_src.Click
        Dim WSRpathSrc As String
        If OFD_WSR_IC.ShowDialog = DialogResult.OK Then
            WSRpathSrc = OFD_WSR_IC.FileName()
            txtWSR_IC_src.Text = WSRpathSrc
        End If
        If txtWSR_IC_dest.Text <> "" Then
            btnNeu_WSR_IC.Enabled = True
        Else
            btnNeu_WSR_IC.Enabled = False
        End If
    End Sub

    Private Sub btnBrow_WSR_IC_dest_Click(sender As Object, e As EventArgs) Handles btnBrow_WSR_IC_dest.Click
        Dim datenow As DateTime = DateTime.Now
        Dim filename As String = "Neutralize_IC_WaiveStore_" & datenow.ToString("ddMMyyyy_HHmm")
        SFD_WSR_IC.FileName = filename
        Dim WSRpath_Dest As String
        If SFD_WSR_IC.ShowDialog = DialogResult.OK Then
            WSRpath_Dest = SFD_WSR_IC.FileName
            txtWSR_IC_dest.Text = WSRpath_Dest
        End If
        If txtWSR_IC_src.Text <> "" Then
            btnNeu_WSR_IC.Enabled = True
        Else
            btnNeu_WSR_IC.Enabled = False
        End If
        If txtWSR_IC_dest.Text <> "" Then
            btnNeu_WSR_IC.Enabled = True
        Else
            btnNeu_WSR_IC.Enabled = False
        End If
    End Sub

    Private Sub btnNeu_WSR_IC_Click(sender As Object, e As EventArgs) Handles btnNeu_WSR_IC.Click
        PicBar_WSR_IC.Visible = True
        BWWSR_IC.RunWorkerAsync()
    End Sub

    Private Sub BWWSR_IC_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BWWSR_IC.DoWork
        Select Case AppsOffice
            Case "XL_NotInstalled"
                Dim xlAppWSR_IC As Object
                Dim xlWbookWSR_IC As Object
                Dim xlWsheetWSR_IC As Object

                Try
                    xlAppWSR_IC = CreateObject("Ket.Application")
                    xlWbookWSR_IC = xlAppWSR_IC.Workbooks.Open(txtWSR_IC_src.Text)
                    xlWsheetWSR_IC = xlWbookWSR_IC.Worksheets("UID IC Waive Store Report")

                    xlWsheetWSR_IC.UsedRange.UnMerge()
                    xlWsheetWSR_IC.UsedRange.WrapText = False
                    xlWsheetWSR_IC.UsedRange.ColumnWidth = 15
                    xlWsheetWSR_IC.UsedRange.RowHeight = 15

                    Dim rg_head_cut1 As Object = xlWsheetWSR_IC.Range("C2")
                    Dim rg_head_paste1 As Object = xlWsheetWSR_IC.Range("A2")
                    rg_head_cut1.Select()
                    rg_head_cut1.Cut(rg_head_paste1)

                    Dim rg_head_cut2 As Object = xlWsheetWSR_IC.Range("C4")
                    Dim rg_head_paste2 As Object = xlWsheetWSR_IC.Range("A3")
                    rg_head_cut2.Select()
                    rg_head_cut2.Cut(rg_head_paste2)

                    xlWsheetWSR_IC.Range("A2").RowHeight = 27

                    Dim paramhead1, paramhead2, paramhead3 As String
                    paramhead1 = xlWsheetWSR_IC.Range("B6").Value & " " & xlWsheetWSR_IC.Range("H6").Value & "; "
                    paramhead1 = paramhead1 & xlWsheetWSR_IC.Range("J6").Value & " " & xlWsheetWSR_IC.Range("M6").Value

                    paramhead2 = xlWsheetWSR_IC.Range("O6").Value & " " & xlWsheetWSR_IC.Range("T6").Value & ";"
                    paramhead2 = paramhead2 & xlWsheetWSR_IC.Range("Z6").Value & " " & xlWsheetWSR_IC.Range("AC6").Value

                    paramhead3 = xlWsheetWSR_IC.Range("B9").Value & " " & xlWsheetWSR_IC.Range("E8").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetWSR_IC.Range("J9").Value & " " & xlWsheetWSR_IC.Range("M8").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetWSR_IC.Range("P8").Value & " " & xlWsheetWSR_IC.Range("Q8").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetWSR_IC.Range("V8").Value & " " & xlWsheetWSR_IC.Range("W8").Value

                    xlWsheetWSR_IC.Range("A5").Value = paramhead1
                    xlWsheetWSR_IC.Range("A5").EntireRow.Font.Name = "Calibri"
                    xlWsheetWSR_IC.Range("A6").Value = paramhead2
                    xlWsheetWSR_IC.Range("A6").EntireRow.Font.Name = "Calibri"
                    xlWsheetWSR_IC.Range("A7").Value = paramhead3
                    xlWsheetWSR_IC.Range("A7").EntireRow.Font.Name = "Calibri"

                    xlWsheetWSR_IC.Range("B6:AD9").value = ""
                    xlWsheetWSR_IC.Range("A9:A11").EntireRow.Delete()

                    Dim xlfunc As Excel.WorksheetFunction
                    xlfunc = xlAppWSR_IC.WorksheetFunction()
                    Dim lnCol As Long
                    Dim i As Long
                    Dim rnarea As Object = xlWsheetWSR_IC.UsedRange

                    lnCol = rnarea.Columns.Count
                    For i = lnCol To 1 Step -1
                        If xlfunc.CountA(rnarea.Columns(i)) = 0 Then
                            rnarea.Columns(i).Delete()
                        End If
                    Next

                    xlWbookWSR_IC.SaveAs(txtWSR_IC_dest.Text)
                    xlWbookWSR_IC.Close()
                    xlAppWSR_IC.Quit()

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetWSR_IC)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookWSR_IC)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppWSR_IC)

                    xlWsheetWSR_IC = Nothing
                    xlWbookWSR_IC = Nothing
                    xlAppWSR_IC = Nothing

                    MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                Catch ex As Exception
                    MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try

            Case "XL_Installed"
                Dim xlAppWSR_IC As Object
                Dim xlWbookWSR_IC As Object
                Dim xlWsheetWSR_IC As Object

                Try
                    xlAppWSR_IC = CreateObject("Excel.Application")
                    xlWbookWSR_IC = xlAppWSR_IC.Workbooks.Open(txtWSR_IC_src.Text)
                    xlWsheetWSR_IC = xlWbookWSR_IC.Worksheets("UID IC Waive Store Report")

                    xlWsheetWSR_IC.UsedRange.UnMerge()
                    xlWsheetWSR_IC.UsedRange.WrapText = False
                    xlWsheetWSR_IC.UsedRange.ColumnWidth = 15
                    xlWsheetWSR_IC.UsedRange.RowHeight = 15

                    Dim rg_head_cut1 As Excel.Range = xlWsheetWSR_IC.Range("C2")
                    Dim rg_head_paste1 As Excel.Range = xlWsheetWSR_IC.Range("A2")
                    rg_head_cut1.Select()
                    rg_head_cut1.Cut(rg_head_paste1)

                    Dim rg_head_cut2 As Excel.Range = xlWsheetWSR_IC.Range("C4")
                    Dim rg_head_paste2 As Excel.Range = xlWsheetWSR_IC.Range("A3")
                    rg_head_cut2.Select()
                    rg_head_cut2.Cut(rg_head_paste2)

                    xlWsheetWSR_IC.Range("A2").RowHeight = 27

                    Dim paramhead1, paramhead2, paramhead3 As String
                    paramhead1 = xlWsheetWSR_IC.Range("B6").Value & " " & xlWsheetWSR_IC.Range("H6").Value & "; "
                    paramhead1 = paramhead1 & xlWsheetWSR_IC.Range("J6").Value & " " & xlWsheetWSR_IC.Range("M6").Value

                    paramhead2 = xlWsheetWSR_IC.Range("O6").Value & " " & xlWsheetWSR_IC.Range("T6").Value & ";"
                    paramhead2 = paramhead2 & xlWsheetWSR_IC.Range("Z6").Value & " " & xlWsheetWSR_IC.Range("AC6").Value

                    paramhead3 = xlWsheetWSR_IC.Range("B9").Value & " " & xlWsheetWSR_IC.Range("E8").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetWSR_IC.Range("J9").Value & " " & xlWsheetWSR_IC.Range("M8").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetWSR_IC.Range("P8").Value & " " & xlWsheetWSR_IC.Range("Q8").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetWSR_IC.Range("V8").Value & " " & xlWsheetWSR_IC.Range("W8").Value

                    xlWsheetWSR_IC.Range("A5").Value = paramhead1
                    xlWsheetWSR_IC.Range("A5").EntireRow.Font.Name = "Calibri"
                    xlWsheetWSR_IC.Range("A6").Value = paramhead2
                    xlWsheetWSR_IC.Range("A6").EntireRow.Font.Name = "Calibri"
                    xlWsheetWSR_IC.Range("A7").Value = paramhead3
                    xlWsheetWSR_IC.Range("A7").EntireRow.Font.Name = "Calibri"

                    xlWsheetWSR_IC.Range("B6:AD9").value = ""
                    xlWsheetWSR_IC.Range("A9:A11").EntireRow.Delete()

                    Dim xlfunc As Excel.WorksheetFunction
                    xlfunc = xlAppWSR_IC.WorksheetFunction()
                    Dim lnCol As Long
                    Dim i As Long
                    Dim rnarea As Excel.Range = xlWsheetWSR_IC.UsedRange

                    lnCol = rnarea.Columns.Count
                    For i = lnCol To 1 Step -1
                        If xlfunc.CountA(rnarea.Columns(i)) = 0 Then
                            rnarea.Columns(i).Delete()
                        End If
                    Next

                    xlWbookWSR_IC.SaveAs(txtWSR_IC_dest.Text)
                    xlWbookWSR_IC.Close()
                    xlAppWSR_IC.Quit()

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetWSR_IC)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookWSR_IC)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppWSR_IC)

                    xlWsheetWSR_IC = Nothing
                    xlWbookWSR_IC = Nothing
                    xlAppWSR_IC = Nothing

                    MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                Catch ex As Exception
                    MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
        End Select
    End Sub

    Private Sub BWWSR_IC_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BWWSR_IC.RunWorkerCompleted
        PicBar_WSR_IC.Visible = False
        btnNeu_WSR_IC.Enabled = False
        txtWSR_IC_dest.Text = ""
        txtWSR_IC_src.Text = ""
    End Sub
End Class