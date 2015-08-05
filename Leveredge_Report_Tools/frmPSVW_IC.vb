Imports Microsoft.Office.Interop
Imports System.Windows.Forms
Imports System.Text
Imports System.Drawing
Imports Microsoft.Win32
Public Class frmPSVW_IC
    Dim AppsOffice As String
    Private Sub frmPSVW_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim TargetKey As RegistryKey
        TargetKey = Registry.ClassesRoot.OpenSubKey("Excel.Application")
        If TargetKey Is Nothing Then
            AppsOffice = "XL_NotInstalled"
        Else
            AppsOffice = "XL_Installed"
            TargetKey.Close()
        End If
    End Sub

    Private Sub btnBrow_PSVW_IC_src_Click(sender As Object, e As EventArgs) Handles btnBrow_PSVW_IC_src.Click
        Dim PSCWpathSrc As String
        If OFD_PSVW_IC.ShowDialog = DialogResult.OK Then
            PSCWpathSrc = OFD_PSVW_IC.FileName()
            txtPSVW_IC_src.Text = PSCWpathSrc
        End If
        If txtPSVW_IC_dest.Text <> "" Then
            btnNeu_PSVW_IC.Enabled = True
        Else
            btnNeu_PSVW_IC.Enabled = False
        End If
    End Sub

    Private Sub btnBrow_PSVW_IC_dest_Click(sender As Object, e As EventArgs) Handles btnBrow_PSVW_IC_dest.Click
        Dim datenow As DateTime = DateTime.Now
        Dim filename As String = "Neutralize_IC_ProdSalesByVolume_" & datenow.ToString("ddMMyyyy_HHmm")
        SFD_PSVW_IC.FileName = filename
        Dim PSCWpath_Dest As String
        If SFD_PSVW_IC.ShowDialog = DialogResult.OK Then
            PSCWpath_Dest = SFD_PSVW_IC.FileName
            txtPSVW_IC_dest.Text = PSCWpath_Dest
        End If
        If txtPSVW_IC_src.Text <> "" Then
            btnNeu_PSVW_IC.Enabled = True
        Else
            btnNeu_PSVW_IC.Enabled = False
        End If
        If txtPSVW_IC_dest.Text <> "" Then
            btnNeu_PSVW_IC.Enabled = True
        Else
            btnNeu_PSVW_IC.Enabled = False
        End If
    End Sub

    Private Sub btnNeu_PSVW_IC_Click(sender As Object, e As EventArgs) Handles btnNeu_PSVW_IC.Click
        PicBar_PSVW_IC.Visible = True
        BWPSVW_IC.RunWorkerAsync()
    End Sub

    Private Sub BWPSVW_IC_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BWPSVW_IC.DoWork
        Select Case AppsOffice
            Case "XL_NotInstalled"
                Dim xlAppPSVW_IC As Object
                Dim xlWbookPSVW_IC As Object
                Dim xlWsheetPSVW_IC As Object

                Try
                    xlAppPSVW_IC = CreateObject("Ket.Application")
                    xlWbookPSVW_IC = xlAppPSVW_IC.Workbooks.Open(txtPSVW_IC_src.Text)
                    xlWsheetPSVW_IC = xlWbookPSVW_IC.Worksheets("UID IC Product Sales By Volume ")

                    xlWsheetPSVW_IC.UsedRange.UnMerge()
                    xlWsheetPSVW_IC.UsedRange.WrapText = False
                    xlWsheetPSVW_IC.UsedRange.ColumnWidth = 15
                    xlWsheetPSVW_IC.UsedRange.RowHeight = 15

                    Dim rg_head_cut1 As Object = xlWsheetPSVW_IC.Range("C2")
                    Dim rg_head_paste1 As Object = xlWsheetPSVW_IC.Range("A2")
                    rg_head_cut1.Select()
                    rg_head_cut1.Cut(rg_head_paste1)

                    Dim rg_head_cut2 As Object = xlWsheetPSVW_IC.Range("C4")
                    Dim rg_head_paste2 As Object = xlWsheetPSVW_IC.Range("A3")
                    rg_head_cut2.Select()
                    rg_head_cut2.Cut(rg_head_paste2)

                    xlWsheetPSVW_IC.Range("A2").RowHeight = 27

                    Dim paramhead1, paramhead2, paramhead3 As String
                    paramhead1 = xlWsheetPSVW_IC.Range("B6").Value & " " & xlWsheetPSVW_IC.Range("F6").Value & "; "
                    paramhead1 = paramhead1 & xlWsheetPSVW_IC.Range("I6").Value & " " & xlWsheetPSVW_IC.Range("K6").Value

                    paramhead2 = xlWsheetPSVW_IC.Range("O6").Value & " " & xlWsheetPSVW_IC.Range("R6").Value & ";"
                    paramhead2 = paramhead2 & xlWsheetPSVW_IC.Range("V6").Value & " " & xlWsheetPSVW_IC.Range("Y6").Value & "; "
                    paramhead2 = paramhead2 & xlWsheetPSVW_IC.Range("AB6").Value & " " & xlWsheetPSVW_IC.Range("AD6").Value & "; "

                    paramhead3 = xlWsheetPSVW_IC.Range("B8").Value & " " & xlWsheetPSVW_IC.Range("F8").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetPSVW_IC.Range("I8").Value & " " & xlWsheetPSVW_IC.Range("K8").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetPSVW_IC.Range("O8").Value & " " & xlWsheetPSVW_IC.Range("R8").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetPSVW_IC.Range("V8").Value & " " & xlWsheetPSVW_IC.Range("Y8").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetPSVW_IC.Range("AB8").Value & " " & xlWsheetPSVW_IC.Range("AD8").Value

                    xlWsheetPSVW_IC.Range("A5").Value = paramhead1
                    xlWsheetPSVW_IC.Range("A5").EntireRow.Font.Name = "Calibri"
                    xlWsheetPSVW_IC.Range("A6").Value = paramhead2
                    xlWsheetPSVW_IC.Range("A6").EntireRow.Font.Name = "Calibri"
                    xlWsheetPSVW_IC.Range("A7").Value = paramhead3
                    xlWsheetPSVW_IC.Range("A7").EntireRow.Font.Name = "Calibri"

                    xlWsheetPSVW_IC.Range("B6:AD8").value = ""
                    xlWsheetPSVW_IC.Range("A9").EntireRow.Delete()

                    Dim xlfunc As Object
                    xlfunc = xlAppPSVW_IC.WorksheetFunction
                    Dim lnCol As Long
                    Dim i, j As Long
                    Dim rnarea As Object = xlWsheetPSVW_IC.UsedRange

                    lnCol = rnarea.Columns.Count
                    For i = lnCol To 1 Step -1
                        If xlfunc.CountA(rnarea.Columns(i)) = 0 Then
                            rnarea.Columns(i).Delete()
                            j = j + 1
                        End If
                    Next

                    xlWsheetPSVW_IC.Range("K10").Value = "Outlet Name"

                    Dim lastcol As Long
                    Dim x, y As Integer
                    Dim lastval As String = ""
                    Dim rn As Object
                    y = 15
                    Do Until lastval = "Overall Result"
                        lastcol = xlWsheetPSVW_IC.Cells(9, y).End(Excel.XlDirection.xlToRight).Column
                        x = lastcol - 1
                        rn = xlWsheetPSVW_IC.Range(xlWsheetPSVW_IC.cells(9, y), xlWsheetPSVW_IC.cells(9, x))
                        rn.Merge()
                        rn.HorizontalAlignment = Excel.Constants.xlCenter
                        lastval = xlWsheetPSVW_IC.Cells(9, lastcol).Value
                        y = y + 3
                    Loop
                    x = x + 3
                    rn = xlWsheetPSVW_IC.Range(xlWsheetPSVW_IC.Cells(9, lastcol), xlWsheetPSVW_IC.Cells(9, x))
                    rn.Merge()
                    rn.HorizontalAlignment = Excel.Constants.xlCenter

                    xlWbookPSVW_IC.SaveAs(txtPSVW_IC_dest.Text)
                    xlWbookPSVW_IC.Close()
                    xlAppPSVW_IC.Quit()

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetPSVW_IC)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookPSVW_IC)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppPSVW_IC)

                    xlWsheetPSVW_IC = Nothing
                    xlWbookPSVW_IC = Nothing
                    xlAppPSVW_IC = Nothing

                    MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                Catch ex As Exception
                    MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try

            Case "XL_Installed"
                Dim xlAppPSVW_IC As Object
                Dim xlWbookPSVW_IC As Object
                Dim xlWsheetPSVW_IC As Object

                Try
                    xlAppPSVW_IC = CreateObject("Excel.Application")
                    xlWbookPSVW_IC = xlAppPSVW_IC.Workbooks.Open(txtPSVW_IC_src.Text)
                    xlWsheetPSVW_IC = xlWbookPSVW_IC.Worksheets("UID IC Product Sales By Volume ")

                    xlWsheetPSVW_IC.UsedRange.UnMerge()
                    xlWsheetPSVW_IC.UsedRange.WrapText = False
                    xlWsheetPSVW_IC.UsedRange.ColumnWidth = 15
                    xlWsheetPSVW_IC.UsedRange.RowHeight = 15

                    Dim rg_head_cut1 As Object = xlWsheetPSVW_IC.Range("C2")
                    Dim rg_head_paste1 As Object = xlWsheetPSVW_IC.Range("A2")
                    rg_head_cut1.Select()
                    rg_head_cut1.Cut(rg_head_paste1)

                    Dim rg_head_cut2 As Object = xlWsheetPSVW_IC.Range("C4")
                    Dim rg_head_paste2 As Object = xlWsheetPSVW_IC.Range("A3")
                    rg_head_cut2.Select()
                    rg_head_cut2.Cut(rg_head_paste2)

                    xlWsheetPSVW_IC.Range("A2").RowHeight = 27

                    Dim paramhead1, paramhead2, paramhead3 As String
                    paramhead1 = xlWsheetPSVW_IC.Range("B6").Value & " " & xlWsheetPSVW_IC.Range("F6").Value & "; "
                    paramhead1 = paramhead1 & xlWsheetPSVW_IC.Range("I6").Value & " " & xlWsheetPSVW_IC.Range("K6").Value

                    paramhead2 = xlWsheetPSVW_IC.Range("O6").Value & " " & xlWsheetPSVW_IC.Range("R6").Value & ";"
                    paramhead2 = paramhead2 & xlWsheetPSVW_IC.Range("V6").Value & " " & xlWsheetPSVW_IC.Range("Y6").Value & "; "
                    paramhead2 = paramhead2 & xlWsheetPSVW_IC.Range("AB6").Value & " " & xlWsheetPSVW_IC.Range("AD6").Value & "; "

                    paramhead3 = xlWsheetPSVW_IC.Range("B8").Value & " " & xlWsheetPSVW_IC.Range("F8").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetPSVW_IC.Range("I8").Value & " " & xlWsheetPSVW_IC.Range("K8").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetPSVW_IC.Range("O8").Value & " " & xlWsheetPSVW_IC.Range("R8").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetPSVW_IC.Range("V8").Value & " " & xlWsheetPSVW_IC.Range("Y8").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetPSVW_IC.Range("AB8").Value & " " & xlWsheetPSVW_IC.Range("AD8").Value

                    xlWsheetPSVW_IC.Range("A5").Value = paramhead1
                    xlWsheetPSVW_IC.Range("A5").EntireRow.Font.Name = "Calibri"
                    xlWsheetPSVW_IC.Range("A6").Value = paramhead2
                    xlWsheetPSVW_IC.Range("A6").EntireRow.Font.Name = "Calibri"
                    xlWsheetPSVW_IC.Range("A7").Value = paramhead3
                    xlWsheetPSVW_IC.Range("A7").EntireRow.Font.Name = "Calibri"

                    xlWsheetPSVW_IC.Range("B6:AD8").value = ""
                    xlWsheetPSVW_IC.Range("A9").EntireRow.Delete()

                    Dim xlfunc As Object
                    xlfunc = xlAppPSVW_IC.WorksheetFunction
                    Dim lnCol As Long
                    Dim i, j As Long
                    Dim rnarea As Object = xlWsheetPSVW_IC.UsedRange

                    lnCol = rnarea.Columns.Count
                    For i = lnCol To 1 Step -1
                        If xlfunc.CountA(rnarea.Columns(i)) = 0 Then
                            rnarea.Columns(i).Delete()
                            j = j + 1
                        End If
                    Next

                    xlWsheetPSVW_IC.Range("K10").Value = "Outlet Name"

                    Dim lastcol As Long
                    Dim x, y As Integer
                    Dim lastval As String = ""
                    Dim rn As Excel.Range
                    y = 15
                    Do Until lastval = "Overall Result"
                        lastcol = xlWsheetPSVW_IC.Cells(9, y).End(Excel.XlDirection.xlToRight).Column
                        x = lastcol - 1
                        rn = xlWsheetPSVW_IC.Range(xlWsheetPSVW_IC.cells(9, y), xlWsheetPSVW_IC.cells(9, x))
                        rn.Merge()
                        rn.HorizontalAlignment = Excel.Constants.xlCenter
                        lastval = xlWsheetPSVW_IC.Cells(9, lastcol).Value
                        y = y + 2
                    Loop
                    x = x + 2
                    rn = xlWsheetPSVW_IC.Range(xlWsheetPSVW_IC.Cells(9, lastcol), xlWsheetPSVW_IC.Cells(9, x))
                    rn.Merge()
                    rn.HorizontalAlignment = Excel.Constants.xlCenter

                    xlWbookPSVW_IC.SaveAs(txtPSVW_IC_dest.Text)
                    xlWbookPSVW_IC.Close()
                    xlAppPSVW_IC.Quit()

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetPSVW_IC)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookPSVW_IC)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppPSVW_IC)

                    xlWsheetPSVW_IC = Nothing
                    xlWbookPSVW_IC = Nothing
                    xlAppPSVW_IC = Nothing

                    MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                Catch ex As Exception
                    MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
        End Select
    End Sub

    Private Sub BWPSVW_IC_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BWPSVW_IC.RunWorkerCompleted
        PicBar_PSVW_IC.Visible = False
        btnNeu_PSVW_IC.Enabled = False
        txtPSVW_IC_dest.Text = ""
        txtPSVW_IC_src.Text = ""
    End Sub
End Class