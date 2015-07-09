Imports Microsoft.Office.Interop
Imports System.Windows.Forms
Imports System.Text
Imports System.Drawing
Imports Microsoft.Win32
Public Class frmPSCW_IC
    Dim AppsOffice As String
    Private Sub frmPSCW_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim TargetKey As RegistryKey
        TargetKey = Registry.ClassesRoot.OpenSubKey("Excel.Application")
        If TargetKey Is Nothing Then
            AppsOffice = "XL_NotInstalled"
        Else
            AppsOffice = "XL_Installed"
            TargetKey.Close()
        End If
    End Sub

    Private Sub btnBrow_PSCW_IC_src_Click(sender As Object, e As EventArgs) Handles btnBrow_PSCW_IC_src.Click
        Dim PSCWpathSrc As String
        If OFD_PSCW_IC.ShowDialog = DialogResult.OK Then
            PSCWpathSrc = OFD_PSCW_IC.FileName()
            txtPSCW_IC_src.Text = PSCWpathSrc
        End If
        If txtPSCW_IC_dest.Text <> "" Then
            btnNeu_PSCW_IC.Enabled = True
        Else
            btnNeu_PSCW_IC.Enabled = False
        End If
    End Sub

    Private Sub btnBrow_PSCW_IC_dest_Click(sender As Object, e As EventArgs) Handles btnBrow_PSCW_IC_dest.Click
        Dim datenow As DateTime = DateTime.Now
        Dim filename As String = "Neutralize_IC_ProdSalesByCase_" & datenow.ToString("ddMMyyyy_HHmm")
        SFD_PSCW_IC.FileName = filename
        Dim PSCWpath_Dest As String
        If SFD_PSCW_IC.ShowDialog = DialogResult.OK Then
            PSCWpath_Dest = SFD_PSCW_IC.FileName
            txtPSCW_IC_dest.Text = PSCWpath_Dest
        End If
        If txtPSCW_IC_src.Text <> "" Then
            btnNeu_PSCW_IC.Enabled = True
        Else
            btnNeu_PSCW_IC.Enabled = False
        End If
        If txtPSCW_IC_dest.Text <> "" Then
            btnNeu_PSCW_IC.Enabled = True
        Else
            btnNeu_PSCW_IC.Enabled = False
        End If
    End Sub

    Private Sub btnNeu_PSCW_IC_Click(sender As Object, e As EventArgs) Handles btnNeu_PSCW_IC.Click
        PicBar_PSCW_IC.Visible = True
        BWPSCW_IC.RunWorkerAsync()
    End Sub

    Private Sub BWPSCW_IC_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BWPSCW_IC.DoWork
        Select Case AppsOffice
            Case "XL_NotInstalled"
                Dim xlAppPSCW_IC As Object
                Dim xlWbookPSCW_IC As Object
                Dim xlWsheetPSCW_IC As Object

                Try
                    xlAppPSCW_IC = CreateObject("Ket.Application")
                    xlWbookPSCW_IC = xlAppPSCW_IC.Workbooks.Open(txtPSCW_IC_src.Text)
                    xlWsheetPSCW_IC = xlWbookPSCW_IC.Worksheets("UID IC Product Sales By Case We")

                    xlWsheetPSCW_IC.UsedRange.UnMerge()
                    xlWsheetPSCW_IC.UsedRange.WrapText = False
                    xlWsheetPSCW_IC.UsedRange.ColumnWidth = 15
                    xlWsheetPSCW_IC.UsedRange.RowHeight = 15

                    Dim rg_head_cut1 As Object = xlWsheetPSCW_IC.Range("B2")
                    Dim rg_head_paste1 As Object = xlWsheetPSCW_IC.Range("A2")
                    rg_head_cut1.Select()
                    rg_head_cut1.Cut(rg_head_paste1)

                    Dim rg_head_cut2 As Object = xlWsheetPSCW_IC.Range("B4")
                    Dim rg_head_paste2 As Object = xlWsheetPSCW_IC.Range("A3")
                    rg_head_cut2.Select()
                    rg_head_cut2.Cut(rg_head_paste2)

                    xlWsheetPSCW_IC.Range("A2").RowHeight = 27

                    Dim paramhead1, paramhead2, paramhead3 As String
                    paramhead1 = xlWsheetPSCW_IC.Range("B6").Value & " " & xlWsheetPSCW_IC.Range("D6").Value & "; "
                    paramhead1 = paramhead1 & xlWsheetPSCW_IC.Range("G6").Value & " " & xlWsheetPSCW_IC.Range("I6").Value

                    paramhead2 = xlWsheetPSCW_IC.Range("M6").Value & " " & xlWsheetPSCW_IC.Range("P6").Value & ";"
                    paramhead2 = paramhead2 & xlWsheetPSCW_IC.Range("T6").Value & " " & xlWsheetPSCW_IC.Range("V6").Value & "; "

                    paramhead3 = xlWsheetPSCW_IC.Range("B8").Value & " " & xlWsheetPSCW_IC.Range("D8").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetPSCW_IC.Range("G8").Value & " " & xlWsheetPSCW_IC.Range("I8").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetPSCW_IC.Range("M8").Value & " " & xlWsheetPSCW_IC.Range("P8").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetPSCW_IC.Range("T8").Value & " " & xlWsheetPSCW_IC.Range("V8").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetPSCW_IC.Range("Z8").Value & " " & xlWsheetPSCW_IC.Range("AB8").Value

                    xlWsheetPSCW_IC.Range("A5").Value = paramhead1
                    xlWsheetPSCW_IC.Range("A5").EntireRow.Font.Name = "Calibri"
                    xlWsheetPSCW_IC.Range("A6").Value = paramhead2
                    xlWsheetPSCW_IC.Range("A6").EntireRow.Font.Name = "Calibri"
                    xlWsheetPSCW_IC.Range("A7").Value = paramhead3
                    xlWsheetPSCW_IC.Range("A7").EntireRow.Font.Name = "Calibri"

                    xlWsheetPSCW_IC.Range("B6:AB8").value = ""
                    xlWsheetPSCW_IC.Range("A9").EntireRow.Delete()

                    Dim xlfunc As Object
                    xlfunc = xlAppPSCW_IC.WorksheetFunction
                    Dim lnCol As Long
                    Dim i, j As Long
                    Dim rnarea As Object = xlWsheetPSCW_IC.UsedRange

                    lnCol = rnarea.Columns.Count
                    For i = lnCol To 1 Step -1
                        If xlfunc.CountA(rnarea.Columns(i)) = 0 Then
                            rnarea.Columns(i).Delete()
                            j = j + 1
                        End If
                    Next

                    xlWsheetPSCW_IC.Range("K10").Value = "Outlet Name"

                    Dim lastcol As Long
                    Dim x, y As Integer
                    Dim lastval As String = ""
                    Dim rn As Object
                    y = 15
                    Do Until lastval = "Overall Result"
                        lastcol = xlWsheetPSCW_IC.Cells(9, y).End(Excel.XlDirection.xlToRight).Column
                        x = lastcol - 1
                        rn = xlWsheetPSCW_IC.Range(xlWsheetPSCW_IC.cells(9, y), xlWsheetPSCW_IC.cells(9, x))
                        rn.Merge()
                        rn.HorizontalAlignment = Excel.Constants.xlCenter
                        lastval = xlWsheetPSCW_IC.Cells(9, lastcol).Value
                        y = y + 3
                    Loop
                    x = x + 3
                    rn = xlWsheetPSCW_IC.Range(xlWsheetPSCW_IC.Cells(9, lastcol), xlWsheetPSCW_IC.Cells(9, x))
                    rn.Merge()
                    rn.HorizontalAlignment = Excel.Constants.xlCenter

                    xlWbookPSCW_IC.SaveAs(txtPSCW_IC_dest.Text)
                    xlWbookPSCW_IC.Close()
                    xlAppPSCW_IC.Quit()

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetPSCW_IC)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookPSCW_IC)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppPSCW_IC)

                    xlWsheetPSCW_IC = Nothing
                    xlWbookPSCW_IC = Nothing
                    xlAppPSCW_IC = Nothing

                    MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                Catch ex As Exception
                    MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try

            Case "XL_Installed"
                Dim xlAppPSCW_IC As Object
                Dim xlWbookPSCW_IC As Object
                Dim xlWsheetPSCW_IC As Object

                Try
                    xlAppPSCW_IC = CreateObject("Excel.Application")
                    xlWbookPSCW_IC = xlAppPSCW_IC.Workbooks.Open(txtPSCW_IC_src.Text)
                    xlWsheetPSCW_IC = xlWbookPSCW_IC.Worksheets("UID IC Product Sales By Case We")

                    xlWsheetPSCW_IC.UsedRange.UnMerge()
                    xlWsheetPSCW_IC.UsedRange.WrapText = False
                    xlWsheetPSCW_IC.UsedRange.ColumnWidth = 15
                    xlWsheetPSCW_IC.UsedRange.RowHeight = 15

                    Dim rg_head_cut1 As Excel.Range = xlWsheetPSCW_IC.Range("B2")
                    Dim rg_head_paste1 As Excel.Range = xlWsheetPSCW_IC.Range("A2")
                    rg_head_cut1.Select()
                    rg_head_cut1.Cut(rg_head_paste1)

                    Dim rg_head_cut2 As Excel.Range = xlWsheetPSCW_IC.Range("B4")
                    Dim rg_head_paste2 As Excel.Range = xlWsheetPSCW_IC.Range("A3")
                    rg_head_cut2.Select()
                    rg_head_cut2.Cut(rg_head_paste2)

                    xlWsheetPSCW_IC.Range("A2").RowHeight = 27

                    Dim paramhead1, paramhead2, paramhead3 As String
                    paramhead1 = xlWsheetPSCW_IC.Range("B6").Value & " " & xlWsheetPSCW_IC.Range("D6").Value & "; "
                    paramhead1 = paramhead1 & xlWsheetPSCW_IC.Range("G6").Value & " " & xlWsheetPSCW_IC.Range("I6").Value

                    paramhead2 = xlWsheetPSCW_IC.Range("M6").Value & " " & xlWsheetPSCW_IC.Range("P6").Value & ";"
                    paramhead2 = paramhead2 & xlWsheetPSCW_IC.Range("T6").Value & " " & xlWsheetPSCW_IC.Range("V6").Value & "; "

                    paramhead3 = xlWsheetPSCW_IC.Range("B8").Value & " " & xlWsheetPSCW_IC.Range("D8").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetPSCW_IC.Range("G8").Value & " " & xlWsheetPSCW_IC.Range("I8").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetPSCW_IC.Range("M8").Value & " " & xlWsheetPSCW_IC.Range("P8").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetPSCW_IC.Range("T8").Value & " " & xlWsheetPSCW_IC.Range("V8").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetPSCW_IC.Range("Z8").Value & " " & xlWsheetPSCW_IC.Range("AB8").Value

                    xlWsheetPSCW_IC.Range("A5").Value = paramhead1
                    xlWsheetPSCW_IC.Range("A5").EntireRow.Font.Name = "Calibri"
                    xlWsheetPSCW_IC.Range("A6").Value = paramhead2
                    xlWsheetPSCW_IC.Range("A6").EntireRow.Font.Name = "Calibri"
                    xlWsheetPSCW_IC.Range("A7").Value = paramhead3
                    xlWsheetPSCW_IC.Range("A7").EntireRow.Font.Name = "Calibri"

                    xlWsheetPSCW_IC.Range("B6:AB8").value = ""
                    xlWsheetPSCW_IC.Range("A9").EntireRow.Delete()

                    Dim xlfunc As Object
                    xlfunc = xlAppPSCW_IC.WorksheetFunction
                    Dim lnCol As Long
                    Dim i, j As Long
                    Dim rnarea As Object = xlWsheetPSCW_IC.UsedRange

                    lnCol = rnarea.Columns.Count
                    For i = lnCol To 1 Step -1
                        If xlfunc.CountA(rnarea.Columns(i)) = 0 Then
                            rnarea.Columns(i).Delete()
                            j = j + 1
                        End If
                    Next

                    xlWsheetPSCW_IC.Range("K10").Value = "Outlet Name"

                    Dim lastcol As Long
                    Dim x, y As Integer
                    Dim lastval As String = ""
                    Dim rn As Excel.Range
                    y = 15
                    Do Until lastval = "Overall Result"
                        lastcol = xlWsheetPSCW_IC.Cells(9, y).End(Excel.XlDirection.xlToRight).Column
                        x = lastcol - 1
                        rn = xlWsheetPSCW_IC.Range(xlWsheetPSCW_IC.cells(9, y), xlWsheetPSCW_IC.cells(9, x))
                        rn.Merge()
                        rn.HorizontalAlignment = Excel.Constants.xlCenter
                        lastval = xlWsheetPSCW_IC.Cells(9, lastcol).Value
                        y = y + 3
                    Loop
                    x = x + 3
                    rn = xlWsheetPSCW_IC.Range(xlWsheetPSCW_IC.Cells(9, lastcol), xlWsheetPSCW_IC.Cells(9, x))
                    rn.Merge()
                    rn.HorizontalAlignment = Excel.Constants.xlCenter

                    xlWbookPSCW_IC.SaveAs(txtPSCW_IC_dest.Text)
                    xlWbookPSCW_IC.Close()
                    xlAppPSCW_IC.Quit()

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetPSCW_IC)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookPSCW_IC)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppPSCW_IC)

                    xlWsheetPSCW_IC = Nothing
                    xlWbookPSCW_IC = Nothing
                    xlAppPSCW_IC = Nothing

                    MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                Catch ex As Exception
                    MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
        End Select
    End Sub

    Private Sub BWPSCW_IC_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BWPSCW_IC.RunWorkerCompleted
        PicBar_PSCW_IC.Visible = False
        btnNeu_PSCW_IC.Enabled = False
        txtPSCW_IC_dest.Text = ""
        txtPSCW_IC_src.Text = ""
    End Sub
End Class