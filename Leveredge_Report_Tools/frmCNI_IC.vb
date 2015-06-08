Imports Microsoft.Office.Interop
Imports System.Windows.Forms
Imports System.Text
Imports System.Drawing
Imports Microsoft.Win32
Public Class frmCNI_IC
    Dim AppsOffice As String
    Private Sub frmCNI_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim TargetKey As RegistryKey
        TargetKey = Registry.ClassesRoot.OpenSubKey("Excel.Application")
        If TargetKey Is Nothing Then
            AppsOffice = "XL_NotInstalled"
        Else
            AppsOffice = "XL_Installed"
            TargetKey.Close()
        End If
    End Sub
    Dim CNI_Type As String

    Private Sub RBCNI_Detail_CheckedChanged(sender As Object, e As EventArgs) Handles RBCNI_Detail.CheckedChanged
        CNI_Type = "DTL"
    End Sub

    Private Sub RBCNI_Rekap_CheckedChanged(sender As Object, e As EventArgs) Handles RBCNI_Rekap.CheckedChanged
        CNI_Type = "SUM"
    End Sub

    Private Sub btnBrow_CNI_IC_src_Click(sender As Object, e As EventArgs) Handles btnBrow_CNI_IC_src.Click
        Dim CNIpathSrc As String
        If OFD_CNI_IC.ShowDialog = DialogResult.OK Then
            CNIpathSrc = OFD_CNI_IC.FileName()
            txtCNI_IC_src.Text = CNIpathSrc
        End If
        If txtCNI_IC_dest.Text <> "" Then
            btnNeu_CNI_IC.Enabled = True
        Else
            btnNeu_CNI_IC.Enabled = False
        End If
    End Sub

    Private Sub btnBrow_CNI_IC_dest_Click(sender As Object, e As EventArgs) Handles btnBrow_CNI_IC_dest.Click
        Dim datenow As DateTime = DateTime.Now
        Dim filename As String = "Neutralize_IC_CabinetNetInc_" & CNI_Type & "_" & datenow.ToString("ddMMyyyy_HHmm")
        SFD_CNI_IC.FileName = filename
        Dim CNIpath_Dest As String
        If SFD_CNI_IC.ShowDialog = DialogResult.OK Then
            CNIpath_Dest = SFD_CNI_IC.FileName
            txtCNI_IC_dest.Text = CNIpath_Dest
        End If
        If txtCNI_IC_src.Text <> "" Then
            btnNeu_CNI_IC.Enabled = True
        Else
            btnNeu_CNI_IC.Enabled = False
        End If
        If txtCNI_IC_dest.Text <> "" Then
            btnNeu_CNI_IC.Enabled = True
        Else
            btnNeu_CNI_IC.Enabled = False
        End If
    End Sub

    Private Sub btnNeu_CNI_IC_Click(sender As Object, e As EventArgs) Handles btnNeu_CNI_IC.Click
        PicBar_CNI_IC.Visible = True
        BWCNI_IC.RunWorkerAsync()
    End Sub

    Private Sub BWCNI_IC_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BWCNI_IC.DoWork
        Select Case AppsOffice
            Case "XL_NotInstalled"
                Dim xlAppCNI_IC As Object
                Dim xlWbookCNI_IC As Object
                Dim xlWsheetCNI_IC As Object
                Select Case CNI_Type
                    Case "DTL"
                        Try
                            xlAppCNI_IC = CreateObject("Ket.Application")
                            xlWbookCNI_IC = xlAppCNI_IC.Workbooks.Open(txtCNI_IC_src.Text)
                            xlWsheetCNI_IC = xlWbookCNI_IC.Worksheets("UID IC Cabinet Net Increase Det")

                            xlWsheetCNI_IC.UsedRange.UnMerge()
                            xlWsheetCNI_IC.UsedRange.WrapText = False
                            xlWsheetCNI_IC.UsedRange.ColumnWidth = 15
                            xlWsheetCNI_IC.UsedRange.RowHeight = 15

                            Dim rg_head_cut1 As Object = xlWsheetCNI_IC.Range("B2")
                            Dim rg_head_paste1 As Object = xlWsheetCNI_IC.Range("A2")
                            rg_head_cut1.Select()
                            rg_head_cut1.Cut(rg_head_paste1)

                            Dim rg_head_cut2 As Object = xlWsheetCNI_IC.Range("B4")
                            Dim rg_head_paste2 As Object = xlWsheetCNI_IC.Range("A3")
                            rg_head_cut2.Select()
                            rg_head_cut2.Cut(rg_head_paste2)

                            xlWsheetCNI_IC.Range("A2").RowHeight = 27

                            Dim paramhead1, paramhead2, paramhead3 As String
                            paramhead1 = xlWsheetCNI_IC.Range("B6").Value & " " & xlWsheetCNI_IC.Range("D6").Value & "; "
                            paramhead1 = paramhead1 & xlWsheetCNI_IC.Range("G6").Value & " " & xlWsheetCNI_IC.Range("I6").Value

                            paramhead2 = xlWsheetCNI_IC.Range("M6").Value & " " & xlWsheetCNI_IC.Range("P6").Value & ";"
                            paramhead2 = paramhead2 & xlWsheetCNI_IC.Range("T6").Value & " " & xlWsheetCNI_IC.Range("W6").Value & "; "
                            paramhead2 = paramhead2 & xlWsheetCNI_IC.Range("Y6").Value & " " & xlWsheetCNI_IC.Range("AA6").Value & "; "


                            paramhead3 = xlWsheetCNI_IC.Range("B8").Value & " " & xlWsheetCNI_IC.Range("D8").Value & "; "
                            paramhead3 = paramhead3 & xlWsheetCNI_IC.Range("G8").Value & " " & xlWsheetCNI_IC.Range("I8").Value & "; "
                            paramhead3 = paramhead3 & xlWsheetCNI_IC.Range("M8").Value & " " & xlWsheetCNI_IC.Range("P8").Value & "; "
                            paramhead3 = paramhead3 & xlWsheetCNI_IC.Range("T8").Value & " " & xlWsheetCNI_IC.Range("W8").Value & "; "
                            paramhead3 = paramhead3 & xlWsheetCNI_IC.Range("Y8").Value & " " & xlWsheetCNI_IC.Range("AA8").Value & "; "

                            xlWsheetCNI_IC.Range("A5").Value = paramhead1
                            xlWsheetCNI_IC.Range("A5").EntireRow.Font.Name = "Calibri"
                            xlWsheetCNI_IC.Range("A6").Value = paramhead2
                            xlWsheetCNI_IC.Range("A6").EntireRow.Font.Name = "Calibri"
                            xlWsheetCNI_IC.Range("A7").Value = paramhead3
                            xlWsheetCNI_IC.Range("A7").EntireRow.Font.Name = "Calibri"

                            xlWsheetCNI_IC.Range("B6:AA8").value = ""
                            xlWsheetCNI_IC.Range("A8:A9").EntireRow.Delete()

                            Dim xlfunc As Excel.WorksheetFunction
                            xlfunc = xlAppCNI_IC.WorksheetFunction()
                            Dim lnCol As Long
                            Dim i As Long
                            Dim rnarea As Object = xlWsheetCNI_IC.UsedRange

                            lnCol = rnarea.Columns.Count
                            For i = lnCol To 1 Step -1
                                If xlfunc.CountA(rnarea.Columns(i)) = 0 Then
                                    rnarea.Columns(i).Delete()
                                End If
                            Next

                            Dim colend As Long
                            colend = xlWsheetCNI_IC.Range("IV8").End(Excel.XlDirection.xlToLeft).Column
                            colend = colend + 3
                            xlWsheetCNI_IC.Cells(8, colend).Value = "End"
                            'xlWsheetCNI_IC.Cells(8, colend).Font.ColorIndex = System.Drawing.ColorTranslator.ToOle(Color.GhostWhite)

                            Dim lastcol1 As Long
                            Dim x, y As Integer
                            Dim lasval As String = ""
                            Dim rn As Object
                            y = 18
                            Do Until lasval = "End"
                                lastcol1 = xlWsheetCNI_IC.Cells(8, y).End(Excel.XlDirection.xlToRight).Column
                                x = lastcol1 - 1
                                rn = xlWsheetCNI_IC.Range(xlWsheetCNI_IC.Cells(8, y), xlWsheetCNI_IC.Cells(8, x))
                                rn.Merge()
                                rn.HorizontalAlignment = Excel.Constants.xlCenter
                                lasval = xlWsheetCNI_IC.Cells(8, lastcol1).Value
                                y = y + 3
                            Loop

                            Dim colend2 As Long
                            colend2 = xlWsheetCNI_IC.Range("IV8").End(Excel.XlDirection.xlToLeft).Column
                            xlWsheetCNI_IC.Cells(8, colend).Value = ""

                            xlWbookCNI_IC.SaveAs(txtCNI_IC_dest.Text)
                            xlWbookCNI_IC.Close()
                            xlAppCNI_IC.Quit()

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetCNI_IC)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookCNI_IC)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppCNI_IC)

                            xlWsheetCNI_IC = Nothing
                            xlWbookCNI_IC = Nothing
                            xlAppCNI_IC = Nothing

                            MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                        Catch ex As Exception
                            MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try

                    Case "SUM"
                        Try
                            xlAppCNI_IC = CreateObject("Excel.Application")
                            xlWbookCNI_IC = xlAppCNI_IC.Workbooks.Open(txtCNI_IC_src.Text)
                            xlWsheetCNI_IC = xlWbookCNI_IC.Worksheets("UID IC Cabinet Net Increase Sum")

                            xlWsheetCNI_IC.UsedRange.UnMerge()
                            xlWsheetCNI_IC.UsedRange.WrapText = False
                            xlWsheetCNI_IC.UsedRange.ColumnWidth = 15
                            xlWsheetCNI_IC.UsedRange.RowHeight = 15

                            Dim rg_head_cut1 As Object = xlWsheetCNI_IC.Range("B2")
                            Dim rg_head_paste1 As Object = xlWsheetCNI_IC.Range("A2")
                            rg_head_cut1.Select()
                            rg_head_cut1.Cut(rg_head_paste1)

                            Dim rg_head_cut2 As Object = xlWsheetCNI_IC.Range("B4")
                            Dim rg_head_paste2 As Object = xlWsheetCNI_IC.Range("A3")
                            rg_head_cut2.Select()
                            rg_head_cut2.Cut(rg_head_paste2)

                            xlWsheetCNI_IC.Range("A2").RowHeight = 27

                            Dim paramhead1, paramhead2, paramhead3 As String
                            paramhead1 = xlWsheetCNI_IC.Range("B6").Value & " " & xlWsheetCNI_IC.Range("D6").Value & "; "
                            paramhead1 = paramhead1 & xlWsheetCNI_IC.Range("G6").Value & " " & xlWsheetCNI_IC.Range("I6").Value

                            paramhead2 = xlWsheetCNI_IC.Range("L6").Value & " " & xlWsheetCNI_IC.Range("O6").Value & ";"
                            paramhead2 = paramhead2 & xlWsheetCNI_IC.Range("R6").Value & " " & xlWsheetCNI_IC.Range("V6").Value & "; "
                            paramhead2 = paramhead2 & xlWsheetCNI_IC.Range("Y6").Value & " " & xlWsheetCNI_IC.Range("AA6").Value & "; "


                            paramhead3 = xlWsheetCNI_IC.Range("B8").Value & " " & xlWsheetCNI_IC.Range("D8").Value & "; "
                            paramhead3 = paramhead3 & xlWsheetCNI_IC.Range("G8").Value & " " & xlWsheetCNI_IC.Range("I8").Value & "; "
                            paramhead3 = paramhead3 & xlWsheetCNI_IC.Range("L8").Value & " " & xlWsheetCNI_IC.Range("O8").Value & "; "
                            paramhead3 = paramhead3 & xlWsheetCNI_IC.Range("R8").Value & " " & xlWsheetCNI_IC.Range("V8").Value & "; "
                            paramhead3 = paramhead3 & xlWsheetCNI_IC.Range("Y8").Value & " " & xlWsheetCNI_IC.Range("AA8").Value & "; "

                            xlWsheetCNI_IC.Range("A5").Value = paramhead1
                            xlWsheetCNI_IC.Range("A5").EntireRow.Font.Name = "Calibri"
                            xlWsheetCNI_IC.Range("A6").Value = paramhead2
                            xlWsheetCNI_IC.Range("A6").EntireRow.Font.Name = "Calibri"
                            xlWsheetCNI_IC.Range("A7").Value = paramhead3
                            xlWsheetCNI_IC.Range("A7").EntireRow.Font.Name = "Calibri"

                            xlWsheetCNI_IC.Range("B6:AA8").value = ""
                            xlWsheetCNI_IC.Range("A8").EntireRow.Delete()

                            Dim xlfunc As Excel.WorksheetFunction
                            xlfunc = xlAppCNI_IC.WorksheetFunction()
                            Dim lnCol As Long
                            Dim i As Long
                            Dim rnarea As Object = xlWsheetCNI_IC.UsedRange

                            lnCol = rnarea.Columns.Count
                            For i = lnCol To 1 Step -1
                                If xlfunc.CountA(rnarea.Columns(i)) = 0 Then
                                    rnarea.Columns(i).Delete()
                                End If
                            Next

                            xlWsheetCNI_IC.Range("H9").WrapText = True
                            xlWsheetCNI_IC.Range("H:H").EntireColumn.ColumnWidth = 17
                            xlWsheetCNI_IC.range("A9").EntireRow.RowHeight = 36

                            xlWbookCNI_IC.SaveAs(txtCNI_IC_dest.Text)
                            xlWbookCNI_IC.Close()
                            xlAppCNI_IC.Quit()

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetCNI_IC)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookCNI_IC)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppCNI_IC)

                            xlWsheetCNI_IC = Nothing
                            xlWbookCNI_IC = Nothing
                            xlAppCNI_IC = Nothing

                            MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                        Catch ex As Exception
                            MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try
                End Select

            Case "XL_Installed"
                Dim xlAppCNI_IC As Object
                Dim xlWbookCNI_IC As Object
                Dim xlWsheetCNI_IC As Object
                Select Case CNI_Type
                    Case "DTL"
                        Try
                            xlAppCNI_IC = CreateObject("Excel.Application")
                            xlWbookCNI_IC = xlAppCNI_IC.Workbooks.Open(txtCNI_IC_src.Text)
                            xlWsheetCNI_IC = xlWbookCNI_IC.Worksheets("UID IC Cabinet Net Increase Det")

                            xlWsheetCNI_IC.UsedRange.UnMerge()
                            xlWsheetCNI_IC.UsedRange.WrapText = False
                            xlWsheetCNI_IC.UsedRange.ColumnWidth = 15
                            xlWsheetCNI_IC.UsedRange.RowHeight = 15

                            Dim rg_head_cut1 As Excel.Range = xlWsheetCNI_IC.Range("B2")
                            Dim rg_head_paste1 As Excel.Range = xlWsheetCNI_IC.Range("A2")
                            rg_head_cut1.Select()
                            rg_head_cut1.Cut(rg_head_paste1)

                            Dim rg_head_cut2 As Excel.Range = xlWsheetCNI_IC.Range("B4")
                            Dim rg_head_paste2 As Excel.Range = xlWsheetCNI_IC.Range("A3")
                            rg_head_cut2.Select()
                            rg_head_cut2.Cut(rg_head_paste2)

                            xlWsheetCNI_IC.Range("A2").RowHeight = 27

                            Dim paramhead1, paramhead2, paramhead3 As String
                            paramhead1 = xlWsheetCNI_IC.Range("B6").Value & " " & xlWsheetCNI_IC.Range("D6").Value & "; "
                            paramhead1 = paramhead1 & xlWsheetCNI_IC.Range("G6").Value & " " & xlWsheetCNI_IC.Range("I6").Value

                            paramhead2 = xlWsheetCNI_IC.Range("M6").Value & " " & xlWsheetCNI_IC.Range("P6").Value & ";"
                            paramhead2 = paramhead2 & xlWsheetCNI_IC.Range("T6").Value & " " & xlWsheetCNI_IC.Range("W6").Value & "; "
                            paramhead2 = paramhead2 & xlWsheetCNI_IC.Range("Y6").Value & " " & xlWsheetCNI_IC.Range("AA6").Value & "; "


                            paramhead3 = xlWsheetCNI_IC.Range("B8").Value & " " & xlWsheetCNI_IC.Range("D8").Value & "; "
                            paramhead3 = paramhead3 & xlWsheetCNI_IC.Range("G8").Value & " " & xlWsheetCNI_IC.Range("I8").Value & "; "
                            paramhead3 = paramhead3 & xlWsheetCNI_IC.Range("M8").Value & " " & xlWsheetCNI_IC.Range("P8").Value & "; "
                            paramhead3 = paramhead3 & xlWsheetCNI_IC.Range("T8").Value & " " & xlWsheetCNI_IC.Range("W8").Value & "; "
                            paramhead3 = paramhead3 & xlWsheetCNI_IC.Range("Y8").Value & " " & xlWsheetCNI_IC.Range("AA8").Value & "; "

                            xlWsheetCNI_IC.Range("A5").Value = paramhead1
                            xlWsheetCNI_IC.Range("A5").EntireRow.Font.Name = "Calibri"
                            xlWsheetCNI_IC.Range("A6").Value = paramhead2
                            xlWsheetCNI_IC.Range("A6").EntireRow.Font.Name = "Calibri"
                            xlWsheetCNI_IC.Range("A7").Value = paramhead3
                            xlWsheetCNI_IC.Range("A7").EntireRow.Font.Name = "Calibri"

                            xlWsheetCNI_IC.Range("B6:AA8").value = ""
                            xlWsheetCNI_IC.Range("A8:A9").EntireRow.Delete()

                            Dim xlfunc As Excel.WorksheetFunction
                            xlfunc = xlAppCNI_IC.WorksheetFunction()
                            Dim lnCol As Long
                            Dim i As Long
                            Dim rnarea As Excel.Range = xlWsheetCNI_IC.UsedRange

                            lnCol = rnarea.Columns.Count
                            For i = lnCol To 1 Step -1
                                If xlfunc.CountA(rnarea.Columns(i)) = 0 Then
                                    rnarea.Columns(i).Delete()
                                End If
                            Next

                            Dim colend As Long
                            colend = xlWsheetCNI_IC.Range("IV8").End(Excel.XlDirection.xlToLeft).Column
                            colend = colend + 3
                            xlWsheetCNI_IC.Cells(8, colend).Value = "End"
                            'xlWsheetCNI_IC.Cells(8, colend).Font.ColorIndex = System.Drawing.ColorTranslator.ToOle(Color.GhostWhite)

                            Dim lastcol1 As Long
                            Dim x, y As Integer
                            Dim lasval As String = ""
                            Dim rn As Excel.Range
                            y = 18
                            Do Until lasval = "End"
                                lastcol1 = xlWsheetCNI_IC.Cells(8, y).End(Excel.XlDirection.xlToRight).Column
                                x = lastcol1 - 1
                                rn = xlWsheetCNI_IC.Range(xlWsheetCNI_IC.Cells(8, y), xlWsheetCNI_IC.Cells(8, x))
                                rn.Merge()
                                rn.HorizontalAlignment = Excel.Constants.xlCenter
                                lasval = xlWsheetCNI_IC.Cells(8, lastcol1).Value
                                y = y + 3
                            Loop

                            Dim colend2 As Long
                            colend2 = xlWsheetCNI_IC.Range("IV8").End(Excel.XlDirection.xlToLeft).Column
                            xlWsheetCNI_IC.Cells(8, colend).Value = ""

                            xlWbookCNI_IC.SaveAs(txtCNI_IC_dest.Text)
                            xlWbookCNI_IC.Close()
                            xlAppCNI_IC.Quit()

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetCNI_IC)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookCNI_IC)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppCNI_IC)

                            xlWsheetCNI_IC = Nothing
                            xlWbookCNI_IC = Nothing
                            xlAppCNI_IC = Nothing

                            MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                        Catch ex As Exception
                            MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try

                    Case "SUM"
                        Try
                            xlAppCNI_IC = CreateObject("Excel.Application")
                            xlWbookCNI_IC = xlAppCNI_IC.Workbooks.Open(txtCNI_IC_src.Text)
                            xlWsheetCNI_IC = xlWbookCNI_IC.Worksheets("UID IC Cabinet Net Increase Sum")

                            xlWsheetCNI_IC.UsedRange.UnMerge()
                            xlWsheetCNI_IC.UsedRange.WrapText = False
                            xlWsheetCNI_IC.UsedRange.ColumnWidth = 15
                            xlWsheetCNI_IC.UsedRange.RowHeight = 15

                            Dim rg_head_cut1 As Excel.Range = xlWsheetCNI_IC.Range("B2")
                            Dim rg_head_paste1 As Excel.Range = xlWsheetCNI_IC.Range("A2")
                            rg_head_cut1.Select()
                            rg_head_cut1.Cut(rg_head_paste1)

                            Dim rg_head_cut2 As Excel.Range = xlWsheetCNI_IC.Range("B4")
                            Dim rg_head_paste2 As Excel.Range = xlWsheetCNI_IC.Range("A3")
                            rg_head_cut2.Select()
                            rg_head_cut2.Cut(rg_head_paste2)

                            xlWsheetCNI_IC.Range("A2").RowHeight = 27

                            Dim paramhead1, paramhead2, paramhead3 As String
                            paramhead1 = xlWsheetCNI_IC.Range("B6").Value & " " & xlWsheetCNI_IC.Range("D6").Value & "; "
                            paramhead1 = paramhead1 & xlWsheetCNI_IC.Range("G6").Value & " " & xlWsheetCNI_IC.Range("I6").Value

                            paramhead2 = xlWsheetCNI_IC.Range("L6").Value & " " & xlWsheetCNI_IC.Range("O6").Value & ";"
                            paramhead2 = paramhead2 & xlWsheetCNI_IC.Range("R6").Value & " " & xlWsheetCNI_IC.Range("V6").Value & "; "
                            paramhead2 = paramhead2 & xlWsheetCNI_IC.Range("Y6").Value & " " & xlWsheetCNI_IC.Range("AA6").Value & "; "


                            paramhead3 = xlWsheetCNI_IC.Range("B8").Value & " " & xlWsheetCNI_IC.Range("D8").Value & "; "
                            paramhead3 = paramhead3 & xlWsheetCNI_IC.Range("G8").Value & " " & xlWsheetCNI_IC.Range("I8").Value & "; "
                            paramhead3 = paramhead3 & xlWsheetCNI_IC.Range("L8").Value & " " & xlWsheetCNI_IC.Range("O8").Value & "; "
                            paramhead3 = paramhead3 & xlWsheetCNI_IC.Range("R8").Value & " " & xlWsheetCNI_IC.Range("V8").Value & "; "
                            paramhead3 = paramhead3 & xlWsheetCNI_IC.Range("Y8").Value & " " & xlWsheetCNI_IC.Range("AA8").Value & "; "

                            xlWsheetCNI_IC.Range("A5").Value = paramhead1
                            xlWsheetCNI_IC.Range("A5").EntireRow.Font.Name = "Calibri"
                            xlWsheetCNI_IC.Range("A6").Value = paramhead2
                            xlWsheetCNI_IC.Range("A6").EntireRow.Font.Name = "Calibri"
                            xlWsheetCNI_IC.Range("A7").Value = paramhead3
                            xlWsheetCNI_IC.Range("A7").EntireRow.Font.Name = "Calibri"

                            xlWsheetCNI_IC.Range("B6:AA8").value = ""
                            xlWsheetCNI_IC.Range("A8").EntireRow.Delete()

                            Dim xlfunc As Excel.WorksheetFunction
                            xlfunc = xlAppCNI_IC.WorksheetFunction()
                            Dim lnCol As Long
                            Dim i As Long
                            Dim rnarea As Excel.Range = xlWsheetCNI_IC.UsedRange

                            lnCol = rnarea.Columns.Count
                            For i = lnCol To 1 Step -1
                                If xlfunc.CountA(rnarea.Columns(i)) = 0 Then
                                    rnarea.Columns(i).Delete()
                                End If
                            Next

                            xlWsheetCNI_IC.Range("H9").WrapText = True
                            xlWsheetCNI_IC.Range("H:H").EntireColumn.ColumnWidth = 17
                            xlWsheetCNI_IC.range("A9").EntireRow.RowHeight = 36

                            xlWbookCNI_IC.SaveAs(txtCNI_IC_dest.Text)
                            xlWbookCNI_IC.Close()
                            xlAppCNI_IC.Quit()

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetCNI_IC)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookCNI_IC)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppCNI_IC)

                            xlWsheetCNI_IC = Nothing
                            xlWbookCNI_IC = Nothing
                            xlAppCNI_IC = Nothing

                            MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                        Catch ex As Exception
                            MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try
                End Select
        End Select
    End Sub

    Private Sub BWCNI_IC_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BWCNI_IC.RunWorkerCompleted
        PicBar_CNI_IC.Visible = False
        btnNeu_CNI_IC.Enabled = False
        txtCNI_IC_dest.Text = ""
        txtCNI_IC_src.Text = ""
    End Sub
End Class