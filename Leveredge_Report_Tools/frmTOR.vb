Imports Microsoft.Office.Interop
Imports System.Windows.Forms
Imports System.Text
Imports System.Drawing
Imports Microsoft.Win32

Public Class frmTOR
    Dim AppsOffice As String
    Private Sub frmTOR_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim TargetKey As RegistryKey
        TargetKey = Registry.ClassesRoot.OpenSubKey("Excel.Application")
        If TargetKey Is Nothing Then
            AppsOffice = "XL_NotInstalled"
        Else
            AppsOffice = "XL_Installed"
            TargetKey.Close()
        End If
    End Sub
    Dim TOR_Type As String
    Private Sub RBTOR_RUP_CheckedChanged(sender As Object, e As EventArgs) Handles RBTOR_RUP.CheckedChanged
        TOR_Type = "RP"
    End Sub

    Private Sub RBTOR_Weight_CheckedChanged(sender As Object, e As EventArgs) Handles RBTOR_Weight.CheckedChanged
        TOR_Type = "WG"
    End Sub

    Private Sub btnBrow_TOR_src_Click(sender As Object, e As EventArgs) Handles btnBrow_TOR_src.Click
        Dim TORpathSrc As String
        If OFD_TOR.ShowDialog = DialogResult.OK Then
            TORpathSrc = OFD_TOR.FileName()
            txtTOR_src.Text = TORpathSrc
        End If
        If txtTOR_dest.Text <> "" Then
            btnNeu_TOR.Enabled = True
        Else
            btnNeu_TOR.Enabled = False
        End If
    End Sub

    Private Sub btnBrow_TOR_dest_Click(sender As Object, e As EventArgs) Handles btnBrow_TOR_dest.Click
        Dim datenow As DateTime = DateTime.Now
        Dim filename As String = "Neutralize_TurnOver_" & TOR_Type & "_" & datenow.ToString("ddMMyyyy_HHmm")
        SFD_TOR.FileName = filename
        Dim TORpath_Dest As String
        If SFD_TOR.ShowDialog = DialogResult.OK Then
            TORpath_Dest = SFD_TOR.FileName
            txtTOR_dest.Text = TORpath_Dest
        End If
        If txtTOR_src.Text <> "" Then
            btnNeu_TOR.Enabled = True
        Else
            btnNeu_TOR.Enabled = False
        End If
        If txtTOR_dest.Text <> "" Then
            btnNeu_TOR.Enabled = True
        Else
            btnNeu_TOR.Enabled = False
        End If
    End Sub

    Private Sub btnNeu_TOR_Click(sender As Object, e As EventArgs) Handles btnNeu_TOR.Click
        PicBar_TOR.Visible = True
        BWTOR.RunWorkerAsync()
    End Sub

    Private Sub BWTOR_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BWTOR.DoWork
        Select Case AppsOffice
            Case "XL_NotInstalled"
                Dim xlAppTOR As Object
                Dim xlWbookTOR As Object
                Dim xlWsheetTOR As Object

                Select Case TOR_Type
                    Case "RP"
                        Try
                            xlAppTOR = CreateObject("Ket.Application")
                            xlWbookTOR = xlAppTOR.Workbooks.Open(txtTOR_src.Text)
                            xlWsheetTOR = xlWbookTOR.Worksheets("UID Turn Over  Report")

                            xlWsheetTOR.UsedRange.UnMerge()
                            xlWsheetTOR.UsedRange.WrapText = False
                            xlWsheetTOR.UsedRange.ColumnWidth = 15
                            xlWsheetTOR.UsedRange.RowHeight = 15

                            Dim rg_head_cut1 As Object = xlWsheetTOR.Range("B2")
                            Dim rg_head_paste1 As Object = xlWsheetTOR.Range("A2")
                            rg_head_cut1.Select()
                            rg_head_cut1.Cut(rg_head_paste1)

                            Dim rg_head_cut2 As Object = xlWsheetTOR.Range("B3")
                            Dim rg_head_paste2 As Object = xlWsheetTOR.Range("A3")
                            rg_head_cut2.Select()
                            rg_head_cut2.Cut(rg_head_paste2)

                            xlWsheetTOR.Range("A2").RowHeight = 27

                            Dim paramhead1, paramhead2, paramhead3 As String
                            paramhead1 = xlWsheetTOR.Range("B5").Value & " " & xlWsheetTOR.Range("F5").Value & "; "
                            paramhead1 = paramhead1 & xlWsheetTOR.Range("J5").Value & " " & xlWsheetTOR.Range("L5").Value

                            paramhead2 = xlWsheetTOR.Range("M5").Value & " " & xlWsheetTOR.Range("S5").Value & "; "
                            paramhead2 = paramhead2 & xlWsheetTOR.Range("X5").Value & " " & xlWsheetTOR.Range("AB5").Value

                            paramhead3 = xlWsheetTOR.Range("AE5").Value & " " & xlWsheetTOR.Range("AH5").Value & "; "
                            paramhead3 = paramhead3 & xlWsheetTOR.Range("B7").Value & " " & xlWsheetTOR.Range("G7").Value & "; "
                            paramhead3 = paramhead3 & xlWsheetTOR.Range("O7").Value & " " & xlWsheetTOR.Range("U7").Value

                            xlWsheetTOR.Range("A5").Value = paramhead1
                            xlWsheetTOR.Range("A5").EntireRow.Font.Name = "Calibri"
                            xlWsheetTOR.Range("A6").Value = paramhead2
                            xlWsheetTOR.Range("A6").EntireRow.Font.Name = "Calibri"
                            xlWsheetTOR.Range("A7").Value = paramhead3
                            xlWsheetTOR.Range("A7").EntireRow.Font.Name = "Calibri"

                            xlWsheetTOR.range("B5:AH7").value = ""

                            Dim xlfunc As Excel.WorksheetFunction
                            xlfunc = xlAppTOR.WorksheetFunction
                            Dim lnCol As Long
                            Dim i As Long
                            Dim rnarea As Object = xlWsheetTOR.UsedRange

                            lnCol = rnarea.Columns.Count
                            For i = lnCol To 1 Step -1
                                If xlfunc.CountA(rnarea.Columns(i)) = 0 Then
                                    rnarea.Columns(i).Delete()
                                End If
                            Next

                            xlWsheetTOR.Range("A9:K9").Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.LawnGreen)

                            xlWbookTOR.SaveAs(txtTOR_dest.Text)
                            xlWbookTOR.Close()
                            xlAppTOR.Quit()

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetTOR)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookTOR)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppTOR)

                            xlWsheetTOR = Nothing
                            xlWbookTOR = Nothing
                            xlAppTOR = Nothing

                            MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                        Catch ex As Exception
                            MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try

                    Case "WG"
                        Try
                            xlAppTOR = CreateObject("Ket.Application")
                            xlWbookTOR = xlAppTOR.Workbooks.Open(txtTOR_src.Text)
                            xlWsheetTOR = xlWbookTOR.Worksheets("UID Turn Over  Report")

                            xlWsheetTOR.UsedRange.UnMerge()
                            xlWsheetTOR.UsedRange.WrapText = False
                            xlWsheetTOR.UsedRange.ColumnWidth = 15
                            xlWsheetTOR.UsedRange.RowHeight = 15

                            Dim rg_head_cut1 As Object = xlWsheetTOR.Range("B2")
                            Dim rg_head_paste1 As Object = xlWsheetTOR.Range("A2")
                            rg_head_cut1.Select()
                            rg_head_cut1.Cut(rg_head_paste1)

                            Dim rg_head_cut2 As Object = xlWsheetTOR.Range("B3")
                            Dim rg_head_paste2 As Object = xlWsheetTOR.Range("A3")
                            rg_head_cut2.Select()
                            rg_head_cut2.Cut(rg_head_paste2)

                            xlWsheetTOR.Range("A2").RowHeight = 27

                            Dim paramhead1, paramhead2, paramhead3 As String
                            paramhead1 = xlWsheetTOR.Range("B5").Value & " " & xlWsheetTOR.Range("F5").Value & "; "
                            paramhead1 = paramhead1 & xlWsheetTOR.Range("J5").Value & " " & xlWsheetTOR.Range("L5").Value

                            paramhead2 = xlWsheetTOR.Range("M5").Value & " " & xlWsheetTOR.Range("S5").Value & "; "
                            paramhead2 = paramhead2 & xlWsheetTOR.Range("X5").Value & " " & xlWsheetTOR.Range("AB5").Value

                            paramhead3 = xlWsheetTOR.Range("AE5").Value & " " & xlWsheetTOR.Range("AH5").Value & "; "
                            paramhead3 = paramhead3 & xlWsheetTOR.Range("B7").Value & " " & xlWsheetTOR.Range("G7").Value & "; "
                            paramhead3 = paramhead3 & xlWsheetTOR.Range("O7").Value & " " & xlWsheetTOR.Range("U7").Value

                            xlWsheetTOR.Range("A5").Value = paramhead1
                            xlWsheetTOR.Range("A5").EntireRow.Font.Name = "Calibri"
                            xlWsheetTOR.Range("A6").Value = paramhead2
                            xlWsheetTOR.Range("A6").EntireRow.Font.Name = "Calibri"
                            xlWsheetTOR.Range("A7").Value = paramhead3
                            xlWsheetTOR.Range("A7").EntireRow.Font.Name = "Calibri"

                            xlWsheetTOR.range("B5:AH7").value = ""

                            Dim xlfunc As Excel.WorksheetFunction
                            xlfunc = xlAppTOR.WorksheetFunction
                            Dim lnCol As Long
                            Dim i As Long
                            Dim rnarea As Object = xlWsheetTOR.UsedRange

                            lnCol = rnarea.Columns.Count
                            For i = lnCol To 1 Step -1
                                If xlfunc.CountA(rnarea.Columns(i)) = 0 Then
                                    rnarea.Columns(i).Delete()
                                End If
                            Next

                            xlWsheetTOR.Range("A9:K9").Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.LawnGreen)

                            xlWbookTOR.SaveAs(txtTOR_dest.Text)
                            xlWbookTOR.Close()
                            xlAppTOR.Quit()

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetTOR)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookTOR)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppTOR)

                            xlWsheetTOR = Nothing
                            xlWbookTOR = Nothing
                            xlAppTOR = Nothing

                            MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                        Catch ex As Exception
                            MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try

                End Select

            Case "XL_Installed"
                Dim xlAppTOR As Object
                Dim xlWbookTOR As Object
                Dim xlWsheetTOR As Object

                Select Case TOR_Type
                    Case "RP"
                        Try
                            xlAppTOR = CreateObject("Excel.Application")
                            xlWbookTOR = xlAppTOR.Workbooks.Open(txtTOR_src.Text)
                            xlWsheetTOR = xlWbookTOR.Worksheets("UID Turn Over  Report")

                            xlWsheetTOR.UsedRange.UnMerge()
                            xlWsheetTOR.UsedRange.WrapText = False
                            xlWsheetTOR.UsedRange.ColumnWidth = 15
                            xlWsheetTOR.UsedRange.RowHeight = 15

                            Dim rg_head_cut1 As Excel.Range = xlWsheetTOR.Range("B2")
                            Dim rg_head_paste1 As Excel.Range = xlWsheetTOR.Range("A2")
                            rg_head_cut1.Select()
                            rg_head_cut1.Cut(rg_head_paste1)

                            Dim rg_head_cut2 As Excel.Range = xlWsheetTOR.Range("B3")
                            Dim rg_head_paste2 As Excel.Range = xlWsheetTOR.Range("A3")
                            rg_head_cut2.Select()
                            rg_head_cut2.Cut(rg_head_paste2)

                            xlWsheetTOR.Range("A2").RowHeight = 27

                            Dim paramhead1, paramhead2, paramhead3 As String
                            paramhead1 = xlWsheetTOR.Range("B5").Value & " " & xlWsheetTOR.Range("F5").Value & "; "
                            paramhead1 = paramhead1 & xlWsheetTOR.Range("J5").Value & " " & xlWsheetTOR.Range("L5").Value

                            paramhead2 = xlWsheetTOR.Range("M5").Value & " " & xlWsheetTOR.Range("S5").Value & "; "
                            paramhead2 = paramhead2 & xlWsheetTOR.Range("X5").Value & " " & xlWsheetTOR.Range("AB5").Value

                            paramhead3 = xlWsheetTOR.Range("AE5").Value & " " & xlWsheetTOR.Range("AH5").Value & "; "
                            paramhead3 = paramhead3 & xlWsheetTOR.Range("B7").Value & " " & xlWsheetTOR.Range("G7").Value & "; "
                            paramhead3 = paramhead3 & xlWsheetTOR.Range("O7").Value & " " & xlWsheetTOR.Range("U7").Value

                            xlWsheetTOR.Range("A5").Value = paramhead1
                            xlWsheetTOR.Range("A5").EntireRow.Font.Name = "Calibri"
                            xlWsheetTOR.Range("A6").Value = paramhead2
                            xlWsheetTOR.Range("A6").EntireRow.Font.Name = "Calibri"
                            xlWsheetTOR.Range("A7").Value = paramhead3
                            xlWsheetTOR.Range("A7").EntireRow.Font.Name = "Calibri"

                            xlWsheetTOR.range("B5:AH7").value = ""

                            Dim xlfunc As Excel.WorksheetFunction
                            xlfunc = xlAppTOR.WorksheetFunction
                            Dim lnCol As Long
                            Dim i As Long
                            Dim rnarea As Excel.Range = xlWsheetTOR.UsedRange

                            lnCol = rnarea.Columns.Count
                            For i = lnCol To 1 Step -1
                                If xlfunc.CountA(rnarea.Columns(i)) = 0 Then
                                    rnarea.Columns(i).Delete()
                                End If
                            Next

                            xlWsheetTOR.Range("A9:K9").Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.LawnGreen)

                            xlWbookTOR.SaveAs(txtTOR_dest.Text)
                            xlWbookTOR.Close()
                            xlAppTOR.Quit()

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetTOR)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookTOR)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppTOR)

                            xlWsheetTOR = Nothing
                            xlWbookTOR = Nothing
                            xlAppTOR = Nothing

                            MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                        Catch ex As Exception
                            MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try

                    Case "WG"
                        Try
                            xlAppTOR = CreateObject("Excel.Application")
                            xlWbookTOR = xlAppTOR.Workbooks.Open(txtTOR_src.Text)
                            xlWsheetTOR = xlWbookTOR.Worksheets("UID Turn Over  Report")

                            xlWsheetTOR.UsedRange.UnMerge()
                            xlWsheetTOR.UsedRange.WrapText = False
                            xlWsheetTOR.UsedRange.ColumnWidth = 15
                            xlWsheetTOR.UsedRange.RowHeight = 15

                            Dim rg_head_cut1 As Excel.Range = xlWsheetTOR.Range("B2")
                            Dim rg_head_paste1 As Excel.Range = xlWsheetTOR.Range("A2")
                            rg_head_cut1.Select()
                            rg_head_cut1.Cut(rg_head_paste1)

                            Dim rg_head_cut2 As Excel.Range = xlWsheetTOR.Range("B3")
                            Dim rg_head_paste2 As Excel.Range = xlWsheetTOR.Range("A3")
                            rg_head_cut2.Select()
                            rg_head_cut2.Cut(rg_head_paste2)

                            xlWsheetTOR.Range("A2").RowHeight = 27

                            Dim paramhead1, paramhead2, paramhead3 As String
                            paramhead1 = xlWsheetTOR.Range("B5").Value & " " & xlWsheetTOR.Range("F5").Value & "; "
                            paramhead1 = paramhead1 & xlWsheetTOR.Range("J5").Value & " " & xlWsheetTOR.Range("L5").Value

                            paramhead2 = xlWsheetTOR.Range("M5").Value & " " & xlWsheetTOR.Range("S5").Value & "; "
                            paramhead2 = paramhead2 & xlWsheetTOR.Range("X5").Value & " " & xlWsheetTOR.Range("AB5").Value

                            paramhead3 = xlWsheetTOR.Range("AE5").Value & " " & xlWsheetTOR.Range("AH5").Value & "; "
                            paramhead3 = paramhead3 & xlWsheetTOR.Range("B7").Value & " " & xlWsheetTOR.Range("G7").Value & "; "
                            paramhead3 = paramhead3 & xlWsheetTOR.Range("O7").Value & " " & xlWsheetTOR.Range("U7").Value

                            xlWsheetTOR.Range("A5").Value = paramhead1
                            xlWsheetTOR.Range("A5").EntireRow.Font.Name = "Calibri"
                            xlWsheetTOR.Range("A6").Value = paramhead2
                            xlWsheetTOR.Range("A6").EntireRow.Font.Name = "Calibri"
                            xlWsheetTOR.Range("A7").Value = paramhead3
                            xlWsheetTOR.Range("A7").EntireRow.Font.Name = "Calibri"

                            xlWsheetTOR.range("B5:AH7").value = ""

                            Dim xlfunc As Excel.WorksheetFunction
                            xlfunc = xlAppTOR.WorksheetFunction
                            Dim lnCol As Long
                            Dim i As Long
                            Dim rnarea As Excel.Range = xlWsheetTOR.UsedRange

                            lnCol = rnarea.Columns.Count
                            For i = lnCol To 1 Step -1
                                If xlfunc.CountA(rnarea.Columns(i)) = 0 Then
                                    rnarea.Columns(i).Delete()
                                End If
                            Next

                            xlWsheetTOR.Range("A9:K9").Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.LawnGreen)

                            xlWbookTOR.SaveAs(txtTOR_dest.Text)
                            xlWbookTOR.Close()
                            xlAppTOR.Quit()

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetTOR)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookTOR)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppTOR)

                            xlWsheetTOR = Nothing
                            xlWbookTOR = Nothing
                            xlAppTOR = Nothing

                            MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                        Catch ex As Exception
                            MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try

                End Select

        End Select
    End Sub

    Private Sub BWTOR_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BWTOR.RunWorkerCompleted
        PicBar_TOR.Visible = False
        btnNeu_TOR.Enabled = False
        txtTOR_dest.Text = ""
        txtTOR_src.Text = ""
    End Sub
End Class