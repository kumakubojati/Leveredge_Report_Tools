Imports Microsoft.Office.Interop
Imports System.Windows.Forms
Imports System.Text
Imports System.Drawing
Imports Microsoft.Win32
Public Class frmSL
    Dim AppsOffice As String
    Private Sub frmSL_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim TargetKey As RegistryKey
        TargetKey = Registry.ClassesRoot.OpenSubKey("Excel.Application")
        If TargetKey Is Nothing Then
            AppsOffice = "XL_NotInstalled"
        Else
            AppsOffice = "XL_Installed"
            TargetKey.Close()
        End If
    End Sub
    Dim SL_Type As String

    Private Sub RBSL_Rekap_CheckedChanged(sender As Object, e As EventArgs) Handles RBSL_Rekap.CheckedChanged
        SL_Type = "RKP"
    End Sub

    Private Sub RBSL_Detail_CheckedChanged(sender As Object, e As EventArgs) Handles RBSL_Detail.CheckedChanged
        SL_Type = "DTL"
    End Sub

    Private Sub btnBrow_SL_src_Click(sender As Object, e As EventArgs) Handles btnBrow_SL_src.Click
        Dim SLpathSrc As String
        If OFD_SL.ShowDialog = DialogResult.OK Then
            SLpathSrc = OFD_SL.FileName()
            txtSL_src.Text = SLpathSrc
        End If
        If txtSL_dest.Text <> "" Then
            btnNeu_SL.Enabled = True
        Else
            btnNeu_SL.Enabled = False
        End If
    End Sub

    Private Sub btnBrow_SL_dest_Click(sender As Object, e As EventArgs) Handles btnBrow_SL_dest.Click
        Dim datenow As DateTime = DateTime.Now
        Dim filename As String = "Neutralize_ServiceLevel_" & SL_Type & "_" & datenow.ToString("ddMMyyyy_HHmm")
        SFD_SL.FileName = filename
        Dim SLpath_Dest As String
        If SFD_SL.ShowDialog = DialogResult.OK Then
            SLpath_Dest = SFD_SL.FileName
            txtSL_dest.Text = SLpath_Dest
        End If
        If txtSL_src.Text <> "" Then
            btnNeu_SL.Enabled = True
        Else
            btnNeu_SL.Enabled = False
        End If
        If txtSL_dest.Text <> "" Then
            btnNeu_SL.Enabled = True
        Else
            btnNeu_SL.Enabled = False
        End If
    End Sub

    Private Sub btnNeu_SL_Click(sender As Object, e As EventArgs) Handles btnNeu_SL.Click
        PicBar_SL.Visible = True
        BW_SL.RunWorkerAsync()
    End Sub

    Private Sub BW_SL_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BW_SL.DoWork
        Select Case AppsOffice
            Case "XL_NotInstalled"
                Dim xlAppSL As Object
                Dim xlWbookSL As Object
                Dim xlWsheetSL As Object

                Select Case SL_Type
                    Case "RKP"
                        Try
                            xlAppSL = CreateObject("Ket.Application")
                            xlWbookSL = xlAppSL.Workbooks.Open(txtSL_src.Text)
                            xlWsheetSL = xlWbookSL.Worksheets("UID Service Level Report")

                            xlWsheetSL.UsedRange.UnMerge()
                            xlWsheetSL.UsedRange.WrapText = False
                            xlWsheetSL.UsedRange.ColumnWidth = 15
                            xlWsheetSL.UsedRange.RowHeight = 15

                            xlWsheetSL.Range("A8").EntireRow.Delete()

                            Dim rg_head_cut1 As Object = xlWsheetSL.Range("B2")
                            Dim rg_head_paste1 As Object = xlWsheetSL.Range("A2")
                            rg_head_cut1.Select()
                            rg_head_cut1.Cut(rg_head_paste1)

                            Dim rg_head_cut2 As Object = xlWsheetSL.Range("B4")
                            Dim rg_head_paste2 As Object = xlWsheetSL.Range("A3")
                            rg_head_cut2.Select()
                            rg_head_cut2.Cut(rg_head_paste2)

                            xlWsheetSL.Range("A2").RowHeight = 27

                            Dim paramhead1, paramhead2, paramhead3 As String
                            paramhead1 = xlWsheetSL.Range("C6").Value & " " & xlWsheetSL.Range("E6").Value & "; "
                            paramhead1 = paramhead1 & xlWsheetSL.Range("I6").Value & " " & xlWsheetSL.Range("K6").Value

                            paramhead2 = xlWsheetSL.Range("N6").Value & " " & xlWsheetSL.Range("Q6").Value & "; "
                            paramhead2 = paramhead2 & xlWsheetSL.Range("U6").Value & " " & xlWsheetSL.Range("W6").Value

                            paramhead3 = xlWsheetSL.Range("AA6").Value & " " & xlWsheetSL.Range("AC6").Value

                            xlWsheetSL.Range("A5").Value = paramhead1
                            xlWsheetSL.Range("A5").EntireRow.Font.Name = "Calibri"
                            xlWsheetSL.Range("A6").Value = paramhead2
                            xlWsheetSL.Range("A6").EntireRow.Font.Name = "Calibri"
                            xlWsheetSL.Range("A7").Value = paramhead3
                            xlWsheetSL.Range("A7").EntireRow.Font.Name = "Calibri"

                            xlWsheetSL.range("C6:AC6").value = ""

                            Dim xlfunc As Excel.WorksheetFunction
                            xlfunc = xlAppSL.WorksheetFunction
                            Dim lnCol As Long
                            Dim i As Long
                            Dim rnarea As Object = xlWsheetSL.UsedRange

                            lnCol = rnarea.Columns.Count
                            For i = lnCol To 1 Step -1
                                If xlfunc.CountA(rnarea.Columns(i)) = 0 Then
                                    rnarea.Columns(i).Delete()
                                End If
                            Next

                            xlWsheetSL.Range("A8:C8").Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.LawnGreen)

                            xlWbookSL.SaveAs(txtSL_dest.Text)
                            xlWbookSL.Close()
                            xlAppSL.Quit()

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetSL)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookSL)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppSL)

                            xlWsheetSL = Nothing
                            xlWbookSL = Nothing
                            xlAppSL = Nothing

                            MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                        Catch ex As Exception
                            MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try

                    Case "DTL"
                        Try
                            xlAppSL = CreateObject("Ket.Application")
                            xlWbookSL = xlAppSL.Workbooks.Open(txtSL_src.Text)
                            xlWsheetSL = xlWbookSL.Worksheets("UID Service Level Report")

                            xlWsheetSL.UsedRange.UnMerge()
                            xlWsheetSL.UsedRange.WrapText = False
                            xlWsheetSL.UsedRange.ColumnWidth = 15
                            xlWsheetSL.UsedRange.RowHeight = 15

                            Dim rg_head_cut1 As Object = xlWsheetSL.Range("B2")
                            Dim rg_head_paste1 As Object = xlWsheetSL.Range("A2")
                            rg_head_cut1.Select()
                            rg_head_cut1.Cut(rg_head_paste1)

                            Dim rg_head_cut2 As Object = xlWsheetSL.Range("B4")
                            Dim rg_head_paste2 As Object = xlWsheetSL.Range("A3")
                            rg_head_cut2.Select()
                            rg_head_cut2.Cut(rg_head_paste2)

                            xlWsheetSL.Range("A2").RowHeight = 27

                            Dim paramhead1, paramhead2, paramhead3 As String
                            paramhead1 = xlWsheetSL.Range("C6").Value & " " & xlWsheetSL.Range("F6").Value & "; "
                            paramhead1 = paramhead1 & xlWsheetSL.Range("H6").Value & " " & xlWsheetSL.Range("J6").Value

                            paramhead2 = xlWsheetSL.Range("N6").Value & " " & xlWsheetSL.Range("Q6").Value & "; "
                            paramhead2 = paramhead2 & xlWsheetSL.Range("U6").Value & " " & xlWsheetSL.Range("X6").Value

                            paramhead3 = xlWsheetSL.Range("AB6").Value & " " & xlWsheetSL.Range("AF6").Value

                            xlWsheetSL.Range("A5").Value = paramhead1
                            xlWsheetSL.Range("A5").EntireRow.Font.Name = "Calibri"
                            xlWsheetSL.Range("A6").Value = paramhead2
                            xlWsheetSL.Range("A6").EntireRow.Font.Name = "Calibri"
                            xlWsheetSL.Range("A7").Value = paramhead3
                            xlWsheetSL.Range("A7").EntireRow.Font.Name = "Calibri"

                            xlWsheetSL.range("C6:AF6").value = ""

                            Dim xlfunc As Excel.WorksheetFunction
                            xlfunc = xlAppSL.WorksheetFunction
                            Dim lnCol As Long
                            Dim i As Long
                            Dim rnarea As Object = xlWsheetSL.UsedRange

                            lnCol = rnarea.Columns.Count
                            For i = lnCol To 1 Step -1
                                If xlfunc.CountA(rnarea.Columns(i)) = 0 Then
                                    rnarea.Columns(i).Delete()
                                End If
                            Next

                            xlWsheetSL.Range("A8:C8").Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.LawnGreen)

                            xlWbookSL.SaveAs(txtSL_dest.Text)
                            xlWbookSL.Close()
                            xlAppSL.Quit()

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetSL)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookSL)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppSL)

                            xlWsheetSL = Nothing
                            xlWbookSL = Nothing
                            xlAppSL = Nothing

                            MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                        Catch ex As Exception
                            MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try
                End Select
            Case "XL_Installed"
                Dim xlAppSL As Object
                Dim xlWbookSL As Object
                Dim xlWsheetSL As Object

                Select Case SL_Type
                    Case "RKP"
                        Try
                            xlAppSL = CreateObject("Excel.Application")
                            xlWbookSL = xlAppSL.Workbooks.Open(txtSL_src.Text)
                            xlWsheetSL = xlWbookSL.Worksheets("UID Service Level Report")

                            xlWsheetSL.UsedRange.UnMerge()
                            xlWsheetSL.UsedRange.WrapText = False
                            xlWsheetSL.UsedRange.ColumnWidth = 15
                            xlWsheetSL.UsedRange.RowHeight = 15

                            xlWsheetSL.Range("A8").EntireRow.Delete()

                            Dim rg_head_cut1 As Excel.Range = xlWsheetSL.Range("B2")
                            Dim rg_head_paste1 As Excel.Range = xlWsheetSL.Range("A2")
                            rg_head_cut1.Select()
                            rg_head_cut1.Cut(rg_head_paste1)

                            Dim rg_head_cut2 As Excel.Range = xlWsheetSL.Range("B4")
                            Dim rg_head_paste2 As Excel.Range = xlWsheetSL.Range("A3")
                            rg_head_cut2.Select()
                            rg_head_cut2.Cut(rg_head_paste2)

                            xlWsheetSL.Range("A2").RowHeight = 27

                            Dim paramhead1, paramhead2, paramhead3 As String
                            paramhead1 = xlWsheetSL.Range("C6").Value & " " & xlWsheetSL.Range("E6").Value & "; "
                            paramhead1 = paramhead1 & xlWsheetSL.Range("I6").Value & " " & xlWsheetSL.Range("K6").Value

                            paramhead2 = xlWsheetSL.Range("N6").Value & " " & xlWsheetSL.Range("Q6").Value & "; "
                            paramhead2 = paramhead2 & xlWsheetSL.Range("U6").Value & " " & xlWsheetSL.Range("W6").Value

                            paramhead3 = xlWsheetSL.Range("AA6").Value & " " & xlWsheetSL.Range("AC6").Value

                            xlWsheetSL.Range("A5").Value = paramhead1
                            xlWsheetSL.Range("A5").EntireRow.Font.Name = "Calibri"
                            xlWsheetSL.Range("A6").Value = paramhead2
                            xlWsheetSL.Range("A6").EntireRow.Font.Name = "Calibri"
                            xlWsheetSL.Range("A7").Value = paramhead3
                            xlWsheetSL.Range("A7").EntireRow.Font.Name = "Calibri"

                            xlWsheetSL.range("C6:AC6").value = ""

                            Dim xlfunc As Excel.WorksheetFunction
                            xlfunc = xlAppSL.WorksheetFunction
                            Dim lnCol As Long
                            Dim i As Long
                            Dim rnarea As Excel.Range = xlWsheetSL.UsedRange

                            lnCol = rnarea.Columns.Count
                            For i = lnCol To 1 Step -1
                                If xlfunc.CountA(rnarea.Columns(i)) = 0 Then
                                    rnarea.Columns(i).Delete()
                                End If
                            Next

                            xlWsheetSL.Range("A8:C8").Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.LawnGreen)

                            xlWbookSL.SaveAs(txtSL_dest.Text)
                            xlWbookSL.Close()
                            xlAppSL.Quit()

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetSL)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookSL)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppSL)

                            xlWsheetSL = Nothing
                            xlWbookSL = Nothing
                            xlAppSL = Nothing

                            MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                        Catch ex As Exception
                            MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try

                    Case "DTL"
                        Try
                            xlAppSL = CreateObject("Excel.Application")
                            xlWbookSL = xlAppSL.Workbooks.Open(txtSL_src.Text)
                            xlWsheetSL = xlWbookSL.Worksheets("UID Service Level Report")

                            xlWsheetSL.UsedRange.UnMerge()
                            xlWsheetSL.UsedRange.WrapText = False
                            xlWsheetSL.UsedRange.ColumnWidth = 15
                            xlWsheetSL.UsedRange.RowHeight = 15

                            Dim rg_head_cut1 As Excel.Range = xlWsheetSL.Range("B2")
                            Dim rg_head_paste1 As Excel.Range = xlWsheetSL.Range("A2")
                            rg_head_cut1.Select()
                            rg_head_cut1.Cut(rg_head_paste1)

                            Dim rg_head_cut2 As Excel.Range = xlWsheetSL.Range("B4")
                            Dim rg_head_paste2 As Excel.Range = xlWsheetSL.Range("A3")
                            rg_head_cut2.Select()
                            rg_head_cut2.Cut(rg_head_paste2)

                            xlWsheetSL.Range("A2").RowHeight = 27

                            Dim paramhead1, paramhead2, paramhead3 As String
                            paramhead1 = xlWsheetSL.Range("C6").Value & " " & xlWsheetSL.Range("F6").Value & "; "
                            paramhead1 = paramhead1 & xlWsheetSL.Range("H6").Value & " " & xlWsheetSL.Range("J6").Value

                            paramhead2 = xlWsheetSL.Range("N6").Value & " " & xlWsheetSL.Range("Q6").Value & "; "
                            paramhead2 = paramhead2 & xlWsheetSL.Range("U6").Value & " " & xlWsheetSL.Range("X6").Value

                            paramhead3 = xlWsheetSL.Range("AB6").Value & " " & xlWsheetSL.Range("AF6").Value

                            xlWsheetSL.Range("A5").Value = paramhead1
                            xlWsheetSL.Range("A5").EntireRow.Font.Name = "Calibri"
                            xlWsheetSL.Range("A6").Value = paramhead2
                            xlWsheetSL.Range("A6").EntireRow.Font.Name = "Calibri"
                            xlWsheetSL.Range("A7").Value = paramhead3
                            xlWsheetSL.Range("A7").EntireRow.Font.Name = "Calibri"

                            xlWsheetSL.range("C6:AF6").value = ""

                            Dim xlfunc As Excel.WorksheetFunction
                            xlfunc = xlAppSL.WorksheetFunction
                            Dim lnCol As Long
                            Dim i As Long
                            Dim rnarea As Excel.Range = xlWsheetSL.UsedRange

                            lnCol = rnarea.Columns.Count
                            For i = lnCol To 1 Step -1
                                If xlfunc.CountA(rnarea.Columns(i)) = 0 Then
                                    rnarea.Columns(i).Delete()
                                End If
                            Next

                            xlWsheetSL.Range("A8:C8").Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.LawnGreen)

                            xlWbookSL.SaveAs(txtSL_dest.Text)
                            xlWbookSL.Close()
                            xlAppSL.Quit()

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetSL)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookSL)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppSL)

                            xlWsheetSL = Nothing
                            xlWbookSL = Nothing
                            xlAppSL = Nothing

                            MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                        Catch ex As Exception
                            MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try
                End Select
        End Select
    End Sub

    Private Sub BW_SL_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BW_SL.RunWorkerCompleted
        PicBar_SL.Visible = False
        btnNeu_SL.Enabled = False
        txtSL_dest.Text = ""
        txtSL_src.Text = ""
    End Sub
End Class