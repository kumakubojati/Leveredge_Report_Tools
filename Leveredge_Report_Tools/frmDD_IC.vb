Imports Microsoft.Office.Interop
Imports System.Windows.Forms
Imports System.Text
Imports System.Drawing
Imports Microsoft.Win32
Public Class frmDD_IC
    Dim AppsOffice As String
    Private Sub frmDD_IC_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim TargetKey As RegistryKey
        TargetKey = Registry.ClassesRoot.OpenSubKey("Excel.Application")
        If TargetKey Is Nothing Then
            AppsOffice = "XL_NotInstalled"
        Else
            AppsOffice = "XL_Installed"
            TargetKey.Close()
        End If
    End Sub
    Dim DD_type As String

    Private Sub RBDD_Dist_CheckedChanged(sender As Object, e As EventArgs) Handles RBDD_Dist.CheckedChanged
        DD_type = "DIST"
    End Sub

    Private Sub RBDD_NotDist_CheckedChanged(sender As Object, e As EventArgs) Handles RBDD_NotDist.CheckedChanged
        DD_type = "NOTDIST"
    End Sub

    Private Sub btnBrow_DD_IC_src_Click(sender As Object, e As EventArgs) Handles btnBrow_DD_IC_src.Click
        Dim DDpathSrc As String
        If OFD_DD_IC.ShowDialog = DialogResult.OK Then
            DDpathSrc = OFD_DD_IC.FileName()
            txtDD_IC_src.Text = DDpathSrc
        End If
        If txtDD_IC_dest.Text <> "" Then
            btnNeu_DD_IC.Enabled = True
        Else
            btnNeu_DD_IC.Enabled = False
        End If
    End Sub

    Private Sub btnBrow_DD_IC_dest_Click(sender As Object, e As EventArgs) Handles btnBrow_DD_IC_dest.Click
        Dim datenow As DateTime = DateTime.Now
        Dim filename As String = "Neutralize_IC_DistDrive_" & DD_type & "_" & datenow.ToString("ddMMyyyy_HHmm")
        SFD_DD_IC.FileName = filename
        Dim DDpath_Dest As String
        If SFD_DD_IC.ShowDialog = DialogResult.OK Then
            DDpath_Dest = SFD_DD_IC.FileName
            txtDD_IC_dest.Text = DDpath_Dest
        End If
        If txtDD_IC_src.Text <> "" Then
            btnNeu_DD_IC.Enabled = True
        Else
            btnNeu_DD_IC.Enabled = False
        End If
        If txtDD_IC_dest.Text <> "" Then
            btnNeu_DD_IC.Enabled = True
        Else
            btnNeu_DD_IC.Enabled = False
        End If
    End Sub

    Private Sub btnNeu_DD_IC_Click(sender As Object, e As EventArgs) Handles btnNeu_DD_IC.Click
        PicBar_DD_IC.Visible = True
        BWDD_IC.RunWorkerAsync()
    End Sub

    Private Sub BWDD_IC_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BWDD_IC.DoWork
        Select Case AppsOffice
            Case "XL_NotInstalled"
                Dim xlAppDD_IC As Object
                Dim xlWbookDD_IC As Object
                Dim xlWsheetDD_IC As Object
                Select Case DD_type
                    Case "DIST"
                        Try
                            xlAppDD_IC = CreateObject("Ket.Application")
                            xlWbookDD_IC = xlAppDD_IC.Workbooks.Open(txtDD_IC_src.Text)
                            xlWsheetDD_IC = xlWbookDD_IC.Worksheets("UID IC Distribution Drive Repor")

                            xlWsheetDD_IC.UsedRange.UnMerge()
                            xlWsheetDD_IC.UsedRange.WrapText = False
                            xlWsheetDD_IC.UsedRange.ColumnWidth = 15
                            xlWsheetDD_IC.UsedRange.RowHeight = 15

                            Dim rg_head_cut1 As Object = xlWsheetDD_IC.Range("B2")
                            Dim rg_head_paste1 As Object = xlWsheetDD_IC.Range("A2")
                            rg_head_cut1.Select()
                            rg_head_cut1.Cut(rg_head_paste1)

                            Dim rg_head_cut2 As Object = xlWsheetDD_IC.Range("B4")
                            Dim rg_head_paste2 As Object = xlWsheetDD_IC.Range("A3")
                            rg_head_cut2.Select()
                            rg_head_cut2.Cut(rg_head_paste2)

                            xlWsheetDD_IC.Range("A2").RowHeight = 27

                            Dim paramhead1, paramhead2, paramhead3 As String
                            paramhead1 = xlWsheetDD_IC.Range("B6").Value & " " & xlWsheetDD_IC.Range("E6").Value & "; "
                            paramhead1 = paramhead1 & xlWsheetDD_IC.Range("I6").Value & " " & xlWsheetDD_IC.Range("L6").Value

                            paramhead2 = xlWsheetDD_IC.Range("R6").Value & " " & xlWsheetDD_IC.Range("U6").Value & ";"
                            paramhead2 = paramhead2 & xlWsheetDD_IC.Range("Y6").Value & " " & xlWsheetDD_IC.Range("AB6").Value

                            paramhead3 = xlWsheetDD_IC.Range("B8").Value & " " & xlWsheetDD_IC.Range("E8").Value & "; "
                            paramhead3 = paramhead3 & xlWsheetDD_IC.Range("I8").Value & " " & xlWsheetDD_IC.Range("M8").Value & "; "
                            paramhead3 = paramhead3 & xlWsheetDD_IC.Range("R8").Value & " " & xlWsheetDD_IC.Range("U8").Value


                            xlWsheetDD_IC.Range("A5").Value = paramhead1
                            xlWsheetDD_IC.Range("A5").EntireRow.Font.Name = "Calibri"
                            xlWsheetDD_IC.Range("A6").Value = paramhead2
                            xlWsheetDD_IC.Range("A6").EntireRow.Font.Name = "Calibri"
                            xlWsheetDD_IC.Range("A7").Value = paramhead3
                            xlWsheetDD_IC.Range("A7").EntireRow.Font.Name = "Calibri"

                            xlWsheetDD_IC.Range("B6:AB8").value = ""
                            xlWsheetDD_IC.Range("A10").EntireRow.Delete()

                            Dim xlfunc As Excel.WorksheetFunction
                            xlfunc = xlAppDD_IC.WorksheetFunction()
                            Dim lnCol As Long
                            Dim i As Long
                            Dim rnarea As Object = xlWsheetDD_IC.UsedRange

                            lnCol = rnarea.Columns.Count
                            For i = lnCol To 1 Step -1
                                If xlfunc.CountA(rnarea.Columns(i)) = 0 Then
                                    rnarea.Columns(i).Delete()
                                End If
                            Next

                            xlWsheetDD_IC.Range("A10:B10").Merge()
                            xlWsheetDD_IC.Range("C10:D10").Merge()
                            xlWsheetDD_IC.Range("E10:F10").Merge()
                            xlWsheetDD_IC.Range("L10:M10").Merge()

                            xlWbookDD_IC.SaveAs(txtDD_IC_dest.Text)
                            xlWbookDD_IC.Close()
                            xlAppDD_IC.Quit()

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetDD_IC)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookDD_IC)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppDD_IC)

                            xlWsheetDD_IC = Nothing
                            xlWbookDD_IC = Nothing
                            xlAppDD_IC = Nothing

                            MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                        Catch ex As Exception
                            MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try

                    Case "NOTDIST"
                        Try
                            xlAppDD_IC = CreateObject("Ket.Application")
                            xlWbookDD_IC = xlAppDD_IC.Workbooks.Open(txtDD_IC_src.Text)
                            xlWsheetDD_IC = xlWbookDD_IC.Worksheets("UID IC Distribution Drive (Not ")

                            xlWsheetDD_IC.UsedRange.UnMerge()
                            xlWsheetDD_IC.UsedRange.WrapText = False
                            xlWsheetDD_IC.UsedRange.ColumnWidth = 15
                            xlWsheetDD_IC.UsedRange.RowHeight = 15

                            Dim rg_head_cut1 As Object = xlWsheetDD_IC.Range("B2")
                            Dim rg_head_paste1 As Object = xlWsheetDD_IC.Range("A2")
                            rg_head_cut1.Select()
                            rg_head_cut1.Cut(rg_head_paste1)

                            Dim rg_head_cut2 As Object = xlWsheetDD_IC.Range("B4")
                            Dim rg_head_paste2 As Object = xlWsheetDD_IC.Range("A3")
                            rg_head_cut2.Select()
                            rg_head_cut2.Cut(rg_head_paste2)

                            xlWsheetDD_IC.Range("A2").RowHeight = 27

                            Dim paramhead1, paramhead2, paramhead3 As String
                            paramhead1 = xlWsheetDD_IC.Range("B6").Value & " " & xlWsheetDD_IC.Range("E6").Value & "; "
                            paramhead1 = paramhead1 & xlWsheetDD_IC.Range("I6").Value & " " & xlWsheetDD_IC.Range("L6").Value

                            paramhead2 = xlWsheetDD_IC.Range("R6").Value & " " & xlWsheetDD_IC.Range("U6").Value & ";"
                            paramhead2 = paramhead2 & xlWsheetDD_IC.Range("X6").Value & " " & xlWsheetDD_IC.Range("AA6").Value

                            paramhead3 = xlWsheetDD_IC.Range("B8").Value & " " & xlWsheetDD_IC.Range("E8").Value & "; "
                            paramhead3 = paramhead3 & xlWsheetDD_IC.Range("I8").Value & " " & xlWsheetDD_IC.Range("M8").Value & "; "
                            paramhead3 = paramhead3 & xlWsheetDD_IC.Range("R8").Value & " " & xlWsheetDD_IC.Range("U8").Value


                            xlWsheetDD_IC.Range("A5").Value = paramhead1
                            xlWsheetDD_IC.Range("A5").EntireRow.Font.Name = "Calibri"
                            xlWsheetDD_IC.Range("A6").Value = paramhead2
                            xlWsheetDD_IC.Range("A6").EntireRow.Font.Name = "Calibri"
                            xlWsheetDD_IC.Range("A7").Value = paramhead3
                            xlWsheetDD_IC.Range("A7").EntireRow.Font.Name = "Calibri"

                            xlWsheetDD_IC.Range("B6:AB8").value = ""
                            xlWsheetDD_IC.Range("A9").EntireRow.Delete()

                            Dim xlfunc As Excel.WorksheetFunction
                            xlfunc = xlAppDD_IC.WorksheetFunction()
                            Dim lnCol As Long
                            Dim i As Long
                            Dim rnarea As Object = xlWsheetDD_IC.UsedRange

                            lnCol = rnarea.Columns.Count
                            For i = lnCol To 1 Step -1
                                If xlfunc.CountA(rnarea.Columns(i)) = 0 Then
                                    rnarea.Columns(i).Delete()
                                End If
                            Next

                            xlWsheetDD_IC.Range("A9:B9").Merge()
                            xlWsheetDD_IC.Range("C9:D9").Merge()
                            xlWsheetDD_IC.Range("E9:F9").Merge()
                            xlWsheetDD_IC.Range("L:L").EntireColumn.ColumnWidth = 56

                            xlWbookDD_IC.SaveAs(txtDD_IC_dest.Text)
                            xlWbookDD_IC.Close()
                            xlAppDD_IC.Quit()

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetDD_IC)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookDD_IC)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppDD_IC)

                            xlWsheetDD_IC = Nothing
                            xlWbookDD_IC = Nothing
                            xlAppDD_IC = Nothing

                            MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                        Catch ex As Exception
                            MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try

                End Select

            Case "XL_Installed"
                Dim xlAppDD_IC As Object
                Dim xlWbookDD_IC As Object
                Dim xlWsheetDD_IC As Object
                Select Case DD_type
                    Case "DIST"
                        Try
                            xlAppDD_IC = CreateObject("Excel.Application")
                            xlWbookDD_IC = xlAppDD_IC.Workbooks.Open(txtDD_IC_src.Text)
                            xlWsheetDD_IC = xlWbookDD_IC.Worksheets("UID IC Distribution Drive Repor")

                            xlWsheetDD_IC.UsedRange.UnMerge()
                            xlWsheetDD_IC.UsedRange.WrapText = False
                            xlWsheetDD_IC.UsedRange.ColumnWidth = 15
                            xlWsheetDD_IC.UsedRange.RowHeight = 15

                            Dim rg_head_cut1 As Excel.Range = xlWsheetDD_IC.Range("B2")
                            Dim rg_head_paste1 As Excel.Range = xlWsheetDD_IC.Range("A2")
                            rg_head_cut1.Select()
                            rg_head_cut1.Cut(rg_head_paste1)

                            Dim rg_head_cut2 As Excel.Range = xlWsheetDD_IC.Range("B4")
                            Dim rg_head_paste2 As Excel.Range = xlWsheetDD_IC.Range("A3")
                            rg_head_cut2.Select()
                            rg_head_cut2.Cut(rg_head_paste2)

                            xlWsheetDD_IC.Range("A2").RowHeight = 27

                            Dim paramhead1, paramhead2, paramhead3 As String
                            paramhead1 = xlWsheetDD_IC.Range("B6").Value & " " & xlWsheetDD_IC.Range("E6").Value & "; "
                            paramhead1 = paramhead1 & xlWsheetDD_IC.Range("I6").Value & " " & xlWsheetDD_IC.Range("L6").Value

                            paramhead2 = xlWsheetDD_IC.Range("R6").Value & " " & xlWsheetDD_IC.Range("U6").Value & ";"
                            paramhead2 = paramhead2 & xlWsheetDD_IC.Range("Y6").Value & " " & xlWsheetDD_IC.Range("AB6").Value

                            paramhead3 = xlWsheetDD_IC.Range("B8").Value & " " & xlWsheetDD_IC.Range("E8").Value & "; "
                            paramhead3 = paramhead3 & xlWsheetDD_IC.Range("I8").Value & " " & xlWsheetDD_IC.Range("M8").Value & "; "
                            paramhead3 = paramhead3 & xlWsheetDD_IC.Range("R8").Value & " " & xlWsheetDD_IC.Range("U8").Value
                            

                            xlWsheetDD_IC.Range("A5").Value = paramhead1
                            xlWsheetDD_IC.Range("A5").EntireRow.Font.Name = "Calibri"
                            xlWsheetDD_IC.Range("A6").Value = paramhead2
                            xlWsheetDD_IC.Range("A6").EntireRow.Font.Name = "Calibri"
                            xlWsheetDD_IC.Range("A7").Value = paramhead3
                            xlWsheetDD_IC.Range("A7").EntireRow.Font.Name = "Calibri"

                            xlWsheetDD_IC.Range("B6:AB8").value = ""
                            xlWsheetDD_IC.Range("A10").EntireRow.Delete()

                            Dim xlfunc As Excel.WorksheetFunction
                            xlfunc = xlAppDD_IC.WorksheetFunction()
                            Dim lnCol As Long
                            Dim i As Long
                            Dim rnarea As Excel.Range = xlWsheetDD_IC.UsedRange

                            lnCol = rnarea.Columns.Count
                            For i = lnCol To 1 Step -1
                                If xlfunc.CountA(rnarea.Columns(i)) = 0 Then
                                    rnarea.Columns(i).Delete()
                                End If
                            Next

                            xlWsheetDD_IC.Range("A10:B10").Merge()
                            xlWsheetDD_IC.Range("C10:D10").Merge()
                            xlWsheetDD_IC.Range("E10:F10").Merge()
                            xlWsheetDD_IC.Range("L10:M10").Merge()

                            xlWbookDD_IC.SaveAs(txtDD_IC_dest.Text)
                            xlWbookDD_IC.Close()
                            xlAppDD_IC.Quit()

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetDD_IC)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookDD_IC)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppDD_IC)

                            xlWsheetDD_IC = Nothing
                            xlWbookDD_IC = Nothing
                            xlAppDD_IC = Nothing

                            MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                        Catch ex As Exception
                            MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try

                    Case "NOTDIST"
                        Try
                            xlAppDD_IC = CreateObject("Excel.Application")
                            xlWbookDD_IC = xlAppDD_IC.Workbooks.Open(txtDD_IC_src.Text)
                            xlWsheetDD_IC = xlWbookDD_IC.Worksheets("UID IC Distribution Drive (Not ")

                            xlWsheetDD_IC.UsedRange.UnMerge()
                            xlWsheetDD_IC.UsedRange.WrapText = False
                            xlWsheetDD_IC.UsedRange.ColumnWidth = 15
                            xlWsheetDD_IC.UsedRange.RowHeight = 15

                            Dim rg_head_cut1 As Excel.Range = xlWsheetDD_IC.Range("B2")
                            Dim rg_head_paste1 As Excel.Range = xlWsheetDD_IC.Range("A2")
                            rg_head_cut1.Select()
                            rg_head_cut1.Cut(rg_head_paste1)

                            Dim rg_head_cut2 As Excel.Range = xlWsheetDD_IC.Range("B4")
                            Dim rg_head_paste2 As Excel.Range = xlWsheetDD_IC.Range("A3")
                            rg_head_cut2.Select()
                            rg_head_cut2.Cut(rg_head_paste2)

                            xlWsheetDD_IC.Range("A2").RowHeight = 27

                            Dim paramhead1, paramhead2, paramhead3 As String
                            paramhead1 = xlWsheetDD_IC.Range("B6").Value & " " & xlWsheetDD_IC.Range("E6").Value & "; "
                            paramhead1 = paramhead1 & xlWsheetDD_IC.Range("I6").Value & " " & xlWsheetDD_IC.Range("L6").Value

                            paramhead2 = xlWsheetDD_IC.Range("R6").Value & " " & xlWsheetDD_IC.Range("U6").Value & ";"
                            paramhead2 = paramhead2 & xlWsheetDD_IC.Range("X6").Value & " " & xlWsheetDD_IC.Range("AA6").Value

                            paramhead3 = xlWsheetDD_IC.Range("B8").Value & " " & xlWsheetDD_IC.Range("E8").Value & "; "
                            paramhead3 = paramhead3 & xlWsheetDD_IC.Range("I8").Value & " " & xlWsheetDD_IC.Range("M8").Value & "; "
                            paramhead3 = paramhead3 & xlWsheetDD_IC.Range("R8").Value & " " & xlWsheetDD_IC.Range("U8").Value


                            xlWsheetDD_IC.Range("A5").Value = paramhead1
                            xlWsheetDD_IC.Range("A5").EntireRow.Font.Name = "Calibri"
                            xlWsheetDD_IC.Range("A6").Value = paramhead2
                            xlWsheetDD_IC.Range("A6").EntireRow.Font.Name = "Calibri"
                            xlWsheetDD_IC.Range("A7").Value = paramhead3
                            xlWsheetDD_IC.Range("A7").EntireRow.Font.Name = "Calibri"

                            xlWsheetDD_IC.Range("B6:AB8").value = ""
                            xlWsheetDD_IC.Range("A9").EntireRow.Delete()

                            Dim xlfunc As Excel.WorksheetFunction
                            xlfunc = xlAppDD_IC.WorksheetFunction()
                            Dim lnCol As Long
                            Dim i As Long
                            Dim rnarea As Excel.Range = xlWsheetDD_IC.UsedRange

                            lnCol = rnarea.Columns.Count
                            For i = lnCol To 1 Step -1
                                If xlfunc.CountA(rnarea.Columns(i)) = 0 Then
                                    rnarea.Columns(i).Delete()
                                End If
                            Next

                            xlWsheetDD_IC.Range("A9:B9").Merge()
                            xlWsheetDD_IC.Range("C9:D9").Merge()
                            xlWsheetDD_IC.Range("E9:F9").Merge()
                            xlWsheetDD_IC.Range("L:L").EntireColumn.ColumnWidth = 56

                            xlWbookDD_IC.SaveAs(txtDD_IC_dest.Text)
                            xlWbookDD_IC.Close()
                            xlAppDD_IC.Quit()

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetDD_IC)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookDD_IC)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppDD_IC)

                            xlWsheetDD_IC = Nothing
                            xlWbookDD_IC = Nothing
                            xlAppDD_IC = Nothing

                            MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                        Catch ex As Exception
                            MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try

                End Select

        End Select
    End Sub

    Private Sub BWDD_IC_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BWDD_IC.RunWorkerCompleted
        PicBar_DD_IC.Visible = False
        btnNeu_DD_IC.Enabled = False
        txtDD_IC_dest.Text = ""
        txtDD_IC_src.Text = ""
    End Sub
End Class