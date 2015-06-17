Imports Microsoft.Office.Interop
Imports System.Windows.Forms
Imports System.Text
Imports System.Drawing
Imports Microsoft.Win32
Public Class frmONI_IC
    Dim AppsOffice As String
    Private Sub frmONI_IC_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim TargetKey As RegistryKey
        TargetKey = Registry.ClassesRoot.OpenSubKey("Excel.Application")
        If TargetKey Is Nothing Then
            AppsOffice = "XL_NotInstalled"
        Else
            AppsOffice = "XL_Installed"
            TargetKey.Close()
        End If
    End Sub
    Dim NIR_Type As String

    Private Sub RBNIR_Detail_CheckedChanged(sender As Object, e As EventArgs) Handles RBONI_Detail.CheckedChanged
        NIR_Type = "DTL"
    End Sub

    Private Sub RBNIR_Summary_CheckedChanged(sender As Object, e As EventArgs) Handles RBONI_Summary.CheckedChanged
        NIR_Type = "SUM"
    End Sub

    Private Sub btnBrow_ONI_IC_src_Click(sender As Object, e As EventArgs) Handles btnBrow_ONI_IC_src.Click
        Dim ONIpathSrc As String
        If OFD_ONI_IC.ShowDialog = DialogResult.OK Then
            ONIpathSrc = OFD_ONI_IC.FileName()
            txtONI_IC_src.Text = ONIpathSrc
        End If
        If txtONI_IC_dest.Text <> "" Then
            btnNeu_ONI_IC.Enabled = True
        Else
            btnNeu_ONI_IC.Enabled = False
        End If
    End Sub

    Private Sub btnBrow_ONI_IC_dest_Click(sender As Object, e As EventArgs) Handles btnBrow_ONI_IC_dest.Click
        Dim datenow As DateTime = DateTime.Now
        Dim filename As String = "Neutralize_IC_NetIncrease_" & NIR_Type & "_" & datenow.ToString("ddMMyyyy_HHmm")
        SFD_ONI_IC.FileName = filename
        Dim ONIpath_Dest As String
        If SFD_ONI_IC.ShowDialog = DialogResult.OK Then
            ONIpath_Dest = SFD_ONI_IC.FileName
            txtONI_IC_dest.Text = ONIpath_Dest
        End If
        If txtONI_IC_src.Text <> "" Then
            btnNeu_ONI_IC.Enabled = True
        Else
            btnNeu_ONI_IC.Enabled = False
        End If
        If txtONI_IC_dest.Text <> "" Then
            btnNeu_ONI_IC.Enabled = True
        Else
            btnNeu_ONI_IC.Enabled = False
        End If
    End Sub

    Private Sub btnNeu_ONI_IC_Click(sender As Object, e As EventArgs) Handles btnNeu_ONI_IC.Click
        PicBar_ONI_IC.Visible = True
        BWONI_IC.RunWorkerAsync()
    End Sub

    Private Sub BWONI_IC_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BWONI_IC.DoWork
        Select Case AppsOffice
            Case "XL_NotInstalled"
                Dim xlAppONI_IC As Object
                Dim xlWbookONI_IC As Object
                Dim xlWsheetONI_IC As Object
                Select Case NIR_Type
                    Case "DTL"
                        Try
                            xlAppONI_IC = CreateObject("Ket.Application")
                            xlWbookONI_IC = xlAppONI_IC.Workbooks.Open(txtONI_IC_src.Text)
                            xlWsheetONI_IC = xlWbookONI_IC.Worksheets("UID IC Outlet Net Increase Deta")

                            xlWsheetONI_IC.UsedRange.UnMerge()
                            xlWsheetONI_IC.UsedRange.WrapText = False
                            xlWsheetONI_IC.UsedRange.ColumnWidth = 15
                            xlWsheetONI_IC.UsedRange.RowHeight = 15

                            Dim rg_head_cut1 As Object = xlWsheetONI_IC.Range("B2")
                            Dim rg_head_paste1 As Object = xlWsheetONI_IC.Range("A2")
                            rg_head_cut1.Select()
                            rg_head_cut1.Cut(rg_head_paste1)

                            Dim rg_head_cut2 As Object = xlWsheetONI_IC.Range("B4")
                            Dim rg_head_paste2 As Object = xlWsheetONI_IC.Range("A3")
                            rg_head_cut2.Select()
                            rg_head_cut2.Cut(rg_head_paste2)

                            xlWsheetONI_IC.Range("A2").RowHeight = 27

                            Dim paramhead1, paramhead2, paramhead3 As String
                            paramhead1 = xlWsheetONI_IC.Range("B6").Value & " " & xlWsheetONI_IC.Range("E6").Value & "; "
                            paramhead1 = paramhead1 & xlWsheetONI_IC.Range("H6").Value & " " & xlWsheetONI_IC.Range("K6").Value

                            paramhead2 = xlWsheetONI_IC.Range("N6").Value & " " & xlWsheetONI_IC.Range("R6").Value & ";"
                            paramhead2 = paramhead2 & xlWsheetONI_IC.Range("B8").Value & " " & xlWsheetONI_IC.Range("E8").Value & "; "
                            paramhead2 = paramhead2 & xlWsheetONI_IC.Range("H8").Value & " " & xlWsheetONI_IC.Range("K8").Value


                            paramhead3 = xlWsheetONI_IC.Range("N8").Value & " " & xlWsheetONI_IC.Range("R8").Value & "; "
                            paramhead3 = paramhead3 & xlWsheetONI_IC.Range("U8").Value & " " & xlWsheetONI_IC.Range("X8").Value

                            xlWsheetONI_IC.Range("A5").Value = paramhead1
                            xlWsheetONI_IC.Range("A5").EntireRow.Font.Name = "Calibri"
                            xlWsheetONI_IC.Range("A6").Value = paramhead2
                            xlWsheetONI_IC.Range("A6").EntireRow.Font.Name = "Calibri"
                            xlWsheetONI_IC.Range("A7").Value = paramhead3
                            xlWsheetONI_IC.Range("A7").EntireRow.Font.Name = "Calibri"

                            xlWsheetONI_IC.Range("B6:AA8").value = ""
                            xlWsheetONI_IC.Range("A9").EntireRow.Delete()

                            Dim xlfunc As Excel.WorksheetFunction
                            xlfunc = xlAppONI_IC.WorksheetFunction()
                            Dim lnCol As Long
                            Dim i As Long
                            Dim rnarea As Object = xlWsheetONI_IC.UsedRange

                            lnCol = rnarea.Columns.Count
                            For i = lnCol To 1 Step -1
                                If xlfunc.CountA(rnarea.Columns(i)) = 0 Then
                                    rnarea.Columns(i).Delete()
                                End If
                            Next

                            xlWbookONI_IC.SaveAs(txtONI_IC_dest.Text)
                            xlWbookONI_IC.Close()
                            xlAppONI_IC.Quit()

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetONI_IC)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookONI_IC)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppONI_IC)

                            xlWsheetONI_IC = Nothing
                            xlWbookONI_IC = Nothing
                            xlAppONI_IC = Nothing

                            MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                        Catch ex As Exception
                            MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try

                    Case "SUM"
                        Try
                            xlAppONI_IC = CreateObject("Ket.Application")
                            xlWbookONI_IC = xlAppONI_IC.Workbooks.Open(txtONI_IC_src.Text)
                            xlWsheetONI_IC = xlWbookONI_IC.Worksheets("UID IC Outlet Net Increase Summ")

                            xlWsheetONI_IC.UsedRange.UnMerge()
                            xlWsheetONI_IC.UsedRange.WrapText = False
                            xlWsheetONI_IC.UsedRange.ColumnWidth = 15
                            xlWsheetONI_IC.UsedRange.RowHeight = 15

                            Dim rg_head_cut1 As Object = xlWsheetONI_IC.Range("B2")
                            Dim rg_head_paste1 As Object = xlWsheetONI_IC.Range("A2")
                            rg_head_cut1.Select()
                            rg_head_cut1.Cut(rg_head_paste1)

                            Dim rg_head_cut2 As Object = xlWsheetONI_IC.Range("B4")
                            Dim rg_head_paste2 As Object = xlWsheetONI_IC.Range("A3")
                            rg_head_cut2.Select()
                            rg_head_cut2.Cut(rg_head_paste2)

                            xlWsheetONI_IC.Range("A2").RowHeight = 27

                            Dim paramhead1, paramhead2 As String
                            paramhead1 = xlWsheetONI_IC.Range("B6").Value & " " & xlWsheetONI_IC.Range("D6").Value & "; "
                            paramhead1 = paramhead1 & xlWsheetONI_IC.Range("G6").Value & " " & xlWsheetONI_IC.Range("I6").Value & "; "
                            paramhead1 = paramhead1 & xlWsheetONI_IC.Range("N6").Value & " " & xlWsheetONI_IC.Range("Q6").Value

                            paramhead2 = xlWsheetONI_IC.Range("B8").Value & " " & xlWsheetONI_IC.Range("D8").Value & ";"
                            paramhead2 = paramhead2 & xlWsheetONI_IC.Range("G8").Value & " " & xlWsheetONI_IC.Range("I8").Value & ";"
                            paramhead2 = paramhead2 & xlWsheetONI_IC.Range("N8").Value & " " & xlWsheetONI_IC.Range("Q8").Value & ";"
                            paramhead2 = paramhead2 & xlWsheetONI_IC.Range("U8").Value & " " & xlWsheetONI_IC.Range("X8").Value & ";"

                            xlWsheetONI_IC.Range("A5").Value = paramhead1
                            xlWsheetONI_IC.Range("A5").EntireRow.Font.Name = "Calibri"
                            xlWsheetONI_IC.Range("A6").Value = paramhead2
                            xlWsheetONI_IC.Range("A6").EntireRow.Font.Name = "Calibri"

                            xlWsheetONI_IC.Range("B6:AA8").value = ""
                            xlWsheetONI_IC.Range("A8:A9").EntireRow.Delete()

                            Dim xlfunc As Excel.WorksheetFunction
                            xlfunc = xlAppONI_IC.WorksheetFunction()
                            Dim lnCol As Long
                            Dim i As Long
                            Dim rnarea As Object = xlWsheetONI_IC.UsedRange

                            lnCol = rnarea.Columns.Count
                            For i = lnCol To 1 Step -1
                                If xlfunc.CountA(rnarea.Columns(i)) = 0 Then
                                    rnarea.Columns(i).Delete()
                                End If
                            Next

                            xlWbookONI_IC.SaveAs(txtONI_IC_dest.Text)
                            xlWbookONI_IC.Close()
                            xlAppONI_IC.Quit()

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetONI_IC)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookONI_IC)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppONI_IC)

                            xlWsheetONI_IC = Nothing
                            xlWbookONI_IC = Nothing
                            xlAppONI_IC = Nothing

                            MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                        Catch ex As Exception
                            MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try
                End Select

            Case "XL_Installed"
                Dim xlAppONI_IC As Object
                Dim xlWbookONI_IC As Object
                Dim xlWsheetONI_IC As Object
                Select Case NIR_Type
                    Case "DTL"
                        Try
                            xlAppONI_IC = CreateObject("Excel.Application")
                            xlWbookONI_IC = xlAppONI_IC.Workbooks.Open(txtONI_IC_src.Text)
                            xlWsheetONI_IC = xlWbookONI_IC.Worksheets("UID IC Outlet Net Increase Deta")

                            xlWsheetONI_IC.UsedRange.UnMerge()
                            xlWsheetONI_IC.UsedRange.WrapText = False
                            xlWsheetONI_IC.UsedRange.ColumnWidth = 15
                            xlWsheetONI_IC.UsedRange.RowHeight = 15

                            Dim rg_head_cut1 As Excel.Range = xlWsheetONI_IC.Range("B2")
                            Dim rg_head_paste1 As Excel.Range = xlWsheetONI_IC.Range("A2")
                            rg_head_cut1.Select()
                            rg_head_cut1.Cut(rg_head_paste1)

                            Dim rg_head_cut2 As Excel.Range = xlWsheetONI_IC.Range("B4")
                            Dim rg_head_paste2 As Excel.Range = xlWsheetONI_IC.Range("A3")
                            rg_head_cut2.Select()
                            rg_head_cut2.Cut(rg_head_paste2)

                            xlWsheetONI_IC.Range("A2").RowHeight = 27

                            Dim paramhead1, paramhead2, paramhead3 As String
                            paramhead1 = xlWsheetONI_IC.Range("B6").Value & " " & xlWsheetONI_IC.Range("E6").Value & "; "
                            paramhead1 = paramhead1 & xlWsheetONI_IC.Range("H6").Value & " " & xlWsheetONI_IC.Range("K6").Value

                            paramhead2 = xlWsheetONI_IC.Range("N6").Value & " " & xlWsheetONI_IC.Range("R6").Value & ";"
                            paramhead2 = paramhead2 & xlWsheetONI_IC.Range("B8").Value & " " & xlWsheetONI_IC.Range("E8").Value & "; "
                            paramhead2 = paramhead2 & xlWsheetONI_IC.Range("H8").Value & " " & xlWsheetONI_IC.Range("K8").Value


                            paramhead3 = xlWsheetONI_IC.Range("N8").Value & " " & xlWsheetONI_IC.Range("R8").Value & "; "
                            paramhead3 = paramhead3 & xlWsheetONI_IC.Range("U8").Value & " " & xlWsheetONI_IC.Range("X8").Value

                            xlWsheetONI_IC.Range("A5").Value = paramhead1
                            xlWsheetONI_IC.Range("A5").EntireRow.Font.Name = "Calibri"
                            xlWsheetONI_IC.Range("A6").Value = paramhead2
                            xlWsheetONI_IC.Range("A6").EntireRow.Font.Name = "Calibri"
                            xlWsheetONI_IC.Range("A7").Value = paramhead3
                            xlWsheetONI_IC.Range("A7").EntireRow.Font.Name = "Calibri"

                            xlWsheetONI_IC.Range("B6:AA8").value = ""
                            xlWsheetONI_IC.Range("A9").EntireRow.Delete()

                            Dim xlfunc As Excel.WorksheetFunction
                            xlfunc = xlAppONI_IC.WorksheetFunction()
                            Dim lnCol As Long
                            Dim i As Long
                            Dim rnarea As Excel.Range = xlWsheetONI_IC.UsedRange

                            lnCol = rnarea.Columns.Count
                            For i = lnCol To 1 Step -1
                                If xlfunc.CountA(rnarea.Columns(i)) = 0 Then
                                    rnarea.Columns(i).Delete()
                                End If
                            Next

                            xlWbookONI_IC.SaveAs(txtONI_IC_dest.Text)
                            xlWbookONI_IC.Close()
                            xlAppONI_IC.Quit()

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetONI_IC)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookONI_IC)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppONI_IC)

                            xlWsheetONI_IC = Nothing
                            xlWbookONI_IC = Nothing
                            xlAppONI_IC = Nothing

                            MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                        Catch ex As Exception
                            MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try

                    Case "SUM"
                        Try
                            xlAppONI_IC = CreateObject("Excel.Application")
                            xlWbookONI_IC = xlAppONI_IC.Workbooks.Open(txtONI_IC_src.Text)
                            xlWsheetONI_IC = xlWbookONI_IC.Worksheets("UID IC Outlet Net Increase Summ")

                            xlWsheetONI_IC.UsedRange.UnMerge()
                            xlWsheetONI_IC.UsedRange.WrapText = False
                            xlWsheetONI_IC.UsedRange.ColumnWidth = 15
                            xlWsheetONI_IC.UsedRange.RowHeight = 15

                            Dim rg_head_cut1 As Excel.Range = xlWsheetONI_IC.Range("B2")
                            Dim rg_head_paste1 As Excel.Range = xlWsheetONI_IC.Range("A2")
                            rg_head_cut1.Select()
                            rg_head_cut1.Cut(rg_head_paste1)

                            Dim rg_head_cut2 As Excel.Range = xlWsheetONI_IC.Range("B4")
                            Dim rg_head_paste2 As Excel.Range = xlWsheetONI_IC.Range("A3")
                            rg_head_cut2.Select()
                            rg_head_cut2.Cut(rg_head_paste2)

                            xlWsheetONI_IC.Range("A2").RowHeight = 27

                            Dim paramhead1, paramhead2 As String
                            paramhead1 = xlWsheetONI_IC.Range("B6").Value & " " & xlWsheetONI_IC.Range("D6").Value & "; "
                            paramhead1 = paramhead1 & xlWsheetONI_IC.Range("G6").Value & " " & xlWsheetONI_IC.Range("I6").Value & "; "
                            paramhead1 = paramhead1 & xlWsheetONI_IC.Range("N6").Value & " " & xlWsheetONI_IC.Range("Q6").Value

                            paramhead2 = xlWsheetONI_IC.Range("B8").Value & " " & xlWsheetONI_IC.Range("D8").Value & ";"
                            paramhead2 = paramhead2 & xlWsheetONI_IC.Range("G8").Value & " " & xlWsheetONI_IC.Range("I8").Value & ";"
                            paramhead2 = paramhead2 & xlWsheetONI_IC.Range("N8").Value & " " & xlWsheetONI_IC.Range("Q8").Value & ";"
                            paramhead2 = paramhead2 & xlWsheetONI_IC.Range("U8").Value & " " & xlWsheetONI_IC.Range("X8").Value & ";"

                            xlWsheetONI_IC.Range("A5").Value = paramhead1
                            xlWsheetONI_IC.Range("A5").EntireRow.Font.Name = "Calibri"
                            xlWsheetONI_IC.Range("A6").Value = paramhead2
                            xlWsheetONI_IC.Range("A6").EntireRow.Font.Name = "Calibri"

                            xlWsheetONI_IC.Range("B6:AA8").value = ""
                            xlWsheetONI_IC.Range("A8:A9").EntireRow.Delete()

                            Dim xlfunc As Excel.WorksheetFunction
                            xlfunc = xlAppONI_IC.WorksheetFunction()
                            Dim lnCol As Long
                            Dim i As Long
                            Dim rnarea As Excel.Range = xlWsheetONI_IC.UsedRange

                            lnCol = rnarea.Columns.Count
                            For i = lnCol To 1 Step -1
                                If xlfunc.CountA(rnarea.Columns(i)) = 0 Then
                                    rnarea.Columns(i).Delete()
                                End If
                            Next

                            xlWbookONI_IC.SaveAs(txtONI_IC_dest.Text)
                            xlWbookONI_IC.Close()
                            xlAppONI_IC.Quit()

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetONI_IC)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookONI_IC)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppONI_IC)

                            xlWsheetONI_IC = Nothing
                            xlWbookONI_IC = Nothing
                            xlAppONI_IC = Nothing

                            MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                        Catch ex As Exception
                            MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try
                End Select
        End Select

    End Sub

    Private Sub BWONI_IC_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BWONI_IC.RunWorkerCompleted
        PicBar_ONI_IC.Visible = False
        btnNeu_ONI_IC.Enabled = False
        txtONI_IC_dest.Text = ""
        txtONI_IC_src.Text = ""
    End Sub

End Class