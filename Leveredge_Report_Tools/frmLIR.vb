Imports Microsoft.Office.Interop
Imports System.Windows.Forms
Imports System.Text
Imports System.Drawing
Imports Microsoft.Win32
Public Class frmLIR
    Dim AppsOffice As String
    Private Sub frmLIR_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim TargetKey As RegistryKey
        TargetKey = Registry.ClassesRoot.OpenSubKey("Excel.Application")
        If TargetKey Is Nothing Then
            AppsOffice = "XL_NotInstalled"
        Else
            AppsOffice = "XL_Installed"
            TargetKey.Close()
        End If
    End Sub
    Dim LIR_Type As String

    Private Sub RBLIR_Detail_CheckedChanged(sender As Object, e As EventArgs) Handles RBLIR_Detail.CheckedChanged
        LIR_Type = "DTL"
    End Sub

    Private Sub RBLIR_Recap_CheckedChanged(sender As Object, e As EventArgs) Handles RBLIR_Recap.CheckedChanged
        LIR_Type = "RKP"
    End Sub

    Private Sub btnBrow_LIR_src_Click(sender As Object, e As EventArgs) Handles btnBrow_LIR_src.Click
        Dim LIRpathSrc As String
        If OFD_LIR.ShowDialog = DialogResult.OK Then
            LIRpathSrc = OFD_LIR.FileName()
            txtLIR_src.Text = LIRpathSrc
        End If
        If txtLIR_dest.Text <> "" Then
            btnNeu_LIR.Enabled = True
        Else
            btnNeu_LIR.Enabled = False
        End If
    End Sub

    Private Sub btnBrow_LIR_dest_Click(sender As Object, e As EventArgs) Handles btnBrow_LIR_dest.Click
        Dim datenow As DateTime = DateTime.Now
        Dim filename As String = "Neutralize_ListOfInvoice_" & LIR_Type & "_" & datenow.ToString("ddMMyyyy_HHmm")
        SFD_LIR.FileName = filename
        Dim LIRpath_Dest As String
        If SFD_LIR.ShowDialog = DialogResult.OK Then
            LIRpath_Dest = SFD_LIR.FileName
            txtLIR_dest.Text = LIRpath_Dest
        End If
        If txtLIR_src.Text <> "" Then
            btnNeu_LIR.Enabled = True
        Else
            btnNeu_LIR.Enabled = False
        End If
        If txtLIR_dest.Text <> "" Then
            btnNeu_LIR.Enabled = True
        Else
            btnNeu_LIR.Enabled = False
        End If
    End Sub

    Private Sub btnNeu_LIR_Click(sender As Object, e As EventArgs) Handles btnNeu_LIR.Click
        PicBar_LIR.Visible = True
        BWLIR.RunWorkerAsync()
    End Sub

    Private Sub BWLIR_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BWLIR.DoWork
        Select Case AppsOffice
            Case "XL_NotInstalled"
                Dim xlAppLIR As Object
                Dim xlWbookLIR As Object
                Dim xlWsheetLIR As Object

                Select Case LIR_Type
                    Case "DTL"
                        Try
                            xlAppLIR = CreateObject("Ket.Application")
                            xlWbookLIR = xlAppLIR.Workbooks.Open(txtLIR_src.Text)
                            xlWsheetLIR = xlWbookLIR.Worksheets("UID List of Invoice Report")

                            xlWsheetLIR.UsedRange.UnMerge()
                            xlWsheetLIR.UsedRange.WrapText = False
                            xlWsheetLIR.UsedRange.ColumnWidth = 15
                            xlWsheetLIR.UsedRange.RowHeight = 15

                            xlWsheetLIR.Range("A:A").EntireColumn.Delete()

                            Dim rg_head_cut1 As Object = xlWsheetLIR.Range("B2")
                            Dim rg_head_paste1 As Object = xlWsheetLIR.Range("A2")
                            rg_head_cut1.Select()
                            rg_head_cut1.Cut(rg_head_paste1)

                            Dim rg_head_cut2 As Object = xlWsheetLIR.Range("B4")
                            Dim rg_head_paste2 As Object = xlWsheetLIR.Range("A3")
                            rg_head_cut2.Select()
                            rg_head_cut2.Cut(rg_head_paste2)

                            xlWsheetLIR.Range("A2").RowHeight = 27

                            Dim paramhead1, paramhead2, paramhead3 As String
                            paramhead1 = xlWsheetLIR.Range("C6").Value & " " & xlWsheetLIR.Range("F6").Value & "; "
                            paramhead1 = paramhead1 & xlWsheetLIR.Range("I6").Value & " " & xlWsheetLIR.Range("K6").Value

                            paramhead2 = xlWsheetLIR.Range("C8").Value & " " & xlWsheetLIR.Range("F8").Value & "; "
                            paramhead2 = paramhead2 & xlWsheetLIR.Range("I8").Value & " " & xlWsheetLIR.Range("K8").Value

                            paramhead3 = xlWsheetLIR.Range("P8").Value & " " & xlWsheetLIR.Range("S8").Value & "; "
                            paramhead3 = paramhead3 & xlWsheetLIR.Range("X8").Value & " " & xlWsheetLIR.Range("AB8").Value & "; "

                            xlWsheetLIR.Range("A5").Value = paramhead1
                            xlWsheetLIR.Range("A5").EntireRow.Font.Name = "Calibri"
                            xlWsheetLIR.Range("A6").Value = paramhead2
                            xlWsheetLIR.Range("A6").EntireRow.Font.Name = "Calibri"
                            xlWsheetLIR.Range("A7").Value = paramhead3
                            xlWsheetLIR.Range("A7").EntireRow.Font.Name = "Calibri"

                            xlWsheetLIR.range("B5:AB8").value = ""

                            xlWsheetLIR.Range("A9").EntireRow.Delete()

                            Dim xlfunc As Excel.WorksheetFunction
                            xlfunc = xlAppLIR.WorksheetFunction
                            Dim lnCol As Long
                            Dim i As Long
                            Dim rnarea As Object = xlWsheetLIR.UsedRange

                            lnCol = rnarea.Columns.Count
                            For i = lnCol To 1 Step -1
                                If xlfunc.CountA(rnarea.Columns(i)) = 0 Then
                                    rnarea.Columns(i).Delete()
                                End If
                            Next

                            Dim lrow As Long
                            lrow = xlWsheetLIR.Range("B65536").End(Excel.XlDirection.xlUp).Row
                            Dim g As String = "B" & lrow
                            Dim cr As String = g & ":R" & lrow
                            Dim cellrange As Object = xlWsheetLIR.Range(cr)
                            cellrange.Select()
                            cellrange.Merge()
                            cellrange.HorizontalAlignment = 4

                            Dim lrow2 As Long
                            lrow2 = xlWsheetLIR.Range("B65536").End(Excel.XlDirection.xlUp).Row
                            lrow2 = lrow2 - 1
                            Dim g1 As String = "B" & lrow2
                            Dim cr2 As String = g1 & ":R" & lrow2
                            Dim cellrange2 As Object = xlWsheetLIR.Range(cr2)
                            cellrange2.Select()
                            cellrange2.Merge()
                            cellrange2.HorizontalAlignment = 4

                            Dim lrow3 As Long
                            lrow3 = xlWsheetLIR.Range("B65536").End(Excel.XlDirection.xlUp).Row
                            lrow3 = lrow3 - 2
                            Dim g2 As String = "B" & lrow3
                            Dim cr3 As String = g2 & ":R" & lrow3
                            Dim cellrange3 As Object = xlWsheetLIR.Range(cr3)
                            cellrange3.Select()
                            cellrange3.Merge()
                            cellrange3.HorizontalAlignment = 4

                            Dim lrow4 As Long
                            lrow4 = xlWsheetLIR.Range("B65536").End(Excel.XlDirection.xlUp).Row
                            lrow4 = lrow4 - 3
                            Dim g3 As String = "B" & lrow4
                            Dim cr4 As String = g3 & ":R" & lrow4
                            Dim cellrange4 As Object = xlWsheetLIR.Range(cr4)
                            cellrange4.Select()
                            cellrange4.Merge()
                            cellrange4.HorizontalAlignment = 4

                            Dim lrow5 As Long
                            lrow5 = xlWsheetLIR.Range("B65536").End(Excel.XlDirection.xlUp).Row
                            lrow5 = lrow5 - 4
                            Dim g4 As String = "B" & lrow5
                            Dim cr5 As String = g4 & ":L" & lrow5
                            Dim cellrange5 As Object = xlWsheetLIR.Range(cr5)
                            cellrange5.Select()
                            cellrange5.Merge()
                            cellrange5.HorizontalAlignment = 4

                            xlWbookLIR.SaveAs(txtLIR_dest.Text)
                            xlWbookLIR.Close()
                            xlAppLIR.Quit()

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetLIR)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookLIR)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppLIR)

                            xlWsheetLIR = Nothing
                            xlWbookLIR = Nothing
                            xlAppLIR = Nothing

                            MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                        Catch ex As Exception
                            MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try

                    Case "RKP"
                        Try
                            xlAppLIR = CreateObject("Ket.Application")
                            xlWbookLIR = xlAppLIR.Workbooks.Open(txtLIR_src.Text)
                            xlWsheetLIR = xlWbookLIR.Worksheets("UID List of Invoice Report")

                            xlWsheetLIR.UsedRange.UnMerge()
                            xlWsheetLIR.UsedRange.WrapText = False
                            xlWsheetLIR.UsedRange.ColumnWidth = 15
                            xlWsheetLIR.UsedRange.RowHeight = 15

                            xlWsheetLIR.Range("A:A").EntireColumn.Delete()
                            xlWsheetLIR.Range("A10").EntireRow.Delete()

                            Dim rg_head_cut1 As Object = xlWsheetLIR.Range("B2")
                            Dim rg_head_paste1 As Object = xlWsheetLIR.Range("A2")
                            rg_head_cut1.Select()
                            rg_head_cut1.Cut(rg_head_paste1)

                            Dim rg_head_cut2 As Object = xlWsheetLIR.Range("B4")
                            Dim rg_head_paste2 As Object = xlWsheetLIR.Range("A3")
                            rg_head_cut2.Select()
                            rg_head_cut2.Cut(rg_head_paste2)

                            xlWsheetLIR.Range("A2").RowHeight = 27

                            Dim paramhead1, paramhead2, paramhead3 As String
                            paramhead1 = xlWsheetLIR.Range("C6").Value & " " & xlWsheetLIR.Range("F6").Value & "; "
                            paramhead1 = paramhead1 & xlWsheetLIR.Range("I6").Value & " " & xlWsheetLIR.Range("K6").Value

                            paramhead2 = xlWsheetLIR.Range("C8").Value & " " & xlWsheetLIR.Range("F8").Value & "; "
                            paramhead2 = paramhead2 & xlWsheetLIR.Range("I8").Value & " " & xlWsheetLIR.Range("K8").Value

                            paramhead3 = xlWsheetLIR.Range("O8").Value & " " & xlWsheetLIR.Range("R8").Value & "; "
                            paramhead3 = paramhead3 & xlWsheetLIR.Range("W8").Value & " " & xlWsheetLIR.Range("Z8").Value

                            xlWsheetLIR.Range("A5").Value = paramhead1
                            xlWsheetLIR.Range("A5").EntireRow.Font.Name = "Calibri"
                            xlWsheetLIR.Range("A6").Value = paramhead2
                            xlWsheetLIR.Range("A6").EntireRow.Font.Name = "Calibri"
                            xlWsheetLIR.Range("A7").Value = paramhead3
                            xlWsheetLIR.Range("A7").EntireRow.Font.Name = "Calibri"

                            xlWsheetLIR.range("B5:Z8").value = ""

                            Dim xlfunc As Excel.WorksheetFunction
                            xlfunc = xlAppLIR.WorksheetFunction
                            Dim lnCol As Long
                            Dim i As Long
                            Dim rnarea As Object = xlWsheetLIR.UsedRange

                            lnCol = rnarea.Columns.Count
                            For i = lnCol To 1 Step -1
                                If xlfunc.CountA(rnarea.Columns(i)) = 0 Then
                                    rnarea.Columns(i).Delete()
                                End If
                            Next

                            Dim lrow As Long
                            lrow = xlWsheetLIR.Range("B65536").End(Excel.XlDirection.xlUp).Row
                            Dim g As String = "B" & lrow
                            Dim cr As String = g & ":J" & lrow
                            Dim cellrange As Object = xlWsheetLIR.Range(cr)
                            cellrange.Select()
                            cellrange.Merge()
                            cellrange.HorizontalAlignment = 4

                            xlWbookLIR.SaveAs(txtLIR_dest.Text)
                            xlWbookLIR.Close()
                            xlAppLIR.Quit()

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetLIR)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookLIR)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppLIR)

                            xlWsheetLIR = Nothing
                            xlWbookLIR = Nothing
                            xlAppLIR = Nothing

                            MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                        Catch ex As Exception
                            MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try

                End Select

            Case "XL_Installed"
                Dim xlAppLIR As Object
                Dim xlWbookLIR As Object
                Dim xlWsheetLIR As Object

                Select Case LIR_Type
                    Case "DTL"
                        Try
                            xlAppLIR = CreateObject("Excel.Application")
                            xlWbookLIR = xlAppLIR.Workbooks.Open(txtLIR_src.Text)
                            xlWsheetLIR = xlWbookLIR.Worksheets("UID List of Invoice Report")

                            xlWsheetLIR.UsedRange.UnMerge()
                            xlWsheetLIR.UsedRange.WrapText = False
                            xlWsheetLIR.UsedRange.ColumnWidth = 15
                            xlWsheetLIR.UsedRange.RowHeight = 15

                            xlWsheetLIR.Range("A:A").EntireColumn.Delete()

                            Dim rg_head_cut1 As Excel.Range = xlWsheetLIR.Range("B2")
                            Dim rg_head_paste1 As Excel.Range = xlWsheetLIR.Range("A2")
                            rg_head_cut1.Select()
                            rg_head_cut1.Cut(rg_head_paste1)

                            Dim rg_head_cut2 As Excel.Range = xlWsheetLIR.Range("B4")
                            Dim rg_head_paste2 As Excel.Range = xlWsheetLIR.Range("A3")
                            rg_head_cut2.Select()
                            rg_head_cut2.Cut(rg_head_paste2)

                            xlWsheetLIR.Range("A2").RowHeight = 27

                            Dim paramhead1, paramhead2, paramhead3 As String
                            paramhead1 = xlWsheetLIR.Range("C6").Value & " " & xlWsheetLIR.Range("F6").Value & "; "
                            paramhead1 = paramhead1 & xlWsheetLIR.Range("I6").Value & " " & xlWsheetLIR.Range("K6").Value

                            paramhead2 = xlWsheetLIR.Range("C8").Value & " " & xlWsheetLIR.Range("F8").Value & "; "
                            paramhead2 = paramhead2 & xlWsheetLIR.Range("I8").Value & " " & xlWsheetLIR.Range("K8").Value

                            paramhead3 = xlWsheetLIR.Range("P8").Value & " " & xlWsheetLIR.Range("S8").Value & "; "
                            paramhead3 = paramhead3 & xlWsheetLIR.Range("X8").Value & " " & xlWsheetLIR.Range("AB8").Value & "; "

                            xlWsheetLIR.Range("A5").Value = paramhead1
                            xlWsheetLIR.Range("A5").EntireRow.Font.Name = "Calibri"
                            xlWsheetLIR.Range("A6").Value = paramhead2
                            xlWsheetLIR.Range("A6").EntireRow.Font.Name = "Calibri"
                            xlWsheetLIR.Range("A7").Value = paramhead3
                            xlWsheetLIR.Range("A7").EntireRow.Font.Name = "Calibri"

                            xlWsheetLIR.range("B5:AB8").value = ""

                            xlWsheetLIR.Range("A9").EntireRow.Delete()

                            Dim xlfunc As Excel.WorksheetFunction
                            xlfunc = xlAppLIR.WorksheetFunction
                            Dim lnCol As Long
                            Dim i As Long
                            Dim rnarea As Excel.Range = xlWsheetLIR.UsedRange

                            lnCol = rnarea.Columns.Count
                            For i = lnCol To 1 Step -1
                                If xlfunc.CountA(rnarea.Columns(i)) = 0 Then
                                    rnarea.Columns(i).Delete()
                                End If
                            Next

                            Dim lrow As Long
                            lrow = xlWsheetLIR.Range("B65536").End(Excel.XlDirection.xlUp).Row
                            Dim g As String = "B" & lrow
                            Dim cr As String = g & ":R" & lrow
                            Dim cellrange As Excel.Range = xlWsheetLIR.Range(cr)
                            cellrange.Select()
                            cellrange.Merge()
                            cellrange.HorizontalAlignment = 4

                            Dim lrow2 As Long
                            lrow2 = xlWsheetLIR.Range("B65536").End(Excel.XlDirection.xlUp).Row
                            lrow2 = lrow2 - 1
                            Dim g1 As String = "B" & lrow2
                            Dim cr2 As String = g1 & ":R" & lrow2
                            Dim cellrange2 As Excel.Range = xlWsheetLIR.Range(cr2)
                            cellrange2.Select()
                            cellrange2.Merge()
                            cellrange2.HorizontalAlignment = 4

                            Dim lrow3 As Long
                            lrow3 = xlWsheetLIR.Range("B65536").End(Excel.XlDirection.xlUp).Row
                            lrow3 = lrow3 - 2
                            Dim g2 As String = "B" & lrow3
                            Dim cr3 As String = g2 & ":R" & lrow3
                            Dim cellrange3 As Excel.Range = xlWsheetLIR.Range(cr3)
                            cellrange3.Select()
                            cellrange3.Merge()
                            cellrange3.HorizontalAlignment = 4

                            Dim lrow4 As Long
                            lrow4 = xlWsheetLIR.Range("B65536").End(Excel.XlDirection.xlUp).Row
                            lrow4 = lrow4 - 3
                            Dim g3 As String = "B" & lrow4
                            Dim cr4 As String = g3 & ":R" & lrow4
                            Dim cellrange4 As Excel.Range = xlWsheetLIR.Range(cr4)
                            cellrange4.Select()
                            cellrange4.Merge()
                            cellrange4.HorizontalAlignment = 4

                            Dim lrow5 As Long
                            lrow5 = xlWsheetLIR.Range("B65536").End(Excel.XlDirection.xlUp).Row
                            lrow5 = lrow5 - 4
                            Dim g4 As String = "B" & lrow5
                            Dim cr5 As String = g4 & ":L" & lrow5
                            Dim cellrange5 As Excel.Range = xlWsheetLIR.Range(cr5)
                            cellrange5.Select()
                            cellrange5.Merge()
                            cellrange5.HorizontalAlignment = 4

                            xlWbookLIR.SaveAs(txtLIR_dest.Text)
                            xlWbookLIR.Close()
                            xlAppLIR.Quit()

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetLIR)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookLIR)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppLIR)

                            xlWsheetLIR = Nothing
                            xlWbookLIR = Nothing
                            xlAppLIR = Nothing

                            MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                        Catch ex As Exception
                            MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try

                    Case "RKP"
                        Try
                            xlAppLIR = CreateObject("Excel.Application")
                            xlWbookLIR = xlAppLIR.Workbooks.Open(txtLIR_src.Text)
                            xlWsheetLIR = xlWbookLIR.Worksheets("UID List of Invoice Report")

                            xlWsheetLIR.UsedRange.UnMerge()
                            xlWsheetLIR.UsedRange.WrapText = False
                            xlWsheetLIR.UsedRange.ColumnWidth = 15
                            xlWsheetLIR.UsedRange.RowHeight = 15

                            xlWsheetLIR.Range("A:A").EntireColumn.Delete()
                            xlWsheetLIR.Range("A10").EntireRow.Delete()

                            Dim rg_head_cut1 As Excel.Range = xlWsheetLIR.Range("B2")
                            Dim rg_head_paste1 As Excel.Range = xlWsheetLIR.Range("A2")
                            rg_head_cut1.Select()
                            rg_head_cut1.Cut(rg_head_paste1)

                            Dim rg_head_cut2 As Excel.Range = xlWsheetLIR.Range("B4")
                            Dim rg_head_paste2 As Excel.Range = xlWsheetLIR.Range("A3")
                            rg_head_cut2.Select()
                            rg_head_cut2.Cut(rg_head_paste2)

                            xlWsheetLIR.Range("A2").RowHeight = 27

                            Dim paramhead1, paramhead2, paramhead3 As String
                            paramhead1 = xlWsheetLIR.Range("C6").Value & " " & xlWsheetLIR.Range("F6").Value & "; "
                            paramhead1 = paramhead1 & xlWsheetLIR.Range("I6").Value & " " & xlWsheetLIR.Range("K6").Value

                            paramhead2 = xlWsheetLIR.Range("C8").Value & " " & xlWsheetLIR.Range("F8").Value & "; "
                            paramhead2 = paramhead2 & xlWsheetLIR.Range("I8").Value & " " & xlWsheetLIR.Range("K8").Value

                            paramhead3 = xlWsheetLIR.Range("O8").Value & " " & xlWsheetLIR.Range("R8").Value & "; "
                            paramhead3 = paramhead3 & xlWsheetLIR.Range("W8").Value & " " & xlWsheetLIR.Range("Z8").Value

                            xlWsheetLIR.Range("A5").Value = paramhead1
                            xlWsheetLIR.Range("A5").EntireRow.Font.Name = "Calibri"
                            xlWsheetLIR.Range("A6").Value = paramhead2
                            xlWsheetLIR.Range("A6").EntireRow.Font.Name = "Calibri"
                            xlWsheetLIR.Range("A7").Value = paramhead3
                            xlWsheetLIR.Range("A7").EntireRow.Font.Name = "Calibri"

                            xlWsheetLIR.range("B5:Z8").value = ""

                            Dim xlfunc As Excel.WorksheetFunction
                            xlfunc = xlAppLIR.WorksheetFunction
                            Dim lnCol As Long
                            Dim i As Long
                            Dim rnarea As Excel.Range = xlWsheetLIR.UsedRange

                            lnCol = rnarea.Columns.Count
                            For i = lnCol To 1 Step -1
                                If xlfunc.CountA(rnarea.Columns(i)) = 0 Then
                                    rnarea.Columns(i).Delete()
                                End If
                            Next

                            Dim lrow As Long
                            lrow = xlWsheetLIR.Range("B65536").End(Excel.XlDirection.xlUp).Row
                            Dim g As String = "B" & lrow
                            Dim cr As String = g & ":J" & lrow
                            Dim cellrange As Excel.Range = xlWsheetLIR.Range(cr)
                            cellrange.Select()
                            cellrange.Merge()
                            cellrange.HorizontalAlignment = 4

                            xlWbookLIR.SaveAs(txtLIR_dest.Text)
                            xlWbookLIR.Close()
                            xlAppLIR.Quit()

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetLIR)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookLIR)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppLIR)

                            xlWsheetLIR = Nothing
                            xlWbookLIR = Nothing
                            xlAppLIR = Nothing

                            MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                        Catch ex As Exception
                            MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try
                        
                End Select

        End Select
    End Sub

    Private Sub BWLIR_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BWLIR.RunWorkerCompleted
        PicBar_LIR.Visible = False
        btnNeu_LIR.Enabled = False
        txtLIR_dest.Text = ""
        txtLIR_src.Text = ""
    End Sub
End Class