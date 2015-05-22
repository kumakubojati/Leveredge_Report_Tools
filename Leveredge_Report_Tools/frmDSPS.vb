Imports Microsoft.Office.Interop
Imports System.Windows.Forms
Imports System.Text
Imports System.Drawing
Imports Microsoft.Win32

Public Class frmDSPS
    Dim AppsOffice As String
    Private Sub frmDSPS_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim TargetKey As RegistryKey
        TargetKey = Registry.ClassesRoot.OpenSubKey("Excel.Application")
        If TargetKey Is Nothing Then
            AppsOffice = "XL_NotInstalled"
        Else
            AppsOffice = "XL_Installed"
            TargetKey.Close()
        End If
    End Sub

    Private Sub btnBrow_DSPS_src_Click(sender As Object, e As EventArgs) Handles btnBrow_DSPS_src.Click
        Dim DSMpathSrc As String
        If OFD_DSPS.ShowDialog = DialogResult.OK Then
            DSMpathSrc = OFD_DSPS.FileName
            txtDSPS_src.Text = DSMpathSrc
        End If
        If txtDSPS_dest.Text <> "" Then
            btnNeu_DSPS.Enabled = True
        Else
            btnNeu_DSPS.Enabled = False
        End If
    End Sub

    Private Sub btnBrow_DSPS_dest_Click(sender As Object, e As EventArgs) Handles btnBrow_DSPS_dest.Click
        Dim datenow As DateTime = DateTime.Now
        Dim filename As String = "Neutralize_DSPS_" & datenow.ToString("ddMMyyyy_HHmm")
        SFD_DSPS.FileName = filename
        Dim DSMpath_Dest As String
        If SFD_DSPS.ShowDialog = DialogResult.OK Then
            DSMpath_Dest = SFD_DSPS.FileName
            txtDSPS_dest.Text = DSMpath_Dest
        End If
        If txtDSPS_src.Text <> "" Then
            btnNeu_DSPS.Enabled = True
        Else
            btnNeu_DSPS.Enabled = False
        End If
        If txtDSPS_dest.Text <> "" Then
            btnNeu_DSPS.Enabled = True
        Else
            btnNeu_DSPS.Enabled = False
        End If
    End Sub

    Private Sub btnNeu_DSPS_Click(sender As Object, e As EventArgs) Handles btnNeu_DSPS.Click
        PicBar_DSPS.Visible = True
        BWDSPS.RunWorkerAsync()
    End Sub

    Private Sub BWDSPS_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BWDSPS.DoWork
        Select Case AppsOffice
            Case "XL_NotInstalled"
                Dim xlAppDSPS As Object
                Dim xlWbookDSPS As Object
                Dim xlWsheetDSPS As Object

                Try
                    xlAppDSPS = CreateObject("Ket.Application")
                    xlWbookDSPS = xlAppDSPS.Workbooks.Open(txtDSPS_src.Text)
                    xlWsheetDSPS = xlWbookDSPS.Worksheets("UID Daily Sales And Payment Sum")

                    xlWsheetDSPS.UsedRange.UnMerge()
                    xlWsheetDSPS.UsedRange.WrapText = False
                    xlWsheetDSPS.UsedRange.ColumnWidth = 15
                    xlWsheetDSPS.UsedRange.RowHeight = 15

                    Dim rg_head_cut1 As Object = xlWsheetDSPS.Range("B2")
                    Dim rg_head_paste1 As Object = xlWsheetDSPS.Range("A2")
                    rg_head_cut1.Select()
                    rg_head_cut1.Cut(rg_head_paste1)

                    Dim rg_head_cut2 As Object = xlWsheetDSPS.Range("B4")
                    Dim rg_head_paste2 As Object = xlWsheetDSPS.Range("A3")
                    rg_head_cut2.Select()
                    rg_head_cut2.Cut(rg_head_paste2)

                    xlWsheetDSPS.Range("A2").RowHeight = 27

                    Dim paramhead1, paramhead2 As String
                    paramhead1 = xlWsheetDSPS.Range("C6").Value & " " & xlWsheetDSPS.Range("E6").Value & "; "
                    paramhead1 = paramhead1 & xlWsheetDSPS.Range("H6").Value & " " & xlWsheetDSPS.Range("J6").Value

                    paramhead2 = xlWsheetDSPS.Range("M6").Value & " " & xlWsheetDSPS.Range("O6").Value & "; "
                    paramhead2 = paramhead2 & xlWsheetDSPS.Range("T6").Value & " " & xlWsheetDSPS.Range("W6").Value
                    paramhead2 = paramhead2 & xlWsheetDSPS.Range("Z6").Value & " " & xlWsheetDSPS.Range("AC6").Value

                    xlWsheetDSPS.Range("A5").Value = paramhead1
                    xlWsheetDSPS.Range("A5").EntireRow.Font.Name = "Calibri"
                    xlWsheetDSPS.Range("A6").Value = paramhead2
                    xlWsheetDSPS.Range("A6").EntireRow.Font.Name = "Calibri"

                    xlWsheetDSPS.Range("C6:AD6").Value = ""

                    'Dim xlfunc As Object
                    'xlfunc = xlAppDSPS.WorksheetFunction
                    'Dim lnCol As Long
                    'Dim i, j As Long
                    'Dim rnarea As Object = xlWsheetDSPS.UsedRange

                    'lnCol = rnarea.Columns.Count
                    'For i = lnCol To 1 Step -1
                    '    If xlfunc.CountA(rnarea.Columns(i)) = 0 Then
                    '        rnarea.Columns(i).Delete()
                    '        j = j + 1
                    '    End If
                    'Next

                    xlWsheetDSPS.Range("B:C").EntireColumn.Delete()
                    xlWsheetDSPS.Range("C:D").EntireColumn.Delete()
                    xlWsheetDSPS.Range("D:K").EntireColumn.Delete()
                    xlWsheetDSPS.Range("E:E").EntireColumn.Delete()
                    xlWsheetDSPS.Range("F:G").EntireColumn.Delete()
                    xlWsheetDSPS.Range("G:H").EntireColumn.Delete()
                    xlWsheetDSPS.Range("H:I").EntireColumn.Delete()
                    xlWsheetDSPS.Range("I:J").EntireColumn.Delete()
                    xlWsheetDSPS.Range("J:J").EntireColumn.Delete()

                    xlWsheetDSPS.Range("D8").Copy(xlWsheetDSPS.Range("A8"))
                    xlWsheetDSPS.Range("A8").Value = "Date"
                    xlWsheetDSPS.Range("A8").Font.Color = System.Drawing.ColorTranslator.ToOle(Color.White)
                    xlWsheetDSPS.Range("A8").HorizontalAlignment = 3

                    xlWsheetDSPS.Range("D8").Copy(xlWsheetDSPS.Range("B8"))
                    xlWsheetDSPS.Range("B8").Value = "Salesman Code"
                    xlWsheetDSPS.Range("B8").Font.Color = System.Drawing.ColorTranslator.ToOle(Color.White)
                    xlWsheetDSPS.Range("B8").HorizontalAlignment = 3

                    xlWsheetDSPS.Range("D8").Copy(xlWsheetDSPS.Range("C8"))
                    xlWsheetDSPS.Range("C8").Value = "Salesman Name"
                    xlWsheetDSPS.Range("C8").Font.Color = System.Drawing.ColorTranslator.ToOle(Color.White)
                    xlWsheetDSPS.Range("C8").HorizontalAlignment = 3

                    xlWsheetDSPS.Range("D8").Copy(xlWsheetDSPS.Range("K8"))
                    xlWsheetDSPS.Range("K8").Value = "Setor"
                    xlWsheetDSPS.Range("K8").Font.Color = System.Drawing.ColorTranslator.ToOle(Color.White)
                    xlWsheetDSPS.Range("K8").HorizontalAlignment = 3

                    xlWsheetDSPS.Range("L8").Value = "Selisih"

                    xlWbookDSPS.SaveAs(txtDSPS_dest.Text)
                    xlWbookDSPS.Close()
                    xlAppDSPS.Quit()

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetDSPS)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookDSPS)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppDSPS)

                    xlWsheetDSPS = Nothing
                    xlWbookDSPS = Nothing
                    xlAppDSPS = Nothing

                    MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                Catch ex As Exception
                    MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try

            Case "XL_Installed"
                Dim xlAppDSPS As Object
                Dim xlWbookDSPS As Object
                Dim xlWsheetDSPS As Object

                Try
                    xlAppDSPS = CreateObject("Excel.Application")
                    xlWbookDSPS = xlAppDSPS.Workbooks.Open(txtDSPS_src.Text)
                    xlWsheetDSPS = xlWbookDSPS.Worksheets("UID Daily Sales And Payment Sum")

                    xlWsheetDSPS.UsedRange.UnMerge()
                    xlWsheetDSPS.UsedRange.WrapText = False
                    xlWsheetDSPS.UsedRange.ColumnWidth = 15
                    xlWsheetDSPS.UsedRange.RowHeight = 15

                    Dim rg_head_cut1 As Excel.Range = xlWsheetDSPS.Range("B2")
                    Dim rg_head_paste1 As Excel.Range = xlWsheetDSPS.Range("A2")
                    rg_head_cut1.Select()
                    rg_head_cut1.Cut(rg_head_paste1)

                    Dim rg_head_cut2 As Excel.Range = xlWsheetDSPS.Range("B4")
                    Dim rg_head_paste2 As Excel.Range = xlWsheetDSPS.Range("A3")
                    rg_head_cut2.Select()
                    rg_head_cut2.Cut(rg_head_paste2)

                    xlWsheetDSPS.Range("A2").RowHeight = 27

                    Dim paramhead1, paramhead2 As String
                    paramhead1 = xlWsheetDSPS.Range("C6").Value & " " & xlWsheetDSPS.Range("E6").Value & "; "
                    paramhead1 = paramhead1 & xlWsheetDSPS.Range("H6").Value & " " & xlWsheetDSPS.Range("J6").Value

                    paramhead2 = xlWsheetDSPS.Range("M6").Value & " " & xlWsheetDSPS.Range("O6").Value & "; "
                    paramhead2 = paramhead2 & xlWsheetDSPS.Range("T6").Value & " " & xlWsheetDSPS.Range("W6").Value
                    paramhead2 = paramhead2 & xlWsheetDSPS.Range("Z6").Value & " " & xlWsheetDSPS.Range("AC6").Value

                    xlWsheetDSPS.Range("A5").Value = paramhead1
                    xlWsheetDSPS.Range("A5").EntireRow.Font.Name = "Calibri"
                    xlWsheetDSPS.Range("A6").Value = paramhead2
                    xlWsheetDSPS.Range("A6").EntireRow.Font.Name = "Calibri"

                    xlWsheetDSPS.Range("B6:AC6").Value = ""

                    'Dim xlfunc As Object
                    'xlfunc = xlAppDSPS.WorksheetFunction
                    'Dim lnCol As Long
                    'Dim i, j As Long
                    'Dim rnarea As Excel.Range = xlWsheetDSPS.UsedRange

                    'lnCol = rnarea.Columns.Count
                    'For i = lnCol To 1 Step -1
                    '    If xlfunc.CountA(rnarea.Columns(i)) = 0 Then
                    '        rnarea.Columns(i).Delete()
                    '        j = j + 1
                    '    End If
                    'Next

                    xlWsheetDSPS.Range("B:C").EntireColumn.Delete()
                    xlWsheetDSPS.Range("C:D").EntireColumn.Delete()
                    xlWsheetDSPS.Range("D:K").EntireColumn.Delete()
                    xlWsheetDSPS.Range("E:E").EntireColumn.Delete()
                    xlWsheetDSPS.Range("F:G").EntireColumn.Delete()
                    xlWsheetDSPS.Range("G:H").EntireColumn.Delete()
                    xlWsheetDSPS.Range("H:I").EntireColumn.Delete()
                    xlWsheetDSPS.Range("I:J").EntireColumn.Delete()
                    xlWsheetDSPS.Range("J:J").EntireColumn.Delete()

                    xlWsheetDSPS.Range("D8").Copy(xlWsheetDSPS.Range("A8"))
                    xlWsheetDSPS.Range("A8").Value = "Date"
                    xlWsheetDSPS.Range("A8").Font.Color = System.Drawing.ColorTranslator.ToOle(Color.White)
                    xlWsheetDSPS.Range("A8").HorizontalAlignment = 3

                    xlWsheetDSPS.Range("D8").Copy(xlWsheetDSPS.Range("B8"))
                    xlWsheetDSPS.Range("B8").Value = "Salesman Code"
                    xlWsheetDSPS.Range("B8").Font.Color = System.Drawing.ColorTranslator.ToOle(Color.White)
                    xlWsheetDSPS.Range("B8").HorizontalAlignment = 3

                    xlWsheetDSPS.Range("D8").Copy(xlWsheetDSPS.Range("C8"))
                    xlWsheetDSPS.Range("C8").Value = "Salesman Name"
                    xlWsheetDSPS.Range("C8").Font.Color = System.Drawing.ColorTranslator.ToOle(Color.White)
                    xlWsheetDSPS.Range("C8").HorizontalAlignment = 3

                    xlWsheetDSPS.Range("D8").Copy(xlWsheetDSPS.Range("K8"))
                    xlWsheetDSPS.Range("K8").Value = "Setor"
                    xlWsheetDSPS.Range("K8").Font.Color = System.Drawing.ColorTranslator.ToOle(Color.White)
                    xlWsheetDSPS.Range("K8").HorizontalAlignment = 3

                    xlWsheetDSPS.Range("L8").Value = "Selisih"

                    xlWbookDSPS.SaveAs(txtDSPS_dest.Text)
                    xlWbookDSPS.Close()
                    xlAppDSPS.Quit()

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetDSPS)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookDSPS)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppDSPS)

                    xlWsheetDSPS = Nothing
                    xlWbookDSPS = Nothing
                    xlAppDSPS = Nothing

                    MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                Catch ex As Exception
                    MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
        End Select
    End Sub

    Private Sub BWDSPS_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BWDSPS.RunWorkerCompleted
        PicBar_DSPS.Visible = False
        btnNeu_DSPS.Enabled = False
        txtDSPS_dest.Text = ""
        txtDSPS_src.Text = ""
    End Sub
End Class