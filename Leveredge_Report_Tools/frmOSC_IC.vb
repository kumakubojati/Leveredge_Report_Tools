Imports Microsoft.Office.Interop
Imports System.Windows.Forms
Imports System.Text
Imports System.Drawing
Imports Microsoft.Win32

Public Class frmOSC_IC
    Dim AppsOffice As String
    Private Sub frmOSC_IC_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim TargetKey As RegistryKey
        TargetKey = Registry.ClassesRoot.OpenSubKey("Excel.Application")
        If TargetKey Is Nothing Then
            AppsOffice = "XL_NotInstalled"
        Else
            AppsOffice = "XL_Installed"
            TargetKey.Close()
        End If
    End Sub

    Private Sub btnBrow_OSC_IC_src_Click(sender As Object, e As EventArgs) Handles btnBrow_OSC_IC_src.Click
        Dim OSCpathSrc As String
        If OFD_OSC_IC.ShowDialog = DialogResult.OK Then
            OSCpathSrc = OFD_OSC_IC.FileName()
            txtOSC_IC_src.Text = OSCpathSrc
        End If
        If txtOSC_IC_dest.Text <> "" Then
            btnNeu_OSC_IC.Enabled = True
        Else
            btnNeu_OSC_IC.Enabled = False
        End If
    End Sub

    Private Sub btnBrow_OSC_IC_dest_Click(sender As Object, e As EventArgs) Handles btnBrow_OSC_IC_dest.Click
        Dim datenow As DateTime = DateTime.Now
        Dim filename As String = "Neutralize_IC_OutletStoreClass_" & datenow.ToString("ddMMyyyy_HHmm")
        SFD_OSC_IC.FileName = filename
        Dim OSCpath_Dest As String
        If SFD_OSC_IC.ShowDialog = DialogResult.OK Then
            OSCpath_Dest = SFD_OSC_IC.FileName
            txtOSC_IC_dest.Text = OSCpath_Dest
        End If
        If txtOSC_IC_src.Text <> "" Then
            btnNeu_OSC_IC.Enabled = True
        Else
            btnNeu_OSC_IC.Enabled = False
        End If
        If txtOSC_IC_dest.Text <> "" Then
            btnNeu_OSC_IC.Enabled = True
        Else
            btnNeu_OSC_IC.Enabled = False
        End If
    End Sub

    Private Sub btnNeu_OSC_IC_Click(sender As Object, e As EventArgs) Handles btnNeu_OSC_IC.Click
        PicBar_OSC_IC.Visible = True
        BWOSC_IC.RunWorkerAsync()
    End Sub

    Private Sub BWOSC_IC_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BWOSC_IC.DoWork
        Select Case AppsOffice
            Case "XL_NotInstalled"
                Dim xlAppOSC_IC As Object
                Dim xlWbookOSC_IC As Object
                Dim xlWsheetOSC_IC As Object
                Try
                    xlAppOSC_IC = CreateObject("Ket.Application")
                    xlWbookOSC_IC = xlAppOSC_IC.Workbooks.Open(txtOSC_IC_src.Text)
                    xlWsheetOSC_IC = xlWbookOSC_IC.Worksheets("UID IC Outlet Store Class Repor")

                    xlWsheetOSC_IC.UsedRange.UnMerge()
                    xlWsheetOSC_IC.UsedRange.WrapText = False
                    xlWsheetOSC_IC.UsedRange.ColumnWidth = 15
                    xlWsheetOSC_IC.UsedRange.RowHeight = 15

                    Dim rg_head_cut1 As Object = xlWsheetOSC_IC.Range("C2")
                    Dim rg_head_paste1 As Object = xlWsheetOSC_IC.Range("A2")
                    rg_head_cut1.Select()
                    rg_head_cut1.Cut(rg_head_paste1)

                    Dim rg_head_cut2 As Object = xlWsheetOSC_IC.Range("C4")
                    Dim rg_head_paste2 As Object = xlWsheetOSC_IC.Range("A3")
                    rg_head_cut2.Select()
                    rg_head_cut2.Cut(rg_head_paste2)

                    xlWsheetOSC_IC.Range("A2").RowHeight = 27

                    Dim paramhead1, paramhead2, paramhead3 As String
                    paramhead1 = xlWsheetOSC_IC.Range("B6").Value & " " & xlWsheetOSC_IC.Range("E6").Value & "; "
                    paramhead1 = paramhead1 & xlWsheetOSC_IC.Range("H6").Value & " " & xlWsheetOSC_IC.Range("J6").Value

                    paramhead2 = xlWsheetOSC_IC.Range("O6").Value & " " & xlWsheetOSC_IC.Range("S6").Value & ";"
                    paramhead2 = paramhead2 & xlWsheetOSC_IC.Range("B8").Value & " " & xlWsheetOSC_IC.Range("E8").Value & "; "
                    paramhead2 = paramhead2 & xlWsheetOSC_IC.Range("H8").Value & " " & xlWsheetOSC_IC.Range("K8").Value


                    paramhead3 = xlWsheetOSC_IC.Range("O8").Value & " " & xlWsheetOSC_IC.Range("Q8").Value

                    xlWsheetOSC_IC.Range("A5").Value = paramhead1
                    xlWsheetOSC_IC.Range("A5").EntireRow.Font.Name = "Calibri"
                    xlWsheetOSC_IC.Range("A6").Value = paramhead2
                    xlWsheetOSC_IC.Range("A6").EntireRow.Font.Name = "Calibri"
                    xlWsheetOSC_IC.Range("A7").Value = paramhead3
                    xlWsheetOSC_IC.Range("A7").EntireRow.Font.Name = "Calibri"

                    xlWsheetOSC_IC.Range("B6:AA8").value = ""
                    xlWsheetOSC_IC.Range("A9").EntireRow.Delete()

                    Dim xlfunc As Excel.WorksheetFunction
                    xlfunc = xlAppOSC_IC.WorksheetFunction()
                    Dim lnCol As Long
                    Dim i As Long
                    Dim rnarea As Object = xlWsheetOSC_IC.UsedRange

                    lnCol = rnarea.Columns.Count
                    For i = lnCol To 1 Step -1
                        If xlfunc.CountA(rnarea.Columns(i)) = 0 Then
                            rnarea.Columns(i).Delete()
                        End If
                    Next

                    xlWbookOSC_IC.SaveAs(txtOSC_IC_dest.Text)
                    xlWbookOSC_IC.Close()
                    xlAppOSC_IC.Quit()

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetOSC_IC)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookOSC_IC)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppOSC_IC)

                    xlWsheetOSC_IC = Nothing
                    xlWbookOSC_IC = Nothing
                    xlAppOSC_IC = Nothing

                    MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                Catch ex As Exception
                    MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try

            Case "XL_Installed"
                Dim xlAppOSC_IC As Object
                Dim xlWbookOSC_IC As Object
                Dim xlWsheetOSC_IC As Object
                Try
                    xlAppOSC_IC = CreateObject("Excel.Application")
                    xlWbookOSC_IC = xlAppOSC_IC.Workbooks.Open(txtOSC_IC_src.Text)
                    xlWsheetOSC_IC = xlWbookOSC_IC.Worksheets("UID IC Outlet Store Class Repor")

                    xlWsheetOSC_IC.UsedRange.UnMerge()
                    xlWsheetOSC_IC.UsedRange.WrapText = False
                    xlWsheetOSC_IC.UsedRange.ColumnWidth = 15
                    xlWsheetOSC_IC.UsedRange.RowHeight = 15

                    Dim rg_head_cut1 As Excel.Range = xlWsheetOSC_IC.Range("C2")
                    Dim rg_head_paste1 As Excel.Range = xlWsheetOSC_IC.Range("A2")
                    rg_head_cut1.Select()
                    rg_head_cut1.Cut(rg_head_paste1)

                    Dim rg_head_cut2 As Excel.Range = xlWsheetOSC_IC.Range("C4")
                    Dim rg_head_paste2 As Excel.Range = xlWsheetOSC_IC.Range("A3")
                    rg_head_cut2.Select()
                    rg_head_cut2.Cut(rg_head_paste2)

                    xlWsheetOSC_IC.Range("A2").RowHeight = 27

                    Dim paramhead1, paramhead2, paramhead3 As String
                    paramhead1 = xlWsheetOSC_IC.Range("B6").Value & " " & xlWsheetOSC_IC.Range("E6").Value & "; "
                    paramhead1 = paramhead1 & xlWsheetOSC_IC.Range("H6").Value & " " & xlWsheetOSC_IC.Range("J6").Value

                    paramhead2 = xlWsheetOSC_IC.Range("O6").Value & " " & xlWsheetOSC_IC.Range("S6").Value & ";"
                    paramhead2 = paramhead2 & xlWsheetOSC_IC.Range("B8").Value & " " & xlWsheetOSC_IC.Range("E8").Value & "; "
                    paramhead2 = paramhead2 & xlWsheetOSC_IC.Range("H8").Value & " " & xlWsheetOSC_IC.Range("K8").Value


                    paramhead3 = xlWsheetOSC_IC.Range("O8").Value & " " & xlWsheetOSC_IC.Range("Q8").Value

                    xlWsheetOSC_IC.Range("A5").Value = paramhead1
                    xlWsheetOSC_IC.Range("A5").EntireRow.Font.Name = "Calibri"
                    xlWsheetOSC_IC.Range("A6").Value = paramhead2
                    xlWsheetOSC_IC.Range("A6").EntireRow.Font.Name = "Calibri"
                    xlWsheetOSC_IC.Range("A7").Value = paramhead3
                    xlWsheetOSC_IC.Range("A7").EntireRow.Font.Name = "Calibri"

                    xlWsheetOSC_IC.Range("B6:AA8").value = ""
                    xlWsheetOSC_IC.Range("A9").EntireRow.Delete()

                    Dim xlfunc As Excel.WorksheetFunction
                    xlfunc = xlAppOSC_IC.WorksheetFunction()
                    Dim lnCol As Long
                    Dim i As Long
                    Dim rnarea As Excel.Range = xlWsheetOSC_IC.UsedRange

                    lnCol = rnarea.Columns.Count
                    For i = lnCol To 1 Step -1
                        If xlfunc.CountA(rnarea.Columns(i)) = 0 Then
                            rnarea.Columns(i).Delete()
                        End If
                    Next

                    xlWbookOSC_IC.SaveAs(txtOSC_IC_dest.Text)
                    xlWbookOSC_IC.Close()
                    xlAppOSC_IC.Quit()

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetOSC_IC)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookOSC_IC)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppOSC_IC)

                    xlWsheetOSC_IC = Nothing
                    xlWbookOSC_IC = Nothing
                    xlAppOSC_IC = Nothing

                    MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                Catch ex As Exception
                    MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
        End Select
    End Sub

    Private Sub BWOSC_IC_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BWOSC_IC.RunWorkerCompleted
        PicBar_OSC_IC.Visible = False
        btnNeu_OSC_IC.Enabled = False
        txtOSC_IC_dest.Text = ""
        txtOSC_IC_src.Text = ""
    End Sub
End Class