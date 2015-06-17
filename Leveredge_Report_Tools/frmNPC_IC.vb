Imports Microsoft.Office.Interop
Imports System.Windows.Forms
Imports System.Text
Imports System.Drawing
Imports Microsoft.Win32

Public Class frmNPC_IC
    Dim AppsOffice As String
    Private Sub frmNPC_IC_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim TargetKey As RegistryKey
        TargetKey = Registry.ClassesRoot.OpenSubKey("Excel.Application")
        If TargetKey Is Nothing Then
            AppsOffice = "XL_NotInstalled"
        Else
            AppsOffice = "XL_Installed"
            TargetKey.Close()
        End If
    End Sub

    Private Sub btnBrow_NPC_IC_src_Click(sender As Object, e As EventArgs) Handles btnBrow_NPC_IC_src.Click
        Dim NPCpathSrc As String
        If OFD_NPC_IC.ShowDialog = DialogResult.OK Then
            NPCpathSrc = OFD_NPC_IC.FileName()
            txtNPC_IC_src.Text = NPCpathSrc
        End If
        If txtNPC_IC_dest.Text <> "" Then
            btnNeu_NPC_IC.Enabled = True
        Else
            btnNeu_NPC_IC.Enabled = False
        End If
    End Sub

    Private Sub btnBrow_NPC_IC_dest_Click(sender As Object, e As EventArgs) Handles btnBrow_NPC_IC_dest.Click
        Dim datenow As DateTime = DateTime.Now
        Dim filename As String = "Neutralize_IC_NonPerfCab_" & datenow.ToString("ddMMyyyy_HHmm")
        SFD_NPC_IC.FileName = filename
        Dim NPCpath_Dest As String
        If SFD_NPC_IC.ShowDialog = DialogResult.OK Then
            NPCpath_Dest = SFD_NPC_IC.FileName
            txtNPC_IC_dest.Text = NPCpath_Dest
        End If
        If txtNPC_IC_src.Text <> "" Then
            btnNeu_NPC_IC.Enabled = True
        Else
            btnNeu_NPC_IC.Enabled = False
        End If
        If txtNPC_IC_dest.Text <> "" Then
            btnNeu_NPC_IC.Enabled = True
        Else
            btnNeu_NPC_IC.Enabled = False
        End If
    End Sub

    Private Sub btnNeu_NPC_IC_Click(sender As Object, e As EventArgs) Handles btnNeu_NPC_IC.Click
        PicBar_NPC_IC.Visible = True
        BWNPC_IC.RunWorkerAsync()
    End Sub

    Private Sub BWNPC_IC_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BWNPC_IC.DoWork
        Select Case AppsOffice
            Case "XL_NotInstalled"
                Dim xlAppNPC_IC As Object
                Dim xlWbookNPC_IC As Object
                Dim xlWsheetNPC_IC As Object
                Try
                    xlAppNPC_IC = CreateObject("Ket.Application")
                    xlWbookNPC_IC = xlAppNPC_IC.Workbooks.Open(txtNPC_IC_src.Text)
                    xlWsheetNPC_IC = xlWbookNPC_IC.Worksheets("UID IC Non Performance Cabinet ")

                    xlWsheetNPC_IC.UsedRange.UnMerge()
                    xlWsheetNPC_IC.UsedRange.WrapText = False
                    xlWsheetNPC_IC.UsedRange.ColumnWidth = 15
                    xlWsheetNPC_IC.UsedRange.RowHeight = 15

                    Dim rg_head_cut1 As Object = xlWsheetNPC_IC.Range("C2")
                    Dim rg_head_paste1 As Object = xlWsheetNPC_IC.Range("A2")
                    rg_head_cut1.Select()
                    rg_head_cut1.Cut(rg_head_paste1)

                    Dim rg_head_cut2 As Object = xlWsheetNPC_IC.Range("C4")
                    Dim rg_head_paste2 As Object = xlWsheetNPC_IC.Range("A3")
                    rg_head_cut2.Select()
                    rg_head_cut2.Cut(rg_head_paste2)

                    xlWsheetNPC_IC.Range("A2").RowHeight = 27

                    Dim paramhead1, paramhead2, paramhead3 As String
                    paramhead1 = xlWsheetNPC_IC.Range("B6").Value & " " & xlWsheetNPC_IC.Range("H6").Value & "; "
                    paramhead1 = paramhead1 & xlWsheetNPC_IC.Range("L6").Value & " " & xlWsheetNPC_IC.Range("O6").Value

                    paramhead2 = xlWsheetNPC_IC.Range("S6").Value & " " & xlWsheetNPC_IC.Range("V6").Value & "; "
                    paramhead2 = paramhead2 & xlWsheetNPC_IC.Range("AE6").Value & " " & xlWsheetNPC_IC.Range("AG6").Value

                    paramhead3 = xlWsheetNPC_IC.Range("B8").Value & " " & xlWsheetNPC_IC.Range("E8").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetNPC_IC.Range("L8").Value & " " & xlWsheetNPC_IC.Range("O8").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetNPC_IC.Range("Y8").Value & " " & xlWsheetNPC_IC.Range("AB8").Value

                    xlWsheetNPC_IC.Range("A5").Value = paramhead1
                    xlWsheetNPC_IC.Range("A5").EntireRow.Font.Name = "Calibri"
                    xlWsheetNPC_IC.Range("A6").Value = paramhead2
                    xlWsheetNPC_IC.Range("A6").EntireRow.Font.Name = "Calibri"
                    xlWsheetNPC_IC.Range("A7").Value = paramhead3
                    xlWsheetNPC_IC.Range("A7").EntireRow.Font.Name = "Calibri"

                    xlWsheetNPC_IC.Range("B6:AG9").value = ""
                    xlWsheetNPC_IC.Range("A9:A10").EntireRow.Delete()

                    Dim xlfunc As Excel.WorksheetFunction
                    xlfunc = xlAppNPC_IC.WorksheetFunction
                    Dim lnCol As Long
                    Dim i As Long
                    Dim rnarea As Object = xlWsheetNPC_IC.UsedRange

                    lnCol = rnarea.Columns.Count
                    For i = lnCol To 1 Step -1
                        If xlfunc.CountA(rnarea.Columns(i)) = 0 Then
                            rnarea.Columns(i).Delete()
                        End If
                    Next

                    xlWsheetNPC_IC.Range("J9:K9").Merge()

                    xlWbookNPC_IC.SaveAs(txtNPC_IC_dest.Text)
                    xlWbookNPC_IC.Close()
                    xlAppNPC_IC.Quit()

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetNPC_IC)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookNPC_IC)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppNPC_IC)

                    xlWsheetNPC_IC = Nothing
                    xlWbookNPC_IC = Nothing
                    xlAppNPC_IC = Nothing

                    MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                Catch ex As Exception
                    MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try

            Case "XL_Installed"
                Dim xlAppNPC_IC As Object
                Dim xlWbookNPC_IC As Object
                Dim xlWsheetNPC_IC As Object
                Try
                    xlAppNPC_IC = CreateObject("Excel.Application")
                    xlWbookNPC_IC = xlAppNPC_IC.Workbooks.Open(txtNPC_IC_src.Text)
                    xlWsheetNPC_IC = xlWbookNPC_IC.Worksheets("UID IC Non Performance Cabinet ")

                    xlWsheetNPC_IC.UsedRange.UnMerge()
                    xlWsheetNPC_IC.UsedRange.WrapText = False
                    xlWsheetNPC_IC.UsedRange.ColumnWidth = 15
                    xlWsheetNPC_IC.UsedRange.RowHeight = 15

                    Dim rg_head_cut1 As Excel.Range = xlWsheetNPC_IC.Range("C2")
                    Dim rg_head_paste1 As Excel.Range = xlWsheetNPC_IC.Range("A2")
                    rg_head_cut1.Select()
                    rg_head_cut1.Cut(rg_head_paste1)

                    Dim rg_head_cut2 As Excel.Range = xlWsheetNPC_IC.Range("C4")
                    Dim rg_head_paste2 As Excel.Range = xlWsheetNPC_IC.Range("A3")
                    rg_head_cut2.Select()
                    rg_head_cut2.Cut(rg_head_paste2)

                    xlWsheetNPC_IC.Range("A2").RowHeight = 27

                    Dim paramhead1, paramhead2, paramhead3 As String
                    paramhead1 = xlWsheetNPC_IC.Range("B6").Value & " " & xlWsheetNPC_IC.Range("H6").Value & "; "
                    paramhead1 = paramhead1 & xlWsheetNPC_IC.Range("L6").Value & " " & xlWsheetNPC_IC.Range("O6").Value

                    paramhead2 = xlWsheetNPC_IC.Range("S6").Value & " " & xlWsheetNPC_IC.Range("V6").Value & "; "
                    paramhead2 = paramhead2 & xlWsheetNPC_IC.Range("AE6").Value & " " & xlWsheetNPC_IC.Range("AG6").Value

                    paramhead3 = xlWsheetNPC_IC.Range("B8").Value & " " & xlWsheetNPC_IC.Range("E8").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetNPC_IC.Range("L8").Value & " " & xlWsheetNPC_IC.Range("O8").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetNPC_IC.Range("Y8").Value & " " & xlWsheetNPC_IC.Range("AB8").Value

                    xlWsheetNPC_IC.Range("A5").Value = paramhead1
                    xlWsheetNPC_IC.Range("A5").EntireRow.Font.Name = "Calibri"
                    xlWsheetNPC_IC.Range("A6").Value = paramhead2
                    xlWsheetNPC_IC.Range("A6").EntireRow.Font.Name = "Calibri"
                    xlWsheetNPC_IC.Range("A7").Value = paramhead3
                    xlWsheetNPC_IC.Range("A7").EntireRow.Font.Name = "Calibri"

                    xlWsheetNPC_IC.Range("B6:AG9").value = ""
                    xlWsheetNPC_IC.Range("A9:A10").EntireRow.Delete()

                    Dim xlfunc As Excel.WorksheetFunction
                    xlfunc = xlAppNPC_IC.WorksheetFunction
                    Dim lnCol As Long
                    Dim i As Long
                    Dim rnarea As Excel.Range = xlWsheetNPC_IC.UsedRange

                    lnCol = rnarea.Columns.Count
                    For i = lnCol To 1 Step -1
                        If xlfunc.CountA(rnarea.Columns(i)) = 0 Then
                            rnarea.Columns(i).Delete()
                        End If
                    Next

                    xlWsheetNPC_IC.Range("J9:K9").Merge()

                    xlWbookNPC_IC.SaveAs(txtNPC_IC_dest.Text)
                    xlWbookNPC_IC.Close()
                    xlAppNPC_IC.Quit()

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetNPC_IC)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookNPC_IC)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppNPC_IC)

                    xlWsheetNPC_IC = Nothing
                    xlWbookNPC_IC = Nothing
                    xlAppNPC_IC = Nothing

                    MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                Catch ex As Exception
                    MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
        End Select
    End Sub

    Private Sub BWNPC_IC_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BWNPC_IC.RunWorkerCompleted
        PicBar_NPC_IC.Visible = False
        btnNeu_NPC_IC.Enabled = False
        txtNPC_IC_dest.Text = ""
        txtNPC_IC_src.Text = ""
    End Sub
End Class