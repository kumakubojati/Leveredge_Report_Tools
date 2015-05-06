Imports Microsoft.Office.Interop
Imports System.Windows.Forms
Imports System.Text
Imports System.Drawing
Imports Microsoft.Win32
Public Class frmListPro
    Dim AppsOffice As String
    Private Sub frmListPro_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim TargetKey As RegistryKey
        TargetKey = Registry.ClassesRoot.OpenSubKey("Excel.Application")
        If TargetKey Is Nothing Then
            AppsOffice = "XL_NotInstalled"
        Else
            AppsOffice = "XL_Installed"
            TargetKey.Close()
        End If
    End Sub

    Private Sub btnBrowProm_src_Click(sender As Object, e As EventArgs) Handles btnBrowProm_src.Click
        Dim PrmpathSrc As String
        If OFD_Pro.ShowDialog = DialogResult.OK Then
            PrmpathSrc = OFD_Pro.FileName
            txtPromo_src.Text = PrmpathSrc
        End If
        If txtProm_dest.Text <> "" Then
            btnNeu_Promo.Enabled = True
        Else
            btnNeu_Promo.Enabled = False
        End If
    End Sub

    Private Sub btnBrowPromo_dest_Click(sender As Object, e As EventArgs) Handles btnBrowPromo_dest.Click
        Dim datenow As DateTime = DateTime.Now
        Dim filename As String
        filename = "Neutralize_ListOfPromo_" & datenow.ToString("ddMMyyyy_HHmm")
        OFD_Pro.FileName = filename

        Dim Prompath_Dest As String
        If OFD_Pro.ShowDialog = DialogResult.OK Then
            Prompath_Dest = OFD_Pro.FileName
            txtProm_dest.Text = Prompath_Dest
        End If
        If txtPromo_src.Text <> "" Then
            btnNeu_Promo.Enabled = True
        Else
            btnNeu_Promo.Enabled = False
        End If
        If txtProm_dest.Text <> "" Then
            btnNeu_Promo.Enabled = True
        Else
            btnNeu_Promo.Enabled = False
        End If
    End Sub

    Private Sub btnNeu_Promo_Click(sender As Object, e As EventArgs) Handles btnNeu_Promo.Click
        PicBar_Promo.Visible = True
        BWPro.RunWorkerAsync()
    End Sub

    Private Sub BWPro_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BWPro.RunWorkerCompleted
        PicBar_Promo.Visible = False
        btnNeu_Promo.Enabled = False
        txtProm_dest.Text = ""
        txtPromo_src.Text = ""
    End Sub

    Private Sub BWPro_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BWPro.DoWork
        Select Case AppsOffice
            Case "XL_NotInstalled"
                Dim xlAppPRM As OBject
                Dim xlWbookPRM As Object
                Dim xlWsheetPRM As Object

                Try
                    xlAppPRM = CreateObject("Ket.Application")
                    xlWbookPRM = xlAppPRM.Workbooks.Open(txtPromo_src.Text)
                    xlWsheetPRM = xlWbookPRM.Worksheets("UID List Of Promotion Report")

                    xlWsheetPRM.UsedRange.UnMerge()
                    xlWsheetPRM.UsedRange.WrapText = False
                    xlWsheetPRM.UsedRange.ColumnWidth = 15
                    xlWsheetPRM.UsedRange.RowHeight = 15

                    Dim rg_head_cut1 As Object = xlWsheetPRM.Range("B2")
                    Dim rg_head_paste1 As Object = xlWsheetPRM.Range("A2")
                    rg_head_cut1.Select()
                    rg_head_cut1.Cut(rg_head_paste1)

                    Dim rg_head_cut2 As Object = xlWsheetPRM.Range("B4")
                    Dim rg_head_paste2 As Object = xlWsheetPRM.Range("A3")
                    rg_head_cut2.Select()
                    rg_head_cut2.Cut(rg_head_paste2)

                    xlWsheetPRM.Range("A2").RowHeight = 27

                    Dim paramhead1, paramhead2, paramhead3 As String
                    paramhead1 = xlWsheetPRM.Range("C6").Value & " " & xlWsheetPRM.Range("F6").Value & "; "
                    paramhead1 = paramhead1 & xlWsheetPRM.Range("H6").Value & " " & xlWsheetPRM.Range("J6").Value & "; "

                    paramhead2 = xlWsheetPRM.Range("O6").Value & " " & xlWsheetPRM.Range("R6").Value & "; "
                    paramhead2 = paramhead2 & xlWsheetPRM.Range("U6").Value & " " & xlWsheetPRM.Range("X6").Value & "; "

                    paramhead3 = xlWsheetPRM.Range("C8").Value & " " & xlWsheetPRM.Range("F8").Value

                    xlWsheetPRM.Range("A5").Value = paramhead1
                    xlWsheetPRM.Range("A5").EntireRow.Font.Name = "Calibri"
                    xlWsheetPRM.Range("A6").Value = paramhead2
                    xlWsheetPRM.Range("A6").EntireRow.Font.Name = "Calibri"
                    xlWsheetPRM.Range("A7").Value = paramhead3
                    xlWsheetPRM.Range("A7").EntireRow.Font.Name = "Calibri"

                    Dim rg1, rg2, rg3, rg4, rg5, rg6 As Object
                    rg1 = xlWsheetPRM.Range("B:C")
                    rg1.Select()
                    rg1.Delete()
                    rg2 = xlWsheetPRM.Range("C:H")
                    rg2.Select()
                    rg2.Delete()
                    rg3 = xlWsheetPRM.Range("D:E")
                    rg3.Select()
                    rg3.Delete()
                    rg4 = xlWsheetPRM.Range("E:F")
                    rg4.Select()
                    rg4.Delete()
                    xlWsheetPRM.Range("F:F").EntireColumn.Delete()
                    rg5 = xlWsheetPRM.Range("G:H")
                    rg5.Select()
                    rg5.Delete()
                    rg6 = xlWsheetPRM.Range("H:J")
                    rg6.Select()
                    rg6.Delete()
                    xlWsheetPRM.Range("J:J").EntireColumn.Delete()

                    xlWsheetPRM.Range("A11:A12").Merge()
                    xlWsheetPRM.Range("A11:A12").HorizontalAlignment = 3
                    xlWsheetPRM.Range("B11:B12").Merge()
                    xlWsheetPRM.Range("B11:B12").HorizontalAlignment = 3
                    xlWsheetPRM.Range("C11:C12").Merge()
                    xlWsheetPRM.Range("C11:C12").HorizontalAlignment = 3
                    xlWsheetPRM.Range("D11:D12").Merge()
                    xlWsheetPRM.Range("D11:D12").HorizontalAlignment = 3
                    xlWsheetPRM.Range("E11:F11").Merge()
                    xlWsheetPRM.Range("E11:F11").HorizontalAlignment = 3
                    xlWsheetPRM.Range("G11:I11").Merge()
                    xlWsheetPRM.Range("G11:I11").HorizontalAlignment = 3
                    xlWsheetPRM.Range("J11:J12").Merge()
                    xlWsheetPRM.Range("J11:J12").HorizontalAlignment = 3
                    xlWsheetPRM.Range("K11:K12").Merge()
                    xlWsheetPRM.Range("K11:K12").HorizontalAlignment = 3
                    xlWsheetPRM.Range("L11:M11").Merge()
                    xlWsheetPRM.Range("L11:M11").HorizontalAlignment = 3

                    xlWbookPRM.SaveAs(txtProm_dest.Text)
                    xlWbookPRM.Close()
                    xlAppPRM.Quit()

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetPRM)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookPRM)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppPRM)

                    xlWsheetPRM = Nothing
                    xlWbookPRM = Nothing
                    xlAppPRM = Nothing

                    MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Catch ex As Exception
                    MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try

            Case "XL_Installed"
                Dim xlAppPRM As Object
                Dim xlWbookPRM As Object
                Dim xlWsheetPRM As Object

                Try
                    xlAppPRM = CreateObject("Excel.Application")
                    xlWbookPRM = xlAppPRM.Workbooks.Open(txtPromo_src.Text)
                    xlWsheetPRM = xlWbookPRM.Worksheets("UID List Of Promotion Report")

                    xlWsheetPRM.UsedRange.UnMerge()
                    xlWsheetPRM.UsedRange.WrapText = False
                    xlWsheetPRM.UsedRange.ColumnWidth = 15
                    xlWsheetPRM.UsedRange.RowHeight = 15

                    Dim rg_head_cut1 As Excel.Range = xlWsheetPRM.Range("B2")
                    Dim rg_head_paste1 As Excel.Range = xlWsheetPRM.Range("A2")
                    rg_head_cut1.Select()
                    rg_head_cut1.Cut(rg_head_paste1)

                    Dim rg_head_cut2 As Excel.Range = xlWsheetPRM.Range("B4")
                    Dim rg_head_paste2 As Excel.Range = xlWsheetPRM.Range("A3")
                    rg_head_cut2.Select()
                    rg_head_cut2.Cut(rg_head_paste2)

                    xlWsheetPRM.Range("A2").RowHeight = 27

                    Dim paramhead1, paramhead2, paramhead3 As String
                    paramhead1 = xlWsheetPRM.Range("C6").Value & " " & xlWsheetPRM.Range("F6").Value & "; "
                    paramhead1 = paramhead1 & xlWsheetPRM.Range("H6").Value & " " & xlWsheetPRM.Range("J6").Value & "; "

                    paramhead2 = xlWsheetPRM.Range("O6").Value & " " & xlWsheetPRM.Range("R6").Value & "; "
                    paramhead2 = paramhead2 & xlWsheetPRM.Range("U6").Value & " " & xlWsheetPRM.Range("X6").Value & "; "

                    paramhead3 = xlWsheetPRM.Range("C8").Value & " " & xlWsheetPRM.Range("F8").Value

                    xlWsheetPRM.Range("A5").Value = paramhead1
                    xlWsheetPRM.Range("A5").EntireRow.Font.Name = "Calibri"
                    xlWsheetPRM.Range("A6").Value = paramhead2
                    xlWsheetPRM.Range("A6").EntireRow.Font.Name = "Calibri"
                    xlWsheetPRM.Range("A7").Value = paramhead3
                    xlWsheetPRM.Range("A7").EntireRow.Font.Name = "Calibri"

                    Dim rg1, rg2, rg3, rg4, rg5, rg6 As Excel.Range
                    rg1 = xlWsheetPRM.Range("B:C")
                    rg1.Select()
                    rg1.Delete()
                    rg2 = xlWsheetPRM.Range("C:H")
                    rg2.Select()
                    rg2.Delete()
                    rg3 = xlWsheetPRM.Range("D:E")
                    rg3.Select()
                    rg3.Delete()
                    rg4 = xlWsheetPRM.Range("E:F")
                    rg4.Select()
                    rg4.Delete()
                    xlWsheetPRM.Range("F:F").EntireColumn.Delete()
                    rg5 = xlWsheetPRM.Range("G:H")
                    rg5.Select()
                    rg5.Delete()
                    rg6 = xlWsheetPRM.Range("H:J")
                    rg6.Select()
                    rg6.Delete()
                    xlWsheetPRM.Range("J:J").EntireColumn.Delete()

                    xlWsheetPRM.Range("A11:A12").Merge()
                    xlWsheetPRM.Range("A11:A12").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetPRM.Range("B11:B12").Merge()
                    xlWsheetPRM.Range("B11:B12").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetPRM.Range("C11:C12").Merge()
                    xlWsheetPRM.Range("C11:C12").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetPRM.Range("D11:D12").Merge()
                    xlWsheetPRM.Range("D11:D12").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetPRM.Range("E11:F11").Merge()
                    xlWsheetPRM.Range("E11:F11").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetPRM.Range("G11:I11").Merge()
                    xlWsheetPRM.Range("G11:I11").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetPRM.Range("J11:J12").Merge()
                    xlWsheetPRM.Range("J11:J12").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetPRM.Range("K11:K12").Merge()
                    xlWsheetPRM.Range("K11:K12").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetPRM.Range("L11:M11").Merge()
                    xlWsheetPRM.Range("L11:M11").HorizontalAlignment = Excel.Constants.xlCenter

                    xlWbookPRM.SaveAs(txtProm_dest.Text)
                    xlWbookPRM.Close()
                    xlAppPRM.Quit()

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetPRM)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookPRM)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppPRM)

                    xlWsheetPRM = Nothing
                    xlWbookPRM = Nothing
                    xlAppPRM = Nothing

                    MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Catch ex As Exception
                    MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try

        End Select
    End Sub
End Class