Imports Microsoft.Office.Interop
Imports System.Windows.Forms
Imports System.Text
Imports System.Drawing
Imports Microsoft.Win32

Public Class frmProdMas
    Dim AppsOffice As String
    Private Sub btnBrowProdMas_src_Click(sender As Object, e As EventArgs) Handles btnBrowProdMas_src.Click
        Dim prodmaspathSrc As String
        If OFD_ProdMas.ShowDialog = DialogResult.OK Then
            prodmaspathSrc = OFD_ProdMas.FileName
            txtProdMas_src.Text = prodmaspathSrc
        End If
        If txtProdMas_dest.Text <> "" Then
            btnNeu_ProdMas.Enabled = True
        Else
            btnNeu_ProdMas.Enabled = False
        End If
    End Sub

    Private Sub frmProdMas_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim TargetKey As RegistryKey
        TargetKey = Registry.ClassesRoot.OpenSubKey("Excel.Application")
        If TargetKey Is Nothing Then
            AppsOffice = "XL_NotInstalled"
        Else
            AppsOffice = "XL_Installed"
            TargetKey.Close()
        End If
    End Sub

    Private Sub btnBrowProdMas_dest_Click(sender As Object, e As EventArgs) Handles btnBrowProdMas_dest.Click
        Dim datenow As DateTime = DateTime.Now
        Dim filename As String = "Neutralize_ProdMast_" & datenow.ToString("ddMMyyyy_HHmm")
        SFD_ProdMas.FileName = filename
        Dim ProdMasPath_Dest As String
        If SFD_ProdMas.ShowDialog = DialogResult.OK Then
            ProdMasPath_Dest = SFD_ProdMas.FileName
            txtProdMas_dest.Text = ProdMasPath_Dest
        End If
        If txtProdMas_src.Text <> "" Then
            btnNeu_ProdMas.Enabled = True
        Else
            btnNeu_ProdMas.Enabled = False
        End If
        If txtProdMas_dest.Text <> "" Then
            btnNeu_ProdMas.Enabled = True
        Else
            btnNeu_ProdMas.Enabled = False
        End If
    End Sub

    Private Sub btnNeu_ProdMas_Click(sender As Object, e As EventArgs) Handles btnNeu_ProdMas.Click
        PicBar_ProdMas.Visible = True
        BWProdMas.RunWorkerAsync()
    End Sub

    Private Sub BWProdMas_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BWProdMas.RunWorkerCompleted
        PicBar_ProdMas.Visible = False
        btnNeu_ProdMas.Enabled = False
        txtProdMas_dest.Text = ""
        txtProdMas_src.Text = ""
    End Sub

    Private Sub BWProdMas_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BWProdMas.DoWork
        Select Case AppsOffice
            Case "XL_NotInstalled"
                Dim xlAppProdMas As Object
                Dim xlWbookProdMas As Object
                Dim xlWsheetProdMas As Object

                Try
                    xlAppProdMas = CreateObject("Ket.Application")
                    xlWbookProdMas = xlAppProdMas.Workbooks.Open(txtProdMas_src.Text)
                    xlWsheetProdMas = xlWbookProdMas.Worksheets("UID Product Master Report")

                    xlWsheetProdMas.UsedRange.UnMerge()
                    xlWsheetProdMas.UsedRange.WrapText = False
                    xlWsheetProdMas.UsedRange.ColumnWidth = 15
                    xlWsheetProdMas.UsedRange.RowHeight = 15

                    Dim rg_head_cut1 As Object = xlWsheetProdMas.Range("B2")
                    Dim rg_head_paste1 As Object = xlWsheetProdMas.Range("A2")
                    rg_head_cut1.Select()
                    rg_head_cut1.Cut(rg_head_paste1)

                    Dim rg_head_cut2 As Object = xlWsheetProdMas.Range("B4")
                    Dim rg_head_paste2 As Object = xlWsheetProdMas.Range("A3")
                    rg_head_cut2.Select()
                    rg_head_cut2.Cut(rg_head_paste2)

                    Dim paramhead1, paramhead2, paramhead3 As String
                    paramhead1 = xlWsheetProdMas.Range("C6").Value & " " & xlWsheetProdMas.Range("F6").Value & "; "
                    paramhead1 = paramhead1 & xlWsheetProdMas.Range("I6").Value & " " & xlWsheetProdMas.Range("L6").Value & "; "

                    paramhead2 = xlWsheetProdMas.Range("R6").Value & " " & xlWsheetProdMas.Range("V6").Value & "; "
                    paramhead2 = paramhead2 & xlWsheetProdMas.Range("Y6").Value & " " & xlWsheetProdMas.Range("AB6").Value & "; "

                    paramhead3 = xlWsheetProdMas.Range("C8").Value & " " & xlWsheetProdMas.Range("F6").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetProdMas.Range("J8").Value & " " & xlWsheetProdMas.Range("L8").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetProdMas.Range("R8").Value & " " & xlWsheetProdMas.Range("V8").Value & "; "

                    xlWsheetProdMas.Range("A5").Value = paramhead1
                    xlWsheetProdMas.Range("A5").EntireRow.Font.Name = "Calibri"
                    xlWsheetProdMas.Range("A6").Value = paramhead2
                    xlWsheetProdMas.Range("A6").EntireRow.Font.Name = "Calibri"
                    xlWsheetProdMas.Range("A7").Value = paramhead3
                    xlWsheetProdMas.Range("A7").EntireRow.Font.Name = "Calibri"

                    xlWsheetProdMas.Range("A2").RowHeight = 27

                    Dim rg1, rg2, rg3, rg4, rg5, rg6, rg7, rg8, rg9 As Object
                    rg1 = xlWsheetProdMas.Range("B:C")
                    rg1.Select()
                    rg1.Delete()

                    rg2 = xlWsheetProdMas.Range("C:J")
                    rg2.Select()
                    rg2.Delete()

                    rg3 = xlWsheetProdMas.Range("D:F")
                    rg3.Select()
                    rg3.Delete()

                    xlWsheetProdMas.Range("E:E").EntireColumn.Delete()

                    rg4 = xlWsheetProdMas.Range("F:K")
                    rg4.Select()
                    rg4.Delete()

                    rg5 = xlWsheetProdMas.Range("G:H")
                    rg5.Select()
                    rg5.Delete()

                    xlWsheetProdMas.Range("I:I").EntireColumn.Delete()
                    xlWsheetProdMas.Range("AT:AT").EntireColumn.Delete()
                    xlWsheetProdMas.Range("A10").EntireRow.Delete()

                    rg6 = xlWsheetProdMas.Range("A10:A11")
                    xlWsheetProdMas.Range("E10:E11").Copy(rg6)
                    xlWsheetProdMas.Range("A10").Value = "SKU"
                    xlWsheetProdMas.Range("A10").Font.Color = System.Drawing.ColorTranslator.ToOle(Color.White)
                    xlWsheetProdMas.Range("A10:A11").Merge()
                    xlWsheetProdMas.Range("A10:A11").HorizontalAlignment = 3

                    rg7 = xlWsheetProdMas.Range("B10:B11")
                    xlWsheetProdMas.Range("E10:E11").Copy(rg7)
                    xlWsheetProdMas.Range("B10").Value = "Description"
                    xlWsheetProdMas.Range("B10").Font.Color = System.Drawing.ColorTranslator.ToOle(Color.White)
                    xlWsheetProdMas.Range("B10:B11").Merge()
                    xlWsheetProdMas.Range("B10:B11").HorizontalAlignment = 3

                    rg8 = xlWsheetProdMas.Range("C10:C11")
                    xlWsheetProdMas.Range("E10:E11").Copy(rg8)
                    xlWsheetProdMas.Range("C10").Value = "Product Type"
                    xlWsheetProdMas.Range("C10").Font.Color = System.Drawing.ColorTranslator.ToOle(Color.White)
                    xlWsheetProdMas.Range("C10:C11").Merge()
                    xlWsheetProdMas.Range("C10:C11").HorizontalAlignment = 3

                    rg9 = xlWsheetProdMas.Range("D10:D11")
                    xlWsheetProdMas.Range("E10:E11").Copy(rg9)
                    xlWsheetProdMas.Range("D10").Value = "PC/CS"
                    xlWsheetProdMas.Range("D10").Font.Color = System.Drawing.ColorTranslator.ToOle(Color.White)
                    xlWsheetProdMas.Range("D10:D11").Merge()
                    xlWsheetProdMas.Range("D10:D11").HorizontalAlignment = 3

                    xlWsheetProdMas.Range("E10:E11").Merge()
                    xlWsheetProdMas.Range("E10:E11").HorizontalAlignment = 3
                    xlWsheetProdMas.Range("F10:G10").Merge()
                    xlWsheetProdMas.Range("F10:G10").HorizontalAlignment = 3
                    xlWsheetProdMas.Range("F11").HorizontalAlignment = 3
                    xlWsheetProdMas.Range("H10:J10").Merge()
                    xlWsheetProdMas.Range("H10:J10").HorizontalAlignment = 3
                    xlWsheetProdMas.Range("H11").HorizontalAlignment = 3
                    xlWsheetProdMas.Range("K10:N10").Merge()
                    xlWsheetProdMas.Range("K10:N10").HorizontalAlignment = 3
                    xlWsheetProdMas.Range("O10:P10").Merge()
                    xlWsheetProdMas.Range("O10:P10").HorizontalAlignment = 3
                    xlWsheetProdMas.Range("R10:S10").Merge()
                    xlWsheetProdMas.Range("R10:S10").HorizontalAlignment = 3
                    xlWsheetProdMas.Range("AO10:AQ10").Merge()
                    xlWsheetProdMas.Range("AO10:AQ10").HorizontalAlignment = 3
                    xlWsheetProdMas.Range("T10:T11").Merge()
                    xlWsheetProdMas.Range("T10:T11").HorizontalAlignment = 3
                    xlWsheetProdMas.Range("U10:U11").Merge()
                    xlWsheetProdMas.Range("U10:U11").HorizontalAlignment = 3
                    xlWsheetProdMas.Range("V10:V11").Merge()
                    xlWsheetProdMas.Range("V10:V11").HorizontalAlignment = 3
                    xlWsheetProdMas.Range("W10:W11").Merge()
                    xlWsheetProdMas.Range("W10:W11").HorizontalAlignment = 3
                    xlWsheetProdMas.Range("V10:V11").Merge()
                    xlWsheetProdMas.Range("V10:V11").HorizontalAlignment = 3
                    xlWsheetProdMas.Range("X10:X11").Merge()
                    xlWsheetProdMas.Range("X10:X11").HorizontalAlignment = 3
                    xlWsheetProdMas.Range("Y10:Y11").Merge()
                    xlWsheetProdMas.Range("Y10:Y11").HorizontalAlignment = 3
                    xlWsheetProdMas.Range("Z10:Z11").Merge()
                    xlWsheetProdMas.Range("Z10:Z11").HorizontalAlignment = 3
                    xlWsheetProdMas.Range("AA10:AA11").Merge()
                    xlWsheetProdMas.Range("AA10:AA11").HorizontalAlignment = 3
                    xlWsheetProdMas.Range("AA10:AA11").WrapText = True
                    xlWsheetProdMas.Range("AB10:AB11").Merge()
                    xlWsheetProdMas.Range("AB10:AB11").HorizontalAlignment = 3
                    xlWsheetProdMas.Range("AC10:AC11").Merge()
                    xlWsheetProdMas.Range("AC10:AC11").HorizontalAlignment = 3
                    xlWsheetProdMas.Range("AD10:AD11").Merge()
                    xlWsheetProdMas.Range("AD10:AD11").HorizontalAlignment = 3
                    xlWsheetProdMas.Range("AE10:AE11").Merge()
                    xlWsheetProdMas.Range("AE10:AE11").HorizontalAlignment = 3
                    xlWsheetProdMas.Range("AF10:AF11").Merge()
                    xlWsheetProdMas.Range("AF10:AF11").HorizontalAlignment = 3
                    xlWsheetProdMas.Range("AG10:AG11").Merge()
                    xlWsheetProdMas.Range("AG10:AG11").HorizontalAlignment = 3
                    xlWsheetProdMas.Range("AH10:AH11").Merge()
                    xlWsheetProdMas.Range("AH10:AH11").HorizontalAlignment = 3
                    xlWsheetProdMas.Range("AI10:AI11").Merge()
                    xlWsheetProdMas.Range("AI10:AI11").HorizontalAlignment = 3
                    xlWsheetProdMas.Range("AJ10:AJ11").Merge()
                    xlWsheetProdMas.Range("AJ10:AJ11").HorizontalAlignment = 3
                    xlWsheetProdMas.Range("AK10:AK11").Merge()
                    xlWsheetProdMas.Range("AK10:AK11").HorizontalAlignment = 3
                    xlWsheetProdMas.Range("AL10:Al11").Merge()
                    xlWsheetProdMas.Range("AL10:Al11").HorizontalAlignment = 3
                    xlWsheetProdMas.Range("AM10:AM11").Merge()
                    xlWsheetProdMas.Range("AM10:AM11").HorizontalAlignment = 3
                    xlWsheetProdMas.Range("AN10:AN11").Merge()
                    xlWsheetProdMas.Range("AN10:AN11").HorizontalAlignment = 3
                    xlWsheetProdMas.Range("AR10:AR11").Merge()
                    xlWsheetProdMas.Range("AR10:AR11").HorizontalAlignment = 3
                    xlWsheetProdMas.Range("AS10:AS11").Merge()
                    xlWsheetProdMas.Range("AS10:AS11").HorizontalAlignment = 3

                    xlWbookProdMas.SaveAs(txtProdMas_dest.Text)
                    xlWbookProdMas.Close()
                    xlAppProdMas.Quit()

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetProdMas)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookProdMas)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppProdMas)

                    xlWsheetProdMas = Nothing
                    xlWbookProdMas = Nothing
                    xlAppProdMas = Nothing

                    MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                Catch ex As Exception
                    MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try

            Case "XL_Installed"
                Dim xlAppProdMas As Object
                Dim xlWbookProdMas As Object
                Dim xlWsheetProdMas As Object

                Try
                    xlAppProdMas = CreateObject("Excel.Application")
                    xlWbookProdMas = xlAppProdMas.Workbooks.Open(txtProdMas_src.Text)
                    xlWsheetProdMas = xlWbookProdMas.Worksheets("UID Product Master Report")

                    xlWsheetProdMas.UsedRange.UnMerge()
                    xlWsheetProdMas.UsedRange.WrapText = False
                    xlWsheetProdMas.UsedRange.ColumnWidth = 15
                    xlWsheetProdMas.UsedRange.RowHeight = 15

                    Dim rg_head_cut1 As Excel.Range = xlWsheetProdMas.Range("B2")
                    Dim rg_head_paste1 As Excel.Range = xlWsheetProdMas.Range("A2")
                    rg_head_cut1.Select()
                    rg_head_cut1.Cut(rg_head_paste1)

                    Dim rg_head_cut2 As Excel.Range = xlWsheetProdMas.Range("B4")
                    Dim rg_head_paste2 As Excel.Range = xlWsheetProdMas.Range("A3")
                    rg_head_cut2.Select()
                    rg_head_cut2.Cut(rg_head_paste2)

                    Dim paramhead1, paramhead2, paramhead3 As String
                    paramhead1 = xlWsheetProdMas.Range("C6").Value & " " & xlWsheetProdMas.Range("F6").Value & "; "
                    paramhead1 = paramhead1 & xlWsheetProdMas.Range("I6").Value & " " & xlWsheetProdMas.Range("L6").Value & "; "

                    paramhead2 = xlWsheetProdMas.Range("R6").Value & " " & xlWsheetProdMas.Range("V6").Value & "; "
                    paramhead2 = paramhead2 & xlWsheetProdMas.Range("Y6").Value & " " & xlWsheetProdMas.Range("AB6").Value & "; "

                    paramhead3 = xlWsheetProdMas.Range("C8").Value & " " & xlWsheetProdMas.Range("F6").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetProdMas.Range("J8").Value & " " & xlWsheetProdMas.Range("L8").Value & "; "
                    paramhead3 = paramhead3 & xlWsheetProdMas.Range("R8").Value & " " & xlWsheetProdMas.Range("V8").Value & "; "

                    xlWsheetProdMas.Range("A5").Value = paramhead1
                    xlWsheetProdMas.Range("A5").EntireRow.Font.Name = "Calibri"
                    xlWsheetProdMas.Range("A6").Value = paramhead2
                    xlWsheetProdMas.Range("A6").EntireRow.Font.Name = "Calibri"
                    xlWsheetProdMas.Range("A7").Value = paramhead3
                    xlWsheetProdMas.Range("A7").EntireRow.Font.Name = "Calibri"

                    xlWsheetProdMas.Range("A2").RowHeight = 27

                    Dim rg1, rg2, rg3, rg4, rg5, rg6, rg7, rg8, rg9 As Excel.Range
                    rg1 = xlWsheetProdMas.Range("B:C")
                    rg1.Select()
                    rg1.Delete()

                    rg2 = xlWsheetProdMas.Range("C:J")
                    rg2.Select()
                    rg2.Delete()

                    rg3 = xlWsheetProdMas.Range("D:F")
                    rg3.Select()
                    rg3.Delete()

                    xlWsheetProdMas.Range("E:E").EntireColumn.Delete()

                    rg4 = xlWsheetProdMas.Range("F:K")
                    rg4.Select()
                    rg4.Delete()

                    rg5 = xlWsheetProdMas.Range("G:H")
                    rg5.Select()
                    rg5.Delete()

                    xlWsheetProdMas.Range("I:I").EntireColumn.Delete()
                    xlWsheetProdMas.Range("AT:AT").EntireColumn.Delete()
                    xlWsheetProdMas.Range("A10").EntireRow.Delete()

                    rg6 = xlWsheetProdMas.Range("A10:A11")
                    xlWsheetProdMas.Range("E10:E11").Copy(rg6)
                    xlWsheetProdMas.Range("A10").Value = "SKU"
                    xlWsheetProdMas.Range("A10").Font.Color = System.Drawing.ColorTranslator.ToOle(Color.White)
                    xlWsheetProdMas.Range("A10:A11").Merge()
                    xlWsheetProdMas.Range("A10:A11").HorizontalAlignment = Excel.Constants.xlCenter

                    rg7 = xlWsheetProdMas.Range("B10:B11")
                    xlWsheetProdMas.Range("E10:E11").Copy(rg7)
                    xlWsheetProdMas.Range("B10").Value = "Description"
                    xlWsheetProdMas.Range("B10").Font.Color = System.Drawing.ColorTranslator.ToOle(Color.White)
                    xlWsheetProdMas.Range("B10:B11").Merge()
                    xlWsheetProdMas.Range("B10:B11").HorizontalAlignment = Excel.Constants.xlCenter

                    rg8 = xlWsheetProdMas.Range("C10:C11")
                    xlWsheetProdMas.Range("E10:E11").Copy(rg8)
                    xlWsheetProdMas.Range("C10").Value = "Product Type"
                    xlWsheetProdMas.Range("C10").Font.Color = System.Drawing.ColorTranslator.ToOle(Color.White)
                    xlWsheetProdMas.Range("C10:C11").Merge()
                    xlWsheetProdMas.Range("C10:C11").HorizontalAlignment = Excel.Constants.xlCenter

                    rg9 = xlWsheetProdMas.Range("D10:D11")
                    xlWsheetProdMas.Range("E10:E11").Copy(rg9)
                    xlWsheetProdMas.Range("D10").Value = "PC/CS"
                    xlWsheetProdMas.Range("D10").Font.Color = System.Drawing.ColorTranslator.ToOle(Color.White)
                    xlWsheetProdMas.Range("D10:D11").Merge()
                    xlWsheetProdMas.Range("D10:D11").HorizontalAlignment = Excel.Constants.xlCenter

                    xlWsheetProdMas.Range("E10:E11").Merge()
                    xlWsheetProdMas.Range("E10:E11").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetProdMas.Range("F10:G10").Merge()
                    xlWsheetProdMas.Range("F10:G10").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetProdMas.Range("F11").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetProdMas.Range("H10:J10").Merge()
                    xlWsheetProdMas.Range("H10:J10").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetProdMas.Range("H11").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetProdMas.Range("K10:N10").Merge()
                    xlWsheetProdMas.Range("K10:N10").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetProdMas.Range("O10:P10").Merge()
                    xlWsheetProdMas.Range("O10:P10").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetProdMas.Range("R10:S10").Merge()
                    xlWsheetProdMas.Range("R10:S10").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetProdMas.Range("AO10:AQ10").Merge()
                    xlWsheetProdMas.Range("AO10:AQ10").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetProdMas.Range("T10:T11").Merge()
                    xlWsheetProdMas.Range("T10:T11").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetProdMas.Range("U10:U11").Merge()
                    xlWsheetProdMas.Range("U10:U11").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetProdMas.Range("V10:V11").Merge()
                    xlWsheetProdMas.Range("V10:V11").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetProdMas.Range("W10:W11").Merge()
                    xlWsheetProdMas.Range("W10:W11").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetProdMas.Range("V10:V11").Merge()
                    xlWsheetProdMas.Range("V10:V11").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetProdMas.Range("X10:X11").Merge()
                    xlWsheetProdMas.Range("X10:X11").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetProdMas.Range("Y10:Y11").Merge()
                    xlWsheetProdMas.Range("Y10:Y11").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetProdMas.Range("Z10:Z11").Merge()
                    xlWsheetProdMas.Range("Z10:Z11").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetProdMas.Range("AA10:AA11").Merge()
                    xlWsheetProdMas.Range("AA10:AA11").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetProdMas.Range("AA10:AA11").WrapText = True
                    xlWsheetProdMas.Range("AB10:AB11").Merge()
                    xlWsheetProdMas.Range("AB10:AB11").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetProdMas.Range("AC10:AC11").Merge()
                    xlWsheetProdMas.Range("AC10:AC11").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetProdMas.Range("AD10:AD11").Merge()
                    xlWsheetProdMas.Range("AD10:AD11").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetProdMas.Range("AE10:AE11").Merge()
                    xlWsheetProdMas.Range("AE10:AE11").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetProdMas.Range("AF10:AF11").Merge()
                    xlWsheetProdMas.Range("AF10:AF11").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetProdMas.Range("AG10:AG11").Merge()
                    xlWsheetProdMas.Range("AG10:AG11").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetProdMas.Range("AH10:AH11").Merge()
                    xlWsheetProdMas.Range("AH10:AH11").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetProdMas.Range("AI10:AI11").Merge()
                    xlWsheetProdMas.Range("AI10:AI11").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetProdMas.Range("AJ10:AJ11").Merge()
                    xlWsheetProdMas.Range("AJ10:AJ11").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetProdMas.Range("AK10:AK11").Merge()
                    xlWsheetProdMas.Range("AK10:AK11").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetProdMas.Range("AL10:Al11").Merge()
                    xlWsheetProdMas.Range("AL10:Al11").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetProdMas.Range("AM10:AM11").Merge()
                    xlWsheetProdMas.Range("AM10:AM11").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetProdMas.Range("AN10:AN11").Merge()
                    xlWsheetProdMas.Range("AN10:AN11").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetProdMas.Range("AR10:AR11").Merge()
                    xlWsheetProdMas.Range("AR10:AR11").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWsheetProdMas.Range("AS10:AS11").Merge()
                    xlWsheetProdMas.Range("AS10:AS11").HorizontalAlignment = Excel.Constants.xlCenter

                    xlWbookProdMas.SaveAs(txtProdMas_dest.Text)
                    xlWbookProdMas.Close()
                    xlAppProdMas.Quit()

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetProdMas)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookProdMas)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppProdMas)

                    xlWsheetProdMas = Nothing
                    xlWbookProdMas = Nothing
                    xlAppProdMas = Nothing

                    MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                Catch ex As Exception
                    MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
        End Select
        
    End Sub
End Class