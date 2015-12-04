Imports Microsoft.Office.Interop
Imports System.Windows.Forms
Imports System.Text
Imports System.Drawing
Imports Microsoft.Win32
Public Class frmOutMas_New
    Dim AppsOffice As String
    Private Sub frmOutMas_New_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim TargetKey As RegistryKey
        TargetKey = Registry.ClassesRoot.OpenSubKey("Excel.Application")
        If TargetKey Is Nothing Then
            AppsOffice = "XL_NotInstalled"
        Else
            AppsOffice = "XL_Installed"
            TargetKey.Close()
        End If
    End Sub

    Private Sub btnBrow_OutMas_src_Click(sender As Object, e As EventArgs) Handles btnBrow_OutMas_src.Click
        Dim outmassrc_path As String
        If OFD_OUTMAS.ShowDialog() = DialogResult.OK Then
            outmassrc_path = OFD_OUTMAS.FileName
            txtOutMas_src.Text = outmassrc_path
        End If
        If txtOutMas_dest.Text <> "" Then
            btnNeu_OutMas.Enabled = True
        Else
            btnNeu_OutMas.Enabled = False
        End If
    End Sub

    Private Sub btnBrow_OutMas_dest_Click(sender As Object, e As EventArgs) Handles btnBrow_OutMas_dest.Click
        Dim outmasdest_path As String
        Dim datenow As DateTime = DateTime.Now
        SFD_OUTMAS.FileName = "Neutralize_OutMas_" & datenow.ToString("ddMMyyyy_HHmm")
        If SFD_OUTMAS.ShowDialog() = DialogResult.OK Then
            outmasdest_path = SFD_OUTMAS.FileName
            txtOutMas_dest.Text = outmasdest_path
        End If
        If txtOutMas_src.Text <> "" Then
            btnNeu_OutMas.Enabled = True
        Else
            btnNeu_OutMas.Enabled = False
        End If
        If txtOutMas_dest.Text <> "" Then
            btnNeu_OutMas.Enabled = True
        Else
            btnNeu_OutMas.Enabled = False
        End If
    End Sub

    Private Sub btnNeu_OutMas_Click(sender As Object, e As EventArgs) Handles btnNeu_OutMas.Click
        PicBar_OutMas.Visible = True
        BWOUTMAS.RunWorkerAsync()
    End Sub

    Private Sub BWOUTMAS_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BWOUTMAS.DoWork
        Select Case AppsOffice
            Case "XL_NotInstalled"
                Dim objApp_OutMas As Object
                Dim objwbook_OutMas As Object
                Dim objwsheet_OutMas As Object

                Try
                    objApp_OutMas = CreateObject("Ket.Application")
                    objwbook_OutMas = objApp_OutMas.Workbook.Open(txtOutMas_src.Text)
                    objwsheet_OutMas = objwbook_OutMas.Worksheets("UID Outlet Master Report")

                    objwsheet_OutMas.UsedRange.UnMerge()
                    objwsheet_OutMas.UsedRange.WrapText = False
                    objwsheet_OutMas.UsedRange.ColumnWidth = 15
                    objwsheet_OutMas.UsedRange.RowHeight = 15

                    Dim rg_head_cut1 As Object = objwsheet_OutMas.Range("C1")
                    Dim rg_head_paste1 As Object = objwsheet_OutMas.Range("A2")
                    rg_head_cut1.Select()
                    rg_head_cut1.Cut(rg_head_paste1)

                    objwsheet_OutMas.Range("A2").RowHeight = 27

                    Dim paramhead1, paramhead2, paramhead3, paramhead4 As String
                    paramhead1 = objwsheet_OutMas.Range("D4").Value & " " & objwsheet_OutMas.Range("J4").Value & "; "
                    paramhead1 = paramhead1 & objwsheet_OutMas.Range("N4").Value & " " & objwsheet_OutMas.Range("U4").Value & "; "

                    paramhead2 = objwsheet_OutMas.Range("Z4").Value & " " & objwsheet_OutMas.Range("AD4").Value & "; "
                    paramhead2 = paramhead2 & objwsheet_OutMas.Range("AI4").Value & " " & objwsheet_OutMas.Range("AL4").Value & "; "

                    paramhead3 = objwsheet_OutMas.Range("AP4").Value & " " & objwsheet_OutMas.Range("AT4").Value & "; "
                    paramhead3 = paramhead3 & objwsheet_OutMas.Range("AX4").Value & " " & objwsheet_OutMas.Range("BB4").Value & "; "
                    paramhead3 = paramhead3 & objwsheet_OutMas.Range("BE4").Value & " " & objwsheet_OutMas.Range("BI4").Value

                    paramhead4 = objwsheet_OutMas.Range("B7").Value & " " & objwsheet_OutMas.Range("I7").Value & "; "
                    paramhead4 = paramhead4 & objwsheet_OutMas.Range("N7").Value & " " & objwsheet_OutMas.Range("S7").Value

                    objwsheet_OutMas.Range("A3").Value = paramhead1
                    objwsheet_OutMas.Range("A3").EntireRow.Font.Name = "Calibri"
                    objwsheet_OutMas.Range("A4").Value = paramhead2
                    objwsheet_OutMas.Range("A4").EntireRow.Font.Name = "Calibri"
                    objwsheet_OutMas.Range("A5").Value = paramhead3
                    objwsheet_OutMas.Range("A5").EntireRow.Font.Name = "Calibri"
                    objwsheet_OutMas.Range("A6").Value = paramhead4
                    objwsheet_OutMas.Range("A6").EntireRow.Font.Name = "Calibri"

                    objwsheet_OutMas.Range("A10").EntireRow.Delete()

                    Dim rg1 As Object = objwsheet_OutMas.Range("B:D")
                    rg1.Select()
                    rg1.Delete()

                    Dim rg2 As Object = objwsheet_OutMas.Range("D:G")
                    rg2.Select()
                    rg2.Delete()

                    Dim rg3 As Object = objwsheet_OutMas.Range("F:G")
                    rg3.Select()
                    rg3.Delete()

                    objwsheet_OutMas.Range("H:H").EntireColumn.Delete()

                    Dim rg4 As Object = objwsheet_OutMas.Range("I:K")
                    rg4.Select()
                    rg4.Delete()

                    Dim rg5 As Object = objwsheet_OutMas.Range("K:M")
                    rg5.Select()
                    rg5.Delete()

                    Dim rg6 As Object = objwsheet_OutMas.Range("L:N")
                    rg6.Select()
                    rg6.Delete()

                    Dim rg7 As Object = objwsheet_OutMas.Range("O:P")
                    rg7.Select()
                    rg7.Delete()

                    Dim rg8 As Object = objwsheet_OutMas.Range("P:Q")
                    rg8.Select()
                    rg8.Delete()

                    Dim rg9 As Object = objwsheet_OutMas.Range("R:S")
                    rg9.Select()
                    rg9.Delete()

                    Dim rg10 As Object = objwsheet_OutMas.Range("T:U")
                    rg10.Select()
                    rg10.Delete()

                    Dim rg11 As Object = objwsheet_OutMas.Range("V:W")
                    rg11.Select()
                    rg11.Delete()

                    objwsheet_OutMas.Range("W:W").EntireColumn.Delete()
                    objwsheet_OutMas.Range("X:X").EntireColumn.Delete()

                    Dim rg12 As Object = objwsheet_OutMas.Range("Y:Z")
                    rg12.Select()
                    rg12.Delete()

                    Dim rg13 As Object = objwsheet_OutMas.Range("AA:AB")
                    rg13.Select()
                    rg13.Delete()

                    objwsheet_OutMas.Range("AC:AC").EntireColumn.Delete()
                    objwsheet_OutMas.Range("AZ:AZ").EntireColumn.Delete()

                    objwbook_OutMas.SaveAs(txtOutMas_dest.Text)
                    objwbook_OutMas.Close()
                    objApp_OutMas.Quit()

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objwsheet_OutMas)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objwbook_OutMas)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objApp_OutMas)

                    objwsheet_OutMas = Nothing
                    objwbook_OutMas = Nothing
                    objApp_OutMas = Nothing

                    MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Catch ex As Exception
                    MessageBox.Show("Error on : " & ex.Message.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try

            Case "XL_Installed"
                Dim objApp As Object
                Dim objwBook As Object
                Dim objwSheet As Object


                Try
                    objApp = CreateObject("Excel.Application")
                    objwBook = objApp.Workbooks.Open(txtOutMas_src.Text)
                    objwSheet = objwBook.Worksheets("UID Outlet Master Report")

                    objwSheet.UsedRange.UnMerge()
                    objwSheet.UsedRange.WrapText = False
                    objwSheet.UsedRange.ColumnWidth = 15
                    objwSheet.UsedRange.RowHeight = 15

                    Dim rg_head_cut1 As Excel.Range = objwSheet.Range("C1")
                    Dim rg_head_paste1 As Excel.Range = objwSheet.Range("A2")
                    rg_head_cut1.Select()
                    rg_head_cut1.Cut(rg_head_paste1)

                    objwSheet.Range("A2").RowHeight = 27

                    Dim paramhead1, paramhead2, paramhead3, paramhead4 As String
                    paramhead1 = objwSheet.Range("D4").Value & " " & objwSheet.Range("J4").Value & "; "
                    paramhead1 = paramhead1 & objwSheet.Range("N4").Value & " " & objwSheet.Range("U4").Value & "; "

                    paramhead2 = objwSheet.Range("Z4").Value & " " & objwSheet.Range("AD4").Value & "; "
                    paramhead2 = paramhead2 & objwSheet.Range("AI4").Value & " " & objwSheet.Range("AL4").Value & "; "

                    paramhead3 = objwSheet.Range("AP4").Value & " " & objwSheet.Range("AT4").Value & "; "
                    paramhead3 = paramhead3 & objwSheet.Range("AX4").Value & " " & objwSheet.Range("BB4").Value & "; "
                    paramhead3 = paramhead3 & objwSheet.Range("BE4").Value & " " & objwSheet.Range("BI4").Value

                    paramhead4 = objwSheet.Range("B7").Value & " " & objwSheet.Range("I7").Value & "; "
                    paramhead4 = paramhead4 & objwSheet.Range("N7").Value & " " & objwSheet.Range("S7").Value

                    objwSheet.Range("A3").Value = paramhead1
                    objwSheet.Range("A3").EntireRow.Font.Name = "Calibri"
                    objwSheet.Range("A4").Value = paramhead2
                    objwSheet.Range("A4").EntireRow.Font.Name = "Calibri"
                    objwSheet.Range("A5").Value = paramhead3
                    objwSheet.Range("A5").EntireRow.Font.Name = "Calibri"
                    objwSheet.Range("A6").Value = paramhead4
                    objwSheet.Range("A6").EntireRow.Font.Name = "Calibri"

                    objwSheet.Range("A10").EntireRow.Delete()

                    Dim rg1 As Excel.Range = objwSheet.Range("B:D")
                    rg1.Select()
                    rg1.Delete()

                    Dim rg2 As Excel.Range = objwSheet.Range("D:G")
                    rg2.Select()
                    rg2.Delete()

                    Dim rg3 As Excel.Range = objwSheet.Range("F:G")
                    rg3.Select()
                    rg3.Delete()

                    objwSheet.Range("H:H").EntireColumn.Delete()

                    Dim rg4 As Excel.Range = objwSheet.Range("I:K")
                    rg4.Select()
                    rg4.Delete()

                    Dim rg5 As Excel.Range = objwSheet.Range("K:M")
                    rg5.Select()
                    rg5.Delete()

                    Dim rg6 As Excel.Range = objwSheet.Range("L:N")
                    rg6.Select()
                    rg6.Delete()

                    Dim rg7 As Excel.Range = objwSheet.Range("O:P")
                    rg7.Select()
                    rg7.Delete()

                    Dim rg8 As Excel.Range = objwSheet.Range("P:Q")
                    rg8.Select()
                    rg8.Delete()

                    Dim rg9 As Excel.Range = objwSheet.Range("R:S")
                    rg9.Select()
                    rg9.Delete()

                    Dim rg10 As Excel.Range = objwSheet.Range("T:U")
                    rg10.Select()
                    rg10.Delete()

                    Dim rg11 As Excel.Range = objwSheet.Range("V:W")
                    rg11.Select()
                    rg11.Delete()

                    objwSheet.Range("W:W").EntireColumn.Delete()
                    objwSheet.Range("X:X").EntireColumn.Delete()

                    Dim rg12 As Excel.Range = objwSheet.Range("Y:Z")
                    rg12.Select()
                    rg12.Delete()

                    Dim rg13 As Excel.Range = objwSheet.Range("AA:AB")
                    rg13.Select()
                    rg13.Delete()

                    objwSheet.Range("AC:AC").EntireColumn.Delete()
                    objwSheet.Range("AZ:AZ").EntireColumn.Delete()

                    objwBook.SaveAs(txtOutMas_dest.Text)
                    objwBook.Close()
                    objApp.Quit()

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objwSheet)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objwBook)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objApp)

                    objwSheet = Nothing
                    objwBook = Nothing
                    objApp = Nothing
                    MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                Catch ex As Exception
                    MessageBox.Show("Error on : " & ex.Message.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try

        End Select
    End Sub

    Private Sub BWOUTMAS_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BWOUTMAS.RunWorkerCompleted
        PicBar_OutMas.Visible = False
        btnNeu_OutMas.Visible = False
        txtOutMas_src.Text = ""
        txtOutMas_dest.Text = ""
    End Sub
End Class