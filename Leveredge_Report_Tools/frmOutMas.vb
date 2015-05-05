Imports Microsoft.Office.Interop
Imports System.Windows.Forms
Imports System.Text
Imports System.Drawing
Imports Microsoft.Win32
Public Class frmOutMas
    Dim AppsOffice As String
    Private Sub frmOutMas_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim TargetKey As RegistryKey
        TargetKey = Registry.ClassesRoot.OpenSubKey("Excel.Application")
        If TargetKey Is Nothing Then
            AppsOffice = "XL_NotInstalled"
        Else
            AppsOffice = "XL_Installed"
            TargetKey.Close()
        End If
    End Sub

    Private Sub btnbrow_outmassrc_Click(sender As Object, e As EventArgs) Handles btnbrow_outmassrc.Click
        Dim outmassrc_path As String
        If OFD_OutMas.ShowDialog() = DialogResult.OK Then
            outmassrc_path = OFD_OutMas.FileName
            txtoutmas_src.Text = outmassrc_path
        End If
        If txtDest_OutMas.Text <> "" Then
            btnNeuOutMas.Enabled = True
        Else
            btnNeuOutMas.Enabled = False
        End If
    End Sub

    Private Sub btnBrow_OutMasDes_Click(sender As Object, e As EventArgs) Handles btnBrow_OutMasDes.Click
        Dim outmasdest_path As String
        Dim datenow As DateTime = DateTime.Now
        SFD_OutMas.FileName = "Neutrailize_OutMas_" & datenow.ToString("ddMMyyyy_HHmm")
        If SFD_OutMas.ShowDialog() = DialogResult.OK Then
            outmasdest_path = SFD_OutMas.FileName
            txtDest_OutMas.Text = outmasdest_path
        End If
        If txtoutmas_src.Text <> "" Then
            btnNeuOutMas.Enabled = True
        Else
            btnNeuOutMas.Enabled = False
        End If
        If txtDest_OutMas.Text <> "" Then
            btnNeuOutMas.Enabled = True
        Else
            btnNeuOutMas.Enabled = False
        End If
    End Sub

    Private Sub btnNeuOutMas_Click(sender As Object, e As EventArgs) Handles btnNeuOutMas.Click
        PicBar.Visible = True
        BWOutMas.RunWorkerAsync()
    End Sub

    Private Sub BWOutMas_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BWOutMas.RunWorkerCompleted
        PicBar.Visible = False
        btnNeuOutMas.Visible = False
        txtoutmas_src.Text = ""
        txtDest_OutMas.Text = ""
    End Sub

    Private Sub BWOutMas_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BWOutMas.DoWork
        Dim objApp As Excel.Application
        Dim objwBook As Excel.Workbook
        Dim objwSheet As Excel.Worksheet


        Try
            objApp = New Excel.Application
            objwBook = objApp.Workbooks.Open(txtoutmas_src.Text)
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

            objwBook.SaveAs(txtDest_OutMas.Text)
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
    End Sub
End Class