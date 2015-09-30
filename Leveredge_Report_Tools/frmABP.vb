Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Data.SqlClient
Imports System.IO.File

Public Class frmABP
    Dim sqlcon As New SqlClient.SqlConnection
    Dim readdata As String

    Private Sub Get_dbCon()
        Dim strfile As String = My.Application.Info.DirectoryPath & "\Conf.ini"
        readdata = ReadAllLines(strfile)(1)
        Dim constring As String
        constring = readdata
        sqlcon.ConnectionString = constring
        Try
            Using sqlcon
                If ConnectionState.Closed Then
                    sqlcon.Open()
                ElseIf ConnectionState.Open Then
                    Exit Sub
                End If
            End Using

        Catch ex As Exception
            MessageBox.Show("Get_dbCon Error on Connection : " & ex.Message.ToString, "SQL Connection Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub init_dgridSKU()

        Dim com As String = "SELECT SKU+' - '+LDESC as SKU FROM SKU ORDER BY LDESC"

        Try
            Dim cmb As New DataGridViewComboBoxColumn()
            Get_dbCon()
            sqlcon.ConnectionString = readdata
            Dim sqlcom As New SqlCommand(com, sqlcon)
            Dim da As New SqlDataAdapter(sqlcom)
            Dim ds As New DataSet
            da.Fill(ds)
            cmb.DataSource = ds.Tables(0)
            cmb.ValueMember = "SKU"
            dgridSKU.Columns.Add(cmb)
            dgridSKU.ColumnCount = 4
            dgridSKU.Columns(0).Name = "SKU"
            dgridSKU.Columns(1).Name = "CS"
            dgridSKU.Columns(1).DefaultCellStyle.NullValue = "0"
            dgridSKU.Columns(2).Name = "DZ"
            dgridSKU.Columns(2).DefaultCellStyle.NullValue = "0"
            dgridSKU.Columns(3).Name = "PC"
            dgridSKU.Columns(3).DefaultCellStyle.NullValue = "0"
            sqlcon.Close()
        Catch ex As Exception
            MessageBox.Show("Error On Init dgridSKU = " & ex.Message.ToString, "Error Init", MessageBoxButtons.OK, MessageBoxIcon.Error)
            sqlcon.Close()
        End Try

    End Sub
    Private Sub frmABP_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        init_dgridSKU()
    End Sub

    Private Sub rbAND_CheckedChanged(sender As Object, e As EventArgs) Handles rbAND.CheckedChanged
        rep_cond = ""
        rep_cond = "AND"
    End Sub

    Private Sub rbOR_CheckedChanged(sender As Object, e As EventArgs) Handles rbOR.CheckedChanged
        rep_cond = ""
        rep_cond = "OR"
    End Sub

    Private Sub btnPrevRep_Click(sender As Object, e As EventArgs) Handles btnPrevRep.Click
        If dgridSKU.Rows.Count < 1 Then
            MessageBox.Show("Minimum SKU Harus Di Isi", "STOP", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            Exit Sub
        Else
            SSloading.Show()
            frmRepABP.Show()
        End If
    End Sub
    Public rep_cond As String
    Public Sub GetDataABP()
        Dim datefrom, dateto As String
        datefrom = dtpfrom_ABP.Value.ToString("yyyyMMdd")
        dateto = dtpTo_ABP.Value.ToString("yyyyMMdd")
        Dim com1 As String
        com1 = "SELECT DISTINCT a.PJP, b.Name AS DSR_NAME, c.SDESC,COUNT(d.POP) AS JUMLAH_POP, "
        com1 = com1 & "COUNT(CASE WHEN e.ACTIVE = '1' THEN 'x' ELSE NULL END) AS ACTIVE_POP, f.JUM_OUT_BELI, "
        com1 = com1 & "g.JUM_INVOICE AS INVOICE,ISNULL(ts.JUM_OUT,0) AS JUM_OUT_ACHV,ISNULL(ts.JUM_INV,0) AS JUM_INV_ACHV, "
        com1 = com1 & "CAST(((CAST(ISNULL(ts.JUM_OUT,0) AS DECIMAL (10,2)) / COUNT(d.POP))*100) AS DECIMAL (10,2)) AS OUT_ACHV_VS_JUM_OUT "
        com1 = com1 & "FROM PJP_HEAD a INNER JOIN DSR b ON a.DSR = b.DSR INNER JOIN JOB_TYPE c ON b.JOB_TYPE = c.JOB_TYPE "
        com1 = com1 & "INNER JOIN SECTION_POP d ON a.PJP = d.PJP INNER JOIN POP e ON d.POP = e.POP INNER JOIN (SELECT PJP,COUNT(PJP) AS JUM_OUT_BELI FROM "
        com1 = com1 & "(SELECT PJP,POP FROM CASHMEMO WHERE CONVERT(DATE,DOC_DATE) BETWEEN CONVERT (DATE,'" & datefrom & "') AND CONVERT(DATE,'" & dateto & "') "
        com1 = com1 & "AND REF_DOC_NO IS NOT NULL AND REF_DOCUMENT='GN' GROUP BY PJP,POP) a GROUP BY PJP) f ON a.PJP = f.PJP "
        com1 = com1 & "INNER JOIN (SELECT PJP,COUNT(DOC_NO) AS JUM_INVOICE FROM CASHMEMO "
        com1 = com1 & "WHERE CONVERT(DATE,DOC_DATE) BETWEEN CONVERT (DATE,'" & datefrom & "') AND CONVERT(DATE,'" & dateto & "') "
        com1 = com1 & "AND REF_DOC_NO IS NOT NULL AND REF_DOCUMENT='GN' GROUP BY PJP) g ON a.pjp = g.PJP FULL OUTER JOIN "
        com1 = com1 & "(SELECT DISTINCT PJP, COUNT(DISTINCT JUM_OUT) AS JUM_OUT, COUNT(DISTINCT DOC_NO) AS JUM_INV FROM("
        com1 = com1 & "SELECT a.PJP,a.POP AS JUM_OUT, b.DOC_NO as DOC_NO,b.SKU, c.PCSQTY "
        com1 = com1 & "FROM CASHMEMO a INNER JOIN CASHMEMO_DETAIL b ON a.DOC_NO = b.DOC_NO "
        com1 = com1 & "INNER JOIN (SELECT DOC_NO,SKU,((QTY1*SELL_FACTOR1) + (QTY2*SELL_FACTOR2) + QTY3)as PCSQTY "
        com1 = com1 & "FROM (SELECT x.DOC_NO,x.SKU,x.QTY1,x.QTY2,x.QTY3,z.SELL_FACTOR1,z.SELL_FACTOR2 "
        com1 = com1 & "FROM CASHMEMO_DETAIL x INNER JOIN SKU z ON x.SKU=z.SKU "
        com1 = com1 & "WHERE x.SKU IN ("

        Dim totalrow As Integer = dgridSKU.RowCount - 1
        Dim com2, com3, sku As String
        Select Case rep_cond
            Case "AND"
                If totalrow >= 2 Then
                    For i As Integer = 0 To totalrow
                        Dim qty1, qty2, qty3, pcsqty As String
                        If i = 0 Then
                            Dim r1 As String = dgridSKU.Rows(i).Cells(0).Value
                            Dim split = r1.Split("-")
                            'QTY1
                            If dgridSKU.Rows(i).Cells(1).Value > 0 Then
                                qty1 = dgridSKU.Rows(i).Cells(1).Value
                            Else
                                qty1 = "0"
                            End If
                            'QTY2
                            If dgridSKU.Rows(i).Cells(2).Value > 0 Then
                                qty2 = dgridSKU.Rows(i).Cells(2).Value
                            Else
                                qty2 = "0"
                            End If
                            'QTY3
                            If dgridSKU.Rows(i).Cells(3).Value > 0 Then
                                qty3 = dgridSKU.Rows(i).Cells(3).Value
                            Else
                                qty3 = "0"
                            End If
                            sku = "'" & split(0).ToString() & "'"
                            Dim min_qty_com, min_qty As String
                            min_qty_com = "SELECT ((" & qty1 & " * SELL_FACTOR1) + (" & qty2 & " * SELL_FACTOR2) + " & qty3
                            min_qty_com = min_qty_com & ") AS pcsqty FROM SKU WHERE SKU='" & split(0).ToString & "'"
                            Try
                                Get_dbCon()
                                sqlcon.ConnectionString = readdata
                                Dim sqlcom As New SqlCommand(min_qty_com, sqlcon)
                                sqlcon.Open()
                                min_qty = sqlcom.ExecuteScalar
                                sqlcon.Close()
                            Catch ex As Exception
                                MessageBox.Show("Error On : " & ex.Message.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            End Try
                            com3 = "SKU='" & split(0).ToString() & "' AND PCSQTY >= " & min_qty
                            i = i + 1
                        End If
                        Dim r2 As String = dgridSKU.Rows(i).Cells(0).Value
                        Dim split2 = r2.Split("-")
                        If dgridSKU.Rows(i).Cells(1).Value > 0 Then
                            qty1 = dgridSKU.Rows(i).Cells(1).Value
                        Else
                            qty1 = "0"
                        End If
                        'QTY2
                        If dgridSKU.Rows(i).Cells(2).Value > 0 Then
                            qty2 = dgridSKU.Rows(i).Cells(2).Value
                        Else
                            qty2 = "0"
                        End If
                        'QTY3
                        If dgridSKU.Rows(i).Cells(3).Value > 0 Then
                            qty3 = dgridSKU.Rows(i).Cells(3).Value
                        Else
                            qty3 = "0"
                        End If
                        sku = sku & ",'" & split2(0).ToString() & "'"
                        Dim min_qty_com2, min_qty2 As String
                        min_qty_com2 = "SELECT ((" & qty1 & " * SELL_FACTOR1) + (" & qty2 & " * SELL_FACTOR2) + " & qty3
                        min_qty_com2 = min_qty_com2 & ") AS pcsqty FROM SKU WHERE SKU='" & split2(0).ToString & "'"
                        Try
                            Get_dbCon()
                            sqlcon.ConnectionString = readdata
                            Dim sqlcom As New SqlCommand(min_qty_com2, sqlcon)
                            sqlcon.Open()
                            min_qty2 = sqlcom.ExecuteScalar
                            sqlcon.Close()
                        Catch ex As Exception
                            MessageBox.Show("Error On : " & ex.Message.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try
                        com3 = com3 & " AND SKU='" & split2(0).ToString() & "' AND PCSQTY >= " & min_qty2
                        i = i + 1
                    Next
                    com2 = sku & ") GROUP BY x.DOC_NO,x.SKU,x.QTY1,x.QTY2,x.QTY3,z.SELL_FACTOR1,z.SELL_FACTOR2)gg) c ON a.DOC_NO = c.DOC_NO "
                    com2 = com2 & "WHERE CONVERT(DATE,b.DOC_DATE) BETWEEN CONVERT (DATE,'" & datefrom & "') AND CONVERT(DATE,'" & dateto & "') "
                    com2 = com2 & "AND b.SKU IN (" & sku & ") GROUP BY a.PJP,a.POP,b.DOC_NO,b.SKU,b.QTY1,b.QTY2,b.QTY3,c.PCSQTY) tt WHERE " & com3
                    Dim comfinal As String = com1 & com2
                    comfinal = comfinal & " GROUP BY PJP) ts ON ts.pjp = a.PJP "
                    comfinal = comfinal & "GROUP BY a.PJP, b.NAME,c.SDESC,f.JUM_OUT_BELI,g.JUM_INVOICE,ts.JUM_OUT,ts.JUM_INV"

                    Dim comhead1 As String
                    comhead1 = "SELECT NAME as Name, (SELECT 'Achievement By Product') AS repname, (SELECT 'AND') AS condmin, "
                    comhead1 = comhead1 & "(SELECT '" & dtpfrom_ABP.Value.ToString("dd MMM yyyy") & " - " & dtpTo_ABP.Value.ToString("dd MMM yyyy")
                    comhead1 = comhead1 & "') AS Tanggal FROM DISTRIBUTOR"

                    Try
                        Get_dbCon()
                        sqlcon.ConnectionString = readdata
                        Dim scomHeadRead As New SqlCommand(comhead1, sqlcon)
                        Dim saHeadRep As New SqlDataAdapter(scomHeadRead)
                        Dim dsHeadRep As New DataSet
                        saHeadRep.Fill(dsHeadRep)
                        dsHeadRep.Tables(0).TableName = "dtHeadRep"

                        frmRepABP.dtHeadRepBindingSource.DataSource = dsHeadRep

                        Dim dsRep As New DataSet
                        Dim sComRep As New SqlCommand(comfinal, sqlcon)
                        Dim saRep As New SqlDataAdapter(sComRep)
                        saRep.Fill(dsRep)
                        dsRep.Tables(0).TableName = "dtRep"

                        frmRepABP.dtRepBindingSource.DataSource = dsRep

                    Catch ex As Exception
                        MessageBox.Show("Error On Loading Report : " & ex.Message.ToString, "Error Loading Report", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End Try

                ElseIf totalrow < 2 Then
                    Dim qty1, qty2, qty3 As String
                    Dim r1 As String = dgridSKU.Rows(0).Cells(0).Value
                    Dim split = r1.Split("-")
                    'QTY1
                    If dgridSKU.Rows(0).Cells(1).Value > 0 Then
                        qty1 = dgridSKU.Rows(0).Cells(1).Value
                    Else
                        qty1 = "0"
                    End If
                    'QTY2
                    If dgridSKU.Rows(0).Cells(2).Value > 0 Then
                        qty2 = ">=" & dgridSKU.Rows(0).Cells(2).Value
                    Else
                        qty2 = "0"
                    End If
                    'QTY3
                    If dgridSKU.Rows(0).Cells(3).Value > 0 Then
                        qty3 = dgridSKU.Rows(0).Cells(3).Value
                    Else
                        qty3 = "0"
                    End If
                    sku = "'" & split(0).ToString() & "'"
                    Dim min_qty_com, min_qty As String
                    min_qty_com = "SELECT ((" & qty1 & " * SELL_FACTOR1) + (" & qty2 & " * SELL_FACTOR2) + " & qty3
                    min_qty_com = min_qty_com & ") AS pcsqty FROM SKU WHERE SKU='" & split(0).ToString & "'"
                    Try
                        Get_dbCon()
                        sqlcon.ConnectionString = readdata
                        Dim sqlcom As New SqlCommand(min_qty_com, sqlcon)
                        sqlcon.Open()
                        min_qty = sqlcom.ExecuteScalar
                        sqlcon.Close()
                    Catch ex As Exception
                        MessageBox.Show("Error On : " & ex.Message.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End Try
                    com3 = "SKU='" & split(0).ToString() & "' AND PCSQTY >= " & min_qty
                    com2 = sku & ") GROUP BY x.DOC_NO,x.SKU,x.QTY1,x.QTY2,x.QTY3,z.SELL_FACTOR1,z.SELL_FACTOR2)gg) c ON a.DOC_NO = c.DOC_NO "
                    com2 = com2 & "WHERE CONVERT(DATE,b.DOC_DATE) BETWEEN CONVERT (DATE,'" & datefrom & "') AND CONVERT(DATE,'" & dateto & "') "
                    com2 = com2 & "AND b.SKU IN (" & sku & ") GROUP BY a.PJP,a.POP,b.DOC_NO,b.SKU,b.QTY1,b.QTY2,b.QTY3,c.PCSQTY) tt WHERE " & com3

                    Dim comfinal As String = com1 & com2
                    comfinal = comfinal & " GROUP BY PJP) ts ON ts.pjp = a.PJP "
                    comfinal = comfinal & "GROUP BY a.PJP, b.NAME,c.SDESC,f.JUM_OUT_BELI,g.JUM_INVOICE,ts.JUM_OUT,ts.JUM_INV"

                    Dim comhead1 As String
                    comhead1 = "SELECT NAME as Name, (SELECT 'Achievement By Product') AS repname, (SELECT 'AND') AS condmin, "
                    comhead1 = comhead1 & "(SELECT '" & dtpfrom_ABP.Value.ToString("dd MMM yyyy") & " - " & dtpTo_ABP.Value.ToString("dd MMM yyyy")
                    comhead1 = comhead1 & "') AS Tanggal FROM DISTRIBUTOR"

                    Try
                        Get_dbCon()
                        sqlcon.ConnectionString = readdata
                        Dim scomHeadRead As New SqlCommand(comhead1, sqlcon)
                        Dim saHeadRep As New SqlDataAdapter(scomHeadRead)
                        Dim dsHeadRep As New DataSet
                        saHeadRep.Fill(dsHeadRep)
                        dsHeadRep.Tables(0).TableName = "dtHeadRep"

                        frmRepABP.dtHeadRepBindingSource.DataSource = dsHeadRep

                        Dim dsRep As New DataSet
                        Dim sComRep As New SqlCommand(comfinal, sqlcon)
                        Dim saRep As New SqlDataAdapter(sComRep)
                        saRep.Fill(dsRep)
                        dsRep.Tables(0).TableName = "dtRep"

                        frmRepABP.dtRepBindingSource.DataSource = dsRep

                    Catch ex As Exception
                        MessageBox.Show("Error On Loading Report : " & ex.Message.ToString, "Error Loading Report", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End Try

                End If

            Case "OR"
                If totalrow >= 2 Then
                    For i As Integer = 0 To totalrow
                        Dim qty1, qty2, qty3, pcsqty As String
                        If i = 0 Then
                            Dim r1 As String = dgridSKU.Rows(i).Cells(0).Value
                            Dim split = r1.Split("-")
                            'QTY1
                            If dgridSKU.Rows(i).Cells(1).Value > 0 Then
                                qty1 = dgridSKU.Rows(i).Cells(1).Value
                            Else
                                qty1 = "0"
                            End If
                            'QTY2
                            If dgridSKU.Rows(i).Cells(2).Value > 0 Then
                                qty2 = dgridSKU.Rows(i).Cells(2).Value
                            Else
                                qty2 = "0"
                            End If
                            'QTY3
                            If dgridSKU.Rows(i).Cells(3).Value > 0 Then
                                qty3 = dgridSKU.Rows(i).Cells(3).Value
                            Else
                                qty3 = "0"
                            End If
                            sku = "'" & split(0).ToString() & "'"
                            Dim min_qty_com, min_qty As String
                            min_qty_com = "SELECT ((" & qty1 & " * SELL_FACTOR1) + (" & qty2 & " * SELL_FACTOR2) + " & qty3
                            min_qty_com = min_qty_com & ") AS pcsqty FROM SKU WHERE SKU='" & split(0).ToString & "'"
                            Try
                                Get_dbCon()
                                sqlcon.ConnectionString = readdata
                                Dim sqlcom As New SqlCommand(min_qty_com, sqlcon)
                                sqlcon.Open()
                                min_qty = sqlcom.ExecuteScalar
                                sqlcon.Close()
                            Catch ex As Exception
                                MessageBox.Show("Error On : " & ex.Message.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            End Try
                            com3 = "SKU='" & split(0).ToString() & "' AND PCSQTY >= " & min_qty
                            i = i + 1
                        End If
                        Dim r2 As String = dgridSKU.Rows(i).Cells(0).Value
                        Dim split2 = r2.Split("-")
                        If dgridSKU.Rows(i).Cells(1).Value > 0 Then
                            qty1 = dgridSKU.Rows(i).Cells(1).Value
                        Else
                            qty1 = "0"
                        End If
                        'QTY2
                        If dgridSKU.Rows(i).Cells(2).Value > 0 Then
                            qty2 = dgridSKU.Rows(i).Cells(2).Value
                        Else
                            qty2 = "0"
                        End If
                        'QTY3
                        If dgridSKU.Rows(i).Cells(3).Value > 0 Then
                            qty3 = dgridSKU.Rows(i).Cells(3).Value
                        Else
                            qty3 = "0"
                        End If
                        sku = sku & ",'" & split2(0).ToString() & "'"
                        Dim min_qty_com2, min_qty2 As String
                        min_qty_com2 = "SELECT ((" & qty1 & " * SELL_FACTOR1) + (" & qty2 & " * SELL_FACTOR2) + " & qty3
                        min_qty_com2 = min_qty_com2 & ") AS pcsqty FROM SKU WHERE SKU='" & split2(0).ToString & "'"
                        Try
                            Get_dbCon()
                            sqlcon.ConnectionString = readdata
                            Dim sqlcom As New SqlCommand(min_qty_com2, sqlcon)
                            sqlcon.Open()
                            min_qty2 = sqlcom.ExecuteScalar()
                            sqlcon.Close()
                        Catch ex As Exception
                            MessageBox.Show("Error On : " & ex.Message.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try
                        com3 = com3 & " OR SKU='" & split2(0).ToString() & "' AND PCSQTY >= " & min_qty2
                        i = i + 1
                    Next
                    com2 = sku & ") GROUP BY x.DOC_NO,x.SKU,x.QTY1,x.QTY2,x.QTY3,z.SELL_FACTOR1,z.SELL_FACTOR2)gg) c ON a.DOC_NO = c.DOC_NO "
                    com2 = com2 & "WHERE CONVERT(DATE,b.DOC_DATE) BETWEEN CONVERT (DATE,'" & datefrom & "') AND CONVERT(DATE,'" & dateto & "') "
                    com2 = com2 & "AND b.SKU IN (" & sku & ") GROUP BY a.PJP,a.POP,b.DOC_NO,b.SKU,b.QTY1,b.QTY2,b.QTY3,c.PCSQTY) tt WHERE " & com3
                    Dim comfinal As String = com1 & com2
                    comfinal = comfinal & " GROUP BY PJP) ts ON ts.pjp = a.PJP "
                    comfinal = comfinal & "GROUP BY a.PJP, b.NAME,c.SDESC,f.JUM_OUT_BELI,g.JUM_INVOICE,ts.JUM_OUT,ts.JUM_INV"

                    Dim comhead1 As String
                    comhead1 = "SELECT NAME as Name, (SELECT 'Achievement By Product') AS repname, (SELECT 'AND') AS condmin, "
                    comhead1 = comhead1 & "(SELECT '" & dtpfrom_ABP.Value.ToString("dd MMM yyyy") & " - " & dtpTo_ABP.Value.ToString("dd MMM yyyy")
                    comhead1 = comhead1 & "') AS Tanggal FROM DISTRIBUTOR"

                    Try
                        Get_dbCon()
                        sqlcon.ConnectionString = readdata
                        Dim scomHeadRead As New SqlCommand(comhead1, sqlcon)
                        Dim saHeadRep As New SqlDataAdapter(scomHeadRead)
                        Dim dsHeadRep As New DataSet
                        saHeadRep.Fill(dsHeadRep)
                        dsHeadRep.Tables(0).TableName = "dtHeadRep"

                        frmRepABP.dtHeadRepBindingSource.DataSource = dsHeadRep

                        Dim dsRep As New DataSet
                        Dim sComRep As New SqlCommand(comfinal, sqlcon)
                        Dim saRep As New SqlDataAdapter(sComRep)
                        saRep.Fill(dsRep)
                        dsRep.Tables(0).TableName = "dtRep"

                        frmRepABP.dtRepBindingSource.DataSource = dsRep

                    Catch ex As Exception
                        MessageBox.Show("Error On Loading Report : " & ex.Message.ToString, "Error Loading Report", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End Try

                ElseIf totalrow < 2 Then
                    Dim qty1, qty2, qty3 As String
                    Dim r1 As String = dgridSKU.Rows(0).Cells(0).Value
                    Dim split = r1.Split("-")
                    'QTY1
                    If dgridSKU.Rows(0).Cells(1).Value > 0 Then
                        qty1 = dgridSKU.Rows(0).Cells(1).Value
                    Else
                        qty1 = "0"
                    End If
                    'QTY2
                    If dgridSKU.Rows(0).Cells(2).Value > 0 Then
                        qty2 = dgridSKU.Rows(0).Cells(2).Value
                    Else
                        qty2 = "0"
                    End If
                    'QTY3
                    If dgridSKU.Rows(0).Cells(3).Value > 0 Then
                        qty3 = dgridSKU.Rows(0).Cells(3).Value
                    Else
                        qty3 = "0"
                    End If
                    sku = "'" & split(0).ToString() & "'"
                    Dim min_qty_com, min_qty As String
                    min_qty_com = "SELECT ((" & qty1 & " * SELL_FACTOR1) + (" & qty2 & " * SELL_FACTOR2) + " & qty3
                    min_qty_com = min_qty_com & ") AS pcsqty FROM SKU WHERE SKU='" & split(0).ToString & "'"
                    Try
                        Get_dbCon()
                        sqlcon.ConnectionString = readdata
                        Dim sqlcom As New SqlCommand(min_qty_com, sqlcon)
                        sqlcon.Open()
                        min_qty = sqlcom.ExecuteScalar
                        sqlcon.Close()
                    Catch ex As Exception
                        MessageBox.Show("Error On : " & ex.Message.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End Try
                    com3 = "SKU='" & split(0).ToString() & "' AND PCSQTY >= " & min_qty
                    com2 = sku & ") GROUP BY x.DOC_NO,x.SKU,x.QTY1,x.QTY2,x.QTY3,z.SELL_FACTOR1,z.SELL_FACTOR2)gg) c ON a.DOC_NO = c.DOC_NO "
                    com2 = com2 & "WHERE CONVERT(DATE,b.DOC_DATE) BETWEEN CONVERT (DATE,'" & datefrom & "') AND CONVERT(DATE,'" & dateto & "') "
                    com2 = com2 & "AND b.SKU IN (" & sku & ") GROUP BY a.PJP,a.POP,b.DOC_NO,b.SKU,b.QTY1,b.QTY2,b.QTY3,c.PCSQTY) tt WHERE " & com3

                    Dim comfinal As String = com1 & com2
                    comfinal = comfinal & " GROUP BY PJP) ts ON ts.pjp = a.PJP "
                    comfinal = comfinal & "GROUP BY a.PJP, b.NAME,c.SDESC,f.JUM_OUT_BELI,g.JUM_INVOICE,ts.JUM_OUT,ts.JUM_INV"

                    Dim comhead1 As String
                    comhead1 = "SELECT NAME as Name, (SELECT 'Achievement By Product') AS repname, (SELECT 'AND') AS condmin, "
                    comhead1 = comhead1 & "(SELECT '" & dtpfrom_ABP.Value.ToString("dd MMM yyyy") & " - " & dtpTo_ABP.Value.ToString("dd MMM yyyy")
                    comhead1 = comhead1 & "') AS Tanggal FROM DISTRIBUTOR"

                    Try
                        Get_dbCon()
                        sqlcon.ConnectionString = readdata
                        Dim scomHeadRead As New SqlCommand(comhead1, sqlcon)
                        Dim saHeadRep As New SqlDataAdapter(scomHeadRead)
                        Dim dsHeadRep As New DataSet
                        saHeadRep.Fill(dsHeadRep)
                        dsHeadRep.Tables(0).TableName = "dtHeadRep"

                        frmRepABP.dtHeadRepBindingSource.DataSource = dsHeadRep

                        Dim dsRep As New DataSet
                        Dim sComRep As New SqlCommand(comfinal, sqlcon)
                        Dim saRep As New SqlDataAdapter(sComRep)
                        saRep.Fill(dsRep)
                        dsRep.Tables(0).TableName = "dtRep"

                        frmRepABP.dtRepBindingSource.DataSource = dsRep

                    Catch ex As Exception
                        MessageBox.Show("Error On Loading Report : " & ex.Message.ToString, "Error Loading Report", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End Try

                End If
        End Select
        sqlcon.Close()
    End Sub
    'Dim combogrid As ComboBox
    'Private Sub dgridSKU_EditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs) Handles dgridSKU.EditingControlShowing
    '    '-----Error Right Here (Failed to convert from text to combobox)----------------
    '    combogrid = CType(e.Control, ComboBox)
    '    '-------------------------------------
    '    If dgridSKU.CurrentCellAddress.X > 0 Then
    '        Exit Sub
    '    End If
    '    If combogrid IsNot Nothing Then
    '        combogrid.AutoCompleteMode = AutoCompleteMode.SuggestAppend
    '        combogrid.AutoCompleteSource = AutoCompleteSource.ListItems
    '    End If
    'End Sub
End Class