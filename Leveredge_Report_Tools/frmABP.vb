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

    Public rep_cond As String

    'Private Sub rbAND_CheckedChanged(sender As Object, e As EventArgs) Handles rbAND.CheckedChanged
    '    rep_cond = ""
    '    rep_cond = "AND"
    'End Sub

    'Private Sub rbOR_CheckedChanged(sender As Object, e As EventArgs) Handles rbOR.CheckedChanged
    '    rep_cond = ""
    '    rep_cond = "OR"
    'End Sub

    Private Sub btnPrevRep_Click(sender As Object, e As EventArgs) Handles btnPrevRep.Click
        If dgridSKU.Rows.Count < 1 Then
            MessageBox.Show("Minimum SKU Harus Di Isi", "STOP", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            Exit Sub
        Else
            If rbAND.Checked = True And rbOR.Checked = False Then
                rep_cond = "AND"
            ElseIf rbAND.Checked = False And rbOR.Checked = True Then
                rep_cond = "OR"
            End If
            SSUpdateDB.Show()
            frmRepABP.Show()
        End If
    End Sub

    Public Sub GetDataABP()
        Select Case rep_cond
            Case "AND"
                '--- First Query AND
                Dim FromDate, EndDate As String
                FromDate = dtpfrom_ABP.Value.ToString("yyyyMMdd")
                EndDate = dtpTo_ABP.Value.ToString("yyyyMMdd")
                Dim com1 As String
                com1 = "DECLARE @FROMDATE DATE,@ENDDATE DATE " _
                    & "SET @FROMDATE = '" & FromDate & "' " _
                    & "SET @ENDDATE = '" & EndDate & "' " _
                    & "SELECT JUM_OUT.PJP,JUM_OUT.NAME AS DSR_NAME, " _
                    & "JUM_OUT.SDESC,JUM_OUT.JUM_POP AS JUMLAH_POP, " _
                    & "CASE WHEN JUM_OUT_ACT.JUM_POP_ACT IS NULL THEN 0 ELSE JUM_OUT_ACT.JUM_POP_ACT END AS ACTIVE_POP, " _
                    & "CASE WHEN JUM_POP_BUY.JUM_OUT_BUY IS NULL THEN 0 ELSE JUM_POP_BUY.JUM_OUT_BUY END AS JUM_OUT_BELI, " _
                    & "CASE WHEN JUM_OUT_ACHV.JUM_POP_ACHV IS NULL THEN 0 ELSE JUM_OUT_ACHV.JUM_POP_ACHV END AS JUM_OUT_ACHV, " _
                    & "CASE WHEN JUM_INV_TOT.JUM_INV IS NULL THEN 0 ELSE JUM_INV_TOT.JUM_INV END AS INVOICE, " _
                    & "CASE WHEN JUM_INV_ACHV.INV_ACHV IS NULL THEN 0 ELSE JUM_INV_ACHV.INV_ACHV END AS JUM_INV_ACHV " _
                    & "FROM (" _
                    & "SELECT PJP,NAME,SDESC,COUNT(PJP) AS JUM_POP " _
                    & "FROM (" _
                    & "SELECT DISTINCT A.PJP, B.NAME, E.SDESC, C.POP " _
                    & "FROM Centegy_SnDPro_UID.dbo.PJP_HEAD A " _
                    & "LEFT JOIN Centegy_SnDPro_UID.dbo.DSR B ON A.DSR=B.DSR " _
                    & "LEFT JOIN Centegy_SnDPro_UID.dbo.SECTION_POP_PERMANENT C ON A.PJP = C.PJP " _
                    & "LEFT JOIN Centegy_SnDPro_UID.dbo.POP D ON C.POP=D.POP " _
                    & "LEFT JOIN Centegy_SnDPro_UID.dbo.JOB_TYPE E ON B.JOB_TYPE=E.JOB_TYPE " _
                    & "GROUP BY A.PJP,B.NAME,E.SDESC,C.POP)POP " _
                    & "GROUP BY PJP,NAME,SDESC)JUM_OUT " _
                    & "LEFT OUTER JOIN " _
                    & "(SELECT PJP,COUNT(PJP) AS JUM_POP_ACT " _
                    & "FROM (" _
                    & "SELECT DISTINCT A.PJP, B.NAME, E.SDESC, C.POP " _
                    & "FROM Centegy_SnDPro_UID.dbo.PJP_HEAD A " _
                    & "INNER JOIN Centegy_SnDPro_UID.dbo.DSR B ON A.DSR=B.DSR " _
                    & "INNER JOIN Centegy_SnDPro_UID.dbo.SECTION_POP_PERMANENT C ON A.PJP = C.PJP " _
                    & "INNER JOIN Centegy_SnDPro_UID.dbo.POP D ON C.POP=D.POP " _
                    & "INNER JOIN Centegy_SnDPro_UID.dbo.JOB_TYPE E ON B.JOB_TYPE=E.JOB_TYPE " _
                    & "WHERE D.ACTIVE='1' " _
                    & "GROUP BY A.PJP,B.NAME,E.SDESC,C.POP)POP_ACT " _
                    & "GROUP BY PJP)JUM_OUT_ACT ON JUM_OUT.PJP=JUM_OUT_ACT.PJP " _
                    & "LEFT OUTER JOIN " _
                    & "(SELECT DISTINCT PJP,COUNT(PJP) AS JUM_OUT_BUY " _
                    & "FROM " _
                    & "(SELECT A.PJP,B.POP FROM Centegy_SnDPro_UID.dbo.PJP_HEAD A " _
                    & "INNER JOIN (SELECT DISTINCT PJP,POP FROM Centegy_SnDPro_UID.dbo.CASHMEMO " _
                    & "WHERE CASHMEMO_TYPE IN ('01','03') AND REF_DOCUMENT = 'GN' " _
                    & "AND DOC_DATE BETWEEN (@FROMDATE) AND (@ENDDATE) " _
                    & "GROUP BY PJP,POP) B ON A.PJP=B.PJP " _
                    & "GROUP BY A.PJP,B.POP) POP_BUY " _
                    & "GROUP BY PJP) JUM_POP_BUY ON JUM_OUT.PJP=JUM_POP_BUY.PJP " _
                    & "LEFT OUTER JOIN " _
                    & "(SELECT DISTINCT PJP, COUNT(PJP) AS JUM_POP_ACHV " _
                    & "FROM (" _
                    & "SELECT DISTINCT DOC_NO,PJP,POP FROM (" _
                    & "SELECT DISTINCT B.DOC_NO,A.PJP,A.POP,B.SKU,(B.QTY1*D.SELL_FACTOR1) + (B.QTY2 * D.SELL_FACTOR2) + B.QTY3 AS QTY_PCS " _
                    & "FROM Centegy_SnDPro_UID.dbo.CASHMEMO A " _
                    & "LEFT JOIN Centegy_SnDPro_UID.dbo.CASHMEMO_DETAIL B ON A.DOC_NO = B.DOC_NO " _
                    & "LEFT JOIN Centegy_SnDPro_UID.dbo.SECTION_POP_PERMANENT C ON A.PJP=C.PJP " _
                    & "AND A.POP = C.POP " _
                    & "LEFT JOIN Centegy_SnDPro_UID.dbo.SKU D ON B.SKU=D.SKU " _
                    & "WHERE A.DOC_DATE BETWEEN (@FROMDATE) AND (@ENDDATE) "

                '--- Get SKU Condition
                Dim totalrow As Integer = dgridSKU.Rows.Count - 1
                Dim sku As String
                Dim com_having As String
                If totalrow >= 2 Then
                    Dim array_SKU As New ArrayList
                    Dim array_QTYSKU As New ArrayList
                    For i As Integer = 0 To totalrow - 1
                        Dim r1 As String = dgridSKU.Rows(i).Cells(0).Value
                        Dim split = r1.Split(" - ")
                        array_SKU.Add(split(0).ToString())
                        Dim qty1, qty2, qty3 As String
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
                        Dim min_qty_com, min_qty As String
                        min_qty_com = "SELECT ((" & qty1 & " * SELL_FACTOR1) + (" & qty2 & " * SELL_FACTOR2) + " & qty3
                        min_qty_com = min_qty_com & ") AS pcsqty FROM SKU WHERE SKU='" & split(0).ToString & "'"
                        Try
                            Get_dbCon()
                            sqlcon.ConnectionString = readdata
                            Dim sqlcom As New SqlCommand(min_qty_com, sqlcon)
                            sqlcon.Open()
                            min_qty = sqlcom.ExecuteScalar
                            array_QTYSKU.Add(min_qty)
                            sqlcon.Close()
                        Catch ex As Exception
                            MessageBox.Show("Error On get Minimum Quantity : " & ex.Message.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try
                        com_having = com_having & "B.SKU='" & array_SKU(i).ToString() & "' AND ((B.QTY1*D.SELL_FACTOR1) + (B.QTY2 * D.SELL_FACTOR2) + B.QTY3) >= "
                        com_having = com_having & array_QTYSKU(i).ToString()
                        If i <> totalrow - 1 Then
                            com_having = com_having & " OR "
                        End If
                        Array.Clear(split, 0, split.Length)
                    Next
                    sku = String.Join("','", array_SKU.ToArray())

                    com1 = com1 & "AND B.SKU IN ('" & sku & "') AND A.CASHMEMO_TYPE IN ('01','03') AND A.REF_DOCUMENT='GN' " _
                        & "GROUP BY B.DOC_NO,A.PJP,A.POP,B.SKU,B.QTY1,B.QTY2,B.QTY3," _
                        & "D.SELL_FACTOR1,D.SELL_FACTOR2 HAVING "

                    Dim com_final As String
                    com_final = com1 & com_having & ") MAS_POP " _
                        & "GROUP BY DOC_NO,PJP,POP " _
                        & "HAVING COUNT(*) = " & totalrow & ") ACH_POP " _
                        & "GROUP BY PJP) JUM_OUT_ACHV ON JUM_OUT.PJP=JUM_OUT_ACHV.PJP " _
                        & "LEFT OUTER JOIN " _
                        & "(SELECT DISTINCT PJP,COUNT(PJP) AS JUM_INV " _
                        & "FROM " _
                        & "(SELECT DISTINCT A.PJP, B.DOC_NO FROM " _
                        & "Centegy_SnDPro_UID.dbo.PJP_HEAD A " _
                        & "LEFT JOIN Centegy_SnDPro_UID.dbo.CASHMEMO B ON A.PJP=B.PJP " _
                        & "WHERE B.DOC_DATE BETWEEN (@FROMDATE) AND (@ENDDATE) " _
                        & "AND B.CASHMEMO_TYPE IN ('01','03') AND B.REF_DOCUMENT='GN') JUM_INV_ALL " _
                        & "GROUP BY PJP) JUM_INV_TOT ON JUM_OUT.PJP = JUM_INV_TOT.PJP " _
                        & "LEFT OUTER JOIN " _
                        & "(SELECT DISTINCT PJP, COUNT(PJP) AS INV_ACHV FROM " _
                        & "(SELECT DISTINCT PJP,DOC_NO FROM " _
                        & "(SELECT A.PJP,B.DOC_NO,(B.QTY1*D.SELL_FACTOR1) + (B.QTY2 * D.SELL_FACTOR2) + B.QTY3 AS QTY_PCS " _
                        & "FROM Centegy_SnDPro_UID.dbo.CASHMEMO A " _
                        & "LEFT JOIN Centegy_SnDPro_UID.dbo.CASHMEMO_DETAIL B ON A.DOC_NO=B.DOC_NO " _
                        & "LEFT JOIN Centegy_SnDPro_UID.dbo.SKU D ON B.SKU = D.SKU " _
                        & "WHERE A.DOC_DATE BETWEEN (@FROMDATE) AND (@ENDDATE) " _
                        & "AND B.SKU IN ('" & sku & "') " _
                        & "AND A.CASHMEMO_TYPE IN ('01','03') AND A.REF_DOCUMENT='GN' " _
                        & "GROUP BY A.PJP,B.DOC_NO,B.SKU,B.QTY1,B.QTY2,B.QTY3,D.SELL_FACTOR1,D.SELL_FACTOR2 HAVING "
                    com_final = com_final & com_having & ") MAS_DOC " _
                        & "GROUP BY PJP,DOC_NO " _
                        & "HAVING COUNT(*) = " & totalrow & ")JUM_INV_ACHV_ALL " _
                        & "GROUP BY PJP) JUM_INV_ACHV ON JUM_INV_TOT.PJP=JUM_INV_ACHV.PJP " _
                        & "WHERE JUM_OUT.SDESC <> 'DM'"

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
                        Dim sComRep As New SqlCommand(com_final, sqlcon)
                        Dim saRep As New SqlDataAdapter(sComRep)
                        saRep.Fill(dsRep)
                        dsRep.Tables(0).TableName = "dtRep"

                        frmRepABP.dtRepBindingSource.DataSource = dsRep
                        sqlcon.Close()
                    Catch ex As Exception
                        MessageBox.Show("Error On Loading Report : " & ex.Message.ToString, "Error Loading Report", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        sqlcon.Close()
                    End Try

                Else
                    Dim array_SKU As New ArrayList
                    Dim array_QTYSKU As New ArrayList
                    For i As Integer = 0 To totalrow - 1
                        Dim r1 As String = dgridSKU.Rows(i).Cells(0).Value
                        Dim split = r1.Split(" - ")
                        array_SKU.Add(split(0).ToString())
                        Dim qty1, qty2, qty3 As String
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
                        Dim min_qty_com, min_qty As String
                        min_qty_com = "SELECT ((" & qty1 & " * SELL_FACTOR1) + (" & qty2 & " * SELL_FACTOR2) + " & qty3
                        min_qty_com = min_qty_com & ") AS pcsqty FROM SKU WHERE SKU='" & split(0).ToString & "'"
                        Try
                            Get_dbCon()
                            sqlcon.ConnectionString = readdata
                            Dim sqlcom As New SqlCommand(min_qty_com, sqlcon)
                            sqlcon.Open()
                            min_qty = sqlcom.ExecuteScalar
                            array_QTYSKU.Add(min_qty)
                            sqlcon.Close()
                        Catch ex As Exception
                            MessageBox.Show("Error On get Minimum Quantity : " & ex.Message.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try
                        com_having = com_having & "B.SKU='" & array_SKU(i).ToString() & "' AND ((B.QTY1*D.SELL_FACTOR1) + (B.QTY2 * D.SELL_FACTOR2) + B.QTY3) >= "
                        com_having = com_having & array_QTYSKU(i).ToString()
                        'If i <> totalrow - 1 Then
                        '    com_having = com_having & "OR "
                        'End If
                        Array.Clear(split, 0, split.Length)
                    Next
                    sku = array_SKU(0).ToString()
                    com1 = com1 & "AND B.SKU IN ('" & sku & "') AND A.CASHMEMO_TYPE IN ('01','03') AND A.REF_DOCUMENT='GN' " _
                        & "GROUP BY B.DOC_NO,A.PJP,A.POP,B.SKU,B.QTY1,B.QTY2,B.QTY3," _
                        & "D.SELL_FACTOR1,D.SELL_FACTOR2 HAVING "

                    Dim com_final As String
                    com_final = com1 & com_having & ") MAS_POP " _
                         & "GROUP BY PJP,POP, DOC_NO) ACH_POP " _
                         & "GROUP BY PJP) JUM_OUT_ACHV ON JUM_OUT.PJP=JUM_OUT_ACHV.PJP " _
                         & "LEFT OUTER JOIN " _
                         & "(SELECT DISTINCT PJP,COUNT(PJP) AS JUM_INV " _
                         & "FROM " _
                         & "(SELECT DISTINCT A.PJP, B.DOC_NO FROM " _
                         & "Centegy_SnDPro_UID.dbo.PJP_HEAD A " _
                         & "LEFT JOIN Centegy_SnDPro_UID.dbo.CASHMEMO B ON A.PJP=B.PJP " _
                         & "WHERE B.DOC_DATE BETWEEN (@FROMDATE) AND (@ENDDATE) " _
                         & "AND B.CASHMEMO_TYPE IN ('01','03') AND B.REF_DOCUMENT='GN') JUM_INV_ALL " _
                         & "GROUP BY PJP) JUM_INV_TOT ON JUM_OUT.PJP = JUM_INV_TOT.PJP " _
                         & "LEFT OUTER JOIN " _
                         & "(SELECT DISTINCT PJP, COUNT(PJP) AS INV_ACHV FROM " _
                         & "(SELECT DISTINCT PJP,DOC_NO FROM " _
                         & "(SELECT A.PJP,B.DOC_NO,(B.QTY1*D.SELL_FACTOR1) + (B.QTY2 * D.SELL_FACTOR2) + B.QTY3 AS QTY_PCS " _
                         & "FROM Centegy_SnDPro_UID.dbo.CASHMEMO A " _
                         & "LEFT JOIN Centegy_SnDPro_UID.dbo.CASHMEMO_DETAIL B ON A.DOC_NO=B.DOC_NO " _
                         & "LEFT JOIN Centegy_SnDPro_UID.dbo.SKU D ON B.SKU = D.SKU " _
                         & "WHERE A.DOC_DATE BETWEEN (@FROMDATE) AND (@ENDDATE) " _
                         & "AND B.SKU IN ('" & sku & "') " _
                         & "AND A.CASHMEMO_TYPE IN ('01','03') AND A.REF_DOCUMENT='GN' " _
                         & "GROUP BY A.PJP,B.DOC_NO,B.SKU,B.QTY1,B.QTY2,B.QTY3,D.SELL_FACTOR1,D.SELL_FACTOR2 HAVING "
                    com_final = com_final & com_having & ") MAS_DOC " _
                        & "GROUP BY PJP,DOC_NO)JUM_INV_ACHV_ALL " _
                        & "GROUP BY PJP) JUM_INV_ACHV ON JUM_INV_TOT.PJP=JUM_INV_ACHV.PJP " _
                        & "WHERE JUM_OUT.SDESC <> 'DM'"

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
                        Dim sComRep As New SqlCommand(com_final, sqlcon)
                        Dim saRep As New SqlDataAdapter(sComRep)
                        saRep.Fill(dsRep)
                        dsRep.Tables(0).TableName = "dtRep"

                        frmRepABP.dtRepBindingSource.DataSource = dsRep
                        sqlcon.Close()
                    Catch ex As Exception
                        MessageBox.Show("Error On Loading Report : " & ex.Message.ToString, "Error Loading Report", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        sqlcon.Close()
                    End Try

                End If

            Case "OR"
                Dim FromDate, EndDate As String
                FromDate = dtpfrom_ABP.Value.ToString("yyyyMMdd")
                EndDate = dtpTo_ABP.Value.ToString("yyyyMMdd")
                Dim com1 As String
                com1 = "DECLARE @FROMDATE DATE,@ENDDATE DATE " _
                    & "SET @FROMDATE = '" & FromDate & "' " _
                    & "SET @ENDDATE = '" & EndDate & "' " _
                    & "SELECT JUM_OUT.PJP,JUM_OUT.NAME AS DSR_NAME," _
                    & "JUM_OUT.SDESC,JUM_OUT.JUM_POP AS JUMLAH_POP," _
                    & "CASE WHEN JUM_OUT_ACT.JUM_POP_ACT IS NULL THEN 0 ELSE JUM_OUT_ACT.JUM_POP_ACT END AS ACTIVE_POP," _
                    & "CASE WHEN JUM_POP_BUY.JUM_OUT_BUY IS NULL THEN 0 ELSE JUM_POP_BUY.JUM_OUT_BUY END AS JUM_OUT_BELI," _
                    & "CASE WHEN JUM_OUT_ACHV.JUM_POP_ACHV IS NULL THEN 0 ELSE JUM_OUT_ACHV.JUM_POP_ACHV END AS JUM_OUT_ACHV," _
                    & "CASE WHEN JUM_INV_TOT.JUM_INV IS NULL THEN 0 ELSE JUM_INV_TOT.JUM_INV END AS INVOICE," _
                    & "CASE WHEN JUM_INV_ACHV.INV_ACHV IS NULL THEN 0 ELSE JUM_INV_ACHV.INV_ACHV END AS JUM_INV_ACHV " _
                    & "FROM (SELECT PJP,NAME,SDESC,COUNT(PJP) AS JUM_POP " _
                    & "FROM (SELECT DISTINCT A.PJP, B.NAME, E.SDESC, C.POP " _
                    & "FROM Centegy_SnDPro_UID.dbo.PJP_HEAD A " _
                    & "LEFT JOIN Centegy_SnDPro_UID.dbo.DSR B ON A.DSR=B.DSR " _
                    & "LEFT JOIN Centegy_SnDPro_UID.dbo.SECTION_POP_PERMANENT C ON A.PJP = C.PJP " _
                    & "LEFT JOIN Centegy_SnDPro_UID.dbo.POP D ON C.POP=D.POP " _
                    & "LEFT JOIN Centegy_SnDPro_UID.dbo.JOB_TYPE E ON B.JOB_TYPE=E.JOB_TYPE " _
                    & "GROUP BY A.PJP,B.NAME,E.SDESC,C.POP)POP " _
                    & "GROUP BY PJP,NAME,SDESC)JUM_OUT " _
                    & "LEFT OUTER JOIN " _
                    & "(SELECT PJP,COUNT(PJP) AS JUM_POP_ACT " _
                    & "FROM(SELECT DISTINCT A.PJP, B.NAME, E.SDESC, C.POP " _
                    & "FROM Centegy_SnDPro_UID.dbo.PJP_HEAD A " _
                    & "INNER JOIN Centegy_SnDPro_UID.dbo.DSR B ON A.DSR=B.DSR " _
                    & "INNER JOIN Centegy_SnDPro_UID.dbo.SECTION_POP_PERMANENT C ON A.PJP = C.PJP " _
                    & "INNER JOIN Centegy_SnDPro_UID.dbo.POP D ON C.POP=D.POP " _
                    & "INNER JOIN Centegy_SnDPro_UID.dbo.JOB_TYPE E ON B.JOB_TYPE=E.JOB_TYPE " _
                    & "WHERE D.ACTIVE='1' " _
                    & "GROUP BY A.PJP,B.NAME,E.SDESC,C.POP)POP_ACT " _
                    & "GROUP BY PJP)JUM_OUT_ACT ON JUM_OUT.PJP=JUM_OUT_ACT.PJP " _
                    & "LEFT OUTER JOIN " _
                    & "(SELECT DISTINCT PJP,COUNT(PJP) AS JUM_OUT_BUY " _
                    & "FROM (SELECT A.PJP,B.POP " _
                    & "FROM Centegy_SnDPro_UID.dbo.PJP_HEAD A " _
                    & "INNER JOIN (SELECT DISTINCT PJP,POP " _
                    & "FROM Centegy_SnDPro_UID.dbo.CASHMEMO " _
                    & "WHERE CASHMEMO_TYPE IN ('01','03') AND REF_DOCUMENT = 'GN' " _
                    & "AND DOC_DATE BETWEEN (@FROMDATE) AND (@ENDDATE) " _
                    & "GROUP BY PJP,POP) B ON A.PJP=B.PJP " _
                    & "GROUP BY A.PJP,B.POP) POP_BUY " _
                    & "GROUP BY PJP) JUM_POP_BUY ON JUM_OUT.PJP=JUM_POP_BUY.PJP  " _
                    & "LEFT OUTER JOIN " _
                    & "(SELECT DISTINCT PJP, COUNT(PJP) AS JUM_POP_ACHV " _
                    & "FROM (SELECT DISTINCT PJP,POP " _
                    & "FROM (" _
                    & "SELECT DISTINCT A.PJP,A.POP,B.DOC_NO,B.SKU,(B.QTY1*D.SELL_FACTOR1) + (B.QTY2 * D.SELL_FACTOR2) + B.QTY3 AS QTY_PCS " _
                    & "FROM Centegy_SnDPro_UID.dbo.CASHMEMO A " _
                    & "LEFT JOIN Centegy_SnDPro_UID.dbo.CASHMEMO_DETAIL B ON A.DOC_NO = B.DOC_NO " _
                    & "LEFT JOIN Centegy_SnDPro_UID.dbo.SECTION_POP_PERMANENT C ON A.PJP=C.PJP " _
                    & "AND A.POP = C.POP " _
                    & "LEFT JOIN Centegy_SnDPro_UID.dbo.SKU D ON B.SKU=D.SKU " _
                    & "WHERE A.DOC_DATE BETWEEN (@FROMDATE) AND (@ENDDATE) "

                Dim totalrow As Integer = dgridSKU.Rows.Count - 1
                Dim sku As String
                Dim com_having As String
                If totalrow >= 2 Then
                    Dim array_SKU As New ArrayList
                    Dim array_QTYSKU As New ArrayList
                    For i As Integer = 0 To totalrow - 1
                        Dim r1 As String = dgridSKU.Rows(i).Cells(0).Value
                        Dim split = r1.Split(" - ")
                        array_SKU.Add(split(0).ToString())
                        Dim qty1, qty2, qty3 As String
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
                        Dim min_qty_com, min_qty As String
                        min_qty_com = "SELECT ((" & qty1 & " * SELL_FACTOR1) + (" & qty2 & " * SELL_FACTOR2) + " & qty3
                        min_qty_com = min_qty_com & ") AS pcsqty FROM SKU WHERE SKU='" & split(0).ToString & "'"
                        Try
                            Get_dbCon()
                            sqlcon.ConnectionString = readdata
                            Dim sqlcom As New SqlCommand(min_qty_com, sqlcon)
                            sqlcon.Open()
                            min_qty = sqlcom.ExecuteScalar
                            array_QTYSKU.Add(min_qty)
                            sqlcon.Close()
                        Catch ex As Exception
                            MessageBox.Show("Error On get Minimum Quantity : " & ex.Message.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try
                        com_having = com_having & "B.SKU='" & array_SKU(i).ToString() & "' AND ((B.QTY1*D.SELL_FACTOR1) + (B.QTY2 * D.SELL_FACTOR2) + B.QTY3) >= "
                        com_having = com_having & array_QTYSKU(i).ToString()
                        If i <> totalrow - 1 Then
                            com_having = com_having & " OR "
                        End If
                        Array.Clear(split, 0, split.Length)
                    Next
                    sku = String.Join("','", array_SKU.ToArray())

                    com1 = com1 & "AND B.SKU IN ('" & sku & "') AND A.CASHMEMO_TYPE IN ('01','03') AND A.REF_DOCUMENT='GN' " _
                        & "GROUP BY B.DOC_NO,A.PJP,A.POP,B.SKU,B.QTY1,B.QTY2,B.QTY3," _
                        & "D.SELL_FACTOR1,D.SELL_FACTOR2 HAVING "

                    Dim com_final As String
                    com_final = com1 & com_having & ") MAS_POP " _
                        & "GROUP BY PJP,POP) ACH_POP " _
                        & "GROUP BY PJP) JUM_OUT_ACHV ON JUM_OUT.PJP=JUM_OUT_ACHV.PJP " _
                        & "LEFT OUTER JOIN " _
                        & "(SELECT DISTINCT PJP,COUNT(PJP) AS JUM_INV " _
                        & "FROM " _
                        & "(SELECT DISTINCT A.PJP, B.DOC_NO FROM " _
                        & "Centegy_SnDPro_UID.dbo.PJP_HEAD A " _
                        & "LEFT JOIN Centegy_SnDPro_UID.dbo.CASHMEMO B ON A.PJP=B.PJP " _
                        & "WHERE B.DOC_DATE BETWEEN (@FROMDATE) AND (@ENDDATE) " _
                        & "AND B.CASHMEMO_TYPE IN ('01','03') AND B.REF_DOCUMENT='GN') JUM_INV_ALL " _
                        & "GROUP BY PJP) JUM_INV_TOT ON JUM_OUT.PJP = JUM_INV_TOT.PJP " _
                        & "LEFT OUTER JOIN " _
                        & "(SELECT DISTINCT PJP, COUNT(PJP) AS INV_ACHV FROM " _
                        & "(SELECT DISTINCT PJP,DOC_NO FROM " _
                        & "(SELECT A.PJP,B.DOC_NO,(B.QTY1*D.SELL_FACTOR1) + (B.QTY2 * D.SELL_FACTOR2) + B.QTY3 AS QTY_PCS " _
                        & "FROM Centegy_SnDPro_UID.dbo.CASHMEMO A " _
                        & "LEFT JOIN Centegy_SnDPro_UID.dbo.CASHMEMO_DETAIL B ON A.DOC_NO=B.DOC_NO " _
                        & "LEFT JOIN Centegy_SnDPro_UID.dbo.SKU D ON B.SKU = D.SKU " _
                        & "WHERE A.DOC_DATE BETWEEN (@FROMDATE) AND (@ENDDATE) " _
                        & "AND B.SKU IN ('" & sku & "') " _
                        & "AND A.CASHMEMO_TYPE IN ('01','03') AND A.REF_DOCUMENT='GN' " _
                        & "GROUP BY A.PJP,B.DOC_NO,B.SKU,B.QTY1,B.QTY2,B.QTY3,D.SELL_FACTOR1,D.SELL_FACTOR2 HAVING "
                    com_final = com_final & com_having & ") MAS_DOC " _
                        & "GROUP BY PJP,DOC_NO)JUM_INV_ACHV_ALL " _
                        & "GROUP BY PJP) JUM_INV_ACHV ON JUM_INV_TOT.PJP=JUM_INV_ACHV.PJP " _
                        & "WHERE JUM_OUT.SDESC <> 'DM'"

                    Dim comhead1 As String
                    comhead1 = "SELECT NAME as Name, (SELECT 'Achievement By Product') AS repname, (SELECT 'OR') AS condmin, "
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
                        Dim sComRep As New SqlCommand(com_final, sqlcon)
                        Dim saRep As New SqlDataAdapter(sComRep)
                        saRep.Fill(dsRep)
                        dsRep.Tables(0).TableName = "dtRep"

                        frmRepABP.dtRepBindingSource.DataSource = dsRep
                        sqlcon.Close()
                    Catch ex As Exception
                        MessageBox.Show("Error On Loading Report : " & ex.Message.ToString, "Error Loading Report", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        sqlcon.Close()
                    End Try

                Else
                    Dim array_SKU As New ArrayList
                    Dim array_QTYSKU As New ArrayList
                    For i As Integer = 0 To totalrow - 1
                        Dim r1 As String = dgridSKU.Rows(i).Cells(0).Value
                        Dim split = r1.Split(" - ")
                        array_SKU.Add(split(0).ToString())
                        Dim qty1, qty2, qty3 As String
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
                        Dim min_qty_com, min_qty As String
                        min_qty_com = "SELECT ((" & qty1 & " * SELL_FACTOR1) + (" & qty2 & " * SELL_FACTOR2) + " & qty3
                        min_qty_com = min_qty_com & ") AS pcsqty FROM SKU WHERE SKU='" & split(0).ToString & "'"
                        Try
                            Get_dbCon()
                            sqlcon.ConnectionString = readdata
                            Dim sqlcom As New SqlCommand(min_qty_com, sqlcon)
                            sqlcon.Open()
                            min_qty = sqlcom.ExecuteScalar
                            array_QTYSKU.Add(min_qty)
                            sqlcon.Close()
                        Catch ex As Exception
                            MessageBox.Show("Error On get Minimum Quantity : " & ex.Message.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try
                        com_having = com_having & "B.SKU='" & array_SKU(i).ToString() & "' AND ((B.QTY1*D.SELL_FACTOR1) + (B.QTY2 * D.SELL_FACTOR2) + B.QTY3) >= "
                        com_having = com_having & array_QTYSKU(i).ToString()
                        'If i <> totalrow - 1 Then
                        '    com_having = com_having & "OR "
                        'End If
                        Array.Clear(split, 0, split.Length)
                    Next
                    sku = array_SKU(0).ToString()
                    com1 = com1 & "AND B.SKU IN ('" & sku & "') AND A.CASHMEMO_TYPE IN ('01','03') AND A.REF_DOCUMENT='GN' " _
                        & "GROUP BY B.DOC_NO,A.PJP,A.POP,B.SKU,B.QTY1,B.QTY2,B.QTY3," _
                        & "D.SELL_FACTOR1,D.SELL_FACTOR2 HAVING "

                    Dim com_final As String
                    com_final = com1 & com_having & ") MAS_POP " _
                         & "GROUP BY PJP,POP) ACH_POP " _
                         & "GROUP BY PJP) JUM_OUT_ACHV ON JUM_OUT.PJP=JUM_OUT_ACHV.PJP " _
                         & "LEFT OUTER JOIN " _
                         & "(SELECT DISTINCT PJP,COUNT(PJP) AS JUM_INV " _
                         & "FROM " _
                         & "(SELECT DISTINCT A.PJP, B.DOC_NO FROM " _
                         & "Centegy_SnDPro_UID.dbo.PJP_HEAD A " _
                         & "LEFT JOIN Centegy_SnDPro_UID.dbo.CASHMEMO B ON A.PJP=B.PJP " _
                         & "WHERE B.DOC_DATE BETWEEN (@FROMDATE) AND (@ENDDATE) " _
                         & "AND B.CASHMEMO_TYPE IN ('01','03') AND B.REF_DOCUMENT='GN') JUM_INV_ALL " _
                         & "GROUP BY PJP) JUM_INV_TOT ON JUM_OUT.PJP = JUM_INV_TOT.PJP " _
                         & "LEFT OUTER JOIN " _
                         & "(SELECT DISTINCT PJP, COUNT(PJP) AS INV_ACHV FROM " _
                         & "(SELECT DISTINCT PJP,DOC_NO FROM " _
                         & "(SELECT A.PJP,B.DOC_NO,(B.QTY1*D.SELL_FACTOR1) + (B.QTY2 * D.SELL_FACTOR2) + B.QTY3 AS QTY_PCS " _
                         & "FROM Centegy_SnDPro_UID.dbo.CASHMEMO A " _
                         & "LEFT JOIN Centegy_SnDPro_UID.dbo.CASHMEMO_DETAIL B ON A.DOC_NO=B.DOC_NO " _
                         & "LEFT JOIN Centegy_SnDPro_UID.dbo.SKU D ON B.SKU = D.SKU " _
                         & "WHERE A.DOC_DATE BETWEEN (@FROMDATE) AND (@ENDDATE) " _
                         & "AND B.SKU IN ('" & sku & "') " _
                         & "AND A.CASHMEMO_TYPE IN ('01','03') AND A.REF_DOCUMENT='GN' " _
                         & "GROUP BY A.PJP,B.DOC_NO,B.SKU,B.QTY1,B.QTY2,B.QTY3,D.SELL_FACTOR1,D.SELL_FACTOR2 HAVING "
                    com_final = com_final & com_having & ") MAS_DOC " _
                        & "GROUP BY PJP,DOC_NO)JUM_INV_ACHV_ALL " _
                        & "GROUP BY PJP) JUM_INV_ACHV ON JUM_INV_TOT.PJP=JUM_INV_ACHV.PJP " _
                        & "WHERE JUM_OUT.SDESC <> 'DM'"

                    Dim comhead1 As String
                    comhead1 = "SELECT NAME as Name, (SELECT 'Achievement By Product') AS repname, (SELECT 'OR') AS condmin, "
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
                        Dim sComRep As New SqlCommand(com_final, sqlcon)
                        Dim saRep As New SqlDataAdapter(sComRep)
                        saRep.Fill(dsRep)
                        dsRep.Tables(0).TableName = "dtRep"

                        frmRepABP.dtRepBindingSource.DataSource = dsRep
                        sqlcon.Close()
                    Catch ex As Exception
                        MessageBox.Show("Error On Loading Report : " & ex.Message.ToString, "Error Loading Report", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        sqlcon.Close()
                    End Try

                End If

        End Select




    End Sub

    'Public Sub GetDataABP()
    '    Dim datefrom, dateto As String
    '    datefrom = dtpfrom_ABP.Value.ToString("yyyyMMdd")
    '    dateto = dtpTo_ABP.Value.ToString("yyyyMMdd")
    '    Dim com1 As String
    '    com1 = "SELECT DISTINCT a.PJP, b.Name AS DSR_NAME, c.SDESC,COUNT(d.POP) AS JUMLAH_POP, "
    '    com1 = com1 & "COUNT(CASE WHEN e.ACTIVE = '1' THEN 'x' ELSE NULL END) AS ACTIVE_POP, f.JUM_OUT_BELI, "
    '    com1 = com1 & "g.JUM_INVOICE AS INVOICE,ISNULL(ts.JUM_OUT,0) AS JUM_OUT_ACHV,ISNULL(ts.JUM_INV,0) AS JUM_INV_ACHV, "
    '    com1 = com1 & "CAST(((CAST(ISNULL(ts.JUM_OUT,0) AS DECIMAL (10,2)) / COUNT(d.POP))*100) AS DECIMAL (10,2)) AS OUT_ACHV_VS_JUM_OUT "
    '    com1 = com1 & "FROM PJP_HEAD a INNER JOIN DSR b ON a.DSR = b.DSR INNER JOIN JOB_TYPE c ON b.JOB_TYPE = c.JOB_TYPE "
    '    com1 = com1 & "INNER JOIN SECTION_POP d ON a.PJP = d.PJP INNER JOIN POP e ON d.POP = e.POP INNER JOIN (SELECT PJP,COUNT(PJP) AS JUM_OUT_BELI FROM "
    '    com1 = com1 & "(SELECT PJP,POP FROM CASHMEMO WHERE CONVERT(DATE,DOC_DATE) BETWEEN CONVERT (DATE,'" & datefrom & "') AND CONVERT(DATE,'" & dateto & "') "
    '    com1 = com1 & "AND REF_DOC_NO IS NOT NULL AND REF_DOCUMENT='GN' GROUP BY PJP,POP) a GROUP BY PJP) f ON a.PJP = f.PJP "
    '    com1 = com1 & "INNER JOIN (SELECT PJP,COUNT(DOC_NO) AS JUM_INVOICE FROM CASHMEMO "
    '    com1 = com1 & "WHERE CONVERT(DATE,DOC_DATE) BETWEEN CONVERT (DATE,'" & datefrom & "') AND CONVERT(DATE,'" & dateto & "') "
    '    com1 = com1 & "AND REF_DOC_NO IS NOT NULL AND REF_DOCUMENT='GN' GROUP BY PJP) g ON a.pjp = g.PJP FULL OUTER JOIN "
    '    com1 = com1 & "(SELECT DISTINCT PJP, COUNT(DISTINCT JUM_OUT) AS JUM_OUT, COUNT(DISTINCT DOC_NO) AS JUM_INV FROM("
    '    com1 = com1 & "SELECT a.PJP,a.POP AS JUM_OUT, b.DOC_NO as DOC_NO,b.SKU, c.PCSQTY "
    '    com1 = com1 & "FROM CASHMEMO a INNER JOIN CASHMEMO_DETAIL b ON a.DOC_NO = b.DOC_NO "
    '    com1 = com1 & "INNER JOIN (SELECT DOC_NO,SKU,((QTY1*SELL_FACTOR1) + (QTY2*SELL_FACTOR2) + QTY3)as PCSQTY "
    '    com1 = com1 & "FROM (SELECT x.DOC_NO,x.SKU,x.QTY1,x.QTY2,x.QTY3,z.SELL_FACTOR1,z.SELL_FACTOR2 "
    '    com1 = com1 & "FROM CASHMEMO_DETAIL x INNER JOIN SKU z ON x.SKU=z.SKU "
    '    com1 = com1 & "WHERE x.SKU IN ("

    '    Dim totalrow As Integer = dgridSKU.RowCount - 1
    '    Dim com2, com3, sku As String
    '    Select Case rep_cond
    '        Case "AND"
    '            If totalrow >= 2 Then
    '                For i As Integer = 0 To totalrow - 1
    '                    Dim qty1, qty2, qty3, pcsqty As String
    '                    If i = 0 Then
    '                        Dim r1 As String = dgridSKU.Rows(i).Cells(0).Value
    '                        Dim split = r1.Split("-")
    '                        'QTY1
    '                        If dgridSKU.Rows(i).Cells(1).Value > 0 Then
    '                            qty1 = dgridSKU.Rows(i).Cells(1).Value
    '                        Else
    '                            qty1 = "0"
    '                        End If
    '                        'QTY2
    '                        If dgridSKU.Rows(i).Cells(2).Value > 0 Then
    '                            qty2 = dgridSKU.Rows(i).Cells(2).Value
    '                        Else
    '                            qty2 = "0"
    '                        End If
    '                        'QTY3
    '                        If dgridSKU.Rows(i).Cells(3).Value > 0 Then
    '                            qty3 = dgridSKU.Rows(i).Cells(3).Value
    '                        Else
    '                            qty3 = "0"
    '                        End If
    '                        sku = "'" & split(0).ToString() & "'"
    '                        Dim min_qty_com, min_qty As String
    '                        min_qty_com = "SELECT ((" & qty1 & " * SELL_FACTOR1) + (" & qty2 & " * SELL_FACTOR2) + " & qty3
    '                        min_qty_com = min_qty_com & ") AS pcsqty FROM SKU WHERE SKU='" & split(0).ToString & "'"
    '                        Try
    '                            Get_dbCon()
    '                            sqlcon.ConnectionString = readdata
    '                            Dim sqlcom As New SqlCommand(min_qty_com, sqlcon)
    '                            sqlcon.Open()
    '                            min_qty = sqlcom.ExecuteScalar
    '                            sqlcon.Close()
    '                        Catch ex As Exception
    '                            MessageBox.Show("Error On : " & ex.Message.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '                        End Try
    '                        com3 = "SKU='" & split(0).ToString() & "' AND PCSQTY >= " & min_qty
    '                        i = i + 1
    '                    End If
    '                    Dim r2 As String = dgridSKU.Rows(i).Cells(0).Value
    '                    Dim split2 = r2.Split("-")
    '                    If dgridSKU.Rows(i).Cells(1).Value > 0 Then
    '                        qty1 = dgridSKU.Rows(i).Cells(1).Value
    '                    Else
    '                        qty1 = "0"
    '                    End If
    '                    'QTY2
    '                    If dgridSKU.Rows(i).Cells(2).Value > 0 Then
    '                        qty2 = dgridSKU.Rows(i).Cells(2).Value
    '                    Else
    '                        qty2 = "0"
    '                    End If
    '                    'QTY3
    '                    If dgridSKU.Rows(i).Cells(3).Value > 0 Then
    '                        qty3 = dgridSKU.Rows(i).Cells(3).Value
    '                    Else
    '                        qty3 = "0"
    '                    End If
    '                    sku = sku & ",'" & split2(0).ToString() & "'"
    '                    Dim min_qty_com2, min_qty2 As String
    '                    min_qty_com2 = "SELECT ((" & qty1 & " * SELL_FACTOR1) + (" & qty2 & " * SELL_FACTOR2) + " & qty3
    '                    min_qty_com2 = min_qty_com2 & ") AS pcsqty FROM SKU WHERE SKU='" & split2(0).ToString & "'"
    '                    Try
    '                        Get_dbCon()
    '                        sqlcon.ConnectionString = readdata
    '                        Dim sqlcom As New SqlCommand(min_qty_com2, sqlcon)
    '                        sqlcon.Open()
    '                        min_qty2 = sqlcom.ExecuteScalar
    '                        sqlcon.Close()
    '                    Catch ex As Exception
    '                        MessageBox.Show("Error On : " & ex.Message.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '                    End Try
    '                    com3 = com3 & " AND SKU='" & split2(0).ToString() & "' AND PCSQTY >= " & min_qty2
    '                    'i = i + 1
    '                    Array.Clear(split2, 0, split2.Length)
    '                Next
    '                com2 = sku & ") GROUP BY x.DOC_NO,x.SKU,x.QTY1,x.QTY2,x.QTY3,z.SELL_FACTOR1,z.SELL_FACTOR2)gg) c ON a.DOC_NO = c.DOC_NO "
    '                com2 = com2 & "WHERE CONVERT(DATE,b.DOC_DATE) BETWEEN CONVERT (DATE,'" & datefrom & "') AND CONVERT(DATE,'" & dateto & "') "
    '                com2 = com2 & "AND b.SKU IN (" & sku & ") GROUP BY a.PJP,a.POP,b.DOC_NO,b.SKU,b.QTY1,b.QTY2,b.QTY3,c.PCSQTY) tt WHERE " & com3
    '                Dim comfinal As String = com1 & com2
    '                comfinal = comfinal & " GROUP BY PJP) ts ON ts.pjp = a.PJP "
    '                comfinal = comfinal & "GROUP BY a.PJP, b.NAME,c.SDESC,f.JUM_OUT_BELI,g.JUM_INVOICE,ts.JUM_OUT,ts.JUM_INV"

    '                Dim comhead1 As String
    '                comhead1 = "SELECT NAME as Name, (SELECT 'Achievement By Product') AS repname, (SELECT 'AND') AS condmin, "
    '                comhead1 = comhead1 & "(SELECT '" & dtpfrom_ABP.Value.ToString("dd MMM yyyy") & " - " & dtpTo_ABP.Value.ToString("dd MMM yyyy")
    '                comhead1 = comhead1 & "') AS Tanggal FROM DISTRIBUTOR"

    '                Try
    '                    Get_dbCon()
    '                    sqlcon.ConnectionString = readdata
    '                    Dim scomHeadRead As New SqlCommand(comhead1, sqlcon)
    '                    Dim saHeadRep As New SqlDataAdapter(scomHeadRead)
    '                    Dim dsHeadRep As New DataSet
    '                    saHeadRep.Fill(dsHeadRep)
    '                    dsHeadRep.Tables(0).TableName = "dtHeadRep"

    '                    frmRepABP.dtHeadRepBindingSource.DataSource = dsHeadRep

    '                    Dim dsRep As New DataSet
    '                    Dim sComRep As New SqlCommand(comfinal, sqlcon)
    '                    Dim saRep As New SqlDataAdapter(sComRep)
    '                    saRep.Fill(dsRep)
    '                    dsRep.Tables(0).TableName = "dtRep"

    '                    frmRepABP.dtRepBindingSource.DataSource = dsRep

    '                Catch ex As Exception
    '                    MessageBox.Show("Error On Loading Report : " & ex.Message.ToString, "Error Loading Report", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '                End Try

    '            ElseIf totalrow < 2 Then
    '                Dim qty1, qty2, qty3 As String
    '                Dim r1 As String = dgridSKU.Rows(0).Cells(0).Value
    '                Dim split = r1.Split("-")
    '                'QTY1
    '                If dgridSKU.Rows(0).Cells(1).Value > 0 Then
    '                    qty1 = dgridSKU.Rows(0).Cells(1).Value
    '                Else
    '                    qty1 = "0"
    '                End If
    '                'QTY2
    '                If dgridSKU.Rows(0).Cells(2).Value > 0 Then
    '                    qty2 = ">=" & dgridSKU.Rows(0).Cells(2).Value
    '                Else
    '                    qty2 = "0"
    '                End If
    '                'QTY3
    '                If dgridSKU.Rows(0).Cells(3).Value > 0 Then
    '                    qty3 = dgridSKU.Rows(0).Cells(3).Value
    '                Else
    '                    qty3 = "0"
    '                End If
    '                sku = "'" & split(0).ToString() & "'"
    '                Dim min_qty_com, min_qty As String
    '                min_qty_com = "SELECT ((" & qty1 & " * SELL_FACTOR1) + (" & qty2 & " * SELL_FACTOR2) + " & qty3
    '                min_qty_com = min_qty_com & ") AS pcsqty FROM SKU WHERE SKU='" & split(0).ToString & "'"
    '                Try
    '                    Get_dbCon()
    '                    sqlcon.ConnectionString = readdata
    '                    Dim sqlcom As New SqlCommand(min_qty_com, sqlcon)
    '                    sqlcon.Open()
    '                    min_qty = sqlcom.ExecuteScalar
    '                    sqlcon.Close()
    '                Catch ex As Exception
    '                    MessageBox.Show("Error On : " & ex.Message.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '                End Try
    '                com3 = "SKU='" & split(0).ToString() & "' AND PCSQTY >= " & min_qty
    '                com2 = sku & ") GROUP BY x.DOC_NO,x.SKU,x.QTY1,x.QTY2,x.QTY3,z.SELL_FACTOR1,z.SELL_FACTOR2)gg) c ON a.DOC_NO = c.DOC_NO "
    '                com2 = com2 & "WHERE CONVERT(DATE,b.DOC_DATE) BETWEEN CONVERT (DATE,'" & datefrom & "') AND CONVERT(DATE,'" & dateto & "') "
    '                com2 = com2 & "AND b.SKU IN (" & sku & ") GROUP BY a.PJP,a.POP,b.DOC_NO,b.SKU,b.QTY1,b.QTY2,b.QTY3,c.PCSQTY) tt WHERE " & com3

    '                Dim comfinal As String = com1 & com2
    '                comfinal = comfinal & " GROUP BY PJP) ts ON ts.pjp = a.PJP "
    '                comfinal = comfinal & "GROUP BY a.PJP, b.NAME,c.SDESC,f.JUM_OUT_BELI,g.JUM_INVOICE,ts.JUM_OUT,ts.JUM_INV"

    '                Dim comhead1 As String
    '                comhead1 = "SELECT NAME as Name, (SELECT 'Achievement By Product') AS repname, (SELECT 'AND') AS condmin, "
    '                comhead1 = comhead1 & "(SELECT '" & dtpfrom_ABP.Value.ToString("dd MMM yyyy") & " - " & dtpTo_ABP.Value.ToString("dd MMM yyyy")
    '                comhead1 = comhead1 & "') AS Tanggal FROM DISTRIBUTOR"

    '                Try
    '                    Get_dbCon()
    '                    sqlcon.ConnectionString = readdata
    '                    Dim scomHeadRead As New SqlCommand(comhead1, sqlcon)
    '                    Dim saHeadRep As New SqlDataAdapter(scomHeadRead)
    '                    Dim dsHeadRep As New DataSet
    '                    saHeadRep.Fill(dsHeadRep)
    '                    dsHeadRep.Tables(0).TableName = "dtHeadRep"

    '                    frmRepABP.dtHeadRepBindingSource.DataSource = dsHeadRep

    '                    Dim dsRep As New DataSet
    '                    Dim sComRep As New SqlCommand(comfinal, sqlcon)
    '                    Dim saRep As New SqlDataAdapter(sComRep)
    '                    saRep.Fill(dsRep)
    '                    dsRep.Tables(0).TableName = "dtRep"

    '                    frmRepABP.dtRepBindingSource.DataSource = dsRep

    '                Catch ex As Exception
    '                    MessageBox.Show("Error On Loading Report : " & ex.Message.ToString, "Error Loading Report", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '                End Try

    '            End If

    '        Case "OR"
    '            If totalrow >= 2 Then
    '                For i As Integer = 0 To totalrow - 1
    '                    Dim qty1, qty2, qty3, pcsqty As String
    '                    If i = 0 Then
    '                        Dim r1 As String = dgridSKU.Rows(i).Cells(0).Value
    '                        Dim split = r1.Split(" - ")
    '                        'QTY1
    '                        If dgridSKU.Rows(i).Cells(1).Value > 0 Then
    '                            qty1 = dgridSKU.Rows(i).Cells(1).Value
    '                        Else
    '                            qty1 = "0"
    '                        End If
    '                        'QTY2
    '                        If dgridSKU.Rows(i).Cells(2).Value > 0 Then
    '                            qty2 = dgridSKU.Rows(i).Cells(2).Value
    '                        Else
    '                            qty2 = "0"
    '                        End If
    '                        'QTY3
    '                        If dgridSKU.Rows(i).Cells(3).Value > 0 Then
    '                            qty3 = dgridSKU.Rows(i).Cells(3).Value
    '                        Else
    '                            qty3 = "0"
    '                        End If
    '                        sku = "'" & split(0).ToString() & "'"
    '                        Dim min_qty_com, min_qty As String
    '                        min_qty_com = "SELECT ((" & qty1 & " * SELL_FACTOR1) + (" & qty2 & " * SELL_FACTOR2) + " & qty3
    '                        min_qty_com = min_qty_com & ") AS pcsqty FROM SKU WHERE SKU='" & split(0).ToString & "'"
    '                        Try
    '                            Get_dbCon()
    '                            sqlcon.ConnectionString = readdata
    '                            Dim sqlcom As New SqlCommand(min_qty_com, sqlcon)
    '                            sqlcon.Open()
    '                            min_qty = sqlcom.ExecuteScalar
    '                            sqlcon.Close()
    '                        Catch ex As Exception
    '                            MessageBox.Show("Error On : " & ex.Message.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '                        End Try
    '                        com3 = "SKU='" & split(0).ToString() & "' AND PCSQTY >= " & min_qty
    '                        i = i + 1
    '                    End If
    '                    Dim r2 As String = dgridSKU.Rows(i).Cells(0).Value
    '                    Dim split2 = r2.Split(" - ")
    '                    If dgridSKU.Rows(i).Cells(1).Value > 0 Then
    '                        qty1 = dgridSKU.Rows(i).Cells(1).Value
    '                    Else
    '                        qty1 = "0"
    '                    End If
    '                    'QTY2
    '                    If dgridSKU.Rows(i).Cells(2).Value > 0 Then
    '                        qty2 = dgridSKU.Rows(i).Cells(2).Value
    '                    Else
    '                        qty2 = "0"
    '                    End If
    '                    'QTY3
    '                    If dgridSKU.Rows(i).Cells(3).Value > 0 Then
    '                        qty3 = dgridSKU.Rows(i).Cells(3).Value
    '                    Else
    '                        qty3 = "0"
    '                    End If
    '                    sku = sku & ",'" & split2(0).ToString() & "'"
    '                    Dim min_qty_com2, min_qty2 As String
    '                    min_qty_com2 = "SELECT ((" & qty1 & " * SELL_FACTOR1) + (" & qty2 & " * SELL_FACTOR2) + " & qty3
    '                    min_qty_com2 = min_qty_com2 & ") AS pcsqty FROM SKU WHERE SKU='" & split2(0).ToString & "'"
    '                    Try
    '                        Get_dbCon()
    '                        sqlcon.ConnectionString = readdata
    '                        Dim sqlcom As New SqlCommand(min_qty_com2, sqlcon)
    '                        sqlcon.Open()
    '                        min_qty2 = sqlcom.ExecuteScalar()
    '                        sqlcon.Close()
    '                    Catch ex As Exception
    '                        MessageBox.Show("Error On : " & ex.Message.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '                    End Try
    '                    com3 = com3 & " OR SKU='" & split2(0).ToString() & "' AND PCSQTY >= " & min_qty2
    '                    'i = i + 1
    '                    Array.Clear(split2, 0, split2.Length)
    '                Next
    '                com2 = sku & ") GROUP BY x.DOC_NO,x.SKU,x.QTY1,x.QTY2,x.QTY3,z.SELL_FACTOR1,z.SELL_FACTOR2)gg) c ON a.DOC_NO = c.DOC_NO "
    '                com2 = com2 & "WHERE CONVERT(DATE,b.DOC_DATE) BETWEEN CONVERT (DATE,'" & datefrom & "') AND CONVERT(DATE,'" & dateto & "') "
    '                com2 = com2 & "AND b.SKU IN (" & sku & ") GROUP BY a.PJP,a.POP,b.DOC_NO,b.SKU,b.QTY1,b.QTY2,b.QTY3,c.PCSQTY) tt WHERE " & com3
    '                Dim comfinal As String = com1 & com2
    '                comfinal = comfinal & " GROUP BY PJP) ts ON ts.pjp = a.PJP "
    '                comfinal = comfinal & "GROUP BY a.PJP, b.NAME,c.SDESC,f.JUM_OUT_BELI,g.JUM_INVOICE,ts.JUM_OUT,ts.JUM_INV"

    '                Dim comhead1 As String
    '                comhead1 = "SELECT NAME as Name, (SELECT 'Achievement By Product') AS repname, (SELECT 'AND') AS condmin, "
    '                comhead1 = comhead1 & "(SELECT '" & dtpfrom_ABP.Value.ToString("dd MMM yyyy") & " - " & dtpTo_ABP.Value.ToString("dd MMM yyyy")
    '                comhead1 = comhead1 & "') AS Tanggal FROM DISTRIBUTOR"

    '                Try
    '                    Get_dbCon()
    '                    sqlcon.ConnectionString = readdata
    '                    Dim scomHeadRead As New SqlCommand(comhead1, sqlcon)
    '                    Dim saHeadRep As New SqlDataAdapter(scomHeadRead)
    '                    Dim dsHeadRep As New DataSet
    '                    saHeadRep.Fill(dsHeadRep)
    '                    dsHeadRep.Tables(0).TableName = "dtHeadRep"

    '                    frmRepABP.dtHeadRepBindingSource.DataSource = dsHeadRep

    '                    Dim dsRep As New DataSet
    '                    Dim sComRep As New SqlCommand(comfinal, sqlcon)
    '                    Dim saRep As New SqlDataAdapter(sComRep)
    '                    saRep.Fill(dsRep)
    '                    dsRep.Tables(0).TableName = "dtRep"

    '                    frmRepABP.dtRepBindingSource.DataSource = dsRep

    '                Catch ex As Exception
    '                    MessageBox.Show("Error On Loading Report : " & ex.Message.ToString, "Error Loading Report", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '                End Try

    '            ElseIf totalrow < 2 Then
    '                Dim qty1, qty2, qty3 As String
    '                Dim r1 As String = dgridSKU.Rows(0).Cells(0).Value
    '                Dim split = r1.Split("-")
    '                'QTY1
    '                If dgridSKU.Rows(0).Cells(1).Value > 0 Then
    '                    qty1 = dgridSKU.Rows(0).Cells(1).Value
    '                Else
    '                    qty1 = "0"
    '                End If
    '                'QTY2
    '                If dgridSKU.Rows(0).Cells(2).Value > 0 Then
    '                    qty2 = dgridSKU.Rows(0).Cells(2).Value
    '                Else
    '                    qty2 = "0"
    '                End If
    '                'QTY3
    '                If dgridSKU.Rows(0).Cells(3).Value > 0 Then
    '                    qty3 = dgridSKU.Rows(0).Cells(3).Value
    '                Else
    '                    qty3 = "0"
    '                End If
    '                sku = "'" & split(0).ToString() & "'"
    '                Dim min_qty_com, min_qty As String
    '                min_qty_com = "SELECT ((" & qty1 & " * SELL_FACTOR1) + (" & qty2 & " * SELL_FACTOR2) + " & qty3
    '                min_qty_com = min_qty_com & ") AS pcsqty FROM SKU WHERE SKU='" & split(0).ToString & "'"
    '                Try
    '                    Get_dbCon()
    '                    sqlcon.ConnectionString = readdata
    '                    Dim sqlcom As New SqlCommand(min_qty_com, sqlcon)
    '                    sqlcon.Open()
    '                    min_qty = sqlcom.ExecuteScalar
    '                    sqlcon.Close()
    '                Catch ex As Exception
    '                    MessageBox.Show("Error On : " & ex.Message.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '                End Try
    '                com3 = "SKU='" & split(0).ToString() & "' AND PCSQTY >= " & min_qty
    '                com2 = sku & ") GROUP BY x.DOC_NO,x.SKU,x.QTY1,x.QTY2,x.QTY3,z.SELL_FACTOR1,z.SELL_FACTOR2)gg) c ON a.DOC_NO = c.DOC_NO "
    '                com2 = com2 & "WHERE CONVERT(DATE,b.DOC_DATE) BETWEEN CONVERT (DATE,'" & datefrom & "') AND CONVERT(DATE,'" & dateto & "') "
    '                com2 = com2 & "AND b.SKU IN (" & sku & ") GROUP BY a.PJP,a.POP,b.DOC_NO,b.SKU,b.QTY1,b.QTY2,b.QTY3,c.PCSQTY) tt WHERE " & com3

    '                Dim comfinal As String = com1 & com2
    '                comfinal = comfinal & " GROUP BY PJP) ts ON ts.pjp = a.PJP "
    '                comfinal = comfinal & "GROUP BY a.PJP, b.NAME,c.SDESC,f.JUM_OUT_BELI,g.JUM_INVOICE,ts.JUM_OUT,ts.JUM_INV"

    '                Dim comhead1 As String
    '                comhead1 = "SELECT NAME as Name, (SELECT 'Achievement By Product') AS repname, (SELECT 'AND') AS condmin, "
    '                comhead1 = comhead1 & "(SELECT '" & dtpfrom_ABP.Value.ToString("dd MMM yyyy") & " - " & dtpTo_ABP.Value.ToString("dd MMM yyyy")
    '                comhead1 = comhead1 & "') AS Tanggal FROM DISTRIBUTOR"

    '                Try
    '                    Get_dbCon()
    '                    sqlcon.ConnectionString = readdata
    '                    Dim scomHeadRead As New SqlCommand(comhead1, sqlcon)
    '                    Dim saHeadRep As New SqlDataAdapter(scomHeadRead)
    '                    Dim dsHeadRep As New DataSet
    '                    saHeadRep.Fill(dsHeadRep)
    '                    dsHeadRep.Tables(0).TableName = "dtHeadRep"

    '                    frmRepABP.dtHeadRepBindingSource.DataSource = dsHeadRep

    '                    Dim dsRep As New DataSet
    '                    Dim sComRep As New SqlCommand(comfinal, sqlcon)
    '                    Dim saRep As New SqlDataAdapter(sComRep)
    '                    saRep.Fill(dsRep)
    '                    dsRep.Tables(0).TableName = "dtRep"

    '                    frmRepABP.dtRepBindingSource.DataSource = dsRep

    '                Catch ex As Exception
    '                    MessageBox.Show("Error On Loading Report : " & ex.Message.ToString, "Error Loading Report", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '                End Try

    '            End If
    '    End Select
    '    sqlcon.Close()
    'End Sub
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