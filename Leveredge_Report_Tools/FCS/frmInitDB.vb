Imports System.Data
Imports System.IO.File
Imports System.Data.SqlClient
Imports Microsoft.SqlServer.Management.Common
Imports Microsoft.SqlServer.Management.Smo

Public Class frmInitDB
    Dim sqlcon As New SqlConnection
    Dim constr As String

    Private Sub GetDBCon()
        Dim strfile As String = My.Application.Info.DirectoryPath & "\Conf.ini"
        constr = ReadAllLines(strfile)(3)
        sqlcon.ConnectionString = constr
    End Sub

    Private Sub GetYearData()
        Dim comstr As String = "select DISTINCT(DATEPART(YYYY,DOC_DATE)) as DataYear from CASHMEMO " _
                               & "GROUP BY DOC_DATE"
        Dim comstr2 As String = "select DISTINCT(DATEPART(YYYY,DOC_DATE)) as DataYear from CASHMEMO " _
                               & "GROUP BY DOC_DATE"
        Try
            GetDBCon()
            sqlcon.Open()
            Dim cmdcek As New SqlCommand(comstr, sqlcon)
            Dim result As String = cmdcek.ExecuteScalar
            If String.IsNullOrEmpty(result) OrElse IsDBNull(result) Then
                cbYear.Enabled = True
                btnInitDB.Enabled = True
            Else
                Dim da As New SqlDataAdapter(comstr2, sqlcon)
                Dim ds As New DataSet
                da.Fill(ds)
                'cbYear.Items.Clear()
                cbYear.DataSource = ds.Tables(0)
                cbYear.ValueMember = "DataYear"
                cbYear.DisplayMember = "DataYear"
                btnCleanDB.Enabled = True
            End If
            sqlcon.Close()
        Catch ex As Exception
            MessageBox.Show("Error on getting data, with message: " & ex.Message.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            sqlcon.Close()
        End Try
    End Sub
    Private Sub frmInitDB_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        System.Windows.Forms.Control.CheckForIllegalCrossThreadCalls = False
        GetYearData()
    End Sub

    Private Sub BW2_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BW2.DoWork

        Dim comstr_Calendar, comstr_cm, comstr_cmdet, comstr_DSR, comstr_jcw, comstr_pjp, comstr_SR, comstr_SPP, _
            comstr_spbd, comstr_SRdet, comstr_trgt, comstr_Dist, comstr_MobTran As String
        comstr_Calendar = "TRUNCATE TABLE Calendar"
        comstr_cm = "TRUNCATE TABLE CASHMEMO"
        comstr_cmdet = "TRUNCATE TABLE CM_DETAIL"
        comstr_DSR = "DELETE FROM DSR"
        comstr_jcw = "TRUNCATE TABLE JC_WEEK"
        comstr_pjp = "TRUNCATE TABLE PJP"
        comstr_SR = "TRUNCATE TABLE SALES_RETURN"
        comstr_SPP = "TRUNCATE TABLE SECTION_POP_PERMANENT"
        comstr_spbd = "TRUNCATE TABLE SECTION_PRMNT_BYDATE"
        comstr_SRdet = "TRUNCATE TABLE SR_DETAIL"
        comstr_trgt = "TRUNCATE TABLE TARGET_BYDSR"
        comstr_Dist = "TRUNCATE TABLE DISTRIBUTOR"
        comstr_MobTran = "TRUNCATE TABLE MOBILITY_TRANS"

        Dim cmd_cal As New SqlCommand(comstr_Calendar, sqlcon)
        Dim cmd_cm As New SqlCommand(comstr_cm, sqlcon)
        Dim cmd_cmdet As New SqlCommand(comstr_cmdet, sqlcon)
        Dim cmd_DSR As New SqlCommand(comstr_DSR, sqlcon)
        Dim cmd_jcw As New SqlCommand(comstr_jcw, sqlcon)
        Dim cmd_pjp As New SqlCommand(comstr_pjp, sqlcon)
        Dim cmd_SR As New SqlCommand(comstr_SR, sqlcon)
        Dim cmd_SPP As New SqlCommand(comstr_SPP, sqlcon)
        Dim cmd_spbd As New SqlCommand(comstr_spbd, sqlcon)
        Dim cmd_SRdet As New SqlCommand(comstr_SRdet, sqlcon)
        Dim cmd_trgt As New SqlCommand(comstr_trgt, sqlcon)
        Dim cmd_Dsit As New SqlCommand(comstr_Dist, sqlcon)

        Try
            GetDBCon()
            sqlcon.Open()
            cmd_cal.ExecuteNonQuery()
            cmd_cm.ExecuteNonQuery()
            cmd_cmdet.ExecuteNonQuery()
            cmd_DSR.ExecuteNonQuery()
            cmd_jcw.ExecuteNonQuery()
            cmd_pjp.ExecuteNonQuery()
            cmd_SR.ExecuteNonQuery()
            cmd_SPP.ExecuteNonQuery()
            cmd_spbd.ExecuteNonQuery()
            cmd_SRdet.ExecuteNonQuery()
            'cmd_trgt.ExecuteNonQuery()
            cmd_Dsit.ExecuteNonQuery()
            sqlcon.Close()
        Catch ex As Exception
            MessageBox.Show("Error on reset DB, with message: " & ex.Message.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            sqlcon.Close()
        End Try
    End Sub

    Private Sub BW2_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BW2.RunWorkerCompleted
        MessageBox.Show("DB has clean", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
        pbProcess.Visible = False
        Call frmMain.FCSDB()
        frmMain.FCSRepMenuItem.Enabled = False
        frmMain.TargetMenuItem.Enabled = False
        frmMain.Refresh()
        Me.Close()
    End Sub

    Private Sub ImportDT()
        Dim comstr As String
        comstr = "INSERT INTO FCS_REPORT.dbo.DISTRIBUTOR(DTCODE,DTNAME,DTADDRESS) " _
            & "SELECT DISTRIBUTOR,NAME,ADDRESS1 + ' ' + ADDRESS2 AS DTADDRESS FROM Centegy_SnDPro_UID.dbo.DISTRIBUTOR"
        Dim cmdImpDT As New SqlCommand(comstr, sqlcon)
        Try
            GetDBCon()
            sqlcon.Open()
            cmdImpDT.ExecuteNonQuery()
            sqlcon.Close()
        Catch ex As Exception
            MessageBox.Show("Error on importing Distributor, with message: " & ex.Message.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            sqlcon.Close()
        End Try
    End Sub

    Private Sub ImportCM()
        Dim comstr As String
        comstr = "insert into FCS_REPORT.dbo.CASHMEMO(DOC_NO,DOC_DATE,DSR,PJP,SECTION,POP,REF_DOC) " _
            & "select DOC_NO,DOC_DATE,DSR,PJP,SECTION,POP,REF_DOCUMENT from Centegy_SnDPro_UID.dbo.CASHMEMO " _
            & "WHERE CASHMEMO_TYPE IN('01','03') AND REF_DOCUMENT='GN' AND DATENAME(YEAR,DOC_DATE)='" & cbYear.SelectedItem.ToString & "'"
        Dim cmdImpCM As New SqlCommand(comstr, sqlcon)
        Try
            GetDBCon()
            sqlcon.Open()
            cmdImpCM.ExecuteNonQuery()
            sqlcon.Close()
        Catch ex As Exception
            MessageBox.Show("Error on importing Cashmemo, with message: " & ex.Message.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            sqlcon.Close()
        End Try
    End Sub

    Private Sub ImportCM_det()
        Dim comstr As String
        comstr = "insert into FCS_REPORT.dbo.CM_DETAIL(DOC_NO,DOC_DATE,SKU) select A.DOC_NO,A.DOC_DATE,B.SKU " _
            & "from Centegy_SnDPro_UID.dbo.CASHMEMO A left join Centegy_SnDPro_UID.dbo.CASHMEMO_DETAIL B ON A.DOC_NO=B.DOC_NO " _
            & "where A.CASHMEMO_TYPE in ('01','03') and A.REF_DOCUMENT='GN' AND B.AMOUNT > 0 AND DATENAME(year,A.DOC_DATE)='" & cbYear.SelectedItem.ToString & "'"
        Dim cmdImpCMdet As New SqlCommand(comstr, sqlcon)
        Try
            GetDBCon()
            sqlcon.Open()
            cmdImpCMdet.ExecuteNonQuery()
            sqlcon.Close()
        Catch ex As Exception
            MessageBox.Show("Error on importing Cashmemo detail, with message: " & ex.Message.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            sqlcon.Close()
        End Try
    End Sub

    Private Sub ImportDSR()
        Dim comstr As String
        comstr = "INSERT INTO FCS_REPORT.dbo.DSR(DSR,NAME,SELLCAT,SELLCAT_DESC,JOB_TYPE) SELECT A.DSR,A.NAME,B.SELL_CATEGORY,C.SDESC,A.JOB_TYPE " _
            & "FROM Centegy_SnDPro_UID.dbo.DSR A INNER JOIN Centegy_SnDPro_UID.dbo.PJP_HEAD B ON A.DSR=B.DSR " _
            & "INNER JOIN Centegy_SnDPro_UID.dbo.SELLING_CATEGORY C ON B.SELL_CATEGORY=C.SELL_CATEGORY " _
            & "WHERE A.JOB_TYPE IN ('01','03')"
        Dim cmdImpDSR As New SqlCommand(comstr, sqlcon)
        Try
            GetDBCon()
            sqlcon.Open()
            cmdImpDSR.ExecuteNonQuery()
            sqlcon.Close()
        Catch ex As Exception
            MessageBox.Show("Error on importing DSR, with message: " & ex.Message.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            sqlcon.Close()
        End Try
    End Sub

    Private Sub ImportJCW()
        Dim comstr As String
        comstr = "insert into FCS_REPORT.dbo.JC_WEEK(YEAR,WEEK_NO,JCNO,START_DATE,END_DATE) " _
            & "SELECT YEAR,substring(WEEK_NO,PATINDEX('%[^0]%',WEEK_NO),10),JCNO,START_DATE,END_DATE " _
            & "FROM Centegy_SnDPro_UID.dbo.JC_WEEK where YEAR='" & cbYear.SelectedItem.ToString & "'"
        Dim cmdImpJCW As New SqlCommand(comstr, sqlcon)
        Try
            GetDBCon()
            sqlcon.Open()
            cmdImpJCW.ExecuteNonQuery()
            sqlcon.Close()
        Catch ex As Exception
            MessageBox.Show("Error on importing JC Week, with message: " & ex.Message.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            sqlcon.Close()
        End Try
    End Sub

    Private Sub ImportPJP()
        Dim comstr As String
        comstr = "INSERT INTO FCS_REPORT.dbo.PJP(PJP,SECTION,WEEK_NO,DSR,SAT,SUN,MON,TUE,WED,THU,FRI) " _
            & "SELECT A.PJP,b.SECTION,b.WEEK_NO,A.DSR,b.SAT,b.SUN,b.MON,b.TUE,b.WED,b.THU,b.FRI " _
            & "FROM Centegy_SnDPro_UID.dbo.PJP_HEAD A INNER JOIN Centegy_SnDPro_UID.dbo.PJP_DETAIL B ON A.PJP=B.PJP " _
            & "INNER JOIN Centegy_SnDPro_UID.dbo.DSR C ON A.DSR=C.DSR WHERE C.JOB_TYPE IN ('01','03')"
        Dim cmdImpPJP As New SqlCommand(comstr, sqlcon)
        Try
            GetDBCon()
            sqlcon.Open()
            cmdImpPJP.ExecuteNonQuery()
            sqlcon.Close()
        Catch ex As Exception
            MessageBox.Show("Error on importing PJP, with message: " & ex.Message.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            sqlcon.Close()
        End Try
    End Sub

    Private Sub ImportSR()
        Dim comstr As String
        comstr = "insert into FCS_REPORT.dbo.SALES_RETURN(DOC_NO,DOC_DATE,DSR,PJP,SECTION,POP,REF_DOC_NO,REF_DOC_DATE) " _
            & "select A.DOC_NO,A.DOC_DATE,A.DSR,A.PJP,A.SECTION,A.POP,A.REF_DOC_NO,B.DOC_DATE AS REF_DOC_DATE from " _
            & "(select DOC_NO,DOC_DATE,DSR,PJP,SECTION,POP,REF_DOC_NO from Centegy_SnDPro_UID.dbo.CASHMEMO WHERE CASHMEMO_TYPE ='19' AND REF_DOCUMENT='CM' " _
            & "AND REF_DOCUMENT2 = 'GR' AND DATENAME(YEAR,DOC_DATE)='" & cbYear.SelectedItem.ToString & "') A INNER JOIN (select DOC_NO,DOC_DATE from Centegy_SnDPro_UID.dbo.CASHMEMO " _
            & "WHERE DOC_NO IN (select REF_DOC_NO from Centegy_SnDPro_UID.dbo.CASHMEMO WHERE CASHMEMO_TYPE ='19' AND REF_DOCUMENT='CM' AND REF_DOCUMENT2 = 'GR' " _
            & "AND DATENAME(YEAR,DOC_DATE)='" & cbYear.SelectedItem.ToString & "')) B ON A.REF_DOC_NO=B.DOC_NO"
        Dim cmdImpSR As New SqlCommand(comstr, sqlcon)
        Try
            GetDBCon()
            sqlcon.Open()
            cmdImpSR.ExecuteNonQuery()
            sqlcon.Close()
        Catch ex As Exception
            MessageBox.Show("Error on importing Sales Return, with message: " & ex.Message.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            sqlcon.Close()
        End Try
    End Sub

    Private Sub ImportSR_Det()
        Dim comstr As String
        comstr = "insert into FCS_REPORT.dbo.SR_DETAIL(DOC_NO,DOC_DATE,SKU,REF_DOC_NO,REF_DOC_DATE)" _
            & "select x.DOC_NO,x.DOC_DATE,x.SKU,x.REF_DOC_NO,y.DOC_DATE AS REF_DOC_DATE from " _
            & "(select A.DOC_NO,convert(DATE,A.DOC_DATE)as DOC_DATE,A.REF_DOC_NO,B.SKU " _
            & "from Centegy_SnDPro_UID.dbo.CASHMEMO A left join Centegy_SnDPro_UID.dbo.CASHMEMO_DETAIL B ON A.DOC_NO=B.DOC_NO AND A.DOC_DATE=B.DOC_DATE " _
            & "where A.REF_DOCUMENT2='GR' and A.CASHMEMO_TYPE='19' AND B.AMOUNT < 0 AND DATENAME(year,A.DOC_DATE)='" & cbYear.SelectedItem.ToString & "') x " _
            & "Inner join Centegy_SnDPro_UID.dbo.CASHMEMO Y ON x.REF_DOC_NO=Y.DOC_NO"
        Dim cmdImpSR_Det As New SqlCommand(comstr, sqlcon)
        Try
            GetDBCon()
            sqlcon.Open()
            cmdImpSR_Det.ExecuteNonQuery()
            sqlcon.Close()
        Catch ex As Exception
            MessageBox.Show("Error on importing Sales Return Detail, with message: " & ex.Message.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            sqlcon.Close()
        End Try
    End Sub

    Private Sub ImportSPP()
        Dim comstr As String
        comstr = "INSERT INTO FCS_REPORT.dbo.SECTION_POP_PERMANENT(PJP,SECTION,POP) SELECT A.PJP,A.SECTION,A.POP " _
            & "FROM Centegy_SnDPro_UID.dbo.SECTION_POP_PERMANENT A INNER JOIN Centegy_SnDPro_UID.dbo.PJP_HEAD B ON A.PJP=B.PJP " _
            & "INNER JOIN Centegy_SnDPro_UID.dbo.PJP_DETAIL C ON B.PJP= C.PJP INNER JOIN Centegy_SnDPro_UID.dbo.DSR D ON B.DSR = D.DSR " _
            & "WHERE D.JOB_TYPE IN ('01','03') GROUP BY A.PJP,A.SECTION,D.DSR,A.POP"
        Dim cmdImpSPP As New SqlCommand(comstr, sqlcon)
        Try
            GetDBCon()
            sqlcon.Open()
            cmdImpSPP.ExecuteNonQuery()
            sqlcon.Close()
        Catch ex As Exception
            MessageBox.Show("Error on importing POP Permanent, with message: " & ex.Message.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            sqlcon.Close()
        End Try
    End Sub

    Private Sub GenerateSPBD()
        Dim comstr As String
        comstr = "insert into SECTION_PRMNT_BYDATE(DATE,PJP,DSR,DSR_NAME,SECTION,WEEK_NO,JCNO,DAY_VISIT,POP) " _
            & "SELECT F.Date,E.PJP,E.DSR,E.DSR_NAME,E.SECTION,E.WEEK_NO,G.JCNO,E.DAY_VISIT,E.POP FROM " _
            & "(select D.PJP,D.DSR,D.DSR_NAME,D.SECTION,D.WEEK_NO, DAY_VISIT= CASE D.DAY when '1000000' then 'Saturday' " _
            & "when '0100000' then 'Sunday' when '0010000' then 'Monday' when '0001000' then 'Tuesday' when '0000100' then 'Wednesday' " _
            & "when '0000010' then 'Thursday' when '0000001' then 'Friday' ELSE 'N/A' END, D.POP FROM " _
            & "(select A.PJP,A.SECTION,B.WEEK_NO, cast(B.SAT as varchar(2))+cast(B.SUN as varchar(2))+" _
            & "cast(B.MON as varchar(2))+cast(B.TUE as varchar(2))+cast(B.WED as varchar(2))+cast(B.THU as varchar(2))+" _
            & "cast(B.FRI as varchar(2)) as DAY,A.POP,B.DSR,C.NAME AS DSR_NAME from SECTION_POP_PERMANENT A " _
            & "inner join PJP B on a.PJP=b.PJP and A.SECTION=b.SECTION inner join DSR C on B.DSR=C.DSR) D) E " _
            & "Inner Join Calendar F ON E.WEEK_NO = F.WEEK_NO AND E.DAY_VISIT=F.DAY_NAME LEFT OUTER join JC_WEEK G on F.WEEK_NO = G.WEEK_NO Order by F.Date ASC"
        Dim cmdGenSPBD As New SqlCommand(comstr, sqlcon)
        Try
            GetDBCon()
            sqlcon.Open()
            cmdGenSPBD.CommandTimeout = 600
            cmdGenSPBD.ExecuteNonQuery()
            sqlcon.Close()
        Catch ex As Exception
            MessageBox.Show("Error on Generating POP Permanent, with message: " & ex.Message.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            sqlcon.Close()
        End Try
    End Sub

    Private Sub btnBrowCal_Click(sender As Object, e As EventArgs) Handles btnBrowCal.Click
        Dim src_sql As String
        If OFDCal.ShowDialog = Windows.Forms.DialogResult.OK Then
            src_sql = OFDCal.FileName
            txtCalPath.Text = src_sql
        End If
    End Sub

    Private Sub ImportCal()
        Dim filepath As String = txtCalPath.Text.Trim
        Dim file As New System.IO.FileInfo(filepath)
        Dim script As String = file.OpenText().ReadToEnd()
        Try
            GetDBCon()
            Dim hostserver As New Server(New ServerConnection(sqlcon))
            hostserver.ConnectionContext.ExecuteNonQuery(script)
        Catch ex As Exception
            MessageBox.Show("Error on executing Script, with message: " & ex.Message.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            sqlcon.Close()
        End Try
        sqlcon.Close()
    End Sub

    'Private Sub ImportMobPlan()
    '    Dim filename As String = txtRCFPath.Text.Trim
    '    Dim xlapp As Object
    '    Dim xlwbook As Object
    '    Dim xlwsheet As Object
    '    Dim DS As New DataSet
    '    Try
    '        xlapp = CreateObject("Ket.Application")
    '        xlwbook = xlapp.Workbooks.Open(filename)
    '        xlwsheet = xlwbook.Worksheets(1)
    '        xlwsheet.Range("A:A").NumberFormat = "dd/MM/yyyy hh:MM:ss;@"
    '        xlwbook.Save()
    '        xlwbook.Close()
    '        xlapp.Quit()

    '        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlwsheet)
    '        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlwbook)
    '        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlapp)

    '        xlwsheet = Nothing
    '        xlwbook = Nothing
    '        xlapp = Nothing

    '        Dim xlapp1 As Object
    '        Dim xlwbook1 As Object
    '        Dim xlwsheet1 As Object
    '        Try
    '            xlapp1 = CreateObject("Ket.Application")
    '            xlwbook1 = xlapp1.Workbooks.Open(filename)
    '            xlwsheet1 = xlwbook1.Worksheets(1)
    '            Dim constr As String = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source='" & filename & "'; " _
    '                              & "Extended Properties=Excel 8.0;"
    '            Dim OleDBCon As New System.Data.OleDb.OleDbConnection(constr)
    '            Dim OleDBComStr As String = "Select * FROM [" & xlwsheet1.Name & "$]"
    '            Dim DA As New System.Data.OleDb.OleDbDataAdapter(OleDBComStr, OleDBCon)
    '            OleDBCon.Open()
    '            DA.TableMappings.Add("Table", xlwsheet1.Name)
    '            DA.Fill(DS)
    '            OleDBCon.Close()
    '            xlwbook1.Close()
    '            xlapp1.Quit()
    '        Catch ex As Exception
    '            MessageBox.Show("Error on transfering Mobility plan, with message: " & ex.Message.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '        End Try
    '        For Each row As DataRow In DS.Tables(0).Rows
    '            Dim JDate As DateTime = row(0)
    '            'If IsDBNull(row(0)) = True Then
    '            '    Dim DateValue As String
    '            '    Dim dr As DataRow = row
    '            '    DateValue = row(0).ToString
    '            '    JDate = DateTime.ParseExact(DateValue, "MM/dd/yyyy", Globalization.CultureInfo.CurrentCulture.DateTimeFormat)
    '            'Else
    '            '    Dim DateValue As String
    '            '    DateValue = row(0)
    '            '    JDate = DateTime.ParseExact(DateValue, "MM/dd/yyyy", Globalization.CultureInfo.CurrentCulture.DateTimeFormat)
    '            'End If

    '            Dim Route As String = row(1).ToString
    '            Dim DSR As String = row(2).ToString
    '            Dim Planned As Integer = Convert.ToInt32(row(3).ToString)
    '            Dim Visited As Integer = Convert.ToInt32(row(4).ToString)
    '            Dim Unplanned As Integer = Convert.ToInt32(row(5).ToString)
    '            Dim ST As String = row(6).ToString
    '            Dim ET As String = row(7).ToString
    '            Dim SPENT As String = row(8).ToString
    '            Dim NOSALE As Integer = Convert.ToInt32(row(9).ToString)
    '            Dim Productive As Integer = Convert.ToInt32(row(10).ToString)
    '            Dim Geo_Mismatch As Integer = Convert.ToInt32(row(11).ToString)
    '            Dim Line As Integer = Convert.ToInt32(row(12).ToString)
    '            Dim TotQTY As Integer = Convert.ToInt32(row(13).ToString)
    '            Dim TotSale As Decimal = Convert.ToDecimal(row(14).ToString)
    '            If IsDBNull(Route) = True Then
    '                Route = ""
    '            End If
    '            If IsDBNull(Route) = True Then
    '                Route = ""
    '            End If
    '            If IsDBNull(DSR) = True Then
    '                DSR = ""
    '            End If
    '            If IsDBNull(Planned) = True Then
    '                Planned = 0
    '            End If
    '            If IsDBNull(Visited) = True Then
    '                Visited = 0
    '            End If
    '            If IsDBNull(Unplanned) = True Then
    '                Unplanned = 0
    '            End If
    '            If IsDBNull(ST) = True Then
    '                ST = ""
    '            End If
    '            If IsDBNull(ET) = True Then
    '                ET = ""
    '            End If
    '            If IsDBNull(SPENT) = True Then
    '                SPENT = ""
    '            End If
    '            If IsDBNull(NOSALE) = True Then
    '                NOSALE = 0
    '            End If
    '            If IsDBNull(Productive) = True Then
    '                Productive = 0
    '            End If
    '            If IsDBNull(Geo_Mismatch) = True Then
    '                Geo_Mismatch = 0
    '            End If
    '            If IsDBNull(Line) = True Then
    '                Line = 0
    '            End If
    '            If IsDBNull(TotQTY) = True Then
    '                TotQTY = 0
    '            End If
    '            If IsDBNull(TotSale) = True Then
    '                TotSale = 0
    '            End If
    '            Dim comstr As String = "INSERT INTO MOBILITY_TRANS(J_DATE,ROUTE_NAME,DSR,PLANNED,VISITED,UNPLANNED, " _
    '                                    & "START_TIME,END_TIME,SPENT,NOSALE,PRODUCTIVE,GEO_MISMATCH,LINE,TOTAL_QTY,TOTAL_SALE)VALUES(" _
    '                                    & "@JDATE,@ROUTE,@DSR,@PLANNED,@VISITED,@UNPLANNED,@ST,@ET,@SPENT,@NOSALE,@PRODUCTIVE,@GEO_MISMATCH," _
    '                                    & "@LINE,@TOTQTY,@TOTSALE)"
    '            Dim cmdimp_Mob_Plan As New SqlCommand(comstr, sqlcon)
    '            cmdimp_Mob_Plan.Parameters.Add("@JDATE", SqlDbType.Date).Value = JDate
    '            cmdimp_Mob_Plan.Parameters.Add("@ROUTE", SqlDbType.NVarChar, 50).Value = Route
    '            cmdimp_Mob_Plan.Parameters.Add("@DSR", SqlDbType.NVarChar, 50).Value = DSR
    '            cmdimp_Mob_Plan.Parameters.Add("@PLANNED", SqlDbType.Int).Value = Planned
    '            cmdimp_Mob_Plan.Parameters.Add("@VISITED", SqlDbType.Int).Value = Visited
    '            cmdimp_Mob_Plan.Parameters.Add("@UNPLANNED", SqlDbType.Int).Value = Unplanned
    '            cmdimp_Mob_Plan.Parameters.Add("@ST", SqlDbType.NVarChar, 50).Value = ST
    '            cmdimp_Mob_Plan.Parameters.Add("@ET", SqlDbType.NVarChar, 50).Value = ET
    '            cmdimp_Mob_Plan.Parameters.Add("@SPENT", SqlDbType.NVarChar, 50).Value = SPENT
    '            cmdimp_Mob_Plan.Parameters.Add("@NOSALE", SqlDbType.Int).Value = NOSALE
    '            cmdimp_Mob_Plan.Parameters.Add("@PRODUCTIVE", SqlDbType.Int).Value = Productive
    '            cmdimp_Mob_Plan.Parameters.Add("@GEO_MISMATCH", SqlDbType.Int).Value = Geo_Mismatch
    '            cmdimp_Mob_Plan.Parameters.Add("@LINE", SqlDbType.Int).Value = Line
    '            cmdimp_Mob_Plan.Parameters.Add("@TOTQTY", SqlDbType.Int).Value = TotQTY
    '            cmdimp_Mob_Plan.Parameters.Add("@TOTSALE", SqlDbType.Int).Value = TotSale

    '            GetDBCon()
    '            sqlcon.Open()
    '            cmdimp_Mob_Plan.ExecuteNonQuery()
    '            sqlcon.Close()
    '        Next row
    '    Catch ex As Exception
    '        MessageBox.Show("Error on importing Mobility Plan, with message: " & ex.Message.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '        sqlcon.Close()
    '    End Try
    'End Sub
    Delegate Sub UpdatelblProgressDelegate(ByVal mystring As String)
    Public Sub UpdateTextlblProgress(ByVal mystring As String)
        Me.lblProgress.Text = mystring
    End Sub


    Private Sub BW1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BW1.DoWork
        Dim UpdatelblProgDel As UpdatelblProgressDelegate = New UpdatelblProgressDelegate(AddressOf UpdateTextlblProgress)
        Dim newtxt As String

        Try
            newtxt = "Importing POP Permanent"
            Me.Invoke(UpdatelblProgDel, newtxt)
            ImportSPP()
            newtxt = "Importing Calendar"
            Me.Invoke(UpdatelblProgDel, newtxt)
            ImportCal()
            newtxt = "Importing Distributor"
            Me.Invoke(UpdatelblProgDel, newtxt)
            ImportDT()
            newtxt = "Importing DSR"
            Me.Invoke(UpdatelblProgDel, newtxt)
            ImportDSR()
            newtxt = "Importing JC Week"
            Me.Invoke(UpdatelblProgDel, newtxt)
            ImportJCW()
            newtxt = "Importing PJP"
            Me.Invoke(UpdatelblProgDel, newtxt)
            ImportPJP()
            newtxt = "Importing Sales Return"
            Me.Invoke(UpdatelblProgDel, newtxt)
            ImportSR()
            newtxt = "Importing Sales Return Detail"
            Me.Invoke(UpdatelblProgDel, newtxt)
            ImportSR_Det()
            newtxt = "Importing Cashmemo"
            Me.Invoke(UpdatelblProgDel, newtxt)
            ImportCM()
            newtxt = "Importing Cashmemo Detail"
            Me.Invoke(UpdatelblProgDel, newtxt)
            ImportCM_det()
            newtxt = "Generating Section POP Permanent By Date"
            Me.Invoke(UpdatelblProgDel, newtxt)
            GenerateSPBD()
        Catch ex As Exception
            MessageBox.Show("Error On Initialize DB, with message : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
            
    End Sub

    Private Sub BW1_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BW1.RunWorkerCompleted
        MessageBox.Show("Initialize DB Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
        lblProgress.Text = ""
        lblProgress.Visible = True
        pbProcess.Visible = False
        txtCalPath.Clear()
        txtRCFPath.Clear()
        frmMain.Refresh()
        frmMain.FCSDB()
        Me.Close()
    End Sub

    Private Sub btnInitDB_Click(sender As Object, e As EventArgs) Handles btnInitDB.Click
        If cbYear.SelectedItem Is Nothing Then
            MessageBox.Show("Please select year that need to be initiate to DB", "STOP", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            cbYear.Focus()
            Exit Sub
        ElseIf txtCalPath.Text = "" Then
            MessageBox.Show("Please insert Calendar file", "STOP", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            btnBrowCal.Focus()
            Exit Sub
            'ElseIf txtRCFPath.Text = "" Then
            '    MessageBox.Show("Please insert Route Control File", "STOP", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            '    btnBrowRCF.Focus()
            '    Exit Sub
        Else
            pbProcess.Visible = True
            lblProgress.Visible = True
            lblProgress.Text = ""
            BW1.RunWorkerAsync()
        End If
    End Sub

    Private Sub btnCleanDB_Click(sender As Object, e As EventArgs) Handles btnCleanDB.Click
        pbProcess.Visible = True
        BW2.RunWorkerAsync()
    End Sub

    Private Sub frmInitDB_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        sqlcon.Close()
    End Sub

    Private Sub btnBrowRCF_Click(sender As Object, e As EventArgs) Handles btnBrowRCF.Click
        Dim src_excel As String
        If OFDRCF.ShowDialog = Windows.Forms.DialogResult.OK Then
            src_excel = OFDRCF.FileName
            txtRCFPath.Text = src_excel
        End If
    End Sub

    Private Sub BW1_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles BW1.ProgressChanged
        'If e.ProgressPercentage = 0 Then
        '    lblProgress.Text = "Importing Calendar"
        'End If
        'If e.ProgressPercentage = 9.090909 Then
        '    lblProgress.Text = "Importing Mobility Plan"
        'End If
        'If e.ProgressPercentage = 18.18182 Then
        '    lblProgress.Text = "Importing Cashmemo"
        'End If
        'If e.ProgressPercentage = 27.27273 Then
        '    lblProgress.Text = "Importing Cashmemo Detail"
        'End If
        'If e.ProgressPercentage = 36.36364 Then
        '    lblProgress.Text = "Importing DSR"
        'End If
        'If e.ProgressPercentage = 45.45455 Then
        '    lblProgress.Text = "Importing JC Week"
        'End If
        'If e.ProgressPercentage = 54.54545 Then
        '    lblProgress.Text = "Importing PJP"
        'End If
        'If e.ProgressPercentage = 63.63636 Then
        '    lblProgress.Text = "Importing Sales Return"
        'End If
        'If e.ProgressPercentage = 72.72727 Then
        '    lblProgress.Text = "Importing Sales Return"
        'End If
        'If e.ProgressPercentage = 81.81818 Then
        '    lblProgress.Text = "Importing Sales Return Detail"
        'End If
        'If e.ProgressPercentage = 90.90909 Then
        '    lblProgress.Text = "Importing Section POP Permanent & Generating Plan"
        'End If
        'If e.ProgressPercentage = 100 Then
        '    lblProgress.Text = "Complete"
        'End If
    End Sub
End Class