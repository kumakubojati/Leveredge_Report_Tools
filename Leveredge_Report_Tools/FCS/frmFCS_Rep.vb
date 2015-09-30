Imports System.Data
Imports System.Data.SqlClient
Imports System.IO.File

Public Class frmFCS_Rep

    Dim sqlcon As New SqlConnection
    Dim constr As String

    Public Sub GetDBCon()
        Dim strfile As String = My.Application.Info.DirectoryPath & "\Conf.ini"
        constr = ReadAllLines(strfile)(3)
        sqlcon.ConnectionString = constr
    End Sub

    Public Sub loadclbDSR()
        Dim comstr As String = "SELECT DSR + ' - ' + NAME AS DSR FROM DSR"
        Dim da As New SqlDataAdapter(comstr, sqlcon)
        Dim dt As New DataTable
        Try
            GetDBCon()
            sqlcon.Open()
            da.Fill(dt)
            For Each row As DataRow In dt.Rows
                Dim item As String = Convert.ToString(row("DSR"))
                clbDSR.Items.Add(item)
            Next row
            sqlcon.Close()
        Catch ex As Exception
            MessageBox.Show("Error in DSR Check list box, with message : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            sqlcon.Close()
        End Try
    End Sub

    Public Sub Year()
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
                cbYear.Items.Add("")
            Else
                Dim da As New SqlDataAdapter(comstr2, sqlcon)
                Dim ds As New DataSet
                da.Fill(ds)
                'cbYear.Items.Clear()
                cbYear.DataSource = ds.Tables(0)
                cbYear.ValueMember = "DataYear"
                cbYear.DisplayMember = "DataYear"
            End If
            sqlcon.Close()
        Catch ex As Exception
            MessageBox.Show("Error on getting year data, with message: " & ex.Message.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            sqlcon.Close()
        End Try
    End Sub

    Public Sub period()
        Dim comstr As String = "SELECT DISTINCT JCNO FROM JC_WEEK ORDER BY JCNO"
        Dim da As New SqlDataAdapter(comstr, sqlcon)
        Dim ds As New DataSet
        Try
            GetDBCon()
            sqlcon.Open()
            da.Fill(ds, "JC_WEEK")
            cbPeriod.DataSource = ds.Tables(0)
            cbPeriod.ValueMember = "JCNO"
            cbPeriod.DisplayMember = "JCNO"
            sqlcon.Close()
        Catch ex As Exception
            MessageBox.Show("Error on getting period data, with message: " & ex.Message.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            sqlcon.Close()
        End Try
    End Sub

    Public Sub frmFCS_Rep_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Year()
        period()
        loadclbDSR()
    End Sub

    Private Sub clbDSR_ItemCheck(sender As Object, e As ItemCheckEventArgs) Handles clbDSR.ItemCheck
        If e.Index = 0 Then
            Dim newcheckstate As CheckState = e.NewValue
            For i As Integer = 1 To clbDSR.Items.Count - 1
                Me.clbDSR.SetItemCheckState(i, newcheckstate)
            Next i
        End If
        SelectedItem = ""
    End Sub

    Public SelectedItem As String
    Public DSRlist As String

    Public Sub GetSelectedCheckBox()
        Dim ListDSR As New ArrayList
        Dim ListHeadDSR As New ArrayList
        Try
            Dim checked As Integer = clbDSR.CheckedItems.Count
            For i As Integer = 0 To checked - 1
                If clbDSR.CheckedItems(i).ToString = "Select All" Then
                    '-----Do Nothing-----

                Else
                    Dim result As String = clbDSR.CheckedItems(i).ToString
                    ListHeadDSR.Add(result)
                    Dim split = result.Split(" - ")
                    ListDSR.Add(split(0))
                End If
            Next
            SelectedItem = String.Join("; ", TryCast(ListHeadDSR.ToArray(GetType(String)), String()))
            DSRlist = String.Join("','", TryCast(ListDSR.ToArray(GetType(String)), String()))
        Catch ex As Exception
            MessageBox.Show("Error on getting DSR, With message: " & ex.Message.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        
    End Sub
    Public Sub GetDataHead()
        Dim comhead As String
        comhead = "SELECT '" & cbYear.SelectedValue.ToString & "' AS YEAR,'" & cbPeriod.SelectedValue.ToString & "' AS PERIOD, " _
            & "DTNAME, '" & SelectedItem & "' AS DSR_LIST FROM DISTRIBUTOR"
        Dim cmdhead As New SqlCommand(comhead, sqlcon)
        Dim daHeadRep As New SqlDataAdapter(cmdhead)
        Dim dsHeadRep As New DataSet
        Try
            GetDBCon()
            sqlcon.Open()
            daHeadRep.Fill(dsHeadRep)
            dsHeadRep.Tables(0).TableName = "dtHeadRep"
            frmFCS_RepViewer.dsHeadRepBindingSource.DataSource = dsHeadRep
            sqlcon.Close()
        Catch ex As Exception
            MessageBox.Show("Error on passing data to report head, with message: " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            sqlcon.Close()
        End Try
    End Sub

    'Public Sub GetECO()
    '    Dim comstr As String
    '    comstr = "SELECT H.DSR,H.NAME,H.UniqSCH AS UNIQSCH," _
    '        & "CASE WHEN G.UniqProd IS NULL THEN 0 ELSE G.UniqProd END AS UNIQPROD, H.ECO_TARGET FROM (select DISTINCT A.DSR,A.NAME,COUNT(A.POP) as UniqSCH, C.ECO_TARGET from (" _
    '        & "select DISTINCT A.DSR,B.NAME,A.POP from SECTION_PRMNT_BYDATE A inner join DSR B ON A.DSR=B.DSR " _
    '        & "WHERE A.JCNO='" & cbPeriod.SelectedValue.ToString & "' AND A.DSR IN ('" & DSRlist & "') GROUP BY A.DSR,B.NAME,A.POP) A Inner Join TARGET_BYDSR C ON A.DSR = C.DSR GROUP BY A.DSR, A.NAME, C.ECO_TARGET) H " _
    '        & "FULL OUTER JOIN (Select DSR, NAME, COUNT(POP) AS UniqProd FROM (select DISTINCT(A.POP) as POP, A.DSR,D.NAME " _
    '        & "from CASHMEMO A left join SECTION_PRMNT_BYDATE B ON a.PJP = b.PJP and a.SECTION=b.SECTION and a.POP = b.POP and a.DOC_DATE=b.DATE " _
    '        & "left join JC_WEEK C on b.JCNO = c.JCNO left join DSR D on a.DSR=D.DSR " _
    '        & "where A.DOC_DATE between (select MIN(DATE) from SECTION_PRMNT_BYDATE where JCNO = '" & cbPeriod.SelectedValue.ToString & "') " _
    '        & "and (select MAX(DATE) from SECTION_PRMNT_BYDATE where JCNO = '" & cbPeriod.SelectedValue.ToString & "')) X WHERE X.DSR IN ('" & DSRlist & "') " _
    '        & "GROUP BY X.DSR,X.NAME) G ON H.DSR=G.DSR " _
    '        & "GROUP BY H.DSR,H.NAME,H.UniqSCH,G.UniqProd,H.ECO_TARGET ORDER BY H.DSR"
    '    Dim cmdstr As New SqlCommand(comstr, sqlcon)
    '    Dim daECO As New SqlDataAdapter(cmdstr)
    '    Dim dsECO As New DataSet
    '    Try
    '        GetDBCon()
    '        sqlcon.Open()
    '        daECO.Fill(dsECO)
    '        dsECO.Tables(0).TableName = "dtECO"
    '        frmFCS_RepViewer.ECOBindingSource.DataSource = dsECO
    '        sqlcon.Close()
    '    Catch ex As Exception
    '        MessageBox.Show("Error on passing data to ECO, with message: " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '        sqlcon.Close()
    '    End Try
    'End Sub
    ''----------------Query of Ivan Phase-----------------
    'Public Sub GetBPSchPRod()
    '    Dim comstr As String
    '    comstr = "select g.DSR,H.SCH,h.PROD,g.BP_TARGET FROM (select a.DSR," _
    '        & "case when b.ECO_TARGET is null then 0 else b.ECO_TARGET end as ECO_TARGET," _
    '        & "case when b.BP_TARGET is null then 0 else b.BP_TARGET end as BP_TARGET," _
    '        & "case when b.LPPC_TARGET is null then 0 else b.LPPC_TARGET end as LPPC_TARGET from " _
    '        & "(select DSR from DSR) a FULL OUTER JOIN (select DSR,ECO_TARGET,BP_TARGET,LPPC_TARGET from TARGET_BYDSR " _
    '        & "where JC_TARGET='" & cbPeriod.SelectedValue.ToString & "' and TARGET_YEAR='" & cbYear.SelectedValue.ToString _
    '        & "') b ON a.DSR=b.DSR GROUP BY a.DSR,b.ECO_TARGET,b.BP_TARGET,b.LPPC_TARGET) G FULL OUTER JOIN " _
    '        & "(select x.DSR,x.sch as SCH,case when y.prod IS null then 0 else y.prod end as PROD from " _
    '        & "(select bp.dsr, count(bp.pop) sch from (select a.DATE , a.DSR , a.POP from SECTION_PRMNT_BYDATE a " _
    '        & "inner join (select distinct doc_date from CASHMEMO ) c on a.DATE = c.DOC_DATE " _
    '        & "inner join Calendar d on c.DOC_DATE = d.DATE where a.JCNO ='" & cbPeriod.SelectedValue.ToString & "' and d.DAY_NAME = a.DAY_VISIT) bp " _
    '        & "group by bp.DSR) X FULL OUTER JOIN (select cm.dsr , (cm.cms-sr.sr) as prod from " _
    '        & "(select b.DSR, COUNT(pop)CMs from (select distinct (POP), DOC_DATE,DSR from CASHMEMO a, JC_WEEK b " _
    '        & " where a.DOC_DATE >= b.START_DATE and a.DOC_DATE <= b.END_DATE and b.YEAR ='" & cbYear.SelectedValue.ToString _
    '        & "' and b.JCNO ='" & cbPeriod.SelectedValue.ToString & "') b group by b.DSR) cm inner join" _
    '        & "(select DSR, COUNT(DOC_NO) sr from SALES_RETURN a, JC_WEEK b " _
    '        & "where a.DOC_DATE >= b.START_DATE and a.DOC_DATE <= b.END_DATE and b.JCNO ='" & cbPeriod.SelectedValue.ToString _
    '        & "' and b.YEAR ='" & cbYear.SelectedValue.ToString & "' group by DSR) sr " _
    '        & "on cm.DSR=sr.DSR) Y ON x.DSR=y.dsr) H ON G.DSR=H.DSR where G.DSR IN ('" & DSRlist & "') " _
    '        & "GROUP BY G.DSR,H.SCH,H.prod,G.BP_TARGET ORDER BY G.DSR"
    '    Dim cmdstr As New SqlCommand(comstr, sqlcon)
    '    Dim daBP As New SqlDataAdapter(cmdstr)
    '    Dim dsBP As New DataSet
    '    Try
    '        GetDBCon()
    '        sqlcon.Open()
    '        daBP.Fill(dsBP)
    '        dsBP.Tables(0).TableName = "dtBP"
    '        frmFCS_RepViewer.BPBindingSource.DataSource = dsBP
    '        sqlcon.Close()
    '    Catch ex As Exception
    '        MessageBox.Show("Error on passing data to BP, with message: " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '        sqlcon.Close()
    '    End Try

    'End Sub

    'Public Sub GetLPPC()
    '    Dim comstr As String
    '    comstr = "SELECT G.DSR,case when H.bp is null then 0 else h.bp end as CM," _
    '        & "case when H.sku is null then 0 else h.sku end AS SKU,G.LPPC_TARGET FROM " _
    '        & "(select a.DSR,case when b.ECO_TARGET is null then 0 else b.ECO_TARGET end as ECO_TARGET," _
    '        & "case when b.BP_TARGET is null then 0 else b.BP_TARGET end as BP_TARGET," _
    '        & "case when b.LPPC_TARGET is null then 0 else b.LPPC_TARGET end as LPPC_TARGET from (select DSR from DSR) a " _
    '        & "FULL OUTER JOIN (select DSR,ECO_TARGET,BP_TARGET,LPPC_TARGET from TARGET_BYDSR " _
    '        & "where JC_TARGET='" & cbPeriod.SelectedValue.ToString & "' and TARGET_YEAR='" & cbYear.SelectedValue.ToString & "') b ON a.DSR=b.DSR " _
    '        & "GROUP BY a.DSR,b.ECO_TARGET,b.BP_TARGET,b.LPPC_TARGET) G FULL OUTER JOIN (SELECT X.DSR,X.bp,y.sku from " _
    '        & "(select cm.dsr , (cm.cms-sr.sr) bp from (select b.DSR, COUNT(pop)CMs from (select distinct (POP), DOC_DATE,DSR from CASHMEMO a, JC_WEEK b " _
    '        & "where a.DOC_DATE >= b.START_DATE and a.DOC_DATE <= b.END_DATE and b.YEAR = '" & cbYear.SelectedValue.ToString & "' " _
    '        & "and b.JCNO = '07') b group by b.DSR) cm left join (select DSR, COUNT(DOC_NO) sr from SALES_RETURN a, JC_WEEK b " _
    '        & "where a.DOC_DATE >= b.START_DATE and a.DOC_DATE <= b.END_DATE and b.JCNO = '" & cbPeriod.SelectedValue.ToString & "' and b.YEAR = '" & cbYear.SelectedValue.ToString & "' " _
    '        & "group by DSR) sr on cm.DSR=sr.DSR) X Full OUTER JOIN (select cm.dsr , (cm.skus-sr.skues) sku from " _
    '        & "(select c.DSR, COUNT(c.sku) skus from (select distinct ab.sku, ab.POP,ab.DOC_DATE,ab.dsr from ( " _
    '        & "select a.SKU,a.DOC_DATE,b.DSR,a.DOC_NO ,b.POP from cm_detail a left join cashmemo b on  a.DOC_NO = b.DOC_NO, JC_WEEK c " _
    '        & "where b.DOC_DATE >= c.START_DATE and b.DOC_DATE <= c.END_DATE and c.JCNO = '" & cbPeriod.SelectedValue.ToString & "' and c.YEAR='" & cbYear.SelectedValue.ToString & "' " _
    '        & ") ab) c left join DSR d on c.DSR=d.DSR group by c.DSR ) cm left join (select d.dsr, COUNT(d.sku) skues from " _
    '        & "(select a.SKU,b.DSR,b.POP from SR_DETAIL a left join SALES_RETURN b on a.DOC_NO=b.DOC_NO, JC_WEEK c " _
    '        & "where b.DOC_DATE >= c.START_DATE and b.DOC_DATE <= c.END_DATE and c.JCNO = '" & cbPeriod.SelectedValue.ToString & "' and c.YEAR='" & cbYear.SelectedValue.ToString & "') d " _
    '        & "group by d.DSR) sr on cm.DSR=sr.DSR) Y ON x.DSR = y.DSR) H ON G.DSR=H.DSR GROUP BY G.DSR,H.bp,H.sku,G.LPPC_TARGET ORDER By G.DSR"
    '    Dim cmdstr As New SqlCommand(comstr, sqlcon)
    '    Dim daLPPC As New SqlDataAdapter(cmdstr)
    '    Dim dsLPPC As New DataSet
    '    Try
    '        GetDBCon()
    '        sqlcon.Open()
    '        daLPPC.Fill(dsLPPC)
    '        dsLPPC.Tables(0).TableName = "dtLPPC"
    '        frmFCS_RepViewer.LPPCBindingSource.DataSource = dsLPPC
    '        sqlcon.Close()
    '    Catch ex As Exception
    '        MessageBox.Show("Error on passing data to LPPC, with message: " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '        sqlcon.Close()
    '    End Try

    'End Sub

    Public Sub GetFCSData()
        Dim comstr As String
        comstr = "select A.DSR,A.NAME,A.UniqSCH AS UNIQSCH,A.UNIQPROD,A.ECO_TARGET,B.SCH,B.PROD,B.BP_TARGET,C.CM,C.SKU,C.LPPC_TARGET from " _
            & "(SELECT H.DSR,H.NAME,H.UniqSCH,CASE WHEN G.UniqProd IS NULL THEN 0 ELSE G.UniqProd END AS UNIQPROD, H.ECO_TARGET " _
            & "FROM (select DISTINCT A.DSR,A.NAME,COUNT(A.POP) as UniqSCH, C.ECO_TARGET from (" _
            & "select DISTINCT A.DSR,B.NAME,A.POP from SECTION_PRMNT_BYDATE A inner join DSR B ON A.DSR=b.DSR " _
            & "WHERE A.JCNO='" & cbPeriod.SelectedValue.ToString & "' AND A.DSR IN ('" & DSRlist & "') GROUP BY A.DSR,B.NAME,A.POP) A Inner Join TARGET_BYDSR C ON A.DSR = C.DSR " _
            & "WHERE C.JC_TARGET='" & cbPeriod.SelectedValue.ToString & "' GROUP BY A.DSR, A.NAME, C.ECO_TARGET) H FULL OUTER JOIN (Select DSR, NAME, COUNT(POP) AS UniqProd FROM (" _
            & "select DISTINCT(A.POP) as POP, A.DSR,D.NAME from CASHMEMO A " _
            & "left join SECTION_PRMNT_BYDATE B ON a.PJP = b.PJP and a.SECTION=b.SECTION and a.POP = b.POP and a.DOC_DATE=b.DATE " _
            & "left join JC_WEEK C on b.JCNO = c.JCNO left join DSR D on a.DSR=D.DSR " _
            & "where A.DOC_DATE between (select MIN(DATE) from SECTION_PRMNT_BYDATE where JCNO = '" & cbPeriod.SelectedValue.ToString & "') and " _
            & "(select MAX(DATE) from SECTION_PRMNT_BYDATE where JCNO = '" & cbPeriod.SelectedValue.ToString & "')) X GROUP BY X.DSR,X.NAME) G ON H.DSR=G.DSR " _
            & "GROUP BY H.DSR,H.NAME,H.UniqSCH,G.UniqProd, H.ECO_TARGET) A " _
            & "JOIN " _
            & "(select g.DSR,H.SCH,h.PROD,g.BP_TARGET FROM (select a.DSR,case when b.ECO_TARGET is null then 0 else b.ECO_TARGET end as ECO_TARGET," _
            & "case when b.BP_TARGET is null then 0 else b.BP_TARGET end as BP_TARGET," _
            & "case when b.LPPC_TARGET is null then 0 else b.LPPC_TARGET end as LPPC_TARGET from (select DSR from DSR) a " _
            & "FULL OUTER JOIN (select DSR,ECO_TARGET,BP_TARGET,LPPC_TARGET from TARGET_BYDSR " _
            & "where JC_TARGET='" & cbPeriod.SelectedValue.ToString & "' and TARGET_YEAR='" & cbYear.SelectedValue.ToString & "') b ON a.DSR=b.DSR " _
            & "GROUP BY a.DSR,b.ECO_TARGET,b.BP_TARGET,b.LPPC_TARGET) G FULL OUTER JOIN (" _
            & "select x.DSR,x.sch as SCH,case when y.prod IS null then 0 else y.prod end as PROD from " _
            & "(select PRMNT.dsr ,case when unpln.POP is null then PRMNT.sch else (PRMNT.sch+unpln.POP) end AS SCH from " _
            & "(select bp.dsr, count(bp.pop) sch from (select a.DATE , a.DSR , a.POP from SECTION_PRMNT_BYDATE a " _
            & "inner join (select distinct doc_date from CASHMEMO ) c on a.DATE = c.DOC_DATE " _
            & "inner join Calendar d on c.DOC_DATE = d.DATE where a.JCNO = '" & cbPeriod.SelectedValue.ToString & "'and d.DAY_NAME = a.DAY_VISIT) bp " _
            & "group by bp.DSR) PRMNT left join (SELECT Q.DSR_CM, COUNT(Q.POP_CM) AS POP FROM " _
            & "(SELECT DOC_DATE, DSR_CM, POP_CM FROM (select A.DOC_DATE,A.DSR AS DSR_CM,A.POP AS POP_CM,B.DSR as DSR_PERMANENT,B.POP POP_PERMANENT FROM " _
            & "(select DOC_DATE,DSR,POP FROM CASHMEMO , JC_WEEK where CASHMEMO.DOC_DATE>=JC_WEEK.START_DATE " _
            & "and CASHMEMO.DOC_DATE <= JC_WEEK.END_DATE and JC_WEEK.JCNO='" & cbPeriod.SelectedValue.ToString & "' and JC_WEEK.YEAR = '" & cbYear.SelectedValue.ToString & "') A " _
            & "FULL JOIN (select DATE,DSR,POP FROM SECTION_PRMNT_BYDATE where SECTION_PRMNT_BYDATE.JCNO ='" & cbPeriod.SelectedValue.ToString & "') B " _
            & "ON A.DOC_DATE=B.DATE and A.POP=B.POP) W WHERE POP_PERMANENT IS NULL) Q GROUP BY Q.DSR_CM) unpln on PRMNT.dsr = unpln.DSR_CM" _
            & ") X FULL OUTER JOIN (select cm.dsr , case when sr.sr is null then cm.CMs else (cm.cms-sr.sr)end as prod from (select b.DSR, COUNT(pop)CMs from " _
            & "(select distinct (POP), DOC_DATE,DSR from CASHMEMO a, JC_WEEK b " _
            & "where a.DOC_DATE >= b.START_DATE and a.DOC_DATE <= b.END_DATE and b.YEAR ='" & cbYear.SelectedValue.ToString & "' and b.JCNO ='" & cbPeriod.SelectedValue.ToString & "') b group by b.DSR) cm " _
            & "left join (select DSR, COUNT(DOC_NO) sr from SALES_RETURN a, JC_WEEK b " _
            & "where a.DOC_DATE >= b.START_DATE and a.DOC_DATE <= b.END_DATE and b.JCNO ='" & cbPeriod.SelectedValue.ToString & "' and b.YEAR ='" & cbYear.SelectedValue.ToString & "' " _
            & "group by DSR) sr on cm.DSR=sr.DSR) Y ON x.DSR=y.dsr) H ON G.DSR=H.DSR GROUP BY G.DSR,H.SCH,H.prod,G.BP_TARGET) B ON A.DSR=B.DSR " _
            & "JOIN " _
            & "(SELECT G.DSR,case when H.bp is null then 0 else h.bp end as CM,case when H.sku is null then 0 else h.sku end AS SKU,G.LPPC_TARGET FROM " _
            & "(select a.DSR,case when b.ECO_TARGET is null then 0 else b.ECO_TARGET end as ECO_TARGET," _
            & "case when b.BP_TARGET is null then 0 else b.BP_TARGET end as BP_TARGET," _
            & "case when b.LPPC_TARGET is null then 0 else b.LPPC_TARGET end as LPPC_TARGET from (select DSR from DSR) a " _
            & "FULL OUTER JOIN (select DSR,ECO_TARGET,BP_TARGET,LPPC_TARGET from TARGET_BYDSR " _
            & "where JC_TARGET='" & cbPeriod.SelectedValue.ToString & "' and TARGET_YEAR='" & cbYear.SelectedValue.ToString & "') b ON a.DSR=b.DSR " _
            & "GROUP BY a.DSR,b.ECO_TARGET,b.BP_TARGET,b.LPPC_TARGET) G FULL OUTER JOIN (SELECT X.DSR,X.bp,y.sku from " _
            & "(select cm.dsr , case when sr.sr is null then cm.CMs else (cm.cms-sr.sr)end as bp from (select b.DSR, COUNT(pop)CMs from " _
            & "(select distinct (POP), DOC_DATE,DSR from CASHMEMO a, JC_WEEK b where a.DOC_DATE >= b.START_DATE and a.DOC_DATE <= b.END_DATE and b.YEAR = '" & cbYear.SelectedValue.ToString & "' " _
            & "and b.JCNO = '" & cbPeriod.SelectedValue.ToString & "') b group by b.DSR) cm left join " _
            & "(select DSR, COUNT(DOC_NO) sr from SALES_RETURN a, JC_WEEK b " _
            & "where a.DOC_DATE >= b.START_DATE and a.DOC_DATE <= b.END_DATE and b.JCNO = '" & cbPeriod.SelectedValue.ToString & "' and b.YEAR = '" & cbYear.SelectedValue.ToString & "' " _
            & "group by DSR) sr on cm.DSR=sr.DSR) X Full OUTER JOIN (select cm.dsr ," _
            & "case when sr.skues is null then cm.skus else (cm.skus - sr.skues) end as sku from " _
            & "(select c.DSR, COUNT(c.sku) skus from (select distinct ab.sku, ab.POP,ab.DOC_DATE,ab.dsr from ( " _
            & "select a.SKU,a.DOC_DATE,b.DSR,a.DOC_NO ,b.POP from cm_detail a left join cashmemo b on  a.DOC_NO = b.DOC_NO, JC_WEEK c " _
            & "where b.DOC_DATE >= c.START_DATE and b.DOC_DATE <= c.END_DATE and c.JCNO = '" & cbPeriod.SelectedValue.ToString & "' and c.YEAR='" & cbYear.SelectedValue.ToString & "' " _
            & ") ab) c left join DSR d on c.DSR=d.DSR group by c.DSR ) cm left join " _
            & "(select d.dsr, COUNT(d.sku) skues from (select a.SKU,b.DSR,b.POP from SR_DETAIL a left join SALES_RETURN b on a.DOC_NO=b.DOC_NO, JC_WEEK c " _
            & "where b.DOC_DATE >= c.START_DATE and b.DOC_DATE <= c.END_DATE and c.JCNO = '" & cbPeriod.SelectedValue.ToString & "' and c.YEAR='" & cbYear.SelectedValue.ToString & "') d " _
            & "group by d.DSR) sr on cm.DSR=sr.DSR) Y ON x.DSR = y.DSR) H ON G.DSR=H.DSR GROUP BY G.DSR,H.bp,H.sku,G.LPPC_TARGET) C ON c.DSR=b.DSR " _
            & "WHERE A.DSR IN ('" & DSRlist & "') order By A.DSR"

        Dim cmdstr As New SqlCommand(comstr, sqlcon)
        Dim daFCS As New SqlDataAdapter(cmdstr)
        Dim dsFCS As New DataSet
        Try
            GetDBCon()
            sqlcon.Open()
            daFCS.Fill(dsFCS)
            daFCS.SelectCommand.CommandTimeout = 600
            dsFCS.Tables(0).TableName = "dtFCS"
            frmFCS_RepViewer.dsFCSRepBindingSource.DataSource = dsFCS
            sqlcon.Close()
        Catch ex As Exception
            MessageBox.Show("Error on passing data to FCS, with message: " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            sqlcon.Close()
        End Try

    End Sub

    Private Sub btnViewRep_Click(sender As Object, e As EventArgs) Handles btnViewRep.Click
        SSLoading.Show()
        frmFCS_RepViewer.Show()
    End Sub
End Class