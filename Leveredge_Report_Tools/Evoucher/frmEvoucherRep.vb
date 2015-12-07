Imports System.Data
Imports System.Data.SqlClient
Imports System.IO.File
Public Class frmEvoucherRep
    Dim sqlcon As New SqlConnection
    Dim constr As String

    Private Sub getDbCon()
        Dim strfile As String = My.Application.Info.DirectoryPath & "\Conf.ini"
        constr = ReadAllLines(strfile)(1)
        sqlcon.ConnectionString = constr
    End Sub

    Private Sub GetEV_TreeView()
        Dim cmd_EV_CAT As String = "SELECT EVOUCHER_MASTER, LDESC FROM EV_EVOUCHER_MASTER"
        cmd_EV_CAT = "DECLARE @DATE DATE " _
            & "SET @DATE = '" & DTPEV.Value.ToString("yyyy-MM-dd") & "' " _
            & "SELECT B.EVOUCHER_MASTER,A.LDESC FROM EV_EVOUCHER_MASTER A LEFT JOIN EV_EVOUCHER_CATEGORY B ON A.EVOUCHER_MASTER=B.EVOUCHER_MASTER " _
            & "WHERE @DATE BETWEEN (B.FROM_DATE) AND (B.END_DATE)"

        Dim cmd_EV_MAS As String
        cmd_EV_MAS = "DECLARE @DATE DATE " _
            & "SET @DATE='" & DTPEV.Value.ToString("yyyy-MM-dd") & "' " _
            & "SELECT EVOUCHER_MASTER,EVOUCHER_CATEGORY,LDESC FROM EV_EVOUCHER_CATEGORY " _
            & "WHERE @DATE BETWEEN (FROM_DATE) AND (END_DATE)"
        Try
            Dim ds As DataSet
            ds = New DataSet()
            getDbCon()
            sqlcon.Open()
            Dim da_EV_CAT As New SqlDataAdapter(cmd_EV_CAT, sqlcon)
            Dim da_EV_MAS As New SqlDataAdapter(cmd_EV_MAS, sqlcon)
            da_EV_CAT.Fill(ds, "EV_CATEGORY")
            da_EV_MAS.Fill(ds, "EV_MASTER")
            sqlcon.Close()

            'Create data relation
            ds.Relations.Add("CatToMas", ds.Tables("EV_CATEGORY").Columns("EVOUCHER_MASTER"), ds.Tables("EV_MASTER").Columns("EVOUCHER_MASTER"))
            TVevoucherMaster.Nodes.Clear()
            Dim parentrow As DataRow
            Dim parentTable As DataTable
            parentTable = ds.Tables("EV_CATEGORY")

            For Each parentrow In parentTable.Rows
                Dim parentnode As TreeNode
                parentnode = New TreeNode(parentrow(0) & " - " & parentrow(1))
                TVevoucherMaster.Nodes.Add(parentnode)

                Dim childrow As DataRow
                Dim childnode As New TreeNode()
                For Each childrow In parentrow.GetChildRows("CatToMas")
                    childnode = parentnode.Nodes.Add(childrow(1) & " - " & childrow(2))
                Next childrow
            Next parentrow
        Catch ex As Exception
            MessageBox.Show("Error on populating treeview, with message: " & ex.Message.ToString, "Error TreeView", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub DTPEV_ValueChanged(sender As Object, e As EventArgs) Handles DTPEV.ValueChanged
        GetEV_TreeView()
    End Sub

    Public Sub Get_EV_HEAD()
        Dim str_EV_HEAD As String
        str_EV_HEAD = "select DISTRIBUTOR as DTCODE, NAME as DTNAME from DISTRIBUTOR"
        Dim cmd_EV_HEAD As New SqlCommand(str_EV_HEAD, sqlcon)
        Dim da_EV_HEAD As New SqlDataAdapter(cmd_EV_HEAD)
        Dim ds_EV_HEAD As New DataSet
        Try
            getDbCon()
            sqlcon.Open()
            da_EV_HEAD.Fill(ds_EV_HEAD)
            da_EV_HEAD.SelectCommand.CommandTimeout = 600
            ds_EV_HEAD.Tables(0).TableName = "dtHeader_EV"
            If cbRepType.SelectedIndex = 0 Then
                frmEV_RVClaim.dtHeader_EVBindingSource.DataSource = ds_EV_HEAD
            ElseIf cbRepType.SelectedIndex = 1 Then
                frmEV_RVParticipant.dtHeader_EVBindingSource.DataSource = ds_EV_HEAD
            End If
            sqlcon.Close()
        Catch ex As Exception
            MessageBox.Show("Error on passing data to report head, with message: " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            sqlcon.Close()
        End Try
    End Sub

    Public Comp, Dist, EV_CAT As String

    Private Sub GetComp_Dist_EVCAT()
        Dim strcomp, strdist As String
        strcomp = "SELECT COMPANY FROM DISTRIBUTOR"
        strdist = "SELECT DISTRIBUTOR FROM DISTRIBUTOR"
        Dim cmdcomp As New SqlCommand(strcomp, sqlcon)
        Dim cmddist As New SqlCommand(strdist, sqlcon)
        Try
            getDbCon()
            sqlcon.Open()
            Comp = cmdcomp.ExecuteScalar
            Dist = cmddist.ExecuteScalar
            sqlcon.Close()
            Dim selnode_child, selnode_parent As String
            selnode_child = TVevoucherMaster.SelectedNode.Text
            selnode_parent = TVevoucherMaster.SelectedNode.Parent.Text
            Dim Cat_result = selnode_parent.Split(" - ")
            Dim Mas_result = selnode_child.Split(" - ")
            EV_CAT = Cat_result(0) & Mas_result(0)
        Catch ex As Exception
            MessageBox.Show("Error on getting parameter Company, Distributor, Evoucher, With message: " & ex.Message.ToString, "Error", MessageBoxButtons.OK)
            sqlcon.Close()
        End Try
    End Sub

    Public Sub Get_EV_CLAIM()
        Cursor = Cursors.AppStarting
        GetComp_Dist_EVCAT()
        Dim cmd_EV_CLAIM As String
        cmd_EV_CLAIM = "DECLARE @START_DATE13 DATE, @END_DATE13 DATE, @START_DATE DATE, @COMPANY NVARCHAR(2), " _
            & "@DISTRIBUTOR NVARCHAR(8),@EVOUCHER_CATEGORY NVARCHAR(10) " _
            & "SET @EVOUCHER_CATEGORY = '" & EV_CAT & "' " _
            & "SET @START_DATE='" & DTPEV.Value.ToString("yyyy-MM-dd") & "' " _
            & "SET @COMPANY='" & Comp & "' " _
            & "SET @DISTRIBUTOR='" & Dist & "' " _
            & "SELECT MIN(START_DATE) AS START_DATE, MAX(END_DATE) END_DATE INTO #LAST13WEEK FROM " _
            & "(SELECT *, ROW_NUMBER()OVER (PARTITION BY COMPANY ORDER BY YEAR DESC, JCNO DESC, WEEK_NO DESC) RN " _
            & "FROM JC_WEEK WHERE COMPANY+YEAR+JCNO+WEEK_NO < (SELECT COMPANY+YEAR+JCNO+WEEK_NO FROM JC_WEEK  WHERE @START_DATE >= START_DATE AND @START_DATE <= END_DATE)) A WHERE RN <= 13 " _
            & "SELECT @START_DATE13 = (SELECT START_DATE FROM #LAST13WEEK), @END_DATE13 = (SELECT END_DATE FROM #LAST13WEEK) " _
            & "SELECT EP.EVOUCHER_MASTER as EVOUCHER_CATEGORY,M.LDESC AS EVOUCHER_CATEGORY_DESC, EP.EVOUCHER_CATEGORY AS EVOUCHER_MASTER, EC.LDESC AS EVOUCHER_MASTER_DESC, " _
            & "EC.FROM_DATE, EC.END_DATE, EP.TOWN+EP.LOCALITY+EP.SLOCALITY+EP.POP AS POP_CODE, " _
            & "P.NAME AS POP_NAME, SUM(C.LAST_13_WEEK_SALE) LAST_13_WEEK_SALE, EP.REF_DOC_NO, " _
            & "CASE WHEN EP.STATUS = 'A' AND EP.QUALIFY_STATUS = 'Q' THEN 'Qualified' WHEN EP.STATUS = 'A' " _
            & "AND EP.QUALIFY_STATUS = 'U' THEN 'Un-Qualified' WHEN EP.STATUS = 'A' THEN 'Associated' WHEN EP.STATUS = 'P' " _
            & "AND EP.QUALIFY_STATUS = 'Q' THEN 'Processed'WHEN EP.STATUS = 'U' AND EP.QUALIFY_STATUS = 'Q' THEN 'UnProcessed' " _
            & "WHEN EP.STATUS = 'F' AND EP.QUALIFY_STATUS = 'Q' THEN 'Finalized' END  AS STATUS_ASSESMENT, " _
            & "(SELECT MIN(CA.DOC_NO) FROM CASHMEMO CA WHERE EP.COMPANY = CA.COMPANY AND EP.DISTRIBUTOR = CA.DISTRIBUTOR " _
            & "AND EP.TOWN = CA.TOWN AND EP.LOCALITY = CA.LOCALITY AND EP.SLOCALITY = CA.SLOCALITY " _
            & "AND EP.POP = CA.POP AND CA.DOC_DATE >= EP.QUALIFY_DATE AND EP.QUALIFY_STATUS='Q' AND CA.VISIT_TYPE='02' " _
            & "AND CA.DOCUMENT='CM' AND CA.SUB_DOCUMENT='01') AS INVOICE_NO_INFO, R.REF_DOC_NO AS INVOICE_NUMBER_EVOUCHER, " _
            & "SUM(CASE WHEN R.REF_DOC_NO IS NOT NULL THEN EC.VOUCHER_VALUE ELSE 0 END) AS VOUCHER_VALUE1, " _
            & "CASE WHEN ROW_NUMBER() OVER (PARTITION BY EP.EVOUCHER_MASTER, EP.EVOUCHER_CATEGORY ORDER BY EP.TOWN+EP.LOCALITY+EP.SLOCALITY+EP.POP)=1 " _
            & "THEN SUM(EC.VOUCHER_VALUE) ELSE 0 END VOUCHER_VALUE2, EA.POP_ALLOCATION AS BUDGET_ALLOCATED " _
            & "FROM EV_EVOUCHER_POP EP " _
            & "LEFT JOIN EV_EVOUCHER_ALLOCATION EA ON EP.EVOUCHER_CATEGORY = EA.EVOUCHER_CATEGORY " _
            & "LEFT JOIN RECEIPT_ADJ R ON EP.COMPANY = R.COMPANY AND EP.DISTRIBUTOR = R.DISTRIBUTOR AND EP.REF_DOCUMENT = R.DOCUMENT AND EP.REF_SUB_DOCUMENT = R.SUB_DOCUMENT AND EP.REF_DOC_NO = R.DOC_NO " _
            & "LEFT JOIN CASHMEMO CM ON R.COMPANY = CM.COMPANY AND R.DISTRIBUTOR = CM.DISTRIBUTOR AND R.REF_DOCUMENT = CM.DOCUMENT AND R.REF_SUB_DOCUMENT = CM.SUB_DOCUMENT AND R.REF_DOC_NO = CM.DOC_NO " _
            & "INNER JOIN EV_EVOUCHER_MASTER M ON EP.COMPANY = M.COMPANY AND EP.EVOUCHER_MASTER = M.EVOUCHER_MASTER " _
            & "INNER JOIN EV_EVOUCHER_CATEGORY EC ON EP.COMPANY = EC.COMPANY AND EP.EVOUCHER_MASTER = EC.EVOUCHER_MASTER " _
            & "AND EA.VALUE_COMB LIKE ('%~'+@DISTRIBUTOR) " _
            & "AND EP.EVOUCHER_CATEGORY = EC.EVOUCHER_CATEGORY INNER JOIN POP P ON EP.COMPANY = P.COMPANY AND EP.DISTRIBUTOR = P.DISTRIBUTOR " _
            & "AND EP.TOWN = P.TOWN AND EP.LOCALITY = P.LOCALITY AND EP.SLOCALITY = P.SLOCALITY AND EP.POP = P.POP LEFT JOIN (" _
            & "SELECT CM.COMPANY, CM.DISTRIBUTOR, CM.TOWN, CM.LOCALITY, CM.SLOCALITY, CM.POP, " _
            & "SUM(COALESCE(CM.NET_AMOUNT,0)) / 13 LAST_13_WEEK_SALE FROM CASHMEMO CM WHERE CM.COMPANY = @COMPANY AND CM.DOC_DATE >= @START_DATE13 AND CM.DOC_DATE <= @END_DATE13 " _
            & "AND	CM.DISTRIBUTOR IN (@DISTRIBUTOR) AND CM.DOCUMENT IN ('CM') AND CM.VISIT_TYPE ='02' AND CM.SUB_DOCUMENT IN ('01','02','04') " _
            & "GROUP BY CM.COMPANY, CM.DISTRIBUTOR, CM.TOWN, CM.LOCALITY, CM.SLOCALITY, CM.POP) C ON EP.COMPANY = C.COMPANY " _
            & "AND EP.DISTRIBUTOR = C.DISTRIBUTOR AND EP.TOWN = C.TOWN AND EP.LOCALITY = C.LOCALITY AND EP.SLOCALITY = C.SLOCALITY AND EP.POP = C.POP WHERE EP.COMPANY = @COMPANY " _
            & "AND EP.DISTRIBUTOR IN (@DISTRIBUTOR)AND EP.EVOUCHER_MASTER+EP.EVOUCHER_CATEGORY IN (@EVOUCHER_CATEGORY) " _
            & "GROUP BY EP.EVOUCHER_MASTER ,M.LDESC, EP.EVOUCHER_CATEGORY, EC.LDESC, EC.FROM_DATE, EC.END_DATE, EP.COMPANY, EP.DISTRIBUTOR,EP.TOWN,EP.LOCALITY, " _
            & "EP.SLOCALITY,EP.POP, P.NAME, EP.EVOUCHER_MASTER, EP.EVOUCHER_CATEGORY,EP.REF_DOC_NO, R.REF_DOC_NO, " _
            & "EP.QUALIFY_DATE, EP.QUALIFY_STATUS,EA.POP_ALLOCATION, CASE WHEN EP.STATUS = 'A' " _
            & "AND EP.QUALIFY_STATUS = 'Q'  THEN 'Qualified' WHEN EP.STATUS = 'A' " _
            & "AND EP.QUALIFY_STATUS = 'U' THEN 'Un-Qualified' WHEN EP.STATUS = 'A' THEN 'Associated' " _
            & "WHEN EP.STATUS = 'P' AND EP.QUALIFY_STATUS = 'Q' THEN 'Processed'WHEN EP.STATUS = 'U' AND EP.QUALIFY_STATUS = 'Q' THEN 'UnProcessed' " _
            & "WHEN EP.STATUS = 'F' AND EP.QUALIFY_STATUS = 'Q' THEN 'Finalized' END " _
            & "DROP TABLE #LAST13WEEK"

        Dim com_EV_CLAIM As New SqlCommand(cmd_EV_CLAIM, sqlcon)
        Dim da_EV_CLAIM As New SqlDataAdapter(com_EV_CLAIM)
        Dim ds_EV_CLAIM As New DataSet
        Try
            getDbCon()
            sqlcon.Open()
            da_EV_CLAIM.Fill(ds_EV_CLAIM)
            da_EV_CLAIM.SelectCommand.CommandTimeout = 600
            ds_EV_CLAIM.Tables(0).TableName = "EV_REP"
            frmEV_RVClaim.EV_REPBindingSource.DataSource = ds_EV_CLAIM
            sqlcon.Close()
        Catch ex As Exception
            MessageBox.Show("Error on passing data to Evoucher Claim Report, with message: " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            sqlcon.Close()
        End Try
        Cursor = Cursors.Default
    End Sub

    Public Sub Get_EV_Participant()
        Cursor = Cursors.AppStarting
        GetComp_Dist_EVCAT()
        Dim cmd_EV_Participant As String
        cmd_EV_Participant = "DECLARE @DIST NVARCHAR(10) " _
            & "SELECT @DIST = (SELECT DISTRIBUTOR FROM DISTRIBUTOR) " _
            & "select A.EVOUCHER_MASTER," _
            & "E.LDESC AS CAT_DESC,A.EVOUCHER_CATEGORY," _
            & "D.LDESC AS MAS_DESC," _
            & "D.FROM_DATE,D.END_DATE," _
            & "D.VOUCHER_VALUE," _
            & "A.TOWN+A.LOCALITY+A.SLOCALITY+A.POP AS POP_CODE, B.NAME AS POP_NAME," _
            & "C.LDESC AS POP_TYPE_DESC,F.POP_ALLOCATION AS ALLOCATED_BUDGET " _
            & "FROM EV_EVOUCHER_POP A " _
            & "LEFT JOIN POP B ON A.POP = B.POP " _
            & "LEFT JOIN POP_TYPE C ON B.POPTYPE=C.POPTYPE " _
            & "LEFT JOIN EV_EVOUCHER_CATEGORY D ON A.EVOUCHER_CATEGORY=D.EVOUCHER_CATEGORY " _
            & "LEFT JOIN EV_EVOUCHER_MASTER E ON A.EVOUCHER_MASTER = E.EVOUCHER_MASTER " _
            & "LEFT JOIN EV_EVOUCHER_ALLOCATION F ON A.EVOUCHER_CATEGORY = F.EVOUCHER_CATEGORY " _
            & "WHERE A.EVOUCHER_MASTER+A.EVOUCHER_CATEGORY IN ('" & EV_CAT & "') " _
            & "AND F.VALUE_COMB LIKE ('%'+@DIST)"
        Dim com_EV_Participant As New SqlCommand(cmd_EV_Participant, sqlcon)
        Dim da_EV_Participant As New SqlDataAdapter(com_EV_Participant)
        Dim ds_EV_PArticipant As New DataSet
        Try
            getDbCon()
            sqlcon.Open()
            da_EV_Participant.Fill(ds_EV_PArticipant)
            da_EV_Participant.SelectCommand.CommandTimeout = 600
            ds_EV_PArticipant.Tables(0).TableName = "EV_PARTI"
            frmEV_RVParticipant.EV_PARTIBindingSource.DataSource = ds_EV_PArticipant
            sqlcon.Close()
        Catch ex As Exception
            MessageBox.Show("Error on passing data to Evoucher Claim Report, with message: " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            sqlcon.Close()
        End Try
        Cursor = Cursors.Default
    End Sub

    Private Sub btnViewRep_Click(sender As Object, e As EventArgs) Handles btnViewRep.Click
        If cbRepType.SelectedIndex = 0 Then
            frmEV_RVClaim.Show()
        ElseIf cbRepType.SelectedIndex = 1 Then
            frmEV_RVParticipant.Show()
        End If
    End Sub
End Class