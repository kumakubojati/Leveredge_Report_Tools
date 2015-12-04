Imports System.Data
Imports System.Data.SqlClient
Imports System.IO.File

Public NotInheritable Class SSUploadDB

    'TODO: This form can easily be set as the splash screen for the application by going to the "Application" tab
    '  of the Project Designer ("Properties" under the "Project" menu).
    Dim sqlcon As New SqlConnection
    Dim constr As String

    Public Sub GetDBCon()
        Dim strfile As String = My.Settings("FCSDBCon").ToString
        'constr = ReadAllLines(strfile)(3)
        sqlcon.ConnectionString = strfile
    End Sub

    Private Sub SSUploadDB_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'Set up the dialog text at runtime according to the application's assembly information.  

        'TODO: Customize the application's assembly information in the "Application" pane of the project 
        '  properties dialog (under the "Project" menu).

        'Application title
        'If My.Application.Info.Title <> "" Then
        '    ApplicationTitle.Text = My.Application.Info.Title
        'Else
        '    'If the application title is missing, use the application name, without the extension
        '    ApplicationTitle.Text = System.IO.Path.GetFileNameWithoutExtension(My.Application.Info.AssemblyName)
        'End If

        'Format the version information using the text set into the Version control at design time as the
        '  formatting string.  This allows for effective localization if desired.
        '  Build and revision information could be included by using the following code and changing the 
        '  Version control's designtime text to "Version {0}.{1:00}.{2}.{3}" or something similar.  See
        '  String.Format() in Help for more information.
        '
        '    Version.Text = System.String.Format(Version.Text, My.Application.Info.Version.Major, My.Application.Info.Version.Minor, My.Application.Info.Version.Build, My.Application.Info.Version.Revision)

        'Version.Text = System.String.Format(Version.Text, My.Application.Info.Version.Major, My.Application.Info.Version.Minor)

        ''Copyright info
        'Copyright.Text = My.Application.Info.Copyright
        System.Windows.Forms.Control.CheckForIllegalCrossThreadCalls = False
        BWUpdateDB.RunWorkerAsync()
    End Sub

    Public Sub UpdateDistributor()
        Dim comstr As String
        comstr = "MERGE FCS_REPORT.dbo.DISTRIBUTOR AS TARGET USING " _
            & "(SELECT DISTRIBUTOR AS DTCODE,NAME AS DTNAME, " _
            & "ADDRESS1 + ' ' + ADDRESS2 AS DTADDRESS FROM Centegy_SnDPro_UID.dbo.DISTRIBUTOR) AS SOURCE " _
            & "ON (TARGET.DTCODE = SOURCE.DTCODE) WHEN MATCHED AND TARGET.DTCODE <> SOURCE.DTCODE OR " _
            & "TARGET.DTNAME <> SOURCE.DTNAME OR TARGET.DTADDRESS <> SOURCE.DTADDRESS THEN " _
            & "UPDATE SET TARGET.DTNAME = SOURCE.DTNAME, TARGET.DTADDRESS = SOURCE.DTADDRESS " _
            & "WHEN NOT MATCHED BY TARGET THEN INSERT(DTCODE,DTNAME,DTADDRESS)VALUES (SOURCE.DTCODE,SOURCE.DTNAME,SOURCE.DTADDRESS);"
        Dim cmdstr As New SqlCommand(comstr, sqlcon)
        Try
            GetDBCon()
            sqlcon.Open()
            cmdstr.ExecuteNonQuery()
            sqlcon.Close()
        Catch ex As Exception
            MessageBox.Show("Error on Update Distributor, with message: " & ex.Message.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            sqlcon.Close()
        End Try
    End Sub

    Public Sub UpdateCalendar()
        Dim comstr As String
        comstr = "CREATE TABLE #TempCal(DAYS_NO INT IDENTITY(1,1) NOT NULL PRIMARY KEY,DATE DATE NOT NULL,DAY_NAME NVARCHAR(50) NULL, " _
            & "WEEK_NO NVARCHAR(50) NULL,YEAR NVARCHAR(4) NULL) SET IDENTITY_INSERT #TempCal OFF; DECLARE @date1 DATE, @date2 DATE " _
            & "SELECT @date1 = MIN(START_DATE) FROM Centegy_SnDPro_UID.dbo.JC_WEEK SELECT @date2 = MAX(END_DATE) FROM Centegy_SnDPro_UID.dbo.JC_WEEK " _
            & "INSERT INTO #TempCal(DATE,DAY_NAME,WEEK_NO,YEAR) EXEC GetCalendar @date1,@date2 " _
            & "MERGE FCS_REPORT.dbo.Calendar AS TARGET USING #TempCal AS SOURCE ON (TARGET.DATE = SOURCE.DATE) " _
            & "WHEN MATCHED AND TARGET.WEEK_NO <> SOURCE.WEEK_NO OR TARGET.YEAR <> SOURCE.YEAR THEN " _
            & "UPDATE SET TARGET.WEEK_NO = SOURCE.WEEK_NO,TARGET.YEAR = SOURCE.YEAR WHEN NOT MATCHED BY TARGET THEN " _
            & "INSERT(DATE,DAY_NAME,WEEK_NO,YEAR)VALUES(SOURCE.DATE,SOURCE.DAY_NAME,SOURCE.WEEK_NO,SOURCE.YEAR); " _
            & "DROP TABLE #TempCal"
        Dim cmdstr As New SqlCommand(comstr, sqlcon)
        Try
            GetDBCon()
            sqlcon.Open()
            cmdstr.ExecuteNonQuery()
            sqlcon.Close()
        Catch ex As Exception
            MessageBox.Show("Error on Update Calendar, with message: " & ex.Message.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            sqlcon.Close()
        End Try

    End Sub

    Public Sub UpdateCashmemo()
        Dim comstr As String
        comstr = "MERGE FCS_REPORT.dbo.CASHMEMO AS TARGET USING " _
            & "(SELECT DOC_NO,DOC_DATE,DSR,PJP,SECTION,POP,REF_DOCUMENT FROM Centegy_SnDPro_UID.dbo.CASHMEMO " _
            & "WHERE CASHMEMO_TYPE IN('01','03') AND REF_DOCUMENT='GN') AS SOURCE " _
            & "ON (TARGET.DOC_NO = SOURCE.DOC_NO AND TARGET.DOC_DATE = SOURCE.DOC_DATE) " _
            & "WHEN MATCHED AND TARGET.DOC_NO <> SOURCE.DOC_NO OR " _
            & "TARGET.DOC_DATE <> SOURCE.DOC_DATE OR TARGET.DSR <> SOURCE.DSR OR " _
            & "TARGET.PJP <> SOURCE.PJP OR TARGET.SECTION <> SOURCE.SECTION OR " _
            & "TARGET.POP <> SOURCE.POP OR TARGET.REF_DOC<> SOURCE.REF_DOCUMENT THEN " _
            & "UPDATE SET TARGET.DOC_NO = SOURCE.DOC_NO,TARGET.DSR = SOURCE.DSR," _
            & "TARGET.PJP=SOURCE.PJP,TARGET.SECTION=SOURCE.SECTION," _
            & "TARGET.POP=SOURCE.POP,TARGET.REF_DOC=SOURCE.REF_DOCUMENT " _
            & "WHEN NOT MATCHED BY TARGET THEN INSERT(DOC_NO,DOC_DATE,DSR,PJP,SECTION,POP,REF_DOC)" _
            & "VALUES(SOURCE.DOC_NO,SOURCE.DOC_DATE,SOURCE.DSR,SOURCE.PJP,SOURCE.SECTION,SOURCE.POP,SOURCE.REF_DOCUMENT);"
        Dim cmdstr As New SqlCommand(comstr, sqlcon)
        Try
            GetDBCon()
            sqlcon.Open()
            cmdstr.ExecuteNonQuery()
            sqlcon.Close()
        Catch ex As Exception
            MessageBox.Show("Error on Update Cashmemo, with message: " & ex.Message.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            sqlcon.Close()
        End Try

    End Sub

    Public Sub UpdateCM_Detail()
        Dim comstr1, comstr2 As String
        comstr1 = "TRUNCATE TABLE FCS_REPORT.dbo.CM_DETAIL"
        comstr2 = "insert into FCS_REPORT.dbo.CM_DETAIL(DOC_NO,DOC_DATE,SKU) select A.DOC_NO,A.DOC_DATE,B.SKU " _
            & "from Centegy_SnDPro_UID.dbo.CASHMEMO A left join Centegy_SnDPro_UID.dbo.CASHMEMO_DETAIL B ON A.DOC_NO=B.DOC_NO " _
            & "where A.CASHMEMO_TYPE in ('01','03') and A.REF_DOCUMENT='GN' AND B.AMOUNT > 0"
        Dim cmdstr1 As New SqlCommand(comstr1, sqlcon)
        Dim cmdstr2 As New SqlCommand(comstr2, sqlcon)
        Try
            GetDBCon()
            sqlcon.Open()
            cmdstr1.ExecuteNonQuery()
            cmdstr2.CommandTimeout = 600
            cmdstr2.ExecuteNonQuery()
            sqlcon.Close()
        Catch ex As Exception
            MessageBox.Show("Error on Update Cashmemo Detail, with message: " & ex.Message.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            sqlcon.Close()
        End Try
    End Sub

    Public Sub UpdateDSR()
        Dim comstr As String
        comstr = "MERGE FCS_REPORT.dbo.DSR AS TARGET USING " _
            & "(SELECT B.PJP,A.DSR,B.LDESC,B.SELL_CATEGORY AS SELLCAT,C.SDESC AS SELLCAT_DESC,A.JOB_TYPE " _
            & "FROM Centegy_SnDPro_UID.dbo.DSR A INNER JOIN Centegy_SnDPro_UID.dbo.PJP_HEAD B ON A.DSR=B.DSR " _
            & "INNER JOIN Centegy_SnDPro_UID.dbo.SELLING_CATEGORY C ON B.SELL_CATEGORY=C.SELL_CATEGORY " _
            & "WHERE A.JOB_TYPE IN ('01','03')) AS SOURCE ON (TARGET.PJP = SOURCE.PJP) " _
            & "WHEN MATCHED AND TARGET.PJP <> SOURCE.PJP OR " _
            & "TARGET.DSR <> SOURCE.DSR OR TARGET.LDESC <> SOURCE.LDESC OR " _
            & "TARGET.SELLCAT <> SOURCE.SELLCAT OR " _
            & "TARGET.SELLCAT_DESC <> SOURCE.SELLCAT_DESC OR TARGET.JOB_TYPE <> SOURCE.JOB_TYPE THEN " _
            & "UPDATE SET TARGET.PJP=SOURCE.PJP,TARGET.DSR=SOURCE.DSR,TARGET.LDESC=SOURCE.LDESC," _
            & "TARGET.SELLCAT=SOURCE.SELLCAT,TARGET.SELLCAT_DESC=SOURCE.SELLCAT_DESC,TARGET.JOB_TYPE=SOURCE.JOB_TYPE " _
            & "WHEN NOT MATCHED BY TARGET THEN INSERT(PJP,DSR,LDESC,SELLCAT,SELLCAT_DESC,JOB_TYPE)" _
            & "VALUES(SOURCE.PJP,SOURCE.DSR,SOURCE.LDESC,SOURCE.SELLCAT,SOURCE.SELLCAT_DESC,SOURCE.JOB_TYPE);"
        Dim cmdstr As New SqlCommand(comstr, sqlcon)
        Try
            GetDBCon()
            sqlcon.Open()
            cmdstr.ExecuteNonQuery()
            sqlcon.Close()
        Catch ex As Exception
            MessageBox.Show("Error on Update DSR, with message: " & ex.Message.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            sqlcon.Close()
        End Try

    End Sub

    Public Sub UpdateJC_Week()
        Dim comstr As String
        comstr = "MERGE FCS_REPORT.dbo.JC_WEEK AS TARGET USING " _
            & "(SELECT YEAR,substring(WEEK_NO,PATINDEX('%[^0]%',WEEK_NO),10) AS WEEK_NO,JCNO,START_DATE,END_DATE " _
            & "FROM Centegy_SnDPro_UID.dbo.JC_WEEK) AS SOURCE " _
            & "ON (TARGET.YEAR=SOURCE.YEAR AND TARGET.WEEK_NO=SOURCE.WEEK_NO AND TARGET.JCNO=SOURCE.JCNO) " _
            & "WHEN MATCHED AND TARGET.YEAR <> SOURCE.YEAR OR TARGET.WEEK_NO <> SOURCE.WEEK_NO OR " _
            & "TARGET.JCNO <> SOURCE.JCNO OR TARGET.START_DATE <> SOURCE.START_DATE OR " _
            & "TARGET.END_DATE <> SOURCE.END_DATE THEN UPDATE SET TARGET.YEAR=SOURCE.YEAR, TARGET.WEEK_NO=SOURCE.WEEK_NO," _
            & "TARGET.JCNO=SOURCE.JCNO, TARGET.START_DATE=SOURCE.START_DATE,TARGET.END_DATE=SOURCE.END_DATE " _
            & "WHEN NOT MATCHED BY TARGET THEN INSERT(YEAR,WEEK_NO,JCNO,START_DATE,END_DATE)VALUES" _
            & "(SOURCE.YEAR,SOURCE.WEEK_NO,SOURCE.JCNO,SOURCE.START_DATE,SOURCE.END_DATE);"
        Dim cmdstr As New SqlCommand(comstr, sqlcon)
        Try
            GetDBCon()
            sqlcon.Open()
            cmdstr.ExecuteNonQuery()
            sqlcon.Close()
        Catch ex As Exception
            MessageBox.Show("Error on Update JC, with message: " & ex.Message.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            sqlcon.Close()
        End Try
    End Sub

    Public Sub UpdatePJP()
        Dim comstr1, comstr2 As String
        comstr1 = "TRUNCATE TABLE FCS_REPORT.dbo.PJP"
        comstr2 = "INSERT INTO FCS_REPORT.dbo.PJP(PJP,SECTION,WEEK_NO,DSR,SAT,SUN,MON,TUE,WED,THU,FRI) " _
            & "SELECT A.PJP,b.SECTION,b.WEEK_NO,A.DSR,b.SAT,b.SUN,b.MON,b.TUE,b.WED,b.THU,b.FRI " _
            & "FROM Centegy_SnDPro_UID.dbo.PJP_HEAD A INNER JOIN Centegy_SnDPro_UID.dbo.PJP_DETAIL B ON A.PJP=B.PJP " _
            & "INNER JOIN Centegy_SnDPro_UID.dbo.DSR C ON A.DSR=C.DSR WHERE C.JOB_TYPE IN ('01','03')"
        Dim cmdstr1 As New SqlCommand(comstr1, sqlcon)
        Dim cmdstr2 As New SqlCommand(comstr2, sqlcon)
        Try
            GetDBCon()
            sqlcon.Open()
            cmdstr1.ExecuteNonQuery()
            cmdstr2.ExecuteNonQuery()
            sqlcon.Close()
        Catch ex As Exception
            MessageBox.Show("Error on Update PJP, with message: " & ex.Message.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            sqlcon.Close()
        End Try
    End Sub

    Public Sub UpdateSalesReturn()
        Dim comstr As String
        comstr = "MERGE FCS_REPORT.dbo.SALES_RETURN AS TARGET USING " _
            & "(select A.DOC_NO,A.DOC_DATE,A.DSR,A.PJP,A.SECTION,A.POP,A.REF_DOC_NO,B.DOC_DATE AS REF_DOC_DATE from " _
            & "(select DOC_NO,DOC_DATE,DSR,PJP,SECTION,POP,REF_DOC_NO from Centegy_SnDPro_UID.dbo.CASHMEMO WHERE CASHMEMO_TYPE ='19' AND REF_DOCUMENT='CM' " _
            & "AND REF_DOCUMENT2 = 'GR') A INNER JOIN (select DOC_NO,DOC_DATE from Centegy_SnDPro_UID.dbo.CASHMEMO " _
            & "WHERE DOC_NO IN (select REF_DOC_NO from Centegy_SnDPro_UID.dbo.CASHMEMO WHERE CASHMEMO_TYPE ='19' AND REF_DOCUMENT='CM' AND REF_DOCUMENT2 = 'GR')) B " _
            & "ON A.REF_DOC_NO=B.DOC_NO) AS SOURCE ON (TARGET.DOC_NO=SOURCE.DOC_NO AND TARGET.DOC_DATE=SOURCE.DOC_DATE) " _
            & "WHEN MATCHED AND TARGET.DOC_NO <> SOURCE.DOC_NO OR TARGET.DOC_DATE <> SOURCE.DOC_DATE OR " _
            & "TARGET.DSR <> SOURCE.DSR OR TARGET.PJP <> SOURCE.PJP OR TARGET.SECTION <> SOURCE.SECTION THEN " _
            & "UPDATE SET TARGET.DOC_NO=SOURCE.DOC_NO,TARGET.DOC_DATE=SOURCE.DOC_DATE,TARGET.DSR=SOURCE.DSR," _
            & "TARGET.PJP=SOURCE.PJP,TARGET.SECTION=SOURCE.SECTION,TARGET.POP=SOURCE.POP,TARGET.REF_DOC_NO=SOURCE.REF_DOC_NO," _
            & "TARGET.REF_DOC_DATE=SOURCE.REF_DOC_DATE WHEN NOT MATCHED BY TARGET THEN " _
            & "INSERT(DOC_NO,DOC_DATE,DSR,PJP,SECTION,POP,REF_DOC_NO,REF_DOC_DATE)VALUES" _
            & "(SOURCE.DOC_NO,SOURCE.DOC_DATE,SOURCE.DSR,SOURCE.PJP,SOURCE.SECTION,SOURCE.POP,SOURCE.REF_DOC_NO,SOURCE.REF_DOC_DATE);"
        Dim cmdstr As New SqlCommand(comstr, sqlcon)
        Try
            GetDBCon()
            sqlcon.Open()
            cmdstr.ExecuteNonQuery()
            sqlcon.Close()
        Catch ex As Exception
            MessageBox.Show("Error on Update Sales Return, with message: " & ex.Message.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            sqlcon.Close()
        End Try
    End Sub

    Public Sub UpdateSR_Detail()
        Dim comstr, comstr2 As String
        comstr = "TRUNCATE TABLE FCS_REPORT.dbo.SR_DETAIL"
        comstr2 = "insert into FCS_REPORT.dbo.SR_DETAIL(DOC_NO,DOC_DATE,SKU,REF_DOC_NO,REF_DOC_DATE)" _
            & "select x.DOC_NO,x.DOC_DATE,x.SKU,x.REF_DOC_NO,y.DOC_DATE AS REF_DOC_DATE from " _
            & "(select A.DOC_NO,convert(DATE,A.DOC_DATE)as DOC_DATE,A.REF_DOC_NO,B.SKU " _
            & "from Centegy_SnDPro_UID.dbo.CASHMEMO A left join Centegy_SnDPro_UID.dbo.CASHMEMO_DETAIL B ON A.DOC_NO=B.DOC_NO AND A.DOC_DATE=B.DOC_DATE " _
            & "where A.REF_DOCUMENT2='GR' and A.CASHMEMO_TYPE='19' AND B.AMOUNT < 0) x " _
            & "Inner join Centegy_SnDPro_UID.dbo.CASHMEMO Y ON x.REF_DOC_NO=Y.DOC_NO"
        Dim cmdstr1 As New SqlCommand(comstr, sqlcon)
        Dim cmdstr2 As New SqlCommand(comstr2, sqlcon)
        Try
            GetDBCon()
            sqlcon.Open()
            cmdstr1.ExecuteNonQuery()
            cmdstr2.CommandTimeout = 600
            cmdstr2.ExecuteNonQuery()
            sqlcon.Close()
        Catch ex As Exception
            MessageBox.Show("Error on Update Sales Return Detail, with message: " & ex.Message.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            sqlcon.Close()
        End Try
    End Sub

    Public Sub UpdateSPP()
        Dim comstr1, comstr2 As String
        comstr1 = "TRUNCATE TABLE FCS_REPORT.dbo.SECTION_POP_PERMANENT"
        comstr2 = "INSERT INTO FCS_REPORT.dbo.SECTION_POP_PERMANENT(PJP,SECTION,POP) SELECT A.PJP,A.SECTION,A.POP " _
           & "FROM Centegy_SnDPro_UID.dbo.SECTION_POP_PERMANENT A INNER JOIN Centegy_SnDPro_UID.dbo.PJP_HEAD B ON A.PJP=B.PJP " _
           & "INNER JOIN Centegy_SnDPro_UID.dbo.PJP_DETAIL C ON B.PJP= C.PJP INNER JOIN Centegy_SnDPro_UID.dbo.DSR D ON B.DSR = D.DSR " _
           & "WHERE D.JOB_TYPE IN ('01','03') GROUP BY A.PJP,A.SECTION,D.DSR,A.POP"
        Dim cmdstr1 As New SqlCommand(comstr1, sqlcon)
        Dim cmdstr2 As New SqlCommand(comstr2, sqlcon)
        Try
            GetDBCon()
            sqlcon.Open()
            cmdstr1.ExecuteNonQuery()
            cmdstr2.CommandTimeout = 600
            cmdstr2.ExecuteNonQuery()
            sqlcon.Close()
        Catch ex As Exception
            MessageBox.Show("Error on Update Section POP Permanent, with message: " & ex.Message.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            sqlcon.Close()
        End Try
    End Sub

    Public Sub UpdateSPBD()
        Dim comstr1, comstr2 As String
        comstr1 = "TRUNCATE TABLE FCS_REPORT.dbo.SECTION_PRMNT_BYDATE"
        comstr2 = "insert into SECTION_PRMNT_BYDATE(DATE,PJP,DSR,LDESC,SECTION,WEEK_NO,JCNO,DAY_VISIT,POP) " _
            & "SELECT F.Date,E.PJP,E.DSR,E.PJP_NAME,E.SECTION,E.WEEK_NO,G.JCNO,E.DAY_VISIT,E.POP FROM " _
            & "(select D.PJP,D.DSR,D.PJP_NAME,D.SECTION,D.WEEK_NO, DAY_VISIT= CASE D.DAY when '1000000' then 'Saturday' " _
            & "when '0100000' then 'Sunday' when '0010000' then 'Monday' when '0001000' then 'Tuesday' when '0000100' then 'Wednesday' " _
            & "when '0000010' then 'Thursday' when '0000001' then 'Friday' ELSE 'N/A' END, D.POP FROM " _
            & "(select A.PJP,A.SECTION,B.WEEK_NO, cast(B.SAT as varchar(2))+cast(B.SUN as varchar(2))+" _
            & "cast(B.MON as varchar(2))+cast(B.TUE as varchar(2))+cast(B.WED as varchar(2))+cast(B.THU as varchar(2))+" _
            & "cast(B.FRI as varchar(2)) as DAY,A.POP,B.DSR,C.LDESC AS PJP_NAME from SECTION_POP_PERMANENT A " _
            & "inner join PJP B on a.PJP=b.PJP and A.SECTION=b.SECTION inner join DSR C on B.PJP=C.PJP) D) E " _
            & "Inner Join Calendar F ON E.WEEK_NO = F.WEEK_NO AND E.DAY_VISIT=F.DAY_NAME LEFT OUTER join JC_WEEK G on F.WEEK_NO = G.WEEK_NO Order by F.Date ASC"
        Dim cmdstr1 As New SqlCommand(comstr1, sqlcon)
        Dim cmdstr2 As New SqlCommand(comstr2, sqlcon)
        Try
            GetDBCon()
            sqlcon.Open()
            cmdstr1.ExecuteNonQuery()
            cmdstr2.CommandTimeout = 600
            cmdstr2.ExecuteNonQuery()
            sqlcon.Close()
        Catch ex As Exception
            MessageBox.Show("Error on Update Section POP Permanent By Date, with message: " & ex.Message.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            sqlcon.Close()
        End Try
    End Sub

    Public Sub UpdateTarget()
        Dim comstr1, comstr2 As String
        comstr1 = "TRUNCATE TABLE FCS_REPORT.dbo.TARGET_BYDSR"
        comstr2 = "INSERT INTO FCS_REPORT.dbo.TARGET_BYDSR(TARGET_CODE,TARGET_YEAR,JC_TARGET,PJP,ECO_TARGET,BP_TARGET,LPPC_TARGET) " _
            & "select eco.REF_POLICY_ID,eco.YEAR,eco.JCNO,pjp.PJP ," _
            & "case when eco.range_from is null then 0 else eco.range_from end as target_eco ," _
            & "case when bp.RANGE_FROM is null then 0 else bp.RANGE_FROM end as target_bp," _
            & "case when LPPC.RANGE_FROM is null then 0 else LPPC.RANGE_FROM end as target_lppc " _
            & "from (select a.VALUE_COMB_FROM,b.RANGE_FROM ,c.JCNO,c.YEAR , d.REF_POLICY_ID from Centegy_SnDPro_UID.dbo.POLICY_DETAIL a " _
            & "left join Centegy_SnDPro_UID.dbo.POLICY_SLAB b on a.POLICY_ID = b.POLICY_ID left join (select policy_id," _
            & "LEFT(value_comb_from,4) as YEAR , RIGHT(value_comb_from,2) as JCNO from Centegy_SnDPro_UID.dbo.POLICY_DETAIL WHERE FIELD_COMB = 'YEAR~JCNO') c " _
            & "on a.POLICY_ID = c.POLICY_ID left join Centegy_SnDPro_UID.dbo.POLICY d on a.POLICY_ID = d.POLICY_ID where a.POLICY_ID in (" _
            & "select POLICY_ID from Centegy_SnDPro_UID.dbo.POLICY_DETAIL where  VALUE_COMB_FROM = '4111') " _
            & "and FIELD_COMB = 'pjp' and c.JCNO IN('01','02','03','04','05','06','07','08','09','10','11','12') and d.STATUS='1' " _
            & "and c.YEAR IN (SELECT DISTINCT DATENAME(YEAR,DOC_DATE) AS YEAR FROM Centegy_SnDPro_UID.dbo.CASHMEMO " _
            & "GROUP BY DOC_DATE) and d.REF_POLICY_ID in (select d.REF_POLICY_ID from Centegy_SnDPro_UID.dbo.POLICY_DETAIL a " _
            & "left join Centegy_SnDPro_UID.dbo.POLICY_SLAB b on a.POLICY_ID = b.POLICY_ID left join (select policy_id," _
            & "LEFT(value_comb_from,4) as YEAR , RIGHT(value_comb_from,2) as JCNO from Centegy_SnDPro_UID.dbo.POLICY_DETAIL WHERE FIELD_COMB = 'YEAR~JCNO') c " _
            & "on a.POLICY_ID = c.POLICY_ID left join Centegy_SnDPro_UID.dbo.POLICY d on a.POLICY_ID = d.POLICY_ID where a.POLICY_ID in (" _
            & "select POLICY_ID from Centegy_SnDPro_UID.dbo.POLICY_DETAIL where  VALUE_COMB_FROM = '4111') and FIELD_COMB = 'pjp' " _
            & "and c.JCNO IN('01','02','03','04','05','06','07','08','09','10','11','12') and d.STATUS='1' and c.YEAR IN " _
            & "(SELECT DISTINCT DATENAME(YEAR,DOC_DATE) AS YEAR FROM Centegy_SnDPro_UID.dbo.CASHMEMO GROUP BY DOC_DATE) " _
            & ")) eco left join (select a.VALUE_COMB_FROM,b.RANGE_FROM ,c.JCNO,c.YEAR , d.REF_POLICY_ID " _
            & "from Centegy_SnDPro_UID.dbo.POLICY_DETAIL a left join Centegy_SnDPro_UID.dbo.POLICY_SLAB b on a.POLICY_ID = b.POLICY_ID " _
            & "left join (select policy_id,LEFT(value_comb_from,4) as YEAR , RIGHT(value_comb_from,2) as JCNO from Centegy_SnDPro_UID.dbo.POLICY_DETAIL WHERE FIELD_COMB = 'YEAR~JCNO') c " _
            & "on a.POLICY_ID = c.POLICY_ID left join Centegy_SnDPro_UID.dbo.POLICY d on a.POLICY_ID = d.POLICY_ID where a.POLICY_ID in ( " _
            & "select POLICY_ID from Centegy_SnDPro_UID.dbo.POLICY_DETAIL where  VALUE_COMB_FROM = '4121') and FIELD_COMB = 'pjp' " _
            & "and c.JCNO IN('01','02','03','04','05','06','07','08','09','10','11','12') and d.STATUS='1' and c.YEAR IN " _
            & "(SELECT DISTINCT DATENAME(YEAR,DOC_DATE) AS YEAR FROM Centegy_SnDPro_UID.dbo.CASHMEMO GROUP BY DOC_DATE) and d.REF_POLICY_ID in " _
            & "(select d.REF_POLICY_ID from Centegy_SnDPro_UID.dbo.POLICY_DETAIL a left join Centegy_SnDPro_UID.dbo.POLICY_SLAB b on a.POLICY_ID = b.POLICY_ID " _
            & "left join (select policy_id,LEFT(value_comb_from,4) as YEAR , RIGHT(value_comb_from,2) as JCNO from Centegy_SnDPro_UID.dbo.POLICY_DETAIL WHERE FIELD_COMB = 'YEAR~JCNO') c " _
            & "on a.POLICY_ID = c.POLICY_ID left join Centegy_SnDPro_UID.dbo.POLICY d on a.POLICY_ID = d.POLICY_ID where a.POLICY_ID in (" _
            & "select POLICY_ID from Centegy_SnDPro_UID.dbo.POLICY_DETAIL where  VALUE_COMB_FROM = '4121') and FIELD_COMB = 'pjp' " _
            & "and c.JCNO IN('01','02','03','04','05','06','07','08','09','10','11','12') and d.STATUS='1' and c.YEAR IN " _
            & "(SELECT DISTINCT DATENAME(YEAR,DOC_DATE) AS YEAR FROM Centegy_SnDPro_UID.dbo.CASHMEMO GROUP BY DOC_DATE))) bp " _
            & "on eco.JCNO=bp.JCNO and eco.YEAR=bp.YEAR and eco.VALUE_COMB_FROM = bp.VALUE_COMB_FROM left join (" _
            & "select a.VALUE_COMB_FROM,b.RANGE_FROM ,c.JCNO,c.YEAR , d.REF_POLICY_ID from Centegy_SnDPro_UID.dbo.POLICY_DETAIL a " _
            & "left join Centegy_SnDPro_UID.dbo.POLICY_SLAB b on a.POLICY_ID = b.POLICY_ID left join (" _
            & "select policy_id,LEFT(value_comb_from,4) as YEAR , RIGHT(value_comb_from,2) as JCNO from Centegy_SnDPro_UID.dbo.POLICY_DETAIL WHERE FIELD_COMB = 'YEAR~JCNO') c " _
            & "on a.POLICY_ID = c.POLICY_ID left join Centegy_SnDPro_UID.dbo.POLICY d on a.POLICY_ID = d.POLICY_ID where a.POLICY_ID in (" _
            & "select POLICY_ID from Centegy_SnDPro_UID.dbo.POLICY_DETAIL where  VALUE_COMB_FROM = '4131') and FIELD_COMB = 'pjp' " _
            & "and c.JCNO IN('01','02','03','04','05','06','07','08','09','10','11','12') and d.STATUS='1' and c.YEAR IN " _
            & "(SELECT DISTINCT DATENAME(YEAR,DOC_DATE) AS YEAR FROM Centegy_SnDPro_UID.dbo.CASHMEMO GROUP BY DOC_DATE) " _
            & "and d.REF_POLICY_ID in (select d.REF_POLICY_ID from Centegy_SnDPro_UID.dbo.POLICY_DETAIL a " _
            & "left join Centegy_SnDPro_UID.dbo.POLICY_SLAB b on a.POLICY_ID = b.POLICY_ID left join (select policy_id," _
            & "LEFT(value_comb_from,4) as YEAR , RIGHT(value_comb_from,2) as JCNO from Centegy_SnDPro_UID.dbo.POLICY_DETAIL WHERE FIELD_COMB = 'YEAR~JCNO') c " _
            & "on a.POLICY_ID = c.POLICY_ID left join Centegy_SnDPro_UID.dbo.POLICY d on a.POLICY_ID = d.POLICY_ID where a.POLICY_ID in ( " _
            & "select POLICY_ID from Centegy_SnDPro_UID.dbo.POLICY_DETAIL where  VALUE_COMB_FROM = '4131') and FIELD_COMB = 'pjp' " _
            & "and c.JCNO IN('01','02','03','04','05','06','07','08','09','10','11','12') and d.STATUS='1' and c.YEAR IN " _
            & "(SELECT DISTINCT DATENAME(YEAR,DOC_DATE) AS YEAR FROM Centegy_SnDPro_UID.dbo.CASHMEMO GROUP BY DOC_DATE)" _
            & ")) LPPC on eco.JCNO = LPPC.JCNO and eco.YEAR=LPPC.YEAR and eco.VALUE_COMB_FROM=LPPC.VALUE_COMB_FROM " _
            & "left join Centegy_SnDPro_UID.dbo.PJP_HEAD pjp on eco.VALUE_COMB_FROM = pjp.pjp where eco.JCNO IN('01','02','03','04','05','06','07','08','09','10','11','12') " _
            & "and eco.YEAR IN (SELECT DISTINCT DATENAME(YEAR,DOC_DATE) AS YEAR FROM Centegy_SnDPro_UID.dbo.CASHMEMO GROUP BY DOC_DATE) " _
            & "and eco.REF_POLICY_ID in (select d.REF_POLICY_ID from Centegy_SnDPro_UID.dbo.POLICY_DETAIL a " _
            & "left join Centegy_SnDPro_UID.dbo.POLICY_SLAB b on a.POLICY_ID = b.POLICY_ID left join (select policy_id," _
            & "LEFT(value_comb_from,4) as YEAR , RIGHT(value_comb_from,2) as JCNO from Centegy_SnDPro_UID.dbo.POLICY_DETAIL WHERE FIELD_COMB = 'YEAR~JCNO') c " _
            & "on a.POLICY_ID = c.POLICY_ID left join Centegy_SnDPro_UID.dbo.POLICY d on a.POLICY_ID = d.POLICY_ID " _
            & "where a.POLICY_ID in (select POLICY_ID from Centegy_SnDPro_UID.dbo.POLICY_DETAIL where  VALUE_COMB_FROM = '4111') " _
            & "and FIELD_COMB = 'pjp' and c.JCNO IN('01','02','03','04','05','06','07','08','09','10','11','12') " _
            & "and d.STATUS='1' and c.YEAR IN (SELECT DISTINCT DATENAME(YEAR,DOC_DATE) AS YEAR FROM Centegy_SnDPro_UID.dbo.CASHMEMO GROUP BY DOC_DATE))"
        Dim cmdstr1 As New SqlCommand(comstr1, sqlcon)
        Dim cmdstr2 As New SqlCommand(comstr2, sqlcon)
        Try
            GetDBCon()
            sqlcon.Open()
            cmdstr1.ExecuteNonQuery()
            cmdstr2.CommandTimeout = 600
            cmdstr2.ExecuteNonQuery()
            sqlcon.Close()
        Catch ex As Exception
            MessageBox.Show("Error on Update Target, with message: " & ex.Message.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            sqlcon.Close()
        End Try
    End Sub

    Delegate Sub UpdatelblProgressDelegate(ByVal mystring As String)
    Public Sub UpdateTextlblProgress(ByVal mystring As String)
        Me.lblnotif.Text = mystring
    End Sub

    Private Sub BWUpdateDB_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BWUpdateDB.DoWork
        Dim UpdatelblProgDel As UpdatelblProgressDelegate = New UpdatelblProgressDelegate(AddressOf UpdateTextlblProgress)
        Dim newtxt As String

        Try
            newtxt = "Checking and Updating Distributor DT"
            Me.Invoke(UpdatelblProgDel, newtxt)
            UpdateDistributor()
            newtxt = "Checking and Updating Calendar"
            Me.Invoke(UpdatelblProgDel, newtxt)
            UpdateCalendar()
            newtxt = "Checking and Updating Cashmemo"
            Me.Invoke(UpdatelblProgDel, newtxt)
            UpdateCashmemo()
            newtxt = "Checking and Updating Cashmemo Detail"
            Me.Invoke(UpdatelblProgDel, newtxt)
            UpdateCM_Detail()
            newtxt = "Checking and Updating Sales Return"
            Me.Invoke(UpdatelblProgDel, newtxt)
            UpdateSalesReturn()
            newtxt = "Checking and Updating Sales Return Detail"
            UpdateSR_Detail()
            newtxt = "Checking and Updating DSR"
            Me.Invoke(UpdatelblProgDel, newtxt)
            UpdateDSR()
            newtxt = "Checking and Updating JC WEEK"
            Me.Invoke(UpdatelblProgDel, newtxt)
            UpdateJC_Week()
            newtxt = "Checking and Updating PJP"
            Me.Invoke(UpdatelblProgDel, newtxt)
            UpdatePJP()
            newtxt = "Checking and Updating Section POP Permanent"
            Me.Invoke(UpdatelblProgDel, newtxt)
            UpdateSPP()
            newtxt = "Checking and Updating Section POP Permanent By Date"
            Me.Invoke(UpdatelblProgDel, newtxt)
            UpdateSPBD()
            newtxt = "Checking and Updating FCS Target By DSR"
            Me.Invoke(UpdatelblProgDel, newtxt)
            UpdateTarget()
        Catch ex As Exception
            MessageBox.Show("Error On updating DB, with message : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub BWUpdateDB_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BWUpdateDB.RunWorkerCompleted
        frmFCS_Rep.Show()
        Me.Close()
    End Sub
End Class
