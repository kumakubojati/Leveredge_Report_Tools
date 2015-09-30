Imports System.IO
Imports System.IO.File
Imports System.Data
Imports System.Data.SqlClient

Public Class frmTarget

    Dim sqlcon As New SqlConnection
    Dim readconnection As String
    Dim cekresult As Integer
    Public Sub GetDBCOn()
        Dim strfile As String = My.Application.Info.DirectoryPath & "\Conf.ini"
        readconnection = ReadAllLines(strfile)(3)
        sqlcon.ConnectionString = readconnection
    End Sub

    Public Sub InitdgDetailTarget()
        Dim year As String = cbYear.SelectedValue.ToString
        Dim period As String = cbPeriod.SelectedValue.ToString
        Dim comstr1 As String = "select A.DSR,A.NAME,A.SELLCAT_DESC,B.ECO_TARGET AS ECO,B.BP_TARGET AS BP,B.LPPC_TARGET AS LPPC FROM DSR A " _
                                & "FULL OUTER JOIN TARGET_BYDSR B ON A.DSR=B.DSR WHERE B.TARGET_YEAR = '" _
                                & year & "' AND B.JC_TARGET='" & period & "'"
        Dim comstr2 As String = "SELECT DSR,NAME,SELLCAT_DESC AS 'SELLING CATEGORY',(SELECT '') As ECO,(SELECT '') As BP,(SELECT '') As LPPC " _
                                & "FROM DSR"
        Dim cmdstr1 As New SqlCommand(comstr1, sqlcon)
        Dim cmdstr2 As New SqlCommand(comstr2, sqlcon)
        Try
            GetDBCOn()
            sqlcon.Open()
            Dim result1 As String = cmdstr1.ExecuteScalar
            If Not String.IsNullOrEmpty(result1) OrElse IsDBNull(result1) Then
                Dim da As New SqlDataAdapter(comstr1, sqlcon)
                Dim dt As New DataTable
                da.Fill(dt)
                dgDetailTarget.DataSource = dt
                cekresult = 1
                btnDelete.Enabled = True
                btnUpdate.Enabled = True
            Else
                Dim da As New SqlDataAdapter(comstr2, sqlcon)
                Dim dt As New DataTable
                da.Fill(dt)
                dgDetailTarget.DataSource = dt
                btnSave.Enabled = True
            End If
        Catch ex As Exception
            MessageBox.Show("Error on Filling the Data, with message: " & ex.Message.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            sqlcon.Close()
        End Try
        sqlcon.Close()
    End Sub

    Public Sub GetCBYearPeriod()
        Dim comstr1 As String = "select DISTINCT(DATEPART(YYYY,DOC_DATE)) as DataYear from CASHMEMO " _
                                & "GROUP BY DOC_DATE"
        Dim comstr2 As String = "SELECT DISTINCT JCNO from JC_WEEK GROUP BY JCNO"
        Dim cmdstr1 As New SqlCommand(comstr1, sqlcon)
        Dim cmdstr2 As New SqlCommand(comstr2, sqlcon)
        Dim da As New SqlDataAdapter
        Dim ds As New DataSet
        Try
            GetDBCOn()
            sqlcon.Open()
            da.SelectCommand = cmdstr1
            da.Fill(ds, "CASHMEMO")

            da.SelectCommand = cmdstr2
            da.Fill(ds, "JC_WEEK")

            cbYear.DataSource = ds.Tables(0)
            cbYear.ValueMember = "DataYear"
            cbYear.DisplayMember = "DataYear"

            cbPeriod.DataSource = ds.Tables(1)
            cbPeriod.ValueMember = "JCNO"
            cbPeriod.DisplayMember = "JCNO"
            sqlcon.Close()
        Catch ex As Exception
            MessageBox.Show("Error on getting data, with message : " & ex.Message.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            sqlcon.Close()
        End Try
    End Sub
    Private Sub frmTarget_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        GetCBYearPeriod()
    End Sub

    Private Sub btnFetchData_Click(sender As Object, e As EventArgs) Handles btnFetchData.Click
        If cbYear.SelectedValue Is Nothing Then
            MessageBox.Show("Please Choose the year first", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            cbYear.Focus()
        ElseIf cbPeriod.SelectedValue Is Nothing Then
            MessageBox.Show("Please Choose the period first", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            cbPeriod.Focus()
        Else
            InitdgDetailTarget()
        End If

    End Sub

    Private Sub frmTarget_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        frmMain.Show()
    End Sub

    Private Sub btnForAll_Click(sender As Object, e As EventArgs) Handles btnForAll.Click
        If dgDetailTarget.Rows.Count > 1 Then
            For i As Integer = 0 To dgDetailTarget.Rows.Count - 1
                dgDetailTarget.Rows(i).Cells(3).Value = txtECO.Text
                dgDetailTarget.Rows(i).Cells(4).Value = txtBP.Text
                dgDetailTarget.Rows(i).Cells(5).Value = txtLPPC.Text
            Next i
        Else
            Exit Sub
        End If
        
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Dim year As String = cbYear.SelectedValue.ToString
        Dim period As String = cbPeriod.SelectedValue.ToString
        Dim comsave As String
        comsave = "INSERT INTO TARGET_BYDSR(TARGET_YEAR,JC_TARGET,DSR,ECO_TARGET,BP_TARGET,LPPC_TARGET)VALUES(" _
            & "@YEAR,@PERIOD,@DSR,@ECO,@BP,@LPPC)"
        Dim cmdsave As New SqlCommand(comsave, sqlcon)
        cmdsave.Parameters.Add("@YEAR", SqlDbType.NVarChar, 50)
        cmdsave.Parameters.Add("@PERIOD", SqlDbType.NVarChar, 50)
        cmdsave.Parameters.Add("@DSR", SqlDbType.NVarChar, 50)
        cmdsave.Parameters.Add("@ECO", SqlDbType.NVarChar, 50)
        cmdsave.Parameters.Add("@BP", SqlDbType.NVarChar, 50)
        cmdsave.Parameters.Add("@LPPC", SqlDbType.NVarChar, 50)

        Try
            GetDBCOn()
            sqlcon.Open()
            cmdsave.Prepare()
            For Each row As DataGridViewRow In dgDetailTarget.Rows
                If Not row.IsNewRow Then
                    cmdsave.Parameters("@YEAR").Value = year
                    cmdsave.Parameters("@PERIOD").Value = period
                    cmdsave.Parameters("@DSR").Value = row.Cells(0).Value.ToString
                    cmdsave.Parameters("@ECO").Value = row.Cells(3).Value.ToString
                    cmdsave.Parameters("@BP").Value = row.Cells(4).Value.ToString
                    cmdsave.Parameters("@LPPC").Value = row.Cells(5).Value.ToString
                End If
                cmdsave.ExecuteNonQuery()
            Next row
        Catch ex As Exception
            MessageBox.Show("Error on storing data, with message : " & ex.Message.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            sqlcon.Close()
        End Try
        sqlcon.Close()
        txtBP.Clear()
        txtECO.Clear()
        txtLPPC.Clear()
        btnSave.Enabled = False
        InitdgDetailTarget()
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        If cbYear.SelectedValue Is Nothing Then
            MessageBox.Show("Please select the year first", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        ElseIf cbPeriod.SelectedValue Is Nothing Then
            MessageBox.Show("Please select the year first", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Else
            Dim year As String = cbYear.SelectedValue.ToString
            Dim period As String = cbPeriod.SelectedValue.ToString
            Dim comdel As String
            comdel = "DELETE FROM TARGET_BYDSR WHERE TARGET_YEAR='" & cbYear.SelectedValue.ToString & "' AND JC_TARGET='" _
                & cbPeriod.SelectedValue.ToString & "'"
            Dim cmddel As New SqlCommand(comdel, sqlcon)
            Try
                GetDBCOn()
                sqlcon.Open()
                cmddel.ExecuteNonQuery()
                MessageBox.Show("Data successfully deleted", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                MessageBox.Show("Error on deleting data, with message: " & ex.Message.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
                sqlcon.Close()
            End Try
        End If
        sqlcon.Close()
        InitdgDetailTarget()
    End Sub

    Private Sub btnUpdate_Click(sender As Object, e As EventArgs) Handles btnUpdate.Click
        If cbYear.SelectedValue Is Nothing Then
            MessageBox.Show("Please select the year first", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        ElseIf cbPeriod.SelectedValue Is Nothing Then
            MessageBox.Show("Please select the year first", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Else
            Dim comupd As String
            comupd = "UPDATE TARGET_BYDSR SET ECO_TARGET=@ECO, BP_TARGET=@BP, LPPC_TARGET=@LPPC WHERE " _
                & "TARGET_YEAR=@YEAR AND JC_TARGET=@PERIOD AND DSR=@DSR"
            Dim cmdupd As New SqlCommand(comupd, sqlcon)
            cmdupd.Parameters.Add("@ECO", SqlDbType.NVarChar, 50)
            cmdupd.Parameters.Add("@BP", SqlDbType.NVarChar, 50)
            cmdupd.Parameters.Add("@LPPC", SqlDbType.NVarChar, 50)
            cmdupd.Parameters.Add("@YEAR", SqlDbType.NVarChar, 50)
            cmdupd.Parameters.Add("@PERIOD", SqlDbType.NVarChar, 50)
            cmdupd.Parameters.Add("@DSR", SqlDbType.NVarChar, 50)

            Try
                GetDBCOn()
                sqlcon.Open()
                cmdupd.Prepare()
                For Each row As DataGridViewRow In dgDetailTarget.Rows
                    If Not row.IsNewRow Then
                        cmdupd.Parameters("@ECO").Value = row.Cells(2).Value.ToString
                        cmdupd.Parameters("@BP").Value = row.Cells(3).Value.ToString
                        cmdupd.Parameters("@LPPC").Value = row.Cells(4).Value.ToString
                        cmdupd.Parameters("@YEAR").Value = cbYear.SelectedValue.ToString
                        cmdupd.Parameters("@PERIOD").Value = cbPeriod.SelectedValue.ToString
                        cmdupd.Parameters("@DSR").Value = row.Cells(0).Value.ToString
                    End If
                    cmdupd.ExecuteNonQuery()
                Next row
            Catch ex As Exception
                MessageBox.Show("Error on Updating data, with message: " & ex.Message.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                sqlcon.Close()
            End Try
        End If
        sqlcon.Close()
        MessageBox.Show("Data successfully updated", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
        InitdgDetailTarget()
    End Sub

    Private Sub cbPeriod_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbPeriod.SelectedIndexChanged
        txtECO.Clear()
        txtBP.Clear()
        txtLPPC.Clear()
        dgDetailTarget.DataSource = ""
        btnSave.Enabled = False
        btnUpdate.Enabled = False
        btnDelete.Enabled = False
    End Sub
End Class