Imports System.Data.Odbc
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports CrystalDecisions.Windows.Forms
Public Class _frmViewReports
    Dim db As New ConnectionString
    Dim dt As DataTable
    Public cryRpt As New ReportDocument
    Public GetName As String
    Dim Group As Integer
    Dim DR As Integer
    Dim OLAP As Integer
    Dim validateRangeDate As Integer
    Dim Trans As Integer
    Public MonthSelect As String
    Public TransType As String
    Dim Transaction As String
    Dim Failedtrans As String
    Dim pickyear As Integer
    Dim SSID As Integer
    Public ReportPrint As Integer
    Public date12 As Date
    Public date22 As Date
    Dim getBranch As String
    Public sql As String
    Public connstring1 As String = "DSN=" & My.Settings.db_DSN & ";SERVER=" & My.Settings.db_Server & ";DATABASE=" & My.Settings.db_Name & ";UID=" & My.Settings.db_UName & ";PWD=" & My.Settings.db_Pass & ""
    Public conn As OdbcConnection = New OdbcConnection(connstring1)
#Region "Public Declaration"
    Public Sub ChangeDate()
        If m.Checked = True Then
            pickyear = 2
            DR = 1
            showyearonly()
            getMonth()
        ElseIf y.Checked = True Then
            pickyear = 1
            DR = 1
            showyearonly()
            getMonth()
        Else

        End If

    End Sub
    Public Sub GetNameBranch()
        If cbGroup.Text <> Nothing Then
            getBranch = cbGroup.Text
        ElseIf cbCluster.Text <> Nothing Then
            getBranch = cbCluster.Text
        ElseIf cbBranch.Text <> Nothing Then
            getBranch = cbBranch.Text
        Else
            getBranch = ""
        End If
    End Sub
    Public Sub getMonth()
        Dim day1 As String = DatePart(DateInterval.Month, dtpFrom.Value)
        Dim day2 As String = DatePart(DateInterval.Month, dtpTo.Value)
        Dim year1 As String = DatePart(DateInterval.Year, dtpFrom.Value)
        Dim year2 As String = DatePart(DateInterval.Year, dtpTo.Value)

        If day1 = 1 Then
            date12 = day1 & "/" & "1" & "/" & year1
        ElseIf day1 = 2 Then
            date12 = day1 & "/" & "1" & "/" & year1
        ElseIf day1 = 3 Then
            date12 = day1 & "/" & "1" & "/" & year1
        ElseIf day1 = 3 Then
            date12 = day1 & "/" & "1" & "/" & year1
        ElseIf day1 = 4 Then
            date12 = day1 & "/" & "1" & "/" & year1
        ElseIf day1 = 5 Then
            date12 = day1 & "/" & "1" & "/" & year1
        ElseIf day1 = 6 Then
            date12 = day1 & "/" & "1" & "/" & year1
        ElseIf day1 = 7 Then
            date12 = day1 & "/" & "1" & "/" & year1
        ElseIf day1 = 8 Then
            date12 = day1 & "/" & "1" & "/" & year1
        ElseIf day1 = 9 Then
            date12 = day1 & "/" & "1" & "/" & year1
        ElseIf day1 = 10 Then
            date12 = day1 & "/" & "1" & "/" & year1
        ElseIf day1 = 11 Then
            date12 = day1 & "/" & "1" & "/" & year1
        ElseIf day1 = 12 Then
            date12 = day1 & "/" & "1" & "/" & year1
        End If

        If day2 = 1 Then
            date22 = day2 & "/" & "31" & "/" & year2
        ElseIf day2 = 2 Then
            If Date.IsLeapYear(DatePart(DateInterval.Year, dtpTo.Value)) Then
                date22 = day2 & "/" & "29" & "/" & year2
            Else
                date22 = day2 & "/" & "28" & "/" & year2
            End If
        ElseIf day2 = 3 Then
            date22 = day2 & "/" & "31" & "/" & year2
        ElseIf day2 = 4 Then
            date22 = day2 & "/" & "30" & "/" & year2
        ElseIf day2 = 5 Then
            date22 = day2 & "/" & "31" & "/" & year2
        ElseIf day2 = 6 Then
            date22 = day2 & "/" & "30" & "/" & year2
        ElseIf day2 = 7 Then
            date22 = day2 & "/" & "31" & "/" & year2
        ElseIf day2 = 8 Then
            date22 = day2 & "/" & "31" & "/" & year2
        ElseIf day2 = 9 Then
            date22 = day2 & "/" & "30" & "/" & year2
        ElseIf day2 = 10 Then
            date22 = day2 & "/" & "31" & "/" & year2
        ElseIf day2 = 11 Then
            date22 = day2 & "/" & "30" & "/" & year2
        ElseIf day2 = 12 Then
            date22 = day2 & "/" & "31" & "/" & year2

        End If

    End Sub
    Public Sub showyearonly()
        Select Case pickyear

            Case 1
                dtpFrom.MinDate = New DateTime(1753, 1, 1)
                dtpFrom.MaxDate = New DateTime(9998, 1, 1)
                dtpFrom.Format = DateTimePickerFormat.Custom
                dtpFrom.CustomFormat = "yyyy"
                dtpFrom.Value = Today
                dtpFrom.ShowUpDown = True
                dtpTo.MinDate = New DateTime(1753, 1, 1)
                dtpTo.MaxDate = New DateTime(9998, 1, 1)
                dtpTo.Format = DateTimePickerFormat.Custom
                dtpTo.CustomFormat = "yyyy"
                dtpTo.Value = Today
                dtpTo.ShowUpDown = True

            Case 2
                dtpFrom.MinDate = New DateTime(1753, 1, 1)
                dtpFrom.MaxDate = New DateTime(9998, 1, 1)
                dtpFrom.Format = DateTimePickerFormat.Custom
                dtpFrom.CustomFormat = "MMMM-yyyy"
                'dtpFrom.Value = Today
                dtpFrom.ShowUpDown = True
                dtpTo.MinDate = New DateTime(1753, 1, 1)
                dtpTo.MaxDate = New DateTime(9998, 1, 1)
                dtpTo.Format = DateTimePickerFormat.Custom
                dtpTo.CustomFormat = "MMMM-yyyy"
                'dtpTo.Value = Today
                dtpTo.ShowUpDown = True
            Case Else
                dtpFrom.MinDate = New DateTime(1753, 1, 1)
                dtpFrom.MaxDate = New DateTime(9998, 1, 1)
                dtpFrom.Format = DateTimePickerFormat.Short
                dtpFrom.Value = Today
                dtpFrom.ShowUpDown = False
                dtpTo.MinDate = New DateTime(1753, 1, 1)
                dtpTo.MaxDate = New DateTime(9998, 1, 1)
                dtpTo.Value = Today
                dtpTo.Format = DateTimePickerFormat.Short
                dtpTo.ShowUpDown = False
        End Select
    End Sub
    Public Sub Removealltxt()
        cbSalloan.Text = Nothing
        cbBranch.Text = Nothing
        cbCard.Text = Nothing
        cbCluster.Text = Nothing
        cbSTrans.Text = Nothing
        cbFTrans.Text = Nothing
        cbGroup.Text = Nothing
    End Sub
    Public Sub TransFailed()
        Select Case cbFTrans.Text

            Case "Salary Loan"
                Failedtrans = "SALARY LOAN"
            Case "Maternity Notification"
                Failedtrans = "MEMBER MATERNITY NOTIFICATION"
            Case "Technical Retirement"
                Failedtrans = "TECHNICAL RETIREMENT"
            Case "Annual Confirmation of Pensioner"
                Failedtrans = "ANNUAL CONFIRMATION OF PENSIONER"
            Case "Simplified Web Registration"
                Failedtrans = "SIMPLIFIED WEB REGISTRATION"
        End Select
    End Sub
    Public Sub uncheckRB()
        d.Checked = False
        w.Checked = False
        m.Checked = False
        y.Checked = False
        rbCharter.Checked = False
        rbSMAT.Checked = False
        rbSuc.Checked = False
        rbfb.Checked = False
        fbk.Checked = False
        rbSWR.Checked = False
        rbacopdep.Checked = False
        rbacop.Checked = False
        rbTOP.Checked = False
        rbFT.Checked = False
        rbGroup.Checked = False
        cbGroup.Enabled = False
        cbGroup.Text = Nothing
        cbBranch.Enabled = False
        cbCluster.Enabled = False
        rbAll.Checked = False
        rbBranch.Checked = False
        rbCluster.Checked = False
        rbDaily.Checked = False
        rbWeekly.Checked = False
        rbMonthly.Checked = False
        rbYearly.Checked = False
        rbDownTime.Checked = False
        rbQueryReports.Checked = False
        rbSummarySSS.Checked = False
        rbMonitoring.Checked = False
        rbOnlineInquiry.Checked = False
        rbDDR.Checked = False
        rbEligibility.Checked = False
        rbDDR.Checked = False
        cbFTrans.Text = Nothing
        cbSTrans.Text = Nothing
        rbFBK.Checked = False
        rbFBW.Checked = False
        rbSuccessTrans.Checked = False
        rbfailedTrans.Checked = False
        cbSTrans.Enabled = False
        cbFTrans.Enabled = False
        rbFTR.Checked = False
        rbMonitoringAll.Checked = False
        rbSucFail.Checked = False
        rbErrorLogs.Checked = False
        rdGraphInquiries.Checked = False
        rdGraphTransactions.Checked = False
        rdPinChange.Checked = False
        rdUserFeedBack.Checked = False
        rdPinChange.Checked = False
        rbPRN.Checked = False
        rdExitSurvey.Checked = False
        cboMonth.ResetText()
        cbomonthto.ResetText()
        cboYear.ResetText()
        Group = 0
        DR = 0
        Trans = 0
        OLAP = 0
        SSID = 0
    End Sub
    Public Sub RBCONTROL()
        Select Case TabControl1.SelectedTab.Text

            Case "Down Time Report"
                rbDaily.Enabled = True
                rbMonthly.Enabled = True
                rbWeekly.Enabled = True
                rbYearly.Enabled = True

            Case "Transaction Report"
                rbDaily.Enabled = False
                rbMonthly.Enabled = False
                rbWeekly.Enabled = False
                rbYearly.Enabled = False
            Case Else
                rbDaily.Enabled = False
                rbMonthly.Enabled = False
                rbWeekly.Enabled = False
                rbYearly.Enabled = False
        End Select
    End Sub
    Public Sub TransValidation()
        Select Case TabControl1.SelectedTab.Text

            Case "Down Time Report"
                If Trans = 0 Then
                    MsgBox("Please Choose Transaction Type")
                ElseIf Group = 0 Then
                    MsgBox("Please Choose Group")
                ElseIf DR = 0 Then
                    MsgBox("Please Choose Date Range")
                End If

            Case "Hits Summary"
                If rdPinChange.Checked = True Or rdUserFeedBack.Checked = True Then
                    OLAP = 1
                End If
                If OLAP = 0 Then
                    MsgBox("Please Choose Transaction Type")
                ElseIf Group = 0 Then
                    MsgBox("Please Choose Group")
                    'ElseIf dtpFrom.Value >= Today Or dtpTo.Value >= Today Then
                    '    MsgBox("Invalid date range.")
                End If

            Case "SSID Usage Report"

                If Trans = 0 Then
                    MsgBox("Please choose transaction type")
                ElseIf cbCard.Text = "" And rbSummarySSS.Checked = True Then
                    MsgBox("Please choose card type")
                ElseIf Group = 0 Then
                    MsgBox("Please Choose Group")
                Else
                    SSID = 1
                End If
            Case "Transaction Report"
                If cbSTrans.Text = "Salary Loan" And cbSalloan.Text = "" Then
                    MsgBox("Please choose member status.")
                ElseIf Trans = 0 Then
                    MsgBox("Please Choose Transaction Type")
                ElseIf Group = 0 Then
                    MsgBox("Please choose group.")
                End If
            Case "Graph"
                If rdGraphTransactions.Checked = False And rdGraphInquiries.Checked = False Then
                    MsgBox("Please select transactions or inquiries.")
                Else
                    If cboYear.Text = "" And cboMonth.Text = "" And cbomonthto.Text = "" Then
                        MsgBox("Please select year and month.")
                    End If
                End If
            Case Else
                If Trans = 0 Then
                    MsgBox("Please Choose Transaction Type")
                ElseIf Group = 0 Then
                    MsgBox("Please Choose Group")
                End If

        End Select
    End Sub
    Public Sub GetTrans()
        Select Case cbCard.Text

            Case "SSS CARD"
                TransType = "OLDSSSCARD"

            Case "UMID CARD"
                TransType = "UMIDCARD"
        End Select
    End Sub
    Public Sub fillGroup()
        Try
            cbGroup.Items.Clear()
            db.fillComboBox(db.ExecuteSQLQuery("SELECT GROUP_NM FROM SSINFOTERMGROUP acs"), cbGroup)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Public Sub FillGraphFilter()
        db.fillComboBox(db.ExecuteSQLQuery("SELECT YEAR FROM tblYear ORDER BY YEAR DESC"), cboYear)
        cboMonth.Items.Add("January")
        cboMonth.Items.Add("February")
        cboMonth.Items.Add("March")
        cboMonth.Items.Add("April")
        cboMonth.Items.Add("May")
        cboMonth.Items.Add("June")
        cboMonth.Items.Add("July")
        cboMonth.Items.Add("August")
        cboMonth.Items.Add("September")
        cboMonth.Items.Add("October")
        cboMonth.Items.Add("November")
        cboMonth.Items.Add("December")

        cbomonthto.Items.Add("January")
        cbomonthto.Items.Add("February")
        cbomonthto.Items.Add("March")
        cbomonthto.Items.Add("April")
        cbomonthto.Items.Add("May")
        cbomonthto.Items.Add("June")
        cbomonthto.Items.Add("July")
        cbomonthto.Items.Add("August")
        cbomonthto.Items.Add("September")
        cbomonthto.Items.Add("October")
        cbomonthto.Items.Add("November")
        cbomonthto.Items.Add("December")
    End Sub
    Public Sub fillCluster()
        Try
            cbCluster.Items.Clear()
            db.fillComboBox(db.ExecuteSQLQuery("Select CLSTR_NM from SSINFOTERMCLSTR acs"), cbCluster)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
    Public Sub fillBranch()
        Try
            cbBranch.Items.Clear()
            db.fillComboBox(db.ExecuteSQLQuery("Select BRANCH_NM from SSINFOTERMBR ORDER by BRANCH_NM"), cbBranch)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
#End Region
#Region "ReportsQuery"

    Private Sub OpenReportDbase()
        Dim strLogin As CrystalDecisions.Shared.TableLogOnInfo = Nothing
        Dim strConnectionInfo As New CrystalDecisions.Shared.ConnectionInfo

        strConnectionInfo.ServerName = My.Settings.db_Server
        strConnectionInfo.DatabaseName = My.Settings.db_Name
        strConnectionInfo.UserID = My.Settings.db_UName
        strConnectionInfo.Password = My.Settings.db_Pass

        For Each strTable As CrystalDecisions.CrystalReports.Engine.Table In cryRpt.Database.Tables
            strLogin = strTable.LogOnInfo
            strLogin.ConnectionInfo = strConnectionInfo
            strTable.ApplyLogOnInfo(strLogin)
        Next

        'if report have subreport/s
        For i As Short = 0 To cryRpt.Subreports.Count - 1
            Dim crDoc As ReportDocument = cryRpt.OpenSubreport(cryRpt.Subreports(i).Name)

            Dim strLogin2 As CrystalDecisions.Shared.TableLogOnInfo = Nothing
            For Each strTable2 As CrystalDecisions.CrystalReports.Engine.Table In crDoc.Database.Tables
                strLogin2 = strTable2.LogOnInfo
                strLogin2.ConnectionInfo = strConnectionInfo
                strTable2.ApplyLogOnInfo(strLogin2)
            Next
        Next
    End Sub


    Public Sub fillComboBox(ByVal dt As DataTable, ByVal cb As ComboBox)
        Try
            cb.Items.Clear()

            For row As Integer = 0 To dt.Rows.Count - 1
                cb.Items.Add(dt.Rows(row)(0))
            Next
        Catch ex As Exception
            Dim errorLogs As String = ex.ToString
            errorLogs = errorLogs.Trim
            Sql = "insert into tbl_mgmt_errorlogs values('" & My.Settings.mgmtIP & "', '" & My.Settings.mgmtID & "', '" & My.Settings.mgmtName & "', '" & My.Settings.mgmtBranch & "', '" & errorLogs _
                & "','" & "Class: Connection String" & "', '" & "Database connection error" & "', '" & Date.Today.ToShortDateString & "', '" & TimeOfDay & "') "
            ExecuteSQLQuery(Sql)
            MsgBox("Database connection error, Please contact system Administrator! ", MsgBoxStyle.Information)
        End Try
    End Sub

    Public Function ExecuteSQLQuery(ByVal SQLQuery As String) As DataTable
        Dim sqlDT As New DataTable
        Try
            Dim sqlCon As New OdbcConnection(connstring1)
            Dim sqlDA As New OdbcDataAdapter(SQLQuery, sqlCon)
            Dim sqlCB As New OdbcCommandBuilder(sqlDA)
            sqlDA.SelectCommand.CommandTimeout = 0
            sqlDA.Fill(sqlDT)
        Catch ex As Exception
            'MsgBox("Program Error: " & ex.Message, MsgBoxStyle.Critical)
            Dim errorLogs As String = ex.ToString
            errorLogs = errorLogs.Trim
            sql = "insert into tbl_mgmt_errorlogs values('" & My.Settings.mgmtIP & "', '" & My.Settings.mgmtID & "', '" & My.Settings.mgmtName & "', '" & My.Settings.mgmtBranch & "', '" & errorLogs _
                & "','" & "Class: Connection String" & "', '" & "Database connection error" & "', '" & Date.Today.ToShortDateString & "', '" & TimeOfDay & "') "
            ExecuteSQLQuery(sql)
            MsgBox("Database connection error, Please contact system Administrator! ", MsgBoxStyle.Information)
        End Try

        Return sqlDT
    End Function


    Private Sub btnSubmit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Filter.Click

        Try
            Dim day1 As String
            Dim day2 As String
            day1 = DatePart(DateInterval.Day, dtpFrom.Value)
            day2 = DatePart(DateInterval.Day, dtpTo.Value)
            Filter.Enabled = True
            TransValidation()
            cryRpt.Close()
            cryRpt.Dispose()
            GC.Collect()
            Dim date1 As Date = dtpFrom.Value
            Dim date2 As Date = dtpTo.Value

            If rdGraphTransactions.Checked = True Then
                GetName = "Transactions made through SET"
                'dt = db.ExecuteSQLQuery("SELECT SSINFOTERMCONSTAT.BRANCH_IP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMBR.BRAfNCH_NM,DATESTAMP,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0) as [ONLINE],case when(24 - ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0)) < 0 then 0 else (24 - ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0)) end as [OFFLINE] from SSINFOTERMCONSTAT inner join SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSINFOTERMCONSTAT.BRANCH_IP INNER JOIN SSINFOTERMBR on SSINFOTERMBR.BRANCH_CD = SSINFOTERMCONSTAT.BRANCH_CD INNER JOIN SSINFOTERMCLSTR on SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMCONSTAT.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSINFOTERMCONSTAT.DIVSN where DATESTAMP between '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and SSINFOTERMGROUP.GROUP_NM = '" & cbGroup.Text & "'GROUP by SSINFOTERMCONSTAT.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,DATESTAMP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM")
                Dim rpt As New GraphTransaction
                cryRpt = rpt
                cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", cboMonth.SelectedIndex + 1)
                cryRpt.SetParameterValue("@DateTo", cbomonthto.SelectedIndex + 1)
                cryRpt.SetParameterValue("@Year", cboYear.Text)
            ElseIf rdGraphInquiries.Checked = True Then
                GetName = "Inquiries made through SET"
                'dt = db.ExecuteSQLQuery("SELECT SSINFOTERMCONSTAT.BRANCH_IP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMBR.BRAfNCH_NM,DATESTAMP,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0) as [ONLINE],case when(24 - ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0)) < 0 then 0 else (24 - ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0)) end as [OFFLINE] from SSINFOTERMCONSTAT inner join SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSINFOTERMCONSTAT.BRANCH_IP INNER JOIN SSINFOTERMBR on SSINFOTERMBR.BRANCH_CD = SSINFOTERMCONSTAT.BRANCH_CD INNER JOIN SSINFOTERMCLSTR on SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMCONSTAT.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSINFOTERMCONSTAT.DIVSN where DATESTAMP between '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and SSINFOTERMGROUP.GROUP_NM = '" & cbGroup.Text & "'GROUP by SSINFOTERMCONSTAT.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,DATESTAMP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM")
                Dim rpt As New GraphInquiries
                cryRpt = rpt
                cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", cboMonth.SelectedIndex + 1)
                cryRpt.SetParameterValue("@DateTo", cbomonthto.SelectedIndex + 1)
                cryRpt.SetParameterValue("@Year", cboYear.Text)

            End If

            If rbDownTime.Checked = True And rbGroup.Checked = True And rbDaily.Checked = True Then
                GetName = "DAILY DOWNTIME REPORT BY GROUP"
                'dt = db.ExecuteSQLQuery("SELECT SSINFOTERMCONSTAT.BRANCH_IP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,DATESTAMP,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0) as [ONLINE],case when(24 - ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0)) < 0 then 0 else (24 - ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0)) end as [OFFLINE] from SSINFOTERMCONSTAT inner join SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSINFOTERMCONSTAT.BRANCH_IP INNER JOIN SSINFOTERMBR on SSINFOTERMBR.BRANCH_CD = SSINFOTERMCONSTAT.BRANCH_CD INNER JOIN SSINFOTERMCLSTR on SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMCONSTAT.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSINFOTERMCONSTAT.DIVSN where DATESTAMP between '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and SSINFOTERMGROUP.GROUP_NM = '" & cbGroup.Text & "'GROUP by SSINFOTERMCONSTAT.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,DATESTAMP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM")
                Dim rpt As New NewDowntimeDailyGroup
                cryRpt = rpt
                cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getGroup", cbGroup.Text)
                cryRpt.SetParameterValue("@dateToday", Today)
            ElseIf rbDownTime.Checked = True And rbCluster.Checked = True And rbDaily.Checked = True Then
                GetName = "DAILY DOWNTIME REPORT BY DIVISION"
                'dt = db.ExecuteSQLQuery("SELECT SSINFOTERMCONSTAT.BRANCH_IP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,DATESTAMP,SSINFOTERMCLSTR.CLSTR_NM,ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0) as [ONLINE],case when(24 - ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0)) < 0 then 0 else (24 - ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0)) end as [OFFLINE] from SSINFOTERMCONSTAT inner join SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSINFOTERMCONSTAT.BRANCH_IP INNER JOIN SSINFOTERMBR on SSINFOTERMBR.BRANCH_CD = SSINFOTERMCONSTAT.BRANCH_CD INNER JOIN SSINFOTERMCLSTR on SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMCONSTAT.CLSTR where DATESTAMP between '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and SSINFOTERMCLSTR.CLSTR_NM = '" & cbCluster.Text & "'GROUP by SSINFOTERMCONSTAT.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,DATESTAMP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMCLSTR.CLSTR_NM")
                Dim rpt As New NewDowntimeDailyCluster
                cryRpt = rpt
                'cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getCluster", cbCluster.Text)
                cryRpt.SetParameterValue("@dateToday", Today)

            ElseIf rbDownTime.Checked = True And rbDaily.Checked = True And rbAll.Checked = True Then
                GetName = "DAILY DOWNTIME REPORT SUMMARY"
                'dt = db.ExecuteSQLQuery("SELECT SSINFOTERMCONSTAT.BRANCH_IP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,DATESTAMP,SSINFOTERMCLSTR.CLSTR_NM,ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0) as [ONLINE],case when(24 - ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0)) < 0 then 0 else (24 - ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0)) end as [OFFLINE] from SSINFOTERMCONSTAT inner join SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSINFOTERMCONSTAT.BRANCH_IP INNER JOIN SSINFOTERMBR on SSINFOTERMBR.BRANCH_CD = SSINFOTERMCONSTAT.BRANCH_CD INNER JOIN SSINFOTERMCLSTR on SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMCONSTAT.CLSTR where DATESTAMP between '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' GROUP by SSINFOTERMCONSTAT.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,DATESTAMP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMCLSTR.CLSTR_NM")
                Dim rpt As New NewDowntimeDaily
                cryRpt = rpt
                ' cryRpt.Refresh()@dateToday
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@dateToday", Today)
                'GetName = "DAILY DOWNTIME REPORT SUMMARY"
                ''dt = db.ExecuteSQLQuery("SELECT SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,x.OFFLINE_DT,x.ONLINE_DT,x.DATESTAMP, CASE WHEN x.ONLINE_DT IS NULL THEN  ROUND(cast((datediff(SS, x.OFFLINE_DT, GETDATE())  / 60.0 / 60.0) as FLOAT),2) WHEN x.ONLINE_DT IS NOT NULL THEN  ROUND(cast((datediff(SS, x.OFFLINE_DT, x.ONLINE_DT)  / 60.0 / 60.0) as FLOAT),2) END AS DURATION, SUM(CASE WHEN x.ONLINE_DT IS NULL THEN  ROUND(cast((datediff(SS, y.OFFLINE_DT, GETDATE())  / 60.0 / 60.0) as FLOAT),2) WHEN x.ONLINE_DT IS NOT NULL THEN  ROUND(cast((datediff(SS, y.OFFLINE_DT, y.ONLINE_DT)  / 60.0 / 60.0) as FLOAT),2) END) AS DURATION1 FROM SSINFOTERMKIOSK INNER JOIN SSINFOTERMBR  ON SSINFOTERMKIOSK.BRANCH_CD=SSINFOTERMBR.BRANCH_CD INNER JOIN SSMONITORING x ON SSINFOTERMKIOSK.BRANCH_IP=x.BRANCH_IP JOIN SSMONITORING y on y.BRANCH_IP = x.BRANCH_IP and cast(y.OFFLINE_DT as date) = cast(x.ONLINE_DT as date) where cast(x.DATESTAMP as date) between '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' GROUP BY  SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,x.DATESTAMP,x.OFFLINE_DT,x.ONLINE_DT")
                'Dim rpt As New NewDowntime
                'cryRpt = rpt
                '' cryRpt.Refresh()
                'cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass, My.Settings.db_Host, My.Settings.db_Server)
                'rptView.ReportSource = cryRpt
                'cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                'cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
            ElseIf rbDownTime.Checked = True And rbDaily.Checked = True And rbBranch.Checked = True Then
                GetName = "DAILY DOWNTIME REPORT BY BRANCH"
                'dt = db.ExecuteSQLQuery("SELECT SSINFOTERMCONSTAT.BRANCH_IP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,DATESTAMP,SSINFOTERMCLSTR.CLSTR_NM,ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0) as [ONLINE],case when(24 - ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0)) < 0 then 0 else (24 - ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0)) end as [OFFLINE] from SSINFOTERMCONSTAT inner join SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSINFOTERMCONSTAT.BRANCH_IP INNER JOIN SSINFOTERMBR on SSINFOTERMBR.BRANCH_CD = SSINFOTERMCONSTAT.BRANCH_CD INNER JOIN SSINFOTERMCLSTR on SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMCONSTAT.CLSTR where DATESTAMP between '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and SSINFOTERMBR.BRANCH_NM = '" & cbBranch.Text & "'GROUP by SSINFOTERMCONSTAT.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,DATESTAMP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMCLSTR.CLSTR_NM")
                Dim rpt As New NewDowntimeDailyBranch
                cryRpt = rpt
                ' cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getBranch", cbBranch.Text)
                cryRpt.SetParameterValue("@dateToday", Today)

            ElseIf rbDownTime.Checked = True And rbWeekly.Checked = True And rbAll.Checked = True Then
                GetName = "WEEKLY DOWNTIME REPORT SUMMARY"
                'dt = db.ExecuteSQLQuery("SELECT SSINFOTERMCONSTAT.BRANCH_IP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,DATESTAMP,SSINFOTERMCLSTR.CLSTR_NM,ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0) as [ONLINE],case when(24 - ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0)) < 0 then 0 else (24 - ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0)) end as [OFFLINE] from SSINFOTERMCONSTAT inner join SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSINFOTERMCONSTAT.BRANCH_IP INNER JOIN SSINFOTERMBR on SSINFOTERMBR.BRANCH_CD = SSINFOTERMCONSTAT.BRANCH_CD INNER JOIN SSINFOTERMCLSTR on SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMCONSTAT.CLSTR where DATESTAMP between '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' GROUP by SSINFOTERMCONSTAT.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,DATESTAMP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMCLSTR.CLSTR_NM")
                Dim rpt As New NewDowntimeWeeklyall
                cryRpt = rpt
                '    cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@dateToday", Today)

            ElseIf rbDownTime.Checked = True And rbWeekly.Checked = True And rbBranch.Checked = True Then
                GetName = "WEEKLY DOWNTIME REPORT BY BRANCH"
                'dt = db.ExecuteSQLQuery("SELECT SSINFOTERMCONSTAT.BRANCH_IP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,DATESTAMP,SSINFOTERMCLSTR.CLSTR_NM,ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0) as [ONLINE],case when(24 - ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0)) < 0 then 0 else (24 - ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0)) end as [OFFLINE] from SSINFOTERMCONSTAT inner join SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSINFOTERMCONSTAT.BRANCH_IP INNER JOIN SSINFOTERMBR on SSINFOTERMBR.BRANCH_CD = SSINFOTERMCONSTAT.BRANCH_CD INNER JOIN SSINFOTERMCLSTR on SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMCONSTAT.CLSTR where DATESTAMP between '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and SSINFOTERMBR.BRANCH_NM = '" & cbBranch.Text & "'GROUP by SSINFOTERMCONSTAT.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,DATESTAMP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMCLSTR.CLSTR_NM")
                Dim rpt As New NewDowntimeWeeklyBranch
                cryRpt = rpt
                ' cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@getBranch", cbBranch.Text)
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@dateToday", Today)
            ElseIf rbDownTime.Checked = True And rbGroup.Checked = True And rbWeekly.Checked = True Then
                GetName = "WEEKLY DOWNTIME REPORT BY GROUP"
                'dt = db.ExecuteSQLQuery("SELECT SSINFOTERMCONSTAT.BRANCH_IP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,DATESTAMP,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0) as [ONLINE],case when(24 - ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0)) < 0 then 0 else (24 - ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0)) end as [OFFLINE] from SSINFOTERMCONSTAT inner join SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSINFOTERMCONSTAT.BRANCH_IP INNER JOIN SSINFOTERMBR on SSINFOTERMBR.BRANCH_CD = SSINFOTERMCONSTAT.BRANCH_CD INNER JOIN SSINFOTERMCLSTR on SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMCONSTAT.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSINFOTERMCONSTAT.DIVSN where DATESTAMP between '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and SSINFOTERMGROUP.GROUP_NM = '" & cbGroup.Text & "'GROUP by SSINFOTERMCONSTAT.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,DATESTAMP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM")
                Dim rpt As New NewDowntimeWeeklyGroup
                cryRpt = rpt
                cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@getGroup", cbGroup.Text)
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@dateToday", Today)
            ElseIf rbDownTime.Checked = True And rbWeekly.Checked = True And rbCluster.Checked = True Then
                GetName = "WEEKLY DOWNTIME REPORT BY DIVISION"
                'dt = db.ExecuteSQLQuery("SELECT SSINFOTERMCONSTAT.BRANCH_IP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,DATESTAMP,SSINFOTERMCLSTR.CLSTR_NM,ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0) as [ONLINE],case when(24 - ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0)) < 0 then 0 else (24 - ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0)) end as [OFFLINE] from SSINFOTERMCONSTAT inner join SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSINFOTERMCONSTAT.BRANCH_IP INNER JOIN SSINFOTERMBR on SSINFOTERMBR.BRANCH_CD = SSINFOTERMCONSTAT.BRANCH_CD INNER JOIN SSINFOTERMCLSTR on SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMCONSTAT.CLSTR where DATESTAMP between '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and SSINFOTERMCLSTR.CLSTR_NM = '" & cbCluster.Text & "'GROUP by SSINFOTERMCONSTAT.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,DATESTAMP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMCLSTR.CLSTR_NM")
                Dim rpt As New NewDowntimeWeeklyCluster
                cryRpt = rpt
                '  cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@getCluster", cbCluster.Text)
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@dateToday", Today)
            ElseIf rbDownTime.Checked = True And rbMonthly.Checked = True And rbGroup.Checked = True Then
                getMonth()
                GetName = "MONTHLY DOWNTIME REPORT BY GROUP"
                'dt = db.ExecuteSQLQuery("SELECT SSINFOTERMCONSTAT.BRANCH_IP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,DATESTAMP,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0) as [ONLINE],case when(24 - ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0)) < 0 then 0 else (24 - ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0)) end as [OFFLINE] from SSINFOTERMCONSTAT inner join SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSINFOTERMCONSTAT.BRANCH_IP INNER JOIN SSINFOTERMBR on SSINFOTERMBR.BRANCH_CD = SSINFOTERMCONSTAT.BRANCH_CD INNER JOIN SSINFOTERMCLSTR on SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMCONSTAT.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSINFOTERMCONSTAT.DIVSN where DATESTAMP between '" & date12 & "' and '" & date22 & "' and SSINFOTERMGROUP.GROUP_NM = '" & cbGroup.Text & "'GROUP by SSINFOTERMCONSTAT.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,DATESTAMP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM")
                Dim rpt As New NewDowntimeMonthlyGroup
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@getGroup", cbGroup.Text)
                cryRpt.SetParameterValue("@DateFrom", date12)
                cryRpt.SetParameterValue("@DateTo", date22)
                cryRpt.SetParameterValue("@dateToday", Today)
            ElseIf rbDownTime.Checked = True And rbMonthly.Checked = True And rbCluster.Checked = True Then
                getMonth()
                GetName = "MONTHLY DOWNTIME REPORT BY DIVISION"
                'dt = db.ExecuteSQLQuery("SELECT SSINFOTERMCONSTAT.BRANCH_IP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,DATESTAMP,SSINFOTERMCLSTR.CLSTR_NM,ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0) as [ONLINE],case when(24 - ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0)) < 0 then 0 else (24 - ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0)) end as [OFFLINE] from SSINFOTERMCONSTAT inner join SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSINFOTERMCONSTAT.BRANCH_IP INNER JOIN SSINFOTERMBR on SSINFOTERMBR.BRANCH_CD = SSINFOTERMCONSTAT.BRANCH_CD INNER JOIN SSINFOTERMCLSTR on SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMCONSTAT.CLSTR where DATESTAMP between '" & date12 & "' and '" & date22 & "' and SSINFOTERMCLSTR.CLSTR_NM = '" & cbCluster.Text & "'GROUP by SSINFOTERMCONSTAT.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,DATESTAMP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMCLSTR.CLSTR_NM")
                Dim rpt As New NewDowntimeMonthlyCluster
                cryRpt = rpt
                '  cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@getCluster", cbCluster.Text)
                cryRpt.SetParameterValue("@DateFrom", date12)
                cryRpt.SetParameterValue("@DateTo", date22)
                cryRpt.SetParameterValue("@dateToday", Today)
            ElseIf rbDownTime.Checked = True And rbMonthly.Checked = True And rbBranch.Checked = True Then
                getMonth()
                GetName = "MONTHLY DOWNTIME REPORT BY BRANCH"
                'dt = db.ExecuteSQLQuery("SELECT SSINFOTERMCONSTAT.BRANCH_IP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,DATESTAMP,SSINFOTERMCLSTR.CLSTR_NM,ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0) as [ONLINE],case when(24 - ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0)) < 0 then 0 else (24 - ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0)) end as [OFFLINE] from SSINFOTERMCONSTAT inner join SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSINFOTERMCONSTAT.BRANCH_IP INNER JOIN SSINFOTERMBR on SSINFOTERMBR.BRANCH_CD = SSINFOTERMCONSTAT.BRANCH_CD INNER JOIN SSINFOTERMCLSTR on SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMCONSTAT.CLSTR where DATESTAMP between '" & date12 & "' and '" & date22 & "' and SSINFOTERMBR.BRANCH_NM = '" & cbBranch.Text & "'GROUP by SSINFOTERMCONSTAT.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,DATESTAMP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMCLSTR.CLSTR_NM")
                Dim rpt As New NewDowntimeMonthlyBranch
                cryRpt = rpt
                '  cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@getBranch", cbBranch.Text)
                cryRpt.SetParameterValue("@DateFrom", date12)
                cryRpt.SetParameterValue("@DateTo", date22)
                cryRpt.SetParameterValue("@dateToday", Today)

            ElseIf rbDownTime.Checked = True And rbMonthly.Checked = True And rbAll.Checked = True Then
                getMonth()
                GetName = "MONTHLY DOWNTIME REPORT SUMMARY"
                'dt = db.ExecuteSQLQuery("SELECT SSINFOTERMCONSTAT.BRANCH_IP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,DATESTAMP,SSINFOTERMCLSTR.CLSTR_NM,ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0) as [ONLINE],case when(24 - ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0)) < 0 then 0 else (24 - ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0)) end as [OFFLINE] from SSINFOTERMCONSTAT inner join SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSINFOTERMCONSTAT.BRANCH_IP INNER JOIN SSINFOTERMBR on SSINFOTERMBR.BRANCH_CD = SSINFOTERMCONSTAT.BRANCH_CD INNER JOIN SSINFOTERMCLSTR on SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMCONSTAT.CLSTR where DATESTAMP between '" & date12 & "' and '" & date22 & "' GROUP by SSINFOTERMCONSTAT.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,DATESTAMP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMCLSTR.CLSTR_NM")
                Dim rpt As New NewDowntimeMonthlyAll
                cryRpt = rpt
                '  cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date12)
                cryRpt.SetParameterValue("@DateTo", date22)
                cryRpt.SetParameterValue("@dateToday", Today)
            ElseIf rbYearly.Checked = True And rbDownTime.Checked = True And rbAll.Checked = True Then
                GetName = "MONTHLY DOWNTIME REPORT SUMMARY"
                'dt = db.ExecuteSQLQuery("SELECT SSINFOTERMCONSTAT.BRANCH_IP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,DATESTAMP,SSINFOTERMCLSTR.CLSTR_NM,ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0) as [ONLINE],case when(24 - ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0)) < 0 then 0 else (24 - ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0)) end as [OFFLINE] from SSINFOTERMCONSTAT inner join SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSINFOTERMCONSTAT.BRANCH_IP INNER JOIN SSINFOTERMBR on SSINFOTERMBR.BRANCH_CD = SSINFOTERMCONSTAT.BRANCH_CD INNER JOIN SSINFOTERMCLSTR on SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMCONSTAT.CLSTR where DATESTAMP between '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' GROUP by SSINFOTERMCONSTAT.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,DATESTAMP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMCLSTR.CLSTR_NM")
                Dim rpt As New NewDowntimeYearlyall
                cryRpt = rpt
                '  cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", dtpFrom.Value)
                cryRpt.SetParameterValue("@DateTo", dtpTo.Value)
                cryRpt.SetParameterValue("@dateToday", Today)
            ElseIf rbYearly.Checked = True And rbGroup.Checked = True And rbGroup.Checked = True Then
                GetName = "MONTHLY DOWNTIME REPORT BY GROUP"
                'dt = db.ExecuteSQLQuery("SELECT SSINFOTERMCONSTAT.BRANCH_IP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,DATESTAMP,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0) as [ONLINE],case when(24 - ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0)) < 0 then 0 else (24 - ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0)) end as [OFFLINE] from SSINFOTERMCONSTAT inner join SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSINFOTERMCONSTAT.BRANCH_IP INNER JOIN SSINFOTERMBR on SSINFOTERMBR.BRANCH_CD = SSINFOTERMCONSTAT.BRANCH_CD INNER JOIN SSINFOTERMCLSTR on SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMCONSTAT.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSINFOTERMCONSTAT.DIVSN where DATESTAMP between '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and SSINFOTERMGROUP.GROUP_NM = '" & cbGroup.Text & "'GROUP by SSINFOTERMCONSTAT.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,DATESTAMP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM")
                Dim rpt As New NewDowntimeYearlyGroup
                cryRpt = rpt
                '  cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@getGroup", cbGroup.Text)
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@dateToday", Today)
            ElseIf rbYearly.Checked = True And rbDownTime.Checked = True And rbCluster.Checked = True Then
                GetName = "MONTHLY DOWNTIME REPORT BY DIVISION"
                'dt = db.ExecuteSQLQuery("SELECT SSINFOTERMCONSTAT.BRANCH_IP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,DATESTAMP,SSINFOTERMCLSTR.CLSTR_NM,ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0) as [ONLINE],case when(24 - ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0)) < 0 then 0 else (24 - ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0)) end as [OFFLINE] from SSINFOTERMCONSTAT inner join SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSINFOTERMCONSTAT.BRANCH_IP INNER JOIN SSINFOTERMBR on SSINFOTERMBR.BRANCH_CD = SSINFOTERMCONSTAT.BRANCH_CD INNER JOIN SSINFOTERMCLSTR on SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMCONSTAT.CLSTR where DATESTAMP between '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and SSINFOTERMCLSTR.CLSTR_NM = '" & cbCluster.Text & "'GROUP by SSINFOTERMCONSTAT.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,DATESTAMP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMCLSTR.CLSTR_NM")
                Dim rpt As New NewDowntimeYearlyCluster
                cryRpt = rpt
                '  cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@getCluster", cbCluster.Text)
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@dateToday", Today)
            ElseIf rbDownTime.Checked = True And rbYearly.Checked = True And rbBranch.Checked = True Then
                GetName = "MONTHLY DOWNTIME REPORT BY BRANCH"
                'dt = db.ExecuteSQLQuery("SELECT SSINFOTERMCONSTAT.BRANCH_IP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,DATESTAMP,SSINFOTERMCLSTR.CLSTR_NM,ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0) as [ONLINE],case when(24 - ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0)) < 0 then 0 else (24 - ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0)) end as [OFFLINE] from SSINFOTERMCONSTAT inner join SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSINFOTERMCONSTAT.BRANCH_IP INNER JOIN SSINFOTERMBR on SSINFOTERMBR.BRANCH_CD = SSINFOTERMCONSTAT.BRANCH_CD INNER JOIN SSINFOTERMCLSTR on SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMCONSTAT.CLSTR where DATESTAMP between '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and SSINFOTERMBR.BRANCH_NM = '" & cbBranch.Text & "'GROUP by SSINFOTERMCONSTAT.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,DATESTAMP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMCLSTR.CLSTR_NM")
                Dim rpt As New NewDowntimeYearlyBranch
                cryRpt = rpt
                '  cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@getBranch", cbBranch.Text)
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@dateToday", Today)
            ElseIf rbMonitoring.Checked = True And rbBranch.Checked = True Then
                GetName = "MONITORING REPORT BY BRANCH"
                ''dt = db.ExecuteSQLQuery("select t1.BRANCH_IP,cast(t1.DATESTAMP as date)[Date],SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,t1.ONLINE_TME as [TIME],case when t1.ONLINE_TME <> '' or t1.OFFLINE_TME = '' then 'ONLINE' else 'OFFLINE' end as status,dbo.getinseconds(CAST(NULLIF(t1.MSG,'') AS decimal(13,2)) * 60) as 'Duration' from SSINFOTERMCONSTAT t1 INNER JOIN SSINFOTERMBR on SSINFOTERMBR.BRANCH_CD = t1.BRANCH_CD INNER JOIN SSINFOTERMCLSTR  on SSINFOTERMCLSTR.CLSTR_CD = t1.CLSTR  where t1.ONLINE_TME = t1.ONLINE_TME and t1.DATESTAMP BETWEEN  '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and SSINFOTERMBR.BRANCH_NM = '" & cbBranch.Text & "' GROUP BY t1.BRANCH_CD,t1.BRANCH_IP,t1.DATESTAMP,t1.BRANCH_CD,t1.OFFLINE_TME,t1.ONLINE_TME,t1.MSG,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM UNION select t3.BRANCH_IP,cast(t3.DATESTAMP as date)[Date],SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,t3.OFFLINE_TME, case when t3.ONLINE_TME <> '' or t3.OFFLINE_TME = '' then 'ONLINE' else 'OFFLINE' end as status,dbo.getinseconds(CAST(NULLIF(t3.MSG,'') AS decimal(13,2)) * 60) as 'Duration' from SSINFOTERMCONSTAT t3 INNER JOIN SSINFOTERMBR on SSINFOTERMBR.BRANCH_CD = t3.BRANCH_CD INNER JOIN SSINFOTERMCLSTR on SSINFOTERMCLSTR.CLSTR_CD = t3.CLSTR  where t3.OFFLINE_TME = t3.OFFLINE_TME and t3.DATESTAMP BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and SSINFOTERMBR.BRANCH_NM = '" & cbBranch.Text & "' GROUP BY t3.BRANCH_CD,t3.BRANCH_IP,t3.DATESTAMP,t3.BRANCH_CD,t3.OFFLINE_TME,t3.ONLINE_TME,t3.MSG,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM")
                'Dim rpt As New MonitoringBranch
                'dt = db.ExecuteSQLQuery("select SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,unpvt.BRANCH_IP,unpvt.value,case when unpvt.attribute = 'OFFLINE_DT' then '*** OFFLINE ***' else 'ONLINE' end as [STATUS] FROM SSMONITORING c UNPIVOT ( value for attribute in (ONLINE_DT,OFFLINE_DT) )unpvt INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = unpvt.BRANCH_CD INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.BRANCH_IP = unpvt.BRANCH_IP INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMKIOSK.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSINFOTERMKIOSK.DIVSN where cast(unpvt.[value] as date) BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and SSINFOTERMBR.BRANCH_NM = '" & cbBranch.Text & "' UNION  SELECT SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM, SSINFOTERMKIOSK.BRANCH_IP, GETDATE() AS VALUE, 'ONLINE' AS STATUS  FROM SSINFOTERMKIOSK INNER JOIN SSMONITORING ON SSMONITORING.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSINFOTERMKIOSK.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMKIOSK.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSINFOTERMKIOSK.DIVSN WHERE SSINFOTERMKIOSK.BRANCH_IP NOT IN(SELECT SSMONITORING.BRANCH_IP FROM SSMONITORING) and cast(SSMONITORING.DATESTAMP as date) BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and SSINFOTERMBR.BRANCH_NM = '" & cbBranch.Text & "' ORDER BY BRANCH_NM")
                Dim rpt As New NewMonitoringBranch
                cryRpt = rpt
                ' cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@getBranch", cbBranch.Text)
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@dateToday", Today)
            ElseIf rbMonitoring.Checked = True And rbGroup.Checked = True Then
                GetName = "MONITORING REPORT BY GROUP"
                ''dt = db.ExecuteSQLQuery("select t1.BRANCH_IP,cast(t1.DATESTAMP as date)[Date],SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,t1.ONLINE_TME as [TIME], case when t1.ONLINE_TME <> '' or t1.OFFLINE_TME = '' then 'ONLINE' else 'OFFLINE' end as status,dbo.getinseconds(CAST(NULLIF(t1.MSG,'') AS decimal(13,2)) * 60) as 'Duration' from SSINFOTERMCONSTAT t1 INNER JOIN SSINFOTERMBR on SSINFOTERMBR.BRANCH_CD = t1.BRANCH_CD INNER JOIN SSINFOTERMCLSTR  on SSINFOTERMCLSTR.CLSTR_CD = t1.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = t1.DIVSN where t1.ONLINE_TME = t1.ONLINE_TME and t1.DATESTAMP BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and SSINFOTERMGROUP.GROUP_NM = '" & cbGroup.Text & "' GROUP BY t1.BRANCH_CD,t1.BRANCH_IP,t1.DATESTAMP,t1.BRANCH_CD,t1.OFFLINE_TME,t1.ONLINE_TME,t1.MSG,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM UNION select t3.BRANCH_IP,cast(t3.DATESTAMP as date)[Date],SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,t3.OFFLINE_TME, case when t3.ONLINE_TME <> '' or t3.OFFLINE_TME = '' then 'ONLINE' else 'OFFLINE' end as status,dbo.getinseconds(CAST(NULLIF(t3.MSG,'') AS decimal(13,2)) * 60) as 'Duration' from SSINFOTERMCONSTAT t3 INNER JOIN SSINFOTERMBR  on SSINFOTERMBR.BRANCH_CD = t3.BRANCH_CD INNER JOIN SSINFOTERMCLSTR on SSINFOTERMCLSTR.CLSTR_CD = t3.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = t3.DIVSN where t3.OFFLINE_TME = t3.OFFLINE_TME and t3.DATESTAMP BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and SSINFOTERMGROUP.GROUP_NM = '" & cbGroup.Text & "' GROUP BY t3.BRANCH_CD,t3.BRANCH_IP,t3.DATESTAMP,t3.BRANCH_CD,t3.OFFLINE_TME,t3.ONLINE_TME,t3.MSG,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM")
                'Dim rpt As New MonitoringGroup
                'dt = db.ExecuteSQLQuery("select SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,unpvt.BRANCH_IP,unpvt.value,case when unpvt.attribute = 'OFFLINE_DT' then '*** OFFLINE ***' else 'ONLINE' end as [STATUS] FROM SSMONITORING c UNPIVOT ( value for attribute in (ONLINE_DT,OFFLINE_DT) )unpvt INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = unpvt.BRANCH_CD INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.BRANCH_IP = unpvt.BRANCH_IP INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMKIOSK.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSINFOTERMKIOSK.DIVSN where cast(unpvt.[value] as date) BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and SSINFOTERMGROUP.GROUP_NM = '" & cbGroup.Text & "' UNION  SELECT SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM, SSINFOTERMKIOSK.BRANCH_IP, GETDATE() AS VALUE, 'ONLINE' AS STATUS  FROM SSINFOTERMKIOSK INNER JOIN SSMONITORING ON SSMONITORING.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSINFOTERMKIOSK.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMKIOSK.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSINFOTERMKIOSK.DIVSN WHERE SSINFOTERMKIOSK.BRANCH_IP NOT IN(SELECT SSMONITORING.BRANCH_IP FROM SSMONITORING) and cast(SSMONITORING.DATESTAMP as date) BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and SSINFOTERMGROUP.GROUP_NM = '" & cbGroup.Text & "' ORDER BY BRANCH_NM")
                Dim rpt As New NewMonitoringGroup
                cryRpt = rpt
                ' cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@getGroup", cbGroup.Text)
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@dateToday", Today)
            ElseIf rbMonitoring.Checked = True And rbCluster.Checked = True Then
                GetName = "MONITORING REPORT BY DIVISION"
                ''dt = db.ExecuteSQLQuery("select t1.BRANCH_IP,cast(t1.DATESTAMP as date)[Date],SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,t1.ONLINE_TME as [TIME], case when t1.ONLINE_TME <> '' or t1.OFFLINE_TME = '' then 'ONLINE' else 'OFFLINE' end as status,dbo.getinseconds(CAST(NULLIF(t1.MSG,'') AS decimal(13,2)) * 60) as 'Duration' from SSINFOTERMCONSTAT t1 INNER JOIN SSINFOTERMBR on SSINFOTERMBR.BRANCH_CD = t1.BRANCH_CD INNER JOIN SSINFOTERMCLSTR on SSINFOTERMCLSTR.CLSTR_CD = t1.CLSTR  where t1.ONLINE_TME = t1.ONLINE_TME and t1.DATESTAMP BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "'  and SSINFOTERMCLSTR.CLSTR_NM = '" & cbCluster.Text & "' GROUP BY t1.BRANCH_CD,t1.BRANCH_IP,t1.DATESTAMP,t1.BRANCH_CD,t1.OFFLINE_TME,t1.ONLINE_TME,t1.MSG,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM UNION select t3.BRANCH_IP,cast(t3.DATESTAMP as date)[Date],SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,t3.OFFLINE_TME, case when t3.ONLINE_TME <> '' or t3.OFFLINE_TME = '' then 'ONLINE' else 'OFFLINE' end as status,dbo.getinseconds(CAST(NULLIF(t3.MSG,'') AS decimal(13,2)) * 60) as 'Duration' from SSINFOTERMCONSTAT t3 INNER JOIN SSINFOTERMBR on SSINFOTERMBR.BRANCH_CD = t3.BRANCH_CD INNER JOIN SSINFOTERMCLSTR on SSINFOTERMCLSTR.CLSTR_CD = t3.CLSTR where t3.OFFLINE_TME = t3.OFFLINE_TME and t3.DATESTAMP BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "'  and SSINFOTERMCLSTR.CLSTR_NM = '" & cbCluster.Text & "' GROUP BY t3.BRANCH_CD,t3.BRANCH_IP,t3.DATESTAMP,t3.BRANCH_CD,t3.OFFLINE_TME,t3.ONLINE_TME,t3.MSG,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM")
                'Dim rpt As New KioskStatusAll
                'dt = db.ExecuteSQLQuery("select SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,unpvt.BRANCH_IP,unpvt.value,case when unpvt.attribute = 'OFFLINE_DT' then '*** OFFLINE ***' else 'ONLINE' end as [STATUS] FROM SSMONITORING c UNPIVOT ( value for attribute in (ONLINE_DT,OFFLINE_DT) )unpvt INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = unpvt.BRANCH_CD INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.BRANCH_IP = unpvt.BRANCH_IP INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMKIOSK.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSINFOTERMKIOSK.DIVSN where cast(unpvt.[value] as date) BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and SSINFOTERMCLSTR.CLSTR_NM = '" & cbCluster.Text & "' UNION  SELECT SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM, SSINFOTERMKIOSK.BRANCH_IP, GETDATE() AS VALUE, 'ONLINE' AS STATUS  FROM SSINFOTERMKIOSK INNER JOIN SSMONITORING ON SSMONITORING.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSINFOTERMKIOSK.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMKIOSK.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSINFOTERMKIOSK.DIVSN WHERE SSINFOTERMKIOSK.BRANCH_IP NOT IN(SELECT SSMONITORING.BRANCH_IP FROM SSMONITORING) and cast(SSMONITORING.DATESTAMP as date) BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and SSINFOTERMCLSTR.CLSTR_NM = '" & cbCluster.Text & "' ORDER BY BRANCH_NM")
                Dim rpt As New NewMonitoringCluster
                cryRpt = rpt
                ' cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@getCluster", cbCluster.Text)
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@dateToday", Today)
            ElseIf rbMonitoring.Checked = True And rbAll.Checked = True Then
                GetName = "MONITORING REPORT SUMMARY"
                ''dt = db.ExecuteSQLQuery("select t1.BRANCH_IP,cast(t1.DATESTAMP as date)[Date],SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,t1.ONLINE_TME as [TIME], case when t1.ONLINE_TME <> '' or t1.OFFLINE_TME = '' then 'ONLINE' else 'OFFLINE' end as status,dbo.getinseconds(CAST(NULLIF(t1.MSG,'') AS decimal(13,2)) * 60) as 'Duration' from SSINFOTERMCONSTAT t1 INNER JOIN SSINFOTERMBR  on SSINFOTERMBR.BRANCH_CD = t1.BRANCH_CD INNER JOIN SSINFOTERMCLSTR  on SSINFOTERMCLSTR.CLSTR_CD = t1.CLSTR  where t1.ONLINE_TME = t1.ONLINE_TME and t1.DATESTAMP BETWEEN '" & date1.ToShortDateString & "' AND '" & date2.ToShortDateString & "' GROUP BY t1.BRANCH_CD,t1.BRANCH_IP,t1.DATESTAMP,t1.BRANCH_CD,t1.ONLINE_TME,t1.OFFLINE_TME,t1.MSG,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM UNION select t3.BRANCH_IP,cast(t3.DATESTAMP as date)[Date],SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,t3.OFFLINE_TME, case when t3.ONLINE_TME <> '' or t3.OFFLINE_TME = '' then 'ONLINE' else 'OFFLINE' end as status,dbo.getinseconds(CAST(NULLIF(t3.MSG,'') AS decimal(13,2)) * 60) as 'Duration' from SSINFOTERMCONSTAT t3 INNER JOIN SSINFOTERMBR  on SSINFOTERMBR.BRANCH_CD = t3.BRANCH_CD  INNER JOIN SSINFOTERMCLSTR on SSINFOTERMCLSTR.CLSTR_CD = t3.CLSTR  where t3.OFFLINE_TME = t3.OFFLINE_TME and t3.DATESTAMP BETWEEN '" & date1.ToShortDateString & "' AND '" & date2.ToShortDateString & "' GROUP BY t3.BRANCH_CD,t3.BRANCH_IP,t3.DATESTAMP,t3.BRANCH_CD,t3.OFFLINE_TME,t3.ONLINE_TME,t3.MSG,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM")
                'Dim rpt As New Monitoring
                'dt = db.ExecuteSQLQuery("select SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,unpvt.BRANCH_IP,unpvt.value,case when unpvt.attribute = 'OFFLINE_DT' then '*** OFFLINE ***' else 'ONLINE' end as [STATUS] FROM SSMONITORING c UNPIVOT ( value for attribute in (ONLINE_DT,OFFLINE_DT) )unpvt INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = unpvt.BRANCH_CD INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.BRANCH_IP = unpvt.BRANCH_IP INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMKIOSK.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSINFOTERMKIOSK.DIVSN where cast(unpvt.[value] as date) BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' UNION  SELECT SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM, SSINFOTERMKIOSK.BRANCH_IP, GETDATE() AS VALUE, 'ONLINE' AS STATUS  FROM SSINFOTERMKIOSK INNER JOIN SSMONITORING ON SSMONITORING.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSINFOTERMKIOSK.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMKIOSK.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSINFOTERMKIOSK.DIVSN WHERE SSINFOTERMKIOSK.BRANCH_IP NOT IN(SELECT SSMONITORING.BRANCH_IP FROM SSMONITORING where  cast(SSMONITORING.DATESTAMP as date) = '" & Today & "') and cast(SSMONITORING.DATESTAMP as date) BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' ORDER BY BRANCH_NM")
                Dim rpt As New NewMonitoringAll
                cryRpt = rpt
                '    cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@dateToday", Today)
            ElseIf rbTOP.Checked = True And rbGroup.Checked = True Then
                GetName = "TOP HITS REPORT BY GROUP"
                ''dt = db.ExecuteSQLQuery("SELECT convert(varchar,SSTRANSAT.ENCODE_DT,101) as DateOfTrans,case when SSINFOTERMPROCCD.PROC_NM = 'SSID CLEARANCE' then 'SSS/UMID CARD INFO' else SSINFOTERMPROCCD.PROC_NM end as 'PROC',SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN, COUNT(SSINFOTERMPROCCD.PROC_NM) as [count],(select SSINFOTERMPROCCD.PROC_NM where SSINFOTERMPROCCD.PROC_NM not LIKE '%MY.SSS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOGIN%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%REGISTRATION%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBERSHIP%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOANS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFITS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%CORPORATE PROFILE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PUBLICATIONS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%OTHER SERVICES%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%HOME%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - SS SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - EC SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%EC - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%CHECKLIST OF DOCUMENTS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER MATERNITY NOTIFICATION%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ANNUAL CONFIRMATION OF PENSIONER%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ONLINE INQUIRY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SALARY LOAN%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%TOTAL DISABILITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PARTIAL DISABILITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PENSION MAINTENANCE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%KIOSK FEEDBACK%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%WEBSITE FEEDBACK%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS FUNERAL%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%TECHNICAL RETIREMENT%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%DDR FUNERAL%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SICKNESS/MATERNITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOAN ELIGIBILITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%DDR FUNERAL%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - MATERNITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%CITIZENS CHARTER%') as [ONLINE INQUIRY], (SELECT SSINFOTERMPROCCD.PROC_NM where SSINFOTERMPROCCD.PROC_NM not LIKE 'WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MY.SSS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOGIN%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%REGISTRATION%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBERSHIP%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOANS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFITS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%CORPORATE PROFILE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PUBLICATIONS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%OTHER SERVICES%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%HOME%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - SS SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - EC SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%EC - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER MATERNITY NOTIFICATION%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ANNUAL CONFIRMATION OF PENSIONER%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ONLINE INQUIRY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SALARY LOAN%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%TOTAL DISABILITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PARTIAL DISABILITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PENSION MAINTENANCE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%KIOSK FEEDBACK%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%WEBSITE FEEDBACK%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS FUNERAL%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ACTUAL PREMIUMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%EMPLOYMENT HISTORY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%FLEXI-FUND%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOAN STATUS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MATERNITY CLAIMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER DETAILS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SICKNESS CLAIMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%TECHNICAL RETIREMENT%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSID CLEARANCE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFIT CLAIMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - MATERNITY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%DEATH BENEFITS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%CITIZENS CHARTER%' )AS [ELIGIBILITY], (select SSINFOTERMPROCCD.PROC_NM where SSINFOTERMPROCCD.PROC_NM not LIKE 'WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MY.SSS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOGIN%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%REGISTRATION%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBERSHIP%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOANS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFITS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%CORPORATE PROFILE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PUBLICATIONS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%OTHER SERVICES%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%HOME%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - SS SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - EC SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%EC - SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CHECKLIST OF DOCUMENTS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER MATERNITY NOTIFICATION%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ANNUAL CONFIRMATION OF PENSIONER%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ONLINE INQUIRY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SALARY LOAN%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%DDR FUNERAL%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SICKNESS/MATERNITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PENSION MAINTENANCE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%KIOSK FEEDBACK%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%WEBSITE FEEDBACK%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOAN ELIGIBILITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ACTUAL PREMIUMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%EMPLOYMENT HISTORY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%FLEXI-FUND%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOAN STATUS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MATERNITY CLAIMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER DETAILS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SICKNESS CLAIMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%TECHNICAL RETIREMENT%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSID CLEARANCE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFIT CLAIMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - MATERNITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%CITIZENS CHARTER%')as [DDR FUNERAL], (select SSINFOTERMPROCCD.PROC_NM where SSTRANSAT.PROC_CD not like '%10001%'and SSTRANSAT.PROC_CD not like '%10002%'and SSTRANSAT.PROC_CD not like '%10003%'and SSTRANSAT.PROC_CD not like '%10004%'and SSTRANSAT.PROC_CD not like '%10005%'and SSTRANSAT.PROC_CD not like '%10006%'and SSTRANSAT.PROC_CD not like '%10007%'and SSTRANSAT.PROC_CD not like '%10008%'and SSTRANSAT.PROC_CD not like '%10009%'and SSTRANSAT.PROC_CD not like '%10010%'and SSTRANSAT.PROC_CD not like '%10011%'and SSTRANSAT.PROC_CD not like '%10012%'and SSTRANSAT.PROC_CD not like '%10013%'and SSTRANSAT.PROC_CD not like '%10014%'and SSTRANSAT.PROC_CD not like '%10015%'and SSTRANSAT.PROC_CD not like '%10016%'and SSTRANSAT.PROC_CD not like '%10017%'and SSTRANSAT.PROC_CD not like '%10018%'and SSTRANSAT.PROC_CD not like '%10019%'and SSTRANSAT.PROC_CD not like '%10020%'and SSTRANSAT.PROC_CD not like '%10021%'and SSTRANSAT.PROC_CD not like '%10022%'and SSTRANSAT.PROC_CD not like '%10023%'and SSTRANSAT.PROC_CD not like '%10024%'and SSTRANSAT.PROC_CD not like '%10025%'and SSTRANSAT.PROC_CD not like '%10026%'and SSTRANSAT.PROC_CD not like '%10027%'and SSTRANSAT.PROC_CD not like '%10028%'and SSTRANSAT.PROC_CD not like '%10029%'and SSTRANSAT.PROC_CD not like '%10030%'and SSTRANSAT.PROC_CD not like '%10031%'and SSTRANSAT.PROC_CD not like '%10032%'and SSTRANSAT.PROC_CD not like '%10033%'and SSTRANSAT.PROC_CD not like '%10034%'and SSTRANSAT.PROC_CD not like '%10035%'and SSTRANSAT.PROC_CD not like '%10036%'and SSTRANSAT.PROC_CD not like '%10041%'and SSTRANSAT.PROC_CD not like '%10042%') as [SICKNESS/MATERNITY]	FROM SSTRANSAT INNER JOIN SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAT.BRANCH_IP INNER JOIN SSINFOTERMPROCCD ON SSINFOTERMPROCCD.PROC_CD = SSTRANSAT.PROC_CD where SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and SSINFOTERMGROUP.GROUP_NM = '" & cbGroup.Text & "' and SSINFOTERMPROCCD.PROC_NM not LIKE 'WEBSITE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MY.SSS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOGIN%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%REGISTRATION%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBERSHIP%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOANS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFITS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CORPORATE PROFILE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%PUBLICATIONS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%OTHER SERVICES%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%HOME%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CITIZENS CHARTER%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%ONLINE INQUIRY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SALARY LOAN%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%TECHNICAL RETIREMENT%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%ANNUAL CONFIRMATION OF PENSIONER%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER MATERNITY NOTIFICATION%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%PENSION MAINTENANCE%' GROUP by SSTRANSAT.ENCODE_DT,SSINFOTERMPROCCD.PROC_NM,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN,SSTRANSAT.PROC_CD")
                Dim rpt As New TOTALALLGROUP
                cryRpt = rpt
                '   cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@getGroup", cbGroup.Text)
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
            ElseIf rbTOP.Checked = True And rbBranch.Checked = True Then
                '    GetName = "TOP HITS REPORT BY BRANCH"
                ''dt = db.ExecuteSQLQuery("SELECT convert(varchar,SSTRANSAT.ENCODE_DT,101) as DateOfTrans,case when SSINFOTERMPROCCD.PROC_NM = 'SSID CLEARANCE' then 'SSS/UMID CARD INFO' else SSINFOTERMPROCCD.PROC_NM end as 'PROC',SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN, COUNT(SSINFOTERMPROCCD.PROC_NM) as [count],(select SSINFOTERMPROCCD.PROC_NM where SSINFOTERMPROCCD.PROC_NM not LIKE '%MY.SSS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOGIN%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%REGISTRATION%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBERSHIP%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOANS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFITS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%CORPORATE PROFILE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PUBLICATIONS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%OTHER SERVICES%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%HOME%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - SS SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - EC SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%EC - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%CHECKLIST OF DOCUMENTS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER MATERNITY NOTIFICATION%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ANNUAL CONFIRMATION OF PENSIONER%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ONLINE INQUIRY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SALARY LOAN%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%TOTAL DISABILITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PARTIAL DISABILITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PENSION MAINTENANCE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%KIOSK FEEDBACK%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%WEBSITE FEEDBACK%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS FUNERAL%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%TECHNICAL RETIREMENT%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%DDR FUNERAL%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SICKNESS/MATERNITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOAN ELIGIBILITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%DDR FUNERAL%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - MATERNITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%CITIZENS CHARTER%') as [ONLINE INQUIRY], (SELECT SSINFOTERMPROCCD.PROC_NM where SSINFOTERMPROCCD.PROC_NM not LIKE 'WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MY.SSS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOGIN%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%REGISTRATION%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBERSHIP%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOANS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFITS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%CORPORATE PROFILE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PUBLICATIONS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%OTHER SERVICES%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%HOME%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - SS SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - EC SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%EC - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER MATERNITY NOTIFICATION%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ANNUAL CONFIRMATION OF PENSIONER%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ONLINE INQUIRY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SALARY LOAN%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%TOTAL DISABILITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PARTIAL DISABILITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PENSION MAINTENANCE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%KIOSK FEEDBACK%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%WEBSITE FEEDBACK%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS FUNERAL%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ACTUAL PREMIUMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%EMPLOYMENT HISTORY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%FLEXI-FUND%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOAN STATUS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MATERNITY CLAIMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER DETAILS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SICKNESS CLAIMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%TECHNICAL RETIREMENT%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSID CLEARANCE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFIT CLAIMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - MATERNITY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%DEATH BENEFITS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%CITIZENS CHARTER%' )AS [ELIGIBILITY], (select SSINFOTERMPROCCD.PROC_NM where SSINFOTERMPROCCD.PROC_NM not LIKE 'WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MY.SSS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOGIN%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%REGISTRATION%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBERSHIP%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOANS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFITS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%CORPORATE PROFILE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PUBLICATIONS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%OTHER SERVICES%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%HOME%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - SS SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - EC SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%EC - SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CHECKLIST OF DOCUMENTS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER MATERNITY NOTIFICATION%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ANNUAL CONFIRMATION OF PENSIONER%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ONLINE INQUIRY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SALARY LOAN%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%DDR FUNERAL%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SICKNESS/MATERNITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PENSION MAINTENANCE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%KIOSK FEEDBACK%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%WEBSITE FEEDBACK%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOAN ELIGIBILITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ACTUAL PREMIUMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%EMPLOYMENT HISTORY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%FLEXI-FUND%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOAN STATUS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MATERNITY CLAIMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER DETAILS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SICKNESS CLAIMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%TECHNICAL RETIREMENT%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSID CLEARANCE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFIT CLAIMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - MATERNITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%CITIZENS CHARTER%')as [DDR FUNERAL], (select SSINFOTERMPROCCD.PROC_NM where SSTRANSAT.PROC_CD not like '%10001%'and SSTRANSAT.PROC_CD not like '%10002%'and SSTRANSAT.PROC_CD not like '%10003%'and SSTRANSAT.PROC_CD not like '%10004%'and SSTRANSAT.PROC_CD not like '%10005%'and SSTRANSAT.PROC_CD not like '%10006%'and SSTRANSAT.PROC_CD not like '%10007%'and SSTRANSAT.PROC_CD not like '%10008%'and SSTRANSAT.PROC_CD not like '%10009%'and SSTRANSAT.PROC_CD not like '%10010%'and SSTRANSAT.PROC_CD not like '%10011%'and SSTRANSAT.PROC_CD not like '%10012%'and SSTRANSAT.PROC_CD not like '%10013%'and SSTRANSAT.PROC_CD not like '%10014%'and SSTRANSAT.PROC_CD not like '%10015%'and SSTRANSAT.PROC_CD not like '%10016%'and SSTRANSAT.PROC_CD not like '%10017%'and SSTRANSAT.PROC_CD not like '%10018%'and SSTRANSAT.PROC_CD not like '%10019%'and SSTRANSAT.PROC_CD not like '%10020%'and SSTRANSAT.PROC_CD not like '%10021%'and SSTRANSAT.PROC_CD not like '%10022%'and SSTRANSAT.PROC_CD not like '%10023%'and SSTRANSAT.PROC_CD not like '%10024%'and SSTRANSAT.PROC_CD not like '%10025%'and SSTRANSAT.PROC_CD not like '%10026%'and SSTRANSAT.PROC_CD not like '%10027%'and SSTRANSAT.PROC_CD not like '%10028%'and SSTRANSAT.PROC_CD not like '%10029%'and SSTRANSAT.PROC_CD not like '%10030%'and SSTRANSAT.PROC_CD not like '%10031%'and SSTRANSAT.PROC_CD not like '%10032%'and SSTRANSAT.PROC_CD not like '%10033%'and SSTRANSAT.PROC_CD not like '%10034%'and SSTRANSAT.PROC_CD not like '%10035%'and SSTRANSAT.PROC_CD not like '%10036%'and SSTRANSAT.PROC_CD not like '%10041%'and SSTRANSAT.PROC_CD not like '%10042%') as [SICKNESS/MATERNITY]	FROM SSTRANSAT INNER JOIN SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAT.BRANCH_IP INNER JOIN SSINFOTERMPROCCD ON SSINFOTERMPROCCD.PROC_CD = SSTRANSAT.PROC_CD where SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and SSINFOTERMBR.BRANCH_NM = '" & cbBranch.Text & "' and SSINFOTERMPROCCD.PROC_NM not LIKE 'WEBSITE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MY.SSS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOGIN%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%REGISTRATION%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBERSHIP%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOANS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFITS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CORPORATE PROFILE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%PUBLICATIONS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%OTHER SERVICES%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%HOME%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CITIZENS CHARTER%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%ONLINE INQUIRY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SALARY LOAN%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%TECHNICAL RETIREMENT%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%ANNUAL CONFIRMATION OF PENSIONER%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER MATERNITY NOTIFICATION%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%PENSION MAINTENANCE%' GROUP by SSTRANSAT.ENCODE_DT,SSINFOTERMPROCCD.PROC_NM,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN,SSTRANSAT.PROC_CD")
                '    Dim rpt As New totalBranch
                '    cryRpt = rpt
                '    '   cryRpt.Refresh()
                '   cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                'rptView.ReportSource = cryRpt
                'cryRpt.SetParameterValue("@getBranch", cbBranch.Text)
                '    cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                '    cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
            ElseIf rbTOP.Checked = True And rbCluster.Checked = True Then
                GetName = "TOP HITS REPORT BY DIVISION"
                ''dt = db.ExecuteSQLQuery("SELECT convert(varchar,SSTRANSAT.ENCODE_DT,101) as DateOfTrans,case when SSINFOTERMPROCCD.PROC_NM = 'SSID CLEARANCE' then 'SSS/UMID CARD INFO' else SSINFOTERMPROCCD.PROC_NM end as 'PROC',SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN, COUNT(SSINFOTERMPROCCD.PROC_NM) as [count],(select SSINFOTERMPROCCD.PROC_NM where SSINFOTERMPROCCD.PROC_NM not LIKE '%MY.SSS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOGIN%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%REGISTRATION%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBERSHIP%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOANS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFITS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%CORPORATE PROFILE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PUBLICATIONS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%OTHER SERVICES%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%HOME%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - SS SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - EC SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%EC - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%CHECKLIST OF DOCUMENTS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER MATERNITY NOTIFICATION%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ANNUAL CONFIRMATION OF PENSIONER%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ONLINE INQUIRY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SALARY LOAN%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%TOTAL DISABILITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PARTIAL DISABILITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PENSION MAINTENANCE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%KIOSK FEEDBACK%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%WEBSITE FEEDBACK%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS FUNERAL%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%TECHNICAL RETIREMENT%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%DDR FUNERAL%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SICKNESS/MATERNITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOAN ELIGIBILITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%DDR FUNERAL%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - MATERNITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%CITIZENS CHARTER%') as [ONLINE INQUIRY], (SELECT SSINFOTERMPROCCD.PROC_NM where SSINFOTERMPROCCD.PROC_NM not LIKE 'WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MY.SSS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOGIN%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%REGISTRATION%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBERSHIP%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOANS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFITS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%CORPORATE PROFILE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PUBLICATIONS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%OTHER SERVICES%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%HOME%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - SS SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - EC SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%EC - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER MATERNITY NOTIFICATION%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ANNUAL CONFIRMATION OF PENSIONER%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ONLINE INQUIRY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SALARY LOAN%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%TOTAL DISABILITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PARTIAL DISABILITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PENSION MAINTENANCE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%KIOSK FEEDBACK%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%WEBSITE FEEDBACK%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS FUNERAL%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ACTUAL PREMIUMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%EMPLOYMENT HISTORY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%FLEXI-FUND%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOAN STATUS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MATERNITY CLAIMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER DETAILS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SICKNESS CLAIMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%TECHNICAL RETIREMENT%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSID CLEARANCE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFIT CLAIMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - MATERNITY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%DEATH BENEFITS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%CITIZENS CHARTER%' )AS [ELIGIBILITY], (select SSINFOTERMPROCCD.PROC_NM where SSINFOTERMPROCCD.PROC_NM not LIKE 'WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MY.SSS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOGIN%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%REGISTRATION%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBERSHIP%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOANS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFITS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%CORPORATE PROFILE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PUBLICATIONS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%OTHER SERVICES%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%HOME%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - SS SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - EC SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%EC - SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CHECKLIST OF DOCUMENTS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER MATERNITY NOTIFICATION%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ANNUAL CONFIRMATION OF PENSIONER%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ONLINE INQUIRY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SALARY LOAN%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%DDR FUNERAL%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SICKNESS/MATERNITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PENSION MAINTENANCE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%KIOSK FEEDBACK%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%WEBSITE FEEDBACK%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOAN ELIGIBILITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ACTUAL PREMIUMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%EMPLOYMENT HISTORY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%FLEXI-FUND%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOAN STATUS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MATERNITY CLAIMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER DETAILS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SICKNESS CLAIMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%TECHNICAL RETIREMENT%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSID CLEARANCE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFIT CLAIMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - MATERNITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%CITIZENS CHARTER%')as [DDR FUNERAL], (select SSINFOTERMPROCCD.PROC_NM where SSTRANSAT.PROC_CD not like '%10001%'and SSTRANSAT.PROC_CD not like '%10002%'and SSTRANSAT.PROC_CD not like '%10003%'and SSTRANSAT.PROC_CD not like '%10004%'and SSTRANSAT.PROC_CD not like '%10005%'and SSTRANSAT.PROC_CD not like '%10006%'and SSTRANSAT.PROC_CD not like '%10007%'and SSTRANSAT.PROC_CD not like '%10008%'and SSTRANSAT.PROC_CD not like '%10009%'and SSTRANSAT.PROC_CD not like '%10010%'and SSTRANSAT.PROC_CD not like '%10011%'and SSTRANSAT.PROC_CD not like '%10012%'and SSTRANSAT.PROC_CD not like '%10013%'and SSTRANSAT.PROC_CD not like '%10014%'and SSTRANSAT.PROC_CD not like '%10015%'and SSTRANSAT.PROC_CD not like '%10016%'and SSTRANSAT.PROC_CD not like '%10017%'and SSTRANSAT.PROC_CD not like '%10018%'and SSTRANSAT.PROC_CD not like '%10019%'and SSTRANSAT.PROC_CD not like '%10020%'and SSTRANSAT.PROC_CD not like '%10021%'and SSTRANSAT.PROC_CD not like '%10022%'and SSTRANSAT.PROC_CD not like '%10023%'and SSTRANSAT.PROC_CD not like '%10024%'and SSTRANSAT.PROC_CD not like '%10025%'and SSTRANSAT.PROC_CD not like '%10026%'and SSTRANSAT.PROC_CD not like '%10027%'and SSTRANSAT.PROC_CD not like '%10028%'and SSTRANSAT.PROC_CD not like '%10029%'and SSTRANSAT.PROC_CD not like '%10030%'and SSTRANSAT.PROC_CD not like '%10031%'and SSTRANSAT.PROC_CD not like '%10032%'and SSTRANSAT.PROC_CD not like '%10033%'and SSTRANSAT.PROC_CD not like '%10034%'and SSTRANSAT.PROC_CD not like '%10035%'and SSTRANSAT.PROC_CD not like '%10036%'and SSTRANSAT.PROC_CD not like '%10041%'and SSTRANSAT.PROC_CD not like '%10042%') as [SICKNESS/MATERNITY]	FROM SSTRANSAT INNER JOIN SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAT.BRANCH_IP INNER JOIN SSINFOTERMPROCCD ON SSINFOTERMPROCCD.PROC_CD = SSTRANSAT.PROC_CD where SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "'and SSINFOTERMCLSTR.CLSTR_NM = '" & cbCluster.Text & "' and SSINFOTERMPROCCD.PROC_NM not LIKE 'WEBSITE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MY.SSS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOGIN%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%REGISTRATION%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBERSHIP%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOANS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFITS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CORPORATE PROFILE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%PUBLICATIONS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%OTHER SERVICES%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%HOME%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CITIZENS CHARTER%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%ONLINE INQUIRY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SALARY LOAN%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%TECHNICAL RETIREMENT%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%ANNUAL CONFIRMATION OF PENSIONER%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER MATERNITY NOTIFICATION%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%PENSION MAINTENANCE%' GROUP by SSTRANSAT.ENCODE_DT,SSINFOTERMPROCCD.PROC_NM,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN,SSTRANSAT.PROC_CD")
                Dim rpt As New TOTALLALLCLUSTER
                cryRpt = rpt
                '   cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@getCluster", cbCluster.Text)
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
            ElseIf rbTOP.Checked = True And rbAll.Checked = True Then
                GetName = "TOP HITS REPORT SUMMARY"
                'dt = db.ExecuteSQLQuery("SELECT convert(varchar,SSTRANSAT.ENCODE_DT,101) as DateOfTrans,case when SSINFOTERMPROCCD.PROC_NM = 'SSID CLEARANCE' then 'SSS/UMID CARD INFO' else SSINFOTERMPROCCD.PROC_NM end as 'PROC',SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN, COUNT(SSINFOTERMPROCCD.PROC_NM) as [count],(select SSINFOTERMPROCCD.PROC_NM where SSINFOTERMPROCCD.PROC_NM not LIKE '%MY.SSS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOGIN%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%REGISTRATION%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBERSHIP%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOANS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFITS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%CORPORATE PROFILE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PUBLICATIONS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%OTHER SERVICES%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%HOME%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - SS SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - EC SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%EC - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%CHECKLIST OF DOCUMENTS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER MATERNITY NOTIFICATION%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ANNUAL CONFIRMATION OF PENSIONER%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ONLINE INQUIRY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SALARY LOAN%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%TOTAL DISABILITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PARTIAL DISABILITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PENSION MAINTENANCE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%KIOSK FEEDBACK%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%WEBSITE FEEDBACK%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS FUNERAL%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%TECHNICAL RETIREMENT%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%DDR FUNERAL%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SICKNESS/MATERNITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOAN ELIGIBILITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%DDR FUNERAL%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - MATERNITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%CITIZENS CHARTER%') as [ONLINE INQUIRY], (SELECT SSINFOTERMPROCCD.PROC_NM where SSINFOTERMPROCCD.PROC_NM not LIKE 'WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MY.SSS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOGIN%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%REGISTRATION%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBERSHIP%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOANS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFITS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%CORPORATE PROFILE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PUBLICATIONS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%OTHER SERVICES%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%HOME%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - SS SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - EC SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%EC - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER MATERNITY NOTIFICATION%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ANNUAL CONFIRMATION OF PENSIONER%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ONLINE INQUIRY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SALARY LOAN%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%TOTAL DISABILITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PARTIAL DISABILITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PENSION MAINTENANCE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%KIOSK FEEDBACK%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%WEBSITE FEEDBACK%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS FUNERAL%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ACTUAL PREMIUMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%EMPLOYMENT HISTORY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%FLEXI-FUND%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOAN STATUS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MATERNITY CLAIMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER DETAILS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SICKNESS CLAIMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%TECHNICAL RETIREMENT%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSID CLEARANCE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFIT CLAIMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - MATERNITY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%DEATH BENEFITS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%CITIZENS CHARTER%' )AS [ELIGIBILITY], (select SSINFOTERMPROCCD.PROC_NM where SSINFOTERMPROCCD.PROC_NM not LIKE 'WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MY.SSS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOGIN%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%REGISTRATION%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBERSHIP%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOANS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFITS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%CORPORATE PROFILE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PUBLICATIONS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%OTHER SERVICES%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%HOME%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - SS SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - EC SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%EC - SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CHECKLIST OF DOCUMENTS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER MATERNITY NOTIFICATION%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ANNUAL CONFIRMATION OF PENSIONER%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ONLINE INQUIRY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SALARY LOAN%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%DDR FUNERAL%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SICKNESS/MATERNITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PENSION MAINTENANCE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%KIOSK FEEDBACK%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%WEBSITE FEEDBACK%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOAN ELIGIBILITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ACTUAL PREMIUMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%EMPLOYMENT HISTORY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%FLEXI-FUND%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOAN STATUS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MATERNITY CLAIMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER DETAILS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SICKNESS CLAIMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%TECHNICAL RETIREMENT%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSID CLEARANCE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFIT CLAIMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - MATERNITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%CITIZENS CHARTER%')as [DDR FUNERAL], (select SSINFOTERMPROCCD.PROC_NM where SSTRANSAT.PROC_CD not like '%10001%'and SSTRANSAT.PROC_CD not like '%10002%'and SSTRANSAT.PROC_CD not like '%10003%'and SSTRANSAT.PROC_CD not like '%10004%'and SSTRANSAT.PROC_CD not like '%10005%'and SSTRANSAT.PROC_CD not like '%10006%'and SSTRANSAT.PROC_CD not like '%10007%'and SSTRANSAT.PROC_CD not like '%10008%'and SSTRANSAT.PROC_CD not like '%10009%'and SSTRANSAT.PROC_CD not like '%10010%'and SSTRANSAT.PROC_CD not like '%10011%'and SSTRANSAT.PROC_CD not like '%10012%'and SSTRANSAT.PROC_CD not like '%10013%'and SSTRANSAT.PROC_CD not like '%10014%'and SSTRANSAT.PROC_CD not like '%10015%'and SSTRANSAT.PROC_CD not like '%10016%'and SSTRANSAT.PROC_CD not like '%10017%'and SSTRANSAT.PROC_CD not like '%10018%'and SSTRANSAT.PROC_CD not like '%10019%'and SSTRANSAT.PROC_CD not like '%10020%'and SSTRANSAT.PROC_CD not like '%10021%'and SSTRANSAT.PROC_CD not like '%10022%'and SSTRANSAT.PROC_CD not like '%10023%'and SSTRANSAT.PROC_CD not like '%10024%'and SSTRANSAT.PROC_CD not like '%10025%'and SSTRANSAT.PROC_CD not like '%10026%'and SSTRANSAT.PROC_CD not like '%10027%'and SSTRANSAT.PROC_CD not like '%10028%'and SSTRANSAT.PROC_CD not like '%10029%'and SSTRANSAT.PROC_CD not like '%10030%'and SSTRANSAT.PROC_CD not like '%10031%'and SSTRANSAT.PROC_CD not like '%10032%'and SSTRANSAT.PROC_CD not like '%10033%'and SSTRANSAT.PROC_CD not like '%10034%'and SSTRANSAT.PROC_CD not like '%10035%'and SSTRANSAT.PROC_CD not like '%10036%'and SSTRANSAT.PROC_CD not like '%10041%'and SSTRANSAT.PROC_CD not like '%10042%') as [SICKNESS/MATERNITY]	FROM SSTRANSAT INNER JOIN SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAT.BRANCH_IP INNER JOIN SSINFOTERMPROCCD ON SSINFOTERMPROCCD.PROC_CD = SSTRANSAT.PROC_CD where SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and SSINFOTERMPROCCD.PROC_NM not LIKE 'WEBSITE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MY.SSS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOGIN%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%REGISTRATION%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBERSHIP%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOANS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFITS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CORPORATE PROFILE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%PUBLICATIONS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%OTHER SERVICES%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%HOME%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CITIZENS CHARTER%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%ONLINE INQUIRY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SALARY LOAN%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%TECHNICAL RETIREMENT%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%ANNUAL CONFIRMATION OF PENSIONER%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER MATERNITY NOTIFICATION%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%PENSION MAINTENANCE%' GROUP by SSTRANSAT.ENCODE_DT,SSINFOTERMPROCCD.PROC_NM,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN,SSTRANSAT.PROC_CD")
                Dim rpt As New TOTALLALL
                cryRpt = rpt
                '   cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)

            ElseIf rbQueryReports.Checked = True And rbAll.Checked = True Then
                GetNameBranch()
                If day1 = "1" And day2 = "31" Or day1 = "1" And day2 = "30" Then
                    getMonth()
                    GetName = "APPLICATION USAGE REPORT SUMMARY"
                    'dt = db.ExecuteSQLQuery("SELECT DISTINCT SSINFOTERMKIOSK.BRANCH_IP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM, (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10003' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [ACT PRM], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10007' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [LN STS], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10011' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [SSS/UMDCRD INFO], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10004' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [BNFT CLMS], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10005' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [EMP HSTRY], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10008' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [MTRNTY CLMS], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10006' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [FLX FND], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10010' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [SCK CLMS], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10009' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [MMBR DTLS], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10012' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [DDR FNRL], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10013' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [LN ELGBLTY], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10014' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [SCK/MTRNTY], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10035' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [PRTL DSBLTY], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10034' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [SS FNRL], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10036' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [TTL DSBLTY], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10033' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [SS DTH], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10037' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [SE/V-SICK], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10038' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [SEP-SS SICK], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10039' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [SE/V-MAT], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10040' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [SEP-MAT], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10043' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [SEP-EC SICK], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10044' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [SS SICK], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10045' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [EC SICK] FROM SSINFOTERMKIOSK LEFT OUTER JOIN SSTRANSAT ON SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAT.BRANCH_IP INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSINFOTERMKIOSK.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMKIOSK.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSINFOTERMKIOSK.DIVSN ORDER BY SSINFOTERMKIOSK.KIOSK_ID")
                    Dim rpt As New HitsAll
                    cryRpt = rpt
                    '   cryRpt.Refresh()
                    cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                    rptView.ReportSource = cryRpt
                    cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                    cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                    cryRpt.SetParameterValue("@getName", getBranch)


                    'GetName = "APPLICATION USAGE REPORT SUMMARY"
                    ''dt = db.ExecuteSQLQuery("SELECT SSTRANSAT.ENCODE_DT,case when SSINFOTERMPROCCD.PROC_NM = 'SSID CLEARANCE' then 'SSS/UMID CARD INFO' else SSINFOTERMPROCCD.PROC_NM end as 'PROC',SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN, COUNT(SSINFOTERMPROCCD.PROC_NM) as [count],(select SSINFOTERMPROCCD.PROC_NM where SSINFOTERMPROCCD.PROC_NM not LIKE '%MY.SSS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOGIN%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%REGISTRATION%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBERSHIP%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOANS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFITS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%CORPORATE PROFILE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PUBLICATIONS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%OTHER SERVICES%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%HOME%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - SS SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - EC SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%EC - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%CHECKLIST OF DOCUMENTS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER MATERNITY NOTIFICATION%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ANNUAL CONFIRMATION OF PENSIONER%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ONLINE INQUIRY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SALARY LOAN%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%TOTAL DISABILITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PARTIAL DISABILITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PENSION MAINTENANCE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%KIOSK FEEDBACK%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%WEBSITE FEEDBACK%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS FUNERAL%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%TECHNICAL RETIREMENT%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%DDR FUNERAL%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SICKNESS/MATERNITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOAN ELIGIBILITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%DDR FUNERAL%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - MATERNITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%CITIZENS CHARTER%') as [ONLINE INQUIRY], (SELECT SSINFOTERMPROCCD.PROC_NM where SSINFOTERMPROCCD.PROC_NM not LIKE 'WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MY.SSS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOGIN%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%REGISTRATION%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBERSHIP%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOANS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFITS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%CORPORATE PROFILE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PUBLICATIONS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%OTHER SERVICES%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%HOME%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - SS SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - EC SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%EC - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER MATERNITY NOTIFICATION%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ANNUAL CONFIRMATION OF PENSIONER%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ONLINE INQUIRY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SALARY LOAN%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%TOTAL DISABILITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PARTIAL DISABILITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PENSION MAINTENANCE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%KIOSK FEEDBACK%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%WEBSITE FEEDBACK%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS FUNERAL%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ACTUAL PREMIUMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%EMPLOYMENT HISTORY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%FLEXI-FUND%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOAN STATUS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MATERNITY CLAIMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER DETAILS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SICKNESS CLAIMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%TECHNICAL RETIREMENT%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSID CLEARANCE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFIT CLAIMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - MATERNITY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%DEATH BENEFITS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%CITIZENS CHARTER%' )AS [ELIGIBILITY], (select SSINFOTERMPROCCD.PROC_NM where SSINFOTERMPROCCD.PROC_NM not LIKE 'WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MY.SSS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOGIN%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%REGISTRATION%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBERSHIP%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOANS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFITS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%CORPORATE PROFILE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PUBLICATIONS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%OTHER SERVICES%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%HOME%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - SS SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - EC SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%EC - SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CHECKLIST OF DOCUMENTS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER MATERNITY NOTIFICATION%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ANNUAL CONFIRMATION OF PENSIONER%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ONLINE INQUIRY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SALARY LOAN%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%DDR FUNERAL%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SICKNESS/MATERNITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PENSION MAINTENANCE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%KIOSK FEEDBACK%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%WEBSITE FEEDBACK%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOAN ELIGIBILITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ACTUAL PREMIUMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%EMPLOYMENT HISTORY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%FLEXI-FUND%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOAN STATUS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MATERNITY CLAIMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER DETAILS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SICKNESS CLAIMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%TECHNICAL RETIREMENT%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSID CLEARANCE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFIT CLAIMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - MATERNITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%CITIZENS CHARTER%')as [DDR FUNERAL], (select SSINFOTERMPROCCD.PROC_NM where SSTRANSAT.PROC_CD not like '%10001%'and SSTRANSAT.PROC_CD not like '%10002%'and SSTRANSAT.PROC_CD not like '%10003%'and SSTRANSAT.PROC_CD not like '%10004%'and SSTRANSAT.PROC_CD not like '%10005%'and SSTRANSAT.PROC_CD not like '%10006%'and SSTRANSAT.PROC_CD not like '%10007%'and SSTRANSAT.PROC_CD not like '%10008%'and SSTRANSAT.PROC_CD not like '%10009%'and SSTRANSAT.PROC_CD not like '%10010%'and SSTRANSAT.PROC_CD not like '%10011%'and SSTRANSAT.PROC_CD not like '%10012%'and SSTRANSAT.PROC_CD not like '%10013%'and SSTRANSAT.PROC_CD not like '%10014%'and SSTRANSAT.PROC_CD not like '%10015%'and SSTRANSAT.PROC_CD not like '%10016%'and SSTRANSAT.PROC_CD not like '%10017%'and SSTRANSAT.PROC_CD not like '%10018%'and SSTRANSAT.PROC_CD not like '%10019%'and SSTRANSAT.PROC_CD not like '%10020%'and SSTRANSAT.PROC_CD not like '%10021%'and SSTRANSAT.PROC_CD not like '%10022%'and SSTRANSAT.PROC_CD not like '%10023%'and SSTRANSAT.PROC_CD not like '%10024%'and SSTRANSAT.PROC_CD not like '%10025%'and SSTRANSAT.PROC_CD not like '%10026%'and SSTRANSAT.PROC_CD not like '%10027%'and SSTRANSAT.PROC_CD not like '%10028%'and SSTRANSAT.PROC_CD not like '%10029%'and SSTRANSAT.PROC_CD not like '%10030%'and SSTRANSAT.PROC_CD not like '%10031%'and SSTRANSAT.PROC_CD not like '%10032%'and SSTRANSAT.PROC_CD not like '%10033%'and SSTRANSAT.PROC_CD not like '%10034%'and SSTRANSAT.PROC_CD not like '%10035%'and SSTRANSAT.PROC_CD not like '%10036%'and SSTRANSAT.PROC_CD not like '%10041%'and SSTRANSAT.PROC_CD not like '%10042%') as [SICKNESS/MATERNITY]	FROM SSTRANSAT INNER JOIN SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAT.BRANCH_IP INNER JOIN SSINFOTERMPROCCD ON SSINFOTERMPROCCD.PROC_CD = SSTRANSAT.PROC_CD where SSTRANSAT.ENCODE_DT BETWEEN '" & date12 & "' and '" & date22 & "' and SSINFOTERMPROCCD.PROC_NM not LIKE 'WEBSITE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MY.SSS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOGIN%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%REGISTRATION%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBERSHIP%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOANS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFITS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CORPORATE PROFILE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%PUBLICATIONS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%OTHER SERVICES%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%HOME%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CITIZENS CHARTER%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%ONLINE INQUIRY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SALARY LOAN%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%TECHNICAL RETIREMENT%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%ANNUAL CONFIRMATION OF PENSIONER%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER MATERNITY NOTIFICATION%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%PENSION MAINTENANCE%' GROUP by SSTRANSAT.ENCODE_DT,SSINFOTERMPROCCD.PROC_NM,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN,SSTRANSAT.PROC_CD")
                    'Dim rpt As New SSITQueryReportsMonthly
                    'cryRpt = rpt
                    'cryRpt.Refresh()
                    'cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass, My.Settings.db_Host, My.Settings.db_Server)
                    'rptView.ReportSource = cryRpt
                    'cryRpt.SetParameterValue("@DateFrom", date12)
                    'cryRpt.SetParameterValue("@DateTo", date22)
                Else
                    GetNameBranch()
                    'GetName = "APPLICATION USAGE REPORT SUMMARY"
                    ''dt = db.ExecuteSQLQuery("SELECT SSTRANSAT.ENCODE_DT,case when SSINFOTERMPROCCD.PROC_NM = 'SSID CLEARANCE' then 'SSS/UMID CARD INFO' else SSINFOTERMPROCCD.PROC_NM end as 'PROC',SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN, COUNT(SSINFOTERMPROCCD.PROC_NM) as [count],(select SSINFOTERMPROCCD.PROC_NM where SSINFOTERMPROCCD.PROC_NM not LIKE '%MY.SSS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOGIN%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%REGISTRATION%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBERSHIP%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOANS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFITS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%CORPORATE PROFILE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PUBLICATIONS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%OTHER SERVICES%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%HOME%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - SS SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - EC SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%EC - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%CHECKLIST OF DOCUMENTS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER MATERNITY NOTIFICATION%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ANNUAL CONFIRMATION OF PENSIONER%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ONLINE INQUIRY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SALARY LOAN%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%TOTAL DISABILITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PARTIAL DISABILITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PENSION MAINTENANCE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%KIOSK FEEDBACK%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%WEBSITE FEEDBACK%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS FUNERAL%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%TECHNICAL RETIREMENT%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%DDR FUNERAL%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SICKNESS/MATERNITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOAN ELIGIBILITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%DDR FUNERAL%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - MATERNITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%CITIZENS CHARTER%') as [ONLINE INQUIRY], (SELECT SSINFOTERMPROCCD.PROC_NM where SSINFOTERMPROCCD.PROC_NM not LIKE 'WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MY.SSS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOGIN%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%REGISTRATION%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBERSHIP%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOANS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFITS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%CORPORATE PROFILE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PUBLICATIONS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%OTHER SERVICES%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%HOME%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - SS SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - EC SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%EC - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER MATERNITY NOTIFICATION%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ANNUAL CONFIRMATION OF PENSIONER%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ONLINE INQUIRY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SALARY LOAN%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%TOTAL DISABILITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PARTIAL DISABILITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PENSION MAINTENANCE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%KIOSK FEEDBACK%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%WEBSITE FEEDBACK%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS FUNERAL%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ACTUAL PREMIUMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%EMPLOYMENT HISTORY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%FLEXI-FUND%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOAN STATUS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MATERNITY CLAIMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER DETAILS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SICKNESS CLAIMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%TECHNICAL RETIREMENT%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSID CLEARANCE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFIT CLAIMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - MATERNITY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%DEATH BENEFITS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%CITIZENS CHARTER%' )AS [ELIGIBILITY], (select SSINFOTERMPROCCD.PROC_NM where SSINFOTERMPROCCD.PROC_NM not LIKE 'WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MY.SSS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOGIN%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%REGISTRATION%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBERSHIP%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOANS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFITS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%CORPORATE PROFILE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PUBLICATIONS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%OTHER SERVICES%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%HOME%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - SS SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - EC SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%EC - SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CHECKLIST OF DOCUMENTS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER MATERNITY NOTIFICATION%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ANNUAL CONFIRMATION OF PENSIONER%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ONLINE INQUIRY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SALARY LOAN%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%DDR FUNERAL%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SICKNESS/MATERNITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PENSION MAINTENANCE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%KIOSK FEEDBACK%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%WEBSITE FEEDBACK%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOAN ELIGIBILITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ACTUAL PREMIUMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%EMPLOYMENT HISTORY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%FLEXI-FUND%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOAN STATUS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MATERNITY CLAIMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER DETAILS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SICKNESS CLAIMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%TECHNICAL RETIREMENT%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSID CLEARANCE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFIT CLAIMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - MATERNITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%CITIZENS CHARTER%')as [DDR FUNERAL], (select SSINFOTERMPROCCD.PROC_NM where SSTRANSAT.PROC_CD not like '%10001%'and SSTRANSAT.PROC_CD not like '%10002%'and SSTRANSAT.PROC_CD not like '%10003%'and SSTRANSAT.PROC_CD not like '%10004%'and SSTRANSAT.PROC_CD not like '%10005%'and SSTRANSAT.PROC_CD not like '%10006%'and SSTRANSAT.PROC_CD not like '%10007%'and SSTRANSAT.PROC_CD not like '%10008%'and SSTRANSAT.PROC_CD not like '%10009%'and SSTRANSAT.PROC_CD not like '%10010%'and SSTRANSAT.PROC_CD not like '%10011%'and SSTRANSAT.PROC_CD not like '%10012%'and SSTRANSAT.PROC_CD not like '%10013%'and SSTRANSAT.PROC_CD not like '%10014%'and SSTRANSAT.PROC_CD not like '%10015%'and SSTRANSAT.PROC_CD not like '%10016%'and SSTRANSAT.PROC_CD not like '%10017%'and SSTRANSAT.PROC_CD not like '%10018%'and SSTRANSAT.PROC_CD not like '%10019%'and SSTRANSAT.PROC_CD not like '%10020%'and SSTRANSAT.PROC_CD not like '%10021%'and SSTRANSAT.PROC_CD not like '%10022%'and SSTRANSAT.PROC_CD not like '%10023%'and SSTRANSAT.PROC_CD not like '%10024%'and SSTRANSAT.PROC_CD not like '%10025%'and SSTRANSAT.PROC_CD not like '%10026%'and SSTRANSAT.PROC_CD not like '%10027%'and SSTRANSAT.PROC_CD not like '%10028%'and SSTRANSAT.PROC_CD not like '%10029%'and SSTRANSAT.PROC_CD not like '%10030%'and SSTRANSAT.PROC_CD not like '%10031%'and SSTRANSAT.PROC_CD not like '%10032%'and SSTRANSAT.PROC_CD not like '%10033%'and SSTRANSAT.PROC_CD not like '%10034%'and SSTRANSAT.PROC_CD not like '%10035%'and SSTRANSAT.PROC_CD not like '%10036%'and SSTRANSAT.PROC_CD not like '%10041%'and SSTRANSAT.PROC_CD not like '%10042%') as [SICKNESS/MATERNITY]	FROM SSTRANSAT INNER JOIN SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAT.BRANCH_IP INNER JOIN SSINFOTERMPROCCD ON SSINFOTERMPROCCD.PROC_CD = SSTRANSAT.PROC_CD where SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and SSINFOTERMPROCCD.PROC_NM not LIKE 'WEBSITE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MY.SSS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOGIN%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%REGISTRATION%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBERSHIP%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOANS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFITS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CORPORATE PROFILE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%PUBLICATIONS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%OTHER SERVICES%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%HOME%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CITIZENS CHARTER%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%ONLINE INQUIRY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SALARY LOAN%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%TECHNICAL RETIREMENT%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%ANNUAL CONFIRMATION OF PENSIONER%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER MATERNITY NOTIFICATION%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%PENSION MAINTENANCE%' GROUP by SSTRANSAT.ENCODE_DT,SSINFOTERMPROCCD.PROC_NM,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN,SSTRANSAT.PROC_CD")
                    'Dim rpt As New SSITQueryReports
                    'cryRpt = rpt
                    ''   cryRpt.Refresh()
                    'cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass, My.Settings.db_Host, My.Settings.db_Server)
                    'rptView.ReportSource = cryRpt
                    'cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                    'cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)

                    GetName = "APPLICATION USAGE REPORT SUMMARY"
                    'dt = db.ExecuteSQLQuery("SELECT DISTINCT SSINFOTERMKIOSK.BRANCH_IP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM, (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10003' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [ACT PRM], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10007' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [LN STS], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10011' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [SSS/UMDCRD INFO], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10004' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [BNFT CLMS], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10005' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [EMP HSTRY], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10008' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [MTRNTY CLMS], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10006' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [FLX FND], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10010' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [SCK CLMS], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10009' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [MMBR DTLS], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10012' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [DDR FNRL], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10013' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [LN ELGBLTY], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10014' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [SCK/MTRNTY], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10035' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [PRTL DSBLTY], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10034' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [SS FNRL], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10036' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [TTL DSBLTY], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10033' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [SS DTH], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10037' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [SE/V-SICK], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10038' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [SEP-SS SICK], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10039' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [SE/V-MAT], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10040' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [SEP-MAT], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10043' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [SEP-EC SICK], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10044' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [SS SICK], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10045' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [EC SICK] FROM SSINFOTERMKIOSK LEFT OUTER JOIN SSTRANSAT ON SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAT.BRANCH_IP INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSINFOTERMKIOSK.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMKIOSK.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSINFOTERMKIOSK.DIVSN ORDER BY SSINFOTERMKIOSK.KIOSK_ID")
                    Dim rpt As New HitsAll
                    cryRpt = rpt
                    '   cryRpt.Refresh()
                    cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                    rptView.ReportSource = cryRpt
                    cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                    cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                    cryRpt.SetParameterValue("@getName", getBranch)
                End If
            ElseIf rbQueryReports.Checked = True And rbGroup.Checked = True Then
                'GetName = "APPLICATION USAGE REPORT BY GROUP"
                ''dt = db.ExecuteSQLQuery("SELECT convert(varchar,SSTRANSAT.ENCODE_DT,101) as DateOfTrans,case when SSINFOTERMPROCCD.PROC_NM = 'SSID CLEARANCE' then 'SSS/UMID CARD INFO' else SSINFOTERMPROCCD.PROC_NM end as 'PROC',SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN, COUNT(SSINFOTERMPROCCD.PROC_NM) as [count],(select SSINFOTERMPROCCD.PROC_NM where SSINFOTERMPROCCD.PROC_NM not LIKE '%MY.SSS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOGIN%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%REGISTRATION%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBERSHIP%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOANS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFITS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%CORPORATE PROFILE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PUBLICATIONS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%OTHER SERVICES%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%HOME%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - SS SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - EC SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%EC - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%CHECKLIST OF DOCUMENTS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER MATERNITY NOTIFICATION%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ANNUAL CONFIRMATION OF PENSIONER%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ONLINE INQUIRY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SALARY LOAN%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%TOTAL DISABILITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PARTIAL DISABILITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PENSION MAINTENANCE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%KIOSK FEEDBACK%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%WEBSITE FEEDBACK%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS FUNERAL%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%TECHNICAL RETIREMENT%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%DDR FUNERAL%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SICKNESS/MATERNITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOAN ELIGIBILITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%DDR FUNERAL%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - MATERNITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%CITIZENS CHARTER%') as [ONLINE INQUIRY], (SELECT SSINFOTERMPROCCD.PROC_NM where SSINFOTERMPROCCD.PROC_NM not LIKE 'WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MY.SSS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOGIN%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%REGISTRATION%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBERSHIP%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOANS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFITS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%CORPORATE PROFILE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PUBLICATIONS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%OTHER SERVICES%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%HOME%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - SS SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - EC SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%EC - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER MATERNITY NOTIFICATION%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ANNUAL CONFIRMATION OF PENSIONER%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ONLINE INQUIRY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SALARY LOAN%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%TOTAL DISABILITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PARTIAL DISABILITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PENSION MAINTENANCE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%KIOSK FEEDBACK%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%WEBSITE FEEDBACK%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS FUNERAL%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ACTUAL PREMIUMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%EMPLOYMENT HISTORY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%FLEXI-FUND%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOAN STATUS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MATERNITY CLAIMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER DETAILS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SICKNESS CLAIMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%TECHNICAL RETIREMENT%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSID CLEARANCE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFIT CLAIMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - MATERNITY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%DEATH BENEFITS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%CITIZENS CHARTER%' )AS [ELIGIBILITY], (select SSINFOTERMPROCCD.PROC_NM where SSINFOTERMPROCCD.PROC_NM not LIKE 'WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MY.SSS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOGIN%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%REGISTRATION%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBERSHIP%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOANS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFITS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%CORPORATE PROFILE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PUBLICATIONS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%OTHER SERVICES%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%HOME%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - SS SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - EC SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%EC - SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CHECKLIST OF DOCUMENTS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER MATERNITY NOTIFICATION%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ANNUAL CONFIRMATION OF PENSIONER%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ONLINE INQUIRY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SALARY LOAN%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%DDR FUNERAL%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SICKNESS/MATERNITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PENSION MAINTENANCE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%KIOSK FEEDBACK%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%WEBSITE FEEDBACK%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOAN ELIGIBILITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ACTUAL PREMIUMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%EMPLOYMENT HISTORY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%FLEXI-FUND%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOAN STATUS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MATERNITY CLAIMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER DETAILS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SICKNESS CLAIMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%TECHNICAL RETIREMENT%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSID CLEARANCE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFIT CLAIMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - MATERNITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%CITIZENS CHARTER%')as [DDR FUNERAL], (select SSINFOTERMPROCCD.PROC_NM where SSTRANSAT.PROC_CD not like '%10001%'and SSTRANSAT.PROC_CD not like '%10002%'and SSTRANSAT.PROC_CD not like '%10003%'and SSTRANSAT.PROC_CD not like '%10004%'and SSTRANSAT.PROC_CD not like '%10005%'and SSTRANSAT.PROC_CD not like '%10006%'and SSTRANSAT.PROC_CD not like '%10007%'and SSTRANSAT.PROC_CD not like '%10008%'and SSTRANSAT.PROC_CD not like '%10009%'and SSTRANSAT.PROC_CD not like '%10010%'and SSTRANSAT.PROC_CD not like '%10011%'and SSTRANSAT.PROC_CD not like '%10012%'and SSTRANSAT.PROC_CD not like '%10013%'and SSTRANSAT.PROC_CD not like '%10014%'and SSTRANSAT.PROC_CD not like '%10015%'and SSTRANSAT.PROC_CD not like '%10016%'and SSTRANSAT.PROC_CD not like '%10017%'and SSTRANSAT.PROC_CD not like '%10018%'and SSTRANSAT.PROC_CD not like '%10019%'and SSTRANSAT.PROC_CD not like '%10020%'and SSTRANSAT.PROC_CD not like '%10021%'and SSTRANSAT.PROC_CD not like '%10022%'and SSTRANSAT.PROC_CD not like '%10023%'and SSTRANSAT.PROC_CD not like '%10024%'and SSTRANSAT.PROC_CD not like '%10025%'and SSTRANSAT.PROC_CD not like '%10026%'and SSTRANSAT.PROC_CD not like '%10027%'and SSTRANSAT.PROC_CD not like '%10028%'and SSTRANSAT.PROC_CD not like '%10029%'and SSTRANSAT.PROC_CD not like '%10030%'and SSTRANSAT.PROC_CD not like '%10031%'and SSTRANSAT.PROC_CD not like '%10032%'and SSTRANSAT.PROC_CD not like '%10033%'and SSTRANSAT.PROC_CD not like '%10034%'and SSTRANSAT.PROC_CD not like '%10035%'and SSTRANSAT.PROC_CD not like '%10036%'and SSTRANSAT.PROC_CD not like '%10041%'and SSTRANSAT.PROC_CD not like '%10042%') as [SICKNESS/MATERNITY] FROM SSTRANSAT INNER JOIN SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAT.BRANCH_IP INNER JOIN SSINFOTERMPROCCD ON SSINFOTERMPROCCD.PROC_CD = SSTRANSAT.PROC_CD where SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and SSTRANSAT.DIVSN = '" & cbGroup.Text & "' and SSINFOTERMPROCCD.PROC_NM not LIKE 'WEBSITE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MY.SSS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOGIN%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%REGISTRATION%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBERSHIP%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOANS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFITS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CORPORATE PROFILE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%PUBLICATIONS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%OTHER SERVICES%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%HOME%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - SS SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CITIZENS CHARTER%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%ONLINE INQUIRY%' GROUP by SSTRANSAT.ENCODE_DT,SSINFOTERMPROCCD.PROC_NM,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN,SSTRANSAT.PROC_CD")
                'Dim rpt As New APPLICATIONUSAGEGROUP
                'cryRpt = rpt
                ''      cryRpt.Refresh()
                'cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass, My.Settings.db_Host, My.Settings.db_Server)
                'rptView.ReportSource = cryRpt
                'cryRpt.SetParameterValue("@getGroup", cbGroup.Text)
                'cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                'cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                GetNameBranch()
                GetName = "APPLICATION USAGE REPORT BY BRANCH"
                'dt = db.ExecuteSQLQuery("SELECT DISTINCT SSINFOTERMKIOSK.BRANCH_IP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM, (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10003' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [ACT PRM], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10007' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [LN STS], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10011' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [SSS/UMDCRD INFO], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10004' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [BNFT CLMS], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10005' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [EMP HSTRY], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10008' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [MTRNTY CLMS], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10006' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [FLX FND], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10010' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [SCK CLMS], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10009' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [MMBR DTLS], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10012' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [DDR FNRL], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10013' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [LN ELGBLTY], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10014' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [SCK/MTRNTY], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10035' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [PRTL DSBLTY], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10034' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [SS FNRL], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10036' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [TTL DSBLTY], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10033' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [SS DTH], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10037' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [SE/V-SICK], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10038' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [SEP-SS SICK], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10039' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [SE/V-MAT], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10040' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [SEP-MAT], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10043' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [SEP-EC SICK], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10044' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [SS SICK], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10045' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [EC SICK] FROM SSINFOTERMKIOSK LEFT OUTER JOIN SSTRANSAT ON SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAT.BRANCH_IP INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSINFOTERMKIOSK.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMKIOSK.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSINFOTERMKIOSK.DIVSN WHERE SSINFOTERMGROUP.GROUP_NM = '" & cbGroup.Text & "'ORDER BY SSINFOTERMKIOSK.KIOSK_ID")
                Dim rpt As New HitsAll
                cryRpt = rpt
                '   cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)

            ElseIf rbQueryReports.Checked = True And rbBranch.Checked = True Then
                'GetName = "APPLICATION USAGE REPORT BY BRANCH"
                ''dt = db.ExecuteSQLQuery("SELECT convert(varchar,SSTRANSAT.ENCODE_DT,101) as DateOfTrans,case when SSINFOTERMPROCCD.PROC_NM = 'SSID CLEARANCE' then 'SSS/UMID CARD INFO' else SSINFOTERMPROCCD.PROC_NM end as 'PROC',SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN, COUNT(SSINFOTERMPROCCD.PROC_NM) as [count],(select SSINFOTERMPROCCD.PROC_NM where SSINFOTERMPROCCD.PROC_NM not LIKE '%MY.SSS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOGIN%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%REGISTRATION%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBERSHIP%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOANS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFITS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%CORPORATE PROFILE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PUBLICATIONS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%OTHER SERVICES%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%HOME%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - SS SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - EC SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%EC - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%CHECKLIST OF DOCUMENTS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER MATERNITY NOTIFICATION%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ANNUAL CONFIRMATION OF PENSIONER%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ONLINE INQUIRY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SALARY LOAN%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%TOTAL DISABILITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PARTIAL DISABILITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PENSION MAINTENANCE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%KIOSK FEEDBACK%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%WEBSITE FEEDBACK%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS FUNERAL%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%TECHNICAL RETIREMENT%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%DDR FUNERAL%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SICKNESS/MATERNITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOAN ELIGIBILITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%DDR FUNERAL%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - MATERNITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%CITIZENS CHARTER%') as [ONLINE INQUIRY], (SELECT SSINFOTERMPROCCD.PROC_NM where SSINFOTERMPROCCD.PROC_NM not LIKE 'WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MY.SSS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOGIN%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%REGISTRATION%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBERSHIP%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOANS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFITS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%CORPORATE PROFILE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PUBLICATIONS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%OTHER SERVICES%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%HOME%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - SS SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - EC SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%EC - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER MATERNITY NOTIFICATION%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ANNUAL CONFIRMATION OF PENSIONER%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ONLINE INQUIRY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SALARY LOAN%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%TOTAL DISABILITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PARTIAL DISABILITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PENSION MAINTENANCE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%KIOSK FEEDBACK%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%WEBSITE FEEDBACK%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS FUNERAL%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ACTUAL PREMIUMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%EMPLOYMENT HISTORY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%FLEXI-FUND%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOAN STATUS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MATERNITY CLAIMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER DETAILS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SICKNESS CLAIMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%TECHNICAL RETIREMENT%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSID CLEARANCE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFIT CLAIMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - MATERNITY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%DEATH BENEFITS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%CITIZENS CHARTER%' )AS [ELIGIBILITY], (select SSINFOTERMPROCCD.PROC_NM where SSINFOTERMPROCCD.PROC_NM not LIKE 'WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MY.SSS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOGIN%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%REGISTRATION%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBERSHIP%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOANS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFITS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%CORPORATE PROFILE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PUBLICATIONS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%OTHER SERVICES%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%HOME%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - SS SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - EC SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%EC - SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CHECKLIST OF DOCUMENTS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER MATERNITY NOTIFICATION%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ANNUAL CONFIRMATION OF PENSIONER%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ONLINE INQUIRY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SALARY LOAN%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%DDR FUNERAL%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SICKNESS/MATERNITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PENSION MAINTENANCE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%KIOSK FEEDBACK%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%WEBSITE FEEDBACK%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOAN ELIGIBILITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ACTUAL PREMIUMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%EMPLOYMENT HISTORY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%FLEXI-FUND%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOAN STATUS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MATERNITY CLAIMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER DETAILS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SICKNESS CLAIMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%TECHNICAL RETIREMENT%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSID CLEARANCE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFIT CLAIMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - MATERNITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%CITIZENS CHARTER%')as [DDR FUNERAL], (select SSINFOTERMPROCCD.PROC_NM where SSTRANSAT.PROC_CD not like '%10001%'and SSTRANSAT.PROC_CD not like '%10002%'and SSTRANSAT.PROC_CD not like '%10003%'and SSTRANSAT.PROC_CD not like '%10004%'and SSTRANSAT.PROC_CD not like '%10005%'and SSTRANSAT.PROC_CD not like '%10006%'and SSTRANSAT.PROC_CD not like '%10007%'and SSTRANSAT.PROC_CD not like '%10008%'and SSTRANSAT.PROC_CD not like '%10009%'and SSTRANSAT.PROC_CD not like '%10010%'and SSTRANSAT.PROC_CD not like '%10011%'and SSTRANSAT.PROC_CD not like '%10012%'and SSTRANSAT.PROC_CD not like '%10013%'and SSTRANSAT.PROC_CD not like '%10014%'and SSTRANSAT.PROC_CD not like '%10015%'and SSTRANSAT.PROC_CD not like '%10016%'and SSTRANSAT.PROC_CD not like '%10017%'and SSTRANSAT.PROC_CD not like '%10018%'and SSTRANSAT.PROC_CD not like '%10019%'and SSTRANSAT.PROC_CD not like '%10020%'and SSTRANSAT.PROC_CD not like '%10021%'and SSTRANSAT.PROC_CD not like '%10022%'and SSTRANSAT.PROC_CD not like '%10023%'and SSTRANSAT.PROC_CD not like '%10024%'and SSTRANSAT.PROC_CD not like '%10025%'and SSTRANSAT.PROC_CD not like '%10026%'and SSTRANSAT.PROC_CD not like '%10027%'and SSTRANSAT.PROC_CD not like '%10028%'and SSTRANSAT.PROC_CD not like '%10029%'and SSTRANSAT.PROC_CD not like '%10030%'and SSTRANSAT.PROC_CD not like '%10031%'and SSTRANSAT.PROC_CD not like '%10032%'and SSTRANSAT.PROC_CD not like '%10033%'and SSTRANSAT.PROC_CD not like '%10034%'and SSTRANSAT.PROC_CD not like '%10035%'and SSTRANSAT.PROC_CD not like '%10036%'and SSTRANSAT.PROC_CD not like '%10041%'and SSTRANSAT.PROC_CD not like '%10042%') as [SICKNESS/MATERNITY] FROM SSTRANSAT INNER JOIN SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAT.BRANCH_IP INNER JOIN SSINFOTERMPROCCD ON SSINFOTERMPROCCD.PROC_CD = SSTRANSAT.PROC_CD where SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and SSTRANSAT.BRANCH_CD = '" & cbBranch.Text & "' and SSINFOTERMPROCCD.PROC_NM not LIKE 'WEBSITE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MY.SSS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOGIN%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%REGISTRATION%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBERSHIP%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOANS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFITS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CORPORATE PROFILE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%PUBLICATIONS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%OTHER SERVICES%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%HOME%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - SS SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CITIZENS CHARTER%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%ONLINE INQUIRY%' GROUP by SSTRANSAT.ENCODE_DT,SSINFOTERMPROCCD.PROC_NM,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN,SSTRANSAT.PROC_CD")
                'Dim rpt As New ApplicationUsageTotal
                'cryRpt = rpt
                ''      cryRpt.Refresh()
                'cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass, My.Settings.db_Host, My.Settings.db_Server)
                'rptView.ReportSource = cryRpt
                'cryRpt.SetParameterValue("@getBranch", cbBranch.Text)
                'cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                'cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                GetNameBranch()
                GetName = "APPLICATION USAGE REPORT BY BRANCH"
                'dt = db.ExecuteSQLQuery("SELECT DISTINCT SSINFOTERMKIOSK.BRANCH_IP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM, (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10003' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [ACT PRM], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10007' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [LN STS], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10011' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [SSS/UMDCRD INFO], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10004' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [BNFT CLMS], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10005' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [EMP HSTRY], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10008' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [MTRNTY CLMS], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10006' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [FLX FND], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10010' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [SCK CLMS], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10009' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [MMBR DTLS], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10012' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [DDR FNRL], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10013' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [LN ELGBLTY], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10014' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [SCK/MTRNTY], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10035' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [PRTL DSBLTY], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10034' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [SS FNRL], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10036' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [TTL DSBLTY], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10033' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [SS DTH], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10037' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [SE/V-SICK], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10038' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [SEP-SS SICK], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10039' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [SE/V-MAT], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10040' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [SEP-MAT], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10043' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [SEP-EC SICK], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10044' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [SS SICK], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10045' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [EC SICK] FROM SSINFOTERMKIOSK LEFT OUTER JOIN SSTRANSAT ON SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAT.BRANCH_IP INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSINFOTERMKIOSK.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMKIOSK.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSINFOTERMKIOSK.DIVSN WHERE SSINFOTERMBR.BRANCH_NM = '" & cbBranch.Text & "'ORDER BY SSINFOTERMKIOSK.KIOSK_ID")
                Dim rpt As New HitsAll
                cryRpt = rpt
                '   cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)

            ElseIf rbQueryReports.Checked = True And rbCluster.Checked = True Then

                'GetName = "APPLICATION USAGE REPORT BY DIVISION"
                ''dt = db.ExecuteSQLQuery("SELECT convert(varchar,SSTRANSAT.ENCODE_DT,101) as DateOfTrans,case when SSINFOTERMPROCCD.PROC_NM = 'SSID CLEARANCE' then 'SSS/UMID CARD INFO' else SSINFOTERMPROCCD.PROC_NM end as 'PROC',SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN, COUNT(SSINFOTERMPROCCD.PROC_NM) as [count],(select SSINFOTERMPROCCD.PROC_NM where SSINFOTERMPROCCD.PROC_NM not LIKE '%MY.SSS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOGIN%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%REGISTRATION%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBERSHIP%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOANS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFITS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%CORPORATE PROFILE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PUBLICATIONS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%OTHER SERVICES%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%HOME%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - SS SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - EC SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%EC - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%CHECKLIST OF DOCUMENTS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER MATERNITY NOTIFICATION%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ANNUAL CONFIRMATION OF PENSIONER%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ONLINE INQUIRY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SALARY LOAN%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%TOTAL DISABILITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PARTIAL DISABILITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PENSION MAINTENANCE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%KIOSK FEEDBACK%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%WEBSITE FEEDBACK%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS FUNERAL%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%TECHNICAL RETIREMENT%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%DDR FUNERAL%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SICKNESS/MATERNITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOAN ELIGIBILITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%DDR FUNERAL%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - MATERNITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%CITIZENS CHARTER%') as [ONLINE INQUIRY], (SELECT SSINFOTERMPROCCD.PROC_NM where SSINFOTERMPROCCD.PROC_NM not LIKE 'WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MY.SSS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOGIN%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%REGISTRATION%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBERSHIP%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOANS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFITS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%CORPORATE PROFILE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PUBLICATIONS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%OTHER SERVICES%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%HOME%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - SS SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - EC SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%EC - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER MATERNITY NOTIFICATION%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ANNUAL CONFIRMATION OF PENSIONER%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ONLINE INQUIRY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SALARY LOAN%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%TOTAL DISABILITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PARTIAL DISABILITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PENSION MAINTENANCE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%KIOSK FEEDBACK%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%WEBSITE FEEDBACK%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS FUNERAL%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ACTUAL PREMIUMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%EMPLOYMENT HISTORY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%FLEXI-FUND%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOAN STATUS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MATERNITY CLAIMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER DETAILS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SICKNESS CLAIMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%TECHNICAL RETIREMENT%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSID CLEARANCE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFIT CLAIMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - MATERNITY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%DEATH BENEFITS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%CITIZENS CHARTER%' )AS [ELIGIBILITY], (select SSINFOTERMPROCCD.PROC_NM where SSINFOTERMPROCCD.PROC_NM not LIKE 'WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MY.SSS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOGIN%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%REGISTRATION%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBERSHIP%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOANS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFITS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%CORPORATE PROFILE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PUBLICATIONS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%OTHER SERVICES%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%HOME%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - SS SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - EC SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS - SICKNESS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%EC - SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CHECKLIST OF DOCUMENTS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER MATERNITY NOTIFICATION%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ANNUAL CONFIRMATION OF PENSIONER%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ONLINE INQUIRY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SALARY LOAN%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%DDR FUNERAL%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SICKNESS/MATERNITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%PENSION MAINTENANCE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%KIOSK FEEDBACK%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%WEBSITE FEEDBACK%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOAN ELIGIBILITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%ACTUAL PREMIUMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%EMPLOYMENT HISTORY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%FLEXI-FUND%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOAN STATUS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MATERNITY CLAIMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER DETAILS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SICKNESS CLAIMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%TECHNICAL RETIREMENT%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSID CLEARANCE%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFIT CLAIMS%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - MATERNITY%'and SSINFOTERMPROCCD.PROC_NM not LIKE '%CITIZENS CHARTER%')as [DDR FUNERAL], (select SSINFOTERMPROCCD.PROC_NM where SSTRANSAT.PROC_CD not like '%10001%'and SSTRANSAT.PROC_CD not like '%10002%'and SSTRANSAT.PROC_CD not like '%10003%'and SSTRANSAT.PROC_CD not like '%10004%'and SSTRANSAT.PROC_CD not like '%10005%'and SSTRANSAT.PROC_CD not like '%10006%'and SSTRANSAT.PROC_CD not like '%10007%'and SSTRANSAT.PROC_CD not like '%10008%'and SSTRANSAT.PROC_CD not like '%10009%'and SSTRANSAT.PROC_CD not like '%10010%'and SSTRANSAT.PROC_CD not like '%10011%'and SSTRANSAT.PROC_CD not like '%10012%'and SSTRANSAT.PROC_CD not like '%10013%'and SSTRANSAT.PROC_CD not like '%10014%'and SSTRANSAT.PROC_CD not like '%10015%'and SSTRANSAT.PROC_CD not like '%10016%'and SSTRANSAT.PROC_CD not like '%10017%'and SSTRANSAT.PROC_CD not like '%10018%'and SSTRANSAT.PROC_CD not like '%10019%'and SSTRANSAT.PROC_CD not like '%10020%'and SSTRANSAT.PROC_CD not like '%10021%'and SSTRANSAT.PROC_CD not like '%10022%'and SSTRANSAT.PROC_CD not like '%10023%'and SSTRANSAT.PROC_CD not like '%10024%'and SSTRANSAT.PROC_CD not like '%10025%'and SSTRANSAT.PROC_CD not like '%10026%'and SSTRANSAT.PROC_CD not like '%10027%'and SSTRANSAT.PROC_CD not like '%10028%'and SSTRANSAT.PROC_CD not like '%10029%'and SSTRANSAT.PROC_CD not like '%10030%'and SSTRANSAT.PROC_CD not like '%10031%'and SSTRANSAT.PROC_CD not like '%10032%'and SSTRANSAT.PROC_CD not like '%10033%'and SSTRANSAT.PROC_CD not like '%10034%'and SSTRANSAT.PROC_CD not like '%10035%'and SSTRANSAT.PROC_CD not like '%10036%'and SSTRANSAT.PROC_CD not like '%10041%'and SSTRANSAT.PROC_CD not like '%10042%') as [SICKNESS/MATERNITY] FROM SSTRANSAT INNER JOIN SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAT.BRANCH_IP INNER JOIN SSINFOTERMPROCCD ON SSINFOTERMPROCCD.PROC_CD = SSTRANSAT.PROC_CD where SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and SSTRANSAT.CLSTR = '" & cbCluster.Text & "' and SSINFOTERMPROCCD.PROC_NM not LIKE 'WEBSITE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MY.SSS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOGIN%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%REGISTRATION%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBERSHIP%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOANS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFITS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CORPORATE PROFILE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%PUBLICATIONS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%OTHER SERVICES%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%HOME%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - SS SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CITIZENS CHARTER%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%ONLINE INQUIRY%' GROUP by SSTRANSAT.ENCODE_DT,SSINFOTERMPROCCD.PROC_NM,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN,SSTRANSAT.PROC_CD")
                'Dim rpt As New ApplicationUsageCluster
                'cryRpt = rpt
                ''   cryRpt.Refresh()
                'cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass, My.Settings.db_Host, My.Settings.db_Server)
                'rptView.ReportSource = cryRpt
                'cryRpt.SetParameterValue("@getCluster", cbCluster.Text)
                'cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                'cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                GetNameBranch()
                GetName = "APPLICATION USAGE REPORT BY BRANCH"
                'dt = db.ExecuteSQLQuery("SELECT DISTINCT SSINFOTERMKIOSK.BRANCH_IP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM, (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10003' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [ACT PRM], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10007' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [LN STS], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10011' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [SSS/UMDCRD INFO], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10004' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [BNFT CLMS], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10005' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [EMP HSTRY], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10008' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [MTRNTY CLMS], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10006' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [FLX FND], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10010' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [SCK CLMS], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10009' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [MMBR DTLS], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10012' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [DDR FNRL], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10013' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [LN ELGBLTY], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10014' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [SCK/MTRNTY], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10035' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [PRTL DSBLTY], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10034' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [SS FNRL], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10036' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [TTL DSBLTY], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10033' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [SS DTH], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10037' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [SE/V-SICK], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10038' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [SEP-SS SICK], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10039' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [SE/V-MAT], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10040' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [SEP-MAT], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10043' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [SEP-EC SICK], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10044' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [SS SICK], (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10045' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS [EC SICK] FROM SSINFOTERMKIOSK LEFT OUTER JOIN SSTRANSAT ON SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAT.BRANCH_IP INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSINFOTERMKIOSK.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMKIOSK.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSINFOTERMKIOSK.DIVSN WHERE SSINFOTERMCLSTR.CLSTR_NM = '" & cbCluster.Text & "'ORDER BY SSINFOTERMKIOSK.KIOSK_ID")
                Dim rpt As New HitsAll
                cryRpt = rpt
                '   cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)

            ElseIf rbOnlineInquiry.Checked = True And rbAll.Checked = True Then
                getMonth()
                If rbAll.Checked = True And day1 = "1" And day2 = "31" Or day1 = "1" And day2 = "30" Then
                    GetNameBranch()
                    GetName = "ONLINE INQUIRY REPORT SUMMARY"
                    ''dt = db.ExecuteSQLQuery("SELECT SSTRANSAT.ENCODE_DT,case when SSINFOTERMPROCCD.PROC_NM = 'SSID CLEARANCE' then 'SSS/UMID CARD INFO' else SSINFOTERMPROCCD.PROC_NM end as 'PROC',SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN, COUNT(SSINFOTERMPROCCD.PROC_NM) as [count] FROM SSTRANSAT INNER JOIN SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAT.BRANCH_IP INNER JOIN SSINFOTERMPROCCD ON SSINFOTERMPROCCD.PROC_CD = SSTRANSAT.PROC_CD where SSTRANSAT.ENCODE_DT BETWEEN '" & date12 & "' and '" & date22 & "' and SSINFOTERMPROCCD.PROC_NM not LIKE 'WEBSITE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MY.SSS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOGIN%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%REGISTRATION%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBERSHIP%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOANS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFITS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CORPORATE PROFILE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%PUBLICATIONS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%OTHER SERVICES%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%HOME%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - SS SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - EC SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS - SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%EC - SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CHECKLIST OF DOCUMENTS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER MATERNITY NOTIFICATION%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%ANNUAL CONFIRMATION OF PENSIONER%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%ONLINE INQUIRY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SALARY LOAN%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%TOTAL DISABILITY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%PARTIAL DISABILITY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%PENSION MAINTENANCE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%KIOSK FEEDBACK%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%WEBSITE FEEDBACK%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS FUNERAL%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CITIZENS CHARTER%' GROUP by SSTRANSAT.ENCODE_DT,SSINFOTERMPROCCD.PROC_NM,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN")
                    Dim rpt As New OnlineInquiryall
                    cryRpt = rpt
                    'cryRpt.Refresh()
                    cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                    rptView.ReportSource = cryRpt
                    cryRpt.SetParameterValue("@DateFrom", date12)
                    cryRpt.SetParameterValue("@DateTo", date22)
                    cryRpt.SetParameterValue("@getName", getBranch)
                Else
                    GetName = "ONLINE INQUIRY REPORT SUMMARY"
                    ''dt = db.ExecuteSQLQuery("SELECT SSTRANSAT.ENCODE_DT,case when SSINFOTERMPROCCD.PROC_NM = 'SSID CLEARANCE' then 'SSS/UMID CARD INFO' else SSINFOTERMPROCCD.PROC_NM end as 'PROC',SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN, COUNT(SSINFOTERMPROCCD.PROC_NM) as [count] FROM SSTRANSAT INNER JOIN SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAT.BRANCH_IP INNER JOIN SSINFOTERMPROCCD ON SSINFOTERMPROCCD.PROC_CD = SSTRANSAT.PROC_CD where SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and SSINFOTERMPROCCD.PROC_NM not LIKE 'WEBSITE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MY.SSS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOGIN%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%REGISTRATION%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBERSHIP%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOANS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFITS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CORPORATE PROFILE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%PUBLICATIONS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%OTHER SERVICES%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%HOME%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - SS SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - EC SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS - SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%EC - SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CHECKLIST OF DOCUMENTS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER MATERNITY NOTIFICATION%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%ANNUAL CONFIRMATION OF PENSIONER%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%ONLINE INQUIRY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SALARY LOAN%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%TOTAL DISABILITY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%PARTIAL DISABILITY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%PENSION MAINTENANCE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%KIOSK FEEDBACK%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%WEBSITE FEEDBACK%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS FUNERAL%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CITIZENS CHARTER%' GROUP by SSTRANSAT.ENCODE_DT,SSINFOTERMPROCCD.PROC_NM,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN")
                    GetNameBranch()
                    GetName = "ONLINE INQUIRY REPORT SUMMARY"
                    ''dt = db.ExecuteSQLQuery("SELECT SSTRANSAT.ENCODE_DT,case when SSINFOTERMPROCCD.PROC_NM = 'SSID CLEARANCE' then 'SSS/UMID CARD INFO' else SSINFOTERMPROCCD.PROC_NM end as 'PROC',SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN, COUNT(SSINFOTERMPROCCD.PROC_NM) as [count] FROM SSTRANSAT INNER JOIN SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAT.BRANCH_IP INNER JOIN SSINFOTERMPROCCD ON SSINFOTERMPROCCD.PROC_CD = SSTRANSAT.PROC_CD where SSTRANSAT.ENCODE_DT BETWEEN '" & date12 & "' and '" & date22 & "' and SSINFOTERMPROCCD.PROC_NM not LIKE 'WEBSITE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MY.SSS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOGIN%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%REGISTRATION%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBERSHIP%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOANS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFITS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CORPORATE PROFILE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%PUBLICATIONS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%OTHER SERVICES%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%HOME%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - SS SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - EC SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS - SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%EC - SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CHECKLIST OF DOCUMENTS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER MATERNITY NOTIFICATION%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%ANNUAL CONFIRMATION OF PENSIONER%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%ONLINE INQUIRY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SALARY LOAN%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%TOTAL DISABILITY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%PARTIAL DISABILITY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%PENSION MAINTENANCE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%KIOSK FEEDBACK%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%WEBSITE FEEDBACK%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS FUNERAL%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CITIZENS CHARTER%' GROUP by SSTRANSAT.ENCODE_DT,SSINFOTERMPROCCD.PROC_NM,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN")
                    Dim rpt As New OnlineInquiryall
                    cryRpt = rpt
                    'cryRpt.Refresh()
                    cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                    rptView.ReportSource = cryRpt
                    cryRpt.SetParameterValue("@DateFrom", date12)
                    cryRpt.SetParameterValue("@DateTo", date22)
                    cryRpt.SetParameterValue("@getName", getBranch)
                End If
            ElseIf rbOnlineInquiry.Checked = True And rbGroup.Checked = True Then
                GetNameBranch()
                GetName = "ONLINE INQUIRY REPORT SUMMARY"
                ''dt = db.ExecuteSQLQuery("SELECT SSTRANSAT.ENCODE_DT,case when SSINFOTERMPROCCD.PROC_NM = 'SSID CLEARANCE' then 'SSS/UMID CARD INFO' else SSINFOTERMPROCCD.PROC_NM end as 'PROC',SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN, COUNT(SSINFOTERMPROCCD.PROC_NM) as [count] FROM SSTRANSAT INNER JOIN SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAT.BRANCH_IP INNER JOIN SSINFOTERMPROCCD ON SSINFOTERMPROCCD.PROC_CD = SSTRANSAT.PROC_CD where SSTRANSAT.ENCODE_DT BETWEEN '" & date12 & "' and '" & date22 & "' and SSINFOTERMPROCCD.PROC_NM not LIKE 'WEBSITE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MY.SSS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOGIN%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%REGISTRATION%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBERSHIP%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOANS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFITS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CORPORATE PROFILE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%PUBLICATIONS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%OTHER SERVICES%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%HOME%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - SS SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - EC SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS - SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%EC - SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CHECKLIST OF DOCUMENTS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER MATERNITY NOTIFICATION%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%ANNUAL CONFIRMATION OF PENSIONER%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%ONLINE INQUIRY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SALARY LOAN%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%TOTAL DISABILITY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%PARTIAL DISABILITY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%PENSION MAINTENANCE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%KIOSK FEEDBACK%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%WEBSITE FEEDBACK%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS FUNERAL%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CITIZENS CHARTER%' GROUP by SSTRANSAT.ENCODE_DT,SSINFOTERMPROCCD.PROC_NM,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN")
                Dim rpt As New OnlineInquiryall
                cryRpt = rpt
                'cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date12)
                cryRpt.SetParameterValue("@DateTo", date22)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf rbOnlineInquiry.Checked = True And rbBranch.Checked = True Then
                GetNameBranch()
                GetName = "ONLINE INQUIRY REPORT SUMMARY"
                ''dt = db.ExecuteSQLQuery("SELECT SSTRANSAT.ENCODE_DT,case when SSINFOTERMPROCCD.PROC_NM = 'SSID CLEARANCE' then 'SSS/UMID CARD INFO' else SSINFOTERMPROCCD.PROC_NM end as 'PROC',SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN, COUNT(SSINFOTERMPROCCD.PROC_NM) as [count] FROM SSTRANSAT INNER JOIN SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAT.BRANCH_IP INNER JOIN SSINFOTERMPROCCD ON SSINFOTERMPROCCD.PROC_CD = SSTRANSAT.PROC_CD where SSTRANSAT.ENCODE_DT BETWEEN '" & date12 & "' and '" & date22 & "' and SSINFOTERMPROCCD.PROC_NM not LIKE 'WEBSITE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MY.SSS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOGIN%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%REGISTRATION%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBERSHIP%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOANS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFITS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CORPORATE PROFILE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%PUBLICATIONS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%OTHER SERVICES%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%HOME%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - SS SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - EC SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS - SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%EC - SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CHECKLIST OF DOCUMENTS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER MATERNITY NOTIFICATION%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%ANNUAL CONFIRMATION OF PENSIONER%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%ONLINE INQUIRY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SALARY LOAN%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%TOTAL DISABILITY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%PARTIAL DISABILITY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%PENSION MAINTENANCE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%KIOSK FEEDBACK%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%WEBSITE FEEDBACK%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS FUNERAL%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CITIZENS CHARTER%' GROUP by SSTRANSAT.ENCODE_DT,SSINFOTERMPROCCD.PROC_NM,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN")
                Dim rpt As New OnlineInquiryall
                cryRpt = rpt
                'cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date12)
                cryRpt.SetParameterValue("@DateTo", date22)
                cryRpt.SetParameterValue("@getName", getBranch)

            ElseIf rbOnlineInquiry.Checked = True And rbCluster.Checked = True Then
                GetNameBranch()
                GetName = "ONLINE INQUIRY REPORT SUMMARY"
                ''dt = db.ExecuteSQLQuery("SELECT SSTRANSAT.ENCODE_DT,case when SSINFOTERMPROCCD.PROC_NM = 'SSID CLEARANCE' then 'SSS/UMID CARD INFO' else SSINFOTERMPROCCD.PROC_NM end as 'PROC',SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN, COUNT(SSINFOTERMPROCCD.PROC_NM) as [count] FROM SSTRANSAT INNER JOIN SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAT.BRANCH_IP INNER JOIN SSINFOTERMPROCCD ON SSINFOTERMPROCCD.PROC_CD = SSTRANSAT.PROC_CD where SSTRANSAT.ENCODE_DT BETWEEN '" & date12 & "' and '" & date22 & "' and SSINFOTERMPROCCD.PROC_NM not LIKE 'WEBSITE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MY.SSS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOGIN%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%REGISTRATION%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBERSHIP%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOANS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFITS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CORPORATE PROFILE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%PUBLICATIONS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%OTHER SERVICES%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%HOME%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - SS SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - EC SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS - SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%EC - SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CHECKLIST OF DOCUMENTS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER MATERNITY NOTIFICATION%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%ANNUAL CONFIRMATION OF PENSIONER%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%ONLINE INQUIRY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SALARY LOAN%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%TOTAL DISABILITY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%PARTIAL DISABILITY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%PENSION MAINTENANCE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%KIOSK FEEDBACK%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%WEBSITE FEEDBACK%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS FUNERAL%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CITIZENS CHARTER%' GROUP by SSTRANSAT.ENCODE_DT,SSINFOTERMPROCCD.PROC_NM,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN")
                Dim rpt As New OnlineInquiryall
                cryRpt = rpt
                'cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date12)
                cryRpt.SetParameterValue("@DateTo", date22)
                cryRpt.SetParameterValue("@getName", getBranch)

            ElseIf rbEligibility.Checked = True And rbAll.Checked = True Then

                If rbAll.Checked = True And day1 = "1" And day2 = "31" Or day1 = "1" And day2 = "30" Then
                    getMonth()
                    GetNameBranch()
                    GetName = "ELIGIBILITY REPORT SUMMARY"
                    'dt = db.ExecuteSQLQuery("SELECT SSTRANSAT.ENCODE_DT,SSINFOTERMPROCCD.PROC_NM,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN, COUNT(SSINFOTERMPROCCD.PROC_NM) as count FROM SSTRANSAT INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAT.BRANCH_IP INNER JOIN SSINFOTERMPROCCD ON SSINFOTERMPROCCD.PROC_CD = SSTRANSAT.PROC_CD where SSTRANSAT.ENCODE_DT BETWEEN '" & date12 & "' and '" & date22 & "' and SSINFOTERMPROCCD.PROC_NM not LIKE 'WEBSITE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MY.SSS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOGIN%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%REGISTRATION%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBERSHIP%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOANS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFITS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CORPORATE PROFILE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%PUBLICATIONS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%OTHER SERVICES%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%HOME%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - SS SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - EC SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS - SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%EC - SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CHECKLIST OF DOCUMENTS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER MATERNITY NOTIFICATION%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%ANNUAL CONFIRMATION OF PENSIONER%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%ONLINE INQUIRY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SALARY LOAN%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%TOTAL DISABILITY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%PARTIAL DISABILITY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%PENSION MAINTENANCE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%KIOSK FEEDBACK%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%WEBSITE FEEDBACK%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS FUNERAL%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%ACTUAL PREMIUMS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%EMPLOYMENT HISTORY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%FLEXI-FUND%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOAN STATUS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MATERNITY CLAIMS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER DETAILS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SICKNESS CLAIMS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%TECHNICAL RETIREMENT%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CITIZENS CHARTER%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSID CLEARANCE%' GROUP by SSTRANSAT.ENCODE_DT,SSINFOTERMPROCCD.PROC_NM,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN,SSTRANSAT.BRANCH_IP")
                    Dim rpt As New Eligibilityall
                    cryRpt = rpt
                    cryRpt.Refresh()
                    cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                    rptView.ReportSource = cryRpt
                    cryRpt.SetParameterValue("@DateFrom", date12)
                    cryRpt.SetParameterValue("@DateTo", date22)
                    cryRpt.SetParameterValue("@getName", getBranch)
                Else
                    GetNameBranch()
                    GetName = "ELIGIBILITY REPORT SUMMARY"
                    'dt = db.ExecuteSQLQuery("SELECT SSTRANSAT.ENCODE_DT,SSINFOTERMPROCCD.PROC_NM,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN, COUNT(SSINFOTERMPROCCD.PROC_NM) as count FROM SSTRANSAT INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAT.BRANCH_IP INNER JOIN SSINFOTERMPROCCD ON SSINFOTERMPROCCD.PROC_CD = SSTRANSAT.PROC_CD where SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and SSINFOTERMPROCCD.PROC_NM not LIKE 'WEBSITE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MY.SSS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOGIN%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%REGISTRATION%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBERSHIP%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOANS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFITS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CORPORATE PROFILE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%PUBLICATIONS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%OTHER SERVICES%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%HOME%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - SS SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - EC SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS - SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%EC - SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CHECKLIST OF DOCUMENTS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER MATERNITY NOTIFICATION%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%ANNUAL CONFIRMATION OF PENSIONER%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%ONLINE INQUIRY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SALARY LOAN%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%TOTAL DISABILITY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%PARTIAL DISABILITY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%PENSION MAINTENANCE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%KIOSK FEEDBACK%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%WEBSITE FEEDBACK%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS FUNERAL%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%ACTUAL PREMIUMS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%EMPLOYMENT HISTORY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%FLEXI-FUND%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOAN STATUS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MATERNITY CLAIMS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER DETAILS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SICKNESS CLAIMS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%TECHNICAL RETIREMENT%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CITIZENS CHARTER%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSID CLEARANCE%' GROUP by SSTRANSAT.ENCODE_DT,SSINFOTERMPROCCD.PROC_NM,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN,SSTRANSAT.BRANCH_IP")
                    Dim rpt As New Eligibilityall
                    cryRpt = rpt
                    '    cryRpt.Refresh()
                    cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                    rptView.ReportSource = cryRpt
                    cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                    cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                    cryRpt.SetParameterValue("@getName", getBranch)
                End If

            ElseIf rbEligibility.Checked = True And rbGroup.Checked = True Then
                GetNameBranch()
                GetName = "ELIGIBILITY REPORT SUMMARY"
                'dt = db.ExecuteSQLQuery("SELECT SSTRANSAT.ENCODE_DT,SSINFOTERMPROCCD.PROC_NM,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN, COUNT(SSINFOTERMPROCCD.PROC_NM) as count FROM SSTRANSAT INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAT.BRANCH_IP INNER JOIN SSINFOTERMPROCCD ON SSINFOTERMPROCCD.PROC_CD = SSTRANSAT.PROC_CD where SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and SSINFOTERMPROCCD.PROC_NM not LIKE 'WEBSITE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MY.SSS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOGIN%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%REGISTRATION%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBERSHIP%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOANS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFITS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CORPORATE PROFILE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%PUBLICATIONS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%OTHER SERVICES%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%HOME%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - SS SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - EC SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS - SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%EC - SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CHECKLIST OF DOCUMENTS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER MATERNITY NOTIFICATION%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%ANNUAL CONFIRMATION OF PENSIONER%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%ONLINE INQUIRY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SALARY LOAN%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%TOTAL DISABILITY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%PARTIAL DISABILITY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%PENSION MAINTENANCE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%KIOSK FEEDBACK%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%WEBSITE FEEDBACK%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS FUNERAL%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%ACTUAL PREMIUMS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%EMPLOYMENT HISTORY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%FLEXI-FUND%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOAN STATUS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MATERNITY CLAIMS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER DETAILS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SICKNESS CLAIMS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%TECHNICAL RETIREMENT%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CITIZENS CHARTER%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSID CLEARANCE%' GROUP by SSTRANSAT.ENCODE_DT,SSINFOTERMPROCCD.PROC_NM,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN,SSTRANSAT.BRANCH_IP")
                Dim rpt As New Eligibilityall
                cryRpt = rpt
                '    cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)

            ElseIf rbEligibility.Checked = True And rbCluster.Checked = True Then
                GetNameBranch()
                GetName = "ELIGIBILITY REPORT SUMMARY"
                'dt = db.ExecuteSQLQuery("SELECT SSTRANSAT.ENCODE_DT,SSINFOTERMPROCCD.PROC_NM,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN, COUNT(SSINFOTERMPROCCD.PROC_NM) as count FROM SSTRANSAT INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAT.BRANCH_IP INNER JOIN SSINFOTERMPROCCD ON SSINFOTERMPROCCD.PROC_CD = SSTRANSAT.PROC_CD where SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and SSINFOTERMPROCCD.PROC_NM not LIKE 'WEBSITE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MY.SSS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOGIN%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%REGISTRATION%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBERSHIP%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOANS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFITS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CORPORATE PROFILE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%PUBLICATIONS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%OTHER SERVICES%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%HOME%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - SS SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - EC SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS - SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%EC - SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CHECKLIST OF DOCUMENTS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER MATERNITY NOTIFICATION%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%ANNUAL CONFIRMATION OF PENSIONER%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%ONLINE INQUIRY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SALARY LOAN%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%TOTAL DISABILITY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%PARTIAL DISABILITY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%PENSION MAINTENANCE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%KIOSK FEEDBACK%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%WEBSITE FEEDBACK%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS FUNERAL%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%ACTUAL PREMIUMS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%EMPLOYMENT HISTORY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%FLEXI-FUND%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOAN STATUS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MATERNITY CLAIMS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER DETAILS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SICKNESS CLAIMS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%TECHNICAL RETIREMENT%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CITIZENS CHARTER%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSID CLEARANCE%' GROUP by SSTRANSAT.ENCODE_DT,SSINFOTERMPROCCD.PROC_NM,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN,SSTRANSAT.BRANCH_IP")
                Dim rpt As New Eligibilityall
                cryRpt = rpt
                '    cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)

            ElseIf rbEligibility.Checked = True And rbBranch.Checked = True Then
                GetNameBranch()
                GetName = "ELIGIBILITY REPORT SUMMARY"
                'dt = db.ExecuteSQLQuery("SELECT SSTRANSAT.ENCODE_DT,SSINFOTERMPROCCD.PROC_NM,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN, COUNT(SSINFOTERMPROCCD.PROC_NM) as count FROM SSTRANSAT INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAT.BRANCH_IP INNER JOIN SSINFOTERMPROCCD ON SSINFOTERMPROCCD.PROC_CD = SSTRANSAT.PROC_CD where SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and SSINFOTERMPROCCD.PROC_NM not LIKE 'WEBSITE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MY.SSS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOGIN%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%REGISTRATION%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBERSHIP%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOANS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFITS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CORPORATE PROFILE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%PUBLICATIONS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%OTHER SERVICES%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%HOME%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - SS SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - EC SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS - SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%EC - SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CHECKLIST OF DOCUMENTS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER MATERNITY NOTIFICATION%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%ANNUAL CONFIRMATION OF PENSIONER%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%ONLINE INQUIRY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SALARY LOAN%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%TOTAL DISABILITY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%PARTIAL DISABILITY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%PENSION MAINTENANCE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%KIOSK FEEDBACK%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%WEBSITE FEEDBACK%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS FUNERAL%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%ACTUAL PREMIUMS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%EMPLOYMENT HISTORY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%FLEXI-FUND%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOAN STATUS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MATERNITY CLAIMS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER DETAILS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SICKNESS CLAIMS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%TECHNICAL RETIREMENT%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CITIZENS CHARTER%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSID CLEARANCE%' GROUP by SSTRANSAT.ENCODE_DT,SSINFOTERMPROCCD.PROC_NM,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN,SSTRANSAT.BRANCH_IP")
                Dim rpt As New Eligibilityall
                cryRpt = rpt
                '    cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)

            ElseIf rbPRN.Checked = True And rbAll.Checked = True Then

                If rbAll.Checked = True And day1 = "1" And day2 = "31" Or day1 = "1" And day2 = "30" Then
                    getMonth()
                    GetNameBranch()
                    GetName = "PRN REPORT SUMMARY"
                    Dim rpt As New PRNall
                    cryRpt = rpt
                    cryRpt.Refresh()
                    cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                    rptView.ReportSource = cryRpt
                    cryRpt.SetParameterValue("@DateFrom", date12)
                    cryRpt.SetParameterValue("@DateTo", date22)
                    cryRpt.SetParameterValue("@getName", getBranch)
                Else
                    GetNameBranch()
                    GetName = "PRN REPORT SUMMARY"

                    Dim rpt As New PRNall
                    cryRpt = rpt
                    cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                    rptView.ReportSource = cryRpt
                    cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                    cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                    cryRpt.SetParameterValue("@getName", getBranch)
                End If
            ElseIf rbPRN.Checked = True And rbGroup.Checked = True Then
                GetNameBranch()
                GetName = "PRN REPORT SUMMARY"
                Dim rpt As New PRNall
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)

            ElseIf rbPRN.Checked = True And rbCluster.Checked = True Then
                GetNameBranch()
                GetName = "PRN REPORT SUMMARY"
                Dim rpt As New PRNall
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)

            ElseIf rbPRN.Checked = True And rbBranch.Checked = True Then
                GetNameBranch()
                GetName = "PRN REPORT SUMMARY"
                Dim rpt As New PRNall
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)

            ElseIf rbDDR.Checked = True And rbAll.Checked = True Then

                If day1 = "1" And day2 = "31" Or day1 = "1" And day2 = "30" Then
                    getMonth()
                    GetNameBranch()
                    GetName = "DDR FUNERAL REPORT SUMMARY"
                    'dt = db.ExecuteSQLQuery("SELECT SSTRANSAT.ENCODE_DT,SSINFOTERMPROCCD.PROC_NM,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN, COUNT(SSINFOTERMPROCCD.PROC_NM) as count FROM SSTRANSAT INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAT.BRANCH_IP INNER JOIN SSINFOTERMPROCCD ON SSINFOTERMPROCCD.PROC_CD = SSTRANSAT.PROC_CD where SSTRANSAT.ENCODE_DT BETWEEN '" & date12 & "' and '" & date22 & "' and SSINFOTERMPROCCD.PROC_NM not LIKE 'WEBSITE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MY.SSS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOGIN%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%REGISTRATION%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBERSHIP%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOANS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFITS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CORPORATE PROFILE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%PUBLICATIONS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%OTHER SERVICES%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%HOME%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - SS SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - EC SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS - SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%EC - SICKNESS%'  and SSINFOTERMPROCCD.PROC_NM not LIKE '%CHECKLIST OF DOCUMENTS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER MATERNITY NOTIFICATION%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%ANNUAL CONFIRMATION OF PENSIONER%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%ONLINE INQUIRY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SALARY LOAN%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%DDR FUNERAL%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SICKNESS/MATERNITY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%PENSION MAINTENANCE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%KIOSK FEEDBACK%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%WEBSITE FEEDBACK%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOAN ELIGIBILITY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%ACTUAL PREMIUMS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%EMPLOYMENT HISTORY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%FLEXI-FUND%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOAN STATUS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MATERNITY CLAIMS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER DETAILS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SICKNESS CLAIMS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%TECHNICAL RETIREMENT%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSID CLEARANCE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFIT CLAIMS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - MATERNITY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CITIZENS CHARTER%' GROUP by SSTRANSAT.ENCODE_DT,SSINFOTERMPROCCD.PROC_NM,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN,SSTRANSAT.BRANCH_IP")
                    Dim rpt As New DDRFuneralAll
                    cryRpt = rpt
                    '    cryRpt.Refresh()
                    cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                    rptView.ReportSource = cryRpt
                    cryRpt.SetParameterValue("@DateFrom", date12)
                    cryRpt.SetParameterValue("@DateTo", date22)
                    cryRpt.SetParameterValue("@getName", getBranch)
                Else
                    GetNameBranch()
                    GetName = "DDR FUNERAL REPORT SUMMARY"
                    'dt = db.ExecuteSQLQuery("SELECT convert(varchar,SSTRANSAT.ENCODE_DT,101) as DateOfTrans ,SSINFOTERMPROCCD.PROC_NM,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN, COUNT(SSINFOTERMPROCCD.PROC_NM) as count FROM SSTRANSAT INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAT.BRANCH_IP INNER JOIN SSINFOTERMPROCCD ON SSINFOTERMPROCCD.PROC_CD = SSTRANSAT.PROC_CD where SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date1.ToShortDateString & "' and SSINFOTERMPROCCD.PROC_NM not LIKE 'WEBSITE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MY.SSS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOGIN%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%REGISTRATION%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBERSHIP%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOANS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFITS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CORPORATE PROFILE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%PUBLICATIONS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%OTHER SERVICES%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%HOME%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - SS SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - EC SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS - SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%EC - SICKNESS%'  and SSINFOTERMPROCCD.PROC_NM not LIKE '%CHECKLIST OF DOCUMENTS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER MATERNITY NOTIFICATION%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%ANNUAL CONFIRMATION OF PENSIONER%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%ONLINE INQUIRY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SALARY LOAN%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%DDR FUNERAL%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SICKNESS/MATERNITY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%PENSION MAINTENANCE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%KIOSK FEEDBACK%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%WEBSITE FEEDBACK%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOAN ELIGIBILITY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%ACTUAL PREMIUMS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%EMPLOYMENT HISTORY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%FLEXI-FUND%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOAN STATUS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MATERNITY CLAIMS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER DETAILS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SICKNESS CLAIMS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%TECHNICAL RETIREMENT%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSID CLEARANCE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFIT CLAIMS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - MATERNITY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CITIZENS CHARTER%' GROUP by SSTRANSAT.ENCODE_DT,SSINFOTERMPROCCD.PROC_NM,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN,SSTRANSAT.BRANCH_IP")
                    Dim rpt As New DDRFuneralAll
                    cryRpt = rpt
                    '    cryRpt.Refresh()
                    cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                    rptView.ReportSource = cryRpt
                    cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                    cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                    cryRpt.SetParameterValue("@getName", getBranch)
                End If
            ElseIf rbDDR.Checked = True And rbGroup.Checked = True Then
                GetNameBranch()
                GetName = "DDR FUNERAL REPORT SUMMARY"
                'dt = db.ExecuteSQLQuery("SELECT convert(varchar,SSTRANSAT.ENCODE_DT,101) as DateOfTrans ,SSINFOTERMPROCCD.PROC_NM,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN, COUNT(SSINFOTERMPROCCD.PROC_NM) as count FROM SSTRANSAT INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAT.BRANCH_IP INNER JOIN SSINFOTERMPROCCD ON SSINFOTERMPROCCD.PROC_CD = SSTRANSAT.PROC_CD where SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date1.ToShortDateString & "' and SSINFOTERMPROCCD.PROC_NM not LIKE 'WEBSITE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MY.SSS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOGIN%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%REGISTRATION%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBERSHIP%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOANS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFITS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CORPORATE PROFILE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%PUBLICATIONS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%OTHER SERVICES%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%HOME%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - SS SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - EC SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS - SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%EC - SICKNESS%'  and SSINFOTERMPROCCD.PROC_NM not LIKE '%CHECKLIST OF DOCUMENTS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER MATERNITY NOTIFICATION%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%ANNUAL CONFIRMATION OF PENSIONER%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%ONLINE INQUIRY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SALARY LOAN%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%DDR FUNERAL%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SICKNESS/MATERNITY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%PENSION MAINTENANCE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%KIOSK FEEDBACK%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%WEBSITE FEEDBACK%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOAN ELIGIBILITY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%ACTUAL PREMIUMS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%EMPLOYMENT HISTORY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%FLEXI-FUND%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOAN STATUS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MATERNITY CLAIMS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER DETAILS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SICKNESS CLAIMS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%TECHNICAL RETIREMENT%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSID CLEARANCE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFIT CLAIMS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - MATERNITY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CITIZENS CHARTER%' GROUP by SSTRANSAT.ENCODE_DT,SSINFOTERMPROCCD.PROC_NM,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN,SSTRANSAT.BRANCH_IP")
                Dim rpt As New DDRFuneralAll
                cryRpt = rpt
                '    cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf rbDDR.Checked = True And rbBranch.Checked = True Then
                GetNameBranch()
                GetName = "DDR FUNERAL REPORT SUMMARY"
                'dt = db.ExecuteSQLQuery("SELECT convert(varchar,SSTRANSAT.ENCODE_DT,101) as DateOfTrans ,SSINFOTERMPROCCD.PROC_NM,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN, COUNT(SSINFOTERMPROCCD.PROC_NM) as count FROM SSTRANSAT INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAT.BRANCH_IP INNER JOIN SSINFOTERMPROCCD ON SSINFOTERMPROCCD.PROC_CD = SSTRANSAT.PROC_CD where SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date1.ToShortDateString & "' and SSINFOTERMPROCCD.PROC_NM not LIKE 'WEBSITE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MY.SSS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOGIN%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%REGISTRATION%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBERSHIP%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOANS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFITS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CORPORATE PROFILE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%PUBLICATIONS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%OTHER SERVICES%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%HOME%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - SS SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - EC SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS - SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%EC - SICKNESS%'  and SSINFOTERMPROCCD.PROC_NM not LIKE '%CHECKLIST OF DOCUMENTS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER MATERNITY NOTIFICATION%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%ANNUAL CONFIRMATION OF PENSIONER%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%ONLINE INQUIRY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SALARY LOAN%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%DDR FUNERAL%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SICKNESS/MATERNITY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%PENSION MAINTENANCE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%KIOSK FEEDBACK%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%WEBSITE FEEDBACK%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOAN ELIGIBILITY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%ACTUAL PREMIUMS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%EMPLOYMENT HISTORY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%FLEXI-FUND%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOAN STATUS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MATERNITY CLAIMS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER DETAILS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SICKNESS CLAIMS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%TECHNICAL RETIREMENT%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSID CLEARANCE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFIT CLAIMS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - MATERNITY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CITIZENS CHARTER%' GROUP by SSTRANSAT.ENCODE_DT,SSINFOTERMPROCCD.PROC_NM,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN,SSTRANSAT.BRANCH_IP")
                Dim rpt As New DDRFuneralAll
                cryRpt = rpt
                '    cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)

            ElseIf rbDDR.Checked = True And rbCluster.Checked = True Then
                GetNameBranch()
                GetName = "DDR FUNERAL REPORT SUMMARY"
                'dt = db.ExecuteSQLQuery("SELECT convert(varchar,SSTRANSAT.ENCODE_DT,101) as DateOfTrans ,SSINFOTERMPROCCD.PROC_NM,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN, COUNT(SSINFOTERMPROCCD.PROC_NM) as count FROM SSTRANSAT INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAT.BRANCH_IP INNER JOIN SSINFOTERMPROCCD ON SSINFOTERMPROCCD.PROC_CD = SSTRANSAT.PROC_CD where SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date1.ToShortDateString & "' and SSINFOTERMPROCCD.PROC_NM not LIKE 'WEBSITE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MY.SSS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOGIN%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%REGISTRATION%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBERSHIP%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOANS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFITS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CORPORATE PROFILE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%PUBLICATIONS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%OTHER SERVICES%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%HOME%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - SS SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - EC SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS - SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%EC - SICKNESS%'  and SSINFOTERMPROCCD.PROC_NM not LIKE '%CHECKLIST OF DOCUMENTS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER MATERNITY NOTIFICATION%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%ANNUAL CONFIRMATION OF PENSIONER%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%ONLINE INQUIRY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SALARY LOAN%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%DDR FUNERAL%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SICKNESS/MATERNITY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%PENSION MAINTENANCE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%KIOSK FEEDBACK%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%WEBSITE FEEDBACK%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOAN ELIGIBILITY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%ACTUAL PREMIUMS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%EMPLOYMENT HISTORY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%FLEXI-FUND%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOAN STATUS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MATERNITY CLAIMS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER DETAILS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SICKNESS CLAIMS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%TECHNICAL RETIREMENT%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSID CLEARANCE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFIT CLAIMS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - MATERNITY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CITIZENS CHARTER%' GROUP by SSTRANSAT.ENCODE_DT,SSINFOTERMPROCCD.PROC_NM,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN,SSTRANSAT.BRANCH_IP")
                Dim rpt As New DDRFuneralAll
                cryRpt = rpt
                '    cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)

            ElseIf rbSummarySSS.Checked = True And rbAll.Checked = True And cbCard.Text = "Show All" Then
                GetName = "SSID TRANSACTION REPORT"
                'TransValidation()
                GetTrans()
                GetNameBranch()
                ''dt = db.ExecuteSQLQuery("select SSINFOTERMACCESS.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,CAST(TRANS_DT as date)[date],sum(case when TRANS_TYPE != 'ONLINE' AND TRANS_TYPE = 'UMIDCARD' and REF_NUM <> '' then 1 else 0 end) SUCCESSUMID,sum(case when TRANS_TYPE != 'ONLINE' AND TRANS_TYPE = 'OLDSSSCARD' and REF_NUM <> '' then 1 else 0 end) SUCCESSOLD, (select count(*) from SSTRANSERRORLOGS where CARDTYPE = 'UMIDCARD'and CAST(TRANS_DT as date) = CAST(SSTRANSERRORLOGS.ENCODE_DT as date)) as[failedUmid], (select count(*) from SSTRANSERRORLOGS where CARDTYPE = 'OLDSSSCARD' and CAST(TRANS_DT as date) = CAST(SSTRANSERRORLOGS.ENCODE_DT as date)) as[failedOLD] from SSINFOTERMACCESS LEFT JOIN SSTRANSERRORLOGS ON SSINFOTERMACCESS.KIOSK_ID = SSTRANSERRORLOGS.KIOSK_ID INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID = SSINFOTERMACCESS.KIOSK_ID INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSINFOTERMKIOSK.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMKIOSK.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSINFOTERMKIOSK.DIVSN where TRANS_TYPE not LIKE '%ONLINE%' and CAST(TRANS_DT as date) BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' group by SSINFOTERMACCESS.KIOSK_ID,CAST(TRANS_DT as date),SSINFOTERMBR.BRANCH_NM,SSINFOTERMKIOSK.BRANCH_IP,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM ORDER BY CAST(TRANS_DT as date)")
                Dim rpt As New SSIDCARDALL
                cryRpt = rpt
                '  cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf rbSummarySSS.Checked = True And rbGroup.Checked = True And cbCard.Text = "Show All" Then
                GetName = "SSID TRANSACTION REPORT"
                'TransValidation()
                GetTrans()
                GetNameBranch()
                ''dt = db.ExecuteSQLQuery("select SSINFOTERMACCESS.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,CAST(TRANS_DT as date)[date],sum(case when TRANS_TYPE != 'ONLINE' AND TRANS_TYPE = 'UMIDCARD' and REF_NUM <> '' then 1 else 0 end) SUCCESSUMID,sum(case when TRANS_TYPE != 'ONLINE' AND TRANS_TYPE = 'OLDSSSCARD' and REF_NUM <> '' then 1 else 0 end) SUCCESSOLD, (select count(*) from SSTRANSERRORLOGS where CARDTYPE = 'UMIDCARD'and CAST(TRANS_DT as date) = CAST(SSTRANSERRORLOGS.ENCODE_DT as date)) as[failedUmid], (select count(*) from SSTRANSERRORLOGS where CARDTYPE = 'OLDSSSCARD' and CAST(TRANS_DT as date) = CAST(SSTRANSERRORLOGS.ENCODE_DT as date)) as[failedOLD] from SSINFOTERMACCESS LEFT JOIN SSTRANSERRORLOGS ON SSINFOTERMACCESS.KIOSK_ID = SSTRANSERRORLOGS.KIOSK_ID INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID = SSINFOTERMACCESS.KIOSK_ID INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSINFOTERMKIOSK.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMKIOSK.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSINFOTERMKIOSK.DIVSN where TRANS_TYPE not LIKE '%ONLINE%' and CAST(TRANS_DT as date) BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' group by SSINFOTERMACCESS.KIOSK_ID,CAST(TRANS_DT as date),SSINFOTERMBR.BRANCH_NM,SSINFOTERMKIOSK.BRANCH_IP,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM ORDER BY CAST(TRANS_DT as date)")
                Dim rpt As New SSIDCARDALL
                cryRpt = rpt
                '  cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf rbSummarySSS.Checked = True And rbCluster.Checked = True And cbCard.Text = "Show All" Then
                GetName = "SSID TRANSACTION REPORT"
                'TransValidation()
                GetTrans()
                GetNameBranch()
                ''dt = db.ExecuteSQLQuery("select SSINFOTERMACCESS.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,CAST(TRANS_DT as date)[date],sum(case when TRANS_TYPE != 'ONLINE' AND TRANS_TYPE = 'UMIDCARD' and REF_NUM <> '' then 1 else 0 end) SUCCESSUMID,sum(case when TRANS_TYPE != 'ONLINE' AND TRANS_TYPE = 'OLDSSSCARD' and REF_NUM <> '' then 1 else 0 end) SUCCESSOLD, (select count(*) from SSTRANSERRORLOGS where CARDTYPE = 'UMIDCARD'and CAST(TRANS_DT as date) = CAST(SSTRANSERRORLOGS.ENCODE_DT as date)) as[failedUmid], (select count(*) from SSTRANSERRORLOGS where CARDTYPE = 'OLDSSSCARD' and CAST(TRANS_DT as date) = CAST(SSTRANSERRORLOGS.ENCODE_DT as date)) as[failedOLD] from SSINFOTERMACCESS LEFT JOIN SSTRANSERRORLOGS ON SSINFOTERMACCESS.KIOSK_ID = SSTRANSERRORLOGS.KIOSK_ID INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID = SSINFOTERMACCESS.KIOSK_ID INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSINFOTERMKIOSK.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMKIOSK.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSINFOTERMKIOSK.DIVSN where TRANS_TYPE not LIKE '%ONLINE%' and CAST(TRANS_DT as date) BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' group by SSINFOTERMACCESS.KIOSK_ID,CAST(TRANS_DT as date),SSINFOTERMBR.BRANCH_NM,SSINFOTERMKIOSK.BRANCH_IP,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM ORDER BY CAST(TRANS_DT as date)")
                Dim rpt As New SSIDCARDALL
                cryRpt = rpt
                '  cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf rbSummarySSS.Checked = True And rbBranch.Checked = True And cbCard.Text = "Show All" Then
                GetName = "SSID TRANSACTION REPORT"
                'TransValidation()
                GetTrans()
                GetNameBranch()
                ''dt = db.ExecuteSQLQuery("select SSINFOTERMACCESS.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,CAST(TRANS_DT as date)[date],sum(case when TRANS_TYPE != 'ONLINE' AND TRANS_TYPE = 'UMIDCARD' and REF_NUM <> '' then 1 else 0 end) SUCCESSUMID,sum(case when TRANS_TYPE != 'ONLINE' AND TRANS_TYPE = 'OLDSSSCARD' and REF_NUM <> '' then 1 else 0 end) SUCCESSOLD, (select count(*) from SSTRANSERRORLOGS where CARDTYPE = 'UMIDCARD'and CAST(TRANS_DT as date) = CAST(SSTRANSERRORLOGS.ENCODE_DT as date)) as[failedUmid], (select count(*) from SSTRANSERRORLOGS where CARDTYPE = 'OLDSSSCARD' and CAST(TRANS_DT as date) = CAST(SSTRANSERRORLOGS.ENCODE_DT as date)) as[failedOLD] from SSINFOTERMACCESS LEFT JOIN SSTRANSERRORLOGS ON SSINFOTERMACCESS.KIOSK_ID = SSTRANSERRORLOGS.KIOSK_ID INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID = SSINFOTERMACCESS.KIOSK_ID INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSINFOTERMKIOSK.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMKIOSK.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSINFOTERMKIOSK.DIVSN where TRANS_TYPE not LIKE '%ONLINE%' and CAST(TRANS_DT as date) BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' group by SSINFOTERMACCESS.KIOSK_ID,CAST(TRANS_DT as date),SSINFOTERMBR.BRANCH_NM,SSINFOTERMKIOSK.BRANCH_IP,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM ORDER BY CAST(TRANS_DT as date)")
                Dim rpt As New SSIDCARDALL
                cryRpt = rpt
                '  cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf rbSummarySSS.Checked = True And rbAll.Checked = True And SSID = 1 Then
                'TransValidation()
                GetTrans()
                GetNameBranch()
                GetName = "SSID TRANSACTION"
                ' 'dt = db.ExecuteSQLQuery("select SSINFOTERMACCESS.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,CAST(TRANS_DT as date)[date],CASE WHEN TRANS_TYPE = 'UMIDCARD' THEN 'UMID CARD' else 'SSS CARD' end as [TRANS_TYPE], sum(case when TRANS_TYPE != 'ONLINE' AND TRANS_TYPE = '" & cbCard.Text & "' and REF_NUM <> '' then 1 else 0 end) SUCCESS, (select count(*) from SSTRANSERRORLOGS where CARDTYPE = '" & cbCard.Text & "' and CAST(TRANS_DT as date) = CAST(SSTRANSERRORLOGS.ENCODE_DT as date) and TRANS_TYPE = CARDTYPE) as [FAILED] from SSINFOTERMACCESS LEFT JOIN SSTRANSERRORLOGS ON SSINFOTERMACCESS.KIOSK_ID = SSTRANSERRORLOGS.KIOSK_ID INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID = SSINFOTERMACCESS.KIOSK_ID INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSINFOTERMKIOSK.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMKIOSK.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSINFOTERMKIOSK.DIVSN where TRANS_TYPE not LIKE '%ONLINE%' and TRANS_TYPE = '" & cbCard.Text & "' and CAST(TRANS_DT as date) BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' group by SSINFOTERMACCESS.KIOSK_ID,CAST(TRANS_DT as date),SSINFOTERMBR.BRANCH_NM,SSINFOTERMKIOSK.BRANCH_IP,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,TRANS_TYPE ORDER BY CAST(TRANS_DT as date)")
                Dim rpt As New SSSIDCARD
                cryRpt = rpt
                '   cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@card_type", TransType)
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf rbSummarySSS.Checked = True And rbGroup.Checked = True And SSID = 1 Then
                'TransValidation()
                GetTrans()
                GetNameBranch()
                GetName = "SSID TRANSACTION"
                ' 'dt = db.ExecuteSQLQuery("select SSINFOTERMACCESS.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,CAST(TRANS_DT as date)[date],CASE WHEN TRANS_TYPE = 'UMIDCARD' THEN 'UMID CARD' else 'SSS CARD' end as [TRANS_TYPE], sum(case when TRANS_TYPE != 'ONLINE' AND TRANS_TYPE = '" & cbCard.Text & "' and REF_NUM <> '' then 1 else 0 end) SUCCESS, (select count(*) from SSTRANSERRORLOGS where CARDTYPE = '" & cbCard.Text & "' and CAST(TRANS_DT as date) = CAST(SSTRANSERRORLOGS.ENCODE_DT as date) and TRANS_TYPE = CARDTYPE) as [FAILED] from SSINFOTERMACCESS LEFT JOIN SSTRANSERRORLOGS ON SSINFOTERMACCESS.KIOSK_ID = SSTRANSERRORLOGS.KIOSK_ID INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID = SSINFOTERMACCESS.KIOSK_ID INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSINFOTERMKIOSK.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMKIOSK.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSINFOTERMKIOSK.DIVSN where TRANS_TYPE not LIKE '%ONLINE%' and TRANS_TYPE = '" & cbCard.Text & "' and CAST(TRANS_DT as date) BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' group by SSINFOTERMACCESS.KIOSK_ID,CAST(TRANS_DT as date),SSINFOTERMBR.BRANCH_NM,SSINFOTERMKIOSK.BRANCH_IP,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,TRANS_TYPE ORDER BY CAST(TRANS_DT as date)")
                Dim rpt As New SSSIDCARD
                cryRpt = rpt
                '   cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@card_type", TransType)
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf rbSummarySSS.Checked = True And rbBranch.Checked = True And SSID = 1 Then
                'TransValidation()
                GetTrans()
                GetNameBranch()
                GetName = "SSID TRANSACTION"
                ' 'dt = db.ExecuteSQLQuery("select SSINFOTERMACCESS.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,CAST(TRANS_DT as date)[date],CASE WHEN TRANS_TYPE = 'UMIDCARD' THEN 'UMID CARD' else 'SSS CARD' end as [TRANS_TYPE], sum(case when TRANS_TYPE != 'ONLINE' AND TRANS_TYPE = '" & cbCard.Text & "' and REF_NUM <> '' then 1 else 0 end) SUCCESS, (select count(*) from SSTRANSERRORLOGS where CARDTYPE = '" & cbCard.Text & "' and CAST(TRANS_DT as date) = CAST(SSTRANSERRORLOGS.ENCODE_DT as date) and TRANS_TYPE = CARDTYPE) as [FAILED] from SSINFOTERMACCESS LEFT JOIN SSTRANSERRORLOGS ON SSINFOTERMACCESS.KIOSK_ID = SSTRANSERRORLOGS.KIOSK_ID INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID = SSINFOTERMACCESS.KIOSK_ID INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSINFOTERMKIOSK.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMKIOSK.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSINFOTERMKIOSK.DIVSN where TRANS_TYPE not LIKE '%ONLINE%' and TRANS_TYPE = '" & cbCard.Text & "' and CAST(TRANS_DT as date) BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' group by SSINFOTERMACCESS.KIOSK_ID,CAST(TRANS_DT as date),SSINFOTERMBR.BRANCH_NM,SSINFOTERMKIOSK.BRANCH_IP,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,TRANS_TYPE ORDER BY CAST(TRANS_DT as date)")
                Dim rpt As New SSSIDCARD
                cryRpt = rpt
                '   cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@card_type", TransType)
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)

            ElseIf rbSummarySSS.Checked = True And rbCluster.Checked = True And SSID = 1 Then
                'TransValidation()
                GetTrans()
                GetNameBranch()
                GetName = "SSID TRANSACTION"
                ' 'dt = db.ExecuteSQLQuery("select SSINFOTERMACCESS.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,CAST(TRANS_DT as date)[date],CASE WHEN TRANS_TYPE = 'UMIDCARD' THEN 'UMID CARD' else 'SSS CARD' end as [TRANS_TYPE], sum(case when TRANS_TYPE != 'ONLINE' AND TRANS_TYPE = '" & cbCard.Text & "' and REF_NUM <> '' then 1 else 0 end) SUCCESS, (select count(*) from SSTRANSERRORLOGS where CARDTYPE = '" & cbCard.Text & "' and CAST(TRANS_DT as date) = CAST(SSTRANSERRORLOGS.ENCODE_DT as date) and TRANS_TYPE = CARDTYPE) as [FAILED] from SSINFOTERMACCESS LEFT JOIN SSTRANSERRORLOGS ON SSINFOTERMACCESS.KIOSK_ID = SSTRANSERRORLOGS.KIOSK_ID INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID = SSINFOTERMACCESS.KIOSK_ID INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSINFOTERMKIOSK.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMKIOSK.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSINFOTERMKIOSK.DIVSN where TRANS_TYPE not LIKE '%ONLINE%' and TRANS_TYPE = '" & cbCard.Text & "' and CAST(TRANS_DT as date) BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' group by SSINFOTERMACCESS.KIOSK_ID,CAST(TRANS_DT as date),SSINFOTERMBR.BRANCH_NM,SSINFOTERMKIOSK.BRANCH_IP,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,TRANS_TYPE ORDER BY CAST(TRANS_DT as date)")
                Dim rpt As New SSSIDCARD
                cryRpt = rpt
                '   cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@card_type", TransType)
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)

            ElseIf cbSTrans.Text = "Retirement Pensioner" And rbGroup.Checked = True Then
                GetNameBranch()
                GetName = "Retirement Pensioner"

                'dt = db.ExecuteSQLQuery("SELECT SSNUM,DOBTH,DOCVRG,UMID_BANK,BANK_BR,TRANS_ACCTNO,STREET,BRGAY,POST_CD,EMAIL,PHONE,CELNO,USRTYP,SSTRANSINFOTERMTR.ENCODE_DT,SSTRANSINFOTERMTR.ENCODE_TME ,SSTRANSINFOTERMTR.BRANCH_IP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSINFOTERMTR.TAG,TRANID FROM SSTRANSINFOTERMTR INNER JOIN SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSTRANSINFOTERMTR.BRANCH_IP inner join SSINFOTERMBR on SSTRANSINFOTERMTR.BRANCH_CD = SSINFOTERMBR.BRANCH_CD inner join SSINFOTERMCLSTR on SSTRANSINFOTERMTR.CLSTR = SSINFOTERMCLSTR.CLSTR_CD INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSTRANSINFOTERMTR.DIVSN where SSTRANSINFOTERMTR.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "'")
                Dim rpt As New TECHRETREGALL
                cryRpt = rpt
                '  cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf cbSTrans.Text = "Retirement Pensioner" And rbBranch.Checked = True Then
                GetNameBranch()
                GetName = "Retirement Pensioner"

                'dt = db.ExecuteSQLQuery("SELECT SSNUM,DOBTH,DOCVRG,UMID_BANK,BANK_BR,TRANS_ACCTNO,STREET,BRGAY,POST_CD,EMAIL,PHONE,CELNO,USRTYP,SSTRANSINFOTERMTR.ENCODE_DT,SSTRANSINFOTERMTR.ENCODE_TME ,SSTRANSINFOTERMTR.BRANCH_IP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSINFOTERMTR.TAG,TRANID FROM SSTRANSINFOTERMTR INNER JOIN SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSTRANSINFOTERMTR.BRANCH_IP inner join SSINFOTERMBR on SSTRANSINFOTERMTR.BRANCH_CD = SSINFOTERMBR.BRANCH_CD inner join SSINFOTERMCLSTR on SSTRANSINFOTERMTR.CLSTR = SSINFOTERMCLSTR.CLSTR_CD INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSTRANSINFOTERMTR.DIVSN where SSTRANSINFOTERMTR.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "'")
                Dim rpt As New TECHRETREGALL
                cryRpt = rpt
                '  cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)

            ElseIf cbSTrans.Text = "Retirement Pensioner" And rbCluster.Checked = True Then
                GetNameBranch()
                GetName = "Retirement Pensioner"

                'dt = db.ExecuteSQLQuery("SELECT SSNUM,DOBTH,DOCVRG,UMID_BANK,BANK_BR,TRANS_ACCTNO,STREET,BRGAY,POST_CD,EMAIL,PHONE,CELNO,USRTYP,SSTRANSINFOTERMTR.ENCODE_DT,SSTRANSINFOTERMTR.ENCODE_TME ,SSTRANSINFOTERMTR.BRANCH_IP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSINFOTERMTR.TAG,TRANID FROM SSTRANSINFOTERMTR INNER JOIN SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSTRANSINFOTERMTR.BRANCH_IP inner join SSINFOTERMBR on SSTRANSINFOTERMTR.BRANCH_CD = SSINFOTERMBR.BRANCH_CD inner join SSINFOTERMCLSTR on SSTRANSINFOTERMTR.CLSTR = SSINFOTERMCLSTR.CLSTR_CD INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSTRANSINFOTERMTR.DIVSN where SSTRANSINFOTERMTR.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "'")
                Dim rpt As New TECHRETREGALL
                cryRpt = rpt
                '  cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)

            ElseIf cbSTrans.Text = "Retirement Pensioner" = True And rbAll.Checked = True Then
                GetNameBranch()
                GetName = "Retirement Pensioner"

                'dt = db.ExecuteSQLQuery("SELECT SSNUM,DOBTH,DOCVRG,UMID_BANK,BANK_BR,TRANS_ACCTNO,STREET,BRGAY,POST_CD,EMAIL,PHONE,CELNO,USRTYP,SSTRANSINFOTERMTR.ENCODE_DT,SSTRANSINFOTERMTR.ENCODE_TME ,SSTRANSINFOTERMTR.BRANCH_IP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSINFOTERMTR.TAG,TRANID FROM SSTRANSINFOTERMTR INNER JOIN SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSTRANSINFOTERMTR.BRANCH_IP inner join SSINFOTERMBR on SSTRANSINFOTERMTR.BRANCH_CD = SSINFOTERMBR.BRANCH_CD inner join SSINFOTERMCLSTR on SSTRANSINFOTERMTR.CLSTR = SSINFOTERMCLSTR.CLSTR_CD INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSTRANSINFOTERMTR.DIVSN where SSTRANSINFOTERMTR.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "'")
                Dim rpt As New TECHRETREGALL
                cryRpt = rpt
                '  cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf cbSTrans.Text = "Technical Retirement Lump Sum" And rbAll.Checked = True Then
                GetNameBranch()
                GetName = "Technical Retirment Lump Sum"
                'dt = db.ExecuteSQLQuery("SELECT SSNUM,TRANID,DOBTH,[TIMESTAMP],LUMPSUMAMT,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM, DATECREATED,SSTRANSINFOTERMTRLS.BRANCH_IP,SSINFOTERMKIOSK.KIOSK_ID from SSTRANSINFOTERMTRLS INNER JOIN SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSTRANSINFOTERMTRLS.BRANCH_IP inner join SSINFOTERMBR on SSTRANSINFOTERMTRLS.BRANCH_CD = SSINFOTERMBR.BRANCH_CD inner join SSINFOTERMCLSTR on  SSTRANSINFOTERMTRLS.CLUSTER = SSINFOTERMCLSTR.CLSTR_CD INNER JOIN SSINFOTERMGROUP  on SSINFOTERMGROUP.GROUP_CD = SSTRANSINFOTERMTRLS.DIVISION where DATECREATED BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "'")
                Dim rpt As New TECHRETLUMPALL
                cryRpt = rpt
                '  cryRpt.Refresh()

                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf cbSTrans.Text = "Technical Retirement Lump Sum" And rbBranch.Checked = True Then
                GetNameBranch()
                GetName = "Technical Retirment Lump Sum"
                'dt = db.ExecuteSQLQuery("SELECT SSNUM,TRANID,DOBTH,[TIMESTAMP],LUMPSUMAMT,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM, DATECREATED,SSTRANSINFOTERMTRLS.BRANCH_IP,SSINFOTERMKIOSK.KIOSK_ID from SSTRANSINFOTERMTRLS INNER JOIN SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSTRANSINFOTERMTRLS.BRANCH_IP inner join SSINFOTERMBR on SSTRANSINFOTERMTRLS.BRANCH_CD = SSINFOTERMBR.BRANCH_CD inner join SSINFOTERMCLSTR on  SSTRANSINFOTERMTRLS.CLUSTER = SSINFOTERMCLSTR.CLSTR_CD INNER JOIN SSINFOTERMGROUP  on SSINFOTERMGROUP.GROUP_CD = SSTRANSINFOTERMTRLS.DIVISION where DATECREATED BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "'")
                Dim rpt As New TECHRETLUMPALL
                cryRpt = rpt
                '  cryRpt.Refresh()

                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf cbSTrans.Text = "Technical Retirement Lump Sum" And rbGroup.Checked = True Then
                GetNameBranch()
                GetName = "Technical Retirment Lump Sum"
                'dt = db.ExecuteSQLQuery("SELECT SSNUM,TRANID,DOBTH,[TIMESTAMP],LUMPSUMAMT,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM, DATECREATED,SSTRANSINFOTERMTRLS.BRANCH_IP,SSINFOTERMKIOSK.KIOSK_ID from SSTRANSINFOTERMTRLS INNER JOIN SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSTRANSINFOTERMTRLS.BRANCH_IP inner join SSINFOTERMBR on SSTRANSINFOTERMTRLS.BRANCH_CD = SSINFOTERMBR.BRANCH_CD inner join SSINFOTERMCLSTR on  SSTRANSINFOTERMTRLS.CLUSTER = SSINFOTERMCLSTR.CLSTR_CD INNER JOIN SSINFOTERMGROUP  on SSINFOTERMGROUP.GROUP_CD = SSTRANSINFOTERMTRLS.DIVISION where DATECREATED BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "'")
                Dim rpt As New TECHRETLUMPALL
                cryRpt = rpt
                '  cryRpt.Refresh()

                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf cbSTrans.Text = "Technical Retirement Lump Sum" And rbCluster.Checked = True Then
                GetNameBranch()
                GetName = "Technical Retirment Lump Sum"
                'dt = db.ExecuteSQLQuery("SELECT SSNUM,TRANID,DOBTH,[TIMESTAMP],LUMPSUMAMT,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM, DATECREATED,SSTRANSINFOTERMTRLS.BRANCH_IP,SSINFOTERMKIOSK.KIOSK_ID from SSTRANSINFOTERMTRLS INNER JOIN SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSTRANSINFOTERMTRLS.BRANCH_IP inner join SSINFOTERMBR on SSTRANSINFOTERMTRLS.BRANCH_CD = SSINFOTERMBR.BRANCH_CD inner join SSINFOTERMCLSTR on  SSTRANSINFOTERMTRLS.CLUSTER = SSINFOTERMCLSTR.CLSTR_CD INNER JOIN SSINFOTERMGROUP  on SSINFOTERMGROUP.GROUP_CD = SSTRANSINFOTERMTRLS.DIVISION where DATECREATED BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "'")
                Dim rpt As New TECHRETLUMPALL
                cryRpt = rpt
                '  cryRpt.Refresh()

                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)

            ElseIf cbSTrans.Text = "Salary Loan" And rbBranch.Checked = True And cbSalloan.Text = "Employed" Then
                GetNameBranch()
                GetName = "Salary Loan Employed"
                'dt = db.ExecuteSQLQuery("SELECT SSINFOTERMKIOSK.KIOSK_ID,STRSSSID,strEmpId,LOANAMOUNT,(STRADDRESS1 +', '+ STRADDRESS2) as [address],strPostalcode, (memfname +' '+memmidInit + ' ' +memlname) as [name], TRANSREFNO,IN_IPADD,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSAPPSLEMP.ENCODE_DT,SSTRANSAPPSLEMP.ENCODE_TME FROM SSTRANSAPPSLEMP INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAPPSLEMP.IN_IPADD INNER JOIN SSINFOTERMBR ON SSTRANSAPPSLEMP.BRANCH_CD = SSINFOTERMBR.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSTRANSAPPSLEMP.CLSTR = SSINFOTERMCLSTR.CLSTR_CD INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSTRANSAPPSLEMP.DIVSN where SSTRANSAPPSLEMP.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "'")
                Dim rpt As New SALLOANEMPALL
                cryRpt = rpt
                'cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf cbSTrans.Text = "Salary Loan" And rbGroup.Checked = True And cbSalloan.Text = "Employed" Then
                GetNameBranch()
                GetName = "Salary Loan Employed"
                'dt = db.ExecuteSQLQuery("SELECT SSINFOTERMKIOSK.KIOSK_ID,STRSSSID,strEmpId,LOANAMOUNT,(STRADDRESS1 +', '+ STRADDRESS2) as [address],strPostalcode, (memfname +' '+memmidInit + ' ' +memlname) as [name], TRANSREFNO,IN_IPADD,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSAPPSLEMP.ENCODE_DT,SSTRANSAPPSLEMP.ENCODE_TME FROM SSTRANSAPPSLEMP INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAPPSLEMP.IN_IPADD INNER JOIN SSINFOTERMBR ON SSTRANSAPPSLEMP.BRANCH_CD = SSINFOTERMBR.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSTRANSAPPSLEMP.CLSTR = SSINFOTERMCLSTR.CLSTR_CD INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSTRANSAPPSLEMP.DIVSN where SSTRANSAPPSLEMP.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "'")
                Dim rpt As New SALLOANEMPALL
                cryRpt = rpt
                'cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)

            ElseIf cbSTrans.Text = "Salary Loan" And rbCluster.Checked = True And cbSalloan.Text = "Employed" Then
                GetNameBranch()
                GetName = "Salary Loan Employed"
                'dt = db.ExecuteSQLQuery("SELECT SSINFOTERMKIOSK.KIOSK_ID,STRSSSID,strEmpId,LOANAMOUNT,(STRADDRESS1 +', '+ STRADDRESS2) as [address],strPostalcode, (memfname +' '+memmidInit + ' ' +memlname) as [name], TRANSREFNO,IN_IPADD,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSAPPSLEMP.ENCODE_DT,SSTRANSAPPSLEMP.ENCODE_TME FROM SSTRANSAPPSLEMP INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAPPSLEMP.IN_IPADD INNER JOIN SSINFOTERMBR ON SSTRANSAPPSLEMP.BRANCH_CD = SSINFOTERMBR.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSTRANSAPPSLEMP.CLSTR = SSINFOTERMCLSTR.CLSTR_CD INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSTRANSAPPSLEMP.DIVSN where SSTRANSAPPSLEMP.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "'")
                Dim rpt As New SALLOANEMPALL
                cryRpt = rpt
                'cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)

            ElseIf cbSTrans.Text = "Salary Loan" And rbAll.Checked = True And cbSalloan.Text = "Employed" Then
                GetNameBranch()
                GetName = "Salary Loan Employed"
                'dt = db.ExecuteSQLQuery("SELECT SSINFOTERMKIOSK.KIOSK_ID,STRSSSID,strEmpId,LOANAMOUNT,(STRADDRESS1 +', '+ STRADDRESS2) as [address],strPostalcode, (memfname +' '+memmidInit + ' ' +memlname) as [name], TRANSREFNO,IN_IPADD,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSAPPSLEMP.ENCODE_DT,SSTRANSAPPSLEMP.ENCODE_TME FROM SSTRANSAPPSLEMP INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAPPSLEMP.IN_IPADD INNER JOIN SSINFOTERMBR ON SSTRANSAPPSLEMP.BRANCH_CD = SSINFOTERMBR.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSTRANSAPPSLEMP.CLSTR = SSINFOTERMCLSTR.CLSTR_CD INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSTRANSAPPSLEMP.DIVSN where SSTRANSAPPSLEMP.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "'")
                Dim rpt As New SALLOANEMPALL
                cryRpt = rpt
                'cryRpt.Refresh()
                'cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                OpenReportDbase()
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
                'edel_edel
                OpenReportDbase()
            ElseIf cbSTrans.Text = "Salary Loan" And rbBranch.Checked = True And cbSalloan.Text = "Self-Employed" Then
                GetNameBranch()
                GetName = "Salary Loan SEVM"
                'dt = db.ExecuteSQLQuery("SELECT IN_SSNBR,S_TRANID,SSTRANSAPPSL.IN_IPADD,(IN_BRFNM+ ' '+IN_BRMID+ ' ' + IN_BRSNM) as [name],(IN_HOME1 + ','+ IN_HOME2) as [address],SSINFOTERMKIOSK.KIOSK_ID, SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,IN_TUITN,SSTRANSAPPSL.ENCODE_DT,ENCODE_TME,SSTRANSINFOTERMMEMTAG.MEM_NM from SSTRANSAPPSL INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAPPSL.IN_IPADD INNER JOIN SSTRANSINFOTERMMEMTAG on SSTRANSINFOTERMMEMTAG.MEM_CD = SSTRANSAPPSL.STRMEMBERSTATUS inner join SSINFOTERMBR on SSTRANSAPPSL.BRANCH_CD = SSINFOTERMBR.BRANCH_CD inner join SSINFOTERMCLSTR on SSTRANSAPPSL.CLSTR = SSINFOTERMCLSTR.CLSTR_CD INNER JOIN SSINFOTERMGROUP on SSINFOTERMGROUP.GROUP_CD = SSTRANSAPPSL.DIVSN where SSTRANSAPPSL.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "'and SSTRANSINFOTERMMEMTAG.MEM_NM not LIKE '%COVERED EMPLOYEE%' and SSTRANSINFOTERMMEMTAG.MEM_NM not LIKE '%VOLUNTARY MEMBER%' and SSTRANSINFOTERMMEMTAG.MEM_NM not LIKE '%NON-WORKING SPOUSE%' and SSTRANSINFOTERMMEMTAG.MEM_NM not LIKE '%OVERSEAS CONTRACT WORKER%' and SSTRANSINFOTERMMEMTAG.MEM_NM not LIKE '%PENSIONER%' and SSTRANSINFOTERMMEMTAG.MEM_NM not LIKE '%HOUSEHOLD HELPER%'")
                Dim rpt As New SALLOANSEALL
                cryRpt = rpt
                '    cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf cbSTrans.Text = "Salary Loan" And rbGroup.Checked = True And cbSalloan.Text = "Self-Employed" Then
                GetNameBranch()
                GetName = "Salary Loan SEVM"
                'dt = db.ExecuteSQLQuery("SELECT IN_SSNBR,S_TRANID,SSTRANSAPPSL.IN_IPADD,(IN_BRFNM+ ' '+IN_BRMID+ ' ' + IN_BRSNM) as [name],(IN_HOME1 + ','+ IN_HOME2) as [address],SSINFOTERMKIOSK.KIOSK_ID, SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,IN_TUITN,SSTRANSAPPSL.ENCODE_DT,ENCODE_TME,SSTRANSINFOTERMMEMTAG.MEM_NM from SSTRANSAPPSL INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAPPSL.IN_IPADD INNER JOIN SSTRANSINFOTERMMEMTAG on SSTRANSINFOTERMMEMTAG.MEM_CD = SSTRANSAPPSL.STRMEMBERSTATUS inner join SSINFOTERMBR on SSTRANSAPPSL.BRANCH_CD = SSINFOTERMBR.BRANCH_CD inner join SSINFOTERMCLSTR on SSTRANSAPPSL.CLSTR = SSINFOTERMCLSTR.CLSTR_CD INNER JOIN SSINFOTERMGROUP on SSINFOTERMGROUP.GROUP_CD = SSTRANSAPPSL.DIVSN where SSTRANSAPPSL.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "'and SSTRANSINFOTERMMEMTAG.MEM_NM not LIKE '%COVERED EMPLOYEE%' and SSTRANSINFOTERMMEMTAG.MEM_NM not LIKE '%VOLUNTARY MEMBER%' and SSTRANSINFOTERMMEMTAG.MEM_NM not LIKE '%NON-WORKING SPOUSE%' and SSTRANSINFOTERMMEMTAG.MEM_NM not LIKE '%OVERSEAS CONTRACT WORKER%' and SSTRANSINFOTERMMEMTAG.MEM_NM not LIKE '%PENSIONER%' and SSTRANSINFOTERMMEMTAG.MEM_NM not LIKE '%HOUSEHOLD HELPER%'")
                Dim rpt As New SALLOANSEALL
                cryRpt = rpt
                '    cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf cbSTrans.Text = "Salary Loan" And rbCluster.Checked = True And cbSalloan.Text = "Self-Employed" Then
                GetNameBranch()
                GetName = "Salary Loan SEVM"
                'dt = db.ExecuteSQLQuery("SELECT IN_SSNBR,S_TRANID,SSTRANSAPPSL.IN_IPADD,(IN_BRFNM+ ' '+IN_BRMID+ ' ' + IN_BRSNM) as [name],(IN_HOME1 + ','+ IN_HOME2) as [address],SSINFOTERMKIOSK.KIOSK_ID, SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,IN_TUITN,SSTRANSAPPSL.ENCODE_DT,ENCODE_TME,SSTRANSINFOTERMMEMTAG.MEM_NM from SSTRANSAPPSL INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAPPSL.IN_IPADD INNER JOIN SSTRANSINFOTERMMEMTAG on SSTRANSINFOTERMMEMTAG.MEM_CD = SSTRANSAPPSL.STRMEMBERSTATUS inner join SSINFOTERMBR on SSTRANSAPPSL.BRANCH_CD = SSINFOTERMBR.BRANCH_CD inner join SSINFOTERMCLSTR on SSTRANSAPPSL.CLSTR = SSINFOTERMCLSTR.CLSTR_CD INNER JOIN SSINFOTERMGROUP on SSINFOTERMGROUP.GROUP_CD = SSTRANSAPPSL.DIVSN where SSTRANSAPPSL.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "'and SSTRANSINFOTERMMEMTAG.MEM_NM not LIKE '%COVERED EMPLOYEE%' and SSTRANSINFOTERMMEMTAG.MEM_NM not LIKE '%VOLUNTARY MEMBER%' and SSTRANSINFOTERMMEMTAG.MEM_NM not LIKE '%NON-WORKING SPOUSE%' and SSTRANSINFOTERMMEMTAG.MEM_NM not LIKE '%OVERSEAS CONTRACT WORKER%' and SSTRANSINFOTERMMEMTAG.MEM_NM not LIKE '%PENSIONER%' and SSTRANSINFOTERMMEMTAG.MEM_NM not LIKE '%HOUSEHOLD HELPER%'")
                Dim rpt As New SALLOANSEALL
                cryRpt = rpt
                '    cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)

            ElseIf cbSTrans.Text = "Salary Loan" And rbAll.Checked = True And cbSalloan.Text = "Self-Employed" Then
                GetNameBranch()
                GetName = "Salary Loan SEVM"
                'dt = db.ExecuteSQLQuery("SELECT IN_SSNBR,S_TRANID,SSTRANSAPPSL.IN_IPADD,(IN_BRFNM+ ' '+IN_BRMID+ ' ' + IN_BRSNM) as [name],(IN_HOME1 + ','+ IN_HOME2) as [address],SSINFOTERMKIOSK.KIOSK_ID, SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,IN_TUITN,SSTRANSAPPSL.ENCODE_DT,ENCODE_TME,SSTRANSINFOTERMMEMTAG.MEM_NM from SSTRANSAPPSL INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAPPSL.IN_IPADD INNER JOIN SSTRANSINFOTERMMEMTAG on SSTRANSINFOTERMMEMTAG.MEM_CD = SSTRANSAPPSL.STRMEMBERSTATUS inner join SSINFOTERMBR on SSTRANSAPPSL.BRANCH_CD = SSINFOTERMBR.BRANCH_CD inner join SSINFOTERMCLSTR on SSTRANSAPPSL.CLSTR = SSINFOTERMCLSTR.CLSTR_CD INNER JOIN SSINFOTERMGROUP on SSINFOTERMGROUP.GROUP_CD = SSTRANSAPPSL.DIVSN where SSTRANSAPPSL.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "'and SSTRANSINFOTERMMEMTAG.MEM_NM not LIKE '%COVERED EMPLOYEE%' and SSTRANSINFOTERMMEMTAG.MEM_NM not LIKE '%VOLUNTARY MEMBER%' and SSTRANSINFOTERMMEMTAG.MEM_NM not LIKE '%NON-WORKING SPOUSE%' and SSTRANSINFOTERMMEMTAG.MEM_NM not LIKE '%OVERSEAS CONTRACT WORKER%' and SSTRANSINFOTERMMEMTAG.MEM_NM not LIKE '%PENSIONER%' and SSTRANSINFOTERMMEMTAG.MEM_NM not LIKE '%HOUSEHOLD HELPER%'")
                Dim rpt As New SALLOANSEALL
                cryRpt = rpt
                '    cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf cbSTrans.Text = "Salary Loan" And rbAll.Checked = True And cbSalloan.Text = "Overseas Contract Worker" Then
                GetNameBranch()
                GetName = "SALARY LOAN OFW"
                'dt = db.ExecuteSQLQuery("SELECT IN_SSNBR,S_TRANID,SSTRANSAPPSL.IN_IPADD,(IN_BRFNM+ ' '+IN_BRMID+ ' ' + IN_BRSNM) as [name],(IN_HOME1 + ','+ IN_HOME2) as [address],SSINFOTERMKIOSK.KIOSK_ID, SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,IN_TUITN,SSTRANSAPPSL.ENCODE_DT,ENCODE_TME,SSTRANSINFOTERMMEMTAG.MEM_NM from SSTRANSAPPSL INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAPPSL.IN_IPADD INNER JOIN SSTRANSINFOTERMMEMTAG on SSTRANSINFOTERMMEMTAG.MEM_CD = SSTRANSAPPSL.STRMEMBERSTATUS inner join SSINFOTERMBR on SSTRANSAPPSL.BRANCH_CD = SSINFOTERMBR.BRANCH_CD inner join SSINFOTERMCLSTR on  SSTRANSAPPSL.CLSTR = SSINFOTERMCLSTR.CLSTR_CD INNER JOIN SSINFOTERMGROUP on SSINFOTERMGROUP.GROUP_CD = SSTRANSAPPSL.DIVSN where SSTRANSAPPSL.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and SSTRANSINFOTERMMEMTAG.MEM_NM = 'OVERSEAS CONTRACT WORKER'")
                Dim rpt As New SALLOANOFWALL
                cryRpt = rpt
                '    cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf cbSTrans.Text = "Salary Loan" And rbBranch.Checked = True And cbSalloan.Text = "Overseas Contract Worker" Then
                GetNameBranch()
                GetName = "SALARY LOAN OFW"
                'dt = db.ExecuteSQLQuery("SELECT IN_SSNBR,S_TRANID,SSTRANSAPPSL.IN_IPADD,(IN_BRFNM+ ' '+IN_BRMID+ ' ' + IN_BRSNM) as [name],(IN_HOME1 + ','+ IN_HOME2) as [address],SSINFOTERMKIOSK.KIOSK_ID, SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,IN_TUITN,SSTRANSAPPSL.ENCODE_DT,ENCODE_TME,SSTRANSINFOTERMMEMTAG.MEM_NM from SSTRANSAPPSL INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAPPSL.IN_IPADD INNER JOIN SSTRANSINFOTERMMEMTAG on SSTRANSINFOTERMMEMTAG.MEM_CD = SSTRANSAPPSL.STRMEMBERSTATUS inner join SSINFOTERMBR on SSTRANSAPPSL.BRANCH_CD = SSINFOTERMBR.BRANCH_CD inner join SSINFOTERMCLSTR on  SSTRANSAPPSL.CLSTR = SSINFOTERMCLSTR.CLSTR_CD INNER JOIN SSINFOTERMGROUP on SSINFOTERMGROUP.GROUP_CD = SSTRANSAPPSL.DIVSN where SSTRANSAPPSL.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and SSTRANSINFOTERMMEMTAG.MEM_NM = 'OVERSEAS CONTRACT WORKER'")
                Dim rpt As New SALLOANOFWALL
                cryRpt = rpt
                '    cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf cbSTrans.Text = "Salary Loan" And rbGroup.Checked = True And cbSalloan.Text = "Overseas Contract Worker" Then
                GetNameBranch()
                GetName = "SALARY LOAN OFW"
                'dt = db.ExecuteSQLQuery("SELECT IN_SSNBR,S_TRANID,SSTRANSAPPSL.IN_IPADD,(IN_BRFNM+ ' '+IN_BRMID+ ' ' + IN_BRSNM) as [name],(IN_HOME1 + ','+ IN_HOME2) as [address],SSINFOTERMKIOSK.KIOSK_ID, SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,IN_TUITN,SSTRANSAPPSL.ENCODE_DT,ENCODE_TME,SSTRANSINFOTERMMEMTAG.MEM_NM from SSTRANSAPPSL INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAPPSL.IN_IPADD INNER JOIN SSTRANSINFOTERMMEMTAG on SSTRANSINFOTERMMEMTAG.MEM_CD = SSTRANSAPPSL.STRMEMBERSTATUS inner join SSINFOTERMBR on SSTRANSAPPSL.BRANCH_CD = SSINFOTERMBR.BRANCH_CD inner join SSINFOTERMCLSTR on  SSTRANSAPPSL.CLSTR = SSINFOTERMCLSTR.CLSTR_CD INNER JOIN SSINFOTERMGROUP on SSINFOTERMGROUP.GROUP_CD = SSTRANSAPPSL.DIVSN where SSTRANSAPPSL.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and SSTRANSINFOTERMMEMTAG.MEM_NM = 'OVERSEAS CONTRACT WORKER'")
                Dim rpt As New SALLOANOFWALL
                cryRpt = rpt
                '    cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf cbSTrans.Text = "Salary Loan" And rbCluster.Checked = True And cbSalloan.Text = "Overseas Contract Worker" Then
                GetNameBranch()
                GetName = "SALARY LOAN OFW"
                'dt = db.ExecuteSQLQuery("SELECT IN_SSNBR,S_TRANID,SSTRANSAPPSL.IN_IPADD,(IN_BRFNM+ ' '+IN_BRMID+ ' ' + IN_BRSNM) as [name],(IN_HOME1 + ','+ IN_HOME2) as [address],SSINFOTERMKIOSK.KIOSK_ID, SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,IN_TUITN,SSTRANSAPPSL.ENCODE_DT,ENCODE_TME,SSTRANSINFOTERMMEMTAG.MEM_NM from SSTRANSAPPSL INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAPPSL.IN_IPADD INNER JOIN SSTRANSINFOTERMMEMTAG on SSTRANSINFOTERMMEMTAG.MEM_CD = SSTRANSAPPSL.STRMEMBERSTATUS inner join SSINFOTERMBR on SSTRANSAPPSL.BRANCH_CD = SSINFOTERMBR.BRANCH_CD inner join SSINFOTERMCLSTR on  SSTRANSAPPSL.CLSTR = SSINFOTERMCLSTR.CLSTR_CD INNER JOIN SSINFOTERMGROUP on SSINFOTERMGROUP.GROUP_CD = SSTRANSAPPSL.DIVSN where SSTRANSAPPSL.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and SSTRANSINFOTERMMEMTAG.MEM_NM = 'OVERSEAS CONTRACT WORKER'")
                Dim rpt As New SALLOANOFWALL
                cryRpt = rpt
                '    cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf cbSTrans.Text = "Salary Loan" And rbCluster.Checked = True And cbSalloan.Text = "Voluntary Member" Then
                GetNameBranch()
                GetName = "SALARY LOAN VOL"
                'dt = db.ExecuteSQLQuery("SELECT IN_SSNBR,S_TRANID,SSTRANSAPPSL.IN_IPADD,(IN_BRFNM+ ' '+IN_BRMID+ ' ' + IN_BRSNM) as [name],(IN_HOME1 + ','+ IN_HOME2) as [address],SSINFOTERMKIOSK.KIOSK_ID, SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,IN_TUITN,SSTRANSAPPSL.ENCODE_DT,ENCODE_TME,SSTRANSINFOTERMMEMTAG.MEM_NM from SSTRANSAPPSL INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAPPSL.IN_IPADD INNER JOIN SSTRANSINFOTERMMEMTAG on SSTRANSINFOTERMMEMTAG.MEM_CD = SSTRANSAPPSL.STRMEMBERSTATUS inner join SSINFOTERMBR on SSTRANSAPPSL.BRANCH_CD = SSINFOTERMBR.BRANCH_CD inner join SSINFOTERMCLSTR on  SSTRANSAPPSL.CLSTR = SSINFOTERMCLSTR.CLSTR_CD INNER JOIN SSINFOTERMGROUP on SSINFOTERMGROUP.GROUP_CD = SSTRANSAPPSL.DIVSN where SSTRANSAPPSL.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and SSINFOTERMBR.BRANCH_NM = '" & cbBranch.Text & "'and SSTRANSINFOTERMMEMTAG.MEM_NM not LIKE '%COVERED EMPLOYEE%' and SSTRANSINFOTERMMEMTAG.MEM_NM not LIKE '%SELF-EMPLOYED%' and SSTRANSINFOTERMMEMTAG.MEM_NM not LIKE '%OVERSEAS CONTRACT WORKER%' and SSTRANSINFOTERMMEMTAG.MEM_NM not LIKE '%PENSIONER%' and SSTRANSINFOTERMMEMTAG.MEM_NM not LIKE '%HOUSEHOLD HELPER%'")
                Dim rpt As New SALLOANVOLALL
                cryRpt = rpt
                '    cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf cbSTrans.Text = "Salary Loan" And rbGroup.Checked = True And cbSalloan.Text = "Voluntary Member" Then
                GetNameBranch()
                GetName = "SALARY LOAN VOL"
                'dt = db.ExecuteSQLQuery("SELECT IN_SSNBR,S_TRANID,SSTRANSAPPSL.IN_IPADD,(IN_BRFNM+ ' '+IN_BRMID+ ' ' + IN_BRSNM) as [name],(IN_HOME1 + ','+ IN_HOME2) as [address],SSINFOTERMKIOSK.KIOSK_ID, SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,IN_TUITN,SSTRANSAPPSL.ENCODE_DT,ENCODE_TME,SSTRANSINFOTERMMEMTAG.MEM_NM from SSTRANSAPPSL INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAPPSL.IN_IPADD INNER JOIN SSTRANSINFOTERMMEMTAG on SSTRANSINFOTERMMEMTAG.MEM_CD = SSTRANSAPPSL.STRMEMBERSTATUS inner join SSINFOTERMBR on SSTRANSAPPSL.BRANCH_CD = SSINFOTERMBR.BRANCH_CD inner join SSINFOTERMCLSTR on  SSTRANSAPPSL.CLSTR = SSINFOTERMCLSTR.CLSTR_CD INNER JOIN SSINFOTERMGROUP on SSINFOTERMGROUP.GROUP_CD = SSTRANSAPPSL.DIVSN where SSTRANSAPPSL.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and SSINFOTERMBR.BRANCH_NM = '" & cbBranch.Text & "'and SSTRANSINFOTERMMEMTAG.MEM_NM not LIKE '%COVERED EMPLOYEE%' and SSTRANSINFOTERMMEMTAG.MEM_NM not LIKE '%SELF-EMPLOYED%' and SSTRANSINFOTERMMEMTAG.MEM_NM not LIKE '%OVERSEAS CONTRACT WORKER%' and SSTRANSINFOTERMMEMTAG.MEM_NM not LIKE '%PENSIONER%' and SSTRANSINFOTERMMEMTAG.MEM_NM not LIKE '%HOUSEHOLD HELPER%'")
                Dim rpt As New SALLOANVOLALL
                cryRpt = rpt
                '    cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)

            ElseIf cbSTrans.Text = "Salary Loan" And rbBranch.Checked = True And cbSalloan.Text = "Voluntary Member" Then
                GetNameBranch()
                GetName = "SALARY LOAN VOL"
                'dt = db.ExecuteSQLQuery("SELECT IN_SSNBR,S_TRANID,SSTRANSAPPSL.IN_IPADD,(IN_BRFNM+ ' '+IN_BRMID+ ' ' + IN_BRSNM) as [name],(IN_HOME1 + ','+ IN_HOME2) as [address],SSINFOTERMKIOSK.KIOSK_ID, SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,IN_TUITN,SSTRANSAPPSL.ENCODE_DT,ENCODE_TME,SSTRANSINFOTERMMEMTAG.MEM_NM from SSTRANSAPPSL INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAPPSL.IN_IPADD INNER JOIN SSTRANSINFOTERMMEMTAG on SSTRANSINFOTERMMEMTAG.MEM_CD = SSTRANSAPPSL.STRMEMBERSTATUS inner join SSINFOTERMBR on SSTRANSAPPSL.BRANCH_CD = SSINFOTERMBR.BRANCH_CD inner join SSINFOTERMCLSTR on  SSTRANSAPPSL.CLSTR = SSINFOTERMCLSTR.CLSTR_CD INNER JOIN SSINFOTERMGROUP on SSINFOTERMGROUP.GROUP_CD = SSTRANSAPPSL.DIVSN where SSTRANSAPPSL.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and SSINFOTERMBR.BRANCH_NM = '" & cbBranch.Text & "'and SSTRANSINFOTERMMEMTAG.MEM_NM not LIKE '%COVERED EMPLOYEE%' and SSTRANSINFOTERMMEMTAG.MEM_NM not LIKE '%SELF-EMPLOYED%' and SSTRANSINFOTERMMEMTAG.MEM_NM not LIKE '%OVERSEAS CONTRACT WORKER%' and SSTRANSINFOTERMMEMTAG.MEM_NM not LIKE '%PENSIONER%' and SSTRANSINFOTERMMEMTAG.MEM_NM not LIKE '%HOUSEHOLD HELPER%'")
                Dim rpt As New SALLOANVOLALL
                cryRpt = rpt
                '    cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf cbSTrans.Text = "Salary Loan" And rbAll.Checked = True And cbSalloan.Text = "Voluntary Member" Then
                GetNameBranch()
                GetName = "SALARY LOAN VOL"
                'dt = db.ExecuteSQLQuery("SELECT IN_SSNBR,S_TRANID,SSTRANSAPPSL.IN_IPADD,(IN_BRFNM+ ' '+IN_BRMID+ ' ' + IN_BRSNM) as [name],(IN_HOME1 + ','+ IN_HOME2) as [address],SSINFOTERMKIOSK.KIOSK_ID, SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,IN_TUITN,SSTRANSAPPSL.ENCODE_DT,ENCODE_TME,SSTRANSINFOTERMMEMTAG.MEM_NM from SSTRANSAPPSL INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAPPSL.IN_IPADD INNER JOIN SSTRANSINFOTERMMEMTAG on SSTRANSINFOTERMMEMTAG.MEM_CD = SSTRANSAPPSL.STRMEMBERSTATUS inner join SSINFOTERMBR on SSTRANSAPPSL.BRANCH_CD = SSINFOTERMBR.BRANCH_CD inner join SSINFOTERMCLSTR on  SSTRANSAPPSL.CLSTR = SSINFOTERMCLSTR.CLSTR_CD INNER JOIN SSINFOTERMGROUP on SSINFOTERMGROUP.GROUP_CD = SSTRANSAPPSL.DIVSN where SSTRANSAPPSL.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and SSINFOTERMBR.BRANCH_NM = '" & cbBranch.Text & "'and SSTRANSINFOTERMMEMTAG.MEM_NM not LIKE '%COVERED EMPLOYEE%' and SSTRANSINFOTERMMEMTAG.MEM_NM not LIKE '%SELF-EMPLOYED%' and SSTRANSINFOTERMMEMTAG.MEM_NM not LIKE '%OVERSEAS CONTRACT WORKER%' and SSTRANSINFOTERMMEMTAG.MEM_NM not LIKE '%PENSIONER%' and SSTRANSINFOTERMMEMTAG.MEM_NM not LIKE '%HOUSEHOLD HELPER%'")
                Dim rpt As New SALLOANVOLALL
                cryRpt = rpt
                '    cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf cbSTrans.Text = "Salary Loan" And rbBranch.Checked = True And cbSalloan.Text = "SE/VM/OFW" Then
                GetNameBranch()
                GetName = "Salary Loan All"
                'dt = db.ExecuteSQLQuery("SELECT IN_SSNBR,S_TRANID,SSTRANSAPPSL.IN_IPADD,(IN_BRFNM+ ' '+IN_BRMID+ ' ' + IN_BRSNM) as [name],(IN_HOME1 + ','+ IN_HOME2) as [address],SSINFOTERMKIOSK.KIOSK_ID, SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,IN_TUITN,SSTRANSAPPSL.ENCODE_DT,ENCODE_TME,SSTRANSINFOTERMMEMTAG.MEM_NM from SSTRANSAPPSL INNER JOIN SSTRANSINFOTERMMEMTAG ON SSTRANSAPPSL.STRMEMBERSTATUS = SSTRANSINFOTERMMEMTAG.MEM_CD INNER JOIN SSINFOTERMKIOSK on SSTRANSAPPSL.IN_IPADD = SSINFOTERMKIOSK.BRANCH_IP inner join SSINFOTERMBR on SSTRANSAPPSL.BRANCH_CD = SSINFOTERMBR.BRANCH_CD inner join SSINFOTERMCLSTR on  SSTRANSAPPSL.CLSTR = SSINFOTERMCLSTR.CLSTR_CD INNER JOIN SSINFOTERMGROUP on SSINFOTERMGROUP.GROUP_CD = SSTRANSAPPSL.DIVSN where SSTRANSAPPSL.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "'")
                Dim rpt As New SALLOANALLALL
                cryRpt = rpt
                'cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf cbSTrans.Text = "Salary Loan" And rbGroup.Checked = True And cbSalloan.Text = "SE/VM/OFW" Then
                GetNameBranch()
                GetName = "Salary Loan All"
                'dt = db.ExecuteSQLQuery("SELECT IN_SSNBR,S_TRANID,SSTRANSAPPSL.IN_IPADD,(IN_BRFNM+ ' '+IN_BRMID+ ' ' + IN_BRSNM) as [name],(IN_HOME1 + ','+ IN_HOME2) as [address],SSINFOTERMKIOSK.KIOSK_ID, SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,IN_TUITN,SSTRANSAPPSL.ENCODE_DT,ENCODE_TME,SSTRANSINFOTERMMEMTAG.MEM_NM from SSTRANSAPPSL INNER JOIN SSTRANSINFOTERMMEMTAG ON SSTRANSAPPSL.STRMEMBERSTATUS = SSTRANSINFOTERMMEMTAG.MEM_CD INNER JOIN SSINFOTERMKIOSK on SSTRANSAPPSL.IN_IPADD = SSINFOTERMKIOSK.BRANCH_IP inner join SSINFOTERMBR on SSTRANSAPPSL.BRANCH_CD = SSINFOTERMBR.BRANCH_CD inner join SSINFOTERMCLSTR on  SSTRANSAPPSL.CLSTR = SSINFOTERMCLSTR.CLSTR_CD INNER JOIN SSINFOTERMGROUP on SSINFOTERMGROUP.GROUP_CD = SSTRANSAPPSL.DIVSN where SSTRANSAPPSL.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "'")
                Dim rpt As New SALLOANALLALL
                cryRpt = rpt
                'cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf cbSTrans.Text = "Salary Loan" And rbCluster.Checked = True And cbSalloan.Text = "SE/VM/OFW" Then
                GetNameBranch()
                GetName = "Salary Loan All"
                'dt = db.ExecuteSQLQuery("SELECT IN_SSNBR,S_TRANID,SSTRANSAPPSL.IN_IPADD,(IN_BRFNM+ ' '+IN_BRMID+ ' ' + IN_BRSNM) as [name],(IN_HOME1 + ','+ IN_HOME2) as [address],SSINFOTERMKIOSK.KIOSK_ID, SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,IN_TUITN,SSTRANSAPPSL.ENCODE_DT,ENCODE_TME,SSTRANSINFOTERMMEMTAG.MEM_NM from SSTRANSAPPSL INNER JOIN SSTRANSINFOTERMMEMTAG ON SSTRANSAPPSL.STRMEMBERSTATUS = SSTRANSINFOTERMMEMTAG.MEM_CD INNER JOIN SSINFOTERMKIOSK on SSTRANSAPPSL.IN_IPADD = SSINFOTERMKIOSK.BRANCH_IP inner join SSINFOTERMBR on SSTRANSAPPSL.BRANCH_CD = SSINFOTERMBR.BRANCH_CD inner join SSINFOTERMCLSTR on  SSTRANSAPPSL.CLSTR = SSINFOTERMCLSTR.CLSTR_CD INNER JOIN SSINFOTERMGROUP on SSINFOTERMGROUP.GROUP_CD = SSTRANSAPPSL.DIVSN where SSTRANSAPPSL.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "'")
                Dim rpt As New SALLOANALLALL
                cryRpt = rpt
                'cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf cbSTrans.Text = "Salary Loan" And rbAll.Checked = True And cbSalloan.Text = "SE/VM/OFW" Then
                GetNameBranch()
                GetName = "Salary Loan All"
                'dt = db.ExecuteSQLQuery("SELECT IN_SSNBR,S_TRANID,SSTRANSAPPSL.IN_IPADD,(IN_BRFNM+ ' '+IN_BRMID+ ' ' + IN_BRSNM) as [name],(IN_HOME1 + ','+ IN_HOME2) as [address],SSINFOTERMKIOSK.KIOSK_ID, SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,IN_TUITN,SSTRANSAPPSL.ENCODE_DT,ENCODE_TME,SSTRANSINFOTERMMEMTAG.MEM_NM from SSTRANSAPPSL INNER JOIN SSTRANSINFOTERMMEMTAG ON SSTRANSAPPSL.STRMEMBERSTATUS = SSTRANSINFOTERMMEMTAG.MEM_CD INNER JOIN SSINFOTERMKIOSK on SSTRANSAPPSL.IN_IPADD = SSINFOTERMKIOSK.BRANCH_IP inner join SSINFOTERMBR on SSTRANSAPPSL.BRANCH_CD = SSINFOTERMBR.BRANCH_CD inner join SSINFOTERMCLSTR on  SSTRANSAPPSL.CLSTR = SSINFOTERMCLSTR.CLSTR_CD INNER JOIN SSINFOTERMGROUP on SSINFOTERMGROUP.GROUP_CD = SSTRANSAPPSL.DIVSN where SSTRANSAPPSL.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "'")
                Dim rpt As New SALLOANALLALL
                cryRpt = rpt
                'cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf cbSTrans.Text = "ACOP" And rbGroup.Checked = True Then
                GetNameBranch()
                GetName = "ACOP"
                'dt = db.ExecuteSQLQuery("SELECT SSTRANSACOP.SSNUM,SSTRANSACOP.TRANID,SSTRANSACOP.BRANCH_IP,IKAUDITREPMF.REFNO,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSACOP.ENCODE_DT,SSTRANSACOP.ENCODE_TME,SSTRANSACOP.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,NXTSUBM,SSINFOTERMGROUP.GROUP_NM from SSTRANSACOP INNER JOIN IKAUDITREPMF on SSTRANSACOP.TRANID = IKAUDITREPMF.TRANID inner join SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSTRANSACOP.BRANCH_IP INNER JOIN SSINFOTERMBR ON SSTRANSACOP.BRANCH_CD = SSINFOTERMBR.BRANCH_CD INNER JOIN SSINFOTERMCLSTR on SSTRANSACOP.CLSTR = SSINFOTERMCLSTR.CLSTR_CD INNER JOIN SSINFOTERMGROUP on SSTRANSACOP.DIVSN = SSINFOTERMGROUP.GROUP_CD WHERE SSTRANSACOP.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' ORDER by SSTRANSACOP.ENCODE_DT")
                Dim rpt As New ACOPALL
                cryRpt = rpt
                '    cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf cbSTrans.Text = "ACOP" And rbBranch.Checked = True Then
                GetNameBranch()
                GetName = "ACOP"
                'dt = db.ExecuteSQLQuery("SELECT SSTRANSACOP.SSNUM,SSTRANSACOP.TRANID,SSTRANSACOP.BRANCH_IP,IKAUDITREPMF.REFNO,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSACOP.ENCODE_DT,SSTRANSACOP.ENCODE_TME,SSTRANSACOP.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,NXTSUBM,SSINFOTERMGROUP.GROUP_NM from SSTRANSACOP INNER JOIN IKAUDITREPMF on SSTRANSACOP.TRANID = IKAUDITREPMF.TRANID inner join SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSTRANSACOP.BRANCH_IP INNER JOIN SSINFOTERMBR ON SSTRANSACOP.BRANCH_CD = SSINFOTERMBR.BRANCH_CD INNER JOIN SSINFOTERMCLSTR on SSTRANSACOP.CLSTR = SSINFOTERMCLSTR.CLSTR_CD INNER JOIN SSINFOTERMGROUP on SSTRANSACOP.DIVSN = SSINFOTERMGROUP.GROUP_CD WHERE SSTRANSACOP.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' ORDER by SSTRANSACOP.ENCODE_DT")
                Dim rpt As New ACOPALL
                cryRpt = rpt
                '    cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf cbSTrans.Text = "ACOP" And rbCluster.Checked = True Then
                GetNameBranch()
                GetName = "ACOP"
                'dt = db.ExecuteSQLQuery("SELECT SSTRANSACOP.SSNUM,SSTRANSACOP.TRANID,SSTRANSACOP.BRANCH_IP,IKAUDITREPMF.REFNO,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSACOP.ENCODE_DT,SSTRANSACOP.ENCODE_TME,SSTRANSACOP.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,NXTSUBM,SSINFOTERMGROUP.GROUP_NM from SSTRANSACOP INNER JOIN IKAUDITREPMF on SSTRANSACOP.TRANID = IKAUDITREPMF.TRANID inner join SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSTRANSACOP.BRANCH_IP INNER JOIN SSINFOTERMBR ON SSTRANSACOP.BRANCH_CD = SSINFOTERMBR.BRANCH_CD INNER JOIN SSINFOTERMCLSTR on SSTRANSACOP.CLSTR = SSINFOTERMCLSTR.CLSTR_CD INNER JOIN SSINFOTERMGROUP on SSTRANSACOP.DIVSN = SSINFOTERMGROUP.GROUP_CD WHERE SSTRANSACOP.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' ORDER by SSTRANSACOP.ENCODE_DT")
                Dim rpt As New ACOPALL
                cryRpt = rpt
                '    cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf cbSTrans.Text = "ACOP" And rbAll.Checked = True Then
                GetNameBranch()
                GetName = "ACOP"
                'dt = db.ExecuteSQLQuery("SELECT SSTRANSACOP.SSNUM,SSTRANSACOP.TRANID,SSTRANSACOP.BRANCH_IP,IKAUDITREPMF.REFNO,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSACOP.ENCODE_DT,SSTRANSACOP.ENCODE_TME,SSTRANSACOP.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,NXTSUBM,SSINFOTERMGROUP.GROUP_NM from SSTRANSACOP INNER JOIN IKAUDITREPMF on SSTRANSACOP.TRANID = IKAUDITREPMF.TRANID inner join SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSTRANSACOP.BRANCH_IP INNER JOIN SSINFOTERMBR ON SSTRANSACOP.BRANCH_CD = SSINFOTERMBR.BRANCH_CD INNER JOIN SSINFOTERMCLSTR on SSTRANSACOP.CLSTR = SSINFOTERMCLSTR.CLSTR_CD INNER JOIN SSINFOTERMGROUP on SSTRANSACOP.DIVSN = SSINFOTERMGROUP.GROUP_CD WHERE SSTRANSACOP.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' ORDER by SSTRANSACOP.ENCODE_DT")
                Dim rpt As New ACOPALL
                cryRpt = rpt
                '    cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf cbSTrans.Text = "Maternity Notification" And rbBranch.Checked = True Then
                GetNameBranch()
                GetName = "Maternity Notification"
                'dt = db.ExecuteSQLQuery("SELECT SSNUM,WTRANS_NO,SSINFOTERMBR.BRANCH_NM,SSTRANSINFORTERMMN.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP,DELIV_DATE,LDELIV_NO,LDELIV_DATE,SSTRANSINFORTERMMN.ENCODE_DT,ENCODE_TME,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM from SSTRANSINFORTERMMN INNER JOIN SSINFOTERMBR ON SSTRANSINFORTERMMN.BRANCH_CD = SSINFOTERMBR.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMBR.CLSTR_CD INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSINFOTERMBR.GROUP_CD INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID = SSTRANSINFORTERMMN.KIOSK_ID WHERE SSTRANSINFORTERMMN.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' ORDER BY SSTRANSINFORTERMMN.ENCODE_DT")
                Dim rpt As New MATNOTIFNEW
                cryRpt = rpt
                '    cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf cbSTrans.Text = "Maternity Notification" And rbGroup.Checked = True Then
                GetNameBranch()
                GetName = "Maternity Notification"
                'dt = db.ExecuteSQLQuery("SELECT SSNUM,WTRANS_NO,SSINFOTERMBR.BRANCH_NM,SSTRANSINFORTERMMN.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP,DELIV_DATE,LDELIV_NO,LDELIV_DATE,SSTRANSINFORTERMMN.ENCODE_DT,ENCODE_TME,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM from SSTRANSINFORTERMMN INNER JOIN SSINFOTERMBR ON SSTRANSINFORTERMMN.BRANCH_CD = SSINFOTERMBR.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMBR.CLSTR_CD INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSINFOTERMBR.GROUP_CD INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID = SSTRANSINFORTERMMN.KIOSK_ID WHERE SSTRANSINFORTERMMN.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' ORDER BY SSTRANSINFORTERMMN.ENCODE_DT")
                Dim rpt As New MATNOTIFNEW
                cryRpt = rpt
                '    cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf cbSTrans.Text = "Maternity Notification" And rbCluster.Checked = True Then
                GetNameBranch()
                GetName = "Maternity Notification"
                'dt = db.ExecuteSQLQuery("SELECT SSNUM,WTRANS_NO,SSINFOTERMBR.BRANCH_NM,SSTRANSINFORTERMMN.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP,DELIV_DATE,LDELIV_NO,LDELIV_DATE,SSTRANSINFORTERMMN.ENCODE_DT,ENCODE_TME,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM from SSTRANSINFORTERMMN INNER JOIN SSINFOTERMBR ON SSTRANSINFORTERMMN.BRANCH_CD = SSINFOTERMBR.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMBR.CLSTR_CD INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSINFOTERMBR.GROUP_CD INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID = SSTRANSINFORTERMMN.KIOSK_ID WHERE SSTRANSINFORTERMMN.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' ORDER BY SSTRANSINFORTERMMN.ENCODE_DT")
                Dim rpt As New MATNOTIFNEW
                cryRpt = rpt
                '    cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf cbSTrans.Text = "Maternity Notification" And rbAll.Checked = True Then
                GetNameBranch()
                GetName = "Maternity Notification"
                'dt = db.ExecuteSQLQuery("SELECT SSNUM,WTRANS_NO,SSINFOTERMBR.BRANCH_NM,SSTRANSINFORTERMMN.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP,DELIV_DATE,LDELIV_NO,LDELIV_DATE,SSTRANSINFORTERMMN.ENCODE_DT,ENCODE_TME,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM from SSTRANSINFORTERMMN INNER JOIN SSINFOTERMBR ON SSTRANSINFORTERMMN.BRANCH_CD = SSINFOTERMBR.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMBR.CLSTR_CD INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSINFOTERMBR.GROUP_CD INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID = SSTRANSINFORTERMMN.KIOSK_ID WHERE SSTRANSINFORTERMMN.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' ORDER BY SSTRANSINFORTERMMN.ENCODE_DT")
                Dim rpt As New MATNOTIFNEW
                cryRpt = rpt
                '    cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)

#Region "PIN CHANGE DETAILED"
            ElseIf cbSTrans.Text = "Pin Change" And rbGroup.Checked = True Then
                GetNameBranch()
                GetName = "Pin Change"
                'dt = db.ExecuteSQLQuery("SELECT SSNUM,DOBTH,DOCVRG,UMID_BANK,BANK_BR,TRANS_ACCTNO,STREET,BRGAY,POST_CD,EMAIL,PHONE,CELNO,USRTYP,SSTRANSINFOTERMTR.ENCODE_DT,SSTRANSINFOTERMTR.ENCODE_TME ,SSTRANSINFOTERMTR.BRANCH_IP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSINFOTERMTR.TAG,TRANID FROM SSTRANSINFOTERMTR INNER JOIN SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSTRANSINFOTERMTR.BRANCH_IP inner join SSINFOTERMBR on SSTRANSINFOTERMTR.BRANCH_CD = SSINFOTERMBR.BRANCH_CD inner join SSINFOTERMCLSTR on SSTRANSINFOTERMTR.CLSTR = SSINFOTERMCLSTR.CLSTR_CD INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSTRANSINFOTERMTR.DIVSN where SSTRANSINFOTERMTR.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "'")
                Dim rpt As New PINCHANGEDETAILED
                cryRpt = rpt
                '  cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf cbSTrans.Text = "Pin Change" And rbBranch.Checked = True Then
                GetNameBranch()
                GetName = "Pin Change"

                'dt = db.ExecuteSQLQuery("SELECT SSNUM,DOBTH,DOCVRG,UMID_BANK,BANK_BR,TRANS_ACCTNO,STREET,BRGAY,POST_CD,EMAIL,PHONE,CELNO,USRTYP,SSTRANSINFOTERMTR.ENCODE_DT,SSTRANSINFOTERMTR.ENCODE_TME ,SSTRANSINFOTERMTR.BRANCH_IP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSINFOTERMTR.TAG,TRANID FROM SSTRANSINFOTERMTR INNER JOIN SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSTRANSINFOTERMTR.BRANCH_IP inner join SSINFOTERMBR on SSTRANSINFOTERMTR.BRANCH_CD = SSINFOTERMBR.BRANCH_CD inner join SSINFOTERMCLSTR on SSTRANSINFOTERMTR.CLSTR = SSINFOTERMCLSTR.CLSTR_CD INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSTRANSINFOTERMTR.DIVSN where SSTRANSINFOTERMTR.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "'")
                Dim rpt As New PINCHANGEDETAILED
                cryRpt = rpt
                '  cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)

            ElseIf cbSTrans.Text = "Pin Change" And rbCluster.Checked = True Then
                GetNameBranch()
                GetName = "Pin Change"

                'dt = db.ExecuteSQLQuery("SELECT SSNUM,DOBTH,DOCVRG,UMID_BANK,BANK_BR,TRANS_ACCTNO,STREET,BRGAY,POST_CD,EMAIL,PHONE,CELNO,USRTYP,SSTRANSINFOTERMTR.ENCODE_DT,SSTRANSINFOTERMTR.ENCODE_TME ,SSTRANSINFOTERMTR.BRANCH_IP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSINFOTERMTR.TAG,TRANID FROM SSTRANSINFOTERMTR INNER JOIN SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSTRANSINFOTERMTR.BRANCH_IP inner join SSINFOTERMBR on SSTRANSINFOTERMTR.BRANCH_CD = SSINFOTERMBR.BRANCH_CD inner join SSINFOTERMCLSTR on SSTRANSINFOTERMTR.CLSTR = SSINFOTERMCLSTR.CLSTR_CD INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSTRANSINFOTERMTR.DIVSN where SSTRANSINFOTERMTR.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "'")
                Dim rpt As New PINCHANGEDETAILED
                cryRpt = rpt
                '  cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)

            ElseIf cbSTrans.Text = "Pin Change" = True And rbAll.Checked = True Then
                GetNameBranch()
                GetName = "Pin Change"

                'dt = db.ExecuteSQLQuery("SELECT SSNUM,DOBTH,DOCVRG,UMID_BANK,BANK_BR,TRANS_ACCTNO,STREET,BRGAY,POST_CD,EMAIL,PHONE,CELNO,USRTYP,SSTRANSINFOTERMTR.ENCODE_DT,SSTRANSINFOTERMTR.ENCODE_TME ,SSTRANSINFOTERMTR.BRANCH_IP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSINFOTERMTR.TAG,TRANID FROM SSTRANSINFOTERMTR INNER JOIN SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSTRANSINFOTERMTR.BRANCH_IP inner join SSINFOTERMBR on SSTRANSINFOTERMTR.BRANCH_CD = SSINFOTERMBR.BRANCH_CD inner join SSINFOTERMCLSTR on SSTRANSINFOTERMTR.CLSTR = SSINFOTERMCLSTR.CLSTR_CD INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSTRANSINFOTERMTR.DIVSN where SSTRANSINFOTERMTR.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "'")
                Dim rpt As New PINCHANGEDETAILED
                cryRpt = rpt
                '  cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
#End Region


#Region "Transactions using GSIS Card"
            ElseIf cbSTrans.Text = "Transactions using GSIS Card" And rbGroup.Checked = True Then
                GetNameBranch()
                GetName = "Transactions using GSIS Card"
                'dt = db.ExecuteSQLQuery("SELECT SSNUM,DOBTH,DOCVRG,UMID_BANK,BANK_BR,TRANS_ACCTNO,STREET,BRGAY,POST_CD,EMAIL,PHONE,CELNO,USRTYP,SSTRANSINFOTERMTR.ENCODE_DT,SSTRANSINFOTERMTR.ENCODE_TME ,SSTRANSINFOTERMTR.BRANCH_IP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSINFOTERMTR.TAG,TRANID FROM SSTRANSINFOTERMTR INNER JOIN SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSTRANSINFOTERMTR.BRANCH_IP inner join SSINFOTERMBR on SSTRANSINFOTERMTR.BRANCH_CD = SSINFOTERMBR.BRANCH_CD inner join SSINFOTERMCLSTR on SSTRANSINFOTERMTR.CLSTR = SSINFOTERMCLSTR.CLSTR_CD INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSTRANSINFOTERMTR.DIVSN where SSTRANSINFOTERMTR.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "'")
                Dim rpt As New TRANSUSINGGSIS
                cryRpt = rpt
                '  cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf cbSTrans.Text = "Transactions using GSIS Card" And rbBranch.Checked = True Then
                GetNameBranch()
                GetName = "Transactions using GSIS Card"

                'dt = db.ExecuteSQLQuery("SELECT SSNUM,DOBTH,DOCVRG,UMID_BANK,BANK_BR,TRANS_ACCTNO,STREET,BRGAY,POST_CD,EMAIL,PHONE,CELNO,USRTYP,SSTRANSINFOTERMTR.ENCODE_DT,SSTRANSINFOTERMTR.ENCODE_TME ,SSTRANSINFOTERMTR.BRANCH_IP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSINFOTERMTR.TAG,TRANID FROM SSTRANSINFOTERMTR INNER JOIN SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSTRANSINFOTERMTR.BRANCH_IP inner join SSINFOTERMBR on SSTRANSINFOTERMTR.BRANCH_CD = SSINFOTERMBR.BRANCH_CD inner join SSINFOTERMCLSTR on SSTRANSINFOTERMTR.CLSTR = SSINFOTERMCLSTR.CLSTR_CD INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSTRANSINFOTERMTR.DIVSN where SSTRANSINFOTERMTR.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "'")
                Dim rpt As New TRANSUSINGGSIS
                cryRpt = rpt
                '  cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)

            ElseIf cbSTrans.Text = "Transactions using GSIS Card" And rbCluster.Checked = True Then
                GetNameBranch()
                GetName = "Transactions using GSIS Card"

                'dt = db.ExecuteSQLQuery("SELECT SSNUM,DOBTH,DOCVRG,UMID_BANK,BANK_BR,TRANS_ACCTNO,STREET,BRGAY,POST_CD,EMAIL,PHONE,CELNO,USRTYP,SSTRANSINFOTERMTR.ENCODE_DT,SSTRANSINFOTERMTR.ENCODE_TME ,SSTRANSINFOTERMTR.BRANCH_IP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSINFOTERMTR.TAG,TRANID FROM SSTRANSINFOTERMTR INNER JOIN SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSTRANSINFOTERMTR.BRANCH_IP inner join SSINFOTERMBR on SSTRANSINFOTERMTR.BRANCH_CD = SSINFOTERMBR.BRANCH_CD inner join SSINFOTERMCLSTR on SSTRANSINFOTERMTR.CLSTR = SSINFOTERMCLSTR.CLSTR_CD INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSTRANSINFOTERMTR.DIVSN where SSTRANSINFOTERMTR.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "'")
                Dim rpt As New TRANSUSINGGSIS
                cryRpt = rpt
                '  cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)

            ElseIf cbSTrans.Text = "Transactions using GSIS Card" = True And rbAll.Checked = True Then
                GetNameBranch()
                GetName = "Transactions using GSIS Card"

                'dt = db.ExecuteSQLQuery("SELECT SSNUM,DOBTH,DOCVRG,UMID_BANK,BANK_BR,TRANS_ACCTNO,STREET,BRGAY,POST_CD,EMAIL,PHONE,CELNO,USRTYP,SSTRANSINFOTERMTR.ENCODE_DT,SSTRANSINFOTERMTR.ENCODE_TME ,SSTRANSINFOTERMTR.BRANCH_IP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSINFOTERMTR.TAG,TRANID FROM SSTRANSINFOTERMTR INNER JOIN SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSTRANSINFOTERMTR.BRANCH_IP inner join SSINFOTERMBR on SSTRANSINFOTERMTR.BRANCH_CD = SSINFOTERMBR.BRANCH_CD inner join SSINFOTERMCLSTR on SSTRANSINFOTERMTR.CLSTR = SSINFOTERMCLSTR.CLSTR_CD INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSTRANSINFOTERMTR.DIVSN where SSTRANSINFOTERMTR.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "'")
                Dim rpt As New TRANSUSINGGSIS
                cryRpt = rpt
                '  cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
#End Region

            ElseIf rbFBK.Checked = True And rbGroup.Checked = True Then
                GetNameBranch()
                GetName = "Kiosk Feedback"
                'dt = db.ExecuteSQLQuery("SELECT NAME,EMAIL,POST_CD,SSRATE1,SSRATE2,SSRATE3,SSRATE4,SSRATE5,SSRATE6,HELP_TAG,COMMNT_TAG,SSTRANSINFOTERMFBKIOSK.ENCODE_DT,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSINFOTERMFBKIOSK.KIOSK_ID,SSTRANSINFOTERMFBKIOSK.BRANCH_CD,SSINFOTERMKIOSK.BRANCH_IP FROM SSTRANSINFOTERMFBKIOSK INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID = SSTRANSINFOTERMFBKIOSK.KIOSK_ID INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSINFOTERMFBKIOSK.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSTRANSINFOTERMFBKIOSK.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSTRANSINFOTERMFBKIOSK.DIVSN where cast(SSTRANSINFOTERMFBKIOSK.ENCODE_DT as DATE) BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "'")
                Dim rpt As New FEEDBACKKIOSK
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf rbFBK.Checked = True And rbCluster.Checked = True Then
                GetNameBranch()
                GetName = "Kiosk Feedback"
                'dt = db.ExecuteSQLQuery("SELECT NAME,EMAIL,POST_CD,SSRATE1,SSRATE2,SSRATE3,SSRATE4,SSRATE5,SSRATE6,HELP_TAG,COMMNT_TAG,SSTRANSINFOTERMFBKIOSK.ENCODE_DT,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSINFOTERMFBKIOSK.KIOSK_ID,SSTRANSINFOTERMFBKIOSK.BRANCH_CD,SSINFOTERMKIOSK.BRANCH_IP FROM SSTRANSINFOTERMFBKIOSK INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID = SSTRANSINFOTERMFBKIOSK.KIOSK_ID INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSINFOTERMFBKIOSK.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSTRANSINFOTERMFBKIOSK.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSTRANSINFOTERMFBKIOSK.DIVSN where cast(SSTRANSINFOTERMFBKIOSK.ENCODE_DT as DATE) BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "'")
                Dim rpt As New FEEDBACKKIOSK
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf rbFBK.Checked = True And rbBranch.Checked = True Then
                GetNameBranch()
                GetName = "Kiosk Feedback"
                'dt = db.ExecuteSQLQuery("SELECT NAME,EMAIL,POST_CD,SSRATE1,SSRATE2,SSRATE3,SSRATE4,SSRATE5,SSRATE6,HELP_TAG,COMMNT_TAG,SSTRANSINFOTERMFBKIOSK.ENCODE_DT,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSINFOTERMFBKIOSK.KIOSK_ID,SSTRANSINFOTERMFBKIOSK.BRANCH_CD,SSINFOTERMKIOSK.BRANCH_IP FROM SSTRANSINFOTERMFBKIOSK INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID = SSTRANSINFOTERMFBKIOSK.KIOSK_ID INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSINFOTERMFBKIOSK.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSTRANSINFOTERMFBKIOSK.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSTRANSINFOTERMFBKIOSK.DIVSN where cast(SSTRANSINFOTERMFBKIOSK.ENCODE_DT as DATE) BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "'")
                Dim rpt As New FEEDBACKKIOSK
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf rbFBK.Checked = True And rbAll.Checked = True Then
                GetNameBranch()
                GetName = "Kiosk Feedback"
                'dt = db.ExecuteSQLQuery("SELECT NAME,EMAIL,POST_CD,SSRATE1,SSRATE2,SSRATE3,SSRATE4,SSRATE5,SSRATE6,HELP_TAG,COMMNT_TAG,SSTRANSINFOTERMFBKIOSK.ENCODE_DT,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSINFOTERMFBKIOSK.KIOSK_ID,SSTRANSINFOTERMFBKIOSK.BRANCH_CD,SSINFOTERMKIOSK.BRANCH_IP FROM SSTRANSINFOTERMFBKIOSK INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID = SSTRANSINFOTERMFBKIOSK.KIOSK_ID INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSINFOTERMFBKIOSK.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSTRANSINFOTERMFBKIOSK.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSTRANSINFOTERMFBKIOSK.DIVSN where cast(SSTRANSINFOTERMFBKIOSK.ENCODE_DT as DATE) BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "'")
                Dim rpt As New FEEDBACKKIOSK
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf rbFBW.Checked = True And rbGroup.Checked = True Then
                GetNameBranch()
                GetName = "Website Feedback"
                'dt = db.ExecuteSQLQuery("SELECT NAME,EMAIL,ADDRESS_TYP,POST_CD, SSRATE1,SSRATE2,SSRATE3,SSRATE4,SSRATE5,SSRATE6,SSRATE7,VST_TAG,REASN_TAG,INFO_TAG,COMMNT_TAG,SSTRANSINFOTERMFB.ENCODE_DT,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSINFOTERMFB.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP FROM SSTRANSINFOTERMFB INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID = SSTRANSINFOTERMFB.KIOSK_ID INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSINFOTERMFB.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSTRANSINFOTERMFB.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSTRANSINFOTERMFB.DIVSN WHERE cast(SSTRANSINFOTERMFB.ENCODE_DT as DATE) BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "'")
                Dim rpt As New FEEDBACKWEB
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf rbFBW.Checked = True And rbCluster.Checked = True Then
                GetNameBranch()
                GetName = "Website Feedback"
                'dt = db.ExecuteSQLQuery("SELECT NAME,EMAIL,ADDRESS_TYP,POST_CD, SSRATE1,SSRATE2,SSRATE3,SSRATE4,SSRATE5,SSRATE6,SSRATE7,VST_TAG,REASN_TAG,INFO_TAG,COMMNT_TAG,SSTRANSINFOTERMFB.ENCODE_DT,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSINFOTERMFB.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP FROM SSTRANSINFOTERMFB INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID = SSTRANSINFOTERMFB.KIOSK_ID INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSINFOTERMFB.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSTRANSINFOTERMFB.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSTRANSINFOTERMFB.DIVSN WHERE cast(SSTRANSINFOTERMFB.ENCODE_DT as DATE) BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "'")
                Dim rpt As New FEEDBACKWEB
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf rbFBW.Checked = True And rbBranch.Checked = True Then
                GetNameBranch()
                GetName = "Website Feedback"
                'dt = db.ExecuteSQLQuery("SELECT NAME,EMAIL,ADDRESS_TYP,POST_CD, SSRATE1,SSRATE2,SSRATE3,SSRATE4,SSRATE5,SSRATE6,SSRATE7,VST_TAG,REASN_TAG,INFO_TAG,COMMNT_TAG,SSTRANSINFOTERMFB.ENCODE_DT,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSINFOTERMFB.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP FROM SSTRANSINFOTERMFB INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID = SSTRANSINFOTERMFB.KIOSK_ID INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSINFOTERMFB.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSTRANSINFOTERMFB.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSTRANSINFOTERMFB.DIVSN WHERE cast(SSTRANSINFOTERMFB.ENCODE_DT as DATE) BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "'")
                Dim rpt As New FEEDBACKWEB
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf rbFBW.Checked = True And rbAll.Checked = True Then
                GetNameBranch()
                GetName = "Website Feedback"
                'dt = db.ExecuteSQLQuery("SELECT NAME,EMAIL,ADDRESS_TYP,POST_CD, SSRATE1,SSRATE2,SSRATE3,SSRATE4,SSRATE5,SSRATE6,SSRATE7,VST_TAG,REASN_TAG,INFO_TAG,COMMNT_TAG,SSTRANSINFOTERMFB.ENCODE_DT,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSINFOTERMFB.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP FROM SSTRANSINFOTERMFB INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID = SSTRANSINFOTERMFB.KIOSK_ID INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSINFOTERMFB.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSTRANSINFOTERMFB.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSTRANSINFOTERMFB.DIVSN WHERE cast(SSTRANSINFOTERMFB.ENCODE_DT as DATE) BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "'")
                Dim rpt As New FEEDBACKWEB
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf cbSTrans.Text = "Status updates of retiree pensioner dependents" And rbGroup.Checked = True Then
                GetNameBranch()
                GetName = "Status updates of retiree pensioner dependents"
                'dt = db.ExecuteSQLQuery("SELECT SSTRANSACOPAD.SSNUM,SSTRANSACOPAD.TRANID,SSTRANSACOPAD.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP, IKBENEFREPTF.REFNO, TYPOFRPT, SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM, SSTRANSACOPAD.ENCODE_DT ,SSTRANSACOPAD.ENCODE_TME, DPDNAME,DPDDATE,case when DPDSTAT = 1 then 'DATE OF MARRIAGE' when DPDSTAT = 2 then 'DATE OF DEATH' when DPDSTAT = 3 then 'DATE OF EMPLOYMENT' else 'NO STATUS CHANGE' end as [type] from SSTRANSACOPAD INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSACOPAD.BRANCH_CD INNER JOIN IKBENEFREPTF ON SSTRANSACOPAD.TRANID = IKBENEFREPTF.TRANID INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMBR.CLSTR_CD = SSINFOTERMCLSTR.CLSTR_CD INNER JOIN SSINFOTERMGROUP ON SSINFOTERMBR.GROUP_CD = SSINFOTERMGROUP.GROUP_CD INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID = SSTRANSACOPAD.KIOSK_ID where SSTRANSACOPAD.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' ORDER BY SSTRANSACOPAD.ENCODE_DT")
                Dim rpt As New ACOPDEPALL
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)

            ElseIf cbSTrans.Text = "Status updates of retiree pensioner dependents" And rbCluster.Checked = True Then
                GetNameBranch()
                GetName = "Status updates of retiree pensioner dependents"
                'dt = db.ExecuteSQLQuery("SELECT SSTRANSACOPAD.SSNUM,SSTRANSACOPAD.TRANID,SSTRANSACOPAD.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP, IKBENEFREPTF.REFNO, TYPOFRPT, SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM, SSTRANSACOPAD.ENCODE_DT ,SSTRANSACOPAD.ENCODE_TME, DPDNAME,DPDDATE,case when DPDSTAT = 1 then 'DATE OF MARRIAGE' when DPDSTAT = 2 then 'DATE OF DEATH' when DPDSTAT = 3 then 'DATE OF EMPLOYMENT' else 'NO STATUS CHANGE' end as [type] from SSTRANSACOPAD INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSACOPAD.BRANCH_CD INNER JOIN IKBENEFREPTF ON SSTRANSACOPAD.TRANID = IKBENEFREPTF.TRANID INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMBR.CLSTR_CD = SSINFOTERMCLSTR.CLSTR_CD INNER JOIN SSINFOTERMGROUP ON SSINFOTERMBR.GROUP_CD = SSINFOTERMGROUP.GROUP_CD INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID = SSTRANSACOPAD.KIOSK_ID where SSTRANSACOPAD.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' ORDER BY SSTRANSACOPAD.ENCODE_DT")
                Dim rpt As New ACOPDEPALL
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf cbSTrans.Text = "Status updates of retiree pensioner dependents" And rbBranch.Checked = True Then
                GetNameBranch()
                GetName = "Status updates of retiree pensioner dependents"
                'dt = db.ExecuteSQLQuery("SELECT SSTRANSACOPAD.SSNUM,SSTRANSACOPAD.TRANID,SSTRANSACOPAD.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP, IKBENEFREPTF.REFNO, TYPOFRPT, SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM, SSTRANSACOPAD.ENCODE_DT ,SSTRANSACOPAD.ENCODE_TME, DPDNAME,DPDDATE,case when DPDSTAT = 1 then 'DATE OF MARRIAGE' when DPDSTAT = 2 then 'DATE OF DEATH' when DPDSTAT = 3 then 'DATE OF EMPLOYMENT' else 'NO STATUS CHANGE' end as [type] from SSTRANSACOPAD INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSACOPAD.BRANCH_CD INNER JOIN IKBENEFREPTF ON SSTRANSACOPAD.TRANID = IKBENEFREPTF.TRANID INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMBR.CLSTR_CD = SSINFOTERMCLSTR.CLSTR_CD INNER JOIN SSINFOTERMGROUP ON SSINFOTERMBR.GROUP_CD = SSINFOTERMGROUP.GROUP_CD INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID = SSTRANSACOPAD.KIOSK_ID where SSTRANSACOPAD.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' ORDER BY SSTRANSACOPAD.ENCODE_DT")
                Dim rpt As New ACOPDEPALL
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf cbSTrans.Text = "Status updates of retiree pensioner dependents" And rbAll.Checked = True Then
                GetNameBranch()
                GetName = "Status updates of retiree pensioner dependents"
                'dt = db.ExecuteSQLQuery("SELECT SSTRANSACOPAD.SSNUM,SSTRANSACOPAD.TRANID,SSTRANSACOPAD.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP, IKBENEFREPTF.REFNO, TYPOFRPT, SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM, SSTRANSACOPAD.ENCODE_DT ,SSTRANSACOPAD.ENCODE_TME, DPDNAME,DPDDATE,case when DPDSTAT = 1 then 'DATE OF MARRIAGE' when DPDSTAT = 2 then 'DATE OF DEATH' when DPDSTAT = 3 then 'DATE OF EMPLOYMENT' else 'NO STATUS CHANGE' end as [type] from SSTRANSACOPAD INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSACOPAD.BRANCH_CD INNER JOIN IKBENEFREPTF ON SSTRANSACOPAD.TRANID = IKBENEFREPTF.TRANID INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMBR.CLSTR_CD = SSINFOTERMCLSTR.CLSTR_CD INNER JOIN SSINFOTERMGROUP ON SSINFOTERMBR.GROUP_CD = SSINFOTERMGROUP.GROUP_CD INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID = SSTRANSACOPAD.KIOSK_ID where SSTRANSACOPAD.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' ORDER BY SSTRANSACOPAD.ENCODE_DT")
                Dim rpt As New ACOPDEPALL
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf cbSTrans.Text = "Updating Contact Information" And rbBranch.Checked = True Then
                GetNameBranch()
                GetName = "Updating Contact Information"
                'dt = db.ExecuteSQLQuery("SELECT SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP,SSNUM,HOUSENO +','+ isnull(STREET,'') +','+ SUBDIVISION +','+ CITY as [address],ZIPCODE,LANDLINE,MOBILE,EMAIL,SSTRANSPM.ENCODE_DT,ENCODE_TME,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM FROM SSTRANSPM INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID = SSTRANSPM.KIOSK_ID INNER JOIN SSINFOTERMBR ON SSTRANSPM.BRANCH_CD = SSINFOTERMBR.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSTRANSPM.CLSTR = SSINFOTERMCLSTR.CLSTR_CD INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSTRANSPM.DIVSN where SSTRANSPM.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "'")
                Dim rpt As New CHANGEOFADDALL
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf cbSTrans.Text = "Updating Contact Information" And rbGroup.Checked = True Then
                GetNameBranch()
                GetName = "Updating Contact Information"
                'dt = db.ExecuteSQLQuery("SELECT SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP,SSNUM,HOUSENO +','+ isnull(STREET,'') +','+ SUBDIVISION +','+ CITY as [address],ZIPCODE,LANDLINE,MOBILE,EMAIL,SSTRANSPM.ENCODE_DT,ENCODE_TME,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM FROM SSTRANSPM INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID = SSTRANSPM.KIOSK_ID INNER JOIN SSINFOTERMBR ON SSTRANSPM.BRANCH_CD = SSINFOTERMBR.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSTRANSPM.CLSTR = SSINFOTERMCLSTR.CLSTR_CD INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSTRANSPM.DIVSN where SSTRANSPM.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "'")
                Dim rpt As New CHANGEOFADDALL
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf cbSTrans.Text = "Updating Contact Information" And rbCluster.Checked = True Then
                GetNameBranch()
                GetName = "Updating Contact Information"
                'dt = db.ExecuteSQLQuery("SELECT SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP,SSNUM,HOUSENO +','+ isnull(STREET,'') +','+ SUBDIVISION +','+ CITY as [address],ZIPCODE,LANDLINE,MOBILE,EMAIL,SSTRANSPM.ENCODE_DT,ENCODE_TME,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM FROM SSTRANSPM INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID = SSTRANSPM.KIOSK_ID INNER JOIN SSINFOTERMBR ON SSTRANSPM.BRANCH_CD = SSINFOTERMBR.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSTRANSPM.CLSTR = SSINFOTERMCLSTR.CLSTR_CD INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSTRANSPM.DIVSN where SSTRANSPM.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "'")
                Dim rpt As New CHANGEOFADDALL
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf cbSTrans.Text = "Updating Contact Information" And rbAll.Checked = True Then
                GetNameBranch()
                GetName = "Updating Contact Information"
                'dt = db.ExecuteSQLQuery("SELECT SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP,SSNUM,HOUSENO +','+ isnull(STREET,'') +','+ SUBDIVISION +','+ CITY as [address],ZIPCODE,LANDLINE,MOBILE,EMAIL,SSTRANSPM.ENCODE_DT,ENCODE_TME,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM FROM SSTRANSPM INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID = SSTRANSPM.KIOSK_ID INNER JOIN SSINFOTERMBR ON SSTRANSPM.BRANCH_CD = SSINFOTERMBR.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSTRANSPM.CLSTR = SSINFOTERMCLSTR.CLSTR_CD INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSTRANSPM.DIVSN where SSTRANSPM.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "'")
                Dim rpt As New CHANGEOFADDALL
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf cbFTrans.Text = "Simplified Web Registration" And rbBranch.Checked = True Then
                GetNameBranch()
                GetName = "Failed Registration"
                'dt = db.ExecuteSQLQuery("select COUNT(PROC_CD) as total,SSTRANSAT.ENCODE_DT,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP from SSTRANSAT INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAT.BRANCH_IP where SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and TRANS_DESC <> '' and PROC_CD = '10001' and TRANS_DESC <> '' and PROC_CD = '10001'GROUP BY SSTRANSAT.ENCODE_DT,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP")
                Dim rpt As New FREGALL
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf cbFTrans.Text = "Simplified Web Registration" And rbGroup.Checked = True Then
                GetNameBranch()
                GetName = "Failed Registration"
                'dt = db.ExecuteSQLQuery("select COUNT(PROC_CD) as total,SSTRANSAT.ENCODE_DT,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP from SSTRANSAT INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAT.BRANCH_IP where SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and TRANS_DESC <> '' and PROC_CD = '10001' and TRANS_DESC <> '' and PROC_CD = '10001'GROUP BY SSTRANSAT.ENCODE_DT,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP")
                Dim rpt As New FREGALL
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf cbFTrans.Text = "Simplified Web Registration" And rbCluster.Checked = True Then
                GetNameBranch()
                GetName = "Failed Registration"
                'dt = db.ExecuteSQLQuery("select COUNT(PROC_CD) as total,SSTRANSAT.ENCODE_DT,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP from SSTRANSAT INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAT.BRANCH_IP where SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and TRANS_DESC <> '' and PROC_CD = '10001' and TRANS_DESC <> '' and PROC_CD = '10001'GROUP BY SSTRANSAT.ENCODE_DT,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP")
                Dim rpt As New FREGALL
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf cbFTrans.Text = "Simplified Web Registration" And rbAll.Checked = True Then
                GetNameBranch()
                GetName = "Failed Registration"
                'dt = db.ExecuteSQLQuery("select COUNT(PROC_CD) as total,SSTRANSAT.ENCODE_DT,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP from SSTRANSAT INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAT.BRANCH_IP where SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and TRANS_DESC <> '' and PROC_CD = '10001' and TRANS_DESC <> '' and PROC_CD = '10001'GROUP BY SSTRANSAT.ENCODE_DT,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP")
                Dim rpt As New FREGALL
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf cbFTrans.Text <> "" And rbBranch.Checked = True Then
                TransFailed()
                GetNameBranch()
                GetName = "Failed Transaction"
                'dt = db.ExecuteSQLQuery("SELECT SSNUM,case when SSINFOTERMPROCCD.PROC_NM = 'MEMBER MATERNITY NOTIFICATION' then 'MATERNITY NOTIFICATION' when SSINFOTERMPROCCD.PROC_NM = 'LOAN ELIGIBILITY' then 'SALARY LOAN' else SSINFOTERMPROCCD.PROC_NM end as [Proc],SSTRANSAT.ENCODE_DT,SSTRANSAT.ENCODE_TME,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN,TRANS_DESC from SSTRANSAT INNER JOIN SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAT.BRANCH_IP INNER JOIN SSINFOTERMPROCCD ON SSINFOTERMPROCCD.PROC_CD = SSTRANSAT.PROC_CD where SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and SSINFOTERMPROCCD.PROC_NM = '" & Failedtrans & "' and TRANS_DESC <> '' ORDER by SSTRANSAT.ENCODE_DT, SSTRANSAT.ENCODE_TME")
                Dim rpt As New FAILEDTRANSALL
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@getProc", Failedtrans)
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf cbFTrans.Text <> "" And rbGroup.Checked = True Then
                TransFailed()
                GetNameBranch()
                GetName = "Failed Transaction"
                'dt = db.ExecuteSQLQuery("SELECT SSNUM,case when SSINFOTERMPROCCD.PROC_NM = 'MEMBER MATERNITY NOTIFICATION' then 'MATERNITY NOTIFICATION' when SSINFOTERMPROCCD.PROC_NM = 'LOAN ELIGIBILITY' then 'SALARY LOAN' else SSINFOTERMPROCCD.PROC_NM end as [Proc],SSTRANSAT.ENCODE_DT,SSTRANSAT.ENCODE_TME,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN,TRANS_DESC from SSTRANSAT INNER JOIN SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAT.BRANCH_IP INNER JOIN SSINFOTERMPROCCD ON SSINFOTERMPROCCD.PROC_CD = SSTRANSAT.PROC_CD where SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and SSINFOTERMPROCCD.PROC_NM = '" & Failedtrans & "' and TRANS_DESC <> '' ORDER by SSTRANSAT.ENCODE_DT, SSTRANSAT.ENCODE_TME")
                Dim rpt As New FAILEDTRANSALL
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@getProc", Failedtrans)
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf cbFTrans.Text <> "" And rbCluster.Checked = True Then
                TransFailed()
                GetNameBranch()
                GetName = "Failed Transaction"
                'dt = db.ExecuteSQLQuery("SELECT SSNUM,case when SSINFOTERMPROCCD.PROC_NM = 'MEMBER MATERNITY NOTIFICATION' then 'MATERNITY NOTIFICATION' when SSINFOTERMPROCCD.PROC_NM = 'LOAN ELIGIBILITY' then 'SALARY LOAN' else SSINFOTERMPROCCD.PROC_NM end as [Proc],SSTRANSAT.ENCODE_DT,SSTRANSAT.ENCODE_TME,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN,TRANS_DESC from SSTRANSAT INNER JOIN SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAT.BRANCH_IP INNER JOIN SSINFOTERMPROCCD ON SSINFOTERMPROCCD.PROC_CD = SSTRANSAT.PROC_CD where SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and SSINFOTERMPROCCD.PROC_NM = '" & Failedtrans & "' and TRANS_DESC <> '' ORDER by SSTRANSAT.ENCODE_DT, SSTRANSAT.ENCODE_TME")
                Dim rpt As New FAILEDTRANSALL
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@getProc", Failedtrans)
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf cbFTrans.Text <> "" And rbAll.Checked = True Then
                TransFailed()
                GetNameBranch()
                GetName = "Failed Transaction"
                'dt = db.ExecuteSQLQuery("SELECT SSNUM,case when SSINFOTERMPROCCD.PROC_NM = 'MEMBER MATERNITY NOTIFICATION' then 'MATERNITY NOTIFICATION' when SSINFOTERMPROCCD.PROC_NM = 'LOAN ELIGIBILITY' then 'SALARY LOAN' else SSINFOTERMPROCCD.PROC_NM end as [Proc],SSTRANSAT.ENCODE_DT,SSTRANSAT.ENCODE_TME,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN,TRANS_DESC from SSTRANSAT INNER JOIN SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAT.BRANCH_IP INNER JOIN SSINFOTERMPROCCD ON SSINFOTERMPROCCD.PROC_CD = SSTRANSAT.PROC_CD where SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and SSINFOTERMPROCCD.PROC_NM = '" & Failedtrans & "' and TRANS_DESC <> '' ORDER by SSTRANSAT.ENCODE_DT, SSTRANSAT.ENCODE_TME")
                Dim rpt As New FAILEDTRANSALL
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@getProc", Failedtrans)
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf cbSTrans.Text = "Simplified Web Registration" And rbBranch.Checked = True Then
                GetNameBranch()
                GetName = "Success Registration"
                'dt = db.ExecuteSQLQuery("SELECT SSTRANSINFOTERMRG.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP,SSNUM,givName + ' ' + midName + ' ' + surName [FULLNAME],address1 + ', ' + address2 + ', ' + cityProv [address],postal,mobile, emailAdd,memstatus,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSINFOTERMRG.ENCODE_DT FROM SSTRANSINFOTERMRG INNER JOIN SSINFOTERMKIOSK ON SSTRANSINFOTERMRG.KIOSK_ID = SSINFOTERMKIOSK.KIOSK_ID INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSINFOTERMRG.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSTRANSINFOTERMRG.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD= SSTRANSINFOTERMRG.DIVSN WHERE SSTRANSINFOTERMRG.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "'")
                Dim rpt As New REGALL
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf cbSTrans.Text = "Simplified Web Registration" And rbGroup.Checked = True Then
                GetNameBranch()
                GetName = "Success Registration"
                'dt = db.ExecuteSQLQuery("SELECT SSTRANSINFOTERMRG.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP,SSNUM,givName + ' ' + midName + ' ' + surName [FULLNAME],address1 + ', ' + address2 + ', ' + cityProv [address],postal,mobile, emailAdd,memstatus,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSINFOTERMRG.ENCODE_DT FROM SSTRANSINFOTERMRG INNER JOIN SSINFOTERMKIOSK ON SSTRANSINFOTERMRG.KIOSK_ID = SSINFOTERMKIOSK.KIOSK_ID INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSINFOTERMRG.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSTRANSINFOTERMRG.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD= SSTRANSINFOTERMRG.DIVSN WHERE SSTRANSINFOTERMRG.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "'")
                Dim rpt As New REGALL
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf cbSTrans.Text = "Simplified Web Registration" And rbCluster.Checked = True Then
                GetNameBranch()
                GetName = "Success Registration"
                'dt = db.ExecuteSQLQuery("SELECT SSTRANSINFOTERMRG.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP,SSNUM,givName + ' ' + midName + ' ' + surName [FULLNAME],address1 + ', ' + address2 + ', ' + cityProv [address],postal,mobile, emailAdd,memstatus,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSINFOTERMRG.ENCODE_DT FROM SSTRANSINFOTERMRG INNER JOIN SSINFOTERMKIOSK ON SSTRANSINFOTERMRG.KIOSK_ID = SSINFOTERMKIOSK.KIOSK_ID INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSINFOTERMRG.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSTRANSINFOTERMRG.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD= SSTRANSINFOTERMRG.DIVSN WHERE SSTRANSINFOTERMRG.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "'")
                Dim rpt As New REGALL
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf cbSTrans.Text = "Simplified Web Registration" And rbAll.Checked = True Then
                GetNameBranch()
                GetName = "Success Registration"
                'dt = db.ExecuteSQLQuery("SELECT SSTRANSINFOTERMRG.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP,SSNUM,givName + ' ' + midName + ' ' + surName [FULLNAME],address1 + ', ' + address2 + ', ' + cityProv [address],postal,mobile, emailAdd,memstatus,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSINFOTERMRG.ENCODE_DT FROM SSTRANSINFOTERMRG INNER JOIN SSINFOTERMKIOSK ON SSTRANSINFOTERMRG.KIOSK_ID = SSINFOTERMKIOSK.KIOSK_ID INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSINFOTERMRG.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSTRANSINFOTERMRG.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD= SSTRANSINFOTERMRG.DIVSN WHERE SSTRANSINFOTERMRG.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "'")
                Dim rpt As New REGALL
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf rbCharter.Checked = True And rbBranch.Checked = True Then
                GetNameBranch()
                GetName = "Citizens Charter"
                'dt = db.ExecuteSQLQuery("SELECT convert(varchar,SSTRANSAT.ENCODE_DT,101) as DateOfTrans ,case when SSINFOTERMPROCCD.PROC_NM = 'SSID CLEARANCE' then 'SSS/UMID CARD INFO' else SSINFOTERMPROCCD.PROC_NM end as 'PROC',SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN, COUNT(SSINFOTERMPROCCD.PROC_NM) as [count] FROM SSTRANSAT INNER JOIN SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAT.BRANCH_IP INNER JOIN SSINFOTERMPROCCD ON SSINFOTERMPROCCD.PROC_CD = SSTRANSAT.PROC_CD and SSINFOTERMPROCCD.PROC_NM = 'CITIZENS CHARTER' and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' GROUP by SSTRANSAT.ENCODE_DT,SSINFOTERMPROCCD.PROC_NM,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN")
                Dim rpt As New CITIZENALL
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf rbCharter.Checked = True And rbGroup.Checked = True Then
                GetNameBranch()
                GetName = "Citizens Charter"
                'dt = db.ExecuteSQLQuery("SELECT convert(varchar,SSTRANSAT.ENCODE_DT,101) as DateOfTrans ,case when SSINFOTERMPROCCD.PROC_NM = 'SSID CLEARANCE' then 'SSS/UMID CARD INFO' else SSINFOTERMPROCCD.PROC_NM end as 'PROC',SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN, COUNT(SSINFOTERMPROCCD.PROC_NM) as [count] FROM SSTRANSAT INNER JOIN SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAT.BRANCH_IP INNER JOIN SSINFOTERMPROCCD ON SSINFOTERMPROCCD.PROC_CD = SSTRANSAT.PROC_CD and SSINFOTERMPROCCD.PROC_NM = 'CITIZENS CHARTER' and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' GROUP by SSTRANSAT.ENCODE_DT,SSINFOTERMPROCCD.PROC_NM,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN")
                Dim rpt As New CITIZENALL
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf rbCharter.Checked = True And rbCluster.Checked = True Then
                GetNameBranch()
                GetName = "Citizens Charter"
                'dt = db.ExecuteSQLQuery("SELECT convert(varchar,SSTRANSAT.ENCODE_DT,101) as DateOfTrans ,case when SSINFOTERMPROCCD.PROC_NM = 'SSID CLEARANCE' then 'SSS/UMID CARD INFO' else SSINFOTERMPROCCD.PROC_NM end as 'PROC',SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN, COUNT(SSINFOTERMPROCCD.PROC_NM) as [count] FROM SSTRANSAT INNER JOIN SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAT.BRANCH_IP INNER JOIN SSINFOTERMPROCCD ON SSINFOTERMPROCCD.PROC_CD = SSTRANSAT.PROC_CD and SSINFOTERMPROCCD.PROC_NM = 'CITIZENS CHARTER' and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' GROUP by SSTRANSAT.ENCODE_DT,SSINFOTERMPROCCD.PROC_NM,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN")
                Dim rpt As New CITIZENALL
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf rbCharter.Checked = True And rbAll.Checked = True Then
                GetNameBranch()
                GetName = "Citizens Charter"
                'dt = db.ExecuteSQLQuery("SELECT convert(varchar,SSTRANSAT.ENCODE_DT,101) as DateOfTrans ,case when SSINFOTERMPROCCD.PROC_NM = 'SSID CLEARANCE' then 'SSS/UMID CARD INFO' else SSINFOTERMPROCCD.PROC_NM end as 'PROC',SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN, COUNT(SSINFOTERMPROCCD.PROC_NM) as [count] FROM SSTRANSAT INNER JOIN SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAT.BRANCH_IP INNER JOIN SSINFOTERMPROCCD ON SSINFOTERMPROCCD.PROC_CD = SSTRANSAT.PROC_CD and SSINFOTERMPROCCD.PROC_NM = 'CITIZENS CHARTER' and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' GROUP by SSTRANSAT.ENCODE_DT,SSINFOTERMPROCCD.PROC_NM,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN")
                Dim rpt As New CITIZENALL
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf rbUpdates.Checked = True And rbAll.Checked = True Then
                GetName = "UPDATES"
                'dt = db.ExecuteSQLQuery("SELECT KIOSK_ID,KIOSK_IP,DATE_UPDATED,SSINFOTERMBR.BRANCH_NM,case when STATUS = 1 then 'SUCCESS' else 'FAILED' end as [Status], CASE WHEN tag = 1 then 'CITIZEN CHARTER' when tag = 2 then 'SSS VIDEO' when tag = 3 then 'SSIT APPLICATION' when tag = 4 then 'SSIT SETTINGS' when tag = 5 then 'TERMS AND CONDITION' when tag = 6 then 'SSIT UPDATER' end as [ITEM UPDATE] FROM SSUPDATELOGS INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSUPDATELOGS.KIOSK_BRANCH WHERE cast(DATE_UPDATED as date) BETWEEN '" & date1.ToShortDateString & "' AND '" & date2.ToShortDateString & "'")
                Dim rpt As New UPDATE
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
            ElseIf rbUpdates.Checked = True And rbBranch.Checked = True Then
                GetName = "UPDATES"
                'dt = db.ExecuteSQLQuery("SELECT KIOSK_ID,KIOSK_IP,DATE_UPDATED,SSINFOTERMBR.BRANCH_NM,case when STATUS = 1 then 'SUCCESS' else 'FAILED' end as [Status], CASE WHEN tag = 1 then 'CITIZEN CHARTER' when tag = 2 then 'SSS VIDEO' when tag = 3 then 'SSIT APPLICATION' when tag = 4 then 'SSIT SETTINGS' when tag = 5 then 'TERMS AND CONDITION' when tag = 6 then 'SSIT UPDATER' end as [ITEM UPDATE] FROM SSUPDATELOGS INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSUPDATELOGS.KIOSK_BRANCH WHERE cast(DATE_UPDATED as date) BETWEEN '" & date1.ToShortDateString & "' AND '" & date2.ToShortDateString & "' and SSINFOTERMBR.BRANCH_NM = '" & cbBranch.Text & "'")
                Dim rpt As New UPDATEBR
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@getBranch", cbBranch.Text)
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
            ElseIf rbUpdates.Checked = True And rbCluster.Checked = True Then
                GetName = "UPDATES"
                'dt = db.ExecuteSQLQuery("SELECT KIOSK_ID,KIOSK_IP,DATE_UPDATED,SSINFOTERMBR.BRANCH_NM,case when STATUS = 1 then 'SUCCESS' else 'FAILED' end as [Status], CASE WHEN tag = 1 then 'CITIZEN CHARTER' when tag = 2 then 'SSS VIDEO' when tag = 3 then 'SSIT APPLICATION' when tag = 4 then 'SSIT SETTINGS' when tag = 5 then 'TERMS AND CONDITION' when tag = 6 then 'SSIT UPDATER' end as [ITEM UPDATE] FROM SSUPDATELOGS INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSUPDATELOGS.KIOSK_BRANCH INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMBR.CLSTR_CD WHERE cast(DATE_UPDATED as date) BETWEEN '" & date1.ToShortDateString & "' AND '" & date2.ToShortDateString & "' and SSINFOTERMCLSTR.CLSTR_NM = '" & cbCluster.Text & "'")
                Dim rpt As New UPDATECLUST
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@getCluster", cbCluster.Text)
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
            ElseIf rbUpdates.Checked = True And rbGroup.Checked = True Then
                GetName = "UPDATES"
                'dt = db.ExecuteSQLQuery("SELECT KIOSK_ID,KIOSK_IP,DATE_UPDATED,SSINFOTERMBR.BRANCH_NM,case when STATUS = 1 then 'SUCCESS' else 'FAILED' end as [Status], CASE WHEN tag = 1 then 'CITIZEN CHARTER' when tag = 2 then 'SSS VIDEO' when tag = 3 then 'SSIT APPLICATION' when tag = 4 then 'SSIT SETTINGS' when tag = 5 then 'TERMS AND CONDITION' when tag = 6 then 'SSIT UPDATER' end as [ITEM UPDATE] FROM SSUPDATELOGS INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSUPDATELOGS.KIOSK_BRANCH INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSINFOTERMBR.GROUP_CD WHERE cast(DATE_UPDATED as date) BETWEEN '" & date1.ToShortDateString & "' AND '" & date2.ToShortDateString & "' and SSINFOTERMGROUP.GROUP_NM = '" & cbGroup.Text & "'")
                Dim rpt As New UPDATEGROUP
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@getGroup", cbGroup.Text)
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)


            ElseIf rbSMAT.Checked = True And rbCluster.Checked = True Then
                GetNameBranch()
                GetName = "SicknessMaternity"
                'dt = db.ExecuteSQLQuery("SELECT convert(varchar,SSTRANSAT.ENCODE_DT,101) as DateOfTrans,case when SSINFOTERMPROCCD.PROC_NM = 'SSID CLEARANCE' then 'SSS/UMID CARD INFO' else SSINFOTERMPROCCD.PROC_NM end as 'PROC',SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN, COUNT(SSINFOTERMPROCCD.PROC_NM) as [count] FROM SSTRANSAT INNER JOIN SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAT.BRANCH_IP INNER JOIN SSINFOTERMPROCCD ON SSINFOTERMPROCCD.PROC_CD = SSTRANSAT.PROC_CD where SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and SSTRANSAT.PROC_CD not like '%10001%' and SSTRANSAT.PROC_CD not like '%10002%' and SSTRANSAT.PROC_CD not like '%10003%' and SSTRANSAT.PROC_CD not like '%10004%' and SSTRANSAT.PROC_CD not like '%10005%' and SSTRANSAT.PROC_CD not like '%10006%' and SSTRANSAT.PROC_CD not like '%10007%' and SSTRANSAT.PROC_CD not like '%10008%' and SSTRANSAT.PROC_CD not like '%10009%' and SSTRANSAT.PROC_CD not like '%10010%' and SSTRANSAT.PROC_CD not like '%10011%' and SSTRANSAT.PROC_CD not like '%10012%' and SSTRANSAT.PROC_CD not like '%10013%' and SSTRANSAT.PROC_CD not like '%10014%' and SSTRANSAT.PROC_CD not like '%10015%' and SSTRANSAT.PROC_CD not like '%10016%' and SSTRANSAT.PROC_CD not like '%10017%' and SSTRANSAT.PROC_CD not like '%10018%' and SSTRANSAT.PROC_CD not like '%10019%' and SSTRANSAT.PROC_CD not like '%10020%' and SSTRANSAT.PROC_CD not like '%10021%' and SSTRANSAT.PROC_CD not like '%10022%' and SSTRANSAT.PROC_CD not like '%10023%' and SSTRANSAT.PROC_CD not like '%10024%' and SSTRANSAT.PROC_CD not like '%10025%' and SSTRANSAT.PROC_CD not like '%10026%' and SSTRANSAT.PROC_CD not like '%10027%' and SSTRANSAT.PROC_CD not like '%10028%' and SSTRANSAT.PROC_CD not like '%10029%' and SSTRANSAT.PROC_CD not like '%10030%' and SSTRANSAT.PROC_CD not like '%10031%' and SSTRANSAT.PROC_CD not like '%10032%' and SSTRANSAT.PROC_CD not like '%10033%' and SSTRANSAT.PROC_CD not like '%10034%' and SSTRANSAT.PROC_CD not like '%10035%' and SSTRANSAT.PROC_CD not like '%10036%' and SSTRANSAT.PROC_CD not like '%10041%' and SSTRANSAT.PROC_CD not like '%10042%' GROUP by SSTRANSAT.ENCODE_DT,SSINFOTERMPROCCD.PROC_NM,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN")
                Dim rpt As New SICKMATALL
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf rbSMAT.Checked = True And rbGroup.Checked = True Then
                GetNameBranch()
                GetName = "SicknessMaternity"
                'dt = db.ExecuteSQLQuery("SELECT convert(varchar,SSTRANSAT.ENCODE_DT,101) as DateOfTrans,case when SSINFOTERMPROCCD.PROC_NM = 'SSID CLEARANCE' then 'SSS/UMID CARD INFO' else SSINFOTERMPROCCD.PROC_NM end as 'PROC',SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN, COUNT(SSINFOTERMPROCCD.PROC_NM) as [count] FROM SSTRANSAT INNER JOIN SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAT.BRANCH_IP INNER JOIN SSINFOTERMPROCCD ON SSINFOTERMPROCCD.PROC_CD = SSTRANSAT.PROC_CD where SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and SSTRANSAT.PROC_CD not like '%10001%' and SSTRANSAT.PROC_CD not like '%10002%' and SSTRANSAT.PROC_CD not like '%10003%' and SSTRANSAT.PROC_CD not like '%10004%' and SSTRANSAT.PROC_CD not like '%10005%' and SSTRANSAT.PROC_CD not like '%10006%' and SSTRANSAT.PROC_CD not like '%10007%' and SSTRANSAT.PROC_CD not like '%10008%' and SSTRANSAT.PROC_CD not like '%10009%' and SSTRANSAT.PROC_CD not like '%10010%' and SSTRANSAT.PROC_CD not like '%10011%' and SSTRANSAT.PROC_CD not like '%10012%' and SSTRANSAT.PROC_CD not like '%10013%' and SSTRANSAT.PROC_CD not like '%10014%' and SSTRANSAT.PROC_CD not like '%10015%' and SSTRANSAT.PROC_CD not like '%10016%' and SSTRANSAT.PROC_CD not like '%10017%' and SSTRANSAT.PROC_CD not like '%10018%' and SSTRANSAT.PROC_CD not like '%10019%' and SSTRANSAT.PROC_CD not like '%10020%' and SSTRANSAT.PROC_CD not like '%10021%' and SSTRANSAT.PROC_CD not like '%10022%' and SSTRANSAT.PROC_CD not like '%10023%' and SSTRANSAT.PROC_CD not like '%10024%' and SSTRANSAT.PROC_CD not like '%10025%' and SSTRANSAT.PROC_CD not like '%10026%' and SSTRANSAT.PROC_CD not like '%10027%' and SSTRANSAT.PROC_CD not like '%10028%' and SSTRANSAT.PROC_CD not like '%10029%' and SSTRANSAT.PROC_CD not like '%10030%' and SSTRANSAT.PROC_CD not like '%10031%' and SSTRANSAT.PROC_CD not like '%10032%' and SSTRANSAT.PROC_CD not like '%10033%' and SSTRANSAT.PROC_CD not like '%10034%' and SSTRANSAT.PROC_CD not like '%10035%' and SSTRANSAT.PROC_CD not like '%10036%' and SSTRANSAT.PROC_CD not like '%10041%' and SSTRANSAT.PROC_CD not like '%10042%' GROUP by SSTRANSAT.ENCODE_DT,SSINFOTERMPROCCD.PROC_NM,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN")
                Dim rpt As New SICKMATALL
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf rbSMAT.Checked = True And rbBranch.Checked = True Then
                GetNameBranch()
                GetName = "SicknessMaternity"
                'dt = db.ExecuteSQLQuery("SELECT convert(varchar,SSTRANSAT.ENCODE_DT,101) as DateOfTrans,case when SSINFOTERMPROCCD.PROC_NM = 'SSID CLEARANCE' then 'SSS/UMID CARD INFO' else SSINFOTERMPROCCD.PROC_NM end as 'PROC',SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN, COUNT(SSINFOTERMPROCCD.PROC_NM) as [count] FROM SSTRANSAT INNER JOIN SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAT.BRANCH_IP INNER JOIN SSINFOTERMPROCCD ON SSINFOTERMPROCCD.PROC_CD = SSTRANSAT.PROC_CD where SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and SSTRANSAT.PROC_CD not like '%10001%' and SSTRANSAT.PROC_CD not like '%10002%' and SSTRANSAT.PROC_CD not like '%10003%' and SSTRANSAT.PROC_CD not like '%10004%' and SSTRANSAT.PROC_CD not like '%10005%' and SSTRANSAT.PROC_CD not like '%10006%' and SSTRANSAT.PROC_CD not like '%10007%' and SSTRANSAT.PROC_CD not like '%10008%' and SSTRANSAT.PROC_CD not like '%10009%' and SSTRANSAT.PROC_CD not like '%10010%' and SSTRANSAT.PROC_CD not like '%10011%' and SSTRANSAT.PROC_CD not like '%10012%' and SSTRANSAT.PROC_CD not like '%10013%' and SSTRANSAT.PROC_CD not like '%10014%' and SSTRANSAT.PROC_CD not like '%10015%' and SSTRANSAT.PROC_CD not like '%10016%' and SSTRANSAT.PROC_CD not like '%10017%' and SSTRANSAT.PROC_CD not like '%10018%' and SSTRANSAT.PROC_CD not like '%10019%' and SSTRANSAT.PROC_CD not like '%10020%' and SSTRANSAT.PROC_CD not like '%10021%' and SSTRANSAT.PROC_CD not like '%10022%' and SSTRANSAT.PROC_CD not like '%10023%' and SSTRANSAT.PROC_CD not like '%10024%' and SSTRANSAT.PROC_CD not like '%10025%' and SSTRANSAT.PROC_CD not like '%10026%' and SSTRANSAT.PROC_CD not like '%10027%' and SSTRANSAT.PROC_CD not like '%10028%' and SSTRANSAT.PROC_CD not like '%10029%' and SSTRANSAT.PROC_CD not like '%10030%' and SSTRANSAT.PROC_CD not like '%10031%' and SSTRANSAT.PROC_CD not like '%10032%' and SSTRANSAT.PROC_CD not like '%10033%' and SSTRANSAT.PROC_CD not like '%10034%' and SSTRANSAT.PROC_CD not like '%10035%' and SSTRANSAT.PROC_CD not like '%10036%' and SSTRANSAT.PROC_CD not like '%10041%' and SSTRANSAT.PROC_CD not like '%10042%' GROUP by SSTRANSAT.ENCODE_DT,SSINFOTERMPROCCD.PROC_NM,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN")
                Dim rpt As New SICKMATALL
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf rbSMAT.Checked = True And rbAll.Checked = True Then
                If day1 = "1" And day2 = "31" Or day1 = "1" And day2 = "30" Then
                    GetNameBranch()
                    getMonth()
                    GetName = "SicknessMaternity"
                    'dt = db.ExecuteSQLQuery("SELECT convert(varchar,SSTRANSAT.ENCODE_DT,101) as DateOfTrans,case when SSINFOTERMPROCCD.PROC_NM = 'SSID CLEARANCE' then 'SSS/UMID CARD INFO' else SSINFOTERMPROCCD.PROC_NM end as 'PROC',SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN, COUNT(SSINFOTERMPROCCD.PROC_NM) as [count] FROM SSTRANSAT INNER JOIN SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAT.BRANCH_IP INNER JOIN SSINFOTERMPROCCD ON SSINFOTERMPROCCD.PROC_CD = SSTRANSAT.PROC_CD where SSTRANSAT.ENCODE_DT BETWEEN '" & date12 & "' and '" & date22 & "' and SSTRANSAT.PROC_CD not like '%10001%' and SSTRANSAT.PROC_CD not like '%10002%' and SSTRANSAT.PROC_CD not like '%10003%' and SSTRANSAT.PROC_CD not like '%10004%' and SSTRANSAT.PROC_CD not like '%10005%' and SSTRANSAT.PROC_CD not like '%10006%' and SSTRANSAT.PROC_CD not like '%10007%' and SSTRANSAT.PROC_CD not like '%10008%' and SSTRANSAT.PROC_CD not like '%10009%' and SSTRANSAT.PROC_CD not like '%10010%' and SSTRANSAT.PROC_CD not like '%10011%' and SSTRANSAT.PROC_CD not like '%10012%' and SSTRANSAT.PROC_CD not like '%10013%' and SSTRANSAT.PROC_CD not like '%10014%' and SSTRANSAT.PROC_CD not like '%10015%' and SSTRANSAT.PROC_CD not like '%10016%' and SSTRANSAT.PROC_CD not like '%10017%' and SSTRANSAT.PROC_CD not like '%10018%' and SSTRANSAT.PROC_CD not like '%10019%' and SSTRANSAT.PROC_CD not like '%10020%' and SSTRANSAT.PROC_CD not like '%10021%' and SSTRANSAT.PROC_CD not like '%10022%' and SSTRANSAT.PROC_CD not like '%10023%' and SSTRANSAT.PROC_CD not like '%10024%' and SSTRANSAT.PROC_CD not like '%10025%' and SSTRANSAT.PROC_CD not like '%10026%' and SSTRANSAT.PROC_CD not like '%10027%' and SSTRANSAT.PROC_CD not like '%10028%' and SSTRANSAT.PROC_CD not like '%10029%' and SSTRANSAT.PROC_CD not like '%10030%' and SSTRANSAT.PROC_CD not like '%10031%' and SSTRANSAT.PROC_CD not like '%10032%' and SSTRANSAT.PROC_CD not like '%10033%' and SSTRANSAT.PROC_CD not like '%10034%' and SSTRANSAT.PROC_CD not like '%10035%' and SSTRANSAT.PROC_CD not like '%10036%' and SSTRANSAT.PROC_CD not like '%10041%' and SSTRANSAT.PROC_CD not like '%10042%' GROUP by SSTRANSAT.ENCODE_DT,SSINFOTERMPROCCD.PROC_NM,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN")
                    Dim rpt As New SICKMATALL
                    cryRpt = rpt
                    cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                    rptView.ReportSource = cryRpt
                    cryRpt.SetParameterValue("@DateFrom", date12)
                    cryRpt.SetParameterValue("@DateTo", date22)
                    cryRpt.SetParameterValue("@getName", getBranch)
                Else
                    GetNameBranch()
                    GetName = "SicknessMaternity"
                    'dt = db.ExecuteSQLQuery("SELECT convert(varchar,SSTRANSAT.ENCODE_DT,101) as DateOfTrans,case when SSINFOTERMPROCCD.PROC_NM = 'SSID CLEARANCE' then 'SSS/UMID CARD INFO' else SSINFOTERMPROCCD.PROC_NM end as 'PROC',SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN, COUNT(SSINFOTERMPROCCD.PROC_NM) as [count] FROM SSTRANSAT INNER JOIN SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAT.BRANCH_IP INNER JOIN SSINFOTERMPROCCD ON SSINFOTERMPROCCD.PROC_CD = SSTRANSAT.PROC_CD where SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and SSTRANSAT.PROC_CD not like '%10001%' and SSTRANSAT.PROC_CD not like '%10002%' and SSTRANSAT.PROC_CD not like '%10003%' and SSTRANSAT.PROC_CD not like '%10004%' and SSTRANSAT.PROC_CD not like '%10005%' and SSTRANSAT.PROC_CD not like '%10006%' and SSTRANSAT.PROC_CD not like '%10007%' and SSTRANSAT.PROC_CD not like '%10008%' and SSTRANSAT.PROC_CD not like '%10009%' and SSTRANSAT.PROC_CD not like '%10010%' and SSTRANSAT.PROC_CD not like '%10011%' and SSTRANSAT.PROC_CD not like '%10012%' and SSTRANSAT.PROC_CD not like '%10013%' and SSTRANSAT.PROC_CD not like '%10014%' and SSTRANSAT.PROC_CD not like '%10015%' and SSTRANSAT.PROC_CD not like '%10016%' and SSTRANSAT.PROC_CD not like '%10017%' and SSTRANSAT.PROC_CD not like '%10018%' and SSTRANSAT.PROC_CD not like '%10019%' and SSTRANSAT.PROC_CD not like '%10020%' and SSTRANSAT.PROC_CD not like '%10021%' and SSTRANSAT.PROC_CD not like '%10022%' and SSTRANSAT.PROC_CD not like '%10023%' and SSTRANSAT.PROC_CD not like '%10024%' and SSTRANSAT.PROC_CD not like '%10025%' and SSTRANSAT.PROC_CD not like '%10026%' and SSTRANSAT.PROC_CD not like '%10027%' and SSTRANSAT.PROC_CD not like '%10028%' and SSTRANSAT.PROC_CD not like '%10029%' and SSTRANSAT.PROC_CD not like '%10030%' and SSTRANSAT.PROC_CD not like '%10031%' and SSTRANSAT.PROC_CD not like '%10032%' and SSTRANSAT.PROC_CD not like '%10033%' and SSTRANSAT.PROC_CD not like '%10034%' and SSTRANSAT.PROC_CD not like '%10035%' and SSTRANSAT.PROC_CD not like '%10036%' and SSTRANSAT.PROC_CD not like '%10041%' and SSTRANSAT.PROC_CD not like '%10042%' GROUP by SSTRANSAT.ENCODE_DT,SSINFOTERMPROCCD.PROC_NM,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN")
                    Dim rpt As New SICKMATALL
                    cryRpt = rpt
                    cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                    rptView.ReportSource = cryRpt
                    cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                    cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                    cryRpt.SetParameterValue("@getName", getBranch)
                End If


            ElseIf rbfb.Checked = True And rbBranch.Checked = True Then
                GetNameBranch()
                GetName = "WEBSITE FEEDBACK"
                'dt = db.ExecuteSQLQuery("select count(NAME)[count],SSINFOTERMKIOSK.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSINFOTERMFB.KIOSK_ID,SSTRANSINFOTERMFB.ENCODE_DT from SSTRANSINFOTERMFB INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID = SSTRANSINFOTERMFB.KIOSK_ID INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSINFOTERMFB.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSTRANSINFOTERMFB.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSTRANSINFOTERMFB.DIVSN where cast(SSTRANSINFOTERMFB.ENCODE_DT as date) BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' GROUP BY SSTRANSINFOTERMFB.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSINFOTERMFB.ENCODE_DT,SSTRANSINFOTERMFB.BRANCH_CD")
                Dim rpt As New FIGFB
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf rbfb.Checked = True And rbGroup.Checked = True Then
                GetNameBranch()
                GetName = "WEBSITE FEEDBACK"
                'dt = db.ExecuteSQLQuery("select count(NAME)[count],SSINFOTERMKIOSK.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSINFOTERMFB.KIOSK_ID,SSTRANSINFOTERMFB.ENCODE_DT from SSTRANSINFOTERMFB INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID = SSTRANSINFOTERMFB.KIOSK_ID INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSINFOTERMFB.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSTRANSINFOTERMFB.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSTRANSINFOTERMFB.DIVSN where cast(SSTRANSINFOTERMFB.ENCODE_DT as date) BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' GROUP BY SSTRANSINFOTERMFB.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSINFOTERMFB.ENCODE_DT,SSTRANSINFOTERMFB.BRANCH_CD")
                Dim rpt As New FIGFB
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf rbfb.Checked = True And rbCluster.Checked = True Then
                GetNameBranch()
                GetName = "WEBSITE FEEDBACK"
                'dt = db.ExecuteSQLQuery("select count(NAME)[count],SSINFOTERMKIOSK.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSINFOTERMFB.KIOSK_ID,SSTRANSINFOTERMFB.ENCODE_DT from SSTRANSINFOTERMFB INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID = SSTRANSINFOTERMFB.KIOSK_ID INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSINFOTERMFB.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSTRANSINFOTERMFB.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSTRANSINFOTERMFB.DIVSN where cast(SSTRANSINFOTERMFB.ENCODE_DT as date) BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' GROUP BY SSTRANSINFOTERMFB.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSINFOTERMFB.ENCODE_DT,SSTRANSINFOTERMFB.BRANCH_CD")
                Dim rpt As New FIGFB
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf rbfb.Checked = True And rbAll.Checked = True Then
                If day1 = "1" And day2 = "31" Or day1 = "1" And day2 = "30" Then
                    GetNameBranch()
                    getMonth()
                    GetName = "WEBSITE FEEDBACK"
                    'dt = db.ExecuteSQLQuery("select count(NAME)[count],SSINFOTERMKIOSK.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSINFOTERMFB.KIOSK_ID,SSTRANSINFOTERMFB.ENCODE_DT from SSTRANSINFOTERMFB INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID = SSTRANSINFOTERMFB.KIOSK_ID INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSINFOTERMFB.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSTRANSINFOTERMFB.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSTRANSINFOTERMFB.DIVSN where cast(SSTRANSINFOTERMFB.ENCODE_DT as date) BETWEEN '" & date12 & "' and '" & date22 & "' GROUP BY SSTRANSINFOTERMFB.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSINFOTERMFB.ENCODE_DT,SSTRANSINFOTERMFB.BRANCH_CD")
                    Dim rpt As New FIGFB
                    cryRpt = rpt
                    cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                    rptView.ReportSource = cryRpt
                    cryRpt.SetParameterValue("@DateFrom", date12)
                    cryRpt.SetParameterValue("@DateTo", date22)
                    cryRpt.SetParameterValue("@getName", getBranch)
                Else
                    GetNameBranch()
                    GetName = "WEBSITE FEEDBACK"
                    'dt = db.ExecuteSQLQuery("select count(NAME)[count],SSINFOTERMKIOSK.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSINFOTERMFB.KIOSK_ID,SSTRANSINFOTERMFB.ENCODE_DT from SSTRANSINFOTERMFB INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID = SSTRANSINFOTERMFB.KIOSK_ID INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSINFOTERMFB.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSTRANSINFOTERMFB.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSTRANSINFOTERMFB.DIVSN where cast(SSTRANSINFOTERMFB.ENCODE_DT as date) BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' GROUP BY SSTRANSINFOTERMFB.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSINFOTERMFB.ENCODE_DT,SSTRANSINFOTERMFB.BRANCH_CD")
                    Dim rpt As New FIGFB
                    cryRpt = rpt
                    cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                    rptView.ReportSource = cryRpt
                    cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                    cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                    cryRpt.SetParameterValue("@getName", getBranch)
                End If


            ElseIf fbk.Checked = True And rbGroup.Checked = True Then
                GetNameBranch()
                GetName = "SET FEEDBACK"
                'dt = db.ExecuteSQLQuery("select count(NAME)[count],SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSINFOTERMFBKIOSK.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP,SSTRANSINFOTERMFBKIOSK.ENCODE_DT from SSTRANSINFOTERMFBKIOSK INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID = SSTRANSINFOTERMFBKIOSK.KIOSK_ID INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSINFOTERMFBKIOSK.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSTRANSINFOTERMFBKIOSK.CLSTR INNER JOIN SSINFOTERMGROUP  ON SSINFOTERMGROUP.GROUP_CD = SSTRANSINFOTERMFBKIOSK.DIVSN where cast(SSTRANSINFOTERMFBKIOSK.ENCODE_DT as date) BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' GROUP BY SSINFOTERMKIOSK.BRANCH_IP,SSTRANSINFOTERMFBKIOSK.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSINFOTERMFBKIOSK.ENCODE_DT")
                Dim rpt As New FIGFBK
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)

            ElseIf fbk.Checked = True And rbBranch.Checked = True Then
                GetNameBranch()
                GetName = "SET FEEDBACK"
                'dt = db.ExecuteSQLQuery("select count(NAME)[count],SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSINFOTERMFBKIOSK.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP,SSTRANSINFOTERMFBKIOSK.ENCODE_DT from SSTRANSINFOTERMFBKIOSK INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID = SSTRANSINFOTERMFBKIOSK.KIOSK_ID INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSINFOTERMFBKIOSK.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSTRANSINFOTERMFBKIOSK.CLSTR INNER JOIN SSINFOTERMGROUP  ON SSINFOTERMGROUP.GROUP_CD = SSTRANSINFOTERMFBKIOSK.DIVSN where cast(SSTRANSINFOTERMFBKIOSK.ENCODE_DT as date) BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' GROUP BY SSINFOTERMKIOSK.BRANCH_IP,SSTRANSINFOTERMFBKIOSK.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSINFOTERMFBKIOSK.ENCODE_DT")
                Dim rpt As New FIGFBK
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf fbk.Checked = True And rbCluster.Checked = True Then
                GetNameBranch()
                GetName = "SET FEEDBACK"
                'dt = db.ExecuteSQLQuery("select count(NAME)[count],SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSINFOTERMFBKIOSK.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP,SSTRANSINFOTERMFBKIOSK.ENCODE_DT from SSTRANSINFOTERMFBKIOSK INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID = SSTRANSINFOTERMFBKIOSK.KIOSK_ID INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSINFOTERMFBKIOSK.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSTRANSINFOTERMFBKIOSK.CLSTR INNER JOIN SSINFOTERMGROUP  ON SSINFOTERMGROUP.GROUP_CD = SSTRANSINFOTERMFBKIOSK.DIVSN where cast(SSTRANSINFOTERMFBKIOSK.ENCODE_DT as date) BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' GROUP BY SSINFOTERMKIOSK.BRANCH_IP,SSTRANSINFOTERMFBKIOSK.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSINFOTERMFBKIOSK.ENCODE_DT")
                Dim rpt As New FIGFBK
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf fbk.Checked = True And rbAll.Checked = True Then
                getMonth()
                If day1 = "1" And day2 = "31" Or day1 = "1" And day2 = "30" Then
                    GetNameBranch()
                    GetName = "SET FEEDBACK"
                    'dt = db.ExecuteSQLQuery("select count(NAME)[count],SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSINFOTERMFBKIOSK.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP,SSTRANSINFOTERMFBKIOSK.ENCODE_DT from SSTRANSINFOTERMFBKIOSK INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID = SSTRANSINFOTERMFBKIOSK.KIOSK_ID INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSINFOTERMFBKIOSK.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSTRANSINFOTERMFBKIOSK.CLSTR INNER JOIN SSINFOTERMGROUP  ON SSINFOTERMGROUP.GROUP_CD = SSTRANSINFOTERMFBKIOSK.DIVSN where cast(SSTRANSINFOTERMFBKIOSK.ENCODE_DT as date) BETWEEN '" & date12 & "' and '" & date22 & "' GROUP BY SSINFOTERMKIOSK.BRANCH_IP,SSTRANSINFOTERMFBKIOSK.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSINFOTERMFBKIOSK.ENCODE_DT")
                    Dim rpt As New FIGFBK
                    cryRpt = rpt
                    cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                    rptView.ReportSource = cryRpt
                    cryRpt.SetParameterValue("@DateFrom", date12)
                    cryRpt.SetParameterValue("@DateTo", date22)
                    cryRpt.SetParameterValue("@getName", getBranch)
                Else
                    GetNameBranch()
                    GetName = "SET FEEDBACK"
                    'dt = db.ExecuteSQLQuery("select count(NAME)[count],SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSINFOTERMFBKIOSK.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP,SSTRANSINFOTERMFBKIOSK.ENCODE_DT from SSTRANSINFOTERMFBKIOSK INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID = SSTRANSINFOTERMFBKIOSK.KIOSK_ID INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSINFOTERMFBKIOSK.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSTRANSINFOTERMFBKIOSK.CLSTR INNER JOIN SSINFOTERMGROUP  ON SSINFOTERMGROUP.GROUP_CD = SSTRANSINFOTERMFBKIOSK.DIVSN where cast(SSTRANSINFOTERMFBKIOSK.ENCODE_DT as date) BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' GROUP BY SSINFOTERMKIOSK.BRANCH_IP,SSTRANSINFOTERMFBKIOSK.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSINFOTERMFBKIOSK.ENCODE_DT")
                    Dim rpt As New FIGFBK
                    cryRpt = rpt
                    cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                    rptView.ReportSource = cryRpt
                    cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                    cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                    cryRpt.SetParameterValue("@getName", getBranch)
                End If
            ElseIf rbSWR.Checked = True And rbGroup.Checked = True Then
                GetNameBranch()
                GetName = "SIMPLIFIED WEB REGISTRATION"
                'dt = db.ExecuteSQLQuery("select case when memstatus = 'COVERED EMPLOYEE' then 'CE' when memstatus = 'SELF EMPLOYED' then 'SE' when memstatus = 'VOLUNTARY' then 'VM' when memstatus = 'RETIRMENT PENSIONER' then 'RP' else memstatus end as [status], COUNT(memstatus)[count],SSTRANSINFOTERMRG.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSINFOTERMRG.ENCODE_DT FROM SSTRANSINFOTERMRG INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSINFOTERMRG.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSTRANSINFOTERMRG.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSTRANSINFOTERMRG.DIVSN INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID = SSTRANSINFOTERMRG.KIOSK_ID WHERE SSTRANSINFOTERMRG.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' GROUP BY memstatus,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSINFOTERMRG.ENCODE_DT,SSTRANSINFOTERMRG.S_TRANID,SSINFOTERMKIOSK.BRANCH_IP,SSTRANSINFOTERMRG.KIOSK_ID")
                Dim rpt As New FIGREGALL
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)

            ElseIf rbSWR.Checked = True And rbCluster.Checked = True Then
                GetNameBranch()
                GetName = "SIMPLIFIED WEB REGISTRATION"
                'dt = db.ExecuteSQLQuery("select case when memstatus = 'COVERED EMPLOYEE' then 'CE' when memstatus = 'SELF EMPLOYED' then 'SE' when memstatus = 'VOLUNTARY' then 'VM' when memstatus = 'RETIRMENT PENSIONER' then 'RP' else memstatus end as [status], COUNT(memstatus)[count],SSTRANSINFOTERMRG.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSINFOTERMRG.ENCODE_DT FROM SSTRANSINFOTERMRG INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSINFOTERMRG.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSTRANSINFOTERMRG.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSTRANSINFOTERMRG.DIVSN INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID = SSTRANSINFOTERMRG.KIOSK_ID WHERE SSTRANSINFOTERMRG.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' GROUP BY memstatus,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSINFOTERMRG.ENCODE_DT,SSTRANSINFOTERMRG.S_TRANID,SSINFOTERMKIOSK.BRANCH_IP,SSTRANSINFOTERMRG.KIOSK_ID")
                Dim rpt As New FIGREGALL
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf rbSWR.Checked = True And rbBranch.Checked = True Then
                GetNameBranch()
                GetName = "SIMPLIFIED WEB REGISTRATION"
                'dt = db.ExecuteSQLQuery("select case when memstatus = 'COVERED EMPLOYEE' then 'CE' when memstatus = 'SELF EMPLOYED' then 'SE' when memstatus = 'VOLUNTARY' then 'VM' when memstatus = 'RETIRMENT PENSIONER' then 'RP' else memstatus end as [status], COUNT(memstatus)[count],SSTRANSINFOTERMRG.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSINFOTERMRG.ENCODE_DT FROM SSTRANSINFOTERMRG INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSINFOTERMRG.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSTRANSINFOTERMRG.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSTRANSINFOTERMRG.DIVSN INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID = SSTRANSINFOTERMRG.KIOSK_ID WHERE SSTRANSINFOTERMRG.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' GROUP BY memstatus,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSINFOTERMRG.ENCODE_DT,SSTRANSINFOTERMRG.S_TRANID,SSINFOTERMKIOSK.BRANCH_IP,SSTRANSINFOTERMRG.KIOSK_ID")
                Dim rpt As New FIGREGALL
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf rbSWR.Checked = True And rbAll.Checked = True Then
                getMonth()
                If day1 = "1" And day2 = "31" Or day1 = "1" And day2 = "30" Then
                    GetNameBranch()
                    GetName = "SIMPLIFIED WEB REGISTRATION"
                    'dt = db.ExecuteSQLQuery("select case when memstatus = 'COVERED EMPLOYEE' then 'CE' when memstatus = 'SELF EMPLOYED' then 'SE' when memstatus = 'VOLUNTARY' then 'VM' when memstatus = 'RETIRMENT PENSIONER' then 'RP' else memstatus end as [status], COUNT(memstatus)[count],SSTRANSINFOTERMRG.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSINFOTERMRG.ENCODE_DT FROM SSTRANSINFOTERMRG INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSINFOTERMRG.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSTRANSINFOTERMRG.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSTRANSINFOTERMRG.DIVSN INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID = SSTRANSINFOTERMRG.KIOSK_ID WHERE SSTRANSINFOTERMRG.ENCODE_DT BETWEEN '" & date12 & "' and '" & date22 & "' GROUP BY memstatus,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSINFOTERMRG.ENCODE_DT,SSTRANSINFOTERMRG.S_TRANID,SSINFOTERMKIOSK.BRANCH_IP,SSTRANSINFOTERMRG.KIOSK_ID")
                    Dim rpt As New FIGREGALL
                    cryRpt = rpt
                    cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                    rptView.ReportSource = cryRpt
                    cryRpt.SetParameterValue("@DateFrom", date12)
                    cryRpt.SetParameterValue("@DateTo", date22)
                    cryRpt.SetParameterValue("@getName", getBranch)
                Else
                    GetNameBranch()
                    GetName = "SIMPLIFIED WEB REGISTRATION"
                    'dt = db.ExecuteSQLQuery("select case when memstatus = 'COVERED EMPLOYEE' then 'CE' when memstatus = 'SELF EMPLOYED' then 'SE' when memstatus = 'VOLUNTARY' then 'VM' when memstatus = 'RETIRMENT PENSIONER' then 'RP' else memstatus end as [status], COUNT(memstatus)[count],SSTRANSINFOTERMRG.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSINFOTERMRG.ENCODE_DT FROM SSTRANSINFOTERMRG INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSINFOTERMRG.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSTRANSINFOTERMRG.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSTRANSINFOTERMRG.DIVSN INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID = SSTRANSINFOTERMRG.KIOSK_ID WHERE SSTRANSINFOTERMRG.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' GROUP BY memstatus,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSINFOTERMRG.ENCODE_DT,SSTRANSINFOTERMRG.S_TRANID,SSINFOTERMKIOSK.BRANCH_IP,SSTRANSINFOTERMRG.KIOSK_ID")
                    Dim rpt As New FIGREGALL
                    cryRpt = rpt
                    cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                    rptView.ReportSource = cryRpt
                    cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                    cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                    cryRpt.SetParameterValue("@getName", getBranch)
                End If

            ElseIf rbSuc.Checked = True And rbGroup.Checked = True Then
                GetNameBranch()
                GetName = "SUCCESS TRANSACTIONS"
                'dt = db.ExecuteSQLQuery("SELECT DISTINCT SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM, (SELECT COUNT(ID) from SSTRANSAPPSL where SSTRANSAPPSL.IN_IPADD = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAPPSL.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') +  (SELECT COUNT(INT) from SSTRANSAPPSLEMP where SSTRANSAPPSLEMP.IN_IPADD = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAPPSLEMP.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') as [SL], (SELECT COUNT(ID) FROM SSTRANSACOP where SSTRANSACOP.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSACOP.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') as [ACP], (SELECT COUNT(ID) FROM SSTRANSINFOTERMTR where SSTRANSINFOTERMTR.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSINFOTERMTR.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') +  (SELECT COUNT(ID) FROM SSTRANSINFOTERMTRLS where SSTRANSINFOTERMTRLS.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSINFOTERMTRLS.DATECREATED BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') as [TR], (SELECT COUNT(SSTRANSINFOTERMRG.ID) FROM SSTRANSINFOTERMRG where SSTRANSINFOTERMRG.KIOSK_ID = SSINFOTERMKIOSK.KIOSK_ID and SSTRANSINFOTERMRG.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') as [SWR], (SELECT COUNT(SSTRANSINFORTERMMN.SEQ_NO) FROM SSTRANSINFORTERMMN where SSTRANSINFORTERMMN.KIOSK_ID = SSINFOTERMKIOSK.KIOSK_ID and SSTRANSINFORTERMMN.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') as [MN] FROM SSINFOTERMKIOSK LEFT JOIN SSTRANSACOP ON SSTRANSACOP.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP LEFT JOIN SSTRANSAPPSL ON SSTRANSAPPSL.IN_IPADD = SSINFOTERMKIOSK.BRANCH_IP LEFT JOIN SSTRANSAPPSLEMP ON SSTRANSAPPSLEMP.IN_IPADD = SSINFOTERMKIOSK.BRANCH_IP LEFT JOIN SSTRANSINFORTERMMN ON SSTRANSINFORTERMMN.KIOSK_ID = SSINFOTERMKIOSK.KIOSK_ID LEFT JOIN SSTRANSINFOTERMRG ON SSTRANSINFOTERMRG.KIOSK_ID = SSINFOTERMKIOSK.KIOSK_ID LEFT JOIN SSTRANSINFOTERMTR ON SSTRANSINFOTERMTR.BRANCH_IP = SSINFOTERMKIOSK.KIOSK_ID LEFT JOIN SSTRANSINFOTERMTRLS ON SSTRANSINFOTERMTRLS.BRANCH_IP = SSINFOTERMKIOSK.KIOSK_ID INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSINFOTERMKIOSK.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMKIOSK.CLSTR = SSINFOTERMCLSTR.CLSTR_CD INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSINFOTERMKIOSK.DIVSN ORDER BY SSINFOTERMKIOSK.KIOSK_ID")
                Dim rpt As New SUCCESSTRANSALL
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf rbSuc.Checked = True And rbBranch.Checked = True Then
                GetNameBranch()
                GetName = "SUCCESS TRANSACTIONS"
                'dt = db.ExecuteSQLQuery("SELECT DISTINCT SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM, (SELECT COUNT(ID) from SSTRANSAPPSL where SSTRANSAPPSL.IN_IPADD = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAPPSL.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') +  (SELECT COUNT(INT) from SSTRANSAPPSLEMP where SSTRANSAPPSLEMP.IN_IPADD = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAPPSLEMP.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') as [SL], (SELECT COUNT(ID) FROM SSTRANSACOP where SSTRANSACOP.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSACOP.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') as [ACP], (SELECT COUNT(ID) FROM SSTRANSINFOTERMTR where SSTRANSINFOTERMTR.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSINFOTERMTR.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') +  (SELECT COUNT(ID) FROM SSTRANSINFOTERMTRLS where SSTRANSINFOTERMTRLS.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSINFOTERMTRLS.DATECREATED BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') as [TR], (SELECT COUNT(SSTRANSINFOTERMRG.ID) FROM SSTRANSINFOTERMRG where SSTRANSINFOTERMRG.KIOSK_ID = SSINFOTERMKIOSK.KIOSK_ID and SSTRANSINFOTERMRG.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') as [SWR], (SELECT COUNT(SSTRANSINFORTERMMN.SEQ_NO) FROM SSTRANSINFORTERMMN where SSTRANSINFORTERMMN.KIOSK_ID = SSINFOTERMKIOSK.KIOSK_ID and SSTRANSINFORTERMMN.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') as [MN] FROM SSINFOTERMKIOSK LEFT JOIN SSTRANSACOP ON SSTRANSACOP.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP LEFT JOIN SSTRANSAPPSL ON SSTRANSAPPSL.IN_IPADD = SSINFOTERMKIOSK.BRANCH_IP LEFT JOIN SSTRANSAPPSLEMP ON SSTRANSAPPSLEMP.IN_IPADD = SSINFOTERMKIOSK.BRANCH_IP LEFT JOIN SSTRANSINFORTERMMN ON SSTRANSINFORTERMMN.KIOSK_ID = SSINFOTERMKIOSK.KIOSK_ID LEFT JOIN SSTRANSINFOTERMRG ON SSTRANSINFOTERMRG.KIOSK_ID = SSINFOTERMKIOSK.KIOSK_ID LEFT JOIN SSTRANSINFOTERMTR ON SSTRANSINFOTERMTR.BRANCH_IP = SSINFOTERMKIOSK.KIOSK_ID LEFT JOIN SSTRANSINFOTERMTRLS ON SSTRANSINFOTERMTRLS.BRANCH_IP = SSINFOTERMKIOSK.KIOSK_ID INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSINFOTERMKIOSK.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMKIOSK.CLSTR = SSINFOTERMCLSTR.CLSTR_CD INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSINFOTERMKIOSK.DIVSN ORDER BY SSINFOTERMKIOSK.KIOSK_ID")
                Dim rpt As New SUCCESSTRANSALL
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf rbSuc.Checked = True And rbCluster.Checked = True Then
                GetNameBranch()
                GetName = "SUCCESS TRANSACTIONS"
                'dt = db.ExecuteSQLQuery("SELECT DISTINCT SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM, (SELECT COUNT(ID) from SSTRANSAPPSL where SSTRANSAPPSL.IN_IPADD = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAPPSL.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') +  (SELECT COUNT(INT) from SSTRANSAPPSLEMP where SSTRANSAPPSLEMP.IN_IPADD = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAPPSLEMP.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') as [SL], (SELECT COUNT(ID) FROM SSTRANSACOP where SSTRANSACOP.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSACOP.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') as [ACP], (SELECT COUNT(ID) FROM SSTRANSINFOTERMTR where SSTRANSINFOTERMTR.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSINFOTERMTR.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') +  (SELECT COUNT(ID) FROM SSTRANSINFOTERMTRLS where SSTRANSINFOTERMTRLS.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSINFOTERMTRLS.DATECREATED BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') as [TR], (SELECT COUNT(SSTRANSINFOTERMRG.ID) FROM SSTRANSINFOTERMRG where SSTRANSINFOTERMRG.KIOSK_ID = SSINFOTERMKIOSK.KIOSK_ID and SSTRANSINFOTERMRG.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') as [SWR], (SELECT COUNT(SSTRANSINFORTERMMN.SEQ_NO) FROM SSTRANSINFORTERMMN where SSTRANSINFORTERMMN.KIOSK_ID = SSINFOTERMKIOSK.KIOSK_ID and SSTRANSINFORTERMMN.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') as [MN] FROM SSINFOTERMKIOSK LEFT JOIN SSTRANSACOP ON SSTRANSACOP.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP LEFT JOIN SSTRANSAPPSL ON SSTRANSAPPSL.IN_IPADD = SSINFOTERMKIOSK.BRANCH_IP LEFT JOIN SSTRANSAPPSLEMP ON SSTRANSAPPSLEMP.IN_IPADD = SSINFOTERMKIOSK.BRANCH_IP LEFT JOIN SSTRANSINFORTERMMN ON SSTRANSINFORTERMMN.KIOSK_ID = SSINFOTERMKIOSK.KIOSK_ID LEFT JOIN SSTRANSINFOTERMRG ON SSTRANSINFOTERMRG.KIOSK_ID = SSINFOTERMKIOSK.KIOSK_ID LEFT JOIN SSTRANSINFOTERMTR ON SSTRANSINFOTERMTR.BRANCH_IP = SSINFOTERMKIOSK.KIOSK_ID LEFT JOIN SSTRANSINFOTERMTRLS ON SSTRANSINFOTERMTRLS.BRANCH_IP = SSINFOTERMKIOSK.KIOSK_ID INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSINFOTERMKIOSK.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMKIOSK.CLSTR = SSINFOTERMCLSTR.CLSTR_CD INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSINFOTERMKIOSK.DIVSN ORDER BY SSINFOTERMKIOSK.KIOSK_ID")
                Dim rpt As New SUCCESSTRANSALL
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf rbSuc.Checked = True And rbAll.Checked = True Then
                getMonth()
                If day1 = "1" And day2 = "31" Or day1 = "1" And day2 = "30" Then
                    GetNameBranch()
                    GetName = "SUCCESS TRANSACTIONS"
                    'dt = db.ExecuteSQLQuery("SELECT DISTINCT SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM, (SELECT COUNT(ID) from SSTRANSAPPSL where SSTRANSAPPSL.IN_IPADD = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAPPSL.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') +  (SELECT COUNT(INT) from SSTRANSAPPSLEMP where SSTRANSAPPSLEMP.IN_IPADD = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAPPSLEMP.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') as [SL], (SELECT COUNT(ID) FROM SSTRANSACOP where SSTRANSACOP.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSACOP.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') as [ACP], (SELECT COUNT(ID) FROM SSTRANSINFOTERMTR where SSTRANSINFOTERMTR.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSINFOTERMTR.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') +  (SELECT COUNT(ID) FROM SSTRANSINFOTERMTRLS where SSTRANSINFOTERMTRLS.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSINFOTERMTRLS.DATECREATED BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') as [TR], (SELECT COUNT(SSTRANSINFOTERMRG.ID) FROM SSTRANSINFOTERMRG where SSTRANSINFOTERMRG.KIOSK_ID = SSINFOTERMKIOSK.KIOSK_ID and SSTRANSINFOTERMRG.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') as [SWR], (SELECT COUNT(SSTRANSINFORTERMMN.SEQ_NO) FROM SSTRANSINFORTERMMN where SSTRANSINFORTERMMN.KIOSK_ID = SSINFOTERMKIOSK.KIOSK_ID and SSTRANSINFORTERMMN.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') as [MN] FROM SSINFOTERMKIOSK LEFT JOIN SSTRANSACOP ON SSTRANSACOP.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP LEFT JOIN SSTRANSAPPSL ON SSTRANSAPPSL.IN_IPADD = SSINFOTERMKIOSK.BRANCH_IP LEFT JOIN SSTRANSAPPSLEMP ON SSTRANSAPPSLEMP.IN_IPADD = SSINFOTERMKIOSK.BRANCH_IP LEFT JOIN SSTRANSINFORTERMMN ON SSTRANSINFORTERMMN.KIOSK_ID = SSINFOTERMKIOSK.KIOSK_ID LEFT JOIN SSTRANSINFOTERMRG ON SSTRANSINFOTERMRG.KIOSK_ID = SSINFOTERMKIOSK.KIOSK_ID LEFT JOIN SSTRANSINFOTERMTR ON SSTRANSINFOTERMTR.BRANCH_IP = SSINFOTERMKIOSK.KIOSK_ID LEFT JOIN SSTRANSINFOTERMTRLS ON SSTRANSINFOTERMTRLS.BRANCH_IP = SSINFOTERMKIOSK.KIOSK_ID INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSINFOTERMKIOSK.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMKIOSK.CLSTR = SSINFOTERMCLSTR.CLSTR_CD INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSINFOTERMKIOSK.DIVSN ORDER BY SSINFOTERMKIOSK.KIOSK_ID")
                    Dim rpt As New SUCCESSTRANSALL
                    cryRpt = rpt
                    cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                    rptView.ReportSource = cryRpt
                    cryRpt.SetParameterValue("@DateFrom", date12)
                    cryRpt.SetParameterValue("@DateTo", date22)
                    cryRpt.SetParameterValue("@getName", getBranch)
                Else
                    GetNameBranch()
                    GetName = "SUCCESS TRANSACTIONS"
                    'dt = db.ExecuteSQLQuery("SELECT DISTINCT SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM, (SELECT COUNT(ID) from SSTRANSAPPSL where SSTRANSAPPSL.IN_IPADD = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAPPSL.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') +  (SELECT COUNT(INT) from SSTRANSAPPSLEMP where SSTRANSAPPSLEMP.IN_IPADD = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAPPSLEMP.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') as [SL], (SELECT COUNT(ID) FROM SSTRANSACOP where SSTRANSACOP.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSACOP.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') as [ACP], (SELECT COUNT(ID) FROM SSTRANSINFOTERMTR where SSTRANSINFOTERMTR.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSINFOTERMTR.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') +  (SELECT COUNT(ID) FROM SSTRANSINFOTERMTRLS where SSTRANSINFOTERMTRLS.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSINFOTERMTRLS.DATECREATED BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') as [TR], (SELECT COUNT(SSTRANSINFOTERMRG.ID) FROM SSTRANSINFOTERMRG where SSTRANSINFOTERMRG.KIOSK_ID = SSINFOTERMKIOSK.KIOSK_ID and SSTRANSINFOTERMRG.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') as [SWR], (SELECT COUNT(SSTRANSINFORTERMMN.SEQ_NO) FROM SSTRANSINFORTERMMN where SSTRANSINFORTERMMN.KIOSK_ID = SSINFOTERMKIOSK.KIOSK_ID and SSTRANSINFORTERMMN.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') as [MN] FROM SSINFOTERMKIOSK LEFT JOIN SSTRANSACOP ON SSTRANSACOP.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP LEFT JOIN SSTRANSAPPSL ON SSTRANSAPPSL.IN_IPADD = SSINFOTERMKIOSK.BRANCH_IP LEFT JOIN SSTRANSAPPSLEMP ON SSTRANSAPPSLEMP.IN_IPADD = SSINFOTERMKIOSK.BRANCH_IP LEFT JOIN SSTRANSINFORTERMMN ON SSTRANSINFORTERMMN.KIOSK_ID = SSINFOTERMKIOSK.KIOSK_ID LEFT JOIN SSTRANSINFOTERMRG ON SSTRANSINFOTERMRG.KIOSK_ID = SSINFOTERMKIOSK.KIOSK_ID LEFT JOIN SSTRANSINFOTERMTR ON SSTRANSINFOTERMTR.BRANCH_IP = SSINFOTERMKIOSK.KIOSK_ID LEFT JOIN SSTRANSINFOTERMTRLS ON SSTRANSINFOTERMTRLS.BRANCH_IP = SSINFOTERMKIOSK.KIOSK_ID INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSINFOTERMKIOSK.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMKIOSK.CLSTR = SSINFOTERMCLSTR.CLSTR_CD INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSINFOTERMKIOSK.DIVSN ORDER BY SSINFOTERMKIOSK.KIOSK_ID")
                    Dim rpt As New SUCCESSTRANSALL
                    cryRpt = rpt
                    cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                    rptView.ReportSource = cryRpt
                    cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                    cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                    cryRpt.SetParameterValue("@getName", getBranch)
                End If

            ElseIf rbFT.Checked = True And rbGroup.Checked = True Then
                GetNameBranch()
                GetName = "FAILED TRANSACTIONS"
                'dt = db.ExecuteSQLQuery("select SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,case when SSINFOTERMPROCCD.PROC_NM = 'MEMBER MATERNITY NOTIFICATION' then 'MN' when SSINFOTERMPROCCD.PROC_NM = 'SIMPLIFIED WEB REGISTRATION' then 'SWR' when SSINFOTERMPROCCD.PROC_NM = 'ANNUAL CONFIRMATION OF PENSIONER' then 'ACP' when SSINFOTERMPROCCD.PROC_NM = 'SALARY LOAN' then 'SL'when SSINFOTERMPROCCD.PROC_NM = 'TECHNICAL RETIREMENT' then 'TR' else SSINFOTERMPROCCD.PROC_NM end as [PROC],SSTRANSAT.ENCODE_DT, COUNT(SSINFOTERMPROCCD.PROC_NM) as [count],SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN from SSTRANSAT INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAT.BRANCH_IP INNER JOIN SSINFOTERMPROCCD ON SSINFOTERMPROCCD.PROC_CD = SSTRANSAT.PROC_CD WHERE TRANSNUM <> '' and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' GROUP BY  SSINFOTERMPROCCD.PROC_NM,SSTRANSAT.ENCODE_DT,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN,SSTRANSAT.BRANCH_IP")
                Dim rpt As New FIGFAILALL
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf rbFT.Checked = True And rbBranch.Checked = True Then
                GetNameBranch()
                GetName = "FAILED TRANSACTIONS"
                'dt = db.ExecuteSQLQuery("select SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,case when SSINFOTERMPROCCD.PROC_NM = 'MEMBER MATERNITY NOTIFICATION' then 'MN' when SSINFOTERMPROCCD.PROC_NM = 'SIMPLIFIED WEB REGISTRATION' then 'SWR' when SSINFOTERMPROCCD.PROC_NM = 'ANNUAL CONFIRMATION OF PENSIONER' then 'ACP' when SSINFOTERMPROCCD.PROC_NM = 'SALARY LOAN' then 'SL'when SSINFOTERMPROCCD.PROC_NM = 'TECHNICAL RETIREMENT' then 'TR' else SSINFOTERMPROCCD.PROC_NM end as [PROC],SSTRANSAT.ENCODE_DT, COUNT(SSINFOTERMPROCCD.PROC_NM) as [count],SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN from SSTRANSAT INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAT.BRANCH_IP INNER JOIN SSINFOTERMPROCCD ON SSINFOTERMPROCCD.PROC_CD = SSTRANSAT.PROC_CD WHERE TRANSNUM <> '' and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' GROUP BY  SSINFOTERMPROCCD.PROC_NM,SSTRANSAT.ENCODE_DT,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN,SSTRANSAT.BRANCH_IP")
                Dim rpt As New FIGFAILALL
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf rbFT.Checked = True And rbCluster.Checked = True Then
                GetNameBranch()
                GetName = "FAILED TRANSACTIONS"
                'dt = db.ExecuteSQLQuery("select SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,case when SSINFOTERMPROCCD.PROC_NM = 'MEMBER MATERNITY NOTIFICATION' then 'MN' when SSINFOTERMPROCCD.PROC_NM = 'SIMPLIFIED WEB REGISTRATION' then 'SWR' when SSINFOTERMPROCCD.PROC_NM = 'ANNUAL CONFIRMATION OF PENSIONER' then 'ACP' when SSINFOTERMPROCCD.PROC_NM = 'SALARY LOAN' then 'SL'when SSINFOTERMPROCCD.PROC_NM = 'TECHNICAL RETIREMENT' then 'TR' else SSINFOTERMPROCCD.PROC_NM end as [PROC],SSTRANSAT.ENCODE_DT, COUNT(SSINFOTERMPROCCD.PROC_NM) as [count],SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN from SSTRANSAT INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAT.BRANCH_IP INNER JOIN SSINFOTERMPROCCD ON SSINFOTERMPROCCD.PROC_CD = SSTRANSAT.PROC_CD WHERE TRANSNUM <> '' and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' GROUP BY  SSINFOTERMPROCCD.PROC_NM,SSTRANSAT.ENCODE_DT,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN,SSTRANSAT.BRANCH_IP")
                Dim rpt As New FIGFAILALL
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf rbFT.Checked = True And rbAll.Checked = True Then
                getMonth()
                If day1 = "1" And day2 = "31" Or day1 = "1" And day2 = "30" Then
                    GetNameBranch()
                    GetName = "FAILED TRANSACTIONS"
                    'dt = db.ExecuteSQLQuery("select SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,case when SSINFOTERMPROCCD.PROC_NM = 'MEMBER MATERNITY NOTIFICATION' then 'MN' when SSINFOTERMPROCCD.PROC_NM = 'SIMPLIFIED WEB REGISTRATION' then 'SWR' when SSINFOTERMPROCCD.PROC_NM = 'ANNUAL CONFIRMATION OF PENSIONER' then 'ACP' when SSINFOTERMPROCCD.PROC_NM = 'SALARY LOAN' then 'SL'when SSINFOTERMPROCCD.PROC_NM = 'TECHNICAL RETIREMENT' then 'TR' else SSINFOTERMPROCCD.PROC_NM end as [PROC],SSTRANSAT.ENCODE_DT, COUNT(SSINFOTERMPROCCD.PROC_NM) as [count],SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN from SSTRANSAT INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAT.BRANCH_IP INNER JOIN SSINFOTERMPROCCD ON SSINFOTERMPROCCD.PROC_CD = SSTRANSAT.PROC_CD WHERE TRANSNUM <> '' and SSTRANSAT.ENCODE_DT BETWEEN '" & date12 & "' and '" & date22 & "' GROUP BY  SSINFOTERMPROCCD.PROC_NM,SSTRANSAT.ENCODE_DT,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN,SSTRANSAT.BRANCH_IP")
                    Dim rpt As New FIGFAILALL
                    cryRpt = rpt
                    cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                    rptView.ReportSource = cryRpt
                    cryRpt.SetParameterValue("@DateFrom", date12)
                    cryRpt.SetParameterValue("@DateTo", date22)
                    cryRpt.SetParameterValue("@getName", getBranch)
                Else
                    GetNameBranch()
                    GetName = "FAILED TRANSACTIONS"
                    'dt = db.ExecuteSQLQuery("select SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,case when SSINFOTERMPROCCD.PROC_NM = 'MEMBER MATERNITY NOTIFICATION' then 'MN' when SSINFOTERMPROCCD.PROC_NM = 'SIMPLIFIED WEB REGISTRATION' then 'SWR' when SSINFOTERMPROCCD.PROC_NM = 'ANNUAL CONFIRMATION OF PENSIONER' then 'ACP' when SSINFOTERMPROCCD.PROC_NM = 'SALARY LOAN' then 'SL'when SSINFOTERMPROCCD.PROC_NM = 'TECHNICAL RETIREMENT' then 'TR' else SSINFOTERMPROCCD.PROC_NM end as [PROC],SSTRANSAT.ENCODE_DT, COUNT(SSINFOTERMPROCCD.PROC_NM) as [count],SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN from SSTRANSAT INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAT.BRANCH_IP INNER JOIN SSINFOTERMPROCCD ON SSINFOTERMPROCCD.PROC_CD = SSTRANSAT.PROC_CD WHERE TRANSNUM <> '' and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' GROUP BY  SSINFOTERMPROCCD.PROC_NM,SSTRANSAT.ENCODE_DT,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN,SSTRANSAT.BRANCH_IP")
                    Dim rpt As New FIGFAILALL
                    cryRpt = rpt
                    cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                    rptView.ReportSource = cryRpt
                    cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                    cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                    cryRpt.SetParameterValue("@getName", getBranch)
                End If
            ElseIf rbacop.Checked = True And rbGroup.Checked = True Then
                GetNameBranch()
                getMonth()
                GetName = "COUNT OF ACOP"
                'dt = db.ExecuteSQLQuery("select SSINFOTERMKIOSK.KIOSK_ID,SSTRANSACOP.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSACOP.ENCODE_DT,COUNT(TRANID)as [COUNT] from SSTRANSACOP INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.BRANCH_IP = SSTRANSACOP.BRANCH_IP INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSACOP.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSTRANSACOP.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSTRANSACOP.DIVSN where SSTRANSACOP.ENCODE_DT BETWEEN '" & date12 & "' and '" & date22 & "' GROUP BY SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSACOP.ENCODE_DT,SSTRANSACOP.BRANCH_IP")
                Dim rpt As New FIGACOPALL
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date12)
                cryRpt.SetParameterValue("@DateTo", date22)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf rbacop.Checked = True And rbBranch.Checked = True Then
                GetNameBranch()
                getMonth()
                GetName = "COUNT OF ACOP"
                'dt = db.ExecuteSQLQuery("select SSINFOTERMKIOSK.KIOSK_ID,SSTRANSACOP.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSACOP.ENCODE_DT,COUNT(TRANID)as [COUNT] from SSTRANSACOP INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.BRANCH_IP = SSTRANSACOP.BRANCH_IP INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSACOP.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSTRANSACOP.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSTRANSACOP.DIVSN where SSTRANSACOP.ENCODE_DT BETWEEN '" & date12 & "' and '" & date22 & "' GROUP BY SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSACOP.ENCODE_DT,SSTRANSACOP.BRANCH_IP")
                Dim rpt As New FIGACOPALL
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date12)
                cryRpt.SetParameterValue("@DateTo", date22)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf rbacop.Checked = True And rbCluster.Checked = True Then
                GetNameBranch()
                getMonth()
                GetName = "COUNT OF ACOP"
                'dt = db.ExecuteSQLQuery("select SSINFOTERMKIOSK.KIOSK_ID,SSTRANSACOP.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSACOP.ENCODE_DT,COUNT(TRANID)as [COUNT] from SSTRANSACOP INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.BRANCH_IP = SSTRANSACOP.BRANCH_IP INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSACOP.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSTRANSACOP.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSTRANSACOP.DIVSN where SSTRANSACOP.ENCODE_DT BETWEEN '" & date12 & "' and '" & date22 & "' GROUP BY SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSACOP.ENCODE_DT,SSTRANSACOP.BRANCH_IP")
                Dim rpt As New FIGACOPALL
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date12)
                cryRpt.SetParameterValue("@DateTo", date22)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf rbacop.Checked = True And rbAll.Checked = True Then
                If day1 = "1" And day2 = "31" Or day1 = "1" And day2 = "30" Then
                    GetNameBranch()
                    getMonth()
                    GetName = "COUNT OF ACOP"
                    'dt = db.ExecuteSQLQuery("select SSINFOTERMKIOSK.KIOSK_ID,SSTRANSACOP.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSACOP.ENCODE_DT,COUNT(TRANID)as [COUNT] from SSTRANSACOP INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.BRANCH_IP = SSTRANSACOP.BRANCH_IP INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSACOP.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSTRANSACOP.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSTRANSACOP.DIVSN where SSTRANSACOP.ENCODE_DT BETWEEN '" & date12 & "' and '" & date22 & "' GROUP BY SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSACOP.ENCODE_DT,SSTRANSACOP.BRANCH_IP")
                    Dim rpt As New FIGACOPALL
                    cryRpt = rpt
                    cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                    rptView.ReportSource = cryRpt
                    cryRpt.SetParameterValue("@DateFrom", date12)
                    cryRpt.SetParameterValue("@DateTo", date22)
                    cryRpt.SetParameterValue("@getName", getBranch)
                Else
                    GetNameBranch()
                    getMonth()
                    GetName = "COUNT OF ACOP"
                    'dt = db.ExecuteSQLQuery("select SSINFOTERMKIOSK.KIOSK_ID,SSTRANSACOP.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSACOP.ENCODE_DT,COUNT(TRANID)as [COUNT] from SSTRANSACOP INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.BRANCH_IP = SSTRANSACOP.BRANCH_IP INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSACOP.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSTRANSACOP.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSTRANSACOP.DIVSN where SSTRANSACOP.ENCODE_DT BETWEEN '" & date12 & "' and '" & date22 & "' GROUP BY SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSACOP.ENCODE_DT,SSTRANSACOP.BRANCH_IP")
                    Dim rpt As New FIGACOPALL
                    cryRpt = rpt
                    cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                    rptView.ReportSource = cryRpt
                    cryRpt.SetParameterValue("@DateFrom", date12)
                    cryRpt.SetParameterValue("@DateTo", date22)
                    cryRpt.SetParameterValue("@getName", getBranch)
                End If
            ElseIf rbacopdep.Checked = True And rbBranch.Checked = True Then
                GetNameBranch()
                GetName = "CHANGE OF STATUS"
                'dt = db.ExecuteSQLQuery("select SSINFOTERMBR.BRANCH_NM,SSTRANSACOPAD.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSACOPAD.ENCODE_DT,COUNT(SSTRANSACOPAD.TRANID)as [COUNT] from SSTRANSACOPAD INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSACOPAD.BRANCH_CD INNER JOIN IKBENEFREPTF ON SSTRANSACOPAD.TRANID = IKBENEFREPTF.TRANID INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID = SSTRANSACOPAD.KIOSK_ID INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMBR.CLSTR_CD = SSINFOTERMCLSTR.CLSTR_CD INNER JOIN SSINFOTERMGROUP ON SSINFOTERMBR.GROUP_CD = SSINFOTERMGROUP.GROUP_CD where SSTRANSACOPAD.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' GROUP BY SSINFOTERMBR.BRANCH_NM	,SSTRANSACOPAD.ENCODE_DT	,SSTRANSACOPAD.TRANID,SSINFOTERMKIOSK.BRANCH_IP,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSACOPAD.KIOSK_ID")
                Dim rpt As New FIGACOPDEPALL
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf rbacopdep.Checked = True And rbCluster.Checked = True Then
                GetNameBranch()
                GetName = "CHANGE OF STATUS"
                'dt = db.ExecuteSQLQuery("select SSINFOTERMBR.BRANCH_NM,SSTRANSACOPAD.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSACOPAD.ENCODE_DT,COUNT(SSTRANSACOPAD.TRANID)as [COUNT] from SSTRANSACOPAD INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSACOPAD.BRANCH_CD INNER JOIN IKBENEFREPTF ON SSTRANSACOPAD.TRANID = IKBENEFREPTF.TRANID INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID = SSTRANSACOPAD.KIOSK_ID INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMBR.CLSTR_CD = SSINFOTERMCLSTR.CLSTR_CD INNER JOIN SSINFOTERMGROUP ON SSINFOTERMBR.GROUP_CD = SSINFOTERMGROUP.GROUP_CD where SSTRANSACOPAD.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' GROUP BY SSINFOTERMBR.BRANCH_NM	,SSTRANSACOPAD.ENCODE_DT	,SSTRANSACOPAD.TRANID,SSINFOTERMKIOSK.BRANCH_IP,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSACOPAD.KIOSK_ID")
                Dim rpt As New FIGACOPDEPALL
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf rbacopdep.Checked = True And rbGroup.Checked = True Then
                GetNameBranch()
                GetName = "CHANGE OF STATUS"
                'dt = db.ExecuteSQLQuery("select SSINFOTERMBR.BRANCH_NM,SSTRANSACOPAD.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSACOPAD.ENCODE_DT,COUNT(SSTRANSACOPAD.TRANID)as [COUNT] from SSTRANSACOPAD INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSACOPAD.BRANCH_CD INNER JOIN IKBENEFREPTF ON SSTRANSACOPAD.TRANID = IKBENEFREPTF.TRANID INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID = SSTRANSACOPAD.KIOSK_ID INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMBR.CLSTR_CD = SSINFOTERMCLSTR.CLSTR_CD INNER JOIN SSINFOTERMGROUP ON SSINFOTERMBR.GROUP_CD = SSINFOTERMGROUP.GROUP_CD where SSTRANSACOPAD.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' GROUP BY SSINFOTERMBR.BRANCH_NM	,SSTRANSACOPAD.ENCODE_DT	,SSTRANSACOPAD.TRANID,SSINFOTERMKIOSK.BRANCH_IP,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSACOPAD.KIOSK_ID")
                Dim rpt As New FIGACOPDEPALL
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf rbacopdep.Checked = True And rbAll.Checked = True Then
                If day1 = "1" And day2 = "31" Or day1 = "1" And day2 = "30" Then
                    'getMonth()
                    'GetName = "CHANGE OF STATUS"
                    ''dt = db.ExecuteSQLQuery("select SSINFOTERMBR.BRANCH_NM,SSTRANSACOPAD.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSACOPAD.ENCODE_DT,COUNT(SSTRANSACOPAD.TRANID)as [COUNT] from SSTRANSACOPAD INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSACOPAD.BRANCH_CD INNER JOIN IKBENEFREPTF ON SSTRANSACOPAD.TRANID = IKBENEFREPTF.TRANID INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID = SSTRANSACOPAD.KIOSK_ID INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMBR.CLSTR_CD = SSINFOTERMCLSTR.CLSTR_CD INNER JOIN SSINFOTERMGROUP ON SSINFOTERMBR.GROUP_CD = SSINFOTERMGROUP.GROUP_CD where SSTRANSACOPAD.ENCODE_DT BETWEEN '" & date12 & "' and '" & date22 & "' GROUP BY SSINFOTERMBR.BRANCH_NM	,SSTRANSACOPAD.ENCODE_DT	,SSTRANSACOPAD.TRANID,SSINFOTERMKIOSK.BRANCH_IP,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSACOPAD.KIOSK_ID")
                    'Dim rpt As New FIGACOPDEPALLMONTHLY
                    'cryRpt = rpt
                    'cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass, My.Settings.db_Host, My.Settings.db_Server)
                    'rptView.ReportSource = cryRpt
                    'cryRpt.SetParameterValue("@DateFrom", date12)
                    'cryRpt.SetParameterValue("@DateTo", date22)
                Else
                    GetNameBranch()
                    GetName = "CHANGE OF STATUS"
                    'dt = db.ExecuteSQLQuery("select SSINFOTERMBR.BRANCH_NM,SSTRANSACOPAD.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSACOPAD.ENCODE_DT,COUNT(SSTRANSACOPAD.TRANID)as [COUNT] from SSTRANSACOPAD INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSACOPAD.BRANCH_CD INNER JOIN IKBENEFREPTF ON SSTRANSACOPAD.TRANID = IKBENEFREPTF.TRANID INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID = SSTRANSACOPAD.KIOSK_ID INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMBR.CLSTR_CD = SSINFOTERMCLSTR.CLSTR_CD INNER JOIN SSINFOTERMGROUP ON SSINFOTERMBR.GROUP_CD = SSINFOTERMGROUP.GROUP_CD where SSTRANSACOPAD.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' GROUP BY SSINFOTERMBR.BRANCH_NM	,SSTRANSACOPAD.ENCODE_DT	,SSTRANSACOPAD.TRANID,SSINFOTERMKIOSK.BRANCH_IP,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSACOPAD.KIOSK_ID")
                    Dim rpt As New FIGACOPDEPALL
                    cryRpt = rpt
                    cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                    rptView.ReportSource = cryRpt
                    cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                    cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                    cryRpt.SetParameterValue("@getName", getBranch)
                End If
            ElseIf rbSucFail.Checked = True And rbAll.Checked = True Then
                getMonth()
                If day1 = "1" And day2 = "31" Or day1 = "1" And day2 = "30" Then
                    GetNameBranch()
                    GetName = "SUCCESSANDFAILED"
                    'dt = db.ExecuteSQLQuery("SELECT DISTINCT SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM, (SELECT COUNT(ID) from SSTRANSAPPSL where SSTRANSAPPSL.IN_IPADD = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAPPSL.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') + (SELECT COUNT(INT) from SSTRANSAPPSLEMP where SSTRANSAPPSLEMP.IN_IPADD = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAPPSLEMP.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS SL, (SELECT COUNT(ID) FROM SSTRANSACOP where SSTRANSACOP.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSACOP.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS ACP, (SELECT COUNT(ID) FROM SSTRANSINFOTERMTR where SSTRANSINFOTERMTR.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSINFOTERMTR.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') +  (SELECT COUNT(ID) FROM SSTRANSINFOTERMTRLS where SSTRANSINFOTERMTRLS.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSINFOTERMTRLS.DATECREATED BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS TR, (SELECT COUNT(SSTRANSINFOTERMRG.ID) FROM SSTRANSINFOTERMRG where SSTRANSINFOTERMRG.KIOSK_ID = SSINFOTERMKIOSK.KIOSK_ID and SSTRANSINFOTERMRG.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS SWR, (SELECT COUNT(SSTRANSINFORTERMMN.SEQ_NO) FROM SSTRANSINFORTERMMN where SSTRANSINFORTERMMN.KIOSK_ID = SSINFOTERMKIOSK.KIOSK_ID and SSTRANSINFORTERMMN.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS MN, (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10001' AND TRANS_DESC <> '' AND TRANS_DESC NOT LIKE '%SERVICE IS TEMPORARILY UNAVAILABLE%' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS FSWR, (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10027' AND TRANS_DESC <> '' AND TRANS_DESC NOT LIKE '%SERVICE IS TEMPORARILY UNAVAILABLE%' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS FMAT, (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10032' AND TRANS_DESC <> '' AND TRANS_DESC NOT LIKE '%SERVICE IS TEMPORARILY UNAVAILABLE%' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS FSAL, (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10028' AND TRANS_DESC <> '' AND TRANS_DESC NOT LIKE '%SERVICE IS TEMPORARILY UNAVAILABLE%' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS FTR, (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10029' AND TRANS_DESC <> '' AND TRANS_DESC NOT LIKE '%SERVICE IS TEMPORARILY UNAVAILABLE%' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS FACOP FROM SSINFOTERMKIOSK LEFT JOIN SSTRANSAT ON SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAT.BRANCH_IP LEFT JOIN SSTRANSACOP ON SSTRANSACOP.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP LEFT JOIN SSTRANSAPPSL ON SSTRANSAPPSL.IN_IPADD = SSINFOTERMKIOSK.BRANCH_IP LEFT JOIN SSTRANSAPPSLEMP ON SSTRANSAPPSLEMP.IN_IPADD = SSINFOTERMKIOSK.BRANCH_IP LEFT JOIN SSTRANSINFORTERMMN ON SSTRANSINFORTERMMN.KIOSK_ID = SSINFOTERMKIOSK.KIOSK_ID LEFT JOIN SSTRANSINFOTERMRG ON SSTRANSINFOTERMRG.KIOSK_ID = SSINFOTERMKIOSK.KIOSK_ID LEFT JOIN SSTRANSINFOTERMTR ON SSTRANSINFOTERMTR.BRANCH_IP = SSINFOTERMKIOSK.KIOSK_ID LEFT JOIN SSTRANSINFOTERMTRLS ON SSTRANSINFOTERMTRLS.BRANCH_IP = SSINFOTERMKIOSK.KIOSK_ID INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSINFOTERMKIOSK.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMKIOSK.CLSTR = SSINFOTERMCLSTR.CLSTR_CD INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSINFOTERMKIOSK.DIVSN ORDER BY SSINFOTERMKIOSK.KIOSK_ID")
                    Dim rpt As New SUCFAILEDALL
                    cryRpt = rpt
                    cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                    rptView.ReportSource = cryRpt
                    cryRpt.SetParameterValue("@DateFrom", date12)
                    cryRpt.SetParameterValue("@DateTo", date22)
                    cryRpt.SetParameterValue("@getName", getBranch)
                Else
                    GetNameBranch()
                    GetName = "SUCCESSANDFAILED"
                    'dt = db.ExecuteSQLQuery("SELECT DISTINCT SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM, (SELECT COUNT(ID) from SSTRANSAPPSL where SSTRANSAPPSL.IN_IPADD = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAPPSL.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') + (SELECT COUNT(INT) from SSTRANSAPPSLEMP where SSTRANSAPPSLEMP.IN_IPADD = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAPPSLEMP.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS SL, (SELECT COUNT(ID) FROM SSTRANSACOP where SSTRANSACOP.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSACOP.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS ACP, (SELECT COUNT(ID) FROM SSTRANSINFOTERMTR where SSTRANSINFOTERMTR.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSINFOTERMTR.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') +  (SELECT COUNT(ID) FROM SSTRANSINFOTERMTRLS where SSTRANSINFOTERMTRLS.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSINFOTERMTRLS.DATECREATED BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS TR, (SELECT COUNT(SSTRANSINFOTERMRG.ID) FROM SSTRANSINFOTERMRG where SSTRANSINFOTERMRG.KIOSK_ID = SSINFOTERMKIOSK.KIOSK_ID and SSTRANSINFOTERMRG.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS SWR, (SELECT COUNT(SSTRANSINFORTERMMN.SEQ_NO) FROM SSTRANSINFORTERMMN where SSTRANSINFORTERMMN.KIOSK_ID = SSINFOTERMKIOSK.KIOSK_ID and SSTRANSINFORTERMMN.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS MN, (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10001' AND TRANS_DESC <> '' AND TRANS_DESC NOT LIKE '%SERVICE IS TEMPORARILY UNAVAILABLE%' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS FSWR, (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10027' AND TRANS_DESC <> '' AND TRANS_DESC NOT LIKE '%SERVICE IS TEMPORARILY UNAVAILABLE%' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS FMAT, (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10032' AND TRANS_DESC <> '' AND TRANS_DESC NOT LIKE '%SERVICE IS TEMPORARILY UNAVAILABLE%' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS FSAL, (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10028' AND TRANS_DESC <> '' AND TRANS_DESC NOT LIKE '%SERVICE IS TEMPORARILY UNAVAILABLE%' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS FTR, (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10029' AND TRANS_DESC <> '' AND TRANS_DESC NOT LIKE '%SERVICE IS TEMPORARILY UNAVAILABLE%' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS FACOP FROM SSINFOTERMKIOSK LEFT JOIN SSTRANSAT ON SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAT.BRANCH_IP LEFT JOIN SSTRANSACOP ON SSTRANSACOP.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP LEFT JOIN SSTRANSAPPSL ON SSTRANSAPPSL.IN_IPADD = SSINFOTERMKIOSK.BRANCH_IP LEFT JOIN SSTRANSAPPSLEMP ON SSTRANSAPPSLEMP.IN_IPADD = SSINFOTERMKIOSK.BRANCH_IP LEFT JOIN SSTRANSINFORTERMMN ON SSTRANSINFORTERMMN.KIOSK_ID = SSINFOTERMKIOSK.KIOSK_ID LEFT JOIN SSTRANSINFOTERMRG ON SSTRANSINFOTERMRG.KIOSK_ID = SSINFOTERMKIOSK.KIOSK_ID LEFT JOIN SSTRANSINFOTERMTR ON SSTRANSINFOTERMTR.BRANCH_IP = SSINFOTERMKIOSK.KIOSK_ID LEFT JOIN SSTRANSINFOTERMTRLS ON SSTRANSINFOTERMTRLS.BRANCH_IP = SSINFOTERMKIOSK.KIOSK_ID INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSINFOTERMKIOSK.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMKIOSK.CLSTR = SSINFOTERMCLSTR.CLSTR_CD INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSINFOTERMKIOSK.DIVSN ORDER BY SSINFOTERMKIOSK.KIOSK_ID")
                    Dim rpt As New SUCFAILEDALL
                    cryRpt = rpt
                    cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                    rptView.ReportSource = cryRpt
                    cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                    cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                    'cryRpt.SetParameterValue("@getName", getBranch)
                End If
            ElseIf rbSucFail.Checked = True And rbBranch.Checked = True Then
                GetNameBranch()
                GetName = "SUCCESSANDFAILED"
                'dt = db.ExecuteSQLQuery("SELECT DISTINCT SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM, (SELECT COUNT(ID) from SSTRANSAPPSL where SSTRANSAPPSL.IN_IPADD = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAPPSL.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') + (SELECT COUNT(INT) from SSTRANSAPPSLEMP where SSTRANSAPPSLEMP.IN_IPADD = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAPPSLEMP.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS SL, (SELECT COUNT(ID) FROM SSTRANSACOP where SSTRANSACOP.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSACOP.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS ACP, (SELECT COUNT(ID) FROM SSTRANSINFOTERMTR where SSTRANSINFOTERMTR.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSINFOTERMTR.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') +  (SELECT COUNT(ID) FROM SSTRANSINFOTERMTRLS where SSTRANSINFOTERMTRLS.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSINFOTERMTRLS.DATECREATED BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS TR, (SELECT COUNT(SSTRANSINFOTERMRG.ID) FROM SSTRANSINFOTERMRG where SSTRANSINFOTERMRG.KIOSK_ID = SSINFOTERMKIOSK.KIOSK_ID and SSTRANSINFOTERMRG.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS SWR, (SELECT COUNT(SSTRANSINFORTERMMN.SEQ_NO) FROM SSTRANSINFORTERMMN where SSTRANSINFORTERMMN.KIOSK_ID = SSINFOTERMKIOSK.KIOSK_ID and SSTRANSINFORTERMMN.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS MN, (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10001' AND TRANS_DESC <> '' AND TRANS_DESC NOT LIKE '%SERVICE IS TEMPORARILY UNAVAILABLE%' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS FSWR, (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10027' AND TRANS_DESC <> '' AND TRANS_DESC NOT LIKE '%SERVICE IS TEMPORARILY UNAVAILABLE%' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS FMAT, (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10032' AND TRANS_DESC <> '' AND TRANS_DESC NOT LIKE '%SERVICE IS TEMPORARILY UNAVAILABLE%' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS FSAL, (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10028' AND TRANS_DESC <> '' AND TRANS_DESC NOT LIKE '%SERVICE IS TEMPORARILY UNAVAILABLE%' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS FTR, (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10029' AND TRANS_DESC <> '' AND TRANS_DESC NOT LIKE '%SERVICE IS TEMPORARILY UNAVAILABLE%' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS FACOP FROM SSINFOTERMKIOSK LEFT JOIN SSTRANSAT ON SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAT.BRANCH_IP LEFT JOIN SSTRANSACOP ON SSTRANSACOP.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP LEFT JOIN SSTRANSAPPSL ON SSTRANSAPPSL.IN_IPADD = SSINFOTERMKIOSK.BRANCH_IP LEFT JOIN SSTRANSAPPSLEMP ON SSTRANSAPPSLEMP.IN_IPADD = SSINFOTERMKIOSK.BRANCH_IP LEFT JOIN SSTRANSINFORTERMMN ON SSTRANSINFORTERMMN.KIOSK_ID = SSINFOTERMKIOSK.KIOSK_ID LEFT JOIN SSTRANSINFOTERMRG ON SSTRANSINFOTERMRG.KIOSK_ID = SSINFOTERMKIOSK.KIOSK_ID LEFT JOIN SSTRANSINFOTERMTR ON SSTRANSINFOTERMTR.BRANCH_IP = SSINFOTERMKIOSK.KIOSK_ID LEFT JOIN SSTRANSINFOTERMTRLS ON SSTRANSINFOTERMTRLS.BRANCH_IP = SSINFOTERMKIOSK.KIOSK_ID INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSINFOTERMKIOSK.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMKIOSK.CLSTR = SSINFOTERMCLSTR.CLSTR_CD INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSINFOTERMKIOSK.DIVSN ORDER BY SSINFOTERMKIOSK.KIOSK_ID")
                Dim rpt As New SUCFAILEDALL
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf rbSucFail.Checked = True And rbCluster.Checked = True Then
                GetNameBranch()
                GetName = "SUCCESSANDFAILED"
                'dt = db.ExecuteSQLQuery("SELECT DISTINCT SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM, (SELECT COUNT(ID) from SSTRANSAPPSL where SSTRANSAPPSL.IN_IPADD = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAPPSL.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') + (SELECT COUNT(INT) from SSTRANSAPPSLEMP where SSTRANSAPPSLEMP.IN_IPADD = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAPPSLEMP.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS SL, (SELECT COUNT(ID) FROM SSTRANSACOP where SSTRANSACOP.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSACOP.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS ACP, (SELECT COUNT(ID) FROM SSTRANSINFOTERMTR where SSTRANSINFOTERMTR.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSINFOTERMTR.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') +  (SELECT COUNT(ID) FROM SSTRANSINFOTERMTRLS where SSTRANSINFOTERMTRLS.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSINFOTERMTRLS.DATECREATED BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS TR, (SELECT COUNT(SSTRANSINFOTERMRG.ID) FROM SSTRANSINFOTERMRG where SSTRANSINFOTERMRG.KIOSK_ID = SSINFOTERMKIOSK.KIOSK_ID and SSTRANSINFOTERMRG.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS SWR, (SELECT COUNT(SSTRANSINFORTERMMN.SEQ_NO) FROM SSTRANSINFORTERMMN where SSTRANSINFORTERMMN.KIOSK_ID = SSINFOTERMKIOSK.KIOSK_ID and SSTRANSINFORTERMMN.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS MN, (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10001' AND TRANS_DESC <> '' AND TRANS_DESC NOT LIKE '%SERVICE IS TEMPORARILY UNAVAILABLE%' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS FSWR, (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10027' AND TRANS_DESC <> '' AND TRANS_DESC NOT LIKE '%SERVICE IS TEMPORARILY UNAVAILABLE%' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS FMAT, (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10032' AND TRANS_DESC <> '' AND TRANS_DESC NOT LIKE '%SERVICE IS TEMPORARILY UNAVAILABLE%' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS FSAL, (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10028' AND TRANS_DESC <> '' AND TRANS_DESC NOT LIKE '%SERVICE IS TEMPORARILY UNAVAILABLE%' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS FTR, (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10029' AND TRANS_DESC <> '' AND TRANS_DESC NOT LIKE '%SERVICE IS TEMPORARILY UNAVAILABLE%' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS FACOP FROM SSINFOTERMKIOSK LEFT JOIN SSTRANSAT ON SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAT.BRANCH_IP LEFT JOIN SSTRANSACOP ON SSTRANSACOP.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP LEFT JOIN SSTRANSAPPSL ON SSTRANSAPPSL.IN_IPADD = SSINFOTERMKIOSK.BRANCH_IP LEFT JOIN SSTRANSAPPSLEMP ON SSTRANSAPPSLEMP.IN_IPADD = SSINFOTERMKIOSK.BRANCH_IP LEFT JOIN SSTRANSINFORTERMMN ON SSTRANSINFORTERMMN.KIOSK_ID = SSINFOTERMKIOSK.KIOSK_ID LEFT JOIN SSTRANSINFOTERMRG ON SSTRANSINFOTERMRG.KIOSK_ID = SSINFOTERMKIOSK.KIOSK_ID LEFT JOIN SSTRANSINFOTERMTR ON SSTRANSINFOTERMTR.BRANCH_IP = SSINFOTERMKIOSK.KIOSK_ID LEFT JOIN SSTRANSINFOTERMTRLS ON SSTRANSINFOTERMTRLS.BRANCH_IP = SSINFOTERMKIOSK.KIOSK_ID INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSINFOTERMKIOSK.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMKIOSK.CLSTR = SSINFOTERMCLSTR.CLSTR_CD INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSINFOTERMKIOSK.DIVSN ORDER BY SSINFOTERMKIOSK.KIOSK_ID")
                Dim rpt As New SUCFAILEDALL
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf rbSucFail.Checked = True And rbGroup.Checked = True Then
                GetNameBranch()
                GetName = "SUCCESSANDFAILED"
                'dt = db.ExecuteSQLQuery("SELECT DISTINCT SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM, (SELECT COUNT(ID) from SSTRANSAPPSL where SSTRANSAPPSL.IN_IPADD = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAPPSL.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') + (SELECT COUNT(INT) from SSTRANSAPPSLEMP where SSTRANSAPPSLEMP.IN_IPADD = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAPPSLEMP.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS SL, (SELECT COUNT(ID) FROM SSTRANSACOP where SSTRANSACOP.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSACOP.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS ACP, (SELECT COUNT(ID) FROM SSTRANSINFOTERMTR where SSTRANSINFOTERMTR.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSINFOTERMTR.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') +  (SELECT COUNT(ID) FROM SSTRANSINFOTERMTRLS where SSTRANSINFOTERMTRLS.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSINFOTERMTRLS.DATECREATED BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS TR, (SELECT COUNT(SSTRANSINFOTERMRG.ID) FROM SSTRANSINFOTERMRG where SSTRANSINFOTERMRG.KIOSK_ID = SSINFOTERMKIOSK.KIOSK_ID and SSTRANSINFOTERMRG.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS SWR, (SELECT COUNT(SSTRANSINFORTERMMN.SEQ_NO) FROM SSTRANSINFORTERMMN where SSTRANSINFORTERMMN.KIOSK_ID = SSINFOTERMKIOSK.KIOSK_ID and SSTRANSINFORTERMMN.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS MN, (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10001' AND TRANS_DESC <> '' AND TRANS_DESC NOT LIKE '%SERVICE IS TEMPORARILY UNAVAILABLE%' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS FSWR, (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10027' AND TRANS_DESC <> '' AND TRANS_DESC NOT LIKE '%SERVICE IS TEMPORARILY UNAVAILABLE%' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS FMAT, (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10032' AND TRANS_DESC <> '' AND TRANS_DESC NOT LIKE '%SERVICE IS TEMPORARILY UNAVAILABLE%' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS FSAL, (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10028' AND TRANS_DESC <> '' AND TRANS_DESC NOT LIKE '%SERVICE IS TEMPORARILY UNAVAILABLE%' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS FTR, (SELECT COUNT(PROC_CD) FROM SSTRANSAT WHERE PROC_CD = '10029' AND TRANS_DESC <> '' AND TRANS_DESC NOT LIKE '%SERVICE IS TEMPORARILY UNAVAILABLE%' and SSTRANSAT.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP and SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "') AS FACOP FROM SSINFOTERMKIOSK LEFT JOIN SSTRANSAT ON SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAT.BRANCH_IP LEFT JOIN SSTRANSACOP ON SSTRANSACOP.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP LEFT JOIN SSTRANSAPPSL ON SSTRANSAPPSL.IN_IPADD = SSINFOTERMKIOSK.BRANCH_IP LEFT JOIN SSTRANSAPPSLEMP ON SSTRANSAPPSLEMP.IN_IPADD = SSINFOTERMKIOSK.BRANCH_IP LEFT JOIN SSTRANSINFORTERMMN ON SSTRANSINFORTERMMN.KIOSK_ID = SSINFOTERMKIOSK.KIOSK_ID LEFT JOIN SSTRANSINFOTERMRG ON SSTRANSINFOTERMRG.KIOSK_ID = SSINFOTERMKIOSK.KIOSK_ID LEFT JOIN SSTRANSINFOTERMTR ON SSTRANSINFOTERMTR.BRANCH_IP = SSINFOTERMKIOSK.KIOSK_ID LEFT JOIN SSTRANSINFOTERMTRLS ON SSTRANSINFOTERMTRLS.BRANCH_IP = SSINFOTERMKIOSK.KIOSK_ID INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSINFOTERMKIOSK.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMKIOSK.CLSTR = SSINFOTERMCLSTR.CLSTR_CD INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSINFOTERMKIOSK.DIVSN ORDER BY SSINFOTERMKIOSK.KIOSK_ID")
                Dim rpt As New SUCFAILEDALL
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf rbFTR.Checked = True And rbAll.Checked = True And s.Checked = True Then
                getMonth()
                If day1 = "1" And day2 = "31" Or day1 = "1" And day2 = "30" Then
                    GetNameBranch()
                    Dim rpt As New NEWTECHRETALL
                    cryRpt = rpt
                    cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                    rptView.ReportSource = cryRpt
                    cryRpt.SetParameterValue("@DateFrom", date12)
                    cryRpt.SetParameterValue("@DateTo", date22)
                    cryRpt.SetParameterValue("@getName", getBranch)
                Else
                    GetNameBranch()
                    GetName = "TECHNICAL RETIREMENT"
                    Dim rpt As New NEWTECHRETALL
                    cryRpt = rpt
                    cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                    rptView.ReportSource = cryRpt
                    cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                    cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                    cryRpt.SetParameterValue("@getName", getBranch)
                End If
            ElseIf rbFTR.Checked = True And rbBranch.Checked = True And s.Checked = True Then
                GetNameBranch()
                GetName = "TECHNICAL RETIREMENT"
                Dim rpt As New NEWTECHRETALL
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf rbFTR.Checked = True And rbCluster.Checked = True And s.Checked = True Then
                GetNameBranch()
                GetName = "TECHNICAL RETIREMENT"
                Dim rpt As New NEWTECHRETALL
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf rbFTR.Checked = True And rbGroup.Checked = True And s.Checked = True Then
                GetNameBranch()
                GetName = "TECHNICAL RETIREMENT"
                Dim rpt As New NEWTECHRETALL
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)

            ElseIf rbUFAILED.Checked = True And rbAll.Checked = True Then
                GetNameBranch()
                GetName = "SSIDREASONSALL"
                'dt = db.ExecuteSQLQuery("select SSTRANSERRORLOGS.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,CAST(SSTRANSERRORLOGS.ENCODE_DT as date)[date],CARDTYPE, sum(case when  TRANS_DESC like '%NOT THE LATEST CARD%' and CARDTYPE = 'UMIDCARD' then 1 else 0 end) [NOTLATESTCARD UMID],sum(case when TRANS_DESC like '%NOT ISSUED BY SSS%' and CARDTYPE = 'UMIDCARD' then 1 else 0 end) [NOTISSUEDBYSSS UMID],sum(case when TRANS_DESC like '%NO CARD DETECTED%' and CARDTYPE = 'UMIDCARD' then 1 else 0 end) [NOCARDDETECTED UMID],sum(case when TRANS_DESC like '%NO CARD DETECTED%' and CARDTYPE = 'OLDSSSCARD' then 1 else 0 end) [NOCARDDETECTED OLD],sum(case when TRANS_DESC like '%MEMBER_WITH_UMID_CARD%' and CARDTYPE = 'OLDSSSCARD' then 1 else 0 end) [MEMBERWIHUMID OLD] from SSTRANSERRORLOGS INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID  = SSTRANSERRORLOGS.KIOSK_ID INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSINFOTERMKIOSK.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMKIOSK.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSINFOTERMKIOSK.DIVSN  where CAST(SSTRANSERRORLOGS.ENCODE_DT as date) BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' group by SSTRANSERRORLOGS.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,CAST(SSTRANSERRORLOGS.ENCODE_DT as date),SSINFOTERMKIOSK.BRANCH_IP,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,CARDTYPE ORDER BY CAST(SSTRANSERRORLOGS.ENCODE_DT as date)")
                Dim rpt As New SSIDREASONSALL
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf rbUFAILED.Checked = True And rbBranch.Checked = True Then
                GetNameBranch()
                GetName = "SSIDREASONSALL"
                'dt = db.ExecuteSQLQuery("select SSTRANSERRORLOGS.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,CAST(SSTRANSERRORLOGS.ENCODE_DT as date)[date],CARDTYPE, sum(case when  TRANS_DESC like '%NOT THE LATEST CARD%' and CARDTYPE = 'UMIDCARD' then 1 else 0 end) [NOTLATESTCARD UMID],sum(case when TRANS_DESC like '%NOT ISSUED BY SSS%' and CARDTYPE = 'UMIDCARD' then 1 else 0 end) [NOTISSUEDBYSSS UMID],sum(case when TRANS_DESC like '%NO CARD DETECTED%' and CARDTYPE = 'UMIDCARD' then 1 else 0 end) [NOCARDDETECTED UMID],sum(case when TRANS_DESC like '%NO CARD DETECTED%' and CARDTYPE = 'OLDSSSCARD' then 1 else 0 end) [NOCARDDETECTED OLD],sum(case when TRANS_DESC like '%MEMBER_WITH_UMID_CARD%' and CARDTYPE = 'OLDSSSCARD' then 1 else 0 end) [MEMBERWIHUMID OLD] from SSTRANSERRORLOGS INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID  = SSTRANSERRORLOGS.KIOSK_ID INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSINFOTERMKIOSK.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMKIOSK.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSINFOTERMKIOSK.DIVSN  where CAST(SSTRANSERRORLOGS.ENCODE_DT as date) BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' group by SSTRANSERRORLOGS.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,CAST(SSTRANSERRORLOGS.ENCODE_DT as date),SSINFOTERMKIOSK.BRANCH_IP,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,CARDTYPE ORDER BY CAST(SSTRANSERRORLOGS.ENCODE_DT as date)")
                Dim rpt As New SSIDREASONSALL
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf rbUFAILED.Checked = True And rbCluster.Checked = True Then
                GetNameBranch()
                GetName = "SSIDREASONSALL"
                'dt = db.ExecuteSQLQuery("select SSTRANSERRORLOGS.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,CAST(SSTRANSERRORLOGS.ENCODE_DT as date)[date],CARDTYPE, sum(case when  TRANS_DESC like '%NOT THE LATEST CARD%' and CARDTYPE = 'UMIDCARD' then 1 else 0 end) [NOTLATESTCARD UMID],sum(case when TRANS_DESC like '%NOT ISSUED BY SSS%' and CARDTYPE = 'UMIDCARD' then 1 else 0 end) [NOTISSUEDBYSSS UMID],sum(case when TRANS_DESC like '%NO CARD DETECTED%' and CARDTYPE = 'UMIDCARD' then 1 else 0 end) [NOCARDDETECTED UMID],sum(case when TRANS_DESC like '%NO CARD DETECTED%' and CARDTYPE = 'OLDSSSCARD' then 1 else 0 end) [NOCARDDETECTED OLD],sum(case when TRANS_DESC like '%MEMBER_WITH_UMID_CARD%' and CARDTYPE = 'OLDSSSCARD' then 1 else 0 end) [MEMBERWIHUMID OLD] from SSTRANSERRORLOGS INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID  = SSTRANSERRORLOGS.KIOSK_ID INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSINFOTERMKIOSK.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMKIOSK.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSINFOTERMKIOSK.DIVSN  where CAST(SSTRANSERRORLOGS.ENCODE_DT as date) BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' group by SSTRANSERRORLOGS.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,CAST(SSTRANSERRORLOGS.ENCODE_DT as date),SSINFOTERMKIOSK.BRANCH_IP,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,CARDTYPE ORDER BY CAST(SSTRANSERRORLOGS.ENCODE_DT as date)")
                Dim rpt As New SSIDREASONSALL
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf rbUFAILED.Checked = True And rbGroup.Checked = True Then
                GetNameBranch()
                GetName = "SSIDREASONSALL"
                'dt = db.ExecuteSQLQuery("select SSTRANSERRORLOGS.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,CAST(SSTRANSERRORLOGS.ENCODE_DT as date)[date],CARDTYPE, sum(case when  TRANS_DESC like '%NOT THE LATEST CARD%' and CARDTYPE = 'UMIDCARD' then 1 else 0 end) [NOTLATESTCARD UMID],sum(case when TRANS_DESC like '%NOT ISSUED BY SSS%' and CARDTYPE = 'UMIDCARD' then 1 else 0 end) [NOTISSUEDBYSSS UMID],sum(case when TRANS_DESC like '%NO CARD DETECTED%' and CARDTYPE = 'UMIDCARD' then 1 else 0 end) [NOCARDDETECTED UMID],sum(case when TRANS_DESC like '%NO CARD DETECTED%' and CARDTYPE = 'OLDSSSCARD' then 1 else 0 end) [NOCARDDETECTED OLD],sum(case when TRANS_DESC like '%MEMBER_WITH_UMID_CARD%' and CARDTYPE = 'OLDSSSCARD' then 1 else 0 end) [MEMBERWIHUMID OLD] from SSTRANSERRORLOGS INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID  = SSTRANSERRORLOGS.KIOSK_ID INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSINFOTERMKIOSK.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMKIOSK.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSINFOTERMKIOSK.DIVSN  where CAST(SSTRANSERRORLOGS.ENCODE_DT as date) BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' group by SSTRANSERRORLOGS.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,CAST(SSTRANSERRORLOGS.ENCODE_DT as date),SSINFOTERMKIOSK.BRANCH_IP,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,CARDTYPE ORDER BY CAST(SSTRANSERRORLOGS.ENCODE_DT as date)")
                Dim rpt As New SSIDREASONSALL
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)

            ElseIf rbErrorLogs.Checked = True And rbAll.Checked = True Then
                GetNameBranch()
                GetName = "SSIDERRORLOGS"
                'dt = db.ExecuteSQLQuery("select CAST(SSTRANSERRORLOGS.ENCODE_DT as date)[date],SSTRANSERRORLOGS.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,CASE WHEN CARDTYPE = 'UMIDCARD' then 'UMID CARD' else 'OLD SSS CARD' end as [CARDTYPE],TRANS_DESC FROM SSTRANSERRORLOGS INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID = SSTRANSERRORLOGS.KIOSK_ID INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSINFOTERMKIOSK.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMKIOSK.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSINFOTERMKIOSK.DIVSN where CAST(SSTRANSERRORLOGS.ENCODE_DT as date) BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' ORDER BY SSTRANSERRORLOGS.KIOSK_ID")
                Dim rpt As New SSIDLOGSALL
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf rbErrorLogs.Checked = True And rbBranch.Checked = True Then

                GetNameBranch()
                GetName = "SSIDERRORLOGS"
                'dt = db.ExecuteSQLQuery("select CAST(SSTRANSERRORLOGS.ENCODE_DT as date)[date],SSTRANSERRORLOGS.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,CASE WHEN CARDTYPE = 'UMIDCARD' then 'UMID CARD' else 'OLD SSS CARD' end as [CARDTYPE],TRANS_DESC FROM SSTRANSERRORLOGS INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID = SSTRANSERRORLOGS.KIOSK_ID INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSINFOTERMKIOSK.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMKIOSK.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSINFOTERMKIOSK.DIVSN where CAST(SSTRANSERRORLOGS.ENCODE_DT as date) BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' ORDER BY SSTRANSERRORLOGS.KIOSK_ID")
                Dim rpt As New SSIDLOGSALL
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)

            ElseIf rbErrorLogs.Checked = True And rbCluster.Checked = True Then

                GetNameBranch()
                GetName = "SSIDERRORLOGS"
                'dt = db.ExecuteSQLQuery("select CAST(SSTRANSERRORLOGS.ENCODE_DT as date)[date],SSTRANSERRORLOGS.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,CASE WHEN CARDTYPE = 'UMIDCARD' then 'UMID CARD' else 'OLD SSS CARD' end as [CARDTYPE],TRANS_DESC FROM SSTRANSERRORLOGS INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID = SSTRANSERRORLOGS.KIOSK_ID INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSINFOTERMKIOSK.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMKIOSK.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSINFOTERMKIOSK.DIVSN where CAST(SSTRANSERRORLOGS.ENCODE_DT as date) BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' ORDER BY SSTRANSERRORLOGS.KIOSK_ID")
                Dim rpt As New SSIDLOGSALL
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)

            ElseIf rbErrorLogs.Checked = True And rbGroup.Checked = True Then

                GetNameBranch()
                GetName = "SSIDERRORLOGS"
                'dt = db.ExecuteSQLQuery("select CAST(SSTRANSERRORLOGS.ENCODE_DT as date)[date],SSTRANSERRORLOGS.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,CASE WHEN CARDTYPE = 'UMIDCARD' then 'UMID CARD' else 'OLD SSS CARD' end as [CARDTYPE],TRANS_DESC FROM SSTRANSERRORLOGS INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID = SSTRANSERRORLOGS.KIOSK_ID INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSINFOTERMKIOSK.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMKIOSK.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSINFOTERMKIOSK.DIVSN where CAST(SSTRANSERRORLOGS.ENCODE_DT as date) BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' ORDER BY SSTRANSERRORLOGS.KIOSK_ID")
                Dim rpt As New SSIDLOGSALL
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)

                'ElseIf rbMonitoringAll.Checked = True And rbAll.Checked = True Then

                '    GetName = "MONITORINGGROUP"
                '    'dt = db.ExecuteSQLQuery("select CAST(SSTRANSERRORLOGS.ENCODE_DT as date)[date],SSTRANSERRORLOGS.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,CASE WHEN CARDTYPE = 'UMIDCARD' then 'UMID CARD' else 'OLD SSS CARD' end as [CARDTYPE],TRANS_DESC FROM SSTRANSERRORLOGS INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID = SSTRANSERRORLOGS.KIOSK_ID INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSINFOTERMKIOSK.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMKIOSK.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSINFOTERMKIOSK.DIVSN where CAST(SSTRANSERRORLOGS.ENCODE_DT as date) BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and SSINFOTERMBR.BRANCH_NM = '" & cbBranch.Text & "' ORDER BY SSTRANSERRORLOGS.KIOSK_ID")

                '    Dim rpt As New SSIDLOGSBRANCH
                '    cryRpt = rpt
                '   cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                '    rptView.ReportSource = cryRpt
                '    cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                '    cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)

                'ElseIf rbMonitoringAll.Checked = True And rbBranch.Checked = True Then

                '    GetName = "MONITORINGGROUPBRANCH"
                '    'dt = db.ExecuteSQLQuery("select CAST(SSTRANSERRORLOGS.ENCODE_DT as date)[date],SSTRANSERRORLOGS.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,CASE WHEN CARDTYPE = 'UMIDCARD' then 'UMID CARD' else 'OLD SSS CARD' end as [CARDTYPE],TRANS_DESC FROM SSTRANSERRORLOGS INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID = SSTRANSERRORLOGS.KIOSK_ID INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSINFOTERMKIOSK.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMKIOSK.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSINFOTERMKIOSK.DIVSN where CAST(SSTRANSERRORLOGS.ENCODE_DT as date) BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and SSINFOTERMBR.BRANCH_NM = '" & cbBranch.Text & "' ORDER BY SSTRANSERRORLOGS.KIOSK_ID")
                '    Dim rpt As New SSIDLOGSBRANCH
                '    cryRpt = rpt
                '   cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                '    rptView.ReportSource = cryRpt
                '    cryRpt.SetParameterValue("@getBranch", cbBranch.Text)
                '    cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                '    cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)

            ElseIf rbMonitoringAll.Checked = True And rbCluster.Checked = True Then

                'GetName = "MONITORINGGROUPCLUSTER"
                ''dt = db.ExecuteSQLQuery("select SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSINFOTERMKIOSK.KIOSK_ID,case when STATUS = 1 then LONLINE_DT else null end as [LastOffline],case when STATUS = 1 then '*** OFFLINE ***' else 'ONLINE' end as [STATUS], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE DIVSN = 'N' and isVPN = 0 and TAG = 1 and STATUS = 0)[NCR ONLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE DIVSN = 'L' and isVPN = 0 and TAG = 1 and STATUS = 0)[LUZON ONLINE],  (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE DIVSN = 'V' and isVPN = 0 and TAG = 1 and STATUS = 0)[VISAYAS ONLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE DIVSN = 'M' and isVPN = 0 and TAG = 1 and STATUS = 0)[MINDANAO ONLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE DIVSN = 'F' and isVPN = 0 and TAG = 1 and STATUS = 0)[FOREGIN ONLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE CLSTR = '04' and isVPN = 0 and TAG = 1 and STATUS = 0)[NCR NORTH ONLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE CLSTR = '05' and isVPN = 0 and TAG = 1 and STATUS = 0)[NCR CENTRAL ONLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE CLSTR = '06' and isVPN = 0 and TAG = 1 and STATUS = 0)[NCR SOUTH ONLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE CLSTR = '01' and isVPN = 0 and TAG = 1 and STATUS = 0)[LUZON NORTH ONLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE CLSTR = '02' and isVPN = 0 and TAG = 1 and STATUS = 0)[LUZON CENTRAL ONLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE CLSTR = '07' and isVPN = 0 and TAG = 1 and STATUS = 0)[BICOL REGION ONLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE CLSTR = '03' and isVPN = 0 and TAG = 1 and STATUS = 0)[LUZON SOUTH ONLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE CLSTR = '08' and isVPN = 0 and TAG = 1 and STATUS = 0)[CENTRAL VISAYAS ONLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE CLSTR = '09' and isVPN = 0 and TAG = 1 and STATUS = 0)[WESTERN VISAYAS ONLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE CLSTR = '10' and isVPN = 0 and TAG = 1 and STATUS = 0)[NORTHERN MINDANAO ONLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE CLSTR = '11' and isVPN = 0 and TAG = 1 and STATUS = 0)[SOUTHERN MINDANAO ONLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE CLSTR = '12' and isVPN = 0 and TAG = 1 and STATUS = 0)[WESTERN MINDANAO ONLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE CLSTR = '13' and isVPN = 0 and TAG = 1 and STATUS = 0)[FOREIGN ONLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE DIVSN = 'N' and isVPN = 0 and TAG = 1 and STATUS = 1)[NCR OFFLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE DIVSN = 'L' and isVPN = 0 and TAG = 1 and STATUS = 1)[LUZON OFFLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE DIVSN = 'V' and isVPN = 0 and TAG = 1 and STATUS = 1)[VISAYAS OFFLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE DIVSN = 'M' and isVPN = 0 and TAG = 1 and STATUS = 1)[MINDANAO OFFLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE DIVSN = 'F' and isVPN = 0 and TAG = 1 and STATUS = 1)[FOREGIN OFFLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE CLSTR = '04' and isVPN = 0 and TAG = 1 and STATUS = 1)[NCR NORTH OFFLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE CLSTR = '05' and isVPN = 0 and TAG = 1 and STATUS = 1)[NCR CENTRAL OFFLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE CLSTR = '06' and isVPN = 0 and TAG = 1 and STATUS = 1)[NCR SOUTH OFFLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE CLSTR = '01' and isVPN = 0 and TAG = 1 and STATUS = 1)[LUZON NORTH OFFLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE CLSTR = '02' and isVPN = 0 and TAG = 1 and STATUS = 1)[LUZON CENTRAL OFFLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE CLSTR = '07' and isVPN = 0 and TAG = 1 and STATUS = 1)[BICOL REGION OFFLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE CLSTR = '03' and isVPN = 0 and TAG = 1 and STATUS = 1)[LUZON SOUTH OFFLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE CLSTR = '08' and isVPN = 0 and TAG = 1 and STATUS = 1)[CENTRAL VISAYAS OFFLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE CLSTR = '09' and isVPN = 0 and TAG = 1 and STATUS = 1)[WESTERN VISAYAS OFFLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE CLSTR = '10' and isVPN = 0 and TAG = 1 and STATUS = 1)[NORTHERN MINDANAO OFFLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE CLSTR = '11' and isVPN = 0 and TAG = 1 and STATUS = 1)[SOUTHERN MINDANAO OFFLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE CLSTR = '12' and isVPN = 0 and TAG = 1 and STATUS = 1)[WESTERN MINDANAO OFFLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE CLSTR = '13' and isVPN = 0 and TAG = 1 and STATUS = 1)[FOREIGN OFFLINE] FROM SSINFOTERMKIOSK INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSINFOTERMKIOSK.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMKIOSK.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSINFOTERMKIOSK.DIVSN where isVPN = 0 and TAG = 1 GROUP BY SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMKIOSK.STATUS,CLSTR,LONLINE_DT ORDER BY SSINFOTERMKIOSK.CLSTR")
                'Dim rpt As New KioskMonitoringDivision
                'cryRpt = rpt
                'cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass, My.Settings.db_Host, My.Settings.db_Server)
                'rptView.ReportSource = cryRpt
                ''cryRpt.SetParameterValue("@getCluster", cbCluster.Text)
                ''cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                ''cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)

            ElseIf rbMonitoringAll.Checked = True And rbAll.Checked = True Then

                GetName = "MONITORINGGROUPGROUP"
                'dt = db.ExecuteSQLQuery("select SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSINFOTERMKIOSK.KIOSK_ID,case when STATUS = 1 then LONLINE_DT else null end as [LastOffline],case when STATUS = 1 then '*** OFFLINE ***' else 'ONLINE' end as [STATUS], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE DIVSN = 'N' and isVPN = 0 and TAG = 1 and STATUS = 0)[NCR ONLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE DIVSN = 'L' and isVPN = 0 and TAG = 1 and STATUS = 0)[LUZON ONLINE],  (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE DIVSN = 'V' and isVPN = 0 and TAG = 1 and STATUS = 0)[VISAYAS ONLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE DIVSN = 'M' and isVPN = 0 and TAG = 1 and STATUS = 0)[MINDANAO ONLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE DIVSN = 'F' and isVPN = 0 and TAG = 1 and STATUS = 0)[FOREGIN ONLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE CLSTR = '04' and isVPN = 0 and TAG = 1 and STATUS = 0)[NCR NORTH ONLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE CLSTR = '05' and isVPN = 0 and TAG = 1 and STATUS = 0)[NCR CENTRAL ONLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE CLSTR = '06' and isVPN = 0 and TAG = 1 and STATUS = 0)[NCR SOUTH ONLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE CLSTR = '01' and isVPN = 0 and TAG = 1 and STATUS = 0)[LUZON NORTH ONLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE CLSTR = '02' and isVPN = 0 and TAG = 1 and STATUS = 0)[LUZON CENTRAL ONLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE CLSTR = '07' and isVPN = 0 and TAG = 1 and STATUS = 0)[BICOL REGION ONLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE CLSTR = '03' and isVPN = 0 and TAG = 1 and STATUS = 0)[LUZON SOUTH ONLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE CLSTR = '08' and isVPN = 0 and TAG = 1 and STATUS = 0)[CENTRAL VISAYAS ONLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE CLSTR = '09' and isVPN = 0 and TAG = 1 and STATUS = 0)[WESTERN VISAYAS ONLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE CLSTR = '10' and isVPN = 0 and TAG = 1 and STATUS = 0)[NORTHERN MINDANAO ONLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE CLSTR = '11' and isVPN = 0 and TAG = 1 and STATUS = 0)[SOUTHERN MINDANAO ONLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE CLSTR = '12' and isVPN = 0 and TAG = 1 and STATUS = 0)[WESTERN MINDANAO ONLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE CLSTR = '13' and isVPN = 0 and TAG = 1 and STATUS = 0)[FOREIGN ONLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE DIVSN = 'N' and isVPN = 0 and TAG = 1 and STATUS = 1)[NCR OFFLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE DIVSN = 'L' and isVPN = 0 and TAG = 1 and STATUS = 1)[LUZON OFFLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE DIVSN = 'V' and isVPN = 0 and TAG = 1 and STATUS = 1)[VISAYAS OFFLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE DIVSN = 'M' and isVPN = 0 and TAG = 1 and STATUS = 1)[MINDANAO OFFLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE DIVSN = 'F' and isVPN = 0 and TAG = 1 and STATUS = 1)[FOREGIN OFFLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE CLSTR = '04' and isVPN = 0 and TAG = 1 and STATUS = 1)[NCR NORTH OFFLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE CLSTR = '05' and isVPN = 0 and TAG = 1 and STATUS = 1)[NCR CENTRAL OFFLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE CLSTR = '06' and isVPN = 0 and TAG = 1 and STATUS = 1)[NCR SOUTH OFFLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE CLSTR = '01' and isVPN = 0 and TAG = 1 and STATUS = 1)[LUZON NORTH OFFLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE CLSTR = '02' and isVPN = 0 and TAG = 1 and STATUS = 1)[LUZON CENTRAL OFFLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE CLSTR = '07' and isVPN = 0 and TAG = 1 and STATUS = 1)[BICOL REGION OFFLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE CLSTR = '03' and isVPN = 0 and TAG = 1 and STATUS = 1)[LUZON SOUTH OFFLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE CLSTR = '08' and isVPN = 0 and TAG = 1 and STATUS = 1)[CENTRAL VISAYAS OFFLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE CLSTR = '09' and isVPN = 0 and TAG = 1 and STATUS = 1)[WESTERN VISAYAS OFFLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE CLSTR = '10' and isVPN = 0 and TAG = 1 and STATUS = 1)[NORTHERN MINDANAO OFFLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE CLSTR = '11' and isVPN = 0 and TAG = 1 and STATUS = 1)[SOUTHERN MINDANAO OFFLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE CLSTR = '12' and isVPN = 0 and TAG = 1 and STATUS = 1)[WESTERN MINDANAO OFFLINE], (SELECT COUNT(*) FROM SSINFOTERMKIOSK WHERE CLSTR = '13' and isVPN = 0 and TAG = 1 and STATUS = 1)[FOREIGN OFFLINE] FROM SSINFOTERMKIOSK INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSINFOTERMKIOSK.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMKIOSK.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSINFOTERMKIOSK.DIVSN where isVPN = 0 and TAG = 1 GROUP BY SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMKIOSK.STATUS,CLSTR,LONLINE_DT ORDER BY SSINFOTERMKIOSK.CLSTR")
                Dim rpt As New KioskMonitoring
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@dateToday", Today)
                'cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                'cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)

            ElseIf rbFTR.Checked = True And rbGroup.Checked = True And d.Checked = True Then
                GetNameBranch()
                GetName = "DAILY TECH RET REPORT SUMMARY"
                Dim rpt As New FIGTRALL
                cryRpt = rpt

                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)

            ElseIf rbFTR.Checked = True And rbCluster.Checked = True And d.Checked = True Then
                GetNameBranch()
                GetName = "DAILY TECH RET REPORT SUMMARY"
                Dim rpt As New FIGTRALL
                cryRpt = rpt

                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)


            ElseIf rbFTR.Checked = True And d.Checked = True And rbAll.Checked = True Then
                GetNameBranch()
                GetName = "DAILY TECH RET REPORT SUMMARY"
                Dim rpt As New FIGTRALL
                cryRpt = rpt

                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)

            ElseIf rbFTR.Checked = True And d.Checked = True And rbBranch.Checked = True Then
                GetNameBranch()
                GetName = "DAILY TECH RET REPORT SUMMARY"
                Dim rpt As New FIGTRALL
                cryRpt = rpt

                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)


            ElseIf rbFTR.Checked = True And w.Checked = True And rbAll.Checked = True Then
                '  pickyear = 0
                '  showyearonly()
                GetNameBranch()
                GetName = "WEEKLY TECH RET REPORT SUMMARY"
                'dt = db.ExecuteSQLQuery("SELECT SSINFOTERMCONSTAT.BRANCH_IP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,DATESTAMP,SSINFOTERMCLSTR.CLSTR_NM,ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0) as [ONLINE],case when(24 - ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0)) < 0 then 0 else (24 - ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0)) end as [OFFLINE] from SSINFOTERMCONSTAT inner join SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSINFOTERMCONSTAT.BRANCH_IP INNER JOIN SSINFOTERMBR on SSINFOTERMBR.BRANCH_CD = SSINFOTERMCONSTAT.BRANCH_CD INNER JOIN SSINFOTERMCLSTR on SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMCONSTAT.CLSTR where DATESTAMP between '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' GROUP by SSINFOTERMCONSTAT.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,DATESTAMP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMCLSTR.CLSTR_NM")
                Dim rpt As New FIGTRALLWEEKLY
                cryRpt = rpt
                '    cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)

            ElseIf rbFTR.Checked = True And w.Checked = True And rbBranch.Checked = True Then
                ' pickyear = 0
                ' showyearonly()
                GetNameBranch()
                GetName = "WEEKLY TECH RET REPORT SUMMARY"
                'dt = db.ExecuteSQLQuery("SELECT SSINFOTERMCONSTAT.BRANCH_IP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,DATESTAMP,SSINFOTERMCLSTR.CLSTR_NM,ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0) as [ONLINE],case when(24 - ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0)) < 0 then 0 else (24 - ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0)) end as [OFFLINE] from SSINFOTERMCONSTAT inner join SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSINFOTERMCONSTAT.BRANCH_IP INNER JOIN SSINFOTERMBR on SSINFOTERMBR.BRANCH_CD = SSINFOTERMCONSTAT.BRANCH_CD INNER JOIN SSINFOTERMCLSTR on SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMCONSTAT.CLSTR where DATESTAMP between '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' GROUP by SSINFOTERMCONSTAT.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,DATESTAMP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMCLSTR.CLSTR_NM")
                Dim rpt As New FIGTRALLWEEKLY
                cryRpt = rpt
                '    cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)

            ElseIf rbFTR.Checked = True And rbGroup.Checked = True And rbWeekly.Checked = True Then
                '  pickyear = 0
                '  showyearonly()
                GetNameBranch()
                GetName = "WEEKLY TECH RET REPORT SUMMARY"
                'dt = db.ExecuteSQLQuery("SELECT SSINFOTERMCONSTAT.BRANCH_IP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,DATESTAMP,SSINFOTERMCLSTR.CLSTR_NM,ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0) as [ONLINE],case when(24 - ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0)) < 0 then 0 else (24 - ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0)) end as [OFFLINE] from SSINFOTERMCONSTAT inner join SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSINFOTERMCONSTAT.BRANCH_IP INNER JOIN SSINFOTERMBR on SSINFOTERMBR.BRANCH_CD = SSINFOTERMCONSTAT.BRANCH_CD INNER JOIN SSINFOTERMCLSTR on SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMCONSTAT.CLSTR where DATESTAMP between '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' GROUP by SSINFOTERMCONSTAT.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,DATESTAMP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMCLSTR.CLSTR_NM")
                Dim rpt As New FIGTRALLWEEKLY
                cryRpt = rpt
                '    cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)

            ElseIf rbFTR.Checked = True And w.Checked = True And rbCluster.Checked = True Then
                ' pickyear = 0
                ' showyearonly()
                GetNameBranch()
                GetName = "WEEKLY TECH RET REPORT SUMMARY"
                'dt = db.ExecuteSQLQuery("SELECT SSINFOTERMCONSTAT.BRANCH_IP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,DATESTAMP,SSINFOTERMCLSTR.CLSTR_NM,ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0) as [ONLINE],case when(24 - ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0)) < 0 then 0 else (24 - ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0)) end as [OFFLINE] from SSINFOTERMCONSTAT inner join SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSINFOTERMCONSTAT.BRANCH_IP INNER JOIN SSINFOTERMBR on SSINFOTERMBR.BRANCH_CD = SSINFOTERMCONSTAT.BRANCH_CD INNER JOIN SSINFOTERMCLSTR on SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMCONSTAT.CLSTR where DATESTAMP between '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' GROUP by SSINFOTERMCONSTAT.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,DATESTAMP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMCLSTR.CLSTR_NM")
                Dim rpt As New FIGTRALLWEEKLY
                cryRpt = rpt
                '    cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)

            ElseIf rbFTR.Checked = True And m.Checked = True And rbGroup.Checked = True Then
                ' pickyear = 2
                'showyearonly()
                getMonth()
                GetNameBranch()
                GetName = "MONTHLY TECH RET REPORT SUMMARY"
                'dt = db.ExecuteSQLQuery("SELECT SSINFOTERMCONSTAT.BRANCH_IP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,DATESTAMP,SSINFOTERMCLSTR.CLSTR_NM,ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0) as [ONLINE],case when(24 - ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0)) < 0 then 0 else (24 - ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0)) end as [OFFLINE] from SSINFOTERMCONSTAT inner join SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSINFOTERMCONSTAT.BRANCH_IP INNER JOIN SSINFOTERMBR on SSINFOTERMBR.BRANCH_CD = SSINFOTERMCONSTAT.BRANCH_CD INNER JOIN SSINFOTERMCLSTR on SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMCONSTAT.CLSTR where DATESTAMP between '" & date12 & "' and '" & date22 & "' GROUP by SSINFOTERMCONSTAT.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,DATESTAMP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMCLSTR.CLSTR_NM")
                Dim rpt As New FIGTRALLMONTHLY
                cryRpt = rpt
                '  cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date12)
                cryRpt.SetParameterValue("@DateTo", date22)
                cryRpt.SetParameterValue("@getName", getBranch)

            ElseIf rbFTR.Checked = True And m.Checked = True And rbCluster.Checked = True Then
                ' pickyear = 2
                '  showyearonly()
                getMonth()
                GetNameBranch()
                GetName = "MONTHLY TECH RET REPORT SUMMARY"
                'dt = db.ExecuteSQLQuery("SELECT SSINFOTERMCONSTAT.BRANCH_IP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,DATESTAMP,SSINFOTERMCLSTR.CLSTR_NM,ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0) as [ONLINE],case when(24 - ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0)) < 0 then 0 else (24 - ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0)) end as [OFFLINE] from SSINFOTERMCONSTAT inner join SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSINFOTERMCONSTAT.BRANCH_IP INNER JOIN SSINFOTERMBR on SSINFOTERMBR.BRANCH_CD = SSINFOTERMCONSTAT.BRANCH_CD INNER JOIN SSINFOTERMCLSTR on SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMCONSTAT.CLSTR where DATESTAMP between '" & date12 & "' and '" & date22 & "' GROUP by SSINFOTERMCONSTAT.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,DATESTAMP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMCLSTR.CLSTR_NM")
                Dim rpt As New FIGTRALLMONTHLY
                cryRpt = rpt
                '  cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date12)
                cryRpt.SetParameterValue("@DateTo", date22)
                cryRpt.SetParameterValue("@getName", getBranch)

            ElseIf rbFTR.Checked = True And m.Checked = True And rbBranch.Checked = True Then
                ' pickyear = 2
                ' showyearonly()
                getMonth()
                GetNameBranch()
                GetName = "MONTHLY TECH RET REPORT SUMMARY"
                'dt = db.ExecuteSQLQuery("SELECT SSINFOTERMCONSTAT.BRANCH_IP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,DATESTAMP,SSINFOTERMCLSTR.CLSTR_NM,ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0) as [ONLINE],case when(24 - ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0)) < 0 then 0 else (24 - ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0)) end as [OFFLINE] from SSINFOTERMCONSTAT inner join SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSINFOTERMCONSTAT.BRANCH_IP INNER JOIN SSINFOTERMBR on SSINFOTERMBR.BRANCH_CD = SSINFOTERMCONSTAT.BRANCH_CD INNER JOIN SSINFOTERMCLSTR on SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMCONSTAT.CLSTR where DATESTAMP between '" & date12 & "' and '" & date22 & "' GROUP by SSINFOTERMCONSTAT.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,DATESTAMP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMCLSTR.CLSTR_NM")
                Dim rpt As New FIGTRALLMONTHLY
                cryRpt = rpt
                '  cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date12)
                cryRpt.SetParameterValue("@DateTo", date22)
                cryRpt.SetParameterValue("@getName", getBranch)


            ElseIf rbFTR.Checked = True And m.Checked = True And rbAll.Checked = True Then
                'pickyear = 2
                'showyearonly()
                getMonth()
                GetNameBranch()
                GetName = "MONTHLY TECH RET REPORT SUMMARY"
                'dt = db.ExecuteSQLQuery("SELECT SSINFOTERMCONSTAT.BRANCH_IP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,DATESTAMP,SSINFOTERMCLSTR.CLSTR_NM,ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0) as [ONLINE],case when(24 - ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0)) < 0 then 0 else (24 - ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0)) end as [OFFLINE] from SSINFOTERMCONSTAT inner join SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSINFOTERMCONSTAT.BRANCH_IP INNER JOIN SSINFOTERMBR on SSINFOTERMBR.BRANCH_CD = SSINFOTERMCONSTAT.BRANCH_CD INNER JOIN SSINFOTERMCLSTR on SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMCONSTAT.CLSTR where DATESTAMP between '" & date12 & "' and '" & date22 & "' GROUP by SSINFOTERMCONSTAT.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,DATESTAMP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMCLSTR.CLSTR_NM")
                Dim rpt As New FIGTRALLMONTHLY
                cryRpt = rpt
                '  cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date12)
                cryRpt.SetParameterValue("@DateTo", date22)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf y.Checked = True And rbFTR.Checked = True And rbAll.Checked = True Then
                GetNameBranch()
                getMonth()
                '  pickyear = 1
                '  showyearonly()
                GetName = "YEARLY TECH RET REPORT SUMMARY"
                'dt = db.ExecuteSQLQuery("SELECT SSINFOTERMCONSTAT.BRANCH_IP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,DATESTAMP,SSINFOTERMCLSTR.CLSTR_NM,ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0) as [ONLINE],case when(24 - ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0)) < 0 then 0 else (24 - ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0)) end as [OFFLINE] from SSINFOTERMCONSTAT inner join SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSINFOTERMCONSTAT.BRANCH_IP INNER JOIN SSINFOTERMBR on SSINFOTERMBR.BRANCH_CD = SSINFOTERMCONSTAT.BRANCH_CD INNER JOIN SSINFOTERMCLSTR on SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMCONSTAT.CLSTR where DATESTAMP between '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' GROUP by SSINFOTERMCONSTAT.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,DATESTAMP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMCLSTR.CLSTR_NM")
                Dim rpt As New FIGTRALLYEARLY
                cryRpt = rpt
                '  cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", DatePart(DateInterval.Year, dtpFrom.Value))
                cryRpt.SetParameterValue("@DateTo", DatePart(DateInterval.Year, dtpTo.Value))
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf y.Checked = True And rbFTR.Checked = True And rbGroup.Checked = True Then
                GetNameBranch()
                getMonth()
                '  pickyear = 1
                '  showyearonly()
                GetName = "YEARLY TECH RET REPORT SUMMARY"
                'dt = db.ExecuteSQLQuery("SELECT SSINFOTERMCONSTAT.BRANCH_IP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,DATESTAMP,SSINFOTERMCLSTR.CLSTR_NM,ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0) as [ONLINE],case when(24 - ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0)) < 0 then 0 else (24 - ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0)) end as [OFFLINE] from SSINFOTERMCONSTAT inner join SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSINFOTERMCONSTAT.BRANCH_IP INNER JOIN SSINFOTERMBR on SSINFOTERMBR.BRANCH_CD = SSINFOTERMCONSTAT.BRANCH_CD INNER JOIN SSINFOTERMCLSTR on SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMCONSTAT.CLSTR where DATESTAMP between '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' GROUP by SSINFOTERMCONSTAT.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,DATESTAMP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMCLSTR.CLSTR_NM")
                Dim rpt As New FIGTRALLYEARLY
                cryRpt = rpt
                '  cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", DatePart(DateInterval.Year, dtpFrom.Value))
                cryRpt.SetParameterValue("@DateTo", DatePart(DateInterval.Year, dtpTo.Value))
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf y.Checked = True And rbFTR.Checked = True And rbCluster.Checked = True Then
                GetNameBranch()
                getMonth()
                '  pickyear = 1
                '  showyearonly()
                GetName = "YEARLY TECH RET REPORT SUMMARY"
                'dt = db.ExecuteSQLQuery("SELECT SSINFOTERMCONSTAT.BRANCH_IP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,DATESTAMP,SSINFOTERMCLSTR.CLSTR_NM,ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0) as [ONLINE],case when(24 - ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0)) < 0 then 0 else (24 - ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0)) end as [OFFLINE] from SSINFOTERMCONSTAT inner join SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSINFOTERMCONSTAT.BRANCH_IP INNER JOIN SSINFOTERMBR on SSINFOTERMBR.BRANCH_CD = SSINFOTERMCONSTAT.BRANCH_CD INNER JOIN SSINFOTERMCLSTR on SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMCONSTAT.CLSTR where DATESTAMP between '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' GROUP by SSINFOTERMCONSTAT.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,DATESTAMP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMCLSTR.CLSTR_NM")
                Dim rpt As New FIGTRALLYEARLY
                cryRpt = rpt
                '  cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", DatePart(DateInterval.Year, dtpFrom.Value))
                cryRpt.SetParameterValue("@DateTo", DatePart(DateInterval.Year, dtpTo.Value))
                cryRpt.SetParameterValue("@getName", getBranch)

            ElseIf rbFTR.Checked = True And y.Checked = True And rbBranch.Checked = True Then
                GetNameBranch()
                getMonth()
                '  pickyear = 1s
                '  showyearonly()
                GetName = "YEARLY TECH RET REPORT SUMMARY"
                'dt = db.ExecuteSQLQuery("SELECT SSINFOTERMCONSTAT.BRANCH_IP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,DATESTAMP,SSINFOTERMCLSTR.CLSTR_NM,ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0) as [ONLINE],case when(24 - ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0)) < 0 then 0 else (24 - ISNULL(SUM(case when ONLINE_TME <> '' then cast(round(cast(MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0)) end as [OFFLINE] from SSINFOTERMCONSTAT inner join SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSINFOTERMCONSTAT.BRANCH_IP INNER JOIN SSINFOTERMBR on SSINFOTERMBR.BRANCH_CD = SSINFOTERMCONSTAT.BRANCH_CD INNER JOIN SSINFOTERMCLSTR on SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMCONSTAT.CLSTR where DATESTAMP between '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' GROUP by SSINFOTERMCONSTAT.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,DATESTAMP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMCLSTR.CLSTR_NM")
                Dim rpt As New FIGTRALLYEARLY
                cryRpt = rpt
                '  cryRpt.Refresh()
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", DatePart(DateInterval.Year, dtpFrom.Value))
                cryRpt.SetParameterValue("@DateTo", DatePart(DateInterval.Year, dtpTo.Value))
                cryRpt.SetParameterValue("@getName", getBranch)

#Region "PIN CHANGED"
            ElseIf rdPinChange.Checked = True And rbCluster.Checked = True Then
                GetNameBranch()
                GetName = "PIN CHANGE"
                'dt = db.ExecuteSQLQuery("SELECT convert(varchar,SSTRANSAT.ENCODE_DT,101) as DateOfTrans,case when SSINFOTERMPROCCD.PROC_NM = 'SSID CLEARANCE' then 'SSS/UMID CARD INFO' else SSINFOTERMPROCCD.PROC_NM end as 'PROC',SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN, COUNT(SSINFOTERMPROCCD.PROC_NM) as [count] FROM SSTRANSAT INNER JOIN SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAT.BRANCH_IP INNER JOIN SSINFOTERMPROCCD ON SSINFOTERMPROCCD.PROC_CD = SSTRANSAT.PROC_CD where SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and SSTRANSAT.PROC_CD not like '%10001%' and SSTRANSAT.PROC_CD not like '%10002%' and SSTRANSAT.PROC_CD not like '%10003%' and SSTRANSAT.PROC_CD not like '%10004%' and SSTRANSAT.PROC_CD not like '%10005%' and SSTRANSAT.PROC_CD not like '%10006%' and SSTRANSAT.PROC_CD not like '%10007%' and SSTRANSAT.PROC_CD not like '%10008%' and SSTRANSAT.PROC_CD not like '%10009%' and SSTRANSAT.PROC_CD not like '%10010%' and SSTRANSAT.PROC_CD not like '%10011%' and SSTRANSAT.PROC_CD not like '%10012%' and SSTRANSAT.PROC_CD not like '%10013%' and SSTRANSAT.PROC_CD not like '%10014%' and SSTRANSAT.PROC_CD not like '%10015%' and SSTRANSAT.PROC_CD not like '%10016%' and SSTRANSAT.PROC_CD not like '%10017%' and SSTRANSAT.PROC_CD not like '%10018%' and SSTRANSAT.PROC_CD not like '%10019%' and SSTRANSAT.PROC_CD not like '%10020%' and SSTRANSAT.PROC_CD not like '%10021%' and SSTRANSAT.PROC_CD not like '%10022%' and SSTRANSAT.PROC_CD not like '%10023%' and SSTRANSAT.PROC_CD not like '%10024%' and SSTRANSAT.PROC_CD not like '%10025%' and SSTRANSAT.PROC_CD not like '%10026%' and SSTRANSAT.PROC_CD not like '%10027%' and SSTRANSAT.PROC_CD not like '%10028%' and SSTRANSAT.PROC_CD not like '%10029%' and SSTRANSAT.PROC_CD not like '%10030%' and SSTRANSAT.PROC_CD not like '%10031%' and SSTRANSAT.PROC_CD not like '%10032%' and SSTRANSAT.PROC_CD not like '%10033%' and SSTRANSAT.PROC_CD not like '%10034%' and SSTRANSAT.PROC_CD not like '%10035%' and SSTRANSAT.PROC_CD not like '%10036%' and SSTRANSAT.PROC_CD not like '%10041%' and SSTRANSAT.PROC_CD not like '%10042%' GROUP by SSTRANSAT.ENCODE_DT,SSINFOTERMPROCCD.PROC_NM,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN")
                Dim rpt As New PINCHANGE
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf rdPinChange.Checked = True And rbGroup.Checked = True Then
                GetNameBranch()
                GetName = "PIN CHANGE"
                'dt = db.ExecuteSQLQuery("SELECT convert(varchar,SSTRANSAT.ENCODE_DT,101) as DateOfTrans,case when SSINFOTERMPROCCD.PROC_NM = 'SSID CLEARANCE' then 'SSS/UMID CARD INFO' else SSINFOTERMPROCCD.PROC_NM end as 'PROC',SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN, COUNT(SSINFOTERMPROCCD.PROC_NM) as [count] FROM SSTRANSAT INNER JOIN SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAT.BRANCH_IP INNER JOIN SSINFOTERMPROCCD ON SSINFOTERMPROCCD.PROC_CD = SSTRANSAT.PROC_CD where SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and SSTRANSAT.PROC_CD not like '%10001%' and SSTRANSAT.PROC_CD not like '%10002%' and SSTRANSAT.PROC_CD not like '%10003%' and SSTRANSAT.PROC_CD not like '%10004%' and SSTRANSAT.PROC_CD not like '%10005%' and SSTRANSAT.PROC_CD not like '%10006%' and SSTRANSAT.PROC_CD not like '%10007%' and SSTRANSAT.PROC_CD not like '%10008%' and SSTRANSAT.PROC_CD not like '%10009%' and SSTRANSAT.PROC_CD not like '%10010%' and SSTRANSAT.PROC_CD not like '%10011%' and SSTRANSAT.PROC_CD not like '%10012%' and SSTRANSAT.PROC_CD not like '%10013%' and SSTRANSAT.PROC_CD not like '%10014%' and SSTRANSAT.PROC_CD not like '%10015%' and SSTRANSAT.PROC_CD not like '%10016%' and SSTRANSAT.PROC_CD not like '%10017%' and SSTRANSAT.PROC_CD not like '%10018%' and SSTRANSAT.PROC_CD not like '%10019%' and SSTRANSAT.PROC_CD not like '%10020%' and SSTRANSAT.PROC_CD not like '%10021%' and SSTRANSAT.PROC_CD not like '%10022%' and SSTRANSAT.PROC_CD not like '%10023%' and SSTRANSAT.PROC_CD not like '%10024%' and SSTRANSAT.PROC_CD not like '%10025%' and SSTRANSAT.PROC_CD not like '%10026%' and SSTRANSAT.PROC_CD not like '%10027%' and SSTRANSAT.PROC_CD not like '%10028%' and SSTRANSAT.PROC_CD not like '%10029%' and SSTRANSAT.PROC_CD not like '%10030%' and SSTRANSAT.PROC_CD not like '%10031%' and SSTRANSAT.PROC_CD not like '%10032%' and SSTRANSAT.PROC_CD not like '%10033%' and SSTRANSAT.PROC_CD not like '%10034%' and SSTRANSAT.PROC_CD not like '%10035%' and SSTRANSAT.PROC_CD not like '%10036%' and SSTRANSAT.PROC_CD not like '%10041%' and SSTRANSAT.PROC_CD not like '%10042%' GROUP by SSTRANSAT.ENCODE_DT,SSINFOTERMPROCCD.PROC_NM,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN")
                Dim rpt As New PINCHANGE
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf rdPinChange.Checked = True And rbBranch.Checked = True Then
                GetNameBranch()
                GetName = "PIN CHANGE"
                'dt = db.ExecuteSQLQuery("SELECT convert(varchar,SSTRANSAT.ENCODE_DT,101) as DateOfTrans,case when SSINFOTERMPROCCD.PROC_NM = 'SSID CLEARANCE' then 'SSS/UMID CARD INFO' else SSINFOTERMPROCCD.PROC_NM end as 'PROC',SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN, COUNT(SSINFOTERMPROCCD.PROC_NM) as [count] FROM SSTRANSAT INNER JOIN SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAT.BRANCH_IP INNER JOIN SSINFOTERMPROCCD ON SSINFOTERMPROCCD.PROC_CD = SSTRANSAT.PROC_CD where SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and SSTRANSAT.PROC_CD not like '%10001%' and SSTRANSAT.PROC_CD not like '%10002%' and SSTRANSAT.PROC_CD not like '%10003%' and SSTRANSAT.PROC_CD not like '%10004%' and SSTRANSAT.PROC_CD not like '%10005%' and SSTRANSAT.PROC_CD not like '%10006%' and SSTRANSAT.PROC_CD not like '%10007%' and SSTRANSAT.PROC_CD not like '%10008%' and SSTRANSAT.PROC_CD not like '%10009%' and SSTRANSAT.PROC_CD not like '%10010%' and SSTRANSAT.PROC_CD not like '%10011%' and SSTRANSAT.PROC_CD not like '%10012%' and SSTRANSAT.PROC_CD not like '%10013%' and SSTRANSAT.PROC_CD not like '%10014%' and SSTRANSAT.PROC_CD not like '%10015%' and SSTRANSAT.PROC_CD not like '%10016%' and SSTRANSAT.PROC_CD not like '%10017%' and SSTRANSAT.PROC_CD not like '%10018%' and SSTRANSAT.PROC_CD not like '%10019%' and SSTRANSAT.PROC_CD not like '%10020%' and SSTRANSAT.PROC_CD not like '%10021%' and SSTRANSAT.PROC_CD not like '%10022%' and SSTRANSAT.PROC_CD not like '%10023%' and SSTRANSAT.PROC_CD not like '%10024%' and SSTRANSAT.PROC_CD not like '%10025%' and SSTRANSAT.PROC_CD not like '%10026%' and SSTRANSAT.PROC_CD not like '%10027%' and SSTRANSAT.PROC_CD not like '%10028%' and SSTRANSAT.PROC_CD not like '%10029%' and SSTRANSAT.PROC_CD not like '%10030%' and SSTRANSAT.PROC_CD not like '%10031%' and SSTRANSAT.PROC_CD not like '%10032%' and SSTRANSAT.PROC_CD not like '%10033%' and SSTRANSAT.PROC_CD not like '%10034%' and SSTRANSAT.PROC_CD not like '%10035%' and SSTRANSAT.PROC_CD not like '%10036%' and SSTRANSAT.PROC_CD not like '%10041%' and SSTRANSAT.PROC_CD not like '%10042%' GROUP by SSTRANSAT.ENCODE_DT,SSINFOTERMPROCCD.PROC_NM,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN")
                Dim rpt As New PINCHANGE
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf rdPinChange.Checked = True And rbAll.Checked = True Then
                If day1 = "1" And day2 = "31" Or day1 = "1" And day2 = "30" Then
                    getMonth()
                    GetNameBranch()
                    GetName = "PIN CHANGE"
                    'dt = db.ExecuteSQLQuery("SELECT SSTRANSAT.ENCODE_DT,SSINFOTERMPROCCD.PROC_NM,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN, COUNT(SSINFOTERMPROCCD.PROC_NM) as count FROM SSTRANSAT INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAT.BRANCH_IP INNER JOIN SSINFOTERMPROCCD ON SSINFOTERMPROCCD.PROC_CD = SSTRANSAT.PROC_CD where SSTRANSAT.ENCODE_DT BETWEEN '" & date12 & "' and '" & date22 & "' and SSINFOTERMPROCCD.PROC_NM not LIKE 'WEBSITE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MY.SSS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOGIN%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%REGISTRATION%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBERSHIP%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOANS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFITS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CORPORATE PROFILE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%PUBLICATIONS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%OTHER SERVICES%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%HOME%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - SS SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - EC SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS - SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%EC - SICKNESS%'  and SSINFOTERMPROCCD.PROC_NM not LIKE '%CHECKLIST OF DOCUMENTS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER MATERNITY NOTIFICATION%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%ANNUAL CONFIRMATION OF PENSIONER%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%ONLINE INQUIRY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SALARY LOAN%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%DDR FUNERAL%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SICKNESS/MATERNITY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%PENSION MAINTENANCE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%KIOSK FEEDBACK%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%WEBSITE FEEDBACK%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOAN ELIGIBILITY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%ACTUAL PREMIUMS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%EMPLOYMENT HISTORY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%FLEXI-FUND%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOAN STATUS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MATERNITY CLAIMS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER DETAILS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SICKNESS CLAIMS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%TECHNICAL RETIREMENT%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSID CLEARANCE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFIT CLAIMS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - MATERNITY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CITIZENS CHARTER%' GROUP by SSTRANSAT.ENCODE_DT,SSINFOTERMPROCCD.PROC_NM,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN,SSTRANSAT.BRANCH_IP")
                    Dim rpt As New PINCHANGE
                    cryRpt = rpt
                    '    cryRpt.Refresh()
                    cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                    rptView.ReportSource = cryRpt
                    cryRpt.SetParameterValue("@DateFrom", date12)
                    cryRpt.SetParameterValue("@DateTo", date22)
                    cryRpt.SetParameterValue("@getName", getBranch)
                Else
                    GetNameBranch()
                    GetName = "PIN CHANGE"
                    'dt = db.ExecuteSQLQuery("SELECT convert(varchar,SSTRANSAT.ENCODE_DT,101) as DateOfTrans ,SSINFOTERMPROCCD.PROC_NM,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_IP,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN, COUNT(SSINFOTERMPROCCD.PROC_NM) as count FROM SSTRANSAT INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.BRANCH_IP = SSTRANSAT.BRANCH_IP INNER JOIN SSINFOTERMPROCCD ON SSINFOTERMPROCCD.PROC_CD = SSTRANSAT.PROC_CD where SSTRANSAT.ENCODE_DT BETWEEN '" & date1.ToShortDateString & "' and '" & date1.ToShortDateString & "' and SSINFOTERMPROCCD.PROC_NM not LIKE 'WEBSITE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MY.SSS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOGIN%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%REGISTRATION%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBERSHIP%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOANS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFITS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CORPORATE PROFILE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%PUBLICATIONS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%OTHER SERVICES%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%HOME%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - SS SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SEPARATED - EC SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SS - SICKNESS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%EC - SICKNESS%'  and SSINFOTERMPROCCD.PROC_NM not LIKE '%CHECKLIST OF DOCUMENTS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER MATERNITY NOTIFICATION%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%ANNUAL CONFIRMATION OF PENSIONER%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%ONLINE INQUIRY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SALARY LOAN%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%DDR FUNERAL%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SICKNESS/MATERNITY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%PENSION MAINTENANCE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%KIOSK FEEDBACK%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%WEBSITE FEEDBACK%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSS WEBSITE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOAN ELIGIBILITY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%ACTUAL PREMIUMS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%EMPLOYMENT HISTORY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%FLEXI-FUND%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%LOAN STATUS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MATERNITY CLAIMS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%MEMBER DETAILS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SICKNESS CLAIMS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%TECHNICAL RETIREMENT%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SSID CLEARANCE%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%BENEFIT CLAIMS%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%SELF-EMPLOYED/VOLUNTARY - MATERNITY%' and SSINFOTERMPROCCD.PROC_NM not LIKE '%CITIZENS CHARTER%' GROUP by SSTRANSAT.ENCODE_DT,SSINFOTERMPROCCD.PROC_NM,SSINFOTERMKIOSK.KIOSK_ID,SSTRANSAT.BRANCH_CD,SSTRANSAT.CLSTR,SSTRANSAT.DIVSN,SSTRANSAT.BRANCH_IP")
                    Dim rpt As New PINCHANGE
                    cryRpt = rpt
                    '    cryRpt.Refresh()
                    cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                    rptView.ReportSource = cryRpt
                    cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                    cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                    cryRpt.SetParameterValue("@getName", getBranch)
                End If
#End Region

#Region "USER FEEDBACK"
            ElseIf rdUserFeedBack.Checked = True And rbGroup.Checked = True Then
                GetNameBranch()
                GetName = "User Feedback"
                'dt = db.ExecuteSQLQuery("SELECT NAME,EMAIL,ADDRESS_TYP,POST_CD, SSRATE1,SSRATE2,SSRATE3,SSRATE4,SSRATE5,SSRATE6,SSRATE7,VST_TAG,REASN_TAG,INFO_TAG,COMMNT_TAG,SSTRANSINFOTERMFB.ENCODE_DT,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSINFOTERMFB.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP FROM SSTRANSINFOTERMFB INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID = SSTRANSINFOTERMFB.KIOSK_ID INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSINFOTERMFB.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSTRANSINFOTERMFB.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSTRANSINFOTERMFB.DIVSN WHERE cast(SSTRANSINFOTERMFB.ENCODE_DT as DATE) BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "'")
                Dim rpt As New FEEDBACKUSER
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf rdUserFeedBack.Checked = True And rbCluster.Checked = True Then
                GetNameBranch()
                GetName = "User Feedback"
                'dt = db.ExecuteSQLQuery("SELECT NAME,EMAIL,ADDRESS_TYP,POST_CD, SSRATE1,SSRATE2,SSRATE3,SSRATE4,SSRATE5,SSRATE6,SSRATE7,VST_TAG,REASN_TAG,INFO_TAG,COMMNT_TAG,SSTRANSINFOTERMFB.ENCODE_DT,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSINFOTERMFB.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP FROM SSTRANSINFOTERMFB INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID = SSTRANSINFOTERMFB.KIOSK_ID INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSINFOTERMFB.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSTRANSINFOTERMFB.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSTRANSINFOTERMFB.DIVSN WHERE cast(SSTRANSINFOTERMFB.ENCODE_DT as DATE) BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "'")
                Dim rpt As New FEEDBACKUSER
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf rdUserFeedBack.Checked = True And rbBranch.Checked = True Then
                GetNameBranch()
                GetName = "User Feedback"
                'dt = db.ExecuteSQLQuery("SELECT NAME,EMAIL,ADDRESS_TYP,POST_CD, SSRATE1,SSRATE2,SSRATE3,SSRATE4,SSRATE5,SSRATE6,SSRATE7,VST_TAG,REASN_TAG,INFO_TAG,COMMNT_TAG,SSTRANSINFOTERMFB.ENCODE_DT,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSINFOTERMFB.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP FROM SSTRANSINFOTERMFB INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID = SSTRANSINFOTERMFB.KIOSK_ID INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSINFOTERMFB.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSTRANSINFOTERMFB.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSTRANSINFOTERMFB.DIVSN WHERE cast(SSTRANSINFOTERMFB.ENCODE_DT as DATE) BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "'")
                Dim rpt As New FEEDBACKUSER
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf rdUserFeedBack.Checked = True And rbAll.Checked = True Then
                GetNameBranch()
                GetName = "User Feedback"
                'dt = db.ExecuteSQLQuery("SELECT NAME,EMAIL,ADDRESS_TYP,POST_CD, SSRATE1,SSRATE2,SSRATE3,SSRATE4,SSRATE5,SSRATE6,SSRATE7,VST_TAG,REASN_TAG,INFO_TAG,COMMNT_TAG,SSTRANSINFOTERMFB.ENCODE_DT,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSINFOTERMFB.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP FROM SSTRANSINFOTERMFB INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID = SSTRANSINFOTERMFB.KIOSK_ID INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSINFOTERMFB.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSTRANSINFOTERMFB.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSTRANSINFOTERMFB.DIVSN WHERE cast(SSTRANSINFOTERMFB.ENCODE_DT as DATE) BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "'")
                Dim rpt As New FEEDBACKUSER
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)

#End Region

#Region "EXIT SURVEY"
            ElseIf rdExitSurvey.Checked = True And rbGroup.Checked = True Then
                GetNameBranch()
                GetName = "Exit Survery"
                'dt = db.ExecuteSQLQuery("SELECT NAME,EMAIL,ADDRESS_TYP,POST_CD, SSRATE1,SSRATE2,SSRATE3,SSRATE4,SSRATE5,SSRATE6,SSRATE7,VST_TAG,REASN_TAG,INFO_TAG,COMMNT_TAG,SSTRANSINFOTERMFB.ENCODE_DT,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSINFOTERMFB.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP FROM SSTRANSINFOTERMFB INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID = SSTRANSINFOTERMFB.KIOSK_ID INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSINFOTERMFB.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSTRANSINFOTERMFB.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSTRANSINFOTERMFB.DIVSN WHERE cast(SSTRANSINFOTERMFB.ENCODE_DT as DATE) BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "'")
                Dim rpt As New EXITSURVEY
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf rdExitSurvey.Checked = True And rbCluster.Checked = True Then
                GetNameBranch()
                GetName = "Exit Survery"
                'dt = db.ExecuteSQLQuery("SELECT NAME,EMAIL,ADDRESS_TYP,POST_CD, SSRATE1,SSRATE2,SSRATE3,SSRATE4,SSRATE5,SSRATE6,SSRATE7,VST_TAG,REASN_TAG,INFO_TAG,COMMNT_TAG,SSTRANSINFOTERMFB.ENCODE_DT,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSINFOTERMFB.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP FROM SSTRANSINFOTERMFB INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID = SSTRANSINFOTERMFB.KIOSK_ID INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSINFOTERMFB.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSTRANSINFOTERMFB.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSTRANSINFOTERMFB.DIVSN WHERE cast(SSTRANSINFOTERMFB.ENCODE_DT as DATE) BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "'")
                Dim rpt As New EXITSURVEY
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf rdExitSurvey.Checked = True And rbBranch.Checked = True Then
                GetNameBranch()
                GetName = "Exit Survery"
                'dt = db.ExecuteSQLQuery("SELECT NAME,EMAIL,ADDRESS_TYP,POST_CD, SSRATE1,SSRATE2,SSRATE3,SSRATE4,SSRATE5,SSRATE6,SSRATE7,VST_TAG,REASN_TAG,INFO_TAG,COMMNT_TAG,SSTRANSINFOTERMFB.ENCODE_DT,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSINFOTERMFB.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP FROM SSTRANSINFOTERMFB INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID = SSTRANSINFOTERMFB.KIOSK_ID INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSINFOTERMFB.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSTRANSINFOTERMFB.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSTRANSINFOTERMFB.DIVSN WHERE cast(SSTRANSINFOTERMFB.ENCODE_DT as DATE) BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "'")
                Dim rpt As New EXITSURVEY
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
            ElseIf rdExitSurvey.Checked = True And rbAll.Checked = True Then
                GetNameBranch()
                GetName = "Exit Survery"
                'dt = db.ExecuteSQLQuery("SELECT NAME,EMAIL,ADDRESS_TYP,POST_CD, SSRATE1,SSRATE2,SSRATE3,SSRATE4,SSRATE5,SSRATE6,SSRATE7,VST_TAG,REASN_TAG,INFO_TAG,COMMNT_TAG,SSTRANSINFOTERMFB.ENCODE_DT,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,SSTRANSINFOTERMFB.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP FROM SSTRANSINFOTERMFB INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.KIOSK_ID = SSTRANSINFOTERMFB.KIOSK_ID INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSINFOTERMFB.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSTRANSINFOTERMFB.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSTRANSINFOTERMFB.DIVSN WHERE cast(SSTRANSINFOTERMFB.ENCODE_DT as DATE) BETWEEN '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "'")
                Dim rpt As New EXITSURVEY
                cryRpt = rpt
                cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                rptView.ReportSource = cryRpt
                cryRpt.SetParameterValue("@DateFrom", date1.ToShortDateString)
                cryRpt.SetParameterValue("@DateTo", date2.ToShortDateString)
                cryRpt.SetParameterValue("@getName", getBranch)
#End Region
            End If

            'Filter.Enabled = True
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
#End Region
    Private Sub ButtonX1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonX1.Click
        _frmExportReport.int1 = 0
        Dim reportload As String = Convert.ToString(rptView.ReportSource)
        If reportload = "" Then
            MsgBox("Please generate a report before exporting it to a different format.", MsgBoxStyle.Information)
        Else
            _frmBlock.Show()
            _frmExportReport.ShowDialog()
        End If
    End Sub
    Private Sub _frmViewReports_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        cbBranch.Text = "ABU DHABI"
        cbCluster.Text = "NCR CENTRAL"
        uncheckRB()
        fillBranch()
        fillCluster()
        fillGroup()
        FillGraphFilter()
        GroupBox2.Visible = False
        GroupBox3.Visible = False
        cbCard.Enabled = False
        cbBranch.Enabled = False
        cbCluster.Enabled = False

    End Sub
    Private Sub rbWeekly_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbWeekly.CheckedChanged
        pickyear = 0
        DR = 1
        showyearonly()
    End Sub
    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        uncheckRB()
        RBCONTROL()
        Removealltxt()
        cbBranch.Enabled = False
        cbCluster.Enabled = False
        cbSalloan.Enabled = False
        dtpFrom.MinDate = New DateTime(1753, 1, 1)
        dtpFrom.MaxDate = New DateTime(9998, 1, 1)
        dtpFrom.Format = DateTimePickerFormat.Short
        dtpFrom.Value = Today
        dtpFrom.ShowUpDown = False
        dtpTo.MinDate = New DateTime(1753, 1, 1)
        dtpTo.MaxDate = New DateTime(9998, 1, 1)
        dtpTo.Value = Today
        dtpTo.Format = DateTimePickerFormat.Short
        dtpTo.ShowUpDown = False
        Select Case TabControl1.SelectedTab.Text

            Case "SET Update Logs"
                rbCluster.Visible = True
                cbCluster.Visible = True
                rbGroup.Visible = True
                cbGroup.Visible = True
            Case "SSID Usage Report"
                cbCard.Enabled = True
                GroupBox2.Visible = True
            Case "Transaction Report"
                GroupBox2.Visible = False
                GroupBox3.Visible = True
            Case Else
                cbCard.Enabled = False
                rbBranch.Visible = True
                rbBranch.Enabled = True
                cbBranch.Visible = True
                cbBranch.Enabled = False
                rbCluster.Visible = True
                rbCluster.Enabled = True
                rbGroup.Visible = True
                cbGroup.Visible = True
                cbGroup.Enabled = False
                cbCluster.Visible = True
                cbCluster.Enabled = False
                GroupBox2.Visible = False
                GroupBox3.Visible = False
        End Select
    End Sub

    Private Sub rbAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbAll.CheckedChanged
        Group = 1
        Select Case rbAll.Checked

            Case True
                cbCard.Items.Clear()
                cbCard.Items.Add("UMID CARD")
                cbCard.Items.Add("SSS CARD")
                cbCard.Items.Add("Show All")
                cbBranch.Enabled = False
                cbCluster.Enabled = False
                cbBranch.Text = Nothing
                cbCluster.Text = Nothing
                cbGroup.Text = Nothing
            Case False
                cbBranch.Enabled = True
                cbCluster.Enabled = True
                cbGroup.Enabled = True
        End Select
    End Sub

    Private Sub rbCluster_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbCluster.CheckedChanged
        Group = 1
        Select Case rbCluster.Checked

            Case True
                cbCard.Items.Clear()
                cbCard.Items.Add("UMID CARD")
                cbCard.Items.Add("SSS CARD")
                cbCard.Items.Add("Show All")
                cbCluster.Enabled = True
                cbGroup.Enabled = False
                cbBranch.Enabled = False
                cbBranch.Text = Nothing
                cbGroup.Text = Nothing
            Case False
                cbCluster.Enabled = False
        End Select
    End Sub
    Private Sub rbBranch_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbBranch.CheckedChanged
        Group = 1
        Select Case rbBranch.Checked
            Case True
                cbCard.Items.Clear()
                cbCard.Items.Add("UMID CARD")
                cbCard.Items.Add("SSS CARD")
                cbCard.Items.Add("Show All")
                cbBranch.Enabled = True
                cbCluster.Enabled = False
                cbGroup.Enabled = False
                cbCluster.Text = Nothing
                cbGroup.Text = Nothing
            Case False
                cbBranch.Enabled = False
        End Select
    End Sub
#Region "Transactions Text"
    Private Sub rbQueryReports_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbQueryReports.CheckedChanged
        OLAP = 1
    End Sub
    Private Sub rbOnlineInquiry_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbOnlineInquiry.CheckedChanged
        OLAP = 1
    End Sub

    Private Sub rbDownTime_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbDownTime.CheckedChanged
        Trans = 1
    End Sub

    Private Sub rbMontoring_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbMonitoring.CheckedChanged
        Trans = 1
    End Sub

    Private Sub rbSummarySSS_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbSummarySSS.CheckedChanged
        Trans = 1
    End Sub

    Private Sub rbWebLog_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Trans = 1
    End Sub

    Private Sub rbDaily_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbDaily.CheckedChanged
        pickyear = 0
        DR = 1
        showyearonly()
    End Sub

    Private Sub rbMonthly_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbMonthly.CheckedChanged
        pickyear = 2
        DR = 1
        showyearonly()
    End Sub

    Private Sub rbyearly_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbYearly.CheckedChanged
        DR = 1
        pickyear = 1
        showyearonly()
    End Sub
    Private Sub rbEligibility_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbEligibility.CheckedChanged
        OLAP = 1
    End Sub

    Private Sub rbDDR_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbDDR.CheckedChanged
        OLAP = 1
    End Sub

    Private Sub rbWebAccess_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Trans = 1

    End Sub

    Private Sub rbFBK_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbFBK.CheckedChanged
        Trans = 1

    End Sub
    Private Sub cbSTrans_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbSTrans.SelectedIndexChanged
        Select Case cbSTrans.Text
            Case "Salary Loan"
                cbSalloan.Enabled = True
                Trans = 1
                rbCluster.Visible = True
                cbCluster.Visible = True
            Case "Maternity Notification"
                Trans = 1
                cbSalloan.Text = Nothing
                cbSalloan.Enabled = False
            Case "Technical Retirement Pensioner"
                Trans = 1
                rbCluster.Visible = True
                cbCluster.Visible = True
                cbSalloan.Text = Nothing
                cbSalloan.Enabled = False
            Case "Technical Retirement Lump Sum"
                Trans = 1
                cbSalloan.Text = Nothing
                cbSalloan.Enabled = False
                rbCluster.Visible = True
                cbCluster.Visible = True
            Case "Status updates of retiree pensioner dependents"
                Trans = 1
                rbCluster.Visible = True
                cbCluster.Visible = True
                cbSalloan.Text = Nothing
                cbSalloan.Enabled = False
            Case "ACOP"
                Trans = 1
                cbSalloan.Text = Nothing
                cbSalloan.Enabled = False
                rbCluster.Visible = True
                cbCluster.Visible = True
            Case "Updating Contact Information"
                Trans = 1
                cbSalloan.Text = Nothing
                cbSalloan.Enabled = False
                rbCluster.Visible = True
                cbCluster.Visible = True
            Case "Simplified Web Registration"
                Trans = 1
                cbSalloan.Text = Nothing
                cbSalloan.Enabled = False
                rbCluster.Visible = True
                cbCluster.Visible = True
        End Select

    End Sub
#End Region

    Private Sub rbSuccessTrans_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbSuccessTrans.CheckedChanged
        Trans = 1
        Select Case rbSuccessTrans.Checked
            Case True
                rbCluster.Checked = False
                cbSTrans.Enabled = True
                cbFTrans.Text = Nothing
                cbFTrans.Enabled = False
                cbSalloan.Text = Nothing
            Case False
                cbSTrans.Enabled = False
                cbFTrans.Enabled = True
        End Select
    End Sub

    Private Sub rbfailedTrans_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbfailedTrans.CheckedChanged
        Trans = 1
        Select Case rbfailedTrans.Checked
            Case True
                If rbCluster.Visible = False Then
                    rbCluster.Visible = True
                    rbCluster.Checked = False
                End If
                If cbCluster.Visible = False Then
                    cbCluster.Visible = True
                    rbCluster.Checked = False
                End If
                cbFTrans.Enabled = True
                cbSTrans.Text = Nothing
                cbSalloan.Text = Nothing
                cbSTrans.Enabled = False
                cbSalloan.Enabled = False
            Case False
                cbFTrans.Enabled = False
                cbSTrans.Enabled = True
        End Select
    End Sub

    Private Sub cbSalloan_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbSalloan.SelectedIndexChanged
        Trans = 1
    End Sub

    Private Sub cbFTrans_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbFTrans.SelectedIndexChanged

    End Sub

    Private Sub ButtonX2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonX2.Click
        Dim reportload As String = Convert.ToString(rptView.ReportSource)

        If reportload = "" Then
            MsgBox("Please generate a report before print", MsgBoxStyle.Information)
        Else
            ReportPrint = 0
            _frmBlock.Show()
            _frmPrinter.ShowDialog()
        End If
    End Sub

    Private Sub rbCharter_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbCharter.CheckedChanged
        OLAP = 1

    End Sub


    Private Sub rbFBW_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbFBW.CheckedChanged
        Trans = 1

    End Sub

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbUpdates.CheckedChanged
        Trans = 1

    End Sub

    Private Sub rbSMAT_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbSMAT.CheckedChanged
        OLAP = 1

    End Sub

    Private Sub fbk_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles fbk.CheckedChanged
        OLAP = 1

    End Sub

    Private Sub rbfb_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbfb.CheckedChanged
        OLAP = 1

    End Sub

    Private Sub rbSWR_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbSWR.CheckedChanged
        OLAP = 1

    End Sub

    Private Sub rbSuc_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbSuc.CheckedChanged
        OLAP = 1

    End Sub

    Private Sub rbFT_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbFT.CheckedChanged
        OLAP = 1

    End Sub

    Private Sub rbacop_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbacop.CheckedChanged
        OLAP = 1

    End Sub

    Private Sub rbacopdep_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbacopdep.CheckedChanged
        OLAP = 1

    End Sub

    Private Sub rbGroup_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbGroup.CheckedChanged
        Group = 1
        Select Case rbGroup.Checked

            Case True
                cbGroup.Enabled = True
                cbBranch.Enabled = False
                cbCluster.Enabled = False
                cbBranch.Text = Nothing
                cbCluster.Enabled = False
                cbCluster.Text = Nothing
            Case False
                cbGroup.Enabled = False

        End Select


    End Sub

    Private Sub rbTOP_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbTOP.CheckedChanged
        OLAP = 1

    End Sub

    Private Sub rbSucFail_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbSucFail.CheckedChanged
        OLAP = 1
    End Sub

    Private Sub rbFTR_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        OLAP = 1
    End Sub

    Private Sub rbUFAILED_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbUFAILED.CheckedChanged
        Trans = 1
        cbCard.Text = "Show All"
        If rbUFAILED.Checked = True Then
            cbCard.Enabled = False
        Else
            cbCard.Enabled = True
        End If
    End Sub

    Private Sub rbErrorLogs_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbErrorLogs.CheckedChanged
        Trans = 1
        cbCard.Text = "Show All"
        If rbErrorLogs.Checked = True Then
            cbCard.Enabled = False
        Else
            cbCard.Enabled = True
        End If
    End Sub

    Private Sub rbMonitoringAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbMonitoringAll.CheckedChanged
        Trans = 1
        If rbMonitoringAll.Checked = True Then
            rbAll.Checked = True
            rbBranch.Checked = False
            rbGroup.Checked = False
            rbCluster.Checked = False
            rbBranch.Enabled = False
            cbBranch.Enabled = False
            rbGroup.Enabled = False
            cbGroup.Enabled = False
            rbCluster.Enabled = False
            cbCluster.Enabled = False
        Else
            rbBranch.Enabled = True
            rbGroup.Enabled = True
            rbCluster.Enabled = True
        End If
    End Sub

    Private Sub rbFTR_CheckedChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbFTR.CheckedChanged
        Trans = 1
    End Sub

    Private Sub d_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles d.CheckedChanged
        pickyear = 0
        DR = 1
        showyearonly()
    End Sub

    Private Sub w_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles w.CheckedChanged
        pickyear = 0
        DR = 1
        showyearonly()
    End Sub

    Private Sub m_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles m.CheckedChanged
        ChangeDate()
    End Sub

    Private Sub y_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles y.CheckedChanged
        ChangeDate()
    End Sub

    Private Sub rbS_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles s.CheckedChanged
        DR = 1
        pickyear = 0
        showyearonly()
    End Sub

    Private Sub rbPRN_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbPRN.CheckedChanged
        OLAP = 1
    End Sub
    Private Sub TabPage6_Click(sender As Object, e As EventArgs) Handles TabPage6.Click
        uncheckRB()
    End Sub
End Class