Imports System.Threading
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports CrystalDecisions.Windows.Forms
Public Class _frmMonitoring
    Dim cnt As Integer
    Dim statusChange As Integer
    Dim db As New ConnectionString
    Public trd As Thread
    Dim accessLock As New Object
    Public i2 As Integer
    Dim endThread As Boolean = False
    Dim dt As DataTable
    Dim cryRpt As New ReportDocument
#Region "Listview Clicked event"
    Public Sub loadDetailslvPC()
        Try

            Dim ipadd As String = lvPC.SelectedItems(0).SubItems(1).Text
            lblKname.Text = db.putSingleValue("Select KIOSK_NM FROM SSINFOTERMKIOSK inner join SSINFOTERMBR on SSINFOTERMKIOSK.BRANCH_CD = SSINFOTERMBR.BRANCH_CD where BRANCH_IP = '" & ipadd & "'")
            lblKIP.Text = db.putSingleValue("Select BRANCH_IP FROM SSINFOTERMKIOSK inner join SSINFOTERMBR on SSINFOTERMKIOSK.BRANCH_CD = SSINFOTERMBR.BRANCH_CD where BRANCH_IP = '" & ipadd & "'")
            lblKbranch.Text = db.putSingleValue("Select SSINFOTERMBR.BRANCH_NM FROM SSINFOTERMKIOSK inner join SSINFOTERMBR on SSINFOTERMKIOSK.BRANCH_CD = SSINFOTERMBR.BRANCH_CD where BRANCH_IP = '" & ipadd & "'")
            lblStatus.Text = db.putSingleValue("Select DISTINCT STATUS FROM SSMONITORING where DATESTAMP = (select MAX(DATESTAMP) from SSMONITORING WHERE BRANCH_IP = '" & ipadd & "')")
            lblCluster.Text = db.putSingleValue("Select SSINFOTERMCLSTR.CLSTR_NM FROM SSINFOTERMKIOSK inner join SSINFOTERMBR on SSINFOTERMKIOSK.BRANCH_CD = SSINFOTERMBR.BRANCH_CD inner join SSINFOTERMCLSTR on SSINFOTERMKIOSK.CLSTR = SSINFOTERMCLSTR.CLSTR_CD where BRANCH_IP = '" & ipadd & "'")
            lblGroup.Text = db.putSingleValue("Select SSINFOTERMGROUP.GROUP_NM FROM SSINFOTERMKIOSK inner join SSINFOTERMBR on SSINFOTERMKIOSK.BRANCH_CD = SSINFOTERMBR.BRANCH_CD INNER JOIN SSINFOTERMGROUP on SSINFOTERMGROUP.GROUP_CD = SSINFOTERMKIOSK.DIVSN where BRANCH_IP = '" & ipadd & "'")
            LASTONLINE.Text = db.putSingleValue("Select max(ONLINE_DT) FROM SSMONITORING where BRANCH_IP = '" & ipadd & "'")
            LASTOFFLINE.Text = db.putSingleValue("Select max(OFFLINE_DT) FROM SSMONITORING where BRANCH_IP = '" & ipadd & "'")
            lbldate.Text = db.putSingleValue("Select ENCODE_DT from SSINFOTERMKIOSK where BRANCH_IP = '" & ipadd & "'")
            lblSerial.Text = db.putSingleValue("Select SERIALNUM FROM SSINFOTERMKIOSK WHERE BRANCH_IP = '" & ipadd & "'")
            lblTrans.Text = db.putSingleValue("SELECT COUNT(REF_NUM) from SSINFOTERMACCESS where KIOSK_ID = '" & lvPC.SelectedItems(0).SubItems(0).Text & "' and REF_NUM <>'' ")
            OpStart1.Text = db.putSingleValue("Select OP_START from SSINFOTERMKIOSK where BRANCH_IP = '" & ipadd & "'")
            OpEnd1.Text = db.putSingleValue("Select OP_END from SSINFOTERMKIOSK where BRANCH_IP = '" & ipadd & "'")

            DURATION.Text = "OFFLINE DURATION"

            Dim date1 As DateTime = _
        System.Convert.ToDateTime(LASTOFFLINE.Text)

            Dim date2 As DateTime = _
             System.Convert.ToDateTime(DateAndTime.Now)

            Dim ts As New TimeSpan()

            ts = date2.Subtract(date1)

            lblOD.Text = ts.Days & " Day(s) " & _
              ts.Hours & " Hr(s) " & ts.Minutes & " Min(s) " & ts.Seconds & _
            " Sec(s) Ago"
            LASTOFFLINE.ForeColor = Color.Black
            LASTOFFLINE.Visible = True
            Dim DateFrom As Date = dtpFrom.Value
            Dim DateTo As Date = dtpTo.Value
            dgvLogs.Visible = True
            db.FillDataGridView("select unpvt.value, case when unpvt.attribute = 'OFFLINE_DT' then '*** OFFLINE ***' else 'ONLINE' end as [STATUS] FROM SSMONITORING c UNPIVOT ( value for attribute in (ONLINE_DT,OFFLINE_DT) )unpvt INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.BRANCH_IP = unpvt.BRANCH_IP where cast(unpvt.[value] as date) BETWEEN '" & dtpFrom.Value & "' and '" & dtpTo.Value & "' and SSINFOTERMKIOSK.BRANCH_IP = '" & ipadd & "' UNION SELECT GETDATE() AS VALUE, 'ONLINE' AS STATUS  FROM SSINFOTERMKIOSK INNER JOIN SSMONITORING ON SSMONITORING.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP WHERE SSINFOTERMKIOSK.BRANCH_IP NOT IN(SELECT SSMONITORING.BRANCH_IP FROM SSMONITORING where cast(DATESTAMP as date) = '" & Today & "') and SSINFOTERMKIOSK.BRANCH_IP = '" & ipadd & "' ORDER BY unpvt.value", dgvLogs)

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
    Public Sub loadDetailslvList()
        Try

            Dim ipadd As String = lvList.SelectedItems(0).SubItems(1).Text
            lblKname.Text = db.putSingleValue("Select KIOSK_NM FROM SSINFOTERMKIOSK inner join SSINFOTERMBR on SSINFOTERMKIOSK.BRANCH_CD = SSINFOTERMBR.BRANCH_CD where BRANCH_IP = '" & ipadd & "'")
            lblKIP.Text = db.putSingleValue("Select BRANCH_IP FROM SSINFOTERMKIOSK inner join SSINFOTERMBR on SSINFOTERMKIOSK.BRANCH_CD = SSINFOTERMBR.BRANCH_CD where BRANCH_IP = '" & ipadd & "'")
            lblKbranch.Text = db.putSingleValue("Select SSINFOTERMBR.BRANCH_NM FROM SSINFOTERMKIOSK inner join SSINFOTERMBR on SSINFOTERMKIOSK.BRANCH_CD = SSINFOTERMBR.BRANCH_CD where BRANCH_IP = '" & ipadd & "'")
            lblStatus.Text = db.putSingleValue("Select DISTINCT STATUS FROM SSMONITORING where DATESTAMP = (select MAX(DATESTAMP) from SSMONITORING WHERE BRANCH_IP = '" & ipadd & "')")
            lblCluster.Text = db.putSingleValue("Select SSINFOTERMCLSTR.CLSTR_NM FROM SSINFOTERMKIOSK inner join SSINFOTERMBR on SSINFOTERMKIOSK.BRANCH_CD = SSINFOTERMBR.BRANCH_CD inner join SSINFOTERMCLSTR on SSINFOTERMKIOSK.CLSTR = SSINFOTERMCLSTR.CLSTR_CD where BRANCH_IP = '" & ipadd & "'")
            lblGroup.Text = db.putSingleValue("Select SSINFOTERMGROUP.GROUP_NM FROM SSINFOTERMKIOSK inner join SSINFOTERMBR on SSINFOTERMKIOSK.BRANCH_CD = SSINFOTERMBR.BRANCH_CD INNER JOIN SSINFOTERMGROUP on SSINFOTERMGROUP.GROUP_CD = SSINFOTERMKIOSK.DIVSN where BRANCH_IP = '" & ipadd & "'")
            LASTONLINE.Text = db.putSingleValue("Select max(ONLINE_DT) FROM SSMONITORING where BRANCH_IP = '" & ipadd & "'")
            LASTOFFLINE.Text = db.putSingleValue("Select max(OFFLINE_DT) FROM SSMONITORING where BRANCH_IP = '" & ipadd & "'")
            lbldate.Text = db.putSingleValue("Select ENCODE_DT from SSINFOTERMKIOSK where BRANCH_IP = '" & ipadd & "'")
            lblSerial.Text = db.putSingleValue("Select SERIALNUM FROM SSINFOTERMKIOSK WHERE BRANCH_IP = '" & ipadd & "'")
            lblTrans.Text = db.putSingleValue("SELECT COUNT(REF_NUM) from SSINFOTERMACCESS where KIOSK_ID = '" & lvList.SelectedItems(0).SubItems(0).Text & "' and REF_NUM <>''")
            OpStart1.Text = db.putSingleValue("Select OP_START from SSINFOTERMKIOSK where BRANCH_IP = '" & ipadd & "'")
            OpEnd1.Text = db.putSingleValue("Select OP_END from SSINFOTERMKIOSK where BRANCH_IP = '" & ipadd & "'")

            DURATION.Text = "ONLINE DURATION"

            If IsDate(LASTONLINE.Text) Then
                Dim date1 As DateTime =
         System.Convert.ToDateTime(LASTONLINE.Text)
                Dim date2 As DateTime =
             System.Convert.ToDateTime(DateAndTime.Now)

                Dim ts As New TimeSpan()

                ts = date2.Subtract(date1)

                lblOD.Text = ts.Days & " Day(s) " &
              ts.Hours & " Hr(s) " & ts.Minutes & " Min(s) " & ts.Seconds &
            " Sec(s) Ago"
                LASTONLINE.ForeColor = Color.Black
                LASTONLINE.Visible = True

                Dim DateFrom As Date = dtpFrom.Value
                Dim DateTo As Date = dtpTo.Value
                dgvLogs.Visible = True
                db.FillDataGridView("select unpvt.value, case when unpvt.attribute = 'OFFLINE_DT' then '*** OFFLINE ***' else 'ONLINE' end as [STATUS] FROM SSMONITORING c UNPIVOT ( value for attribute in (ONLINE_DT,OFFLINE_DT) )unpvt INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.BRANCH_IP = unpvt.BRANCH_IP where cast(unpvt.[value] as date) BETWEEN '" & dtpFrom.Value & "' and '" & dtpTo.Value & "' and SSINFOTERMKIOSK.BRANCH_IP = '" & ipadd & "' UNION SELECT GETDATE() AS VALUE, 'ONLINE' AS STATUS  FROM SSINFOTERMKIOSK INNER JOIN SSMONITORING ON SSMONITORING.BRANCH_IP = SSINFOTERMKIOSK.BRANCH_IP WHERE SSINFOTERMKIOSK.BRANCH_IP NOT IN(SELECT SSMONITORING.BRANCH_IP FROM SSMONITORING where cast(DATESTAMP as date) = '" & Today & "') and SSINFOTERMKIOSK.BRANCH_IP = '" & ipadd & "' ORDER BY unpvt.value", dgvLogs)
                Exit Sub
            End If

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
    Private Sub lvList_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lvList.Click
        Try
            loadDetailslvList()
            lvList.Size = New System.Drawing.Size(521, 99)
            lvPC.Size = New System.Drawing.Size(552, 99)
            btnBack.Visible = True
            dgvLogs.Visible = True
            gbKioskHistory.Visible = True
            lblFilter.Visible = True
            dtpFrom.Visible = True
            dtpTo.Visible = True
            Select Case lblStatus.Text

                Case "True"
                    lblStatus.Text = "OFFLINE"
                    lblStatus.ForeColor = Color.Red
                    lblLastOnline.Visible = True
                    LASTOFFLINE.Visible = True
                    lblOffline.Visible = True
                    LASTONLINE.Visible = False

                Case "False"
                    lblStatus.Text = "ONLINE"
                    lblStatus.ForeColor = Color.Green
                    lblLastOnline.Visible = False
                    LASTOFFLINE.Visible = False
                    lblOffline.Visible = False
                    LASTONLINE.Visible = True
            End Select
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
    Private Sub lvPC_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lvPC.Click
        Try
            loadDetailslvPC()
            lvList.Size = New System.Drawing.Size(521, 99)
            lvPC.Size = New System.Drawing.Size(552, 99)
            btnBack.Visible = True
            dgvLogs.Visible = True
            gbKioskHistory.Visible = True
            lblFilter.Visible = True
            dtpFrom.Visible = True
            dtpTo.Visible = True
            Select Case lblStatus.Text

                Case "True"
                    lblStatus.Text = "OFFLINE"
                    lblStatus.ForeColor = Color.Red
                    lblLastOnline.Visible = False
                    LASTOFFLINE.Visible = True
                    lblOffline.Visible = True
                    LASTONLINE.Visible = False

                Case "False"
                    lblStatus.Text = "ONLINE"
                    lblStatus.ForeColor = Color.Green
                    lblLastOnline.Visible = True
                    LASTOFFLINE.Visible = False
                    lblOffline.Visible = False
                    LASTONLINE.Visible = True
            End Select
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
#End Region
#Region "Thread Event"
    Private Sub _frmMonitoring_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            'dtpFrom.Value = Today
            'dtpTo.Value = Today
            GC.Collect()
            Control.CheckForIllegalCrossThreadCalls = False
            trd = New Thread(AddressOf ThreadTask)
            trd.IsBackground = True
            trd.Start()

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
    Private Sub ThreadTask()
        Do
            Try
                runTime()
                System.Threading.Thread.Sleep(1000 * 300)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        Loop
    End Sub
    Public Sub runTime()
        Button1.Enabled = False
        lvPC.Enabled = False
        lvList.Enabled = False
        loadBranches()
        lvPC.Enabled = True
        lvList.Enabled = True
        Button1.Enabled = True
    End Sub
#End Region
    Public Sub loadBranches()
        Try
            db.FillListView(db.ExecuteSQLQuery("SELECT DISTINCT SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP,SSINFOTERMBR.BRANCH_NM, TAG = 'ONLINE' FROM SSINFOTERMKIOSK INNER JOIN SSINFOTERMBR ON SSINFOTERMKIOSK.BRANCH_CD=SSINFOTERMBR.BRANCH_CD WHERE (SSINFOTERMKIOSK.BRANCH_IP NOT IN(SELECT BRANCH_IP FROM SSMONITORING) OR BRANCH_IP NOT IN(SELECT DISTINCT BRANCH_IP FROM SSMONITORING WHERE STATUS=1 AND CAST(DATESTAMP AS DATE) = '" & Today & "')) and isvpn = 0 ORDER BY SSINFOTERMBR.BRANCH_NM"), lvList)
            db.FillListView(db.ExecuteSQLQuery("SELECT DISTINCT SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMKIOSK.BRANCH_IP,SSINFOTERMBR.BRANCH_NM, SSMONITORING.OFFLINE_DT, TAG = 'OFFLINE' FROM SSINFOTERMKIOSK INNER JOIN SSINFOTERMBR ON SSINFOTERMKIOSK.BRANCH_CD=SSINFOTERMBR.BRANCH_CD INNER JOIN SSMONITORING ON SSINFOTERMKIOSK.BRANCH_IP=SSMONITORING.BRANCH_IP WHERE SSMONITORING.STATUS = 1 AND CAST(DATESTAMP AS DATE) = '" & Today & "' ORDER BY SSINFOTERMBR.BRANCH_NM"), lvPC)

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBack.Click
        dgvLogs.Visible = False
        btnBack.Visible = False
        gbKioskHistory.Visible = False
        lblFilter.Visible = False
        dtpFrom.Visible = False
        dtpTo.Visible = False
        lvList.Size = New System.Drawing.Size(521, 677)
        lvPC.Size = New System.Drawing.Size(552, 677)
    End Sub

    Private Sub dtpTo_ValueChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpTo.ValueChanged
        Dim DateFrom As Date = dtpFrom.Value
        Dim DateTo As Date = dtpTo.Value
        dgvLogs.Refresh()
        If dtpTo.Value < dtpFrom.Value Then
            MsgBox("Invalid Date Range.")

            dtpTo.Value = Today.Date
        Else
            db.FillDataGridView("select SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,unpvt.BRANCH_IP,unpvt.value, case when unpvt.attribute = 'OFFLINE_DT' then '*** OFFLINE ***' else 'ONLINE' end as [STATUS] FROM SSMONITORING c UNPIVOT ( value for attribute in (ONLINE_DT,OFFLINE_DT) )unpvt INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = unpvt.BRANCH_CD INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.BRANCH_IP = unpvt.BRANCH_IP INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMKIOSK.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSINFOTERMKIOSK.DIVSN where cast(unpvt.[value] as date) BETWEEN '" & DateFrom.ToShortDateString & "' and '" & DateTo.ToShortDateString & "' and unpvt.BRANCH_IP ='" & lblKIP.Text & "' UNION SELECT SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM, SSINFOTERMKIOSK.BRANCH_IP, GETDATE() AS VALUE, 'ONLINE' AS STATUS  FROM SSINFOTERMKIOSK INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSINFOTERMKIOSK.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMKIOSK.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSINFOTERMKIOSK.DIVSN WHERE SSINFOTERMKIOSK.BRANCH_IP NOT IN(SELECT SSMONITORING.BRANCH_IP FROM SSMONITORING where cast(SSMONITORING.DATESTAMP as date) = '" & Today & "') and SSINFOTERMKIOSK.BRANCH_IP = '" & lblKIP.Text & "' ORDER BY unpvt.value,BRANCH_NM", dgvLogs)

        End If
    End Sub

    Private Sub dtpFrom_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpFrom.ValueChanged
        Dim DateFrom As Date = dtpFrom.Value
        Dim DateTo As Date = dtpTo.Value
        dgvLogs.Refresh()
        If dtpTo.Value < dtpFrom.Value Then
            MsgBox("Invalid Date Range.")
            dtpFrom.Value = Today.Date

        Else
            db.FillDataGridView("select SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,unpvt.BRANCH_IP,unpvt.value, case when unpvt.attribute = 'OFFLINE_DT' then '*** OFFLINE ***' else 'ONLINE' end as [STATUS] FROM SSMONITORING c UNPIVOT ( value for attribute in (ONLINE_DT,OFFLINE_DT) )unpvt INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = unpvt.BRANCH_CD INNER JOIN SSINFOTERMKIOSK ON SSINFOTERMKIOSK.BRANCH_IP = unpvt.BRANCH_IP INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMKIOSK.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSINFOTERMKIOSK.DIVSN where cast(unpvt.[value] as date) BETWEEN '" & DateFrom.ToShortDateString & "' and '" & DateTo.ToShortDateString & "' and unpvt.BRANCH_IP ='" & lblKIP.Text & "' UNION SELECT SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM, SSINFOTERMKIOSK.BRANCH_IP, GETDATE() AS VALUE, 'ONLINE' AS STATUS  FROM SSINFOTERMKIOSK INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSINFOTERMKIOSK.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMKIOSK.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSINFOTERMKIOSK.DIVSN WHERE SSINFOTERMKIOSK.BRANCH_IP NOT IN(SELECT SSMONITORING.BRANCH_IP FROM SSMONITORING where cast(SSMONITORING.DATESTAMP as date) = '" & Today & "') and SSINFOTERMKIOSK.BRANCH_IP = '" & lblKIP.Text & "' ORDER BY unpvt.value,BRANCH_NM", dgvLogs)
        End If
    End Sub
    Public Sub operationFileTrans()
        Dim getFileName As String
        Dim newview As New CrystalReportViewer

        Try
            getFileName = "LIST OF SSIT OFFLINE"
            'dt = db.ExecuteSQLQuery("SELECT t1.BRANCH_IP,case when t1.STATUS = 1 then 'ONLINE' else '*** OFFLINE ****' end as [Status],SSINFOTERMBR.BRANCH_NM,(select ISNULL(SUM(case when t2.ONLINE_TME <> '' then cast(round(cast(t2.MSG as FLOAT),2)/ 60.0 as DECIMAL(12,2)) end),0) from SSINFOTERMCONSTAT t2 where DATESTAMP = '" & Date.Now & "' and t1.BRANCH_IP = t2.BRANCH_IP) as [ONLINE HRS],(select case when t1.STATUS = '0' then t1.LOFFLINE_DT else NULL end as [OFFLINE] from SSINFOTERMKIOSK t1 where t1.BRANCH_IP = t2.BRANCH_IP) as [OFFLINE DATE],(select sum(case when t1.STATUS = 1 then 1 else 0 end) from SSINFOTERMKIOSK t1  where isVPN = 'false' and tag = 1) as [ONLINE COUNT], (select sum(case when t1.STATUS = 0 then 1 else 0 end) from SSINFOTERMKIOSK t1  where isVPN = 'false' and tag = 1) as [OFFLINE COUNT] FROM SSINFOTERMKIOSK t1 INNER JOIN SSINFOTERMCONSTAT t2 ON t2.BRANCH_IP = t1.BRANCH_IP INNER JOIN SSINFOTERMBR ON t1.BRANCH_CD = SSINFOTERMBR.BRANCH_CD where isVPN = 'false' and tag = 1 GROUP BY t1.BRANCH_IP,SSINFOTERMBR.BRANCH_NM,t1.LOFFLINE_DT,t2.BRANCH_IP,t1.STATUS ORDER BY SSINFOTERMBR.BRANCH_NM asc")
            Dim rpt As New MonitoringKioskStatus
            cryRpt = rpt
            '  cryRpt.Refresh()
            cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
            newview.ReportSource = cryRpt
            cryRpt.SetParameterValue("@getDate", Date.Now)

            Dim getdate As String = Date.Today.ToShortDateString
            Dim gettime As String = TimeOfDay
            getdate = getdate.Replace("/", "-")
            gettime = gettime.Replace(":", ".")
            Dim myPath As String = Application.StartupPath & "\Monitoring"
            If (Not System.IO.Directory.Exists(myPath)) Then
                System.IO.Directory.CreateDirectory(myPath)
            End If


            Dim CrExportOptions As ExportOptions
            Dim CrDiskFileDestinationOptions As New  _
   DiskFileDestinationOptions()
            Dim CrFormatTypeOptions As New ExcelFormatOptions
            CrFormatTypeOptions.ExcelUseConstantColumnWidth = True
            CrFormatTypeOptions.ExcelConstantColumnWidth = 2000
            CrFormatTypeOptions.ShowGridLines = True
            CrDiskFileDestinationOptions.DiskFileName = _
                                        myPath & "\" & getFileName & " " & getdate & " " & gettime & ".xls"
            CrExportOptions = cryRpt.ExportOptions
            With CrExportOptions
                .ExportDestinationType = ExportDestinationType.DiskFile
                .ExportFormatType = ExportFormatType.Excel
                .DestinationOptions = CrDiskFileDestinationOptions
                .FormatOptions = CrFormatTypeOptions
            End With
            cryRpt.Export()
            MsgBox("Scan complete.")
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try


    End Sub
    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Button1.Enabled = False
        lvList.Enabled = False
        lvPC.Enabled = False
        loadBranches()
        lvList.Enabled = True
        lvPC.Enabled = True
        operationFileTrans()
        Button1.Enabled = True
    End Sub

End Class
