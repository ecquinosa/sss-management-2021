Option Explicit On
Imports System.IO
Imports System.Net.Sockets
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports CrystalDecisions.Windows.Forms
Public Class _frmMain
#Region "DIM DIM DIM GAMING"
    Dim cnt As Integer
    Dim Mreport As New GenerateMonReport
    Dim db As New ConnectionString
    Dim FormName As String
    Dim am As New auditModule
    Dim getAutoGenID As String
    Dim taskDone As String
    Dim duration As Date
    Dim LastRecieve As String
    Dim timestarted As Date
    Dim it As New IdleTime
    Dim dt As DataTable
    Public cryRpt As New ReportDocument
    Dim ur As New userRole
    Dim getUserRole As String
    Dim Dash As Integer
    Dim kiosk As Integer
    Dim monitoring As Integer
    Dim reports As Integer
    Dim URole As Integer
    Dim Dset As Integer
    Dim User As Integer
    Dim FB As Integer
    Public tada As Integer
#End Region
#Region "PUBLIC SUBS"
    Public Function userMgmtAll(ByVal getUserRole As String)
        Dim mgmtAll As String = db.putSingleValue("Select user_mgmt_all from tbl_mgmt_userrole1 where user_role = '" & getUserRole & "' ")
        Return mgmtAll
    End Function
    Public Function convertToBit(ByVal getBool As String)
        Dim convertedBit As String
        Select Case getBool

            Case "True"
                convertedBit = 1
            Case "False"
                convertedBit = 0
            Case Else
                convertedBit = 0
        End Select
        Return convertedBit

    End Function
    Public Sub validationOfUserRole()
        Try

            Dim string1 As String = db.putSingleValue("Select mgmt_UserType from tbl_Mgmt_UserRegistration where mgmt_Username = '" & _frmLogin.txtUname.Text & "'")
            Dim mgmtAll As Boolean = Convert.ToString(ur.userMgmtAll(string1)).Equals(Boolean.TrueString)
            If mgmtAll = True Then
                btnUser.Enabled = True
                btnKiosk.Enabled = True
                btnMonitoring.Enabled = True
                btnReports.Enabled = True
                btnUserRole.Enabled = True
                btnDatabaseSet.Enabled = True
                btnDash.Enabled = True
                btnFB.Enabled = True
            ElseIf _frmLogin.txtUname.Text = My.Settings.bypassUname And _frmLogin.txtPass.Text = My.Settings.bypassPassword Then
                btnUser.Enabled = True
                btnKiosk.Enabled = True
                btnMonitoring.Enabled = True
                btnReports.Enabled = True
                btnUserRole.Enabled = True
                btnDatabaseSet.Enabled = True
                btnDash.Enabled = True
                btnFB.Enabled = True
            Else

                Dim mgmtMonitoring As String = db.putSingleValue("Select view_monitoring from tbl_mgmt_userrole1 where user_role = '" & string1 & "' ")
                mgmtMonitoring = convertToBit(mgmtMonitoring)
                Dim mgmtReports As String = db.putSingleValue("Select view_reports from tbl_mgmt_userrole1 where user_role = '" & string1 & "' ")
                mgmtReports = convertToBit(mgmtReports)
                Dim mgmtKiosk As String = db.putSingleValue("Select view_kioskmgmt from tbl_mgmt_userrole1 where user_role = '" & string1 & "' ")
                mgmtKiosk = convertToBit(mgmtKiosk)
                Dim mgmtUser As String = db.putSingleValue("Select view_usermgmt from tbl_mgmt_userrole1 where user_role = '" & string1 & "' ")
                mgmtUser = convertToBit(mgmtUser)
                Dim mgmtUserRole As String = db.putSingleValue("Select view_userrole from tbl_mgmt_userrole1 where user_role = '" & string1 & "' ")
                mgmtUserRole = convertToBit(mgmtUserRole)
                Dim mgmtDB As String = db.putSingleValue("Select view_dbset from tbl_mgmt_userrole1 where user_role = '" & string1 & "' ")
                mgmtDB = convertToBit(mgmtDB)
                Dim mgmtFB As String = db.putSingleValue("Select view_FeedBack from tbl_mgmt_userrole1 where user_role = '" & string1 & "'")
                mgmtFB = convertToBit(mgmtFB)
                btnMonitoring.Enabled = mgmtMonitoring
                btnReports.Enabled = mgmtReports
                btnKiosk.Enabled = mgmtKiosk
                btnUserRole.Enabled = mgmtUserRole
                btnUser.Enabled = mgmtUser
                btnDatabaseSet.Enabled = mgmtDB
                btnFB.Enabled = mgmtFB

                'monitoring = mgmtMonitoring
                'reports = mgmtReports
                'kiosk = mgmtKiosk
                'URole = mgmtUserRole
                'User = mgmtUser
                'Dset = mgmtDB
                'FB = mgmtFB


            End If
            '_frmUserRegistration.dvgUsers.Columns("Password").Visible = False
        Catch ex As Exception
            MsgBox("Error Occured Please Contact your administrator!", ex.ToString)
        End Try
    End Sub
    Protected Overrides Sub OnFormClosing(ByVal e As FormClosingEventArgs)
        MyBase.OnFormClosing(e)
        If Not e.Cancel AndAlso e.CloseReason = CloseReason.UserClosing Then
            db.ExecuteSQLQuery("Update tbl_Mgmt_UserRegistration set TAG ='0' where mgmt_Username ='" & My.Settings.LastUserNameLogin & "'")
            _frmDashboard.Timer1.Stop()
            _frmDashboard.Dispose()
            Timer1.Stop()
            Timer2.Stop()
            GC.Collect()
            e.Cancel = True
            _frmPrinter.Dispose()
            _frmExportReport.Dispose()
            _frmLogin.RememberMeAgain()
            _frmLogin.RememberMeTgla()
            _frmLogin.Show()
            Me.Hide()
        End If

    End Sub
    'Protected Overrides Function ProcessCmdKey(ByRef msg As System.Windows.Forms.Message, ByVal keyData As System.Windows.Forms.Keys) As Boolean

    '    If msg.Msg = WM_KEYDOWN Or msg.Msg = WM_SYSKEYDOWN Then
    '        If keyData = (Keys.Alt Or Keys.F4) Then
    '            Timer1.Stop()
    '            'Listener.Stop()
    '            db.ExecuteSQLQuery("Update tbl_Mgmt_UserRegistration set TAG ='0' where mgmt_Username ='" & My.Settings.LastUserNameLogin & "'")
    '            _frmDashboard.Timer1.Stop()
    '            _frmDashboard.cryRpt3.Close()
    '            _frmDashboard.cryRpt3.Dispose()
    '            _frmDashboard.cryRpt3.Close()
    '            _frmDashboard.cryRpt3.Dispose()
    '            GC.Collect()
    '            _frmLogin.RememberMeTgla()
    '            _frmLogin.Show()
    '            Me.Close()
    '        End If
    '        Return True
    '    End If
    '    Return MyBase.ProcessCmdKey(msg, keyData)
    'End Function

    Public Sub buttonclicked()
        If cnt = 1 Then

            btnDash.FlatStyle = FlatStyle.Flat
            btnDash.BackColor = Color.Blue
            btnDash.ForeColor = Color.White
        Else
            btnDash.BackColor = Color.White
            btnDash.FlatStyle = FlatStyle.Standard
            btnDash.ForeColor = Color.Black
        End If
        If cnt = 2 Then
            btnKiosk.FlatStyle = FlatStyle.Flat
            btnKiosk.BackColor = Color.Blue
            btnKiosk.ForeColor = Color.White
        Else
            btnKiosk.BackColor = Color.White
            btnKiosk.FlatStyle = FlatStyle.Standard
            btnKiosk.ForeColor = Color.Black
        End If
        If cnt = 3 Then

            btnMonitoring.FlatStyle = FlatStyle.Flat
            btnMonitoring.BackColor = Color.Blue
            btnMonitoring.ForeColor = Color.White
        Else
            btnMonitoring.BackColor = Color.White
            btnMonitoring.FlatStyle = FlatStyle.Standard
            btnMonitoring.ForeColor = Color.Black
        End If
        If cnt = 4 Then

            btnReports.FlatStyle = FlatStyle.Flat
            btnReports.BackColor = Color.Blue
            btnReports.ForeColor = Color.White
        Else
            btnReports.BackColor = Color.White
            btnReports.FlatStyle = FlatStyle.Standard
            btnReports.ForeColor = Color.Black
        End If

        If cnt = 5 Then

            btnUser.FlatStyle = FlatStyle.Flat
            btnUser.BackColor = Color.Blue
            btnUser.ForeColor = Color.White
        Else
            btnUser.BackColor = Color.White
            btnUser.FlatStyle = FlatStyle.Standard
            btnUser.ForeColor = Color.Black
        End If

        If cnt = 6 Then
            btnUserRole.FlatStyle = FlatStyle.Flat
            btnUserRole.BackColor = Color.Blue
            btnUserRole.ForeColor = Color.White
        Else
            btnUserRole.BackColor = Color.White
            btnUserRole.FlatStyle = FlatStyle.Standard
            btnUserRole.ForeColor = Color.Black
        End If

        If cnt = 7 Then
            btnDatabaseSet.FlatStyle = FlatStyle.Flat
            btnDatabaseSet.BackColor = Color.Blue
            btnDatabaseSet.ForeColor = Color.White
        Else
            btnDatabaseSet.BackColor = Color.White
            btnDatabaseSet.FlatStyle = FlatStyle.Standard
            btnDatabaseSet.ForeColor = Color.Black
        End If

        If cnt = 8 Then
            btnFB.FlatStyle = FlatStyle.Flat
            btnFB.BackColor = Color.Blue
            btnFB.ForeColor = Color.White
        Else
            btnFB.BackColor = Color.White
            btnFB.FlatStyle = FlatStyle.Standard
            btnFB.ForeColor = Color.Black


        End If
    End Sub
    Public Sub fillListview()
        Try
            _frmDashboard.lvDiv.Items.Clear()
            _frmDashboard.lvGroup.Items.Clear()
            _frmDashboard.lvList.Items.Clear()
            db.FillListView(db.ExecuteSQLQuery("Select SSINFOTERMGROUP.GROUP_NM as 'GROUP' from SSINFOTERMKIOSK INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSINFOTERMKIOSK.DIVSN  WHERE STATUS = 'false' and TAG = 1"), _frmDashboard.lvDivsn)
            db.RemoveDup(_frmDashboard.lvDivsn)
            db.FillListView(db.ExecuteSQLQuery("Select SSINFOTERMBR.BRANCH_NM as 'BRANCH' from SSINFOTERMKIOSK inner join SSINFOTERMBR on SSINFOTERMKIOSK.BRANCH_CD = SSINFOTERMBR.BRANCH_CD where STATUS ='false' and TAG = 1 "), _frmDashboard.lvDiv)
            db.RemoveDup(_frmDashboard.lvDiv)
            db.FillListView(db.ExecuteSQLQuery("Select SSINFOTERMCLSTR.CLSTR_NM as 'DIVISION' from SSINFOTERMKIOSK inner join SSINFOTERMCLSTR on SSINFOTERMKIOSK.CLSTR = SSINFOTERMCLSTR.CLSTR_CD where STATUS ='false' and TAG = 1"), _frmDashboard.lvGroup)
            db.RemoveDup(_frmDashboard.lvGroup)
            db.FillListView(db.ExecuteSQLQuery("SELECT KIOSK_ID as'Terminal ID',STATUS,BRANCH_IP as 'IP ADDRESS', SSINFOTERMBR.BRANCH_NM as 'Branch Name',LOFFLINE_DT as 'Offline date' FROM SSINFOTERMKIOSK inner join SSINFOTERMBR on SSINFOTERMKIOSK.BRANCH_CD = SSINFOTERMBR.BRANCH_CD where STATUS = 'false' and TAG = 1 ORDER BY LONLINE_DT ASC"), _frmDashboard.lvList)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
#End Region
    Private Sub _frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SplitContainer1.SetBounds(100, 100, 100, 100)
        SplitContainer1.VerticalScroll.Enabled = False
        getUserRole = db.putSingleValue("select mgmt_usertype from tbl_Mgmt_UserRegistration where mgmt_username ='" & _frmLogin.txtUname.Text & "'")
        validationOfUserRole()
        My.Settings.Module_Tag = 3
        getAutoGenID = am.getAutoID()
        Timer2.Start()
        Timer1.Interval = 1000
        Timer1.Start()
        SplitContainer1.Panel2.Controls.Clear()
        _frmDashboard.TopLevel = False
        _frmDashboard.Parent = Me.SplitContainer1.Panel2
        _frmDashboard.Dock = DockStyle.Fill
        _frmDashboard.Show()
        btnDash.FlatStyle = FlatStyle.Flat
        btnDash.BackColor = Color.Blue
        btnDash.ForeColor = Color.White
    End Sub

    Private Sub ExpandablePanel1_ExpandedChanged(ByVal sender As Object, ByVal e As DevComponents.DotNetBar.ExpandedChangeEventArgs)
        exPan2.Dock = DockStyle.Top
    End Sub
    Private Sub ExpandablePanel1_ExpandedChanging(ByVal sender As Object, ByVal e As DevComponents.DotNetBar.ExpandedChangeEventArgs)
        exPan2.Dock = DockStyle.Bottom
    End Sub
    Private Sub btnSettings_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        FormName = am.getModuleName(My.Settings.Module_Tag)
        taskDone = "Click System Settings Button"
        db.sql = "Insert into tbl_Mgmt_auditTrail values('" & getAutoGenID & "','" & My.Settings.LastUserNameLogin & "','" & taskDone & "','" & FormName & "', '" & "" & "',  '" & DateTime.Today.ToShortDateString & "', '" & TimeOfDay & "' )"
        db.ExecuteSQLQuery(db.sql)
    End Sub
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick

        Dim a As New Label
        a.Text = it.IdleTime
        If My.Settings.Idletime = a.Text Then
            If tada = 1 Then
                _frmMonitoring.trd.Suspend()
            End If
            Timer2.Stop()
            Timer1.Stop()
            taskDone = "User Idle"
            MsgBox("Idle for " & My.Settings.Idletime / 60 & " minutes.")
            db.ExecuteSQLQuery("Update tbl_Mgmt_UserRegistration set TAG ='0' where mgmt_Username ='" & My.Settings.LastUserNameLogin & "'")
            db.sql = "Insert into tbl_Mgmt_auditTrail values('" & getAutoGenID & "','" & My.Settings.LastUserNameLogin & "','" & taskDone & "','" & FormName & "', '" & "" & "',  '" & DateTime.Today.ToShortDateString & "', '" & TimeOfDay & "' )"
            db.ExecuteSQLQuery(db.sql)
            _frmViewReports.rptView.Dispose()
            _frmPrinter.Dispose()
            _frmExportReport.Dispose()
            _frmDashboard.Dispose()
            _frmLogin.RememberMeTgla()
            _frmLogin.Show()
            Me.Dispose()
        End If
    End Sub
    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Mreport.ConvertToExcel()
        Mreport.convertToPDf()
    End Sub
    Private Sub btnDash_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDash.Click
        cnt = 1
        buttonclicked()
        _frmDashboard.lblTot.Text = Nothing
        _frmDashboard.lblOnline1.Text = Nothing
        _frmDashboard.lblOffline1.Text = Nothing
        _frmDashboard.lblBranch.Text = Nothing
        _frmDashboard.fillListview()
        _frmDashboard.fillKioskNum()
        _frmDashboard.lvGroup.Columns(0).Width = -2
        _frmDashboard.lvDiv.Columns(0).Width = -2
        _frmDashboard.lvDivsn.Columns(0).Width = -2
        SplitContainer1.Panel2.Controls.Clear()
        _frmDashboard.fillData()
        _frmDashboard.TopLevel = False
        _frmDashboard.Parent = Me.SplitContainer1.Panel2
        _frmDashboard.Dock = DockStyle.Fill
        _frmDashboard.Show()
    End Sub

    Private Sub btnKiosk_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnKiosk.Click
        cnt = 2
        buttonclicked()
        FormName = am.getModuleName(My.Settings.Module_Tag)
        taskDone = "Click Kiosk Registration Button"
        db.sql = "Insert into tbl_Mgmt_auditTrail values('" & getAutoGenID & "','" & My.Settings.LastUserNameLogin & "','" & taskDone & "','" & FormName & "', '" & "" & "',  '" & DateTime.Today.ToShortDateString & "', '" & TimeOfDay & "' )"
        db.ExecuteSQLQuery(db.sql)
        SplitContainer1.Panel2.Controls.Clear()
        _frmKiosk.TopLevel = False
        _frmKiosk.Parent = Me.SplitContainer1.Panel2
        _frmKiosk.Dock = DockStyle.Fill
        _frmKiosk.Show()
    End Sub
    Private Sub btnMonitoring_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMonitoring.Click
        cnt = 3
        buttonclicked()

        FormName = am.getModuleName(My.Settings.Module_Tag)
        taskDone = "Click Monitoring Button"
        db.sql = "Insert into tbl_Mgmt_auditTrail values('" & getAutoGenID & "','" & My.Settings.LastUserNameLogin & "','" & taskDone & "','" & FormName & "', '" & "" & "',  '" & DateTime.Today.ToShortDateString & "', '" & TimeOfDay & "' )"
        db.ExecuteSQLQuery(db.sql)

        SplitContainer1.Panel2.Controls.Clear()
        _frmMonitoring.TopLevel = False
        _frmMonitoring.Parent = Me.SplitContainer1.Panel2
        _frmMonitoring.Dock = DockStyle.Fill
        _frmMonitoring.Show()
        tada = 1
    End Sub

    Private Sub btnReports_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReports.Click
        cnt = 4
        buttonclicked()
        FormName = am.getModuleName(My.Settings.Module_Tag)
        taskDone = "Click Report Generator Button"
        db.sql = "Insert into tbl_Mgmt_auditTrail values('" & getAutoGenID & "','" & My.Settings.LastUserNameLogin & "','" & taskDone & "','" & FormName & "', '" & "" & "',  '" & DateTime.Today.ToShortDateString & "', '" & TimeOfDay & "' )"
        db.ExecuteSQLQuery(db.sql)
        SplitContainer1.Panel2.Controls.Clear()
        _frmViewReports.TopLevel = False
        _frmViewReports.Parent = Me.SplitContainer1.Panel2
        _frmViewReports.Dock = DockStyle.Fill
        _frmViewReports.Show()
    End Sub

    Private Sub btnFB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFB.Click
        cnt = 8
        buttonclicked()
        SplitContainer1.Panel2.Controls.Clear()
        _frmFBK.TopLevel = False
        _frmFBK.Parent = Me.SplitContainer1.Panel2
        _frmFBK.Dock = DockStyle.Fill
        _frmFBK.Show()
    End Sub

    Private Sub btnUser_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUser.Click
        cnt = 5
        buttonclicked()
        FormName = am.getModuleName(My.Settings.Module_Tag)
        taskDone = "Click User Registration Button"
        db.sql = "Insert into tbl_Mgmt_auditTrail values('" & getAutoGenID & "','" & My.Settings.LastUserNameLogin & "','" & taskDone & "','" & FormName & "', '" & "" & "',  '" & DateTime.Today.ToShortDateString & "', '" & TimeOfDay & "' )"
        db.ExecuteSQLQuery(db.sql)
        SplitContainer1.Panel2.Controls.Clear()
        _frmUserRegistration.TopLevel = False
        _frmUserRegistration.Parent = Me.SplitContainer1.Panel2
        _frmUserRegistration.Dock = DockStyle.Fill
        _frmUserRegistration.Show()
    End Sub

    Private Sub btnUserRole_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUserRole.Click
        cnt = 6
        buttonclicked()
        FormName = am.getModuleName(My.Settings.Module_Tag)
        taskDone = "Click User Role Button"
        db.sql = "Insert into tbl_Mgmt_auditTrail values('" & getAutoGenID & "','" & My.Settings.LastUserNameLogin & "','" & taskDone & "','" & FormName & "', '" & "" & "',  '" & DateTime.Today.ToShortDateString & "', '" & TimeOfDay & "' )"
        db.ExecuteSQLQuery(db.sql)
        SplitContainer1.Panel2.Controls.Clear()
        _frmUserRole.TopLevel = False
        _frmUserRole.Parent = Me.SplitContainer1.Panel2
        _frmUserRole.Dock = DockStyle.Fill
        _frmUserRole.Show()
    End Sub

    Private Sub btnDatabaseSet_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDatabaseSet.Click
        cnt = 7
        buttonclicked()
        FormName = am.getModuleName(My.Settings.Module_Tag)
        taskDone = "Click System Settings Button"

        SplitContainer1.Panel2.Controls.Clear()
        _frmSettings.TopLevel = False
        _frmSettings.Parent = Me.SplitContainer1.Panel2
        _frmSettings.Dock = DockStyle.Fill
        _frmSettings.Show()
        db.sql = "Insert into tbl_Mgmt_auditTrail values('" & getAutoGenID & "','" & My.Settings.LastUserNameLogin & "','" & taskDone & "','" & FormName & "', '" & "" & "',  '" & DateTime.Today.ToShortDateString & "', '" & TimeOfDay & "' )"
        db.ExecuteSQLQuery(db.sql)
    End Sub

    Private Sub btnLogout_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLogout.Click
        If tada = 1 Then
            _frmMonitoring.trd.Suspend()
        End If
        Timer1.Stop()
        'Listener.Stop()
        FormName = am.getModuleName(My.Settings.Module_Tag)
        taskDone = "Click Logout Button"
        db.ExecuteSQLQuery("update tbl_Mgmt_UserRegistration set TAG = '0' where mgmt_Username = '" & My.Settings.LastUserNameLogin & "'")
        db.sql = "Insert into tbl_Mgmt_auditTrail values('" & getAutoGenID & "','" & My.Settings.LastUserNameLogin & "','" & taskDone & "','" & FormName & "', '" & "" & "',  '" & DateTime.Today.ToShortDateString & "', '" & TimeOfDay & "' )"
        db.ExecuteSQLQuery(db.sql)
        _frmViewReports.rptView.Dispose()
        _frmPrinter.Dispose()
        _frmExportReport.Dispose()
        _frmDashboard.Dispose()
        _frmLogin.RememberMeTgla()
        _frmLogin.Show()
        Me.Dispose()

    End Sub
    Dim i As Integer
    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
            Dim temp As String = "DSN=" & My.Settings.db_DSN & ";SERVER=" & My.Settings.db_Server & ";DATABASE=" & My.Settings.db_Name & ";UID=" & My.Settings.db_UName & ";PWD=" & My.Settings.db_Pass & ""
            'Dim temp As String = "SERVER=" & TextBox2.Text & ";DATABASE=" & TextBox5.Text & ";UID=" & TextBox3.Text & ";PWD=" & TextBox4.Text & ""
            ''Dim temp As String = "Data Source=" & TextBox2.Text & ";Initial Catalog=" & TextBox5.Text & ";User ID=" & TextBox3.Text & ";Password=" & TextBox4.Text & ""
        If db.webisconnected(temp) Then
            If i = 1 Then
                validationOfUserRole()
            End If
            i = 0
            'Dim status As String = db.putSingleValue("Select TAG from tbl_Mgmt_UserRegistration where mgmt_Username = '" & My.Settings.LastUserNameLogin & "'")
            Dim status As String = db.putSingleValue("Select TAG from tbl_Mgmt_UserRegistration where mgmt_Username = '" & "admin" & "'")
            If status = 0 Then
                'db.ExecuteSQLQuery("Update tbl_Mgmt_UserRegistration set TAG ='1' where mgmt_Username ='" & My.Settings.LastUserNameLogin & "'")
                Timer2.Stop()
                Timer1.Stop()
                If tada = 1 Then
                    _frmMonitoring.trd.Abort()
                End If
                _frmPrinter.Dispose()
                _frmExportReport.Dispose()
                _frmDashboard.Dispose()
                _frmBlock.Dispose()
                _frmLogin.RememberMeTgla()
                _frmLogin.Show()
                Me.Dispose()
            End If
        Else
            _frmDashboard.Timer1.Stop()
            Timer2.Stop()
            Timer1.Stop()

            If i = 0 Then
                i = 1
                btnDash.Enabled = False
                btnMonitoring.Enabled = False
                btnUser.Enabled = False
                btnKiosk.Enabled = False
                btnDatabaseSet.Enabled = False
                btnReports.Enabled = False
                btnFB.Enabled = False
                btnUserRole.Enabled = False
                SplitContainer1.Panel2.Controls.Clear()
                _frmLoading.TopLevel = False
                _frmLoading.Parent = Me.SplitContainer1.Panel2
                _frmLoading.Dock = DockStyle.Fill
                _frmLoading.Show()
                MsgBox("Connection time-out.", MsgBoxStyle.Exclamation)
            End If

        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub
End Class