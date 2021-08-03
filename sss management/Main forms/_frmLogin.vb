Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports CrystalDecisions.Windows.Forms
Imports System.Net
Public Class _frmLogin
    Dim db As New ConnectionString
    Dim taskDone As String
    Dim am As New auditModule
    Dim getAutoGenID As String
    Dim FormName As String
    Private int As Integer
    Dim ur As New userRole
    Public getUserRole As String
    Dim dt As DataTable
    Dim cryRpt As New ReportDocument
    Dim up As New UploadFTP
    Dim a As String
    Dim b As String
    Dim plainText As String
    Dim cipherText As String
    Dim passOld As String
    Dim passPhrase As String
    Dim saltValue As String
    Dim hashAlgorithm As String
    Dim passwordIterations As Integer
    Dim initVector As String
    Dim keySize As Integer
    'Public Sub operationFileTrans()
    '    Dim getFileName As String
    '    Dim newview As New CrystalReportViewer

    '    Try
    '        getFileName = "Technical Retirement"
    '        dt = db.ExecuteSQLQuery("SELECT SSNUM,DOBTH,DOCVRG,UMID_BANK,BANK_BR,TRANS_ACCTNO,STREET,BRGAY,POST_CD,EMAIL,PHONE,CELNO,USRTYP,SSTRANSINFOTERMTR.ENCODE_DT,SSTRANSINFOTERMTR.ENCODE_TME ,SSTRANSINFOTERMTR.BRANCH_IP,SSINFOTERMKIOSK.KIOSK_ID,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSTRANSINFOTERMTR.TAG,TRANID FROM SSTRANSINFOTERMTR INNER JOIN SSINFOTERMKIOSK on SSINFOTERMKIOSK.BRANCH_IP = SSTRANSINFOTERMTR.BRANCH_IP inner join SSINFOTERMBR on SSTRANSINFOTERMTR.BRANCH_CD = SSINFOTERMBR.BRANCH_CD inner join SSINFOTERMCLSTR on SSTRANSINFOTERMTR.CLSTR = SSINFOTERMCLSTR.CLSTR_CD where SSTRANSINFOTERMTR.ENCODE_DT BETWEEN '" & Today.Date & "' and '" & Today.Date & "'")
    '        Dim rpt As New FTPTRANS
    '        cryRpt = rpt
    '        '  cryRpt.Refresh()
    '        cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
    '        newview.ReportSource = cryRpt
    '        cryRpt.SetParameterValue("@DateFrom", Today.Date)
    '        cryRpt.SetParameterValue("@DateTo", Today.Date)


    '        Dim getdate As String = Date.Today.ToShortDateString
    '        Dim gettime As String = TimeOfDay
    '        getdate = getdate.Replace("/", "-")
    '        gettime = gettime.Replace(":", ".")
    '        Dim myPath As String = Application.StartupPath & "\PDF"
    '        If (Not System.IO.Directory.Exists(myPath)) Then
    '            System.IO.Directory.CreateDirectory(myPath)
    '        End If


    '        Dim CrExportOptions As ExportOptions
    '        Dim CrDiskFileDestinationOptions As New  _
    '        DiskFileDestinationOptions()
    '        Dim CrFormatTypeOptions As New PdfRtfWordFormatOptions()
    '        CrDiskFileDestinationOptions.DiskFileName = _
    '                                    myPath & "\" & getFileName & " " & getdate & ".pdf"
    '        CrExportOptions = cryRpt.ExportOptions
    '        With CrExportOptions
    '            .ExportDestinationType = ExportDestinationType.DiskFile
    '            .ExportFormatType = ExportFormatType.PortableDocFormat
    '            .DestinationOptions = CrDiskFileDestinationOptions
    '            .FormatOptions = CrFormatTypeOptions
    '        End With
    '        cryRpt.Export()


    '        Dim fileName As String = myPath & "\" & getFileName & " " & getdate & ".pdf"
    '        fileName = fileName.Replace("\", "/")
    '        Dim DestFile As String = My.Settings.FtpFileUpload & getFileName & " " & getdate & ".pdf"
    '        up.UploadFile(fileName, DestFile, My.Settings.FTPUN, My.Settings.FTPPASS)

    '    Catch ex As Exception

    '        MsgBox(ex.ToString)
    '    End Try


    'End Sub
    Public Sub RememberMeTgla()
        Select Case chkRemember.Checked
            Case True
                If Text.Contains(txtUname.Text) = False Then
                    My.Settings.RememberMe1 = True
                    My.Settings.RememberMe = txtUname.Text
                    My.Settings.Save()
                Else
                    txtUname.Text = My.Settings.RememberMe
                    My.Settings.RememberMe1 = True
                    My.Settings.Save()
                End If
            Case False
                My.Settings.RememberMe1 = False
                My.Settings.Save()
        End Select
    End Sub
    Public Sub RememberMeAgain()
        If chkRemember.Checked = True Then
            If Text.Contains(txtUname.Text) = False Then
                My.Settings.RememberMe = txtUname.Text
                My.Settings.Save()
            Else
                txtUname.Text = My.Settings.RememberMe
            End If
        End If

        If My.Settings.RememberMe1 = True Then
            chkRemember.Checked = True
        Else
            chkRemember.Checked = False
        End If
    End Sub
    Public Sub checkDbSettings()
        My.Settings.Reload()
        If My.Settings.db_Server = "" Or My.Settings.db_UName = "" Or My.Settings.db_Name = "" Or My.Settings.db_DSN = "" Then
            MsgBox("Please Check your Database Settings before Loggin in!")
        End If
    End Sub
    Private Sub _frmLogin2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtUname.Focus()
        Me.BackColor = Nothing
        TransparencyKey = BackColor
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.None
        My.Settings.Module_Tag = 1
        RememberMeAgain()
        If My.Settings.firstRun = 0 Then
            _frmServerCon.Dispose()
            _frmBlock.Show()
            _frmServerCon.ShowDialog()
            '_frmServerCon.TopMost = True
            _frmServerCon.btnBack.Visible = True
            _frmBlock.Close()

        End If
    End Sub

    Private Sub ButtonX3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonX3.Click
        Try
            Dim Year As String = db.putSingleValue("Select Year from tblYear where Year = '" & Date.Today.Year & "'")

            If (String.IsNullOrEmpty(Year)) Then
                db.sql = "Insert into tblYear values('" & Date.Today.Year & "' )"
                db.ExecuteSQLQuery(db.sql)
            End If

            If My.Settings.firstRun = 0 Then
                _frmServerCon.Dispose()
                _frmBlock.Show()
                _frmServerCon.ShowDialog()
                '_frmServerCon.TopMost = True
                _frmServerCon.btnBack.Visible = True
                _frmBlock.Close()
            Else

                checkDbSettings()
                getAutoGenID = am.getAutoID()
                RememberMeTgla()
                a = txtUname.Text
                Dim _counter As String = db.putSingleValue("Select attempt from tbl_Mgmt_UserRegistration where mgmt_Username = '" & txtUname.Text & "'")
                Dim checkBlocked As String = db.putSingleValue("select mgmt_status from tbl_Mgmt_UserRegistration where mgmt_Username = '" & txtUname.Text & "'")
                Dim checkPass As String = db.putSingleValue("select mgmt_Password from tbl_Mgmt_UserRegistration where mgmt_Username = '" & txtUname.Text & "'")
                Dim firstlog As Integer = db.putSingleValue("select FirstLog from tbl_Mgmt_UserRegistration where mgmt_Username = '" & txtUname.Text & "'")
                If checkBlocked = "LOCKED" Then
                    Dim timeLock1 As Date = db.putSingleValue("Select iSlocked from tbl_mgmt_UserRegistration where mgmt_Username = '" & txtUname.Text & "'")
                    'If Date.Now < timeLock1 Then
                    '    MsgBox("Your account has been locked. Please try again after 15 minutes.", MsgBoxStyle.Exclamation)
                    'Else
                    '    db.ExecuteSQLQuery("UPDATE tbl_mgmt_UserRegistration set mgmt_STATUS = 'ACTIVE',iSlocked = '',Attempt = '0' where mgmt_Username = '" & txtUname.Text & "' ")
                    'End If
                    db.ExecuteSQLQuery("UPDATE tbl_mgmt_UserRegistration set mgmt_STATUS = 'ACTIVE',iSlocked = '',Attempt = '0' where mgmt_Username = '" & txtUname.Text & "' ")
                ElseIf txtUname.Text = "" Then
                    MsgBox("Username is empty.", MsgBoxStyle.Information)
                ElseIf txtPass.Text = "" Then
                    MsgBox("Password is empty.", MsgBoxStyle.Information)

                ElseIf firstlog = "0" And txtPass.Text = "password" Then

                    _frmBlock.Show()
                    _frmChangePassLogin.ShowDialog()
                    '_frmChangePassLogin.TopMost = True
                    _frmBlock.Close()

                    FormName = am.getModuleName(My.Settings.Module_Tag)
                    taskDone = "User Login"
                    db.sql = "Insert into tbl_Mgmt_auditTrail values('" & getAutoGenID & "','" & txtUname.Text & "','" & taskDone & "','" & FormName & "', '" & "" & "',  '" & DateTime.Today.ToShortDateString & "', '" & TimeOfDay & "' )"
                    db.ExecuteSQLQuery(db.sql)
                    'validationOfUserRole()

                    'audit trail


                    My.Settings.LastUserNameLogin = txtUname.Text
                    My.Settings.LastPasswordLogin = txtPass.Text
                    _counter = 0
                    db.ExecuteSQLQuery("Update tbl_Mgmt_UserRegistration set Attempt = '" & _counter & "' where mgmt_Username = '" & txtUname.Text & "'")
                Else
                    If My.Settings.bypassUname = txtUname.Text And My.Settings.bypassPassword = txtPass.Text Then
                        MsgBox("Welcome ADMINISTRATOR", MsgBoxStyle.Information)
                        Me.Hide()
                        _frmMain.Show()
                        _frmMain.SplitContainer1.Panel2.Controls.Clear()
                        _frmDashboard.TopLevel = False
                        _frmDashboard.Parent = _frmMain.SplitContainer1.Panel2
                        _frmDashboard.Dock = DockStyle.Fill
                        _frmDashboard.Show()
                        _counter = 0
                    Else

                        If db.checkExistence("Select mgmt_Username from tbl_Mgmt_UserRegistration where mgmt_Username = '" & txtUname.Text & "'") Then
                            If txtPass.Text = Nothing Then
                                MsgBox("Password is empty.", MsgBoxStyle.Information)
                            Else
                                cipherText = db.putSingleValue("Select mgmt_Password from tbl_Mgmt_UserRegistration where mgmt_Username = '" & txtUname.Text & "'")
                                passPhrase = "Pas5pr@se"        ' can be any string
                                saltValue = "s@1tValue"         ' can be any string
                                hashAlgorithm = "SHA1"          ' can be "MD5"
                                passwordIterations = 2          ' can be any number
                                initVector = "@1B2c3D4e5F6g7H8" ' must be 16 bytes
                                keySize = 256                   ' can be 192 or 128
                                plainText = RijndaelSimple.Decrypt _
                                    ( _
              cipherText, _
               passPhrase, _
               saltValue, _
               hashAlgorithm, _
               passwordIterations, _
               initVector, _
               keySize _
            )
                                If txtPass.Text = plainText Then
                                    If db.checkExistence("select mgmt_Username,TAG from tbl_Mgmt_UserRegistration where mgmt_Username='" & txtUname.Text & "' and TAG = '0'") = True Then
                                        db.ExecuteSQLQuery("Update tbl_Mgmt_UserRegistration set TAG ='0' where mgmt_Username ='" & txtUname.Text & "'")
                                        db.sql = "Insert into tbl_Mgmt_auditTrail values('" & getAutoGenID & "','" & txtUname.Text & "','" & taskDone & "','" & FormName & "', '" & "" & "',  '" & DateTime.Today.ToShortDateString & "', '" & TimeOfDay & "' )"
                                        db.ExecuteSQLQuery(db.sql)
                                        Me.Hide()
                                        _frmMain.Show()
                                        _frmMain.SplitContainer1.Panel2.Controls.Clear()
                                        _frmDashboard.TopLevel = False
                                        _frmDashboard.Parent = _frmMain.SplitContainer1.Panel2
                                        _frmDashboard.Dock = DockStyle.Fill
                                        _frmDashboard.Show()

                                        My.Settings.LastUserNameLogin = txtUname.Text
                                        My.Settings.LastPasswordLogin = txtPass.Text

                                        db.ExecuteSQLQuery("Update tbl_Mgmt_UserRegistration set TAG ='1' where mgmt_Username ='" & txtUname.Text & "'")


                                        _counter = 0
                                        db.ExecuteSQLQuery("Update tbl_Mgmt_UserRegistration set Attempt = '" & _counter & "' where mgmt_Username = '" & txtUname.Text & "'")
                                        txtPass.Text = ""
                                        txtUname.Text = ""
                                    Else
                                        If MessageBox.Show("Your account is currently logged in on another computer. Do you want to continue and terminate the current login?", "SSIT MANAGEMENT SERVER", MessageBoxButtons.YesNo, MessageBoxIcon.Information) = MsgBoxResult.Yes Then
                                            db.ExecuteSQLQuery("Update tbl_Mgmt_UserRegistration set TAG ='0' where mgmt_Username ='" & txtUname.Text & "'")
                                            db.sql = "Insert into tbl_Mgmt_auditTrail values('" & getAutoGenID & "','" & txtUname.Text & "','" & taskDone & "','" & FormName & "', '" & "" & "',  '" & DateTime.Today.ToShortDateString & "', '" & TimeOfDay & "' )"
                                            db.ExecuteSQLQuery(db.sql)
                                            Me.Hide()
                                            _frmMain.Show()
                                            _frmMain.SplitContainer1.Panel2.Controls.Clear()
                                            _frmDashboard.TopLevel = False
                                            _frmDashboard.Parent = _frmMain.SplitContainer1.Panel2
                                            _frmDashboard.Dock = DockStyle.Fill
                                            _frmDashboard.Show()

                                            My.Settings.LastUserNameLogin = txtUname.Text
                                            My.Settings.LastPasswordLogin = txtPass.Text

                                            db.ExecuteSQLQuery("Update tbl_Mgmt_UserRegistration set TAG ='1' where mgmt_Username ='" & txtUname.Text & "'")

                                            _counter = 0
                                            db.ExecuteSQLQuery("Update tbl_Mgmt_UserRegistration set Attempt = '" & _counter & "' where mgmt_Username = '" & txtUname.Text & "'")
                                            txtPass.Text = ""
                                            txtUname.Text = ""
                                        End If

                                    End If
                                Else
                                    If _counter = 3 Then
                                        MsgBox(" Your account has been locked. Please try again after 15 minutes.", MsgBoxStyle.Exclamation)
                                        Dim timeLock As Date = DateAdd(DateInterval.Minute, 15, Date.Now)
                                        _counter = 3
                                        db.ExecuteSQLQuery("update tbl_Mgmt_UserRegistration set mgmt_status = '" & "LOCKED" & "',iSlocked = '" & timeLock & "', Attempt = '" & _counter & "' where mgmt_Username = '" & txtUname.Text & "'")

                                    Else

                                        If a = b Then
                                            _counter += 1
                                            db.ExecuteSQLQuery("Update tbl_Mgmt_UserRegistration set Attempt = '" & _counter & "' where mgmt_Username = '" & txtUname.Text & "'")
                                            MsgBox("The username or password you entered is incorrect. You have used " & _counter & " of 3 login attempts.", MsgBoxStyle.Information)
                                        Else
                                            _counter += 1
                                            db.ExecuteSQLQuery("Update tbl_Mgmt_UserRegistration set Attempt = '" & _counter & "' where mgmt_Username = '" & txtUname.Text & "'")
                                            MsgBox("The username or password you entered is incorrect. You have used " & _counter & " of 3 login attempts.", MsgBoxStyle.Information)
                                            b = txtUname.Text
                                        End If

                                    End If
                                End If
                            End If
                        Else
                            MsgBox("Invalid username or password.", MsgBoxStyle.Information)
                        End If
                    End If
                End If



                'If checkBlocked = "BLOCKED" Then
                '    MsgBox("ACCOUNT BLOCK" & vbNewLine & "Please contact your system administrator regarding this issue", MsgBoxStyle.Exclamation)
                'ElseIf checkPass = "0000" Then

                '    _frmBlock.Show()
                '    _frmChangePassLogin.ShowDialog()
                '    _frmChangePassLogin.TopMost = True
                '    _frmBlock.Close()


                '    FormName = am.getModuleName(My.Settings.Module_Tag)
                '    taskDone = "User Login"
                '    db.sql = "Insert into tbl_Mgmt_auditTrail values('" & getAutoGenID & "','" & txtUname.Text & "','" & taskDone & "','" & FormName & "', '" & "" & "',  '" & DateTime.Today.ToShortDateString & "', '" & TimeOfDay & "' )"
                '    db.ExecuteSQLQuery(db.sql)
                '    Me.Hide()
                '    'validationOfUserRole()
                '    _frmMain.Show()

                '    'audit trail


                '    My.Settings.LastUserNameLogin = txtUname.Text
                '    My.Settings.LastPasswordLogin = txtPass.Text

                '    txtPass.Text = ""
                '    txtUname.Text = ""
                'Else



                '    If My.Settings.bypassUname = txtUname.Text And My.Settings.bypassPassword = txtPass.Text Then
                '        MsgBox("Welcome ADMINISTRATOR", MsgBoxStyle.Information)
                '        Me.Hide()
                '        _frmMain.Show()
                '        _frmMain.SplitContainer1.Panel2.Controls.Clear()
                '        _frmDashboard.TopLevel = False
                '        _frmDashboard.Parent = _frmMain.SplitContainer1.Panel2
                '        _frmDashboard.Dock = DockStyle.Fill
                '        _frmDashboard.Show()
                '    Else
                '        If txtUname.Text = "" Then
                '            int = 0
                '            tmr1.Enabled = True
                '            tmr1.Start()
                '            lblErrorMessage.Text = "Invalid Username or Password"
                '            _counter -= 1
                '            'pag wla laman same
                '            If _counter = 0 Then
                '                MsgBox("You have entered 3 incorrect password", MsgBoxStyle.Exclamation)
                '                MsgBox("System will automatically close", MsgBoxStyle.Exclamation)


                '                FormName = am.getModuleName(My.Settings.Module_Tag)
                '                taskDone = "Login failed"
                '                db.sql = "Insert into tbl_Mgmt_auditTrail values('" & getAutoGenID & "','" & txtUname.Text & "','" & taskDone & "','" & FormName & "', '" & "" & "',  '" & DateTime.Today.ToShortDateString & "', '" & TimeOfDay & "' )"
                '                db.ExecuteSQLQuery(db.sql)
                '                txtPass.Clear()
                '                txtUname.Clear()
                '                txtUname.Focus()
                '                pbError.Hide()
                '                lblErrorMessage.Hide()
                '                lblErrorMessage.Text = "ERROR MESSAGE"
                '                _counter = 3
                '            End If
                '            txtUname.Focus()
                '        ElseIf txtPass.Text = "" Then
                '            int = 0
                '            tmr1.Enabled = True
                '            tmr1.Start()
                '            lblErrorMessage.Text = "Invalid Username or Password"
                '            _counter -= 1
                '            'pag wala laman pass
                '            If _counter = 0 Then
                '                MsgBox("You have entered 3 incorrect password", MsgBoxStyle.Exclamation)
                '                MsgBox("ACCOUNT BLOCK" & vbNewLine & "Please contact your system administrator regarding this issue", MsgBoxStyle.Exclamation)

                '                db.ExecuteSQLQuery("update tbl_Mgmt_UserRegistration set mgmt_status = '" & "BLOCKED" & "' where mgmt_Username = '" & txtUname.Text & "'")

                '                FormName = am.getModuleName(My.Settings.Module_Tag)
                '                taskDone = "Login failed"
                '                db.sql = "Insert into tbl_Mgmt_auditTrail values('" & getAutoGenID & "','" & txtUname.Text & "','" & taskDone & "','" & FormName & "', '" & "" & "',  '" & DateTime.Today.ToShortDateString & "', '" & TimeOfDay & "' )"
                '                db.ExecuteSQLQuery(db.sql)
                '                txtPass.Clear()
                '                txtUname.Clear()
                '                txtUname.Focus()
                '                pbError.Hide()
                '                lblErrorMessage.Hide()
                '                lblErrorMessage.Text = "ERROR MESSAGE"
                '                _counter = 3
                '            End If
                '            txtPass.Focus()
                '        Else
                '            If db.checkExistence("select mgmt_Username,mgmt_password,TAG from tbl_Mgmt_UserRegistration where mgmt_Username = '" & txtUname.Text.ToUpper & "' and mgmt_Password = '" & txtPass.Text.ToUpper & "'") = True Then
                '                If db.checkExistence("select mgmt_Username,TAG from tbl_Mgmt_UserRegistration where mgmt_Username='" & txtUname.Text & "' and TAG = '0'") = True Then

                '                    FormName = am.getModuleName(My.Settings.Module_Tag)
                '                    taskDone = "User Login"
                '                    db.ExecuteSQLQuery("Update tbl_Mgmt_UserRegistration set TAG ='1' where mgmt_Username ='" & txtUname.Text & "'")
                '                    db.sql = "Insert into tbl_Mgmt_auditTrail values('" & getAutoGenID & "','" & txtUname.Text & "','" & taskDone & "','" & FormName & "', '" & "" & "',  '" & DateTime.Today.ToShortDateString & "', '" & TimeOfDay & "' )"
                '                    db.ExecuteSQLQuery(db.sql)
                '                    Me.Hide()
                '                    _frmMain.Show()
                '                    _frmMain.SplitContainer1.Panel2.Controls.Clear()
                '                    _frmDashboard.TopLevel = False
                '                    _frmDashboard.Parent = _frmMain.SplitContainer1.Panel2
                '                    _frmDashboard.Dock = DockStyle.Fill
                '                    _frmDashboard.Show()

                '                    My.Settings.LastUserNameLogin = txtUname.Text
                '                    My.Settings.LastPasswordLogin = txtPass.Text

                '                    txtPass.Text = ""
                '                    txtUname.Text = ""
                '                    _counter = 3
                '                    'audit trail
                '                Else
                '                    If MsgBox("User is currently logged from a different Computer, the previous session will close if you continue?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then

                '                        FormName = am.getModuleName(My.Settings.Module_Tag)
                '                        taskDone = "User Login"
                '                        db.ExecuteSQLQuery("Update tbl_Mgmt_UserRegistration set TAG ='0' where mgmt_Username ='" & txtUname.Text & "'")
                '                        db.sql = "Insert into tbl_Mgmt_auditTrail values('" & getAutoGenID & "','" & txtUname.Text & "','" & taskDone & "','" & FormName & "', '" & "" & "',  '" & DateTime.Today.ToShortDateString & "', '" & TimeOfDay & "' )"
                '                        db.ExecuteSQLQuery(db.sql)
                '                        Me.Hide()
                '                        _frmMain.SplitContainer1.Panel2.Controls.Clear()
                '                        _frmMain.Show()
                '                        _frmDashboard.TopLevel = False
                '                        _frmDashboard.Parent = _frmMain.SplitContainer1.Panel2
                '                        _frmDashboard.Dock = DockStyle.Fill
                '                        _frmDashboard.Show()

                '                        My.Settings.LastUserNameLogin = txtUname.Text
                '                        My.Settings.LastPasswordLogin = txtPass.Text

                '                        txtPass.Text = ""
                '                        txtUname.Text = ""
                '                        _counter = 3
                '                    End If

                '                    'MsgBox("User name or Password is incorrect! ", MsgBoxStyle.Information)
                '                    'FormName = am.getModuleName(My.Settings.Module_Tag)
                '                    'taskDone = "Login failed || "
                '                    'db.sql = "Insert into tbl_Mgmt_auditTrail values('" & getAutoGenID & "','" & txtUname.Text & "','" & taskDone & "','" & FormName & "', '" & "" & "',  '" & DateTime.Today.ToShortDateString & "', '" & TimeOfDay & "' )"
                '                    'db.ExecuteSQLQuery(db.sql)
                '                    'txtPass.Text = ""
                '                    'txtUname.Focus()
                '                End If
                '            Else
                '                int = 0
                '                tmr1.Enabled = True
                '                tmr1.Start()
                '                lblErrorMessage.Text = "Invalid Username or Password"
                '                _counter -= 1
                '                '2 pag may laman both
                '                If _counter = 0 Then
                '                    MsgBox("You have entered 3 incorrect password", MsgBoxStyle.Exclamation)
                '                    If db.checkExistence("Select mgmt_Username from tbl_mgmt_UserRegistration where mgmt_Username = '" & txtUname.Text & "'") Then
                '                        MsgBox("ACCOUNT BLOCK" & vbNewLine & "Please contact your system administrator regarding this issue", MsgBoxStyle.Exclamation)

                '                        db.ExecuteSQLQuery("update tbl_Mgmt_UserRegistration set mgmt_status = '" & "BLOCKED" & "' where mgmt_Username = '" & txtUname.Text & "'")

                '                        FormName = am.getModuleName(My.Settings.Module_Tag)
                '                        taskDone = "Login failed"
                '                        db.sql = "Insert into tbl_Mgmt_auditTrail values('" & getAutoGenID & "','" & txtUname.Text & "','" & taskDone & "','" & FormName & "', '" & "" & "',  '" & DateTime.Today.ToShortDateString & "', '" & TimeOfDay & "' )"
                '                        db.ExecuteSQLQuery(db.sql)
                '                        txtPass.Clear()
                '                        txtUname.Clear()
                '                        txtUname.Focus()
                '                        pbError.Hide()
                '                        lblErrorMessage.Hide()
                '                        lblErrorMessage.Text = "ERROR MESSAGE"
                '                        _counter = 3
                '                    Else
                '                        MsgBox("invalid username or password.", MsgBoxStyle.Information)
                '                        _counter = 3
                '                    End If
                '                End If
                '        End If
                '        End If
                '    End If
                'End If
            End If
        Catch ex As Exception
            MsgBox("Invalid username or password.", MsgBoxStyle.Information)
        End Try
    End Sub

    Private Sub ButtonX2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click

        _frmViewReports.rptView.Dispose()
        Application.Exit()

    End Sub

    Private Sub chkRememberMe_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkRemember.CheckedChanged
        RememberMeTgla()
    End Sub


    Private Sub llblForgot_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llblForgot.LinkClicked
        _frmBlock.Show()
        _frmForgotPass.ShowDialog()
        '_frmForgotPass.TopMost = True
        _frmBlock.Close()
    End Sub

    Private Sub tmr1_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tmr1.Tick
        If int = 0 Then
            pbError.Visible = True
            lblErrorMessage.Visible = True
            int = 1
            Exit Sub
        End If

        If int = 1 Then
            pbError.Visible = False
            lblErrorMessage.Visible = False
            int = 2
            Exit Sub
        End If

        If int = 2 Then
            pbError.Visible = True
            lblErrorMessage.Visible = True
            int = 3
            Exit Sub
        End If

        If int = 3 Then
            pbError.Visible = False
            lblErrorMessage.Visible = False
            int = 4
            Exit Sub
        End If

        If int = 4 Then
            pbError.Visible = True
            lblErrorMessage.Visible = True
            int = 5
            Exit Sub
        End If
    End Sub

    Private Sub txtPass_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPass.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.ChrW(Keys.Enter) Then
            ButtonX3.PerformClick()
        End If
    End Sub


End Class