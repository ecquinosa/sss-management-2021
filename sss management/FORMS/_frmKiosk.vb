Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports CrystalDecisions.Windows.Forms
Public Class _frmKiosk
    Dim dt As DataTable
    Dim cryRpt As New ReportDocument
    Dim db As New ConnectionString
    Dim FormName As String
    Dim am As New auditModule
    Dim getAutoGenID As String
    Dim taskDone As String
    Dim Status As String
    Dim isVpn As String
    Dim saveUpdate As Integer
#Region "Miscellaneous"
    Public Sub hideall()
        lanvip.Visible = False
        kip.Visible = False
        kname.Visible = False
        kbranch.Visible = False
        kcluster.Visible = False
        kgroup.Visible = False
        ksnum.Visible = False
        ohrs.Visible = False
        eHrs.Visible = False
        stats.Visible = False
    End Sub
    Public Sub update1()
        disablall()
        clearAll()
        btnAdd.Enabled = True
        btnEdit.Enabled = True
        btnCancel.Enabled = False
        btnSave.Enabled = False
        rbLAN.Checked = False
        rbVPN.Checked = False
        LoadData()
    End Sub
    Public Sub disablall()
        btnSave.Enabled = False
        btnClear.Enabled = False
        btnAdd.Enabled = True
        btnEdit.Enabled = True
        txtmgmtID.Enabled = False
        txtKName.Enabled = False
        txtKSerNum.Enabled = False
        txtIpAdd.Enabled = False
        cbCluster.Enabled = False
        cbBranch.Enabled = False
        cbGroup.Enabled = False
        txtStats.Enabled = False
        cbstart.Enabled = False
        rbLAN.Enabled = False
        rbVPN.Enabled = False
        cbstart1.Enabled = False
    End Sub
    Public Sub enableall()

        txtmgmtID.Enabled = True
        txtKName.Enabled = True
        txtKSerNum.Enabled = True
        txtIpAdd.Enabled = True
        cbCluster.Enabled = True
        cbBranch.Enabled = True
        cbGroup.Enabled = True
        cbstart.Enabled = True
        cbstart1.Enabled = True
        rbLAN.Enabled = True
        rbVPN.Enabled = True
    End Sub
    Public Sub edit()
        dgvKiosk.Enabled = False
        btnEdit.Enabled = True
        btnAdd.Enabled = False
        btnSave.Enabled = True
        btnClear.Enabled = True
        btnCancel.Enabled = True
        txtStats.Enabled = True
    End Sub
    Public Sub add()
        enableall()
        clearAll()
        dgvKiosk.Enabled = False
        btnAdd.Enabled = False
        btnEdit.Enabled = False
        btnSave.Enabled = True
        btnClear.Enabled = True
        btnCancel.Enabled = True
        txtStats.Enabled = True
        txtStats.Text = "INACTIVE"
    End Sub
    Public Sub save()
        clearAll()
        disablall()
        btnAdd.Enabled = True
        btnSave.Enabled = False
        btnClear.Enabled = False
        btnCancel.Enabled = False
        btnEdit.Enabled = True
        txtStats.Enabled = True
        LoadData()
    End Sub
    Public Sub clearAll()
        txtmgmtID.Text = Nothing
        txtKName.Text = Nothing
        txtKSerNum.Text = Nothing
        txtIpAdd.Text = Nothing
        cbCluster.Text = Nothing
        cbBranch.Text = Nothing
        cbGroup.Text = Nothing
        txtStats.Text = Nothing
        cbstart.Text = Nothing
        cbstart1.Text = Nothing
        isVpn = Nothing
    End Sub
    Public Sub cancel()
        rbLAN.Checked = False
        rbVPN.Checked = False
        clearAll()
        btnEdit.Enabled = False
        btnCancel.Enabled = False
        btnSave.Enabled = False
        btnClear.Enabled = False
        btnAdd.Enabled = True
        btnEdit.Enabled = True
        txtmgmtID.Enabled = False
        txtKName.Enabled = False
        txtKSerNum.Enabled = False
        txtIpAdd.Enabled = False
        cbCluster.Enabled = False
        cbBranch.Enabled = False
        cbGroup.Enabled = False
        txtStats.Enabled = False
        cbstart.Enabled = False
        cbstart1.Enabled = False
        isVpn = Nothing
    End Sub
    Public Sub LoadData()
        db.FillDataGridView("Select KIOSK_ID as 'Host ID',SERIALNUM as 'Serial No.',KIOSK_NM as 'Host Name',BRANCH_IP as 'IP Address',SSINFOTERMGROUP.GROUP_NM as 'Group',SSINFOTERMCLSTR.CLSTR_NM as 'Division',BRANCH_NM as 'Branch',ENCODE_DT as 'Date created',case when TAG = 1 then 'ACTIVE' when TAG = 0 then 'INACTIVE' else 'DEACTIVE' end as [Status] from SSINFOTERMKIOSK inner join SSINFOTERMBR on SSINFOTERMKIOSK.BRANCH_CD = SSINFOTERMBR.BRANCH_CD inner join SSINFOTERMGROUP on SSINFOTERMKIOSK.DIVSN = SSINFOTERMGROUP.GROUP_CD inner join SSINFOTERMCLSTR on SSINFOTERMKIOSK.CLSTR = SSINFOTERMCLSTR.CLSTR_CD ORDER BY KIOSK_ID", dgvKiosk)
    End Sub
    Public Sub clear1()
        If txtmgmtID.Text.Contains(Text) = True Then
            txtKName.Text = Nothing
            txtKSerNum.Text = Nothing
            txtIpAdd.Text = Nothing
            cbCluster.Text = Nothing
            cbBranch.Text = Nothing
            cbGroup.Text = Nothing
            txtStats.Text = Nothing
            rbLAN.Checked = False
            rbVPN.Checked = False
        End If
    End Sub
    Public Sub fillFields()
        Try

            Dim useridExist As String = dgvKiosk.CurrentRow.Cells(0).Value

            Dim getGroup As String = db.putSingleValue("Select SSINFOTERMGROUP.GROUP_NM from SSINFOTERMKIOSK inner join SSINFOTERMGROUP on SSINFOTERMKIOSK.DIVSN = SSINFOTERMGROUP.GROUP_CD where KIOSK_ID = '" & useridExist & "'")
            Dim getCluster As String = db.putSingleValue("select SSINFOTERMCLSTR.CLSTR_NM from SSINFOTERMKIOSK inner join SSINFOTERMCLSTR on SSINFOTERMKIOSK.CLSTR = SSINFOTERMCLSTR.CLSTR_CD where KIOSK_ID = '" & useridExist & "'")
            Dim getBranch As String = db.putSingleValue("select BRANCH_NM from SSINFOTERMKIOSK inner join SSINFOTERMBR on SSINFOTERMKIOSK.BRANCH_CD = SSINFOTERMBR.BRANCH_CD where KIOSK_ID = '" & useridExist & "'")

            txtmgmtID.Text = db.putSingleValue("select KIOSK_ID from SSINFOTERMKIOSK where KIOSK_ID = '" & useridExist & "'")
            txtKName.Text = db.putSingleValue("select KIOSK_NM from SSINFOTERMKIOSK where KIOSK_ID = '" & useridExist & "'")
            txtIpAdd.Text = db.putSingleValue("select BRANCH_IP from SSINFOTERMKIOSK where KIOSK_ID = '" & useridExist & "'")
            cbGroup.Text = getGroup
            cbCluster.Text = getCluster
            cbBranch.Text = getBranch
            txtKSerNum.Text = db.putSingleValue("Select SERIALNUM from SSINFOTERMKIOSK where KIOSK_ID = '" & useridExist & "'")
            Dim a As String
            a = db.putSingleValue("Select TAG FROM SSINFOTERMKIOSK where KIOSK_ID = '" & useridExist & "'")
            Select Case a

                Case 1
                    txtStats.Text = "ACTIVE"

                Case 0
                    txtStats.Text = "INACTIVE"

                Case Else
                    txtStats.Text = "DEACTIVATED"

            End Select
            Dim b As String = db.putSingleValue("select isVPN from SSINFOTERMKIOSK where KIOSK_ID = '" & useridExist & "'")
            Select Case b
                Case "True"
                    rbVPN.Checked = True
                Case "False"
                    rbLAN.Checked = True
            End Select
            cbstart.Text = db.putSingleValue("Select OP_START FROM SSINFOTERMKIOSK WHERE KIOSK_ID = '" & useridExist & "'")
            cbstart1.Text = db.putSingleValue("Select OP_END FROM SSINFOTERMKIOSK WHERE KIOSK_ID = '" & useridExist & "'")

        Catch ex As Exception

        End Try

    End Sub
    Public Sub clearall1()

    End Sub
#End Region

    Private Sub _frmKiosk_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        My.Settings.Module_Tag = 5
        getAutoGenID = am.getAutoID()
        db.fillComboBox(db.ExecuteSQLQuery("select GROUP_NM from SSINFOTERMGROUP"), cbGroup)
        db.fillComboBox(db.ExecuteSQLQuery("select CLSTR_NM from SSINFOTERMCLSTR"), cbCluster)
        db.fillComboBox(db.ExecuteSQLQuery("select BRANCH_NM from SSINFOTERMBR"), cbBranch)
        LoadData()
    End Sub

    Private Sub ButtonX1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub

    Private Sub ButtonX2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.MinimizeBox = True
    End Sub

    Private Sub ButtonX3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.MaximizeBox = True
    End Sub

    Public Function GetIP() As String

        Dim strIPAddress As String = System.Net.Dns.GetHostByName(System.Net.Dns.GetHostName).AddressList(0).ToString()
        Return strIPAddress
    End Function

    Private Sub btnSave_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click

        Select Case saveUpdate
            Case 0
                Dim a As Integer = 0

                If txtKName.Text = Nothing Then
                    a = a - 1
                    txtKName.Focus()
                    kname.Visible = True
                Else
                    a = a + 1
                    kname.Visible = False
                End If
                If cbBranch.Text = Nothing Then
                    a = a - 1
                    cbBranch.Focus()
                    kbranch.Visible = True
                Else
                    a = a + 1
                    kbranch.Visible = False
                End If
                If cbCluster.Text = Nothing Then
                    a = a - 1
                    cbCluster.Focus()
                    kcluster.Visible = True
                Else
                    a = a + 1
                    kcluster.Visible = False
                End If
                If cbGroup.Text = Nothing Then
                    a = a - 1
                    cbGroup.Focus()
                    kgroup.Visible = True
                Else
                    a = a + 1
                    kgroup.Visible = False
                End If
                If txtKSerNum.Text = Nothing Then
                    a = a - 1
                    txtKSerNum.Focus()
                    ksnum.Visible = True
                Else
                    a = a + 1
                    ksnum.Visible = False
                End If

                If txtStats.Text = Nothing Then
                    a = a - 1
                    stats.Visible = True
                    txtStats.Focus()
                Else
                    a = a + 1
                    stats.Visible = False
                End If
                If cbstart.Text = Nothing Then
                    a = a - 1
                    ohrs.Visible = True
                Else
                    a = a + 1
                    ohrs.Visible = False
                End If

                If cbstart1.Text = Nothing Then
                    a = a - 1
                    eHrs.Visible = True
                Else
                    a = a + 1
                    eHrs.Visible = False
                End If
                If txtIpAdd.Text = Nothing Then
                    a = a - 1
                    kip.Visible = True
                Else
                    a = a + 1
                    kip.Visible = False
                End If
                If isVpn = Nothing Then
                    a = a - 1
                    lanvip.Visible = True
                Else
                    a = a + 1
                    lanvip.Visible = False
                End If
                If a = 10 Then

                    'If db.checkExistence("Select KIOSK_ID, KIOSK_NM,BRANCH_IP FROM SSINFOTERMKIOSK Where KIOSK_ID ='" & txtmgmtID.Text & "' or KIOSK_NM = '" & txtKName.Text & "' or BRANCH_IP = '" & txtIpAdd.Text & "'") = True Then
                    '    MsgBox("terminal is already registered.")
                    'Else
                    If MsgBox("Click OK to save kiosk details.", MsgBoxStyle.OkCancel) = MsgBoxResult.Ok Then
                        Select Case txtStats.Text

                            Case "ACTIVE"
                                Status = "1"
                            Case "INACTIVE"
                                Status = "0"
                            Case "DEACTIVE"
                                Status = "2"
                        End Select
                        If rbLAN.Checked = True Then
                            isVpn = "False"
                        ElseIf rbVPN.Checked = True Then
                            isVpn = "True"
                        End If
                        Dim OPSTART As Date = cbstart.Text
                        Dim OPEND As Date = cbstart1.Text
                        Dim todayDate As Date = DateTime.Today.ToShortDateString & "  " & TimeOfDay
                        Dim getBranch As String = db.putSingleValue("select BRANCH_CD from SSINFOTERMBR where BRANCH_NM = '" & cbBranch.Text & "'")
                        Dim getCluster As String = db.putSingleValue("select CLSTR_CD from SSINFOTERMCLSTR where CLSTR_NM = '" & cbCluster.Text & "'")
                        Dim getGroup As String = db.putSingleValue("select GROUP_CD from SSINFOTERMGROUP where GROUP_NM = '" & cbGroup.Text & "'")
                        Dim IdleTime As String = db.putSingleValue("select DISTINCT IDLE_TIME FROM SSINFOTERMKIOSK")
                        db.ExecuteSQLQuery("insert into SSINFOTERMKIOSK (KIOSK_ID,BRANCH_CD,KIOSK_NM,BRANCH_IP,IDLE_TIME,STATUS,ENCODE_DT,CLSTR,DIVSN,TAG,OP_START,OP_END,SERIALNUM,isVPN) values('" & txtmgmtID.Text & "','" & getBranch & "','" & txtKName.Text & "','" & txtIpAdd.Text & _
                      "','" & IdleTime & "','0','" & todayDate & "','" & getCluster & "', '" & getGroup & "','" & Status & "','" & OPSTART & "','" & OPEND & "','" & txtKSerNum.Text & "','" & isVpn & "')")
                        'db.ExecuteSQLQuery(db.sql)

                        MsgBox("Successfully saved.")
                        FormName = am.getModuleName(My.Settings.Module_Tag)
                        taskDone = "Add New Kiosk"
                        db.sql = "Insert into tbl_Mgmt_auditTrail values('" & getAutoGenID & "','" & My.Settings.LastUserNameLogin & "','" & taskDone & "','" & FormName & "', '" & "" & "',  '" & DateTime.Today.ToShortDateString & "', '" & TimeOfDay & "' )"
                        db.ExecuteSQLQuery(db.sql)
                        save()
                        txtmgmtID.ReadOnly = True
                        dgvKiosk.Enabled = True
                    End If
                End If
                'Else
                MsgBox("Please fill-out required fields.")
                'End If

            Case 1
                Dim a As Integer = 0

                If txtKName.Text = Nothing Then
                    a = a - 1
                    txtKName.Focus()
                    kname.Visible = True
                Else
                    a = a + 1
                    kname.Visible = False
                End If
                If cbBranch.Text = Nothing Then
                    a = a - 1
                    cbBranch.Focus()
                    kbranch.Visible = True
                Else
                    a = a + 1
                    kbranch.Visible = False
                End If
                If cbCluster.Text = Nothing Then
                    a = a - 1
                    cbCluster.Focus()
                    kcluster.Visible = True
                Else
                    a = a + 1
                    kcluster.Visible = False
                End If
                If cbGroup.Text = Nothing Then
                    a = a - 1
                    cbGroup.Focus()
                    kgroup.Visible = True
                Else
                    a = a + 1
                    kgroup.Visible = False
                End If
                If txtKSerNum.Text = Nothing Then
                    a = a - 1
                    txtKSerNum.Focus()
                    ksnum.Visible = True
                Else
                    a = a + 1
                    ksnum.Visible = False
                End If
                If cbstart.Text = Nothing Then
                    a = a - 1
                    ohrs.Visible = True
                Else
                    a = a + 1
                    ohrs.Visible = False
                End If

                If cbstart1.Text = Nothing Then
                    a = a - 1
                    eHrs.Visible = True
                Else
                    a = a + 1
                    eHrs.Visible = False
                End If

                If txtStats.Text = Nothing Then
                    a = a - 1
                    stats.Visible = True
                    txtStats.Focus()
                Else
                    a = a + 1
                    stats.Visible = False
                End If
                If isVpn = Nothing Then
                    a = a - 1
                    lanvip.Visible = True
                Else
                    a = a + 1
                    lanvip.Visible = False
                End If
                If a = 9 Then

                    If db.checkExistence("select KIOSK_ID from SSINFOTERMKIOSK where KIOSK_ID = '" & txtmgmtID.Text & "'") = True Then
                        If MsgBox("Click OK to update kiosk information.", MsgBoxStyle.OkCancel) = MsgBoxResult.Ok Then
                            Select Case txtStats.Text

                                Case "ACTIVE"
                                    Status = "1"
                                Case "INACTIVE"
                                    Status = "0"
                                Case "DEACTIVE"
                                    Status = "2"
                            End Select
                            If rbLAN.Checked = True Then
                                isVpn = "False"
                            ElseIf rbVPN.Checked = True Then
                                isVpn = "True"
                            End If
                            Dim OPSTART As Date = cbstart.Text
                            Dim OPEND As Date = cbstart1.Text
                            Dim getBranchCD As String = db.putSingleValue("select BRANCH_CD from SSINFOTERMBR where BRANCH_NM = '" & cbBranch.Text & "'")
                            Dim getClusterCD As String = db.putSingleValue("select CLSTR_CD from SSINFOTERMCLSTR where CLSTR_NM = '" & cbCluster.Text & "'")
                            Dim getGroupCD As String = db.putSingleValue("select GROUP_CD from SSINFOTERMGROUP where GROUP_NM = '" & cbGroup.Text & "'")

                            db.ExecuteSQLQuery("Update SSINFOTERMKIOSK set KIOSK_NM = '" & txtKName.Text & "', BRANCH_CD = '" & getBranchCD & _
                                         "', BRANCH_IP = '" & txtIpAdd.Text & "', CLSTR = '" & getClusterCD & "', DIVSN = '" & getGroupCD & "', OP_START = '" & OPSTART & "', OP_END = '" & OPEND & "',SERIALNUM = '" & txtKSerNum.Text & "',TAG = '" & Status & "',isVPN = '" & isVpn & "'  where KIOSK_ID = '" & txtmgmtID.Text & "' ")
                            MsgBox("Successfully updated.", MsgBoxStyle.Information, "Registration")
                            LoadData()
                            update1()
                            dgvKiosk.Enabled = True
                            txtmgmtID.ReadOnly = False
                        End If
                    Else
                        MsgBox("Terminal doesn't exist.")
                    End If
                End If


        End Select



    End Sub
    Private Sub ButtonX1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        saveUpdate = 0
        txtmgmtID.ReadOnly = False
        add()
        'txtmgmtID.Text = db.putSingleNumber("select max(KIOSK_ID) from SSINFOTERMKIOSK")
        'If txtmgmtID.Text = "0" Or txtmgmtID.Text = "1" Then
        '    txtmgmtID.Text = "1000"
        'Else
        '    txtmgmtID.Text = txtmgmtID.Text + 1
        'End If
    End Sub
    Private Sub dgvKiosk_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dgvKiosk.DoubleClick
        fillFields()
        btnEdit.Enabled = True
    End Sub


    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        Select Case cbFields.Text

            Case "HOST ID"
                db.FillDataGridView("Select KIOSK_ID as 'Host ID',SERIALNUM as 'Serial No.',KIOSK_NM as 'Host Name',BRANCH_IP as 'IP Address',SSINFOTERMGROUP.GROUP_NM as 'Group',SSINFOTERMCLSTR.CLSTR_NM as 'Division',BRANCH_NM as 'Branch',ENCODE_DT as 'Date created',case when TAG = 1 then 'ACTIVE' when TAG = 0 then 'INACTIVE' else 'DEACTIVE' end as [Status] from SSINFOTERMKIOSK inner join SSINFOTERMBR on SSINFOTERMKIOSK.BRANCH_CD = SSINFOTERMBR.BRANCH_CD inner join SSINFOTERMGROUP on SSINFOTERMKIOSK.DIVSN = SSINFOTERMGROUP.GROUP_CD inner join SSINFOTERMCLSTR on SSINFOTERMKIOSK.CLSTR = SSINFOTERMCLSTR.CLSTR_CD    where KIOSK_ID like '%" & txtSearch.Text & "%' ", dgvKiosk)
            Case "HOST IP ADDRESS"
                db.FillDataGridView("Select KIOSK_ID as 'Host ID',SERIALNUM as 'Serial No.',KIOSK_NM as 'Host Name',BRANCH_IP as 'IP Address',SSINFOTERMGROUP.GROUP_NM as 'Group',SSINFOTERMCLSTR.CLSTR_NM as 'Division',BRANCH_NM as 'Branch',ENCODE_DT as 'Date created',case when TAG = 1 then 'ACTIVE' when TAG = 0 then 'INACTIVE' else 'DEACTIVE' end as [Status] from SSINFOTERMKIOSK inner join SSINFOTERMBR on SSINFOTERMKIOSK.BRANCH_CD = SSINFOTERMBR.BRANCH_CD inner join SSINFOTERMGROUP on SSINFOTERMKIOSK.DIVSN = SSINFOTERMGROUP.GROUP_CD inner join SSINFOTERMCLSTR on SSINFOTERMKIOSK.CLSTR = SSINFOTERMCLSTR.CLSTR_CD where BRANCH_IP like '%" & txtSearch.Text & "%' ", dgvKiosk)
            Case "HOST SERIAL NUMBER"
                db.FillDataGridView("Select KIOSK_ID as 'Host ID',SERIALNUM as 'Serial No.',KIOSK_NM as 'Host Name',BRANCH_IP as 'IP Address',SSINFOTERMGROUP.GROUP_NM as 'Group',SSINFOTERMCLSTR.CLSTR_NM as 'Division',BRANCH_NM as 'Branch',ENCODE_DT as 'Date created',case when TAG = 1 then 'ACTIVE' when TAG = 0 then 'INACTIVE' else 'DEACTIVE' end as [Status] from SSINFOTERMKIOSK inner join SSINFOTERMBR on SSINFOTERMKIOSK.BRANCH_CD = SSINFOTERMBR.BRANCH_CD inner join SSINFOTERMGROUP on SSINFOTERMKIOSK.DIVSN = SSINFOTERMGROUP.GROUP_CD inner join SSINFOTERMCLSTR on SSINFOTERMKIOSK.CLSTR = SSINFOTERMCLSTR.CLSTR_CD where SERIALNUM like '%" & txtSearch.Text & "%' ", dgvKiosk)
            Case "HOST NAME"
                db.FillDataGridView("Select KIOSK_ID as 'Host ID',SERIALNUM as 'Serial No.',KIOSK_NM as 'Host Name',BRANCH_IP as 'IP Address',SSINFOTERMGROUP.GROUP_NM as 'Group',SSINFOTERMCLSTR.CLSTR_NM as 'Division',BRANCH_NM as 'Branch',ENCODE_DT as 'Date created',case when TAG = 1 then 'ACTIVE' when TAG = 0 then 'INACTIVE' else 'DEACTIVE' end as [Status] from SSINFOTERMKIOSK inner join SSINFOTERMBR on SSINFOTERMKIOSK.BRANCH_CD = SSINFOTERMBR.BRANCH_CD inner join SSINFOTERMGROUP on SSINFOTERMKIOSK.DIVSN = SSINFOTERMGROUP.GROUP_CD inner join SSINFOTERMCLSTR on SSINFOTERMKIOSK.CLSTR = SSINFOTERMCLSTR.CLSTR_CD where KIOSK_NM like '%" & txtSearch.Text & "%' ", dgvKiosk)
            Case "BRANCH"
                db.FillDataGridView("Select KIOSK_ID as 'Host ID',SERIALNUM as 'Serial No.',KIOSK_NM as 'Host Name',BRANCH_IP as 'IP Address',SSINFOTERMGROUP.GROUP_NM as 'Group',SSINFOTERMCLSTR.CLSTR_NM as 'Division',BRANCH_NM as 'Branch',ENCODE_DT as 'Date created',case when TAG = 1 then 'ACTIVE' when TAG = 0 then 'INACTIVE' else 'DEACTIVE' end as [Status] from SSINFOTERMKIOSK inner join SSINFOTERMBR on SSINFOTERMKIOSK.BRANCH_CD = SSINFOTERMBR.BRANCH_CD inner join SSINFOTERMGROUP on SSINFOTERMKIOSK.DIVSN = SSINFOTERMGROUP.GROUP_CD inner join SSINFOTERMCLSTR on SSINFOTERMKIOSK.CLSTR = SSINFOTERMCLSTR.CLSTR_CD where SSINFOTERMBR.BRANCH_NM like '%" & txtSearch.Text & "%' ORDER BY KIOSK_ID", dgvKiosk)
            Case "DIVISION"
                db.FillDataGridView("Select KIOSK_ID as 'Host ID',SERIALNUM as 'Serial No.',KIOSK_NM as 'Host Name',BRANCH_IP as 'IP Address',SSINFOTERMGROUP.GROUP_NM as 'Group',SSINFOTERMCLSTR.CLSTR_NM as 'Division',BRANCH_NM as 'Branch',ENCODE_DT as 'Date created',case when TAG = 1 then 'ACTIVE' when TAG = 0 then 'INACTIVE' else 'DEACTIVE' end as [Status] from SSINFOTERMKIOSK inner join SSINFOTERMBR on SSINFOTERMKIOSK.BRANCH_CD = SSINFOTERMBR.BRANCH_CD inner join SSINFOTERMGROUP on SSINFOTERMKIOSK.DIVSN = SSINFOTERMGROUP.GROUP_CD inner join SSINFOTERMCLSTR on SSINFOTERMKIOSK.CLSTR = SSINFOTERMCLSTR.CLSTR_CD where SSINFOTERMCLSTR.CLSTR_NM like '%" & txtSearch.Text & "%' ORDER BY KIOSK_ID", dgvKiosk)
            Case "GROUP"
                db.FillDataGridView("Select KIOSK_ID as 'Host ID',SERIALNUM as 'Serial No.',KIOSK_NM as 'Host Name',BRANCH_IP as 'IP Address',SSINFOTERMGROUP.GROUP_NM as 'Group',SSINFOTERMCLSTR.CLSTR_NM as 'Division',BRANCH_NM as 'Branch',ENCODE_DT as 'Date created',case when TAG = 1 then 'ACTIVE' when TAG = 0 then 'INACTIVE' else 'DEACTIVE' end as [Status] from SSINFOTERMKIOSK inner join SSINFOTERMBR on SSINFOTERMKIOSK.BRANCH_CD = SSINFOTERMBR.BRANCH_CD inner join SSINFOTERMGROUP on SSINFOTERMKIOSK.DIVSN = SSINFOTERMGROUP.GROUP_CD inner join SSINFOTERMCLSTR on SSINFOTERMKIOSK.CLSTR = SSINFOTERMCLSTR.CLSTR_CD where SSINFOTERMGROUP.GROUP_NM like '%" & txtSearch.Text & "%' ORDER BY KIOSK_ID ", dgvKiosk)
            Case "Show all"
                db.FillDataGridView("Select KIOSK_ID as 'Host ID',SERIALNUM as 'Serial No.',KIOSK_NM as 'Host Name',BRANCH_IP as 'IP Address',SSINFOTERMGROUP.GROUP_NM as 'Group',SSINFOTERMCLSTR.CLSTR_NM as 'Division',BRANCH_NM as 'Branch',ENCODE_DT as 'Date created',case when TAG = 1 then 'ACTIVE' when TAG = 0 then 'INACTIVE' else 'DEACTIVE' end as [Status] from SSINFOTERMKIOSK inner join SSINFOTERMBR on SSINFOTERMKIOSK.BRANCH_CD = SSINFOTERMBR.BRANCH_CD inner join SSINFOTERMGROUP on SSINFOTERMKIOSK.DIVSN = SSINFOTERMGROUP.GROUP_CD inner join SSINFOTERMCLSTR on SSINFOTERMKIOSK.CLSTR = SSINFOTERMCLSTR.CLSTR_CD where KIOSK_ID like '%" & txtSearch.Text & "%' or  BRANCH_IP like '%" & txtSearch.Text & "%' or  SERIALNUM like '%" & txtSearch.Text & "%' or  KIOSK_NM like '%" & txtSearch.Text & "%' or SSINFOTERMBR.BRANCH_NM like '%" & txtSearch.Text & "%' or  SSINFOTERMCLSTR.CLSTR_NM like '%" & txtSearch.Text & "%' ORDER BY KIOSK_ID", dgvKiosk)

        End Select
    End Sub

    Private Sub cbFields_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbFields.SelectedIndexChanged
        If cbFields.Text = "Show all" Then
            LoadData()
        End If
    End Sub
    Private Sub btnClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClear.Click
        clear1()
    End Sub

    Private Sub btnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        saveUpdate = 1
        txtmgmtID.ReadOnly = True
        fillFields()
        edit()
        enableall()
        btnEdit.Enabled = False
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        cancel()
        hideall()
        dgvKiosk.Enabled = True
        btnEdit.Enabled = False
        txtmgmtID.ReadOnly = False
    End Sub

    Private Sub cbCluster_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbCluster.Click
        If cbGroup.Text <> "" Then
            db.fillComboBox(db.ExecuteSQLQuery("SELECT CLSTR_NM FROM SSINFOTERMCLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSINFOTERMCLSTR.GROUP_CD where SSINFOTERMGROUP.GROUP_NM = '" & cbGroup.Text & "'"), cbCluster)
        Else
            db.fillComboBox(db.ExecuteSQLQuery("Select CLSTR_NM FROM SSINFOTERMCLSTR"), cbCluster)
        End If
    End Sub

    Private Sub cbBranch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbBranch.Click
        If cbCluster.Text <> "" Then
            db.fillComboBox(db.ExecuteSQLQuery("SELECT BRANCH_NM FROM SSINFOTERMBR INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMBR.CLSTR_CD where SSINFOTERMCLSTR.CLSTR_NM = '" & cbCluster.Text & "'"), cbBranch)
        ElseIf cbGroup.Text <> "" Then
            db.fillComboBox(db.ExecuteSQLQuery("SELECT BRANCH_NM FROM SSINFOTERMBR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSINFOTERMBR.GROUP_CD where SSINFOTERMGROUP.GROUP_NM = '" & cbGroup.Text & "'"), cbBranch)
        Else
            db.fillComboBox(db.ExecuteSQLQuery("SELECT BRANCH_NM FROM SSINFOTERMBR"), cbBranch)
        End If
    End Sub

    Private Sub cbGroup_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbGroup.SelectedIndexChanged
        db.fillComboBox(db.ExecuteSQLQuery("SELECT CLSTR_NM FROM SSINFOTERMCLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSINFOTERMCLSTR.GROUP_CD where SSINFOTERMGROUP.GROUP_NM = '" & cbGroup.Text & "'"), cbCluster)
        cbBranch.Text = Nothing
    End Sub

    Private Sub cbCluster_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbCluster.SelectedIndexChanged
        db.fillComboBox(db.ExecuteSQLQuery("SELECT BRANCH_NM FROM SSINFOTERMBR INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMBR.CLSTR_CD where SSINFOTERMCLSTR.CLSTR_NM = '" & cbCluster.Text & "'"), cbBranch)
    End Sub
    Private Sub rbLAN_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbLAN.CheckedChanged
        isVpn = "false"
    End Sub

    Private Sub rbVPN_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbVPN.CheckedChanged
        isVpn = "true"
    End Sub

    Private Sub ButtonX1_Click_3(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonX1.Click
        _frmBlock.Show()
        _frmAddBranch.ShowDialog()
    End Sub
    Public Sub operationFileTrans()
        Dim getFileName As String
        Dim newview As New CrystalReportViewer

        Try
            getFileName = "LIST OF ACITVE AND INACTIVE SSIT"
            dt = db.ExecuteSQLQuery("select KIOSK_ID,BRANCH_IP,KIOSK_NM,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,CASE WHEN TAG = 0 then 'INACTIVE' WHEN TAG = 1 then 'ACTIVE' else 'DEACTIVE' END AS [TAG] from SSINFOTERMKIOSK INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSINFOTERMKIOSK.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMKIOSK.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSINFOTERMKIOSK.DIVSN where isVPN = 'false' ORDER BY SSINFOTERMKIOSK.KIOSK_ID")
            Dim rpt As New KioskList
            cryRpt = rpt
            '  cryRpt.Refresh()
            cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
            newview.ReportSource = cryRpt
            Dim getdate As String = Date.Today.ToShortDateString
            Dim gettime As String = TimeOfDay
            getdate = getdate.Replace("/", "-")
            gettime = gettime.Replace(":", ".")
            Dim test As New FolderBrowserDialog
            Dim filepath As String
            If test.ShowDialog = DialogResult.OK Then

                filepath = test.SelectedPath
                filepath = filepath & "\"

                Dim CrExportOptions As ExportOptions
                Dim CrDiskFileDestinationOptions As New  _
       DiskFileDestinationOptions()
                Dim CrFormatTypeOptions As New ExcelFormatOptions
                CrFormatTypeOptions.ExcelUseConstantColumnWidth = False
                CrFormatTypeOptions.ShowGridLines = True
                CrDiskFileDestinationOptions.DiskFileName = _
                                            filepath & getFileName & " " & getdate & " " & gettime & ".xls"
                CrExportOptions = cryRpt.ExportOptions
                With CrExportOptions
                    .ExportDestinationType = ExportDestinationType.DiskFile
                    .ExportFormatType = ExportFormatType.Excel
                    .DestinationOptions = CrDiskFileDestinationOptions
                    .FormatOptions = CrFormatTypeOptions
                End With
                cryRpt.Export()
                MsgBox("Export complete.", MsgBoxStyle.Information)
            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try


    End Sub
    Private Sub ButtonX2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonX2.Click
        operationFileTrans()
    End Sub

    Private Sub txtmgmtID_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtmgmtID.KeyPress
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If
    End Sub
End Class