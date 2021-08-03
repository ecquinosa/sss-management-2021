Public Class _frmUserRole
    Dim db As New ConnectionString
    Dim userID As String
    Dim FormName As String
    Dim am As New auditModule
    Dim getAutoGenID As String
    Dim taskDone As String
    Dim ur As userRole
    Public Sub checkboxes()
        cbKioskMonitoring.Checked = False
        cbReports.Checked = False
        cbKioskMonitoring.Checked = False
        cbUserRole.Checked = False
        cbUserManagement.Checked = False
        cbDatabase.Checked = False
        cbFBV.Checked = False

        Dim mgmtMonitoring As String = db.putSingleValue("Select view_monitoring from tbl_mgmt_userrole1 where user_role = '" & cbUserType.Text & "' ")
        Dim mgmtReports As String = db.putSingleValue("Select view_reports from tbl_mgmt_userrole1 where user_role = '" & cbUserType.Text & "' ")
        Dim mgmtKiosk As String = db.putSingleValue("Select view_kioskmgmt from tbl_mgmt_userrole1 where user_role = '" & cbUserType.Text & "' ")
        Dim mgmtUser As String = db.putSingleValue("Select view_usermgmt from tbl_mgmt_userrole1 where user_role = '" & cbUserType.Text & "' ")
        Dim mgmtUserRole As String = db.putSingleValue("Select view_userrole from tbl_mgmt_userrole1 where user_role = '" & cbUserType.Text & "' ")
        Dim mgmtDB As String = db.putSingleValue("Select view_dbset from tbl_mgmt_userrole1 where user_role = '" & cbUserType.Text & "'")
        Dim mgmtFB As String = db.putSingleValue("Select view_Feedback from tbl_mgmt_userrole1 where user_role = '" & cbUserType.Text & "'")

        cbKioskMonitoring.Checked = Convert.ToString(mgmtMonitoring).Equals(Boolean.TrueString)
        cbReports.Checked = Convert.ToString(mgmtReports).Equals(Boolean.TrueString)
        cbKioskManagement.Checked = Convert.ToString(mgmtKiosk).Equals(Boolean.TrueString)
        cbUserRole.Checked = Convert.ToString(mgmtUserRole).Equals(Boolean.TrueString)
        cbUserManagement.Checked = Convert.ToString(mgmtUser).Equals(Boolean.TrueString)
        cbDatabase.Checked = Convert.ToString(mgmtDB).Equals(Boolean.TrueString)
        cbFBV.Checked = Convert.ToString(mgmtFB).Equals(Boolean.TrueString)
    End Sub
    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        _frmBlock.Show()
        _frmUserType.ShowDialog()
        '_frmUserType.TopMost = True
        _frmBlock.Close()
    End Sub

    Private Sub cbUserType_DropDown(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbUserType.DropDown
        db.fillComboBox(db.ExecuteSQLQuery("select user_type from tbl_mgmt_usertype"), cbUserType)
    End Sub

    Private Sub ButtonX1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonX1.Click

        'db.sql = "Insert into tbl_mgmt_userrole (user_role, allow_add_user) values('" _
        If cbUserType.Text = Nothing Then
            MsgBox("Please create a new user role.", MsgBoxStyle.Information)
        Else
            If db.checkExistence("select user_role from tbl_mgmt_userrole1 where user_role = '" & cbUserType.Text & "'") = True Then
                If MsgBox("Do you want to update the user role?", MsgBoxStyle.OkCancel) = MsgBoxResult.Ok Then
                    db.ExecuteSQLQuery("Update tbl_mgmt_userrole1 set user_mgmt_all = '" & cbAll.Checked & "',view_monitoring = '" & cbKioskMonitoring.Checked & "',view_kioskmgmt = '" & cbKioskManagement.Checked & "' ,view_reports = '" & cbReports.Checked & "',view_usermgmt = '" & cbUserManagement.Checked & "',view_userrole = '" & cbUserRole.Checked & "' ,view_dbset = '" & cbDatabase.Checked & "',view_Feedback = '" & cbFBV.Checked & "' where user_role = '" & cbUserType.Text & "'")
                    MsgBox("Successfully updated.")
                End If
            Else

                db.ExecuteSQLQuery("Insert into tbl_mgmt_userrole1 (user_role, user_mgmt_all, view_monitoring, view_kioskmgmt, view_reports, view_usermgmt, view_userrole, view_dbset,view_Feedback) values('" _
                 & cbUserType.Text & "','" & cbAll.Checked & "', '" & cbKioskMonitoring.Checked & "', '" & cbKioskManagement.Checked & "', '" & cbReports.Checked & "', '" & cbUserManagement.Checked & _
                 "', '" & cbUserRole.Checked & "', '" & cbDatabase.Checked & "','" & cbFBV.Checked & "')")

                MsgBox("Successfully saved.", MsgBoxStyle.Information)
                clearX()


                FormName = am.getModuleName(My.Settings.Module_Tag)
                taskDone = "New User Role Added"
                db.sql = "Insert into tbl_Mgmt_auditTrail values('" & getAutoGenID & "','" & My.Settings.LastUserNameLogin & "','" & taskDone & "','" & FormName & "', '" & "" & "',  '" & DateTime.Today.ToShortDateString & "', '" & TimeOfDay & "' )"
                db.ExecuteSQLQuery(db.sql)
            End If
        End If
    End Sub

    Public Sub clearX()

        cbUserType.Text = ""
        cbKioskManagement.Checked = False
        cbKioskMonitoring.Checked = False
        cbUserManagement.Checked = False
        cbFBV.Checked = False
        cbDatabase.Checked = False
        cbReports.Checked = False
        cbUserRole.Checked = False
        cbAll.Checked = False

    End Sub

    Private Sub ButtonX2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonX2.Click
        clearX()
    End Sub
    Private Sub _frmUserRole_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        My.Settings.Module_Tag = 9
        getAutoGenID = am.getAutoID()
    End Sub

    Private Sub cbAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbAll.CheckedChanged
        Select Case cbAll.Checked
            Case True
                cbKioskManagement.Checked = True
                cbKioskMonitoring.Checked = True
                cbUserManagement.Checked = True
                cbDatabase.Checked = True
                cbReports.Checked = True
                cbUserRole.Checked = True
                cbFBV.Checked = True
                cbKioskManagement.Enabled = False
                cbKioskMonitoring.Enabled = False
                cbUserManagement.Enabled = False
                cbDatabase.Enabled = False
                cbReports.Enabled = False
                cbUserRole.Enabled = False
                cbFBV.Enabled = False
            Case False
                cbKioskManagement.Checked = False
                cbKioskMonitoring.Checked = False
                cbUserManagement.Checked = False
                cbDatabase.Checked = False
                cbReports.Checked = False
                cbUserRole.Checked = False
                cbFBV.Checked = False
                cbKioskManagement.Enabled = True
                cbKioskMonitoring.Enabled = True
                cbUserManagement.Enabled = True
                cbDatabase.Enabled = True
                cbReports.Enabled = True
                cbUserRole.Enabled = True
                cbFBV.Enabled = True
        End Select
      
    End Sub

    Private Sub cbUserType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbUserType.SelectedIndexChanged
        checkboxes()
    End Sub
End Class