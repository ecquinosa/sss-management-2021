Public Class _frmUserType

    Dim db As New ConnectionString
    Dim FormName As String
    Dim am As New auditModule
    Dim getAutoGenID As String
    Dim taskDone As String
    Public a As String
    Private Sub loadUserTypes()
        db.FillListView(db.ExecuteSQLQuery("SELECT ID as 'ID',user_type as 'USER TYPE', description as 'DESCRIPTION' FROM tbl_mgmt_usertype "), lvList)
    End Sub
    Private Sub enableText()
        txtDesc.Enabled = True
        txtUsertType.Enabled = True

    End Sub

    Private Sub disableText()
        txtDesc.Enabled = False
        txtUsertType.Enabled = False

    End Sub
    Private Sub clearText()
        txtDesc.Text = ""
        txtRecordID.Text = ""
        txtUsertType.Text = ""

    End Sub
    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If btnSave.Tag = "Add" Then
            btnReset.Enabled = True
            txtRecordID.Text = db.putSingleNumber("select max(id) from tbl_mgmt_usertype")
            txtRecordID.Text += 1
            enableText()
            txtDesc.Text = Nothing
            txtUsertType.Text = Nothing
            lvList.Enabled = False
            btnSave.Text = "Save"
            btnSave.Tag = "Save"
            btnRemove.Enabled = False

            btnUpdate.Enabled = False
        ElseIf btnSave.Tag = "Save" Then
            If txtUsertType.Text = "" Then
                MsgBox("Please input a user type.", MsgBoxStyle.Information, "User Type")
            ElseIf txtDesc.Text = "" Then
                MsgBox("Please input a description.", MsgBoxStyle.Information, "User Type")
            ElseIf txtDesc.TextLength > 50 Then
                MsgBox("Description must be equal or less than 50 characters only.", MsgBoxStyle.Information, "User Type")
            ElseIf db.checkExistence("Select user_type from tbl_mgmt_usertype where user_type = '" & txtUsertType.Text & "'") Then
                MsgBox("User type already exists.", MsgBoxStyle.Exclamation)
            Else
                db.ExecuteSQLQuery("INSERT INTO tbl_mgmt_usertype (user_type, description) VALUES ('" & txtUsertType.Text & "' , '" & txtDesc.Text & "')")
                MsgBox("Successfully added new user type.", MsgBoxStyle.Information, "User Type")
                loadUserTypes()
                btnSave.Text = "Add"
                btnSave.Tag = "Add"
                clearText()
                disableText()
                lvList.Enabled = True
                btnReset.Enabled = False
                btnRemove.Enabled = True
                btnUpdate.Enabled = True
                FormName = am.getModuleName(My.Settings.Module_Tag)
                taskDone = "New User Type Added"
                db.sql = "Insert into tbl_Mgmt_auditTrail values('" & getAutoGenID & "','" & My.Settings.LastUserNameLogin & "','" & taskDone & "','" & FormName & "', '" & "" & "',  '" & DateTime.Today.ToShortDateString & "', '" & TimeOfDay & "' )"
                db.ExecuteSQLQuery(db.sql)
            End If
        End If

    End Sub

    Private Sub _frmUserType_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        My.Settings.Module_Tag = 1.2
        getAutoGenID = am.getAutoID()
        btnRemove.Enabled = True
        btnUpdate.Enabled = True
        btnSave.Tag = "Add"
        btnUpdate.Tag = "Edit"
        btnUpdate.Text = "Edit"
        loadUserTypes()
        lvList.Columns(0).Width = 40
    End Sub

    Private Sub btnRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemove.Click
        If lvList.Items.Count = 0 Then
            MsgBox("There are no items to be deleted.", MsgBoxStyle.Information, "Delete User Type")
        ElseIf txtRecordID.Text = "" Then
            MsgBox("Please select a user type.", MsgBoxStyle.Information, "Delete User Type")
        Else
            If MsgBox("Do you want to remove this user type?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                db.ExecuteSQLQuery("DELETE FROM tbl_mgmt_userrole1 where user_role = '" & txtUsertType.Text & "'")
                db.ExecuteSQLQuery("DELETE FROM tbl_mgmt_usertype where id = '" & lvList.SelectedItems(0).Text & "'")
                MsgBox("Successfully deleted.", MsgBoxStyle.Information, "Remove User Type")
                loadUserTypes()
                FormName = am.getModuleName(My.Settings.Module_Tag)
                taskDone = "User Type Deleted"
                db.sql = "Insert into tbl_Mgmt_auditTrail values('" & getAutoGenID & "','" & My.Settings.LastUserNameLogin & "','" & taskDone & "','" & FormName & "', '" & "" & "',  '" & DateTime.Today.ToShortDateString & "', '" & TimeOfDay & "' )"
                db.ExecuteSQLQuery(db.sql)
                clearText()
            End If
        End If
    End Sub

    Private Sub lvList_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lvList.Click
        txtRecordID.Text = db.putSingleValue("SELECT * FROM tbl_mgmt_usertype where id = '" & lvList.SelectedItems(0).Text & "'")
        txtUsertType.Text = db.putSingleValue("SELECT user_type FROM tbl_mgmt_usertype where id = '" & lvList.SelectedItems(0).Text & "'")
        a = db.putSingleValue("SELECT user_type FROM tbl_mgmt_usertype where id = '" & lvList.SelectedItems(0).Text & "'")
        txtDesc.Text = db.putSingleValue("SELECT description FROM tbl_mgmt_usertype WHERE id = '" & lvList.SelectedItems(0).Text & "'")
    End Sub

    Private Sub ButtonX1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonX1.Click
        Me.Close()

        loadUserTypes()
        btnSave.Text = "Add"
        btnSave.Tag = "Add"
        clearText()
        disableText()
        lvList.Enabled = True
        btnRemove.Enabled = True
        btnUpdate.Enabled = True
        btnSave.Enabled = True
    End Sub

    Private Sub btnReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReset.Click
        btnReset.Enabled = False
        txtDesc.Text = ""
        txtUsertType.Text = ""
        btnSave.Text = "Add"
        btnSave.Tag = "Add"
        btnUpdate.Tag = "Edit"
        btnUpdate.Text = "Edit"
        btnUpdate.Enabled = True
        btnSave.Enabled = True
        btnRemove.Enabled = True
        txtRecordID.Text = ""
        lvList.Enabled = True
    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        If btnUpdate.Tag = "Edit" Then
            If txtRecordID.Text = "" Then
                MsgBox("Please select a user type.", MsgBoxStyle.Information, "Update User Type")
            ElseIf txtDesc.TextLength > 50 Then
                MsgBox("Description must be equal or less than 50 characters only.", MsgBoxStyle.Information, "User Type")
            Else
                btnReset.Enabled = True
                lvList.Enabled = False
                btnUpdate.Tag = "Update"
                txtUsertType.Focus()
                btnUpdate.Text = "Update"
                btnSave.Enabled = False
                btnRemove.Enabled = False
                enableText()
                txtUsertType.Focus()
                loadUserTypes()
            End If

        ElseIf btnUpdate.Tag = "Update" Then
            btnReset.Enabled = False
            db.FillListView(db.ExecuteSQLQuery("SELECT * FROM tbl_mgmt_usertype "), lvList)
            db.ExecuteSQLQuery("Delete tbl_mgmt_userrole1 where user_role ='" & a & "'")
            db.ExecuteSQLQuery("UPDATE tbl_mgmt_usertype SET user_type  = '" & txtUsertType.Text & "',description = '" & txtDesc.Text & "' WHERE id ='" & txtRecordID.Text & "'")
            MsgBox("Successfully updated.", MsgBoxStyle.Information, "Update User Type")
            clearText()
            loadUserTypes()
            disableText()
            btnUpdate.Tag = "Edit"
            btnUpdate.Text = "Edit"
            btnSave.Enabled = True
            btnRemove.Enabled = True
            lvList.Enabled = True
            FormName = am.getModuleName(My.Settings.Module_Tag)
            taskDone = "User Type Updated"
            db.sql = "Insert into tbl_Mgmt_auditTrail values('" & getAutoGenID & "','" & My.Settings.LastUserNameLogin & "','" & taskDone & "','" & FormName & "', '" & "" & "',  '" & DateTime.Today.ToShortDateString & "', '" & TimeOfDay & "' )"
            db.ExecuteSQLQuery(db.sql)
        End If
    End Sub
End Class