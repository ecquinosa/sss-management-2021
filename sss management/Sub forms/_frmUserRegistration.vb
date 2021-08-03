Public Class _frmUserRegistration
    Dim db As New ConnectionString
    Dim tag As Integer
    Dim userID As String
    Dim FormName As String
    Dim am As New auditModule
    Dim getAutoGenID As String
    Dim taskDone As String
    Dim plainText As String
    Dim cipherText As String
    Dim passPhrase As String
    Dim saltValue As String
    Dim hashAlgorithm As String
    Dim passwordIterations As Integer
    Dim initVector As String
    Dim keySize As Integer
    Dim saveupdate As Integer
    Public Sub LoadData()

        db.FillDataGridView("select mgmt_Username as 'Username',mgmt_Password as 'Password', mgmt_EmployeeNumber as 'Employee Number',mgmt_userID as 'User ID', mgmt_Firstname as 'First Name', mgmt_Middlename as 'Middle Name',  mgmt_Lastname as 'Last Name', mgmt_UserType as 'User Type',mgmt_status as 'Status' from tbl_Mgmt_UserRegistration", dvgUsers)

    End Sub
    Private Sub reset1()
        txtAnswer.Text = ""
        txtEmpNum.Text = ""
        txtFirstName.Text = ""
        txtLastName.Text = ""
        txtMiddleName.Text = ""
        txtSearch.Text = ""
        txtStatus.Text = ""
        txtUname.Text = ""
        cbUserType.Text = ""
        cbSecurityQuestion.Text = ""
    End Sub
    Private Sub disableall()
        txtUserID.Enabled = False
        txtEmpNum.Enabled = False
        txtFirstName.Enabled = False
        txtLastName.Enabled = False
        txtMiddleName.Enabled = False
        cbSecurityQuestion.Enabled = False
        cbUserType.Enabled = False
        txtUname.Enabled = False
        txtAnswer.Enabled = False
        txtPassword.Enabled = False
    End Sub
    Private Sub enableall()
        'txtUserID.Enabled = True
        txtEmpNum.Enabled = True
        txtEmpNum.ReadOnly = False
        txtFirstName.Enabled = True
        txtLastName.Enabled = True
        txtMiddleName.Enabled = True
        cbSecurityQuestion.Enabled = True
        cbUserType.Enabled = True
        txtUname.Enabled = True
        txtAnswer.Enabled = True
        txtPassword.Enabled = True
    End Sub
    Private Sub hideall()
        valen.Visible = False
        valfn.Visible = False
        valmn.Visible = False
        valln.Visible = False
        valut.Visible = False
        valun.Visible = False
        valsa.Visible = False
        valsq.Visible = False
    End Sub
    Private Sub resetall()
        cbUserType.Text = Nothing
        cbSecurityQuestion.Text = Nothing


        txtAnswer.Text = ""
        txtEmpNum.Text = ""
        txtFirstName.Text = ""
        txtLastName.Text = ""
        txtMiddleName.Text = ""
        txtSearch.Text = ""
        txtStatus.Text = ""
        txtUname.Text = ""
        txtPassword.Text = ""
        txtUserID.Text = ""
    End Sub
    Public Sub btnAdd1()
        enableall()
        txtEmpNum.Focus()
        txtEmpNum.ReadOnly = False
        resetall()
        txtUserID.Text = userID
        txtStatus.Text = "INACTIVE"
        btnSave.Enabled = True
        btnadd.Enabled = False
        btnEdit.Enabled = False
        btnReset.Enabled = True
        btnCancel.Enabled = True
        btnRP.Enabled = False
        txtstatus.Enabled = True
        txtPassword.Text = "password"
        txtPassword.ReadOnly = True
      
    End Sub
    Public Sub btnSave1()
        resetall()
        hideall()
        disableall()
        btnSave.Enabled = False
        btnadd.Enabled = True
        btnEdit.Enabled = True
        btnReset.Enabled = False
        btnRP.Enabled = False
        btnCancel.Enabled = False
        txtPassword.ReadOnly = False
        txtstatus.Enabled = False
    End Sub
    Public Sub btnUpdate1()
        resetall()
        txtEmpNum.Focus()
        hideall()
        disableall()
        btnSave.Enabled = False
        btnCancel.Enabled = False
        btnEdit.Enabled = True
        btnadd.Enabled = True
        btnReset.Enabled = False
        btnRP.Enabled = False
        txtstatus.Enabled = False
        btnRP.Enabled = False

    End Sub
    Public Sub btnCancel1()
        dvgUsers.Enabled = True
        btnEdit.Enabled = False
        resetall()
        disableall()
        hideall()
        btnEdit.Enabled = False
        txtEmpNum.ReadOnly = False
        btnadd.Enabled = True
        btnSave.Enabled = False
        btnCancel.Enabled = False
        btnReset.Enabled = False
        btnRP.Enabled = False
        txtstatus.Enabled = False
        txtUserID.Text = ""
        txtPassword.ReadOnly = False
    End Sub
    Public Sub btnEdit1()
        dvgUsers.Enabled = False
        btnSave.Enabled = True
        txtEmpNum.ReadOnly = True
        txtPassword.ReadOnly = False
        btnadd.Enabled = False
        btnReset.Enabled = True
        btnCancel.Enabled = True
        btnRP.Enabled = False
        txtstatus.Enabled = True
        btnRP.Enabled = True
        enableall()
    End Sub
    Public Sub fillFields()
        Try
            saveupdate = 1
            Dim useridExist As String = dvgUsers.CurrentRow.Cells(3).Value
            cbSecurityQuestion.Text = db.putSingleValue("select mgmt_SecurityQuestion from tbl_Mgmt_UserRegistration where mgmt_UserID = '" & useridExist & "'")
            cbUserType.Text = db.putSingleValue("select mgmt_UserType from tbl_Mgmt_UserRegistration where mgmt_UserID = '" & useridExist & "'")
            txtUserID.Text = db.putSingleValue("select mgmt_UserID from tbl_Mgmt_UserRegistration where mgmt_UserID = '" & useridExist & "'")
            txtEmpNum.Text = db.putSingleValue("select mgmt_EmployeeNumber from tbl_Mgmt_UserRegistration where mgmt_UserID = '" & useridExist & "'")
            txtFirstName.Text = db.putSingleValue("select mgmt_Firstname from tbl_Mgmt_UserRegistration where mgmt_UserID = '" & useridExist & "'")
            txtMiddleName.Text = db.putSingleValue("select mgmt_Middlename from tbl_Mgmt_UserRegistration where mgmt_UserID = '" & useridExist & "'")
            txtLastName.Text = db.putSingleValue("select mgmt_Lastname from tbl_Mgmt_UserRegistration where mgmt_UserID = '" & useridExist & "'")
            txtUname.Text = db.putSingleValue("select mgmt_Username from tbl_Mgmt_UserRegistration where mgmt_UserID = '" & useridExist & "'")
            cipherText = db.putSingleValue("select mgmt_Password from tbl_Mgmt_UserRegistration where mgmt_UserID = '" & useridExist & "'")
            txtAnswer.Text = db.putSingleValue("select mgmt_SecurityAnswer from tbl_Mgmt_UserRegistration where mgmt_UserID = '" & useridExist & "'")
            txtStatus.Text = db.putSingleValue("select mgmt_status from tbl_Mgmt_UserRegistration where mgmt_UserID = '" & useridExist & "'")
            tag = db.putSingleValue("Select TAG from tbl_mgmt_UserRegistration where mgmt_UserID = '" & useridExist & "'")

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
            txtPassword.Text = plainText
        Catch ex As Exception

        End Try

    End Sub
    Public Sub btnResetPass1()
        txtEmpNum.Focus()
        hideall()
        disableall()
        btnCancel.Enabled = False
        btnEdit.Enabled = True
        btnadd.Enabled = True
        btnReset.Enabled = True
        btnRP.Enabled = False
        txtstatus.Enabled = False
        btnRP.Enabled = False
        btnReset.Enabled = False
    End Sub
    Private Sub _frmUserSettings_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        My.Settings.Module_Tag = 4
        getAutoGenID = am.getAutoID()
        LoadData()
        txtEmpNum.Focus()
        dvgUsers.Columns("Password").Visible = False
        'password = "0000"
        'con.fillComboBox(con.ExecuteSQLQuery("SELECT position FROM tbl_mgmt_position "), cbpostion)
        db.fillComboBox(db.ExecuteSQLQuery("SELECT user_type FROM tbl_mgmt_usertype "), cbUserType)
        'btnSave.Enabled = False
        'btnUpdate.Enabled = False
        'txtstatus.Enabled = False

    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Select Case saveupdate
            Case 0
                Dim a As Integer = 0

                If txtEmpNum.Text = Nothing Then
                    txtEmpNum.Focus()
                    valen.Visible = True
                    a = a - 1
                Else
                    a = a + 1
                    valen.Visible = False
                End If
                If txtFirstName.Text = Nothing Then
                    txtFirstName.Focus()
                    valfn.Visible = True
                    a = a - 1
                Else
                    a = a + 1
                    valfn.Visible = False
                End If
                If txtMiddleName.Text = Nothing Then
                    txtFirstName.Focus()
                    valmn.Visible = True
                    a = a - 1
                Else
                    a = a + 1
                    valmn.Visible = False
                End If

                If txtLastName.Text = Nothing Then
                    txtLastName.Focus()
                    valln.Visible = True
                    a = a - 1
                Else
                    a = a + 1
                    valln.Visible = False
                End If

                If txtUname.Text = Nothing Then
                    txtUname.Focus()
                    valun.Visible = True
                    a = a - 1
                Else
                    a = a + 1
                    valun.Visible = False
                End If
                If txtPassword.Text = Nothing Then
                    Label14.Visible = True
                    a = a - 1
                Else
                    a = a + 1
                    Label14.Visible = False
                End If

                If txtAnswer.Text = Nothing Then
                    txtAnswer.Focus()
                    valsa.Visible = True
                    a = a - 1
                Else
                    a = a + 1
                    valsa.Visible = False
                End If
                If cbUserType.Text = Nothing Then
                    cbUserType.Focus()
                    valut.Visible = True
                    a = a - 1
                Else
                    a = a + 1
                    valut.Visible = False
                End If
                If cbSecurityQuestion.Text = Nothing Then
                    cbSecurityQuestion.Focus()
                    valsq.Visible = True
                    a = a - 1
                Else
                    a = a + 1
                    valsq.Visible = False
                End If
                If txtUname.Text = "ADMIN" Then
                    MsgBox("Invalid username.", MsgBoxStyle.Information)
                End If

                If txtUname.TextLength < 4 Then
                    valun.Visible = True
                    a = a - 1
                Else
                    valun.Visible = False
                    a = a + 1
                End If

                If a = 10 Then
                    If db.checkExistence("Select mgmt_Username from tbl_Mgmt_UserRegistration where mgmt_Username = '" & txtUname.Text & "'") Then
                        MsgBox("Username is already existed.", MsgBoxStyle.Information, "Registration")
                    Else
                        If db.checkExistence("select mgmt_EmployeeNumber from tbl_Mgmt_UserRegistration where mgmt_EmployeeNumber = '" & txtEmpNum.Text & "'") = True Or db.checkExistence("Select mgmt_UserID from tbl_Mgmt_UserRegistration where mgmt_UserID ='" & txtUserID.Text & "'") = True Then
                            MsgBox("Employee number is already existed.", MsgBoxStyle.Information, "Registration")
                        Else

                            plainText = My.Settings.firstPassword    ' original plaintext
                            passPhrase = "Pas5pr@se"        ' can be any string
                            saltValue = "s@1tValue"         ' can be any string
                            hashAlgorithm = "SHA1"          ' can be "MD5"
                            passwordIterations = 2          ' can be any number
                            initVector = "@1B2c3D4e5F6g7H8" ' must be 16 bytes
                            keySize = 256                   ' can be 192 or 128
                            'MsgBox(String.Format("Plaintext : {0}", plainText))


                            cipherText = RijndaelSimple.Encrypt _
                            ( _
                                plainText, _
                                passPhrase, _
                                saltValue, _
                                hashAlgorithm, _
                                passwordIterations, _
                                initVector, _
                                keySize _
                            )
                            db.ExecuteSQLQuery("INSERT INTO tbl_Mgmt_UserRegistration (mgmt_username,mgmt_password,mgmt_employeenumber, mgmt_UserID, mgmt_FirstName, mgmt_MiddleName, mgmt_LastName, mgmt_UserType, mgmt_SecurityAnswer,mgmt_Status,mgmt_dateCreated, mgmt_SecurityQuestion,TAG,Attempt,FirstLog) values ('" _
                                                & txtUname.Text & "', '" & cipherText & "', '" & txtEmpNum.Text & "', '" & txtUserID.Text & "', '" & txtFirstName.Text & "', '" & txtMiddleName.Text & "', '" & txtLastName.Text & "', '" & cbUserType.Text & "','" & txtAnswer.Text & _
                                                "', 'ACTIVE', '" & DateAndTime.Now & "','" & cbSecurityQuestion.Text & "','0','0','0')")
                            MsgBox("Successfully saved.", MsgBoxStyle.Information, "Registration")
                            resetall()
                            btnSave1()

                            FormName = am.getModuleName(My.Settings.Module_Tag)
                            taskDone = "Add New User"
                            db.sql = "Insert into tbl_Mgmt_auditTrail values('" & getAutoGenID & "','" & My.Settings.LastUserNameLogin & "','" & taskDone & "','" & FormName & "', '" & "" & "',  '" & DateTime.Today.ToShortDateString & "', '" & TimeOfDay & "' )"
                            db.ExecuteSQLQuery(db.sql)
                            disableall()
                            dvgUsers.Enabled = True
                        End If
                    End If
                Else
                    MsgBox("Please fill-out required fields.")
                End If

                LoadData()
            Case 1
                txtEmpNum.ReadOnly = True
                txtUname.ReadOnly = True
                Dim a As Integer = 0

                If txtFirstName.Text = Nothing Then
                    txtFirstName.Focus()
                    valfn.Visible = True
                    a = a - 1
                Else
                    a = a + 1
                    valfn.Visible = False
                End If
                If txtMiddleName.Text = Nothing Then
                    txtFirstName.Focus()
                    valmn.Visible = True
                    a = a - 1
                Else
                    a = a + 1
                    valmn.Visible = False
                End If

                If txtLastName.Text = Nothing Then
                    txtLastName.Focus()
                    valln.Visible = True
                    a = a - 1
                Else
                    a = a + 1
                    valln.Visible = False
                End If
                If txtPassword.Text = Nothing Then
                    Label14.Visible = True
                    a = a - 1
                Else
                    a = a + 1
                    Label14.Visible = False
                End If

                If txtAnswer.Text = Nothing Then
                    txtAnswer.Focus()
                    valsa.Visible = True
                    a = a - 1
                Else
                    a = a + 1
                    valsa.Visible = False
                End If
                If cbUserType.Text = Nothing Then
                    cbUserType.Focus()
                    valut.Visible = True
                    a = a - 1
                Else
                    a = a + 1
                    valut.Visible = False
                End If
                If cbSecurityQuestion.Text = Nothing Then
                    cbSecurityQuestion.Focus()
                    valsq.Visible = True
                    a = a - 1
                Else
                    a = a + 1
                    valsq.Visible = False
                End If

                If a = 7 Then
                    If db.checkExistence("select mgmt_userID from tbl_Mgmt_UserRegistration where mgmt_userID = '" & txtUserID.Text & "'") = True Then
                        If MsgBox("Do you want to update this user?", MsgBoxStyle.OkCancel) = MsgBoxResult.Ok Then
                            Select Case txtstatus.Text

                                Case "ACTIVE"
                                    tag = 1
                                Case "INACTIVE"
                                    tag = 0
                                Case "DEACTIVE"
                                    tag = 2
                            End Select
                            plainText = txtPassword.Text    ' original plaintext
                            passPhrase = "Pas5pr@se"        ' can be any string
                            saltValue = "s@1tValue"         ' can be any string
                            hashAlgorithm = "SHA1"          ' can be "MD5"
                            passwordIterations = 2          ' can be any number
                            initVector = "@1B2c3D4e5F6g7H8" ' must be 16 bytes
                            keySize = 256                   ' can be 192 or 128
                            'MsgBox(String.Format("Plaintext : {0}", plainText))


                            cipherText = RijndaelSimple.Encrypt _
                            ( _
                                plainText, _
                                passPhrase, _
                                saltValue, _
                                hashAlgorithm, _
                                passwordIterations, _
                                initVector, _
                                keySize _
                            )
                            db.ExecuteSQLQuery("Update tbl_Mgmt_UserRegistration set mgmt_UserID = '" & txtUserID.Text & "', mgmt_Firstname = '" & txtFirstName.Text & "', mgmt_Middlename = '" & txtMiddleName.Text & _
                                         "', mgmt_Lastname = '" & txtLastName.Text & "', mgmt_UserType = '" & cbUserType.Text & "', mgmt_SecurityAnswer = '" & txtAnswer.Text & _
                                         "', mgmt_status = '" & txtstatus.Text & "',firstLog = '" & tag & "', mgmt_Password = '" & cipherText & "', mgmt_SecurityQuestion ='" & cbSecurityQuestion.Text & "' where mgmt_UserID = '" & txtUserID.Text & "' ")
                            MsgBox("Successfully updated.", MsgBoxStyle.Information, "Registration")


                            FormName = am.getModuleName(My.Settings.Module_Tag)
                            taskDone = "Update User"
                            db.sql = "Insert into tbl_Mgmt_auditTrail values('" & getAutoGenID & "','" & My.Settings.LastUserNameLogin & "','" & taskDone & "','" & FormName & "', '" & "" & "',  '" & DateTime.Today.ToShortDateString & "', '" & TimeOfDay & "' )"
                            db.ExecuteSQLQuery(db.sql)


                            taskDone = "changed the user role of " & txtFirstName.Text & " to " & cbUserType.Text
                            db.sql = "Insert into tbl_Mgmt_auditTrail values('" & getAutoGenID & "','" & My.Settings.LastUserNameLogin & "','" & taskDone & "','" & FormName & "', '" & "" & "',  '" & DateTime.Today.ToShortDateString & "', '" & TimeOfDay & "' )"
                            db.ExecuteSQLQuery(db.sql)
                            resetall()
                            btnUpdate1()
                            dvgUsers.Enabled = True
                        End If
                    Else
                        MsgBox("User doesn't exist.")
                    End If
                Else
                    MsgBox("Please fill-out required fields.")
                End If

                LoadData()
        End Select
       
    End Sub

    Private Sub btnReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReset.Click
        If Text.Contains(txtUserID.Text) = True Or txtPassword.Text.Contains(txtPassword.Text) = True Then
            txtAnswer.Text = ""
            txtEmpNum.Text = ""
            txtFirstName.Text = ""
            txtLastName.Text = ""
            txtMiddleName.Text = ""
            txtSearch.Text = ""
            txtStatus.Text = ""
            txtUname.Text = ""
            cbUserType.Text = ""
            cbSecurityQuestion.Text = ""
        Else
            resetall()
        End If

    End Sub
    Private Sub dvgUsers_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dvgUsers.DoubleClick
        fillFields()
        btnEdit.Enabled = True
    End Sub
    Private Sub cbFields_DropDown(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbFields.DropDown
        db.fillComboBox(db.ExecuteSQLQuery("select field_search from tbl_mgmt_fields where field_search not LIKE '%User ID%'"), cbFields)
    End Sub

    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        Select Case cbFields.Text

            Case "First Name"
                db.FillDataGridView("select mgmt_Username as 'Username',mgmt_Password as 'Password', mgmt_EmployeeNumber as 'Employee Number',mgmt_userID as 'User ID', mgmt_Firstname as 'First Name', mgmt_Middlename as 'Middle Name',  mgmt_Lastname as 'Last Name', mgmt_UserType as 'User Type',mgmt_status as 'Status' from tbl_Mgmt_UserRegistration where mgmt_Firstname like '%" & txtSearch.Text & "%' ", dvgUsers)

            Case "Last Name"
                db.FillDataGridView("select mgmt_Username as 'Username',mgmt_Password as 'Password', mgmt_EmployeeNumber as 'Employee Number',mgmt_userID as 'User ID', mgmt_Firstname as 'First Name', mgmt_Middlename as 'Middle Name',  mgmt_Lastname as 'Last Name', mgmt_UserType as 'User Type',mgmt_status as 'Status' from tbl_Mgmt_UserRegistration where mgmt_Lastname like '%" & txtSearch.Text & "%' ", dvgUsers)

            Case "Username"
                db.FillDataGridView("select mgmt_Username as 'Username',mgmt_Password as 'Password', mgmt_EmployeeNumber as 'Employee Number',mgmt_userID as 'User ID', mgmt_Firstname as 'First Name', mgmt_Middlename as 'Middle Name',  mgmt_Lastname as 'Last Name', mgmt_UserType as 'User Type',mgmt_status as 'Status' from tbl_Mgmt_UserRegistration where mgmt_Username like '%" & txtSearch.Text & "%' ", dvgUsers)

            Case "Employee Number"
                db.FillDataGridView("select mgmt_Username as 'Username',mgmt_Password as 'Password', mgmt_EmployeeNumber as 'Employee Number',mgmt_userID as 'User ID', mgmt_Firstname as 'First Name', mgmt_Middlename as 'Middle Name',  mgmt_Lastname as 'Last Name', mgmt_UserType as 'User Type',mgmt_status as 'Status' from tbl_Mgmt_UserRegistration where mgmt_EmployeeNumber like '%" & txtSearch.Text & "%' ", dvgUsers)

            Case "ACTIVE"
                db.FillDataGridView("select mgmt_Username as 'Username',mgmt_Password as 'Password', mgmt_EmployeeNumber as 'Employee Number',mgmt_userID as 'User ID', mgmt_Firstname as 'First Name', mgmt_Middlename as 'Middle Name',  mgmt_Lastname as 'Last Name', mgmt_UserType as 'User Type',mgmt_status as 'Status' from tbl_Mgmt_UserRegistration where mgmt_status = '" & "ACTIVE" & "' ", dvgUsers)

            Case "INACTIVE"
                db.FillDataGridView("select mgmt_Username as 'Username',mgmt_Password as 'Password', mgmt_EmployeeNumber as 'Employee Number',mgmt_userID as 'User ID', mgmt_Firstname as 'First Name', mgmt_Middlename as 'Middle Name',  mgmt_Lastname as 'Last Name', mgmt_UserType as 'User Type',mgmt_status as 'Status' from tbl_Mgmt_UserRegistration where mgmt_status = '" & "INACTIVE" & "' ", dvgUsers)

            Case "Position"
                db.FillDataGridView("select mgmt_Username as 'Username',mgmt_Password as 'Password', mgmt_EmployeeNumber as 'Employee Number',mgmt_userID as 'User ID', mgmt_Firstname as 'First Name', mgmt_Middlename as 'Middle Name',  mgmt_Lastname as 'Last Name', mgmt_UserType as 'User Type',mgmt_status as 'Status' from tbl_Mgmt_UserRegistration where mgmt_Position like '%" & txtSearch.Text & "%' ", dvgUsers)

            Case "Show All"
                db.FillDataGridView("select mgmt_Username as 'Username',mgmt_Password as 'Password', mgmt_EmployeeNumber as 'Employee Number',mgmt_userID as 'User ID', mgmt_Firstname as 'First Name', mgmt_Middlename as 'Middle Name',  mgmt_Lastname as 'Last Name', mgmt_UserType as 'User Type',mgmt_status as 'Status' from tbl_Mgmt_UserRegistration", dvgUsers)
                txtSearch.Text = ""
            Case Else
                MsgBox("No record found.")
        End Select
    End Sub
    Private Sub btnadd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnadd.Click
        saveupdate = 0
        If db.checkExistence("select mgmt_UserID from tbl_Mgmt_UserRegistration where mgmt_UserID = '" & txtUserID.Text & "'") = True Then
            userID = db.GenRegID(My.Settings.firstUserID)
            txtUserID.Text = userID
        Else
            userID = db.GenRegID(My.Settings.firstUserID)
            txtUserID.Text = userID
        End If
        dvgUsers.Enabled = False
        reset1()
        enableall()
        btnAdd1()
    End Sub
    Private Sub btnRP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRP.Click

        If db.checkExistence("Select mgmt_EmployeeNumber from tbl_Mgmt_UserRegistration where mgmt_EmployeeNumber = '" & txtEmpNum.Text & "'") = True Then
            plainText = My.Settings.firstPassword    ' original plaintext
            passPhrase = "Pas5pr@se"        ' can be any string
            saltValue = "s@1tValue"         ' can be any string
            hashAlgorithm = "SHA1"          ' can be "MD5"
            passwordIterations = 2          ' can be any number
            initVector = "@1B2c3D4e5F6g7H8" ' must be 16 bytes
            keySize = 256                   ' can be 192 or 128
            'MsgBox(String.Format("Plaintext : {0}", plainText))


            cipherText = RijndaelSimple.Encrypt _
            ( _
                plainText, _
                passPhrase, _
                saltValue, _
                hashAlgorithm, _
                passwordIterations, _
                initVector, _
                keySize _
            )
            db.ExecuteSQLQuery("Update tbl_Mgmt_UserRegistration set mgmt_Password = '" & cipherText & "' where mgmt_EmployeeNumber = '" & txtEmpNum.Text & "'")
            MsgBox("Password is successfully reset.")
            LoadData()
            resetall()
            btnResetPass1()

            FormName = am.getModuleName(My.Settings.Module_Tag)
            taskDone = "Reset Password"
            db.sql = "Insert into tbl_Mgmt_auditTrail values('" & getAutoGenID & "','" & My.Settings.LastUserNameLogin & "','" & taskDone & "','" & FormName & "', '" & "" & "',  '" & DateTime.Today.ToShortDateString & "', '" & TimeOfDay & "' )"
            db.ExecuteSQLQuery(db.sql)
        Else
            MsgBox("Password reset failed.")
        End If
    End Sub
    Private Sub btnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        saveupdate = 1
        fillFields()
        btnEdit1()
        enableall()
        btnEdit.Enabled = False
    End Sub
    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        btnCancel1()
    End Sub

    Private Sub cbFields_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbFields.SelectedIndexChanged
        If cbFields.Text = "Show All" = True Then
            txtSearch.Text = ""
            LoadData()
        End If

        Select Case cbFields.Text

            Case "ACTIVE"
                txtSearch.Text = ""
                db.FillDataGridView("select mgmt_Username as 'Username',mgmt_Password as 'Password', mgmt_EmployeeNumber as 'Employee Number',mgmt_userID as 'User ID', mgmt_Firstname as 'First Name', mgmt_Middlename as 'Middle Name',  mgmt_Lastname as 'Last Name', mgmt_UserType as 'User Type',mgmt_status as 'Status' from tbl_Mgmt_UserRegistration where mgmt_status = '" & "ACTIVE" & "' ", dvgUsers)

            Case "INACTIVE"

                txtSearch.Text = ""
                db.FillDataGridView("select mgmt_Username as 'Username',mgmt_Password as 'Password', mgmt_EmployeeNumber as 'Employee Number',mgmt_userID as 'User ID', mgmt_Firstname as 'First Name', mgmt_Middlename as 'Middle Name',  mgmt_Lastname as 'Last Name', mgmt_UserType as 'User Type',mgmt_status as 'Status' from tbl_Mgmt_UserRegistration where mgmt_status = '" & "INACTIVE" & "' ", dvgUsers)
        End Select
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        _frmBlock.Show()
        _frmUserType.ShowDialog()
        '_frmUserType.TopMost = True
        _frmBlock.Close()
    End Sub

    Private Sub cbUserType_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbUserType.Click
        db.fillComboBox(db.ExecuteSQLQuery("SELECT user_type FROM tbl_mgmt_usertype "), cbUserType)
    End Sub
End Class