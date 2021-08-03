Public Class _frmChangePassLogin
    Dim db As New ConnectionString
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
    Private Sub _frmChangePassLogin_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        My.Settings.Module_Tag = 1.3
        getAutoGenID = am.getAutoID()
        Me.Height = 186
        Panel1.Size = New System.Drawing.Size(460, 125)
    End Sub
    Private Sub btnSave_Click(sender As System.Object, e As System.EventArgs) Handles btnSave.Click
        If txtPass1.Text = Nothing Then
            MsgBox("Password is empty.", MsgBoxStyle.Exclamation)
            Exit Sub
        End If
        If txtPass2.Text = Nothing Then
            MsgBox("Confirm new password is empty.", MsgBoxStyle.Exclamation)
            Exit Sub
        End If

        If txtPass1.Text <> txtPass2.Text Then
            MsgBox("Password does not match.", MsgBoxStyle.Exclamation)
            Exit Sub
        End If
        If txtPass1.TextLength < 6 Then
            MsgBox("Password should be at least 6 characters.", MsgBoxStyle.Exclamation)
            Exit Sub
        End If
        Dim a As String = db.putSingleValue("select mgmt_Password from tbl_mgmt_UserRegistration where mgmt_Username = '" & _frmLogin.txtUname.Text & "'")
        If txtPass1.Text = "password" And txtPass2.Text = "password" Then
            MsgBox("Invalid Password.")
            Exit Sub
        End If
        My.Settings.LastPasswordLogin = txtPass1.Text
        plainText = txtPass1.Text    ' original plaintext
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
        db.ExecuteSQLQuery("Update tbl_Mgmt_UserRegistration SET mgmt_Password='" & cipherText & "',FirstLog = '1',tag = '1',mgmt_Status = 'ACTIVE' where mgmt_Username='" & _frmLogin.txtUname.Text & "'")

        MsgBox("Successfully updated.", MsgBoxStyle.Information)
        My.Settings.LastUserNameLogin = _frmLogin.txtUname.Text
            FormName = am.getModuleName(My.Settings.Module_Tag)
            taskDone = "User Changed Password"
            db.sql = "Insert into tbl_Mgmt_auditTrail values('" & getAutoGenID & "','" & My.Settings.LastUserNameLogin & "','" & taskDone & "','" & FormName & "', '" & "" & "',  '" & DateTime.Today.ToShortDateString & "', '" & TimeOfDay & "' )"
            db.ExecuteSQLQuery(db.sql)
        Me.Dispose()
        Me.Close()
        _frmLogin.Hide()
        _frmMain.Show()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Dispose()
        _frmLogin.Show()
    End Sub
End Class