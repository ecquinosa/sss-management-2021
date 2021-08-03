Public Class _frmForgotPass
    Dim db As New ConnectionString
    Dim a As Integer
    Dim plainText As String
    Dim cipherText As String

    Dim passPhrase As String
    Dim saltValue As String
    Dim hashAlgorithm As String
    Dim passwordIterations As Integer
    Dim initVector As String
    Dim keySize As Integer
    Public Sub clearall()
        txtAns.Enabled = True
        txtPass1.Enabled = True
        txtPass2.Enabled = True
        txtUser.Enabled = True
        txtque.Enabled = True
        btnSave.Enabled = True
        btnSubmit.Enabled = True
        btnans.Enabled = True
        txtAns.Text = Nothing
        txtque.Text = Nothing
        txtUser.Text = Nothing
        txtPass1.Text = Nothing
        txtPass2.Text = Nothing
        lblError1.Visible = False
        lblerror2.Visible = False
        lblerror3.Visible = False
    End Sub
    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Dim a As Integer = 0
        If txtPass1.Text = Nothing Then
            a = a - 1
            MsgBox("Password is empty.", MsgBoxStyle.Exclamation)
        Else
            a = a + 1

        End If

        If txtPass2.Text = Nothing Then
            a = a - 1
            MsgBox("Please re-type password.", MsgBoxStyle.Exclamation)

        Else
            a = a + 1

        End If

        If txtPass1.Text <> txtPass2.Text Then
            MsgBox("Password didn't match.", MsgBoxStyle.Exclamation)
            a = a - 1
      
        Else
            a = a + 1

        End If
        If txtPass1.TextLength < 6 Then
            MsgBox("Password should be at least 6 characters.", MsgBoxStyle.Exclamation)
            a = a - 1
        Else
            a = a + 1
        End If
        If a = 4 Then
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
            db.ExecuteSQLQuery("Update tbl_Mgmt_UserRegistration set mgmt_Password = '" & cipherText & "' where mgmt_Username = '" & txtUser.Text & "'")
            MsgBox("You have successfully changed your password.", MsgBoxStyle.Information)
            txtPass1.Enabled = False
            txtPass2.Enabled = False
            btnSave.Enabled = False
            Me.Hide()
            _frmLogin.Show()
        End If
    End Sub

    Private Sub Panel3_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs)


    End Sub

    Private Sub _frmForgotPass_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        clearall()
        Me.Height = 143
        Panel1.Size = New System.Drawing.Size(460, 90)
    End Sub

    Private Sub Panel1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel1.Paint

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSubmit.Click
        If txtUser.Text = Nothing Then
            MsgBox("Username is empty.", MsgBoxStyle.Information)
        Else
            If db.checkExistence("Select * from tbl_Mgmt_UserRegistration where mgmt_Username = '" & txtUser.Text & "'") = True Then
                Dim a As Date = db.putSingleValue("Select isnull(ForgotLocked,0) from tbl_Mgmt_UserRegistration where mgmt_Username = '" & txtUser.Text & "' ")
                If a <> "1/1/1900" Then
                    If a > Date.Now Then
                        MsgBox("Your account has been locked. Please try again tomorrow.", MsgBoxStyle.Information)
                    Else
                        db.ExecuteSQLQuery("UPDATE tbl_Mgmt_UserRegistration set ForgotLocked = '1900-01-01 00:00:00.000' where mgmt_Username = '" & txtUser.Text & "'")
                    End If
                Else
                    txtque.Text = db.putSingleValue("Select mgmt_SecurityQuestion from tbl_Mgmt_UserRegistration where mgmt_Username = '" & txtUser.Text & "'")
                    Me.Height = 259
                    Panel1.Size = New System.Drawing.Size(460, 218)
                    txtUser.Enabled = False
                    btnSubmit.Enabled = False
                    lblError1.Visible = False
                    txtAns.Enabled = True
                    btnans.Enabled = True
                End If

        Else
            MsgBox("Invalid Username", MsgBoxStyle.Exclamation)
            txtUser.Enabled = True
            btnSubmit.Enabled = True
            End If
        End If
    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnans.Click


        If txtque.Text = Nothing Then
            MsgBox("Security Answer is empty.", MsgBoxStyle.Exclamation)
        Else
            If db.checkExistence("Select mgmt_Username, mgmt_SecurityAnswer, mgmt_SecurityQuestion from tbl_Mgmt_UserRegistration where mgmt_Username = '" & txtUser.Text & "' and mgmt_SecurityAnswer = '" & txtAns.Text & "' and mgmt_SecurityQuestion ='" & txtque.Text & "'") = True Then
                Me.Height = 397
                Panel1.Size = New System.Drawing.Size(460, 346)
                btnans.Enabled = False
                txtAns.Enabled = False
                txtque.Enabled = False
                txtPass1.Enabled = True
                txtPass2.Enabled = True
                btnSave.Enabled = True
            Else
                If a >= 5 Then
                    Dim ForgotDay As Date = DateAdd(DateInterval.Day, 1, DateAndTime.Now)
                    db.ExecuteSQLQuery("UPDATE tbl_Mgmt_UserRegistration set ForgotLocked = '" & ForgotDay & "' where mgmt_Username = '" & txtUser.Text & "'")
                    MsgBox("Your account has been locked. Please try again tomorrow.", MsgBoxStyle.Information)
                    Me.Dispose()
                    Me.Close()
                Else
                    a = a + 1
                    MsgBox("Invalid security answer.", MsgBoxStyle.Exclamation)
                    btnans.Enabled = True
                    txtAns.Enabled = True
                    txtque.Enabled = True
                End If
            End If
        End If
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Dispose()
        Me.Close()

    End Sub
End Class