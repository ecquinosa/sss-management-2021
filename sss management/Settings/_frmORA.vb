Public Class _frmORA
    Dim db As New ConnectionString
    Dim er As Integer
    Dim up As New UploadFTP
    Dim oracle As New Constring
    Dim plainText As String
    Dim cipherText As String
    Dim passPhrase As String
    Dim saltValue As String
    Dim hashAlgorithm As String
    Dim passwordIterations As Integer
    Dim initVector As String
    Dim keySize As Integer
    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        Try
            If cbOraSettings.Text = Nothing Then
                MsgBox("Please Choose settings you want to remove", MsgBoxStyle.Information)
            Else
                If MsgBox("Do you want to remove " & cbOraSettings.Text & "?", MsgBoxStyle.OkCancel) = MsgBoxResult.Ok Then
                    db.ExecuteSQLQuery("Delete From SSINFOTERMDBSETTINGS where Name = '" & cbOraSettings.Text & "'")
                    db.fillComboBox(db.ExecuteSQLQuery("Select Name from SSINFOTERMDBSETTINGS"), cbOraSettings)
                    MsgBox("Successfully removed.", MsgBoxStyle.Information)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

 

    Private Sub cbOraSettings_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbOraSettings.Click
        Try
            db.fillComboBox(db.ExecuteSQLQuery("Select Name from SSINFOTERMDBSETTINGS where type = 'ORA'"), cbOraSettings)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Try

            er = 0
            If txtHost.Text = Nothing Then
                MsgBox("Host is empty.", MsgBoxStyle.Information)
                er = er - 1
            Else
                er = er + 1
            End If
            If txtPort.Text = Nothing Then
                MsgBox("Port is empty.", MsgBoxStyle.Information)
                er = er - 1
            Else
                er = er + 1
            End If

            If txtServiceName.Text = Nothing Then
                MsgBox("Service name is empty.", MsgBoxStyle.Information)
                er = er - 1
            Else
                er = er + 1
            End If

            If txtUserID.Text = Nothing Then
                MsgBox("UserID is empty.", MsgBoxStyle.Information)
                er = er - 1
            Else
                er = er + 1
            End If

            If txtOrapass.Text = Nothing Then
                MsgBox("Password is empty.", MsgBoxStyle.Information)
                er = er - 1
            Else
                er = er + 1
            End If
            _frmBlock.Show()
            _frmPassword.ShowDialog()
            If er = 5 And _frmPassword.x = 1 Then
                If db.checkExistence("Select Name from SSINFOTERMDBSETTINGS where Name = '" & cbOraSettings.Text & "'") = True Then
                    If MsgBox("Do you want to Update '" & cbOraSettings.Text & "?", MsgBoxStyle.OkCancel) = MsgBoxResult.Ok Then
                        My.Settings.lastSettingsORA = db.putSingleValue("Select Name from SSINFOTERMDBSETTINGS where tag = '1' and type = 'ORA'")
                        plainText = txtOrapass.Text    ' original plaintext
                        passPhrase = "Pas5pr@se"        ' can be any string
                        saltValue = "s@1tValue"         ' can be any string
                        hashAlgorithm = "SHA1"          ' can be "MD5"
                        passwordIterations = 2          ' can be any number
                        initVector = "@1B2c3D4e5F6g7H8" ' must be 16 bytes
                        keySize = 256                   ' can be 192 or 128

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
                        db.ExecuteSQLQuery("UPDATE SSINFOTERMDBSETTINGS set tag = '0' where Name = '" & My.Settings.lastSettingsORA & "'")
                        db.ExecuteSQLQuery("Update SSINFOTERMDBSETTINGS Set Name='" & txtOra.Text & "' ,Dsn = '" & txtHost.Text & "', Server = '" & txtPort.Text & "', Username = '" & txtUserID.Text & "', Password = '" & cipherText & "',[Database] = '" & txtServiceName.Text & "',tag = '1' where Name = '" & cbOraSettings.Text & "' and type = 'ORA'")
                        MsgBox(cbOraSettings.Text & "Has been Successfully Updated!")
                        db.fillComboBox(db.ExecuteSQLQuery("Select Name from SSINFOTERMDBSETTINGS"), cbOraSettings)
                        er = 0
                        btnNewOra.Enabled = False
                        Dim myPath As String = Application.StartupPath & "\Settings"
                        Using SW As New IO.StreamWriter(myPath & "\ORASettings.txt", False)
                            SW.Write("Host: " & txtHost.Text & "|Port: " & txtPort.Text & "|UserID: " & txtUserID.Text & "|Password: " & txtOrapass.Text & "|Service: " & txtServiceName.Text)
                        End Using

                        Dim fileName As String = myPath & "\ORASettings.txt"
                        fileName = fileName.Replace("\", "/")
                        Dim DestFile As String = ""
                        up.UploadFile(fileName, "ftp://" & My.Settings.FTPIPSETTINGS & "/Settings/ORASettings.txt", My.Settings.FTPUSERSETTINGS, My.Settings.FTPPASSSETTINGS)
                        My.Settings.lastSettingsORA = db.putSingleValue("Select Name from SSINFOTERMDBSETTINGS where tag = '1' and type = 'ORA'")
                        db.ExecuteSQLQuery("UPDATE SSINFOTERMUPDATE set DATEUPDATED ='" & Date.Today & "',TIMEUPDATED = '" & TimeOfDay & "' where TAG = 4")
                    End If
                Else
                    If MsgBox("Do you want to save a new Settings?", MsgBoxStyle.OkCancel) = MsgBoxResult.Ok Then
                        My.Settings.lastSettingsORA = db.putSingleValue("Select Name from SSINFOTERMDBSETTINGS where tag = '1'")
                        db.ExecuteSQLQuery("UPDATE SSINFOTERMDBSETTINGS set tag = '0' where Name = '" & My.Settings.lastSettingsORA & "' and type = 'ORA'")
                        plainText = txtOrapass.Text    ' original plaintext
                        passPhrase = "Pas5pr@se"        ' can be any string
                        saltValue = "s@1tValue"         ' can be any string
                        hashAlgorithm = "SHA1"          ' can be "MD5"
                        passwordIterations = 2          ' can be any number
                        initVector = "@1B2c3D4e5F6g7H8" ' must be 16 bytes
                        keySize = 256                   ' can be 192 or 128

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
                        db.ExecuteSQLQuery("INSERT INTO SSINFOTERMDBSETTINGS (Name,Dsn,Server,Username,Password,[Database],tag,type) values('" & txtOra.Text & "','" & txtHost.Text & "','" & txtPort.Text & "','" & txtUserID.Text & "','" & cipherText & "','" & txtServiceName.Text & "','1','ORA')")
                        MsgBox("successfully saved.")
                        db.fillComboBox(db.ExecuteSQLQuery("Select Name from SSINFOTERMDBSETTINGS"), cbOraSettings)
                        er = 0
                        btnNewOra.Enabled = False
                        Dim myPath As String = Application.StartupPath & "\Settings"
                        Using SW As New IO.StreamWriter(myPath & "\ORASettings.txt", False)
                            SW.Write("Host: " & txtHost.Text & "|Port: " & txtPort.Text & "|UserID: " & txtUserID.Text & "|Password: " & txtOrapass.Text & "|Service: " & txtServiceName.Text)
                        End Using

                        Dim fileName As String = myPath & "\ORASettings.txt"
                        fileName = fileName.Replace("\", "/")
                        Dim DestFile As String = ""
                        up.UploadFile(fileName, "ftp://" & My.Settings.FTPIPSETTINGS & "/Settings/ORASettings.txt", My.Settings.FTPUSERSETTINGS, My.Settings.FTPPASSSETTINGS)
                        My.Settings.lastSettingsORA = db.putSingleValue("Select Name from SSINFOTERMDBSETTINGS where tag = '1' and type = 'ORA'")
                        db.ExecuteSQLQuery("UPDATE SSINFOTERMUPDATE set DATEUPDATED ='" & Date.Today & "',TIMEUPDATED = '" & TimeOfDay & "' where TAG = 4")
                        btnNewOra.Enabled = False
                    End If
                End If
            End If
            My.Settings.Save()
            My.Settings.Reload()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub btnNewOra_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNewOra.Click
        cbOraSettings.Text = Nothing
        txtOra.Text = Nothing
        txtHost.Text = Nothing
        txtPort.Text = Nothing
        txtUserID.Text = Nothing
        txtOrapass.Text = Nothing
        txtServiceName.Text = Nothing
    End Sub

    Private Sub _frmORA_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim a As String = db.putSingleValue("select NAME FROM SSINFOTERMDBSETTINGS where type = 'ora' and tag = '1'")
        If a = Nothing Then

        Else
            cbOraSettings.Text = db.putSingleValue("Select Name from SSINFOTERMDBSETTINGS where type = 'ORA' and tag = '1'")
            txtOra.Text = db.putSingleValue("Select Name from SSINFOTERMDBSETTINGS where type = 'ORA' and tag = '1'")
            txtHost.Text = db.putSingleValue("Select Dsn from SSINFOTERMDBSETTINGS where type = 'ORA' and tag = '1'")
            txtPort.Text = db.putSingleValue("Select Server from SSINFOTERMDBSETTINGS where type = 'ORA' and tag = '1'")
            txtUserID.Text = db.putSingleValue("Select Username from SSINFOTERMDBSETTINGS where type = 'ORA' and tag = '1'")
            cipherText = db.putSingleValue("Select Password from SSINFOTERMDBSETTINGS where type = 'ORA' and tag = '1'")
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
            txtOrapass.Text = plainText
            txtServiceName.Text = db.putSingleValue("Select [Database] from SSINFOTERMDBSETTINGS where type = 'ORA' and tag = '1'")

        End If
       
    End Sub

    Private Sub cbOraSettings_SelectedIndexChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbOraSettings.SelectedIndexChanged
        If cbOraSettings.Text = Nothing Then
        Else

            btnNewOra.Enabled = True
            txtOra.Text = db.putSingleValue("Select Name from SSINFOTERMDBSETTINGS where Name = '" & cbOraSettings.Text & "' and type = 'ORA'")
            txtHost.Text = db.putSingleValue("Select Dsn from SSINFOTERMDBSETTINGS where Name = '" & cbOraSettings.Text & "' and type = 'ORA'")
            txtPort.Text = db.putSingleValue("Select Server from SSINFOTERMDBSETTINGS where Name = '" & cbOraSettings.Text & "' and type = 'ORA'")
            txtUserID.Text = db.putSingleValue("Select Username from SSINFOTERMDBSETTINGS where Name = '" & cbOraSettings.Text & "' and type = 'ORA'")
            cipherText = db.putSingleValue("Select Password from SSINFOTERMDBSETTINGS where Name = '" & cbOraSettings.Text & "' and type = 'ORA'")
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
            txtServiceName.Text = db.putSingleValue("Select [Database] from SSINFOTERMDBSETTINGS where Name = '" & cbOraSettings.Text & "' and type = 'ORA'")

        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim temp As String = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=" & txtHost.Text & ")(PORT=" & txtPort.Text & "))(CONNECT_DATA=(SERVICE_NAME=" & txtServiceName.Text & ")));User Id=" & txtUserID.Text & ";Password=" & txtOrapass.Text & ";"


        If Oracle.webisconnected(temp) Then
            MsgBox("Parameters are correct" & vbNewLine & "You are now connected to server", MsgBoxStyle.Information, "Information")
            Button8.Enabled = True
        Else
            MsgBox("Unable to connect." & vbNewLine & "Please check if all the parameters are correct.", MsgBoxStyle.Exclamation, "Information")
        End If
    End Sub
End Class