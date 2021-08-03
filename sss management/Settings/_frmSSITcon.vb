Imports System.IO
Public Class _frmSSITcon
    Dim db As New ConnectionString
    Dim er As Integer
    Dim up As New UploadFTP
    Dim plainText As String
    Dim cipherText As String
    Dim passPhrase As String
    Dim saltValue As String
    Dim hashAlgorithm As String
    Dim passwordIterations As Integer
    Dim initVector As String
    Dim keySize As Integer
    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Dim temp As String = "DSN=" & txtDsn1.Text & ";SERVER=" & txtServer1.Text & ";DATABASE=" & txtDB1.Text & ";UID=" & txtUsername1.Text & ";PWD=" & txtPword1.Text & ""
        'Dim temp As String = "SERVER=" & TextBox2.Text & ";DATABASE=" & TextBox5.Text & ";UID=" & TextBox3.Text & ";PWD=" & TextBox4.Text & ""
        ''Dim temp As String = "Data Source=" & TextBox2.Text & ";Initial Catalog=" & TextBox5.Text & ";User ID=" & TextBox3.Text & ";Password=" & TextBox4.Text & ""

        If db.webisconnected(temp) Then
            MsgBox("Parameters are correct" & vbNewLine & "You are now connected to server", MsgBoxStyle.Information)
            btnsave1.Enabled = True
        Else
            MsgBox("Unable to connect!" & vbNewLine & "Please check if all the parameters are correct", MsgBoxStyle.Exclamation)
            btnsave1.Enabled = False
        End If
    End Sub

    Private Sub btnNewSSIT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNewSSIT.Click
        cbSettings.Text = Nothing
        txtName1.Text = Nothing
        txtDsn1.Text = Nothing
        txtServer1.Text = Nothing
        txtUsername1.Text = Nothing
        txtPword1.Text = Nothing
        txtDB1.Text = Nothing
    End Sub

    Private Sub btnsave1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnsave1.Click
        Try
            er = 0
            If txtName1.Text = Nothing Then
                MsgBox("Name settings is empty.", MsgBoxStyle.Information)
                er = er - 1
            Else
                er = er + 1
            End If
            If txtDsn1.Text = Nothing Then
                MsgBox("DSN is empty.", MsgBoxStyle.Information)
                er = er - 1
            Else
                er = er + 1
            End If

            If txtServer1.Text = Nothing Then
                MsgBox("Server is empty.", MsgBoxStyle.Information)
                er = er - 1
            Else
                er = er + 1
            End If

            If txtUsername1.Text = Nothing Then
                MsgBox("Username is empty.", MsgBoxStyle.Information)
                er = er - 1
            Else
                er = er + 1
            End If

            If txtPword1.Text = Nothing Then
                MsgBox("Password is empty.", MsgBoxStyle.Information)
                er = er - 1
            Else
                er = er + 1
            End If

            If txtDB1.Text = Nothing Then
                MsgBox("", MsgBoxStyle.Information)
                er = er - 1
            Else
                er = er + 1
            End If
            _frmBlock.Show()
            _frmPassword.ShowDialog()
            If er = 6 And _frmPassword.x = 1 Then
                'If Not Directory.Exists("D:\ftp\Settings\") Then
                '    Directory.CreateDirectory("D:\ftp\Settings\")
                'End If

                If db.checkExistence("Select Name from SSINFOTERMDBSETTINGS where Name = '" & txtName1.Text & "'") = True Then

                    If MsgBox("Do you want to Update '" & txtName1.Text & "?", MsgBoxStyle.OkCancel) = MsgBoxResult.Ok Then
                        plainText = txtPword1.Text    ' original plaintext
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
                        My.Settings.lastSettingsSQL = db.putSingleValue("Select Name from SSINFOTERMDBSETTINGS where tag = '1' and type = 'SQL'")
                        db.ExecuteSQLQuery("UPDATE SSINFOTERMDBSETTINGS set tag = '0' where Name = '" & My.Settings.lastSettingsSQL & "'")
                        db.ExecuteSQLQuery("Update SSINFOTERMDBSETTINGS Set Name = '" & txtName1.Text & "', Dsn = '" & txtDsn1.Text & "', Server = '" & txtServer1.Text & "', Username = '" & txtUsername1.Text & "', Password = '" & cipherText & "',[Database] = '" & txtDB1.Text & "',Time = '" & cbTime.Text & "',tag = '1' where Name = '" & cbSettings.Text & "'")
                        MsgBox(cbSettings.Text & "Has been Successfully Updated!")
                        db.fillComboBox(db.ExecuteSQLQuery("Select Name from SSINFOTERMDBSETTINGS"), cbSettings)
                        btnNewSSIT.Enabled = False
                        er = 0
                        Dim myPath As String = Application.StartupPath & "\Settings"
                        If (Not System.IO.Directory.Exists(myPath)) Then
                            System.IO.Directory.CreateDirectory(myPath)
                        End If
                        Dim date1 As String = Today.Date.ToString("MM-dd-yyyy")
                        Using SW As New IO.StreamWriter(myPath & "\SQLSettings.txt", False)
                            SW.Write("DSN: " & txtDsn1.Text & "|Server: " & txtServer1.Text & "|Username: " & txtUsername1.Text & "|Password: " & txtPword1.Text & "|Database: " & txtDB1.Text)
                        End Using

                        Dim fileName As String = myPath & "\SQLSettings.txt"
                        fileName = fileName.Replace("\", "/")
                        Dim DestFile As String = ""
                        up.UploadFile(fileName, "ftp://" & My.Settings.FTPIPSETTINGS & "/Settings/SQLSettings.txt", My.Settings.FTPUSERSETTINGS, My.Settings.FTPPASSSETTINGS)
                        My.Settings.lastSettingsSQL = db.putSingleValue("Select Name from SSINFOTERMDBSETTINGS where tag = '1' and type = 'SQL'")
                        If db.checkExistence("Select isnull(tag,0) from SSINFOTERMUPDATE where tag = 4") = 0 Then
                            db.ExecuteSQLQuery("INSERT INTO SSINFOTERMUPDATE(TAG,DATEUPDATED,TIMEUPDATED) values(4,'" & Today.Date & "','" & TimeOfDay & "')")
                        Else
                            db.ExecuteSQLQuery("UPDATE SSINFOTERMUPDATE set DATEUPDATED ='" & Date.Today & "',TIMEUPDATED = '" & TimeOfDay & "' where TAG = 4")
                        End If
                 
                    End If
                Else
                    If MsgBox("Do you want to save a new Settings?", MsgBoxStyle.OkCancel) = MsgBoxResult.Ok Then
                        My.Settings.lastSettingsSQL = db.putSingleValue("Select Name from SSINFOTERMDBSETTINGS where tag = '1' and type = 'SQL'")
                        db.ExecuteSQLQuery("UPDATE SSINFOTERMDBSETTINGS set tag = '0' where Name = '" & My.Settings.lastSettingsSQL & "'")
                        plainText = txtPword1.Text    ' original plaintext
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
                        db.ExecuteSQLQuery("INSERT INTO SSINFOTERMDBSETTINGS (Name,Dsn,Server,Username,Password,[Database],[Time],tag,type) values('" & txtName1.Text & "','" & txtDsn1.Text & "','" & txtServer1.Text & "','" & txtUsername1.Text & "','" & cipherText & "','" & txtDB1.Text & "','" & cbTime.Text & "','1','SQL')")
                        MsgBox("Successfully saved.")
                        db.fillComboBox(db.ExecuteSQLQuery("Select Name from SSINFOTERMDBSETTINGS"), cbSettings)
                        er = 0
                        btnNewSSIT.Enabled = False

                        Dim myPath As String = Application.StartupPath & "\Settings"
                        If (Not System.IO.Directory.Exists(myPath)) Then
                            System.IO.Directory.CreateDirectory(myPath)
                        End If
                        Dim date1 As String = Today.Date.ToString("MM-dd-yyyy")
                        Using SW As New IO.StreamWriter(myPath & "\SQLSettings.txt", False)
                            SW.Write("DSN: " & txtDsn1.Text & "|Server: " & txtServer1.Text & "|Username: " & txtUsername1.Text & "|Password: " & txtPword1.Text & "|Database: " & txtDB1.Text)
                        End Using
                        Dim fileName As String = myPath & "\Settings\SQLSettings.txt"
                        fileName = fileName.Replace("\", "/")
                        Dim DestFile As String = ""
                        up.UploadFile(fileName, "ftp://" & My.Settings.FTPIPSETTINGS & "/Settings/SQLSettings.txt", My.Settings.FTPUSERSETTINGS, My.Settings.FTPPASSSETTINGS)
                        My.Settings.lastSettingsSQL = db.putSingleValue("Select Name from SSINFOTERMDBSETTINGS where tag = '1' and type = 'SQL'")
                        If db.checkExistence("Select isnull(tag,0) from SSINFOTERMUPDATE where tag = 4") = 0 Then
                            db.ExecuteSQLQuery("INSERT INTO SSINFOTERMUPDATE(TAG,DATEUPDATED,TIMEUPDATED) values(4,'" & Today.Date & "','" & TimeOfDay & "')")
                        Else
                            db.ExecuteSQLQuery("UPDATE SSINFOTERMUPDATE set DATEUPDATED ='" & Date.Today & "',TIMEUPDATED = '" & TimeOfDay & "' where TAG = 4")
                        End If
                  
                        btnNewSSIT.Enabled = False
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Try
            If cbSettings.Text = Nothing Then
                MsgBox("Please Choose settings you want to remove", MsgBoxStyle.Information)
            Else
                If MsgBox("Do you want to remove " & cbSettings.Text & "?", MsgBoxStyle.OkCancel) = MsgBoxResult.Ok Then
                    db.ExecuteSQLQuery("Delete From SSINFOTERMDBSETTINGS where Name = '" & cbSettings.Text & "'")
                    db.fillComboBox(db.ExecuteSQLQuery("Select Name from SSINFOTERMDBSETTINGS"), cbSettings)
                    MsgBox("Successfully removed.", MsgBoxStyle.Information)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub cbSettings_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbSettings.SelectedIndexChanged
        btnNewSSIT.Enabled = True
        If cbSettings.Text = Nothing Then

        Else
            txtName1.Text = db.putSingleValue("Select Name from SSINFOTERMDBSETTINGS where Name = '" & cbSettings.Text & "' and type = 'SQL'")
            txtDsn1.Text = db.putSingleValue("Select Dsn from SSINFOTERMDBSETTINGS where Name = '" & cbSettings.Text & "' and type = 'SQL'")
            txtServer1.Text = db.putSingleValue("Select Server from SSINFOTERMDBSETTINGS where Name = '" & cbSettings.Text & "' and type = 'SQL'")
            txtUsername1.Text = db.putSingleValue("Select Username from SSINFOTERMDBSETTINGS where Name = '" & cbSettings.Text & "' and type = 'SQL'")
            cipherText = db.putSingleValue("Select Password from SSINFOTERMDBSETTINGS where Name = '" & cbSettings.Text & "' and type = 'SQL'")
            txtDB1.Text = db.putSingleValue("Select [Database] from SSINFOTERMDBSETTINGS where Name = '" & cbSettings.Text & "' and type = 'SQL'")
            cbTime.Text = db.putSingleValue("select [Time] from SSINFOTERMDBSETTINGS where Name = '" & cbSettings.Text & "' and type = 'SQL'")

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
            txtPword1.Text = plainText
        End If
    End Sub

    Private Sub cbSettings_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbSettings.Click
        Try
            db.fillComboBox(db.ExecuteSQLQuery("Select Name from SSINFOTERMDBSETTINGS where type = 'SQL'"), cbSettings)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub _frmSSITcon_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim a As String = db.putSingleValue("Select Name from SSINFOTERMDBSETTINGS where type = 'SQL' and tag = '1'")
        If a = Nothing Then
        Else
            db.fillComboBox(db.ExecuteSQLQuery("Select Name from SSINFOTERMDBSETTINGS where type = 'SQL' and tag = '1'"), cbSettings)
            cbSettings.Text = db.putSingleValue("Select Name from SSINFOTERMDBSETTINGS where type = 'SQL' and tag = '1'")
            txtName1.Text = db.putSingleValue("Select Name from SSINFOTERMDBSETTINGS where type = 'SQL' and tag = '1'")
            txtDsn1.Text = db.putSingleValue("Select Dsn from SSINFOTERMDBSETTINGS where type = 'SQL' and tag = '1'")
            txtServer1.Text = db.putSingleValue("Select Server from SSINFOTERMDBSETTINGS where type = 'SQL' and tag = '1'")
            txtUsername1.Text = db.putSingleValue("Select Username from SSINFOTERMDBSETTINGS where type = 'SQL' and tag = '1'")
            cipherText = db.putSingleValue("Select Password from SSINFOTERMDBSETTINGS where type = 'SQL' and tag = '1'")
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
            txtPword1.Text = plainText

            txtDB1.Text = db.putSingleValue("Select [Database] from SSINFOTERMDBSETTINGS where type = 'SQL' and tag = '1'")
            cbTime.Text = db.putSingleValue("select [Time] from SSINFOTERMDBSETTINGS where type = 'SQL' and tag = '1'")

        End If

    End Sub
End Class