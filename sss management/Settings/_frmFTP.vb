Imports System.Net
Public Class _frmFTP
    Dim db As New ConnectionString
    Dim isConnected As Boolean
    Private Function checkCredentials()

        Try
            Dim request = _
    DirectCast(WebRequest.Create _
    ("ftp://" + FTPIP.Text + "/"), FtpWebRequest)

            request.Credentials = _
            New NetworkCredential(FTPUN.Text, FTPPASS.Text)
            request.Method = WebRequestMethods.Ftp.ListDirectory

            Dim response As FtpWebResponse = _
            DirectCast(request.GetResponse(), FtpWebResponse)
            ' Folder exists here

            ' If response.StatusCode = FtpStatusCode.OpeningData Then
            MsgBox("Connected to FTP Server.", vbInformation, "Information")
            isConnected = True
            '  End If

        Catch ex As WebException
            Dim response As FtpWebResponse = _
            DirectCast(ex.Response, FtpWebResponse)
            'Does not exist
            If response.StatusCode = _
            FtpStatusCode.ActionNotTakenFileUnavailable Then
                MsgBox("Unable to connect. Please check your parameters.", vbInformation, "Information")
            ElseIf response.StatusCode = FtpStatusCode.ServiceTemporarilyNotAvailable Then
                MsgBox("Can't connect to FTP server.", vbInformation, "Information")
            ElseIf response.StatusCode = FtpStatusCode.SendPasswordCommand Then
                MsgBox("Incorrect password.", vbInformation, "Information")
            ElseIf response.StatusCode = FtpStatusCode.NotLoggedIn Then
                MsgBox("Unable to connect. Please check your parameters.", vbInformation, "Information")
            ElseIf response.StatusCode = FtpStatusCode.ServiceNotAvailable Then
                MsgBox("Can't connect to FTP server.", vbInformation, "Information")
            Else
                MsgBox("Invalid permission.", vbInformation, "Information")
                ' MsgBox(ex.ToString, vbInformation, "Information")

            End If
            isConnected = False

        End Try

        Return isConnected

    End Function
    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        _frmBlock.Show()
        _frmPassword.ShowDialog()
        If _frmPassword.x = 1 Then
           
            If db.checkExistence("Select NAME from SSINFOTERMFTP where Name = '" & cbSettings.Text & "'") = True Then
                If MsgBox("Do you want to Update '" & cbSettings.Text & "?", MsgBoxStyle.OkCancel) = MsgBoxResult.Ok Then
                    db.ExecuteSQLQuery("UPDATE SSINFOTERMFTP SET FTPIP = '" & FTPIP.Text & "', FTPUSER = '" & FTPUN.Text & "', FTPPASS = '" & FTPPASS.Text & "' where NAME = '" & cbSettings.Text & "'")
                    My.Settings.FTPIPSETTINGS = FTPIP.Text
                    My.Settings.FTPUSERSETTINGS = FTPUN.Text
                    My.Settings.FTPPASSSETTINGS = FTPPASS.Text

                    My.Settings.Save()
                    My.Settings.Reload()

                    MsgBox("Successfully saved.", MsgBoxStyle.Information)
                End If
            Else

                db.ExecuteSQLQuery("INSERT INTO SSINFOTERMFTP(NAME,FTPUSER,FTPPASS,FTPIP) values ('" & txtFTPName.Text & "','" & FTPUN.Text & "','" & FTPPASS.Text & "','" & FTPIP.Text & "')")
                'db.ExecuteSQLQuery("UPDATE SSINFOTERMFTP SET FTPIP = '" & FTPIP.Text & "', FTPUSER = '" & FTPUN.Text & "', FTPPASS = '" & FTPPASS.Text & "' where NAME = '" & cbSettings.Text & "'")
                My.Settings.FTPIPSETTINGS = FTPIP.Text
                My.Settings.FTPUSERSETTINGS = FTPUN.Text
                My.Settings.FTPPASSSETTINGS = FTPPASS.Text

                My.Settings.Save()
                My.Settings.Reload()
                MsgBox("New settings has been saved.", MsgBoxStyle.Information)
                db.fillComboBox(db.ExecuteSQLQuery("Select NAME FROM SSINFOTERMFTP"), cbSettings)
            End If
        End If

    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Try
            If cbSettings.Text = Nothing Then
                MsgBox("Please Choose settings you want to remove.", MsgBoxStyle.Information, MsgBoxStyle.Information)
            Else
                _frmBlock.Show()
                _frmPassword.ShowDialog()
                If _frmPassword.x = 1 Then
                    If MsgBox("Are you sure you want to remove " & cbSettings.Text & "?", MsgBoxStyle.Information, MsgBoxStyle.OkCancel) = MsgBoxResult.Ok Then
                        db.ExecuteSQLQuery("Delete From SSINFOTERMFTP where Name = '" & cbSettings.Text & "'")
                        db.fillComboBox(db.ExecuteSQLQuery("Select Name from SSINFOTERMFTP"), cbSettings)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub cbSettings_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbSettings.SelectedIndexChanged
        cbFtime.Text = Nothing
        FTPUN.Text = db.putSingleValue("Select FTPUSER from SSINFOTERMFTP where NAME = '" & cbSettings.Text & "'")
        FTPPASS.Text = db.putSingleValue("Select FTPPASS from SSINFOTERMFTP where NAME = '" & cbSettings.Text & "'")
        FTPIP.Text = db.putSingleValue("Select FTPIP from SSINFOTERMFTP where NAME = '" & cbSettings.Text & "'")
        cbFtime.Text = db.putSingleValue("Select FTPTIME from SSINFOTERMFTP where NAME = '" & cbSettings.Text & "'")
        txtFTPName.Text = db.putSingleValue("Select NAME from SSINFOTERMFTP where NAME = '" & cbSettings.Text & "'")
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

    End Sub

    Private Sub _frmFTP_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        db.fillComboBox(db.ExecuteSQLQuery("SELECT NAME FROM SSINFOTERMFTP"), cbSettings)
    End Sub

    Private Sub rbPDF_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        cbFtime.Enabled = True
        cbFtime.Text = Nothing
    End Sub

    Private Sub rbOther_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        cbFtime.Enabled = False
        cbFtime.Text = Nothing
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If FTPIP.Text = "" Then
            MsgBox("FTP IP address is empty.", MsgBoxStyle.Information)
        ElseIf FTPPASS.Text = "" Then
            MsgBox("FTP Password is empty.", MsgBoxStyle.Information)
        ElseIf FTPUN.Text = "" Then
            MsgBox("FTP Username is empty. ", MsgBoxStyle.Information)
        Else
            checkCredentials()
            If isConnected = True Then
                Button7.Enabled = True
            End If
        End If
    End Sub
End Class