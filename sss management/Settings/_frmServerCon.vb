Public Class _frmServerCon
    Dim db As New ConnectionString
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim temp As String = "DSN=" & txtDSN.Text & ";SERVER=" & txtSERVER.Text & ";DATABASE=" & TextBox5.Text & ";UID=" & txtUserName.Text & ";PWD=" & txtPassword.Text & ""
        'Dim temp As String = "SERVER=" & TextBox2.Text & ";DATABASE=" & TextBox5.Text & ";UID=" & TextBox3.Text & ";PWD=" & TextBox4.Text & ""
        ''Dim temp As String = "Data Source=" & TextBox2.Text & ";Initial Catalog=" & TextBox5.Text & ";User ID=" & TextBox3.Text & ";Password=" & TextBox4.Text & ""

        If db.webisconnected(temp) Then
            MsgBox("Parameters are correct" & vbNewLine & "You are now connected to server", MsgBoxStyle.Information)
            btnSave.Enabled = True
        Else
            MsgBox("Unable to connect!" & vbNewLine & "Please check if all the parameters are correct", MsgBoxStyle.Exclamation)
            btnSave.Enabled = False
        End If
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click



        Try
            If txtDSN.Text = Nothing Then
                Throw New Exception("DSN is empty.")
            End If
            If txtSERVER.Text = Nothing Then
                Throw New Exception("Server is empty.")
            End If

            If txtUserName.Text = Nothing Then
                Throw New Exception("Username is empty.")
            End If

            If txtPassword.Text = Nothing Then
                Throw New Exception("Password is empty.")
            End If
            If My.Settings.firstRun = 0 Then
                My.Settings.db_Server = txtSERVER.Text
                My.Settings.db_UName = txtUserName.Text
                My.Settings.db_Pass = txtPassword.Text
                My.Settings.db_Name = TextBox5.Text
                My.Settings.db_DSN = txtDSN.Text
                My.Settings.firstRun = 1
                My.Settings.Save()
                My.Settings.Reload()

                MsgBox("New paramenters has been saved.", MsgBoxStyle.Information)
                MsgBox("System automatically refresh your settings", MsgBoxStyle.Information)
                Application.Exit()
            Else
                _frmBlock.Show()
                _frmPassword.ShowDialog()
                If _frmPassword.x = 1 Then
                    My.Settings.db_Server = txtSERVER.Text
                    My.Settings.db_UName = txtUserName.Text
                    My.Settings.db_Pass = txtPassword.Text
                    My.Settings.db_Name = TextBox5.Text
                    My.Settings.db_DSN = txtDSN.Text
                    My.Settings.firstRun = 1
                    My.Settings.Save()
                    My.Settings.Reload()

                    MsgBox("New paramenters has been saved.", MsgBoxStyle.Information)
                    MsgBox("System automatically refresh your settings", MsgBoxStyle.Information)

                    Application.Exit()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Exclamation)
        End Try
    End Sub

    Private Sub btnBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBack.Click
        _frmLogin.Show()
        Me.Hide()
    End Sub
End Class