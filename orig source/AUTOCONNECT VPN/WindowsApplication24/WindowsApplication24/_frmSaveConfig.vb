Public Class _frmSaveConfig
    Public xs As New MySettings
    Private Sub Form2_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim VPN_USERNAME, VPN_PASSWORD As String
        xs.xml_Filename = Application.StartupPath & "\Mysettings.xml" '"Mysettings.xml"
        xs.value_path = "Configuration/Settings"
        xs.writeSettings(xs.xml_Filename)
        VPN_USERNAME = xs.readSettings(xs.xml_Filename, xs.value_path, "VPN_USERNAME")
        VPN_PASSWORD = xs.readSettings(xs.xml_Filename, xs.value_path, "VPN_PASSWORD")
        If VPN_USERNAME = " " Or VPN_USERNAME = Nothing Then
            Me.WindowState = FormWindowState.Normal
        Else
            Me.Hide()
            _frmConnect.Show()

        End If

    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Try
            If txtUserName.Text = "" Or txtUserName.Text = Nothing Then
                MsgBox("Please enter a Username.", vbInformation, "SAVE VPN CONFIGURATION")
            ElseIf txtPassword.Text = "" Or txtPassword.Text = Nothing Then
                MsgBox("Please enter a Password.", vbInformation, "SAVE VPN CONFIGURATION")
            Else
                xs.editSettings(xs.xml_Filename, xs.value_path, "VPN_USERNAME", txtUserName.Text)
                xs.editSettings(xs.xml_Filename, xs.value_path, "VPN_PASSWORD", txtPassword.Text)

                MsgBox("Configuration Saved!", vbInformation, "SAVE VPN CONFIGURATION")

                _frmConnect.Show()
                Me.Hide()

            End If

        Catch ex As Exception
            MsgBox(ex.Message, vbInformation, "SAVE VPN CONFIGURATION")

        End Try


    End Sub


End Class