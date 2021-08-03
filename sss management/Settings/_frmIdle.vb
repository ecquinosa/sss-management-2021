Imports Microsoft.VisualBasic
Public Class _frmIdle
    Dim db As New ConnectionString
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveMGMT.Click
        If TextBox1.Text = Nothing Then
            MsgBox("Please Input idle time.")
        Else
            _frmBlock.Show()
            _frmPassword.ShowDialog()
            If _frmPassword.x = 1 Then
                My.Settings.Idletime = TextBox1.Text * 60
                My.Settings.Save()
                My.Settings.Reload()
                MsgBox("Idle time has change to " & My.Settings.Idletime / 60 & " Minute(s)")
                lblIT.Text = "IDLE TIME = " & My.Settings.Idletime / 60 & " Minute(s)"
                btnServer.Enabled = True
                TextBox1.Enabled = False
            Else

            End If
        End If
    End Sub

    Private Sub _frmIdle_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        lblIT.Text = "IDLE TIME = " & My.Settings.Idletime / 60 & " Minute(s)"
        Dim IdleTime As String = db.putSingleValue("Select DISTINCT IDLE_TIME from SSINFOTERMKIOSK")
        lblSIT.Text = "IDLE TIME = " & IdleTime / 60

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnServer.Click
        TextBox1.Enabled = True
        btnServer.Enabled = False
        btnSaveMGMT.Enabled = True
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSSIT.Click
        txtSSITtime.Enabled = True
        btnSSIT.Enabled = False
        btnSSITSave.Enabled = True
    End Sub

    Private Sub btnSSITSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSSITSave.Click
        Dim NewTime As String = txtSSITtime.Text * 60
        If txtSSITtime.Text = Nothing Then
            MsgBox("Please Input idle time.")
        Else
            _frmBlock.Show()
            _frmPassword.ShowDialog()
            If _frmPassword.x = 1 Then
                db.ExecuteSQLQuery("UPDATE SSINFOTERMKIOSK SET IDLE_TIME = '" & NewTime & "'")
                txtSSITtime.Enabled = False
                btnSSIT.Enabled = True
            Else

            End If
        End If
    End Sub

    Private Sub txtSSITtime_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSSITtime.KeyPress
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        btnServer.Enabled = True
        TextBox1.Enabled = False
        btnSaveMGMT.Enabled = False
    End Sub

    Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        txtSSITtime.Enabled = False
        btnSSIT.Enabled = True
        btnSSITSave.Enabled = False
    End Sub
End Class