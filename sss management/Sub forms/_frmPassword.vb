Public Class _frmPassword
    Public x As Integer
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If txtpass.Text = "" Then
            MsgBox("Password is empty.", MsgBoxStyle.Information)
        Else
            If txtpass.Text = My.Settings.LastPasswordLogin Then
                x = 1
                Me.Close()
                _frmBlock.Close()
            Else
                MsgBox("Incorrect password.", MsgBoxStyle.Information)
            End If
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Dispose()
        _frmBlock.Close()
    End Sub

    Private Sub _frmPassword_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        x = 0
        txtpass.Text = Nothing
    End Sub
End Class