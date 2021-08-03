Public Class _frmLoading
    Dim db As New ConnectionString
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim temp As String = "DSN=" & My.Settings.db_DSN & ";SERVER=" & My.Settings.db_Server & ";DATABASE=" & My.Settings.db_Name & ";UID=" & My.Settings.db_UName & ";PWD=" & My.Settings.db_Pass & ""
        'Dim temp As String = "SERVER=" & TextBox2.Text & ";DATABASE=" & TextBox5.Text & ";UID=" & TextBox3.Text & ";PWD=" & TextBox4.Text & ""
        ''Dim temp As String = "Data Source=" & TextBox2.Text & ";Initial Catalog=" & TextBox5.Text & ";User ID=" & TextBox3.Text & ";Password=" & TextBox4.Text & ""

        If db.webisconnected(temp) Then
            _frmMain.validationOfUserRole()
            _frmMain.Timer1.Start()
            _frmMain.Timer2.Start()
            _frmDashboard.Timer1.Start()
        Else
            MsgBox("Connection time-out")
        End If
    End Sub
End Class