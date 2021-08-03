Public Class _frmSettings
    Dim cnt As Integer
    Public Sub buttonclicked()
        If cnt = 1 Then

            btnServer.FlatStyle = FlatStyle.Flat
            btnServer.BackColor = Color.Blue
            btnServer.ForeColor = Color.White
        Else
            btnServer.BackColor = Color.White
            btnServer.FlatStyle = FlatStyle.Standard
            btnServer.ForeColor = Color.Black
        End If
        If cnt = 2 Then
            btnSSIT.FlatStyle = FlatStyle.Flat
            btnSSIT.BackColor = Color.Blue
            btnSSIT.ForeColor = Color.White
        Else
            btnSSIT.BackColor = Color.White
            btnSSIT.FlatStyle = FlatStyle.Standard
            btnSSIT.ForeColor = Color.Black
        End If
        If cnt = 3 Then

            btnORA.FlatStyle = FlatStyle.Flat
            btnORA.BackColor = Color.Blue
            btnORA.ForeColor = Color.White
        Else
            btnORA.BackColor = Color.White
            btnORA.FlatStyle = FlatStyle.Standard
            btnORA.ForeColor = Color.Black
        End If
        If cnt = 4 Then

            btnFTP.FlatStyle = FlatStyle.Flat
            btnFTP.BackColor = Color.Blue
            btnFTP.ForeColor = Color.White
        Else
            btnFTP.BackColor = Color.White
            btnFTP.FlatStyle = FlatStyle.Standard
            btnFTP.ForeColor = Color.Black
        End If

        If cnt = 5 Then

            btnIdle.FlatStyle = FlatStyle.Flat
            btnIdle.BackColor = Color.Blue
            btnIdle.ForeColor = Color.White
        Else
            btnIdle.BackColor = Color.White
            btnIdle.FlatStyle = FlatStyle.Standard
            btnIdle.ForeColor = Color.Black
        End If
   
        If cnt = 6 Then
            btnUP.FlatStyle = FlatStyle.Flat
            btnUP.BackColor = Color.Blue
            btnUP.ForeColor = Color.White
        Else
            btnUP.BackColor = Color.White
            btnUP.FlatStyle = FlatStyle.Standard
            btnUP.ForeColor = Color.Black
        End If
    End Sub
    Private Sub btnServer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnServer.Click
        cnt = 1
        buttonclicked()
        SplitContainer1.Panel2.Controls.Clear()
        _frmServerCon.TopLevel = False
        _frmServerCon.Parent = Me.SplitContainer1.Panel2
        _frmServerCon.Dock = DockStyle.Fill
        _frmServerCon.Show()
        _frmServerCon.btnBack.Visible = False
    End Sub

    Private Sub btnSSIT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSSIT.Click
        cnt = 2
        buttonclicked()
        SplitContainer1.Panel2.Controls.Clear()
        _frmSSITcon.TopLevel = False
        _frmSSITcon.Parent = Me.SplitContainer1.Panel2
        _frmSSITcon.Dock = DockStyle.Fill
        _frmSSITcon.Show()
    End Sub

    Private Sub btnORA_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnORA.Click
        cnt = 3
        buttonclicked()
        SplitContainer1.Panel2.Controls.Clear()
        _frmORA.TopLevel = False
        _frmORA.Parent = Me.SplitContainer1.Panel2
        _frmORA.Dock = DockStyle.Fill
        _frmORA.Show()
    End Sub

    Private Sub btnFTP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFTP.Click
        cnt = 4
        buttonclicked()
        SplitContainer1.Panel2.Controls.Clear()
        _frmFTP.TopLevel = False
        _frmFTP.Parent = Me.SplitContainer1.Panel2
        _frmFTP.Dock = DockStyle.Fill
        _frmFTP.Show()
    End Sub

    Private Sub btnIdle_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnIdle.Click
        cnt = 5
        buttonclicked()
        SplitContainer1.Panel2.Controls.Clear()
        _frmIdle.TopLevel = False
        _frmIdle.Parent = Me.SplitContainer1.Panel2
        _frmIdle.Dock = DockStyle.Fill
        _frmIdle.Show()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUP.Click
        cnt = 6
        buttonclicked()
        SplitContainer1.Panel2.Controls.Clear()
        _frmUpload.TopLevel = False
        _frmUpload.Parent = Me.SplitContainer1.Panel2
        _frmUpload.Dock = DockStyle.Fill
        _frmUpload.Show()
    End Sub
End Class