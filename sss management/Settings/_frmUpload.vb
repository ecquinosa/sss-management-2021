Imports System.Net
Imports System.IO
Public Class _frmUpload
    Dim up As New UploadFTP
    Dim FileName As String
    Dim db As New ConnectionString
    Private Sub Panel1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel1.Paint

    End Sub

    Private Sub btnsave1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpload.Click
        Try
            Dim myPath As String = Application.StartupPath & "\UpdateLogs"
            If (Not System.IO.Directory.Exists(myPath)) Then
                System.IO.Directory.CreateDirectory(myPath)
            End If
            Dim getdate As String = Date.Today.ToShortDateString
            getdate = getdate.Replace("/", "-")
            Using SW As New IO.StreamWriter(myPath & "\SSIT UPDATES" & " " & getdate & ".txt", True)

                'Dim str As String = "D:\ftp\Ext\ABC\Stv\text.txt"
                'Dim newstr As String = str.Substring(0, 1)
                'Debug.Print(newstr)
                'Dim new1 As String = str.Substring(3, str.LastIndexOf("\") - 2)
                'Debug.Print(new1)

                Select Case FileName

                    Case 1
                        Dim list As New ListBox
                        'If Not Directory.Exists("D:\ftp\CITIZEN_CHARTER\") Then
                        '    Directory.CreateDirectory("D:\ftp\CITIZEN_CHARTER\")
                        'End If

                        Dim OpenFolder As New FolderBrowserDialog
                        Dim clsDialog As New OpenFileDialog
                        clsDialog.Title = "Select File(s)"
                        clsDialog.Multiselect = False
                        clsDialog.FileName = "CITIZENS_CHARTER.pdf"
                        clsDialog.Filter = "Pdf Files (*.PDF)|*.PDF"

                        If clsDialog.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then

                            If clsDialog.SafeFileName = "CITIZENS_CHARTER.pdf" Then
                                _frmBlock.Show()
                                _frmPassword.ShowDialog()
                                If _frmPassword.x = 1 Then
                                    For Each sFile As String In clsDialog.FileNames

                                        Dim pos As Integer = sFile.LastIndexOf("\"c)
                                        list.Items.Add(sFile.Substring(pos + 1))
                                        up.UploadFile(sFile, "ftp://" & My.Settings.FTPIPSETTINGS & "/Citizen_Charter/" & sFile.Substring(pos + 1), My.Settings.FTPUSERSETTINGS, My.Settings.FTPPASSSETTINGS)
                                        _frmPassword.x = 0
                                    Next

                                    If db.checkExistence("Select isnull(tag,0) from SSINFOTERMUPDATE where tag = 1") = 0 Then
                                        db.ExecuteSQLQuery("INSERT INTO SSINFOTERMUPDATE(TAG,DATEUPDATED,TIMEUPDATED) values(1,'" & Today.Date & "','" & TimeOfDay & "')")
                                    Else
                                        db.ExecuteSQLQuery("UPDATE SSINFOTERMUPDATE set DATEUPDATED ='" & Date.Today & "',TIMEUPDATED = '" & TimeOfDay & "' where TAG = 1")
                                    End If
                                    MsgBox("File transfer complete")
                                    SW.WriteLine("CITIZEN CHARTER" & DateAndTime.Now)
                                End If
                            Else
                                MsgBox("Incorrect file name.", MsgBoxStyle.Information)
                            End If
                        End If
                        clsDialog.Dispose()
                    Case 2
                        Dim list As New ListBox
                        'If Not Directory.Exists("D:\ftp\Video\") Then
                        '    Directory.CreateDirectory("D:\ftp\Video\")
                        'End If

                        Dim OpenFolder As New FolderBrowserDialog
                        Dim clsDialog As New OpenFileDialog
                        clsDialog.Title = "Select File(s)"
                        clsDialog.Multiselect = True
                        clsDialog.FileName = "SSS.mp4"
                        clsDialog.Filter = "Mp4 Files (*.MP4)|*.MP4"

                        If clsDialog.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then

                            If clsDialog.SafeFileName = "SSS.mp4" Then
                                _frmBlock.Show()
                                _frmPassword.ShowDialog()
                                If _frmPassword.x = 1 Then
                                    For Each sFile As String In clsDialog.FileNames

                                        Dim pos As Integer = sFile.LastIndexOf("\"c)
                                        list.Items.Add(sFile.Substring(pos + 1))
                                        up.UploadFile(sFile, "ftp://" & My.Settings.FTPIPSETTINGS & "/Video/" & sFile.Substring(pos + 1), My.Settings.FTPUSERSETTINGS, My.Settings.FTPPASSSETTINGS)
                                        _frmPassword.x = 0
                                    Next
                                    If db.checkExistence("Select isnull(tag,0) from SSINFOTERMUPDATE where tag = 2") = 0 Then
                                        db.ExecuteSQLQuery("INSERT INTO SSINFOTERMUPDATE(TAG,DATEUPDATED,TIMEUPDATED) values(2,'" & Today.Date & "','" & TimeOfDay & "')")
                                    Else
                                        db.ExecuteSQLQuery("UPDATE SSINFOTERMUPDATE set DATEUPDATED ='" & Date.Today & "',TIMEUPDATED = '" & TimeOfDay & "' where TAG = 2")
                                    End If
                                    MsgBox("File transfer complete.")
                                    SW.WriteLine("SSS VIDEO" & DateAndTime.Now)
                                End If

                            End If
                        Else
                            MsgBox("Incorrect file name.", MsgBoxStyle.Information)
                        End If
                        clsDialog.Dispose()
                    Case 3
                        Dim list As New ListBox
                        'If Not Directory.Exists("D:\ftp\Exe\") Then
                        '    Directory.CreateDirectory("D:\ftp\Exe\")
                        'End If

                        Dim OpenFolder As New FolderBrowserDialog
                        Dim clsDialog As New OpenFileDialog
                        clsDialog.Title = "Select File(s)"
                        clsDialog.Multiselect = True
                        clsDialog.FileName = "SSIT_SERVER.exe"
                        clsDialog.Filter = "Exe Files (*.EXE)|*.EXE"
                        If clsDialog.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then

                            If clsDialog.SafeFileName = "SSIT_SERVER.exe" Then
                                _frmBlock.Show()
                                _frmPassword.ShowDialog()
                                If _frmPassword.x = 1 Then
                                    For Each sFile As String In clsDialog.FileNames

                                        Dim pos As Integer = sFile.LastIndexOf("\"c)
                                        list.Items.Add(sFile.Substring(pos + 1))
                                        up.UploadFile(sFile, "ftp://" & My.Settings.FTPIPSETTINGS & "/Application/" & sFile.Substring(pos + 1), My.Settings.FTPUSERSETTINGS, My.Settings.FTPPASSSETTINGS)
                                        _frmPassword.x = 0
                                    Next
                                    If db.checkExistence("Select isnull(tag,0) from SSINFOTERMUPDATE where tag = 3") = 0 Then
                                        db.ExecuteSQLQuery("INSERT INTO SSINFOTERMUPDATE(TAG,DATEUPDATED,TIMEUPDATED) values(3,'" & Today.Date & "','" & TimeOfDay & "')")
                                    Else
                                        db.ExecuteSQLQuery("UPDATE SSINFOTERMUPDATE set DATEUPDATED ='" & Date.Today & "',TIMEUPDATED = '" & TimeOfDay & "' where TAG = 3")
                                    End If
                                    MsgBox("File transfer complete.")
                                    SW.WriteLine("SSIT_SERVER " & DateAndTime.Now)
                                End If
                            Else
                                MsgBox("Incorrect file name.", MsgBoxStyle.Information)
                            End If
                        End If
                        clsDialog.Dispose()
                    Case 4
                        Dim list As New ListBox
                        'If Not Directory.Exists("D:\ftp\Exe\") Then
                        '    Directory.CreateDirectory("D:\ftp\Exe\")
                        'End If

                        Dim OpenFolder As New FolderBrowserDialog
                        Dim clsDialog As New OpenFileDialog
                        clsDialog.Title = "Select File(s)"
                        clsDialog.Multiselect = True
                        clsDialog.FileName = "terms_and_conditions.txt"
                        clsDialog.Filter = "PDF (*.pdf)|*.pdf"

                        If clsDialog.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then

                            If clsDialog.SafeFileName = "terms_and_conditions.pdf" Then
                                _frmBlock.Show()
                                _frmPassword.ShowDialog()
                                If _frmPassword.x = 1 Then
                                    For Each sFile As String In clsDialog.FileNames

                                        Dim pos As Integer = sFile.LastIndexOf("\"c)
                                        list.Items.Add(sFile.Substring(pos + 1))
                                        up.UploadFile(sFile, "ftp://" & My.Settings.FTPIPSETTINGS & "/Terms/" & sFile.Substring(pos + 1), My.Settings.FTPUSERSETTINGS, My.Settings.FTPPASSSETTINGS)
                                        _frmPassword.x = 0
                                    Next

                                    If db.checkExistence("Select isnull(tag,0) from SSINFOTERMUPDATE where tag = 5") = 0 Then
                                        db.ExecuteSQLQuery("INSERT INTO SSINFOTERMUPDATE(TAG,DATEUPDATED,TIMEUPDATED) values(5,'" & Today.Date & "','" & TimeOfDay & "')")
                                    Else
                                        db.ExecuteSQLQuery("UPDATE SSINFOTERMUPDATE set DATEUPDATED ='" & Date.Today & "',TIMEUPDATED = '" & TimeOfDay & "' where TAG = 5")
                                    End If
                                    MsgBox("File Transfer Complete")
                                    SW.WriteLine("Terms and Condition " & DateAndTime.Now)
                                End If
                            End If
                        Else
                            MsgBox("Incorrect file name.", MsgBoxStyle.Information)
                        End If
                        clsDialog.Dispose()

                    Case 6
                        Dim list As New ListBox
                        'If Not Directory.Exists("D:\ftp\Exe\") Then
                        '    Directory.CreateDirectory("D:\ftp\Exe\")
                        'End If

                        Dim OpenFolder As New FolderBrowserDialog
                        Dim clsDialog As New OpenFileDialog
                        clsDialog.Title = "Select File(s)"
                        clsDialog.Multiselect = True
                        clsDialog.FileName = "SSIT_UPDATER.exe"
                        clsDialog.Filter = "Exe Files (*.EXE)|*.EXE"

                        If clsDialog.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then

                            If clsDialog.SafeFileName = "SSIT_UPDATER.exe" Then
                                _frmBlock.Show()
                                _frmPassword.ShowDialog()
                                If _frmPassword.x = 1 Then
                                    For Each sFile As String In clsDialog.FileNames

                                        Dim pos As Integer = sFile.LastIndexOf("\"c)
                                        list.Items.Add(sFile.Substring(pos + 1))
                                        up.UploadFile(sFile, "ftp://" & My.Settings.FTPIPSETTINGS & "/Updater/" & sFile.Substring(pos + 1), My.Settings.FTPUSERSETTINGS, My.Settings.FTPPASSSETTINGS)
                                        _frmPassword.x = 0
                                    Next

                                    If db.checkExistence("Select isnull(tag,0) from SSINFOTERMUPDATE where tag = 6") = 0 Then
                                        db.ExecuteSQLQuery("INSERT INTO SSINFOTERMUPDATE(TAG,DATEUPDATED,TIMEUPDATED) values(6,'" & Today.Date & "','" & TimeOfDay & "')")
                                    Else
                                        db.ExecuteSQLQuery("UPDATE SSINFOTERMUPDATE set DATEUPDATED ='" & Date.Today & "',TIMEUPDATED = '" & TimeOfDay & "' where TAG = 6")
                                    End If
                                    MsgBox("File Transfer Complete")
                                    SW.WriteLine("SSIT_UPDATER " & DateAndTime.Now)
                                End If
                            Else
                                MsgBox("Incorrect file name.", MsgBoxStyle.Information)
                            End If
                        End If
                        clsDialog.Dispose()
                        SW.Close()
                End Select



                'Dim files As List(Of String) = New List(Of String)

                'files.Add(OpenFileDialog1.FileName.ToString)

                'For Each file In files
                '    files.Add(OpenFileDialog1.FileName.ToString)
                '    up.UploadFile(OpenFileDialog1.FileName.ToString, "ftp://" & My.Settings.FTPIPSETTINGS & "/" & OpenFileDialog1.SafeFileName, My.Settings.FTPUSERSETTINGS, My.Settings.FTPPASSSETTINGS)


                'Next
            End Using
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub rbPDF_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbPDF.CheckedChanged
        FileName = 1
    End Sub

    Private Sub rbVideo_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbVideo.CheckedChanged
        FileName = 2
    End Sub

    Private Sub rbEXE_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbEXE.CheckedChanged
        FileName = 3
    End Sub

    Private Sub rbTerms_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbTerms.CheckedChanged
        FileName = 4
    End Sub

    Private Sub rbUpdater_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbUpdater.CheckedChanged
        FileName = 6
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        rbEXE.Checked = False
        rbPDF.Checked = False
        rbUpdater.Checked = False
        rbVideo.Checked = False
        rbTerms.Checked = False
    End Sub
End Class