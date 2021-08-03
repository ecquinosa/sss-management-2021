Public Class _frmAddBranch
    Dim db As New ConnectionString
    Dim x As Integer
    Dim y As Integer
    Public Sub clear()
        cbGroup.Text = Nothing
        cbCluster.Text = Nothing
        txtBranch.Text = Nothing
        txtBranchCode.Text = Nothing
    End Sub
    Public Sub edit()
        btnadd.Enabled = False
        btnEdit.Enabled = False
    End Sub
    Public Sub cancel()
        Select Case y
            Case 1
                disable()
                clear()
            Case Else
                Me.Dispose()
                Me.Close()
        End Select
     
    End Sub
    Public Sub add()
        btnEdit.Enabled = False
        btnadd.Enabled = False

    End Sub
    Public Sub enable()
        cbGroup.Enabled = True
        cbCluster.Enabled = True
        txtBranch.Enabled = True
        txtBranchCode.Enabled = True
    End Sub
    Public Sub disable()
        cbGroup.Enabled = False
        cbCluster.Enabled = False
        txtBranch.Enabled = False
        txtBranchCode.Enabled = False
    End Sub
    Public Sub loadBranches()
        db.FillListView(db.ExecuteSQLQuery("Select SSINFOTERMGROUP.GROUP_NM as 'Group',SSINFOTERMCLSTR.CLSTR_NM as 'Division',SSINFOTERMBR.BRANCH_NM as 'Branch' from SSINFOTERMBR INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSINFOTERMBR.CLSTR_CD INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSINFOTERMBR.GROUP_CD"), lvList)
    End Sub
    Private Sub _frmAddBranch_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        db.fillComboBox(db.ExecuteSQLQuery("select GROUP_NM from SSINFOTERMGROUP"), cbGroup)
        db.fillComboBox(db.ExecuteSQLQuery("select CLSTR_NM from SSINFOTERMCLSTR"), cbCluster)
    End Sub

    Private Sub cbGroup_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbGroup.SelectedIndexChanged
        db.fillComboBox(db.ExecuteSQLQuery("SELECT CLSTR_NM FROM SSINFOTERMCLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSINFOTERMCLSTR.GROUP_CD where SSINFOTERMGROUP.GROUP_NM = '" & cbGroup.Text & "'"), cbCluster)
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Try
            Select Case x
                Case 1
                    Dim getClusterCode As String = db.putSingleValue("select CLSTR_CD from SSINFOTERMCLSTR where CLSTR_NM = '" & cbCluster.Text & "'")
                    Dim getGroupCode As String = db.putSingleValue("select GROUP_CD from SSINFOTERMGROUP where GROUP_NM = '" & cbGroup.Text & "'")
                    If cbGroup.Text = "" Then
                        MsgBox("Group is empty.", MsgBoxStyle.Exclamation)
                    ElseIf cbCluster.Text = "" Then
                        MsgBox("Division is empty.", MsgBoxStyle.Exclamation)
                    ElseIf txtBranch.Text = "" Then
                        MsgBox("Branch is empty.", MsgBoxStyle.Exclamation)
                    ElseIf txtBranchCode.Text = "" Then
                        MsgBox("Branch code is empty.", MsgBoxStyle.Exclamation)
                    Else
                        db.ExecuteSQLQuery("INSERT INTO SSINFOTERMBR (BRANCH_CD,BRANCH_NM,CLSTR_CD,GROUP_CD) values ('" & txtBranchCode.Text & "','" & txtBranch.Text & "','" & getClusterCode & "','" & getGroupCode & "')")
                        MsgBox("Successfully saved.", MsgBoxStyle.Information)
                        y = 0
                        btnCancel.Text = "Close"
                    End If
                Case 2
                    Dim getClusterCode As String = db.putSingleValue("select CLSTR_CD from SSINFOTERMCLSTR where CLSTR_NM = '" & cbCluster.Text & "'")
                    Dim getGroupCode As String = db.putSingleValue("select GROUP_CD from SSINFOTERMGROUP where GROUP_NM = '" & cbGroup.Text & "'")
                    If cbGroup.Text = "" Then
                        MsgBox("Group is empty.", MsgBoxStyle.Exclamation)
                    ElseIf cbCluster.Text = "" Then
                        MsgBox("Division is empty.", MsgBoxStyle.Exclamation)
                    ElseIf txtBranch.Text = "" Then
                        MsgBox("Branch is empty.", MsgBoxStyle.Exclamation)
                    ElseIf txtBranchCode.Text = "" Then
                        MsgBox("Branch code is empty.", MsgBoxStyle.Exclamation)
                    Else
                        If db.checkExistence("Select BRANCH_CD FROM SSINFOTERMBR WHERE BRANCH_CD = '" & txtBranchCode.Text & "'") Then
                            db.ExecuteSQLQuery("UPDATE SSINFOTERMBR SET BRANCH_NM ='" & txtBranch.Text & "',BRANCH_CD = '" & txtBranchCode.Text & "',GROUP_NM = '" & getGroupCode & "',CLSTR_CD = '" & getClusterCode & "' where BRANCH_CD = '" & txtBranchCode.Text & "' ")
                        Else
                            MsgBox("Branch code doesn't exist.")
                        End If
                    End If

            End Select
           
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClear.Click
        _frmBlock.Close()
        Me.Dispose()
        Me.Close()
    End Sub

    Private Sub btnadd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnadd.Click
        y = 1
        x = 1
        btnCancel.Text = "Cancel"
        btnSave.Text = "Save"
        enable()
    End Sub

    Private Sub btnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        y = 1
        x = 2
        btnCancel.Text = "Cancel"
        btnSave.Text = "Update"
        enable()
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        cancel()
    End Sub

    Private Sub lvList_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lvList.DoubleClick
        Dim branchName As String = lvList.SelectedItems(0).SubItems(2).Text
        Dim GroupName As String = lvList.SelectedItems(0).SubItems(0).Text
        Dim Cluster As String = lvList.SelectedItems(0).SubItems(1).Text
        Dim BranchCode As String = db.putSingleValue("Select BRANCH_CD FROM SSINFOTERMBR WHERE BRANCH_NM = '" & branchName & "'")
        cbGroup.Text = GroupName
        txtBranch.Text = branchName
        cbCluster.Text = Cluster
        txtBranchCode.Text = BranchCode
    End Sub
End Class