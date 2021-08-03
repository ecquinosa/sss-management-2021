Public Class auditModule
    Dim formName As String
    Dim db As New ConnectionString
    Dim autogen As String
    'List of Modules

    Public Function getModuleName(ByVal moduleTag As String)
        Select Case moduleTag

            Case "1"
                formName = "Login Form"
            Case "2"
                formName = "Management Server Settings Form"
            Case "3"
                formName = "Main Form"
            Case "4"
                formName = "User Registration Form"
            Case "5"
                formName = "Kiosk Registration Form"
            Case "6"
                formName = "Monitoring Management Form"
            Case "7"
                formName = "Database Settings Form"
            Case "8"
                formName = "System Settings Form"
            Case "9"
                formName = "User Role Form"
            Case "10"

            Case "1.1"
                formName = "Add Position Form"
            Case "1.2"
                formName = "Add User Type Form"
            Case "1.3"
                formName = "Change Password Form"

            Case Else
                formName = "No form selected"
        End Select
        Return formName
    End Function

    Public Function getAutoID()
        Try
            Dim getGen As String = db.putSingleNumber("select autogenID from tbl_Mgmt_auditTrail")
            If getGen = 0 Or getGen = "" Then
                autogen = My.Settings.firstGen
                My.Settings.autoGenID = autogen
            Else
                autogen = db.GenRecID(getGen)
                My.Settings.autoGenID = autogen
            End If

            Return autogen
        Catch ex As Exception
            MsgBox("Connection Time-out")
        End Try
    End Function

End Class
