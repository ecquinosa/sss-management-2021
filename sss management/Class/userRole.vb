Public Class userRole
    Dim db As New ConnectionString
    Public Function convertToBit(ByVal getBool As String)
        Dim convertedBit As String
        Select Case getBool

            Case "True"
                convertedBit = 1
            Case "False"
                convertedBit = 0
            Case Else
                convertedBit = 0
        End Select
        Return convertedBit

    End Function

    Public Function userMgmtAll(ByVal getUserRole As String)
        Dim mgmtAll As String = db.putSingleValue("Select user_mgmt_all from tbl_mgmt_userrole1 where user_role = '" & getUserRole & "' ")
        Return mgmtAll
    End Function

    Public Function userMgmtAddEdit(ByVal getUserRole As String)
        Dim mgmtMonitoring As String = db.putSingleValue("Select view_monitoring from tbl_mgmt_userrole1 where user_role = '" & getUserRole & "' ")
        mgmtMonitoring = convertToBit(mgmtMonitoring)

        Dim mgmtReports As String = db.putSingleValue("Select view_reports from tbl_mgmt_userrole1 where user_role = '" & getUserRole & "' ")
        mgmtReports = convertToBit(mgmtReports)

        Dim mgmtKiosk As String = db.putSingleValue("Select view_kioskmgmt from tbl_mgmt_userrole1 where user_role = '" & getUserRole & "' ")
        mgmtKiosk = convertToBit(mgmtKiosk)

        Dim mgmtView As String = db.putSingleValue("Select view_reports from tbl_mgmt_userrole1 where user_role = '" & getUserRole & "' ")
        mgmtView = convertToBit(mgmtView)

        Dim mgmtUser As String = db.putSingleValue("Select view_usermgmt from tbl_mgmt_userrole1 where user_role = '" & getUserRole & "' ")
        mgmtUser = convertToBit(mgmtUser)

        Dim mgmtUserRole As String = db.putSingleValue("Select view_userrole from tbl_mgmt_userrole1 where user_role = '" & getUserRole & "' ")
        mgmtUserRole = convertToBit(mgmtUserRole)

        Dim mgmtDB As String = db.putSingleValue("Select view_dbset from tbl_mgmt_userrole1 where user_role = '" & getUserRole & "' ")
        mgmtDB = convertToBit(mgmtDB)

        Dim userMgmtFinal As String = mgmtMonitoring & mgmtKiosk & mgmtView & mgmtUser & mgmtUserRole & mgmtDB
        Return userMgmtFinal
    End Function



End Class
