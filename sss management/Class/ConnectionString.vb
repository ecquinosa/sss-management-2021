Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.Sql
Imports System.Text.RegularExpressions
Imports System.Data.Odbc

Public Class ConnectionString

    Public sql As String
    Public task As String

    'Public connString1 As String = My.Settings.Setting
    'Public connstring1 As String = "DSN=" & My.Settings.db_DSN & ";SERVER=" & My.Settings.db_Server & ";DATABASE=" & My.Settings.db_Name & ";UID=" & My.Settings.db_UName & ";PWD=" & My.Settings.db_Pass & ""
    'Public connString1 As String = "Data Source=" & tpd.DecryptData(My.Settings.db_Server) & ";Initial Catalog=" & tpd.DecryptData(My.Settings.db_Name) & ";User ID=" & tpd.DecryptData(My.Settings.db_UName) & ";Password=" & tpd.DecryptData(My.Settings.db_Pass) & ";"
    ' Public connString1 As String = "Data Source=" & (My.Settings.db_Server) & ";Initial Catalog=" & (My.Settings.db_Name) & ";User ID=" & (My.Settings.db_UName) & ";Password=" & (My.Settings.db_Pass) & ";"
    'Public conn As SqlConnection = New SqlConnection(connString1)
    Public connstring1 As String = "DSN=" & My.Settings.db_DSN & ";SERVER=" & My.Settings.db_Server & ";DATABASE=" & My.Settings.db_Name & ";UID=" & My.Settings.db_UName & ";PWD=" & My.Settings.db_Pass & ""
    Public conn As OdbcConnection = New OdbcConnection(connstring1)

    Private _connectionError As String

    Public Property connectionError() As String
        Get
            Return _connectionError
        End Get
        Set(ByVal value As String)
            _connectionError = value
        End Set
    End Property

    Public Function getDataTable(ByVal sql As String, ByVal tbl As String) As DataTable
        Dim dt As New DataTable

        Try

            If conn.State = ConnectionState.Open Then
                conn.Close()
                conn.Open()
            Else
                conn.Open()
            End If
            Dim da As OdbcDataAdapter = New OdbcDataAdapter(sql, conn)
            Dim ds As New DataSet
            da.Fill(ds, tbl)
            dt = ds.Tables(tbl)
        Catch ex As Exception
            'MsgBox("Error:" & ex.ToString)
            Dim errorLogs As String = ex.ToString
            errorLogs = errorLogs.Trim
            sql = "insert into tbl_mgmt_errorlogs values('" & My.Settings.mgmtIP & "', '" & My.Settings.mgmtID & "', '" & My.Settings.mgmtName & "', '" & My.Settings.mgmtBranch & "', '" & errorLogs _
                & "','" & "Class: Connection String" & "', '" & "Database connection error" & "', '" & Date.Today.ToShortDateString & "', '" & TimeOfDay & "') "
            ExecuteSQLQuery(sql)

            MsgBox("Database connection error, Please contact system Administrator! ", MsgBoxStyle.Information)
        Finally
            conn.Close()
        End Try
        Return dt
    End Function

    Public Function ExecuteSQLQuery(ByVal SQLQuery As String) As DataTable
        Dim sqlDT As New DataTable
        Try
            Dim sqlCon As New OdbcConnection(connstring1)
            Dim sqlDA As New OdbcDataAdapter(SQLQuery, sqlCon)
            Dim sqlCB As New OdbcCommandBuilder(sqlDA)
            sqlDA.SelectCommand.CommandTimeout = 0
            sqlDA.Fill(sqlDT)
        Catch ex As Exception
            'MsgBox("Program Error: " & ex.Message, MsgBoxStyle.Critical)
            Dim errorLogs As String = ex.ToString
            errorLogs = errorLogs.Trim
            sql = "insert into tbl_mgmt_errorlogs values('" & My.Settings.mgmtIP & "', '" & My.Settings.mgmtID & "', '" & My.Settings.mgmtName & "', '" & My.Settings.mgmtBranch & "', '" & errorLogs _
                & "','" & "Class: Connection String" & "', '" & "Database connection error" & "', '" & Date.Today.ToShortDateString & "', '" & TimeOfDay & "') "
            ExecuteSQLQuery(sql)
            MsgBox("Database connection error, Please contact system Administrator! ", MsgBoxStyle.Information)
        End Try

        Return sqlDT
    End Function

    Public Sub FillListView(ByVal sqlData As DataTable, ByVal lvList As ListView)
        Try


            lvList.Items.Clear()
            lvList.Columns.Clear()
            Dim i As Integer
            Dim j As Integer
            For i = 0 To sqlData.Columns.Count - 1
                lvList.Columns.Add(sqlData.Columns(i).ColumnName)
            Next i
            For i = 0 To sqlData.Rows.Count - 1
                lvList.Items.Add(sqlData.Rows(i).Item(0))
                For j = 1 To sqlData.Columns.Count - 1
                    If Not IsDBNull(sqlData.Rows(i).Item(j)) Then
                        lvList.Items(i).SubItems.Add(sqlData.Rows(i).Item(j))
                    Else
                        lvList.Items(i).SubItems.Add("")
                    End If

                Next j
            Next i
            For i = 0 To sqlData.Columns.Count - 1
                lvList.Columns(i).AutoResize(ColumnHeaderAutoResizeStyle.HeaderSize)
            Next i
        Catch ex As Exception
            MsgBox(ex.ToString)
            Dim errorLogs As String = ex.ToString
            errorLogs = errorLogs.Trim
            sql = "insert into tbl_mgmt_errorlogs values('" & My.Settings.mgmtIP & "', '" & My.Settings.mgmtID & "', '" & My.Settings.mgmtName & "', '" & My.Settings.mgmtBranch & "', '" & errorLogs _
                & "','" & "Class: Connection String" & "', '" & "Database connection error" & "', '" & Date.Today.ToShortDateString & "', '" & TimeOfDay & "') "
            ExecuteSQLQuery(sql)
            MsgBox("Database connection error, Please contact system Administrator! ", MsgBoxStyle.Information)
        End Try

    End Sub

    Public Function doNonQuery(ByVal sql As String, Optional ByVal process As String = "") As Boolean
        Try
            If conn.State = ConnectionState.Open Then
                conn.Close()
                conn.Open()
            Else
                conn.Open()
            End If
            Dim cmd As OdbcCommand = New OdbcCommand(sql, conn)
            cmd.ExecuteNonQuery()
            doNonQuery = True
        Catch ex As Exception
            doNonQuery = False

            Dim errorLogs As String = ex.ToString
            errorLogs = errorLogs.Trim
            sql = "insert into tbl_mgmt_errorlogs values('" & My.Settings.mgmtIP & "', '" & My.Settings.mgmtID & "', '" & My.Settings.mgmtName & "', '" & My.Settings.mgmtBranch & "', '" & errorLogs _
                & "','" & "Class: Connection String" & "', '" & "Database connection error" & "', '" & Date.Today.ToShortDateString & "', '" & TimeOfDay & "') "
            ExecuteSQLQuery(sql)
            MsgBox("Database connection error, Please contact system Administrator! ", MsgBoxStyle.Information)
        Finally
            conn.Close()
        End Try
    End Function
    Public Function checkExistence(ByVal sql As String) As Boolean
        Dim ans As Boolean = False
        Try
            If conn.State = ConnectionState.Open Then
                conn.Close()
                conn.Open()
            Else
                conn.Open()
            End If
            Dim cmd As OdbcCommand = New OdbcCommand(sql, conn)
            Dim rdr As OdbcDataReader = cmd.ExecuteReader
            If rdr.Read Then
                ans = True
            End If
        Catch ex As Exception
            'MsgBox("Unable to connect to server" & vbNewLine & "Please check your database connection", MsgBoxStyle.Exclamation)
            'MsgBox("Error on(checkExistence): " & ex.ToString)
            Dim errorLogs As String = ex.ToString
            errorLogs = errorLogs.Trim
            sql = "insert into tbl_mgmt_errorlogs values('" & My.Settings.mgmtIP & "', '" & My.Settings.mgmtID & "', '" & My.Settings.mgmtName & "', '" & My.Settings.mgmtBranch & "', '" & errorLogs _
                & "','" & "Class: Connection String" & "', '" & "Database connection error" & "', '" & Date.Today.ToShortDateString & "', '" & TimeOfDay & "') "
            ExecuteSQLQuery(sql)
            MsgBox("Database connection error, Please contact system Administrator! ", MsgBoxStyle.Information)
        Finally
            conn.Close()
        End Try
        Return ans
    End Function

    Public Function checkExistence2(ByVal sql As String) As Boolean
        Dim ans As Boolean = False
        Try
            If conn.State = ConnectionState.Open Then
                conn.Close()
                conn.Open()
            Else
                conn.Open()
            End If
            Dim cmd As OdbcCommand = New OdbcCommand(sql, conn)
            Dim rdr As OdbcDataReader = cmd.ExecuteReader
            If rdr.Read Then
                ans = True
            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
            'MsgBox("Unable to connect to server" & vbNewLine & "Please check your database connection", MsgBoxStyle.Exclamation)
            'MsgBox("Error on(checkExistence): " & ex.ToString)
            Dim errorLogs As String = ex.ToString
            errorLogs = errorLogs.Trim
            sql = "insert into tbl_mgmt_errorlogs values('" & My.Settings.mgmtIP & "', '" & My.Settings.mgmtID & "', '" & My.Settings.mgmtName & "', '" & My.Settings.mgmtBranch & "', '" & errorLogs _
                & "','" & "Class: Connection String" & "', '" & "Database connection error" & "', '" & Date.Today.ToShortDateString & "', '" & TimeOfDay & "') "
            ExecuteSQLQuery(sql)
            MsgBox("Database connection error, Please contact system Administrator! ", MsgBoxStyle.Information)
        Finally
            conn.Close()
        End Try
        Return ans
    End Function

    Public Function putSingleValue(ByVal sql As String, Optional ByVal tbl As String = "") As String
        Dim result As String = ""
        Try
            If tbl <> "" Then
                Dim dt As DataTable = getDataTable(sql, tbl)
                If dt.Rows.Count <> 0 Then
                    If Not IsDBNull(dt.Rows(0)(0)) Then
                        result = dt.Rows(dt.Rows.Count - 1)(0)
                    End If
                End If
            Else
                If conn.State = ConnectionState.Open Then
                    conn.Close()
                    conn.Open()
                Else
                    conn.Open()
                End If
                Dim cmd As OdbcCommand = New OdbcCommand(sql, conn)
                cmd.CommandTimeout = 0
                Dim rdr As OdbcDataReader = cmd.ExecuteReader
                If rdr.Read Then
                    If Not IsDBNull(rdr(0)) Then
                        result = rdr(0)
                    End If
                End If
            End If
        Catch ex As Exception
            'MsgBox("Error 1: " & ex.ToString)
            result = ""
            Dim errorLogs As String = ex.ToString
            errorLogs = errorLogs.Trim
            sql = "insert into tbl_mgmt_errorlogs values('" & My.Settings.mgmtIP & "', '" & My.Settings.mgmtID & "', '" & My.Settings.mgmtName & "', '" & My.Settings.mgmtBranch & "', '" & errorLogs _
                & "','" & "Class: Connection String" & "', '" & "Database connection error" & "', '" & Date.Today.ToShortDateString & "', '" & TimeOfDay & "') "
            ExecuteSQLQuery(sql)
            MsgBox("Database connection error, Please contact system Administrator! ", MsgBoxStyle.Information)
        Finally
            conn.Close()
        End Try
        Return result
    End Function
    Dim conStr As String
    Function webisconnected(ByVal constring As String) As Boolean
        conStr = constring
        Return isconnected()
    End Function

    Function isconnected() As Boolean
        Try
            If DBConnectionStatus() = True Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            'MsgBox(ex.ToString)
            Dim errorLogs As String = ex.ToString
            errorLogs = errorLogs.Trim
            sql = "insert into tbl_mgmt_errorlogs values('" & My.Settings.mgmtIP & "', '" & My.Settings.mgmtID & "', '" & My.Settings.mgmtName & "', '" & My.Settings.mgmtBranch & "', '" & errorLogs _
                & "','" & "Class: Connection String" & "', '" & "Database connection error" & "', '" & Date.Today.ToShortDateString & "', '" & TimeOfDay & "') "
            ExecuteSQLQuery(sql)
            MsgBox("Database connection error, Please contact system Administrator! ", MsgBoxStyle.Information)
        End Try

    End Function

    Private Function DBConnectionStatus() As Boolean
        Try
            Using conStr1 As New OdbcConnection(conStr)
                '("Server=" & txtServer.Text & ";" & _
                '"uid=" & txtLogin.Text & ";pwd=" & txtPassword.Text & "")
                conStr1.Close()
                conStr1.Open()
                Return (conStr1.State = ConnectionState.Open)

            End Using

        Catch e1 As OdbcException
            Return False
        Catch e2 As Exception
            Return False
        End Try
    End Function

    Public Sub FillListBox(ByVal sqlData As DataTable, ByVal lvList As ListBox)
        Try


            lvList.Items.Clear()
            Dim i As Integer
            Dim j As Integer
            For i = 0 To sqlData.Columns.Count - 1
                lvList.Items.Add(sqlData.Columns(i).ColumnName)
            Next i
            For i = 0 To sqlData.Rows.Count - 1
                If IsDBNull(sqlData.Rows(i).Item(0)) Then
                    lvList.Items.Add("")
                Else
                    lvList.Items.Add(sqlData.Rows(i).Item(0))
                End If

                For j = 1 To sqlData.Columns.Count - 1
                    If Not IsDBNull(sqlData.Rows(i).Item(j)) Then
                        'lvList.Items(i).SubItems.Add(sqlData.Rows(i).Item(j))
                    Else
                        'lvList.Items(i).SubItems.Add("")
                    End If
                Next j
            Next i
            'For i = 0 To sqlData.Columns.Count - 1
            '    lvList.Items(i).AutoResize(ColumnHeaderAutoResizeStyle.HeaderSize)
            'Next i
        Catch ex As Exception
            Dim errorLogs As String = ex.ToString
            errorLogs = errorLogs.Trim
            sql = "insert into tbl_mgmt_errorlogs values('" & My.Settings.mgmtIP & "', '" & My.Settings.mgmtID & "', '" & My.Settings.mgmtName & "', '" & My.Settings.mgmtBranch & "', '" & errorLogs _
                & "','" & "Class: Connection String" & "', '" & "Database connection error" & "', '" & Date.Today.ToShortDateString & "', '" & TimeOfDay & "') "
            ExecuteSQLQuery(sql)
            MsgBox("Database connection error, Please contact system Administrator! ", MsgBoxStyle.Information)
        End Try
    End Sub

    Public Sub fillComboBox(ByVal dt As DataTable, ByVal cb As ComboBox)
        Try
            cb.Items.Clear()

            For row As Integer = 0 To dt.Rows.Count - 1
                cb.Items.Add(dt.Rows(row)(0))
            Next
        Catch ex As Exception
            Dim errorLogs As String = ex.ToString
            errorLogs = errorLogs.Trim
            sql = "insert into tbl_mgmt_errorlogs values('" & My.Settings.mgmtIP & "', '" & My.Settings.mgmtID & "', '" & My.Settings.mgmtName & "', '" & My.Settings.mgmtBranch & "', '" & errorLogs _
                & "','" & "Class: Connection String" & "', '" & "Database connection error" & "', '" & Date.Today.ToShortDateString & "', '" & TimeOfDay & "') "
            ExecuteSQLQuery(sql)
            MsgBox("Database connection error, Please contact system Administrator! ", MsgBoxStyle.Information)
        End Try
    End Sub

    Function EmailAddressCheck(ByVal emailAddress As String) As Boolean
        Dim pattern As String = "^[a-zA-Z][\w\.-]*[a-zA-Z0-9]@[a-zA-Z0-9][\w\.-]*[a-zA-Z0-9]\.[a-zA-Z][a-zA-Z\.]*[a-zA-Z]$"
        Dim emailAddressMatch As Match = Regex.Match(emailAddress, pattern)
        If emailAddressMatch.Success Then
            EmailAddressCheck = True
        Else
            EmailAddressCheck = False
        End If
    End Function

    Public Function EncryptText(ByVal sSTR As System.String) As String
        Dim sTmp As System.String
        Dim sResult As System.String
        Dim iCnt As System.Int32

        sTmp = StrReverse(sSTR)
        sResult = ""
        For iCnt = 1 To Len(sTmp)
            sResult = sResult & Chr(Asc(Mid(sTmp, iCnt, 1)) + Asc("g"))
        Next
        EncryptText = sResult
    End Function

    Public Function DecryptText(ByVal sSTR As String) As String

        Dim sTmp As String
        Dim sResult As String
        Dim icnt As Integer

        sTmp = StrReverse(sSTR)

        sResult = ""

        For icnt = 1 To Len(sTmp)
            sResult = sResult & Chr(Asc(Mid(sTmp, icnt, 1)) - Asc("g"))
        Next

        DecryptText = sResult

    End Function

    Public Function putSingleNumber(ByVal sql As String, Optional ByVal tbl As String = "") As String
        Dim result As Double = 0.0
        Try
            If tbl <> "" Then
                Dim dt As DataTable = getDataTable(sql, tbl)
                If dt.Rows.Count <> 0 Then
                    If Not IsDBNull(dt.Rows(0)(0)) Then
                        result = Val(dt.Rows(dt.Rows.Count - 1)(0))
                    End If
                End If
            Else
                conn.Open()
                Dim cmd As OdbcCommand = New OdbcCommand(sql, conn)
                Dim rdr As OdbcDataReader = cmd.ExecuteReader
                If rdr.Read Then
                    If Not IsDBNull(rdr(0)) Then
                        result = rdr(0)
                    End If
                End If
            End If

        Catch ex As Exception
            'MsgBox("Error 1: " & ex.ToString)
            Dim errorLogs As String = ex.ToString
            errorLogs = errorLogs.Trim
            sql = "insert into tbl_mgmt_errorlogs values('" & My.Settings.mgmtIP & "', '" & My.Settings.mgmtID & "', '" & My.Settings.mgmtName & "', '" & My.Settings.mgmtBranch & "', '" & errorLogs _
                & "','" & "Class: Connection String" & "', '" & "Database connection error" & "', '" & Date.Today.ToShortDateString & "', '" & TimeOfDay & "') "
            ExecuteSQLQuery(sql)
            MsgBox("Database connection error, Please contact system Administrator! ", MsgBoxStyle.Information)
        Finally
            conn.Close()
        End Try
        Return result
    End Function

    Public Function selectTable(ByVal sql As String) As Boolean
        Dim ans As Boolean = False
        Try
            conn.Open()
            Dim cmd As OdbcCommand = New OdbcCommand(sql, conn)
            Dim rdr As OdbcDataReader = cmd.ExecuteReader
            If rdr.Read Then
                ans = True
            End If
        Catch ex As Exception
            'MsgBox("Error on(selectTable): " & ex.ToString)
            Dim errorLogs As String = ex.ToString
            errorLogs = errorLogs.Trim
            sql = "insert into tbl_mgmt_errorlogs values('" & My.Settings.mgmtIP & "', '" & My.Settings.mgmtID & "', '" & My.Settings.mgmtName & "', '" & My.Settings.mgmtBranch & "', '" & errorLogs _
                & "','" & "Class: Connection String" & "', '" & "Database connection error" & "', '" & Date.Today.ToShortDateString & "', '" & TimeOfDay & "') "
            ExecuteSQLQuery(sql)
            MsgBox("Database connection error, Please contact system Administrator! ", MsgBoxStyle.Information)
        Finally
            conn.Close()
        End Try
        Return ans
    End Function

    Public Sub FillDataGridView(ByVal selectCommand As String, ByVal dgv As DataGridView)

        Try

            Dim cnn As String = connString1
            Dim da = New OdbcDataAdapter(selectCommand, cnn)
            Dim cb As New OdbcCommandBuilder(da)
            Dim tbl As New DataTable()
            tbl.Locale = System.Globalization.CultureInfo.InvariantCulture
            da.Fill(tbl)
            dgv.DataSource = tbl
            dgv.AutoResizeColumns( _
                DataGridViewAutoSizeColumnsMode.ColumnHeader)
        Catch ex As OdbcException
            Dim errorLogs As String = ex.ToString
            errorLogs = errorLogs.Trim
            sql = "insert into tbl_mgmt_errorlogs values('" & My.Settings.mgmtIP & "', '" & My.Settings.mgmtID & "', '" & My.Settings.mgmtName & "', '" & My.Settings.mgmtBranch & "', '" & errorLogs _
                & "','" & "Class: Connection String" & "', '" & "Database connection error" & "', '" & Date.Today.ToShortDateString & "', '" & TimeOfDay & "') "
            ExecuteSQLQuery(sql)
            MsgBox("Database connection error, Please contact system Administrator! ", MsgBoxStyle.Information)
        End Try
    End Sub

    Public Function GenRecID(ByVal getGen As String)
        Try


            Dim id As String = ""
            For x As Integer = getGen To 999999
                Dim tempID As String = x
                If checkExistence("select * from tbl_mgmt_errorlogs where mgmtserverID='" & tempID & "'") = False Then
                    id = tempID
                    Exit For
                End If
            Next

            Return id
        Catch ex As Exception
            Dim errorLogs As String = ex.ToString
            errorLogs = errorLogs.Trim
            sql = "insert into tbl_mgmt_errorlogs values('" & My.Settings.mgmtIP & "', '" & My.Settings.mgmtID & "', '" & My.Settings.mgmtName & "', '" & My.Settings.mgmtBranch & "', '" & errorLogs _
                & "','" & "Class: Connection String" & "', '" & "Database connection error" & "', '" & Date.Today.ToShortDateString & "', '" & TimeOfDay & "') "
            ExecuteSQLQuery(sql)
            MsgBox("Database connection error, Please contact system Administrator! ", MsgBoxStyle.Information)
        End Try
    End Function

    Public Function GenRecIDMaternity(ByVal getGen As String)
        Try


            Dim id As String = ""
            For x As Integer = getGen To 999999
                Dim tempID As String = x
                If checkExistence("select * from tbl_Kiosk_Maternity where transactionIDmaternity='" & tempID & "'") = False Then
                    id = tempID
                    Exit For
                End If
            Next
            Return id
        Catch ex As Exception
            Dim errorLogs As String = ex.ToString
            errorLogs = errorLogs.Trim
            sql = "insert into tbl_mgmt_errorlogs values('" & My.Settings.mgmtIP & "', '" & My.Settings.mgmtID & "', '" & My.Settings.mgmtName & "', '" & My.Settings.mgmtBranch & "', '" & errorLogs _
                & "','" & "Class: Connection String" & "', '" & "Database connection error" & "', '" & Date.Today.ToShortDateString & "', '" & TimeOfDay & "') "
            ExecuteSQLQuery(sql)
            MsgBox("Database connection error, Please contact system Administrator! ", MsgBoxStyle.Information)
        End Try
    End Function

    Public Function GenFeedback(ByVal getGen As String)
        Try

            Dim id As String = ""
            For x As Integer = getGen To 999999
                Dim tempID As String = x
                If checkExistence("select * from tbl_Kiosk_feedback where autogenIDfeedback='" & tempID & "'") = False Then
                    id = tempID
                    Exit For
                End If
            Next
            Return id
        Catch ex As Exception
            Dim errorLogs As String = ex.ToString
            errorLogs = errorLogs.Trim
            sql = "insert into tbl_mgmt_errorlogs values('" & My.Settings.mgmtIP & "', '" & My.Settings.mgmtID & "', '" & My.Settings.mgmtName & "', '" & My.Settings.mgmtBranch & "', '" & errorLogs _
                & "','" & "Class: Connection String" & "', '" & "Database connection error" & "', '" & Date.Today.ToShortDateString & "', '" & TimeOfDay & "') "
            ExecuteSQLQuery(sql)
            MsgBox("Database connection error, Please contact system Administrator! ", MsgBoxStyle.Information)
        End Try
    End Function

    Public Function GenRegistration(ByVal getGen As String)
        Try

            Dim id As String = ""
            For x As Integer = getGen To 999999
                Dim tempID As String = x
                If checkExistence("select * from tbl_Kiosk_feedback where autogenIDregistration='" & tempID & "'") = False Then
                    id = tempID
                    Exit For
                End If
            Next
            Return id
        Catch ex As Exception
            Dim errorLogs As String = ex.ToString
            errorLogs = errorLogs.Trim
            sql = "insert into tbl_mgmt_errorlogs values('" & My.Settings.mgmtIP & "', '" & My.Settings.mgmtID & "', '" & My.Settings.mgmtName & "', '" & My.Settings.mgmtBranch & "', '" & errorLogs _
                & "','" & "Class: Connection String" & "', '" & "Database connection error" & "', '" & Date.Today.ToShortDateString & "', '" & TimeOfDay & "') "
            ExecuteSQLQuery(sql)
            MsgBox("Database connection error, Please contact system Administrator! ", MsgBoxStyle.Information)
        End Try
    End Function


    Public Sub AuditTrail(ByVal user_ID As String, ByVal username As String, ByVal _module As String, ByVal task As String, ByVal affected_Table As String, ByVal status As String, ByVal transaction_Date As String, ByVal transaction_Time As String)
        ExecuteSQLQuery("Insert into tbl_Audit_Trail(User_ID, User_Name, Module, Task, Affected_Table, Status, Transaction_Date, Transaction_Time) values('" & user_ID & "', '" & username & "', '" & _module & "', '" & task & "', '" & affected_Table & "', '" & status & "', '" & transaction_Date & "', '" & transaction_Time & "')")
    End Sub
    Public Function GenRegID(ByVal getReg As String)
        Try


            Dim id As String = ""
            For x As Integer = getReg To 999999
                Dim tempID As String = x
                If checkExistence("select * from tbl_mgmt_UserRegistration where mgmt_UserID='" & tempID & "'") = False Then
                    id = tempID
                    Exit For
                End If
            Next

            Return id
        Catch ex As Exception
            Dim errorLogs As String = ex.ToString
            errorLogs = errorLogs.Trim
            sql = "insert into tbl_mgmt_errorlogs values('" & My.Settings.mgmtIP & "', '" & My.Settings.mgmtID & "', '" & My.Settings.mgmtName & "', '" & My.Settings.mgmtBranch & "', '" & errorLogs _
                & "','" & "Class: Connection String" & "', '" & "Database connection error" & "', '" & Date.Today.ToShortDateString & "', '" & TimeOfDay & "') "
            ExecuteSQLQuery(sql)
            MsgBox("Database connection error, Please contact system Administrator! ", MsgBoxStyle.Information)
        End Try
    End Function


    Public Function RemoveDup(ByVal lvItems As ListView)
        Dim hTable As Hashtable = New Hashtable()
        Dim duplicateList As ArrayList = New ArrayList()
        Dim itm As ListViewItem
        For Each itm In lvItems.Items
            If hTable.ContainsKey(itm.Text) Then
                duplicateList.Add(itm)
            Else
                hTable.Add(itm.Text, String.Empty)
            End If
        Next

        For Each itm In duplicateList
            lvItems.Items.Remove(itm)
        Next

    End Function

End Class
