Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports CrystalDecisions.Windows.Forms
Public Class _frmFBK
    Dim dt As DataTable
    Public cryRpt As New ReportDocument
    Dim db As New ConnectionString
    Dim FBClick As Integer
    Dim string1 As String
    Public getFileName As String
    Dim PorE As Integer
    Public Sub operationFileTrans()

        Dim newview As New CrystalReportViewer
        Dim date1 As Date = dtpFrom.Value
        Dim date2 As Date = dtpTo.Value
        Try

            Select Case cbFB.Text
                Case "SSIT FEEDBACK"
                    getFileName = cbFB.Text
                    dt = db.ExecuteSQLQuery("SELECT NAME,ADD_1,ADD_2,ADDRESS_TYP,EMAIL,POST_CD,SSRATE1,SSRATE2,SSRATE3,SSRATE4,SSRATE5,SSRATE6,HELP_TAG,COMMNT_TAG,ENCODE_DT,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,KIOSK_ID FROM SSTRANSINFOTERMFBKIOSK INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSINFOTERMFBKIOSK.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSTRANSINFOTERMFBKIOSK.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSTRANSINFOTERMFBKIOSK.DIVSN where NAME = '" & lblName.Text & "' and ENCODE_DT = '" & lbldate.Text & "'")
                    Dim rpt As New PRINTKIOSKFB
                    cryRpt = rpt
                    '  cryRpt.Refresh()
                    cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                    newview.ReportSource = cryRpt
                    cryRpt.SetParameterValue("@getName", lblName.Text)
                    cryRpt.SetParameterValue("@getDate", lbldate.Text)
                Case "WEBSITE FEEDBACK"
                    getFileName = cbFB.Text
                    dt = db.ExecuteSQLQuery("SELECT NAME,EMAIL,ADDRESS_TYP,ADD_1,ADD_2,POST_CD, SSRATE1,SSRATE2,SSRATE3,SSRATE4,SSRATE5,SSRATE6,SSRATE7,VST_TAG,REASN_TAG,INFO_TAG,COMMNT_TAG,SSTRANSINFOTERMFB.ENCODE_DT,SSINFOTERMBR.BRANCH_NM,SSINFOTERMCLSTR.CLSTR_NM,SSINFOTERMGROUP.GROUP_NM,KIOSK_ID FROM SSTRANSINFOTERMFB INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSINFOTERMFB.BRANCH_CD INNER JOIN SSINFOTERMCLSTR ON SSINFOTERMCLSTR.CLSTR_CD = SSTRANSINFOTERMFB.CLSTR INNER JOIN SSINFOTERMGROUP ON SSINFOTERMGROUP.GROUP_CD = SSTRANSINFOTERMFB.DIVSN where NAME = '" & lblName.Text & "' and ENCODE_DT = '" & lbldate.Text & "'")
                    Dim rpt As New PRINTKIOSKWEB
                    cryRpt = rpt
                    '  cryRpt.Refresh()
                    cryRpt.SetDatabaseLogon(My.Settings.db_UName, My.Settings.db_Pass)
                    newview.ReportSource = cryRpt
                    cryRpt.SetParameterValue("@getName", lblName.Text)
                    cryRpt.SetParameterValue("@getDate", lbldate.Text)
            End Select

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub
    Public Sub labelShow()
        lblName.Visible = True
        lbladd.Visible = True
        lblEmail.Visible = True
        lblPcode.Visible = True
        lbladd1.Visible = True
        lbldate.Visible = True
        lblBranch.Visible = True
        lblKiosk_ID.Visible = True
    End Sub
    Public Sub labelhide()
        lblName.Visible = False
        lbladd.Visible = False
        lblEmail.Visible = False
        lblPcode.Visible = False
        lbladd1.Visible = False
        lbldate.Visible = False
        lblBranch.Visible = False
        lblKiosk_ID.Visible = False
    End Sub
    Public Sub fillDataWebsite()
        Try
            Dim Name As String = dgvFB.CurrentRow.Cells(0).Value
            lbldate.Text = db.putSingleValue("Select ENCODE_DT from SSTRANSINFOTERMFB where NAME = '" & Name & "'")
            lblName.Text = db.putSingleValue("Select NAME from SSTRANSINFOTERMFB where NAME = '" & Name & "'")
            lblEmail.Text = db.putSingleValue("Select EMAIL from SSTRANSINFOTERMFB where NAME  = '" & Name & "'")
            lblKiosk_ID.Text = db.putSingleValue("Select KIOSK_ID from SSTRANSINFOTERMFB where NAME = '" & Name & "'")
            lblBranch.Text = db.putSingleValue("Select SSINFOTERMBR.BRANCH_NM from SSTRANSINFOTERMFB INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSINFOTERMFB.BRANCH_CD where NAME = '" & Name & "'")
            Dim b As String = db.putSingleValue("Select ADDRESS_TYP From SSTRANSINFOTERMFB where NAME = '" & Name & "'")
            Select Case b
                Case "1"
                    lbladd.Text = "Business Address"
                Case "0"
                    lbladd.Text = "Home Address"
            End Select
            lbladd1.Text = db.putSingleValue("Select ADD_1 from SSTRANSINFOTERMFB where NAME = '" & Name & "'") & vbNewLine & _
            db.putSingleValue("Select ADD_2 from SSTRANSINFOTERMFB where NAME = '" & Name & "'")
            lblPcode.Text = db.putSingleValue("Select POST_CD from SSTRANSINFOTERMFB where NAME = '" & Name & "'")
            lblqq1.Text = "1. Ability to inform members of their" & vbNewLine & _
        "benefits and privileges under the System," & vbNewLine & _
        "as well as developments or new programs." & vbNewLine & _
         "Rate :" & db.putSingleValue("Select SSRATE1 from SSTRANSINFOTERMFB where NAME = '" & Name & "'")
            lblqq2.Text = "2. Attractiveness of design or layout." & vbNewLine & _
         "Rate :" & db.putSingleValue("Select SSRATE2 fROM SSTRANSINFOTERMFB Where NAME = '" & Name & "'")
            lblqq3.Text = "3. Organization of information or" & vbNewLine & _
        "categories." & vbNewLine & _
         "Rate :" & db.putSingleValue("Select SSRATE3 FROM SSTRANSINFOTERMFB where NAME = '" & Name & "'")
            lblqq4.Text = "4. Speed of downloading " & vbNewLine & _
        "material." & vbNewLine & _
         "Rate :" & db.putSingleValue("Select SSRATE4 FROM SSTRANSINFOTERMFB where NAME = '" & Name & "'")
            lblqq5.Text = "5. Usefulness of information." & vbNewLine & _
         "Rate :" & db.putSingleValue("Select SSRATE5 FROM SSTRANSINFOTERMFB where NAME = '" & Name & "'")
            lblqq6.Text = "6. Adequacy or exhaustiveness " & vbNewLine & _
         "of information provided." & vbNewLine & _
          "Rate :" & db.putSingleValue("Select SSRATE6 from SSTRANSINFOTERMFB where NAME = '" & Name & "'")
            lblqq7.Text = "7. Ease of locating topics." & vbNewLine & _
         "Rate :" & db.putSingleValue("Select SSRATE7 from SSTRANSINFOTERMFB where NAME = '" & Name & "'")
            Dim a As String = db.putSingleValue("Select VST_TAG from SSTRANSINFOTERMFB where NAME = '" & Name & "'")
            Select Case a
                Case "1"
                    lblVisit.Text = "Do you intend to visit the site again?" & vbNewLine & _
                        "Answer: YES"
                Case "0"
                    lblVisit.Text = "Do you intend to visit the site again?" & vbNewLine & _
                        "Answer: NO"
                Case Else
                    lblVisit.Text = "Do you intend to visit the site again?" & vbNewLine & _
                        "Answer: UNDECIDED"
            End Select
            rtbWhy.Text = db.putSingleValue("Select REASN_TAG from SSTRANSINFOTERMFB where NAME = '" & Name & "'")
            rtbComments.Text = db.putSingleValue("Select COMMNT_TAG from SSTRANSINFOTERMFB where NAME = '" & Name & "'")
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub
    Public Sub loadDataWebsite()
        db.FillDataGridView("Select NAME as  'Name', EMAIL as 'Email',case when ADDRESS_TYP = 0 then 'Home Address' else 'Business Address' end as [Address type],ADD_1 + ',' + ADD_2 as Address, POST_CD as 'Postal Code',SSINFOTERMBR.BRANCH_NM as 'Branch',KIOSK_ID as 'Terminal ID',SSRATE1 as 'Q1', SSRATE2 as 'Q2', SSRATE3 as 'Q3', SSRATE4 as 'Q4', SSRATE5 as 'Q5', SSRATE6 as 'Q6', SSRATE7 as 'Q7',case when VST_TAG = 1 then 'YES' when VST_TAG = 0 then 'NO' ELSE 'UNDECIDED' end as [Re-visit],REASN_TAG as 'reason',ENCODE_DT as 'Date Submitted' from SSTRANSINFOTERMFB INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSINFOTERMFB.BRANCH_CD where cast(ENCODE_DT as DATE) between '" & dtpFrom.Value & "' and '" & dtpTo.Value & "'ORDER BY ENCODE_DT", dgvFB)
    End Sub
    Public Sub loadDataSSIT()
        db.FillDataGridView("Select NAME as 'Name', EMAIL as 'Email',case when ADDRESS_TYP = 0 then 'Home Address' else 'Business Address' end as [Address type], ADD_1 +','+ADD_2 as Address,POST_CD as 'Postal Code',SSINFOTERMBR.BRANCH_NM as 'Branch',KIOSK_ID as 'Terminal ID', SSRATE1 as 'Q1', SSRATE2 as 'Q2', SSRATE3 as 'Q3', SSRATE4 as 'Q4', SSRATE5 as 'Q5', SSRATE6 as 'Q6', case when HELP_TAG = 1 then 'YES' else 'NO' end as Helpful,COMMNT_TAG as 'Comment',ENCODE_DT as 'Date Submitted' from SSTRANSINFOTERMFBKIOSK INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSINFOTERMFBKIOSK.BRANCH_CD where cast(ENCODE_DT as DATE) between '" & dtpFrom.Value & "' and '" & dtpTo.Value & "'ORDER BY ENCODE_DT", dgvFB)
    End Sub
    Public Sub fillDataSSIT()
        Try

            Dim Name As String = dgvFB.CurrentRow.Cells(0).Value
            lbldate.Text = db.putSingleValue("Select ENCODE_DT from SSTRANSINFOTERMFBKIOSK where NAME = '" & Name & "'")
            lblName.Text = db.putSingleValue("Select NAME from SSTRANSINFOTERMFBKIOSK where NAME = '" & Name & "'")
            lblEmail.Text = db.putSingleValue("Select EMAIL from SSTRANSINFOTERMFBKIOSK where NAME  = '" & Name & "'")
            lblKiosk_ID.Text = db.putSingleValue("Select KIOSK_ID from SSTRANSINFOTERMFBKIOSK where NAME = '" & Name & "'")
            lblBranch.Text = db.putSingleValue("Select SSINFOTERMBR.BRANCH_NM from SSTRANSINFOTERMFBKIOSK INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSINFOTERMFBKIOSK.BRANCH_CD where NAME = '" & Name & "'")
            Dim b As String = db.putSingleValue("Select ADDRESS_TYP From SSTRANSINFOTERMFBKIOSK where NAME = '" & Name & "'")
            Select Case b
                Case "1"
                    lbladd.Text = "Business Address"
                Case "0"
                    lbladd.Text = "Home Address"
            End Select
            lbladd1.Text = db.putSingleValue("Select ADD_1 from SSTRANSINFOTERMFBKIOSK where NAME = '" & Name & "'") & vbNewLine & _
                db.putSingleValue("Select ADD_2 from SSTRANSINFOTERMFBKIOSK where NAME = '" & Name & "'")

            lblPcode.Text = db.putSingleValue("Select POST_CD from SSTRANSINFOTERMFBKIOSK where NAME = '" & Name & "'")
            lblqq1.Text = "1. Attractiveness of design or layout" & vbNewLine & _
               "Rate :" & db.putSingleValue("Select SSRATE1 from SSTRANSINFOTERMFBKIOSK Where Name = '" & Name & "'")
            lblqq2.Text = "2. Organization of information or categories :" & vbNewLine & _
               "Rate :" & db.putSingleValue("Select SSRATE2 from SSTRANSINFOTERMFBKIOSK Where Name = '" & Name & "'")
            lblqq3.Text = "3. Speed of loading materials to be viewed :" & vbNewLine & _
               "Rate :" & db.putSingleValue("Select SSRATE3 from SSTRANSINFOTERMFBKIOSK Where Name = '" & Name & "'")
            lblqq4.Text = "4. Exhaustiveness of information provided :" & vbNewLine & _
               "Rate :" & db.putSingleValue("Select SSRATE4 from SSTRANSINFOTERMFBKIOSK Where Name = '" & Name & "'")
            lblqq5.Text = "5. Ease of locating topics " & vbNewLine & _
               "Rate :" & db.putSingleValue("Select SSRATE5 from SSTRANSINFOTERMFBKIOSK Where Name = '" & Name & "'")
            lblqq6.Text = " The SSS in terms of delivery service to member? :" & vbNewLine & _
               "Rate :" & db.putSingleValue("Select SSRATE6 from SSTRANSINFOTERMFBKIOSK Where Name = '" & Name & "'")
            Dim a As String = db.putSingleValue("Select HELP_TAG from SSTRANSINFOTERMFBKIOSK Where NAME = '" & Name & "'")
            Select Case a
                Case 1
                    lblVisit.Text = "Do you find SSS Information Terminal" & vbNewLine & _
     "helpful?" & vbNewLine & _
     "Answer: YES"
                Case 0
                    lblVisit.Text = "Do you find SSS Information Terminal " & vbNewLine & _
    "helpful?" & vbNewLine & _
    "Answer: NO"
            End Select

            rtbWhy.Text = db.putSingleValue("Select COMMNT_TAG from SSTRANSINFOTERMFBKIOSK Where Name = '" & Name & "'")
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
    Private Sub _frmFBK_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        gpPI.Visible = False
        gpQ.Visible = False
        gpCS.Visible = False
        labelhide()
        cbFB.Text = "SSIT FEEDBACK"
        db.FillDataGridView("Select NAME as 'Name', EMAIL as 'Email',case when ADDRESS_TYP = 0 then 'Home Address' else 'Business Address' end as [Address type], ADD_1 +','+ADD_2 as Address,POST_CD as 'Postal Code',SSINFOTERMBR.BRANCH_NM AS 'BRANCH',KIOSK_ID as 'Terminal ID', SSRATE1 as 'Q1', SSRATE2 as 'Q2', SSRATE3 as 'Q3', SSRATE4 as 'Q4', SSRATE5 as 'Q5', SSRATE6 as 'Q6', case when HELP_TAG = 1 then 'YES' else 'NO' end as Helpful,COMMNT_TAG as 'Comment',ENCODE_DT as 'Date Submitted' from SSTRANSINFOTERMFBKIOSK INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSINFOTERMFBKIOSK.BRANCH_CD where cast(ENCODE_DT as date) between '" & dtpFrom.Value & "' and '" & dtpFrom.Value & "' order by ENCODE_DT", dgvFB)

    End Sub
    Private Sub cbFB_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbFB.SelectedIndexChanged
        gpPI.Visible = True
        gpQ.Visible = True
        gpCS.Visible = True
        Try
            Select Case cbFB.Text

                Case "SSIT FEEDBACK"
                    cbFields.Items.Clear()
                    cbFields.Items.Add("Name")
                    cbFields.Items.Add("Email")
                    cbFields.Items.Add("Postal Code")
                    cbFields.Items.Add("Branch")
                    cbFields.Items.Add("Terminal ID")
                    cbFields.Items.Add("No")
                    cbFields.Items.Add("Yes")
                    cbFields.Items.Add("Show all")


                    labelhide()
                    rtbComments.Visible = False
                    lblqq7.Visible = False
                    Label17.Visible = False
                    FBClick = 1
                    loadDataSSIT()
                    lblqq1.Text = "1. Attractiveness of design or layout"
                    lblqq2.Text = "2. Organization of information or categories"
                    lblqq3.Text = "3. Speed of loading materials to be viewed "
                    lblqq4.Text = "4. Exhaustiveness of information provided "
                    lblqq5.Text = "5. Ease of locating topics "
                    lblqq6.Text = " The SSS in terms of delivery service to member?"
                    lblVisit.Text = "Do you find SSS Information Terminal" & vbNewLine & _
     "helpful?" & vbNewLine & _
                     "Answer :"
                Case "WEBSITE FEEDBACK"
                    cbFields.Items.Clear()
                    cbFields.Items.Add("Name")
                    cbFields.Items.Add("Email")
                    cbFields.Items.Add("Postal Code")
                    cbFields.Items.Add("Branch")
                    cbFields.Items.Add("Terminal ID")
                    cbFields.Items.Add("No")
                    cbFields.Items.Add("Yes")
                    cbFields.Items.Add("Undecided")
                    cbFields.Items.Add("Show all")

                    labelhide()
                    rtbComments.Visible = True
                    lblqq7.Visible = True
                    Label17.Visible = True
                    FBClick = 0
                    loadDataWebsite()
                    lblqq1.Text = "1. Ability to inform members of their" & vbNewLine & _
        "benefits and privileges under the System," & vbNewLine & _
        "as well as developments or new programs. :"

                    lblqq2.Text = "2. Attractiveness of design or layout."
                    lblqq3.Text = "3. Organization of information or" & vbNewLine & _
        "categories."
                    lblqq4.Text = "4. Speed of downloading " & vbNewLine & _
        "material."
                    lblqq5.Text = "5. Usefulness of information."
                    lblqq6.Text = "6. Adequacy or exhaustiveness " & vbNewLine & _
        "of information provided."
                    lblqq7.Text = "7. Ease of locating topics."
                    lblVisit.Text = "Do you intend to visit the site again?" & vbNewLine & _
                     "Answer :"
            End Select
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
    Private Sub dgvFB_DoubleClick_1(ByVal sender As System.Object, ByVal e As System.EventArgs)
        labelShow()
        Try
            Select Case FBClick

                Case 1
                    fillDataSSIT()
                Case 0
                    fillDataWebsite()
            End Select
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub dtpFrom_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpFrom.ValueChanged
        Try
            Dim date1 As Date = dtpFrom.Value
            Dim date2 As Date = dtpTo.Value
            date1.ToShortDateString()
            date2.ToShortDateString()
            Select Case FBClick

                Case 0
                    If date2 < date1 Then
                        MsgBox("Invalid date range.")
                        dtpFrom.Value = Date.Today
                    Else
                        db.FillDataGridView("Select NAME as  'Name', EMAIL as 'Email',case when ADDRESS_TYP = 0 then 'Home Address' else 'Business Address' end as [Address type],ADD_1 + ',' + ADD_2 as Address, POST_CD as 'Postal Code',SSINFOTERMBR.BRANCH_NM as 'Branch',KIOSK_ID as 'Terminal ID',SSRATE1 as 'Q1', SSRATE2 as 'Q2', SSRATE3 as 'Q3', SSRATE4 as 'Q4', SSRATE5 as 'Q5', SSRATE6 as 'Q6', SSRATE7 as 'Q7',case when VST_TAG = 1 then 'YES' when VST_TAG = 0 then 'NO' ELSE 'UNDECIDED' end as [Re-visit],REASN_TAG as 'reason',ENCODE_DT as 'Date Submitted' from SSTRANSINFOTERMFB INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSINFOTERMFB.BRANCH_CD where cast(ENCODE_DT as DATE) between '" & dtpFrom.Value & "' and '" & dtpTo.Value & "' ORDER BY ENCODE_DT", dgvFB)
                    End If
                Case 1
                    If date2 < date1 Then
                        MsgBox("Invalid Date Range")
                        dtpFrom.Value = Date.Today
                    Else
                        db.FillDataGridView("Select NAME as 'Name', EMAIL as 'Email',case when ADDRESS_TYP = 0 then 'Home Address' else 'Business Address' end as [Address type], ADD_1 +','+ADD_2 as Address,POST_CD as 'Postal Code',SSINFOTERMBR.BRANCH_NM as 'Branch',KIOSK_ID as 'Terminal ID', SSRATE1 as 'Q1', SSRATE2 as 'Q2', SSRATE3 as 'Q3', SSRATE4 as 'Q4', SSRATE5 as 'Q5', SSRATE6 as 'Q6', case when HELP_TAG = 1 then 'YES' else 'NO' end as Helpful,COMMNT_TAG as 'Comment',ENCODE_DT as 'Date Submitted' from SSTRANSINFOTERMFBKIOSK INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSINFOTERMFBKIOSK.BRANCH_CD where cast(ENCODE_DT as DATE) between '" & dtpFrom.Value & "' and '" & dtpTo.Value & "' ORDER BY ENCODE_DT", dgvFB)
                    End If
            End Select
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Private Sub dtpTo_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpTo.ValueChanged
        Try
            Dim date1 As Date = dtpFrom.Value
            Dim date2 As Date = dtpTo.Value
            date1.ToShortDateString()
            date2.ToShortDateString()
            Select Case FBClick

                Case 0
                    If date2 < date1 Then
                        MsgBox("Invalid Date Range")
                        dtpTo.Value = Date.Today
                    Else
                        db.FillDataGridView("Select NAME as 'Name', EMAIL as 'Email',case when ADDRESS_TYP = 0 then 'Home Address' else 'Business Address' end as [Address type],ADD_1 + ',' + ADD_2 as Address, POST_CD as 'Postal Code',SSINFOTERMBR.BRANCH_NM as 'Branch',KIOSK_ID as 'Terminal ID',SSRATE1 as 'Q1', SSRATE2 as 'Q2', SSRATE3 as 'Q3', SSRATE4 as 'Q4', SSRATE5 as 'Q5', SSRATE6 as 'Q6', SSRATE7 as 'Q7',case when VST_TAG = 1 then 'YES' when VST_TAG = 0 then 'NO' ELSE 'UNDECIDED' end as [Re-visit],REASN_TAG as 'reason',ENCODE_DT as 'Date Submitted' from SSTRANSINFOTERMFB INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSINFOTERMFB.BRANCH_CD where cast(ENCODE_DT as DATE) between '" & dtpFrom.Value & "' and '" & dtpTo.Value & "' ORDER BY ENCODE_DT ", dgvFB)
                    End If
                Case 1
                    If date2 < date1 Then
                        MsgBox("Invalid Date Range")
                        dtpTo.Value = Date.Today
                    Else
                        db.FillDataGridView("Select NAME as 'Name', EMAIL as 'Email',case when ADDRESS_TYP = 0 then 'Home Address' else 'Business Address' end as [Address type], ADD_1 +','+ADD_2 as Address,POST_CD as 'Postal Code',SSINFOTERMBR.BRANCH_NM as 'Branch',KIOSK_ID as 'Terminal ID', SSRATE1 as 'Q1', SSRATE2 as 'Q2', SSRATE3 as 'Q3', SSRATE4 as 'Q4', SSRATE5 as 'Q5', SSRATE6 as 'Q6', case when HELP_TAG = 1 then 'YES' else 'NO' end as Helpful,COMMNT_TAG as 'Comment',ENCODE_DT as 'Date Submitted' from SSTRANSINFOTERMFBKIOSK INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSINFOTERMFBKIOSK.BRANCH_CD where cast(ENCODE_DT as DATE) between '" & dtpFrom.Value & "' and '" & dtpTo.Value & "' ORDER BY ENCODE_DT", dgvFB)
                    End If

            End Select
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click

        Dim date1 As Date = dtpFrom.Value
        Dim date2 As Date = dtpTo.Value
        Try
            Select Case FBClick

                Case 0
                    Select Case cbFields.Text

                        Case "Name"
                            db.FillDataGridView("Select NAME as  'Name', EMAIL as 'Email',case when ADDRESS_TYP = 0 then 'Home Address' else 'Business Address' end as [Address type],ADD_1 + ',' + ADD_2 as Address, POST_CD as 'Postal Code',SSINFOTERMBR.BRANCH_NM as 'Branch',KIOSK_ID as 'Terminal ID',SSRATE1 as 'Q1', SSRATE2 as 'Q2', SSRATE3 as 'Q3', SSRATE4 as 'Q4', SSRATE5 as 'Q5', SSRATE6 as 'Q6', SSRATE7 as 'Q7',case when VST_TAG = 1 then 'YES' when VST_TAG = 0 then 'NO' ELSE 'UNDECIDED' end as [Re-visit],REASN_TAG as 'reason',ENCODE_DT as 'Date Submitted' from SSTRANSINFOTERMFB INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSINFOTERMFB.BRANCH_CD where cast(ENCODE_DT as DATE) between '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and NAME LIKE '%" & txtSearch.Text & "%'ORDER BY ENCODE_DT", dgvFB)
                        Case "Email"
                            db.FillDataGridView("Select NAME as  'Name', EMAIL as 'Email',case when ADDRESS_TYP = 0 then 'Home Address' else 'Business Address' end as [Address type],ADD_1 + ',' + ADD_2 as Address, POST_CD as 'Postal Code',SSINFOTERMBR.BRANCH_NM as 'Branch',KIOSK_ID as 'Terminal ID',SSRATE1 as 'Q1', SSRATE2 as 'Q2', SSRATE3 as 'Q3', SSRATE4 as 'Q4', SSRATE5 as 'Q5', SSRATE6 as 'Q6', SSRATE7 as 'Q7',case when VST_TAG = 1 then 'YES' when VST_TAG = 0 then 'NO' ELSE 'UNDECIDED' end as [Re-visit],REASN_TAG as 'reason',ENCODE_DT as 'Date Submitted' from SSTRANSINFOTERMFB INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSINFOTERMFB.BRANCH_CD where cast(ENCODE_DT as DATE) between '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and EMAIL LIKE '%" & txtSearch.Text & "%'ORDER BY ENCODE_DT", dgvFB)
                        Case "Postal Code"
                            db.FillDataGridView("Select NAME as  'Name', EMAIL as 'Email',case when ADDRESS_TYP = 0 then 'Home Address' else 'Business Address' end as [Address type],ADD_1 + ',' + ADD_2 as Address, POST_CD as 'Postal Code',SSINFOERMBR.BRANCH_NM as 'Branch',KIOSK_ID as 'Terminal ID',SSRATE1 as 'Q1', SSRATE2 as 'Q2', SSRATE3 as 'Q3', SSRATE4 as 'Q4', SSRATE5 as 'Q5', SSRATE6 as 'Q6', SSRATE7 as 'Q7',case when VST_TAG = 1 then 'YES' when VST_TAG = 0 then 'NO' ELSE 'UNDECIDED' end as [Re-visit],REASN_TAG as 'reason',ENCODE_DT as 'Date Submitted' from SSTRANSINFOTERMFB INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSINFOTERMFB.BRANCH_CD where cast(ENCODE_DT as DATE) between '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and POST_CD LIKE '%" & txtSearch.Text & "%'ORDER BY ENCODE_DT", dgvFB)
                        Case "Terminal ID"
                            db.FillDataGridView("Select NAME as  'Name', EMAIL as 'Email',case when ADDRESS_TYP = 0 then 'Home Address' else 'Business Address' end as [Address type],ADD_1 + ',' + ADD_2 as Address, POST_CD as 'Postal Code',SSINFOERMBR.BRANCH_NM as 'Branch',KIOSK_ID as 'Terminal ID',SSRATE1 as 'Q1', SSRATE2 as 'Q2', SSRATE3 as 'Q3', SSRATE4 as 'Q4', SSRATE5 as 'Q5', SSRATE6 as 'Q6', SSRATE7 as 'Q7',case when VST_TAG = 1 then 'YES' when VST_TAG = 0 then 'NO' ELSE 'UNDECIDED' end as [Re-visit],REASN_TAG as 'reason',ENCODE_DT as 'Date Submitted' from SSTRANSINFOTERMFB INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSINFOTERMFB.BRANCH_CD where cast(ENCODE_DT as DATE) between '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and KIOSK_ID LIKE '%" & txtSearch.Text & "%'ORDER BY ENCODE_DT", dgvFB)
                        Case "Branch"
                            db.FillDataGridView("Select NAME as  'Name', EMAIL as 'Email',case when ADDRESS_TYP = 0 then 'Home Address' else 'Business Address' end as [Address type],ADD_1 + ',' + ADD_2 as Address, POST_CD as 'Postal Code',SSINFOERMBR.BRANCH_NM as 'Branch',KIOSK_ID as 'Terminal ID',SSRATE1 as 'Q1', SSRATE2 as 'Q2', SSRATE3 as 'Q3', SSRATE4 as 'Q4', SSRATE5 as 'Q5', SSRATE6 as 'Q6', SSRATE7 as 'Q7',case when VST_TAG = 1 then 'YES' when VST_TAG = 0 then 'NO' ELSE 'UNDECIDED' end as [Re-visit],REASN_TAG as 'reason',ENCODE_DT as 'Date Submitted' from SSTRANSINFOTERMFB INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSINFOTERMFB.BRANCH_CD where cast(ENCODE_DT as DATE) between '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and SSINFOTERMBR.BRANCH_NM LIKE '%" & txtSearch.Text & "%'ORDER BY ENCODE_DT", dgvFB)
                    End Select

                Case 1

                    Select Case cbFields.Text


                        Case "Name"
                            db.FillDataGridView("Select NAME as 'Name', EMAIL as 'Email',case when ADDRESS_TYP = 0 then 'Home Address' else 'Business Address' end as [Address type], ADD_1 +','+ADD_2 as Address,POST_CD as 'Postal Code',SSINFOTERMBR.BRANCH_NM as 'Branch',KIOSK_ID as 'Terminal ID', SSRATE1 as 'Q1', SSRATE2 as 'Q2', SSRATE3 as 'Q3', SSRATE4 as 'Q4', SSRATE5 as 'Q5', SSRATE6 as 'Q6', case when HELP_TAG = 1 then 'YES' else 'NO' end as Helpful,COMMNT_TAG as 'Comment',ENCODE_DT as 'Date Submitted' from SSTRANSINFOTERMFBKIOSK INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSINFOTERMFBKIOSK.BRANCH_CD where cast(ENCODE_DT as DATE) between '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and NAME LIKE '%" & txtSearch.Text & "%'ORDER BY ENCODE_DT", dgvFB)
                        Case "Email"
                            db.FillDataGridView("Select NAME as 'Name', EMAIL as 'Email',case when ADDRESS_TYP = 0 then 'Home Address' else 'Business Address' end as [Address type], ADD_1 +','+ADD_2 as Address,POST_CD as 'Postal Code',SSINFOTERMBR.BRANCH_NM as 'Branch',KIOSK_ID as 'Terminal ID', SSRATE1 as 'Q1', SSRATE2 as 'Q2', SSRATE3 as 'Q3', SSRATE4 as 'Q4', SSRATE5 as 'Q5', SSRATE6 as 'Q6', case when HELP_TAG = 1 then 'YES' else 'NO' end as Helpful,COMMNT_TAG as 'Comment',ENCODE_DT as 'Date Submitted' from SSTRANSINFOTERMFBKIOSK INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSINFOTERMFBKIOSK.BRANCH_CD where cast(ENCODE_DT as DATE) between '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and EMAIL LIKE '%" & txtSearch.Text & "%'ORDER BY ENCODE_DT", dgvFB)
                        Case "Postal Code"
                            db.FillDataGridView("Select NAME as 'Name', EMAIL as 'Email',case when ADDRESS_TYP = 0 then 'Home Address' else 'Business Address' end as [Address type], ADD_1 +','+ADD_2 as Address,POST_CD as 'Postal Code',SSINFOTERMBR.BRANCH_NM as 'Branch',KIOSK_ID as 'Terminal ID', SSRATE1 as 'Q1', SSRATE2 as 'Q2', SSRATE3 as 'Q3', SSRATE4 as 'Q4', SSRATE5 as 'Q5', SSRATE6 as 'Q6', case when HELP_TAG = 1 then 'YES' else 'NO' end as Helpful,COMMNT_TAG as 'Comment',ENCODE_DT as 'Date Submitted' from SSTRANSINFOTERMFBKIOSK INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSINFOTERMFBKIOSK.BRANCH_CD where cast(ENCODE_DT as DATE) between '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and POST_CD LIKE '%" & txtSearch.Text & "%'ORDER BY ENCODE_DT", dgvFB)
                        Case "Terminal ID"
                            db.FillDataGridView("Select NAME as 'Name', EMAIL as 'Email',case when ADDRESS_TYP = 0 then 'Home Address' else 'Business Address' end as [Address type], ADD_1 +','+ADD_2 as Address,POST_CD as 'Postal Code',SSINFOTERMBR.BRANCH_NM as 'Branch',KIOSK_ID as 'Terminal ID', SSRATE1 as 'Q1', SSRATE2 as 'Q2', SSRATE3 as 'Q3', SSRATE4 as 'Q4', SSRATE5 as 'Q5', SSRATE6 as 'Q6', case when HELP_TAG = 1 then 'YES' else 'NO' end as Helpful,COMMNT_TAG as 'Comment',ENCODE_DT as 'Date Submitted' from SSTRANSINFOTERMFBKIOSK INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSINFOTERMFBKIOSK.BRANCH_CD where cast(ENCODE_DT as DATE) between '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and KIOSK_ID LIKE '%" & txtSearch.Text & "%'ORDER BY ENCODE_DT", dgvFB)
                        Case "Branch"
                            db.FillDataGridView("Select NAME as 'Name', EMAIL as 'Email',case when ADDRESS_TYP = 0 then 'Home Address' else 'Business Address' end as [Address type], ADD_1 +','+ADD_2 as Address,POST_CD as 'Postal Code',SSINFOTERMBR.BRANCH_NM as 'Branch',KIOSK_ID as 'Terminal ID', SSRATE1 as 'Q1', SSRATE2 as 'Q2', SSRATE3 as 'Q3', SSRATE4 as 'Q4', SSRATE5 as 'Q5', SSRATE6 as 'Q6', case when HELP_TAG = 1 then 'YES' else 'NO' end as Helpful,COMMNT_TAG as 'Comment',ENCODE_DT as 'Date Submitted' from SSTRANSINFOTERMFBKIOSK INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSINFOTERMFBKIOSK.BRANCH_CD where cast(ENCODE_DT as DATE) between '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and SSINFOTERMBR.BRANCH_NM LIKE '%" & txtSearch.Text & "%'ORDER BY ENCODE_DT", dgvFB)

                    End Select

            End Select
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub cbFields_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbFields.SelectedIndexChanged
        Dim date1 As Date = dtpFrom.Value
        Dim date2 As Date = dtpTo.Value
        Try
            Select Case FBClick

                Case 0
                    Select Case cbFields.Text

                        Case "Yes"
                            db.FillDataGridView("Select NAME as  'Name', EMAIL as 'Email',case when ADDRESS_TYP = 0 then 'Home Address' else 'Business Address' end as [Address type],ADD_1 + ',' + ADD_2 as Address, POST_CD as 'Postal Code',SSINFOTERMBR.BRANCH_NM as 'Branch',KIOSK_ID as 'Terminal ID',SSRATE1 as 'Q1', SSRATE2 as 'Q2', SSRATE3 as 'Q3', SSRATE4 as 'Q4', SSRATE5 as 'Q5', SSRATE6 as 'Q6', SSRATE7 as 'Q7',case when VST_TAG = 1 then 'YES' when VST_TAG = 0 then 'NO' ELSE 'UNDECIDED' end as [Re-visit],REASN_TAG as 'reason',ENCODE_DT as 'Date Submitted' from SSTRANSINFOTERMFB INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSINFOTERMFB.BRANCH_CD where cast(ENCODE_DT as DATE) between '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and VST_TAG = '1'ORDER BY ENCODE_DT", dgvFB)

                        Case "No"
                            db.FillDataGridView("Select NAME as  'Name', EMAIL as 'Email',case when ADDRESS_TYP = 0 then 'Home Address' else 'Business Address' end as [Address type],ADD_1 + ',' + ADD_2 as Address, POST_CD as 'Postal Code',SSINFOTERMBR.BRANCH_NM as 'Branch',KIOSK_ID as 'Terminal ID',SSRATE1 as 'Q1', SSRATE2 as 'Q2', SSRATE3 as 'Q3', SSRATE4 as 'Q4', SSRATE5 as 'Q5', SSRATE6 as 'Q6', SSRATE7 as 'Q7',case when VST_TAG = 1 then 'YES' when VST_TAG = 0 then 'NO' ELSE 'UNDECIDED' end as [Re-visit],REASN_TAG as 'reason',ENCODE_DT as 'Date Submitted' from SSTRANSINFOTERMFB INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSINFOTERMFB.BRANCH_CD where cast(ENCODE_DT as DATE) between '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and VST_TAG = '0'ORDER BY ENCODE_DT", dgvFB)

                        Case "Undecided"
                            db.FillDataGridView("Select NAME as  'Name', EMAIL as 'Email',case when ADDRESS_TYP = 0 then 'Home Address' else 'Business Address' end as [Address type],ADD_1 + ',' + ADD_2 as Address, POST_CD as 'Postal Code',SSINFOTERMBR.BRANCH_NM as 'Branch',KIOSK_ID as 'Terminal ID',SSRATE1 as 'Q1', SSRATE2 as 'Q2', SSRATE3 as 'Q3', SSRATE4 as 'Q4', SSRATE5 as 'Q5', SSRATE6 as 'Q6', SSRATE7 as 'Q7',case when VST_TAG = 1 then 'YES' when VST_TAG = 0 then 'NO' ELSE 'UNDECIDED' end as [Re-visit],REASN_TAG as 'reason',ENCODE_DT as 'Date Submitted' from SSTRANSINFOTERMFB INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSINFOTERMFB.BRANCH_CD where cast(ENCODE_DT as DATE) between '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and VST_TAG = '2'ORDER BY ENCODE_DT", dgvFB)
                        Case "Show all"
                            db.FillDataGridView("Select NAME as  'Name', EMAIL as 'Email',case when ADDRESS_TYP = 0 then 'Home Address' else 'Business Address' end as [Address type],ADD_1 + ',' + ADD_2 as Address, POST_CD as 'Postal Code',SSINFOTERMBR.BRANCH_NM as 'Branch',KIOSK_ID as 'Terminal ID',SSRATE1 as 'Q1', SSRATE2 as 'Q2', SSRATE3 as 'Q3', SSRATE4 as 'Q4', SSRATE5 as 'Q5', SSRATE6 as 'Q6', SSRATE7 as 'Q7',case when VST_TAG = 1 then 'YES' when VST_TAG = 0 then 'NO' ELSE 'UNDECIDED' end as [Re-visit],REASN_TAG as 'reason',ENCODE_DT as 'Date Submitted' from SSTRANSINFOTERMFB INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSINFOTERMFB.BRANCH_CD where cast(ENCODE_DT as DATE) between '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "'", dgvFB)

                    End Select
                Case 1

                    Select Case cbFields.Text

                        Case "Yes"
                            db.FillDataGridView("Select NAME as 'Name', EMAIL as 'Email',case when ADDRESS_TYP = 0 then 'Home Address' else 'Business Address' end as [Address type], ADD_1 +','+ADD_2 as Address,POST_CD as 'Postal Code',SSINFOTERMBR.BRANCH_NM as 'Branch',KIOSK_ID as 'Terminal ID', SSRATE1 as 'Q1', SSRATE2 as 'Q2', SSRATE3 as 'Q3', SSRATE4 as 'Q4', SSRATE5 as 'Q5', SSRATE6 as 'Q6', case when HELP_TAG = 1 then 'YES' else 'NO' end as Helpful,COMMNT_TAG as 'Comment',ENCODE_DT as 'Date Submitted' from SSTRANSINFOTERMFBKIOSK INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSINFOTERMFBKIOSK.BRANCH_CD where cast(ENCODE_DT as DATE) between '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and HELP_TAG = '1'ORDER BY ENCODE_DT", dgvFB)

                        Case "No"
                            db.FillDataGridView("Select NAME as 'Name', EMAIL as 'Email',case when ADDRESS_TYP = 0 then 'Home Address' else 'Business Address' end as [Address type], ADD_1 +','+ADD_2 as Address,POST_CD as 'Postal Code',SSINFOTERMBR.BRANCH_NM as 'Branch',KIOSK_ID as 'Terminal ID', SSRATE1 as 'Q1', SSRATE2 as 'Q2', SSRATE3 as 'Q3', SSRATE4 as 'Q4', SSRATE5 as 'Q5', SSRATE6 as 'Q6', case when HELP_TAG = 1 then 'YES' else 'NO' end as Helpful,COMMNT_TAG as 'Comment',ENCODE_DT as 'Date Submitted' from SSTRANSINFOTERMFBKIOSK INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSINFOTERMFBKIOSK.BRANCH_CD where cast(ENCODE_DT as DATE) between '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "' and HELP_TAG = '0'ORDER BY ENCODE_DT", dgvFB)

                        Case "Show all"
                            db.FillDataGridView("Select NAME as 'Name', EMAIL as 'Email', case when ADDRESS_TYP = 0 then 'Home Address' else 'Business Address' end as [Address type], ADD_1 +','+ADD_2 as Address,POST_CD as 'Postal Code',SSINFOTERMBR.BRANCH_NM as 'Branch',KIOSK_ID as 'Terminal ID', SSRATE1 as 'Q1', SSRATE2 as 'Q2', SSRATE3 as 'Q3', SSRATE4 as 'Q4', SSRATE5 as 'Q5', SSRATE6 as 'Q6', case when HELP_TAG = 1 then 'YES' else 'NO' end as Helpful,COMMNT_TAG as 'Comment',ENCODE_DT as 'Date Submitted' from SSTRANSINFOTERMFBKIOSK INNER JOIN SSINFOTERMBR ON SSINFOTERMBR.BRANCH_CD = SSTRANSINFOTERMFBKIOSK.BRANCH_CD where cast(ENCODE_DT as DATE) between '" & date1.ToShortDateString & "' and '" & date2.ToShortDateString & "'ORDER BY ENCODE_DT", dgvFB)

                    End Select

            End Select
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub ButtonX1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonX1.Click
        If lblName.Text = "Name :" Then
            MsgBox("Please select a memeber to print.", MsgBoxStyle.Information)
        Else
            operationFileTrans()
            _frmViewReports.ReportPrint = 1
            _frmBlock.Show()
            _frmPrinter.Show()
        End If
    End Sub

    Private Sub dgvFB_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dgvFB.DoubleClick
        labelShow()
        Select Case cbFB.Text

            Case "SSIT FEEDBACK"
                fillDataSSIT()
            Case "WEBSITE FEEDBACK"
                fillDataWebsite()
        End Select
    End Sub

    Private Sub Label7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label7.Click

    End Sub

    Private Sub ButtonX2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonX2.Click
        If lblName.Text = "Name :" Then
            MsgBox("Please select a memeber to print.", MsgBoxStyle.Information)
        Else
            _frmExportReport.int1 = 1
            operationFileTrans()
            ' _frmViewReports.ReportPrint = 1
            _frmBlock.Show()
            _frmExportReport.Show()
        End If
    End Sub
End Class