

Public Class GenerateMonReport
    Dim db As New ConnectionString
    Dim FILE_NAME As String '= Application.StartupPath & "\REPORT\DOWNTIME.txt"
    Dim dateToday As String
    Dim itemCount As String

    Private Function createFile()
        dateToday = Format(Now, "yyyy_MM_dd")
        Dim filePAth As String = Application.StartupPath & "\REPORT\TEXT_FORMAT\MONITORING_" & dateToday & ".txt"
        If Not System.IO.File.Exists(filePAth) Then
            System.IO.File.Create(filePAth).Dispose()
        End If
        Return filePAth

    End Function
 
    Public Function getData(ByVal timeStarted As Date)
        'createFile()
        FILE_NAME = createFile()

        Dim Lview, lview1 As New ListView
        Dim dtoday As String
        Dim timeDiff As String
        Dim times As Date

        'Dim FILE_NAME As String = Application.StartupPath & "\REPORT\DOWNTIME.txt"
        Dim Ttoday As String
        Dim stat As String
        Dim ocount As Integer


        Dim ipLen As String = "000000000000000"
        Dim dTimeLen As String = "0000000000000000000000000"
        Dim branchLen As String = "000000000000000000000000000000"

        Dim len2, len3 As Integer
        Dim len1 As Integer
        Dim Space0, space1, space3 As String

        Ttoday = Format(Now, "HH:mm")
        dtoday = Format(Now, "yyyy/MM/dd")

        itemCount = itemCount + 1

        Dim proc As System.Diagnostics.Process

        For Each proc In System.Diagnostics.Process.GetProcessesByName("NOTEPAD")
            proc.Kill()
        Next


        db.FillListView(db.ExecuteSQLQuery("SELECT online_Time,offline_time, ipaddress, branch from tbl_kiosk_connection  where online_date = '" & dtoday & "'"), Lview)
        Using SW As New IO.StreamWriter(FILE_NAME, True)
            SW.WriteLine(itemCount & "." & "SSS Infokiosks - " & dtoday & "-" & Ttoday)
            For j = 1 To Lview.Items.Count
                times = Lview.Items(j - 1).Text
                times = times.ToString("t")
                timeDiff = DateDiff(DateInterval.Minute, times, timeStarted)

                If (timeDiff <= 15 And timeDiff >= 0) Then

                    If Lview.Items(j - 1).Text <> "" Then
                        stat = "Online               "
                        ocount = ocount + 1
                    ElseIf Lview.Items(j - 1).SubItems(1).Text <> "" Then
                        stat = "****** Offline ******"
                    End If
                    Dim iplen22 As String
                    iplen22 = Len(ipLen)
                    If Len(Lview.Items(j - 1).SubItems(2).Text) <= Len(ipLen) Then
                        len1 = Len(ipLen) - Len(Lview.Items(j - 1).SubItems(2).Text)
                        Space0 = ""
                        For k = 1 To len1
                            Space0 = Space0 & " "
                        Next
                    End If

                    If Len(Lview.Items(j - 1).Text) <= Len(dTimeLen) Then
                        len2 = Len(dTimeLen) - Len(Lview.Items(j - 1).Text)
                        space1 = ""
                        For o = 1 To len2
                            space1 = space1 & " "
                        Next
                    End If


                    If Len(Lview.Items(j - 1).SubItems(3).Text) <= Len(branchLen) Then
                        len3 = Len(branchLen) - Len(Lview.Items(j - 1).SubItems(3).Text)
                        space3 = ""
                        For u = 1 To len3
                            space3 = space3 & " "
                        Next
                    End If


                    SW.WriteLine((Lview.Items(j - 1).SubItems(2).Text & Space0 & " , " & stat _
                             & " , " & Lview.Items(j - 1).Text & space1 & " , " & Lview.Items(j - 1).SubItems(3).Text) & space3)

                End If
            Next
            SW.WriteLine("")
            SW.WriteLine("Scan End Time: " & dtoday & "-" & Ttoday & " Workstations Online = " & ocount & " Of " & Lview.Items.Count)
            SW.WriteLine("")

        End Using

        ConvertToExcel()

    End Function
    Public Sub ConvertToExcel()

        Dim proc As System.Diagnostics.Process

        For Each proc In System.Diagnostics.Process.GetProcessesByName("EXCEL")
            proc.Kill()
        Next
        For Each proc In System.Diagnostics.Process.GetProcessesByName("TXT")
            proc.Kill()
        Next



        dateToday = Format(Now, "yyyy_MM_dd")
        Dim xlApp As Microsoft.Office.Interop.Excel.Application
        Dim xlWorkBook As Microsoft.Office.Interop.Excel.Workbook
        Dim xlWorkSheet As Microsoft.Office.Interop.Excel.Worksheet
        Dim FILE_NAME1 As String = Application.StartupPath & "\REPORT\TEXT_FORMAT\MONITORING_" & dateToday & ".txt"
        Dim cnt As Integer
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim rng As Microsoft.Office.Interop.Excel.Range
        Dim rchtxtbox As New RichTextBox
        xlApp = New Microsoft.Office.Interop.Excel.Application
        xlWorkBook = xlApp.Workbooks.Add(misValue)
        xlWorkSheet = xlWorkBook.Sheets("sheet1")




        rchtxtbox.AppendText("IP ADDRESS" & " , " _
                                         & "STATUS" & " , " _
                                        & "ONLINE TIME" & " , " _
                                        & "BRANCH" & " , " _
                                        & vbNewLine)

        rchtxtbox.LoadFile(FILE_NAME1, RichTextBoxStreamType.PlainText)

        For Each s As String In rchtxtbox.Lines
            If s.Contains("SSS Infokiosks") Then
                Dim ttle As String
                ttle = s.Substring(0, 35)
                cnt = cnt + 1
                xlWorkSheet.Cells(cnt, 1).Value = ttle
                cnt = cnt + 1
            ElseIf s.Contains("Scan End Time") Then
                cnt = cnt + 1
                xlWorkSheet.Cells(cnt, 1).Value = s
                cnt = cnt + 1
            ElseIf s.Trim <> "" Then
                Dim ipadd As String = s.Substring(0, 16)
                Dim type As String = s.Substring(18, 22)
                Dim tme As String = s.Substring(50, 17)
                Dim brnch As String = s.Substring(69, 30)

                cnt = cnt + 1
                xlWorkSheet.Cells(cnt, 1).Value = ipadd
                xlWorkSheet.Cells(cnt, 2).Value = type
                xlWorkSheet.Cells(cnt, 3).Value = tme
                xlWorkSheet.Cells(cnt, 4).Value = brnch

            End If
        Next
        rng = xlWorkSheet.Range("A1", "D6000")
        rng.EntireColumn.AutoFit()
        Dim excelPath As String = Application.StartupPath & "\REPORT\EXCEL_FORMAT\MONITORING_" & dateToday & ".xlsx"
        If System.IO.File.Exists(excelPath) Then

            My.Computer.FileSystem.DeleteFile(excelPath)
        End If

        xlWorkSheet.SaveAs(Application.StartupPath & "\REPORT\EXCEL_FORMAT\MONITORING_" & dateToday & ".xlsx")

        xlWorkBook.Close()
        xlApp.Quit()

        releaseObject(xlApp)
        releaseObject(xlWorkBook)
        releaseObject(xlWorkSheet)

        xlApp = Nothing
        xlWorkBook = Nothing
        xlWorkSheet = Nothing
        convertToPDf()
    End Sub


    Public Sub convertToPDf()
        

        dateToday = Format(Now, "yyyy_MM_dd")
        Dim xlApp As Microsoft.Office.Interop.Excel.Application
        Dim xlWorkBook As Microsoft.Office.Interop.Excel.Workbook
        Dim xlWorkSheet As Microsoft.Office.Interop.Excel.Worksheet
        Dim proc As System.Diagnostics.Process


        For Each proc In System.Diagnostics.Process.GetProcessesByName("AcroRd32")
            proc.Kill()
        Next


        xlApp = New Microsoft.Office.Interop.Excel.Application
        xlWorkBook = xlApp.Workbooks.Add(Application.StartupPath & "\REPORT\EXCEL_FORMAT\MONITORING_" & dateToday & ".xlsx")

        xlWorkSheet = xlWorkBook.Sheets("sheet1")
        xlWorkSheet.PageSetup.RightMargin = 0
        xlWorkSheet.PageSetup.LeftMargin = 10

        xlWorkSheet.SaveAs(Application.StartupPath & "\REPORT\PDF_FORMAT\MONITORING_" & dateToday & ".xlsx")
        xlWorkSheet.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, Application.StartupPath & "\REPORT\PDF_FORMAT\MONITORING_" & dateToday)


        xlWorkBook.Close()
        xlApp.Quit()
        releaseObject(xlApp)
        releaseObject(xlWorkBook)
        releaseObject(xlWorkSheet)


        xlApp = Nothing
        xlWorkBook = Nothing
        xlWorkSheet = Nothing

        Dim excelPath As String = Application.StartupPath & "\REPORT\PDF_FORMAT\MONITORING_" & dateToday & ".xlsx"
        If System.IO.File.Exists(excelPath) Then
            My.Computer.FileSystem.DeleteFile(excelPath)
        End If
    End Sub
    Private Sub releaseObject(ByVal obj As Object)
        Try

            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing

        Catch ex As Exception
            obj = Nothing
        Finally

            GC.Collect()

        End Try
    End Sub

End Class
