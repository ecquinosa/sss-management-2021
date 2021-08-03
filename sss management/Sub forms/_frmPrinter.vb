Option Explicit On
Imports System.Management
Imports System.Drawing.Printing
Imports CrystalDecisions.CrystalReports.Engine
Public Class _frmPrinter
    Private Sub PrinterList()
        cbPrinter.Items.Clear()
        ' USING WMI. (WINDOWS MANAGEMENT INSTRUMENTATION)

        Dim objMS As System.Management.ManagementScope = _
            New System.Management.ManagementScope(ManagementPath.DefaultPath)
        objMS.Connect()

        Dim objQuery As SelectQuery = New SelectQuery("SELECT * FROM Win32_Printer")
        Dim objMOS As ManagementObjectSearcher = New ManagementObjectSearcher(objMS, objQuery)
        Dim objMOC As System.Management.ManagementObjectCollection = objMOS.Get()

        For Each Printers As ManagementObject In objMOC
            If CBool(Printers("Local")) Then                        ' LOCAL PRINTERS.
                cbPrinter.Items.Add(Printers("Name"))
            End If
            If CBool(Printers("Network")) Then                      ' ALL NETWORK PRINTERS.
                cbPrinter.Items.Add(Printers("Name"))
            End If
        Next Printers
    End Sub
    Public Property PaperSize As PaperSize
    'Public Function GetPapersizeID(ByVal PrinterName As String, ByVal PaperSizeName As String) As Integer
    '    Dim doctoprint As New System.Drawing.Printing.PrintDocument()
    '    Dim PaperSizeID As Integer = 0
    '    Dim ppname As String = ""
    '    Dim s As String = ""
    '    doctoprint.PrinterSettings.PrinterName = PrinterName '(ex."EpsonSQ-1170ESC/P2")
    '    For i As Integer = 0 To doctoprint.PrinterSettings.PaperSizes.Count - 1
    '        Dim rawKind As Integer
    '        ppname = PaperSizeName
    '        If doctoprint.PrinterSettings.PaperSizes(i).PaperName = ppname Then
    '            rawKind = CInt(doctoprint.PrinterSettings.PaperSizes(i).GetType().GetField("kind", Reflection.BindingFlags.NonPublic).GetValue(doctoprint.PrinterSettings.PaperSizes(i)))
    '            PaperSizeID = rawKind
    '            Exit For
    '        End If
    '    Next
    '    Return PaperSizeID
    'End Function
    Public Sub FillPaperSizes()
        cbPaperSize.Items.Clear()
        Dim doctoprint As New System.Drawing.Printing.PrintDocument()
        For i As Integer = 0 To doctoprint.PrinterSettings.PaperSizes.Count - 1
            cbPaperSize.Items.Add(doctoprint.PrinterSettings.PaperSizes(i).PaperName)
        Next
    End Sub
    Private Sub _frmPrinter_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        PrinterList()
        FillPaperSizes()
        lblFileName.Text = _frmViewReports.GetName
    End Sub

    Private Sub btnsave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnsave.Click
        Try
            If cbPrinter.Text = "" Then
                MsgBox("Please select a printer.", MsgBoxStyle.OkOnly)
            ElseIf cbPaperSize.Text = "" Then
                MsgBox("Please Select a paper size.", MsgBoxStyle.OkOnly)
            Else
                Dim doctoprint As New System.Drawing.Printing.PrintDocument()
                doctoprint.PrinterSettings.PrinterName = cbPrinter.Text
                Dim rawKind As Integer
                For i = 0 To doctoprint.PrinterSettings.PaperSizes.Count - 1
                    If doctoprint.PrinterSettings.PaperSizes(i).PaperName = cbPaperSize.Text Then
                        rawKind = CInt(doctoprint.PrinterSettings.PaperSizes(i).GetType().GetField("kind", Reflection.BindingFlags.Instance Or Reflection.BindingFlags.NonPublic).GetValue(doctoprint.PrinterSettings.PaperSizes(i)))
                        Exit For
                    End If
                Next
                Select Case _frmViewReports.ReportPrint
                    Case 0
                        _frmViewReports.cryRpt.PrintOptions.PaperSize = CType(rawKind, CrystalDecisions.Shared.PaperSize)
                        _frmViewReports.cryRpt.PrintOptions.PrinterName = cbPrinter.Text
                        _frmViewReports.cryRpt.PrintToPrinter(nudCopy.Value, True, 0, 0)
                    Case 1
                        _frmFBK.cryRpt.PrintOptions.PaperSize = CType(rawKind, CrystalDecisions.Shared.PaperSize)
                        _frmFBK.cryRpt.PrintOptions.PrinterName = cbPrinter.Text
                        _frmFBK.cryRpt.PrintToPrinter(nudCopy.Value, True, 0, 0)
                End Select
            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub ComboBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbPrinter.Click

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Dispose()
        _frmBlock.Close()
    End Sub
End Class