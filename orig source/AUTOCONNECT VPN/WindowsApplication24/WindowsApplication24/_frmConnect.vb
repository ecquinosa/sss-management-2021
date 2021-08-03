Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System
Imports System.Object
Imports System.Windows
Imports System.Net.NetworkInformation
Imports System.Text
Imports System.Collections.Generic


Public Class _frmConnect
    Dim xs As New MySettings
    Public Const GW_HWNDPREV = 3
    Private Const SW_SHOW = 5
    Private Const SW_RESTORE = 9
    Dim y As Integer

    Private Declare Function GetClassName _
    Lib "user32" Alias "GetClassNameA" _
    (ByVal hwnd As Integer, ByVal lpClassName As StringBuilder, ByVal nMaxCount As Integer) As Integer

    '<System.Runtime.InteropServices.DllImport("user32. dll", EntryPoint:="SetForegroundWindow1", CallingConvention:=Runtime.InteropServices.Calling ,Convention.StdCall,CharSet:=Runtime.InteropServices.CharSet.Unicode, SetLastError:=True)> 
    
   
    Declare Function SetForegroundWindow1 Lib "user32" Alias "SetForegroundWindow1" _
                   (ByVal handle As IntPtr) As Boolean


    Declare Function ShowWindow Lib "user32" Alias "ShowWindow" _
                   (ByVal handle As IntPtr, ByVal nCmd As Int32) As Boolean

    Declare Function IsIconic Lib "user32" _
                    (ByVal handle As IntPtr) As Boolean

    'Declare Function IsZoomed Lib "user32" _
    '            (ByVal hWnd As IntPtr) As Boolean

    'Declare Function EnumChildProc Lib "user32" (ByVal hwnd As Long, ByVal lParam As Long) As Long



    ' NASA TAAS BAGO

    Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
                    (ByVal lpClassName As String, ByVal lpWindowName As String) As IntPtr

    Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
                    (ByVal hWnd As IntPtr, ByVal hWndChildAfterA As IntPtr, ByVal lpszClass As String, ByVal lpszWindow As String) As IntPtr
    Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
                     (ByVal hWnd As IntPtr, ByVal wMsg As Integer, ByVal wParam As IntPtr, ByVal lParam As String) As IntPtr
    Declare Function SetForegroundWindow Lib "user32" _
                  (ByVal hWnd As IntPtr) As Boolean


    Const WM_SETTEXT As Integer = &HC

    Dim VPN_USERNAME As String
    Dim VPN_PASSWORD As String

    'Declare Function EnumChildProc Lib "user32" (ByVal hwnd As Long, ByVal lParam As Long) As Long
    ' program-defined code goes here


    'Private Sub test1(ByVal hwnd As IntPtr)
    '    Dim strStatus As String


    '    ' this will return true or false.

    '    'If IntPtr.Zero.Equals(hwnd) Then
    '    '    Return strStatus = ""
    '    '    Exit Sub
    '    'End If
    '    'If IsIconic(hwnd) Then
    '    '    Return strStatus = True
    '    'Else
    '    '    strStatus = False
    '    'End If

    '    'Return strStatus

    '    If strStatus = "MIN" Then
    '        'mimized
    '        ShowWindow(hwnd, SW_RESTORE)
    '        SetForegroundWindow(hwnd)
    '    Else
    '        'maximzed or restored
    '        SetForegroundWindow(hwnd)
    '    End If

    'End Sub

    Public Class ApiWindow
        Public MainWindowTitle As String = ""
        Public ClassName As String = "Internet_Explorer_Server"
        Public hWnd As Int32
    End Class


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        connect()
    End Sub

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        xs.xml_Filename = Application.StartupPath & "\Mysettings.xml" '"Mysettings.xml"
        xs.value_path = "Configuration/Settings"

        Timer1.Start()


    End Sub

    Private Sub loadFortiClient()
        System.Diagnostics.Process.Start("C:\Program Files\Fortinet\FortiClient\FortiClient.exe")
        'System.Diagnostics.Process.Start("C:\Windows\notepad.exe")

    End Sub
    Private Sub connect()

        GC.Collect()
        Dim destination As IntPtr = FindWindow(Nothing, "FortiClient")



        If destination <> 0 Then

            SetForegroundWindow(destination)


            VPN_USERNAME = xs.readSettings(xs.xml_Filename, xs.value_path, "VPN_USERNAME")
            VPN_PASSWORD = xs.readSettings(xs.xml_Filename, xs.value_path, "VPN_PASSWORD")

            'MsgBox(VPN_PASSWORD.Trim)

            'SendKeys.Send(" " & CStr(My.Settings.VPN_PASSWORD))
            SendKeys.Send(" " & CStr(VPN_PASSWORD))
            SendKeys.SendWait("{ENTER}")
            'SendKeys.SendWait("{TAB}")

        Else
            '  MsgBox("The forticlient is not Open")
            Button1.PerformClick()
        End If

    End Sub

    Private Sub connect1()

        GC.Collect()
        Dim destination As IntPtr = FindWindow(Nothing, "FortiClient")
        Dim window As New ApiWindow()
        Dim classBuilder As New StringBuilder(64)


        Dim hasErrorWindow_1 As Boolean
        GetClassName(destination, classBuilder, 64)
        window.ClassName = classBuilder.ToString()

        'MsgBox(window.ClassName)



        'Dim parent_Destination As IntPtr = FindWindow(, "")
        Dim destination1 As IntPtr = FindWindowEx(destination, IntPtr.Zero, "Static", Nothing)
        Dim destination2 As IntPtr = FindWindowEx(destination1, IntPtr.Zero, "Shell Embedding", Nothing)
        Dim destination3 As IntPtr = FindWindowEx(destination2, IntPtr.Zero, "Shell DocObject View", Nothing)
        Dim destination4 As IntPtr = FindWindowEx(destination3, IntPtr.Zero, "Internet Explorer_Server", Nothing)


        If destination4 <> 0 Then

            SetForegroundWindow(destination)


            VPN_USERNAME = xs.readSettings(xs.xml_Filename, xs.value_path, "VPN_USERNAME")
            VPN_PASSWORD = xs.readSettings(xs.xml_Filename, xs.value_path, "VPN_PASSWORD")

            'MsgBox(VPN_PASSWORD.Trim)
            SendKeys.SendWait("{TAB}")
            SendKeys.SendWait("{TAB}")
            'SendKeys.Send(" " & CStr(My.Settings.VPN_PASSWORD))
            SendKeys.Send(" " & CStr(VPN_PASSWORD))
            SendKeys.SendWait("{ENTER}")
            'SendKeys.SendWait("{TAB}")

        Else
            '  MsgBox("The forticlient is not Open")
            Button1.PerformClick()
        End If

    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick  ' INTERVAL OPENING OF FORTICLIENT
        Timer1.Interval = Timer1.Interval + 100
        If Timer1.Interval = 1500 Then
            loadFortiClient()
            Timer2.Start()
            ' Button1.PerformClick()
            Timer1.Stop()
        End If
    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick '  Connect to forticlient
        Timer2.Interval = Timer2.Interval + 100
        If Timer2.Interval = 1000 Then
            'Button1.PerformClick()
            connect()
            'Timer2.Interval = 0
            Timer2.Stop()
            Timer3.Start()
        End If
    End Sub

    Private Sub Timer3_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer3.Tick ' open updater to open SSIT WHEN connected to the SSS servers.
      
        Dim ifConnected As Boolean = checkNetworkConnection()

        If ifConnected = True Then
            openUpdater()
            'Timer3.Stop()
            'Timer4.Start()
        ElseIf ifConnected = False Then
            ' do nothing 
            y = Timer4.Interval
            Timer4.Start()
            Timer3.Stop()

        End If
      
    End Sub


    Private Sub openUpdater()
        Try

            Dim listbox1 As New ListBox
            Dim myPath As String = Application.StartupPath & "\files"


            listbox1.Items.Clear()
            For Each p As Process In Process.GetProcesses
                If p.MainWindowTitle = String.Empty = False Then
                    listbox1.Items.Add(p.MainWindowTitle)
                End If
            Next

            'Dim Mystring1 As String = "SSIT SERVER"
            Dim MyFilePath = "C:\Users\Public\Desktop\SSIT_UPDATER.lnk"
            Shell("RUNDLL32.EXE URL.DLL,FileProtocolHandler " & MyFilePath, vbMaximizedFocus)



            ' End If


        Catch ex As Exception

            MsgBox(ex.Message, vbOKOnly, "CLICK_UPDATER")


        End Try

    End Sub


    Private Function checkNetworkConnection() As Boolean

        Try
            Dim adapters As NetworkInterface() = NetworkInterface.GetAllNetworkInterfaces()
            Dim adapter As NetworkInterface
            For Each adapter In adapters

                If adapter.Description.Contains("Fortinet") Then
                    Dim adapterProperties As IPInterfaceProperties = adapter.GetIPProperties()
                    Dim dnsServers As IPAddressCollection = adapterProperties.DnsAddresses

                    If adapter.OperationalStatus = OperationalStatus.Up Then
                        ' MsgBox("Connected!")
                        checkNetworkConnection = True
                    ElseIf adapter.OperationalStatus = OperationalStatus.Down Then
                        ' do nothing
                        checkNetworkConnection = False

                    End If
                End If

            Next

            Return checkNetworkConnection

        Catch ex As Exception
            MsgBox(ex.Message, "Check Network Connection")
        End Try


    End Function

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click

    End Sub

    Private Sub Timer4_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer4.Tick
        Try



            If y = 5000 Then
                'MsgBox("3 Minutes na.")
                Dim ifConnected As Boolean = checkNetworkConnection()

                If ifConnected = True Then
                    ' do nothing
                    Dim destination As IntPtr = FindWindow(Nothing, "FortiClient")
                    Dim haserror1 As Boolean = getErrWindow_1()
                    Dim haserror2 As Boolean = getErrWindow_2()
                    If haserror1 = True Then
                        MsgBox("Error Window 1")

                    ElseIf haserror2 = True Then
                        MsgBox("Error Window 2")
                    Else
                        'ShowWindow(destination, SW_RESTORE)
                        'connect()
                    End If
                ElseIf ifConnected = False Then
                    Dim destination As IntPtr = FindWindow(Nothing, "FortiClient")
                    Dim haserror1 As Boolean = getErrWindow_1()
                    Dim haserror2 As Boolean = getErrWindow_2()

                    If haserror1 = True Then

                        connect1()
                        Timer5.Start()

                    ElseIf haserror2 = True Then
                        'MsgBox("Error 2")
                        KillProgram("FortiTray")
                        connect1()
                        Timer5.Start()
                    Else
                        ShowWindow(destination, SW_RESTORE)
                        'connect()
                    End If
                    ' getErrWindow_1()
                    ' getErrWindow_2()
                    ' ShowWindow(destination, SW_RESTORE)
                    ' SetForegroundWindow1(destination)
                    '     connect()

                End If
                y = 0

            Else
                y = y + 1000

            End If



        Catch ex As Exception
            MsgBox(ex.Message, "Timer 4 ")
        End Try
       
    End Sub

    Private Function getErrWindow_1()
        Dim destination As IntPtr = FindWindow(Nothing, "FortiClient")
        Dim window As New ApiWindow()
        Dim classBuilder As New StringBuilder(64)


        Dim hasErrorWindow_1 As Boolean
        GetClassName(destination, classBuilder, 64)
        window.ClassName = classBuilder.ToString()

        'MsgBox(window.ClassName)



        Dim parent_Destination As IntPtr = FindWindow(window.ClassName, "")
        Dim destination1 As IntPtr = FindWindowEx(parent_Destination, IntPtr.Zero, "Static", Nothing)
        Dim destination2 As IntPtr = FindWindowEx(destination1, IntPtr.Zero, "Shell Embedding", Nothing)
        Dim destination3 As IntPtr = FindWindowEx(destination2, IntPtr.Zero, "Shell DocObject View", Nothing)
        Dim destination4 As IntPtr = FindWindowEx(destination3, IntPtr.Zero, "Internet Explorer_Server", Nothing)

        If parent_Destination <> 0 And destination1 <> 0 And destination2 <> 0 And destination3 <> 0 And destination4 <> 0 Then
            '  MsgBox("Parent: " & parent_Destination.ToString & "Child 1: " & destination1.ToString & _
            ' vbNewLine & "Child 2: " & destination2.ToString & "Child 3: " & destination3.ToString & "Child 4: " & destination4.ToString)
            ' MsgBox("Has error 1")
            hasErrorWindow_1 = True
            SetForegroundWindow(destination4)
            SendKeys.SendWait("{ENTER}")
        Else
            hasErrorWindow_1 = False
        End If



        'Dim destination4 As IntPtr = FindWindowEx(destination3, IntPtr.Zero, "Internet Explorer_Server", Nothing)



        Return hasErrorWindow_1


    End Function

    Private Function getErrWindow_2()


        Dim destination As IntPtr = FindWindow(Nothing, "FortiClient")
        Dim parent_Destination1 As IntPtr = FindWindow(Nothing, "VPN Connection Failure")
        Dim destination12 As IntPtr = FindWindowEx(parent_Destination1, IntPtr.Zero, "Button", Nothing)
        Dim hasErrWindow2 As Boolean
        hasErrWindow2 = False


        '   MsgBox(parent_Destination1.ToString & vbNewLine.ToString & destination12.ToString)
        If parent_Destination1 <> 0 And destination12 <> 0 Then
            'KillProgram("FortiTray")
            hasErrWindow2 = True
            SendKeys.SendWait("{ENTER}")

            '   MsgBox(parent_Destination1)
            ' MsgBox("Parent: " & parent_Destination1.ToString & "Child 1: " & destination12.ToString)
            ' SetForegroundWindow(destination12)
            'SendKeys.SendWait("{ENTER}")
        End If


        Return hasErrWindow2
    End Function


    Public Sub KillProgram(ByVal Program As String, Optional ByVal blnLog As Boolean = True)
        Try
            Dim programs() As Process = Process.GetProcessesByName(Program.Replace(".exe", "").Replace(".EXE", ""))
            For Each _program As Process In programs
                _program.Kill()
            Next

            'If blnLog Then SaveToLog(TimeStamp() & Program & " is killed by " & App)
        Catch ex As Exception
            'SaveToErrorLog(TimeStamp() & "KillProgram(): Runtime catched error " & ex.Message)
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        getErrWindow_1()
    End Sub

    Private Sub Timer5_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer5.Tick
        'connect()
        'Timer5.Stop()
    End Sub
End Class
