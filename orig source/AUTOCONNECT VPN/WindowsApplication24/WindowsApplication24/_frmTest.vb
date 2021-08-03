Imports System.Windows
Imports System.Management
Imports System.Runtime.InteropServices
Imports System.Net.NetworkInformation
Imports System.Reflection
Imports System
Imports System.Object
Imports System.Text
Imports System.Collections.Generic

Public Class _frmTest
    Private Const SW_RESTORE = 9
    Private Declare Auto Function FindWindowEx Lib "user32.dll" ( _
  ByVal hwndParent As IntPtr, _
  ByVal hwndChildAfter As IntPtr, _
  ByVal lpszClass As String, _
  ByVal lpszWindow As String _
  ) As IntPtr

    Declare Function SetForegroundWindow Lib "user32" _
                (ByVal hWnd As IntPtr) As Boolean



    Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
                    (ByVal lpClassName As String, ByVal lpWindowName As String) As IntPtr

    Public Declare Function EnumWindows Lib "User32.dll" (ByVal WNDENUMPROC As EnumWindowDelegate, ByVal lparam As IntPtr) As Boolean
    Delegate Function EnumWindowDelegate(ByVal hWnd As IntPtr, ByVal Lparam As IntPtr) As Boolean

    Declare Function ShowWindow Lib "user32" Alias "ShowWindow" _
                  (ByVal handle As IntPtr, ByVal nCmd As Int32) As Boolean


    Public Function EnumChildWindows _
        (ByVal WindowHandle As IntPtr, ByVal Callback As EnumWindowProcess, _
        ByVal lParam As IntPtr) As Boolean
    End Function

    Private Declare Function GetClassName _
    Lib "user32" Alias "GetClassNameA" _
    (ByVal hwnd As Integer, ByVal lpClassName As StringBuilder, ByVal nMaxCount As Integer) As Integer



    Private Declare Function GetWindow Lib "user32" _
  (ByVal hwnd As Long, _
   ByVal wCmd As Long) As Long
  
    Public Delegate Function EnumWindowProcess(ByVal Handle As IntPtr, ByVal Parameter As IntPtr) As Boolean

    ' Get window text length signature.
    Public Declare Function SendMessage _
     Lib "user32" Alias "SendMessageA" _
     (ByVal hwnd As Int32, ByVal wMsg As Int32, ByVal wParam As Int32, ByVal lParam As Int32) As Int32

    ' Get window text signature.
    Public Declare Function SendMessage _
     Lib "user32" Alias "SendMessageA" _
     (ByVal hwnd As Int32, ByVal wMsg As Int32, ByVal wParam As Int32, ByVal lParam As StringBuilder) As Int32


    Private Const GW_CHILD = 5

    Public Structure ApiWindow
        Public MainWindowTitle As String
        Public ClassName As String
        Public hWnd As Int32
    End Structure

    Public Function GetChildWindows(ByVal ParentHandle As IntPtr) As IntPtr()
        Dim ChildrenList As New List(Of IntPtr)
        Dim ListHandle As GCHandle = GCHandle.Alloc(ChildrenList)
        Try
            EnumChildWindows(ParentHandle, AddressOf EnumWindow, GCHandle.ToIntPtr(ListHandle))
        Finally
            If ListHandle.IsAllocated Then ListHandle.Free()
        End Try
        Return ChildrenList.ToArray
    End Function

    Public Function EnumWindow(ByVal Handle As IntPtr, ByVal Parameter As IntPtr) As Boolean
        Dim ChildrenList As List(Of IntPtr) = GCHandle.FromIntPtr(Parameter).Target
        If ChildrenList Is Nothing Then Throw New Exception("GCHandle Target could not be cast as List(Of IntPtr)")
        ChildrenList.Add(Handle)
        Return True
    End Function

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        ' Dim parent_de1 As IntPtr = FindWindow(Nothing, "FortiClient")
        'Dim destination As IntPtr = FindWindow(Nothing, "FortiClient")

        'MsgBox(destination)

        'Dim r As Long

        'Dim hwndStart As Long = 0
        'Dim sWindowtext As String = "FortiClient"
        'Dim hwnd1 As Long = GetWindow(hwndStart, GW_CHILD)
        '' Dim r As String = GetClassName(destination, Nothing, 0)

        'Dim sclassName As String = Space(255)
        'r = GetClassName(hwnd1, sclassName, 255)




        'MsgBox(r)
        ' Dim parent_de1 As IntPtr
        'Dim parent_Destination As IntPtr = FindWindow("ATL:014431C0", "")
        'Dim destination1 As IntPtr = FindWindowEx(parent_Destination, IntPtr.Zero, "Static", Nothing)
        'Dim destination2 As IntPtr = FindWindowEx(destination1, IntPtr.Zero, "Shell Embedding", Nothing)
        'Dim destination3 As IntPtr = FindWindowEx(destination2, IntPtr.Zero, "Shell DocObject View", Nothing)
        'Dim destination4 As IntPtr = FindWindowEx(destination3, IntPtr.Zero, "Internet Explorer_Server", Nothing)
        ''Dim destination4 As IntPtr = FindWindowEx(destination3, IntPtr.Zero, "Internet Explorer_Server", Nothing)

        'MsgBox("Parent: " & parent_Destination.ToString & "Child 1: " & destination1.ToString & _
        '       vbNewLine & "Child 2: " & destination2.ToString & "Child 3: " & destination3.ToString & "Child 4: " & destination4.ToString)
        'SetForegroundWindow(destination4)
        'SendKeys.SendWait("{ENTER}")


        ''System.Diagnostics.Process.Start("C:\Users\ncjaluag\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Accessories\Accessibility\On-Screen Keyboard.lnk")
        'System.Diagnostics.Process.Start("C:\Users\User\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Accessories\Accessibility\On-Screen Keyboard.lnk")
        'Dim parent_Destination As IntPtr = FindWindow(Nothing, "On-Screen Keyboard")
        ''MsgBox(parent_Destination.ToString)
        'SetForegroundWindow(parent_Destination)
        'SendKeys.SendWait("{ENTER}")



        ''  Dim parentHandle As IntPtr = FindWindow(Nothing, "FortiClient")
        'Dim hndls() As IntPtr = GetChildWindows(destination)
        'Dim window As ApiWindow

        'For Each hnd In hndls
        '    window = GetWindowIdentification(hnd)
        '    'Add Code Here 
        'Next

        'SendKeys.SendWait("{ENTER
        getErrWindow_2()
        Dim destination As IntPtr = FindWindow(Nothing, "FortiClient")
        'ShowWindow(destination, SW_RESTORE)
        'SetForegroundWindow(destination)
        ' connect1()
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


        If destination <> 0 Then

            SetForegroundWindow(destination)


            ' VPN_USERNAME = xs.readSettings(xs.xml_Filename, xs.value_path, "VPN_USERNAME")
            ' VPN_PASSWORD = xs.readSettings(xs.xml_Filename, xs.value_path, "VPN_PASSWORD")

            'MsgBox(VPN_PASSWORD.Trim)
            ' SendKeys.SendWait("{TAB}")
            'SendKeys.SendWait("{TAB}")
            'SendKeys.Send(" " & CStr(My.Settings.VPN_PASSWORD))
            SendKeys.Send("t3rminal")
            SendKeys.SendWait("{ENTER}")
            'SendKeys.SendWait("{TAB}")

        Else
            '  MsgBox("The forticlient is not Open")
            Button1.PerformClick()
        End If

    End Sub

    Private Function getErrWindow_2()


        Dim destination As IntPtr = FindWindow(Nothing, "FortiClient")
        Dim parent_Destination1 As IntPtr = FindWindow(Nothing, "VPN Connection Failure")
        Dim destination12 As IntPtr = FindWindowEx(parent_Destination1, IntPtr.Zero, "Button", Nothing)
        Dim hasErrWindow2 As Boolean
        hasErrWindow2 = False


        '   MsgBox(parent_Destination1.ToString & vbNewLine.ToString & destination12.ToString)
        If parent_Destination1 <> 0 And destination12 <> 0 Then
            KillProgram("FortiTray")
            hasErrWindow2 = True

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
    Public Function GetWindowIdentification(ByVal hwnd As Integer) As ApiWindow

        Const WM_GETTEXT As Int32 = &HD
        Const WM_GETTEXTLENGTH As Int32 = &HE

        Dim window As New ApiWindow()

        Dim title As New StringBuilder()

        ' Get the size of the string required to hold the window title.
        Dim size As Int32 = SendMessage(hwnd, WM_GETTEXTLENGTH, 0, 0)

        ' If the return is 0, there is no title.
        If size > 0 Then
            title = New StringBuilder(size + 1)

            SendMessage(hwnd, WM_GETTEXT, title.Capacity, title)
        End If

        ' Set the properties for the ApiWindow object.
        window.MainWindowTitle = title.ToString()
        window.hWnd = hwnd

        Return window

    End Function

    Private Sub _frmTest_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick

    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick

    End Sub
End Class