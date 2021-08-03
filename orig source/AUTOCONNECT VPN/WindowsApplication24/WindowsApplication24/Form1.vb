Imports System.Windows
Imports System.Management
Imports System.Runtime.InteropServices
Imports System.Net.NetworkInformation
Imports System.Reflection
Imports System
Imports System.Object
Imports System.Text
Imports System.Collections.Generic

Public Class Form1
    Private Declare Function GetClassName _
     Lib "user32" Alias "GetClassNameA" _
     (ByVal hwnd As Integer, ByVal lpClassName As StringBuilder, ByVal nMaxCount As Integer) As Integer
    Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
                    (ByVal lpClassName As String, ByVal lpWindowName As String) As IntPtr

    Private Declare Auto Function FindWindowEx Lib "user32.dll" ( _
ByVal hwndParent As IntPtr, _
ByVal hwndChildAfter As IntPtr, _
ByVal lpszClass As String, _
ByVal lpszWindow As String _
) As IntPtr


    Declare Function SetForegroundWindow Lib "user32" _
                (ByVal hWnd As IntPtr) As Boolean

    Public Class ApiWindow
        Public MainWindowTitle As String = ""
        Public ClassName As String = "Internet_Explorer_Server"
        Public hWnd As Int32
    End Class

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim destination As IntPtr = FindWindow(Nothing, "FortiClient")
        'Dim window As New ApiWindow()
        'Dim classBuilder As New StringBuilder(64)

        'GetClassName(destination, classBuilder, 64)
        'window.ClassName = classBuilder.ToString()

        'MsgBox(window.ClassName)

        'Dim parent_Destination As IntPtr = FindWindow(window.ClassName, "")
        'Dim destination1 As IntPtr = FindWindowEx(parent_Destination, IntPtr.Zero, "Static", Nothing)
        'Dim destination2 As IntPtr = FindWindowEx(destination1, IntPtr.Zero, "Shell Embedding", Nothing)
        'Dim destination3 As IntPtr = FindWindowEx(destination2, IntPtr.Zero, "Shell DocObject View", Nothing)
        'Dim destination4 As IntPtr = FindWindowEx(destination3, IntPtr.Zero, "Internet Explorer_Server", Nothing)
        ''Dim destination4 As IntPtr = FindWindowEx(destination3, IntPtr.Zero, "Internet Explorer_Server", Nothing)

        'MsgBox("Parent: " & parent_Destination.ToString & "Child 1: " & destination1.ToString & _
        '       vbNewLine & "Child 2: " & destination2.ToString & "Child 3: " & destination3.ToString & "Child 4: " & destination4.ToString)
        'SetForegroundWindow(destination4)
        'SendKeys.SendWait("{ENTER}")


        Dim parent_Destination1 As IntPtr = FindWindow(Nothing, "VPN Connection Failure")
        Dim destination12 As IntPtr = FindWindowEx(parent_Destination1, IntPtr.Zero, "Button", Nothing)


        KillProgram("FortiTray")
        '   MsgBox(parent_Destination1)

        MsgBox("Parent: " & parent_Destination1.ToString & "Child 1: " & destination12.ToString)
        SetForegroundWindow(destination12)
        SendKeys.SendWait("{ENTER}")

    End Sub
    Private Sub KillProgram(ByVal Program As String, Optional ByVal blnLog As Boolean = True)
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
End Class