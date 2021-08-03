Imports System.IO
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System
Imports System.Object
Imports System.Windows

Public Class Class1
    Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer

    Private Declare Function SendMessageByString Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As String) As Integer

    Private Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hwnd As Integer) As Integer

    Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hwnd As Integer, ByVal lpString As String, ByVal cch As Integer) As Integer

    Public Const WM_SETTEXT As Integer = &HC

    Public Const WM_GETTEXT As Integer = &HD

    Public Const WM_GETTEXTLENGTH As Integer = &HE






    Private Sub GetText(ByVal WindowHandle As Integer)

        Dim TextLen As Integer

        TextLen = SendMessage(WindowHandle, WM_GETTEXTLENGTH, 0, 0) + 1

        Dim Buf As String = New String(" "c, TextLen)

        MsgBox(TextLen)

        MsgBox(Buf)

    End Sub



    Private Sub SetText(ByVal WindowHandle As Integer, ByVal str As String)

        Dim myint As Integer = str.Length

        SendMessageByString(WindowHandle, WM_SETTEXT, myint, str)

    End Sub


    Public Sub TESTFUNCTION()

        'For Each wn In Systemwindow.AllToplevelWindows

        '    If wn.Visible = True Then

        '        If InStr(wn.Title, "Ares Client") > 0 Then

        '        ElseIf InStr(wn.Title, "Item") > 0 Then

        '            wn.TopMost = True

        '            For Each it In wn.AllChildWindows

        '                If it.ClassName = "TPanel" Then

        '                    Dim i As Integer = 0

        '                    For Each itm In it.AllDescendantWindows

        '                        If itm.ClassName = "TcxCustomInnerTextEdit" Then

        '                            i = i + 1

        '                            SetText(itm.HWnd, i.ToString)

        '                        End If

        '                    Next

        '                End If

        '            Next

        '        End If

        '    End If

        'Next

    End Sub
End Class
