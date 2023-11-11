Imports System.Diagnostics
Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports System.Collections.Generic
Imports System.Threading

Public Class Form1
    Private Const WH_KEYBOARD_LL As Integer = 13
    Private Const WM_KEYDOWN As Integer = &H100
    Private Const WM_ACTIVATE As Integer = &H6

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function SetWindowsHookEx(ByVal idHook As Integer, ByVal lpfn As LowLevelKeyboardProc, ByVal hMod As IntPtr, ByVal dwThreadId As UInteger) As IntPtr
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function UnhookWindowsHookEx(ByVal hhk As IntPtr) As Boolean
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function CallNextHookEx(ByVal hhk As IntPtr, ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
    End Function

    <DllImport("kernel32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function GetModuleHandle(ByVal lpModuleName As String) As IntPtr
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function SetForegroundWindow(ByVal hWnd As IntPtr) As Boolean
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function GetForegroundWindow() As IntPtr
    End Function

    Private Delegate Function LowLevelKeyboardProc(ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr

    Private hookHandle As IntPtr = IntPtr.Zero
    Private lastKeyPressed As Char = " " ' Default value
    Private isFormFocused As Boolean = True ' Flag to track form focus

    ' Dictionary to map regular characters to accented characters
    Private accentMap As New Dictionary(Of Char, String) From {
        {"a"c, "á"},
        {"e"c, "é"},
        {"i"c, "í"},
        {"o"c, "ó"},
        {"u"c, "ú"}}

    Private Sub InstallHook()
        Dim hookProc As New LowLevelKeyboardProc(AddressOf HookCallback)
        Dim hMod As IntPtr = GetModuleHandle(Process.GetCurrentProcess().MainModule.ModuleName)

        hookHandle = SetWindowsHookEx(WH_KEYBOARD_LL, hookProc, hMod, 0)
    End Sub

    Private Sub UninstallHook()
        If hookHandle <> IntPtr.Zero Then
            UnhookWindowsHookEx(hookHandle)
        End If
    End Sub

    Private Function HookCallback(ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
        If nCode >= 0 AndAlso wParam = CType(WM_KEYDOWN, IntPtr) Then
            Dim vkCode As Integer = Marshal.ReadInt32(lParam)

            ' Check if the key is not the backspace key
            If vkCode <> Keys.Back Then
                lastKeyPressed = ChrW(vkCode).ToString().ToLower() ' Store the lowercase version of the last pressed key
                Debug.WriteLine("Key Pressed (Out of Focus): " & lastKeyPressed)
                isFormFocused = True ' Reset the flag when a key is pressed
            End If
        End If

        Return CallNextHookEx(hookHandle, nCode, wParam, lParam)
    End Function

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.TopMost = True
        Me.KeyPreview = True
        MoveToBottomRightCorner()
        InstallHook()
    End Sub

    Private Sub MoveToBottomRightCorner()
        Dim primaryScreen As Screen = Screen.PrimaryScreen
        Dim x As Integer = primaryScreen.WorkingArea.Right - Me.Width
        Dim y As Integer = primaryScreen.WorkingArea.Bottom - Me.Height
        Me.Location = New Point(x, y)
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        UninstallHook()
    End Sub

    Private Sub Form1_Deactivate(sender As Object, e As EventArgs) Handles MyBase.Deactivate
        ' Form lost focus
        isFormFocused = False
    End Sub

    Private Sub Form1_Click(sender As Object, e As EventArgs) Handles MyBase.Click
        ' Set the last active window to the foreground
        Dim lastActiveWindow As IntPtr = GetForegroundWindow()
        SetForegroundWindow(lastActiveWindow)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ' Wait until the form loses focus
        While isFormFocused
            Thread.Sleep(100)
            Application.DoEvents()
        End While

        ' Simulate typing the backspace key and then the accented character based on the last pressed key
        SendKeys.SendWait("{BACKSPACE}")
        If accentMap.ContainsKey(lastKeyPressed) Then
            SendKeys.SendWait(accentMap(lastKeyPressed))
        End If
    End Sub
End Class
