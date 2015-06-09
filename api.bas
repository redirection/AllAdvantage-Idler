Attribute VB_Name = "api"
'#If Win32 Then
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long
'#Else
'Public Declare Function ShellExecute Lib "shell.dll" (ByVal hwnd As Integer, _
    ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Integer) As Integer
'#End If

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMEssageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Public Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long)
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function EnumWindows& Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long)
Public Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal cmd As Long) As Long
Public Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function showwindow Lib "user32" Alias "ShowWindow" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long

Public Declare Function GetTopWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Public Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Function DestroyMenu Lib "user32" (ByVal hMenu%) As Integer

Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const Flags = SWP_NOMOVE Or SWP_NOSIZE

Public Const LB_GETITEMDATA = &H199
Public Const LB_GETCOUNT = &H18B
Public Const LB_ADDSTRING = &H180
Public Const LB_DELETESTRING = &H182
Public Const LB_FINDSTRING = &H18F
Public Const LB_FINDSTRINGEXACT = &H1A2
Public Const LB_GETCURSEL = &H188
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_SELECTSTRING = &H18C
Public Const LB_SETCOUNT = &H1A7
Public Const LB_SETCURSEL = &H186
Public Const LB_SETSEL = &H185
Public Const LB_INSERTSTRING = &H181

Public Const SND_SYNC = &H0
Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_LOOP = &H8
Public Const SND_NOSTOP = &H10
Public Const SND_FLAG = SND_ASYNC Or SND_NODEFAULT

Public Const GW_CHILD = 5
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const GW_MAX = 5
Public Const GW_OWNER = 4
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_HIDE = 0
Public Const SW_RESTORE = 9
Public Const SW_SHOW = 5
Public Const SW_SHOWDEFAULT = 10
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOWNORMAL = 1

Public Const VK_HOME = &H24
Public Const VK_SHIFT = &H10
Public Const VK_CONTROL = &H11
Public Const VK_DELETE = &H2E
Public Const VK_LEFT = &H25
Public Const VK_UP = &H26
Public Const VK_RIGHT = &H27
Public Const VK_DOWN = &H28
Public Const VK_RETURN = &HD
Public Const VK_SPACE = &H20
Public Const VK_TAB = &H9
Public Const VK_MENU = &H12

Global Const WM_CUT = &H300
Global Const WM_COPY = &H301
Global Const WM_PASTE = &H302
Public Const WM_CHAR = &H102
Public Const WM_SETTEXT = &HC
Public Const WM_USER = &H400
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_MOVE = &HF012
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_CLEAR = &H303
Public Const WM_DESTROY = &H2
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_SYSCOMMAND = &H112
Public Const BM_GETCHECK = &HF0
Public Const BM_GETSTATE = &HF2
Public Const BM_SETCHECK = &HF1
Public Const BM_SETSTATE = &HF3

Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Sub ShellOpen(What As String)
On Error Resume Next
Dim ret As Long
  '"mailto:someone@somewhere.com?Subject=" & sSub & ""
  ret = ShellExecute(0&, vbNullString, What, vbNullString, App.Path, SW_SHOWNORMAL)
End Sub

Public Sub PlayMIDI(MIDIFile As String)
On Error Resume Next
Dim SafeFile As String
  SafeFile$ = Dir(MIDIFile$)
  If SafeFile$ <> "" Then
    Call mciSendString("play " & MIDIFile$, 0&, 0, 0)
  End If
End Sub

Public Sub StopMIDI(MIDIFile As String)
On Error Resume Next
Dim SafeFile As String
  SafeFile$ = Dir(MIDIFile$)
  If SafeFile$ <> "" Then
    Call mciSendString("stop " & MIDIFile$, 0&, 0, 0)
  End If
End Sub

Public Sub Playwav(WavFile As String)
On Error Resume Next
Dim SafeFile As String
  SafeFile$ = Dir(WavFile$)
  If SafeFile$ <> "" Then
    Call sndPlaySound(WavFile$, SND_FLAG)
  End If
End Sub

Public Sub ReleaseForm(Form As Form)
On Error Resume Next
  Call ReleaseCapture
  Call SendMessage(Form.hWnd, WM_SYSCOMMAND, WM_MOVE, 0)
End Sub

Public Sub StayOnTop(Form As Form, OnTop As Boolean)
On Error Resume Next
  If OnTop = True Then
    Call SetWindowPos(Form.hWnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, Flags)
  Else
    Call SetWindowPos(Form.hWnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, Flags)
  End If
End Sub

Public Function GetText(WindowHandle As Long) As String
On Error Resume Next
Dim buffer As String, TextLength As Long
  TextLength& = SendMessage(WindowHandle&, WM_GETTEXTLENGTH, 0&, 0&)
  buffer$ = String(TextLength&, 0&)
  Call SendMEssageByString(WindowHandle&, WM_GETTEXT, TextLength& + 1, buffer$)
  GetText$ = buffer$
End Function

Function GetCaption(WindowHandle As Long)
On Error Resume Next
  hwndLength% = GetWindowTextLength(WindowHandle)
  hwndTitle$ = String$(hwndLength%, 0)
  A% = GetWindowText(WindowHandle, hwndTitle$, (hwndLength% + 1))
  GetCaption = hwndTitle$
End Function

Public Function GetClass(WindowHandle As Long)
On Error Resume Next
  buffer$ = String$(250, 0)
  getclas% = GetClassName(WindowHandle, buffer$, 250)
  GetClass = buffer$
End Function

Public Sub SetText(WindowHandle As Long, Text As String)
On Error Resume Next
  Call SendMEssageByString(WindowHandle&, WM_SETTEXT, 0&, Text$)
End Sub

Public Sub CloseWindow(WindowHandle As Long)
On Error Resume Next
  Call PostMessage(WindowHandle&, WM_CLOSE, 0&, 0&)
End Sub

Public Sub Click(WindowHandle As Long)
On Error Resume Next
  Call SendMessage(WindowHandle&, WM_KEYDOWN, VK_SPACE, 0&)
  Call SendMessage(WindowHandle&, WM_KEYUP, VK_SPACE, 0&)
End Sub

Public Sub MouseDown(WindowHandle As Long)
On Error Resume Next
  Call SendMessage(WindowHandle&, WM_LBUTTONDOWN, 0&, 0&)
  Call SendMessage(WindowHandle&, WM_LBUTTONUP, 0&, 0&)
End Sub

Public Sub WriteToINI(Section As String, Key As String, KeyValue As String, fullpath As String)
On Error Resume Next
  Call WritePrivateProfileString(Section$, LCase$(Key$), KeyValue$, fullpath$)
End Sub

Public Function ReadFromINI(Section As String, Key As String, fullpath As String) As String
On Error Resume Next
Dim buffer As String
  buffer$ = String$(750, Chr$(0&))
  ReadFromINI = Left$(buffer$, GetPrivateProfileString(Section$, ByVal LCase$(Key$), "", buffer, Len(buffer), fullpath$))
End Function

Public Function fileexists(sFileName As String) As Boolean
On Error Resume Next
  If Len(sFileName$) = 0 Then
    fileexists = False
    Exit Function
  End If
  If Len(Dir$(sFileName$)) Then
    fileexists = True
  Else
    fileexists = False
  End If
End Function
