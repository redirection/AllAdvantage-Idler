VERSION 5.00
Begin VB.Form idler 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "idler"
   ClientHeight    =   2085
   ClientLeft      =   3315
   ClientTop       =   2745
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   4710
   Begin VB.ListBox List2 
      BackColor       =   &H00FFC0C0&
      Height          =   645
      Left            =   0
      TabIndex        =   7
      Top             =   1440
      Width           =   3255
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2880
      Top             =   240
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Height          =   2200
      Left            =   3250
      TabIndex        =   1
      Top             =   -100
      Width           =   1455
      Begin VB.TextBox appval 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   1840
         Width           =   1215
      End
      Begin VB.TextBox mouseval 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Text            =   "30"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox cycleval 
         BackColor       =   &H00FFC0C0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Text            =   "200"
         Top             =   600
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF8080&
         Caption         =   "cycle pages"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "app interval"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1600
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "mouse interval"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "cycle interval"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00FFC0C0&
      Height          =   1425
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3255
   End
   Begin VB.Menu mnufile 
      Caption         =   "&file"
      Visible         =   0   'False
      Begin VB.Menu mnuopen 
         Caption         =   "&open"
      End
      Begin VB.Menu mnuadd 
         Caption         =   "&add"
      End
      Begin VB.Menu mnudelete 
         Caption         =   "&delete"
      End
      Begin VB.Menu mnuclearall 
         Caption         =   "&clear all"
      End
      Begin VB.Menu mnusavesettings 
         Caption         =   "&save settings"
      End
   End
End
Attribute VB_Name = "idler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'initialize variables as 0
Dim MouseInterval As Integer
Dim CycleInterval As Integer
Dim ApplicationInterval(99) As Integer
Dim WinCaptionBuffer As String
Dim WinHandleBuffer As Long
Dim TempString As String

Public Sub AddSetup()
If AddAction.Combo1.ListIndex = 0 Then
  AddAction.Label1 = "Path"
  AddAction.Text1 = ""
  AddAction.Label2.Visible = True
  AddAction.Text2.Visible = True
ElseIf AddAction.Combo1.ListIndex = 1 Then
  AddAction.Label1 = "Url"
  AddAction.Text1 = "www."
  AddAction.Label2.Visible = False
  AddAction.Text2.Visible = False
End If
End Sub

Public Sub LoadSettings()
On Error Resume Next
Dim sFileName As String
sFileName = App.Path & "\settings.ini" '"c:\windows\desktop\settings.ini"
  Temp = ReadFromINI("Settings", "cycle pages", sFileName)
    If Val(Temp) = 0 Or Val(Temp) = 1 Then Check1.Value = Temp
  Temp = ReadFromINI("Settings", "cycle interval", sFileName)
    If Len(Temp) > 0 Then
      If Temp > 9 Then cycleval.Text = Val(Temp)
    End If
  Temp = ReadFromINI("Settings", "mouse interval", sFileName)
    If Len(Temp) > 0 Then mouseval.Text = Val(Temp)
  Temp = "Sites"
  i = 0
  While Len(Temp) > 0
    DoEvents
    Temp = ReadFromINI("Sites", "num" & i, sFileName)
    If Len(Temp) > 3 Then List1.AddItem Temp
    i = i + 1
  Wend
  Temp = "Applications"
  i = 0
  While Len(Temp) > 0
    DoEvents
    Temp = ReadFromINI("Applications", "Num" & i, sFileName)
    If Len(Temp) > 3 Then List2.AddItem Temp
    Temp = ReadFromINI("Applications", "Int" & i, sFileName)
    If Len(Temp) > 0 Then List2.ItemData(i) = Val(Temp)
    i = i + 1
  Wend
End Sub

Private Sub appval_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyReturn And Len(appval) > 0 Then
  List2.ItemData(List2.ListIndex) = Val(appval)
End If
End Sub

Private Sub Check1_Click()
If Check1.Value = 0 Then
  cycleval.Enabled = False
Else
  cycleval.Enabled = True
End If
End Sub

Private Sub Form_Load()
'App.Path & "\settings.ini"
StayOnTop Me, True
StayOnTop AddAction, True
AddAction.Hide
AddAction.Combo1.AddItem "application"
AddAction.Combo1.AddItem "web page"
WinCaptionBuffer = GetForegroundWindow
If fileexists(App.Path & "\settings.ini") = True Then
  LoadSettings
Else
  List1.AddItem "www.puffcool.com"
  List1.AddItem "www.dogpile.com"
  List1.AddItem "www.hotmail.com"
  List1.AddItem "www.yahoo.com"
  List1.AddItem "www.wired.com"
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload AddAction
Unload Me
End
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
AddAction.Combo1.ListIndex = 1
TempString = "List1"
If Button = 2 Then PopupMenu mnufile
End Sub

Private Sub List2_DblClick()
appval = List2.ItemData(List2.ListIndex)
End Sub

Private Sub List2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
AddAction.Combo1.ListIndex = 0
TempString = "List2"
If Button = 2 Then PopupMenu mnufile
End Sub

Private Sub mnuadd_Click()
AddSetup
AddAction.Show vbModal, Me
End Sub

Private Sub mnuclearall_Click()
User = MsgBox("are you sure?", vbYesNo, "clear all")
If User = 6 Then
  If TempString = "List1" Then
    List1.Clear
  ElseIf TempString = "List2" Then
    List2.Clear
  End If
End If
End Sub

Private Sub mnudelete_Click()
On Error Resume Next
If TempString = "List1" Then
  List1.RemoveItem List1.ListIndex
ElseIf TempString = "List2" Then
  List2.RemoveItem List2.ListIndex
End If
End Sub

Private Sub mnuopen_Click()
On Error Resume Next
Dim ShellApp As String
If TempString = "List1" Then
  ShellApp = List1.List(List1.ListIndex)
ElseIf TempString = "List2" Then
  ShellApp = List2.List(List2.ListIndex)
End If
If Len(ShellApp) > 0 Then ShellOpen ShellApp
End Sub

Private Sub mnusavesettings_Click()
Dim sFileName As String
sFileName = App.Path & "\settings.ini"
If fileexists(sFileName) = True Then Kill sFileName
Call WriteToINI("Settings", "Cycle Pages", Check1.Value, sFileName)
Call WriteToINI("Settings", "Cycle Interval", cycleval.Text, sFileName)
Call WriteToINI("Settings", "Mouse Interval", mouseval.Text, sFileName)
For i = 0 To List1.ListCount - 1
  DoEvents
  Call WriteToINI("Sites", "Num" & i, List1.List(i), sFileName)
Next i
For i = 0 To List2.ListCount - 1
  DoEvents
  Call WriteToINI("Applications", "Num" & i, List2.List(i), sFileName)
  Call WriteToINI("Applications", "Int" & i, List2.ItemData(i), sFileName)
Next i
End Sub

Private Sub Timer1_Timer()
DoEvents
If cycleval < 9 Then cycleval.Text = 10
If GetForegroundWindow <> WinHandleBuffer Then
  Call SetText(WinHandleBuffer, WinCaptionBuffer)
End If
ActiveWinCaption = GetCaption(GetForegroundWindow)
If InStr(ActiveWinCaption, "Microsoft Internet Explorer") = 0 Then
  WinHandleBuffer = GetForegroundWindow
  WinCaptionBuffer = ActiveWinCaption
  Call SetText(WinHandleBuffer, ActiveWinCaption & " - Microsoft Internet Explorer")
End If

'Move Cursor
MouseInterval = MouseInterval + 1
If MouseInterval >= Val(mouseval) Then
  MouseInterval = 0
  ScreenWidth = Screen.Width / Screen.TwipsPerPixelX
  ScreenHeight = Screen.Height / Screen.TwipsPerPixelY
  SetCursorPos Int(ScreenWidth * Rnd), Int(ScreenHeight * Rnd)
End If

'Cycle WebPages
If Check1.Value = 1 Then
  CycleInterval = CycleInterval + 1
  If CycleInterval >= Val(cycleval) Then
    CycleInterval = 0
    ShellOpen List1.List(Int(List1.ListCount * Rnd))
  End If
End If

'Start Application
For i = 0 To List2.ListCount - 1
  DoEvents
  ApplicationInterval(i) = ApplicationInterval(i) + 1
  If ApplicationInterval(i) >= Val(List2.ItemData(i)) Then
    ApplicationInterval(i) = 0
    ShellOpen List2.List(i)
  End If
Next i
'If Check2.Value = 1 Then
'  ApplicationInterval = ApplicationInterval + 1
'  If ApplicationInterval > Val(appval) And InStr(Text2, ".") > 0 Then
'    ApplicationInterval = 0
'    X = FindWindow(vbNullString, Text3)
'    If X = 0 Then ShellOpen Text2
'  End If
'End If
End Sub
