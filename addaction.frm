VERSION 5.00
Begin VB.Form AddAction 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Action"
   ClientHeight    =   2340
   ClientLeft      =   3900
   ClientTop       =   2355
   ClientWidth     =   3705
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   3705
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   720
      TabIndex        =   6
      Top             =   960
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   720
      TabIndex        =   3
      Top             =   480
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   1920
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   0
      Width           =   3495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Interval"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Path"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   615
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   3600
      Y1              =   1800
      Y2              =   1800
   End
End
Attribute VB_Name = "AddAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
idler.AddSetup
End Sub

Private Sub Command1_Click()
If Len(Text1) > 4 Then
  If Combo1.ListIndex = 0 Then
    If Val(Text2) > 9 Then
      idler.List2.AddItem Text1
      idler.List2.ItemData(idler.List2.ListCount - 1) = Val(Text2)
    End If
  ElseIf Combo1.ListIndex = 1 Then
    idler.List1.AddItem Text1
  End If
End If
Me.Hide
End Sub

Private Sub Command2_Click()
Me.Hide
End Sub

Private Sub Text1_DblClick()
If Combo1.ListIndex = 0 Then
  Temp = OpenDialog(Me, "Executables (*.exe)|*.exe", "Select Application", App.Path)
  If Len(Temp) > 0 Then Text1 = Temp
End If
End Sub
