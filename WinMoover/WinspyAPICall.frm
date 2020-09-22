VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   30
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   780
   FillStyle       =   0  'Solid
   Icon            =   "WinspyAPICall.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   30
   ScaleWidth      =   780
   Visible         =   0   'False
   Begin VB.PictureBox moving 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   0
      Picture         =   "WinspyAPICall.frx":030A
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   9
      Top             =   480
      Width           =   300
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1680
      Top             =   1560
   End
   Begin VB.PictureBox i 
      AutoSize        =   -1  'True
      Height          =   300
      Index           =   1
      Left            =   240
      Picture         =   "WinspyAPICall.frx":0454
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   8
      Top             =   240
      Width           =   300
   End
   Begin VB.PictureBox i 
      AutoSize        =   -1  'True
      Height          =   300
      Index           =   0
      Left            =   0
      Picture         =   "WinspyAPICall.frx":059E
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   7
      Top             =   240
      Width           =   300
   End
   Begin VB.PictureBox a 
      AutoSize        =   -1  'True
      Height          =   300
      Index           =   1
      Left            =   240
      Picture         =   "WinspyAPICall.frx":06E8
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   6
      Top             =   0
      Width           =   300
   End
   Begin VB.PictureBox a 
      AutoSize        =   -1  'True
      Height          =   300
      Index           =   0
      Left            =   0
      Picture         =   "WinspyAPICall.frx":0832
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   5
      Top             =   0
      Width           =   300
   End
   Begin VB.Label bb1 
      Caption         =   "0"
      Height          =   465
      Left            =   3660
      TabIndex        =   4
      Top             =   5940
      Width           =   990
   End
   Begin VB.Label b4 
      Caption         =   "0"
      Height          =   300
      Left            =   2790
      TabIndex        =   3
      Top             =   6420
      Width           =   735
   End
   Begin VB.Label b3 
      Caption         =   "-1"
      Height          =   420
      Left            =   2130
      TabIndex        =   2
      Top             =   6225
      Width           =   540
   End
   Begin VB.Label b2 
      Height          =   405
      Left            =   1440
      TabIndex        =   1
      Top             =   6225
      Width           =   570
   End
   Begin VB.Label b1 
      Height          =   480
      Left            =   630
      TabIndex        =   0
      Top             =   6150
      Width           =   630
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
PutOnReg
Me.Hide
If App.PrevInstance = True Then
Unload Me
End If
With nid
.cbSize = Len(nid)
.hwnd = Me.hwnd
.uId = vbNull
.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
.uCallBackMessage = WM_MOUSEMOVE
.hIcon = i(0).Picture
.szTip = "WinMoover - http://redib.surf.to" & vbNullChar
End With
Active = False

Shell_NotifyIcon NIM_ADD, nid
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Dim Result As Long
Dim msg As Long
If Me.ScaleMode = vbPixels Then
msg = X
Else
msg = X / Screen.TwipsPerPixelX
End If
     
Select Case msg
Case WM_LBUTTONDBLCLK
Form6.Show
Form6.Timer1.Enabled = True
Active = True
Case WM_RBUTTONUP
Unload Me
End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
Shell_NotifyIcon NIM_DELETE, nid
End Sub

Private Sub Timer1_Timer()
If Active = True Then
  If Icondos = 0 Then
    nid.hIcon = a(0).Picture
    Shell_NotifyIcon NIM_MODIFY, nid
    Icondos = 1
  Else
    nid.hIcon = a(1).Picture
    Shell_NotifyIcon NIM_MODIFY, nid
    Icondos = 0
  End If
Else
  If Icondos = 0 Then
    nid.hIcon = i(0).Picture
    Shell_NotifyIcon NIM_MODIFY, nid
    Icondos = 1
  Else
    nid.hIcon = i(1).Picture
    Shell_NotifyIcon NIM_MODIFY, nid
    Icondos = 0
  End If
End If
End Sub
