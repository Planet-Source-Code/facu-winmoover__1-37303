VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   Caption         =   "Form6"
   ClientHeight    =   90
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   90
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   90
   ScaleWidth      =   90
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   -525
      Top             =   -195
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
StayOnTop Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.Visible = False
Timer1.Enabled = False
nid.hIcon = Form4.moving.Picture
Shell_NotifyIcon NIM_MODIFY, nid
DoEvents
abd = WinFromXY
Call ReleaseCapture
ReturnVal = SendMessage(abd, &HA1, 2, 0)
Active = False
Unload Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.Top = MouseY * Screen.TwipsPerPixelX - (Me.Height / 2)
Me.Left = MouseX * Screen.TwipsPerPixelY - (Me.Width / 2)
StayOnTop Me
End Sub

Private Sub Timer1_Timer()
Me.Top = MouseY * Screen.TwipsPerPixelX - (Me.Height / 2)
Me.Left = MouseX * Screen.TwipsPerPixelY - (Me.Width / 2)
End Sub
