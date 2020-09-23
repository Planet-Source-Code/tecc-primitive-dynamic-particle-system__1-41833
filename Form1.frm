VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5715
   ClientLeft      =   1650
   ClientTop       =   1830
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   6585
   Begin VB.PictureBox kl 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3780
      Left            =   3780
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   3720
      ScaleWidth      =   3750
      TabIndex        =   3
      Top             =   3420
      Visible         =   0   'False
      Width           =   3810
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   3420
      Top             =   2820
   End
   Begin VB.PictureBox PicMax 
      AutoRedraw      =   -1  'True
      Height          =   1935
      Left            =   180
      ScaleHeight     =   125
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   185
      TabIndex        =   2
      Top             =   2400
      Width           =   2835
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   435
      Left            =   2520
      TabIndex        =   1
      Top             =   4200
      Width           =   675
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   2175
      Left            =   0
      ScaleHeight     =   141
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   433
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   6555
   End
   Begin VB.Menu mnuV 
      Caption         =   "&View"
      Begin VB.Menu mnuDBT 
         Caption         =   "&Debug Test"
      End
      Begin VB.Menu mnuVRWD 
         Caption         =   "&Real World Demo"
      End
   End
   Begin VB.Menu mnuH 
      Caption         =   "& Help"
      Begin VB.Menu sssss 
         Caption         =   "Help"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private oldx As Long
Private oldy As Long
Private fx As Long
Private fy As Long
Private md As Boolean

Private Sub Form_Load()

Me.Move Me.Left, Me.Top, 400 * 15, 400 * 15
Picture1.Picture = Nothing
Me.Show
InitParticles 500, 0, 1, Picture1.ScaleWidth / 2, 0, 1, 1, 500, RGB(100, 255, 100)
Timer1.Enabled = False
MsgBox "Input: " & vbCrLf & vbCrLf & " Mouse buttons: " & vbCrLf & "  Right: Draw Border" & vbCrLf & "  Left: Erase border" & vbCrLf & "  Double Click: Clear all borders" & vbCrLf & vbCrLf & " Note: Particles cannot penetrate borders!"
MainLoop

End Sub

Private Sub Form_Resize()
Picture1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
PicMax.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub mnuDBT_Click()
Me.Move Me.Left, Me.Top, 400 * 15, 400 * 15
Picture1.Picture = Nothing
Me.Show
InitParticles 500, 0, 1, Picture1.ScaleWidth / 2, 0, 1, 1, 500, RGB(100, 255, 100)
Timer1.Enabled = False
MainLoop

End Sub

Private Sub mnuVRWD_Click()
Me.Move Me.Left, Me.Top, 255 * 15, 290 * 15
Me.Show
Me.Caption = "Snow!"
Picture1.Picture = kl.Picture
InitParticles 500, 0, 1, Picture1.ScaleWidth / 2, 0, 1, 1, 500, RGB(200, 200, 255)
Timer1.Enabled = True
MainLoop
End Sub

Private Sub PicMax_DblClick()
Picture1.Cls
End Sub

Private Sub PicMax_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    
    'Picture1.Line (oldx, oldy)-(X, Y)
    'Picture1.Line (oldx + 2, oldy + 2)-(X + 1, Y + 1)
    'Picture1.Line (oldx + 2, oldy + 2)-(X + 2, Y + 2)
    'oldx = X
    'oldy = Y
    Picture1.FillColor = vbBlack
    Picture1.ForeColor = vbBlack
    Picture1.Circle (X, Y), 3
ElseIf Button = 2 Then
    Picture1.FillColor = Picture1.BackColor
    Picture1.ForeColor = Picture1.BackColor
    Picture1.Circle (X, Y), 3
End If
End Sub

Private Sub PicMax_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
md = False
End Sub

Private Sub Picture1_DblClick()
Picture1.Cls
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
md = True
fx = X
fy = Y
Picture1.CurrentX = fx
     Picture1.CurrentY = fy
oldx = X
oldy = Y
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    
    'Picture1.Line (oldx, oldy)-(X, Y)
    'Picture1.Line (oldx + 2, oldy + 2)-(X + 1, Y + 1)
    'Picture1.Line (oldx + 2, oldy + 2)-(X + 2, Y + 2)
    'oldx = X
    'oldy = Y
    Picture1.FillColor = vbBlack
    Picture1.ForeColor = vbBlack
    Picture1.Circle (X, Y), 3
ElseIf Button = 2 Then
    Picture1.FillColor = Picture1.BackColor
    Picture1.ForeColor = Picture1.BackColor
    Picture1.Circle (X, Y), 3
End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
md = False
End Sub

Private Sub sssss_Click()
MsgBox "Use the mouse and it's 2 buttons like in paint!"
End Sub

Private Sub Timer1_Timer()
SourceP.X = Rnd * Picture1.ScaleWidth
End Sub
