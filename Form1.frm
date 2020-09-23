VERSION 5.00
Begin VB.Form FrmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'Kein
   Caption         =   "Form1"
   ClientHeight    =   5775
   ClientLeft      =   5415
   ClientTop       =   4275
   ClientWidth     =   10065
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   385
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   671
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picdownmask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000010&
      BorderStyle     =   0  'Kein
      Height          =   3390
      Left            =   480
      ScaleHeight     =   226
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   224
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   3360
   End
   Begin VB.PictureBox jackmask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000010&
      BorderStyle     =   0  'Kein
      Height          =   3750
      Index           =   1
      Left            =   2400
      ScaleHeight     =   250
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   230
      TabIndex        =   15
      Top             =   5280
      Visible         =   0   'False
      Width           =   3450
   End
   Begin VB.PictureBox jack 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000010&
      BorderStyle     =   0  'Kein
      Height          =   3390
      Index           =   1
      Left            =   480
      ScaleHeight     =   226
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   224
      TabIndex        =   14
      Top             =   360
      Visible         =   0   'False
      Width           =   3360
   End
   Begin VB.PictureBox jack 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000010&
      BorderStyle     =   0  'Kein
      Height          =   3750
      Index           =   0
      Left            =   2520
      ScaleHeight     =   250
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   230
      TabIndex        =   13
      Top             =   5280
      Visible         =   0   'False
      Width           =   3450
   End
   Begin VB.PictureBox jackmask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000010&
      BorderStyle     =   0  'Kein
      Height          =   3390
      Index           =   0
      Left            =   360
      ScaleHeight     =   226
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   224
      TabIndex        =   12
      Top             =   360
      Visible         =   0   'False
      Width           =   3360
   End
   Begin VB.PictureBox Piccrack 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000010&
      BorderStyle     =   0  'Kein
      Height          =   3390
      Index           =   2
      Left            =   240
      ScaleHeight     =   226
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   224
      TabIndex        =   11
      Top             =   360
      Visible         =   0   'False
      Width           =   3360
   End
   Begin VB.PictureBox Piccrackmask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000010&
      BorderStyle     =   0  'Kein
      Height          =   3390
      Index           =   2
      Left            =   360
      ScaleHeight     =   226
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   224
      TabIndex        =   10
      Top             =   360
      Visible         =   0   'False
      Width           =   3360
   End
   Begin VB.PictureBox Piccrackmask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000010&
      BorderStyle     =   0  'Kein
      Height          =   3390
      Index           =   1
      Left            =   480
      ScaleHeight     =   226
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   224
      TabIndex        =   9
      Top             =   360
      Visible         =   0   'False
      Width           =   3360
   End
   Begin VB.PictureBox Piccrack 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000010&
      BorderStyle     =   0  'Kein
      Height          =   3390
      Index           =   1
      Left            =   240
      ScaleHeight     =   226
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   224
      TabIndex        =   8
      Top             =   360
      Visible         =   0   'False
      Width           =   3360
   End
   Begin VB.PictureBox Piccrack 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000010&
      BorderStyle     =   0  'Kein
      Height          =   3390
      Index           =   0
      Left            =   480
      ScaleHeight     =   226
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   224
      TabIndex        =   7
      Top             =   480
      Visible         =   0   'False
      Width           =   3360
   End
   Begin VB.PictureBox Piccrackmask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000010&
      BorderStyle     =   0  'Kein
      Height          =   3390
      Index           =   0
      Left            =   360
      ScaleHeight     =   226
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   224
      TabIndex        =   6
      Top             =   720
      Visible         =   0   'False
      Width           =   3360
   End
   Begin VB.PictureBox Picdown 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000010&
      BorderStyle     =   0  'Kein
      Height          =   3750
      Left            =   2400
      ScaleHeight     =   250
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   230
      TabIndex        =   5
      Top             =   5280
      Visible         =   0   'False
      Width           =   3450
   End
   Begin VB.PictureBox Picbg 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000010&
      BorderStyle     =   0  'Kein
      Height          =   3750
      Left            =   5880
      ScaleHeight     =   250
      ScaleMode       =   0  'Benutzerdefiniert
      ScaleWidth      =   250
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   3750
   End
   Begin VB.PictureBox Picbuffer 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000010&
      BorderStyle     =   0  'Kein
      Height          =   3750
      Left            =   5520
      ScaleHeight     =   250
      ScaleMode       =   0  'Benutzerdefiniert
      ScaleWidth      =   250
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   3750
   End
   Begin VB.PictureBox PicMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000010&
      BorderStyle     =   0  'Kein
      Height          =   3750
      Left            =   2400
      ScaleHeight     =   250
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   230
      TabIndex        =   1
      Top             =   5280
      Visible         =   0   'False
      Width           =   3450
   End
   Begin VB.PictureBox Pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000010&
      BorderStyle     =   0  'Kein
      Height          =   3750
      Left            =   2400
      ScaleHeight     =   250
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   230
      TabIndex        =   0
      Top             =   5280
      Visible         =   0   'False
      Width           =   3450
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'some variables
Dim oldx As Long, oldy As Long
Dim down As Boolean
Dim jackh As Boolean
Dim pos As Integer, hammering As Boolean, first As Boolean, sound As Integer

Private Sub Form_Activate()
Call get_desktop
ShowCursor 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
ShowCursor 1
End
End Sub
Private Sub Form_Load()
                                'Resize all elements that needs to be resized
Picbg.Move 0, 0, Screen.Width / 15, Screen.Height / 15
Picbuffer.Move 0, 0, Pic.ScaleWidth, jack(0).ScaleHeight
Me.Move 0, 0, Screen.Width, Screen.Height
'Gotta load them pictures...
Piccrackmask(0).Picture = LoadPicture("crackmask0.gif")
FrmSplash.ProgressBar1.Value = FrmSplash.ProgressBar1.Value + 1
Piccrackmask(1).Picture = LoadPicture("crackmask1.gif")
FrmSplash.ProgressBar1.Value = FrmSplash.ProgressBar1.Value + 1
Piccrackmask(2).Picture = LoadPicture("crackmask2.gif")
FrmSplash.ProgressBar1.Value = FrmSplash.ProgressBar1.Value + 1
Piccrack(0).Picture = LoadPicture("crack0.gif")
FrmSplash.ProgressBar1.Value = FrmSplash.ProgressBar1.Value + 1
Piccrack(1).Picture = LoadPicture("crack1.gif")
FrmSplash.ProgressBar1.Value = FrmSplash.ProgressBar1.Value + 1
Piccrack(2).Picture = LoadPicture("crack2.gif")
FrmSplash.ProgressBar1.Value = FrmSplash.ProgressBar1.Value + 1
Pic.Picture = LoadPicture("pic.gif")
FrmSplash.ProgressBar1.Value = FrmSplash.ProgressBar1.Value + 1
PicMask.Picture = LoadPicture("picmask.gif")
FrmSplash.ProgressBar1.Value = FrmSplash.ProgressBar1.Value + 1
Picdown.Picture = LoadPicture("picdown.gif")
FrmSplash.ProgressBar1.Value = FrmSplash.ProgressBar1.Value + 1
Picdownmask.Picture = LoadPicture("picdownmask.gif")
FrmSplash.ProgressBar1.Value = FrmSplash.ProgressBar1.Value + 1
jack(0).Picture = LoadPicture("jack0.gif")
FrmSplash.ProgressBar1.Value = FrmSplash.ProgressBar1.Value + 1
jackmask(0).Picture = LoadPicture("jackmask0.gif")
FrmSplash.ProgressBar1.Value = FrmSplash.ProgressBar1.Value + 1
jack(1).Picture = LoadPicture("jack1.gif")
FrmSplash.ProgressBar1.Value = FrmSplash.ProgressBar1.Value + 1
jackmask(1).Picture = LoadPicture("jackmask1.gif")
FrmSplash.ProgressBar1.Value = FrmSplash.ProgressBar1.Value + 1
FrmSplash.Label1.Caption = FrmSplash.Label1.Caption & " Done!"
FrmSplash.Label2.Visible = True
FrmSplash.Command1.Visible = True
pos = 0
first = False

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim sound As Integer
    Randomize
    If Button = 1 Then
    
    If jackh = False Then
    down = True
    'Need some sound, and random too. We dont want to bore anyone :)
    sound = 4 * Rnd(Now)
    Select Case sound
    Case 0
    sndPlaySound "1.wav", &H1
    Case 1
    sndPlaySound "2.wav", &H1
    Case 2
    sndPlaySound "3.wav", &H1
    Case 3
    sndPlaySound "4.wav", &H1
    Case 4
    sndPlaySound "5.wav", &H1
    End Select
    'Randomize a index for the crack to be shown, again, no boredoom here..
    num = 2 * Rnd(Now)
    'Draw the crack on the desktop buffer
    Call BitBlt(Picbg.hdc, x - 200, y - 60, Piccrack(1).ScaleWidth, Piccrack(1).ScaleHeight, Piccrackmask(num).hdc, 0, 0, vbSrcAnd)
    Call BitBlt(Picbg.hdc, x - 200, y - 60, Piccrack(1).ScaleWidth, Piccrack(1).ScaleHeight, Piccrack(num).hdc, 0, 0, vbSrcInvert)
    'Draw the desktop on the buffer
    Call BitBlt(Picbuffer.hdc, -x + (Pic.ScaleWidth / 2), -y + (Pic.ScaleHeight / 2), Picbg.ScaleWidth, Picbg.ScaleHeight, Picbg.hdc, 0, 0, vbSrcCopy)
    'And the hammer...
    Call BitBlt(Picbuffer.hdc, 0, 0, Pic.ScaleWidth, Pic.ScaleHeight, Picdownmask.hdc, 0, 0, vbSrcAnd)
    Call BitBlt(Picbuffer.hdc, 0, 0, Pic.ScaleWidth, Pic.ScaleHeight, Picdown.hdc, 0, 0, vbSrcInvert)
    'Draw the crack on the form
    Call BitBlt(Me.hdc, x - (Pic.ScaleWidth / 2) - 100, y - (Pic.ScaleHeight / 2), Pic.ScaleWidth + 50, Pic.ScaleHeight + 150, Picbg.hdc, x - (Pic.ScaleWidth / 2) - 100, y - (Pic.ScaleHeight / 2), vbSrcCopy)
    'And the hammer ontop
    Call BitBlt(Me.hdc, x - (Pic.ScaleWidth / 2), y - (Pic.ScaleHeight / 2), Pic.ScaleWidth, Pic.ScaleHeight, Picbuffer.hdc, 0, 0, SRCCOPY)
    'Make everything show
    Me.Refresh
    Else
    hammering = True
    End If
    End If
    If Button = 2 Then
    If jackh = False Then
    jackh = True
    Else
    jackh = False
    End If
    Call BitBlt(Me.hdc, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Picbg.hdc, 0, 0, vbSrcCopy)
    Me.Refresh
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'DoEvents
'Dont want the hammer to move when it's down
If down = False Then
If jackh = False Then
    'Fill the last place the hammer was on with a portion of the desktop picture
    Call StretchBlt(Me.hdc, oldx - (Pic.ScaleWidth / 2), oldy - (Pic.ScaleHeight / 2), Pic.ScaleWidth, Pic.ScaleHeight, Picbg.hdc, oldx - (Pic.ScaleWidth / 2), oldy - (Pic.ScaleHeight / 2), Pic.ScaleWidth, Pic.ScaleHeight, vbSrcCopy)
    'Draw the desktop on the buffer
    Call BitBlt(Picbuffer.hdc, -x + (Pic.ScaleWidth / 2), -y + (Pic.ScaleHeight / 2), Picbg.ScaleWidth, Picbg.ScaleHeight, Picbg.hdc, 0, 0, vbSrcCopy)
    'And the hammer
    Call BitBlt(Picbuffer.hdc, 0, 0, Pic.ScaleWidth, Pic.ScaleHeight, PicMask.hdc, 0, 0, vbSrcAnd)
    Call BitBlt(Picbuffer.hdc, 0, 0, Pic.ScaleWidth, Pic.ScaleHeight, Pic.hdc, 0, 0, vbSrcInvert)
    'Draw the works on the form
    Call BitBlt(Me.hdc, x - (Pic.ScaleWidth / 2), y - (Pic.ScaleHeight / 2), Pic.ScaleWidth, Pic.ScaleHeight, Picbuffer.hdc, 0, 0, SRCCOPY)
    'Need to know where the hammer was last time so i save the variables
    oldx = x
    oldy = y
    'Make everything show
    Me.Refresh
Else
    If hammering = True Then
    If pos = 0 Then
    pos = 1
    Else
    pos = 0
    End If
    Call hammer
    
    'Fill the last place the hammer was on with a portion of the desktop picture
    Call StretchBlt(Me.hdc, oldx - (jack(0).ScaleWidth / 2), oldy - (jack(0).ScaleHeight / 2), jack(0).ScaleWidth, jack(0).ScaleHeight, Picbg.hdc, oldx - (jack(0).ScaleWidth / 2), oldy - (jack(0).ScaleHeight / 2), jack(0).ScaleWidth, jack(0).ScaleHeight, vbSrcCopy)
    'Draw the desktop on the buffer
    Call BitBlt(Picbuffer.hdc, -x + (jack(0).ScaleWidth / 2), -y + (jack(0).ScaleHeight / 2), Picbg.ScaleWidth, Picbg.ScaleHeight, Picbg.hdc, 0, 0, vbSrcCopy)
    'And the hammer
    Call BitBlt(Picbuffer.hdc, 0, 0, jack(0).ScaleWidth, jack(0).ScaleHeight, jackmask(pos).hdc, 0, 0, vbSrcAnd)
    Call BitBlt(Picbuffer.hdc, 0, 0, jack(0).ScaleWidth, jack(0).ScaleHeight, jack(pos).hdc, 0, 0, vbSrcInvert)
    'Draw the works on the form
    Call BitBlt(Me.hdc, x - (jack(0).ScaleWidth / 2), y - (jack(0).ScaleHeight / 2), jack(0).ScaleWidth, jack(0).ScaleHeight, Picbuffer.hdc, 0, 0, SRCCOPY)
    'Need to know where the hammer was last time so i save the variables
    oldx = x
    oldy = y
    'Make everything show
    Me.Refresh
    Else
    'Fill the last place the hammer was on with a portion of the desktop picture
    Call StretchBlt(Me.hdc, oldx - (jack(0).ScaleWidth / 2), oldy - (jack(0).ScaleHeight / 2), jack(0).ScaleWidth, jack(0).ScaleHeight, Picbg.hdc, oldx - (jack(0).ScaleWidth / 2), oldy - (jack(0).ScaleHeight / 2), jack(0).ScaleWidth, jack(0).ScaleHeight, vbSrcCopy)
    'Draw the desktop on the buffer
    Call BitBlt(Picbuffer.hdc, -x + (jack(0).ScaleWidth / 2), -y + (jack(0).ScaleHeight / 2), Picbg.ScaleWidth, Picbg.ScaleHeight, Picbg.hdc, 0, 0, vbSrcCopy)
    'And the hammer
    Call BitBlt(Picbuffer.hdc, 0, 0, jack(0).ScaleWidth, jack(0).ScaleHeight, jackmask(pos).hdc, 0, 0, vbSrcAnd)
    Call BitBlt(Picbuffer.hdc, 0, 0, jack(0).ScaleWidth, jack(0).ScaleHeight, jack(pos).hdc, 0, 0, vbSrcInvert)
    'Draw the works on the form
    Call BitBlt(Me.hdc, x - (jack(0).ScaleWidth / 2), y - (jack(0).ScaleHeight / 2), jack(0).ScaleWidth, jack(0).ScaleHeight, Picbuffer.hdc, 0, 0, SRCCOPY)
    'Need to know where the hammer was last time so i save the variables
    oldx = x
    oldy = y
    'Make everything show
    Me.Refresh
    End If
    End If
End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'OK to move hammer again
hammering = False
down = False

pos = 0
End Sub

Private Sub hammer()
sndPlaySound "burst.wav", &H1


num = 2 * Rnd(Now)

    'Draw the crack on the desktop buffer
    Call BitBlt(Picbg.hdc, oldx - 200, oldy - 20, Piccrack(num).ScaleWidth, Piccrack(num).ScaleHeight, Piccrackmask(num).hdc, 0, 0, vbSrcAnd)
    Call BitBlt(Picbg.hdc, oldx - 200, oldy - 20, Piccrack(num).ScaleWidth, Piccrack(num).ScaleHeight, Piccrack(num).hdc, 0, 0, vbSrcInvert)
    'Draw the desktop on the buffer
    Call BitBlt(Picbuffer.hdc, -oldx + (jack(0).ScaleWidth / 2), -oldy + (jack(0).ScaleHeight / 2), Picbg.ScaleWidth, Picbg.ScaleHeight, Picbg.hdc, 0, 0, vbSrcCopy)
    'And the hammer...
    Call BitBlt(Picbuffer.hdc, 0, 0, jack(0).ScaleWidth, jack(0).ScaleHeight, jackmask(pos).hdc, 0, 0, vbSrcAnd)
    Call BitBlt(Picbuffer.hdc, 0, 0, jack(0).ScaleWidth, jack(0).ScaleHeight, jack(pos).hdc, 0, 0, vbSrcInvert)
    'Draw the crack on the form
    Call BitBlt(Me.hdc, oldx - (jack(0).ScaleWidth / 2) - 125, oldy + 21 - (jack(0).ScaleHeight / 2), jack(0).ScaleWidth + 50, jack(0).ScaleHeight + 150, Picbg.hdc, oldx - (Pic.ScaleWidth / 2) - 100, oldy - (Pic.ScaleHeight / 2), vbSrcCopy)
    'And the hammer ontop
    Call BitBlt(Me.hdc, oldx - (jack(0).ScaleWidth / 2), oldy - (jack(0).ScaleHeight / 2), jack(0).ScaleWidth, jack(0).ScaleHeight, Picbuffer.hdc, 0, 0, SRCCOPY)
    'Make everything show
    Me.Refresh

End Sub

