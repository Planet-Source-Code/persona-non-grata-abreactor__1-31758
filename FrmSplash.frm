VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmSplash 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Initializing..."
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   240
      Top             =   1560
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      Max             =   14
      Scrolling       =   1
   End
   Begin VB.Label Label2 
      Caption         =   $"FrmSplash.frx":0000
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   "Loading graphics..."
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "FrmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.Hide
Unload Me
FrmMain.Show 1
End Sub

Private Sub Form_Activate()
Me.Show

End Sub

Private Sub Form_Load()

Me.Show
End Sub

Private Sub Timer1_Timer()
Load FrmMain
End Sub
