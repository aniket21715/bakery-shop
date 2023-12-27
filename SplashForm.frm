VERSION 5.00
Begin VB.Form SPLASHFORM 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SPLASH"
   ClientHeight    =   11670
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   21555
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   14.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11670
   ScaleWidth      =   21555
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   0
      Top             =   0
   End
   Begin VB.Image Image3 
      Height          =   2535
      Left            =   9120
      Stretch         =   -1  'True
      Top             =   480
      Width           =   4695
   End
   Begin VB.Image Image1 
      Height          =   8175
      Left            =   15600
      Picture         =   "SplashForm.frx":0000
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   6975
   End
   Begin VB.Image Image2 
      Height          =   7935
      Left            =   1080
      Picture         =   "SplashForm.frx":300042
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   6855
   End
   Begin VB.Line Line4 
      BorderWidth     =   5
      X1              =   8280
      X2              =   15240
      Y1              =   8040
      Y2              =   8040
   End
   Begin VB.Line Line3 
      BorderWidth     =   5
      X1              =   15240
      X2              =   15240
      Y1              =   3600
      Y2              =   8040
   End
   Begin VB.Line Line2 
      BorderWidth     =   5
      X1              =   8280
      X2              =   15240
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line1 
      BorderWidth     =   5
      X1              =   8280
      X2              =   8280
      Y1              =   3600
      Y2              =   8040
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "START"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   11280
      TabIndex        =   4
      Top             =   6120
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "            WELCOME  TO           CAKE SHOP MANAGEMENT SYSTEM"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1815
      Left            =   9840
      TabIndex        =   3
      Top             =   4200
      Width           =   4215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Your Application is loading -"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   9120
      TabIndex        =   2
      Top             =   6840
      Width           =   4575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   13680
      TabIndex        =   1
      Top             =   6840
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   9000
      TabIndex        =   0
      Top             =   7320
      Width           =   5655
   End
End
Attribute VB_Name = "SplashForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
con.ConnectionString = "Provider=MSDAORA.1;Password=PURBEY;User ID=HARSH;Persist Security Info=True"
con.Open
Label1.Width = 0
Label2.Caption = 0 & " % "
Label3.Visible = False
Label2.Visible = False
End Sub

Private Sub Label5_Click()
Label5.Visible = False
Label3.Visible = True
Label2.Visible = True
End Sub
Private Sub Timer1_Timer()
If Label5.Visible = False Then
Label1.Width = Label1.Width + 56
Label2.Caption = Val(Label2.Caption) + 1 & " % "
If Val(Label2.Caption) > 100 Then
Timer1.Enabled = False
LoginForm.Show
Unload Me
Else
End If
End If
End Sub
