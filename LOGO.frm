VERSION 5.00
Begin VB.Form LOGO 
   Caption         =   "Form8"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form8"
   MDIChild        =   -1  'True
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3960
      Top             =   11880
   End
   Begin VB.Shape Shape2 
      Height          =   375
      Left            =   20640
      Top             =   10800
      Width           =   1935
   End
   Begin VB.Shape Shape3 
      Height          =   975
      Left            =   20520
      Top             =   10320
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      Height          =   375
      Left            =   20640
      Top             =   10440
      Width           =   1935
   End
   Begin VB.Image Image5 
      Height          =   7575
      Left            =   17400
      Picture         =   "LOGO.frx":0000
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   5295
   End
   Begin VB.Image Image4 
      Height          =   7095
      Left            =   960
      Picture         =   "LOGO.frx":2D2EAA
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   4695
   End
   Begin VB.Image Image3 
      Height          =   2175
      Left            =   20280
      Picture         =   "LOGO.frx":5A5D54
      Stretch         =   -1  'True
      Top             =   360
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   2175
      Left            =   720
      Picture         =   "LOGO.frx":67D016
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   20640
      TabIndex        =   2
      Top             =   10440
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   20640
      TabIndex        =   1
      Top             =   10800
      Width           =   1935
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CAKE  SHOP  MANAGEMENT  SYSTEM"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   42
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1095
      Left            =   2160
      TabIndex        =   0
      Top             =   600
      Width           =   18855
   End
   Begin VB.Image Image2 
      Height          =   12735
      Left            =   -120
      Picture         =   "LOGO.frx":7542D8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   23160
   End
End
Attribute VB_Name = "LOGO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Label1.Caption = Format(Date, "DD-MMM-YYYY")
Label2.Caption = Format$(Time$, "hh:mm:ss AM/PM")
Timer1.Enabled = True

End Sub

Private Sub Timer1_Timer()
Label2.Caption = Format$(Time$, "hh:mm:ss AM/PM")
End Sub
