VERSION 5.00
Begin VB.Form FORGETPASSWORDFORM 
   Caption         =   "FORGET"
   ClientHeight    =   3420
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   ScaleHeight     =   3420
   ScaleWidth      =   5415
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "FORGET PASSWORD"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9960
      TabIndex        =   5
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11520
      TabIndex        =   4
      Top             =   6840
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11280
      MaxLength       =   10
      TabIndex        =   3
      Top             =   6000
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11280
      TabIndex        =   1
      Top             =   5160
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   10680
      Picture         =   "FORGETPASSWORDFORM.frx":0000
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Line Line4 
      BorderWidth     =   3
      X1              =   9000
      X2              =   13920
      Y1              =   7680
      Y2              =   7680
   End
   Begin VB.Line Line3 
      BorderWidth     =   3
      X1              =   13920
      X2              =   13920
      Y1              =   3600
      Y2              =   7680
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   9000
      X2              =   13920
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   9000
      X2              =   9000
      Y1              =   3600
      Y2              =   7680
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MOBILE NO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      TabIndex        =   2
      Top             =   6120
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "USERID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      TabIndex        =   0
      Top             =   5280
      Width           =   1695
   End
End
Attribute VB_Name = "FORGETPASSWORDFORM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Unload Me
LoginForm.Show
End Sub

Private Sub Command2_Click()
Dim COUNT As Integer
PASS = Text1.Text
rs.MoveFirst
While rs.EOF <> True
If rs(0) = Text1.Text And rs(2) = Text2.Text Then
COUNT = 1
End If
rs.MoveNext
Wend
FRGTPASSUSER = Text1.Text
If (COUNT = 1) Then
Unload Me
RESETPASSWORDFORM.Show
End If
End Sub

Private Sub Form_Load()
rs.MoveFirst
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack Then
Else
KeyAscii = 0
End If
End Sub
