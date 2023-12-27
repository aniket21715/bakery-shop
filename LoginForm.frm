VERSION 5.00
Begin VB.Form LOGINFORM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "3"
   ClientHeight    =   11670
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   20985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11670
   ScaleMode       =   0  'User
   ScaleWidth      =   67613.1
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "      NOT        NOW"
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
      Left            =   12120
      TabIndex        =   6
      Top             =   7440
      Width           =   1095
   End
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
      Left            =   9480
      TabIndex        =   5
      Top             =   7440
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Check1"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   13440
      TabIndex        =   4
      Top             =   6720
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "LOGIN"
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
      Left            =   10800
      TabIndex        =   3
      Top             =   7440
      Width           =   1095
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
      IMEMode         =   3  'DISABLE
      Left            =   11040
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "HARSH3132"
      Top             =   6600
      Width           =   2295
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
      Left            =   11040
      TabIndex        =   0
      Text            =   "HARSH PURBEY"
      Top             =   5760
      Width           =   2295
   End
   Begin VB.Label Label3 
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
      Left            =   9120
      TabIndex        =   7
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Line Line5 
      BorderWidth     =   3
      X1              =   27451.21
      X2              =   46009.77
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Image Image1 
      Height          =   1815
      Left            =   10440
      Picture         =   "LoginForm.frx":0000
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Line Line4 
      BorderWidth     =   3
      X1              =   46009.77
      X2              =   46009.77
      Y1              =   4320
      Y2              =   8400
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   27451.21
      X2              =   46009.77
      Y1              =   8400
      Y2              =   8400
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   27451.21
      X2              =   27451.21
      Y1              =   8400
      Y2              =   4320
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   9120
      TabIndex        =   2
      Top             =   6720
      Width           =   1695
   End
End
Attribute VB_Name = "LoginForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = 1 Then
Text2.PasswordChar = ""
Else
Text2.PasswordChar = "*"
End If
End Sub

Private Sub Command1_Click()
Dim COUNT As Integer
While rs.EOF <> True
If rs(0) = Text1.Text And rs(1) = Text2.Text Then
COUNT = 1
End If
rs.MoveNext
Wend
If (COUNT = 1) Then
Unload Me
MDIForm1.Show
End If
End Sub

Private Sub Command2_Click()
Unload Me
FORGETPASSWORDFORM.Show
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Form_Load()
If rs.State = 1 Then rs.Close
rs.Open "select * from Login", con, adOpenDynamic
rs.MoveFirst
End Sub

