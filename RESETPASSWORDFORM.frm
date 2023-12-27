VERSION 5.00
Begin VB.Form RESETPASSWORDFORM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RESET"
   ClientHeight    =   11175
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   18870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11175
   ScaleWidth      =   18870
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Check1"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   13440
      TabIndex        =   5
      Top             =   4680
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "RESET PASSWORD"
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
      Left            =   10560
      TabIndex        =   4
      Top             =   6360
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
      Left            =   11160
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   5400
      Width           =   2175
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
      IMEMode         =   3  'DISABLE
      Left            =   11160
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   4560
      Width           =   2175
   End
   Begin VB.Line Line4 
      BorderWidth     =   3
      X1              =   14040
      X2              =   14040
      Y1              =   3720
      Y2              =   7200
   End
   Begin VB.Line Line3 
      BorderWidth     =   3
      X1              =   8280
      X2              =   14040
      Y1              =   7200
      Y2              =   7200
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   8280
      X2              =   14040
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   8280
      X2              =   8280
      Y1              =   7200
      Y2              =   3720
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CONFORM PASSWORD"
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
      Left            =   8640
      TabIndex        =   1
      Top             =   5520
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NEW PASSWORD"
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
      Left            =   9000
      TabIndex        =   0
      Top             =   4680
      Width           =   1695
   End
End
Attribute VB_Name = "RESETPASSWORDFORM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()
If Check1.Value = 1 Then
Text1.PasswordChar = ""
Else
Text1.PasswordChar = "*"
End If
End Sub

Private Sub Command2_Click()
If Len(Text1.Text) < 8 Then
MsgBox "MAKE ATLEAST 6 CHARACTERS PASSWORD"
ElseIf Text1.Text <> Text2.Text Then
MsgBox "PASSWORD NOT MATCHED ENTER CAREFULLY"
Else
Dim STR As String
STR = "UPDATE Login set password = '" & Text2.Text & "' WHERE USERID = '" & PASS & "' "
con.Execute STR
con.Execute "commit"
Unload Me
LoginForm.Show
End If
End Sub

