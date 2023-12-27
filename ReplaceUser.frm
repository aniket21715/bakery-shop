VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Find & Update User"
   ClientHeight    =   5295
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8400
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   5295
   ScaleWidth      =   8400
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
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
      Left            =   10080
      TabIndex        =   5
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "UPDATE"
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
      Top             =   6960
      Width           =   1095
   End
   Begin VB.TextBox Text3 
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
      Left            =   11880
      TabIndex        =   3
      Top             =   5280
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
      IMEMode         =   3  'DISABLE
      Left            =   11880
      TabIndex        =   1
      Top             =   4320
      Width           =   2295
   End
   Begin VB.Shape Shape4 
      Height          =   735
      Left            =   11400
      Top             =   6840
      Width           =   1335
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Left            =   9960
      Top             =   6840
      Width           =   1335
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   3
      Height          =   975
      Left            =   9840
      Top             =   6720
      Width           =   3015
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   2655
      Left            =   7680
      Top             =   3720
      Width           =   7215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ENTER NEW MOBILE_NO"
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
      Left            =   7920
      TabIndex        =   2
      Top             =   5400
      Width           =   3735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ENTER REGISTERED USER_ID"
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
      Left            =   8040
      TabIndex        =   0
      Top             =   4440
      Width           =   3855
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
rs.MovePrevious
If rs.BOF = True Then
rs.MoveFirst
MsgBox "FIRST USER"
Else
Text1.Text = rs(0)
Text3.Text = rs(2)
End If
End Sub

Private Sub Command2_Click()
If Text1.Text = rs(0) Then
Dim STR As String
STR = "UPDATE LOGIN SET MOBILE = '" & Text3.Text & "' WHERE USERID = '" & Text1.Text & "' "
con.Execute STR
con.Execute "COMMIT"
MsgBox "UPDATED SUCCESSFULLY"
Else
MsgBox "USER_ID CAN NOT BE UPDATED"
End If
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
Text1.Text = rs(0)
Text3.Text = rs(2)
End Sub

Private Sub Command5_Click()
rs.MoveNext
If rs.EOF = True Then
rs.MoveLast
MsgBox "LAST USER"
Else
Text1.Text = rs(0)
Text3.Text = rs(2)
End If
End Sub

Private Sub Form_Load()
rs.MoveFirst
End Sub
