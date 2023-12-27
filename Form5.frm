VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "ProductsUpdate"
   ClientHeight    =   4650
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7170
   LinkTopic       =   "Form5"
   MDIChild        =   -1  'True
   ScaleHeight     =   12495
   ScaleWidth      =   22920
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command9 
      Caption         =   "CLOSE"
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
      Left            =   9840
      TabIndex        =   19
      Top             =   9480
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
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
      Left            =   11280
      TabIndex        =   18
      Top             =   9480
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "FILTER OFF"
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
      Left            =   7200
      TabIndex        =   16
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "FILTER ON"
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
      Left            =   8640
      TabIndex        =   15
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "UPDATE PRODUCTS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Left            =   6960
      TabIndex        =   0
      Top             =   2520
      Width           =   7815
      Begin VB.CommandButton Command7 
         Caption         =   "SHOW"
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
         Left            =   3360
         TabIndex        =   17
         Top             =   5520
         Width           =   1095
      End
      Begin VB.TextBox Text5 
         Enabled         =   0   'False
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
         Left            =   3840
         TabIndex        =   14
         Top             =   3960
         Width           =   1095
      End
      Begin VB.TextBox Text4 
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
         Left            =   3840
         TabIndex        =   13
         Top             =   3240
         Width           =   1455
      End
      Begin VB.CommandButton Command4 
         Caption         =   "NEXT"
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
         Left            =   4800
         TabIndex        =   12
         Top             =   5520
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "LAST"
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
         Left            =   6240
         TabIndex        =   6
         Top             =   5520
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "PREVIOUS"
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
         Left            =   1920
         TabIndex        =   5
         Top             =   5520
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
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
         Left            =   3840
         TabIndex        =   4
         Top             =   2400
         Width           =   2055
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
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
         Left            =   3840
         TabIndex        =   3
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
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
         Left            =   3840
         TabIndex        =   2
         Top             =   960
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "FIRST"
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
         Left            =   480
         TabIndex        =   1
         Top             =   5520
         Width           =   1095
      End
      Begin VB.Shape Shape6 
         BorderWidth     =   3
         Height          =   4215
         Left            =   1440
         Top             =   720
         Width           =   4695
      End
      Begin VB.Shape Shape5 
         Height          =   735
         Left            =   3240
         Top             =   5400
         Width           =   1335
      End
      Begin VB.Shape Shape4 
         Height          =   735
         Left            =   4680
         Top             =   5400
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "GST"
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
         Left            =   1800
         TabIndex        =   11
         Top             =   4080
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "PRICE"
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
         Left            =   1800
         TabIndex        =   10
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "VARIETIES UNIT"
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
         Left            =   1800
         TabIndex        =   9
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "CAKE VARIETIES"
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
         Left            =   1800
         TabIndex        =   8
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "PRODUCT_NO"
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
         Left            =   1800
         TabIndex        =   7
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Line Line5 
         BorderWidth     =   3
         X1              =   240
         X2              =   240
         Y1              =   5280
         Y2              =   6240
      End
      Begin VB.Line Line6 
         BorderWidth     =   3
         X1              =   7560
         X2              =   7560
         Y1              =   5280
         Y2              =   6240
      End
      Begin VB.Line Line7 
         BorderWidth     =   3
         X1              =   240
         X2              =   7560
         Y1              =   5280
         Y2              =   5280
      End
      Begin VB.Line Line8 
         BorderWidth     =   3
         X1              =   240
         X2              =   7560
         Y1              =   6240
         Y2              =   6240
      End
      Begin VB.Shape Shape1 
         Height          =   735
         Left            =   360
         Top             =   5400
         Width           =   1335
      End
      Begin VB.Shape Shape2 
         Height          =   735
         Left            =   1800
         Top             =   5400
         Width           =   1335
      End
      Begin VB.Shape Shape3 
         Height          =   735
         Left            =   6120
         Top             =   5400
         Width           =   1335
      End
   End
   Begin VB.Shape Shape13 
      BorderWidth     =   3
      Height          =   9975
      Left            =   6120
      Top             =   840
      Width           =   9615
   End
   Begin VB.Shape Shape12 
      BorderWidth     =   3
      Height          =   975
      Left            =   9600
      Top             =   9240
      Width           =   3015
   End
   Begin VB.Shape Shape11 
      Height          =   735
      Left            =   11160
      Top             =   9360
      Width           =   1335
   End
   Begin VB.Shape Shape10 
      Height          =   735
      Left            =   9720
      Top             =   9360
      Width           =   1335
   End
   Begin VB.Shape Shape9 
      BorderWidth     =   3
      Height          =   975
      Left            =   6960
      Top             =   1320
      Width           =   3015
   End
   Begin VB.Shape Shape8 
      Height          =   735
      Left            =   8520
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Shape Shape7 
      Height          =   735
      Left            =   7080
      Top             =   1440
      Width           =   1335
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim GST As Double
Private Sub Command1_Click()
rs.MovePrevious
If rs.BOF = True Then
rs.MoveFirst
MsgBox "FIRST USER"
Else
Text1.Text = rs(0)
Text2.Text = rs(1)
Text3.Text = rs(2)
Text4.Text = rs(3)
Text5.Text = rs(4)
End If
End Sub
Private Sub Command2_Click()
rs.MoveLast
Text1.Text = rs(0)
Text2.Text = rs(1)
Text3.Text = rs(2)
Text4.Text = rs(3)
Text5.Text = rs(4)
End Sub
Private Sub Command3_Click()
rs.MoveFirst
Text1.Text = rs(0)
Text2.Text = rs(1)
Text3.Text = rs(2)
Text4.Text = rs(3)
Text5.Text = rs(4)
End Sub

Private Sub Command4_Click()
rs.MoveNext
If rs.EOF = True Then
rs.MoveLast
MsgBox "LAST USER"
Else
Text1.Text = rs(0)
Text2.Text = rs(1)
Text3.Text = rs(2)
Text4.Text = rs(3)
Text5.Text = rs(4)
End If
End Sub

Private Sub Command5_Click()
Dim SEARCH1 As String
Dim SEARCH2 As String
SEARCH1 = InputBox("ENTER CAKE VARIETY NAME")
SEARCH2 = InputBox("ENTER CAKE VARIETY UNIT")
If rs.State = 1 Then rs.Close
rs.Open "SELECT * FROM PRODUCT where P_NAME LIKE '%" & SEARCH1 & "%' AND P_UNIT LIKE '%" & SEARCH2 & "%' ", con, adOpenDynamic
Text1.Text = rs(0)
Text2.Text = rs(1)
Text3.Text = rs(2)
Text4.Text = rs(3)
Text5.Text = rs(4)
End Sub

Private Sub Command6_Click()
If rs.State = 1 Then rs.Close
rs.Open "SELECT * FROM PRODUCT", con, adOpenDynamic
End Sub

Private Sub Command7_Click()
Text1.Text = rs(0)
Text2.Text = rs(1)
Text3.Text = rs(2)
Text4.Text = rs(3)
Text5.Text = rs(4)
End Sub

Private Sub Command8_Click()
Dim STR As String
STR = "UPDATE PRODUCT SET P_RATE = '" & Text4.Text & "' , P_GST = '" & Text5.Text & "'  WHERE P_NO = '" & Text1.Text & "' "
con.Execute STR
con.Execute "COMMIT"
MsgBox "UPDATED SUCCESSFULLY"
End Sub

Private Sub Command9_Click()
Unload Me
End Sub

Private Sub Form_Load()
If rs.State = 1 Then rs.Close
rs.Open "SELECT * FROM PRODUCT", con, adOpenDynamic
End Sub

Private Sub Text4_KeyUp(KeyCode As Integer, Shift As Integer)
GST = Val(Text4.Text) * 0.12
Text5.Text = GST
End Sub
