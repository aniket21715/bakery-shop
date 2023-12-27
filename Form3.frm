VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5835
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11355
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5835
   ScaleWidth      =   11355
   Begin VB.CommandButton Command3 
      Caption         =   "ADD"
      Height          =   495
      Left            =   6240
      TabIndex        =   6
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "NEW"
      Height          =   495
      Left            =   4800
      TabIndex        =   5
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BACK"
      Height          =   495
      Left            =   3360
      TabIndex        =   4
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   5280
      TabIndex        =   3
      Top             =   2520
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   5280
      TabIndex        =   2
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "PRODUCT VARIANTS"
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "PRODUCT_ID"
      Height          =   495
      Left            =   3360
      TabIndex        =   0
      Top             =   1560
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
MDIForm1.Show
End Sub

Private Sub Command3_Click()
Dim STR As String
STR = "INSERT INTO PRODUCT_VARIANTS('" & Text1.Text & "' ,'" & Text2.Text & "')"
CON.Execute STR
CON.Execute "COMMIT"
MsgBox "ADDED SUCESSFULLY"
Text1.Text = ""
Text2.Text = ""

End Sub
