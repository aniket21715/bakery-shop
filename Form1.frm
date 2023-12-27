VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4785
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9030
   LinkTopic       =   "Form1"
   ScaleHeight     =   4785
   ScaleWidth      =   9030
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   3840
      TabIndex        =   2
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   3600
      TabIndex        =   1
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3600
      TabIndex        =   0
      Top             =   1200
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CON As New ADODB.Connection
Dim RS As New ADODB.Recordset

Private Sub Command1_Click()
Dim A As String
A = " UPDATE Login set password = '" & Text1.Text & "' "
CON.Execute A
End Sub

Private Sub Form_Load()
CON.ConnectionString = "Provider=MSDAORA.1;Password=PURBEY;User ID=HARSH;Persist Security Info=True"
CON.Open
RS.Open "SELECT * FROM Login", CON, adOpenDynamic
End Sub
