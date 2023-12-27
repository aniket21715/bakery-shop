VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form SuppPrdt 
   Caption         =   "AddSuppPrdt"
   ClientHeight    =   10965
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   20265
   LinkTopic       =   "Form7"
   MDIChild        =   -1  'True
   ScaleHeight     =   10965
   ScaleWidth      =   20265
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command6 
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
      Left            =   9720
      TabIndex        =   19
      Top             =   9840
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "PRINT"
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
      Left            =   11160
      TabIndex        =   18
      Top             =   9840
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "SUPPLIER PRODUCT PRICE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   7200
      TabIndex        =   0
      Top             =   360
      Width           =   7455
      Begin VB.CommandButton Command4 
         Caption         =   "DELETE"
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
         Left            =   5520
         TabIndex        =   16
         Top             =   5760
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
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
         Left            =   4080
         TabIndex        =   15
         Top             =   5760
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "ADD"
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
         Left            =   2160
         TabIndex        =   14
         Top             =   5760
         Width           =   1095
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
         Left            =   720
         TabIndex        =   13
         Top             =   5760
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "SuppPrdt.frx":0000
         Left            =   3480
         List            =   "SuppPrdt.frx":0002
         TabIndex        =   11
         Text            =   "SELECT SUPPLIER"
         Top             =   960
         Width           =   2295
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "SuppPrdt.frx":0004
         Left            =   3480
         List            =   "SuppPrdt.frx":0011
         TabIndex        =   5
         Text            =   "CHOOSE UNIT"
         Top             =   3000
         Width           =   2055
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "SuppPrdt.frx":0050
         Left            =   3480
         List            =   "SuppPrdt.frx":0066
         TabIndex        =   4
         Text            =   "SELECT VARIETY"
         Top             =   2280
         Width           =   2055
      End
      Begin VB.TextBox Text3 
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
         Left            =   3480
         TabIndex        =   3
         Top             =   4440
         Width           =   1215
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
         Left            =   3480
         TabIndex        =   2
         Text            =   "*****"
         Top             =   3600
         Width           =   735
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
         Left            =   3480
         TabIndex        =   1
         Text            =   "*****"
         Top             =   1560
         Width           =   735
      End
      Begin VB.Shape Shape7 
         BorderWidth     =   3
         Height          =   975
         Left            =   3840
         Top             =   5520
         Width           =   3015
      End
      Begin VB.Shape Shape6 
         BorderWidth     =   3
         Height          =   975
         Left            =   480
         Top             =   5520
         Width           =   3015
      End
      Begin VB.Shape Shape5 
         Height          =   735
         Left            =   5400
         Top             =   5640
         Width           =   1335
      End
      Begin VB.Shape Shape4 
         Height          =   735
         Left            =   3960
         Top             =   5640
         Width           =   1335
      End
      Begin VB.Shape Shape3 
         Height          =   735
         Left            =   2040
         Top             =   5640
         Width           =   1335
      End
      Begin VB.Shape Shape2 
         Height          =   735
         Left            =   600
         Top             =   5640
         Width           =   1335
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   3
         Height          =   4455
         Left            =   960
         Top             =   720
         Width           =   5415
      End
      Begin VB.Label Label6 
         Caption         =   "SUPPLIER NAME"
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
         Left            =   1320
         TabIndex        =   12
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label5 
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
         Left            =   1320
         TabIndex        =   10
         Top             =   4560
         Width           =   1455
      End
      Begin VB.Label Label4 
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
         Left            =   1320
         TabIndex        =   9
         Top             =   3720
         Width           =   1335
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
         Left            =   1320
         TabIndex        =   8
         Top             =   3000
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
         Left            =   1320
         TabIndex        =   7
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "SUPPLIER_ID"
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
         Left            =   1320
         TabIndex        =   6
         Top             =   1680
         Width           =   1335
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "SuppPrdt.frx":00B7
      Height          =   2175
      Left            =   7680
      TabIndex        =   17
      Top             =   7320
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   3836
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   19
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "S_ID"
         Caption         =   "Supplier Id"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "P_NO"
         Caption         =   "Product No"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "RATE"
         Caption         =   "Price"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1964.976
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2340.284
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   3195.213
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   15840
      Top             =   10680
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1296
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   0
      Connect         =   "Provider=MSDAORA.1;Password=PURBEY;User ID=HARSH;Persist Security Info=True"
      OLEDBString     =   "Provider=MSDAORA.1;Password=PURBEY;User ID=HARSH;Persist Security Info=True"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * FROM SUPPLIER_PRODUCT ORDER BY S_ID"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Shape Shape11 
      BorderWidth     =   3
      Height          =   10695
      Left            =   5760
      Top             =   120
      Width           =   10335
   End
   Begin VB.Shape Shape10 
      BorderWidth     =   3
      Height          =   975
      Left            =   9480
      Top             =   9600
      Width           =   3015
   End
   Begin VB.Shape Shape9 
      Height          =   735
      Left            =   11040
      Top             =   9720
      Width           =   1335
   End
   Begin VB.Shape Shape8 
      Height          =   735
      Left            =   9600
      Top             =   9720
      Width           =   1335
   End
End
Attribute VB_Name = "SuppPrdt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Click()
If rs.State = 1 Then rs.Close
rs.Open "select S_ID from SUPPLIER WHERE S_NAME= '" & Combo1.Text & "' ", con, adOpenDynamic
Text1.Text = rs.Fields(0)
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Combo3_Click()
If rs.State = 1 Then rs.Close
rs.Open "select P_NO from PRODUCT WHERE P_NAME= '" & Combo2.Text & "' AND P_UNIT ='" & Combo3.Text & "' ", con, adOpenDynamic
Text2.Text = rs.Fields(0)
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Combo3_LostFocus()
If rs.State = 1 Then rs.Close
rs.Open "SELECT RATE FROM SUPPLIER_PRODUCT WHERE S_ID ='" & Text1.Text & "' AND P_NO ='" & Text2.Text & "'"

End Sub

Private Sub Command1_Click()
Combo1.Text = "SELECT SUPPLIER"
Text1.Text = "*****"
Combo2.Text = "SELECT VARIETY"
Combo3.Text = "CHOOSE UNIT"
Text2.Text = "*****"
Text3.Text = ""
End Sub

Private Sub Command2_Click()
Dim STR As String
STR = "INSERT INTO SUPPLIER_PRODUCT VALUES('" & Text1.Text & "','" & Text2.Text & "'," & Val(Text3.Text) & ")"
con.Execute STR
con.Execute "COMMIT"
Adodc1.Refresh
End Sub

Private Sub Command3_Click()
Dim STR As String
STR = "UPDATE SUPPLIER_PRODUCT SET RATE = " & Val(Text3.Text) & " WHERE S_ID ='" & Text1.Text & "' AND P_NO ='" & Text2.Text & "' "
con.Execute STR
con.Execute "COMMIT"
Adodc1.Refresh
End Sub

Private Sub Command4_Click()
Dim STR As String
STR = "DELETE FROM SUPPLIER_PRODUCT WHERE S_ID ='" & Text1.Text & "' AND P_NO ='" & Text2.Text & "' "
con.Execute STR
con.Execute "COMMIT"
Adodc1.Refresh
End Sub

Private Sub Command6_Click()
Unload Me
End Sub

Private Sub Form_Activate()
If rs.State = 1 Then rs.Close
rs.Open "select * from SUPPLIER ", con, adOpenDynamic
While rs.EOF = False
Combo1.AddItem rs.Fields(1)
rs.MoveNext
Wend
End Sub

Private Sub Text6_Change()

End Sub

