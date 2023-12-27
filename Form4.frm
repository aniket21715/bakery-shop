VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form4 
   Caption         =   "AddProducts"
   ClientHeight    =   4065
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8205
   LinkTopic       =   "Form4"
   MDIChild        =   -1  'True
   ScaleHeight     =   4065
   ScaleWidth      =   8205
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
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
      Left            =   10920
      TabIndex        =   14
      Top             =   10920
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "ADD PRODUCTS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7455
      Left            =   7920
      TabIndex        =   0
      Top             =   480
      Width           =   6615
      Begin VB.TextBox Text5 
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
         TabIndex        =   19
         Top             =   5280
         Width           =   1695
      End
      Begin VB.TextBox Text4 
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
         TabIndex        =   18
         Text            =   "12"
         Top             =   3840
         Width           =   1215
      End
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
         Left            =   1320
         TabIndex        =   13
         Top             =   6480
         Width           =   1095
      End
      Begin VB.TextBox Text1 
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
         TabIndex        =   7
         Text            =   "*****"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox Text2 
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
         TabIndex        =   6
         Top             =   3120
         Width           =   1695
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
         TabIndex        =   5
         Top             =   4560
         Width           =   1215
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
         ItemData        =   "Form4.frx":0000
         Left            =   3480
         List            =   "Form4.frx":0016
         TabIndex        =   4
         Text            =   "SELECT VARIETIES"
         Top             =   1680
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
         ItemData        =   "Form4.frx":0067
         Left            =   3480
         List            =   "Form4.frx":0074
         TabIndex        =   3
         Text            =   "CHOOSE UNIT"
         Top             =   2400
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "NEW"
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
         Left            =   2760
         TabIndex        =   2
         Top             =   6480
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
         Left            =   4200
         TabIndex        =   1
         Top             =   6480
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "SELLING PRICE"
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
         TabIndex        =   17
         Top             =   5400
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "GST AMOUNT"
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
         TabIndex        =   16
         Top             =   4680
         Width           =   1455
      End
      Begin VB.Shape Shape3 
         Height          =   735
         Left            =   4080
         Top             =   6360
         Width           =   1335
      End
      Begin VB.Shape Shape2 
         Height          =   735
         Left            =   2640
         Top             =   6360
         Width           =   1335
      End
      Begin VB.Shape Shape1 
         Height          =   735
         Left            =   1200
         Top             =   6360
         Width           =   1335
      End
      Begin VB.Line Line8 
         BorderWidth     =   3
         X1              =   960
         X2              =   5640
         Y1              =   7200
         Y2              =   7200
      End
      Begin VB.Line Line7 
         BorderWidth     =   3
         X1              =   960
         X2              =   5640
         Y1              =   6240
         Y2              =   6240
      End
      Begin VB.Line Line6 
         BorderWidth     =   3
         X1              =   5640
         X2              =   5640
         Y1              =   6240
         Y2              =   7200
      End
      Begin VB.Line Line5 
         BorderWidth     =   3
         X1              =   960
         X2              =   960
         Y1              =   6240
         Y2              =   7200
      End
      Begin VB.Line Line4 
         BorderWidth     =   3
         X1              =   720
         X2              =   6000
         Y1              =   6000
         Y2              =   6000
      End
      Begin VB.Line Line3 
         BorderWidth     =   3
         X1              =   6000
         X2              =   6000
         Y1              =   600
         Y2              =   6000
      End
      Begin VB.Line Line2 
         BorderWidth     =   3
         X1              =   720
         X2              =   6000
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         X1              =   720
         X2              =   720
         Y1              =   600
         Y2              =   6000
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
         Left            =   1320
         TabIndex        =   12
         Top             =   960
         Width           =   1335
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
         TabIndex        =   11
         Top             =   1680
         Width           =   1815
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
         TabIndex        =   10
         Top             =   2400
         Width           =   1575
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
         Left            =   1320
         TabIndex        =   9
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "GST(%)"
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
         TabIndex        =   8
         Top             =   3960
         Width           =   1455
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   10320
      Top             =   11520
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
      RecordSource    =   "SELECT * FROM PRODUCT ORDER BY P_NO"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form4.frx":00B3
      Height          =   2175
      Left            =   6000
      TabIndex        =   15
      Top             =   8400
      Width           =   10695
      _ExtentX        =   18865
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
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "P_NO"
         Caption         =   "ProductNo"
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
         DataField       =   "P_NAME"
         Caption         =   "CakeVarieties"
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
         DataField       =   "P_UNIT"
         Caption         =   "VarietiesUnit"
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
      BeginProperty Column03 
         DataField       =   "P_RATE"
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
      BeginProperty Column04 
         DataField       =   "P_GST"
         Caption         =   "Gst(%)"
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
      BeginProperty Column05 
         DataField       =   "GST_RATE"
         Caption         =   "GstPrice"
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
      BeginProperty Column06 
         DataField       =   "S_RATE"
         Caption         =   "SellingPrice"
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
            ColumnWidth     =   1170.142
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2025.071
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2099.906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1604.976
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   959.811
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column06 
         EndProperty
      EndProperty
   End
   Begin VB.Shape Shape5 
      Height          =   735
      Left            =   10800
      Top             =   10800
      Width           =   1335
   End
   Begin VB.Shape Shape4 
      BorderWidth     =   3
      Height          =   7815
      Left            =   7560
      Top             =   240
      Width           =   7335
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim GST As Double
Dim P As String

Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
KeyAscii = 0

End Sub

Private Sub Command2_Click()
Dim STR As String
STR = "INSERT INTO PRODUCT VALUES('" & Text1.Text & "','" & Combo1.Text & "','" & Combo2.Text & "'," & Val(Text2.Text) & "," & Val(Text4.Text) & "," & Val(Text3.Text) & "," & Val(Text5.Text) & ")"
con.Execute STR
con.Execute "INSERT INTO STOCK VALUES('" & Text1.Text & "','" & Combo1.Text & "','" & Combo2.Text & "'," & 0 & ")"
con.Execute "COMMIT"
MsgBox "ADDED SUCCESSFULLY"
Adodc1.Refresh
Unload Me
Load Me
End Sub

Private Sub Command3_Click()
Text1.Text = "*****"
Combo1.Text = "SELECT"
Combo2.Text = "CHOOSE"
Text2.Text = ""
Text3.Text = ""
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Load()
If rs.State = 1 Then rs.Close
rs.Open "SELECT * FROM PRODUCT", con, adOpenDynamic
If rs.BOF = True Then
If rs.State = 1 Then rs.Close
rs.Open "SELECT count(p_no) FROM PRODUCT", con, adOpenDynamic
A = rs.Fields(0)
Text1.Text = "P" & "0" & (A + 1)
P = Text1.Text
ElseIf P >= "P09" Then
If rs.State = 1 Then rs.Close
rs.Open "SELECT max(p_no) FROM PRODUCT", con, adOpenDynamic
A = Right(rs.Fields(0), 2)
Text1.Text = "P" & (A + 1)
P = Text1.Text
Else
If rs.State = 1 Then rs.Close
rs.Open "SELECT max(p_no) FROM PRODUCT", con, adOpenDynamic
A = Right(rs.Fields(0), 2)
Text1.Text = "P" & "0" & (A + 1)
P = Text1.Text
End If
End Sub

Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
GST = Val(Text2.Text) * (Val(Text4.Text) / 100)
Text3.Text = GST
Text5.Text = Val(Text2.Text) + Val(Text3.Text)
End Sub

