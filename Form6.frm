VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form6 
   Caption         =   "Form"
   ClientHeight    =   5820
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   16815
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   11.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form6"
   MDIChild        =   -1  'True
   ScaleHeight     =   12495
   ScaleWidth      =   22920
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "PRINT"
      Height          =   495
      Left            =   11400
      TabIndex        =   16
      Top             =   10320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CLOSE"
      Height          =   495
      Left            =   9840
      TabIndex        =   15
      Top             =   10320
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "LAST"
      Height          =   495
      Left            =   12720
      TabIndex        =   14
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "NEXT"
      Height          =   495
      Left            =   11160
      TabIndex        =   13
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "PREVIOUS"
      Height          =   495
      Left            =   9600
      TabIndex        =   12
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "FIRST"
      Height          =   495
      Left            =   8040
      TabIndex        =   11
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Caption         =   "PRODUCTS LIST"
      Height          =   7335
      Left            =   7560
      TabIndex        =   0
      Top             =   360
      Width           =   6855
      Begin VB.TextBox Text5 
         DataField       =   "S_RATE"
         DataSource      =   "Adodc1"
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
         Left            =   3360
         TabIndex        =   21
         Top             =   5160
         Width           =   1695
      End
      Begin VB.TextBox Text3 
         DataField       =   "GST_RATE"
         DataSource      =   "Adodc1"
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
         Left            =   3360
         TabIndex        =   20
         Top             =   4440
         Width           =   1215
      End
      Begin VB.TextBox Text12 
         DataField       =   "P_GST"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   495
         Left            =   3360
         TabIndex        =   10
         Top             =   3720
         Width           =   1095
      End
      Begin VB.TextBox Text11 
         DataField       =   "P_RATE"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   495
         Left            =   3360
         TabIndex        =   9
         Top             =   3000
         Width           =   1335
      End
      Begin VB.TextBox Text10 
         DataField       =   "P_UNIT"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   495
         Left            =   3360
         TabIndex        =   8
         Top             =   2280
         Width           =   2055
      End
      Begin VB.TextBox Text9 
         DataField       =   "P_NAME"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   495
         Left            =   3360
         TabIndex        =   7
         Top             =   1560
         Width           =   2055
      End
      Begin VB.TextBox Text8 
         DataField       =   "P_NO"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   495
         Left            =   3360
         TabIndex        =   6
         Top             =   840
         Width           =   855
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
         Left            =   1440
         TabIndex        =   19
         Top             =   5160
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
         Left            =   1440
         TabIndex        =   18
         Top             =   4560
         Width           =   1455
      End
      Begin VB.Shape Shape6 
         BorderWidth     =   3
         Height          =   975
         Left            =   120
         Top             =   6240
         Width           =   6615
      End
      Begin VB.Shape Shape5 
         Height          =   735
         Left            =   5040
         Top             =   6360
         Width           =   1455
      End
      Begin VB.Shape Shape4 
         Height          =   735
         Left            =   3480
         Top             =   6360
         Width           =   1455
      End
      Begin VB.Shape Shape3 
         Height          =   735
         Left            =   1920
         Top             =   6360
         Width           =   1455
      End
      Begin VB.Shape Shape2 
         Height          =   735
         Left            =   360
         Top             =   6360
         Width           =   1455
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   3
         Height          =   5415
         Left            =   960
         Top             =   600
         Width           =   4935
      End
      Begin VB.Label Label15 
         Caption         =   "GST"
         Height          =   495
         Left            =   1440
         TabIndex        =   5
         Top             =   3840
         Width           =   1575
      End
      Begin VB.Label Label14 
         Caption         =   "PRICE"
         Height          =   495
         Left            =   1440
         TabIndex        =   4
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Label Label13 
         Caption         =   "VARIETY TYPE"
         Height          =   495
         Left            =   1440
         TabIndex        =   3
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label Label12 
         Caption         =   "CAKE VARIETY"
         Height          =   495
         Left            =   1440
         TabIndex        =   2
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label11 
         Caption         =   "PRODUCT_NO"
         Height          =   495
         Left            =   1440
         TabIndex        =   1
         Top             =   960
         Width           =   1695
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   10080
      Top             =   11640
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
      RecordSource    =   "SELECT * FROM PRODUCT order by P_NO"
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
      Bindings        =   "Form6.frx":0000
      Height          =   2175
      Index           =   1
      Left            =   5880
      TabIndex        =   17
      Top             =   7800
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
   Begin VB.Shape Shape10 
      BorderWidth     =   3
      Height          =   975
      Left            =   9600
      Top             =   10080
      Width           =   3255
   End
   Begin VB.Shape Shape9 
      Height          =   735
      Left            =   11280
      Top             =   10200
      Width           =   1455
   End
   Begin VB.Shape Shape8 
      Height          =   735
      Left            =   9720
      Top             =   10200
      Width           =   1455
   End
   Begin VB.Shape Shape7 
      BorderWidth     =   3
      Height          =   11055
      Left            =   4680
      Top             =   120
      Width           =   12975
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
DataReport1.Show
End Sub

Private Sub Command5_Click()
Adodc1.Recordset.MoveFirst
End Sub

Private Sub Command6_Click()
Adodc1.Recordset.MovePrevious
If Adodc1.Recordset.BOF = True Then
Adodc1.Recordset.MoveFirst
End If
End Sub

Private Sub Command7_Click()
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF = True Then
Adodc1.Recordset.MoveLast
End If
End Sub

Private Sub Command8_Click()
Adodc1.Recordset.MoveLast
End Sub


