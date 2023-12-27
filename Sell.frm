VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Sell 
   Caption         =   "SELLING"
   ClientHeight    =   9075
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13245
   LinkTopic       =   "Sell"
   MDIChild        =   -1  'True
   ScaleHeight     =   9075
   ScaleWidth      =   13245
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text18 
      Height          =   495
      Left            =   1320
      TabIndex        =   63
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text17 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10560
      TabIndex        =   62
      Top             =   4680
      Width           =   975
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   735
      Left            =   9960
      TabIndex        =   57
      Top             =   11640
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      _Version        =   393216
      Format          =   117833729
      CurrentDate     =   44957
   End
   Begin VB.CommandButton Command4 
      Caption         =   "CANCEL"
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
      Left            =   17040
      TabIndex        =   56
      Top             =   10560
      Width           =   1095
   End
   Begin VB.TextBox Text14 
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
      Left            =   13560
      TabIndex        =   55
      Top             =   9720
      Width           =   1455
   End
   Begin VB.TextBox Text13 
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
      Left            =   13560
      TabIndex        =   54
      Top             =   10560
      Width           =   1455
   End
   Begin VB.TextBox Text12 
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
      Left            =   8280
      TabIndex        =   51
      Top             =   10560
      Width           =   2415
   End
   Begin VB.ListBox List12 
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
      Height          =   2220
      ItemData        =   "Sell.frx":0000
      Left            =   16680
      List            =   "Sell.frx":0002
      TabIndex        =   49
      Top             =   6000
      Width           =   1815
   End
   Begin VB.ListBox List11 
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
      Height          =   2220
      ItemData        =   "Sell.frx":0004
      Left            =   12960
      List            =   "Sell.frx":0006
      TabIndex        =   44
      Top             =   6000
      Width           =   1815
   End
   Begin VB.ListBox List10 
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
      Height          =   2220
      ItemData        =   "Sell.frx":0008
      Left            =   11520
      List            =   "Sell.frx":000A
      TabIndex        =   42
      Top             =   6000
      Width           =   1455
   End
   Begin VB.ListBox List9 
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
      Height          =   2220
      ItemData        =   "Sell.frx":000C
      Left            =   9720
      List            =   "Sell.frx":000E
      TabIndex        =   40
      Top             =   6000
      Width           =   1815
   End
   Begin VB.ListBox List8 
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
      Height          =   2220
      ItemData        =   "Sell.frx":0010
      Left            =   7200
      List            =   "Sell.frx":0012
      TabIndex        =   38
      Top             =   6000
      Width           =   2535
   End
   Begin VB.ListBox List7 
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
      Height          =   2220
      ItemData        =   "Sell.frx":0014
      Left            =   4920
      List            =   "Sell.frx":0016
      TabIndex        =   36
      Top             =   6000
      Width           =   2295
   End
   Begin VB.ListBox List6 
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
      Height          =   2220
      ItemData        =   "Sell.frx":0018
      Left            =   3240
      List            =   "Sell.frx":001A
      TabIndex        =   34
      Top             =   6000
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
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
      Left            =   20280
      TabIndex        =   31
      Top             =   6480
      Width           =   1095
   End
   Begin VB.TextBox Text11 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      TabIndex        =   29
      Top             =   4680
      Width           =   975
   End
   Begin VB.TextBox Text10 
      Alignment       =   2  'Center
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
      Left            =   8280
      TabIndex        =   28
      Top             =   9720
      Width           =   855
   End
   Begin VB.ListBox List5 
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
      Height          =   2220
      ItemData        =   "Sell.frx":001C
      Left            =   15480
      List            =   "Sell.frx":001E
      TabIndex        =   26
      Top             =   6000
      Width           =   1215
   End
   Begin VB.ListBox List4 
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
      Height          =   2220
      ItemData        =   "Sell.frx":0020
      Left            =   14760
      List            =   "Sell.frx":0022
      TabIndex        =   25
      Top             =   6000
      Width           =   735
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
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
      Height          =   375
      Left            =   15600
      TabIndex        =   23
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "SELL"
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
      Left            =   17040
      TabIndex        =   22
      Top             =   9720
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
      Index           =   0
      Left            =   20280
      TabIndex        =   21
      Top             =   5640
      Width           =   1095
   End
   Begin VB.ListBox List1 
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
      Height          =   2220
      ItemData        =   "Sell.frx":0024
      Left            =   2520
      List            =   "Sell.frx":0026
      TabIndex        =   20
      Top             =   6000
      Width           =   735
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
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
      Height          =   375
      Left            =   17400
      TabIndex        =   18
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
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
      Height          =   375
      Left            =   14160
      TabIndex        =   17
      Top             =   4680
      Width           =   735
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
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
      Height          =   375
      Left            =   12360
      TabIndex        =   15
      Top             =   4680
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
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
      Height          =   375
      Left            =   7080
      TabIndex        =   13
      Top             =   4680
      Width           =   855
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
      ItemData        =   "Sell.frx":0028
      Left            =   4440
      List            =   "Sell.frx":0035
      TabIndex        =   10
      Text            =   "CHOOSE UNIT"
      Top             =   4680
      Width           =   2055
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
      ItemData        =   "Sell.frx":0074
      Left            =   2040
      List            =   "Sell.frx":008A
      TabIndex        =   8
      Text            =   "SELECT VARIETIES"
      Top             =   4680
      Width           =   2055
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
      Left            =   13920
      TabIndex        =   7
      Top             =   840
      Width           =   2175
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
      Left            =   13920
      TabIndex        =   5
      Top             =   1680
      Width           =   2415
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
      Left            =   7560
      TabIndex        =   3
      Top             =   1680
      Width           =   2415
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
      Left            =   7560
      TabIndex        =   1
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label32 
      Alignment       =   2  'Center
      Caption         =   "AVL QUANTITY"
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
      Left            =   10080
      TabIndex        =   61
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      Caption         =   "CALCULATIONS"
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
      Left            =   4680
      TabIndex        =   60
      Top             =   9240
      Width           =   2535
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Caption         =   "ADD PRODUCTS"
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
      TabIndex        =   59
      Top             =   3000
      Width           =   2535
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      Caption         =   "MAIN INFORMATION"
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
      Left            =   4680
      TabIndex        =   58
      Top             =   240
      Width           =   2535
   End
   Begin VB.Shape Shape9 
      BorderWidth     =   3
      Height          =   1815
      Left            =   16800
      Top             =   9480
      Width           =   1575
   End
   Begin VB.Shape Shape8 
      Height          =   735
      Left            =   16920
      Top             =   10440
      Width           =   1335
   End
   Begin VB.Shape Shape7 
      Height          =   735
      Left            =   16920
      Top             =   9600
      Width           =   1335
   End
   Begin VB.Shape Shape6 
      BorderWidth     =   3
      Height          =   2055
      Left            =   4560
      Top             =   9360
      Width           =   11775
   End
   Begin VB.Label Label27 
      Caption         =   "TOTAL AMOUNT"
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
      Left            =   11160
      TabIndex        =   53
      Top             =   9840
      Width           =   1815
   End
   Begin VB.Label Label25 
      Caption         =   "TOTAL WITH TAX"
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
      Left            =   11160
      TabIndex        =   52
      Top             =   10680
      Width           =   1815
   End
   Begin VB.Label Label23 
      Caption         =   "PAYMENT MODE"
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
      Left            =   6000
      TabIndex        =   50
      Top             =   10680
      Width           =   1815
   End
   Begin VB.Shape Shape5 
      BorderWidth     =   3
      Height          =   1815
      Left            =   20040
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Shape Shape4 
      Height          =   735
      Left            =   20160
      Top             =   6360
      Width           =   1335
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Left            =   20160
      Top             =   5520
      Width           =   1335
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   3
      Height          =   5775
      Left            =   1320
      Top             =   3120
      Width           =   18255
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "         TOTAL           PRICE  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   16680
      TabIndex        =   48
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "AMOUNT"
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
      Left            =   15480
      TabIndex        =   47
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "%"
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
      Left            =   14760
      TabIndex        =   46
      Top             =   5760
      Width           =   735
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
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
      Left            =   14760
      TabIndex        =   45
      Top             =   5400
      Width           =   1935
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
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
      Height          =   615
      Left            =   12960
      TabIndex        =   43
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "QUANTITY"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11520
      TabIndex        =   41
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "RATE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9720
      TabIndex        =   39
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   2295
      Left            =   4560
      Top             =   360
      Width           =   13095
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
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
      Height          =   615
      Left            =   7200
      TabIndex        =   37
      Top             =   5400
      Width           =   2535
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
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
      Height          =   615
      Left            =   4920
      TabIndex        =   35
      Top             =   5400
      Width           =   2295
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PRODUCT NO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      TabIndex        =   33
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Label LABEL100 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SERIAL    NO."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      TabIndex        =   32
      Top             =   5400
      Width           =   735
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "QUANTITY"
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
      Left            =   8520
      TabIndex        =   30
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label13 
      Caption         =   "NO OF PRODUCTS"
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
      Left            =   6000
      TabIndex        =   27
      Top             =   9840
      Width           =   1815
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
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
      Left            =   15480
      TabIndex        =   24
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label10 
      Caption         =   "TOTAL PRICE"
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
      Left            =   17280
      TabIndex        =   19
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "GST (%)"
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
      Left            =   13920
      TabIndex        =   16
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
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
      Left            =   12240
      TabIndex        =   14
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label Label7 
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
      Left            =   6840
      TabIndex        =   12
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label Label6 
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
      Height          =   375
      Left            =   4680
      TabIndex        =   11
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label Label5 
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
      Left            =   2160
      TabIndex        =   9
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "SELL DATE"
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
      Left            =   11520
      TabIndex        =   6
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label3 
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
      Height          =   375
      Left            =   11520
      TabIndex        =   4
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "CUSTOMER NAME"
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
      Left            =   5280
      TabIndex        =   2
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "SELL_NO"
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
      Left            =   5280
      TabIndex        =   0
      Top             =   840
      Width           =   1335
   End
End
Attribute VB_Name = "Sell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim W As Integer
Dim ZZ As String
Dim KK As Integer
Dim S As Double
Dim P As String
Dim C As Integer




Private Sub Combo1_Click()
If Combo2.Text = "Half Pound(226 gm)" Or Combo2.Text = "Full Pound(453 gm)" Or Combo2.Text = "Two Pound(907 gm)" Then
If rs.State = 1 Then rs.Close
rs.Open "select * from PRODUCT WHERE P_NAME= '" & Combo1.Text & "' AND P_UNIT ='" & Combo2.Text & "' ", con, adOpenDynamic
Text5.Text = rs.Fields(0)

End If

End Sub

Private Sub Combo2_Click()
If rs.State = 1 Then rs.Close
rs.Open "select * from PRODUCT WHERE P_NAME= '" & Combo1.Text & "' AND P_UNIT ='" & Combo2.Text & "' ", con, adOpenDynamic
Text5.Text = rs.Fields(0)
If rs.State = 1 Then rs.Close
rs.Open "select * from STOCK WHERE CAKE_VARITIES= '" & Combo1.Text & "' AND VARITIES_UNIT ='" & Combo2.Text & "' ", con, adOpenDynamic
Text17.Text = rs.Fields(3)

End Sub
Private Sub Command1_Click()
Text14.Text = (Val(Text14.Text) - Val(List11.List(List11.ListCount - 1)))
Text13.Text = (Val(Text13.Text) - Val(List12.List(List12.ListCount - 1)))
List1.RemoveItem (List1.ListCount - 1)
List4.RemoveItem (List4.ListCount - 1)
List5.RemoveItem (List5.ListCount - 1)
List6.RemoveItem (List6.ListCount - 1)
List7.RemoveItem (List7.ListCount - 1)
List8.RemoveItem (List8.ListCount - 1)
List9.RemoveItem (List9.ListCount - 1)
List10.RemoveItem (List10.ListCount - 1)
List11.RemoveItem (List11.ListCount - 1)
List12.RemoveItem (List12.ListCount - 1)
Text10.Text = List1.ListCount
C = C - 1
End Sub

Private Sub Command2_Click(Index As Integer)

C = C + 1
List1.AddItem (C)
List6.AddItem Text5.Text
List7.AddItem Combo1.Text
List8.AddItem Combo2.Text
List4.AddItem Text7.Text
List5.AddItem Text9.Text
List10.AddItem Text11.Text
List11.AddItem Text6.Text
List12.AddItem Text8.Text
List9.AddItem (Val(Text6.Text) / Val(Text11.Text))
Text10.Text = List1.ListCount
Combo1.Text = "SELECT VARIETIES"
Combo2.Text = "CHOOSE UNIT"
Text11.Text = ""
Text5.Text = ""
A = Val(Text6.Text)
b = Val(Text8.Text)
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text17.Text = ""
Text14.Text = A + Val(Text14.Text)
Text13.Text = b + Val(Text13.Text)
End Sub

Private Sub Command3_Click()
Dim STR As String
STR = "INSERT INTO Sell_MASTER VALUES('" & Text1.Text & "','" & Text4.Text & "','" & Text2.Text & "','" & Text3.Text & "'," & Val(Text10.Text) & ",'" & Text12.Text & "'," & Val(Text14.Text) & "," & Val(Text13.Text) & ",'" & Text18.Text & "')"
con.Execute STR
con.Execute "COMMIT"

For I = 0 To List6.ListCount - 1
con.Execute "INSERT INTO Sell_DETAILS VALUES('" & Text1.Text & "','" & List6.List(I) & "'," & Val(List10.List(I)) & "," & Val(List12.List(I)) & ",'" & List7.List(I) & "'," & Val(List9.List(I)) & ",'" & List4.List(I) & "','" & List8.List(I) & "')"
con.Execute "COMMIT"
Next

Dim STR2 As String
For I = 0 To List6.ListCount - 1
STR2 = "UPDATE STOCK SET AVL_QUANITY = " & Val(List10.List(I)) & " WHERE P_NO = '" & List6.List(I) & "' "
con.Execute STR2
con.Execute "COMMIT"
Next


MsgBox "SELL COMPLETED"
STR = Trim(Text1.Text)
If DataEnvironment1.rsCommand4.State = 1 Then DataEnvironment1.rsCommand4.Close
DataEnvironment1.Command4 STR
DataReport4.Show

Unload Me
Load Me
End Sub

Private Sub Command4_Click()
Unload Me
Load Me
End Sub

Private Sub Command5_Click()
Text14.Text = (Val(Text14.Text) - Val(List11.List(Val(Text15.Text) - 1)))
Text13.Text = (Val(Text13.Text) - Val(List12.List(Val(Text15.Text) - 1)))
List1.RemoveItem (Val(Text15.Text) - 1)
List4.RemoveItem (Val(Text15.Text) - 1)
List5.RemoveItem (Val(Text15.Text) - 1)
List6.RemoveItem (Val(Text15.Text) - 1)
List7.RemoveItem (Val(Text15.Text) - 1)
List8.RemoveItem (Val(Text15.Text) - 1)
List9.RemoveItem (Val(Text15.Text) - 1)
List10.RemoveItem (Val(Text15.Text) - 1)
List11.RemoveItem (Val(Text15.Text) - 1)
List12.RemoveItem (Val(Text15.Text) - 1)
Text10.Text = List1.ListCount
C = C - 1
For I = (Val(Text15.Text) - 1) To (List1.ListCount - 1)
List1.List(I) = (Val(List1.List(I)) - 1)
Next
End Sub

Private Sub Command6_Click()
K = Val(List11.List(Val(Text15.Text) - 1)) / Val(List10.List(Val(Text15.Text) - 1))
J = Val(List12.List(Val(Text15.Text) - 1)) / Val(List10.List(Val(Text15.Text) - 1))
List10.RemoveItem (Val(Text15.Text) - 1)
List10.List(Val(Text15.Text) - 1) = Val(Text16.Text)
List11.List(Val(Text15.Text) - 1) = (K * (Val(List11.List(Val(Text15.Text) - 1))))
List12.List(Val(Text15.Text) - 1) = (J * (Val(List12.List(Val(Text15.Text) - 1))))

End Sub

Private Sub Form_Activate()
C = 0
End Sub

Private Sub Form_Load()
If rs.State = 1 Then rs.Close
rs.Open "SELECT * FROM Sell_MASTER", con, adOpenDynamic
If rs.BOF = True Then
If rs.State = 1 Then rs.Close
rs.Open "SELECT count(S_NO) FROM Sell_MASTER", con, adOpenDynamic
A = rs.Fields(0)
Text1.Text = "S" & "O" & "0" & (A + 1)
Text18.Text = "I" & "N" & "V" & "0" & "0" & (A + 1)

P = Text1.Text
ElseIf P >= "SO09" Then
If rs.State = 1 Then rs.Close
rs.Open "SELECT max(S_NO) FROM Sell_MASTER", con, adOpenDynamic
A = Right(rs.Fields(0), 2)
Text1.Text = "S" & "O" & (A + 1)
Text18.Text = "I" & "N" & "V" & "0" & "0" & (A + 1)

P = Text1.Text
Else
If rs.State = 1 Then rs.Close
rs.Open "SELECT max(S_NO) FROM Sell_MASTER", con, adOpenDynamic
A = Right(rs.Fields(0), 2)
Text1.Text = "S" & "O" & "0" & (A + 1)
Text18.Text = "I" & "N" & "V" & "0" & "0" & (A + 1)

P = Text1.Text
End If
Text4.Text = Format(Date, "DD-MMM-YYYY")

'Text18.Text = "I" & "N" & "V" & "0" & "0" & (A + 1)







End Sub

Private Sub Text11_KeyUp(KeyCode As Integer, Shift As Integer)
If rs.State = 1 Then rs.Close
rs.Open "select * from PRODUCT WHERE P_NAME= '" & Combo1.Text & "' AND P_UNIT ='" & Combo2.Text & "' ", con, adOpenDynamic
Text6.Text = rs.Fields(3) * Val(Text11.Text)
Text7.Text = rs.Fields(4) & "%"
Text9.Text = rs.Fields(5) * Val(Text11.Text)
Text8.Text = rs.Fields(6) * Val(Text11.Text)
End Sub

