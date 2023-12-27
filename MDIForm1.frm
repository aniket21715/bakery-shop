VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Cake Shop Managment System"
   ClientHeight    =   8655
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   14595
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Menu HAccount 
      Caption         =   "Account"
      Begin VB.Menu HCreateUser 
         Caption         =   "CreateUser"
      End
      Begin VB.Menu HUpdateUser 
         Caption         =   "Update User"
      End
      Begin VB.Menu HReplaceUser 
         Caption         =   "ReplaceUser"
      End
   End
   Begin VB.Menu HProduct 
      Caption         =   "Product"
      Begin VB.Menu HAddProducts 
         Caption         =   "Add Products"
      End
      Begin VB.Menu HProductsList 
         Caption         =   "Products List"
      End
      Begin VB.Menu HProductsUpdate 
         Caption         =   "Update/Delete Products"
      End
   End
   Begin VB.Menu HPurchase 
      Caption         =   "Purchase"
      Begin VB.Menu HProductPurchase 
         Caption         =   "Product Ordered To Suppliers"
      End
      Begin VB.Menu HPRFS 
         Caption         =   "Product Recived From Suppliers"
      End
   End
   Begin VB.Menu HSupplier 
      Caption         =   "Supplier"
      Begin VB.Menu HAddSupplier 
         Caption         =   "Add Suppliers"
      End
      Begin VB.Menu HSupplierList 
         Caption         =   "Suppliers List"
      End
      Begin VB.Menu HAddSuppPrdt 
         Caption         =   "Add Supplier Product"
      End
   End
   Begin VB.Menu HStock 
      Caption         =   "Stock"
      Begin VB.Menu HStockAvailable 
         Caption         =   "Stock Availabe"
      End
   End
   Begin VB.Menu HSell 
      Caption         =   "Sell"
      Begin VB.Menu HSelling 
         Caption         =   "Sellling To Customers"
      End
   End
   Begin VB.Menu HReport 
      Caption         =   "Report"
      Begin VB.Menu HProductRpt 
         Caption         =   "Product Report"
      End
      Begin VB.Menu creport 
         Caption         =   "Conditional Report"
      End
      Begin VB.Menu HStockRpt 
         Caption         =   "Stock Report"
      End
   End
   Begin VB.Menu HExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub creport_Click()
Dim str As String
str = InputBox("Enter S No")
If DataEnvironment1.rsCommand3.State = 1 Then DataEnvironment1.rsCommand3.Close
DataEnvironment1.Command3 str
DataReport3.Show
End Sub

Private Sub HAddProducts_Click()
Form4.Show
End Sub

Private Sub HAddSupplier_Click()
AddSupp.Show
End Sub

Private Sub HAddVariants_Click()
Form4.Show
End Sub

Private Sub HAddSuppPrdt_Click()
SuppPrdt.Show
End Sub

Private Sub HCreateUser_Click()
Form1.Show
End Sub

Private Sub HExit_Click()
End
End Sub

Private Sub HProducts_Click()
End Sub

Private Sub HPRFS_Click()
PrdtRecevied.Show
End Sub

Private Sub HProductPurchase_Click()
PrdtPurchase.Show
End Sub

Private Sub HProductRpt_Click()
DataReport1.Show
End Sub

Private Sub HProductsList_Click()
Form6.Show
End Sub

Private Sub HProductsUpdate_Click()
Form5.Show
End Sub

Private Sub HReplaceUser_Click()
Form2.Show
End Sub

Private Sub HSelling_Click()
Sell.Show
End Sub

Private Sub HStockAvailable_Click()
Stockfrm.Show
End Sub

Private Sub HStockRpt_Click()
DataReport2.Show
End Sub

Private Sub HSupplierList_Click()
SuppList.Show
End Sub

Private Sub HUpdateUser_Click()
Form3.Show
End Sub

Private Sub HUpdateVariants_Click()
Form5.Show
End Sub

Private Sub MDIForm_Load()
LOGO.Show
End Sub
