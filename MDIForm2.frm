VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   3780
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   6405
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu hprodct 
      Caption         =   "Product"
      Begin VB.Menu productvariants 
         Caption         =   "add product variants"
      End
      Begin VB.Menu producttypeprice 
         Caption         =   "add product type & price"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub productvariants_Click()
Form1.Show


End Sub
