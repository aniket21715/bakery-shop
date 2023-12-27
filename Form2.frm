VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

CON.ConnectionString = "Provider=MSDAORA.1;Password=SHOP;User ID=CAKE;Persist Security Info=True"
CON.Open
RS.Open "SELECT * FROM PRODUCT_VARIANTS ", CON, adOpenDynamic
Hide Me
MDIForm1.Show
End Sub
