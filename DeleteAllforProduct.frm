VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DeleteAllforProduct 
   Caption         =   "DeleteAll(訂購產品)"
   ClientHeight    =   3510
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   5925
   OleObjectBlob   =   "DeleteAllforProduct.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "DeleteAllforProduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub CommandButton1ssss_Click()
'清除CustomerAdd所輸入之值,並且關閉此畫面
ProductBookSystem.cusnamee.Text = ""

ProductBookSystem.cusphonee.Text = ""

ProductBookSystem.productbookk.Text = ""

ProductBookSystem.bookknum.Text = ""

Me.Hide
Unload Me
 
End Sub

Sub CommandButton2ssss_Click()
'否的按鈕,關閉此畫面,打開CustomerAdd
Me.Hide
Unload Me

End Sub

