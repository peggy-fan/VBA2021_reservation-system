VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DeleteAllforProduct 
   Caption         =   "DeleteAll(�q�ʲ��~)"
   ClientHeight    =   3510
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   5925
   OleObjectBlob   =   "DeleteAllforProduct.frx":0000
   StartUpPosition =   1  '���ݵ�������
End
Attribute VB_Name = "DeleteAllforProduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub CommandButton1ssss_Click()
'�M��CustomerAdd�ҿ�J����,�åB�������e��
ProductBookSystem.cusnamee.Text = ""

ProductBookSystem.cusphonee.Text = ""

ProductBookSystem.productbookk.Text = ""

ProductBookSystem.bookknum.Text = ""

Me.Hide
Unload Me
 
End Sub

Sub CommandButton2ssss_Click()
'�_�����s,�������e��,���}CustomerAdd
Me.Hide
Unload Me

End Sub

