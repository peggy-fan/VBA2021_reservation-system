VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DeleteAll 
   Caption         =   "DeleteAll"
   ClientHeight    =   3690
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   5550
   OleObjectBlob   =   "DeleteAll.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "DeleteAll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub CommandButton1ss_Click()
'清除CustomerAdd所輸入之值,並且關閉此畫面
CustomerAdd.CustomerName.Text = ""

CustomerAdd.CustomerPhone.Text = ""

CustomerAdd.CustomerBirth.Text = ""

CustomerAdd.CustomerAddDate.Text = ""


Me.Hide
Unload Me
 
End Sub

Sub CommandButton2ss_Click()
'否的按鈕,關閉此畫面,打開CustomerAdd
Me.Hide
Unload Me

End Sub

