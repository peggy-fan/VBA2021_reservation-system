VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ManageSystem 
   Caption         =   "熊大理髮廳"
   ClientHeight    =   4360
   ClientLeft      =   75
   ClientTop       =   300
   ClientWidth     =   6615
   OleObjectBlob   =   "ManageSystem.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "ManageSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub CommandButton1sa_Click()

Dim password As String
Dim ans As String

ans = "1234"
password = InputBox("請輸入密碼 : ", "使用者輸入", " ")

'防止非本店職員進入店家管理系統介面
If password <> ans Then
   MsgBox "密碼輸入錯誤，請重新確認 !"
Else
'按下店家管理系統的按鈕,關閉此畫面,打開CustomerManagement
Me.Hide
Unload Me
CustomerManagement.Show
End If

 
End Sub

Sub CommandButton2sa_Click()

'按下顧客使用介面的按鈕,關閉此畫面,打開CustomerUse
Me.Hide
Unload Me
CustomerUse.Show

End Sub
