VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DeleteAllforbook 
   Caption         =   "DeleteALL(預約服務)"
   ClientHeight    =   3850
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   5640
   OleObjectBlob   =   "DeleteAllforbook.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "DeleteAllforbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub CommandButton1sss_Click()
'清除CustomerAdd所輸入之值,並且關閉此畫面
CustomerBookSystem.cusname.Text = ""

CustomerBookSystem.cusphone.Text = ""

CustomerBookSystem.bookmonth.Text = ""

CustomerBookSystem.bookdate.Text = ""

CustomerBookSystem.booktime.Text = ""

CustomerBookSystem.bookitem.ControlSource = ""

CustomerBookSystem.memo.Text = ""

CustomerBookSystem.OB01.Value = False
CustomerBookSystem.OB02.Value = False
CustomerBookSystem.OB03.Value = False
CustomerBookSystem.OB04.Value = False
CustomerBookSystem.OB05.Value = False

Me.Hide
Unload Me
 
End Sub

Sub CommandButton2sss_Click()
'否的按鈕,關閉此畫面,打開CustomerAdd
Me.Hide
Unload Me

End Sub

