VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DeleteAllforbook 
   Caption         =   "DeleteALL(�w���A��)"
   ClientHeight    =   3850
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   5640
   OleObjectBlob   =   "DeleteAllforbook.frx":0000
   StartUpPosition =   1  '���ݵ�������
End
Attribute VB_Name = "DeleteAllforbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub CommandButton1sss_Click()
'�M��CustomerAdd�ҿ�J����,�åB�������e��
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
'�_�����s,�������e��,���}CustomerAdd
Me.Hide
Unload Me

End Sub

