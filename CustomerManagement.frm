VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CustomerManagement 
   Caption         =   "���a�޲z�t��"
   ClientHeight    =   7020
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   10200
   OleObjectBlob   =   "CustomerManagement.frx":0000
   StartUpPosition =   1  '���ݵ�������
End
Attribute VB_Name = "CustomerManagement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub backfornt_Click()

'2021/7/18�A���X����'
Me.Hide
Unload Me
ManageSystem.Show

End Sub

Private Sub DataAnalysis_Click()

End Sub

Sub NewCustomer_Click()
Me.Hide
Unload Me
CustomerAdd.Show
End Sub
