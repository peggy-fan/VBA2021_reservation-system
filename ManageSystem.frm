VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ManageSystem 
   Caption         =   "���j�z�v�U"
   ClientHeight    =   4360
   ClientLeft      =   75
   ClientTop       =   300
   ClientWidth     =   6615
   OleObjectBlob   =   "ManageSystem.frx":0000
   StartUpPosition =   1  '���ݵ�������
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
password = InputBox("�п�J�K�X : ", "�ϥΪ̿�J", " ")

'����D����¾���i�J���a�޲z�t�Τ���
If password <> ans Then
   MsgBox "�K�X��J���~�A�Э��s�T�{ !"
Else
'���U���a�޲z�t�Ϊ����s,�������e��,���}CustomerManagement
Me.Hide
Unload Me
CustomerManagement.Show
End If

 
End Sub

Sub CommandButton2sa_Click()

'���U�U�ȨϥΤ��������s,�������e��,���}CustomerUse
Me.Hide
Unload Me
CustomerUse.Show

End Sub
