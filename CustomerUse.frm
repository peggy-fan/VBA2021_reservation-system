VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CustomerUse 
   Caption         =   "�U�ȨϥΤ���"
   ClientHeight    =   6930
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   9750.001
   OleObjectBlob   =   "CustomerUse.frx":0000
   StartUpPosition =   1  '���ݵ�������
End
Attribute VB_Name = "CustomerUse"
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

Private Sub booksys_Click()

'2021/7/18�A���X����'
Me.Hide
Unload Me
CustomerBookSystem.Show

End Sub

Private Sub Duedate_Click()

Dim rIdx As Integer

Dim rowCnt As Integer

Dim dtrange As Range

Dim ans As Integer

Dim datee As Date

Set dtrange = Sheets("�|���򥻸��").UsedRange                          '�]�w��������u�@��νd��

rowCnt = dtrange.Rows.Count

Sheets("�|���򥻸��").Select

For rIdx = 2 To rowCnt                                                  '�|������鬰�J�|��@�~
 Cells(rIdx, "F").Value = DateAdd("yyyy", 1, Cells(rIdx, "E").Value)
Next

End Sub

Private Sub ProductBook_Click()

'2021/7/18�A���X����'
Me.Hide
Unload Me
ProductBookSystem.Show

End Sub

Private Sub UserForm_Click()

End Sub
