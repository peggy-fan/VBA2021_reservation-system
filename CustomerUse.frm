VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CustomerUse 
   Caption         =   "顧客使用介面"
   ClientHeight    =   6930
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   9750.001
   OleObjectBlob   =   "CustomerUse.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "CustomerUse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub backfornt_Click()

'2021/7/18，跳出視窗'
Me.Hide
Unload Me
ManageSystem.Show

End Sub

Private Sub booksys_Click()

'2021/7/18，跳出視窗'
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

Set dtrange = Sheets("會員基本資料").UsedRange                          '設定欲選取的工作表及範圍

rowCnt = dtrange.Rows.Count

Sheets("會員基本資料").Select

For rIdx = 2 To rowCnt                                                  '會員到期日為入會後一年
 Cells(rIdx, "F").Value = DateAdd("yyyy", 1, Cells(rIdx, "E").Value)
Next

End Sub

Private Sub ProductBook_Click()

'2021/7/18，跳出視窗'
Me.Hide
Unload Me
ProductBookSystem.Show

End Sub

Private Sub UserForm_Click()

End Sub
