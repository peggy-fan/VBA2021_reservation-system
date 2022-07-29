VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CustomerAdd 
   Caption         =   "新增會員資料"
   ClientHeight    =   4680
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   7995
   OleObjectBlob   =   "CustomerAdd.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "CustomerAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Sub CTAdd_Click()

Dim rowCntx As Integer                                                                   '宣告變數'

Dim dtRangex As Range                                                                    '宣告範圍'

Sheets("會員基本資料").Select                                                            '選取會員基本資料工作表

Set dtRangex = Sheets("會員基本資料").UsedRange                                          '設定已使用範圍'

rowCntx = dtRangex.Rows.Count                                                            '設定已使用列欄位'

Cells(rowCntx + 1, "B").Value = customenam.Text                                          '輸入顧客姓名'

Cells(rowCntx + 1, "C").Value = customephon.Text                                        '輸入電話號碼'

Cells(rowCntx + 1, "D").Value = customebirt.Text                                        '輸入顧客生日'

Cells(rowCntx + 1, "E").Value = customeadddat.Text                                      '輸入顧客服務日期'

                           
'2021/7/17，新增編號序列'
Dim rIdx As Integer
Dim i As Integer
Dim rowCnt As Integer
Dim dtrange As Range
Set dtrange = Sheets("會員基本資料").UsedRange
rowCnt = dtrange.Rows.Count
Sheets("會員基本資料").Select
If Cells(rowCnt, "B").Value <> "" Then                                                     '判斷，如果B欄位的值不等於空字串'

   For i = 2 To rowCnt                                                                     '設定i的值等於2到最後一列'

    Cells(i, "A").Value = i - 1                                                            'A欄位的值等於現有儲存格位置-1'
    
   Next
 
End If

'2021/7/17，建立表單及表單新增表格內容'
'2021/07/17 防呆
'空白輸入
 If customenam.Text = " " Or customephon.Text = " " Or customebirt.Text = " " Or customeadddat.Text = " " Then
  MsgBox ("請正確填寫資料"), vbInformation
 Exit Sub
 End If
  
'顧客電話型態
 If VBA.IsNumeric(customephon.Text) = False Then
 MsgBox ("顧客電話請輸入數字"), vbInformation               'vbInformation為警示聲
 Exit Sub
 End If

'日期型態
 If VBA.IsDate(customebirt.Text) = False Then
 MsgBox ("請填寫正確日期型態 Ex:1999/01/01"), vbInformation
 Exit Sub
 End If
 If VBA.IsDate(customeadddat.Text) = False Then
 MsgBox ("請填寫正確日期型態 Ex:1999/01/01"), vbInformation
 Exit Sub
 End If
'重複資料
 Dim rIx As Integer
 Dim rowCt As Integer
 Dim dtRangenew As Range
 Set dtRangenew = Sheets("會員基本資料").UsedRange
 rowCt = dtRangenew.Rows.Count
 Sheets("會員基本資料").Select
 For rIx = 2 To rowCt
 If Cells(rIx, "B").Value = customephon.Text Then
 MsgBox "電話已重複", vbInformation
 Exit Sub
 End If
 Next


customenam.Text = ""
customephon.Text = ""
customebirt.Text = ""
customeadddat.Text = ""

End Sub


'2021/7/17，讓新增完得值，裡面的值消失'

Sub Delete_Click()
DeleteAll.Show
End Sub

Sub back_Click()
'2021/7/18，跳出視窗'
Me.Hide
Unload Me
CustomerManagement.Show
End Sub

Private Sub UserForm_Click()

End Sub
