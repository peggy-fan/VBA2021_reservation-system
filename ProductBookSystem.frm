VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProductBookSystem 
   Caption         =   "產品訂購系統"
   ClientHeight    =   6980
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   8760.001
   OleObjectBlob   =   "ProductBookSystem.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "ProductBookSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub CommandButton1s_Click()

Dim rowCntxx As Integer                                                                   '宣告變數'

Dim dtRangexx As Range                                                                    '宣告範圍'

'空白輸入
  If cusnamee.Text = " " Or cusphonee.Text = " " Or productbookk.Text = " " Or bookknum.Text = " " Then
  MsgBox ("請正確填寫資料"), vbInformation
  Exit Sub
  End If

Set dtRangexx = Sheets("產品訂購系統介面").UsedRange                                      '設定已使用範圍'
Sheets("產品訂購系統介面").Select
rowCntxx = dtRangexx.Rows.Count                                                           '設定已使用列欄位'

Cells(rowCntxx + 1, "B").Value = cusnamee.Text                                             '輸入顧客姓名'
Cells(rowCntxx + 1, "C").Value = cusphonee.Text                                            '輸入電話號碼'
Cells(rowCntxx + 1, "D").Value = productbookk.Text                                         '輸入訂購產品'
Cells(rowCntxx + 1, "E").Value = bookknum.Text                                             '輸入訂購數量'



                           
'2021/7/17，新增編號序列'
Dim rIdxx As Integer
Dim i As Integer
Dim rowCntxxx As Integer
Dim dtRangexxx As Range
Set dtRangexxx = Sheets("產品訂購系統介面").UsedRange
rowCntxxx = dtRangexxx.Rows.Count
Sheets("產品訂購系統介面").Select
If Cells(rowCntxxx, "B").Value <> "" Then                                                   '判斷，如果B欄位的值不等於空字串'

   For i = 2 To rowCntxxx                                                                   '設定i的值等於2到最後一列'

    Cells(i, "A").Value = i - 1                                                             'A欄位的值等於現有儲存格位置-1'
    
   Next
 
End If

cusnamee.Text = " "        '送出資料後，系統介面重回初始值
cusphonee.Text = " "
productbookk.Text = " "
bookknum.Text = " "

End Sub


Private Sub UserForm_Initialize()

'2021/7/18，新增下拉選項'

productbookk.AddItem "日本FIOLE洗髮乳"  '選擇訂購產品'
productbookk.AddItem "日本FIOLE潤髮乳"
productbookk.AddItem "日本FIOLE染劑"

bookknum.AddItem "1"  '選擇訂購數量'
bookknum.AddItem "2"
bookknum.AddItem "3"
bookknum.AddItem "4"
bookknum.AddItem "5"
bookknum.AddItem "6"
bookknum.AddItem "7"
bookknum.AddItem "8"
bookknum.AddItem "9"
bookknum.AddItem "10"

End Sub

'2021/7/17，讓新增完的值，裡面的值消失'

Sub CommandButton2s_Click()
DeleteAllforProduct.Show
End Sub

Sub CommandButton3s_Click()
'2021/7/18，跳出視窗'
Me.Hide
Unload Me
CustomerUse.Show
End Sub

