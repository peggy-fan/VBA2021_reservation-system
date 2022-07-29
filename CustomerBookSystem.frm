VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CustomerBookSystem 
   Caption         =   "預約系統介面"
   ClientHeight    =   7060
   ClientLeft      =   75
   ClientTop       =   300
   ClientWidth     =   8820.001
   OleObjectBlob   =   "CustomerBookSystem.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "CustomerBookSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Sub CommandButton1_Click()

Dim rowCnt As Integer
Dim dtrange As Range
Set dtrange = Sheets("預約系統介面").UsedRange '設定已使用範圍'
rowCnt = dtrange.Rows.Count                    '設定已使用列欄位'
Sheets("預約系統介面").Select
'空白輸入
  If cusname.Text = " " Or cusphone.Text = " " Or bookmonth.Text = " " Or bookdate.Text = " " Or booktime.Text = " " Then
  MsgBox ("請正確填寫資料"), vbInformation
  Exit Sub
  End If

Cells(rowCnt + 1, "B").Value = cusname.Text                                             '輸入顧客姓名'
Cells(rowCnt + 1, "C").Value = cusphone.Text                                            '輸入顧客電話'
Cells(rowCnt + 1, "D").Value = bookmonth.Text                                           '輸入預約月份'
Cells(rowCnt + 1, "E").Value = bookdate.Text                                            '輸入預約日期'
Cells(rowCnt + 1, "F").Value = booktime.Text                                            '輸入預約時間'
Cells(rowCnt + 1, "L").Value = memo.Text                                                '輸入備註'
datasets

'2021/7/17，新增編號序列'
Dim rIdxx As Integer
Dim i As Integer
Dim rowCntxxx As Integer
Dim dtRangexxx As Range
Set dtRangexxx = Sheets("預約系統介面").UsedRange
rowCntxxx = dtRangexxx.Rows.Count
Sheets("預約系統介面").Select
If Cells(rowCntxxx, "B").Value <> "" Then                                                   '判斷，如果B欄位的值不等於空字串'

   For i = 2 To rowCntxxx                                                                   '設定i的值等於2到最後一列'

    Cells(i, "A").Value = i - 1                                                             'A欄位的值等於現有儲存格位置-1'
    
   Next
 
End If

'判定是否為會員
Dim e As Integer
Dim f As Integer
Dim rowCntt As Integer
Dim dtrangee As Range
Set dtrangee = Sheets("預約系統介面").UsedRange '設定已使用範圍'
rowCntt = dtrangee.Rows.Count                    '設定已使用列欄位'
Sheets("預約系統介面").Select
For e = rowCntt To rowCntt                                           ' 確定是否為會員
    For f = 2 To 100
        If Sheets("預約系統介面").Cells(e, "C").Value = Sheets("會員基本資料").Cells(f, "C").Value Then
        Cells(e, "M").Value = "Y"
        MsgBox ("會員")
        Exit Sub
        Else
        e = e
        Cells(e, "M").Value = "N"
        End If
    Next
Next



cusname.Text = " "     '送出資料後，系統介面重回初始值
cusphone.Text = " "
bookmonth.Text = " "
bookdate.Text = " "
booktime.Text = " "
memo.Text = " "

End Sub


Private Sub UserForm_Initialize()

'2021/7/18，新增下拉選項'

bookmonth.AddItem "1月"  '選擇預約月份'
bookmonth.AddItem "2月"
bookmonth.AddItem "3月"
bookmonth.AddItem "4月"
bookmonth.AddItem "5月"
bookmonth.AddItem "6月"
bookmonth.AddItem "7月"
bookmonth.AddItem "8月"
bookmonth.AddItem "9月"
bookmonth.AddItem "10月"
bookmonth.AddItem "11月"
bookmonth.AddItem "12月"

bookdate.AddItem "1"  '選擇預約日期'
bookdate.AddItem "2"
bookdate.AddItem "3"
bookdate.AddItem "4"
bookdate.AddItem "5"
bookdate.AddItem "6"
bookdate.AddItem "7"
bookdate.AddItem "8"
bookdate.AddItem "9"
bookdate.AddItem "10"
bookdate.AddItem "11"
bookdate.AddItem "12"
bookdate.AddItem "13"
bookdate.AddItem "14"
bookdate.AddItem "15"
bookdate.AddItem "16"
bookdate.AddItem "17"
bookdate.AddItem "18"
bookdate.AddItem "19"
bookdate.AddItem "20"
bookdate.AddItem "21"
bookdate.AddItem "22"
bookdate.AddItem "23"
bookdate.AddItem "24"
bookdate.AddItem "25"
bookdate.AddItem "26"
bookdate.AddItem "27"
bookdate.AddItem "28"
bookdate.AddItem "29"
bookdate.AddItem "30"
bookdate.AddItem "31"

booktime.AddItem "11:00"  '選擇預約時間'
booktime.AddItem "12:00"
booktime.AddItem "13:00"
booktime.AddItem "14:00"
booktime.AddItem "15:00"
booktime.AddItem "16:00"
booktime.AddItem "17:00"
booktime.AddItem "18:00"
booktime.AddItem "19:00"



End Sub

'2021/7/17，讓新增完的值，裡面的值消失'

Sub CommandButton2_Click()
DeleteAllforbook.Show
End Sub

Sub CommandButton3_Click()
'2021/7/18，跳出視窗'
Me.Hide
Unload Me
CustomerUse.Show
End Sub
