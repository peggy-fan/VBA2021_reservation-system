Attribute VB_Name = "Module1"
Option Explicit

Sub datasets() '資料匯入介面設置
Dim rowCnt  As Integer '定義rowCnt為整數
Dim comment As String  '定義comment為字串
Dim dtrange As Range   '定義dtRange為Range值
Set dtrange = Sheets("預約系統介面").UsedRange
Dim mon As Integer

rowCnt = dtrange.Rows.Count
Sheets("預約系統介面").Select
'請選擇開銷類別(洗髮(OB01),燙髮(OB02),剪髮(OB03),染髮(OB04),護髮(OB05))

If (CustomerBookSystem.OB01.Value = True) Then
   Cells(rowCnt, "G").Value = "300"
End If

If (CustomerBookSystem.OB02.Value = True) Then
   Cells(rowCnt, "H").Value = "2800"
End If

If (CustomerBookSystem.OB03.Value = True) Then
   Cells(rowCnt, "I").Value = "800"
End If

If (CustomerBookSystem.OB04.Value = True) Then
   Cells(rowCnt, "J").Value = "1800"
End If

If (CustomerBookSystem.OB05.Value = True) Then
   Cells(rowCnt, "K").Value = "2400"
End If



Cells(rowCnt, "N").Value = Cells(rowCnt, "G").Value + Cells(rowCnt, "H").Value + Cells(rowCnt, "I").Value + Cells(rowCnt, "J").Value + Cells(rowCnt, "K").Value

If Cells(rowCnt, "M").Value = "Y" Then
   Cells(rowCnt, "N").Value = Cells(rowCnt, "N").Value * 0.85
Else
   Cells(rowCnt, "N").Value = Cells(rowCnt, "N").Value * 1
 
End If

CustomerBookSystem.OB01.Value = False
CustomerBookSystem.OB02.Value = False
CustomerBookSystem.OB03.Value = False
CustomerBookSystem.OB04.Value = False
CustomerBookSystem.OB05.Value = False

End Sub


