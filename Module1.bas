Attribute VB_Name = "Module1"
Option Explicit

Sub datasets() '��ƶפJ�����]�m
Dim rowCnt  As Integer '�w�qrowCnt�����
Dim comment As String  '�w�qcomment���r��
Dim dtrange As Range   '�w�qdtRange��Range��
Set dtrange = Sheets("�w���t�Τ���").UsedRange
Dim mon As Integer

rowCnt = dtrange.Rows.Count
Sheets("�w���t�Τ���").Select
'�п�ܶ}�P���O(�~�v(OB01),�S�v(OB02),�žv(OB03),�V�v(OB04),�@�v(OB05))

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


