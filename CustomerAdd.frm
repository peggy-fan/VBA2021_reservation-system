VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CustomerAdd 
   Caption         =   "�s�W�|�����"
   ClientHeight    =   4680
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   7995
   OleObjectBlob   =   "CustomerAdd.frx":0000
   StartUpPosition =   1  '���ݵ�������
End
Attribute VB_Name = "CustomerAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Sub CTAdd_Click()

Dim rowCntx As Integer                                                                   '�ŧi�ܼ�'

Dim dtRangex As Range                                                                    '�ŧi�d��'

Sheets("�|���򥻸��").Select                                                            '����|���򥻸�Ƥu�@��

Set dtRangex = Sheets("�|���򥻸��").UsedRange                                          '�]�w�w�ϥνd��'

rowCntx = dtRangex.Rows.Count                                                            '�]�w�w�ϥΦC���'

Cells(rowCntx + 1, "B").Value = customenam.Text                                          '��J�U�ȩm�W'

Cells(rowCntx + 1, "C").Value = customephon.Text                                        '��J�q�ܸ��X'

Cells(rowCntx + 1, "D").Value = customebirt.Text                                        '��J�U�ȥͤ�'

Cells(rowCntx + 1, "E").Value = customeadddat.Text                                      '��J�U�ȪA�Ȥ��'

                           
'2021/7/17�A�s�W�s���ǦC'
Dim rIdx As Integer
Dim i As Integer
Dim rowCnt As Integer
Dim dtrange As Range
Set dtrange = Sheets("�|���򥻸��").UsedRange
rowCnt = dtrange.Rows.Count
Sheets("�|���򥻸��").Select
If Cells(rowCnt, "B").Value <> "" Then                                                     '�P�_�A�p�GB��쪺�Ȥ�����Ŧr��'

   For i = 2 To rowCnt                                                                     '�]�wi���ȵ���2��̫�@�C'

    Cells(i, "A").Value = i - 1                                                            'A��쪺�ȵ���{���x�s���m-1'
    
   Next
 
End If

'2021/7/17�A�إߪ��Ϊ��s�W��椺�e'
'2021/07/17 ���b
'�ťտ�J
 If customenam.Text = " " Or customephon.Text = " " Or customebirt.Text = " " Or customeadddat.Text = " " Then
  MsgBox ("�Х��T��g���"), vbInformation
 Exit Sub
 End If
  
'�U�ȹq�ܫ��A
 If VBA.IsNumeric(customephon.Text) = False Then
 MsgBox ("�U�ȹq�ܽп�J�Ʀr"), vbInformation               'vbInformation��ĵ���n
 Exit Sub
 End If

'������A
 If VBA.IsDate(customebirt.Text) = False Then
 MsgBox ("�ж�g���T������A Ex:1999/01/01"), vbInformation
 Exit Sub
 End If
 If VBA.IsDate(customeadddat.Text) = False Then
 MsgBox ("�ж�g���T������A Ex:1999/01/01"), vbInformation
 Exit Sub
 End If
'���Ƹ��
 Dim rIx As Integer
 Dim rowCt As Integer
 Dim dtRangenew As Range
 Set dtRangenew = Sheets("�|���򥻸��").UsedRange
 rowCt = dtRangenew.Rows.Count
 Sheets("�|���򥻸��").Select
 For rIx = 2 To rowCt
 If Cells(rIx, "B").Value = customephon.Text Then
 MsgBox "�q�ܤw����", vbInformation
 Exit Sub
 End If
 Next


customenam.Text = ""
customephon.Text = ""
customebirt.Text = ""
customeadddat.Text = ""

End Sub


'2021/7/17�A���s�W���o�ȡA�̭����Ȯ���'

Sub Delete_Click()
DeleteAll.Show
End Sub

Sub back_Click()
'2021/7/18�A���X����'
Me.Hide
Unload Me
CustomerManagement.Show
End Sub

Private Sub UserForm_Click()

End Sub
